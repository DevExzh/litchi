//! Package-scoped memos of values derived from one main-document payload
//! generation.
//!
//! Each memo holds at most one entry, keyed exactly like the paragraph index
//! memo: by the identity of the raw main-part allocation on the owning
//! package, or by the immutable source revision on the source-backed package,
//! together with the semantic route. A key never keeps a payload alive by
//! itself, and a value is only ever reused for the generation it was computed
//! from.
//!
//! The memos follow the weighted clean-value rules of ADR 0005 and of the
//! paragraph index memo (performance changes 0680 and 0682):
//!
//! - an entry is clean while only the memo holds its value; a live view pins
//!   it, and dropping an entry never frees a pinned value;
//! - every value is weighted, and a value above the memo's ceiling is returned
//!   to its caller but never retained, so the ceiling controls retention only
//!   and never turns a valid read into a refusal;
//! - a build failure is never retained: the next request recomputes it, so
//!   every error keeps its identity;
//! - [`GenerationMemo::trim_clean`] is the explicit release seam; generation
//!   replacement is not the only way to drop a clean value.
//!
//! Concurrent first requests for one generation are coalesced: the build runs
//! while the memo's lock is held. A build must therefore never reach back into
//! the same memo.

use super::document_part::{ParagraphIndexCacheKey, ParagraphIndexRoute};
use crate::error::Result;
use litchi_core::SourceVersion;
use std::sync::{Arc, Mutex, MutexGuard};

/// Fixed modeled cost of one retained entry beyond its value's payload.
pub(crate) const MEMO_ENTRY_OVERHEAD_BYTES: usize = 128;

/// A value a [`GenerationMemo`] can retain.
pub(crate) trait MemoValue {
    /// The logical bytes the memo retains for this value.
    fn weight_bytes(&self) -> usize;

    /// Whether a caller outside the memo still holds part of this value.
    ///
    /// The memo's own `Arc` is counted separately; a value only needs to
    /// answer for shared parts it hands out by themselves.
    fn shares_parts(&self) -> bool {
        false
    }
}

struct Entry<V> {
    key: ParagraphIndexCacheKey,
    value: Arc<V>,
}

impl<V: MemoValue> Entry<V> {
    fn is_clean(&self) -> bool {
        Arc::strong_count(&self.value) == 1 && !self.value.shares_parts()
    }
}

struct State<V> {
    entry: Option<Entry<V>>,
    hits: u64,
    misses: u64,
    builds: u64,
    clean_evictions: u64,
    pinned_evictions: u64,
    oversized_bypasses: u64,
}

impl<V> Default for State<V> {
    fn default() -> Self {
        Self {
            entry: None,
            hits: 0,
            misses: 0,
            builds: 0,
            clean_evictions: 0,
            pinned_evictions: 0,
            oversized_bypasses: 0,
        }
    }
}

/// A single-entry, weighted, clean-value memo for one payload generation.
pub(crate) struct GenerationMemo<V> {
    state: Mutex<State<V>>,
    ceiling: usize,
}

impl<V: MemoValue> GenerationMemo<V> {
    /// Create an empty memo that retains values of at most `ceiling` bytes.
    pub(crate) fn new(ceiling: usize) -> Self {
        Self {
            state: Mutex::new(State::default()),
            ceiling,
        }
    }

    fn lock_state(&self) -> MutexGuard<'_, State<V>> {
        self.state
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
    }

    fn evict(state: &mut State<V>) {
        let Some(entry) = state.entry.take() else {
            return;
        };
        if entry.is_clean() {
            state.clean_evictions = state.clean_evictions.saturating_add(1);
        } else {
            state.pinned_evictions = state.pinned_evictions.saturating_add(1);
        }
    }

    /// Drop the entry when no caller outside the memo holds any part of it,
    /// or when its raw generation no longer exists.
    ///
    /// Dropping an entry releases only the memo's reference: a value a view
    /// still holds stays alive with that view.
    pub(crate) fn trim_clean(&self) {
        let mut state = self.lock_state();
        if state
            .entry
            .as_ref()
            .is_some_and(|entry| entry.is_clean() || entry.key.is_dead())
        {
            Self::evict(&mut state);
        }
    }

    fn get_or_build_with_key<F>(&self, key: ParagraphIndexCacheKey, build: F) -> Result<Arc<V>>
    where
        F: FnOnce() -> Result<V>,
    {
        let mut state = self.lock_state();
        if let Some(entry) = state.entry.as_ref()
            && entry.key.matches(&key)
        {
            let value = Arc::clone(&entry.value);
            state.hits = state.hits.saturating_add(1);
            return Ok(value);
        }
        state.misses = state.misses.saturating_add(1);
        Self::evict(&mut state);
        let value = Arc::new(build()?);
        state.builds = state.builds.saturating_add(1);
        if value
            .weight_bytes()
            .saturating_add(MEMO_ENTRY_OVERHEAD_BYTES)
            > self.ceiling
        {
            state.oversized_bypasses = state.oversized_bypasses.saturating_add(1);
            return Ok(value);
        }
        state.entry = Some(Entry {
            key,
            value: Arc::clone(&value),
        });
        Ok(value)
    }

    /// Return the value for the payload allocation `raw`, building it on a
    /// miss.
    pub(crate) fn get_or_build<F>(
        &self,
        raw: &Arc<Vec<u8>>,
        route: ParagraphIndexRoute,
        build: F,
    ) -> Result<Arc<V>>
    where
        F: FnOnce() -> Result<V>,
    {
        self.get_or_build_with_key(ParagraphIndexCacheKey::allocation(raw, route), build)
    }

    /// Return the value for the immutable source revision `version`,
    /// building it on a miss.
    pub(crate) fn get_or_build_source<F>(
        &self,
        version: SourceVersion,
        route: ParagraphIndexRoute,
        build: F,
    ) -> Result<Arc<V>>
    where
        F: FnOnce() -> Result<V>,
    {
        self.get_or_build_with_key(ParagraphIndexCacheKey::source(version, route), build)
    }

    #[cfg(test)]
    pub(crate) fn diagnostics(&self) -> MemoDiagnostics {
        let state = self.lock_state();
        let (clean_entries, pinned_entries, retained_weight) =
            state.entry.as_ref().map_or((0, 0, 0), |entry| {
                let weight = entry
                    .value
                    .weight_bytes()
                    .saturating_add(MEMO_ENTRY_OVERHEAD_BYTES);
                if entry.is_clean() {
                    (1, 0, weight)
                } else {
                    (0, 1, weight)
                }
            });
        MemoDiagnostics {
            hits: state.hits,
            misses: state.misses,
            builds: state.builds,
            clean_evictions: state.clean_evictions,
            pinned_evictions: state.pinned_evictions,
            oversized_bypasses: state.oversized_bypasses,
            clean_entries,
            pinned_entries,
            retained_weight,
        }
    }
}

/// Content-free counters for one [`GenerationMemo`].
#[cfg(test)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct MemoDiagnostics {
    pub(crate) hits: u64,
    pub(crate) misses: u64,
    pub(crate) builds: u64,
    pub(crate) clean_evictions: u64,
    pub(crate) pinned_evictions: u64,
    pub(crate) oversized_bypasses: u64,
    pub(crate) clean_entries: usize,
    pub(crate) pinned_entries: usize,
    pub(crate) retained_weight: usize,
}

/// Whether the projection keeps every element of the part where the part
/// has it.
///
/// The document transaction addresses blocks by their position among the
/// direct children of their container in the raw part, while readers count
/// the markup-compatibility (MCE) projection. The two agree exactly when the
/// projection neither selected an `mc:AlternateContent` branch nor unwrapped
/// `mc:ProcessContent` content: the projection then only drops attributes and
/// whole ignorable subtrees, and WordprocessingML elements, which a profile
/// never ignores, keep their parents and their order. They also agree when
/// every selection is an `mc:AlternateContent` inside a run (record 0307):
/// the selected branch replaces its wrapper inside that run, so every block
/// and every run keeps its parent and its order. A part the document
/// transaction admits has only these shapes, so every admitted part is
/// `Preserved`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum RawPositions {
    /// Raw and visible positions are proven equal.
    Preserved,
    /// The projection selected or unwrapped content outside runs; positions
    /// may differ.
    Unproven,
}

impl RawPositions {
    pub(crate) const fn is_preserved(self) -> bool {
        matches!(self, Self::Preserved)
    }
}

/// The MCE-visible main-document payload of one raw generation.
pub(crate) struct VisiblePayload {
    /// The projection, or `None` when it is the raw allocation itself.
    projected: Option<Arc<Vec<u8>>>,
    positions: RawPositions,
}

impl VisiblePayload {
    /// Select markup-compatibility branches for a main-document payload with
    /// the visible main-document profile.
    ///
    /// Both materialized OPC parts and source-backed pinned documents use
    /// this projection, so their semantics remain identical; its counters are
    /// kept for the position verdict, and its errors are the projection's
    /// own.
    pub(crate) fn project(raw: &[u8]) -> Result<Self> {
        let output = litchi_ooxml_common::mce::process_markup_compatibility(
            raw,
            &super::document_part::visible_document_capabilities(),
            &litchi_ooxml_common::mce::Limits::default(),
        )?;
        let preserved = (output.report.alternate_content_count == 0
            && output.report.unwrapped_elements == 0)
            || (output.report.alternate_content_count > 0
                && super::document_part::holds_run_alternate_content(raw));
        Ok(match output.xml {
            std::borrow::Cow::Borrowed(_) => Self {
                projected: None,
                positions: RawPositions::Preserved,
            },
            std::borrow::Cow::Owned(xml) => Self {
                projected: Some(Arc::new(xml)),
                positions: if preserved {
                    RawPositions::Preserved
                } else {
                    RawPositions::Unproven
                },
            },
        })
    }

    /// The visible bytes for the raw generation `raw` this value was built
    /// from.
    pub(crate) fn xml(&self, raw: Arc<Vec<u8>>) -> Arc<Vec<u8>> {
        self.projected.as_ref().map_or(raw, Arc::clone)
    }

    /// The projection when it is not the raw allocation.
    pub(crate) const fn projected(&self) -> Option<&Arc<Vec<u8>>> {
        self.projected.as_ref()
    }

    pub(crate) const fn positions(&self) -> RawPositions {
        self.positions
    }

    /// Whether `cached`, a memoized projection of `raw`, equals this fresh
    /// projection of the same bytes; a debug cross-check for reused values.
    #[cfg(debug_assertions)]
    pub(crate) fn same_projection(&self, cached: &Self, raw: &Arc<Vec<u8>>) -> bool {
        self.positions == cached.positions
            && match (&self.projected, &cached.projected) {
                (None, None) => true,
                (Some(fresh), Some(kept)) => fresh.as_slice() == kept.as_slice(),
                (None, Some(kept)) => kept.as_slice() == raw.as_slice(),
                (Some(_), None) => false,
            }
    }
}

impl MemoValue for VisiblePayload {
    fn weight_bytes(&self) -> usize {
        self.projected.as_ref().map_or(0, |xml| xml.capacity())
    }

    fn shares_parts(&self) -> bool {
        self.projected
            .as_ref()
            .is_some_and(|xml| Arc::strong_count(xml) > 1)
    }
}

impl std::fmt::Debug for VisiblePayload {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str("VisiblePayload { .. }")
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::error::Error;

    #[derive(Debug)]
    struct Weighted(usize);

    impl MemoValue for Weighted {
        fn weight_bytes(&self) -> usize {
            self.0
        }
    }

    fn allocation(bytes: &[u8]) -> Arc<Vec<u8>> {
        Arc::new(bytes.to_vec())
    }

    #[test]
    fn hits_require_the_same_allocation_and_route() {
        let memo = GenerationMemo::<Weighted>::new(1024);
        let raw = allocation(b"<a/>");
        let first = memo
            .get_or_build(&raw, ParagraphIndexRoute::Visible, || Ok(Weighted(1)))
            .unwrap();
        let second = memo
            .get_or_build(&raw, ParagraphIndexRoute::Visible, || {
                Err(Error::InvalidFormat("same allocation must hit".into()))
            })
            .unwrap();
        assert!(Arc::ptr_eq(&first, &second));
        let equal_bytes = allocation(b"<a/>");
        let third = memo
            .get_or_build(&equal_bytes, ParagraphIndexRoute::Visible, || {
                Ok(Weighted(1))
            })
            .unwrap();
        assert!(!Arc::ptr_eq(&first, &third));
        let managed = memo
            .get_or_build(&equal_bytes, ParagraphIndexRoute::Managed, || {
                Ok(Weighted(1))
            })
            .unwrap();
        assert!(!Arc::ptr_eq(&third, &managed));
        let diagnostics = memo.diagnostics();
        assert_eq!(
            (diagnostics.hits, diagnostics.misses, diagnostics.builds),
            (1, 3, 3)
        );
        // The replaced entries were still held by the test, so they were
        // pinned when the memo dropped them.
        assert_eq!(diagnostics.pinned_evictions, 2);
    }

    #[test]
    fn source_revisions_key_independently_of_allocations() {
        let memo = GenerationMemo::<Weighted>::new(1024);
        let version = SourceVersion::new(7, 0);
        memo.get_or_build_source(version, ParagraphIndexRoute::Visible, || Ok(Weighted(2)))
            .unwrap();
        memo.get_or_build_source(version, ParagraphIndexRoute::Visible, || {
            Err(Error::InvalidFormat("same revision must hit".into()))
        })
        .unwrap();
        memo.get_or_build_source(
            SourceVersion::new(7, 1),
            ParagraphIndexRoute::Visible,
            || Ok(Weighted(2)),
        )
        .unwrap();
        let diagnostics = memo.diagnostics();
        assert_eq!((diagnostics.hits, diagnostics.misses), (1, 2));
        assert_eq!(diagnostics.clean_evictions, 1);
        assert_eq!(diagnostics.clean_entries, 1);
        assert_eq!(diagnostics.retained_weight, 2 + MEMO_ENTRY_OVERHEAD_BYTES);
    }

    #[test]
    fn failures_are_not_retained_and_oversized_values_bypass_retention() {
        let memo = GenerationMemo::<Weighted>::new(256);
        let raw = allocation(b"<a/>");
        assert!(
            memo.get_or_build(&raw, ParagraphIndexRoute::Visible, || {
                Err(Error::InvalidFormat("retryable".into()))
            })
            .is_err()
        );
        assert_eq!(memo.diagnostics().clean_entries, 0);
        let large = memo
            .get_or_build(&raw, ParagraphIndexRoute::Visible, || Ok(Weighted(200)))
            .unwrap();
        assert_eq!(large.0, 200, "an oversized value is still returned");
        let diagnostics = memo.diagnostics();
        assert_eq!(diagnostics.oversized_bypasses, 1);
        assert_eq!(diagnostics.clean_entries + diagnostics.pinned_entries, 0);
        let rebuilt = memo
            .get_or_build(&raw, ParagraphIndexRoute::Visible, || Ok(Weighted(100)))
            .unwrap();
        assert_eq!(rebuilt.0, 100, "the bypassed value is recomputed");
        assert_eq!(memo.diagnostics().builds, 2);
    }

    #[test]
    fn trimming_releases_clean_and_dead_entries_only() {
        let memo = GenerationMemo::<Weighted>::new(1024);
        let raw = allocation(b"<a/>");
        let pinned = memo
            .get_or_build(&raw, ParagraphIndexRoute::Visible, || Ok(Weighted(1)))
            .unwrap();
        memo.trim_clean();
        assert_eq!(memo.diagnostics().pinned_entries, 1, "a live value is kept");
        drop(pinned);
        memo.trim_clean();
        let diagnostics = memo.diagnostics();
        assert_eq!(diagnostics.clean_evictions, 1);
        assert_eq!(diagnostics.clean_entries, 0);

        // An entry whose raw generation is gone can never hit again.
        let pinned = memo
            .get_or_build(&raw, ParagraphIndexRoute::Visible, || Ok(Weighted(1)))
            .unwrap();
        drop(raw);
        memo.trim_clean();
        assert_eq!(memo.diagnostics().pinned_evictions, 1);
        assert_eq!(pinned.0, 1, "the caller keeps its value");
    }

    #[test]
    fn concurrent_first_requests_build_once() {
        let memo = Arc::new(GenerationMemo::<Weighted>::new(1024));
        let raw = allocation(b"<a/>");
        let builds = Arc::new(std::sync::atomic::AtomicUsize::new(0));
        std::thread::scope(|scope| {
            for _ in 0..8 {
                let memo = Arc::clone(&memo);
                let raw = Arc::clone(&raw);
                let builds = Arc::clone(&builds);
                scope.spawn(move || {
                    memo.get_or_build(&raw, ParagraphIndexRoute::Visible, || {
                        builds.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
                        Ok(Weighted(1))
                    })
                    .unwrap();
                });
            }
        });
        assert_eq!(builds.load(std::sync::atomic::Ordering::Relaxed), 1);
        assert_eq!(memo.diagnostics().hits, 7);
    }

    #[test]
    fn visible_payload_records_the_projection_and_its_position_verdict() {
        const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        let plain = allocation(
            format!(r#"<w:document xmlns:w="{W}"><w:body><w:p/></w:body></w:document>"#).as_bytes(),
        );
        let payload = VisiblePayload::project(&plain).unwrap();
        assert!(payload.projected().is_none());
        assert!(Arc::ptr_eq(&payload.xml(Arc::clone(&plain)), &plain));
        assert_eq!(payload.weight_bytes(), 0);
        assert!(payload.positions().is_preserved());

        let declared = allocation(
            format!(
                r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}"><w:body><w:p/></w:body></w:document>"#
            )
            .as_bytes(),
        );
        let payload = VisiblePayload::project(&declared).unwrap();
        let projected = Arc::clone(payload.projected().unwrap());
        assert!(payload.positions().is_preserved());
        assert_eq!(payload.weight_bytes(), projected.capacity());
        assert!(payload.shares_parts(), "the test holds the projection");
        drop(projected);
        assert!(!payload.shares_parts());

        let selected = allocation(
            format!(
                r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:x="urn:unknown"><w:body><mc:AlternateContent><mc:Choice Requires="x"><w:tbl/></mc:Choice><mc:Fallback><w:p/></mc:Fallback></mc:AlternateContent></w:body></w:document>"#
            )
            .as_bytes(),
        );
        let payload = VisiblePayload::project(&selected).unwrap();
        assert_eq!(payload.positions(), RawPositions::Unproven);
        assert!(payload.same_projection(&VisiblePayload::project(&selected).unwrap(), &selected));

        // Record 0307: selections inside runs keep every block and run where
        // it is; one outside a run does not.
        let in_run = allocation(
            format!(
                r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:x="urn:unknown"><w:body><w:p><w:r><mc:AlternateContent><mc:Choice Requires="x"><w:t>C</w:t></mc:Choice><mc:Fallback><w:t>F</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p><w:p/></w:body></w:document>"#
            )
            .as_bytes(),
        );
        let payload = VisiblePayload::project(&in_run).unwrap();
        assert!(payload.projected().is_some());
        assert!(payload.positions().is_preserved());
        let in_paragraph = allocation(
            format!(
                r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:x="urn:unknown"><w:body><w:p><mc:AlternateContent><mc:Choice Requires="x"><w:r/></mc:Choice></mc:AlternateContent></w:p></w:body></w:document>"#
            )
            .as_bytes(),
        );
        let payload = VisiblePayload::project(&in_paragraph).unwrap();
        assert_eq!(payload.positions(), RawPositions::Unproven);
    }
}
