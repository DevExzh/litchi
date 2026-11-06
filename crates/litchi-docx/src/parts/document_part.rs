#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::shadow_unrelated,
    reason = "local parser names mirror the OOXML role currently being decoded"
)]
//! DocumentPart - the main document.xml part of a Word document.

use super::memo::{GenerationMemo, RawPositions, VisiblePayload};
use crate::alt::{Chunk, scan};
use crate::error::Result;
use crate::namespace::{
    NamespaceBindings, NamespaceCapture, is_wordprocessing_namespace, scan_word_element_ranges,
    scan_word_element_ranges_with_context,
};
use crate::paragraph::{Paragraph, extract_word_text};
use crate::table::Table;
use litchi_core::xml::ReaderOrigin;
use litchi_core::{ExecutionContext, ExecutionError, Reservation, Resource, SourceVersion};
use litchi_opc::part::Part;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;
use smallvec::SmallVec;
use std::mem::size_of;
use std::sync::{Arc, Mutex, OnceLock, Weak};

/// Maximum number of paragraph ranges retained by the reusable semantic
/// index.  Documents beyond this bound continue to use the established
/// streaming selectors; the cache is an optimization and never changes the
/// accepted document surface.
const MAX_PARAGRAPH_INDEX_RANGES: usize = 1_000_000;

/// One byte range for a visible `w:p` element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ParagraphRange {
    pub(crate) start: u32,
    pub(crate) length: u32,
    pub(crate) namespaces: NamespaceBindings,
}

/// Bounded paragraph offsets and namespace snapshots for one visible XML view.
///
/// The index never owns or exposes XML bytes, so a cached lookup cannot extend
/// the lifetime of a payload or bypass the existing semantic/lossless
/// ownership boundaries.
#[derive(Debug, Clone)]
pub(crate) struct ParagraphIndex {
    ranges: Arc<[ParagraphRange]>,
}

impl ParagraphIndex {
    /// Build an index using the same bounded namespace scanner as the legacy
    /// paragraph selectors.
    pub(crate) fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut ranges = Vec::new();
        scan_word_element_ranges_with_context(
            xml,
            &[],
            &[b"p".as_slice()],
            |_, start, length, namespaces| {
                if ranges.len() >= MAX_PARAGRAPH_INDEX_RANGES {
                    return Err(crate::Error::InvalidFormat(format!(
                        "document paragraph index exceeds {MAX_PARAGRAPH_INDEX_RANGES} ranges"
                    )));
                }
                ranges
                    .try_reserve(1)
                    .map_err(|source| crate::Error::Allocation {
                        resource: "document paragraph index",
                        source,
                    })?;
                ranges.push(ParagraphRange {
                    start,
                    length,
                    namespaces,
                });
                Ok(())
            },
        )?;
        Ok(Self {
            ranges: ranges.into(),
        })
    }

    pub(crate) fn len(&self) -> usize {
        self.ranges.len()
    }

    pub(crate) fn get(&self, index: usize) -> Option<ParagraphRange> {
        self.ranges.get(index).cloned()
    }

    pub(crate) fn iter(&self) -> impl Iterator<Item = &ParagraphRange> + '_ {
        self.ranges.iter()
    }

    fn weight_bytes(&self) -> usize {
        self.ranges
            .len()
            .saturating_mul(size_of::<ParagraphRange>())
            .saturating_add(PARAGRAPH_INDEX_MEMO_OVERHEAD_BYTES)
    }
}

/// A semantic route whose paragraph offsets are valid for one source view.
///
/// Managed source-backed payloads are already in their source XML view and
/// therefore must never share an index with a materialized MCE-visible view,
/// even when both happen to refer to the same allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ParagraphIndexRoute {
    Visible,
    Managed,
}

/// A retained paragraph index and the reservation that admitted it.
///
/// The reservation is attached to the memo rather than to the cache itself.
/// A cache entry can be evicted while a document still uses it; the document's
/// `Arc` then keeps the admitted resource live until that view is dropped.
pub(crate) struct ParagraphIndexMemo {
    pub(crate) index: ParagraphIndex,
    pub(crate) _admission: Option<Arc<DocumentIndexAdmission>>,
}

pub(crate) struct ParagraphIndexBuild {
    pub(crate) memo: Option<Arc<ParagraphIndexMemo>>,
    pub(crate) admission: Option<Arc<DocumentIndexAdmission>>,
}

impl ParagraphIndexBuild {
    pub(crate) fn indexed(
        index: ParagraphIndex,
        admission: Option<Arc<DocumentIndexAdmission>>,
    ) -> Self {
        Self {
            memo: Some(Arc::new(ParagraphIndexMemo {
                index,
                _admission: admission,
            })),
            admission: None,
        }
    }

    pub(crate) const fn unavailable(admission: Option<Arc<DocumentIndexAdmission>>) -> Self {
        Self {
            memo: None,
            admission,
        }
    }
}

const PARAGRAPH_INDEX_MEMO_OVERHEAD_BYTES: usize = 128;
const PARAGRAPH_INDEX_CACHE_MAX_BYTES: usize = 8 * 1024 * 1024;

struct ParagraphIndexCacheEntry {
    key: ParagraphIndexCacheKey,
    memo: Arc<ParagraphIndexMemo>,
}

/// Identity of the payload generation a memoized value was derived from: the
/// raw main-part allocation (eager package) or the immutable source revision
/// (source-backed package), with the semantic route.
pub(crate) enum ParagraphIndexCacheKey {
    Allocation {
        raw: Weak<Vec<u8>>,
        route: ParagraphIndexRoute,
    },
    Source {
        version: SourceVersion,
        route: ParagraphIndexRoute,
    },
}

impl ParagraphIndexCacheKey {
    pub(crate) fn allocation(raw: &Arc<Vec<u8>>, route: ParagraphIndexRoute) -> Self {
        Self::Allocation {
            raw: Arc::downgrade(raw),
            route,
        }
    }

    pub(crate) const fn source(version: SourceVersion, route: ParagraphIndexRoute) -> Self {
        Self::Source { version, route }
    }

    /// Whether this key names a raw allocation that no longer exists, so no
    /// later request can match it.
    pub(crate) fn is_dead(&self) -> bool {
        match self {
            Self::Allocation { raw, .. } => raw.strong_count() == 0,
            Self::Source { .. } => false,
        }
    }

    pub(crate) fn matches(&self, other: &Self) -> bool {
        match (self, other) {
            (
                Self::Allocation {
                    raw: left,
                    route: left_route,
                },
                Self::Allocation {
                    raw: right,
                    route: right_route,
                },
            ) => {
                *left_route == *right_route
                    && left.upgrade().is_some_and(|candidate| {
                        right
                            .upgrade()
                            .is_some_and(|other| Arc::ptr_eq(&candidate, &other))
                    })
            },
            (
                Self::Source {
                    version: left,
                    route: left_route,
                },
                Self::Source {
                    version: right,
                    route: right_route,
                },
            ) => left == right && left_route == right_route,
            _ => false,
        }
    }
}

#[derive(Default)]
struct ParagraphIndexCacheState {
    entry: Option<ParagraphIndexCacheEntry>,
    hits: u64,
    misses: u64,
    builds: u64,
    clean_evictions: u64,
    pinned_evictions: u64,
    oversized_bypasses: u64,
}

/// Logical clean-retention ceiling for one MCE-visible main-document payload.
///
/// A larger projection is still returned to its document view but is not
/// retained, so each such view recomputes it as before; the ceiling bounds
/// retention and never turns a read into a refusal.
const VISIBLE_PAYLOAD_MEMO_MAX_BYTES: usize = 16 * 1024 * 1024;
/// Logical clean-retention ceiling for one body block index, the paragraph
/// index's ceiling.
const BODY_INDEX_MEMO_MAX_BYTES: usize = PARAGRAPH_INDEX_CACHE_MAX_BYTES;

/// Package-scoped, single-entry memos of values derived from the main
/// document: the paragraph offsets, the markup-compatibility-visible payload
/// and the body block index.
///
/// The paragraph cache retains only a weak identity key and the bounded
/// offset memo. It never keeps XML alive by itself. Holding the state lock
/// while building coalesces concurrent first queries without exposing an
/// in-progress or failed result to another caller; malformed input and
/// admission failures remain retryable and are never negatively cached. The
/// visible-payload and block-index memos follow the same rules
/// ([`super::memo`]) with their own ceilings, and only the unmanaged routes
/// use them: a managed view reads its source bytes and builds neither.
pub(crate) struct ParagraphIndexCache {
    state: Mutex<ParagraphIndexCacheState>,
    visible: GenerationMemo<VisiblePayload>,
    body: GenerationMemo<crate::document::BlockIndex>,
}

impl ParagraphIndexCache {
    pub(crate) fn new() -> Self {
        Self {
            state: Mutex::new(ParagraphIndexCacheState::default()),
            visible: GenerationMemo::new(VISIBLE_PAYLOAD_MEMO_MAX_BYTES),
            body: GenerationMemo::new(BODY_INDEX_MEMO_MAX_BYTES),
        }
    }

    /// The MCE-visible payload of the raw main-part allocation `raw`,
    /// projected once per allocation.
    pub(crate) fn visible_payload(&self, raw: &Arc<Vec<u8>>) -> Result<Arc<VisiblePayload>> {
        let payload = self
            .visible
            .get_or_build(raw, ParagraphIndexRoute::Visible, || {
                VisiblePayload::project(raw.as_slice())
            })?;
        #[cfg(debug_assertions)]
        debug_assert!(
            VisiblePayload::project(raw.as_slice())
                .is_ok_and(|fresh| fresh.same_projection(&payload, raw)),
            "a memoized visible payload differs from a fresh projection"
        );
        Ok(payload)
    }

    /// The MCE-visible payload of the main document of source revision
    /// `version`; `project` loads and projects it on a miss.
    pub(crate) fn visible_payload_source<F>(
        &self,
        version: SourceVersion,
        project: F,
    ) -> Result<Arc<VisiblePayload>>
    where
        F: FnOnce() -> Result<VisiblePayload>,
    {
        self.visible
            .get_or_build_source(version, ParagraphIndexRoute::Visible, project)
    }

    /// The body block index of the raw main-part allocation `raw`, built from
    /// `visible` once per allocation.
    pub(crate) fn body_index(
        &self,
        raw: &Arc<Vec<u8>>,
        visible: &[u8],
        positions: RawPositions,
    ) -> Result<Arc<crate::document::BlockIndex>> {
        let index = self
            .body
            .get_or_build(raw, ParagraphIndexRoute::Visible, || {
                body_block_index(visible, positions)
            })?;
        debug_assert_reused_body_index(&index, visible, positions);
        Ok(index)
    }

    /// The body block index of the main document of source revision
    /// `version`, built from `visible` once per revision.
    pub(crate) fn body_index_source(
        &self,
        version: SourceVersion,
        visible: &[u8],
        positions: RawPositions,
    ) -> Result<Arc<crate::document::BlockIndex>> {
        let index = self
            .body
            .get_or_build_source(version, ParagraphIndexRoute::Visible, || {
                body_block_index(visible, positions)
            })?;
        debug_assert_reused_body_index(&index, visible, positions);
        Ok(index)
    }

    /// Release every clean document memo: the paragraph offsets, the visible
    /// payload and the body block index that no live view pins.
    pub(crate) fn trim_clean_all(&self) {
        self.trim_clean();
        self.visible.trim_clean();
        self.body.trim_clean();
    }

    #[cfg(test)]
    pub(crate) fn visible_diagnostics(&self) -> crate::parts::memo::MemoDiagnostics {
        self.visible.diagnostics()
    }

    #[cfg(test)]
    pub(crate) fn body_diagnostics(&self) -> crate::parts::memo::MemoDiagnostics {
        self.body.diagnostics()
    }

    fn lock_state(&self) -> std::sync::MutexGuard<'_, ParagraphIndexCacheState> {
        self.state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
    }

    fn evict_entry(state: &mut ParagraphIndexCacheState) {
        let Some(entry) = state.entry.take() else {
            return;
        };
        if Arc::strong_count(&entry.memo) == 1 {
            state.clean_evictions = state.clean_evictions.saturating_add(1);
        } else {
            state.pinned_evictions = state.pinned_evictions.saturating_add(1);
        }
    }

    /// Drop an unpinned clean memo before another package operation reserves
    /// workspace. Active documents keep their memo and admission alive.
    pub(crate) fn trim_clean(&self) {
        let mut state = self.lock_state();
        let clean = state
            .entry
            .as_ref()
            .is_some_and(|entry| Arc::strong_count(&entry.memo) == 1);
        if clean {
            Self::evict_entry(&mut state);
        }
    }

    fn get_or_build_with_key<F>(
        &self,
        key: ParagraphIndexCacheKey,
        build: F,
    ) -> Result<ParagraphIndexBuild>
    where
        F: FnOnce() -> Result<ParagraphIndexBuild>,
    {
        let mut state = self.lock_state();
        let hit = state
            .entry
            .as_ref()
            .is_some_and(|entry| entry.key.matches(&key));
        if hit {
            state.hits = state.hits.saturating_add(1);
            let Some(entry) = state.entry.as_ref() else {
                return Err(crate::Error::InvalidFormat(
                    "paragraph index cache hit lost its entry".into(),
                ));
            };
            let memo = Arc::clone(&entry.memo);
            return Ok(ParagraphIndexBuild {
                memo: Some(memo),
                admission: None,
            });
        }

        state.misses = state.misses.saturating_add(1);
        Self::evict_entry(&mut state);
        let built = build()?;
        state.builds = state.builds.saturating_add(1);
        let Some(memo) = built.memo.as_ref() else {
            return Ok(built);
        };
        let weight = memo.index.weight_bytes();
        if weight > PARAGRAPH_INDEX_CACHE_MAX_BYTES {
            state.oversized_bypasses = state.oversized_bypasses.saturating_add(1);
            return Ok(built);
        }
        state.entry = Some(ParagraphIndexCacheEntry {
            key,
            memo: Arc::clone(memo),
        });
        Ok(built)
    }

    pub(crate) fn get_or_build<F>(
        &self,
        raw: &Arc<Vec<u8>>,
        route: ParagraphIndexRoute,
        build: F,
    ) -> Result<ParagraphIndexBuild>
    where
        F: FnOnce() -> Result<ParagraphIndexBuild>,
    {
        self.get_or_build_with_key(ParagraphIndexCacheKey::allocation(raw, route), build)
    }

    pub(crate) fn get_or_build_source<F>(
        &self,
        version: SourceVersion,
        route: ParagraphIndexRoute,
        build: F,
    ) -> Result<ParagraphIndexBuild>
    where
        F: FnOnce() -> Result<ParagraphIndexBuild>,
    {
        self.get_or_build_with_key(ParagraphIndexCacheKey::source(version, route), build)
    }

    #[cfg(test)]
    pub(crate) fn diagnostics(&self) -> ParagraphIndexCacheDiagnostics {
        let state = self.lock_state();
        let (clean_entries, pinned_entries, retained_weight) =
            state.entry.as_ref().map_or((0, 0, 0), |entry| {
                if Arc::strong_count(&entry.memo) == 1 {
                    (1, 0, entry.memo.index.weight_bytes())
                } else {
                    (0, 1, entry.memo.index.weight_bytes())
                }
            });
        ParagraphIndexCacheDiagnostics {
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

#[cfg(test)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct ParagraphIndexCacheDiagnostics {
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

/// Shared execution-budget reservation for a managed paragraph index.
pub(crate) struct DocumentIndexAdmission {
    _memory: Reservation,
    _objects: Reservation,
}

impl DocumentIndexAdmission {
    pub(crate) fn new(context: &ExecutionContext, xml_len: usize) -> Result<Arc<Self>> {
        let ranges = xml_len as u64 / 4 + 1;
        let bytes = ranges.saturating_mul(24).saturating_add(1024);
        let memory = context
            .reserve(Resource::Memory, bytes)
            .map_err(index_execution_error)?;
        let objects = context
            .reserve(Resource::Objects, ranges + 3)
            .map_err(index_execution_error)?;
        Ok(Arc::new(Self {
            _memory: memory,
            _objects: objects,
        }))
    }
}

fn index_execution_error(error: ExecutionError) -> crate::Error {
    match error {
        ExecutionError::Cancelled => crate::Error::Opc(litchi_opc::OpcError::Cancelled),
        error => crate::Error::Opc(litchi_opc::OpcError::Execution(error)),
    }
}

const MAX_DOCUMENT_SEMANTIC_VALUES: usize = 1_000_000;

fn reserve_document_value<T>(values: &mut Vec<T>, resource: &'static str) -> Result<()> {
    if values.len() >= MAX_DOCUMENT_SEMANTIC_VALUES {
        return Err(crate::Error::InvalidFormat(format!(
            "document semantic value count exceeds {MAX_DOCUMENT_SEMANTIC_VALUES}"
        )));
    }
    values
        .try_reserve(1)
        .map_err(|source| crate::Error::Allocation { resource, source })
}

fn push_document_smallvec<T, const N: usize>(
    inline: &mut SmallVec<[T; N]>,
    spill: &mut Option<Vec<T>>,
    value: T,
    resource: &'static str,
) -> Result<()> {
    if let Some(values) = spill {
        reserve_document_value(values, resource)?;
        values.push(value);
        return Ok(());
    }
    if inline.len() < N {
        inline.push(value);
        return Ok(());
    }
    if inline.len() >= MAX_DOCUMENT_SEMANTIC_VALUES {
        return Err(crate::Error::InvalidFormat(format!(
            "document semantic value count exceeds {MAX_DOCUMENT_SEMANTIC_VALUES}"
        )));
    }
    let capacity = N.checked_add(1).ok_or_else(|| {
        crate::Error::InvalidFormat("document semantic value capacity overflow".into())
    })?;
    let mut values = Vec::new();
    values
        .try_reserve_exact(capacity)
        .map_err(|source| crate::Error::Allocation { resource, source })?;
    values.extend(inline.drain(..));
    values.push(value);
    *spill = Some(values);
    Ok(())
}

pub(crate) fn is_xml_outer_whitespace(bytes: &[u8]) -> bool {
    bytes
        .iter()
        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
}

/// The main document part of a Word document.
///
/// This corresponds to the `/word/document.xml` part in the package.
/// It contains the main document content including paragraphs, tables,
/// sections, and other block-level elements.
pub struct DocumentPart<'a> {
    /// Borrowed or owned physical part snapshot.
    part: DocumentPartOwner<'a>,
    xml: Arc<Vec<u8>>,
    /// The source allocation identity used to find a package-scoped memo.
    /// This weak key never keeps the physical payload alive.
    paragraph_source: Weak<Vec<u8>>,
    paragraph_cache: Arc<ParagraphIndexCache>,
    paragraph_route: ParagraphIndexRoute,
    /// Best-effort cache built from the validated visible XML on the first
    /// paragraph query. A malformed or over-bound payload leaves the cached
    /// value empty so the legacy query path keeps reporting the same error at
    /// query time.
    paragraph_memo: OnceLock<Option<Arc<ParagraphIndexMemo>>>,
    /// Whether the visible XML keeps every element where the raw part has it.
    positions: RawPositions,
    /// The body block index, pinned by this view after its first body query.
    body_memo: OnceLock<Arc<crate::document::BlockIndex>>,
}

/// The markup-compatibility profile a main document is read with: the OOXML
/// baseline, plus Word 2010 (`w14`), Word 2012 (`w15`) and Word 2023 UTC-date
/// (`w16du`) namespaces.
///
/// Word 2010 paragraph and drawing extensions are declared ignorable by
/// producers. Keeping those namespaces understood at the package boundary
/// lets `w14:paraId`, `w14:textId`, `w14:noSpellErr` and unrelated `w14`
/// markup survive the visibility pass for typed and opaque readers. Word
/// 2012 supplies the typed collapsed-heading and footnote-column properties.
pub(crate) fn visible_document_capabilities() -> litchi_ooxml_common::mce::Capabilities {
    let mut capabilities = litchi_ooxml_common::mce::Capabilities::default();
    capabilities.understand_namespace(crate::paragraph::extensions::WORD_2010_NAMESPACE);
    capabilities.understand_namespace(crate::revision::WORD_2023_DATE_UTC_NAMESPACE);
    capabilities.understand_namespace("http://schemas.microsoft.com/office/word/2012/wordml");
    // The text and drawing inventories understand these owners. Selecting
    // their VML fallbacks would hide the supported DrawingML anchor metadata.
    // The flat drawing inventory retains group roots without flattening children.
    capabilities
        .understand_namespace("http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
    capabilities
        .understand_namespace("http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
    capabilities
}

/// Index the direct block children of the `w:body` of a visible main
/// document.
pub(crate) fn body_block_index(
    visible: &[u8],
    positions: RawPositions,
) -> Result<crate::document::BlockIndex> {
    crate::document::index_blocks(
        visible,
        0,
        &[],
        crate::document::ContainerRoot::Document,
        positions,
    )
}

/// Re-derive a body index a memo returned and compare it, in test and debug
/// builds, so the suites prove that a reused index equals a fresh one.
#[cfg_attr(
    not(debug_assertions),
    expect(unused_variables, reason = "debug-only check")
)]
fn debug_assert_reused_body_index(
    index: &crate::document::BlockIndex,
    visible: &[u8],
    positions: RawPositions,
) {
    #[cfg(debug_assertions)]
    debug_assert!(
        body_block_index(visible, positions).is_ok_and(|fresh| fresh.same_blocks(index)),
        "a memoized body block index differs from a fresh scan"
    );
}

/// The element a main document's opaque alternate content lies in: a
/// WordprocessingML run, Transitional or Strict (record 0307).
const RUN_HOLDER: litchi_ooxml_common::mce::admission::Holder<'static> =
    litchi_ooxml_common::mce::admission::Holder {
        namespaces: &[
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "http://purl.oclc.org/ooxml/wordprocessingml/main",
        ],
        local: "r",
    };

/// Whether every `mc:AlternateContent` of a main-document payload lies
/// inside a run, and everything outside those wrappers agrees with the
/// visible projection as the shared admission proof states, with at least
/// one wrapper (record 0307).
///
/// The projection replaces each wrapper, inside its run, with the content
/// of the branch it selects, so every block and every run keeps its parent
/// and its order: raw and visible positions agree. The proof does not cover
/// processing inside the wrappers; a caller runs the projection, which
/// fails on a malformed wrapper.
pub(crate) fn holds_run_alternate_content(raw: &[u8]) -> bool {
    litchi_ooxml_common::mce::admission::admit_held(
        raw,
        &visible_document_capabilities(),
        &litchi_ooxml_common::mce::Limits::default(),
        RUN_HOLDER,
    )
    .is_ok_and(|held| held.wrappers > 0)
}

/// How [`admit_document_source`] admitted a main-document part.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SourceAdmission {
    /// The shared admission proof holds for the whole part.
    Proven(litchi_ooxml_common::mce::admission::Admission),
    /// Every `mc:AlternateContent` lies inside a run
    /// ([`holds_run_alternate_content`]) and the projection succeeds. The
    /// transaction scanners treat each wrapper as one opaque inline object.
    RunAlternateContent,
}

impl SourceAdmission {
    /// Whether the projection differs from the part only in spelling and in
    /// the document element's `Ignorable` attribute.
    pub(crate) const fn is_lexical(self) -> bool {
        match self {
            Self::Proven(admission) => admission.is_lexical(),
            Self::RunAlternateContent => false,
        }
    }
}

/// Admit a main-document part to a transaction that reads and publishes its
/// source bytes while readers see the visible projection.
///
/// The transaction may edit the source only where the two agree on
/// everything either of them reads. The shared admission proof
/// ([`litchi_ooxml_common::mce::admission`]) establishes that in one pass
/// without building the projection: a part that never mentions the
/// markup-compatibility namespace, one that binds it and uses nothing in it,
/// and one whose only constructs are valid `Ignorable` directives: one on
/// `w:document`, whose dropped markup holds no element or attribute of an
/// understood namespace, and any below it listing only understood namespaces
/// (as litchi's revision writer emits). The transaction scanners treat that
/// foreign markup as opaque, as the projection drops it, so positions, text
/// and every WordprocessingML element and attribute agree.
///
/// A part the proof cannot admit runs the projection, whose error is
/// returned as is (for example an unsatisfiable `MustUnderstand`). A
/// projection equal to the source is admitted. One that differs is admitted
/// when its only selections are `mc:AlternateContent` wrappers inside runs
/// (record 0307: drawings and text boxes with a fallback), which the
/// transaction treats as opaque inline objects: their branches' text is
/// never document text to an edit, and an edit that would have to replace
/// it is refused. Everything else is refused as an unsafe edit with
/// `reason`. Decisions are made on decoded names, so a namespace declaration
/// that spells the namespace with character references is a binding like
/// any other.
///
/// # Errors
///
/// Returns the projection's markup-compatibility error, or
/// [`crate::Error::Refused`] for a part whose projection selects, unwraps
/// or hides content outside run-level alternate content.
pub(crate) fn admit_document_source(
    raw: &[u8],
    operation: &'static str,
    reason: &'static str,
) -> Result<SourceAdmission> {
    use litchi_ooxml_common::mce;

    if !mce::mentions_namespace(raw) {
        return Ok(SourceAdmission::Proven(mce::admission::Admission::Unbound));
    }
    let capabilities = visible_document_capabilities();
    let limits = mce::Limits::default();
    match mce::admission::admit(raw, &capabilities, &limits) {
        Ok(admission) => Ok(SourceAdmission::Proven(admission)),
        Err(_) => match mce::process_markup_compatibility(raw, &capabilities, &limits)?.xml {
            std::borrow::Cow::Borrowed(_) => {
                Ok(SourceAdmission::Proven(mce::admission::Admission::Unbound))
            },
            std::borrow::Cow::Owned(_) if holds_run_alternate_content(raw) => {
                Ok(SourceAdmission::RunAlternateContent)
            },
            std::borrow::Cow::Owned(_) => Err(crate::Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                operation,
                reason,
            )),
        },
    }
}

/// Extract all visible paragraphs from a normalized main-document payload.
pub(crate) fn document_paragraphs(xml: Arc<Vec<u8>>) -> Result<SmallVec<[Paragraph; 32]>> {
    let mut inline = SmallVec::new();
    let mut spill = None;
    scan_word_element_ranges_with_context(
        xml.as_slice(),
        &[],
        &[b"p".as_slice()],
        |_, start, length, namespaces| {
            push_document_smallvec(
                &mut inline,
                &mut spill,
                Paragraph::from_arc_range_with_context(Arc::clone(&xml), start, length, namespaces),
                "document paragraph views",
            )
        },
    )?;
    match spill {
        Some(values) => Ok(SmallVec::from_vec(values)),
        None => Ok(inline),
    }
}

/// Materialize paragraph views from an already validated range index.
pub(crate) fn document_paragraphs_from_index(
    xml: Arc<Vec<u8>>,
    index: &ParagraphIndex,
) -> Result<SmallVec<[Paragraph; 32]>> {
    let mut inline = SmallVec::new();
    let mut spill = None;
    for range in index.iter() {
        push_document_smallvec(
            &mut inline,
            &mut spill,
            Paragraph::from_arc_range_with_context(
                Arc::clone(&xml),
                range.start,
                range.length,
                Arc::clone(&range.namespaces),
            ),
            "document paragraph views",
        )?;
    }
    match spill {
        Some(values) => Ok(SmallVec::from_vec(values)),
        None => Ok(inline),
    }
}

/// Select one paragraph from an already validated range index.
pub(crate) fn document_paragraph_from_index(
    xml: Arc<Vec<u8>>,
    index: &ParagraphIndex,
    position: usize,
) -> Option<Paragraph> {
    index.get(position).map(|range| {
        Paragraph::from_arc_range_with_context(xml, range.start, range.length, range.namespaces)
    })
}

/// Select one visible paragraph without materializing every paragraph view.
///
/// The scanner still consumes the complete payload so malformed trailing XML
/// and the shared depth/node limits retain their established error timing.
pub(crate) fn document_paragraph(xml: Arc<Vec<u8>>, index: usize) -> Result<Option<Paragraph>> {
    let mut position = 0usize;
    let mut paragraph = None;
    scan_word_element_ranges_with_context(
        xml.as_slice(),
        &[],
        &[b"p".as_slice()],
        |_, start, length, namespaces| {
            if position == index {
                paragraph = Some(Paragraph::from_arc_range_with_context(
                    Arc::clone(&xml),
                    start,
                    length,
                    namespaces,
                ));
            }
            position = position.checked_add(1).ok_or_else(|| {
                crate::Error::InvalidFormat("document paragraph counter overflow".into())
            })?;
            Ok(())
        },
    )?;
    Ok(paragraph)
}

/// Count visible paragraphs without allocating paragraph range objects.
pub(crate) fn document_paragraph_count(xml: &[u8]) -> Result<usize> {
    let mut count = 0usize;
    scan_word_element_ranges(xml, &[b"p".as_slice()], |_, _, _| {
        count = count.checked_add(1).ok_or_else(|| {
            crate::Error::InvalidFormat("document paragraph counter overflow".into())
        })?;
        Ok(())
    })?;
    Ok(count)
}

/// Select active direct children of the main document body in source order.
///
/// `DocumentPart::from_part` has already selected MCE branches for this
/// source, so every returned range addresses the visible XML and unmodeled
/// body children can remain lossless instead of being silently discarded.
pub(crate) fn body_block_ranges(xml: &[u8]) -> Result<Vec<(usize, u32, u32, NamespaceBindings)>> {
    const PARAGRAPH: usize = 0;
    const TABLE: usize = 1;
    const ALT: usize = 2;
    const UNKNOWN: usize = 3;
    const CONTENT_CONTROL: usize = 4;
    const MAX_DEPTH: usize = 256;
    const MAX_NODES: usize = 1_000_000;

    let mut reader = NsReader::from_reader(xml);
    // Ranges are byte offsets into `xml`, whose leading byte-order mark
    // precedes reader position zero.
    let origin = ReaderOrigin::of(xml);
    let mut ranges = Vec::new();
    let mut namespace_capture = NamespaceCapture::default();
    let mut body_depth = None;
    let mut pending = None::<(usize, usize, NamespaceBindings)>;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut saw_root = false;
    let mut root_closed = false;

    loop {
        let start = origin.offset(reader.buffer_position()).ok_or_else(|| {
            crate::Error::InvalidFormat("document XML offset does not fit usize".into())
        })?;
        let event = reader
            .read_event()
            .map_err(|error| crate::Error::Xml(error.to_string()))?;
        let (namespace, event) = reader.resolver().resolve_event(event);
        let end = origin.offset(reader.buffer_position()).ok_or_else(|| {
            crate::Error::InvalidFormat("document XML offset does not fit usize".into())
        })?;

        if matches!(event, Event::Start(_) | Event::Empty(_)) {
            nodes = nodes.checked_add(1).ok_or_else(|| {
                crate::Error::InvalidFormat("document XML element counter overflow".into())
            })?;
            if nodes > MAX_NODES {
                return Err(crate::Error::InvalidFormat(format!(
                    "document XML exceeds {MAX_NODES} elements"
                )));
            }
        }

        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if saw_root || root_closed {
                        return Err(crate::Error::InvalidFormat(
                            "document XML has multiple roots".into(),
                        ));
                    }
                    saw_root = true;
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    crate::Error::InvalidFormat("document XML nesting is too deep".into())
                })?;
                if depth > MAX_DEPTH {
                    return Err(crate::Error::InvalidFormat(format!(
                        "document XML nesting exceeds the {MAX_DEPTH} depth limit"
                    )));
                }
                let is_word = is_wordprocessing_namespace(&namespace);
                let local = element.local_name();
                if body_depth.is_none() && is_word && local.as_ref() == b"body" {
                    body_depth = Some(depth);
                } else if body_depth.is_some_and(|body| depth == body + 1) {
                    let kind = if is_word && local.as_ref() == b"p" {
                        PARAGRAPH
                    } else if is_word && local.as_ref() == b"tbl" {
                        TABLE
                    } else if is_word && local.as_ref() == b"altChunk" {
                        ALT
                    } else if is_word && local.as_ref() == b"sdt" {
                        CONTENT_CONTROL
                    } else {
                        UNKNOWN
                    };
                    pending = Some((kind, start, namespace_capture.capture(reader.resolver())?));
                }
            },
            Event::Empty(element) => {
                if depth == 0 {
                    if saw_root || root_closed {
                        return Err(crate::Error::InvalidFormat(
                            "document XML has multiple roots".into(),
                        ));
                    }
                    saw_root = true;
                    root_closed = true;
                }
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    crate::Error::InvalidFormat("document XML nesting is too deep".into())
                })?;
                if child_depth > MAX_DEPTH {
                    return Err(crate::Error::InvalidFormat(format!(
                        "document XML nesting exceeds the {MAX_DEPTH} depth limit"
                    )));
                }
                if body_depth.is_some_and(|body| child_depth == body + 1) {
                    let local = element.local_name();
                    let kind = if is_wordprocessing_namespace(&namespace) && local.as_ref() == b"p"
                    {
                        PARAGRAPH
                    } else if is_wordprocessing_namespace(&namespace) && local.as_ref() == b"tbl" {
                        TABLE
                    } else if is_wordprocessing_namespace(&namespace)
                        && local.as_ref() == b"altChunk"
                    {
                        ALT
                    } else if is_wordprocessing_namespace(&namespace) && local.as_ref() == b"sdt" {
                        CONTENT_CONTROL
                    } else {
                        UNKNOWN
                    };
                    let range_start = u32::try_from(start).map_err(|_source_error| {
                        crate::Error::InvalidFormat("document XML offset does not fit u32".into())
                    })?;
                    let length = u32::try_from(end.checked_sub(start).ok_or_else(|| {
                        crate::Error::InvalidFormat("document XML range underflow".into())
                    })?)
                    .map_err(|_source_error| {
                        crate::Error::InvalidFormat("document XML range does not fit u32".into())
                    })?;
                    reserve_document_value(&mut ranges, "document body block ranges")?;
                    ranges.push((
                        kind,
                        range_start,
                        length,
                        namespace_capture.capture(reader.resolver())?,
                    ));
                }
            },
            Event::End(element) => {
                if body_depth.is_some_and(|body| depth == body + 1) {
                    let (kind, range_start, namespaces) = pending.take().ok_or_else(|| {
                        crate::Error::InvalidFormat("missing document body block".into())
                    })?;
                    let start = u32::try_from(range_start).map_err(|_source_error| {
                        crate::Error::InvalidFormat("document XML offset does not fit u32".into())
                    })?;
                    let length = u32::try_from(end.checked_sub(range_start).ok_or_else(|| {
                        crate::Error::InvalidFormat("document XML range underflow".into())
                    })?)
                    .map_err(|_source_error| {
                        crate::Error::InvalidFormat("document XML range does not fit u32".into())
                    })?;
                    reserve_document_value(&mut ranges, "document body block ranges")?;
                    ranges.push((kind, start, length, namespaces));
                }
                if body_depth == Some(depth)
                    && is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"body"
                {
                    body_depth = None;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    crate::Error::InvalidFormat("invalid document XML nesting".into())
                })?;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Eof if !saw_root || depth != 0 || pending.is_some() => {
                return Err(crate::Error::InvalidFormat(
                    "document XML does not contain exactly one root".into(),
                ));
            },
            Event::Eof => break,
            Event::Text(text) => {
                if depth == 0 && !is_xml_outer_whitespace(text.as_ref()) {
                    return Err(crate::Error::InvalidFormat(
                        "document XML has character data outside its root".into(),
                    ));
                }
            },
            Event::CData(_) if depth == 0 => {
                return Err(crate::Error::InvalidFormat(
                    "document XML has CDATA outside its root".into(),
                ));
            },
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Ok(ranges)
}

/// Extract all visible table views from a normalized main-document payload.
pub(crate) fn document_tables(xml: Arc<Vec<u8>>) -> Result<SmallVec<[Table; 8]>> {
    let mut inline = SmallVec::new();
    let mut spill = None;
    scan_word_element_ranges_with_context(
        xml.as_slice(),
        &[],
        &[b"tbl".as_slice()],
        |_, start, length, namespaces| {
            push_document_smallvec(
                &mut inline,
                &mut spill,
                Table::from_arc_range_with_context(Arc::clone(&xml), start, length, namespaces),
                "document table views",
            )
        },
    )?;
    match spill {
        Some(values) => Ok(SmallVec::from_vec(values)),
        None => Ok(inline),
    }
}

/// Extract all visible blocks in direct body order from a normalized payload.
///
/// Every active direct body child is represented. Typed paragraphs, tables,
/// block content controls and `altChunk` anchors retain their semantic models
/// while unmodeled children remain inert [`crate::OpaqueBlock`] values. The
/// helper is shared by eager and source-backed documents so MCE-visible
/// ordering and unknown-block retention cannot drift between the two paths.
/// `positions` is the projection's verdict, carried by content controls.
pub(crate) fn document_blocks(
    xml: Arc<Vec<u8>>,
    positions: RawPositions,
) -> Result<Vec<crate::Block>> {
    use crate::Block;

    let mut alts = scan(xml.as_slice())?;
    let mut elements = Vec::new();
    for (target, start, length, namespaces) in body_block_ranges(xml.as_slice())? {
        reserve_document_value(&mut elements, "document block views")?;
        let block_source = Arc::clone(&xml);
        elements.push(if target == 0 {
            Block::Paragraph(Box::new(Paragraph::from_arc_range_with_context(
                block_source,
                start,
                length,
                namespaces,
            )))
        } else if target == 1 {
            Block::Table(Box::new(Table::from_arc_range_with_context(
                block_source,
                start,
                length,
                namespaces,
            )))
        } else if target == 2 {
            let chunk = alts.remove(&start).ok_or_else(|| {
                crate::error::Error::InvalidFormat(
                    "ordered altChunk lacks parsed anchor metadata".into(),
                )
            })?;
            Block::Alt(Box::new(chunk))
        } else if target == 4 {
            Block::Sdt(Box::new(crate::document::BlockContentControl::new(
                block_source,
                start,
                length,
                namespaces,
                positions,
            )))
        } else {
            Block::Unknown(Box::new(crate::OpaqueBlock::from_arc_range(
                block_source,
                start,
                length,
            )))
        });
    }
    Ok(elements)
}

/// Extract visible paragraph/table/unknown elements in direct body order.
/// Alternative-format anchors remain available through [`document_blocks`]
/// but are intentionally omitted from this historical element view.
pub(crate) fn document_elements(xml: Arc<Vec<u8>>) -> Result<Vec<crate::Element>> {
    use crate::Element;

    let blocks = document_blocks(xml, RawPositions::Unproven)?;
    let mut elements = Vec::new();
    for block in blocks {
        let element = match block {
            crate::Block::Paragraph(paragraph) => Some(Element::Paragraph(paragraph)),
            crate::Block::Table(table) => Some(Element::Table(table)),
            crate::Block::Alt(_) => None,
            // The historical element view keeps a content control as the
            // inert exact XML it always returned.
            crate::Block::Sdt(control) => {
                Some(Element::Unknown(Box::new(control.to_opaque_block())))
            },
            crate::Block::Unknown(value) => Some(Element::Unknown(value)),
        };
        if let Some(element) = element {
            reserve_document_value(&mut elements, "document element views")?;
            elements.push(element);
        }
    }
    Ok(elements)
}

enum DocumentPartOwner<'a> {
    Borrowed(&'a dyn Part),
    Owned(Box<dyn Part + Send + Sync>),
}

impl DocumentPart<'static> {
    /// Retain a part snapshot without borrowing the package facade.
    pub(crate) fn from_owned_part_with_cache(
        part: Box<dyn Part + Send + Sync>,
        paragraph_cache: Arc<ParagraphIndexCache>,
    ) -> Result<Self> {
        Self::from_owner_with_cache(DocumentPartOwner::Owned(part), paragraph_cache)
    }
}

impl<'a> DocumentPart<'a> {
    /// Return the original OPC part; semantic reads use the cached MCE view.
    #[must_use]
    pub fn part(&self) -> &dyn Part {
        match &self.part {
            DocumentPartOwner::Borrowed(part) => *part,
            DocumentPartOwner::Owned(part) => part.as_ref(),
        }
    }

    /// Create a `DocumentPart` from a Part.
    ///
    /// # Arguments
    ///
    /// * `part` - The part containing the document.xml content
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        Self::from_part_with_cache(part, Arc::new(ParagraphIndexCache::new()))
    }

    /// Create a document part sharing a package's document memos.
    ///
    /// The markup-compatibility projection of the part's payload is memoized
    /// per payload allocation, so repeated views of an unchanged part share
    /// one projection instead of rewriting the part each time.
    pub(crate) fn from_part_with_cache(
        part: &'a dyn Part,
        paragraph_cache: Arc<ParagraphIndexCache>,
    ) -> Result<Self> {
        Self::from_owner_with_cache(DocumentPartOwner::Borrowed(part), paragraph_cache)
    }

    fn from_owner_with_cache(
        part: DocumentPartOwner<'a>,
        paragraph_cache: Arc<ParagraphIndexCache>,
    ) -> Result<Self> {
        let raw = match &part {
            DocumentPartOwner::Borrowed(part) => part.blob_arc(),
            DocumentPartOwner::Owned(part) => part.blob_arc(),
        };
        let paragraph_source = Arc::downgrade(&raw);
        let visible = paragraph_cache.visible_payload(&raw)?;
        let positions = visible.positions();
        let xml = visible.xml(raw);
        Ok(Self {
            part,
            xml,
            paragraph_source,
            paragraph_cache,
            paragraph_route: ParagraphIndexRoute::Visible,
            paragraph_memo: OnceLock::new(),
            positions,
            body_memo: OnceLock::new(),
        })
    }

    /// The direct block children of the main document's `w:body`, indexed
    /// once per payload generation and shared by every view of it.
    ///
    /// # Errors
    ///
    /// Returns an error when the visible XML is malformed or a bound on the
    /// index is reached; nothing is retained then.
    pub fn body_blocks(&self) -> Result<crate::document::Blocks> {
        if let Some(index) = self.body_memo.get() {
            return Ok(crate::document::Blocks::new(
                Arc::clone(&self.xml),
                Arc::clone(index),
            ));
        }
        let index = match self.paragraph_source.upgrade() {
            Some(raw) => {
                self.paragraph_cache
                    .body_index(&raw, self.xml.as_slice(), self.positions)?
            },
            None => Arc::new(body_block_index(self.xml.as_slice(), self.positions)?),
        };
        let index = self.body_memo.get_or_init(|| index);
        Ok(crate::document::Blocks::new(
            Arc::clone(&self.xml),
            Arc::clone(index),
        ))
    }

    /// Borrow the bounded paragraph offsets, scanning the visible XML once on
    /// the first paragraph query.
    ///
    /// Index construction is deliberately best-effort and deliberately
    /// deferred. `DocumentPart` has always reported structural XML errors from
    /// the paragraph selectors rather than from construction, so a swallowed
    /// index failure keeps the legacy scanner - and its error - on exactly the
    /// same operations as before. Text, table, block and element reads never
    /// consult the index, so they no longer pay for a scan they cannot use.
    fn paragraph_ranges(&self) -> Option<&ParagraphIndex> {
        self.paragraph_memo
            .get_or_init(|| {
                let Some(raw) = self.paragraph_source.upgrade() else {
                    return ParagraphIndex::from_xml(self.xml.as_slice())
                        .ok()
                        .map(|index| {
                            Arc::new(ParagraphIndexMemo {
                                index,
                                _admission: None,
                            })
                        });
                };
                self.paragraph_cache
                    .get_or_build(&raw, self.paragraph_route, || {
                        Ok(ParagraphIndexBuild::indexed(
                            ParagraphIndex::from_xml(self.xml.as_slice())?,
                            None,
                        ))
                    })
                    .ok()
                    .and_then(|build| build.memo)
            })
            .as_ref()
            .map(|memo| &memo.index)
    }

    /// Get the shared Arc of XML bytes (zero-copy from Part).
    #[inline]
    fn get_xml_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.xml)
    }

    /// Get the XML bytes of the document.
    #[inline]
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        self.xml.as_slice()
    }

    /// Extract all paragraph text from the document.
    ///
    /// This performs a quick extraction of all text content by finding
    /// `<w:t>` elements in the XML.
    ///
    /// # Performance
    ///
    /// Uses `quick-xml` for efficient streaming XML parsing with pre-allocated
    /// buffers and validated text decoding.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn extract_text(&self) -> Result<String> {
        extract_word_text(self.xml_bytes())
    }

    /// Extract all paragraph text, with `separator` between paragraphs.
    ///
    /// The text is [`Self::extract_text`]'s, with `separator` before every
    /// paragraph after the first; for well-formed XML it is the paragraph
    /// texts joined by `separator`. Malformed XML fails as
    /// [`Self::extract_text`] fails.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::extract_text`].
    pub fn extract_text_with_paragraph_separator(&self, separator: &str) -> Result<String> {
        crate::paragraph::extract_word_paragraph_text(self.xml_bytes(), separator)
    }

    /// Read recognized tracked revisions throughout the visible main document.
    /// Includes body-final section properties and table-grid/row metadata that
    /// paragraph-only traversal cannot visit. Unknown markup remains in source.
    ///
    /// # Errors
    /// Returns malformed metadata, XML, or revision resource-limit errors.
    pub fn revisions(&self) -> Result<SmallVec<[crate::revision::Revision; 4]>> {
        self.revisions_bounded(crate::revision::Limits::default())
    }

    /// Read tracked revision metadata with explicit parser and retention limits.
    ///
    /// # Errors
    /// Returns malformed metadata, XML, or revision resource-limit errors.
    pub(crate) fn revisions_bounded(
        &self,
        limits: crate::revision::Limits,
    ) -> Result<SmallVec<[crate::revision::Revision; 4]>> {
        crate::revision::parse_revisions_with_limits(self.xml_bytes(), &[], limits)
    }

    /// Count the number of paragraphs in the document.
    ///
    /// Counts `<w:p>` elements in the document body.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn paragraph_count(&self) -> Result<usize> {
        if let Some(index) = self.paragraph_ranges() {
            return Ok(index.len());
        }
        document_paragraph_count(self.xml_bytes())
    }

    /// Count the number of tables in the document.
    ///
    /// Counts `<w:tbl>` elements in the document body.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn table_count(&self) -> Result<usize> {
        let mut count = 0;
        scan_word_element_ranges(self.xml_bytes(), &[b"tbl".as_slice()], |_, _, _| {
            count += 1;
            Ok(())
        })?;
        Ok(count)
    }

    /// Get all paragraphs in the document.
    ///
    /// Extracts all `<w:p>` elements from the document body.
    ///
    /// # Performance
    ///
    /// Uses namespace-aware streaming XML parsing and shared byte ranges.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn paragraphs(&self) -> Result<SmallVec<[Paragraph; 32]>> {
        if let Some(index) = self.paragraph_ranges() {
            return document_paragraphs_from_index(self.get_xml_arc(), index);
        }
        document_paragraphs(self.get_xml_arc())
    }

    /// Get one paragraph by zero-based index without allocating the complete
    /// paragraph collection.
    ///
    /// # Errors
    ///
    /// Returns an error if the document XML is malformed or exceeds bounds.
    pub fn paragraph(&self, index: usize) -> Result<Option<Paragraph>> {
        if let Some(paragraph_index) = self.paragraph_ranges() {
            return Ok(document_paragraph_from_index(
                self.get_xml_arc(),
                paragraph_index,
                index,
            ));
        }
        document_paragraph(self.get_xml_arc(), index)
    }

    /// Get all tables in the document.
    ///
    /// Extracts all `<w:tbl>` elements from the document body.
    ///
    /// # Performance
    ///
    /// Uses namespace-aware streaming XML parsing and shared byte ranges.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn tables(&self) -> Result<SmallVec<[Table; 8]>> {
        document_tables(self.get_xml_arc())
    }

    /// Get all document elements (paragraphs and tables) in document order.
    ///
    /// This method parses the XML once and extracts both paragraphs and tables,
    /// returning an ordered vector that preserves the document structure.
    /// This is more efficient than calling `paragraphs()` and `tables()` separately,
    /// and it maintains the correct order of elements for sequential processing.
    ///
    /// # Performance
    ///
    /// Uses a single-pass XML parser that extracts both `<w:p>` and `<w:tbl>` elements
    /// in document order, which is significantly faster than parsing the XML twice.
    ///
    /// # Performance
    ///
    /// Uses one-pass, namespace-aware zero-copy parsing.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn elements(&self) -> Result<Vec<crate::Element>> {
        document_elements(self.get_xml_arc())
    }

    /// Get paragraphs, tables, and alternative-format anchors in document order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn blocks(&self) -> Result<Vec<crate::Block>> {
        document_blocks(self.get_xml_arc(), self.positions)
    }

    /// Return all alternative-format anchors in XML order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn alts(&self) -> Result<Vec<Chunk>> {
        let mut alts = Vec::new();
        for block in self.blocks()? {
            if let crate::Block::Alt(chunk) = block {
                reserve_document_value(&mut alts, "document alternative-format anchors")?;
                alts.push(*chunk);
            }
        }
        Ok(alts)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::Element;
    use litchi_opc::packuri::PackURI;
    use litchi_opc::part::BlobPart;

    fn document_part(xml: &[u8]) -> BlobPart {
        BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
                .to_string(),
            xml.to_vec(),
        )
    }

    #[test]
    fn extracts_aliased_word_elements_in_document_order_without_copying_text() {
        let xml = br#"<wp:document xmlns:wp="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:false="urn:not-wordprocessingml"><wp:body><false:p><false:r><false:t>ignored</false:t></false:r></false:p><wp:p><wp:r><wp:t><![CDATA[A < B]]></wp:t></wp:r></wp:p><wp:tbl><wp:tr><wp:tc><wp:p><wp:r><wp:t>cell</wp:t></wp:r></wp:p></wp:tc></wp:tr></wp:tbl><wp:p><wp:r><wp:t>tail</wp:t></wp:r></wp:p><wp:p/><false:tbl/></wp:body></wp:document>"#;
        let part = document_part(xml);
        let document = DocumentPart::from_part(&part).unwrap();

        assert_eq!(document.paragraph_count().unwrap(), 4);
        assert_eq!(document.table_count().unwrap(), 1);
        assert_eq!(document.tables().unwrap().len(), 1);

        let paragraphs = document.paragraphs().unwrap();
        assert_eq!(paragraphs.len(), 4);
        assert_eq!(paragraphs[0].text().unwrap(), "A < B");
        assert_eq!(paragraphs[0].runs().unwrap()[0].text().unwrap(), "A < B");
        assert_eq!(paragraphs[1].text().unwrap(), "cell");
        assert_eq!(paragraphs[2].text().unwrap(), "tail");
        assert_eq!(paragraphs[3].text().unwrap(), "");
        assert_eq!(
            document.paragraph(0).unwrap().unwrap().text().unwrap(),
            "A < B"
        );
        assert_eq!(
            document.paragraph(2).unwrap().unwrap().text().unwrap(),
            "tail"
        );
        assert!(document.paragraph(4).unwrap().is_none());

        let elements = document.elements().unwrap();
        assert_eq!(elements.len(), 6);
        assert!(matches!(elements[0], Element::Unknown(_)));
        assert!(matches!(elements[1], Element::Paragraph(_)));
        assert!(matches!(elements[2], Element::Table(_)));
        assert!(matches!(elements[3], Element::Paragraph(_)));
        assert!(matches!(elements[4], Element::Paragraph(_)));
        assert!(matches!(elements[5], Element::Unknown(_)));

        let blocks = document.blocks().unwrap();
        assert!(
            matches!(&blocks[0], crate::Block::Unknown(value) if value.xml_bytes() == br#"<false:p><false:r><false:t>ignored</false:t></false:r></false:p>"#)
        );
    }

    #[test]
    fn accepts_strict_wordprocessingml_and_self_closing_blocks() {
        let xml = br#"<s:document xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:body><s:p/><s:tbl/></s:body></s:document>"#;
        let part = document_part(xml);
        let document = DocumentPart::from_part(&part).unwrap();

        assert_eq!(document.paragraph_count().unwrap(), 1);
        assert_eq!(document.table_count().unwrap(), 1);
        assert_eq!(document.elements().unwrap().len(), 2);
    }

    #[test]
    fn rejects_unterminated_selected_elements() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r/>"#;
        let part = document_part(xml);
        let document = DocumentPart::from_part(&part).unwrap();

        assert!(document.paragraphs().is_err());
        assert!(document.paragraph(0).is_err());
        assert!(document.elements().is_err());
    }

    #[test]
    fn text_and_block_reads_never_build_the_paragraph_index() {
        let xml = br#"<wp:document xmlns:wp="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><wp:body><wp:p><wp:r><wp:t>outer</wp:t></wp:r></wp:p><wp:tbl><wp:tr><wp:tc><wp:p><wp:r><wp:t>cell</wp:t></wp:r></wp:p></wp:tc></wp:tr></wp:tbl><wp:p><wp:r><wp:t>tail</wp:t></wp:r></wp:p></wp:body></wp:document>"#;
        let part = document_part(xml);
        let document = DocumentPart::from_part(&part).unwrap();

        assert!(
            document.paragraph_memo.get().is_none(),
            "construction must not scan for paragraph ranges"
        );
        assert_eq!(document.extract_text().unwrap(), "outercelltail");
        assert_eq!(document.table_count().unwrap(), 1);
        assert_eq!(document.tables().unwrap().len(), 1);
        assert_eq!(document.elements().unwrap().len(), 3);
        assert_eq!(document.blocks().unwrap().len(), 3);
        assert_eq!(document.xml_bytes(), xml);
        assert!(
            document.paragraph_memo.get().is_none(),
            "text, table, element and block reads must not build the index"
        );

        assert_eq!(document.paragraph_count().unwrap(), 3);
        assert!(
            document.paragraph_memo.get().is_some(),
            "the first paragraph query builds the index once"
        );
    }

    #[test]
    fn deferred_index_answers_every_paragraph_query_like_the_streaming_scan() {
        let xml = br#"<wp:document xmlns:wp="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><wp:body><wp:p><wp:r><wp:t>outer</wp:t></wp:r></wp:p><wp:tbl><wp:tr><wp:tc><wp:p><wp:r><wp:t>cell</wp:t></wp:r></wp:p></wp:tc></wp:tr></wp:tbl><wp:p><wp:r><wp:t>tail</wp:t></wp:r></wp:p><wp:p/></wp:body></wp:document>"#;
        let source = Arc::new(xml.to_vec());
        let part = document_part(xml);
        let document = DocumentPart::from_part(&part).unwrap();

        let expected_count = document_paragraph_count(xml).unwrap();
        let expected: Vec<String> = document_paragraphs(Arc::clone(&source))
            .unwrap()
            .iter()
            .map(|paragraph| paragraph.text().unwrap())
            .collect();

        assert_eq!(document.paragraph_count().unwrap(), expected_count);
        let observed: Vec<String> = document
            .paragraphs()
            .unwrap()
            .iter()
            .map(|paragraph| paragraph.text().unwrap())
            .collect();
        assert_eq!(observed, expected);
        for position in 0..=expected_count {
            let streamed = document_paragraph(Arc::clone(&source), position)
                .unwrap()
                .map(|paragraph| paragraph.text().unwrap());
            let selected = document
                .paragraph(position)
                .unwrap()
                .map(|paragraph| paragraph.text().unwrap());
            assert_eq!(selected, streamed, "paragraph {position} differs");
        }
        // The second round runs entirely off the cached ranges.
        assert_eq!(document.paragraph_count().unwrap(), expected_count);
        assert_eq!(
            document.paragraph(0).unwrap().unwrap().text().unwrap(),
            expected[0]
        );
    }

    #[test]
    fn concurrent_first_paragraph_queries_agree_on_one_index() {
        const fn assert_shared<T: Send + Sync>() {}
        assert_shared::<DocumentPart<'_>>();

        let xml = br#"<wp:document xmlns:wp="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><wp:body><wp:p><wp:r><wp:t>outer</wp:t></wp:r></wp:p><wp:p><wp:r><wp:t>tail</wp:t></wp:r></wp:p></wp:body></wp:document>"#;
        let part = document_part(xml);
        let document = DocumentPart::from_part(&part).unwrap();
        let shared = &document;

        std::thread::scope(|scope| {
            let handles: Vec<_> = (0..8)
                .map(|_| {
                    scope.spawn(move || {
                        let count = shared.paragraph_count().unwrap();
                        let first = shared.paragraph(0).unwrap().unwrap().text().unwrap();
                        let all: Vec<String> = shared
                            .paragraphs()
                            .unwrap()
                            .iter()
                            .map(|paragraph| paragraph.text().unwrap())
                            .collect();
                        (count, first, all)
                    })
                })
                .collect();
            for handle in handles {
                let (count, first, all) = handle.join().unwrap();
                assert_eq!(count, 2);
                assert_eq!(first, "outer");
                assert_eq!(all, vec!["outer".to_owned(), "tail".to_owned()]);
            }
        });
        assert!(document.paragraph_memo.get().is_some());
    }

    #[test]
    fn paragraph_index_matches_independent_text_oracle() {
        let xml = br#"<wp:document xmlns:wp="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><wp:body><wp:p><wp:r><wp:t>outer</wp:t></wp:r></wp:p><wp:tbl><wp:tr><wp:tc><wp:p><wp:r><wp:t>cell</wp:t></wp:r></wp:p></wp:tc></wp:tr></wp:tbl><wp:p><wp:r><wp:t>tail</wp:t></wp:r></wp:p><wp:p/></wp:body></wp:document>"#;
        let index = ParagraphIndex::from_xml(xml).unwrap();
        let source = Arc::new(xml.to_vec());
        let expected = ["outer", "cell", "tail", ""];

        assert_eq!(index.len(), expected.len());
        for (position, expected_text) in expected.into_iter().enumerate() {
            let paragraph = document_paragraph_from_index(Arc::clone(&source), &index, position)
                .expect("indexed paragraph is present");
            assert_eq!(paragraph.text().unwrap(), expected_text);
        }
        assert!(document_paragraph_from_index(source, &index, expected.len()).is_none());
    }

    #[test]
    fn package_cache_reuses_allocation_identity_and_rejects_equal_new_allocations() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:p/></w:body></w:document>"#;
        let cache = Arc::new(ParagraphIndexCache::new());
        let raw = Arc::new(xml.to_vec());
        let first = cache
            .get_or_build(&raw, ParagraphIndexRoute::Visible, || {
                Ok(ParagraphIndexBuild::indexed(
                    ParagraphIndex::from_xml(raw.as_slice())?,
                    None,
                ))
            })
            .unwrap();
        let second = cache
            .get_or_build(&raw, ParagraphIndexRoute::Visible, || {
                Err(crate::Error::InvalidFormat(
                    "same allocation unexpectedly missed the package cache".into(),
                ))
            })
            .unwrap();
        assert!(Arc::ptr_eq(
            first.memo.as_ref().unwrap(),
            second.memo.as_ref().unwrap()
        ));

        let equal_new_allocation = Arc::new(xml.to_vec());
        let third = cache
            .get_or_build(&equal_new_allocation, ParagraphIndexRoute::Visible, || {
                Ok(ParagraphIndexBuild::indexed(
                    ParagraphIndex::from_xml(equal_new_allocation.as_slice())?,
                    None,
                ))
            })
            .unwrap();
        assert!(!Arc::ptr_eq(
            first.memo.as_ref().unwrap(),
            third.memo.as_ref().unwrap()
        ));
        let diagnostics = cache.diagnostics();
        assert_eq!(diagnostics.hits, 1);
        assert_eq!(diagnostics.misses, 2);
        assert_eq!(diagnostics.builds, 2);
    }

    #[test]
    fn package_cache_coalesces_builds_and_trims_only_clean_entries() {
        let cache = Arc::new(ParagraphIndexCache::new());
        let raw = Arc::new(b"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p/></w:body></w:document>".to_vec());
        let builds = Arc::new(std::sync::atomic::AtomicUsize::new(0));
        std::thread::scope(|scope| {
            let handles: Vec<_> = (0..8)
                .map(|_| {
                    let cache = Arc::clone(&cache);
                    let raw = Arc::clone(&raw);
                    let builds = Arc::clone(&builds);
                    scope.spawn(move || {
                        cache
                            .get_or_build(&raw, ParagraphIndexRoute::Visible, || {
                                builds.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
                                Ok(ParagraphIndexBuild::indexed(
                                    ParagraphIndex::from_xml(raw.as_slice())?,
                                    None,
                                ))
                            })
                            .unwrap()
                    })
                })
                .collect();
            for handle in handles {
                drop(handle.join().unwrap());
            }
        });
        assert_eq!(builds.load(std::sync::atomic::Ordering::Relaxed), 1);
        assert_eq!(cache.diagnostics().hits, 7);

        cache.trim_clean();
        assert_eq!(cache.diagnostics().clean_evictions, 1);
        assert_eq!(cache.diagnostics().clean_entries, 0);

        let active = cache
            .get_or_build(&raw, ParagraphIndexRoute::Visible, || {
                Ok(ParagraphIndexBuild::indexed(
                    ParagraphIndex::from_xml(raw.as_slice())?,
                    None,
                ))
            })
            .unwrap();
        cache.trim_clean();
        assert_eq!(cache.diagnostics().pinned_entries, 1);
        drop(active);
        cache.trim_clean();
        assert_eq!(cache.diagnostics().clean_evictions, 2);
    }

    #[test]
    fn package_cache_retries_failed_builds_and_bypasses_oversized_indexes() {
        let cache = Arc::new(ParagraphIndexCache::new());
        let raw = Arc::new(vec![0_u8; 1]);
        let attempts = std::sync::atomic::AtomicUsize::new(0);
        assert!(
            cache
                .get_or_build(&raw, ParagraphIndexRoute::Visible, || {
                    attempts.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
                    Err(crate::Error::InvalidFormat("retryable test failure".into()))
                })
                .is_err()
        );
        let retry = cache
            .get_or_build(&raw, ParagraphIndexRoute::Visible, || {
                attempts.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
                Ok(ParagraphIndexBuild::indexed(
                    ParagraphIndex {
                        ranges: Vec::<ParagraphRange>::new().into(),
                    },
                    None,
                ))
            })
            .unwrap();
        assert_eq!(retry.memo.unwrap().index.len(), 0);
        assert_eq!(attempts.load(std::sync::atomic::Ordering::Relaxed), 2);
        assert_eq!(cache.diagnostics().misses, 2);

        let oversized_raw = Arc::new(vec![1_u8; 1]);
        let oversized_ranges = vec![
            ParagraphRange {
                start: 0,
                length: 1,
                namespaces: Arc::from(Vec::new()),
            };
            1_100_000
        ];
        let oversized = cache
            .get_or_build(&oversized_raw, ParagraphIndexRoute::Visible, || {
                Ok(ParagraphIndexBuild::indexed(
                    ParagraphIndex {
                        ranges: oversized_ranges.into(),
                    },
                    None,
                ))
            })
            .unwrap();
        assert!(oversized.memo.is_some());
        let diagnostics = cache.diagnostics();
        assert_eq!(diagnostics.oversized_bypasses, 1);
        assert_eq!(diagnostics.clean_entries, 0);
    }
}
