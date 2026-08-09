#![expect(
    clippy::let_underscore_must_use,
    reason = "the builder return is intentionally ignored after mutation"
)]
#![expect(
    clippy::ptr_arg,
    reason = "the internal vector API preserves capacity-aware mutation"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::shadow_unrelated,
    reason = "local parser names mirror the OOXML role currently being decoded"
)]
//! Failure-atomic, source-bound conflict-markup edits.
//!
//! Removing an annotation only removes its markup. Inline and range payload is
//! retained byte-for-byte; this module never accepts, rejects, merges, resolves,
//! evaluates, fetches, or activates conflict content.

use std::{
    collections::HashMap,
    sync::{
        Arc,
        atomic::{AtomicU8, Ordering},
    },
};

use crate::{Error, Result};

use super::model::{Binding, Source};
use super::{AttributeSpan, Limits, Metadata, Scope, Snapshot, Span};

const GATE_READY: u8 = 0;
const GATE_IN_FLIGHT: u8 = 1;
const GATE_APPLIED: u8 = 2;

/// One source-coordinate edit queued by a [`Transaction`].
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Edit {
    /// Replace the author of a conflict element.
    SetConflictAuthor { index: usize, author: String },
    /// Replace or remove the optional date of a conflict element.
    SetConflictDate { index: usize, date: Option<String> },
    /// Remove a conflict annotation while retaining inline payload.
    RemoveConflict { index: usize },
    /// Replace the author on a paired range's start marker.
    SetRangeAuthor { index: usize, author: String },
    /// Replace or remove the optional date on a paired range's start marker.
    SetRangeDate { index: usize, date: Option<String> },
    /// Remove both range markers while retaining all enclosed payload.
    RemoveRange { index: usize },
}

/// An isolated edit over one exact conflict-markup source snapshot.
#[derive(Debug, Clone)]
pub struct Transaction {
    base: Snapshot,
    edits: Vec<Option<Edit>>,
    edit_index: HashMap<EditKey, usize>,
}

impl Transaction {
    pub(crate) fn new(base: Snapshot) -> Self {
        Self {
            base,
            edits: Vec::new(),
            edit_index: HashMap::new(),
        }
    }

    /// Borrow the immutable source snapshot.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.base
    }

    /// Iterate over the coalesced edits in request order.
    #[must_use]
    pub fn edits(&self) -> impl DoubleEndedIterator<Item = &Edit> {
        self.edits.iter().filter_map(Option::as_ref)
    }

    /// Whether this transaction has a semantic edit to publish.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.edit_index.is_empty()
    }

    /// Queue one typed edit after validating its target and metadata.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn apply(&mut self, edit: Edit) -> Result<&mut Self> {
        match edit {
            Edit::SetConflictAuthor { index, author } => self.set_conflict_author(index, author),
            Edit::SetConflictDate { index, date } => self.set_conflict_date(index, date),
            Edit::RemoveConflict { index } => self.remove_conflict(index),
            Edit::SetRangeAuthor { index, author } => self.set_range_author(index, author),
            Edit::SetRangeDate { index, date } => self.set_range_date(index, date),
            Edit::RemoveRange { index } => self.remove_range(index),
        }
    }

    /// Replace the author of one source-ordered conflict element.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_conflict_author(
        &mut self,
        index: usize,
        author: impl Into<String>,
    ) -> Result<&mut Self> {
        let author = author.into();
        let conflict = conflict_at(&self.base, index)?;
        refuse_removed_conflict(&self.edit_index, index)?;
        Metadata::new(
            conflict.metadata.id,
            author.clone(),
            conflict.metadata.date.clone(),
        )?;
        replace_field_edit(
            &mut self.edits,
            &mut self.edit_index,
            EditKey::ConflictAuthor(index),
            (author != conflict.metadata.author)
                .then_some(Edit::SetConflictAuthor { index, author }),
        )?;
        Ok(self)
    }

    /// Replace or remove the optional date of one conflict element.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_conflict_date(&mut self, index: usize, date: Option<String>) -> Result<&mut Self> {
        let conflict = conflict_at(&self.base, index)?;
        refuse_removed_conflict(&self.edit_index, index)?;
        Metadata::new(
            conflict.metadata.id,
            conflict.metadata.author.clone(),
            date.clone(),
        )?;
        replace_field_edit(
            &mut self.edits,
            &mut self.edit_index,
            EditKey::ConflictDate(index),
            (date != conflict.metadata.date).then_some(Edit::SetConflictDate { index, date }),
        )?;
        Ok(self)
    }

    /// Remove one annotation without interpreting its insertion/deletion kind.
    /// Inline child bytes are retained exactly; property markers are leaf-only.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn remove_conflict(&mut self, index: usize) -> Result<&mut Self> {
        conflict_at(&self.base, index)?;
        let removal = EditKey::RemoveConflict(index);
        if !self.edit_index.contains_key(&removal) {
            reserve_new_edit(&mut self.edits, &mut self.edit_index)?;
            remove_edit(
                &mut self.edits,
                &mut self.edit_index,
                EditKey::ConflictAuthor(index),
            );
            remove_edit(
                &mut self.edits,
                &mut self.edit_index,
                EditKey::ConflictDate(index),
            );
            insert_reserved_edit(
                &mut self.edits,
                &mut self.edit_index,
                removal,
                Edit::RemoveConflict { index },
            );
        }
        Ok(self)
    }

    /// Replace the author on one source-ordered paired range start marker.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_range_author(
        &mut self,
        index: usize,
        author: impl Into<String>,
    ) -> Result<&mut Self> {
        let author = author.into();
        let range = range_at(&self.base, index)?;
        refuse_removed_range(&self.edit_index, index)?;
        Metadata::new(
            range.metadata.id,
            author.clone(),
            range.metadata.date.clone(),
        )?;
        replace_field_edit(
            &mut self.edits,
            &mut self.edit_index,
            EditKey::RangeAuthor(index),
            (author != range.metadata.author).then_some(Edit::SetRangeAuthor { index, author }),
        )?;
        Ok(self)
    }

    /// Replace or remove the optional date on one paired range start marker.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_range_date(&mut self, index: usize, date: Option<String>) -> Result<&mut Self> {
        let range = range_at(&self.base, index)?;
        refuse_removed_range(&self.edit_index, index)?;
        Metadata::new(
            range.metadata.id,
            range.metadata.author.clone(),
            date.clone(),
        )?;
        replace_field_edit(
            &mut self.edits,
            &mut self.edit_index,
            EditKey::RangeDate(index),
            (date != range.metadata.date).then_some(Edit::SetRangeDate { index, date }),
        )?;
        Ok(self)
    }

    /// Remove both paired markers without changing any byte between them.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn remove_range(&mut self, index: usize) -> Result<&mut Self> {
        range_at(&self.base, index)?;
        let removal = EditKey::RemoveRange(index);
        if !self.edit_index.contains_key(&removal) {
            reserve_new_edit(&mut self.edits, &mut self.edit_index)?;
            remove_edit(
                &mut self.edits,
                &mut self.edit_index,
                EditKey::RangeAuthor(index),
            );
            remove_edit(
                &mut self.edits,
                &mut self.edit_index,
                EditKey::RangeDate(index),
            );
            insert_reserved_edit(
                &mut self.edits,
                &mut self.edit_index,
                removal,
                Edit::RemoveRange { index },
            );
        }
        Ok(self)
    }

    /// Validate and materialize the complete candidate without mutating the
    /// source. A failed commit leaves this transaction available for retry.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn commit(&self) -> Result<Commit> {
        let before = self.base.source_owner();
        let source_limits = self.base.limits();
        if self.edit_index.is_empty() {
            return Ok(Commit::new(
                self.base.clone(),
                before.clone(),
                before,
                source_limits,
                source_limits,
            ));
        }

        let mut splices = Vec::new();
        splices
            .try_reserve(self.edit_index.len().saturating_mul(2))
            .map_err(|source| Error::Allocation {
                resource: "conflict XML splices",
                source,
            })?;
        for edit in self.edits.iter().filter_map(Option::as_ref) {
            plan_edit(&self.base, edit, &mut splices)?;
        }
        let after = apply_splices(
            self.base.source(),
            &mut splices,
            source_limits.max_output_bytes,
        )?;
        let after = Source::Blob(Arc::new(after));
        let mut target_limits = source_limits;
        target_limits.max_source_bytes = source_limits.max_source_bytes.max(after.len());
        let target_limits = target_limits.validate()?;
        let mut snapshot = Snapshot::from_source_with_limits(after.clone(), target_limits)?;
        if let Some(binding) = self.base.binding().cloned() {
            snapshot = snapshot.with_binding(binding);
        }
        Ok(Commit::new(
            snapshot,
            before,
            after,
            source_limits,
            target_limits,
        ))
    }
}

/// A successful publication containing a parsed snapshot and reversible patch.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    fn new(
        snapshot: Snapshot,
        before: Source,
        after: Source,
        source_limits: Limits,
        target_limits: Limits,
    ) -> Self {
        let patch = Patch::new(
            before,
            after,
            source_limits,
            target_limits,
            snapshot.binding().cloned(),
        );
        Self { snapshot, patch }
    }

    /// Whether the publication changes any source byte.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Borrow the fully reparsed candidate snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible source-checked patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Move the candidate snapshot out of this commit.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Move the patch out of this commit.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }

    /// Split this commit into its candidate snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// An exact, reversible, source-checked XML replacement.
///
/// Clones share an apply-once gate. Stale or invalid attempts do not consume
/// that gate, so callers may correct the target and retry safely.
#[derive(Debug, Clone)]
pub struct Patch {
    source_fingerprint: u64,
    target_fingerprint: u64,
    before: Source,
    after: Source,
    source_limits: Limits,
    target_limits: Limits,
    binding: Option<Binding>,
    gate: Arc<AtomicU8>,
}

impl Patch {
    fn new(
        before: Source,
        after: Source,
        source_limits: Limits,
        target_limits: Limits,
        binding: Option<Binding>,
    ) -> Self {
        Self {
            source_fingerprint: fingerprint(before.as_ref()),
            target_fingerprint: fingerprint(after.as_ref()),
            before,
            after,
            source_limits,
            target_limits,
            binding,
            gate: Arc::new(AtomicU8::new(GATE_READY)),
        }
    }

    /// Return the expected exact-source fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Return the candidate fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Borrow the exact source bytes required by this patch.
    #[must_use]
    pub fn before_bytes(&self) -> &[u8] {
        self.before.as_ref()
    }

    /// Borrow the exact candidate bytes produced by this patch.
    #[must_use]
    pub fn after_bytes(&self) -> &[u8] {
        self.after.as_ref()
    }

    /// Clone the exact source owner without copying its bytes.
    pub(crate) fn before_source(&self) -> Source {
        self.before.clone()
    }

    /// Clone the exact candidate owner without copying its bytes.
    pub(crate) fn after_source(&self) -> Source {
        self.after.clone()
    }

    /// Return the parser and output limits captured by the transaction.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.target_limits
    }

    /// Return the parser limits required by the exact patch source.
    pub(crate) const fn source_limits(&self) -> Limits {
        self.source_limits
    }

    /// Return the parser limits governing the candidate publication.
    pub(crate) const fn target_limits(&self) -> Limits {
        self.target_limits
    }

    /// Whether the patch preserves every source byte.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.as_ref() == self.after.as_ref()
    }

    /// Whether this patch, or any clone of it, has applied successfully.
    #[must_use]
    pub fn is_applied(&self) -> bool {
        self.gate.load(Ordering::Acquire) == GATE_APPLIED
    }

    /// Apply this patch once to its exact source snapshot.
    ///
    /// Exact bytes, the compact fingerprint, parser limits, and optional story
    /// binding are all preconditions. The candidate is reparsed before the
    /// shared apply-once gate is consumed.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        let claim = self.claim_publication()?;
        if source.limits() != self.source_limits
            || source.binding() != self.binding.as_ref()
            || fingerprint(source.source()) != self.source_fingerprint
            || source.source() != self.before.as_ref()
        {
            return Err(Error::Invalid(
                "conflict patch source does not match its exact precondition".into(),
            ));
        }

        let mut candidate =
            Snapshot::from_source_with_limits(self.after.clone(), self.target_limits)?;
        if let Some(binding) = self.binding.clone() {
            candidate = candidate.with_binding(binding);
        }
        claim.finalize();
        Ok(candidate)
    }

    /// Build a fresh inverse patch. Applying the forward patch does not consume
    /// or pre-apply its inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self::new(
            self.after.clone(),
            self.before.clone(),
            self.target_limits,
            self.source_limits,
            self.binding.clone(),
        )
    }

    /// Claim this patch for one semantic or package publication attempt.
    /// Dropping the claim releases it for retry; finalizing marks every clone
    /// permanently applied, including exact no-op patches.
    pub(crate) fn claim_publication(&self) -> Result<PublicationClaim> {
        match self.gate.compare_exchange(
            GATE_READY,
            GATE_IN_FLIGHT,
            Ordering::AcqRel,
            Ordering::Acquire,
        ) {
            Ok(_) => Ok(PublicationClaim {
                gate: self.gate.clone(),
                finalized: false,
            }),
            Err(GATE_IN_FLIGHT) => Err(Error::Invalid(
                "conflict patch publication is already in flight".into(),
            )),
            Err(_) => Err(Error::Invalid(
                "conflict patch has already been applied".into(),
            )),
        }
    }

    pub(crate) const fn story_binding(&self) -> Option<&Binding> {
        self.binding.as_ref()
    }
}

/// Retry-safe ownership of one patch publication attempt.
pub(crate) struct PublicationClaim {
    gate: Arc<AtomicU8>,
    finalized: bool,
}

impl PublicationClaim {
    /// Permanently consume the patch after publication has succeeded.
    pub(crate) fn finalize(mut self) {
        self.gate.store(GATE_APPLIED, Ordering::Release);
        self.finalized = true;
    }
}

impl Drop for PublicationClaim {
    fn drop(&mut self) {
        if !self.finalized {
            let _ = self.gate.compare_exchange(
                GATE_IN_FLIGHT,
                GATE_READY,
                Ordering::AcqRel,
                Ordering::Acquire,
            );
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum EditKey {
    ConflictAuthor(usize),
    ConflictDate(usize),
    RemoveConflict(usize),
    RangeAuthor(usize),
    RangeDate(usize),
    RemoveRange(usize),
}

#[derive(Debug)]
struct Splice {
    span: Span,
    replacement: Vec<u8>,
}

fn plan_edit(snapshot: &Snapshot, edit: &Edit, splices: &mut Vec<Splice>) -> Result<()> {
    match edit {
        Edit::SetConflictAuthor { index, author } => {
            let conflict = conflict_at(snapshot, *index)?;
            replace_attribute(
                snapshot.source(),
                conflict.start_tag,
                conflict.author_span,
                conflict.author_span,
                "author",
                Some(author),
                splices,
            )
        },
        Edit::SetConflictDate { index, date } => {
            let conflict = conflict_at(snapshot, *index)?;
            replace_attribute(
                snapshot.source(),
                conflict.start_tag,
                conflict.date_span,
                conflict.author_span,
                "date",
                date.as_ref(),
                splices,
            )
        },
        Edit::RemoveConflict { index } => {
            let conflict = conflict_at(snapshot, *index)?;
            if conflict.scope == Scope::Property || conflict.span == conflict.start_tag {
                splices.push(Splice {
                    span: conflict.span,
                    replacement: Vec::new(),
                });
            } else {
                splices.push(Splice {
                    span: conflict.start_tag,
                    replacement: Vec::new(),
                });
                if conflict.content.end() < conflict.span.end() {
                    splices.push(Splice {
                        span: Span::new(conflict.content.end(), conflict.span.end())?,
                        replacement: Vec::new(),
                    });
                }
            }
            Ok(())
        },
        Edit::SetRangeAuthor { index, author } => {
            let range = range_at(snapshot, *index)?;
            replace_attribute(
                snapshot.source(),
                range.start_span,
                range.author_span,
                range.author_span,
                "author",
                Some(author),
                splices,
            )
        },
        Edit::SetRangeDate { index, date } => {
            let range = range_at(snapshot, *index)?;
            replace_attribute(
                snapshot.source(),
                range.start_span,
                range.date_span,
                range.author_span,
                "date",
                date.as_ref(),
                splices,
            )
        },
        Edit::RemoveRange { index } => {
            let range = range_at(snapshot, *index)?;
            splices.push(Splice {
                span: range.start_span,
                replacement: Vec::new(),
            });
            splices.push(Splice {
                span: range.end_span,
                replacement: Vec::new(),
            });
            Ok(())
        },
    }
}

fn replace_attribute(
    source: &[u8],
    tag: Span,
    current: Option<AttributeSpan>,
    author: Option<AttributeSpan>,
    local_name: &str,
    value: Option<&String>,
    splices: &mut Vec<Splice>,
) -> Result<()> {
    match (current, value) {
        (Some(current), Some(value)) => splices.push(Splice {
            span: current.value,
            replacement: escape_attribute(value)?,
        }),
        (Some(current), None) => splices.push(Splice {
            span: current.attribute,
            replacement: Vec::new(),
        }),
        (None, None) => {},
        (None, Some(value)) => {
            let author = author.ok_or_else(|| {
                Error::Invalid("conflict metadata has no qualified author attribute span".into())
            })?;
            let prefix = attribute_prefix(source, author.attribute)?;
            let insertion = tag_insertion_offset(source, tag)?;
            let escaped = escape_attribute(value)?;
            let capacity = 5usize
                .checked_add(prefix.len())
                .and_then(|size| size.checked_add(local_name.len()))
                .and_then(|size| size.checked_add(escaped.len()))
                .ok_or_else(|| Error::Invalid("conflict attribute size overflow".into()))?;
            let mut replacement = Vec::new();
            replacement
                .try_reserve_exact(capacity)
                .map_err(|source| Error::Allocation {
                    resource: "conflict attribute",
                    source,
                })?;
            replacement.push(b' ');
            replacement.extend_from_slice(prefix);
            replacement.extend_from_slice(local_name.as_bytes());
            replacement.extend_from_slice(b"=\"");
            replacement.extend_from_slice(&escaped);
            replacement.push(b'"');
            splices.push(Splice {
                span: Span::new(insertion, insertion)?,
                replacement,
            });
        },
    }
    Ok(())
}

fn attribute_prefix(source: &[u8], span: Span) -> Result<&[u8]> {
    let bytes = checked_slice(source, span)?;
    let equal = bytes
        .iter()
        .position(|byte| *byte == b'=')
        .ok_or_else(|| Error::Invalid("conflict attribute has no equals delimiter".into()))?;
    let name = trim_ascii(&bytes[..equal]);
    let colon = name
        .iter()
        .rposition(|byte| *byte == b':')
        .ok_or_else(|| Error::Invalid("conflict metadata attribute is not qualified".into()))?;
    Ok(&name[..=colon])
}

fn tag_insertion_offset(source: &[u8], span: Span) -> Result<usize> {
    let bytes = checked_slice(source, span)?;
    if bytes.last() != Some(&b'>') {
        return Err(Error::Invalid(
            "conflict start tag has no closing delimiter".into(),
        ));
    }
    let relative = if bytes.len() >= 2 && bytes[bytes.len() - 2] == b'/' {
        bytes.len() - 2
    } else {
        bytes.len() - 1
    };
    span.start()
        .checked_add(relative)
        .ok_or_else(|| Error::Invalid("conflict tag offset overflow".into()))
}

fn escape_attribute(value: &str) -> Result<Vec<u8>> {
    let capacity = value
        .len()
        .checked_mul(6)
        .ok_or_else(|| Error::Invalid("conflict metadata size overflow".into()))?;
    let mut escaped = Vec::new();
    escaped
        .try_reserve(capacity)
        .map_err(|source| Error::Allocation {
            resource: "escaped conflict metadata",
            source,
        })?;
    for byte in value.bytes() {
        match byte {
            b'&' => escaped.extend_from_slice(b"&amp;"),
            b'<' => escaped.extend_from_slice(b"&lt;"),
            b'>' => escaped.extend_from_slice(b"&gt;"),
            b'"' => escaped.extend_from_slice(b"&quot;"),
            b'\'' => escaped.extend_from_slice(b"&apos;"),
            _ => escaped.push(byte),
        }
    }
    Ok(escaped)
}

fn apply_splices(source: &[u8], splices: &mut [Splice], limit: usize) -> Result<Vec<u8>> {
    splices.sort_unstable_by_key(|splice| (splice.span.start(), splice.span.end()));
    let mut cursor = 0usize;
    let mut removed = 0usize;
    let mut added = 0usize;
    for splice in splices.iter() {
        checked_slice(source, splice.span)?;
        if splice.span.start() < cursor {
            return Err(Error::Invalid(
                "conflict XML edits overlap or target the same source span".into(),
            ));
        }
        cursor = splice.span.end();
        removed = removed
            .checked_add(splice.span.len())
            .ok_or_else(|| Error::Invalid("conflict splice size overflow".into()))?;
        added = added
            .checked_add(splice.replacement.len())
            .ok_or_else(|| Error::Invalid("conflict splice size overflow".into()))?;
    }
    let output_len = source
        .len()
        .checked_sub(removed)
        .and_then(|size| size.checked_add(added))
        .ok_or_else(|| Error::Invalid("conflict output size overflow".into()))?;
    if output_len > limit {
        return Err(Error::Invalid(format!(
            "conflict output exceeds the {limit}-byte source limit"
        )));
    }

    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "conflict XML output",
            source,
        })?;
    cursor = 0;
    for splice in splices {
        output.extend_from_slice(&source[cursor..splice.span.start()]);
        output.extend_from_slice(&splice.replacement);
        cursor = splice.span.end();
    }
    output.extend_from_slice(&source[cursor..]);
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

fn checked_slice(source: &[u8], span: Span) -> Result<&[u8]> {
    source
        .get(span.start()..span.end())
        .ok_or_else(|| Error::Invalid("conflict source span is stale or out of bounds".into()))
}

fn trim_ascii(mut bytes: &[u8]) -> &[u8] {
    while bytes.first().is_some_and(u8::is_ascii_whitespace) {
        bytes = &bytes[1..];
    }
    while bytes.last().is_some_and(u8::is_ascii_whitespace) {
        bytes = &bytes[..bytes.len() - 1];
    }
    bytes
}

fn conflict_at(snapshot: &Snapshot, index: usize) -> Result<&super::Conflict> {
    snapshot
        .inventory()
        .conflicts
        .get(index)
        .ok_or(Error::OutOfBounds {
            object: "conflict annotation",
            index,
            len: snapshot.inventory().conflicts.len(),
        })
}

fn range_at(snapshot: &Snapshot, index: usize) -> Result<&super::Range> {
    snapshot
        .inventory()
        .ranges
        .get(index)
        .ok_or(Error::OutOfBounds {
            object: "conflict range",
            index,
            len: snapshot.inventory().ranges.len(),
        })
}

fn refuse_removed_conflict(edit_index: &HashMap<EditKey, usize>, index: usize) -> Result<()> {
    if edit_index.contains_key(&EditKey::RemoveConflict(index)) {
        return Err(Error::Invalid(
            "cannot edit metadata on a removed conflict annotation".into(),
        ));
    }
    Ok(())
}

fn refuse_removed_range(edit_index: &HashMap<EditKey, usize>, index: usize) -> Result<()> {
    if edit_index.contains_key(&EditKey::RemoveRange(index)) {
        return Err(Error::Invalid(
            "cannot edit metadata on a removed conflict range".into(),
        ));
    }
    Ok(())
}

fn replace_field_edit(
    edits: &mut Vec<Option<Edit>>,
    edit_index: &mut HashMap<EditKey, usize>,
    key: EditKey,
    replacement: Option<Edit>,
) -> Result<()> {
    if let Some(index) = edit_index.get(&key).copied() {
        if let Some(replacement) = replacement {
            edits[index] = Some(replacement);
        } else {
            remove_edit(edits, edit_index, key);
        }
    } else if let Some(replacement) = replacement {
        reserve_new_edit(edits, edit_index)?;
        insert_reserved_edit(edits, edit_index, key, replacement);
    }
    Ok(())
}

fn reserve_new_edit(
    edits: &mut Vec<Option<Edit>>,
    edit_index: &mut HashMap<EditKey, usize>,
) -> Result<()> {
    if edits.len() > edit_index.len().saturating_mul(2) {
        edits.retain(Option::is_some);
        for (position, edit) in edits.iter().enumerate() {
            if let Some(edit) = edit {
                edit_index.insert(edit_key(edit), position);
            }
        }
    }
    edits.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "conflict transaction edits",
        source,
    })?;
    edit_index
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "conflict transaction edit index",
            source,
        })
}

fn insert_reserved_edit(
    edits: &mut Vec<Option<Edit>>,
    edit_index: &mut HashMap<EditKey, usize>,
    key: EditKey,
    edit: Edit,
) {
    let position = edits.len();
    edits.push(Some(edit));
    edit_index.insert(key, position);
}

fn remove_edit(
    edits: &mut Vec<Option<Edit>>,
    edit_index: &mut HashMap<EditKey, usize>,
    key: EditKey,
) {
    let Some(position) = edit_index.remove(&key) else {
        return;
    };
    edits[position] = None;
}

fn edit_key(edit: &Edit) -> EditKey {
    match edit {
        Edit::SetConflictAuthor { index, .. } => EditKey::ConflictAuthor(*index),
        Edit::SetConflictDate { index, .. } => EditKey::ConflictDate(*index),
        Edit::RemoveConflict { index } => EditKey::RemoveConflict(*index),
        Edit::SetRangeAuthor { index, .. } => EditKey::RangeAuthor(*index),
        Edit::SetRangeDate { index, .. } => EditKey::RangeDate(*index),
        Edit::RemoveRange { index } => EditKey::RemoveRange(*index),
    }
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}

#[cfg(test)]
mod tests {
    use super::{Edit, Snapshot};

    #[test]
    fn projected_edits_retain_request_order_after_o1_removal() {
        let source = br#"<w:document
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="w14"><w:body><w:p>
            <w14:conflictIns w:id="1" w:author="first"/>
            <w14:conflictDel w:id="2" w:author="second"/>
            </w:p></w:body></w:document>"#;
        let snapshot = Snapshot::from_xml(source.as_slice()).expect("valid conflict source");
        let mut transaction = snapshot.edit();
        transaction
            .set_conflict_author(0, "temporary")
            .expect("first edit");
        transaction
            .set_conflict_author(1, "second edited")
            .expect("second edit");
        transaction
            .set_conflict_author(0, "first")
            .expect("remove projected edit");
        transaction
            .set_conflict_author(0, "first edited")
            .expect("re-add projected edit");

        let order: Vec<_> = transaction
            .edits()
            .map(|edit| match edit {
                Edit::SetConflictAuthor { index, .. } => *index,
                _ => usize::MAX,
            })
            .collect();
        assert_eq!(order, [1, 0]);
    }
}
