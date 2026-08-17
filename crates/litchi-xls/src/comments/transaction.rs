//! Source-checked updates of existing BIFF8 worksheet comments.
//!
//! This owner deliberately updates existing NOTE/TXO author and text fields
//! only. Visibility changes are typed refusals because their `SpContainer` and
//! `Opt` structure is not owned by this record-level editor.
//! Adding or deleting a note also changes the worksheet and workbook-global
//! OfficeArt object catalog; those lifecycle operations remain typed refusals
//! until that complete closure has a lossless editor.

use super::{CONTINUE_TYPE, Comment, MSODRAWING_TYPE, OBJ_TYPE, RECORD_TYPE, TXO_TYPE, Visibility};
use crate::cell_values::{Reference, Selector};
use crate::records::{BoundSheetRecord, Encoding, SheetType};
use crate::{Error, Result, Workbook};
use litchi_biff::Records;
use litchi_cfb::{
    ArtifactFingerprint, ComposedOverlaySource, OverlayError, PublishReport,
    SameLengthStreamSplice, StreamSpliceLimits, ValidatedOverlayPlan,
};
use litchi_core::binary;
use litchi_core::{ReadAt, SourceVersion};
use litchi_ole_common::object::{Editor as PackageEditor, Limits, Targets};
use litchi_ole_common::source_backed_overlay::SourceBackedOverlayPublisher;
use std::collections::{BTreeMap, HashMap, HashSet};
use std::fmt;
use std::io::{Cursor, Read, Seek, SeekFrom, Write};
use std::path::Path;
use std::sync::Arc;

const EOF: u16 = 0x000A;
const CODE_PAGE: u16 = 0x0042;
const BOUND_SHEET: u16 = 0x0085;
const FILE_PASS: u16 = 0x002F;
const DBCELL: u16 = 0x00D7;
const WORKBOOK_GLOBALS: u16 = 0x0005;
const WORKSHEET: u16 = 0x0010;
const COMMENT_OBJECT_TYPE: u16 = 0x0019;
const MAX_EDITS: usize = 256;
const MAX_RECORD_BYTES: usize = 8_224;

/// The author and text fields editable on one existing legacy comment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Value {
    author: String,
    text: String,
}

impl Value {
    /// Constructs a checked comment value.
    ///
    /// # Errors
    ///
    /// Returns an error when the author or text is outside BIFF8 bounds.
    pub fn new(author: impl Into<String>, text: impl Into<String>) -> Result<Self> {
        let value = Self {
            author: author.into(),
            text: text.into(),
        };
        validate_value(&value)?;
        Ok(value)
    }

    /// Builds the editable value retained by a parsed comment.
    #[must_use]
    pub fn from_comment(comment: &Comment) -> Self {
        Self {
            author: comment.author().to_string(),
            text: comment.text().to_string(),
        }
    }

    #[must_use]
    pub fn author(&self) -> &str {
        &self.author
    }

    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

/// One selector-first replacement in a bounded worksheet batch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Update {
    reference: Reference,
    value: Value,
}

impl Update {
    #[must_use]
    pub const fn new(reference: Reference, value: Value) -> Self {
        Self { reference, value }
    }

    #[must_use]
    pub const fn reference(&self) -> Reference {
        self.reference
    }

    #[must_use]
    pub const fn value(&self) -> &Value {
        &self.value
    }
}

#[derive(Debug, Clone)]
struct RawRecord {
    start: usize,
    end: usize,
}

#[derive(Debug, Clone)]
struct Site {
    note: RawRecord,
    txo_family: Vec<RawRecord>,
    source_run_bytes: Vec<u8>,
}

#[derive(Debug, Clone)]
struct Entry {
    comment: Comment,
    site: Site,
}

#[derive(Debug, Clone)]
struct Sheet {
    name: String,
    workbook_position: usize,
    protected: bool,
    entries: Vec<Entry>,
}

struct Inner {
    bytes: Arc<[u8]>,
    workbook_path: Vec<String>,
    workbook_stream: Arc<[u8]>,
    sheets: Vec<Sheet>,
}

/// Immutable inventory of existing, safely updatable BIFF8 comments.
#[derive(Clone)]
pub struct Snapshot {
    inner: Arc<Inner>,
}

impl Snapshot {
    /// Opens an unencrypted, unsigned XLS package and validates all comments.
    ///
    /// # Errors
    ///
    /// Returns a package, workbook, comment-linkage, drawing-ambiguity, or
    /// finite-bound error before publishing the snapshot.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = PackageEditor::open(bytes, Targets::default(), Limits::default())?;
        Self::from_package(package)
    }

    fn from_package(package: PackageEditor) -> Result<Self> {
        let workbook_path = [vec!["Workbook".to_string()], vec!["Book".to_string()]]
            .into_iter()
            .find(|path| package.stream(path).is_some())
            .ok_or_else(|| {
                Error::InvalidData("XLS package has no Workbook or Book stream".into())
            })?;
        let workbook_stream = package
            .stream_shared(&workbook_path)
            .ok_or_else(|| Error::InvalidData("selected XLS Workbook stream disappeared".into()))?;
        let bytes = package.finish()?;
        let workbook = Workbook::new(Cursor::new(bytes.as_slice()))?;
        let sheets = parse_inventory(&workbook_stream, &workbook)?;
        Ok(Self {
            inner: Arc::new(Inner {
                bytes: Arc::from(bytes),
                workbook_path,
                workbook_stream,
                sheets,
            }),
        })
    }

    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.inner.bytes
    }

    #[must_use]
    pub fn workbook_stream(&self) -> &[u8] {
        &self.inner.workbook_stream
    }

    #[must_use]
    pub fn worksheet_count(&self) -> usize {
        self.inner.sheets.len()
    }

    /// Resolves a worksheet through its semantic name or workbook tab position.
    ///
    /// # Errors
    ///
    /// Returns an ambiguity error for duplicate case-insensitive names.
    pub fn worksheet<'a>(&'a self, selector: Selector<'_>) -> Result<Option<Worksheet<'a>>> {
        Ok(self.resolve_sheet(selector)?.map(|index| Worksheet {
            snapshot: self,
            index,
        }))
    }

    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            source: self.clone(),
            changes: Vec::new(),
        }
    }

    fn resolve_sheet(&self, selector: Selector<'_>) -> Result<Option<usize>> {
        match selector {
            Selector::Position(position) => Ok(self
                .inner
                .sheets
                .iter()
                .position(|sheet| sheet.workbook_position == position)),
            Selector::Name(name) => {
                let mut matches = self
                    .inner
                    .sheets
                    .iter()
                    .enumerate()
                    .filter(|(_, sheet)| sheet.name.eq_ignore_ascii_case(name));
                let found = matches.next().map(|(index, _)| index);
                if matches.next().is_some() {
                    return Err(Error::UnsafeEdit(format!(
                        "worksheet name {name:?} is ambiguous"
                    )));
                }
                Ok(found)
            },
        }
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.bytes() == other.bytes()
    }
}

impl Eq for Snapshot {}

impl fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("artifact_bytes", &self.bytes().len())
            .field("workbook_bytes", &self.workbook_stream().len())
            .field("worksheets", &self.worksheet_count())
            .finish()
    }
}

/// Borrowed comment inventory for one worksheet.
#[derive(Debug, Clone, Copy)]
pub struct Worksheet<'a> {
    snapshot: &'a Snapshot,
    index: usize,
}

impl<'a> Worksheet<'a> {
    fn data(self) -> &'a Sheet {
        &self.snapshot.inner.sheets[self.index]
    }

    #[must_use]
    pub fn name(self) -> &'a str {
        &self.data().name
    }

    #[must_use]
    pub fn position(self) -> usize {
        self.data().workbook_position
    }

    #[must_use]
    pub fn comments(self) -> impl ExactSizeIterator<Item = &'a Comment> {
        self.data().entries.iter().map(|entry| &entry.comment)
    }

    /// Finds one existing comment by its exact cell identity.
    ///
    /// # Errors
    ///
    /// Returns an ambiguity error if the source contains duplicate NOTE cells.
    pub fn comment(self, reference: Reference) -> Result<Option<&'a Comment>> {
        Ok(unique_entry(&self.data().entries, reference)?.map(|entry| &entry.comment))
    }
}

#[derive(Debug, Clone)]
struct Change {
    sheet: usize,
    entry: usize,
    before: Value,
    after: Value,
}

/// Detached, failure-atomic replacement of existing BIFF8 comments.
#[derive(Clone)]
pub struct Edit {
    source: Snapshot,
    changes: Vec<Change>,
}

impl Edit {
    /// Replaces the complete text and author of one existing comment.
    ///
    /// # Errors
    ///
    /// Returns a selector, protection, source-ambiguity, encoding, drawing, or
    /// transaction-bound refusal without partially staging this replacement.
    pub fn replace(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        value: Value,
    ) -> Result<()> {
        let sheet = self
            .source
            .resolve_sheet(selector)?
            .ok_or_else(|| Error::WorksheetNotFound("comment edit worksheet selector".into()))?;
        self.stage(sheet, reference, value)
    }

    /// Atomically stages up to 256 distinct replacements on one worksheet.
    ///
    /// # Errors
    ///
    /// Every selector, duplicate, value, protection state, and finite bound is
    /// checked against a detached clone before this edit changes.
    pub fn replace_many(
        &mut self,
        selector: Selector<'_>,
        updates: impl IntoIterator<Item = Update>,
    ) -> Result<()> {
        let sheet = self
            .source
            .resolve_sheet(selector)?
            .ok_or_else(|| Error::WorksheetNotFound("comment edit worksheet selector".into()))?;
        let mut staged = self.clone();
        let mut seen = HashSet::new();
        for update in updates {
            if !seen.insert(update.reference) {
                return Err(Error::UnsafeEdit(format!(
                    "duplicate comment selector ({}, {}) in one batch",
                    update.reference.row(),
                    update.reference.column()
                )));
            }
            staged.stage(sheet, update.reference, update.value)?;
        }
        *self = staged;
        Ok(())
    }

    /// Refuses deletion until the complete OfficeArt catalog closure is owned.
    ///
    /// # Errors
    ///
    /// Always returns [`Error::UnsafeEdit`]; the transaction is unchanged.
    pub fn remove(&mut self, _selector: Selector<'_>, _reference: Reference) -> Result<()> {
        Err(Error::UnsafeEdit(
            "deleting an XLS NOTE requires lossless worksheet and workbook-global OfficeArt catalog updates"
                .into(),
        ))
    }

    fn stage(&mut self, sheet: usize, reference: Reference, value: Value) -> Result<()> {
        validate_value(&value)?;
        let owner = self
            .source
            .inner
            .sheets
            .get(sheet)
            .ok_or_else(|| Error::UnsafeEdit("comment worksheet inventory is stale".into()))?;
        let entry = unique_entry_index(&owner.entries, reference)?.ok_or_else(|| {
            Error::UnsupportedFeature(format!(
                "cell ({}, {}) has no existing BIFF8 NOTE comment",
                reference.row(),
                reference.column()
            ))
        })?;
        let before = Value::from_comment(&owner.entries[entry].comment);
        if value == before {
            if let Some(existing) = self
                .changes
                .iter_mut()
                .find(|change| change.sheet == sheet && change.entry == entry)
            {
                existing.after = value;
            }
            return Ok(());
        }
        if before.text.is_empty() && !value.text.is_empty() {
            return Err(Error::UnsafeEdit(
                "an empty source TXO has no formatting-run identity for nonempty comment text"
                    .into(),
            ));
        }
        if owner.protected {
            return Err(Error::UnsafeEdit(
                "comments on a protected BIFF8 worksheet cannot be changed".into(),
            ));
        }
        if let Some(existing) = self
            .changes
            .iter_mut()
            .find(|change| change.sheet == sheet && change.entry == entry)
        {
            existing.after = value;
            return Ok(());
        }
        if self.changes.len() >= MAX_EDITS {
            return Err(Error::UnsafeEdit(format!(
                "comment edit count exceeds the finite limit of {MAX_EDITS}"
            )));
        }
        self.changes
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("staging XLS comment changes"))?;
        self.changes.push(Change {
            sheet,
            entry,
            before,
            after: value,
        });
        Ok(())
    }

    /// Checks a visibility request without mutating OfficeArt.
    ///
    /// Exact no-ops succeed, including on protected worksheets. A real
    /// visibility change is refused because it requires structurally owning
    /// the complete `SpContainer`/`Opt` shape-property closure.
    ///
    /// # Errors
    ///
    /// Returns a selector, missing-comment, ambiguity, or unsupported drawing
    /// ownership error without staging any bytes.
    pub fn set_visibility(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        visibility: Visibility,
    ) -> Result<()> {
        let sheet = self
            .source
            .resolve_sheet(selector)?
            .ok_or_else(|| Error::WorksheetNotFound("comment edit worksheet selector".into()))?;
        let owner = &self.source.inner.sheets[sheet];
        let entry = unique_entry(&owner.entries, reference)?.ok_or_else(|| {
            Error::UnsupportedFeature(format!(
                "cell ({}, {}) has no existing BIFF8 NOTE comment",
                reference.row(),
                reference.column()
            ))
        })?;
        if entry.comment.visibility() == visibility {
            return Ok(());
        }
        Err(Error::UnsafeEdit(
            "changing XLS comment visibility requires complete SpContainer/Opt ownership".into(),
        ))
    }

    /// Publishes one reopened CFB package and an exact reversible patch.
    ///
    /// # Errors
    ///
    /// Returns a stale-source, BIFF splice, package, complete-reopen, or
    /// semantic-readback error without publishing a partial result.
    pub fn commit(self) -> Result<Commit> {
        let effective: Vec<_> = self
            .changes
            .iter()
            .filter(|change| change.before != change.after)
            .collect();
        if effective.is_empty() {
            let patch = Patch::new(
                Arc::clone(&self.source.inner.bytes),
                Arc::clone(&self.source.inner.bytes),
                Vec::new(),
            );
            return Ok(Commit {
                snapshot: self.source,
                patch,
                diagnostics: Diagnostics::default(),
            });
        }

        let mut replacements = Vec::new();
        for change in &effective {
            let entry = &self.source.inner.sheets[change.sheet].entries[change.entry];
            if change.before.author != change.after.author {
                replacements.push((
                    entry.site.note.start,
                    entry.site.note.end,
                    encode_note(&entry.comment, &change.after)?,
                ));
            }
            if change.before.text != change.after.text {
                let first = entry.site.txo_family.first().ok_or_else(|| {
                    Error::InvalidData("comment TXO source family is empty".into())
                })?;
                let last = entry.site.txo_family.last().ok_or_else(|| {
                    Error::InvalidData("comment TXO source family is empty".into())
                })?;
                replacements.push((
                    first.start,
                    last.end,
                    encode_txo_family(
                        &self.source.inner.workbook_stream,
                        entry,
                        &change.after.text,
                        None,
                    )?,
                ));
            }
        }
        replacements.sort_by_key(|replacement| std::cmp::Reverse(replacement.0));
        ensure_disjoint(&replacements)?;
        let mut workbook = self.source.inner.workbook_stream.to_vec();
        crate::cell_values::replace_workbook_ranges_and_adjust_bounds(
            &mut workbook,
            &replacements,
        )?;

        let mut package = PackageEditor::open(
            self.source.inner.bytes.to_vec(),
            Targets::default(),
            Limits::default(),
        )?;
        package.put_stream_shared(&self.source.inner.workbook_path, Arc::from(workbook))?;
        let snapshot = Snapshot::from_package(package)?;
        verify_readback(&snapshot, &self.source, &effective)?;
        let mut operations = Vec::new();
        operations
            .try_reserve_exact(effective.len())
            .map_err(|_error| Error::Allocation("retaining XLS comment patch operations"))?;
        for change in &effective {
            let sheet = &self.source.inner.sheets[change.sheet];
            let comment = &sheet.entries[change.entry].comment;
            operations.push(Operation {
                sheet_position: sheet.workbook_position,
                reference: Reference::new(u32::from(comment.row()), u32::from(comment.column()))?,
                before: change.before.clone(),
                after: change.after.clone(),
            });
        }
        let patch = Patch::new(
            Arc::clone(&self.source.inner.bytes),
            Arc::clone(&snapshot.inner.bytes),
            operations,
        );
        Ok(Commit {
            snapshot,
            patch,
            diagnostics: Diagnostics {
                changed_comments: effective.len(),
                touched_streams: 1,
            },
        })
    }

    /// Plans a protected, source-backed same-length comment publication.
    ///
    /// This path owns only existing NOTE/TXO families. Every generated
    /// replacement must retain its exact source range length and its source
    /// compressed/UTF-16 encoding width. Each range is submitted as a bounded
    /// [`SameLengthStreamSplice`] to [`SourceBackedOverlayPublisher`]. No CFB
    /// render fallback is attempted; callers needing length-changing edits
    /// must explicitly call [`Self::commit`].
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for protected CFB markers, stale or malformed
    /// source state, unsupported encoding transitions, length/topology changes,
    /// overlapping NOTE/TXO ownership, or failed semantic readback.
    pub fn commit_source_backed(self) -> Result<SourceBackedCommit> {
        let effective: Vec<_> = self
            .changes
            .iter()
            .filter(|change| change.before != change.after)
            .collect();

        let mut replacements = Vec::new();
        let replacement_capacity = effective
            .len()
            .checked_mul(2)
            .ok_or_else(|| Error::UnsafeEdit("comment replacement count overflow".into()))?;
        replacements
            .try_reserve_exact(replacement_capacity)
            .map_err(|_error| Error::Allocation("retaining source-backed XLS comment ranges"))?;
        for change in &effective {
            let sheet =
                self.source.inner.sheets.get(change.sheet).ok_or_else(|| {
                    Error::UnsafeEdit("source-backed comment sheet is stale".into())
                })?;
            let entry = sheet
                .entries
                .get(change.entry)
                .ok_or_else(|| Error::UnsafeEdit("source-backed comment entry is stale".into()))?;

            if change.before.author != change.after.author {
                let compressed = ensure_author_encoding_width(
                    &self.source.inner.workbook_stream,
                    entry,
                    &change.after.author,
                )?;
                let replacement =
                    encode_note_with_compression(&entry.comment, &change.after, compressed)?;
                ensure_equal_range_length(
                    entry.site.note.start,
                    entry.site.note.end,
                    &replacement,
                )?;
                replacements.push((entry.site.note.start, entry.site.note.end, replacement));
            }
            if change.before.text != change.after.text {
                let compressed = ensure_text_encoding_width(
                    &self.source.inner.workbook_stream,
                    entry,
                    &change.after.text,
                )?;
                let first = entry.site.txo_family.first().ok_or_else(|| {
                    Error::InvalidData("comment TXO source family is empty".into())
                })?;
                let last = entry.site.txo_family.last().ok_or_else(|| {
                    Error::InvalidData("comment TXO source family is empty".into())
                })?;
                let replacement = encode_txo_family(
                    &self.source.inner.workbook_stream,
                    entry,
                    &change.after.text,
                    Some(compressed),
                )?;
                ensure_equal_range_length(first.start, last.end, &replacement)?;
                replacements.push((first.start, last.end, replacement));
            }
        }
        replacements.sort_by_key(|replacement| std::cmp::Reverse(replacement.0));
        ensure_disjoint(&replacements)?;

        let mut splices = Vec::new();
        splices
            .try_reserve_exact(replacements.len())
            .map_err(|_error| Error::Allocation("staging source-backed comment splices"))?;
        let splice_count = replacements.len();
        let mut replacement_bytes = 0_u64;
        for (start, end, replacement) in replacements {
            ensure_equal_range_length(start, end, &replacement)?;
            let expected = self
                .source
                .inner
                .workbook_stream
                .get(start..end)
                .ok_or_else(|| {
                    Error::UnsafeEdit("source-backed comment range is outside Workbook".into())
                })?;
            let offset = u64::try_from(start).map_err(|_error| {
                Error::InvalidData("source-backed comment range offset exceeds u64".into())
            })?;
            replacement_bytes = replacement_bytes
                .checked_add(u64::try_from(replacement.len()).map_err(|_error| {
                    Error::InvalidData("source-backed comment range exceeds u64".into())
                })?)
                .ok_or_else(|| {
                    Error::InvalidData("source-backed comment replacement bytes overflow".into())
                })?;
            splices.push(SameLengthStreamSplice::new(
                self.source.inner.workbook_path.clone(),
                offset,
                Arc::from(expected),
                Arc::from(replacement),
            ));
        }

        // Preserve the snapshot's immutable Arc ownership through CFB. This
        // permits direct sequential publication to omit redundant outer
        // mutation fences while retaining emission hashes; composed views and
        // atomic save remain fully fenced.
        let source = Arc::clone(&self.source.inner.bytes);
        let source_version =
            SourceVersion::new(source.as_ptr() as usize as u64, source.len() as u64);
        let publisher = SourceBackedOverlayPublisher::open_owned(source, source_version)
            .map_err(overlay_to_error)?;
        let (plan, owner_validated) = publisher
            .plan_splices_with_owner(splices, StreamSpliceLimits::default(), |candidate| {
                verify_source_backed_candidate(candidate.clone(), &self.source, &effective)
                    .map_err(CommentOwnerError::Owner)
            })
            .map_err(CommentOwnerError::into_error)?;
        if owner_validated.is_none() && !effective.is_empty() {
            verify_source_backed_readback(&plan, &self.source, &effective)?;
        }

        let source_workbook_bytes = u64::try_from(self.source.inner.workbook_stream.len())
            .map_err(|_error| Error::InvalidData("Workbook stream length exceeds u64".into()))?;
        let target_workbook_bytes = source_workbook_bytes;
        let source_bytes = u64::try_from(self.source.inner.bytes.len())
            .map_err(|_error| Error::InvalidData("source CFB length exceeds u64".into()))?;
        let diagnostics = SourceBackedDiagnostics {
            changed_comments: effective.len(),
            touched_streams: usize::from(!plan.is_noop()),
            splice_count,
            replacement_bytes,
            changed_spans: plan.changed_spans(),
            source_bytes,
            source_workbook_bytes,
            target_workbook_bytes,
            source_fingerprint: plan.source_fingerprint(),
            target_fingerprint: plan.target_fingerprint(),
        };
        Ok(SourceBackedCommit {
            source: self.source,
            plan,
            diagnostics,
        })
    }
}

impl fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("source", &self.source)
            .field("staged_changes", &self.changes.len())
            .finish()
    }
}

/// One reversible semantic comment replacement retained by a patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Operation {
    sheet_position: usize,
    reference: Reference,
    before: Value,
    after: Value,
}

impl Operation {
    #[must_use]
    pub const fn sheet_position(&self) -> usize {
        self.sheet_position
    }

    #[must_use]
    pub const fn reference(&self) -> Reference {
        self.reference
    }

    #[must_use]
    pub const fn before(&self) -> &Value {
        &self.before
    }

    #[must_use]
    pub const fn after(&self) -> &Value {
        &self.after
    }
}

/// Exact-source, reversible replacement of one XLS artifact.
#[derive(Clone)]
pub struct Patch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
    operations: Arc<[Operation]>,
}

impl Patch {
    fn new(before: Arc<[u8]>, after: Arc<[u8]>, operations: Vec<Operation>) -> Self {
        Self {
            before,
            after,
            operations: Arc::from(operations),
        }
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before == self.after
    }

    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    #[must_use]
    pub fn operations(&self) -> &[Operation] {
        &self.operations
    }

    /// Applies this patch only to its exact source artifact.
    ///
    /// # Errors
    ///
    /// Returns an exact-source conflict or candidate validation error.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.bytes() != self.before() {
            return Err(Error::UnsafeEdit(
                "XLS comment patch source does not match its base snapshot".into(),
            ));
        }
        if self.is_empty() {
            return Ok(source.clone());
        }
        Snapshot::from_bytes(self.after.to_vec())
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        let operations: Vec<_> = self
            .operations
            .iter()
            .rev()
            .map(|operation| Operation {
                sheet_position: operation.sheet_position,
                reference: operation.reference,
                before: operation.after.clone(),
                after: operation.before.clone(),
            })
            .collect();
        Self::new(
            Arc::clone(&self.after),
            Arc::clone(&self.before),
            operations,
        )
    }
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("before_bytes", &self.before.len())
            .field("after_bytes", &self.after.len())
            .field("operations", &self.operations.len())
            .finish()
    }
}

/// Content-free publication counters.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Diagnostics {
    changed_comments: usize,
    touched_streams: usize,
}

impl Diagnostics {
    #[must_use]
    pub const fn changed_comments(self) -> usize {
        self.changed_comments
    }

    #[must_use]
    pub const fn touched_streams(self) -> usize {
        self.touched_streams
    }
}

/// Successful immutable XLS comment publication.
#[derive(Clone, Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch, Diagnostics) {
        (self.snapshot, self.patch, self.diagnostics)
    }
}

/// Content-free evidence for a same-length source-backed comment publication.
///
/// The diagnostics intentionally retain only counts, lengths, and opaque CFB
/// fingerprints. Comment authors, text, and physical payload bytes are never
/// copied into this value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourceBackedDiagnostics {
    changed_comments: usize,
    touched_streams: usize,
    splice_count: usize,
    replacement_bytes: u64,
    changed_spans: usize,
    source_bytes: u64,
    source_workbook_bytes: u64,
    target_workbook_bytes: u64,
    source_fingerprint: ArtifactFingerprint,
    target_fingerprint: ArtifactFingerprint,
}

impl SourceBackedDiagnostics {
    /// Number of semantic comments changed by the transaction.
    #[must_use]
    pub const fn changed_comments(self) -> usize {
        self.changed_comments
    }

    /// Number of logical streams selected for the overlay.
    #[must_use]
    pub const fn touched_streams(self) -> usize {
        self.touched_streams
    }

    /// Number of source-relative NOTE/TXO splices submitted to CFB.
    #[must_use]
    pub const fn splice_count(self) -> usize {
        self.splice_count
    }

    /// Aggregate replacement bytes retained by the splice plan.
    #[must_use]
    pub const fn replacement_bytes(self) -> u64 {
        self.replacement_bytes
    }

    /// Number of physical CFB spans changed by the overlay.
    #[must_use]
    pub const fn changed_spans(self) -> usize {
        self.changed_spans
    }

    /// Complete source CFB length.
    #[must_use]
    pub const fn source_bytes(self) -> u64 {
        self.source_bytes
    }

    /// Source Workbook/Book stream length.
    #[must_use]
    pub const fn source_workbook_bytes(self) -> u64 {
        self.source_workbook_bytes
    }

    /// Candidate Workbook/Book stream length.
    #[must_use]
    pub const fn target_workbook_bytes(self) -> u64 {
        self.target_workbook_bytes
    }

    /// Exact source CFB fingerprint.
    #[must_use]
    pub const fn source_fingerprint(self) -> ArtifactFingerprint {
        self.source_fingerprint
    }

    /// Exact candidate CFB fingerprint.
    #[must_use]
    pub const fn target_fingerprint(self) -> ArtifactFingerprint {
        self.target_fingerprint
    }
}

/// A checked source-bound same-length XLS comment publication plan.
///
/// Unlike [`Commit`], this value does not retain a rendered replacement CFB
/// artifact. It retains the exact source-backed overlay plan and streams the
/// validated candidate to either a sequential sink or an atomic destination.
/// A caller that needs variable-length record edits must explicitly use
/// [`Edit::commit`].
pub struct SourceBackedCommit {
    source: Snapshot,
    plan: ValidatedOverlayPlan,
    diagnostics: SourceBackedDiagnostics,
}

impl SourceBackedCommit {
    /// Exact immutable source snapshot used by this plan.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Underlying checked common CFB overlay plan.
    #[must_use]
    pub const fn plan(&self) -> &ValidatedOverlayPlan {
        &self.plan
    }

    /// Content-free publication evidence.
    #[must_use]
    pub const fn diagnostics(&self) -> SourceBackedDiagnostics {
        self.diagnostics
    }

    /// Whether this publication is an exact source identity no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.plan.is_noop()
    }

    /// Streams the complete source-backed candidate to a sequential sink.
    ///
    /// The common overlay publisher hashes the exact source and target during
    /// output. Immutable snapshot provenance makes an additional outer
    /// mutation preflight redundant for this direct sequential path; atomic
    /// save retains its complete pre-rename fences. A sink failure retains the
    /// typed [`litchi_cfb::OutputProgress`] inside [`OverlayError`].
    pub fn write_to<W: Write>(&self, writer: &mut W) -> Result<PublishReport, OverlayError> {
        self.plan.write_to(writer)
    }

    /// Alias for [`Self::write_to`] emphasizing that this is publication.
    pub fn publish_to_stream<W: Write>(
        &self,
        writer: &mut W,
    ) -> Result<PublishReport, OverlayError> {
        self.write_to(writer)
    }

    /// Streams the candidate through the common synced sibling-file and
    /// atomic-rename path.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<PublishReport, OverlayError> {
        self.plan.save(path)
    }
}

impl fmt::Debug for SourceBackedCommit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceBackedCommit")
            .field("source", &self.source)
            .field("plan", &self.plan)
            .field("diagnostics", &self.diagnostics)
            .finish()
    }
}

struct PositionalReader {
    source: ComposedOverlaySource,
    position: u64,
}

impl PositionalReader {
    fn new(source: ComposedOverlaySource) -> Self {
        Self {
            source,
            position: 0,
        }
    }
}

impl Read for PositionalReader {
    fn read(&mut self, output: &mut [u8]) -> std::io::Result<usize> {
        let count = self.source.read_at(self.position, output)?;
        self.position = self
            .position
            .checked_add(count as u64)
            .ok_or_else(|| std::io::Error::other("positional reader offset overflow"))?;
        Ok(count)
    }
}

impl Seek for PositionalReader {
    fn seek(&mut self, position: SeekFrom) -> std::io::Result<u64> {
        let length = self.source.len()?;
        let target = match position {
            SeekFrom::Start(offset) => i128::from(offset),
            SeekFrom::Current(offset) => i128::from(self.position) + i128::from(offset),
            SeekFrom::End(offset) => i128::from(length) + i128::from(offset),
        };
        if target < 0 || target > i128::from(u64::MAX) {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "positional reader seek is outside the source",
            ));
        }
        self.position = target as u64;
        Ok(self.position)
    }
}

fn overlay_to_error(error: OverlayError) -> Error {
    match error {
        OverlayError::Ole(error) => Error::Cfb(error),
        OverlayError::Io(error) => Error::Io(error),
        other => Error::UnsafeEdit(format!("source-backed comment overlay refused: {other}")),
    }
}

enum CommentOwnerError {
    Overlay(OverlayError),
    Owner(Error),
}

impl CommentOwnerError {
    fn into_error(self) -> Error {
        match self {
            Self::Overlay(error) => overlay_to_error(error),
            Self::Owner(error) => error,
        }
    }
}

impl From<OverlayError> for CommentOwnerError {
    fn from(error: OverlayError) -> Self {
        Self::Overlay(error)
    }
}

fn ensure_equal_range_length(start: usize, end: usize, replacement: &[u8]) -> Result<()> {
    let source_length = end.checked_sub(start).ok_or_else(|| {
        Error::UnsafeEdit("source-backed comment range has reversed bounds".into())
    })?;
    if source_length != replacement.len() {
        return Err(Error::UnsafeEdit(
            "source-backed comment replacement changes a NOTE/TXO range length".into(),
        ));
    }
    Ok(())
}

fn ensure_author_encoding_width(stream: &[u8], entry: &Entry, author: &str) -> Result<bool> {
    let payload = stream
        .get(entry.site.note.start + 4..entry.site.note.end)
        .ok_or_else(|| Error::InvalidData("comment NOTE source payload is truncated".into()))?;
    let source_compressed = payload
        .get(10)
        .ok_or_else(|| Error::InvalidData("comment NOTE source flags are truncated".into()))?
        & 1
        == 0;
    let target_compressed = author.encode_utf16().all(|unit| unit <= 0x00FF);
    if source_compressed && !target_compressed {
        return Err(Error::UnsafeEdit(
            "source-backed comment author changes encoding width".into(),
        ));
    }
    Ok(source_compressed)
}

fn ensure_text_encoding_width(stream: &[u8], entry: &Entry, text: &str) -> Result<bool> {
    let source_compressed = source_text_compression(stream, entry)?;
    let target_compressed = text.encode_utf16().all(|unit| unit <= 0x00FF);
    if source_compressed.is_some_and(|value| value && !target_compressed) {
        return Err(Error::UnsafeEdit(
            "source-backed comment text changes encoding width".into(),
        ));
    }
    source_compressed.ok_or_else(|| {
        Error::UnsafeEdit(
            "source-backed comment text has no source encoding width for replacement".into(),
        )
    })
}

fn source_text_compression(stream: &[u8], entry: &Entry) -> Result<Option<bool>> {
    let txo = entry
        .site
        .txo_family
        .first()
        .ok_or_else(|| Error::InvalidData("comment TXO source family is empty".into()))?;
    let txo_payload = stream
        .get(txo.start + 4..txo.end)
        .ok_or_else(|| Error::InvalidData("comment TXO source payload is truncated".into()))?;
    if txo_payload.len() < 14 {
        return Err(Error::InvalidData("comment TXO source is truncated".into()));
    }
    let character_count =
        usize::from(u16::from_le_bytes(txo_payload[10..12].try_into().map_err(
            |_error| Error::InvalidData("comment TXO cchText is truncated".into()),
        )?));
    if character_count == 0 {
        return Ok(None);
    }
    let mut remaining = character_count;
    let mut source_compressed = None;
    for continuation in entry.site.txo_family.iter().skip(1) {
        if remaining == 0 {
            break;
        }
        let payload = stream
            .get(continuation.start + 4..continuation.end)
            .ok_or_else(|| Error::InvalidData("comment TXO continuation is truncated".into()))?;
        let flags = *payload.first().ok_or_else(|| {
            Error::InvalidData("comment TXO continuation flags are missing".into())
        })?;
        if flags & !1 != 0 {
            return Err(Error::InvalidData(
                "comment TXO continuation has reserved encoding bits".into(),
            ));
        }
        let compressed = flags & 1 == 0;
        if source_compressed.is_some_and(|value| value != compressed) {
            return Err(Error::UnsafeEdit(
                "source-backed comment TXO uses mixed encoding widths".into(),
            ));
        }
        source_compressed = Some(compressed);
        let width = if compressed { 1 } else { 2 };
        let bytes = payload.len().checked_sub(1).ok_or_else(|| {
            Error::InvalidData("comment TXO continuation has no text payload".into())
        })?;
        if bytes == 0 || bytes % width != 0 {
            return Err(Error::InvalidData(
                "comment TXO continuation has an invalid text width".into(),
            ));
        }
        let count = bytes / width;
        if count > remaining {
            return Err(Error::InvalidData(
                "comment TXO continuation exceeds source text length".into(),
            ));
        }
        remaining -= count;
    }
    if remaining != 0 {
        return Err(Error::InvalidData(
            "comment TXO source text continuation is incomplete".into(),
        ));
    }
    source_compressed
        .ok_or_else(|| Error::InvalidData("comment TXO source text continuation is missing".into()))
        .map(Some)
}

fn verify_source_backed_readback(
    plan: &ValidatedOverlayPlan,
    source: &Snapshot,
    changes: &[&Change],
) -> Result<()> {
    let candidate = plan.composed_source().map_err(overlay_to_error)?;
    verify_source_backed_candidate(candidate, source, changes)
}

fn verify_source_backed_candidate(
    candidate: ComposedOverlaySource,
    source: &Snapshot,
    changes: &[&Change],
) -> Result<()> {
    let workbook = Workbook::new(PositionalReader::new(candidate))?;
    let worksheet_count = workbook
        .sheets()
        .iter()
        .filter(|sheet| sheet.parsed_worksheet_index().is_some())
        .count();
    if worksheet_count != source.inner.sheets.len() {
        return Err(Error::UnsafeEdit(
            "source-backed comment publication changed the worksheet inventory".into(),
        ));
    }
    for change in changes {
        let sheet = source.inner.sheets.get(change.sheet).ok_or_else(|| {
            Error::UnsafeEdit("source-backed comment sheet disappeared during readback".into())
        })?;
        let metadata = workbook.sheet(sheet.workbook_position).ok_or_else(|| {
            Error::UnsafeEdit("source-backed comment worksheet disappeared during readback".into())
        })?;
        let worksheet_index = metadata.parsed_worksheet_index().ok_or_else(|| {
            Error::UnsafeEdit(
                "source-backed comment worksheet is not parsed during readback".into(),
            )
        })?;
        let comments = workbook.xls_worksheet(worksheet_index)?.comments();
        let comment = comments.iter().find(|comment| {
            comment.row()
                == source.inner.sheets[change.sheet].entries[change.entry]
                    .comment
                    .row()
                && comment.column()
                    == source.inner.sheets[change.sheet].entries[change.entry]
                        .comment
                        .column()
        });
        let comment = comment.ok_or_else(|| {
            Error::UnsafeEdit("source-backed comment NOTE disappeared during readback".into())
        })?;
        if Value::from_comment(comment) != change.after {
            return Err(Error::UnsafeEdit(
                "source-backed comment semantic readback disagreed with the staged value".into(),
            ));
        }
    }
    Ok(())
}

fn validate_value(value: &Value) -> Result<()> {
    let author_units = value.author.encode_utf16().count();
    if !(1..=54).contains(&author_units) {
        return Err(Error::InvalidData(
            "comment author length must be 1..=54 UTF-16 code units".into(),
        ));
    }
    if value.text.encode_utf16().count() > usize::from(u16::MAX) {
        return Err(Error::InvalidData(
            "comment text exceeds 65535 UTF-16 code units".into(),
        ));
    }
    if value.author.contains('\0') || value.text.contains('\0') {
        return Err(Error::InvalidData(
            "comment author and text must not contain NUL".into(),
        ));
    }
    Ok(())
}

fn unique_entry(entries: &[Entry], reference: Reference) -> Result<Option<&Entry>> {
    let mut matches = entries.iter().filter(|entry| {
        entry.comment.row() == reference.row() && entry.comment.column() == reference.column()
    });
    let found = matches.next();
    if matches.next().is_some() {
        return Err(Error::UnsafeEdit(format!(
            "duplicate NOTE cell ({}, {})",
            reference.row(),
            reference.column()
        )));
    }
    Ok(found)
}

fn unique_entry_index(entries: &[Entry], reference: Reference) -> Result<Option<usize>> {
    let mut matches = entries.iter().enumerate().filter(|(_, entry)| {
        entry.comment.row() == reference.row() && entry.comment.column() == reference.column()
    });
    let found = matches.next().map(|(index, _)| index);
    if matches.next().is_some() {
        return Err(Error::UnsafeEdit(format!(
            "duplicate NOTE cell ({}, {})",
            reference.row(),
            reference.column()
        )));
    }
    Ok(found)
}

fn parse_inventory<R: Read + Seek>(
    stream: &Arc<[u8]>,
    workbook: &Workbook<R>,
) -> Result<Vec<Sheet>> {
    let mut records = Records::new(stream);
    let first = records.next().ok_or(Error::Eof("Workbook globals BOF"))??;
    require_bof(first.payload(), WORKBOOK_GLOBALS)?;
    let mut encoding = Encoding::from_codepage(1252)?;
    let mut bounds = Vec::new();
    let mut globals_eof = false;
    for record in records.by_ref() {
        let record = record?;
        match record.kind().get() {
            CODE_PAGE if record.payload().len() == 2 => {
                encoding = Encoding::from_codepage(binary::read_u16_le_at(record.payload(), 0)?)?;
            },
            BOUND_SHEET => bounds.push(BoundSheetRecord::parse(record.payload(), &encoding)?),
            FILE_PASS => return Err(Error::PasswordRequired),
            EOF => {
                globals_eof = true;
                break;
            },
            _ => {},
        }
    }
    if !globals_eof {
        return Err(Error::Eof("Workbook globals EOF"));
    }
    if bounds.len() != workbook.sheets().len() {
        return Err(Error::UnsafeEdit(
            "BoundSheet inventory disagrees with the complete workbook reader".into(),
        ));
    }
    let mut positions = HashSet::new();
    let mut sheets = Vec::new();
    for (position, bound) in bounds.iter().enumerate() {
        if !positions.insert(bound.position) {
            return Err(Error::UnsafeEdit(
                "multiple BoundSheet records share one substream position".into(),
            ));
        }
        if bound.sheet_type != SheetType::WorkSheet {
            continue;
        }
        let metadata = workbook.sheet(position).ok_or_else(|| {
            Error::UnsafeEdit("workbook sheet metadata inventory is incomplete".into())
        })?;
        let parsed_index = metadata.parsed_worksheet_index().ok_or_else(|| {
            Error::UnsafeEdit("a worksheet substream was not completely parsed".into())
        })?;
        let semantic = workbook.xls_worksheet(parsed_index)?;
        let start = usize::try_from(bound.position)
            .map_err(|_error| Error::InvalidData("worksheet position exceeds usize".into()))?;
        let end = bounds
            .iter()
            .filter_map(|candidate| {
                let candidate = usize::try_from(candidate.position).ok()?;
                (candidate > start).then_some(candidate)
            })
            .min()
            .unwrap_or(stream.len());
        let entries = parse_sheet_sites(stream, start, end, semantic.comments())?;
        sheets.push(Sheet {
            name: bound.name.clone(),
            workbook_position: position,
            protected: semantic.protection().is_protected()
                || semantic.protection().objects_protected(),
            entries,
        });
    }
    Ok(sheets)
}

#[derive(Debug)]
struct PartialSite {
    object: RawRecord,
    drawing: Option<RawRecord>,
    txo_family: Vec<RawRecord>,
    note: Option<RawRecord>,
    source_run_bytes: Vec<u8>,
}

#[derive(Debug)]
struct PendingText {
    object_id: u16,
    characters_remaining: usize,
    runs_remaining: usize,
}

fn parse_sheet_sites(
    stream: &[u8],
    start: usize,
    end: usize,
    comments: &[Comment],
) -> Result<Vec<Entry>> {
    let data = stream
        .get(start..end)
        .ok_or_else(|| Error::InvalidData("worksheet range is outside Workbook".into()))?;
    let mut records = Records::new(data);
    let first = records.next().ok_or(Error::Eof("worksheet BOF"))??;
    require_bof(first.payload(), WORKSHEET)?;
    let mut collector = super::CommentCollector::new();
    collector.feed_record(first.kind().get(), first.payload())?;
    let mut sites: HashMap<u16, PartialSite> = HashMap::new();
    let mut pending_object = None;
    let mut pending_text: Option<PendingText> = None;
    let mut last_dbcell_end = start;
    let mut found_eof = false;
    for record in records {
        let record = record?;
        let kind = record.kind().get();
        collector.feed_record(kind, record.payload())?;
        let raw = RawRecord {
            start: start
                .checked_add(record.offset())
                .ok_or_else(|| Error::InvalidData("comment record offset overflow".into()))?,
            end: start
                .checked_add(record.offset())
                .and_then(|value| value.checked_add(4 + record.payload().len()))
                .ok_or_else(|| Error::InvalidData("comment record extent overflow".into()))?,
        };
        if kind == DBCELL {
            last_dbcell_end = raw.end;
        }
        if let Some(pending) = pending_text.as_mut() {
            if kind != CONTINUE_TYPE {
                return Err(Error::InvalidData(
                    "comment TXO family is not contiguous".into(),
                ));
            }
            let site = sites
                .get_mut(&pending.object_id)
                .ok_or_else(|| Error::InvalidData("pending TXO has no comment site".into()))?;
            site.txo_family.push(raw.clone());
            if pending.characters_remaining != 0 {
                let (&flags, bytes) = record
                    .payload()
                    .split_first()
                    .ok_or_else(|| Error::InvalidData("empty comment text CONTINUE".into()))?;
                if flags & !1 != 0 {
                    return Err(Error::InvalidData(
                        "comment text CONTINUE has reserved encoding bits".into(),
                    ));
                }
                let width = if flags & 1 == 0 { 1 } else { 2 };
                if bytes.len() % width != 0 {
                    return Err(Error::InvalidData(
                        "comment text CONTINUE has a partial UTF-16 unit".into(),
                    ));
                }
                let count = bytes.len() / width;
                if count == 0 || count > pending.characters_remaining {
                    return Err(Error::InvalidData(
                        "comment text CONTINUE exceeds TXO cchText".into(),
                    ));
                }
                pending.characters_remaining -= count;
            } else {
                if record.payload().is_empty() || record.payload().len() > pending.runs_remaining {
                    return Err(Error::InvalidData(
                        "comment run CONTINUE exceeds TXO cbRuns".into(),
                    ));
                }
                site.source_run_bytes.extend_from_slice(record.payload());
                pending.runs_remaining -= record.payload().len();
            }
            if pending.characters_remaining == 0 && pending.runs_remaining == 0 {
                pending_text = None;
            }
            continue;
        }
        match kind {
            OBJ_TYPE => {
                let Some(object_id) = comment_object_id(record.payload())? else {
                    pending_object = None;
                    continue;
                };
                if sites.contains_key(&object_id) {
                    return Err(Error::UnsafeEdit(format!(
                        "duplicate comment object id {object_id}"
                    )));
                }
                sites.insert(
                    object_id,
                    PartialSite {
                        object: raw.clone(),
                        drawing: None,
                        txo_family: Vec::new(),
                        note: None,
                        source_run_bytes: Vec::new(),
                    },
                );
                pending_object = Some(object_id);
            },
            MSODRAWING_TYPE => {
                let boundary = record.payload() == [0, 0, 0x0D, 0xF0, 0, 0, 0, 0];
                let Some(object_id) = pending_object else {
                    if boundary {
                        return Err(Error::UnsafeEdit(
                            "orphan or duplicate comment ClientTextbox drawing boundary".into(),
                        ));
                    }
                    continue;
                };
                if !boundary {
                    return Err(Error::UnsafeEdit(
                        "comment object has an ambiguous ClientTextbox drawing boundary".into(),
                    ));
                }
                let site = sites
                    .get_mut(&object_id)
                    .ok_or_else(|| Error::InvalidData("comment site disappeared".into()))?;
                if site.drawing.is_some() {
                    return Err(Error::UnsafeEdit(
                        "comment object has duplicate ClientTextbox drawing boundaries".into(),
                    ));
                }
                site.drawing = Some(raw.clone());
            },
            TXO_TYPE => {
                let Some(object_id) = pending_object.take() else {
                    return Err(Error::UnsafeEdit(
                        "orphan or duplicate comment TXO record".into(),
                    ));
                };
                let payload = record.payload();
                if payload.len() < 18 {
                    return Err(Error::InvalidData("comment TXO is truncated".into()));
                }
                let characters = usize::from(binary::read_u16_le_at(payload, 10)?);
                let runs = usize::from(binary::read_u16_le_at(payload, 12)?);
                let site = sites
                    .get_mut(&object_id)
                    .ok_or_else(|| Error::InvalidData("comment site disappeared".into()))?;
                if site.drawing.is_none() {
                    return Err(Error::UnsafeEdit(
                        "comment TXO has no ClientTextbox drawing boundary".into(),
                    ));
                }
                if !site.txo_family.is_empty() {
                    return Err(Error::UnsafeEdit(
                        "comment object has duplicate TXO records".into(),
                    ));
                }
                site.txo_family.push(raw.clone());
                if characters != 0 || runs != 0 {
                    pending_text = Some(PendingText {
                        object_id,
                        characters_remaining: characters,
                        runs_remaining: runs,
                    });
                }
            },
            RECORD_TYPE => {
                if record.payload().len() < 8 {
                    return Err(Error::InvalidData("NOTE is truncated".into()));
                }
                let object_id = binary::read_u16_le_at(record.payload(), 6)?;
                let site = sites.get_mut(&object_id).ok_or_else(|| {
                    Error::UnsafeEdit(format!(
                        "NOTE {object_id} has no uniquely located OfficeArt object"
                    ))
                })?;
                if site.note.replace(raw.clone()).is_some() {
                    return Err(Error::UnsafeEdit(format!(
                        "comment object {object_id} has duplicate NOTE records"
                    )));
                }
            },
            EOF => {
                found_eof = true;
                break;
            },
            _ => {},
        }
    }
    if !found_eof || pending_object.is_some() || pending_text.is_some() {
        return Err(Error::InvalidData(
            "worksheet ended with an incomplete comment sequence".into(),
        ));
    }
    let linked = collector.finish()?;
    if linked != comments {
        return Err(Error::UnsafeEdit(
            "source comment linkage disagrees with the complete workbook reader".into(),
        ));
    }
    let mut entries = Vec::new();
    entries
        .try_reserve_exact(comments.len())
        .map_err(|_error| Error::Allocation("retaining comment transaction inventory"))?;
    for comment in comments {
        let object_id = comment.identity().object_id();
        let site = sites.remove(&object_id).ok_or_else(|| {
            Error::UnsafeEdit(format!("comment object {object_id} has no source site"))
        })?;
        let drawing = site.drawing.ok_or_else(|| {
            Error::UnsafeEdit(format!("comment object {object_id} has no ClientTextbox"))
        })?;
        let note = site
            .note
            .ok_or_else(|| Error::UnsafeEdit(format!("comment object {object_id} has no NOTE")))?;
        if site.object.start < last_dbcell_end
            || drawing.start < last_dbcell_end
            || note.start < last_dbcell_end
            || site
                .txo_family
                .iter()
                .any(|record| record.start < last_dbcell_end)
        {
            return Err(Error::UnsafeEdit(
                "comment records precede the worksheet's final DBCELL and would require INDEX regeneration"
                    .into(),
            ));
        }
        let retained = Site {
            note,
            txo_family: site.txo_family,
            source_run_bytes: site.source_run_bytes,
        };
        entries.push(Entry {
            comment: comment.clone(),
            site: retained,
        });
    }
    if !sites.is_empty() {
        return Err(Error::UnsafeEdit(
            "worksheet contains unlinked comment drawing sites".into(),
        ));
    }
    Ok(entries)
}

fn require_bof(payload: &[u8], expected: u16) -> Result<()> {
    if payload.len() < 4 || binary::read_u16_le_at(payload, 0)? != 0x0600 {
        return Err(Error::UnsupportedFeature(
            "comment transactions require a BIFF8 Workbook stream".into(),
        ));
    }
    if binary::read_u16_le_at(payload, 2)? != expected {
        return Err(Error::InvalidData(
            "unexpected BIFF BOF substream type".into(),
        ));
    }
    Ok(())
}

fn comment_object_id(payload: &[u8]) -> Result<Option<u16>> {
    if payload.len() < 22 {
        return Ok(None);
    }
    if binary::read_u16_le_at(payload, 4)? != COMMENT_OBJECT_TYPE {
        return Ok(None);
    }
    let object_id = binary::read_u16_le_at(payload, 6)?;
    if object_id == 0 {
        return Err(Error::InvalidData("comment OBJ id must not be zero".into()));
    }
    Ok(Some(object_id))
}

fn encode_note(comment: &Comment, value: &Value) -> Result<Vec<u8>> {
    let units: Vec<_> = value.author.encode_utf16().collect();
    let compressed = units.iter().all(|unit| *unit <= 0x00FF);
    encode_note_with_compression(comment, value, compressed)
}

fn encode_note_with_compression(
    comment: &Comment,
    value: &Value,
    compressed: bool,
) -> Result<Vec<u8>> {
    let units: Vec<_> = value.author.encode_utf16().collect();
    let mut payload = Vec::new();
    payload.extend_from_slice(&comment.row().to_le_bytes());
    payload.extend_from_slice(&u16::from(comment.column()).to_le_bytes());
    let metadata = comment.note_metadata();
    let mut flags = metadata.reserved_flags();
    if comment.row_hidden() {
        flags |= 0x0080;
    }
    if comment.column_hidden() {
        flags |= 0x0100;
    }
    if comment.visibility() == Visibility::Visible {
        flags |= 0x0002;
    }
    payload.extend_from_slice(&flags.to_le_bytes());
    payload.extend_from_slice(&comment.identity().object_id().to_le_bytes());
    payload.extend_from_slice(
        &u16::try_from(units.len())
            .map_err(|_error| Error::InvalidData("comment author exceeds u16".into()))?
            .to_le_bytes(),
    );
    payload.push(metadata.reserved_string_flags() | u8::from(!compressed));
    if compressed {
        for unit in units {
            payload.push(u8::try_from(unit).map_err(|_error| {
                Error::InvalidData("compressed comment author unit exceeds u8".into())
            })?);
        }
    } else {
        for unit in units {
            payload.extend_from_slice(&unit.to_le_bytes());
        }
    }
    payload.push(metadata.unused_byte());
    encode_record(RECORD_TYPE, &payload)
}

fn encode_txo_family(
    stream: &[u8],
    entry: &Entry,
    text: &str,
    compression: Option<bool>,
) -> Result<Vec<u8>> {
    let txo = entry
        .site
        .txo_family
        .first()
        .ok_or_else(|| Error::InvalidData("comment TXO family is empty".into()))?;
    let mut payload = stream
        .get(txo.start + 4..txo.end)
        .ok_or_else(|| Error::InvalidData("comment TXO payload is truncated".into()))?
        .to_vec();
    if payload.len() < 18 {
        return Err(Error::InvalidData(
            "comment TXO payload is truncated".into(),
        ));
    }
    let units: Vec<_> = text.encode_utf16().collect();
    let run_bytes = retain_runs(&entry.site.source_run_bytes, units.len())?;
    payload[10..12].copy_from_slice(
        &u16::try_from(units.len())
            .map_err(|_error| Error::InvalidData("comment text exceeds u16".into()))?
            .to_le_bytes(),
    );
    payload[12..14].copy_from_slice(
        &u16::try_from(run_bytes.len())
            .map_err(|_error| Error::InvalidData("comment run bytes exceed u16".into()))?
            .to_le_bytes(),
    );
    let mut output = encode_record(TXO_TYPE, &payload)?;
    if !units.is_empty() {
        let compressed = compression.unwrap_or_else(|| units.iter().all(|unit| *unit <= 0x00FF));
        let per_record = if compressed { 8_223 } else { 4_111 };
        let mut offset = 0;
        while offset < units.len() {
            let mut end = (offset + per_record).min(units.len());
            if !compressed && end < units.len() && (0xD800..=0xDBFF).contains(&units[end - 1]) {
                end -= 1;
            }
            let mut segment = Vec::new();
            segment.push(u8::from(!compressed));
            if compressed {
                for unit in &units[offset..end] {
                    segment.push(u8::try_from(*unit).map_err(|_error| {
                        Error::InvalidData("compressed comment text unit exceeds u8".into())
                    })?);
                }
            } else {
                for unit in &units[offset..end] {
                    segment.extend_from_slice(&unit.to_le_bytes());
                }
            }
            output.extend_from_slice(&encode_record(CONTINUE_TYPE, &segment)?);
            offset = end;
        }
        for chunk in run_bytes.chunks(MAX_RECORD_BYTES) {
            output.extend_from_slice(&encode_record(CONTINUE_TYPE, chunk)?);
        }
    }
    Ok(output)
}

fn retain_runs(source: &[u8], character_count: usize) -> Result<Vec<u8>> {
    if character_count == 0 {
        return Ok(Vec::new());
    }
    if source.len() < 16 || !source.len().is_multiple_of(8) {
        return Err(Error::UnsafeEdit(
            "comment source has an unsupported formatting-run layout".into(),
        ));
    }
    let terminal = source.len() - 8;
    let mut output = Vec::new();
    for run in source[..terminal].chunks_exact(8) {
        let index = usize::from(u16::from_le_bytes([run[0], run[1]]));
        if index <= character_count {
            output.extend_from_slice(run);
        }
    }
    if output.is_empty() || output[0..2] != [0, 0] {
        return Err(Error::UnsafeEdit(
            "comment formatting runs do not begin at character zero".into(),
        ));
    }
    let mut last = source[terminal..].to_vec();
    last[0..2].copy_from_slice(
        &u16::try_from(character_count)
            .map_err(|_error| Error::InvalidData("comment text exceeds u16".into()))?
            .to_le_bytes(),
    );
    output.extend_from_slice(&last);
    if output.len() > usize::from(u16::MAX) {
        return Err(Error::UnsafeEdit(
            "comment formatting runs exceed the BIFF8 cbRuns bound".into(),
        ));
    }
    Ok(output)
}

fn encode_record(kind: u16, payload: &[u8]) -> Result<Vec<u8>> {
    if payload.len() > MAX_RECORD_BYTES {
        return Err(Error::InvalidData(format!(
            "BIFF8 comment record 0x{kind:04X} exceeds {MAX_RECORD_BYTES} bytes"
        )));
    }
    let mut record = Vec::new();
    record.extend_from_slice(&kind.to_le_bytes());
    record.extend_from_slice(
        &u16::try_from(payload.len())
            .map_err(|_error| Error::InvalidData("BIFF8 record length exceeds u16".into()))?
            .to_le_bytes(),
    );
    record.extend_from_slice(payload);
    Ok(record)
}

fn ensure_disjoint(replacements: &[(usize, usize, Vec<u8>)]) -> Result<()> {
    for pair in replacements.windows(2) {
        let later = &pair[0];
        let earlier = &pair[1];
        if earlier.1 > later.0 {
            return Err(Error::UnsafeEdit(
                "comment replacement ranges overlap".into(),
            ));
        }
    }
    Ok(())
}

fn verify_readback(target: &Snapshot, source: &Snapshot, changes: &[&Change]) -> Result<()> {
    if target.inner.sheets.len() != source.inner.sheets.len() {
        return Err(Error::UnsafeEdit(
            "comment publication changed the worksheet inventory".into(),
        ));
    }
    let expected: BTreeMap<_, _> = changes
        .iter()
        .map(|change| {
            let sheet = &source.inner.sheets[change.sheet];
            let comment = &sheet.entries[change.entry].comment;
            (
                (sheet.workbook_position, comment.row(), comment.column()),
                &change.after,
            )
        })
        .collect();
    for sheet in &target.inner.sheets {
        for entry in &sheet.entries {
            let key = (
                sheet.workbook_position,
                entry.comment.row(),
                entry.comment.column(),
            );
            if let Some(value) = expected.get(&key) {
                if &Value::from_comment(&entry.comment) != *value {
                    return Err(Error::UnsafeEdit(format!(
                        "comment publication readback failed at sheet {} cell ({}, {})",
                        sheet.workbook_position,
                        entry.comment.row(),
                        entry.comment.column()
                    )));
                }
            }
        }
    }
    for key in expected.keys() {
        if !target.inner.sheets.iter().any(|sheet| {
            sheet.workbook_position == key.0
                && sheet
                    .entries
                    .iter()
                    .any(|entry| entry.comment.row() == key.1 && entry.comment.column() == key.2)
        }) {
            return Err(Error::UnsafeEdit(
                "comment publication removed a targeted NOTE".into(),
            ));
        }
    }
    Ok(())
}
