//! Lossless source-backed worksheet visibility edits for BIFF8 XLS packages.
//!
//! This owner changes only the one-byte `hsState` field of existing
//! `BoundSheet8` worksheet entries. It does not move sheet substreams, rewrite
//! workbook views, or reinterpret macro/chart/module tabs. Every publication
//! reopens the complete CFB and XLS workbook and retains an exact-source,
//! reversible artifact patch.

#![cfg_attr(
    not(test),
    deny(
        clippy::expect_used,
        clippy::panic,
        clippy::todo,
        clippy::unimplemented,
        clippy::unwrap_used,
        reason = "source-backed visibility edits return typed refusals"
    )
)]

use crate::records::{BoundSheetRecord, Encoding, SheetType, SheetVisible};
use crate::{Error, Result, SheetKind, SheetVisibility, Workbook};
use litchi_biff::Records;
use litchi_cfb::{
    ArtifactFingerprint, ComposedOverlaySource, OleFile, OverlayError, OverlayLimits,
    PublishReport, SameLengthStreamOverlay, ValidatedOverlayPlan,
};
use litchi_core::binary;
use litchi_core::{ReadAt, SourceVersion};
use litchi_ole_common::object::{Editor as PackageEditor, Limits, Targets};
use litchi_ole_common::source_backed_overlay::SourceBackedOverlayPublisher;
use std::fmt;
use std::io::{Cursor, Read, Seek, SeekFrom, Write};
use std::path::Path;
use std::sync::Arc;

const BOF: u16 = 0x0809;
const EOF: u16 = 0x000a;
const CODE_PAGE: u16 = 0x0042;
const BOUND_SHEET: u16 = 0x0085;
const FILE_PASS: u16 = 0x002f;
const BIFF8: u16 = 0x0600;
const WORKBOOK_GLOBALS: u16 = 0x0005;
const PATCH_MAGIC: &[u8; 8] = b"LXSV0001";

/// Maximum distinct worksheet visibility owners in one atomic publication.
pub const MAX_VISIBILITY_CHANGES: usize = 64;

#[derive(Debug, Clone)]
struct SheetEntry {
    position: usize,
    name: String,
    visibility: SheetVisibility,
    kind: SheetKind,
    visibility_offset: usize,
    substream_offset: u32,
}

struct Inner {
    bytes: Arc<[u8]>,
    workbook_path: Vec<String>,
    workbook_stream: Arc<[u8]>,
    sheets: Vec<SheetEntry>,
}

/// Immutable, cheaply cloned source-backed XLS visibility snapshot.
#[derive(Clone)]
pub struct Snapshot {
    inner: Arc<Inner>,
}

impl Snapshot {
    /// Opens an unencrypted, unsigned and unprotected BIFF8 XLS package.
    ///
    /// # Errors
    ///
    /// Returns a typed CFB, BIFF, protection, duplicate-owner, allocation, or
    /// complete-workbook validation error before publishing a snapshot.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = PackageEditor::open(bytes, Targets::default(), Limits::default())?;
        Self::from_package_editor(package)
    }

    fn from_package_editor(package: PackageEditor) -> Result<Self> {
        let workbook_path = select_workbook_path(&package)?;
        let workbook_stream = package
            .stream_shared(&workbook_path)
            .ok_or_else(|| Error::InvalidData("selected XLS Workbook stream disappeared".into()))?;
        let sheets = parse_directory(&workbook_stream)?;
        require_visible_worksheet(&sheets)?;
        let bytes = package.finish()?;
        let workbook = Workbook::new(Cursor::new(bytes.as_slice()))?;
        require_unprotected(&workbook)?;
        require_public_readback(&workbook, &sheets)?;
        Ok(Self {
            inner: Arc::new(Inner {
                bytes: Arc::from(bytes),
                workbook_path,
                workbook_stream,
                sheets,
            }),
        })
    }

    /// Exact source CFB artifact bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.inner.bytes
    }

    /// Exact source Workbook stream.
    #[must_use]
    pub fn workbook_stream(&self) -> &[u8] {
        &self.inner.workbook_stream
    }

    /// Number of worksheet owners (chart, macro, and module tabs excluded).
    #[must_use]
    pub fn worksheet_count(&self) -> usize {
        self.inner
            .sheets
            .iter()
            .filter(|sheet| sheet.kind == SheetKind::WorksheetOrDialog)
            .count()
    }

    /// Iterates worksheet owners in workbook tab order.
    pub fn worksheets(&self) -> impl Iterator<Item = Worksheet<'_>> {
        self.inner
            .sheets
            .iter()
            .enumerate()
            .filter_map(|(index, sheet)| {
                (sheet.kind == SheetKind::WorksheetOrDialog).then_some(Worksheet {
                    snapshot: self,
                    index,
                })
            })
    }

    /// Resolves one worksheet by case-folded name or exact tab position.
    ///
    /// # Errors
    ///
    /// Returns an ambiguity error for duplicate case-folded worksheet names.
    /// Positions selecting chart, macro, or module tabs return `None`.
    pub fn worksheet<'a>(&'a self, selector: Selector<'_>) -> Result<Option<Worksheet<'a>>> {
        Ok(self.resolve(selector)?.map(|index| Worksheet {
            snapshot: self,
            index,
        }))
    }

    /// Starts a detached bounded visibility transaction.
    #[must_use]
    pub fn transaction(&self) -> Transaction {
        Transaction {
            source: self.clone(),
            changes: Vec::new(),
        }
    }

    /// Alias for [`Self::transaction`].
    #[must_use]
    pub fn edit(&self) -> Transaction {
        self.transaction()
    }

    fn resolve(&self, selector: Selector<'_>) -> Result<Option<usize>> {
        match selector {
            Selector::Position(position) => Ok(self.inner.sheets.get(position).and_then(|sheet| {
                (sheet.kind == SheetKind::WorksheetOrDialog).then_some(position)
            })),
            Selector::Name(name) => {
                let mut matches = self.inner.sheets.iter().enumerate().filter(|(_, sheet)| {
                    sheet.kind == SheetKind::WorksheetOrDialog && caseless_eq(&sheet.name, name)
                });
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

impl fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("artifact_bytes", &self.inner.bytes.len())
            .field("workbook_bytes", &self.inner.workbook_stream.len())
            .field("worksheet_count", &self.worksheet_count())
            .finish()
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.bytes() == other.bytes()
    }
}

impl Eq for Snapshot {}

/// Borrowed worksheet visibility owner.
#[derive(Debug, Clone, Copy)]
pub struct Worksheet<'a> {
    snapshot: &'a Snapshot,
    index: usize,
}

impl<'a> Worksheet<'a> {
    fn entry(self) -> &'a SheetEntry {
        &self.snapshot.inner.sheets[self.index]
    }

    /// Zero-based workbook tab position.
    #[must_use]
    pub fn position(self) -> usize {
        self.entry().position
    }

    /// Developer-visible worksheet name.
    #[must_use]
    pub fn name(self) -> &'a str {
        &self.entry().name
    }

    /// Exact BIFF8 visible, hidden, or very-hidden state.
    #[must_use]
    pub fn visibility(self) -> SheetVisibility {
        self.entry().visibility
    }
}

/// Semantic selector for an existing worksheet owner.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Selector<'a> {
    /// Case-folded developer-visible tab name.
    Name(&'a str),
    /// Exact zero-based workbook tab position.
    Position(usize),
}

impl<'a> From<&'a str> for Selector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Name(value)
    }
}

impl From<usize> for Selector<'_> {
    fn from(value: usize) -> Self {
        Self::Position(value)
    }
}

#[derive(Debug, Clone, Copy)]
struct Change {
    sheet: usize,
    before: SheetVisibility,
    after: SheetVisibility,
}

/// Detached, bounded and failure-atomic worksheet visibility transaction.
#[derive(Clone)]
pub struct Transaction {
    source: Snapshot,
    changes: Vec<Change>,
}

impl Transaction {
    /// Stages one exact visible, hidden, or very-hidden outcome.
    ///
    /// # Errors
    ///
    /// Returns a missing/ambiguous/non-worksheet selector or finite-bound
    /// refusal without changing the source artifact.
    pub fn set_visibility(
        &mut self,
        selector: Selector<'_>,
        visibility: SheetVisibility,
    ) -> Result<&mut Self> {
        let sheet = self
            .source
            .resolve(selector)?
            .ok_or_else(|| Error::WorksheetNotFound("visibility edit worksheet selector".into()))?;
        if let Some(change) = self.changes.iter_mut().find(|change| change.sheet == sheet) {
            change.after = visibility;
            if change.before == change.after {
                let sheet = change.sheet;
                self.changes.retain(|candidate| candidate.sheet != sheet);
            }
            return Ok(self);
        }
        let before = self.source.inner.sheets[sheet].visibility;
        if before == visibility {
            return Ok(self);
        }
        if self.changes.len() >= MAX_VISIBILITY_CHANGES {
            return Err(Error::UnsafeEdit(format!(
                "worksheet visibility transaction exceeds its {MAX_VISIBILITY_CHANGES}-owner limit"
            )));
        }
        self.changes
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("staging worksheet visibility changes"))?;
        self.changes.push(Change {
            sheet,
            before,
            after: visibility,
        });
        Ok(self)
    }

    /// Stages an ordinary hidden outcome.
    pub fn hide(&mut self, selector: Selector<'_>) -> Result<&mut Self> {
        self.set_visibility(selector, SheetVisibility::Hidden)
    }

    /// Stages a very-hidden outcome.
    pub fn very_hide(&mut self, selector: Selector<'_>) -> Result<&mut Self> {
        self.set_visibility(selector, SheetVisibility::VeryHidden)
    }

    /// Stages a visible outcome.
    pub fn unhide(&mut self, selector: Selector<'_>) -> Result<&mut Self> {
        self.set_visibility(selector, SheetVisibility::Visible)
    }

    /// Stages a whole caller-provided batch or leaves this transaction intact.
    ///
    /// # Errors
    ///
    /// Returns the first selector, allocation, or bound error atomically.
    pub fn set_visibility_batch<'a, I>(&mut self, changes: I) -> Result<&mut Self>
    where
        I: IntoIterator<Item = (Selector<'a>, SheetVisibility)>,
    {
        let mut candidate = self.clone();
        for (selector, visibility) in changes {
            candidate.set_visibility(selector, visibility)?;
        }
        *self = candidate;
        Ok(self)
    }

    /// Reopens and publishes the complete package with an exact inverse patch.
    ///
    /// # Errors
    ///
    /// Returns a final-visible-sheet, stale-field, CFB, complete-reopen, or
    /// semantic-closure error without publishing partial bytes.
    pub fn commit(self) -> Result<Commit> {
        if self.changes.is_empty() {
            let patch = Patch::new(
                Arc::clone(&self.source.inner.bytes),
                Arc::clone(&self.source.inner.bytes),
            );
            return Ok(Commit {
                snapshot: self.source,
                patch,
                diagnostics: Diagnostics::default(),
            });
        }
        require_final_visible_worksheet(&self.source, &self.changes)?;
        let mut workbook = self.source.inner.workbook_stream.to_vec();
        for change in &self.changes {
            let entry =
                self.source.inner.sheets.get(change.sheet).ok_or_else(|| {
                    Error::UnsafeEdit("worksheet visibility owner is stale".into())
                })?;
            let field = workbook.get_mut(entry.visibility_offset).ok_or_else(|| {
                Error::InvalidData("BoundSheet8 visibility field is outside Workbook".into())
            })?;
            if *field != encode_visibility(change.before) {
                return Err(Error::UnsafeEdit(
                    "BoundSheet8 visibility precondition is stale".into(),
                ));
            }
            *field = encode_visibility(change.after);
        }
        let mut package = PackageEditor::open(
            self.source.inner.bytes.to_vec(),
            Targets::default(),
            Limits::default(),
        )?;
        package.put_stream_shared(&self.source.inner.workbook_path, Arc::from(workbook))?;
        let snapshot = Snapshot::from_package_editor(package)?;
        for change in &self.changes {
            if snapshot
                .inner
                .sheets
                .get(change.sheet)
                .map(|entry| entry.visibility)
                != Some(change.after)
            {
                return Err(Error::UnsafeEdit(
                    "worksheet visibility failed semantic readback".into(),
                ));
            }
        }
        certify_workbook_transition(&self.source, &snapshot)?;
        let patch = Patch::new(
            Arc::clone(&self.source.inner.bytes),
            Arc::clone(&snapshot.inner.bytes),
        );
        Ok(Commit {
            snapshot,
            patch,
            diagnostics: Diagnostics {
                changed_worksheets: self.changes.len(),
                touched_streams: 1,
            },
        })
    }

    /// Plans a protected, source-backed same-length visibility publication.
    ///
    /// The complete `Workbook`/`Book` stream is staged first and then submitted
    /// as one [`SameLengthStreamOverlay`]. The common CFB publisher retains the
    /// source topology and all unselected streams, while the composed target is
    /// reopened through the complete XLS reader before the plan is returned.
    /// This method never falls back to [`Self::commit`]; callers that need a
    /// length-changing or topology-changing edit must explicitly choose that
    /// eager path.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for a protected/signed/encrypted source, stale
    /// `BoundSheet8` state, an invalid final-visible-sheet outcome, a changed
    /// stream length, failed complete semantic readback, or a bounded CFB
    /// overlay failure.
    pub fn commit_source_backed(self) -> Result<SourceBackedCommit> {
        let effective: Vec<_> = self
            .changes
            .iter()
            .filter(|change| change.before != change.after)
            .collect();

        if !effective.is_empty() {
            require_final_visible_worksheet(&self.source, &self.changes)?;
        }

        let mut workbook = Vec::new();
        workbook
            .try_reserve_exact(self.source.inner.workbook_stream.len())
            .map_err(|_error| Error::Allocation("staging source-backed Workbook stream"))?;
        workbook.extend_from_slice(&self.source.inner.workbook_stream);
        for change in &effective {
            let entry = self.source.inner.sheets.get(change.sheet).ok_or_else(|| {
                Error::UnsafeEdit("source-backed worksheet visibility owner is stale".into())
            })?;
            let field = workbook.get_mut(entry.visibility_offset).ok_or_else(|| {
                Error::InvalidData("BoundSheet8 visibility field is outside Workbook".into())
            })?;
            if *field != encode_visibility(change.before) {
                return Err(Error::UnsafeEdit(
                    "BoundSheet8 visibility precondition is stale".into(),
                ));
            }
            *field = encode_visibility(change.after);
        }
        if workbook.len() != self.source.inner.workbook_stream.len() {
            return Err(Error::UnsafeEdit(
                "source-backed visibility publication changed Workbook stream length".into(),
            ));
        }

        let source: Arc<dyn ReadAt> =
            Arc::new(SnapshotSource::new(Arc::clone(&self.source.inner.bytes)));
        let publisher = SourceBackedOverlayPublisher::open(source).map_err(overlay_to_error)?;
        let replacement = SameLengthStreamOverlay::new(
            self.source.inner.workbook_path.clone(),
            Arc::from(workbook),
        );
        let plan = publisher
            .plan(vec![replacement], OverlayLimits::default())
            .map_err(overlay_to_error)?;
        verify_source_backed_readback(&plan, &self.source, &effective)?;

        let source_workbook_bytes = u64::try_from(self.source.inner.workbook_stream.len())
            .map_err(|_error| Error::InvalidData("Workbook stream length exceeds u64".into()))?;
        let source_bytes = u64::try_from(self.source.inner.bytes.len())
            .map_err(|_error| Error::InvalidData("source CFB length exceeds u64".into()))?;
        let diagnostics = SourceBackedDiagnostics {
            changed_worksheets: effective.len(),
            touched_streams: usize::from(!plan.is_noop()),
            changed_spans: plan.changed_spans(),
            source_bytes,
            source_workbook_bytes,
            target_workbook_bytes: source_workbook_bytes,
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

impl fmt::Debug for Transaction {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Transaction")
            .field("source", &self.source)
            .field("staged_changes", &self.changes.len())
            .finish()
    }
}

/// Content-free visibility publication diagnostics.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Diagnostics {
    changed_worksheets: usize,
    touched_streams: usize,
}

impl Diagnostics {
    /// Number of distinct changed worksheet owners.
    #[must_use]
    pub const fn changed_worksheets(self) -> usize {
        self.changed_worksheets
    }

    /// Number of changed logical CFB streams.
    #[must_use]
    pub const fn touched_streams(self) -> usize {
        self.touched_streams
    }
}

/// Successful complete-package visibility publication.
#[derive(Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Reopened target snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Exact-source reversible artifact patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Content-free publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Splits this publication into its components.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch, Diagnostics) {
        (self.snapshot, self.patch, self.diagnostics)
    }
}

impl fmt::Debug for Commit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Commit")
            .field("snapshot", &self.snapshot)
            .field("patch", &self.patch)
            .field("diagnostics", &self.diagnostics)
            .finish()
    }
}

/// Content-free evidence for a same-length source-backed visibility
/// publication.
///
/// The diagnostics retain only bounded counts, lengths, and opaque CFB
/// fingerprints. Workbook payload bytes and worksheet names are not copied
/// into this value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourceBackedDiagnostics {
    changed_worksheets: usize,
    touched_streams: usize,
    changed_spans: usize,
    source_bytes: u64,
    source_workbook_bytes: u64,
    target_workbook_bytes: u64,
    source_fingerprint: ArtifactFingerprint,
    target_fingerprint: ArtifactFingerprint,
}

impl SourceBackedDiagnostics {
    /// Number of distinct worksheet owners changed by the transaction.
    #[must_use]
    pub const fn changed_worksheets(self) -> usize {
        self.changed_worksheets
    }

    /// Number of logical CFB streams selected for the overlay.
    #[must_use]
    pub const fn touched_streams(self) -> usize {
        self.touched_streams
    }

    /// Number of physical CFB spans retained by the overlay plan.
    #[must_use]
    pub const fn changed_spans(self) -> usize {
        self.changed_spans
    }

    /// Complete source CFB length.
    #[must_use]
    pub const fn source_bytes(self) -> u64 {
        self.source_bytes
    }

    /// Source `Workbook`/`Book` stream length.
    #[must_use]
    pub const fn source_workbook_bytes(self) -> u64 {
        self.source_workbook_bytes
    }

    /// Candidate `Workbook`/`Book` stream length.
    #[must_use]
    pub const fn target_workbook_bytes(self) -> u64 {
        self.target_workbook_bytes
    }

    /// Exact source CFB fingerprint checked by the overlay plan.
    #[must_use]
    pub const fn source_fingerprint(self) -> ArtifactFingerprint {
        self.source_fingerprint
    }

    /// Exact composed target CFB fingerprint checked by the overlay plan.
    #[must_use]
    pub const fn target_fingerprint(self) -> ArtifactFingerprint {
        self.target_fingerprint
    }
}

/// A checked source-bound same-length XLS visibility publication plan.
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
    /// A sink failure may leave a typed prefix in the sink; inspect the
    /// returned [`OverlayError`] before deciding whether to retry.
    pub fn write_to<W: Write>(&self, writer: &mut W) -> Result<PublishReport, OverlayError> {
        self.plan.write_to(writer)
    }

    /// Alias for [`Self::write_to`] emphasizing sequential publication.
    pub fn publish_to_stream<W: Write>(
        &self,
        writer: &mut W,
    ) -> Result<PublishReport, OverlayError> {
        self.write_to(writer)
    }

    /// Publishes through the common synced sibling-file and atomic-rename
    /// path.
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

/// Exact-source reversible replacement of one XLS visibility artifact.
#[derive(Clone)]
pub struct Patch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl Patch {
    fn new(before: Arc<[u8]>, after: Arc<[u8]>) -> Self {
        Self { before, after }
    }

    /// Parses the bounded stable `LXSV0001` artifact envelope and certifies
    /// that its only logical change is worksheet `BoundSheet8.hsState`.
    ///
    /// # Errors
    ///
    /// Returns a wire, package, protection, or semantic-closure error.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        if bytes.get(..PATCH_MAGIC.len()) != Some(PATCH_MAGIC) {
            return Err(Error::InvalidData(
                "worksheet visibility patch has the wrong magic".into(),
            ));
        }
        let before_len = read_patch_len(bytes, 8)?;
        let after_len = read_patch_len(bytes, 16)?;
        let body_start = 24_usize;
        let before_end = body_start
            .checked_add(before_len)
            .ok_or_else(|| Error::InvalidData("visibility patch length overflow".into()))?;
        let after_end = before_end
            .checked_add(after_len)
            .ok_or_else(|| Error::InvalidData("visibility patch length overflow".into()))?;
        if after_end != bytes.len() {
            return Err(Error::InvalidData(
                "worksheet visibility patch length is not canonical".into(),
            ));
        }
        let limits = Limits::default();
        let max = usize::try_from(limits.max_total_size)
            .map_err(|_error| Error::InvalidData("package limit exceeds usize".into()))?;
        if before_len > max || after_len > max {
            return Err(Error::UnsafeEdit(
                "worksheet visibility patch artifact exceeds package bounds".into(),
            ));
        }
        let before = Snapshot::from_bytes(bytes[body_start..before_end].to_vec())?;
        let after = Snapshot::from_bytes(bytes[before_end..after_end].to_vec())?;
        certify_workbook_transition(&before, &after)?;
        compare_other_streams(&before, &after)?;
        Ok(Self::new(
            Arc::clone(&before.inner.bytes),
            Arc::clone(&after.inner.bytes),
        ))
    }

    /// Serializes the stable bounded artifact envelope.
    ///
    /// # Errors
    ///
    /// Returns an allocation or platform-length error.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let before_len = u64::try_from(self.before.len())
            .map_err(|_error| Error::InvalidData("source artifact exceeds u64".into()))?;
        let after_len = u64::try_from(self.after.len())
            .map_err(|_error| Error::InvalidData("target artifact exceeds u64".into()))?;
        let capacity = 24_usize
            .checked_add(self.before.len())
            .and_then(|value| value.checked_add(self.after.len()))
            .ok_or_else(|| Error::InvalidData("visibility patch length overflow".into()))?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(capacity)
            .map_err(|_error| Error::Allocation("serializing worksheet visibility patch"))?;
        bytes.extend_from_slice(PATCH_MAGIC);
        bytes.extend_from_slice(&before_len.to_le_bytes());
        bytes.extend_from_slice(&after_len.to_le_bytes());
        bytes.extend_from_slice(&self.before);
        bytes.extend_from_slice(&self.after);
        Ok(bytes)
    }

    /// Exact source artifact required by this patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Exact target artifact produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Whether the exact artifact transition is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before == self.after
    }

    /// Applies only to the exact retained immutable source artifact.
    ///
    /// # Errors
    ///
    /// Returns a stale-source or complete-reopen validation error.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.bytes() != self.before() {
            return Err(Error::UnsafeEdit(
                "worksheet visibility patch source does not match its base snapshot".into(),
            ));
        }
        if self.is_empty() {
            return Ok(source.clone());
        }
        let target = Snapshot::from_bytes(self.after.to_vec())?;
        certify_workbook_transition(source, &target)?;
        Ok(target)
    }

    /// Exact durable inverse artifact transition.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("before_bytes", &self.before.len())
            .field("after_bytes", &self.after.len())
            .field("is_empty", &self.is_empty())
            .finish()
    }
}

fn select_workbook_path(package: &PackageEditor) -> Result<Vec<String>> {
    let workbook = vec!["Workbook".to_string()];
    let book = vec!["Book".to_string()];
    match (
        package.stream(&workbook).is_some(),
        package.stream(&book).is_some(),
    ) {
        (true, false) => Ok(workbook),
        (false, true) => Ok(book),
        (false, false) => Err(Error::InvalidData(
            "XLS package has no Workbook or Book stream".into(),
        )),
        (true, true) => Err(Error::UnsafeEdit(
            "XLS package has ambiguous Workbook and Book streams".into(),
        )),
    }
}

fn parse_directory(workbook: &[u8]) -> Result<Vec<SheetEntry>> {
    let mut records = Records::new(workbook);
    let first = records.next().ok_or(Error::Eof("Workbook globals BOF"))??;
    require_bof(first.payload())?;
    let mut encoding = Encoding::from_codepage(1252)?;
    let mut sheets = Vec::new();
    let mut found_eof = false;
    for record_result in records {
        let record = record_result?;
        match record.kind().get() {
            CODE_PAGE if record.payload().len() == 2 => {
                encoding = Encoding::from_codepage(binary::read_u16_le_at(record.payload(), 0)?)?;
            },
            BOUND_SHEET => {
                let raw_visibility =
                    record
                        .payload()
                        .get(4)
                        .copied()
                        .ok_or_else(|| Error::InvalidRecord {
                            record_type: BOUND_SHEET,
                            message: "BoundSheet8 is truncated before hsState".into(),
                        })?;
                if raw_visibility > 2 {
                    return Err(Error::InvalidRecord {
                        record_type: BOUND_SHEET,
                        message: format!(
                            "BoundSheet8 hsState must be exactly 0, 1, or 2; found 0x{raw_visibility:02X}"
                        ),
                    });
                }
                let bound = BoundSheetRecord::parse(record.payload(), &encoding)?;
                let kind = match bound.sheet_type {
                    SheetType::WorkSheet => SheetKind::WorksheetOrDialog,
                    SheetType::MacroSheet => SheetKind::MacroSheet,
                    SheetType::ChartSheet => SheetKind::ChartSheet,
                    SheetType::VBModule => SheetKind::VbaModule,
                };
                let visibility = decode_visibility(bound.visible);
                let visibility_offset = record
                    .offset()
                    .checked_add(8)
                    .ok_or_else(|| Error::InvalidData("BoundSheet8 offset overflow".into()))?;
                sheets
                    .try_reserve(1)
                    .map_err(|_error| Error::Allocation("indexing BoundSheet8 owners"))?;
                sheets.push(SheetEntry {
                    position: sheets.len(),
                    name: bound.name,
                    visibility,
                    kind,
                    visibility_offset,
                    substream_offset: bound.position,
                });
            },
            FILE_PASS => return Err(Error::PasswordRequired),
            EOF => {
                found_eof = true;
                break;
            },
            _ => {},
        }
    }
    if !found_eof {
        return Err(Error::Eof("Workbook globals EOF"));
    }
    if sheets.is_empty() {
        return Err(Error::UnsafeEdit(
            "worksheet visibility owner requires a BoundSheet8 directory".into(),
        ));
    }
    let mut positions = std::collections::HashSet::new();
    for sheet in &sheets {
        if !positions.insert(sheet.substream_offset) {
            return Err(Error::UnsafeEdit(
                "duplicate BoundSheet8 substream owner".into(),
            ));
        }
        let offset = usize::try_from(sheet.substream_offset)
            .map_err(|_error| Error::InvalidData("sheet substream offset exceeds usize".into()))?;
        if offset >= workbook.len() {
            return Err(Error::InvalidRecord {
                record_type: BOUND_SHEET,
                message: "BoundSheet8 points outside the Workbook stream".into(),
            });
        }
    }
    Ok(sheets)
}

fn require_bof(payload: &[u8]) -> Result<()> {
    if payload.len() < 4 {
        return Err(Error::InvalidLength {
            expected: 4,
            found: payload.len(),
        });
    }
    let version = binary::read_u16_le_at(payload, 0)?;
    let substream = binary::read_u16_le_at(payload, 2)?;
    if version != BIFF8 || substream != WORKBOOK_GLOBALS {
        return Err(Error::InvalidRecord {
            record_type: BOF,
            message: "visibility edits require a BIFF8 workbook-globals BOF".into(),
        });
    }
    Ok(())
}

fn require_unprotected<R: Read + Seek>(workbook: &Workbook<R>) -> Result<()> {
    let protection = workbook.protection();
    if protection.structure_protected()
        || protection.windows_protected()
        || protection.password().is_set()
        || protection.revisions_protected()
        || protection.revision_password().is_set()
        || protection.write_protected()
        || protection.file_sharing().is_some()
    {
        return Err(Error::UnsafeEdit(
            "protected or shared workbooks are not eligible for visibility edits".into(),
        ));
    }
    for metadata in workbook.sheets() {
        let Some(index) = metadata.parsed_worksheet_index() else {
            continue;
        };
        let protection = workbook.xls_worksheet(index)?.protection();
        if protection.is_protected()
            || protection.objects_protected()
            || protection.scenarios_protected()
            || protection.has_password()
        {
            return Err(Error::UnsafeEdit(
                "protected worksheets are not eligible for visibility edits".into(),
            ));
        }
    }
    Ok(())
}

fn require_public_readback<R: Read + Seek>(
    workbook: &Workbook<R>,
    sheets: &[SheetEntry],
) -> Result<()> {
    if workbook.sheets().len() != sheets.len() {
        return Err(Error::UnsafeEdit(
            "BoundSheet8 directory is incomplete on public readback".into(),
        ));
    }
    for (metadata, sheet) in workbook.sheets().iter().zip(sheets) {
        if metadata.workbook_index() != sheet.position
            || metadata.name() != sheet.name
            || metadata.kind() != sheet.kind
            || metadata.visibility() != sheet.visibility
        {
            return Err(Error::UnsafeEdit(
                "BoundSheet8 owner failed exact public readback".into(),
            ));
        }
        if sheet.kind == SheetKind::WorksheetOrDialog && metadata.parsed_worksheet_index().is_none()
        {
            return Err(Error::UnsafeEdit(format!(
                "worksheet at tab position {} was not published by the complete XLS reader",
                sheet.position
            )));
        }
    }
    Ok(())
}

fn require_visible_worksheet(sheets: &[SheetEntry]) -> Result<()> {
    if sheets.iter().any(|sheet| {
        sheet.kind == SheetKind::WorksheetOrDialog && sheet.visibility == SheetVisibility::Visible
    }) {
        Ok(())
    } else {
        Err(Error::UnsafeEdit(
            "an XLS workbook must retain at least one visible worksheet".into(),
        ))
    }
}

fn require_final_visible_worksheet(source: &Snapshot, changes: &[Change]) -> Result<()> {
    if source
        .inner
        .sheets
        .iter()
        .enumerate()
        .any(|(index, sheet)| {
            sheet.kind == SheetKind::WorksheetOrDialog
                && changes
                    .iter()
                    .find(|change| change.sheet == index)
                    .map_or(sheet.visibility, |change| change.after)
                    == SheetVisibility::Visible
        })
    {
        Ok(())
    } else {
        Err(Error::UnsafeEdit(
            "visibility batch would hide every worksheet".into(),
        ))
    }
}

fn certify_workbook_transition(before: &Snapshot, after: &Snapshot) -> Result<()> {
    if before.inner.workbook_path != after.inner.workbook_path
        || before.inner.sheets.len() != after.inner.sheets.len()
        || before.workbook_stream().len() != after.workbook_stream().len()
    {
        return Err(Error::UnsafeEdit(
            "visibility patch changes the Workbook owner inventory".into(),
        ));
    }
    let mut allowed = Vec::new();
    for (left, right) in before.inner.sheets.iter().zip(&after.inner.sheets) {
        if left.position != right.position
            || left.name != right.name
            || left.kind != right.kind
            || left.visibility_offset != right.visibility_offset
            || left.substream_offset != right.substream_offset
        {
            return Err(Error::UnsafeEdit(
                "visibility patch changes a BoundSheet8 owner identity".into(),
            ));
        }
        if left.visibility != right.visibility {
            if left.kind != SheetKind::WorksheetOrDialog {
                return Err(Error::UnsafeEdit(
                    "visibility patch targets a non-worksheet owner".into(),
                ));
            }
            allowed
                .try_reserve(1)
                .map_err(|_error| Error::Allocation("certifying visibility patch fields"))?;
            allowed.push(left.visibility_offset);
        }
        if before.workbook_stream().get(left.visibility_offset)
            != Some(&encode_visibility(left.visibility))
            || after.workbook_stream().get(right.visibility_offset)
                != Some(&encode_visibility(right.visibility))
        {
            return Err(Error::UnsafeEdit(
                "visibility patch contains a noncanonical BoundSheet8.hsState byte".into(),
            ));
        }
    }
    if allowed.len() > MAX_VISIBILITY_CHANGES {
        return Err(Error::UnsafeEdit(
            "visibility patch exceeds its worksheet-owner limit".into(),
        ));
    }
    allowed.sort_unstable();
    let mut allowed_index = 0_usize;
    for (offset, (left, right)) in before
        .workbook_stream()
        .iter()
        .zip(after.workbook_stream())
        .enumerate()
    {
        if left == right {
            continue;
        }
        while allowed
            .get(allowed_index)
            .is_some_and(|field| *field < offset)
        {
            allowed_index += 1;
        }
        if allowed.get(allowed_index) != Some(&offset) {
            return Err(Error::UnsafeEdit(
                "visibility patch changes Workbook bytes outside BoundSheet8.hsState".into(),
            ));
        }
    }
    Ok(())
}

fn compare_other_streams(before: &Snapshot, after: &Snapshot) -> Result<()> {
    let mut left = OleFile::open(Cursor::new(before.bytes()))?;
    let mut right = OleFile::open(Cursor::new(after.bytes()))?;
    if left.sector_size() != right.sector_size()
        || root_identity(&left) != root_identity(&right)
        || directory_catalog(&left)? != directory_catalog(&right)?
    {
        return Err(Error::UnsafeEdit(
            "visibility patch changes the CFB directory or root metadata".into(),
        ));
    }
    let mut left_paths = left.list_streams();
    let mut right_paths = right.list_streams();
    left_paths.sort();
    right_paths.sort();
    if left_paths != right_paths {
        return Err(Error::UnsafeEdit(
            "visibility patch changes the CFB stream catalog".into(),
        ));
    }
    for path in left_paths {
        if path == before.inner.workbook_path {
            continue;
        }
        let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
        if left.open_stream(&refs)? != right.open_stream(&refs)? {
            return Err(Error::UnsafeEdit(
                "visibility patch changes an unrelated CFB stream".into(),
            ));
        }
    }
    Ok(())
}

#[derive(Clone)]
struct SnapshotSource {
    bytes: Arc<[u8]>,
    version: SourceVersion,
}

impl SnapshotSource {
    fn new(bytes: Arc<[u8]>) -> Self {
        let identity = bytes.as_ptr() as usize as u64;
        let length = bytes.len() as u64;
        Self {
            bytes,
            version: SourceVersion::new(identity, length),
        }
    }
}

impl ReadAt for SnapshotSource {
    fn len(&self) -> std::io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_error| std::io::Error::other("source length exceeds u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
        if output.is_empty() || offset >= self.bytes.len() as u64 {
            return Ok(0);
        }
        let start = usize::try_from(offset)
            .map_err(|_error| std::io::Error::other("source offset exceeds usize"))?;
        let count = output.len().min(self.bytes.len() - start);
        output[..count].copy_from_slice(&self.bytes[start..start + count]);
        Ok(count)
    }

    fn version(&self) -> std::io::Result<SourceVersion> {
        Ok(self.version)
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
        other => Error::UnsafeEdit(format!("source-backed visibility overlay refused: {other}")),
    }
}

fn verify_source_backed_readback(
    plan: &ValidatedOverlayPlan,
    source: &Snapshot,
    changes: &[&Change],
) -> Result<()> {
    let candidate = plan.composed_source().map_err(overlay_to_error)?;
    let workbook = Workbook::new(PositionalReader::new(candidate))?;
    require_unprotected(&workbook)?;

    let mut sheets = Vec::new();
    sheets
        .try_reserve_exact(source.inner.sheets.len())
        .map_err(|_error| Error::Allocation("retaining source-backed sheet readback"))?;
    sheets.extend(source.inner.sheets.iter().cloned());
    for change in changes {
        let sheet = sheets.get_mut(change.sheet).ok_or_else(|| {
            Error::UnsafeEdit("source-backed worksheet disappeared during readback".into())
        })?;
        sheet.visibility = change.after;
    }
    require_public_readback(&workbook, &sheets)?;
    for change in changes {
        let metadata = workbook.sheets().get(change.sheet).ok_or_else(|| {
            Error::UnsafeEdit("source-backed worksheet disappeared during readback".into())
        })?;
        if metadata.visibility() != change.after {
            return Err(Error::UnsafeEdit(
                "source-backed worksheet visibility readback disagreed with the staged state"
                    .into(),
            ));
        }
    }
    Ok(())
}

fn root_identity(ole: &OleFile<Cursor<&[u8]>>) -> Option<(String, u8, String)> {
    ole.root_entry()
        .map(|entry| (entry.name.clone(), entry.entry_type, entry.clsid.clone()))
}

fn directory_catalog(ole: &OleFile<Cursor<&[u8]>>) -> Result<Vec<(Vec<String>, u8, String)>> {
    let limits = Limits::default();
    let maximum = limits
        .max_streams
        .saturating_add(limits.max_objects.saturating_mul(limits.max_storage_depth));
    let mut catalog = Vec::new();
    let mut pending = vec![Vec::<String>::new()];
    while let Some(parent) = pending.pop() {
        let refs = parent.iter().map(String::as_str).collect::<Vec<_>>();
        for entry in ole.list_directory_entries(&refs)? {
            if catalog.len() >= maximum {
                return Err(Error::UnsafeEdit(
                    "CFB directory catalog exceeds visibility patch bounds".into(),
                ));
            }
            let mut path = parent.clone();
            path.try_reserve(1)
                .map_err(|_error| Error::Allocation("certifying CFB directory paths"))?;
            path.push(entry.name.clone());
            catalog
                .try_reserve(1)
                .map_err(|_error| Error::Allocation("certifying CFB directory entries"))?;
            catalog.push((path.clone(), entry.entry_type, entry.clsid.clone()));
            if entry.entry_type == 0x01 {
                pending
                    .try_reserve(1)
                    .map_err(|_error| Error::Allocation("walking CFB storage catalog"))?;
                pending.push(path);
            }
        }
    }
    catalog.sort();
    Ok(catalog)
}

fn encode_visibility(visibility: SheetVisibility) -> u8 {
    match visibility {
        SheetVisibility::Visible => 0,
        SheetVisibility::Hidden => 1,
        SheetVisibility::VeryHidden => 2,
    }
}

fn decode_visibility(visibility: SheetVisible) -> SheetVisibility {
    match visibility {
        SheetVisible::Visible => SheetVisibility::Visible,
        SheetVisible::Hidden => SheetVisibility::Hidden,
        SheetVisible::VeryHidden => SheetVisibility::VeryHidden,
    }
}

fn caseless_eq(left: &str, right: &str) -> bool {
    left.chars()
        .flat_map(char::to_lowercase)
        .eq(right.chars().flat_map(char::to_lowercase))
}

fn read_patch_len(bytes: &[u8], offset: usize) -> Result<usize> {
    let field = bytes.get(offset..offset + 8).ok_or_else(|| {
        Error::InvalidData("worksheet visibility patch header is truncated".into())
    })?;
    usize::try_from(u64::from_le_bytes([
        field[0], field[1], field[2], field[3], field[4], field[5], field[6], field[7],
    ]))
    .map_err(|_error| Error::InvalidData("visibility patch length exceeds usize".into()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_cfb::{OleFile, OleWriter};
    use std::io;

    fn package(sheet_count: usize) -> Vec<u8> {
        let mut writer = crate::Writer::new();
        for index in 0..sheet_count {
            let sheet = writer.add_worksheet(&format!("Sheet{index}")).unwrap();
            writer.write_number(sheet, 0, 0, index as f64).unwrap();
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        let mut package =
            PackageEditor::open(output.into_inner(), Targets::default(), Limits::default())
                .unwrap();
        package
            .add_stream(vec!["Opaque".to_string()], b"untouched".to_vec())
            .unwrap();
        package.finish().unwrap()
    }

    fn stream(bytes: &[u8], path: &[&str]) -> Vec<u8> {
        OleFile::open(Cursor::new(bytes))
            .unwrap()
            .open_stream(path)
            .unwrap()
    }

    fn edit_workbook(source: &[u8], edit: impl FnOnce(&mut Vec<u8>)) -> Vec<u8> {
        let mut package =
            PackageEditor::open(source.to_vec(), Targets::default(), Limits::default()).unwrap();
        let path = vec!["Workbook".to_string()];
        let mut workbook = package.stream(&path).unwrap().to_vec();
        edit(&mut workbook);
        package.put_stream(&path, workbook).unwrap();
        package.finish().unwrap()
    }

    fn signed(source: &[u8]) -> Vec<u8> {
        rebuild(source, |writer| {
            writer
                .create_stream(&["DigitalSignature"], b"signature")
                .unwrap();
        })
    }

    fn rebuild(source: &[u8], configure: impl FnOnce(&mut OleWriter)) -> Vec<u8> {
        let mut ole = OleFile::open(Cursor::new(source)).unwrap();
        let mut writer = OleWriter::new();
        for path in ole.list_streams() {
            let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
            writer
                .create_stream(&refs, &ole.open_stream(&refs).unwrap())
                .unwrap();
        }
        configure(&mut writer);
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    #[test]
    fn hides_unhides_and_very_hides_exact_fixed_width_owners() {
        let source = Snapshot::from_bytes(package(3)).unwrap();
        let before = source.workbook_stream().to_vec();
        let mut transaction = source.transaction();
        transaction.hide("sHeEt1".into()).unwrap();
        transaction.very_hide(2usize.into()).unwrap();
        let commit = transaction.commit().unwrap();
        assert_eq!(commit.diagnostics().changed_worksheets(), 2);
        assert_eq!(commit.diagnostics().touched_streams(), 1);
        assert_eq!(
            commit
                .snapshot()
                .worksheet("Sheet1".into())
                .unwrap()
                .unwrap()
                .visibility(),
            SheetVisibility::Hidden
        );
        assert_eq!(
            commit
                .snapshot()
                .worksheet(2usize.into())
                .unwrap()
                .unwrap()
                .visibility(),
            SheetVisibility::VeryHidden
        );
        assert_eq!(before.len(), commit.snapshot().workbook_stream().len());
        assert_eq!(
            before
                .iter()
                .zip(commit.snapshot().workbook_stream())
                .filter(|(left, right)| left != right)
                .count(),
            2
        );
        assert_eq!(stream(source.bytes(), &["Opaque"]), b"untouched");
        assert_eq!(stream(commit.snapshot().bytes(), &["Opaque"]), b"untouched");

        let mut unhide = commit.snapshot().transaction();
        unhide.unhide("SHEET1".into()).unwrap();
        let unhidden = unhide.commit().unwrap();
        assert_eq!(
            unhidden
                .snapshot()
                .worksheet(1usize.into())
                .unwrap()
                .unwrap()
                .visibility(),
            SheetVisibility::Visible
        );
    }

    #[test]
    fn source_backed_visibility_publishes_one_and_bounded_batches() {
        let source = Snapshot::from_bytes(package(3)).unwrap();
        let source_directory = {
            let ole = OleFile::open(Cursor::new(source.bytes())).unwrap();
            directory_catalog(&ole).unwrap()
        };
        let mut edit = source.transaction();
        edit.hide(1usize.into()).unwrap();
        let commit = edit.commit_source_backed().unwrap();
        assert_eq!(commit.diagnostics().changed_worksheets(), 1);
        assert_eq!(commit.diagnostics().touched_streams(), 1);
        assert!(commit.diagnostics().changed_spans() > 0);
        assert_ne!(
            commit.diagnostics().source_fingerprint(),
            commit.diagnostics().target_fingerprint()
        );

        let mut output = Vec::new();
        let report = commit.write_to(&mut output).unwrap();
        assert_eq!(report.changed_spans(), commit.diagnostics().changed_spans());
        let reopened = Snapshot::from_bytes(output.clone()).unwrap();
        assert_eq!(
            reopened
                .worksheet(1usize.into())
                .unwrap()
                .unwrap()
                .visibility(),
            SheetVisibility::Hidden
        );
        assert_eq!(stream(&output, &["Opaque"]), b"untouched");
        let target_directory = {
            let ole = OleFile::open(Cursor::new(output.as_slice())).unwrap();
            directory_catalog(&ole).unwrap()
        };
        assert_eq!(source_directory, target_directory);

        let changed_offsets = source
            .workbook_stream()
            .iter()
            .zip(reopened.workbook_stream())
            .enumerate()
            .filter_map(|(offset, (before, after))| (before != after).then_some(offset))
            .collect::<Vec<_>>();
        assert_eq!(
            changed_offsets,
            vec![source.inner.sheets[1].visibility_offset]
        );

        let source = Snapshot::from_bytes(package(MAX_VISIBILITY_CHANGES + 2)).unwrap();
        let mut bounded = source.transaction();
        bounded
            .set_visibility_batch(
                (1..=MAX_VISIBILITY_CHANGES)
                    .map(|position| (Selector::Position(position), SheetVisibility::Hidden)),
            )
            .unwrap();
        let bounded = bounded.commit_source_backed().unwrap();
        assert_eq!(
            bounded.diagnostics().changed_worksheets(),
            MAX_VISIBILITY_CHANGES
        );
        let mut output = Vec::new();
        bounded.write_to(&mut output).unwrap();
        let reopened = Snapshot::from_bytes(output).unwrap();
        assert_eq!(
            reopened
                .worksheets()
                .filter(|sheet| sheet.visibility() == SheetVisibility::Visible)
                .count(),
            2
        );
    }

    #[test]
    fn source_backed_visibility_noop_and_cap_plus_one_are_atomic() {
        let source = Snapshot::from_bytes(package(2)).unwrap();
        let noop = source.transaction().commit_source_backed().unwrap();
        assert!(noop.is_noop());
        assert_eq!(noop.diagnostics().changed_worksheets(), 0);
        assert_eq!(noop.diagnostics().touched_streams(), 0);
        assert_eq!(
            noop.diagnostics().source_fingerprint(),
            noop.diagnostics().target_fingerprint()
        );
        let mut output = Vec::new();
        noop.write_to(&mut output).unwrap();
        assert_eq!(output, source.bytes());

        let source = Snapshot::from_bytes(package(MAX_VISIBILITY_CHANGES + 2)).unwrap();
        let mut refused = source.transaction();
        assert!(
            refused
                .set_visibility_batch(
                    (0..=MAX_VISIBILITY_CHANGES).map(|position| {
                        (Selector::Position(position), SheetVisibility::Hidden)
                    })
                )
                .is_err()
        );
        let noop = refused.commit_source_backed().unwrap();
        assert!(noop.is_noop());
        assert_eq!(noop.source().bytes(), source.bytes());
    }

    struct PartialSink {
        accepted: usize,
        limit: usize,
    }

    impl Write for PartialSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            if self.accepted >= self.limit {
                return Err(io::Error::new(
                    io::ErrorKind::BrokenPipe,
                    "test sink stopped",
                ));
            }
            let count = bytes.len().min(self.limit - self.accepted);
            self.accepted += count;
            Ok(count)
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn source_backed_visibility_reports_partial_sink_and_save_failures() {
        let source = Snapshot::from_bytes(package(2)).unwrap();
        let mut edit = source.transaction();
        edit.hide(1usize.into()).unwrap();
        let commit = edit.commit_source_backed().unwrap();
        let mut sink = PartialSink {
            accepted: 0,
            limit: 23,
        };
        let error = commit.write_to(&mut sink).unwrap_err();
        assert!(matches!(
            error,
            OverlayError::IncompleteOutput {
                progress: litchi_cfb::OutputProgress::Prefix { .. },
                ..
            }
        ));

        let path = std::env::temp_dir().join(format!(
            "litchi-xls-visibility-missing-parent-{}",
            std::process::id()
        ));
        let missing = path.join("child").join("output.xls");
        let error = commit.save(&missing).unwrap_err();
        assert!(matches!(
            error,
            OverlayError::Ole(litchi_cfb::OleError::Io(_))
                | OverlayError::Io(_)
                | OverlayError::Unavailable { .. }
        ));
    }

    #[test]
    fn refuses_hiding_every_worksheet_atomically() {
        let source = Snapshot::from_bytes(package(2)).unwrap();
        let mut transaction = source.transaction();
        transaction.hide(0usize.into()).unwrap();
        transaction.hide(1usize.into()).unwrap();
        assert!(transaction.commit().is_err());
        assert_eq!(
            source
                .worksheets()
                .filter(|sheet| sheet.visibility() == SheetVisibility::Visible)
                .count(),
            2
        );
    }

    #[test]
    fn batch_accepts_cap_and_refuses_cap_plus_one_without_staging_prefix() {
        let source = Snapshot::from_bytes(package(MAX_VISIBILITY_CHANGES + 2)).unwrap();
        let mut transaction = source.transaction();
        transaction
            .set_visibility_batch(
                (1..=MAX_VISIBILITY_CHANGES)
                    .map(|position| (Selector::Position(position), SheetVisibility::Hidden)),
            )
            .unwrap();
        let commit = transaction.commit().unwrap();
        assert_eq!(
            commit.diagnostics().changed_worksheets(),
            MAX_VISIBILITY_CHANGES
        );

        let mut refused = source.transaction();
        assert!(
            refused
                .set_visibility_batch(
                    (0..=MAX_VISIBILITY_CHANGES).map(|position| {
                        (Selector::Position(position), SheetVisibility::Hidden)
                    })
                )
                .is_err()
        );
        let noop = refused.commit().unwrap();
        assert!(noop.patch().is_empty());
        assert_eq!(noop.snapshot().bytes(), source.bytes());
    }

    #[test]
    fn exact_patch_rejects_stale_and_durable_inverse_restores_source() {
        let source = Snapshot::from_bytes(package(3)).unwrap();
        let mut transaction = source.transaction();
        transaction.hide(2usize.into()).unwrap();
        let commit = transaction.commit().unwrap();
        let wire = commit.patch().to_bytes().unwrap();
        let durable = Patch::from_bytes(&wire).unwrap();
        assert!(durable.apply(commit.snapshot()).is_err());
        assert_eq!(
            durable.apply(&source).unwrap().bytes(),
            commit.snapshot().bytes()
        );
        let inverse_wire = durable.inverse().to_bytes().unwrap();
        let inverse = Patch::from_bytes(&inverse_wire).unwrap();
        assert_eq!(
            inverse.apply(commit.snapshot()).unwrap().bytes(),
            source.bytes()
        );
    }

    #[test]
    fn selector_rules_and_nonworksheet_positions_are_exact() {
        let bytes = package(2);
        let source = Snapshot::from_bytes(bytes.clone()).unwrap();
        assert_eq!(
            source
                .worksheet("sHeEt0".into())
                .unwrap()
                .unwrap()
                .position(),
            0
        );
        assert!(source.worksheet(99usize.into()).unwrap().is_none());

        let chart_typed = edit_workbook(&bytes, |workbook| {
            let offset = Records::new(workbook)
                .filter_map(std::result::Result::ok)
                .filter(|record| record.kind().get() == BOUND_SHEET)
                .nth(1)
                .unwrap()
                .offset();
            workbook[offset + 9] = 0x02;
        });
        let chart_typed = Snapshot::from_bytes(chart_typed).unwrap();
        assert!(chart_typed.worksheet(1usize.into()).unwrap().is_none());
        assert!(chart_typed.transaction().hide(1usize.into()).is_err());
    }

    #[test]
    fn protected_signed_and_encrypted_sources_are_refused() {
        let bytes = package(2);
        assert!(Snapshot::from_bytes(signed(&bytes)).is_err());

        let protected = edit_workbook(&bytes, |workbook| {
            let offset = Records::new(workbook)
                .filter_map(std::result::Result::ok)
                .find(|record| record.kind().get() == crate::protection::PROTECT_TYPE)
                .unwrap()
                .offset();
            workbook[offset + 4..offset + 6].copy_from_slice(&1_u16.to_le_bytes());
        });
        assert!(Snapshot::from_bytes(protected).is_err());

        let encrypted = edit_workbook(&bytes, |workbook| {
            let offset = Records::new(workbook)
                .filter_map(std::result::Result::ok)
                .find(|record| record.kind().get() == CODE_PAGE)
                .unwrap()
                .offset();
            workbook[offset..offset + 2].copy_from_slice(&FILE_PASS.to_le_bytes());
        });
        assert!(matches!(
            Snapshot::from_bytes(encrypted),
            Err(Error::PasswordRequired)
        ));

        let reserved_visibility = edit_workbook(&bytes, |workbook| {
            let offset = Records::new(workbook)
                .filter_map(std::result::Result::ok)
                .find(|record| record.kind().get() == BOUND_SHEET)
                .unwrap()
                .offset();
            workbook[offset + 8] = 0x81;
        });
        assert!(Snapshot::from_bytes(reserved_visibility).is_err());
    }

    #[test]
    fn duplicate_case_folded_names_are_refused_or_reported_ambiguous() {
        let duplicate = edit_workbook(&package(2), |workbook| {
            let (offset, payload_len) = {
                let second = Records::new(workbook)
                    .filter_map(std::result::Result::ok)
                    .filter(|record| record.kind().get() == BOUND_SHEET)
                    .nth(1)
                    .unwrap();
                (second.offset(), second.payload().len())
            };
            let last_name_byte = offset + 4 + payload_len - 1;
            workbook[last_name_byte] = b'0';
        });
        match Snapshot::from_bytes(duplicate) {
            Ok(snapshot) => assert!(snapshot.worksheet("SHEET0".into()).is_err()),
            Err(error) => assert!(error.to_string().to_lowercase().contains("sheet")),
        }
    }

    #[test]
    fn deserialization_rejects_unrelated_stream_changes() {
        let source = Snapshot::from_bytes(package(2)).unwrap();
        let mut package = PackageEditor::open(
            source.bytes().to_vec(),
            Targets::default(),
            Limits::default(),
        )
        .unwrap();
        package
            .put_stream(&["Opaque".to_string()], b"changed".to_vec())
            .unwrap();
        let unrelated = Snapshot::from_bytes(package.finish().unwrap()).unwrap();
        let forged = Patch::new(
            Arc::clone(&source.inner.bytes),
            Arc::clone(&unrelated.inner.bytes),
        );
        assert!(Patch::from_bytes(&forged.to_bytes().unwrap()).is_err());
    }

    #[test]
    fn deserialization_binds_empty_storages_and_root_metadata() {
        let source = Snapshot::from_bytes(package(2)).unwrap();
        for changed in [
            rebuild(source.bytes(), |writer| {
                writer.create_storage(&["EmptyStorage"]).unwrap();
            }),
            rebuild(source.bytes(), |writer| writer.set_root_clsid([0x5a; 16])),
        ] {
            let changed = Snapshot::from_bytes(changed).unwrap();
            let forged = Patch::new(
                Arc::clone(&source.inner.bytes),
                Arc::clone(&changed.inner.bytes),
            );
            assert!(Patch::from_bytes(&forged.to_bytes().unwrap()).is_err());
        }
    }
}
