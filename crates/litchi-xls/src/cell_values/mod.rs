//! Lossless opened-workbook edits for BIFF8 cell records and sheets.
//!
//! The owner covers `Number` (`[MS-XLS]` 2.4.180), standalone and packed `RK`,
//! `BoolErr`, `Blank`, `LabelSst`, and non-string `Formula` caches. Every
//! fixed-width changes retain their source slots. Structural cell insertion,
//! removal (including packed `MulRk`/`MulBlank` members), row/column movement,
//! and sheet rename regenerate the complete
//! affected row-block/`INDEX`/`DBCELL`/dimension closure. Operations that meet
//! an unsupported formula, range, or drawing dependency are refused
//! before bytes are changed. The complete CFB package is reopened
//! before publication, and every other captured stream retains its exact
//! payload.

#![cfg_attr(
    not(test),
    deny(
        clippy::expect_used,
        clippy::panic,
        clippy::todo,
        clippy::unimplemented,
        clippy::unwrap_used,
        reason = "opened-workbook mutation paths must return typed refusals instead of terminating"
    )
)]

mod format;
mod layout;
mod structural;
mod tabs;

pub use format::FormatChange;

use crate::records::{BoundSheetRecord, Encoding, SheetType};
use crate::workbook::{KeptCells, ValidationWorkbook};
use crate::{Error, Result, SheetKind, raw::Reader};
use litchi_biff::Records;
use litchi_cfb::{
    ArtifactFingerprint, ComposedOverlaySource, OverlayError, OverlayOperationShape, PublishReport,
    SameLengthStreamSplice, StreamSpliceLimits, ValidatedOverlayPlan,
};
use litchi_cfb::{OleFileLimits, SharedOleFileLimits, consts::STGTY_STORAGE};
use litchi_core::binary;
use litchi_core::edit::RefusalKind;
pub use litchi_core::patch::HistoryLimits;
use litchi_core::patch::{
    BlobBundle, BlobLimits, DiagnosticFingerprint, Patch as CorePatch, PatchLimits, PatchOperation,
    Reversible, ReversibleOperation,
};
use litchi_core::sheet::{Cell as _, CellValue};
use litchi_core::{Position, ReadAt, SourceVersion};
/// The object and record limits a cell-value snapshot is read under.
pub use litchi_ole_common::object::Limits;
use litchi_ole_common::object::{OpenOptions, Package as ObjectPackage, Probe, Targets};
use litchi_ole_common::source_backed_overlay::SourceBackedOverlayPublisher;
use std::collections::BTreeMap;
use std::fmt;
use std::io::{Cursor, Read, Seek, SeekFrom, Write};
use std::sync::{Arc, OnceLock};

const BOF: u16 = 0x0809;
const EOF: u16 = 0x000a;
const CODE_PAGE: u16 = 0x0042;
const BOUND_SHEET: u16 = 0x0085;
const FILE_PASS: u16 = 0x002f;
const SST: u16 = 0x00fc;
const XF: u16 = 0x00e0;
const FORMULA: u16 = 0x0006;
const STRING: u16 = 0x0207;
const CONTINUE: u16 = 0x003c;
const MUL_RK: u16 = 0x00bd;
const MUL_BLANK: u16 = 0x00be;
const BLANK: u16 = 0x0201;
const NUMBER: u16 = 0x0203;
const BOOL_ERR: u16 = 0x0205;
const RK: u16 = 0x027e;
const LABEL_SST: u16 = 0x00fd;
const MAX_FORMULA_TOKEN_BYTES: usize = 8_202;
const BIFF8: u16 = 0x0600;
const WORKBOOK_GLOBALS: u16 = 0x0005;
const WORKSHEET: u16 = 0x0010;
const NUMBER_PAYLOAD_BYTES: usize = 14;
const NUMBER_VALUE_OFFSET: usize = 6;
const MAX_STAGED_CHANGES: usize = 4_096;
const MAX_SCALAR_TRANSFER_CELLS: usize = 4_096;
const MAX_SCALAR_TRANSFER_BYTES: usize = 4 * 1024 * 1024;

static SNAPSHOT_SOURCE_ID: OnceLock<std::sync::atomic::AtomicU64> = OnceLock::new();

fn fresh_snapshot_source_version() -> SourceVersion {
    let next = SNAPSHOT_SOURCE_ID
        .get_or_init(|| std::sync::atomic::AtomicU64::new(1))
        .fetch_add(1, std::sync::atomic::Ordering::Relaxed);
    SourceVersion::new(next, 0)
}

pub(crate) fn replace_workbook_ranges_and_adjust_bounds(
    workbook: &mut Vec<u8>,
    replacements: &[(usize, usize, Vec<u8>)],
) -> Result<()> {
    structural::replace_ranges_and_adjust_bounds(workbook, replacements)
}

/// A checked zero-based BIFF8 cell reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Reference {
    row: u16,
    column: u8,
}

impl Reference {
    /// Constructs a reference inside the 65,536 by 256 BIFF8 grid.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidCellReference`] when either coordinate is
    /// outside the BIFF8 grid.
    pub fn new(row: u32, column: u32) -> Result<Self> {
        let checked_row = u16::try_from(row).map_err(|_error| {
            Error::InvalidCellReference(format!(
                "cell row {row} is outside the BIFF8 worksheet grid"
            ))
        })?;
        let checked_column = u8::try_from(column).map_err(|_error| {
            Error::InvalidCellReference(format!(
                "cell column {column} is outside the BIFF8 worksheet grid"
            ))
        })?;
        Ok(Self {
            row: checked_row,
            column: checked_column,
        })
    }

    /// Returns the zero-based row.
    #[must_use]
    pub const fn row(self) -> u16 {
        self.row
    }

    /// Returns the zero-based column.
    #[must_use]
    pub const fn column(self) -> u8 {
        self.column
    }
}

/// An inclusive, checked BIFF8 worksheet rectangle.
///
/// Cross-workbook cell transfer deliberately uses an explicit rectangle
/// instead of accepting an unbounded iterator.  This keeps the selector
/// deterministic and lets the transaction preflight the complete destination
/// before it stages its first operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct CellRange {
    start: Reference,
    end: Reference,
}

impl CellRange {
    /// Creates an inclusive rectangle from its top-left and bottom-right
    /// references.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidCellReference`] when the end precedes the
    /// start on either axis.
    pub fn new(start: Reference, end: Reference) -> Result<Self> {
        if end.row() < start.row() || end.column() < start.column() {
            return Err(Error::InvalidCellReference(
                "cell range end precedes its start".into(),
            ));
        }
        Ok(Self { start, end })
    }

    /// Returns the inclusive top-left reference.
    #[must_use]
    pub const fn start(self) -> Reference {
        self.start
    }

    /// Returns the inclusive bottom-right reference.
    #[must_use]
    pub const fn end(self) -> Reference {
        self.end
    }

    /// Returns the number of cells in this rectangle, if representable.
    #[must_use]
    pub fn cell_count(self) -> Option<usize> {
        let rows = usize::from(self.end.row())
            .checked_sub(usize::from(self.start.row()))?
            .checked_add(1)?;
        let columns = usize::from(self.end.column())
            .checked_sub(usize::from(self.start.column()))?
            .checked_add(1)?;
        rows.checked_mul(columns)
    }

    fn contains(self, reference: Reference) -> bool {
        reference.row() >= self.start.row()
            && reference.row() <= self.end.row()
            && reference.column() >= self.start.column()
            && reference.column() <= self.end.column()
    }

    fn target_reference(self, source: Reference, anchor: Reference) -> Result<Reference> {
        let row_offset = u32::from(source.row())
            .checked_sub(u32::from(self.start.row()))
            .ok_or_else(|| Error::InvalidCellReference("source row precedes range start".into()))?;
        let column_offset = u32::from(source.column())
            .checked_sub(u32::from(self.start.column()))
            .ok_or_else(|| {
                Error::InvalidCellReference("source column precedes range start".into())
            })?;
        Reference::new(
            u32::from(anchor.row())
                .checked_add(row_offset)
                .ok_or_else(|| Error::InvalidCellReference("target row overflows BIFF8".into()))?,
            u32::from(anchor.column())
                .checked_add(column_offset)
                .ok_or_else(|| {
                    Error::InvalidCellReference("target column overflows BIFF8".into())
                })?,
        )
    }
}

/// A semantic worksheet selector.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Selector<'a> {
    /// Select a worksheet by case-insensitive developer-visible tab name.
    Name(&'a str),
    /// Select a worksheet by checked zero-based workbook tab position.
    Position(Position),
}

impl<'a> From<&'a str> for Selector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Name(value)
    }
}

impl From<Position> for Selector<'_> {
    fn from(value: Position) -> Self {
        Self::Position(value)
    }
}

impl From<usize> for Selector<'_> {
    fn from(value: usize) -> Self {
        Self::Position(Position::new(value))
    }
}

/// An existing BIFF8 `Number` record in Workbook-stream order.
#[derive(Debug, Clone, Copy)]
pub struct Number {
    reference: Reference,
    value: f64,
}

impl Number {
    /// Returns the cell location.
    #[must_use]
    pub const fn reference(self) -> Reference {
        self.reference
    }

    /// Returns the exact stored IEEE-754 value.
    #[must_use]
    pub const fn value(self) -> f64 {
        self.value
    }
}

impl PartialEq for Number {
    fn eq(&self, other: &Self) -> bool {
        self.reference == other.reference && self.value.to_bits() == other.value.to_bits()
    }
}

impl Eq for Number {}

/// BIFF8 cell storage family retained by the source workbook.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[non_exhaustive]
pub enum Storage {
    /// Standalone IEEE-754 `Number` record.
    Number,
    /// Standalone compressed numeric `RK` record.
    Rk,
    /// One compressed numeric value inside a `MulRk` record.
    MulRk,
    /// Boolean or error `BoolErr` record.
    BoolErr,
    /// Formatting-only `Blank` record or member of a `MulBlank` record.
    Blank,
    /// Shared-string reference `LabelSst` record.
    LabelSst,
    /// Inert formula with an editable non-string cached result.
    Formula,
}

/// A checked workbook-global BIFF8 extended-format resource index.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct StyleIndex(u16);

impl StyleIndex {
    /// Creates an index after checking it against an opened workbook.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] if the resource does not exist.
    pub fn new(snapshot: &Snapshot, index: u16) -> Result<Self> {
        if usize::from(index) >= snapshot.inner.xf_records.len() {
            return Err(Error::refused(
                RefusalKind::NotFound,
                "edit cells",
                "the BIFF8 XF index is outside the workbook resources",
            ));
        }
        Ok(Self(index))
    }

    /// Returns the workbook-global XF index.
    #[must_use]
    pub const fn get(self) -> u16 {
        self.0
    }
}

/// A checked BIFF8 error value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct CellError(u8);

impl CellError {
    /// Constructs one of the error codes defined for BIFF8 cached values.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidData`] for an unassigned error code.
    pub fn new(code: u8) -> Result<Self> {
        if matches!(code, 0x00 | 0x07 | 0x0f | 0x17 | 0x1d | 0x24 | 0x2a | 0x2b) {
            Ok(Self(code))
        } else {
            Err(Error::InvalidData(format!(
                "0x{code:02X} is not a defined BIFF8 cell error"
            )))
        }
    }

    /// Returns the BIFF8 error byte.
    #[must_use]
    pub const fn code(self) -> u8 {
        self.0
    }
}

/// An inert, non-string cached result stored in a `Formula` record.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum FormulaCache {
    /// IEEE-754 cached numeric result.
    Number(f64),
    /// Cached Boolean result.
    Boolean(bool),
    /// Cached spreadsheet error.
    Error(CellError),
    /// Empty cached result.
    Empty,
    /// Cached string result owned by the following `String`/`Continue` records.
    String(String),
}

/// A semantic value supported by the opened-workbook transaction.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Value {
    /// Numeric value stored by `Number`, `RK`, or `MulRk`.
    Number(f64),
    /// Boolean value stored by `BoolErr`.
    Boolean(bool),
    /// Spreadsheet error stored by `BoolErr`.
    Error(CellError),
    /// Formatting-only blank cell.
    Blank,
    /// Text resolved through an existing SST entry.
    Text(String),
    /// Inert cached result of an existing formula.
    FormulaCache(FormulaCache),
}

/// One existing BIFF8 cell owned by the bounded transaction.
#[derive(Debug, Clone, PartialEq)]
pub struct Cell {
    reference: Reference,
    storage: Storage,
    style: StyleIndex,
    value: Value,
}

impl Cell {
    /// Cell location.
    #[must_use]
    pub const fn reference(&self) -> Reference {
        self.reference
    }

    /// Exact source storage family.
    #[must_use]
    pub const fn storage(&self) -> Storage {
        self.storage
    }

    /// Workbook-global formatting resource used by this record.
    #[must_use]
    pub const fn style(&self) -> StyleIndex {
        self.style
    }

    /// Semantic value or inert formula cache.
    #[must_use]
    pub const fn value(&self) -> &Value {
        &self.value
    }
}

#[derive(Debug, Clone)]
struct Entry {
    cell: Cell,
    value_offset: Option<usize>,
    kind_offset: usize,
    sst_index: Option<u32>,
}

#[derive(Debug, Clone)]
struct SheetData {
    name: String,
    workbook_index: usize,
    entries: Arc<Vec<Entry>>,
    cell_index: Arc<CellIndex>,
}

/// Immutable coordinate index for one worksheet's source-order cell inventory.
///
/// `entries` intentionally remains in Workbook-stream order because that is
/// the public iteration order and the source-slot order used by the physical
/// patcher.  These sidecars are sorted only by coordinate, so lookups and
/// bounded row/column occupancy checks do not repeatedly scan that inventory.
#[derive(Debug, Clone)]
struct CellIndex {
    by_reference: Vec<IndexedCell>,
    by_column: Vec<IndexedCell>,
}

#[derive(Debug, Clone, Copy)]
struct IndexedCell {
    reference: Reference,
    entry: usize,
}

impl CellIndex {
    fn from_entries(entries: &[Entry]) -> Result<Self> {
        let mut by_reference = Vec::new();
        by_reference
            .try_reserve_exact(entries.len())
            .map_err(|_error| Error::Allocation("indexing worksheet cell coordinates"))?;
        let mut by_column = Vec::new();
        by_column
            .try_reserve_exact(entries.len())
            .map_err(|_error| Error::Allocation("indexing worksheet column coordinates"))?;
        for (entry, item) in entries.iter().enumerate() {
            let indexed = IndexedCell {
                reference: item.cell.reference,
                entry,
            };
            by_reference.push(indexed);
            by_column.push(indexed);
        }
        by_reference.sort_unstable_by_key(|item| (item.reference, item.entry));
        by_column.sort_unstable_by_key(|item| {
            (item.reference.column(), item.reference.row(), item.entry)
        });
        Ok(Self {
            by_reference,
            by_column,
        })
    }

    fn unique_entry_index(&self, entries: &[Entry], reference: Reference) -> Result<Option<usize>> {
        let start = self
            .by_reference
            .partition_point(|item| item.reference < reference);
        let Some(item) = self.by_reference.get(start) else {
            return Ok(None);
        };
        if item.reference != reference {
            return Ok(None);
        }
        let duplicate = self
            .by_reference
            .get(start + 1)
            .is_some_and(|candidate| candidate.reference == reference);
        if duplicate {
            return Err(Error::refused(
                RefusalKind::UnsupportedContent,
                "edit cells",
                "a cell has duplicate BIFF8 Number records",
            ));
        }
        let entry = entries.get(item.entry).ok_or(Error::PatchConflict {
            operation: "edit cells",
            reason: "worksheet coordinate index is stale",
        })?;
        if entry.cell.reference != reference {
            return Err(Error::Internal(
                "worksheet coordinate index disagrees with its source inventory",
            ));
        }
        Ok(Some(item.entry))
    }

    fn any_rows(&self, start: u16, end: u16) -> bool {
        let first = self.by_reference.partition_point(|item| {
            item.reference
                < Reference {
                    row: start,
                    column: 0,
                }
        });
        self.by_reference[first..]
            .iter()
            .take_while(|item| item.reference.row() <= end)
            .any(|item| item.reference.row() >= start)
    }

    fn any_columns(&self, start: u8, end: u8) -> bool {
        let first = self.by_column.partition_point(|item| {
            (item.reference.column(), item.reference.row()) < (start, 0_u16)
        });
        self.by_column[first..]
            .iter()
            .take_while(|item| item.reference.column() <= end)
            .any(|item| item.reference.column() >= start)
    }
}

impl SheetData {
    fn unique_entry_index(&self, reference: Reference) -> Result<Option<usize>> {
        self.cell_index.unique_entry_index(&self.entries, reference)
    }

    fn any_rows(&self, start: u16, end: u16) -> bool {
        self.cell_index.any_rows(start, end)
    }

    fn any_columns(&self, start: u8, end: u8) -> bool {
        self.cell_index.any_columns(start, end)
    }
}

#[derive(Clone, Copy)]
struct SourcePolicyFacts {
    public_worksheet_coverage: bool,
    protection: SourceProtectionPolicy,
    macro_free_workbook: bool,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum SourceProtectionPolicy {
    Unprotected,
    WorkbookOrShared,
    Worksheet,
}

impl SourcePolicyFacts {
    fn from_workbook<R: Read + Seek>(
        workbook: &ValidationWorkbook<R>,
        sheets: &[SheetData],
    ) -> Result<Self> {
        require_public_worksheet_coverage(workbook, sheets)?;
        Ok(Self {
            public_worksheet_coverage: true,
            protection: workbook_protection_policy(workbook)?,
            macro_free_workbook: workbook_is_macro_free(workbook),
        })
    }

    fn require(self) -> Result<()> {
        if !self.public_worksheet_coverage {
            return Err(Error::Internal(
                "source-backed numeric source lost public worksheet coverage",
            ));
        }
        match self.protection {
            SourceProtectionPolicy::Unprotected => {},
            SourceProtectionPolicy::WorkbookOrShared => {
                return Err(Error::refused(
                    RefusalKind::Protected,
                    "edit cells",
                    "protected or shared workbooks are not eligible for source-backed numeric edits",
                ));
            },
            SourceProtectionPolicy::Worksheet => {
                return Err(Error::refused(
                    RefusalKind::Protected,
                    "edit cells",
                    "protected worksheets are not eligible for source-backed numeric edits",
                ));
            },
        }
        if !self.macro_free_workbook {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "macro-bearing XLS sources are not eligible for source-backed numeric edits",
            ));
        }
        Ok(())
    }
}

/// A source-backed numeric commit's verification of its own materialized
/// target, deferred into the target's snapshot construction.
///
/// Every check below ran, in this order, inside
/// `verify_source_backed_numeric_target` after the target snapshot was built;
/// the only thing that changed is which complete `Reader` parse answers
/// them. The two parses read the same sealed bytes (`expected_bytes`, proved
/// identical by the first check) through the same `Reader::new` under the
/// same retained object and CFB limits.
#[derive(Clone, Copy)]
struct NumericTargetVerification<'a> {
    expected_bytes: &'a [u8],
    source: &'a Snapshot,
    changes: &'a [Change],
}

impl NumericTargetVerification<'_> {
    /// The cells the final readback reads: the only cells the target's
    /// validation-only open has to decode.
    fn kept_cells(self) -> Result<KeptCells> {
        numeric_readback_cells(self.source, self.changes)
    }

    fn run<R: Read + Seek>(
        self,
        workbook: &ValidationWorkbook<R>,
        target_bytes: &[u8],
        target_workbook_path: &[String],
        target_workbook_stream: &[u8],
    ) -> Result<()> {
        if target_bytes != self.expected_bytes {
            return Err(Error::PatchConflict {
                operation: "edit cells",
                reason: "source-backed numeric target changed during complete reopen",
            });
        }
        if self.source.inner.workbook_path != target_workbook_path {
            return Err(Error::Internal(
                "source-backed numeric publication changed the Workbook owner",
            ));
        }
        require_public_worksheet_coverage(workbook, &self.source.inner.sheets)?;
        require_unprotected_workbook(workbook)?;
        require_macro_free_workbook(workbook)?;
        let _ = carry_fixed_numeric_inventory(self.source, target_workbook_stream, self.changes)?;
        verify_public_numeric_readback(workbook, self.source, self.changes)
    }
}

struct Inner {
    // Keep the complete artifact in the same shared-vector shape accepted by
    // `OwnedSource::from_arc`. Source-backed publication can therefore move
    // this handle into its reopened CFB owner without copying the artifact.
    bytes: Arc<Vec<u8>>,
    source_version: SourceVersion,
    limits: Limits,
    cfb_limits: OleFileLimits,
    workbook_path: Vec<String>,
    workbook_stream: Arc<[u8]>,
    shared_strings: Arc<Vec<String>>,
    shared_string_properties: Arc<Vec<Option<Box<crate::records::SharedStringProperties>>>>,
    sst_total_offset: Option<usize>,
    xf_records: Arc<Vec<Vec<u8>>>,
    sheets: Vec<SheetData>,
    source_policy: SourcePolicyFacts,
    /// Signed sources are read through the private, read-only probe.  The
    /// marker never authorizes changed publication; it only lets a final
    /// Workbook byte comparison retain the original signed artifact.
    signed: bool,
}

/// Immutable, cheaply cloned snapshot of an editable XLS package.
///
/// A snapshot retains the object-capture and low-level CFB admission profiles
/// accepted at ingress. Those profiles are part of the snapshot's policy
/// identity: equal bytes opened under different profiles cannot be joined,
/// and every candidate derived from a snapshot remains subject to its
/// retained profiles.
#[derive(Clone)]
pub struct Snapshot {
    inner: Arc<Inner>,
}

/// Explicit bounded undo/redo history of immutable XLS snapshots.
#[cfg_attr(
    not(test),
    allow(
        dead_code,
        reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
    )
)]
pub(crate) type History = litchi_core::patch::History<Snapshot>;

impl Snapshot {
    /// Adapts this editor's retained object and CFB admission profiles for
    /// variable-length raw BIFF records.  Structural certification and any
    /// nested record reader must observe the same limits as the snapshot's
    /// source capture; it must not silently fall back to a broad default.
    pub(crate) fn raw_record_limits(&self) -> Result<crate::SourceLimits> {
        crate::SourceLimits::try_from_profile(
            self.inner.cfb_limits.max_input_bytes(),
            self.inner.limits.max_stream_size,
            self.inner.limits.max_total_size,
        )
    }

    /// Opens an unencrypted XLS package under an explicit object capture
    /// profile before publishing the editable snapshot.
    ///
    /// The bounded corpus and other security-sensitive callers use this entry
    /// point instead of inheriting the broad convenience defaults.
    pub(crate) fn from_bytes(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        let package = ObjectPackage::open_with_options(
            bytes,
            OpenOptions::new()
                .targets(Targets::default())
                .limits(limits)
                .cfb_limits(OleFileLimits::default()),
        )?;
        Self::open_package(package, None)
    }

    /// Opens an unencrypted XLS package under explicit object and CFB
    /// admission profiles before publishing the editable snapshot.
    pub(crate) fn from_bytes_with_cfb(
        bytes: Vec<u8>,
        limits: Limits,
        cfb_limits: OleFileLimits,
    ) -> Result<Self> {
        let package = ObjectPackage::open_with_options(
            bytes,
            OpenOptions::new()
                .targets(Targets::default())
                .limits(limits)
                .cfb_limits(cfb_limits),
        )?;
        Self::open_package(package, None)
    }

    /// Opens the editable snapshot over a positional CFB owner that is
    /// already open, such as the one a source-backed workbook or a format
    /// detector holds.
    ///
    /// This is the re-index fallback of unified-ops record 0003 (risk R2):
    /// the package editor captures its streams through the cursor-based CFB
    /// reader, so this constructor copies the source once, under the CFB
    /// owner's version fence, and the editor indexes that copy again. Both
    /// costs fall when an edit begins, never when the workbook is opened. The
    /// copy is bounded by `cfb_limits`, which also govern the editor's index.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Cfb`] when the source changed or exceeds
    /// `cfb_limits`, [`Error::Allocation`] when the copy cannot be reserved,
    /// and otherwise the errors of [`Self::from_bytes_with_cfb`].
    pub(crate) fn from_shared_ole_file(
        cfb: &litchi_cfb::SharedOleFile,
        limits: Limits,
        cfb_limits: OleFileLimits,
    ) -> Result<Self> {
        let bytes = Self::copy_shared_source(cfb, cfb_limits)?;
        Self::from_bytes_with_cfb(bytes, limits, cfb_limits)
    }

    /// Opens a source-backed signed image through the read-only probe while
    /// retaining the source freshness fence used by the ordinary editor.
    pub(crate) fn from_shared_probe(
        cfb: &litchi_cfb::SharedOleFile,
        limits: Limits,
        cfb_limits: OleFileLimits,
    ) -> Result<Self> {
        let bytes = Self::copy_shared_source(cfb, cfb_limits)?;
        let probe = Probe::open_with_options(
            bytes,
            OpenOptions::new().limits(limits).cfb_limits(cfb_limits),
        )
        .map_err(Error::Object)?;
        Self::from_probe(probe)
    }

    fn copy_shared_source(
        cfb: &litchi_cfb::SharedOleFile,
        cfb_limits: OleFileLimits,
    ) -> Result<Vec<u8>> {
        let expected = cfb.source_version()?;
        let length = cfb.file_size();
        if length > cfb_limits.max_input_bytes() {
            return Err(Error::Cfb(litchi_cfb::OleError::LimitExceeded {
                resource: "input bytes",
                observed: length,
                maximum: cfb_limits.max_input_bytes(),
            }));
        }
        let length = usize::try_from(length)
            .map_err(|_error| Error::Allocation("XLS editor source exceeds this platform"))?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(length)
            .map_err(|_error| Error::Allocation("XLS editor source copy"))?;
        bytes.resize(length, 0);
        let source = cfb.source_arc();
        let mut filled = 0_usize;
        while filled < length {
            let offset = u64::try_from(filled)
                .map_err(|_error| Error::InvalidData("XLS editor offset exceeds u64".into()))?;
            let read = source.read_at(offset, &mut bytes[filled..])?;
            if read == 0 {
                cfb.source_version()?;
                return Err(Error::InvalidData(
                    "XLS source ended while the editor copied it".into(),
                ));
            }
            filled += read;
        }
        let observed = cfb.source_version()?;
        if observed != expected {
            return Err(Error::Cfb(litchi_cfb::OleError::SourceChanged {
                expected,
                observed,
            }));
        }
        Ok(bytes)
    }

    /// Build a private signed read snapshot from a bounded, read-only probe.
    /// The returned owner retains the original CFB bytes and signature state;
    /// changed publication is still refused by [`Edit::commit`].
    pub(crate) fn from_probe(probe: Probe) -> Result<Self> {
        Self::open_probe(probe)
    }

    /// Opens a captured package, optionally verifying a source-backed numeric
    /// commit's materialized target on the complete `raw::Reader` this open
    /// already parses.
    ///
    /// The optional verification is the exact sequence the source-backed
    /// numeric commit used to run on a *second* complete parse of the same
    /// bytes, in the same order and at the same point in the construction. It
    /// reads the parse that is already here instead of repeating it; the parse
    /// is a pure function of the sealed bytes and the retained limits, so the
    /// second one could only ever have reproduced this one.
    fn open_package(
        package: ObjectPackage,
        verification: Option<NumericTargetVerification<'_>>,
    ) -> Result<Self> {
        let limits = package.limits_value();
        let cfb_limits = package.cfb_limits_value();
        // Allocate only the selected path: BIFF8 packages name the stream
        // `Workbook`, and only a BIFF5-style package falls back to `Book`.
        let workbook_path = ["Workbook", "Book"]
            .into_iter()
            .map(|name| vec![name.to_string()])
            .find(|path| package.stream(path).is_some())
            .ok_or_else(|| {
                Error::InvalidData("XLS package has no Workbook or Book stream".into())
            })?;
        let workbook_stream = package
            .stream_shared(&workbook_path)
            .ok_or_else(|| Error::InvalidData("selected XLS Workbook stream disappeared".into()))?;
        let source = package.source_shared();
        Self::open_workbook_source(
            source,
            workbook_path,
            workbook_stream,
            limits,
            cfb_limits,
            false,
            verification,
        )
    }

    /// Opens a signed source captured by the read-only OLE probe.  The probe
    /// retains the original signature streams, while this owner indexes only
    /// the Workbook stream needed for semantic no-op comparison.
    fn open_probe(probe: Probe) -> Result<Self> {
        let workbook_path = [vec!["Workbook".to_string()], vec!["Book".to_string()]]
            .into_iter()
            .find(|path| probe.stream(path).is_some())
            .ok_or_else(|| {
                Error::InvalidData("XLS package has no Workbook or Book stream".into())
            })?;
        let workbook_stream = probe
            .stream_shared(&workbook_path)
            .ok_or_else(|| Error::InvalidData("selected XLS Workbook stream disappeared".into()))?;
        let limits = probe.limits_value();
        let cfb_limits = probe.cfb_limits_value();
        Self::open_workbook_source(
            probe.source_shared(),
            workbook_path,
            workbook_stream,
            limits,
            cfb_limits,
            probe.is_signed(),
            None,
        )
    }

    fn open_workbook_source(
        source: Arc<Vec<u8>>,
        workbook_path: Vec<String>,
        workbook_stream: Arc<[u8]>,
        limits: Limits,
        cfb_limits: OleFileLimits,
        signed: bool,
        verification: Option<NumericTargetVerification<'_>>,
    ) -> Result<Self> {
        let source_version = fresh_snapshot_source_version();

        let (mut sheets, sst_total_offset, xf_records) = parse_workbook_stream(&workbook_stream)?;
        // A full semantic open catches cross-stream and workbook-global
        // dependencies before the narrower source-offset inventory is kept.
        // The legacy reader intentionally skips some malformed optional sheet
        // projections, so this edit owner additionally requires every sheet
        // it can mutate to have survived that complete semantic open. It is
        // the complete reader's validation-only mode: every record is
        // validated exactly as `Reader::new` validates it, and no cell is
        // decoded except the ones a verification reads back.
        let kept = match verification {
            Some(verification) => verification.kept_cells()?,
            None => KeptCells::none(),
        };
        let workbook = Reader::validation_only_with_cfb_limits(
            Cursor::new(source.as_ref()),
            kept,
            cfb_limits,
        )?;
        let source_policy = SourcePolicyFacts::from_workbook(&workbook, &sheets)?;
        let shared_strings = workbook.shared_strings_shared();
        let shared_string_properties =
            retained_shared_string_properties(&workbook, shared_strings.len())?;
        // A plain open drops the complete reader here, exactly where it
        // dropped before. Only a verification owner keeps it alive, and only
        // for the checks it would otherwise have re-parsed the same bytes to
        // run. Pairing the two in one `Option` is deliberate: there is no way
        // to keep the reader without also running the checks that asked for
        // it, or to ask for the checks and find the reader already gone.
        let pending = verification.map(|verification| (verification, workbook));
        resolve_shared_strings(&mut sheets, &shared_strings)?;
        if let Some((verification, workbook)) = &pending {
            verification.run(workbook, &source, &workbook_path, &workbook_stream)?;
        }
        drop(pending);
        Ok(Self {
            inner: Arc::new(Inner {
                bytes: source,
                source_version,
                limits,
                cfb_limits,
                workbook_path,
                workbook_stream,
                shared_strings,
                shared_string_properties,
                sst_total_offset,
                xf_records: Arc::new(xf_records),
                sheets,
                source_policy,
                signed,
            }),
        })
    }

    /// Builds a fixed-numeric commit's target snapshot from the artifact its
    /// package publication rendered and validated for `package`'s state.
    fn from_fixed_numeric_package_editor(
        package: ObjectPackage,
        source_snapshot: &Self,
        changes: &[Change],
    ) -> Result<Self> {
        let workbook_stream = package
            .stream_shared(&source_snapshot.inner.workbook_path)
            .ok_or_else(|| Error::InvalidData("selected XLS Workbook stream disappeared".into()))?;
        let sheets = carry_fixed_numeric_inventory(source_snapshot, &workbook_stream, changes)?;
        let source = package.source_shared();

        // Keep the complete public reader as an independent validation owner.
        // Only the private offset inventory is carried forward after proving
        // that every other Workbook-stream byte is unchanged. Its
        // validation-only mode decodes just the cells the readback reads.
        let workbook = Reader::validation_only_with_cfb_limits(
            Cursor::new(source.as_slice()),
            numeric_readback_cells(source_snapshot, changes)?,
            source_snapshot.inner.cfb_limits,
        )?;
        let source_policy = SourcePolicyFacts::from_workbook(&workbook, &sheets)?;
        verify_public_numeric_readback(&workbook, source_snapshot, changes)?;

        Ok(Self {
            inner: Arc::new(Inner {
                bytes: source,
                source_version: fresh_snapshot_source_version(),
                limits: package.limits_value(),
                cfb_limits: package.cfb_limits_value(),
                workbook_path: source_snapshot.inner.workbook_path.clone(),
                workbook_stream,
                shared_strings: Arc::clone(&source_snapshot.inner.shared_strings),
                shared_string_properties: Arc::clone(
                    &source_snapshot.inner.shared_string_properties,
                ),
                sst_total_offset: source_snapshot.inner.sst_total_offset,
                xf_records: Arc::clone(&source_snapshot.inner.xf_records),
                sheets,
                source_policy,
                signed: source_snapshot.inner.signed,
            }),
        })
    }

    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
        )
    )]
    fn retag_source_version(mut self, source_version: SourceVersion) -> Self {
        if self.inner.source_version == source_version {
            return self;
        }
        if let Some(inner) = Arc::get_mut(&mut self.inner) {
            inner.source_version = source_version;
            return self;
        }
        let inner = &self.inner;
        Self {
            inner: Arc::new(Inner {
                bytes: Arc::clone(&inner.bytes),
                source_version,
                limits: inner.limits,
                cfb_limits: inner.cfb_limits,
                workbook_path: inner.workbook_path.clone(),
                workbook_stream: Arc::clone(&inner.workbook_stream),
                shared_strings: Arc::clone(&inner.shared_strings),
                shared_string_properties: Arc::clone(&inner.shared_string_properties),
                sst_total_offset: inner.sst_total_offset,
                xf_records: Arc::clone(&inner.xf_records),
                sheets: inner.sheets.clone(),
                source_policy: inner.source_policy,
                signed: inner.signed,
            }),
        }
    }

    /// Returns exact source CFB bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.inner.bytes
    }

    /// Returns shared ownership of the exact immutable artifact bytes.
    ///
    /// This is crate-private so source-backed publication can hand the same
    /// allocation to `OwnedSource::from_arc`; public callers continue to use
    /// the borrowed [`Self::bytes`] projection.
    pub(crate) fn bytes_shared(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.inner.bytes)
    }

    /// Effective complete-package capture profile retained by this snapshot.
    pub(crate) fn object_limits(&self) -> Limits {
        self.inner.limits
    }

    /// Effective CFB admission profile retained by this snapshot.
    pub(crate) fn cfb_limits(&self) -> OleFileLimits {
        self.inner.cfb_limits
    }

    /// Retain the narrower profile; reuse the validated owner when unchanged.
    pub(crate) fn constrain_limits(
        &self,
        requested: Limits,
        requested_cfb: OleFileLimits,
    ) -> Result<Self> {
        let (limits, cfb_limits) = self.effective_limits(requested, requested_cfb)?;
        let retained = self.inner.limits;
        let retained_cfb = self.inner.cfb_limits;
        if limits == retained && cfb_limits == retained_cfb {
            return Ok(self.clone());
        }
        if self.inner.signed {
            let probe = Probe::open_with_options(
                self.bytes().to_vec(),
                OpenOptions::new().limits(limits).cfb_limits(cfb_limits),
            )
            .map_err(Error::Object)?;
            return Self::from_probe(probe);
        }
        Self::from_slice_with_limits_and_cfb(self.bytes(), limits, cfb_limits)
    }

    /// Copy an immutable image only after the selected physical input bound.
    pub(crate) fn from_slice_with_limits_and_cfb(
        bytes: &[u8],
        limits: Limits,
        cfb_limits: OleFileLimits,
    ) -> Result<Self> {
        let observed = u64::try_from(bytes.len()).unwrap_or(u64::MAX);
        if observed > cfb_limits.max_input_bytes() {
            return Err(Error::Cfb(litchi_cfb::OleError::LimitExceeded {
                resource: "input bytes",
                observed,
                maximum: cfb_limits.max_input_bytes(),
            }));
        }
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(bytes.len())
            .map_err(|_| Error::Allocation("XLS primary edit source"))?;
        owned.extend_from_slice(bytes);
        Self::from_bytes_with_cfb(owned, limits, cfb_limits)
    }

    /// Intersect selected profiles without reparsing the immutable image.
    pub(crate) fn effective_limits(
        &self,
        requested: Limits,
        requested_cfb: OleFileLimits,
    ) -> Result<(Limits, OleFileLimits)> {
        let retained = self.inner.limits;
        let limits = Limits {
            max_objects: retained.max_objects.min(requested.max_objects),
            max_storage_depth: retained.max_storage_depth.min(requested.max_storage_depth),
            max_streams_per_object: retained
                .max_streams_per_object
                .min(requested.max_streams_per_object),
            max_streams: retained.max_streams.min(requested.max_streams),
            max_stream_size: retained.max_stream_size.min(requested.max_stream_size),
            max_object_size: retained.max_object_size.min(requested.max_object_size),
            max_total_size: retained.max_total_size.min(requested.max_total_size),
        };
        let retained_cfb = self.inner.cfb_limits;
        let cfb_limits = OleFileLimits::new(
            retained_cfb
                .max_input_bytes()
                .min(requested_cfb.max_input_bytes()),
        )?
        .with_max_directory_bytes(
            retained_cfb
                .max_directory_bytes()
                .min(requested_cfb.max_directory_bytes()),
        )?
        .with_max_allocation_table_bytes(
            retained_cfb
                .max_allocation_table_bytes()
                .min(requested_cfb.max_allocation_table_bytes()),
        )?;
        Ok((limits, cfb_limits))
    }

    /// Opaque source version captured for this immutable snapshot lineage.
    ///
    /// `Snapshot::from_bytes` mints this process-local token over its owned,
    /// immutable byte allocation; it is not an external-file change watcher.
    /// The token is stable across cheap snapshot and worksheet-handle clones.
    /// A source-backed commit derives its target token from this lineage and
    /// the validated overlay target fingerprint.
    #[must_use]
    pub fn source_version(&self) -> SourceVersion {
        self.inner.source_version
    }

    /// Returns the exact source Workbook stream.
    #[must_use]
    pub fn workbook_stream(&self) -> &[u8] {
        &self.inner.workbook_stream
    }

    /// Returns the number of worksheet tabs exposed by this bounded owner.
    #[must_use]
    pub fn worksheet_count(&self) -> usize {
        self.inner.sheets.len()
    }

    /// Iterates worksheet tabs in workbook order.
    #[must_use]
    pub fn worksheets(&self) -> impl ExactSizeIterator<Item = Worksheet<'_>> {
        (0..self.inner.sheets.len()).map(|index| Worksheet {
            snapshot: self,
            index,
        })
    }

    /// Resolves a worksheet without exposing a physical stream identifier.
    ///
    /// # Errors
    ///
    /// Returns an error if a malformed source produced an ambiguous name.
    pub fn worksheet<'a>(&'a self, selector: Selector<'_>) -> Result<Option<Worksheet<'a>>> {
        Ok(self.resolve_sheet(selector)?.map(|index| Worksheet {
            snapshot: self,
            index,
        }))
    }

    /// Starts a detached edit against this immutable artifact.
    #[must_use]
    pub(crate) fn edit(&self) -> Edit {
        Edit {
            source: self.clone(),
            changes: Vec::new(),
            structural_changes: Vec::new(),
            resource_changes: Vec::new(),
            fixed_index: FixedChangeIndex::default(),
            structural_index: StructuralChangeIndex::default(),
            tab_change: None,
            layout_changes: Vec::new(),
            source_strings: OnceLock::new(),
        }
    }

    /// Starts bounded undo/redo history at this immutable snapshot.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
        )
    )]
    #[must_use]
    pub(crate) fn history(&self, limits: HistoryLimits) -> History {
        History::new(self.clone(), limits)
    }

    fn resolve_sheet(&self, selector: Selector<'_>) -> Result<Option<usize>> {
        match selector {
            Selector::Position(position) => Ok(self
                .inner
                .sheets
                .iter()
                .position(|sheet| sheet.workbook_index == position.get())),
            Selector::Name(name) => {
                let mut matches = self
                    .inner
                    .sheets
                    .iter()
                    .enumerate()
                    .filter(|(_, sheet)| caseless_eq(&sheet.name, name));
                let found = matches.next().map(|(index, _)| index);
                if matches.next().is_some() {
                    return Err(Error::refused(
                        RefusalKind::InvalidValue,
                        "edit cells",
                        "worksheet name is ambiguous",
                    ));
                }
                Ok(found)
            },
        }
    }

    /// Resolves one cell against this immutable source inventory.
    ///
    /// The package edit facade uses this read-only helper while its mutable
    /// edit engine is still lazy.  Keeping the lookup on the snapshot avoids
    /// forcing a mutable borrow merely to inspect the original source.
    pub(crate) fn source_cell(
        &self,
        selector: Selector<'_>,
        reference: Reference,
    ) -> Result<Option<&Cell>> {
        let sheet = self.resolve_sheet(selector)?.ok_or_else(|| {
            Error::WorksheetNotFound("cell-value source lookup worksheet selector".into())
        })?;
        let sheet_data = &self.inner.sheets[sheet];
        Ok(sheet_data
            .unique_entry_index(reference)?
            .map(|index| &sheet_data.entries[index].cell))
    }
}

impl fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("artifact_bytes", &self.inner.bytes.len())
            .field("workbook_bytes", &self.inner.workbook_stream.len())
            .field("worksheet_count", &self.inner.sheets.len())
            .finish()
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.bytes() == other.bytes()
    }
}

impl Eq for Snapshot {}

/// Borrowed view of one worksheet's editable `Number` records.
#[derive(Debug, Clone, Copy)]
pub struct Worksheet<'a> {
    snapshot: &'a Snapshot,
    index: usize,
}

impl<'a> Worksheet<'a> {
    fn data(self) -> &'a SheetData {
        &self.snapshot.inner.sheets[self.index]
    }

    /// Returns the zero-based workbook tab position.
    #[must_use]
    pub fn position(self) -> usize {
        self.data().workbook_index
    }

    /// Returns the developer-visible worksheet name.
    #[must_use]
    pub fn name(self) -> &'a str {
        &self.data().name
    }

    /// Iterates editable numeric values in source order.
    #[must_use]
    pub fn numbers(self) -> impl Iterator<Item = Number> + 'a {
        self.data().entries.iter().filter_map(|entry| {
            if entry.cell.storage == Storage::Number {
                let Value::Number(value) = &entry.cell.value else {
                    return None;
                };
                Some(Number {
                    reference: entry.cell.reference,
                    value: *value,
                })
            } else {
                None
            }
        })
    }

    /// Iterates all transaction-owned cells in source order.
    #[must_use]
    pub fn cells(self) -> impl ExactSizeIterator<Item = &'a Cell> {
        self.data().entries.iter().map(|entry| &entry.cell)
    }

    /// Looks up an editable numeric value.
    ///
    /// # Errors
    ///
    /// Returns an exact-source ambiguity error for duplicate `Number` records.
    pub fn number(self, reference: Reference) -> Result<Option<Number>> {
        self.data()
            .unique_entry_index(reference)
            .map(|entry| entry.map(|index| &self.data().entries[index]))
            .map(|entry| {
                entry.and_then(|item| {
                    if item.cell.storage != Storage::Number {
                        return None;
                    }
                    let Value::Number(value) = &item.cell.value else {
                        return None;
                    };
                    Some(Number {
                        reference,
                        value: *value,
                    })
                })
            })
    }

    /// Looks up any transaction-owned cell.
    ///
    /// # Errors
    ///
    /// Returns an exact-source ambiguity error for duplicate cell records.
    pub fn cell(self, reference: Reference) -> Result<Option<&'a Cell>> {
        self.data()
            .unique_entry_index(reference)
            .map(|entry| entry.map(|index| &self.data().entries[index].cell))
    }
}

#[derive(Debug, Clone)]
struct Change {
    sheet: usize,
    entry: usize,
    reference: Reference,
    storage: Storage,
    value: Value,
}

#[derive(Debug, Clone)]
enum StructuralChange {
    Cell {
        sheet: usize,
        reference: Reference,
        before: Option<(Storage, Value, StyleIndex)>,
        after: Option<(Storage, Value, StyleIndex)>,
    },
    Rows {
        sheet: usize,
        start: u16,
        count: u16,
        insert: bool,
    },
    Columns {
        sheet: usize,
        start: u8,
        count: u8,
        insert: bool,
    },
    RenameSheet {
        sheet: usize,
        before: String,
        after: String,
    },
}

#[derive(Debug, Clone)]
enum ResourceChange {
    SharedString {
        text: String,
        insert: bool,
    },
    RichSharedString {
        text: String,
        formatting_runs: Vec<crate::records::SharedStringFormatRun>,
        insert: bool,
    },
    FormulaCell {
        sheet: usize,
        reference: Reference,
        style: StyleIndex,
        tokens: Vec<u8>,
        insert: bool,
    },
    ExtendedFormat {
        index: StyleIndex,
        payload: Vec<u8>,
        insert: bool,
    },
}

#[derive(Debug, Clone, Copy)]
struct StagedCell {
    sheet: usize,
    reference: Reference,
    change: usize,
}

/// Sorted sidecar for staged fixed-width cell operations.
///
/// The transaction vector remains the canonical operation storage.  This
/// index only maps a source coordinate to its operation position, preserving
/// repeated-write coalescing while making overlap checks logarithmic.
#[derive(Debug, Clone, Default)]
struct FixedChangeIndex {
    cells: Vec<StagedCell>,
}

impl FixedChangeIndex {
    fn sort(&mut self) {
        self.cells
            .sort_unstable_by_key(|item| (item.sheet, item.reference, item.change));
    }

    fn range(&self, sheet: usize, reference: Reference) -> std::ops::Range<usize> {
        let start = self.lower_bound(sheet, reference);
        let end = self.upper_bound(sheet, reference);
        start..end
    }

    fn lower_bound(&self, sheet: usize, reference: Reference) -> usize {
        self.cells
            .partition_point(|item| (item.sheet, item.reference) < (sheet, reference))
    }

    fn upper_bound(&self, sheet: usize, reference: Reference) -> usize {
        let mut index = self.lower_bound(sheet, reference);
        while self
            .cells
            .get(index)
            .is_some_and(|item| item.sheet == sheet && item.reference == reference)
        {
            index += 1;
        }
        index
    }

    fn first(&self, sheet: usize, reference: Reference) -> Option<usize> {
        self.cells
            .get(self.lower_bound(sheet, reference))
            .filter(|item| item.sheet == sheet && item.reference == reference)
            .map(|item| item.change)
    }

    fn any_effective(
        &self,
        changes: &[Change],
        source: &Snapshot,
        sheet: usize,
        reference: Reference,
    ) -> bool {
        self.range(sheet, reference).any(|position| {
            self.cells
                .get(position)
                .and_then(|item| changes.get(item.change))
                .is_some_and(|change| change_is_effective(source, change))
        })
    }

    fn has_sheet(&self, sheet: usize) -> bool {
        let start = self.cells.partition_point(|item| item.sheet < sheet);
        self.cells
            .get(start)
            .is_some_and(|item| item.sheet == sheet)
    }

    fn any_in_range(&self, sheet: usize, range: CellRange) -> bool {
        let start = self.lower_bound(
            sheet,
            Reference {
                row: range.start().row(),
                column: 0,
            },
        );
        self.cells[start..]
            .iter()
            .take_while(|item| item.sheet == sheet && item.reference.row() <= range.end().row())
            .any(|item| range.contains(item.reference))
    }

    fn reserve_one(&mut self) -> Result<()> {
        self.cells
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("indexing staged XLS cell changes"))
    }

    fn reserve_for_join(&mut self, additions: usize) -> Result<()> {
        self.cells
            .try_reserve(additions)
            .map_err(|_error| Error::Allocation("indexing staged XLS cell changes"))
    }

    fn insert(&mut self, item: StagedCell) {
        let mut position = self.lower_bound(item.sheet, item.reference);
        while self.cells.get(position).is_some_and(|existing| {
            existing.sheet == item.sheet && existing.reference == item.reference
        }) {
            position += 1;
        }
        self.cells.insert(position, item);
    }

    fn rebuild_in_place(&mut self, changes: &[Change]) {
        self.cells.clear();
        for (change, item) in changes.iter().enumerate() {
            self.cells.push(StagedCell {
                sheet: item.sheet,
                reference: item.reference,
                change,
            });
        }
        self.sort();
    }
}

#[derive(Debug, Clone, Default)]
struct StructuralChangeIndex {
    cells: Vec<StagedStructuralCell>,
    axes: Vec<StagedStructuralSheet>,
    renames: Vec<StagedStructuralSheet>,
}

#[derive(Debug, Clone, Copy)]
struct StagedStructuralCell {
    sheet: usize,
    reference: Reference,
    change: usize,
}

#[derive(Debug, Clone, Copy)]
struct StagedStructuralSheet {
    sheet: usize,
    change: usize,
}

impl StructuralChangeIndex {
    fn push(&mut self, item: &StructuralChange, change: usize) {
        match item {
            StructuralChange::Cell {
                sheet, reference, ..
            } => self.cells.push(StagedStructuralCell {
                sheet: *sheet,
                reference: *reference,
                change,
            }),
            StructuralChange::Rows { sheet, .. } | StructuralChange::Columns { sheet, .. } => {
                self.axes.push(StagedStructuralSheet {
                    sheet: *sheet,
                    change,
                });
            },
            StructuralChange::RenameSheet { sheet, .. } => {
                self.renames.push(StagedStructuralSheet {
                    sheet: *sheet,
                    change,
                });
            },
        }
    }

    fn sort(&mut self) {
        self.cells
            .sort_unstable_by_key(|item| (item.sheet, item.reference, item.change));
        self.axes
            .sort_unstable_by_key(|item| (item.sheet, item.change));
        self.renames
            .sort_unstable_by_key(|item| (item.sheet, item.change));
    }

    fn cell(&self, sheet: usize, reference: Reference) -> Option<usize> {
        let index = self
            .cells
            .partition_point(|item| (item.sheet, item.reference) < (sheet, reference));
        self.cells
            .get(index)
            .filter(|item| item.sheet == sheet && item.reference == reference)
            .map(|item| item.change)
    }

    fn axis(&self, sheet: usize) -> Option<usize> {
        let index = self.axes.partition_point(|item| item.sheet < sheet);
        self.axes
            .get(index)
            .filter(|item| item.sheet == sheet)
            .map(|item| item.change)
    }

    fn rename(&self, sheet: usize) -> Option<usize> {
        let index = self.renames.partition_point(|item| item.sheet < sheet);
        self.renames
            .get(index)
            .filter(|item| item.sheet == sheet)
            .map(|item| item.change)
    }

    fn has_cell_sheet(&self, sheet: usize) -> bool {
        let index = self.cells.partition_point(|item| item.sheet < sheet);
        self.cells
            .get(index)
            .is_some_and(|item| item.sheet == sheet)
    }

    fn overlaps(&self, change: &StructuralChange) -> bool {
        let sheet = operation_sheet_index(change);
        match change {
            StructuralChange::Cell { reference, .. } => {
                self.axis(sheet).is_some() || self.cell(sheet, *reference).is_some()
            },
            StructuralChange::Rows { .. } | StructuralChange::Columns { .. } => {
                self.axis(sheet).is_some() || self.has_cell_sheet(sheet)
            },
            StructuralChange::RenameSheet { .. } => self.rename(sheet).is_some(),
        }
    }

    fn conflict_with(
        &self,
        changes: &[StructuralChange],
        right: &StructuralChange,
    ) -> Option<usize> {
        let sheet = operation_sheet_index(right);
        match right {
            StructuralChange::Cell { reference, .. } => {
                if let Some(index) = self.axis(sheet)
                    && !structural_changes_equal(&changes[index], right)
                {
                    return Some(index);
                }
                self.cell(sheet, *reference)
                    .filter(|index| !structural_changes_equal(&changes[*index], right))
            },
            StructuralChange::Rows { .. } | StructuralChange::Columns { .. } => {
                if let Some(index) = self.axis(sheet) {
                    if !structural_changes_equal(&changes[index], right) {
                        return Some(index);
                    }
                }
                self.first_cell_sheet(sheet)
            },
            StructuralChange::RenameSheet { .. } => self
                .rename(sheet)
                .filter(|index| !structural_changes_equal(&changes[*index], right)),
        }
    }

    fn equal_with(&self, changes: &[StructuralChange], right: &StructuralChange) -> bool {
        let sheet = operation_sheet_index(right);
        match right {
            StructuralChange::Cell { reference, .. } => self
                .cell(sheet, *reference)
                .is_some_and(|index| structural_changes_equal(&changes[index], right)),
            StructuralChange::Rows { .. } | StructuralChange::Columns { .. } => self
                .axis(sheet)
                .is_some_and(|index| structural_changes_equal(&changes[index], right)),
            StructuralChange::RenameSheet { .. } => self
                .rename(sheet)
                .is_some_and(|index| structural_changes_equal(&changes[index], right)),
        }
    }

    fn first_cell_sheet(&self, sheet: usize) -> Option<usize> {
        let index = self.cells.partition_point(|item| item.sheet < sheet);
        self.cells
            .get(index)
            .filter(|item| item.sheet == sheet)
            .map(|item| item.change)
    }

    fn reserve_one(&mut self, change: &StructuralChange) -> Result<()> {
        match change {
            StructuralChange::Cell { .. } => self
                .cells
                .try_reserve(1)
                .map_err(|_error| Error::Allocation("indexing structural XLS cells")),
            StructuralChange::Rows { .. } | StructuralChange::Columns { .. } => self
                .axes
                .try_reserve(1)
                .map_err(|_error| Error::Allocation("indexing structural XLS axes")),
            StructuralChange::RenameSheet { .. } => self
                .renames
                .try_reserve(1)
                .map_err(|_error| Error::Allocation("indexing structural XLS renames")),
        }
    }

    fn insert(&mut self, change: &StructuralChange, index: usize) {
        match change {
            StructuralChange::Cell {
                sheet, reference, ..
            } => {
                let position = self.cells.partition_point(|item| {
                    (item.sheet, item.reference, item.change) < (*sheet, *reference, index)
                });
                self.cells.insert(
                    position,
                    StagedStructuralCell {
                        sheet: *sheet,
                        reference: *reference,
                        change: index,
                    },
                );
            },
            StructuralChange::Rows { sheet, .. } | StructuralChange::Columns { sheet, .. } => {
                let position = self
                    .axes
                    .partition_point(|item| (item.sheet, item.change) < (*sheet, index));
                self.axes.insert(
                    position,
                    StagedStructuralSheet {
                        sheet: *sheet,
                        change: index,
                    },
                );
            },
            StructuralChange::RenameSheet { sheet, .. } => {
                let position = self
                    .renames
                    .partition_point(|item| (item.sheet, item.change) < (*sheet, index));
                self.renames.insert(
                    position,
                    StagedStructuralSheet {
                        sheet: *sheet,
                        change: index,
                    },
                );
            },
        }
    }

    fn reserve_for_join(&mut self, additions: usize) -> Result<()> {
        self.cells
            .try_reserve(additions)
            .map_err(|_error| Error::Allocation("indexing structural XLS cells"))?;
        self.axes
            .try_reserve(additions)
            .map_err(|_error| Error::Allocation("indexing structural XLS axes"))?;
        self.renames
            .try_reserve(additions)
            .map_err(|_error| Error::Allocation("indexing structural XLS renames"))
    }

    fn rebuild_in_place(&mut self, changes: &[StructuralChange]) {
        self.cells.clear();
        self.axes.clear();
        self.renames.clear();
        for (change, item) in changes.iter().enumerate() {
            self.push(item, change);
        }
        self.sort();
    }
}

/// Detached, failure-atomic edits of existing fixed-width BIFF8 cell fields.
#[derive(Clone)]
pub(crate) struct Edit {
    source: Snapshot,
    changes: Vec<Change>,
    structural_changes: Vec<StructuralChange>,
    resource_changes: Vec<ResourceChange>,
    fixed_index: FixedChangeIndex,
    structural_index: StructuralChangeIndex,
    /// A complete-catalog tab-structure change.  It is exclusive: an edit
    /// carrying one stages nothing else, because every other staged
    /// operation addresses tabs by their source catalog position.
    tab_change: Option<Box<tabs::TabChange>>,
    /// Row and column layout transitions, one per row or column.
    layout_changes: Vec<layout::LayoutChange>,
    /// Lazily built membership index of the source SST, shared by clones.
    source_strings: OnceLock<Arc<SourceStringIndex>>,
}

/// Hash index of one immutable source SST, for exact text membership.
///
/// Entries are `(hash, index)` pairs sorted by hash; a lookup compares the
/// candidate text with every source string of an equal hash, so collisions
/// cost a comparison and never a wrong answer.  Small tables are scanned
/// directly and build no index.
#[derive(Debug, Default)]
struct SourceStringIndex {
    entries: Vec<(u64, u32)>,
}

impl SourceStringIndex {
    /// Tables at or below this size are scanned without an index.
    const SCAN_LIMIT: usize = 32;

    fn hash(text: &str) -> u64 {
        use std::hash::{BuildHasher, BuildHasherDefault};
        BuildHasherDefault::<std::collections::hash_map::DefaultHasher>::default().hash_one(text)
    }

    fn new(strings: &[String]) -> Arc<Self> {
        if strings.len() <= Self::SCAN_LIMIT {
            return Arc::default();
        }
        let mut entries = Vec::new();
        // The source table was admitted under the snapshot's limits; an
        // index that cannot be retained falls back to scanning.
        if entries.try_reserve_exact(strings.len()).is_err() {
            return Arc::default();
        }
        for (index, text) in strings.iter().enumerate() {
            let Ok(index) = u32::try_from(index) else {
                return Arc::default();
            };
            entries.push((Self::hash(text), index));
        }
        entries.sort_unstable();
        Arc::new(Self { entries })
    }

    /// Counts the occurrences of `text` in `strings`, stopping at `limit`.
    fn count_at_most(&self, strings: &[String], text: &str, limit: usize) -> usize {
        if self.entries.is_empty() {
            return strings
                .iter()
                .filter(|candidate| candidate.as_str() == text)
                .take(limit)
                .count();
        }
        let hash = Self::hash(text);
        let start = self
            .entries
            .partition_point(|(candidate, _)| *candidate < hash);
        self.entries[start..]
            .iter()
            .take_while(|(candidate, _)| *candidate == hash)
            .filter(|(_, index)| {
                usize::try_from(*index)
                    .ok()
                    .and_then(|index| strings.get(index))
                    .is_some_and(|candidate| candidate.as_str() == text)
            })
            .take(limit)
            .count()
    }
}

/// One exact cell targeted differently by independently prepared transactions.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellConflict {
    sheet_position: usize,
    reference: Reference,
    left_storage: Storage,
    right_storage: Storage,
}

/// One structural semantic target with incompatible requested outcomes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OperationConflict {
    target: String,
}

impl OperationConflict {
    /// Canonical workbook-relative conflict target.
    #[must_use]
    pub fn target(&self) -> &str {
        &self.target
    }
}

impl CellConflict {
    /// Workbook tab position containing the conflict.
    #[must_use]
    pub const fn sheet_position(self) -> usize {
        self.sheet_position
    }

    /// Conflicting cell location.
    #[must_use]
    pub const fn reference(self) -> Reference {
        self.reference
    }

    /// Left transaction's requested storage.
    #[must_use]
    pub const fn left_storage(self) -> Storage {
        self.left_storage
    }

    /// Right transaction's requested storage.
    #[must_use]
    pub const fn right_storage(self) -> Storage {
        self.right_storage
    }
}

/// Failure to compose independently prepared workbook transactions.
#[derive(Debug)]
#[non_exhaustive]
pub enum JoinError {
    /// Transactions do not share the same exact immutable source artifact.
    DifferentSource,
    /// Transactions have equal source bytes but different retained admission
    /// profiles and therefore cannot be composed without discarding policy.
    DifferentLimits,
    /// At least one cell has incompatible requested outcomes.
    Conflicts(Box<[CellConflict]>),
    /// Structural operations overlap or request divergent outcomes.
    StructuralConflicts(Box<[OperationConflict]>),
    /// The joined transaction would exceed its finite change bound.
    Limit { observed: usize, limit: usize },
    /// Retaining deterministic conflict or joined state failed.
    Allocation,
}

impl Edit {
    fn fixed_overlaps_structural(&self, change: &StructuralChange) -> bool {
        let sheet = operation_sheet_index(change);
        match change {
            StructuralChange::Cell { reference, .. } => {
                self.fixed_index.first(sheet, *reference).is_some()
            },
            StructuralChange::Rows { .. } | StructuralChange::Columns { .. } => {
                self.fixed_index.has_sheet(sheet)
            },
            StructuralChange::RenameSheet { .. } => false,
        }
    }

    /// Joins independently prepared changes against the same exact snapshot.
    ///
    /// Disjoint and byte-equivalent cell outcomes compose deterministically.
    /// Conflicts and bounds are checked before this transaction changes.
    ///
    /// # Errors
    ///
    /// Returns the exact-source, structured cell-conflict, allocation, or
    /// finite-bound reason for refusing the incoming transaction.
    pub(crate) fn join(&mut self, incoming: Self) -> std::result::Result<&mut Self, JoinError> {
        if self.source.bytes() != incoming.source.bytes() {
            return Err(JoinError::DifferentSource);
        }
        if !self.layout_changes.is_empty() || !incoming.layout_changes.is_empty() {
            return Err(JoinError::StructuralConflicts(Box::new([
                OperationConflict {
                    target: "layout".to_string(),
                },
            ])));
        }
        if self.tab_change.is_some() || incoming.tab_change.is_some() {
            // A tab-structure change renumbers every catalog position the
            // other transaction addresses; it is never composed.
            return Err(JoinError::StructuralConflicts(Box::new([
                OperationConflict {
                    target: tabs::SEMANTIC_TARGET.to_string(),
                },
            ])));
        }
        if self.source.inner.limits != incoming.source.inner.limits
            || self.source.inner.cfb_limits != incoming.source.inner.cfb_limits
        {
            return Err(JoinError::DifferentLimits);
        }
        let mut conflicts = Vec::new();
        let mut additions = 0_usize;
        for right in &incoming.changes {
            if !change_is_effective(&incoming.source, right) {
                continue;
            }
            let left = self
                .fixed_index
                .range(right.sheet, right.reference)
                .filter_map(|position| self.fixed_index.cells.get(position))
                .filter_map(|item| self.changes.get(item.change))
                .find(|left| change_is_effective(&self.source, left));
            if let Some(left) = left {
                if left.storage != right.storage || !values_equal(&left.value, &right.value) {
                    conflicts
                        .try_reserve(1)
                        .map_err(|_error| JoinError::Allocation)?;
                    conflicts.push(CellConflict {
                        sheet_position: self.source.inner.sheets[left.sheet].workbook_index,
                        reference: left.reference,
                        left_storage: left.storage,
                        right_storage: right.storage,
                    });
                }
            } else if self
                .fixed_index
                .range(right.sheet, right.reference)
                .next()
                .is_none()
            {
                additions = additions.saturating_add(1);
            }
        }
        if !conflicts.is_empty() {
            return Err(JoinError::Conflicts(conflicts.into_boxed_slice()));
        }
        let mut structural_conflicts = Vec::new();
        for right in &incoming.structural_changes {
            if self
                .structural_index
                .conflict_with(&self.structural_changes, right)
                .is_some()
                || self.fixed_overlaps_structural(right)
            {
                structural_conflicts
                    .try_reserve(1)
                    .map_err(|_error| JoinError::Allocation)?;
                structural_conflicts.push(OperationConflict {
                    target: structural_target(&self.source, right),
                });
            }
        }
        for right in &incoming.changes {
            if !change_is_effective(&incoming.source, right) {
                continue;
            }
            if self.structural_index.overlaps(&StructuralChange::Cell {
                sheet: right.sheet,
                reference: right.reference,
                before: None,
                after: None,
            }) {
                structural_conflicts
                    .try_reserve(1)
                    .map_err(|_error| JoinError::Allocation)?;
                structural_conflicts.push(OperationConflict {
                    target: semantic_target(
                        self.source.inner.sheets[right.sheet].workbook_index,
                        right.reference,
                    ),
                });
            }
        }
        for right in &incoming.resource_changes {
            for left in &self.resource_changes {
                if (resource_target(left) == resource_target(right)
                    || resource_text(left).is_some_and(|text| {
                        resource_text(right).is_some_and(|candidate| candidate == text)
                    }))
                    && !resource_changes_equal(left, right)
                {
                    structural_conflicts
                        .try_reserve(1)
                        .map_err(|_error| JoinError::Allocation)?;
                    structural_conflicts.push(OperationConflict {
                        target: resource_target(right),
                    });
                }
            }
        }
        for (resource, structural) in self
            .resource_changes
            .iter()
            .flat_map(|resource| {
                incoming
                    .structural_changes
                    .iter()
                    .map(move |structural| (resource, structural))
            })
            .chain(incoming.resource_changes.iter().flat_map(|resource| {
                self.structural_changes
                    .iter()
                    .map(move |structural| (resource, structural))
            }))
        {
            if let Some((resource_sheet, _)) = formula_resource_cell(resource)
                && operation_sheet_index(structural) == resource_sheet
                && !matches!(structural, StructuralChange::RenameSheet { .. })
            {
                structural_conflicts
                    .try_reserve(1)
                    .map_err(|_error| JoinError::Allocation)?;
                structural_conflicts.push(OperationConflict {
                    target: structural_target(&self.source, structural),
                });
            }
        }
        structural_conflicts.sort_by(|left, right| left.target.cmp(&right.target));
        structural_conflicts.dedup();
        if !structural_conflicts.is_empty() {
            return Err(JoinError::StructuralConflicts(
                structural_conflicts.into_boxed_slice(),
            ));
        }
        let observed = self.changes.len().saturating_add(additions);
        let structural_additions = incoming
            .structural_changes
            .iter()
            .filter(|right| {
                !self
                    .structural_index
                    .equal_with(&self.structural_changes, right)
            })
            .count();
        let observed = observed
            .saturating_add(self.structural_changes.len())
            .saturating_add(structural_additions);
        let resource_additions = incoming
            .resource_changes
            .iter()
            .filter(|right| {
                !self
                    .resource_changes
                    .iter()
                    .any(|left| resource_changes_equal(left, right))
            })
            .count();
        let observed = observed
            .saturating_add(self.resource_changes.len())
            .saturating_add(resource_additions);
        if observed > MAX_STAGED_CHANGES {
            return Err(JoinError::Limit {
                observed,
                limit: MAX_STAGED_CHANGES,
            });
        }
        self.changes
            .try_reserve(additions)
            .map_err(|_error| JoinError::Allocation)?;
        self.fixed_index
            .reserve_for_join(additions)
            .map_err(|_error| JoinError::Allocation)?;
        self.structural_changes
            .try_reserve(structural_additions)
            .map_err(|_error| JoinError::Allocation)?;
        self.structural_index
            .reserve_for_join(structural_additions)
            .map_err(|_error| JoinError::Allocation)?;
        self.resource_changes
            .try_reserve(resource_additions)
            .map_err(|_error| JoinError::Allocation)?;
        for change in incoming.changes {
            if !change_is_effective(&incoming.source, &change) {
                continue;
            }
            if self.fixed_index.any_effective(
                &self.changes,
                &self.source,
                change.sheet,
                change.reference,
            ) {
                continue;
            }
            if let Some(change_index) = self.fixed_index.first(change.sheet, change.reference) {
                self.changes[change_index] = change;
                continue;
            }
            let change_index = self.changes.len();
            self.changes.push(change);
            let item = &self.changes[change_index];
            self.fixed_index.insert(StagedCell {
                sheet: item.sheet,
                reference: item.reference,
                change: change_index,
            });
        }
        self.changes.sort_by_key(|change| {
            (
                self.source.inner.sheets[change.sheet].workbook_index,
                change.reference,
            )
        });
        self.fixed_index.rebuild_in_place(&self.changes);
        for change in incoming.structural_changes {
            if !self
                .structural_index
                .equal_with(&self.structural_changes, &change)
            {
                let index = self.structural_changes.len();
                self.structural_changes.push(change);
                self.structural_index
                    .insert(&self.structural_changes[index], index);
            }
        }
        self.structural_changes
            .sort_by_key(|change| structural_target(&self.source, change));
        self.structural_index
            .rebuild_in_place(&self.structural_changes);
        for change in incoming.resource_changes {
            if !self
                .resource_changes
                .iter()
                .any(|left| resource_changes_equal(left, &change))
            {
                self.resource_changes.push(change);
            }
        }
        Ok(self)
    }

    /// Sets one existing BIFF8 `Number` value.
    ///
    /// The replacement must be a normal value or positive zero, as required
    /// by `Xnum`. An absent cell, or a cell represented by `RK`, `MulRk`, `Formula`,
    /// string, Boolean, error, or blank records, is a typed refusal rather than
    /// a request to change record families.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for an invalid Xnum, a missing or ambiguous
    /// worksheet cell, a duplicate cell, or a non-`Number` cell family, or an
    /// allocation failure.
    pub(crate) fn set_number(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        value: f64,
    ) -> Result<()> {
        if !valid_xnum(value) {
            return Err(Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "BIFF8 Number edits require a normal IEEE-754 value or positive zero",
            ));
        }
        let sheet_index = self
            .source
            .resolve_sheet(selector)?
            .ok_or_else(|| Error::WorksheetNotFound("cell-value edit worksheet selector".into()))?;
        let entries = &self.source.inner.sheets[sheet_index].entries;
        let entry = self.source.inner.sheets[sheet_index]
            .unique_entry_index(reference)?
            .ok_or_else(|| {
                Error::refused(
                    RefusalKind::UnsupportedTarget,
                    "edit cells",
                    "the selected cell is not an existing standalone BIFF8 Number",
                )
            })?;
        if entries[entry].cell.storage != Storage::Number {
            return Err(Error::refused(
                RefusalKind::UnsupportedTarget,
                "edit cells",
                "the selected cell is not an existing standalone BIFF8 Number",
            ));
        }
        self.stage(sheet_index, entry, Storage::Number, Value::Number(value))
    }

    /// Sets one existing numeric cell without changing its BIFF8 storage
    /// family.
    ///
    /// `Number` values retain their eight-byte IEEE-754 field. `RK` and
    /// `MulRk` values retain their four-byte compressed field and therefore
    /// accept only values that are exactly representable by the source RK
    /// encoding. This is the ordinary semantic entry point for the
    /// source-backed numeric publication path; callers needing a guaranteed
    /// standalone `Number` record may use [`Self::set_number`].
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for an absent or ambiguous cell, a nonnumeric
    /// source family, an invalid `Xnum`, a non-RK-representable replacement,
    /// or a transaction staging bound.
    pub(crate) fn set_numeric(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        value: f64,
    ) -> Result<()> {
        self.set_value(selector, reference, Value::Number(value))
    }

    /// Copies a bounded dependency-free scalar rectangle from another opened
    /// XLS snapshot into this transaction.
    ///
    /// Only existing standalone `Number` and formatting-only `Blank` records
    /// in the donor are representable.  Source styles are required to use
    /// canonical default cell XF, and no donor XF, SST, formula, drawing, relationship, or
    /// unknown BIFF owner is carried across the workbook boundary.  A missing
    /// target cell is authored with the target workbook's canonical default XF.  An
    /// existing target is accepted only when both source and target use the
    /// same fixed-width family (`Number` to `Number`, or `Blank` to `Blank`).
    /// Any family change, occupied target with no corresponding source cell,
    /// unsupported source owner, protected input, stale staged operation, or
    /// range beyond the finite transfer bound is rejected before this edit is
    /// changed.
    ///
    /// The operation is failure-atomic.  Its normal [`Self::commit`] path
    /// fully reopens the composed CFB and emits the existing exact-source
    /// reversible patch; same-width Number targets therefore retain the
    /// fixed-field publication fast path, while inserted cells use the
    /// checked structural closure already owned by this module.
    ///
    /// `Ok(None)` means that either worksheet selector did not resolve.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, security, dependency, occupancy, allocation,
    /// or representability refusal without partially staging a transfer.
    pub(crate) fn copy_scalar_cells_from<'source, 'target>(
        &mut self,
        donor: &Snapshot,
        source_selector: impl Into<Selector<'source>>,
        source_range: CellRange,
        target_selector: impl Into<Selector<'target>>,
        target_anchor: Reference,
    ) -> Result<Option<&mut Self>> {
        let source_selector = source_selector.into();
        let target_selector = target_selector.into();
        let Some(source_sheet) = donor.resolve_sheet(source_selector)? else {
            return Ok(None);
        };
        let Some(target_sheet) = self.source.resolve_sheet(target_selector)? else {
            return Ok(None);
        };
        let target_range = translated_range(source_range, target_anchor)?;
        let cell_count = source_range
            .cell_count()
            .ok_or_else(|| Error::InvalidData("scalar transfer area overflows usize".into()))?;
        if cell_count > MAX_SCALAR_TRANSFER_CELLS {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "the cross-workbook scalar transfer exceeds its cell limit",
            ));
        }

        ensure_scalar_transfer_container(donor, source_sheet)?;
        ensure_scalar_transfer_container(&self.source, target_sheet)?;
        ensure_transfer_has_no_staged_target_overlap(self, target_sheet, target_range)?;

        let source_entries = &donor.inner.sheets[source_sheet].entries;
        let target_entries = &self.source.inner.sheets[target_sheet].entries;
        let mut source_cells = Vec::new();
        source_cells
            .try_reserve_exact(cell_count.min(source_entries.len()))
            .map_err(|_error| Error::Allocation("staging scalar transfer source cells"))?;
        let mut estimated_bytes = 0_usize;
        let source_style = scalar_default_style(donor)?;
        for entry in source_entries
            .iter()
            .filter(|entry| source_range.contains(entry.cell.reference))
        {
            if source_cells.len() >= cell_count {
                return Err(Error::refused(
                    RefusalKind::InvalidValue,
                    "edit cells",
                    "cross-workbook scalar transfer source inventory exceeds its checked range",
                ));
            }
            if entry.cell.style != source_style {
                return Err(Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit cells",
                    "cross-workbook scalar transfer refuses source XF remapping",
                ));
            }
            if !matches!(entry.cell.storage, Storage::Number | Storage::Blank) {
                return Err(Error::UnsupportedFeature(
                    "cross-workbook scalar transfer accepts only Number and Blank cells".into(),
                ));
            }
            let value_bytes: usize = match entry.cell.storage {
                Storage::Number => 8,
                Storage::Blank => 0,
                _ => 0,
            };
            estimated_bytes = estimated_bytes
                .checked_add(value_bytes.saturating_add(32))
                .ok_or_else(|| Error::InvalidData("scalar transfer size overflows usize".into()))?;
            if estimated_bytes > MAX_SCALAR_TRANSFER_BYTES {
                return Err(Error::refused(
                    RefusalKind::InvalidValue,
                    "edit cells",
                    "the cross-workbook scalar transfer exceeds its staged byte limit",
                ));
            }
            push_bounded(
                &mut source_cells,
                (
                    entry.cell.reference,
                    entry.cell.storage,
                    entry.cell.value.clone(),
                ),
                cell_count,
                "staging scalar transfer source cells",
            )?;
        }
        source_cells.sort_unstable_by_key(|(reference, _, _)| *reference);
        for duplicate in source_cells.windows(2) {
            if duplicate[0].0 == duplicate[1].0 {
                return Err(Error::refused(
                    RefusalKind::UnsupportedContent,
                    "edit cells",
                    "a cell has duplicate BIFF8 owners",
                ));
            }
        }
        let source_cell_count = source_cells.len();

        let mut target_cells = Vec::new();
        target_cells
            .try_reserve_exact(cell_count.min(target_entries.len()))
            .map_err(|_error| Error::Allocation("indexing scalar transfer target cells"))?;
        for (entry_index, entry) in target_entries
            .iter()
            .enumerate()
            .filter(|(_, entry)| target_range.contains(entry.cell.reference))
        {
            if target_cells.len() >= cell_count {
                return Err(Error::refused(
                    RefusalKind::InvalidValue,
                    "edit cells",
                    "cross-workbook scalar transfer target inventory exceeds its checked range",
                ));
            }
            push_bounded(
                &mut target_cells,
                (entry.cell.reference, entry_index),
                cell_count,
                "indexing scalar transfer target cells",
            )?;
        }
        target_cells.sort_unstable_by_key(|(reference, _)| *reference);
        for duplicate in target_cells.windows(2) {
            if duplicate[0].0 == duplicate[1].0 {
                return Err(Error::refused(
                    RefusalKind::UnsupportedContent,
                    "edit cells",
                    "a cell has duplicate BIFF8 target owners",
                ));
            }
        }

        // Every existing target owner is checked before any target operation
        // is staged.  Sparse donor coordinates are intentionally not allowed
        // to overwrite unrelated target owners.
        for &(target_reference, target_index) in &target_cells {
            let target_entry = &target_entries[target_index];
            let source_reference =
                inverse_target_reference(source_range, target_anchor, target_reference)?;
            let Some(source_index) = source_cells
                .binary_search_by_key(&source_reference, |(reference, _, _)| *reference)
                .ok()
            else {
                return Err(Error::refused(
                    RefusalKind::UnsupportedTarget,
                    "edit cells",
                    "cross-workbook scalar transfer target range is occupied by an unrelated cell",
                ));
            };
            let source_storage = source_cells[source_index].1;
            if !matches!(target_entry.cell.storage, Storage::Number | Storage::Blank)
                || target_entry.cell.storage != source_storage
            {
                return Err(Error::refused(
                    RefusalKind::UnsupportedContent,
                    "edit cells",
                    "cross-workbook scalar transfer would change a BIFF cell record width or family",
                ));
            }
        }

        // Resolve every source-to-target mapping while the source and target
        // inventories are still immutable.  The detached clone below is the
        // only object touched during operation staging.
        let mut planned = Vec::new();
        planned
            .try_reserve_exact(source_cells.len())
            .map_err(|_error| Error::Allocation("staging scalar transfer targets"))?;
        for (source_reference, storage, value) in source_cells {
            let target_reference =
                source_range.target_reference(source_reference, target_anchor)?;
            let existing = target_cells
                .binary_search_by_key(&target_reference, |(reference, _)| *reference)
                .ok()
                .map(|target_index| &target_entries[target_cells[target_index].1]);
            match (storage, value, existing) {
                (Storage::Number, Value::Number(value), None) => {
                    push_bounded(
                        &mut planned,
                        TransferCell::Insert {
                            reference: target_reference,
                            value: Value::Number(value),
                        },
                        source_cell_count,
                        "staging scalar transfer targets",
                    )?;
                },
                (Storage::Number, Value::Number(value), Some(entry))
                    if entry.cell.storage == Storage::Number =>
                {
                    push_bounded(
                        &mut planned,
                        TransferCell::SetNumber {
                            reference: target_reference,
                            value,
                        },
                        source_cell_count,
                        "staging scalar transfer targets",
                    )?;
                },
                (Storage::Blank, Value::Blank, None) => {
                    push_bounded(
                        &mut planned,
                        TransferCell::Insert {
                            reference: target_reference,
                            value: Value::Blank,
                        },
                        source_cell_count,
                        "staging scalar transfer targets",
                    )?;
                },
                (Storage::Blank, Value::Blank, Some(entry))
                    if entry.cell.storage == Storage::Blank => {},
                _ => {
                    return Err(Error::refused(
                        RefusalKind::UnsupportedTarget,
                        "edit cells",
                        "cross-workbook scalar transfer source value is not representable in the target slot",
                    ));
                },
            }
        }

        let mut staged = self.clone();
        let target_position = staged.inner_sheet_position(target_sheet)?;
        let target_style = scalar_default_style(&staged.source)?;
        for operation in planned {
            match operation {
                TransferCell::Insert { reference, value } => {
                    staged.insert_cell_with_style(
                        Selector::Position(Position::new(target_position)),
                        reference,
                        value,
                        target_style,
                    )?;
                },
                TransferCell::SetNumber { reference, value } => {
                    staged.set_number(
                        Selector::Position(Position::new(target_position)),
                        reference,
                        value,
                    )?;
                },
            }
        }
        *self = staged;
        Ok(Some(self))
    }

    /// Copies one exact scalar cell from another snapshot.
    pub(crate) fn copy_scalar_cell_from<'source, 'target>(
        &mut self,
        donor: &Snapshot,
        source_selector: impl Into<Selector<'source>>,
        source_reference: Reference,
        target_selector: impl Into<Selector<'target>>,
        target_reference: Reference,
    ) -> Result<Option<&mut Self>> {
        self.copy_scalar_cells_from(
            donor,
            source_selector,
            CellRange::new(source_reference, source_reference)?,
            target_selector,
            target_reference,
        )
    }

    fn inner_sheet_position(&self, sheet: usize) -> Result<usize> {
        self.source
            .inner
            .sheets
            .get(sheet)
            .map(|item| item.workbook_index)
            .ok_or(Error::PatchConflict {
                operation: "edit cells",
                reason: "target worksheet inventory is stale",
            })
    }

    /// Sets an existing scalar cell through a storage-aware,
    /// dependency-checked path.
    ///
    /// Numeric `Number`, `RK`, and `MulRk` cells, `BoolErr`, SST
    /// references, and non-string Formula caches use their fixed-width fast
    /// paths. Scalar type changes that do not fit those record families use a
    /// structural cell replacement that retains the existing XF/style and
    /// row-block closure. A standalone `RK` and `LabelSst` can be converted
    /// into each other because both have the same ten-byte payload; the SST
    /// reference count is updated atomically. Missing simple text is interned
    /// into a bounded SST tail resource when an absent or canonical adjacent
    /// `ExtSST` can be updated without changing its bucket size. Formula
    /// records accept cached-result edits only and are never converted into
    /// ordinary scalar cells.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or duplicate cell, an unrepresentable
    /// storage conversion, an invalid numeric/cache value, or a finite bound.
    /// Failed validation leaves this edit unchanged.
    pub(crate) fn set_value(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        value: Value,
    ) -> Result<()> {
        // A missing text stages an SST resource before the cell operation;
        // that multi-stage path rolls the resource back on a later refusal
        // (every other staging step checks before it pushes), so an overlap
        // refusal cannot leave the resource behind. Fixed-width and
        // structural-only edits are already failure-atomic.
        let needs_candidate = match &value {
            Value::Text(text) => !self.has_shared_string(text),
            _ => false,
        };
        if needs_candidate {
            self.with_resource_rollback(|edit| edit.set_value_in_place(selector, reference, value))
        } else {
            self.set_value_in_place(selector, reference, value)
        }
    }

    fn set_value_in_place(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        value: Value,
    ) -> Result<()> {
        let sheet_index = self
            .source
            .resolve_sheet(selector)?
            .ok_or_else(|| Error::WorksheetNotFound("cell-value edit worksheet selector".into()))?;
        let entries = &self.source.inner.sheets[sheet_index].entries;
        let entry_index = self.source.inner.sheets[sheet_index]
            .unique_entry_index(reference)?
            .ok_or_else(|| {
                Error::UnsupportedFeature(format!(
                    "cell ({}, {}) has no representable source record slot",
                    reference.row(),
                    reference.column()
                ))
            })?;
        let source_cell = entries[entry_index].cell.clone();
        if source_cell.storage == Storage::Formula
            && (matches!(
                &source_cell.value,
                Value::FormulaCache(FormulaCache::String(_))
            ) || matches!(&value, Value::FormulaCache(FormulaCache::String(_))))
        {
            let Value::FormulaCache(cache) = &value else {
                return Err(Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit cells",
                    "Formula storage accepts only cached formula results",
                ));
            };
            if !valid_formula_cache(cache) {
                return Err(Error::refused(
                    RefusalKind::UnsupportedTarget,
                    "edit cells",
                    "formula cache is not representable in BIFF8",
                ));
            }
            return self.stage_structural(StructuralChange::Cell {
                sheet: sheet_index,
                reference,
                before: Some((
                    source_cell.storage,
                    source_cell.value.clone(),
                    source_cell.style,
                )),
                after: Some((source_cell.storage, value, source_cell.style)),
            });
        }
        let missing_text = match &value {
            Value::Text(text) if !self.has_shared_string(text) => Some(text.clone()),
            _ => None,
        };
        // The value's text is present once the missing resource is staged.
        let text_present = |text: &str| missing_text.is_some() || self.has_shared_string(text);
        match target_storage(source_cell.storage, &value, text_present) {
            Ok(target) => {
                if let Some(text) = missing_text {
                    self.stage_resource(ResourceChange::SharedString { text, insert: true })?;
                }
                if self.staged_removal(sheet_index, reference) {
                    return self.restore_removed_cell(sheet_index, reference, target, value);
                }
                self.stage(sheet_index, entry_index, target, value)
            },
            Err(error)
                if source_cell.storage != Storage::Formula
                    && !matches!(&value, Value::FormulaCache(_)) =>
            {
                // Fixed-width replacement is the common fast path above.  A
                // scalar whose BIFF record family cannot be converted in
                // place is still representable through the native structural
                // cell rewrite; retain the existing XF/style while replacing
                // the record and row-block closure atomically.  Formula
                // records stay on the dependency-checked cache path and are
                // never silently converted into ordinary scalar cells.
                let target = storage_for_new_value(&value, text_present).map_err(|_| error)?;
                if let Some(text) = missing_text {
                    self.stage_resource(ResourceChange::SharedString { text, insert: true })?;
                }
                if self.staged_removal(sheet_index, reference) {
                    return self.restore_removed_cell(sheet_index, reference, target, value);
                }
                self.stage_structural(StructuralChange::Cell {
                    sheet: sheet_index,
                    reference,
                    before: Some((
                        source_cell.storage,
                        source_cell.value.clone(),
                        source_cell.style,
                    )),
                    after: Some((target, value, source_cell.style)),
                })
            },
            Err(error) => Err(error),
        }
    }

    /// Clears a cell only when its existing storage is already semantically blank.
    ///
    /// BIFF8 has no fixed-width non-cell tombstone for the other supported
    /// record families. Removing them would require rebuilding row blocks,
    /// `INDEX`/`DBCELL`, dimensions, and formula dependencies, so that request
    /// is refused rather than emitted as an unknown record.
    ///
    /// # Errors
    ///
    /// Returns an error unless the selected source cell is an existing `Blank`.
    pub(crate) fn clear_cell(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
    ) -> Result<()> {
        let sheet_index = self
            .source
            .resolve_sheet(selector)?
            .ok_or_else(|| Error::WorksheetNotFound("cell clear worksheet selector".into()))?;
        let entries = &self.source.inner.sheets[sheet_index].entries;
        let entry_index = self.source.inner.sheets[sheet_index]
            .unique_entry_index(reference)?
            .ok_or_else(|| {
                Error::UnsupportedFeature("an absent BIFF8 cell is already clear".into())
            })?;
        if entries[entry_index].cell.storage != Storage::Blank {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "removing this cell requires a whole row-block rewrite",
            ));
        }
        self.stage(sheet_index, entry_index, Storage::Blank, Value::Blank)
    }

    /// Inserts a previously absent scalar cell using workbook XF zero.
    ///
    /// The affected worksheet's row blocks, `INDEX`, `DBCELL`, row extents,
    /// and `DIMENSIONS` are regenerated as one checked closure. Missing simple
    /// text is interned into a bounded SST tail resource; formulas and
    /// packed-cell overlaps are deliberately refused.
    ///
    /// # Errors
    ///
    /// Returns a typed grid, resource, occupancy, dependency, or bound error.
    pub(crate) fn insert_cell(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        value: Value,
    ) -> Result<()> {
        let style = StyleIndex::new(&self.source, 0)?;
        self.insert_cell_with_style(selector, reference, value, style)
    }

    /// Inserts one text cell backed by an exact rich SST identity.
    ///
    /// Formatting runs use UTF-16 character positions and must be strictly
    /// increasing, inside the text, and reference a font used by an effective
    /// XF. An absent text authors a new rich SST entry. Existing text is reused
    /// only when it occurs once and its formatting identity matches exactly.
    ///
    /// # Errors
    ///
    /// Returns a typed text, formatting-run, resource, occupancy, dependency,
    /// allocation, or transaction-bound refusal without partially staging.
    pub(crate) fn insert_rich_text_cell(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        text: String,
        formatting_runs: Vec<crate::records::SharedStringFormatRun>,
    ) -> Result<()> {
        validate_rich_text(self, &text, &formatting_runs)?;
        if self.has_shared_string(&text) {
            let occurrences = self
                .effective_shared_strings()
                .iter()
                .filter(|candidate| candidate.as_str() == text.as_str())
                .count();
            if occurrences != 1 || !self.has_rich_shared_string(&text, &formatting_runs) {
                return Err(Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit cells",
                    "rich SST reuse requires one exact formatting identity",
                ));
            }
            let mut staged = self.clone();
            staged.insert_cell(selector, reference, Value::Text(text))?;
            *self = staged;
            return Ok(());
        }
        let mut staged = self.clone();
        staged.stage_resource(ResourceChange::RichSharedString {
            text: text.clone(),
            formatting_runs,
            insert: true,
        })?;
        staged.insert_cell(selector, reference, Value::Text(text))?;
        *self = staged;
        Ok(())
    }

    /// Inserts one absent Formula cell from the writer's bounded tokenizer.
    ///
    /// The authored record uses an empty cached result and `fAlwaysCalc`, so a
    /// spreadsheet application can recalculate it without this crate claiming
    /// to evaluate the expression. No shared/array/table owner is synthesized.
    ///
    /// # Errors
    ///
    /// Returns a tokenizer, token-size, occupancy, style, allocation, or
    /// transaction-bound refusal without partially staging.
    pub(crate) fn insert_formula(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        expression: &str,
    ) -> Result<()> {
        let style = StyleIndex::new(&self.source, 0)?;
        self.insert_formula_with_style(selector, reference, expression, style)
    }

    /// Inserts one absent Formula cell using an effective XF resource.
    ///
    /// # Errors
    ///
    /// Returns a tokenizer, token-size, occupancy, style, allocation, or
    /// transaction-bound refusal without partially staging.
    pub(crate) fn insert_formula_with_style(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        expression: &str,
        style: StyleIndex,
    ) -> Result<()> {
        self.require_effective_style(style)?;
        if expression.is_empty() {
            return Err(Error::InvalidData(
                "authored Formula expression must be nonempty".into(),
            ));
        }
        let sheet = self.require_sheet(selector, "Formula insertion")?;
        if self.source.inner.sheets[sheet]
            .unique_entry_index(reference)?
            .is_some()
        {
            return Err(Error::refused(
                RefusalKind::UnsupportedTarget,
                "edit cells",
                "authored Formula target is occupied",
            ));
        }
        let parsed = crate::writer::FormulaTokenizer::new().tokenize(expression)?;
        let tokens = crate::writer::formula::encode_ptg_tokens(&parsed)?;
        if tokens.is_empty() || tokens.len() > MAX_FORMULA_TOKEN_BYTES {
            return Err(Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "authored Formula token bytes are empty or exceed the BIFF8 record limit",
            ));
        }
        self.stage_resource(ResourceChange::FormulaCell {
            sheet,
            reference,
            style,
            tokens,
            insert: true,
        })
    }

    /// Writes one formula cell from the writer's bounded BIFF8 tokenizer.
    ///
    /// An absent coordinate receives a new Formula record with the
    /// workbook's default XF, exactly as [`Self::insert_formula`] does. An
    /// existing scalar cell (including a packed `MulRk`/`MulBlank` member) is
    /// replaced in the same edit: its record is removed and the Formula is
    /// authored with the cell's existing XF. The authored record carries an
    /// empty cached result and `fAlwaysCalc`, so a spreadsheet application
    /// recalculates it; this crate does not evaluate the expression.
    /// Existing Formula records are refused because shared, array and table
    /// owners and string caches need a dependency-aware rewrite.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_xls::cell_values::{Reference, Selector, Storage};
    /// # let mut builder = litchi_xls::Builder::new();
    /// # let sheet = builder.add_worksheet("Data")?;
    /// # builder.write_number(sheet, 0, 0, 2.0)?;
    /// # builder.write_number(sheet, 0, 1, 7.0)?;
    /// # let bytes = builder.build()?.to_bytes()?;
    /// let source = litchi_xls::Package::from_bytes(&bytes)?;
    /// let mut edit = source.edit()?;
    /// let b1 = Reference::new(0, 1)?;
    /// edit.set_formula(Selector::Name("Data"), b1, "A1*3")?;
    /// let target = edit.commit()?.into_package().cell_values()?;
    /// let cell = target.worksheet(Selector::Position(Position::new(0)))?.unwrap().cell(b1)?.unwrap();
    /// assert_eq!(cell.storage(), Storage::Formula);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns a tokenizer, token-size, existing-formula, style, dependency,
    /// allocation, or transaction-bound refusal without partially staging.
    pub(crate) fn set_formula(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        expression: &str,
    ) -> Result<()> {
        let sheet = self.require_sheet(selector, "Formula write")?;
        let current = self.source.inner.sheets[sheet]
            .unique_entry_index(reference)?
            .map(|index| self.source.inner.sheets[sheet].entries[index].cell.clone());
        let Some(cell) = current else {
            return self.insert_formula(
                Selector::Position(Position::new(
                    self.source.inner.sheets[sheet].workbook_index,
                )),
                reference,
                expression,
            );
        };
        if cell.storage == Storage::Formula {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "existing Formula records require a dependency-aware formula rewrite",
            ));
        }
        if expression.is_empty() {
            return Err(Error::InvalidData(
                "authored Formula expression must be nonempty".into(),
            ));
        }
        let parsed = crate::writer::FormulaTokenizer::new().tokenize(expression)?;
        let tokens = crate::writer::formula::encode_ptg_tokens(&parsed)?;
        if tokens.is_empty() || tokens.len() > MAX_FORMULA_TOKEN_BYTES {
            return Err(Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "authored Formula token bytes are empty or exceed the BIFF8 record limit",
            ));
        }
        let mut staged = self.clone();
        staged.stage_structural(StructuralChange::Cell {
            sheet,
            reference,
            before: Some((cell.storage, cell.value.clone(), cell.style)),
            after: None,
        })?;
        staged.stage_resource(ResourceChange::FormulaCell {
            sheet,
            reference,
            style: cell.style,
            tokens,
            insert: true,
        })?;
        *self = staged;
        Ok(())
    }

    /// Whether this edit removes the authored Formula at one coordinate.
    fn staged_formula_removal(&self, sheet: usize, reference: Reference) -> bool {
        self.resource_changes.iter().any(|resource| {
            matches!(resource, ResourceChange::FormulaCell { insert: false, .. })
                && formula_resource_cell(resource) == Some((sheet, reference))
        })
    }

    /// Whether this edit removes the source cell at one coordinate.
    fn staged_removal(&self, sheet: usize, reference: Reference) -> bool {
        self.structural_index
            .cell(sheet, reference)
            .and_then(|index| self.structural_changes.get(index))
            .is_some_and(|change| {
                matches!(
                    change,
                    StructuralChange::Cell {
                        before: Some(_),
                        after: None,
                        ..
                    }
                )
            })
    }

    /// Replace a staged structural removal with its final cell state. A
    /// remove-then-reinsert sequence is one reversible structural operation,
    /// not an overlap between the fixed-width and structural families. This
    /// lets a coalesced signed edit prove an exact cancellation while keeping
    /// the ordinary overlap refusal for independent mixed operations.
    fn restore_removed_cell(
        &mut self,
        sheet: usize,
        reference: Reference,
        storage: Storage,
        value: Value,
    ) -> Result<()> {
        if self.resource_changes.iter().any(|resource| {
            matches!(
                resource,
                ResourceChange::FormulaCell {
                    sheet: resource_sheet,
                    reference: resource_reference,
                    insert: true,
                    ..
                } if *resource_sheet == sheet && *resource_reference == reference
            )
        }) {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "a cell with an authored Formula cannot receive another staged value",
            ));
        }
        let index = self
            .structural_index
            .cell(sheet, reference)
            .ok_or(Error::PatchConflict {
                operation: "edit cells",
                reason: "staged cell removal disappeared before reinsertion",
            })?;
        let change = self
            .structural_changes
            .get_mut(index)
            .ok_or(Error::PatchConflict {
                operation: "edit cells",
                reason: "staged cell removal index is stale",
            })?;
        let StructuralChange::Cell { before, after, .. } = change else {
            return Err(Error::PatchConflict {
                operation: "edit cells",
                reason: "staged cell removal changed family before reinsertion",
            });
        };
        if after.is_some() {
            return Err(Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "a cell has more than one structural outcome",
            ));
        }
        let style = before
            .as_ref()
            .map(|(_, _, style)| *style)
            .ok_or(Error::PatchConflict {
                operation: "edit cells",
                reason: "staged cell removal has no source state",
            })?;
        *after = Some((storage, value, style));
        Ok(())
    }

    /// Inserts a previously absent scalar cell with an existing XF resource.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for an occupied location, unsupported value,
    /// stale style resource, dependency closure, or finite transaction bound.
    /// Failed validation leaves this edit unchanged.
    pub(crate) fn insert_cell_with_style(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        value: Value,
        style: StyleIndex,
    ) -> Result<()> {
        let needs_candidate = match &value {
            Value::Text(text) => !self.has_shared_string(text),
            _ => false,
        };
        if needs_candidate {
            self.with_resource_rollback(|edit| {
                edit.insert_cell_with_style_in_place(selector, reference, value, style)
            })
        } else {
            self.insert_cell_with_style_in_place(selector, reference, value, style)
        }
    }

    fn insert_cell_with_style_in_place(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        value: Value,
        style: StyleIndex,
    ) -> Result<()> {
        self.require_effective_style(style)?;
        let sheet = self.require_sheet(selector, "cell insertion")?;
        if self.source.inner.sheets[sheet]
            .unique_entry_index(reference)?
            .is_some()
            && !self.staged_formula_removal(sheet, reference)
        {
            return Err(Error::refused(
                RefusalKind::UnsupportedTarget,
                "edit cells",
                "BIFF8 cell insertion target is occupied",
            ));
        }
        if let Value::Text(text) = &value
            && !self.has_shared_string(text)
        {
            self.stage_resource(ResourceChange::SharedString {
                text: text.clone(),
                insert: true,
            })?;
        }
        let storage = storage_for_new_value(&value, |text| self.has_shared_string(text))?;
        self.stage_structural(StructuralChange::Cell {
            sheet,
            reference,
            before: None,
            after: Some((storage, value, style)),
        })
    }

    /// Physically removes an existing BIFF8 cell record.
    ///
    /// Members of `MulRk` and `MulBlank` can be removed at either edge or
    /// interior position. Interior removal splits the packed record while
    /// preserving every neighboring cell's coordinate and XF; a final pair
    /// reduces to the corresponding standalone record. Formula ownership
    /// groups remain refused because their dependency closure is wider than a
    /// cell record.
    ///
    /// # Errors
    ///
    /// Returns a typed absence, ambiguity, dependency, or bound error.
    pub(crate) fn remove_cell(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
    ) -> Result<()> {
        let sheet = self.require_sheet(selector, "cell removal")?;
        let entry = self.source.inner.sheets[sheet]
            .unique_entry_index(reference)?
            .ok_or_else(|| {
                Error::refused(
                    RefusalKind::NotFound,
                    "edit cells",
                    "BIFF8 cell removal target is absent",
                )
            })?;
        let cell = &self.source.inner.sheets[sheet].entries[entry].cell;
        if cell.storage == Storage::Formula {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "formula records require a wider dependency rewrite",
            ));
        }
        self.stage_structural(StructuralChange::Cell {
            sheet,
            reference,
            before: Some((cell.storage, cell.value.clone(), cell.style)),
            after: None,
        })
    }

    /// Changes an existing cell to an existing workbook XF resource.
    ///
    /// # Errors
    ///
    /// Returns a typed absence, ambiguity, stale-resource, or bound error.
    pub(crate) fn set_style(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        style: StyleIndex,
    ) -> Result<()> {
        self.require_effective_style(style)?;
        let sheet = self.require_sheet(selector, "cell style edit")?;
        let entry = self.source.inner.sheets[sheet]
            .unique_entry_index(reference)?
            .ok_or_else(|| {
                Error::refused(
                    RefusalKind::NotFound,
                    "edit cells",
                    "BIFF8 style target is absent",
                )
            })?;
        let cell = &self.source.inner.sheets[sheet].entries[entry].cell;
        self.stage_structural(StructuralChange::Cell {
            sheet,
            reference,
            before: Some((cell.storage, cell.value.clone(), cell.style)),
            after: Some((cell.storage, cell.value.clone(), style)),
        })
    }

    /// Changes one cell's alignment, border styles or fill.
    ///
    /// The cell's effective XF is rewritten with the requested fields; an
    /// identical effective XF is reused, otherwise one XF resource is
    /// authored in this edit. An absent coordinate gains a formatting-only
    /// `Blank` cell whose base is the workbook's default cell XF. An empty or
    /// ineffective change is an exact no-op.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_xls::cell_values::{FormatChange, Reference, Selector};
    /// use litchi_xls::HorizontalAlignment;
    /// # let mut builder = litchi_xls::Builder::new();
    /// # let sheet = builder.add_worksheet("Data")?;
    /// # builder.write_number(sheet, 0, 0, 1.0)?;
    /// # let bytes = builder.build()?.to_bytes()?;
    /// let source = litchi_xls::Package::from_bytes(&bytes)?;
    /// let mut edit = source.edit()?;
    /// let mut change = FormatChange::default();
    /// change.horizontal = Some(HorizontalAlignment::Right);
    /// edit.set_cell_format(Selector::Position(Position::new(0)), Reference::new(0, 0)?, &change)?;
    /// assert!(edit.commit()?.diagnostics().changed());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns a typed indent, palette-color, selector, formula-dependency,
    /// style-resource or bound refusal without staging.
    pub(crate) fn set_cell_format(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        change: &FormatChange,
    ) -> Result<()> {
        if change.is_empty() {
            return Ok(());
        }
        let sheet = self.require_sheet(selector, "cell format")?;
        let current = self.source.inner.sheets[sheet]
            .unique_entry_index(reference)?
            .map(|index| self.source.inner.sheets[sheet].entries[index].cell.style);
        let base = match current {
            Some(style) => style,
            None => scalar_default_style(&self.source)?,
        };
        let payload = self
            .effective_xf_payload(base)
            .ok_or_else(|| {
                Error::refused(
                    RefusalKind::NotFound,
                    "edit cells",
                    "cell XF resource is absent",
                )
            })?
            .to_vec();
        let palette = if change.fill_foreground.is_some() || change.fill_background.is_some() {
            format::workbook_palette(&self.source.inner.workbook_stream)?
        } else {
            crate::Palette::default()
        };
        let target = format::apply(&payload, change, &palette)?;
        if target == payload {
            return Ok(());
        }
        let mut staged = self.clone();
        let style = match staged.effective_style_with_payload(&target) {
            Some(style) => style,
            None => staged.author_style(&target)?,
        };
        match current {
            Some(_) => staged.set_style(
                Selector::Position(Position::new(
                    staged.source.inner.sheets[sheet].workbook_index,
                )),
                reference,
                style,
            )?,
            None => staged.insert_cell_with_style(
                Selector::Position(Position::new(
                    staged.source.inner.sheets[sheet].workbook_index,
                )),
                reference,
                Value::Blank,
                style,
            )?,
        }
        *self = staged;
        Ok(())
    }

    /// An effective cell XF whose payload equals `payload`.
    fn effective_style_with_payload(&self, payload: &[u8]) -> Option<StyleIndex> {
        (0..self.effective_xf_count())
            .filter_map(|index| u16::try_from(index).ok())
            .map(StyleIndex)
            .find(|style| self.effective_xf_payload(*style) == Some(payload))
    }

    /// Duplicates an existing XF into a newly authored workbook resource.
    ///
    /// The returned index can be used by later cell insert/style operations
    /// in this transaction. Exact payload identity makes the operation
    /// reversible and dependency-checkable during transfer.
    ///
    /// # Errors
    ///
    /// Returns a stale source resource, index overflow, allocation, or bound
    /// error without staging a partial resource.
    pub(crate) fn duplicate_style(&mut self, source: StyleIndex) -> Result<StyleIndex> {
        self.require_effective_style(source)?;
        let payload = self
            .effective_xf_payload(source)
            .ok_or_else(|| {
                Error::refused(
                    RefusalKind::NotFound,
                    "edit cells",
                    "source XF resource is absent",
                )
            })?
            .to_vec();
        self.author_style(&payload)
    }

    /// Authors one validated BIFF8 XF payload as the next workbook resource.
    ///
    /// The payload is the exact 20-byte body of an XF record. Its six-byte
    /// font, number-format, protection, kind, and parent-style dependency prefix
    /// must match an effective XF; the remaining formatting bytes are retained
    /// exactly.
    ///
    /// # Errors
    ///
    /// Returns a payload-length, index-overflow, allocation, or transaction
    /// bound error without staging a partial resource.
    pub(crate) fn author_style(&mut self, payload: &[u8]) -> Result<StyleIndex> {
        validate_xf_payload(payload)?;
        let mut retained = Vec::new();
        retained
            .try_reserve_exact(payload.len())
            .map_err(|_error| Error::Allocation("retaining authored XF payload"))?;
        retained.extend_from_slice(payload);
        let pending = self
            .resource_changes
            .iter()
            .filter(|change| matches!(change, ResourceChange::ExtendedFormat { insert: true, .. }))
            .count();
        let index = self
            .source
            .inner
            .xf_records
            .len()
            .checked_add(pending)
            .and_then(|value| u16::try_from(value).ok())
            .map(StyleIndex)
            .ok_or_else(|| {
                Error::refused(
                    RefusalKind::InvalidValue,
                    "edit cells",
                    "new XF index exceeds u16",
                )
            })?;
        self.stage_resource(ResourceChange::ExtendedFormat {
            index,
            payload: retained,
            insert: true,
        })?;
        Ok(index)
    }

    /// Runs one multi-stage operation whose only staging before its final
    /// fallible step is SST resources; on failure the resources it staged
    /// are dropped, leaving the edit exactly as before.
    ///
    /// Every staging step checks its refusals before it pushes, so the
    /// resource tail is the only state a failed operation can have changed.
    /// This replaces cloning the whole edit per new text, which made each
    /// text insertion O(staged changes).
    fn with_resource_rollback(
        &mut self,
        stage: impl FnOnce(&mut Self) -> Result<()>,
    ) -> Result<()> {
        let resources = self.resource_changes.len();
        let (changes, structural) = (self.changes.len(), self.structural_changes.len());
        let result = stage(self);
        if result.is_err() {
            debug_assert_eq!(
                (self.changes.len(), self.structural_changes.len()),
                (changes, structural),
                "a failed staging step pushed a cell change"
            );
            self.resource_changes.truncate(resources);
        }
        result
    }

    /// Whether `text` is present in the effective SST.
    ///
    /// This is the membership of [`Self::effective_shared_strings`] without
    /// materializing it: a staged insertion of `text` makes it present, and
    /// otherwise it is present while the source holds more occurrences than
    /// the staged removals of the same text (each removal drops one source
    /// occurrence).  The source occurrences come from a per-edit hash index,
    /// so a text insertion costs O(staged resources) instead of cloning the
    /// source table and fingerprinting every staged string.
    fn has_shared_string(&self, text: &str) -> bool {
        let mut removed = 0_usize;
        for change in &self.resource_changes {
            match change {
                ResourceChange::SharedString {
                    text: candidate,
                    insert,
                }
                | ResourceChange::RichSharedString {
                    text: candidate,
                    insert,
                    ..
                } if candidate == text => {
                    if *insert {
                        return true;
                    }
                    removed = removed.saturating_add(1);
                },
                ResourceChange::SharedString { .. }
                | ResourceChange::RichSharedString { .. }
                | ResourceChange::ExtendedFormat { .. }
                | ResourceChange::FormulaCell { .. } => {},
            }
        }
        self.source_strings
            .get_or_init(|| SourceStringIndex::new(&self.source.inner.shared_strings))
            .count_at_most(
                &self.source.inner.shared_strings,
                text,
                removed.saturating_add(1),
            )
            > removed
    }

    fn has_rich_shared_string(
        &self,
        text: &str,
        formatting_runs: &[crate::records::SharedStringFormatRun],
    ) -> bool {
        self.source
            .inner
            .shared_strings
            .iter()
            .enumerate()
            .any(|(index, candidate)| {
                candidate == text
                    && self
                        .source
                        .inner
                        .shared_string_properties
                        .get(index)
                        .and_then(Option::as_deref)
                        .is_some_and(|properties| {
                            properties.phonetic.is_none()
                                && properties.formatting_runs.as_slice() == formatting_runs
                        })
            })
            || self.resource_changes.iter().any(|change| {
                matches!(
                    change,
                    ResourceChange::RichSharedString {
                        text: candidate,
                        formatting_runs: candidate_runs,
                        insert: true,
                    } if candidate == text && candidate_runs == formatting_runs
                )
            })
    }

    fn effective_shared_strings(&self) -> Vec<String> {
        let mut strings = self.source.inner.shared_strings.as_ref().clone();
        for text in self
            .resource_changes
            .iter()
            .filter_map(|change| match change {
                ResourceChange::SharedString {
                    text,
                    insert: false,
                }
                | ResourceChange::RichSharedString {
                    text,
                    insert: false,
                    ..
                } => Some(text),
                ResourceChange::SharedString { insert: true, .. }
                | ResourceChange::RichSharedString { insert: true, .. }
                | ResourceChange::ExtendedFormat { .. }
                | ResourceChange::FormulaCell { .. } => None,
            })
        {
            if let Some(index) = strings.iter().position(|candidate| candidate == text) {
                strings.remove(index);
            }
        }
        let mut inserted: Vec<_> = self
            .resource_changes
            .iter()
            .filter_map(|change| match change {
                ResourceChange::SharedString { text, insert: true }
                | ResourceChange::RichSharedString {
                    text, insert: true, ..
                } => Some((resource_target(change), text.clone())),
                ResourceChange::SharedString { insert: false, .. }
                | ResourceChange::RichSharedString { insert: false, .. }
                | ResourceChange::ExtendedFormat { .. }
                | ResourceChange::FormulaCell { .. } => None,
            })
            .collect();
        inserted.sort_by(|left, right| left.0.cmp(&right.0));
        strings.extend(inserted.into_iter().map(|(_, text)| text));
        strings
    }

    fn effective_xf_payload(&self, style: StyleIndex) -> Option<&[u8]> {
        let index = usize::from(style.get());
        self.source
            .inner
            .xf_records
            .get(index)
            .map(Vec::as_slice)
            .or_else(|| {
                self.resource_changes
                    .iter()
                    .find_map(|change| match change {
                        ResourceChange::ExtendedFormat {
                            index: candidate,
                            payload,
                            insert: true,
                        } if *candidate == style => Some(payload.as_slice()),
                        ResourceChange::SharedString { .. }
                        | ResourceChange::RichSharedString { .. }
                        | ResourceChange::ExtendedFormat { .. }
                        | ResourceChange::FormulaCell { .. } => None,
                    })
            })
    }

    fn require_effective_style(&self, style: StyleIndex) -> Result<()> {
        self.effective_xf_payload(style)
            .map(|_| ())
            .ok_or(Error::Internal("XF index is not staged"))
    }

    fn effective_xf_count(&self) -> usize {
        let inserted = self
            .resource_changes
            .iter()
            .filter(|change| matches!(change, ResourceChange::ExtendedFormat { insert: true, .. }))
            .count();
        let removed = self
            .resource_changes
            .iter()
            .filter(|change| matches!(change, ResourceChange::ExtendedFormat { insert: false, .. }))
            .count();
        self.source
            .inner
            .xf_records
            .len()
            .saturating_add(inserted)
            .saturating_sub(removed)
    }

    fn has_xf_dependency(&self, payload: &[u8]) -> bool {
        let Some(prefix) = payload.get(..6) else {
            return false;
        };
        self.source
            .inner
            .xf_records
            .iter()
            .any(|candidate| candidate.get(..6) == Some(prefix))
            || self.resource_changes.iter().any(|candidate| {
                matches!(
                    candidate,
                    ResourceChange::ExtendedFormat {
                        payload: existing,
                        insert: true,
                        ..
                    } if existing.get(..6) == Some(prefix)
                )
            })
    }

    fn stage_resource(&mut self, change: ResourceChange) -> Result<()> {
        self.require_no_tab_change()?;
        if self
            .resource_changes
            .iter()
            .any(|existing| resource_changes_equal(existing, &change))
        {
            return Ok(());
        }
        match &change {
            ResourceChange::SharedString { text, insert } => {
                if self.has_shared_string(text) == *insert {
                    return Err(Error::refused(
                        RefusalKind::InvalidValue,
                        "edit cells",
                        "shared-string resource outcome is already present",
                    ));
                }
                if *insert {
                    structural::certify_sst_insertion(&change)?;
                }
            },
            ResourceChange::RichSharedString {
                text,
                formatting_runs,
                insert,
            } => {
                let present = self.has_rich_shared_string(text, formatting_runs);
                if present == *insert {
                    return Err(Error::refused(
                        RefusalKind::InvalidValue,
                        "edit cells",
                        "rich shared-string resource outcome is already present",
                    ));
                }
                if *insert {
                    if self.has_shared_string(text) {
                        return Err(Error::refused(
                            RefusalKind::InvalidValue,
                            "edit cells",
                            "rich SST text is already present in the effective SST",
                        ));
                    }
                    validate_rich_text(self, text, formatting_runs)?;
                    structural::certify_sst_insertion(&change)?;
                }
            },
            ResourceChange::ExtendedFormat {
                index,
                payload,
                insert,
            } => {
                let count = self.effective_xf_count();
                if *insert {
                    if usize::from(index.get()) != count {
                        return Err(Error::refused(
                            RefusalKind::UnsupportedOperation,
                            "edit cells",
                            "new XF index is not the next effective resource",
                        ));
                    }
                    validate_xf_payload(payload)?;
                    if !self.has_xf_dependency(payload) {
                        return Err(Error::refused(
                            RefusalKind::UnsupportedOperation,
                            "edit cells",
                            "new XF resource dependencies are not effective",
                        ));
                    }
                } else if usize::from(index.get()).checked_add(1) != Some(count)
                    || self.effective_xf_payload(*index) != Some(payload.as_slice())
                {
                    return Err(Error::refused(
                        RefusalKind::UnsupportedOperation,
                        "edit cells",
                        "XF removal is not the exact effective tail resource",
                    ));
                }
            },
            ResourceChange::FormulaCell {
                sheet,
                reference,
                style,
                tokens,
                insert,
            } => {
                self.require_effective_style(*style)?;
                // A staged removal of the same coordinate is the one cell
                // operation an authored Formula may follow: together they
                // replace a scalar cell with a formula.
                let replaced = self.staged_removal(*sheet, *reference);
                if self.structural_index.axis(*sheet).is_some()
                    || (self.structural_index.cell(*sheet, *reference).is_some() && !replaced)
                {
                    return Err(Error::refused(
                        RefusalKind::UnsupportedOperation,
                        "edit cells",
                        "authored Formula cannot overlap another structural worksheet edit",
                    ));
                }
                if tokens.is_empty() || tokens.len() > MAX_FORMULA_TOKEN_BYTES {
                    return Err(Error::InvalidData(
                        "authored Formula token size is outside its BIFF8 record".into(),
                    ));
                }
                if self
                    .resource_changes
                    .iter()
                    .any(|resource| formula_resource_cell(resource) == Some((*sheet, *reference)))
                {
                    return Err(Error::refused(
                        RefusalKind::UnsupportedTarget,
                        "edit cells",
                        "authored Formula target already has a staged owner",
                    ));
                }
                let sheet_data =
                    self.source
                        .inner
                        .sheets
                        .get(*sheet)
                        .ok_or(Error::PatchConflict {
                            operation: "edit cells",
                            reason: "authored Formula sheet dependency is stale",
                        })?;
                let current = sheet_data
                    .unique_entry_index(*reference)?
                    .map(|index| &sheet_data.entries[index]);
                if *insert && current.is_some() && !replaced {
                    return Err(Error::refused(
                        RefusalKind::UnsupportedTarget,
                        "edit cells",
                        "authored Formula insertion target is occupied",
                    ));
                }
                if !*insert {
                    let entry = current.ok_or_else(|| {
                        Error::refused(
                            RefusalKind::NotFound,
                            "edit cells",
                            "authored Formula inverse target is absent",
                        )
                    })?;
                    if !authored_formula_record_matches(
                        &self.source,
                        entry,
                        *reference,
                        *style,
                        tokens,
                    )? {
                        return Err(Error::PatchConflict {
                            operation: "edit cells",
                            reason: "authored Formula inverse record dependency is stale",
                        });
                    }
                }
            },
        }
        let observed = self
            .changes
            .len()
            .checked_add(self.structural_changes.len())
            .and_then(|value| value.checked_add(self.resource_changes.len()))
            .ok_or_else(|| Error::InvalidData("transaction resource size overflow".into()))?;
        if observed >= MAX_STAGED_CHANGES {
            return Err(Error::Limit {
                kind: crate::LimitKind::Operations,
                observed: observed.saturating_add(1) as u64,
                maximum: MAX_STAGED_CHANGES as u64,
            });
        }
        self.resource_changes
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("staging XLS resources"))?;
        self.resource_changes.push(change);
        Ok(())
    }

    /// Inserts empty row coordinates and moves every retained row below them.
    ///
    /// # Errors
    ///
    /// Returns a typed grid, dependency-closure, or finite-bound error.
    pub(crate) fn insert_rows(
        &mut self,
        selector: Selector<'_>,
        start: u16,
        count: u16,
    ) -> Result<()> {
        self.stage_rows(selector, start, count, true)
    }

    /// Deletes row coordinates and their cells, moving later rows upward.
    ///
    /// # Errors
    ///
    /// Returns a typed dependency or reversibility refusal. Deletion is
    /// durable only when the deleted coordinate span contains no cells.
    pub(crate) fn delete_rows(
        &mut self,
        selector: Selector<'_>,
        start: u16,
        count: u16,
    ) -> Result<()> {
        self.stage_rows(selector, start, count, false)
    }

    /// Inserts empty column coordinates and moves every retained cell right.
    ///
    /// # Errors
    ///
    /// Returns a typed grid, dependency-closure, or finite-bound error.
    pub(crate) fn insert_columns(
        &mut self,
        selector: Selector<'_>,
        start: u8,
        count: u8,
    ) -> Result<()> {
        self.stage_columns(selector, start, count, true)
    }

    /// Deletes column coordinates and their cells, moving later cells left.
    ///
    /// # Errors
    ///
    /// Returns a typed dependency or reversibility refusal. Deletion is
    /// durable only when the deleted coordinate span contains no cells.
    pub(crate) fn delete_columns(
        &mut self,
        selector: Selector<'_>,
        start: u8,
        count: u8,
    ) -> Result<()> {
        self.stage_columns(selector, start, count, false)
    }

    /// Renames one worksheet tab while preserving its `BoundSheet` identity.
    ///
    /// # Errors
    ///
    /// Returns a typed name, ambiguity, allocation, or finite-bound error.
    pub(crate) fn rename_sheet(&mut self, selector: Selector<'_>, name: &str) -> Result<()> {
        let sheet = self.require_sheet(selector, "worksheet rename")?;
        structural::validate_sheet_name(name)?;
        if self
            .source
            .inner
            .sheets
            .iter()
            .enumerate()
            .any(|(index, item)| index != sheet && caseless_eq(&item.name, name))
        {
            return Err(Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "worksheet name would be ambiguous",
            ));
        }
        let before = self.source.inner.sheets[sheet].name.clone();
        if before == name {
            return Ok(());
        }
        self.stage_structural(StructuralChange::RenameSheet {
            sheet,
            before,
            after: name.to_string(),
        })
    }

    /// Sets one row's height in twips (1/20 point) and marks it custom.
    ///
    /// The row need not hold cells: a missing `Row` record is created. Row
    /// layout composes with cell edits of the same worksheet but not with an
    /// axis shift.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_xls::cell_values::{Selector};
    /// # let bytes = litchi_xls::Package::new()?.to_bytes()?;
    /// let source = litchi_xls::Package::from_bytes(&bytes)?;
    /// let mut edit = source.edit()?;
    /// edit.set_row_height(Selector::Position(Position::new(0)), 4, 600)?;
    /// edit.set_row_hidden(Selector::Position(Position::new(0)), 5, true)?;
    /// let committed = edit.commit()?;
    /// assert!(committed.diagnostics().changed());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns a typed range, selector, axis-overlap, or bound refusal
    /// without staging.
    pub(crate) fn set_row_height(
        &mut self,
        selector: Selector<'_>,
        row: u16,
        twips: u16,
    ) -> Result<()> {
        if twips == 0 || twips > layout::MAX_ROW_TWIPS {
            return Err(Error::InvalidData(
                "row height must be 1 through 8,180 twips (409 pt)".into(),
            ));
        }
        self.stage_row_layout(selector, row, |state| {
            state.present = true;
            state.height = twips;
            state.custom = true;
        })
    }

    /// Hides or shows one row, keeping its height.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, axis-overlap, or bound refusal without
    /// staging.
    pub(crate) fn set_row_hidden(
        &mut self,
        selector: Selector<'_>,
        row: u16,
        hidden: bool,
    ) -> Result<()> {
        self.stage_row_layout(selector, row, |state| {
            if state.present || hidden {
                state.present = true;
                state.hidden = hidden;
            }
        })
    }

    /// Sets one column's width in 1/256 of a character and marks it user-set.
    /// A `ColInfo` range covering other columns is split around it.
    ///
    /// # Errors
    ///
    /// Returns a typed range, selector, record-layout, axis-overlap, or bound
    /// refusal without staging.
    pub(crate) fn set_column_width(
        &mut self,
        selector: Selector<'_>,
        column: u8,
        width: u16,
    ) -> Result<()> {
        if width > layout::MAX_COLUMN_WIDTH {
            return Err(Error::InvalidData(
                "column width must be at most 255 characters".into(),
            ));
        }
        self.stage_column_layout(selector, column, |state| {
            state.present = true;
            state.width = width;
            state.user_set = true;
        })
    }

    /// Hides or shows one column, keeping its width.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, record-layout, axis-overlap, or bound
    /// refusal without staging.
    pub(crate) fn set_column_hidden(
        &mut self,
        selector: Selector<'_>,
        column: u8,
        hidden: bool,
    ) -> Result<()> {
        let default_width = self.default_column_width(selector)?;
        self.stage_column_layout(selector, column, |state| {
            if state.present || hidden {
                if !state.present {
                    state.width = default_width;
                }
                state.present = true;
                state.hidden = hidden;
            }
        })
    }

    /// The source substream of one editable worksheet.
    fn sheet_stream(&self, sheet: usize) -> Result<&[u8]> {
        let workbook_index = self.source.inner.sheets[sheet].workbook_index;
        let catalog = tabs::catalog(&self.source.inner.workbook_stream)?;
        let tab = catalog.get(workbook_index).ok_or(Error::PatchConflict {
            operation: "edit cells",
            reason: "worksheet tab position is stale",
        })?;
        self.source
            .inner
            .workbook_stream
            .get(tab.start..tab.end)
            .ok_or_else(|| Error::InvalidData("worksheet substream is outside Workbook".into()))
    }

    /// The `DefColWidth` width in 1/256 character, for a newly covered column.
    fn default_column_width(&self, selector: Selector<'_>) -> Result<u16> {
        let sheet = self.require_sheet(selector, "column layout")?;
        let stream = self.sheet_stream(sheet)?;
        let mut width = 8_u16;
        for item in Records::new(stream, litchi_biff::Limits::default())? {
            let record = item?;
            match record.kind().get() {
                0x0055 if record.payload().len() >= 2 => {
                    width = binary::read_u16_le_at(record.payload(), 0)?;
                    break;
                },
                EOF => break,
                _ => {},
            }
        }
        Ok(width.saturating_mul(256))
    }

    fn require_layout_sheet(&self, selector: Selector<'_>, context: &str) -> Result<usize> {
        self.require_no_tab_change()?;
        let sheet = self.require_sheet(selector, context)?;
        if self.structural_index.axis(sheet).is_some() {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "a row or column layout change cannot follow an axis shift of its worksheet",
            ));
        }
        Ok(sheet)
    }

    fn stage_row_layout(
        &mut self,
        selector: Selector<'_>,
        row: u16,
        update: impl FnOnce(&mut layout::RowState),
    ) -> Result<()> {
        let sheet = self.require_layout_sheet(selector, "row layout")?;
        let existing = self.layout_changes.iter().position(|change| {
            matches!(change, layout::LayoutChange::Row { sheet: owner, row: target, .. } if *owner == sheet && *target == row)
        });
        match existing {
            Some(index) => {
                if let layout::LayoutChange::Row { after, .. } = &mut self.layout_changes[index] {
                    update(after);
                }
                Ok(())
            },
            None => {
                let before = layout::row_state(self.sheet_stream(sheet)?, row)?;
                let mut after = before;
                update(&mut after);
                self.push_layout(layout::LayoutChange::Row {
                    sheet,
                    row,
                    before,
                    after,
                })
            },
        }
    }

    fn stage_column_layout(
        &mut self,
        selector: Selector<'_>,
        column: u8,
        update: impl FnOnce(&mut layout::ColumnState),
    ) -> Result<()> {
        let sheet = self.require_layout_sheet(selector, "column layout")?;
        let existing = self.layout_changes.iter().position(|change| {
            matches!(change, layout::LayoutChange::Column { sheet: owner, column: target, .. } if *owner == sheet && *target == column)
        });
        match existing {
            Some(index) => {
                if let layout::LayoutChange::Column { after, .. } = &mut self.layout_changes[index]
                {
                    update(after);
                }
                Ok(())
            },
            None => {
                let before = layout::column_state(self.sheet_stream(sheet)?, column)?;
                let mut after = before;
                update(&mut after);
                self.push_layout(layout::LayoutChange::Column {
                    sheet,
                    column,
                    before,
                    after,
                })
            },
        }
    }

    fn push_layout(&mut self, change: layout::LayoutChange) -> Result<()> {
        let observed = self
            .changes
            .len()
            .saturating_add(self.structural_changes.len())
            .saturating_add(self.resource_changes.len())
            .saturating_add(self.layout_changes.len());
        if observed >= MAX_STAGED_CHANGES {
            return Err(Error::Limit {
                kind: crate::LimitKind::Operations,
                observed: observed.saturating_add(1) as u64,
                maximum: MAX_STAGED_CHANGES as u64,
            });
        }
        self.layout_changes
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("staging XLS layout changes"))?;
        self.layout_changes.push(change);
        Ok(())
    }

    /// Reorders the complete workbook tab catalog.
    ///
    /// `order` names every tab exactly once, chart, macro and dialog tabs
    /// included: the tab selected by `order[i]` moves to catalog position
    /// `i`. Names compare by the workbook's Unicode caseless sheet identity.
    /// The identity permutation is an exact no-op. Workbook-global tab
    /// addresses (`Window1`, `RRTabId`, internal `ExternSheet` XTI ranges,
    /// sheet-scoped names) and relocated worksheet `Index`/`ExtSST` offsets
    /// are rewritten in the same closure.
    ///
    /// A tab-structure change must be the only operation of its edit.
    ///
    /// ```
    /// use litchi_xls::cell_values::{Selector};
    /// # let mut builder = litchi_xls::Builder::new();
    /// # builder.add_worksheet("First")?;
    /// # builder.add_worksheet("Second")?;
    /// # let bytes = builder.build()?.to_bytes()?;
    /// let source = litchi_xls::Package::from_bytes(&bytes)?;
    /// let mut edit = source.edit()?;
    /// edit.reorder_sheets(&[Selector::Name("Second"), Selector::Name("first")])?;
    /// let target = edit.commit()?.into_package().cell_values()?;
    /// let names: Vec<_> = target.worksheets().map(|sheet| sheet.name().to_owned()).collect();
    /// assert_eq!(names, ["Second", "First"]);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns a typed selector, ambiguity, protection, dependency-closure
    /// (for example a 3-D reference range made noncontiguous), exclusivity or
    /// allocation refusal without staging.
    pub(crate) fn reorder_sheets(&mut self, order: &[Selector<'_>]) -> Result<()> {
        self.require_tab_edit(false)?;
        let catalog = tabs::catalog(&self.source.inner.workbook_stream)?;
        if order.len() != catalog.len() {
            return Err(Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "a tab reorder must name every catalog tab exactly once",
            ));
        }
        let mut positions = Vec::new();
        positions
            .try_reserve_exact(order.len())
            .map_err(|_error| Error::Allocation("staging an XLS tab reorder"))?;
        for selector in order {
            positions.push(resolve_tab(&catalog, *selector)?);
        }
        if positions.iter().enumerate().all(|(new, old)| new == *old) {
            return Ok(());
        }
        let kind = tabs::TabChangeKind::Reorder { order: positions };
        self.stage_tab_change(&catalog, kind, None)
    }

    /// Inserts one empty ordinary worksheet at a complete-catalog position.
    ///
    /// `position` may equal the tab count to append. The name must satisfy
    /// the BIFF8 sheet-name rule and must not equal another tab's name under
    /// Unicode caseless comparison. The inserted worksheet is visible and is
    /// not selected; 3-D reference ranges that would gain it are refused.
    ///
    /// ```
    /// use litchi_core::Position;
    ///
    /// # let bytes = litchi_xls::Package::new()?.to_bytes()?;
    /// let source = litchi_xls::Package::from_bytes(&bytes)?;
    /// let mut edit = source.edit()?;
    /// edit.insert_sheet(Position::new(0), "Summary")?;
    /// let target = edit.commit()?.into_package().cell_values()?;
    /// assert_eq!(target.worksheets().next().map(|sheet| sheet.name()), Some("Summary"));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns a typed name, uniqueness, position, protection, macro-project,
    /// dependency-closure, exclusivity or allocation refusal without staging.
    pub(crate) fn insert_sheet(&mut self, position: Position, name: &str) -> Result<()> {
        self.require_tab_edit(true)?;
        structural::validate_sheet_name(name)?;
        let catalog = tabs::catalog(&self.source.inner.workbook_stream)?;
        if catalog
            .iter()
            .any(|tab| crate::sheet_name::equivalent(&tab.name, name))
        {
            return Err(Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "the inserted worksheet name would be ambiguous",
            ));
        }
        if position.get() > catalog.len() {
            return Err(Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "worksheet insertion position is outside the tab catalog",
            ));
        }
        let template = tabs::empty_worksheet_template()?;
        let kind = tabs::TabChangeKind::Insert {
            position: position.get(),
            name: name.to_string(),
        };
        self.stage_tab_change(&catalog, kind, Some(&template))
    }

    /// Removes one ordinary worksheet tab and its complete substream.
    ///
    /// Removal is refused while any workbook dependency still addresses the
    /// tab: an internal 3-D reference range, a sheet-scoped defined name, a
    /// pivot data connection, or a drawing/pivot owner on the worksheet. The
    /// workbook must retain a visible tab. When the active tab is removed the
    /// first remaining visible tab becomes active and selected.
    ///
    /// ```
    /// use litchi_xls::cell_values::{Selector};
    /// # let mut builder = litchi_xls::Builder::new();
    /// # builder.add_worksheet("Keep")?;
    /// # builder.add_worksheet("Drop")?;
    /// # let bytes = builder.build()?.to_bytes()?;
    /// let source = litchi_xls::Package::from_bytes(&bytes)?;
    /// let mut edit = source.edit()?;
    /// edit.remove_sheet(Selector::Name("Drop"))?;
    /// let target = edit.commit()?.into_package().cell_values()?;
    /// assert_eq!(target.worksheet_count(), 1);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns a typed selector, dependency, last-tab, last-visible-tab,
    /// protection, macro-project, exclusivity or allocation refusal without
    /// staging.
    pub(crate) fn remove_sheet(&mut self, selector: Selector<'_>) -> Result<()> {
        self.require_tab_edit(true)?;
        let catalog = tabs::catalog(&self.source.inner.workbook_stream)?;
        let position = resolve_tab(&catalog, selector)?;
        let tab = &catalog[position];
        tabs::certify_removal(&self.source.inner.workbook_stream, &catalog, tab)?;
        let kind = tabs::TabChangeKind::Remove {
            position,
            name: tab.name.clone(),
            fingerprint: tabs::fingerprint(&self.source.inner.workbook_stream, tab),
        };
        self.stage_tab_change(&catalog, kind, None)
    }

    fn require_no_tab_change(&self) -> Result<()> {
        if self.tab_change.is_some() {
            return Err(Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "a worksheet tab-structure change must be the only operation of its edit",
            ));
        }
        Ok(())
    }

    fn require_tab_edit(&self, creates_or_destroys: bool) -> Result<()> {
        self.require_no_tab_change()?;
        if !self.changes.is_empty()
            || !self.structural_changes.is_empty()
            || !self.resource_changes.is_empty()
            || !self.layout_changes.is_empty()
        {
            return Err(Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "a worksheet tab-structure change must be the only operation of its edit",
            ));
        }
        if self.source.inner.source_policy.protection == SourceProtectionPolicy::WorkbookOrShared {
            return Err(Error::refused(
                RefusalKind::Protected,
                "edit cells",
                "protected or shared workbook structure cannot change its tabs",
            ));
        }
        if creates_or_destroys && !self.source.inner.source_policy.macro_free_workbook {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "adding or removing tabs would desynchronize the workbook's macro project",
            ));
        }
        Ok(())
    }

    fn stage_tab_change(
        &mut self,
        catalog: &[tabs::Tab],
        kind: tabs::TabChangeKind,
        template: Option<&tabs::Template>,
    ) -> Result<()> {
        let workbook = tabs::rewrite(&self.source.inner.workbook_stream, &kind, template)?;
        let source_names = catalog.iter().map(|tab| tab.name.clone()).collect();
        let target_names = tabs::catalog(&workbook)?
            .into_iter()
            .map(|tab| tab.name)
            .collect();
        self.tab_change = Some(Box::new(tabs::TabChange {
            kind,
            source_names,
            target_names,
            workbook,
        }));
        Ok(())
    }

    fn require_sheet(&self, selector: Selector<'_>, context: &str) -> Result<usize> {
        self.source
            .resolve_sheet(selector)?
            .ok_or_else(|| Error::WorksheetNotFound(context.into()))
    }

    fn stage_rows(
        &mut self,
        selector: Selector<'_>,
        start: u16,
        count: u16,
        insert: bool,
    ) -> Result<()> {
        if count == 0 {
            return Err(Error::InvalidData(
                "row operation count must be nonzero".into(),
            ));
        }
        if insert && start.checked_add(count).is_none() {
            return Err(Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "row insertion exceeds the BIFF8 grid",
            ));
        }
        let sheet = self.require_sheet(selector, "row operation")?;
        structural::certify_shift(
            &self.source,
            sheet,
            structural::AxisShift::Rows {
                start,
                count,
                insert,
            },
        )?;
        self.stage_structural(StructuralChange::Rows {
            sheet,
            start,
            count,
            insert,
        })
    }

    fn stage_columns(
        &mut self,
        selector: Selector<'_>,
        start: u8,
        count: u8,
        insert: bool,
    ) -> Result<()> {
        if count == 0 {
            return Err(Error::InvalidData(
                "column operation count must be nonzero".into(),
            ));
        }
        if insert && start.checked_add(count).is_none() {
            return Err(Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "column insertion exceeds the BIFF8 grid",
            ));
        }
        let sheet = self.require_sheet(selector, "column operation")?;
        structural::certify_shift(
            &self.source,
            sheet,
            structural::AxisShift::Columns {
                start,
                count,
                insert,
            },
        )?;
        self.stage_structural(StructuralChange::Columns {
            sheet,
            start,
            count,
            insert,
        })
    }

    fn stage_structural(&mut self, change: StructuralChange) -> Result<()> {
        self.require_no_tab_change()?;
        if matches!(
            change,
            StructuralChange::Rows { .. } | StructuralChange::Columns { .. }
        ) && self
            .layout_changes
            .iter()
            .any(|layout| layout.sheet() == operation_sheet_index(&change))
        {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "an axis shift cannot follow a row or column layout change of its worksheet",
            ));
        }
        let sheet = operation_sheet_index(&change);
        // An authored Formula is placed after the worksheet's cell operations
        // in one row-block rebuild. Axis shifts would move it, and a cell
        // operation at its own coordinate would race it; operations at other
        // coordinates compose.
        if self.resource_changes.iter().any(|resource| {
            formula_resource_cell(resource).is_some_and(|(resource_sheet, reference)| {
                resource_sheet == sheet
                    && match &change {
                        StructuralChange::Rows { .. } | StructuralChange::Columns { .. } => true,
                        // A scalar may take the place of a removed authored
                        // Formula: removals are rebuilt before cell inserts.
                        StructuralChange::Cell { before: None, .. }
                            if matches!(
                                resource,
                                ResourceChange::FormulaCell { insert: false, .. }
                            ) =>
                        {
                            false
                        },
                        StructuralChange::Cell {
                            reference: target, ..
                        } => *target == reference,
                        StructuralChange::RenameSheet { .. } => false,
                    }
            })
        }) {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "structural worksheet edit overlaps an authored Formula",
            ));
        }
        if self.structural_index.overlaps(&change) {
            return Err(Error::refused(
                RefusalKind::InvalidValue,
                "edit cells",
                "overlapping structural operations must be prepared and joined separately",
            ));
        }
        let fixed_overlap = match &change {
            StructuralChange::Cell { reference, .. } => {
                self.fixed_index.first(sheet, *reference).is_some()
            },
            StructuralChange::Rows { .. } | StructuralChange::Columns { .. } => {
                self.fixed_index.has_sheet(sheet)
            },
            StructuralChange::RenameSheet { .. } => false,
        };
        if fixed_overlap {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "fixed-width and structural operations overlap in one worksheet",
            ));
        }
        let observed = self
            .changes
            .len()
            .checked_add(self.structural_changes.len())
            .and_then(|value| value.checked_add(self.resource_changes.len()))
            .ok_or_else(|| Error::InvalidData("cell transaction size overflow".into()))?;
        if observed >= MAX_STAGED_CHANGES {
            return Err(Error::Limit {
                kind: crate::LimitKind::Operations,
                observed: observed.saturating_add(1) as u64,
                maximum: MAX_STAGED_CHANGES as u64,
            });
        }
        self.structural_changes
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("staging structural XLS changes"))?;
        self.structural_index.reserve_one(&change)?;
        let index = self.structural_changes.len();
        self.structural_changes.push(change);
        self.structural_index
            .insert(&self.structural_changes[index], index);
        Ok(())
    }

    fn stage(
        &mut self,
        sheet_index: usize,
        entry: usize,
        storage: Storage,
        value: Value,
    ) -> Result<()> {
        self.require_no_tab_change()?;
        let reference = self.source.inner.sheets[sheet_index].entries[entry]
            .cell
            .reference;
        if self.structural_index.overlaps(&StructuralChange::Cell {
            sheet: sheet_index,
            reference,
            before: None,
            after: None,
        }) {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "fixed-width and structural operations overlap in one worksheet",
            ));
        }
        if let Some(change_index) = self.fixed_index.first(sheet_index, reference) {
            let change = &mut self.changes[change_index];
            change.storage = storage;
            change.value = value;
        } else {
            let observed = self
                .changes
                .len()
                .saturating_add(self.structural_changes.len())
                .saturating_add(self.resource_changes.len());
            if observed >= MAX_STAGED_CHANGES {
                return Err(Error::Limit {
                    kind: crate::LimitKind::Operations,
                    observed: observed.saturating_add(1) as u64,
                    maximum: MAX_STAGED_CHANGES as u64,
                });
            }
            self.changes
                .try_reserve(1)
                .map_err(|_error| Error::Allocation("staging Number cell changes"))?;
            self.fixed_index.reserve_one()?;
            let change_index = self.changes.len();
            self.changes.push(Change {
                sheet: sheet_index,
                entry,
                reference,
                storage,
                value,
            });
            self.fixed_index.insert(StagedCell {
                sheet: sheet_index,
                reference,
                change: change_index,
            });
        }
        Ok(())
    }

    /// Publishes unchanged-family numeric edits through a validated,
    /// source-backed CFB splice plan.
    ///
    /// Only existing `Number`, `RK`, and `MulRk` value fields are eligible.
    /// The transaction submits the exact source-relative ranges recorded by
    /// the private `Entry` inventory to
    /// [`SourceBackedOverlayPublisher`]; it never renders a replacement
    /// Workbook stream or falls back to [`Self::commit`]. Unsupported,
    /// structural, resource, protected, signed, encrypted, macro-bearing,
    /// stale, and length-changing edits are rejected before a candidate is
    /// published.
    ///
    /// A successful result retains the ordinary exact-source [`Patch`] and
    /// semantic inverse while also exposing the reusable bounded overlay plan
    /// and content-free source/target provenance diagnostics. An exact
    /// semantic no-op shares the source snapshot and artifact allocation.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for unsupported staged operations, protected,
    /// signed/encrypted, or macro-bearing sources, stale source ranges,
    /// invalid RK values, CFB limits, failed complete reopen, worksheet
    /// coverage, or numeric readback.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
        )
    )]
    pub(crate) fn commit_source_backed(self) -> Result<SourceBackedCommit> {
        commit_source_backed_numeric(self)
    }

    /// Plans unchanged-family numeric edits without materializing a target
    /// CFB artifact or retaining a target [`Snapshot`].
    ///
    /// This additive publication path accepts only existing `Number`, `RK`,
    /// and `MulRk` value fields. It retains the immutable source snapshot and
    /// a validated [`ValidatedOverlayPlan`] containing the compact exact
    /// physical replacements derived from the numeric source splices. The
    /// composed candidate is reopened
    /// through the public [`crate::raw::Reader`] reader over a positional source before
    /// this method returns. The overlay plan computes the exact source and
    /// target fingerprints once, over the snapshot's sealed immutable bytes;
    /// reading or publishing it does not re-hash an artifact that cannot
    /// change.
    ///
    /// The result deliberately does not expose the ordinary reversible
    /// [`Patch`], because that contract retains complete before/after CFB
    /// artifacts. Callers that require an in-memory target snapshot or an
    /// artifact inverse can continue to use [`Self::commit_source_backed`].
    /// The plan removes the complete target CFB allocation; the bounded
    /// semantic validation itself may still allocate the raw reader's
    /// stream and parsed model while this method is running.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for unsupported staged operations, protected,
    /// signed, encrypted, or macro-bearing sources, stale numeric ranges,
    /// invalid RK values, CFB limits, failed public raw-reader validation, or
    /// source freshness/fingerprint changes.
    pub(crate) fn commit_source_backed_plan(self) -> Result<SourceBackedPlanCommit> {
        commit_source_backed_numeric_plan(self)
    }

    /// Publishes a fully reopened snapshot and reversible exact-source patch.
    ///
    /// # Errors
    ///
    /// Returns an error if stream patching, CFB reconstruction, complete XLS
    /// reopen, or typed semantic readback fails.
    pub(crate) fn commit(self) -> Result<Commit> {
        if let Some(change) = self.tab_change {
            return commit_tab_change(self.source, *change);
        }
        let fixed_cells = self
            .changes
            .iter()
            .filter(|change| change_is_effective(&self.source, change))
            .count();
        let changed_cells = fixed_cells
            .saturating_add(
                self.structural_changes
                    .iter()
                    .filter(|change| matches!(change, StructuralChange::Cell { .. }))
                    .count(),
            )
            .saturating_add(
                self.resource_changes
                    .iter()
                    .filter(|change| matches!(change, ResourceChange::FormulaCell { .. }))
                    .count(),
            );
        let layout: Vec<&layout::LayoutChange> = self
            .layout_changes
            .iter()
            .filter(|change| !change.is_noop())
            .collect();
        let semantic = SemanticPatch::from_transaction(
            &self.source,
            &self.changes,
            &self.structural_changes,
            &self.resource_changes,
            &layout,
        )?;
        if changed_cells == 0
            && self.structural_changes.is_empty()
            && self.resource_changes.is_empty()
            && layout.is_empty()
        {
            let patch = Patch::new(
                Arc::clone(&self.source.inner.bytes),
                Arc::clone(&self.source.inner.bytes),
                semantic,
            );
            return Ok(Commit {
                snapshot: self.source,
                patch,
                diagnostics: Diagnostics::default(),
            });
        }

        let mut workbook_bytes = self.source.inner.workbook_stream.to_vec();
        let shared_strings = self.effective_shared_strings();
        let mut sst_delta = 0_i64;
        for change in &self.changes {
            let entry = &self.source.inner.sheets[change.sheet].entries[change.entry];
            if !change_is_effective(&self.source, change) {
                continue;
            }
            if entry.cell.storage != change.storage {
                let kind_end = entry
                    .kind_offset
                    .checked_add(2)
                    .ok_or_else(|| Error::InvalidData("cell record kind range overflow".into()))?;
                workbook_bytes
                    .get_mut(entry.kind_offset..kind_end)
                    .ok_or_else(|| {
                        Error::InvalidData("cell record kind is outside Workbook".into())
                    })?
                    .copy_from_slice(&storage_record_kind(change.storage).to_le_bytes());
                sst_delta += match (entry.cell.storage, change.storage) {
                    (Storage::LabelSst, Storage::Rk) => -1,
                    (Storage::Rk, Storage::LabelSst) => 1,
                    _ => 0,
                };
            }
            write_cell_value(&mut workbook_bytes, entry, change, &shared_strings)?;
        }
        for change in &self.structural_changes {
            if let StructuralChange::Cell { before, after, .. } = change {
                sst_delta += i64::from(matches!(after, Some((Storage::LabelSst, _, _))))
                    - i64::from(matches!(before, Some((Storage::LabelSst, _, _))));
            }
        }
        if sst_delta != 0
            && let Some(offset) = self.source.inner.sst_total_offset
        {
            // A structural LabelSst replacement can synthesize the first SST
            // for a workbook that previously had no shared-string table. In
            // that case there is no source cstTotal field to patch; the
            // structural compactor writes the complete header after the
            // worksheet record has been rewritten.
            update_sst_total(&mut workbook_bytes, Some(offset), sst_delta)?;
        }

        let fixed_sst_change = self
            .changes
            .iter()
            .filter(|change| change_is_effective(&self.source, change))
            .any(|change| {
                let before = self
                    .source
                    .inner
                    .sheets
                    .get(change.sheet)
                    .and_then(|sheet| sheet.entries.get(change.entry))
                    .map(|entry| entry.cell.storage);
                before == Some(Storage::LabelSst) || change.storage == Storage::LabelSst
            });
        if !self.structural_changes.is_empty()
            || !self.resource_changes.is_empty()
            || !layout.is_empty()
            || fixed_sst_change
        {
            workbook_bytes = structural::apply(
                workbook_bytes,
                &self.source,
                &self.structural_changes,
                &self.changes,
                &self.resource_changes,
                &layout,
                &shared_strings,
            )?;
        }

        // A coalesced edit may briefly intern a text that is restored by a
        // later command. The final Workbook equality proves that no resource
        // survived publication; do not validate the transient resource ledger
        // as if it were part of the exact no-op. Canonicalize this path for
        // both unsigned and signed owners so the patch is genuinely empty.
        if workbook_bytes.as_slice() == self.source.inner.workbook_stream.as_ref() {
            verify_readback(&self.source, &self.changes)?;
            verify_structural_readback(
                &self.source,
                &self.source,
                &self.structural_changes,
                &self.resource_changes,
            )?;
            verify_layout_readback(&self.source, &layout)?;
            let semantic = SemanticPatch::from_transaction(&self.source, &[], &[], &[], &[])?;
            let changed_number_fields = self
                .changes
                .iter()
                .filter(|change| {
                    change_is_effective(&self.source, change)
                        && self.source.inner.sheets[change.sheet].entries[change.entry]
                            .cell
                            .storage
                            == Storage::Number
                })
                .count();
            let patch = Patch::new(
                Arc::clone(&self.source.inner.bytes),
                Arc::clone(&self.source.inner.bytes),
                semantic,
            );
            return Ok(Commit {
                snapshot: self.source,
                patch,
                diagnostics: Diagnostics {
                    changed: false,
                    changed_cells,
                    changed_number_fields,
                    touched_streams: 0,
                },
            });
        }
        if self.source.inner.signed {
            // A signed owner may stage and validate a logical cell/tab plan,
            // but it must never enter the object editor with a changed Workbook
            // stream. Equality was checked above, so an actual change remains
            // refused before publication.
            return Err(Error::Signed);
        }

        let workbook: Arc<[u8]> = Arc::from(workbook_bytes);
        // Candidate admission must retain both profiles accepted by the
        // immutable source owner.  The CFB profile is part of the source
        // contract even when the object profile alone would accept the
        // replacement stream.
        let package = ObjectPackage::open_with_options(
            self.source.inner.bytes.to_vec(),
            OpenOptions::new()
                .targets(Targets::default())
                .limits(self.source.inner.limits)
                .cfb_limits(self.source.inner.cfb_limits),
        )?;
        let mut editor = package.edit()?;
        editor.put_stream_shared(&self.source.inner.workbook_path, workbook)?;
        let package = editor.commit()?.into_package();
        let snapshot = if fixed_cells != 0
            && self.structural_changes.is_empty()
            && self.resource_changes.is_empty()
            && layout.is_empty()
            && changes_are_fixed_numeric(&self.source, &self.changes)
        {
            Snapshot::from_fixed_numeric_package_editor(package, &self.source, &self.changes)?
        } else {
            Snapshot::open_package(package, None)?
        };
        verify_readback(&snapshot, &self.changes)?;
        verify_structural_readback(
            &snapshot,
            &self.source,
            &self.structural_changes,
            &self.resource_changes,
        )?;
        verify_resource_readback(&snapshot, &self.resource_changes)?;
        verify_layout_readback(&snapshot, &layout)?;
        let patch = Patch::new(
            Arc::clone(&self.source.inner.bytes),
            Arc::clone(&snapshot.inner.bytes),
            semantic,
        );
        let changed = !patch.is_noop();
        Ok(Commit {
            snapshot,
            patch,
            diagnostics: Diagnostics {
                changed,
                changed_cells,
                changed_number_fields: self
                    .changes
                    .iter()
                    .filter(|change| {
                        change_is_effective(&self.source, change)
                            && self.source.inner.sheets[change.sheet].entries[change.entry]
                                .cell
                                .storage
                                == Storage::Number
                    })
                    .count(),
                touched_streams: 1,
            },
        })
    }
}

/// Resolves a semantic selector against the complete tab catalog.
fn resolve_tab(catalog: &[tabs::Tab], selector: Selector<'_>) -> Result<usize> {
    match selector {
        Selector::Position(position) if position.get() < catalog.len() => Ok(position.get()),
        Selector::Position(_) => Err(Error::WorksheetNotFound("tab catalog position".into())),
        Selector::Name(name) => {
            let mut matches = catalog
                .iter()
                .enumerate()
                .filter(|(_, tab)| caseless_eq(&tab.name, name));
            let found = matches
                .next()
                .map(|(index, _)| index)
                .ok_or_else(|| Error::WorksheetNotFound("tab catalog name".into()))?;
            if matches.next().is_some() {
                return Err(Error::refused(
                    RefusalKind::InvalidValue,
                    "edit cells",
                    "tab catalog name is ambiguous",
                ));
            }
            Ok(found)
        },
    }
}

/// Publishes one exclusive tab-structure change through the same checked
/// package editor, complete reopen and exact patch as every other commit.
fn commit_tab_change(source: Snapshot, change: tabs::TabChange) -> Result<Commit> {
    let semantic = SemanticPatch::from_tab_change(&change)?;
    let tabs::TabChange {
        target_names,
        workbook,
        ..
    } = change;
    let workbook: Arc<[u8]> = Arc::from(workbook);
    if source.inner.signed {
        if workbook.as_ref() == source.inner.workbook_stream.as_ref() {
            let patch = Patch::new(
                Arc::clone(&source.inner.bytes),
                Arc::clone(&source.inner.bytes),
                semantic,
            );
            return Ok(Commit {
                snapshot: source,
                patch,
                diagnostics: Diagnostics::default(),
            });
        }
        return Err(Error::Signed);
    }
    let package = ObjectPackage::open_with_options(
        source.inner.bytes.to_vec(),
        OpenOptions::new()
            .targets(Targets::default())
            .limits(source.inner.limits)
            .cfb_limits(source.inner.cfb_limits),
    )?;
    let mut editor = package.edit()?;
    editor.put_stream_shared(&source.inner.workbook_path, workbook)?;
    let snapshot = Snapshot::open_package(editor.commit()?.into_package(), None)?;
    let observed = tabs::catalog(&snapshot.inner.workbook_stream)?;
    if observed.len() != target_names.len()
        || observed
            .iter()
            .zip(&target_names)
            .any(|(tab, name)| tab.name != *name)
    {
        return Err(Error::Internal(
            "worksheet tab structure failed catalog readback",
        ));
    }
    let patch = Patch::new(
        Arc::clone(&source.inner.bytes),
        Arc::clone(&snapshot.inner.bytes),
        semantic,
    );
    let changed = !patch.is_noop();
    Ok(Commit {
        snapshot,
        patch,
        diagnostics: Diagnostics {
            changed,
            changed_cells: 0,
            changed_number_fields: 0,
            touched_streams: 1,
        },
    })
}

impl fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("source", &self.source)
            .field("staged_changes", &self.changes.len())
            .field("structural_changes", &self.structural_changes.len())
            .field("resource_changes", &self.resource_changes.len())
            .finish()
    }
}

#[derive(Debug)]
enum TransferCell {
    Insert { reference: Reference, value: Value },
    SetNumber { reference: Reference, value: f64 },
}

fn scalar_default_style(snapshot: &Snapshot) -> Result<StyleIndex> {
    // BIFF8 keeps fifteen style XFs before the first cell XF.  The existing
    // writer emits the canonical default cell XF at index 15; very small
    // foreign producers sometimes expose only a single XF, where index zero
    // is the only safe default.  No other source style can cross a workbook
    // boundary without a resource remapping closure.
    let index = if snapshot.inner.xf_records.len() > 15 {
        15
    } else {
        0
    };
    StyleIndex::new(snapshot, index)
}

fn push_bounded<T>(items: &mut Vec<T>, item: T, limit: usize, context: &'static str) -> Result<()> {
    if items.len() >= limit {
        return Err(Error::refused(
            RefusalKind::InvalidValue,
            "edit cells",
            "cross-workbook scalar transfer inventory exceeds its checked bound",
        ));
    }
    items
        .try_reserve(1)
        .map_err(|_error| Error::Allocation(context))?;
    items.push(item);
    Ok(())
}

fn translated_range(source: CellRange, anchor: Reference) -> Result<CellRange> {
    let end = source.target_reference(source.end(), anchor)?;
    CellRange::new(anchor, end)
}

fn inverse_target_reference(
    source: CellRange,
    anchor: Reference,
    target: Reference,
) -> Result<Reference> {
    if target.row() < anchor.row() || target.column() < anchor.column() {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "target cell precedes the scalar transfer anchor",
        ));
    }
    let source_row = u32::from(source.start().row())
        .checked_add(u32::from(target.row() - anchor.row()))
        .ok_or_else(|| Error::InvalidCellReference("source row overflows BIFF8".into()))?;
    let source_column = u32::from(source.start().column())
        .checked_add(u32::from(target.column() - anchor.column()))
        .ok_or_else(|| Error::InvalidCellReference("source column overflows BIFF8".into()))?;
    let reference = Reference::new(source_row, source_column)?;
    if !source.contains(reference) {
        return Err(Error::refused(
            RefusalKind::InvalidValue,
            "edit cells",
            "target cell lies outside the scalar transfer source range",
        ));
    }
    Ok(reference)
}

fn ensure_transfer_has_no_staged_target_overlap(
    transaction: &Edit,
    sheet: usize,
    target_range: CellRange,
) -> Result<()> {
    if transaction.fixed_index.any_in_range(sheet, target_range)
        || transaction.structural_index.has_cell_sheet(sheet)
        || transaction.structural_index.axis(sheet).is_some()
        || transaction.resource_changes.iter().any(|change| {
            formula_resource_cell(change).is_some_and(|(resource_sheet, reference)| {
                resource_sheet == sheet && target_range.contains(reference)
            })
        })
    {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "cross-workbook scalar transfer overlaps a staged target operation",
        ));
    }
    Ok(())
}

fn ensure_scalar_transfer_container(snapshot: &Snapshot, sheet: usize) -> Result<()> {
    let mut records = Records::new(
        &snapshot.inner.workbook_stream,
        litchi_biff::Limits::default(),
    )?;
    let _globals = records.next().ok_or(Error::Eof("Workbook globals BOF"))??;
    let mut bound_sheet_position = None;
    let mut bound_sheet_index = 0_usize;
    for record_result in records.by_ref() {
        let record = record_result?;
        match record.kind().get() {
            BOUND_SHEET => {
                if bound_sheet_index == snapshot.inner.sheets[sheet].workbook_index {
                    bound_sheet_position = Some(binary::read_u32_le_at(record.payload(), 0)?);
                }
                bound_sheet_index = bound_sheet_index
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidData("BoundSheet index overflow".into()))?;
            },
            EOF => break,
            kind if transfer_unsafe_record(kind)
                || transfer_protected_record(kind, record.payload()) =>
            {
                return Err(Error::refused(
                    RefusalKind::Protected,
                    "edit cells",
                    "cross-workbook scalar transfer refuses protected or dependency BIFF record",
                ));
            },
            kind if !transfer_known_global_record(kind) => {
                return Err(Error::refused(
                    RefusalKind::UnsupportedContent,
                    "edit cells",
                    "cross-workbook scalar transfer refuses unknown global BIFF record",
                ));
            },
            _ => {},
        }
    }
    let position = bound_sheet_position.ok_or_else(|| {
        Error::refused(
            RefusalKind::UnsupportedTarget,
            "edit cells",
            "selected worksheet has no unambiguous BoundSheet owner",
        )
    })?;
    let start = usize::try_from(position)
        .map_err(|_error| Error::InvalidData("BoundSheet position does not fit usize".into()))?;
    let worksheet = snapshot.inner.workbook_stream.get(start..).ok_or_else(|| {
        Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "BoundSheet owner lies outside Workbook",
        )
    })?;
    let mut worksheet_records = Records::new(worksheet, litchi_biff::Limits::default())?;
    let first = worksheet_records
        .next()
        .ok_or(Error::Eof("worksheet BOF"))??;
    require_bof(first.payload(), WORKSHEET)?;
    let mut found_eof = false;
    for record_result in worksheet_records {
        let record = record_result?;
        let kind = record.kind().get();
        if transfer_unsafe_record(kind) || transfer_protected_record(kind, record.payload()) {
            return Err(Error::refused(
                RefusalKind::Protected,
                "edit cells",
                "cross-workbook scalar transfer refuses protected, dependency, drawing, or unsupported BIFF record",
            ));
        }
        if kind == EOF {
            found_eof = true;
            break;
        }
        if !transfer_known_worksheet_record(kind) {
            return Err(Error::refused(
                RefusalKind::NotFound,
                "edit cells",
                "cross-workbook scalar transfer refuses unknown worksheet BIFF record",
            ));
        }
    }
    if !found_eof {
        return Err(Error::Eof("worksheet EOF"));
    }
    Ok(())
}

fn transfer_known_global_record(kind: u16) -> bool {
    matches!(
        kind,
        BOF | EOF
            | CODE_PAGE
            | BOUND_SHEET
            | 0x00C1 // MMS
            | 0x00E1 // INTERFACEHDR
            | 0x00E2 // INTERFACEEND
            | 0x005C // WRITEACCESS
            | 0x005B // FILESHARING (password verifier checked separately)
            | 0x0161 // DSF
            | 0x01C0 // EXCEL9FILE
            | 0x013D // TABID
            | 0x009C // FNGROUPCOUNT
            | 0x009A // FNGROUPNAME
            | 0x0898 // FNGROUPNAME (ContinueFRT12 form)
            | 0x0019 // WINDOWPROTECT (must be disabled)
            | 0x0012 // PROTECT (must be disabled)
            | 0x0013 // PASSWORD (must be zero)
            | 0x01AF // PROT4REV (must be disabled)
            | 0x01BC // PROT4REVPASS (must be zero)
            | 0x003D // WINDOW1
            | 0x0040 // BACKUP
            | 0x008D // HIDEOBJ
            | 0x0022 // DATE1904
            | 0x000E // PRECISION
            | 0x00DA // BOOKBOOL
            | 0x0031 // FONT
            | 0x041E // FORMAT
            | XF
            | 0x0293 // STYLE
            | 0x087C // XFCRC
            | 0x0160 // USESELFS
            | 0x008C // COUNTRY
            | 0x01C1 // RECALCID
    )
}

fn transfer_known_worksheet_record(kind: u16) -> bool {
    matches!(
        kind,
        BOF | EOF
            | 0x0000 // Unused padding occasionally emitted by legacy writers.
            | 0x000C // CalcCount
            | 0x000D // CalcMode
            | 0x000E // Precision
            | 0x000F // RefMode
            | 0x0010 // Iteration
            | 0x0011 // IterationEnabled
            | 0x0014 // Header
            | 0x0015 // Footer
            | 0x001D // Selection
            | 0x0026 // LeftMargin
            | 0x0027 // RightMargin
            | 0x0028 // TopMargin
            | 0x0029 // BottomMargin
            | 0x002A // PrintHeaders
            | 0x002B // PrintGridlines
            | 0x0041 // Pane
            | 0x005E // Uncalced
            | 0x005F // RecalcBeforeSave
            | 0x0055 // DefColWidth
            | 0x007D // ColInfo
            | 0x0080 // Guts
            | 0x0081 // WSBool
            | 0x0082 // Gridset
            | 0x0083 // HCenter
            | 0x0084 // VCenter
            | 0x0099 // RowDefaults
            | 0x00A1 // Uncalced
            | 0x00A0 // PrintTitles
            | 0x00BD // MulRk (also listed explicitly for audit clarity)
            | 0x00D7 // DBCELL
            | 0x00E0 // XF in malformed worksheet-local streams; validation still checks owner.
            | 0x013D // Window2
            | 0x0140 // SCL
            | 0x0161 // DSF
            | 0x0012 // PROTECT (must be disabled)
            | 0x0063 // OBJECTPROTECT (must be disabled)
            | 0x00DD // SCENPROTECT (must be disabled)
            | 0x01AF // PROT4REV (must be disabled)
            | 0x0200 // Dimensions
            | 0x0201 // Blank
            | 0x0203 // Number
            | 0x0205 // BoolErr
            | 0x0208 // Row
            | 0x020B // Index
            | 0x0225 // DefaultRowHeight
            | 0x0236 // Array/legacy formula metadata
            | 0x023E // Window2 extension
            | 0x027E // RK
            | 0x041E // Format
            | 0x087C // XFCRC
            | 0x089C // HeaderFooter
            | 0x00FC // SST is rejected by transfer_unsafe_record before this list.
            | 0x00FD // LabelSst is rejected by transfer_unsafe_record before this list.
            | 0x0006 // Formula is rejected by transfer_unsafe_record before this list.
            | 0x0207 // String is rejected by transfer_unsafe_record before this list.
            | 0x003C // Continue is rejected by transfer_unsafe_record before this list.
    )
}

fn transfer_unsafe_record(kind: u16) -> bool {
    matches!(
        kind,
        FILE_PASS
            | SST
            | LABEL_SST
            | FORMULA
            | STRING
            | CONTINUE
            | 0x0018 // NAME
            | 0x0017 // ExternSheet
            | 0x001C // Note
            | 0x0023 // ExternName
            | 0x01AE // SUPBOOK
            | 0x009D // AutoFilterInfo
            | 0x009E // AutoFilter
            | 0x009F // AutoFilter compatibility owner
            | 0x0059 // XCT
            | 0x005A // CRN
            | 0x005D // Obj
            | 0x00E5 // MergedCells
            | 0x007F // IMDATA
            | 0x0086 // WriteProtect
            | 0x00D1 // DCon
            | 0x00EB // MsoDrawingGroup
            | 0x00EC // MsoDrawing
            | 0x00ED // MsoDrawingSelection
            | 0x00E9 // UserSViewBegin
            | 0x00EA // UserSViewEnd
            | 0x00AF // AutoFilterInfo/legacy owner
            | 0x00B0 // AutoFilter
            | 0x01B6 // TXO
            | 0x01B8 // LinkTable
            | 0x01B9 // DdeLink
            | 0x01B7 // REFRESHALL may trigger external/dependency refresh.
            | 0x00AE // SCENMAN
            | 0x01AA // CUSTOMVIEW begin
            | 0x01AB // CUSTOMVIEW end
            | 0x01B2 // DVAL
            | 0x01BE // ScenarioProtect extension
            | 0x01C2 // CodeName is a macro/dependency owner.
            | 0x01C4 // RealTimeData
            | 0x0813 // RealTimeData (BIFF8 record family)
            | 0x0218 // Hyperlink
            | 0x0221 // Array formula extension
            | 0x0236 // Shared/array formula extension
            | 0x04BC // Table formula extension
            | 0x0091 // Shared formula extension
            | 0x0866 // HeaderFooter picture
    )
}

fn transfer_protected_record(kind: u16, payload: &[u8]) -> bool {
    match kind {
        // WINDOWPROTECT, PROTECT, OBJECTPROTECT, SCENPROTECT, and
        // PROTECTIONREV4 carry a checked BIFF Boolean.  The writer emits
        // these records even when the protection bit is false.
        0x0019 | 0x0012 | 0x0063 | 0x00DD | 0x01AF => payload != [0_u8, 0].as_slice(),
        // PASSWORD and PASSWORDREV4 are inert when their verifier is zero.
        0x0013 | 0x01BC => payload != [0_u8, 0].as_slice(),
        // FILESHARING's write-password verifier is bytes 2..4.  A
        // read-only recommendation alone does not lock the workbook.
        0x005B => payload
            .get(2..4)
            .is_none_or(|password| password != [0_u8, 0].as_slice()),
        _ => false,
    }
}

/// Content-free publication diagnostics.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    changed_cells: usize,
    changed_number_fields: usize,
    touched_streams: usize,
}

impl Diagnostics {
    /// Whether the committed artifact changed exact bytes.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of semantically changed cells.
    #[must_use]
    pub const fn changed_cells(self) -> usize {
        self.changed_cells
    }

    /// Number of eight-byte `Xnum` fields changed.
    #[must_use]
    pub const fn changed_number_fields(self) -> usize {
        self.changed_number_fields
    }

    /// Number of CFB streams in the mutation closure.
    #[must_use]
    pub const fn touched_streams(self) -> usize {
        self.touched_streams
    }
}

/// Successful immutable cell-value publication.
#[derive(Clone)]
pub(crate) struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    diagnostics: Diagnostics,
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

impl Commit {
    /// Records this commit in bounded history after checking its exact base
    /// and matching retained admission profiles.
    ///
    /// # Errors
    ///
    /// Returns a stale-history error or a finite history-weight refusal
    /// without changing the history.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
        )
    )]
    pub(crate) fn record_in(self, history: &mut History) -> Result<Vec<Snapshot>> {
        if history.current().bytes() != self.patch.before() {
            return Err(Error::PatchConflict {
                operation: "edit cells",
                reason: "XLS cell history current snapshot does not match commit base",
            });
        }
        if history.current().inner.limits != self.snapshot.inner.limits
            || history.current().inner.cfb_limits != self.snapshot.inner.cfb_limits
        {
            return Err(Error::PatchConflict {
                operation: "edit cells",
                reason: "XLS cell history source profiles do not match commit profiles",
            });
        }
        let weight = u64::try_from(self.patch.before.len())
            .ok()
            .and_then(|before| {
                u64::try_from(self.patch.after.len())
                    .ok()
                    .and_then(|after| before.checked_add(after))
            })
            .ok_or_else(|| Error::InvalidData("XLS history weight overflow".into()))?;
        history.record(self.snapshot, weight).map_err(|_error| {
            Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "the cell edit history cannot record this commit",
            )
        })
    }

    /// Returns the reopened target snapshot.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
        )
    )]
    #[must_use]
    pub(crate) const fn package(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Returns the exact-source reversible patch.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
        )
    )]
    #[must_use]
    pub(crate) const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Returns content-free publication diagnostics.
    #[must_use]
    pub(crate) const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Splits this commit into its snapshot and exact-source patch.
    ///
    /// Publication diagnostics remain available through [`Self::diagnostics`]
    /// before consuming the commit.
    #[must_use]
    pub(crate) fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// The content-free diagnostics of a source-backed publication, published
/// as a source-backed family's `Diagnostics` (`crate::source_backed`).
pub(crate) mod source_diagnostics {
    #[allow(
        unused_imports,
        reason = "the definition keeps its parent module's vocabulary"
    )]
    use super::*;

    /// Content-free evidence for a source-backed fixed-width numeric publication.
    ///
    /// The value retains bounded counts, lengths, opaque CFB fingerprints, and
    /// source lineage tokens only. Its operation shape is a low-level contract
    /// derived from the validated CFB plan; it reports logical pass scopes rather
    /// than runtime counters or allocator/syscall activity. It never copies
    /// workbook payloads or semantic cell values.
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    pub struct Diagnostics {
        pub(super) changed_cells: usize,
        pub(super) touched_streams: usize,
        pub(super) splice_count: usize,
        pub(super) replacement_bytes: u64,
        pub(super) changed_spans: usize,
        pub(super) source_bytes: u64,
        pub(super) source_workbook_bytes: u64,
        pub(super) target_workbook_bytes: u64,
        pub(super) source_version: SourceVersion,
        pub(super) target_version: SourceVersion,
        pub(super) source_fingerprint: ArtifactFingerprint,
        pub(super) target_fingerprint: ArtifactFingerprint,
        pub(super) operation_shape: OverlayOperationShape,
    }

    impl Diagnostics {
        /// Whether the commit changed the workbook.
        #[must_use]
        pub const fn changed(self) -> bool {
            self.changed_spans > 0
        }

        /// Number of semantically changed numeric cells.
        #[must_use]
        pub const fn changed_cells(self) -> usize {
            self.changed_cells
        }

        /// Number of logical CFB streams selected for publication.
        #[must_use]
        pub const fn touched_streams(self) -> usize {
            self.touched_streams
        }

        /// Number of source-relative numeric ranges submitted to CFB.
        #[must_use]
        pub const fn splice_count(self) -> usize {
            self.splice_count
        }

        /// Aggregate replacement bytes submitted to CFB.
        #[must_use]
        pub const fn replacement_bytes(self) -> u64 {
            self.replacement_bytes
        }

        /// Number of physical CFB spans retained by the validated plan.
        #[must_use]
        pub const fn changed_spans(self) -> usize {
            self.changed_spans
        }

        /// Complete source CFB artifact length.
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

        /// Source lineage/version checked before and during publication.
        #[must_use]
        pub const fn source_version(self) -> SourceVersion {
            self.source_version
        }

        /// Target lineage/version derived by the validated overlay.
        #[must_use]
        pub const fn target_version(self) -> SourceVersion {
            self.target_version
        }

        /// Complete source CFB fingerprint.
        #[must_use]
        pub const fn source_fingerprint(self) -> ArtifactFingerprint {
            self.source_fingerprint
        }

        /// Complete composed target CFB fingerprint.
        #[must_use]
        pub const fn target_fingerprint(self) -> ArtifactFingerprint {
            self.target_fingerprint
        }

        /// Exact low-level CFB pass shape for this source-backed operation.
        #[must_use]
        pub const fn operation_shape(self) -> OverlayOperationShape {
            self.operation_shape
        }
    }
}

use source_diagnostics::Diagnostics as SourceBackedDiagnostics;
/// Successful source-backed fixed-width numeric publication.
///
/// The reusable [`ValidatedOverlayPlan`] owns only the immutable positional
/// source and changed physical spans. The target snapshot is retained to keep
/// the ordinary XLS semantic and exact reversible-patch contract available to
/// callers that need immediate readback or inverse application. Consequently,
/// this API avoids Workbook-stream reserialization but still materializes one
/// complete target CFB allocation; it is not a bounded-artifact-memory API.
#[cfg_attr(
    not(test),
    allow(
        dead_code,
        reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
    )
)]
pub(crate) struct SourceBackedCommit {
    source: Snapshot,
    snapshot: Snapshot,
    plan: ValidatedOverlayPlan,
    patch: Patch,
    diagnostics: SourceBackedDiagnostics,
}

impl SourceBackedCommit {
    /// Exact immutable source snapshot used by this publication.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
        )
    )]
    #[must_use]
    pub(crate) const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Fully reopened target snapshot.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
        )
    )]
    #[must_use]
    pub(crate) const fn package(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Exact-source reversible semantic/artifact patch.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
        )
    )]
    #[must_use]
    pub(crate) const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Content-free source/target publication evidence.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
        )
    )]
    #[must_use]
    pub(crate) const fn diagnostics(&self) -> SourceBackedDiagnostics {
        self.diagnostics
    }

    /// Whether every staged numeric replacement was an exact byte no-op.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
        )
    )]
    #[must_use]
    pub(crate) fn is_noop(&self) -> bool {
        self.plan.is_noop()
    }

    /// Streams the complete validated target to a sequential sink.
    ///
    /// A sink failure may leave a typed prefix in the sink. The returned
    /// [`OverlayError`] retains that partial-output progress.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
        )
    )]
    pub(crate) fn write_to<W: Write>(
        &self,
        writer: &mut W,
    ) -> std::result::Result<PublishReport, OverlayError> {
        self.plan.write_to(writer)
    }

    /// Publishes through the common synced sibling-temp/atomic-rename path.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
        )
    )]
    pub(crate) fn save<P: AsRef<std::path::Path>>(
        &self,
        path: P,
    ) -> std::result::Result<PublishReport, OverlayError> {
        self.plan.save(path)
    }
}

impl fmt::Debug for SourceBackedCommit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceBackedCommit")
            .field("source", &self.source)
            .field("snapshot", &self.snapshot)
            .field("plan", &self.plan)
            .field("patch", &self.patch)
            .field("diagnostics", &self.diagnostics)
            .finish()
    }
}

/// Successful source-bound fixed-width numeric publication plan.
///
/// Unlike [`SourceBackedCommit`], this value never retains a rendered target
/// [`Snapshot`], complete target `Vec<u8>`, or target artifact `Arc<Vec<u8>>`.
/// It keeps the immutable source and the checked CFB overlay plan. The plan
/// owns the compact exact physical replacement spans derived from the numeric
/// splices; logical path and expected-range descriptors are consumed during
/// validation rather than duplicated here. The plan's source and target
/// fingerprints are computed once, at planning, over the source's sealed
/// immutable `Arc<Vec<u8>>`; direct sequential publication, checked composed
/// views and atomic saves rely on that ownership instead of re-hashing the
/// artifact, and atomic saves keep every durability step.
///
/// The plan is intentionally forward-only. An exact artifact inverse would
/// need a source-bound reverse plan rooted in the composed target; the
/// existing [`Patch`] contract instead retains complete before/after
/// artifacts. Use [`Edit::commit_source_backed`] when that inverse
/// contract is required.
pub(crate) struct SourceBackedPlanCommit {
    source: Snapshot,
    plan: ValidatedOverlayPlan,
    diagnostics: SourceBackedDiagnostics,
}

impl SourceBackedPlanCommit {
    /// The checked overlay plan and its diagnostics.
    pub(crate) fn into_plan(self) -> (ValidatedOverlayPlan, SourceBackedDiagnostics) {
        (self.plan, self.diagnostics)
    }

    /// Exact immutable source snapshot used by this plan.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
        )
    )]
    #[must_use]
    pub(crate) const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Content-free publication evidence.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
        )
    )]
    #[must_use]
    pub(crate) const fn diagnostics(&self) -> SourceBackedDiagnostics {
        self.diagnostics
    }

    /// Whether every staged numeric replacement was an exact byte no-op.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
        )
    )]
    #[must_use]
    pub(crate) fn is_noop(&self) -> bool {
        self.plan.is_noop()
    }

    /// Streams the complete validated target to a sequential sink.
    ///
    /// A sink failure may leave a typed prefix in the sink; inspect the
    /// returned [`OverlayError`] before deciding whether to retry.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
        )
    )]
    pub(crate) fn write_to<W: Write>(
        &self,
        writer: &mut W,
    ) -> std::result::Result<PublishReport, OverlayError> {
        self.plan.write_to(writer)
    }
}

impl fmt::Debug for SourceBackedPlanCommit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceBackedPlanCommit")
            .field("source", &self.source)
            .field("plan", &self.plan)
            .field("diagnostics", &self.diagnostics)
            .finish()
    }
}

/// Versioned deterministic semantic patch for fixed-width cell operations.
#[derive(Clone)]
pub struct SemanticPatch {
    inner: CorePatch<Reversible>,
}

impl SemanticPatch {
    fn from_transaction(
        snapshot: &Snapshot,
        changes: &[Change],
        structural_changes: &[StructuralChange],
        resource_changes: &[ResourceChange],
        layout_changes: &[&layout::LayoutChange],
    ) -> Result<Self> {
        let limits = semantic_patch_limits();
        let mut operations = Vec::new();
        operations
            .try_reserve(
                changes
                    .len()
                    .saturating_add(structural_changes.len())
                    .saturating_add(resource_changes.len()),
            )
            .map_err(|_error| Error::Allocation("retaining semantic cell operations"))?;
        for change in changes {
            if !change_is_effective(snapshot, change) {
                continue;
            }
            let sheet = &snapshot.inner.sheets[change.sheet];
            let source = &sheet.entries[change.entry].cell;
            let target = semantic_target(sheet.workbook_index, change.reference);
            let mut forward_preconditions = BTreeMap::new();
            forward_preconditions.insert(
                "sheet_name".to_string(),
                serde_json::Value::String(sheet.name.clone()),
            );
            forward_preconditions.insert(
                "state".to_string(),
                cell_state(source.storage, &source.value),
            );
            append_formula_dependency(
                snapshot,
                &sheet.entries[change.entry],
                &mut forward_preconditions,
            )?;
            let forward = PatchOperation::new(
                limits,
                "cell.set",
                target.clone(),
                forward_preconditions,
                cell_state(change.storage, &change.value),
            )
            .map_err(patch_error)?;
            let mut inverse_preconditions = BTreeMap::new();
            inverse_preconditions.insert(
                "sheet_name".to_string(),
                serde_json::Value::String(sheet.name.clone()),
            );
            inverse_preconditions.insert(
                "state".to_string(),
                cell_state(change.storage, &change.value),
            );
            append_formula_dependency(
                snapshot,
                &sheet.entries[change.entry],
                &mut inverse_preconditions,
            )?;
            let inverse = PatchOperation::new(
                limits,
                "cell.set",
                target,
                inverse_preconditions,
                cell_state(source.storage, &source.value),
            )
            .map_err(patch_error)?;
            operations.push(ReversibleOperation::new(forward, inverse));
        }
        for change in structural_changes {
            append_structural_semantic(
                snapshot,
                change,
                resource_changes,
                limits,
                &mut operations,
            )?;
        }
        for change in resource_changes {
            append_resource_semantic(snapshot, resource_changes, change, limits, &mut operations)?;
        }
        for change in layout_changes {
            append_layout_semantic(snapshot, change, limits, &mut operations)?;
        }
        operations.sort_by(|left, right| {
            semantic_operation_order(left.forward())
                .cmp(&semantic_operation_order(right.forward()))
                .then_with(|| left.forward().target.cmp(&right.forward().target))
                .then_with(|| left.forward().op.cmp(&right.forward().op))
        });
        let blobs = BlobBundle::new(limits.blobs());
        let reverse_blobs = BlobBundle::new(limits.blobs());
        let inner = CorePatch::<Reversible>::new(
            limits,
            "litchi-xls.cell-values",
            operations,
            blobs,
            reverse_blobs,
        )
        .map_err(patch_error)?;
        Ok(Self { inner })
    }

    fn from_tab_change(change: &tabs::TabChange) -> Result<Self> {
        let limits = semantic_patch_limits();
        let names = |names: &[String]| {
            serde_json::Value::Array(
                names
                    .iter()
                    .cloned()
                    .map(serde_json::Value::String)
                    .collect(),
            )
        };
        let preconditions = |names_value: serde_json::Value| {
            let mut map = BTreeMap::new();
            map.insert("tab_names".to_string(), names_value);
            map
        };
        let (forward, inverse) = match &change.kind {
            tabs::TabChangeKind::Reorder { order } => {
                let mut inverse_order = vec![0_usize; order.len()];
                for (new, old) in order.iter().enumerate() {
                    inverse_order[*old] = new;
                }
                (
                    PatchOperation::new(
                        limits,
                        "sheet.reorder",
                        tabs::SEMANTIC_TARGET.to_string(),
                        preconditions(names(&change.source_names)),
                        serde_json::json!(order),
                    ),
                    PatchOperation::new(
                        limits,
                        "sheet.reorder",
                        tabs::SEMANTIC_TARGET.to_string(),
                        preconditions(names(&change.target_names)),
                        serde_json::json!(inverse_order),
                    ),
                )
            },
            tabs::TabChangeKind::Insert { position, name } => {
                let target = tabs::semantic_tab_target(*position);
                let mut inverse_preconditions = preconditions(names(&change.target_names));
                inverse_preconditions.insert("empty".to_string(), serde_json::Value::Bool(true));
                (
                    PatchOperation::new(
                        limits,
                        "sheet.insert",
                        target.clone(),
                        preconditions(names(&change.source_names)),
                        serde_json::json!({ "name": name }),
                    ),
                    PatchOperation::new(
                        limits,
                        "sheet.remove",
                        target,
                        inverse_preconditions,
                        serde_json::json!({ "name": name }),
                    ),
                )
            },
            tabs::TabChangeKind::Remove {
                position,
                name,
                fingerprint,
            } => {
                let target = tabs::semantic_tab_target(*position);
                let mut forward_preconditions = preconditions(names(&change.source_names));
                forward_preconditions.insert(
                    "fingerprint".to_string(),
                    serde_json::Value::String(fingerprint.clone()),
                );
                (
                    PatchOperation::new(
                        limits,
                        "sheet.remove",
                        target.clone(),
                        forward_preconditions,
                        serde_json::json!({ "name": name }),
                    ),
                    PatchOperation::new(
                        limits,
                        "sheet.restore",
                        target,
                        preconditions(names(&change.target_names)),
                        serde_json::json!({ "name": name, "fingerprint": fingerprint }),
                    ),
                )
            },
        };
        let operations = vec![ReversibleOperation::new(
            forward.map_err(patch_error)?,
            inverse.map_err(patch_error)?,
        )];
        let inner = CorePatch::<Reversible>::new(
            limits,
            "litchi-xls.cell-values",
            operations,
            BlobBundle::new(limits.blobs()),
            BlobBundle::new(limits.blobs()),
        )
        .map_err(patch_error)?;
        Ok(Self { inner })
    }

    /// Parses canonical deterministic JSON under the XLS cell-patch bounds.
    ///
    /// # Errors
    ///
    /// Returns a bounded wire, canonicality, schema, or operation error.
    pub fn from_deterministic_json(bytes: &[u8]) -> Result<Self> {
        let inner =
            CorePatch::<Reversible>::from_deterministic_json(bytes, semantic_patch_limits())
                .map_err(patch_error)?;
        if inner.format() != "litchi-xls.cell-values" {
            return Err(Error::InvalidData(
                "durable patch has the wrong format namespace".into(),
            ));
        }
        Ok(Self { inner })
    }

    /// Serializes canonical deterministic JSON.
    ///
    /// # Errors
    ///
    /// Returns an error if the retained operations exceed their wire bounds.
    pub fn to_deterministic_json(&self) -> Result<Vec<u8>> {
        self.inner.to_deterministic_json().map_err(patch_error)
    }

    /// Number of semantic cell operations.
    #[must_use]
    #[expect(
        clippy::len_without_is_empty,
        reason = "patch no-op introspection uses the canonical is_noop vocabulary"
    )]
    pub fn len(&self) -> usize {
        self.inner.operations().len()
    }

    /// Whether this patch has no semantic operation.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.inner.operations().is_empty()
    }

    /// Returns the saturating byte charge for retained semantic operations.
    ///
    /// Both forward and inverse operation data are included by the shared
    /// patch owner. The calculation walks retained strings, JSON values, and
    /// blobs directly; it does not serialize the patch or construct an
    /// inverse clone.
    #[must_use]
    pub fn retained_bytes(&self) -> usize {
        self.inner.retained_bytes()
    }

    /// Exact semantic inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            inner: self.inner.inverse(),
        }
    }

    /// Applies semantic operations after checking every typed precondition.
    ///
    /// Unlike the artifact patch, this can apply to a byte-different snapshot
    /// with the same sheet identities and exact source cell states.
    ///
    /// # Errors
    ///
    /// Returns a malformed-operation, stale-precondition, storage, or commit
    /// error without changing the source snapshot.
    pub(crate) fn stage(&self, source: &Snapshot) -> Result<Edit> {
        let mut transaction = source.edit();
        let mut sheet_names = semantic_sheet_names(source);
        for operation in self.inner.operations() {
            apply_semantic_operation(source, &mut transaction, operation, &mut sheet_names)?;
        }
        Ok(transaction)
    }

    /// Plans a deterministic three-way semantic merge against a common base.
    ///
    /// Both patches are preflighted against `base`. Identical outcomes
    /// coalesce, disjoint targets compose, and overlapping divergent outcomes
    /// are retained as typed conflicts without publishing bytes.
    ///
    /// # Errors
    ///
    /// Returns a stale-precondition, malformed-operation, allocation, or
    /// finite-patch-bound error before a plan is returned.
    pub(crate) fn plan_three_way(
        base: &Snapshot,
        left: &Self,
        right: &Self,
    ) -> Result<ThreeWayPlan> {
        preflight_semantic(base, left)?;
        preflight_semantic(base, right)?;
        build_three_way_plan(left, right)
    }

    /// Preflights dependency-aware application to another opened workbook.
    ///
    /// The plan records per-operation resource, identity, storage, and
    /// structural-closure refusals. No candidate bytes are built until an
    /// executable plan is applied.
    #[must_use]
    pub(crate) fn plan_transfer(&self, target: &Snapshot) -> TransferPlan {
        let mut refusals = Vec::new();
        let mut transaction = target.edit();
        let mut sheet_names = semantic_sheet_names(target);
        for operation in self.inner.operations() {
            if let Err(error) =
                apply_semantic_operation(target, &mut transaction, operation, &mut sheet_names)
            {
                refusals.push(TransferRefusal {
                    target: operation.target.clone(),
                    reason: error.to_string(),
                });
                break;
            }
        }
        TransferPlan {
            patch: self.clone(),
            refusals: refusals.into_boxed_slice(),
        }
    }
}

/// One deterministic three-way conflict between semantic operations.
#[derive(Debug, Clone, PartialEq)]
pub struct ThreeWayConflict {
    target: String,
    left: serde_json::Value,
    right: serde_json::Value,
}

impl ThreeWayConflict {
    /// Canonical semantic target.
    #[must_use]
    pub fn target(&self) -> &str {
        &self.target
    }

    /// Left requested outcome.
    #[must_use]
    pub const fn left(&self) -> &serde_json::Value {
        &self.left
    }

    /// Right requested outcome.
    #[must_use]
    pub const fn right(&self) -> &serde_json::Value {
        &self.right
    }
}

/// Non-mutating deterministic result of semantic three-way planning.
#[derive(Clone)]
pub struct ThreeWayPlan {
    merged: Option<SemanticPatch>,
    conflicts: Box<[ThreeWayConflict]>,
}

impl ThreeWayPlan {
    /// Conflicts in canonical target order.
    #[must_use]
    pub fn conflicts(&self) -> &[ThreeWayConflict] {
        &self.conflicts
    }

    /// Conflict-free merged patch, when every overlap agreed.
    #[must_use]
    pub const fn merged(&self) -> Option<&SemanticPatch> {
        self.merged.as_ref()
    }
}

impl fmt::Debug for ThreeWayPlan {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ThreeWayPlan")
            .field(
                "merged_operations",
                &self.merged.as_ref().map(SemanticPatch::len),
            )
            .field("conflicts", &self.conflicts)
            .finish()
    }
}

/// One dependency or precondition that prevents semantic transfer.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TransferRefusal {
    target: String,
    reason: String,
}

impl TransferRefusal {
    /// Canonical semantic target.
    #[must_use]
    pub fn target(&self) -> &str {
        &self.target
    }

    /// Stable typed-error description produced during preflight.
    #[must_use]
    pub fn reason(&self) -> &str {
        &self.reason
    }
}

/// Dependency-aware transfer plan for a durable semantic patch.
#[derive(Debug, Clone)]
pub struct TransferPlan {
    patch: SemanticPatch,
    refusals: Box<[TransferRefusal]>,
}

impl TransferPlan {
    /// Whether every operation and resource dependency preflighted.
    #[must_use]
    pub fn is_executable(&self) -> bool {
        self.refusals.is_empty()
    }

    /// Refusals in durable operation order.
    #[must_use]
    pub fn refusals(&self) -> &[TransferRefusal] {
        &self.refusals
    }

    /// Applies the planned semantic transfer and fully reopens the result.
    ///
    /// # Errors
    ///
    /// Returns the first retained refusal or a typed commit/reopen error.
    pub(crate) fn stage(&self, target: &Snapshot) -> Result<Edit> {
        if !self.refusals.is_empty() {
            return Err(Error::refused(
                RefusalKind::UnsupportedTarget,
                "edit cells",
                "a target of the semantic transfer is not executable",
            ));
        }
        self.patch.stage(target)
    }
}

fn append_structural_semantic(
    snapshot: &Snapshot,
    change: &StructuralChange,
    resources: &[ResourceChange],
    limits: PatchLimits,
    operations: &mut Vec<ReversibleOperation>,
) -> Result<()> {
    let sheet_data = &snapshot.inner.sheets[operation_sheet_index(change)];
    let target = structural_target(snapshot, change);
    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        "sheet_name".to_string(),
        serde_json::Value::String(sheet_data.name.clone()),
    );
    let (op, value, inverse_op, inverse_value, inverse_name) = match change {
        StructuralChange::Cell {
            reference,
            before,
            after,
            ..
        } => {
            preconditions.insert("state".to_string(), optional_cell_state(before.as_ref()));
            if matches!(before, Some((Storage::Formula, _, _))) {
                let entry = sheet_data
                    .unique_entry_index(*reference)?
                    .map(|index| &sheet_data.entries[index])
                    .ok_or_else(|| {
                        Error::refused(
                            RefusalKind::NotFound,
                            "edit cells",
                            "Formula dependency owner is absent",
                        )
                    })?;
                append_formula_dependency(snapshot, entry, &mut preconditions)?;
            }
            if let Some((_, _, style)) = after {
                preconditions.insert(
                    "target_xf".to_string(),
                    serde_json::Value::String(effective_xf_fingerprint(
                        snapshot, resources, *style,
                    )?),
                );
            }
            (
                "cell.structural",
                optional_cell_state(after.as_ref()),
                "cell.structural",
                optional_cell_state(before.as_ref()),
                sheet_data.name.clone(),
            )
        },
        StructuralChange::Rows {
            sheet,
            start,
            count,
            insert,
        } => {
            let end = start.saturating_add(*count);
            if !insert && snapshot.inner.sheets[*sheet].any_rows(*start, end.saturating_sub(1)) {
                return Err(Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit cells",
                    "durable row deletion is refused when the deleted span contains cells",
                ));
            }
            (
                if *insert {
                    "rows.insert"
                } else {
                    "rows.delete"
                },
                serde_json::json!({"count": count}),
                if *insert {
                    "rows.delete"
                } else {
                    "rows.insert"
                },
                serde_json::json!({"count": count}),
                sheet_data.name.clone(),
            )
        },
        StructuralChange::Columns {
            sheet,
            start,
            count,
            insert,
        } => {
            let end = start.saturating_add(*count);
            if !insert && snapshot.inner.sheets[*sheet].any_columns(*start, end.saturating_sub(1)) {
                return Err(Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit cells",
                    "durable column deletion is refused when the deleted span contains cells",
                ));
            }
            (
                if *insert {
                    "columns.insert"
                } else {
                    "columns.delete"
                },
                serde_json::json!({"count": count}),
                if *insert {
                    "columns.delete"
                } else {
                    "columns.insert"
                },
                serde_json::json!({"count": count}),
                sheet_data.name.clone(),
            )
        },
        StructuralChange::RenameSheet { before, after, .. } => {
            preconditions.clear();
            preconditions.insert(
                "sheet_name".to_string(),
                serde_json::Value::String(before.clone()),
            );
            (
                "sheet.rename",
                serde_json::Value::String(after.clone()),
                "sheet.rename",
                serde_json::Value::String(before.clone()),
                after.clone(),
            )
        },
    };
    let forward = PatchOperation::new(limits, op, target.clone(), preconditions, value)
        .map_err(patch_error)?;
    let mut inverse_preconditions = BTreeMap::new();
    inverse_preconditions.insert(
        "sheet_name".to_string(),
        serde_json::Value::String(inverse_name),
    );
    if let StructuralChange::Cell { after, .. } = change {
        inverse_preconditions.insert("state".to_string(), optional_cell_state(after.as_ref()));
        if let StructuralChange::Cell {
            reference,
            after: Some((Storage::Formula, _, _)),
            ..
        } = change
        {
            let entry = sheet_data
                .unique_entry_index(*reference)?
                .map(|index| &sheet_data.entries[index])
                .ok_or_else(|| {
                    Error::refused(
                        RefusalKind::NotFound,
                        "edit cells",
                        "Formula dependency owner is absent",
                    )
                })?;
            append_formula_dependency(snapshot, entry, &mut inverse_preconditions)?;
        }
        if let StructuralChange::Cell {
            before: Some((_, _, style)),
            ..
        } = change
        {
            inverse_preconditions.insert(
                "target_xf".to_string(),
                serde_json::Value::String(effective_xf_fingerprint(snapshot, resources, *style)?),
            );
        }
    }
    let inverse = PatchOperation::new(
        limits,
        inverse_op,
        target,
        inverse_preconditions,
        inverse_value,
    )
    .map_err(patch_error)?;
    operations.push(ReversibleOperation::new(forward, inverse));
    Ok(())
}

fn append_resource_semantic(
    snapshot: &Snapshot,
    resources: &[ResourceChange],
    change: &ResourceChange,
    limits: PatchLimits,
    operations: &mut Vec<ReversibleOperation>,
) -> Result<()> {
    let (target, forward_op, inverse_op, value, present, inverse_present) = match change {
        ResourceChange::SharedString { text, insert } => (
            format!("resource/sst/{}", text_fingerprint(text)),
            if *insert { "sst.intern" } else { "sst.remove" },
            if *insert { "sst.remove" } else { "sst.intern" },
            serde_json::Value::String(text.clone()),
            !*insert,
            *insert,
        ),
        ResourceChange::RichSharedString {
            text,
            formatting_runs,
            insert,
        } => (
            format!(
                "resource/sst-rich/{}",
                rich_text_fingerprint(text, formatting_runs)
            ),
            if *insert {
                "sst.intern-rich"
            } else {
                "sst.remove-rich"
            },
            if *insert {
                "sst.remove-rich"
            } else {
                "sst.intern-rich"
            },
            serde_json::json!({
                "text": text,
                "runs": formatting_runs.iter().map(|run| {
                    serde_json::json!([run.character_index, run.font_index])
                }).collect::<Vec<_>>()
            }),
            !*insert,
            *insert,
        ),
        ResourceChange::ExtendedFormat {
            index,
            payload,
            insert,
        } => (
            format!("resource/xf/{:05}", index.get()),
            if *insert { "xf.author" } else { "xf.remove" },
            if *insert { "xf.remove" } else { "xf.author" },
            serde_json::Value::String(bytes_hex(payload)),
            !*insert,
            *insert,
        ),
        ResourceChange::FormulaCell {
            sheet,
            reference,
            style,
            tokens,
            insert,
        } => {
            let sheet_data = snapshot
                .inner
                .sheets
                .get(*sheet)
                .ok_or(Error::PatchConflict {
                    operation: "edit cells",
                    reason: "Formula semantic sheet dependency is stale",
                })?;
            (
                format!(
                    "sheet/{}/formula/{}/{}",
                    sheet_data.workbook_index,
                    reference.row(),
                    reference.column()
                ),
                if *insert {
                    "formula.insert"
                } else {
                    "formula.remove"
                },
                if *insert {
                    "formula.remove"
                } else {
                    "formula.insert"
                },
                serde_json::json!({
                    "tokens": bytes_hex(tokens),
                    "style": style.get(),
                    "target_xf": effective_xf_fingerprint(snapshot, resources, *style)?,
                }),
                !*insert,
                *insert,
            )
        },
    };
    let mut preconditions = BTreeMap::new();
    preconditions.insert("present".to_string(), serde_json::Value::Bool(present));
    if let ResourceChange::FormulaCell { sheet, .. } = change {
        let sheet_name = snapshot
            .inner
            .sheets
            .get(*sheet)
            .ok_or(Error::PatchConflict {
                operation: "edit cells",
                reason: "Formula semantic sheet dependency is stale",
            })?
            .name
            .clone();
        preconditions.insert(
            "sheet_name".to_string(),
            serde_json::Value::String(sheet_name),
        );
    }
    let forward = PatchOperation::new(
        limits,
        forward_op,
        target.clone(),
        preconditions,
        value.clone(),
    )
    .map_err(patch_error)?;
    let mut inverse_preconditions = BTreeMap::new();
    inverse_preconditions.insert(
        "present".to_string(),
        serde_json::Value::Bool(inverse_present),
    );
    if let ResourceChange::FormulaCell { sheet, .. } = change {
        let sheet_name = snapshot
            .inner
            .sheets
            .get(*sheet)
            .ok_or(Error::PatchConflict {
                operation: "edit cells",
                reason: "Formula semantic sheet dependency is stale",
            })?
            .name
            .clone();
        inverse_preconditions.insert(
            "sheet_name".to_string(),
            serde_json::Value::String(sheet_name),
        );
    }
    let inverse = PatchOperation::new(limits, inverse_op, target, inverse_preconditions, value)
        .map_err(patch_error)?;
    operations.push(ReversibleOperation::new(forward, inverse));
    Ok(())
}

fn text_fingerprint(text: &str) -> String {
    DiagnosticFingerprint::of(text.as_bytes()).as_hex()
}

fn rich_text_fingerprint(
    text: &str,
    formatting_runs: &[crate::records::SharedStringFormatRun],
) -> String {
    let capacity = text
        .len()
        .saturating_add(formatting_runs.len().saturating_mul(4).saturating_add(1));
    let mut bytes = Vec::with_capacity(capacity);
    bytes.extend_from_slice(text.as_bytes());
    bytes.push(0);
    for run in formatting_runs {
        bytes.extend_from_slice(&run.character_index.to_le_bytes());
        bytes.extend_from_slice(&run.font_index.to_le_bytes());
    }
    DiagnosticFingerprint::of(&bytes).as_hex()
}

fn bytes_hex(bytes: &[u8]) -> String {
    const HEX: &[u8; 16] = b"0123456789abcdef";
    let mut encoded = String::with_capacity(bytes.len().saturating_mul(2));
    for byte in bytes {
        encoded.push(char::from(HEX[usize::from(*byte >> 4)]));
        encoded.push(char::from(HEX[usize::from(*byte & 0x0f)]));
    }
    encoded
}

fn append_formula_dependency(
    snapshot: &Snapshot,
    entry: &Entry,
    preconditions: &mut BTreeMap<String, serde_json::Value>,
) -> Result<()> {
    if entry.cell.storage != Storage::Formula {
        return Ok(());
    }
    preconditions.insert(
        "formula_dependency".to_string(),
        serde_json::Value::String(formula_dependency_fingerprint(snapshot, entry)?),
    );
    Ok(())
}

fn formula_dependency_fingerprint(snapshot: &Snapshot, entry: &Entry) -> Result<String> {
    let start = entry.kind_offset;
    if binary::read_u16_le_at(&snapshot.inner.workbook_stream, start)? != FORMULA {
        return Err(Error::InvalidData(
            "Formula dependency owner changed record kind".into(),
        ));
    }
    let length_offset = start
        .checked_add(2)
        .ok_or_else(|| Error::InvalidData("Formula length offset overflow".into()))?;
    let payload_len = usize::from(binary::read_u16_le_at(
        &snapshot.inner.workbook_stream,
        length_offset,
    )?);
    if payload_len < 22 {
        return Err(Error::InvalidData(
            "Formula dependency record is truncated".into(),
        ));
    }
    let end = start
        .checked_add(4)
        .and_then(|value| value.checked_add(payload_len))
        .ok_or_else(|| Error::InvalidData("Formula dependency range overflow".into()))?;
    let dependency_start = start
        .checked_add(18)
        .ok_or_else(|| Error::InvalidData("Formula dependency offset overflow".into()))?;
    let dependency = snapshot
        .inner
        .workbook_stream
        .get(dependency_start..end)
        .ok_or_else(|| Error::InvalidData("Formula dependency is outside Workbook".into()))?;
    Ok(DiagnosticFingerprint::of(dependency).as_hex())
}

fn authored_formula_record_matches(
    snapshot: &Snapshot,
    entry: &Entry,
    reference: Reference,
    style: StyleIndex,
    tokens: &[u8],
) -> Result<bool> {
    if entry.cell.storage != Storage::Formula {
        return Ok(false);
    }
    let mut expected = Vec::new();
    crate::writer::biff::write_formula(
        &mut expected,
        u32::from(reference.row()),
        u16::from(reference.column()),
        style.get(),
        tokens,
    )?;
    let start = entry.kind_offset;
    let payload_length_offset = start
        .checked_add(2)
        .ok_or_else(|| Error::InvalidData("authored Formula length offset overflow".into()))?;
    let payload_length = usize::from(binary::read_u16_le_at(
        &snapshot.inner.workbook_stream,
        payload_length_offset,
    )?);
    let payload_start = start
        .checked_add(4)
        .ok_or_else(|| Error::InvalidData("authored Formula payload offset overflow".into()))?;
    let end = payload_start
        .checked_add(payload_length)
        .ok_or_else(|| Error::InvalidData("authored Formula record range overflow".into()))?;
    let payload = snapshot
        .inner
        .workbook_stream
        .get(payload_start..end)
        .ok_or_else(|| {
            Error::InvalidData("authored Formula record lies outside Workbook".into())
        })?;
    let (source_tokens, source_extra) = crate::formula_metadata::formula_payload_parts(payload)?;
    if !source_extra.is_empty() {
        crate::formula_metadata::validate_formula_extra(source_tokens, source_extra)?;
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "Formula resource matching refuses a source Formula with RgbExtra ancillary bytes",
        ));
    }
    Ok(snapshot
        .inner
        .workbook_stream
        .get(start..end)
        .is_some_and(|record| record.len() == expected.len() && record == expected.as_slice()))
}

fn verify_formula_dependency(
    snapshot: &Snapshot,
    entry: &Entry,
    operation: &PatchOperation,
) -> Result<()> {
    if entry.cell.storage != Storage::Formula {
        return Ok(());
    }
    let expected = operation
        .preconditions
        .get("formula_dependency")
        .and_then(serde_json::Value::as_str)
        .ok_or_else(|| Error::InvalidData("Formula patch has no token dependency".into()))?;
    if formula_dependency_fingerprint(snapshot, entry)? != expected {
        return Err(Error::PatchConflict {
            operation: "edit cells",
            reason: "Formula token dependency is stale",
        });
    }
    Ok(())
}

fn effective_xf_fingerprint(
    snapshot: &Snapshot,
    resources: &[ResourceChange],
    style: StyleIndex,
) -> Result<String> {
    if let Some(payload) = snapshot.inner.xf_records.get(usize::from(style.get())) {
        return Ok(bytes_hex(payload));
    }
    resources
        .iter()
        .find_map(|resource| match resource {
            ResourceChange::ExtendedFormat {
                index,
                payload,
                insert: true,
            } if *index == style => Some(bytes_hex(payload)),
            ResourceChange::SharedString { .. }
            | ResourceChange::RichSharedString { .. }
            | ResourceChange::ExtendedFormat { .. }
            | ResourceChange::FormulaCell { .. } => None,
        })
        .ok_or_else(|| {
            Error::refused(
                RefusalKind::NotFound,
                "edit cells",
                "structural cell XF dependency is absent",
            )
        })
}

fn preflight_semantic(source: &Snapshot, patch: &SemanticPatch) -> Result<()> {
    let mut transaction = source.edit();
    let mut sheet_names = semantic_sheet_names(source);
    for operation in patch.inner.operations() {
        apply_semantic_operation(source, &mut transaction, operation, &mut sheet_names)?;
    }
    Ok(())
}

fn build_three_way_plan(left: &SemanticPatch, right: &SemanticPatch) -> Result<ThreeWayPlan> {
    let mut merged = reversible_pairs(left);
    let mut conflicts = Vec::new();
    for pair in reversible_pairs(right) {
        let first_overlap = merged
            .iter()
            .position(|existing| semantic_operations_overlap(&existing.0, &pair.0));
        let Some(first_overlap) = first_overlap else {
            merged.push(pair);
            continue;
        };
        if merged.iter().any(|existing| {
            semantic_operations_overlap(&existing.0, &pair.0) && existing.0 == pair.0
        }) {
            continue;
        }
        conflicts.push(ThreeWayConflict {
            target: pair.0.target.clone(),
            left: merged[first_overlap].0.value.clone(),
            right: pair.0.value.clone(),
        });
    }
    conflicts.sort_by(|left, right| left.target.cmp(&right.target));
    conflicts.dedup_by(|left, right| left.target == right.target);
    if !conflicts.is_empty() {
        return Ok(ThreeWayPlan {
            merged: None,
            conflicts: conflicts.into_boxed_slice(),
        });
    }
    merged.sort_by(|left, right| {
        semantic_operation_order(&left.0)
            .cmp(&semantic_operation_order(&right.0))
            .then_with(|| left.0.target.cmp(&right.0.target))
            .then_with(|| left.0.op.cmp(&right.0.op))
    });
    let inner = CorePatch::<Reversible>::new(
        semantic_patch_limits(),
        "litchi-xls.cell-values",
        merged
            .into_iter()
            .map(|(forward, inverse)| ReversibleOperation::new(forward, inverse)),
        BlobBundle::new(semantic_patch_limits().blobs()),
        BlobBundle::new(semantic_patch_limits().blobs()),
    )
    .map_err(patch_error)?;
    Ok(ThreeWayPlan {
        merged: Some(SemanticPatch { inner }),
        conflicts: Box::default(),
    })
}

fn reversible_pairs(patch: &SemanticPatch) -> Vec<(PatchOperation, PatchOperation)> {
    let inverse = patch.inner.inverse();
    patch
        .inner
        .operations()
        .iter()
        .cloned()
        .zip(inverse.operations().iter().rev().cloned())
        .collect()
}

fn is_tab_operation(op: &str) -> bool {
    matches!(
        op,
        "sheet.reorder" | "sheet.insert" | "sheet.remove" | "sheet.restore"
    )
}

fn semantic_operations_overlap(left: &PatchOperation, right: &PatchOperation) -> bool {
    if is_tab_operation(&left.op) || is_tab_operation(&right.op) {
        // A tab-structure operation renumbers every catalog address.
        return true;
    }
    let left_sheet = semantic_operation_sheet(left);
    let right_sheet = semantic_operation_sheet(right);
    if left_sheet != right_sheet {
        return false;
    }
    let left_axis = matches!(
        left.op.as_str(),
        "rows.insert" | "rows.delete" | "columns.insert" | "columns.delete"
    );
    let right_axis = matches!(
        right.op.as_str(),
        "rows.insert" | "rows.delete" | "columns.insert" | "columns.delete"
    );
    if left_axis || right_axis {
        return left.op != "sheet.rename" && right.op != "sheet.rename";
    }
    if semantic_cell_identity(left).is_some()
        && semantic_cell_identity(left) == semantic_cell_identity(right)
    {
        return true;
    }
    left.target == right.target
}

fn semantic_cell_identity(operation: &PatchOperation) -> Option<(usize, u16, u8)> {
    let mut parts = operation.target.split('/');
    if parts.next()? != "sheet" {
        return None;
    }
    let sheet = parts.next()?.parse().ok()?;
    if !matches!(parts.next()?, "cell" | "formula") {
        return None;
    }
    let row = parts.next()?.parse().ok()?;
    let column = parts.next()?.parse().ok()?;
    parts.next().is_none().then_some((sheet, row, column))
}

fn semantic_operation_order(operation: &PatchOperation) -> u8 {
    if is_tab_operation(&operation.op) {
        4
    } else if operation.op == "sheet.rename" {
        3
    } else if matches!(
        operation.op.as_str(),
        "sst.intern"
            | "sst.remove"
            | "sst.intern-rich"
            | "sst.remove-rich"
            | "xf.author"
            | "xf.duplicate"
            | "xf.remove"
            // A removed authored Formula vacates its cell before cell
            // operations; an inserted one follows them (same order as the
            // `sheet/N/cell` targets, which sort first).
            | "formula.remove"
    ) {
        0
    } else {
        1 + u8::from(matches!(
            operation.op.as_str(),
            "rows.insert" | "rows.delete" | "columns.insert" | "columns.delete"
        ))
    }
}

fn semantic_operation_sheet(operation: &PatchOperation) -> Option<usize> {
    let mut parts = operation.target.split('/');
    (parts.next() == Some("sheet"))
        .then(|| parts.next()?.parse::<usize>().ok())
        .flatten()
}

fn apply_semantic_operation(
    source: &Snapshot,
    transaction: &mut Edit,
    operation: &PatchOperation,
    sheet_names: &mut [String],
) -> Result<()> {
    match operation.op.as_str() {
        "sst.intern" | "sst.remove" => apply_sst_semantic(transaction, operation),
        "sst.intern-rich" | "sst.remove-rich" => apply_rich_sst_semantic(transaction, operation),
        "xf.author" | "xf.duplicate" | "xf.remove" => apply_xf_semantic(transaction, operation),
        "formula.insert" | "formula.remove" => {
            apply_formula_semantic(source, transaction, operation, sheet_names)
        },
        "cell.set" => apply_fixed_semantic(source, transaction, operation, sheet_names),
        "cell.structural" => {
            apply_structural_cell_semantic(source, transaction, operation, sheet_names)
        },
        "rows.insert" | "rows.delete" => {
            let (sheet, start) = parse_axis_target(&operation.target, "rows")?;
            verify_semantic_sheet(source, sheet_names, sheet, operation)?;
            let count = parse_count(&operation.value, "row")?;
            if operation.op == "rows.insert" {
                transaction.insert_rows(Selector::Position(Position::new(sheet)), start, count)
            } else {
                transaction.delete_rows(Selector::Position(Position::new(sheet)), start, count)
            }
        },
        "columns.insert" | "columns.delete" => {
            let (sheet, start) = parse_axis_target(&operation.target, "columns")?;
            verify_semantic_sheet(source, sheet_names, sheet, operation)?;
            let start = u8::try_from(start)
                .map_err(|_error| Error::InvalidData("column patch start exceeds u8".into()))?;
            let count = u8::try_from(parse_count(&operation.value, "column")?)
                .map_err(|_error| Error::InvalidData("column patch count exceeds u8".into()))?;
            if operation.op == "columns.insert" {
                transaction.insert_columns(Selector::Position(Position::new(sheet)), start, count)
            } else {
                transaction.delete_columns(Selector::Position(Position::new(sheet)), start, count)
            }
        },
        "sheet.reorder" | "sheet.insert" | "sheet.remove" | "sheet.restore" => {
            apply_tab_semantic(source, transaction, operation)
        },
        "row.layout" | "column.layout" => {
            apply_layout_semantic(source, transaction, operation, sheet_names)
        },
        "sheet.rename" => {
            let sheet = parse_sheet_target(&operation.target, "name")?;
            let sheet_index = verify_semantic_sheet(source, sheet_names, sheet, operation)?;
            let name = operation
                .value
                .as_str()
                .ok_or_else(|| Error::InvalidData("sheet rename value is not text".into()))?;
            transaction.rename_sheet(Selector::Position(Position::new(sheet)), name)?;
            sheet_names[sheet_index] = name.to_string();
            Ok(())
        },
        _ => Err(Error::InvalidData(format!(
            "unsupported XLS cell patch operation {:?}",
            operation.op
        ))),
    }
}

fn apply_tab_semantic(
    source: &Snapshot,
    transaction: &mut Edit,
    operation: &PatchOperation,
) -> Result<()> {
    let catalog = tabs::catalog(&source.inner.workbook_stream)?;
    let expected = operation
        .preconditions
        .get("tab_names")
        .and_then(serde_json::Value::as_array)
        .ok_or_else(|| Error::InvalidData("tab patch has no tab-name precondition".into()))?;
    if expected.len() != catalog.len()
        || expected
            .iter()
            .zip(&catalog)
            .any(|(name, tab)| name.as_str() != Some(tab.name.as_str()))
    {
        return Err(Error::PatchConflict {
            operation: "edit cells",
            reason: "tab patch precondition does not match the workbook tab catalog",
        });
    }
    let name = || {
        operation
            .value
            .get("name")
            .and_then(serde_json::Value::as_str)
            .ok_or_else(|| Error::InvalidData("tab patch has no sheet name".into()))
    };
    match operation.op.as_str() {
        "sheet.reorder" => {
            if operation.target != tabs::SEMANTIC_TARGET {
                return Err(Error::InvalidData("tab reorder target is malformed".into()));
            }
            let order = operation
                .value
                .as_array()
                .ok_or_else(|| Error::InvalidData("tab reorder value is not a list".into()))?
                .iter()
                .map(|value| {
                    value
                        .as_u64()
                        .and_then(|value| usize::try_from(value).ok())
                        .map(Selector::from)
                        .ok_or_else(|| Error::InvalidData("tab reorder position is invalid".into()))
                })
                .collect::<Result<Vec<_>>>()?;
            transaction.reorder_sheets(&order)
        },
        "sheet.insert" => {
            let position = tabs::parse_semantic_tab_target(&operation.target)?;
            transaction.insert_sheet(Position::new(position), name()?)
        },
        "sheet.remove" => {
            let position = tabs::parse_semantic_tab_target(&operation.target)?;
            let tab = catalog
                .get(position)
                .ok_or_else(|| Error::WorksheetNotFound("tab patch position".into()))?;
            if tab.name != name()? {
                return Err(Error::PatchConflict {
                    operation: "edit cells",
                    reason: "tab removal precondition names another tab",
                });
            }
            if let Some(fingerprint) = operation.preconditions.get("fingerprint") {
                if fingerprint.as_str()
                    != Some(tabs::fingerprint(&source.inner.workbook_stream, tab).as_str())
                {
                    return Err(Error::PatchConflict {
                        operation: "edit cells",
                        reason: "tab removal precondition does not match the tab content",
                    });
                }
            } else if operation.preconditions.get("empty") == Some(&serde_json::Value::Bool(true)) {
                let empty = source
                    .resolve_sheet(Selector::Position(Position::new(position)))?
                    .is_some_and(|sheet| source.inner.sheets[sheet].entries.is_empty());
                if !empty {
                    return Err(Error::PatchConflict {
                        operation: "edit cells",
                        reason: "tab removal precondition requires an empty worksheet",
                    });
                }
            } else {
                return Err(Error::InvalidData(
                    "tab removal patch has no content precondition".into(),
                ));
            }
            transaction.remove_sheet(Selector::Position(Position::new(position)))
        },
        _ => Err(Error::UnsupportedFeature(
            "restoring a removed worksheet requires the exact artifact patch".into(),
        )),
    }
}

fn row_state_json(state: layout::RowState) -> serde_json::Value {
    serde_json::json!({
        "present": state.present,
        "height": state.height,
        "custom": state.custom,
        "hidden": state.hidden,
    })
}

fn column_state_json(state: layout::ColumnState) -> serde_json::Value {
    serde_json::json!({
        "present": state.present,
        "width": state.width,
        "user_set": state.user_set,
        "hidden": state.hidden,
    })
}

fn json_bool(value: &serde_json::Value, key: &str) -> Result<bool> {
    value
        .get(key)
        .and_then(serde_json::Value::as_bool)
        .ok_or_else(|| Error::InvalidData("layout patch state is malformed".into()))
}

fn json_u16(value: &serde_json::Value, key: &str) -> Result<u16> {
    value
        .get(key)
        .and_then(serde_json::Value::as_u64)
        .and_then(|value| u16::try_from(value).ok())
        .ok_or_else(|| Error::InvalidData("layout patch state is malformed".into()))
}

fn parse_row_state(value: &serde_json::Value) -> Result<layout::RowState> {
    Ok(layout::RowState {
        present: json_bool(value, "present")?,
        height: json_u16(value, "height")?,
        custom: json_bool(value, "custom")?,
        hidden: json_bool(value, "hidden")?,
    })
}

fn parse_column_state(value: &serde_json::Value) -> Result<layout::ColumnState> {
    Ok(layout::ColumnState {
        present: json_bool(value, "present")?,
        width: json_u16(value, "width")?,
        user_set: json_bool(value, "user_set")?,
        hidden: json_bool(value, "hidden")?,
    })
}

fn append_layout_semantic(
    snapshot: &Snapshot,
    change: &layout::LayoutChange,
    limits: PatchLimits,
    operations: &mut Vec<ReversibleOperation>,
) -> Result<()> {
    let sheet_data = &snapshot.inner.sheets[change.sheet()];
    let (op, target, before, after) = match change {
        layout::LayoutChange::Row {
            row, before, after, ..
        } => (
            "row.layout",
            format!("sheet/{}/row/{row}", sheet_data.workbook_index),
            row_state_json(*before),
            row_state_json(*after),
        ),
        layout::LayoutChange::Column {
            column,
            before,
            after,
            ..
        } => (
            "column.layout",
            format!("sheet/{}/column/{column}", sheet_data.workbook_index),
            column_state_json(*before),
            column_state_json(*after),
        ),
    };
    let preconditions = |state: serde_json::Value| {
        let mut map = BTreeMap::new();
        map.insert(
            "sheet_name".to_string(),
            serde_json::Value::String(sheet_data.name.clone()),
        );
        map.insert("state".to_string(), state);
        map
    };
    let forward = PatchOperation::new(
        limits,
        op,
        target.clone(),
        preconditions(before.clone()),
        after.clone(),
    )
    .map_err(patch_error)?;
    let inverse = PatchOperation::new(limits, op, target, preconditions(after), before)
        .map_err(patch_error)?;
    operations.push(ReversibleOperation::new(forward, inverse));
    Ok(())
}

fn apply_layout_semantic(
    source: &Snapshot,
    transaction: &mut Edit,
    operation: &PatchOperation,
    sheet_names: &[String],
) -> Result<()> {
    let mut parts = operation.target.split('/');
    let sheet_position = (parts.next() == Some("sheet"))
        .then(|| parts.next()?.parse::<usize>().ok())
        .flatten()
        .ok_or_else(|| Error::InvalidData("layout patch target is malformed".into()))?;
    let axis = parts.next();
    let index = parts
        .next()
        .and_then(|value| value.parse::<u16>().ok())
        .ok_or_else(|| Error::InvalidData("layout patch target is malformed".into()))?;
    if parts.next().is_some() {
        return Err(Error::InvalidData(
            "layout patch target has trailing data".into(),
        ));
    }
    let sheet = verify_semantic_sheet(source, sheet_names, sheet_position, operation)?;
    let expected = operation
        .preconditions
        .get("state")
        .ok_or_else(|| Error::InvalidData("layout patch has no state".into()))?;
    let stream = transaction.sheet_stream(sheet)?;
    let selector = Selector::Position(Position::new(sheet_position));
    match (operation.op.as_str(), axis) {
        ("row.layout", Some("row")) => {
            if layout::row_state(stream, index)? != parse_row_state(expected)? {
                return Err(Error::PatchConflict {
                    operation: "edit cells",
                    reason: "row layout patch precondition is stale",
                });
            }
            let after = parse_row_state(&operation.value)?;
            transaction.stage_row_layout(selector, index, |state| *state = after)
        },
        ("column.layout", Some("column")) => {
            let column = u8::try_from(index)
                .map_err(|_error| Error::InvalidData("layout patch column exceeds u8".into()))?;
            if layout::column_state(stream, column)? != parse_column_state(expected)? {
                return Err(Error::PatchConflict {
                    operation: "edit cells",
                    reason: "column layout patch precondition is stale",
                });
            }
            let after = parse_column_state(&operation.value)?;
            transaction.stage_column_layout(selector, column, |state| *state = after)
        },
        _ => Err(Error::InvalidData(
            "layout patch target is malformed".into(),
        )),
    }
}

/// Every non-no-op layout change reads back from the published candidate.
fn verify_layout_readback(snapshot: &Snapshot, changes: &[&layout::LayoutChange]) -> Result<()> {
    if changes.is_empty() {
        return Ok(());
    }
    let catalog = tabs::catalog(&snapshot.inner.workbook_stream)?;
    for change in changes {
        let workbook_index = snapshot
            .inner
            .sheets
            .get(change.sheet())
            .ok_or(Error::Internal("layout worksheet disappeared"))?
            .workbook_index;
        let tab = catalog
            .get(workbook_index)
            .ok_or(Error::Internal("layout worksheet tab disappeared"))?;
        let stream = &snapshot.inner.workbook_stream[tab.start..tab.end];
        let matches = match change {
            layout::LayoutChange::Row { row, after, .. } => {
                let observed = layout::row_state(stream, *row)?;
                if after.present {
                    observed == *after
                } else {
                    !observed.present || (!observed.custom && !observed.hidden)
                }
            },
            layout::LayoutChange::Column { column, after, .. } => {
                layout::column_state(stream, *column)? == *after
            },
        };
        if !matches {
            return Err(Error::Internal(
                "row or column layout failed semantic readback",
            ));
        }
    }
    Ok(())
}

fn apply_sst_semantic(transaction: &mut Edit, operation: &PatchOperation) -> Result<()> {
    let text = operation
        .value
        .as_str()
        .ok_or_else(|| Error::InvalidData("SST resource value is not text".into()))?;
    if operation.target != format!("resource/sst/{}", text_fingerprint(text)) {
        return Err(Error::InvalidData(
            "SST resource target is malformed".into(),
        ));
    }
    let expected = operation
        .preconditions
        .get("present")
        .and_then(serde_json::Value::as_bool)
        .ok_or_else(|| Error::InvalidData("SST resource has no presence precondition".into()))?;
    if transaction.has_shared_string(text) != expected {
        return Err(Error::PatchConflict {
            operation: "edit cells",
            reason: "SST resource precondition is stale",
        });
    }
    transaction.stage_resource(ResourceChange::SharedString {
        text: text.to_string(),
        insert: operation.op == "sst.intern",
    })
}

fn apply_rich_sst_semantic(transaction: &mut Edit, operation: &PatchOperation) -> Result<()> {
    let object = operation
        .value
        .as_object()
        .ok_or_else(|| Error::InvalidData("rich SST resource value is not an object".into()))?;
    let text = object
        .get("text")
        .and_then(serde_json::Value::as_str)
        .ok_or_else(|| Error::InvalidData("rich SST resource text is malformed".into()))?;
    let encoded_runs = object
        .get("runs")
        .and_then(serde_json::Value::as_array)
        .ok_or_else(|| Error::InvalidData("rich SST resource runs are malformed".into()))?;
    let mut formatting_runs = Vec::new();
    formatting_runs
        .try_reserve_exact(encoded_runs.len())
        .map_err(|_error| Error::Allocation("retaining rich SST semantic runs"))?;
    for encoded in encoded_runs {
        let pair = encoded
            .as_array()
            .filter(|pair| pair.len() == 2)
            .ok_or_else(|| Error::InvalidData("rich SST run is malformed".into()))?;
        let character_index = pair[0]
            .as_u64()
            .and_then(|value| u16::try_from(value).ok())
            .ok_or_else(|| Error::InvalidData("rich SST character index exceeds u16".into()))?;
        let font_index = pair[1]
            .as_u64()
            .and_then(|value| u16::try_from(value).ok())
            .ok_or_else(|| Error::InvalidData("rich SST font index exceeds u16".into()))?;
        formatting_runs.push(crate::records::SharedStringFormatRun {
            character_index,
            font_index,
        });
    }
    if operation.target
        != format!(
            "resource/sst-rich/{}",
            rich_text_fingerprint(text, &formatting_runs)
        )
    {
        return Err(Error::InvalidData(
            "rich SST resource target is malformed".into(),
        ));
    }
    let expected = operation
        .preconditions
        .get("present")
        .and_then(serde_json::Value::as_bool)
        .ok_or_else(|| {
            Error::InvalidData("rich SST resource has no presence precondition".into())
        })?;
    if transaction.has_rich_shared_string(text, &formatting_runs) != expected {
        return Err(Error::PatchConflict {
            operation: "edit cells",
            reason: "rich SST resource precondition is stale",
        });
    }
    transaction.stage_resource(ResourceChange::RichSharedString {
        text: text.to_string(),
        formatting_runs,
        insert: operation.op == "sst.intern-rich",
    })
}

fn apply_xf_semantic(transaction: &mut Edit, operation: &PatchOperation) -> Result<()> {
    let encoded_index = operation
        .target
        .strip_prefix("resource/xf/")
        .ok_or_else(|| Error::InvalidData("XF resource target is malformed".into()))?;
    let index = encoded_index
        .parse::<u16>()
        .map(StyleIndex)
        .map_err(|error| Error::InvalidData(format!("invalid XF resource index: {error}")))?;
    if encoded_index != format!("{:05}", index.get()) {
        return Err(Error::InvalidData(
            "XF resource target is non-canonical".into(),
        ));
    }
    let payload = operation
        .value
        .as_str()
        .ok_or_else(|| Error::InvalidData("XF resource payload is not hexadecimal".into()))
        .and_then(parse_hex)?;
    let expected = operation
        .preconditions
        .get("present")
        .and_then(serde_json::Value::as_bool)
        .ok_or_else(|| Error::InvalidData("XF resource has no presence precondition".into()))?;
    let present = transaction
        .effective_xf_payload(index)
        .is_some_and(|candidate| candidate == payload.as_slice());
    if present != expected {
        return Err(Error::PatchConflict {
            operation: "edit cells",
            reason: "XF resource precondition is stale",
        });
    }
    transaction.stage_resource(ResourceChange::ExtendedFormat {
        index,
        payload,
        insert: matches!(operation.op.as_str(), "xf.author" | "xf.duplicate"),
    })
}

fn apply_formula_semantic(
    source: &Snapshot,
    transaction: &mut Edit,
    operation: &PatchOperation,
    sheet_names: &[String],
) -> Result<()> {
    let mut parts = operation.target.split('/');
    if parts.next() != Some("sheet") {
        return Err(Error::InvalidData("Formula target is malformed".into()));
    }
    let sheet_position = parts
        .next()
        .and_then(|part| part.parse::<usize>().ok())
        .ok_or_else(|| Error::InvalidData("Formula sheet target is malformed".into()))?;
    if parts.next() != Some("formula") {
        return Err(Error::InvalidData("Formula target is malformed".into()));
    }
    let row = parts
        .next()
        .and_then(|part| part.parse::<u16>().ok())
        .ok_or_else(|| Error::InvalidData("Formula row target is malformed".into()))?;
    let column = parts
        .next()
        .and_then(|part| part.parse::<u8>().ok())
        .ok_or_else(|| Error::InvalidData("Formula column target is malformed".into()))?;
    if parts.next().is_some() {
        return Err(Error::InvalidData(
            "Formula target has trailing data".into(),
        ));
    }
    let sheet = verify_semantic_sheet(source, sheet_names, sheet_position, operation)?;
    let object = operation
        .value
        .as_object()
        .ok_or_else(|| Error::InvalidData("Formula resource value is malformed".into()))?;
    let tokens = object
        .get("tokens")
        .and_then(serde_json::Value::as_str)
        .ok_or_else(|| Error::InvalidData("Formula resource tokens are malformed".into()))
        .and_then(parse_hex)?;
    let style = object
        .get("style")
        .and_then(serde_json::Value::as_u64)
        .and_then(|value| u16::try_from(value).ok())
        .map(StyleIndex)
        .ok_or_else(|| Error::InvalidData("Formula resource style is malformed".into()))?;
    let expected_xf = object
        .get("target_xf")
        .and_then(serde_json::Value::as_str)
        .ok_or_else(|| Error::InvalidData("Formula resource has no XF dependency".into()))?;
    let actual_xf = transaction
        .effective_xf_payload(style)
        .map(bytes_hex)
        .ok_or_else(|| {
            Error::refused(
                RefusalKind::NotFound,
                "edit cells",
                "Formula XF resource is absent",
            )
        })?;
    if actual_xf != expected_xf {
        return Err(Error::PatchConflict {
            operation: "edit cells",
            reason: "Formula XF resource dependency is stale",
        });
    }
    let expected = operation
        .preconditions
        .get("present")
        .and_then(serde_json::Value::as_bool)
        .ok_or_else(|| {
            Error::InvalidData("Formula resource has no presence precondition".into())
        })?;
    let reference = Reference::new(u32::from(row), u32::from(column))?;
    let sheet_data = source.inner.sheets.get(sheet).ok_or(Error::PatchConflict {
        operation: "edit cells",
        reason: "Formula semantic sheet dependency is stale",
    })?;
    let current = sheet_data
        .unique_entry_index(reference)?
        .map(|index| &sheet_data.entries[index]);
    let replaced = transaction.staged_removal(sheet, reference);
    if !expected && current.is_some() && !replaced {
        return Err(Error::PatchConflict {
            operation: "edit cells",
            reason: "Formula insertion precondition target is occupied",
        });
    }
    let present = match current {
        Some(_) if replaced => false,
        Some(entry) => authored_formula_record_matches(source, entry, reference, style, &tokens)?,
        None => false,
    };
    if present != expected {
        return Err(Error::PatchConflict {
            operation: "edit cells",
            reason: "Formula resource precondition is stale",
        });
    }
    transaction.stage_resource(ResourceChange::FormulaCell {
        sheet,
        reference,
        style,
        tokens,
        insert: operation.op == "formula.insert",
    })
}

fn parse_hex(value: &str) -> Result<Vec<u8>> {
    if !value.len().is_multiple_of(2)
        || value
            .bytes()
            .any(|byte| !byte.is_ascii_digit() && !(b'a'..=b'f').contains(&byte))
    {
        return Err(Error::InvalidData(
            "resource hexadecimal is non-canonical".into(),
        ));
    }
    value
        .as_bytes()
        .as_chunks::<2>()
        .0
        .iter()
        .map(|pair| {
            let text = std::str::from_utf8(pair)
                .map_err(|error| Error::InvalidData(format!("invalid resource hex: {error}")))?;
            u8::from_str_radix(text, 16)
                .map_err(|error| Error::InvalidData(format!("invalid resource hex: {error}")))
        })
        .collect()
}

fn apply_fixed_semantic(
    source: &Snapshot,
    transaction: &mut Edit,
    operation: &PatchOperation,
    sheet_names: &[String],
) -> Result<()> {
    let (sheet_position, reference) = parse_semantic_target(&operation.target)?;
    let sheet_index = verify_semantic_sheet(source, sheet_names, sheet_position, operation)?;
    let sheet = &source.inner.sheets[sheet_index];
    let entry_index = sheet.unique_entry_index(reference)?.ok_or_else(|| {
        Error::refused(
            RefusalKind::NotFound,
            "edit cells",
            "semantic patch cell is absent",
        )
    })?;
    let cell = &sheet.entries[entry_index].cell;
    let expected = operation
        .preconditions
        .get("state")
        .ok_or_else(|| Error::InvalidData("cell patch has no state precondition".into()))?;
    let (expected_storage, expected_value) = parse_cell_state(expected)?;
    if cell.storage != expected_storage || !values_equal(&cell.value, &expected_value) {
        return Err(Error::PatchConflict {
            operation: "edit cells",
            reason: "semantic patch cell precondition is stale",
        });
    }
    verify_formula_dependency(source, &sheet.entries[entry_index], operation)?;
    let (storage, value) = parse_cell_state(&operation.value)?;
    // A durable inverse may remove the forward-only SST resource before it
    // restores the original text. Let the native setter re-intern that text
    // atomically, rather than requiring the compacted target to retain a dead
    // shared-string entry solely for semantic replay.
    if matches!(cell.storage, Storage::LabelSst | Storage::Rk)
        && storage == Storage::LabelSst
        && matches!(&value, Value::Text(_))
    {
        return transaction.set_value(
            Selector::Position(Position::new(sheet_position)),
            reference,
            value,
        );
    }
    let represented = target_storage(cell.storage, &value, |text| {
        transaction.has_shared_string(text)
    })?;
    if represented != storage {
        return Err(Error::InvalidData(
            "cell patch target storage disagrees with its value".into(),
        ));
    }
    if matches!(cell.value, Value::FormulaCache(FormulaCache::String(_)))
        || matches!(value, Value::FormulaCache(FormulaCache::String(_)))
    {
        return transaction.set_value(
            Selector::Position(Position::new(sheet_position)),
            reference,
            value,
        );
    }
    transaction.stage(sheet_index, entry_index, storage, value)
}

fn apply_structural_cell_semantic(
    source: &Snapshot,
    transaction: &mut Edit,
    operation: &PatchOperation,
    sheet_names: &[String],
) -> Result<()> {
    let (sheet_position, reference) = parse_semantic_target(&operation.target)?;
    let sheet_index = verify_semantic_sheet(source, sheet_names, sheet_position, operation)?;
    let sheet = &source.inner.sheets[sheet_index];
    // A Formula this patch already removed no longer occupies its cell.
    let vacated = transaction.staged_formula_removal(sheet_index, reference);
    let current = sheet
        .unique_entry_index(reference)?
        .filter(|_| !vacated)
        .map(|index| &sheet.entries[index])
        .map(|entry| {
            (
                entry.cell.storage,
                entry.cell.value.clone(),
                entry.cell.style,
            )
        });
    let expected = operation
        .preconditions
        .get("state")
        .ok_or_else(|| Error::InvalidData("structural cell patch has no state".into()))?;
    let expected = parse_optional_cell_state(expected)?;
    if !optional_cell_states_equal(current.as_ref(), expected.as_ref()) {
        return Err(Error::PatchConflict {
            operation: "edit cells",
            reason: "structural cell patch precondition is stale",
        });
    }
    if let Some(index) = sheet.unique_entry_index(reference)?.filter(|_| !vacated) {
        let entry = &sheet.entries[index];
        verify_formula_dependency(source, entry, operation)?;
    }
    let after = parse_optional_cell_state(&operation.value)?;
    if let Some((_, _, style)) = &after {
        let expected_xf = operation
            .preconditions
            .get("target_xf")
            .and_then(serde_json::Value::as_str)
            .ok_or_else(|| {
                Error::InvalidData("structural cell patch has no XF dependency".into())
            })?;
        let actual_xf = transaction
            .effective_xf_payload(*style)
            .map(bytes_hex)
            .ok_or_else(|| {
                Error::refused(
                    RefusalKind::NotFound,
                    "edit cells",
                    "structural cell XF resource is absent",
                )
            })?;
        if actual_xf != expected_xf {
            return Err(Error::PatchConflict {
                operation: "edit cells",
                reason: "structural cell XF resource dependency is stale",
            });
        }
    }
    match (current, after) {
        (None, Some((storage, value, style))) => {
            if storage == Storage::MulRk {
                let Value::Number(number) = value else {
                    return Err(Error::InvalidData(
                        "MulRk structural patch value is not numeric".into(),
                    ));
                };
                if encode_rk(number).is_none() {
                    return Err(Error::InvalidData(
                        "MulRk structural patch value is not RK-representable".into(),
                    ));
                }
                return transaction.stage_structural(StructuralChange::Cell {
                    sheet: sheet_index,
                    reference,
                    before: None,
                    after: Some((Storage::MulRk, Value::Number(number), style)),
                });
            }
            if storage == Storage::LabelSst
                && let Value::Text(text) = &value
                && !transaction.has_shared_string(text)
            {
                transaction.stage_resource(ResourceChange::SharedString {
                    text: text.clone(),
                    insert: true,
                })?;
            }
            if storage_for_new_value(&value, |text| transaction.has_shared_string(text))? != storage
            {
                return Err(Error::InvalidData(
                    "structural cell storage disagrees with its value".into(),
                ));
            }
            transaction.insert_cell_with_style(
                Selector::Position(Position::new(sheet_position)),
                reference,
                value,
                style,
            )
        },
        (Some(_), None) => {
            transaction.remove_cell(Selector::Position(Position::new(sheet_position)), reference)
        },
        (Some(before), Some(after)) if before.0 == after.0 && values_equal(&before.1, &after.1) => {
            transaction.set_style(
                Selector::Position(Position::new(sheet_position)),
                reference,
                after.2,
            )
        },
        (Some(before), Some(after))
            if before.0 == Storage::Formula && after.0 == Storage::Formula =>
        {
            let Value::FormulaCache(cache) = &after.1 else {
                return Err(Error::InvalidData(
                    "Formula structural patch value is not a cache".into(),
                ));
            };
            if !valid_formula_cache(cache) {
                return Err(Error::InvalidData(
                    "Formula structural patch cache is not representable".into(),
                ));
            }
            transaction.stage_structural(StructuralChange::Cell {
                sheet: sheet_index,
                reference,
                before: Some(before),
                after: Some(after),
            })
        },
        (Some(before), Some(after)) => {
            let (storage, value, style) = after;
            if storage == Storage::LabelSst
                && let Value::Text(text) = &value
                && !transaction.has_shared_string(text)
            {
                transaction.stage_resource(ResourceChange::SharedString {
                    text: text.clone(),
                    insert: true,
                })?;
            }
            let represented = if storage == Storage::MulRk {
                matches!(&value, Value::Number(number) if encode_rk(*number).is_some())
            } else {
                storage_for_new_value(&value, |text| transaction.has_shared_string(text))
                    .map(|candidate| candidate == storage)
                    .unwrap_or(false)
            };
            if !represented {
                return Err(Error::InvalidData(
                    "structural cell storage disagrees with its value".into(),
                ));
            }
            transaction.stage_structural(StructuralChange::Cell {
                sheet: sheet_index,
                reference,
                before: Some(before),
                after: Some((storage, value, style)),
            })
        },
        _ => Err(Error::InvalidData(
            "structural cell patch requests an unsupported replacement".into(),
        )),
    }
}

fn verify_semantic_sheet(
    source: &Snapshot,
    sheet_names: &[String],
    sheet_position: usize,
    operation: &PatchOperation,
) -> Result<usize> {
    let sheet_index = source
        .resolve_sheet(Selector::Position(Position::new(sheet_position)))?
        .ok_or_else(|| {
            Error::refused(
                RefusalKind::NotFound,
                "edit cells",
                "semantic patch worksheet is absent",
            )
        })?;
    let expected_name = operation
        .preconditions
        .get("sheet_name")
        .and_then(serde_json::Value::as_str)
        .ok_or_else(|| Error::InvalidData("patch has no sheet-name precondition".into()))?;
    if sheet_names.get(sheet_index).map(String::as_str) != Some(expected_name) {
        return Err(Error::PatchConflict {
            operation: "edit cells",
            reason: "semantic patch worksheet identity is stale",
        });
    }
    Ok(sheet_index)
}

fn semantic_sheet_names(source: &Snapshot) -> Vec<String> {
    source
        .inner
        .sheets
        .iter()
        .map(|sheet| sheet.name.clone())
        .collect()
}

fn optional_cell_state(state: Option<&(Storage, Value, StyleIndex)>) -> serde_json::Value {
    state.map_or(serde_json::Value::Null, |(storage, value, style)| {
        serde_json::json!({
            "storage": storage_name(*storage),
            "style": style.get(),
            "value": encode_value(value),
        })
    })
}

fn parse_optional_cell_state(
    state: &serde_json::Value,
) -> Result<Option<(Storage, Value, StyleIndex)>> {
    if state.is_null() {
        return Ok(None);
    }
    let object = state
        .as_object()
        .ok_or_else(|| Error::InvalidData("structural cell state is not an object".into()))?;
    let storage = parse_storage(
        object
            .get("storage")
            .ok_or_else(|| Error::InvalidData("structural cell state has no storage".into()))?,
    )?;
    let style = object
        .get("style")
        .and_then(serde_json::Value::as_u64)
        .and_then(|value| u16::try_from(value).ok())
        .map(StyleIndex)
        .ok_or_else(|| Error::InvalidData("structural cell style is malformed".into()))?;
    let value = parse_value(
        object
            .get("value")
            .ok_or_else(|| Error::InvalidData("structural cell state has no value".into()))?,
    )?;
    Ok(Some((storage, value, style)))
}

fn parse_axis_target(target: &str, axis: &str) -> Result<(usize, u16)> {
    let mut parts = target.split('/');
    if parts.next() != Some("sheet") {
        return Err(Error::InvalidData(
            "axis patch target has invalid prefix".into(),
        ));
    }
    let sheet = parts
        .next()
        .ok_or_else(|| Error::InvalidData("axis patch target has no sheet".into()))?
        .parse::<usize>()
        .map_err(|error| Error::InvalidData(format!("invalid axis sheet: {error}")))?;
    if parts.next() != Some(axis) {
        return Err(Error::InvalidData(
            "axis patch target has wrong axis".into(),
        ));
    }
    let start = parts
        .next()
        .ok_or_else(|| Error::InvalidData("axis patch target has no start".into()))?
        .parse::<u16>()
        .map_err(|error| Error::InvalidData(format!("invalid axis start: {error}")))?;
    if parts.next().is_some() {
        return Err(Error::InvalidData(
            "axis patch target has trailing data".into(),
        ));
    }
    Ok((sheet, start))
}

fn parse_sheet_target(target: &str, suffix: &str) -> Result<usize> {
    let mut parts = target.split('/');
    if parts.next() != Some("sheet") {
        return Err(Error::InvalidData(
            "sheet patch target has invalid prefix".into(),
        ));
    }
    let sheet = parts
        .next()
        .ok_or_else(|| Error::InvalidData("sheet patch target has no position".into()))?
        .parse::<usize>()
        .map_err(|error| Error::InvalidData(format!("invalid sheet position: {error}")))?;
    if parts.next() != Some(suffix) || parts.next().is_some() {
        return Err(Error::InvalidData("sheet patch target is malformed".into()));
    }
    Ok(sheet)
}

fn parse_count(value: &serde_json::Value, label: &str) -> Result<u16> {
    value
        .get("count")
        .and_then(serde_json::Value::as_u64)
        .and_then(|count| u16::try_from(count).ok())
        .filter(|count| *count != 0)
        .ok_or_else(|| Error::InvalidData(format!("{label} patch count is malformed")))
}

impl fmt::Debug for SemanticPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SemanticPatch")
            .field("operations", &self.len())
            .finish()
    }
}

/// Reversible, exact-source-checked replacement of one XLS artifact.
#[derive(Clone)]
pub(crate) struct Patch {
    before: Arc<Vec<u8>>,
    after: Arc<Vec<u8>>,
    semantic: SemanticPatch,
}

impl Patch {
    fn new(before: Arc<Vec<u8>>, after: Arc<Vec<u8>>, semantic: SemanticPatch) -> Self {
        Self {
            before,
            after,
            semantic,
        }
    }

    /// Returns whether the patch is a byte-exact no-op.
    #[must_use]
    pub(crate) fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Returns the exact required source artifact.
    #[must_use]
    pub(crate) fn before(&self) -> &[u8] {
        &self.before
    }

    /// Returns the exact produced artifact.
    #[must_use]
    pub(crate) fn after(&self) -> &[u8] {
        &self.after
    }

    /// Durable semantic operation patch for this exact artifact transition.
    #[must_use]
    pub(crate) const fn semantic(&self) -> &SemanticPatch {
        &self.semantic
    }

    /// Returns the saturating byte charge for this exact patch.
    ///
    /// The exact before and after artifact buffers are charged in addition to
    /// the forward and inverse semantic data retained by [`Self::semantic`].
    /// No serialization or inverse cloning is performed.
    #[must_use]
    pub(crate) fn retained_bytes(&self) -> usize {
        self.before
            .len()
            .saturating_add(self.after.len())
            .saturating_add(self.semantic.retained_bytes())
    }

    /// Applies this patch only to its exact immutable source artifact.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] for a stale source, or a typed package
    /// validation error if the retained target can no longer be reopened.
    pub(crate) fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.bytes() != self.before() {
            return Err(Error::PatchConflict {
                operation: "edit cells",
                reason: "XLS cell-value patch source does not match its base snapshot",
            });
        }
        if self.is_noop() {
            return Ok(source.clone());
        }
        Snapshot::from_slice_with_limits_and_cfb(
            &self.after,
            source.inner.limits,
            source.inner.cfb_limits,
        )
    }

    /// Returns the exact inverse patch.
    #[must_use]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            semantic: self.semantic.inverse(),
        }
    }
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("before_bytes", &self.before.len())
            .field("after_bytes", &self.after.len())
            .field("is_noop", &self.is_noop())
            .finish()
    }
}

/// A seekable adapter for semantic validation over a lazy composed CFB view.
///
/// `raw::Reader` is intentionally kept generic over `Read + Seek`, while the
/// CFB overlay plan exposes a positional `ReadAt` source. This adapter keeps
/// those contracts separate and allocates no artifact-sized buffer.
struct ComposedPositionalReader {
    source: ComposedOverlaySource,
    position: u64,
}

impl ComposedPositionalReader {
    fn new(source: ComposedOverlaySource) -> Self {
        Self {
            source,
            position: 0,
        }
    }
}

impl Read for ComposedPositionalReader {
    fn read(&mut self, output: &mut [u8]) -> std::io::Result<usize> {
        let count = self.source.read_at(self.position, output)?;
        self.position = self
            .position
            .checked_add(u64::try_from(count).map_err(|_error| {
                std::io::Error::other("composed positional reader count exceeds u64")
            })?)
            .ok_or_else(|| std::io::Error::other("composed positional reader offset overflow"))?;
        Ok(count)
    }
}

impl Seek for ComposedPositionalReader {
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
                "composed positional reader seek is outside the source",
            ));
        }
        self.position = u64::try_from(target).map_err(|_error| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "composed positional reader seek exceeds u64",
            )
        })?;
        Ok(self.position)
    }
}

fn commit_source_backed_numeric_plan(transaction: Edit) -> Result<SourceBackedPlanCommit> {
    let Edit {
        source,
        changes,
        structural_changes,
        resource_changes,
        tab_change,
        layout_changes,
        ..
    } = transaction;

    if !structural_changes.is_empty()
        || !resource_changes.is_empty()
        || tab_change.is_some()
        || !layout_changes.is_empty()
    {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "source-backed numeric plan accepts no structural or resource edits",
        ));
    }
    if !changes_are_fixed_numeric(&source, &changes) {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "source-backed numeric plan requires unchanged Number, RK, or MulRk storage",
        ));
    }

    // The complete source Workbook validation already ran when this immutable
    // snapshot was opened. Reuse its private policy facts here; the target is
    // still checked independently below over the composed positional reader.
    source.inner.source_policy.require()?;

    let publisher = SourceBackedOverlayPublisher::open_owned_vec(
        Arc::clone(&source.inner.bytes),
        source.inner.source_version,
        shared_ole_limits(source.inner.cfb_limits)?,
    )
    .map_err(source_backed_overlay_error)?;
    let source_version = publisher
        .source_version()
        .map_err(source_backed_overlay_error)?;
    if source_version != source.inner.source_version {
        return Err(Error::Internal(
            "source-backed numeric plan lost its source lineage",
        ));
    }
    require_macro_free_container(&publisher)?;

    let mut splices = Vec::new();
    splices
        .try_reserve_exact(changes.len())
        .map_err(|_error| Error::Allocation("staging source-backed numeric plan splices"))?;
    let mut replacement_bytes = 0_u64;
    for change in &changes {
        if !change_is_effective(&source, change) {
            continue;
        }
        let splice = fixed_numeric_splice(&source, change)?;
        replacement_bytes = replacement_bytes
            .checked_add(u64::try_from(splice.replacement().len()).map_err(|_error| {
                Error::InvalidData("numeric plan replacement length exceeds u64".into())
            })?)
            .ok_or_else(|| {
                Error::InvalidData("numeric plan replacement bytes overflow u64".into())
            })?;
        splices.push(splice);
    }

    let splice_count = splices.len();
    let (plan, validated_target_version) =
        publisher.plan_splices_with_owner(splices, StreamSpliceLimits::default(), |candidate| {
            verify_source_backed_numeric_plan_target(candidate, &source, &changes)
        })?;

    // Generic positional adapters retain complete fingerprints around the
    // semantic readback so a dishonest stable version token cannot hide a
    // mutation. This source is a sealed immutable Arc, so the CFB planner may
    // omit only the redundant final scan after validating the composed view.
    let target_version = validated_target_version.unwrap_or(source_version);
    let source_bytes = u64::try_from(source.bytes().len())
        .map_err(|_error| Error::InvalidData("source CFB length exceeds u64".into()))?;
    let source_workbook_bytes = u64::try_from(source.workbook_stream().len())
        .map_err(|_error| Error::InvalidData("source Workbook length exceeds u64".into()))?;
    let diagnostics = SourceBackedDiagnostics {
        changed_cells: changes
            .iter()
            .filter(|change| change_is_effective(&source, change))
            .count(),
        touched_streams: usize::from(!plan.is_noop()),
        splice_count,
        replacement_bytes,
        changed_spans: plan.changed_spans(),
        source_bytes,
        source_workbook_bytes,
        target_workbook_bytes: source_workbook_bytes,
        source_version,
        target_version,
        source_fingerprint: plan.source_fingerprint(),
        target_fingerprint: plan.target_fingerprint(),
        operation_shape: plan.operation_shape(),
    };
    Ok(SourceBackedPlanCommit {
        source,
        plan,
        diagnostics,
    })
}

fn verify_source_backed_numeric_plan_target(
    candidate: &ComposedOverlaySource,
    source: &Snapshot,
    changes: &[Change],
) -> Result<SourceVersion> {
    let workbook = Reader::validation_only_with_cfb_limits(
        ComposedPositionalReader::new(candidate.clone()),
        numeric_readback_cells(source, changes)?,
        source.inner.cfb_limits,
    )?;
    require_public_worksheet_coverage(&workbook, &source.inner.sheets)?;
    require_unprotected_workbook(&workbook)?;
    require_macro_free_workbook(&workbook)?;
    verify_public_numeric_readback(&workbook, source, changes)?;
    drop(workbook);
    candidate.version().map_err(Error::Io)
}

#[cfg_attr(
    not(test),
    allow(
        dead_code,
        reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
    )
)]
fn commit_source_backed_numeric(transaction: Edit) -> Result<SourceBackedCommit> {
    let Edit {
        source,
        changes,
        structural_changes,
        resource_changes,
        tab_change,
        layout_changes,
        ..
    } = transaction;

    if !structural_changes.is_empty()
        || !resource_changes.is_empty()
        || tab_change.is_some()
        || !layout_changes.is_empty()
    {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "source-backed numeric publication accepts no structural or resource edits",
        ));
    }
    if !changes_are_fixed_numeric(&source, &changes) {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "source-backed numeric publication requires unchanged Number, RK, or MulRk storage",
        ));
    }

    let semantic = SemanticPatch::from_transaction(&source, &changes, &[], &[], &[])?;
    // The complete source Workbook validation already ran when this immutable
    // snapshot was opened, over the same sealed bytes under the same limits,
    // and `SourcePolicyFacts` recorded its three verdicts with their exact
    // refusal texts. Reuse them here as the plan-only publication already
    // does; the target is still validated independently below over its own
    // complete reopen.
    source.inner.source_policy.require()?;

    let publisher = SourceBackedOverlayPublisher::open_owned_vec(
        Arc::clone(&source.inner.bytes),
        source.inner.source_version,
        shared_ole_limits(source.inner.cfb_limits)?,
    )
    .map_err(source_backed_overlay_error)?;
    let source_version = publisher
        .source_version()
        .map_err(source_backed_overlay_error)?;
    if source_version != source.inner.source_version {
        return Err(Error::Internal(
            "source-backed numeric publication lost its source lineage",
        ));
    }
    require_macro_free_container(&publisher)?;

    let mut splices = Vec::new();
    splices
        .try_reserve_exact(changes.len())
        .map_err(|_error| Error::Allocation("staging source-backed numeric splices"))?;
    let mut replacement_bytes = 0_u64;
    for change in &changes {
        if !change_is_effective(&source, change) {
            continue;
        }
        let splice = fixed_numeric_splice(&source, change)?;
        replacement_bytes = replacement_bytes
            .checked_add(u64::try_from(splice.replacement().len()).map_err(|_error| {
                Error::InvalidData("numeric replacement length exceeds u64".into())
            })?)
            .ok_or_else(|| Error::InvalidData("numeric replacement bytes overflow u64".into()))?;
        splices.push(splice);
    }
    let splice_count = splices.len();
    let plan = publisher
        .plan_splices(splices, StreamSpliceLimits::default())
        .map_err(source_backed_overlay_error)?;
    let target_version = if plan.is_noop() {
        source_version
    } else {
        plan.composed_source()
            .map_err(source_backed_overlay_error)?
            .version()
            .map_err(Error::Io)?
    };

    let target = if plan.is_noop() {
        source.clone()
    } else {
        let target_bytes = materialize_numeric_plan(&plan, source.bytes().len())?;
        let package = ObjectPackage::open_with_options(
            target_bytes.clone(),
            OpenOptions::new()
                .targets(Targets::default())
                .limits(source.inner.limits)
                .cfb_limits(source.inner.cfb_limits),
        )?;
        // The publication boundary's complete reopen already parses this
        // target in full. The target checks below used to run on a second,
        // independent `Reader::new` over the identical bytes; they now read
        // that first parse, in the same order and at the same point.
        let target = Snapshot::open_package(
            package,
            Some(NumericTargetVerification {
                expected_bytes: &target_bytes,
                source: &source,
                changes: &changes,
            }),
        )?;
        target.retag_source_version(target_version)
    };

    let source_bytes = u64::try_from(source.bytes().len())
        .map_err(|_error| Error::InvalidData("source CFB length exceeds u64".into()))?;
    let source_workbook_bytes = u64::try_from(source.workbook_stream().len())
        .map_err(|_error| Error::InvalidData("source Workbook length exceeds u64".into()))?;
    let target_workbook_bytes = u64::try_from(target.workbook_stream().len())
        .map_err(|_error| Error::InvalidData("target Workbook length exceeds u64".into()))?;
    let patch = Patch::new(
        Arc::clone(&source.inner.bytes),
        Arc::clone(&target.inner.bytes),
        semantic,
    );
    let diagnostics = SourceBackedDiagnostics {
        changed_cells: changes
            .iter()
            .filter(|change| change_is_effective(&source, change))
            .count(),
        touched_streams: usize::from(!plan.is_noop()),
        splice_count,
        replacement_bytes,
        changed_spans: plan.changed_spans(),
        source_bytes,
        source_workbook_bytes,
        target_workbook_bytes,
        source_version,
        target_version,
        source_fingerprint: plan.source_fingerprint(),
        target_fingerprint: plan.target_fingerprint(),
        operation_shape: plan.operation_shape(),
    };
    Ok(SourceBackedCommit {
        source,
        snapshot: target,
        plan,
        patch,
        diagnostics,
    })
}

fn fixed_numeric_splice(source: &Snapshot, change: &Change) -> Result<SameLengthStreamSplice> {
    let sheet = source
        .inner
        .sheets
        .get(change.sheet)
        .ok_or(Error::PatchConflict {
            operation: "edit cells",
            reason: "source-backed numeric worksheet is stale",
        })?;
    let entry = sheet
        .entries
        .get(change.entry)
        .ok_or(Error::PatchConflict {
            operation: "edit cells",
            reason: "source-backed numeric cell is stale",
        })?;
    if entry.cell.storage != change.storage
        || !matches!(
            change.storage,
            Storage::Number | Storage::Rk | Storage::MulRk
        )
    {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "source-backed numeric edit changed its BIFF storage family",
        ));
    }
    if entry.cell.reference != change.reference {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "source-backed numeric edit changed its BIFF cell reference",
        ));
    }
    let Value::Number(value) = change.value else {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "source-backed numeric edit has a nonnumeric replacement",
        ));
    };
    let width = match change.storage {
        Storage::Number => 8,
        Storage::Rk | Storage::MulRk => 4,
        Storage::BoolErr | Storage::Blank | Storage::LabelSst | Storage::Formula => {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "source-backed numeric edit has an unsupported storage family",
            ));
        },
    };
    let kind = binary::read_u16_le_at(&source.inner.workbook_stream, entry.kind_offset)?;
    if kind != storage_record_kind(entry.cell.storage) {
        return Err(Error::PatchConflict {
            operation: "edit cells",
            reason: "source-backed numeric record family is stale",
        });
    }
    let start = entry.value_offset.ok_or_else(|| {
        Error::refused(
            RefusalKind::UnsupportedTarget,
            "edit cells",
            "numeric cell has no fixed-width value field",
        )
    })?;
    let end = start
        .checked_add(width)
        .ok_or_else(|| Error::InvalidData("numeric splice range overflows usize".into()))?;
    let source_bytes = source
        .inner
        .workbook_stream
        .get(start..end)
        .ok_or_else(|| Error::InvalidData("numeric splice range is outside Workbook".into()))?;
    let Value::Number(source_value) = entry.cell.value else {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "source cell does not contain a fixed-width numeric value",
        ));
    };
    let decoded_source_value = match change.storage {
        Storage::Number => {
            let raw: [u8; 8] = source_bytes.try_into().map_err(|_error| {
                Error::InvalidData("Number source field has an unexpected width".into())
            })?;
            f64::from_le_bytes(raw)
        },
        Storage::Rk | Storage::MulRk => {
            let raw: [u8; 4] = source_bytes.try_into().map_err(|_error| {
                Error::InvalidData("RK source field has an unexpected width".into())
            })?;
            crate::utils::rk_to_f64(u32::from_le_bytes(raw))
        },
        Storage::BoolErr | Storage::Blank | Storage::LabelSst | Storage::Formula => {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "source cell does not contain a fixed-width numeric value",
            ));
        },
    };
    if decoded_source_value.to_bits() != source_value.to_bits() {
        return Err(Error::PatchConflict {
            operation: "edit cells",
            reason: "source-backed numeric field failed its exact semantic precondition",
        });
    }
    let expected = source_bytes.to_vec();
    let replacement = match change.storage {
        Storage::Number => {
            if !valid_xnum(value) {
                return Err(Error::refused(
                    RefusalKind::InvalidValue,
                    "edit cells",
                    "source-backed Number replacement is not a valid Xnum",
                ));
            }
            value.to_le_bytes().to_vec()
        },
        Storage::Rk | Storage::MulRk => encode_rk(value)
            .ok_or_else(|| {
                Error::refused(
                    RefusalKind::InvalidValue,
                    "edit cells",
                    "source-backed RK replacement is not exactly representable",
                )
            })?
            .to_le_bytes()
            .to_vec(),
        Storage::BoolErr | Storage::Blank | Storage::LabelSst | Storage::Formula => {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "source-backed numeric edit has an unsupported storage family",
            ));
        },
    };
    let offset = u64::try_from(start)
        .map_err(|_error| Error::InvalidData("numeric splice offset exceeds u64".into()))?;
    Ok(SameLengthStreamSplice::new(
        source.inner.workbook_path.clone(),
        offset,
        Arc::from(expected),
        Arc::from(replacement),
    ))
}

#[cfg_attr(
    not(test),
    allow(
        dead_code,
        reason = "the unit tests drive the record-level publishers; Package and source_backed are the public routes"
    )
)]
fn materialize_numeric_plan(plan: &ValidatedOverlayPlan, capacity: usize) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(capacity)
        .map_err(|_error| Error::Allocation("materializing source-backed numeric target"))?;
    plan.write_to(&mut bytes)
        .map_err(source_backed_overlay_error)?;
    if bytes.len() != capacity {
        return Err(Error::Internal(
            "source-backed numeric publication changed CFB artifact length",
        ));
    }
    Ok(bytes)
}

fn shared_ole_limits(limits: OleFileLimits) -> Result<SharedOleFileLimits> {
    let max_input_bytes = limits.max_input_bytes();
    let max_directory_bytes = limits.max_directory_bytes();
    let max_allocation_table_bytes = limits.max_allocation_table_bytes();
    SharedOleFileLimits::new(max_input_bytes)
        .and_then(|limits| limits.with_max_directory_bytes(max_directory_bytes))
        .and_then(|limits| limits.with_max_allocation_table_bytes(max_allocation_table_bytes))
        .map_err(Error::Cfb)
}

fn source_backed_overlay_error(error: OverlayError) -> Error {
    match error {
        OverlayError::Ole(error) => Error::Cfb(error),
        OverlayError::Io(error) => Error::Io(error),
        OverlayError::Allocation { resource, .. } => Error::Allocation(resource),
        _ => Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "source-backed numeric overlay refused",
        ),
    }
}

impl From<OverlayError> for Error {
    fn from(error: OverlayError) -> Self {
        source_backed_overlay_error(error)
    }
}

fn workbook_protection_policy<R: Read + Seek>(
    workbook: &ValidationWorkbook<R>,
) -> Result<SourceProtectionPolicy> {
    let protection = workbook.protection();
    if protection.structure_protected()
        || protection.windows_protected()
        || protection.password().is_set()
        || protection.revisions_protected()
        || protection.revision_password().is_set()
        || protection.write_protected()
        || protection.file_sharing().is_some()
    {
        return Ok(SourceProtectionPolicy::WorkbookOrShared);
    }
    for metadata in workbook.sheets() {
        let Some(index) = metadata.parsed_worksheet_index() else {
            continue;
        };
        let protection = workbook.worksheet_protection(index)?;
        if protection.is_protected()
            || protection.objects_protected()
            || protection.scenarios_protected()
            || protection.has_password()
        {
            return Ok(SourceProtectionPolicy::Worksheet);
        }
    }
    Ok(SourceProtectionPolicy::Unprotected)
}

fn require_unprotected_workbook<R: Read + Seek>(workbook: &ValidationWorkbook<R>) -> Result<()> {
    match workbook_protection_policy(workbook)? {
        SourceProtectionPolicy::Unprotected => Ok(()),
        SourceProtectionPolicy::WorkbookOrShared => Err(Error::refused(
            RefusalKind::Protected,
            "edit cells",
            "protected or shared workbooks are not eligible for source-backed numeric edits",
        )),
        SourceProtectionPolicy::Worksheet => Err(Error::refused(
            RefusalKind::Protected,
            "edit cells",
            "protected worksheets are not eligible for source-backed numeric edits",
        )),
    }
}

fn workbook_is_macro_free<R: Read + Seek>(workbook: &ValidationWorkbook<R>) -> bool {
    let metadata = workbook.vba_metadata();
    if metadata.has_project_marker()
        || workbook.vba_project_storage().is_some()
        || workbook
            .sheets()
            .iter()
            .any(|sheet| matches!(sheet.kind(), SheetKind::MacroSheet | SheetKind::VbaModule))
    {
        return false;
    }
    true
}

fn require_macro_free_workbook<R: Read + Seek>(workbook: &ValidationWorkbook<R>) -> Result<()> {
    if workbook_is_macro_free(workbook) {
        Ok(())
    } else {
        Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "macro-bearing XLS sources are not eligible for source-backed numeric edits",
        ))
    }
}

fn require_macro_free_container(publisher: &SourceBackedOverlayPublisher) -> Result<()> {
    if publisher.shared().directory_entries().any(|entry| {
        entry.entry_type == STGTY_STORAGE && entry.name.eq_ignore_ascii_case("_VBA_PROJECT_CUR")
    }) {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "macro-bearing XLS CFB storage is not eligible for source-backed numeric edits",
        ));
    }
    Ok(())
}

fn parse_workbook_stream(
    source: &Arc<[u8]>,
) -> Result<(Vec<SheetData>, Option<usize>, Vec<Vec<u8>>)> {
    let mut records = Records::new(source, litchi_biff::Limits::default())?;
    let first = records.next().ok_or(Error::Eof("Workbook globals BOF"))??;
    require_bof(first.payload(), WORKBOOK_GLOBALS)?;

    let mut encoding = Encoding::from_codepage(1252)?;
    let mut bound_payloads = Vec::new();
    let mut sst_total_offset = None;
    let mut xf_records = Vec::new();
    let mut globals_eof = false;
    for record_result in records.by_ref() {
        let record = record_result?;
        match record.kind().get() {
            CODE_PAGE if record.payload().len() == 2 => {
                // The same Windows-1252 fallback the reader applies (and records
                // in its tolerance report) for a page without an exact codec.
                encoding = Encoding::from_codepage_or_fallback(binary::read_u16_le_at(
                    record.payload(),
                    0,
                )?)
                .0;
            },
            BOUND_SHEET => {
                let mut payload = Vec::new();
                payload
                    .try_reserve_exact(record.payload().len())
                    .map_err(|_error| Error::Allocation("retaining BoundSheet8 payload"))?;
                payload.extend_from_slice(record.payload());
                bound_payloads
                    .try_reserve(1)
                    .map_err(|_error| Error::Allocation("indexing BoundSheet8 records"))?;
                bound_payloads.push(payload);
            },
            SST => {
                if record.payload().len() < 8 || sst_total_offset.is_some() {
                    return Err(Error::InvalidRecord {
                        record_type: SST,
                        message: "SST header is truncated or duplicated".into(),
                    });
                }
                sst_total_offset = Some(
                    record
                        .offset()
                        .checked_add(4)
                        .ok_or_else(|| Error::InvalidData("SST total offset overflow".into()))?,
                );
            },
            XF => {
                xf_records
                    .try_reserve(1)
                    .map_err(|_error| Error::Allocation("retaining XF resources"))?;
                xf_records.push(record.payload().to_vec());
            },
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

    let mut bound_sheets = Vec::new();
    bound_sheets
        .try_reserve_exact(bound_payloads.len())
        .map_err(|_error| Error::Allocation("decoding BoundSheet8 records"))?;
    for payload in &bound_payloads {
        bound_sheets.push(BoundSheetRecord::parse(payload, &encoding)?);
    }
    let mut positions = std::collections::HashSet::new();
    let mut sheets = Vec::new();
    for (workbook_index, bound) in bound_sheets.iter().enumerate() {
        if positions.contains(&bound.position) {
            return Err(Error::InvalidRecord {
                record_type: BOUND_SHEET,
                message: "multiple BoundSheet8 records point to one substream".into(),
            });
        }
        positions
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("validating BoundSheet8 positions"))?;
        positions.insert(bound.position);
        if bound.sheet_type != SheetType::WorkSheet {
            continue;
        }
        let start = usize::try_from(bound.position).map_err(|_error| {
            Error::InvalidData("BoundSheet8 position does not fit usize".into())
        })?;
        let data = source.get(start..).ok_or_else(|| Error::InvalidRecord {
            record_type: BOUND_SHEET,
            message: "BoundSheet8 points outside the Workbook stream".into(),
        })?;
        let entries = parse_worksheet(data, start)?;
        let cell_index = CellIndex::from_entries(&entries)?;
        sheets
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("indexing editable worksheets"))?;
        sheets.push(SheetData {
            name: bound.name.clone(),
            workbook_index,
            entries: Arc::new(entries),
            cell_index: Arc::new(cell_index),
        });
    }
    if xf_records.is_empty() {
        return Err(Error::refused(
            RefusalKind::UnsupportedContent,
            "edit cells",
            "opened-workbook transaction requires at least one XF resource",
        ));
    }
    if let Some(cell) = sheets
        .iter()
        .flat_map(|sheet| sheet.entries.iter())
        .find(|entry| usize::from(entry.cell.style.get()) >= xf_records.len())
    {
        return Err(Error::InvalidRecord {
            record_type: storage_record_kind(cell.cell.storage),
            message: format!(
                "cell XF index {} is outside {} workbook resources",
                cell.cell.style.get(),
                xf_records.len()
            ),
        });
    }
    Ok((sheets, sst_total_offset, xf_records))
}

fn parse_worksheet(data: &[u8], base_offset: usize) -> Result<Vec<Entry>> {
    let mut records = Records::new(data, litchi_biff::Limits::default())?;
    let first = records.next().ok_or(Error::Eof("worksheet BOF"))??;
    require_bof(first.payload(), WORKSHEET)?;
    let mut entries = Vec::new();
    let mut found_eof = false;
    let mut pending_string_formula: Option<PendingStringFormula> = None;
    while let Some(record_result) = records.next() {
        let record = record_result?;
        if let Some(pending) = pending_string_formula.take() {
            if record.kind().get() == STRING {
                let mut continues = Vec::new();
                let text = loop {
                    match crate::utils::decode_string_record(record.payload(), &continues)? {
                        crate::utils::StringRecordDecode::Complete(text) => break text,
                        crate::utils::StringRecordDecode::NeedContinue => {
                            let next = records.next().ok_or_else(|| Error::InvalidRecord {
                                record_type: STRING,
                                message: "formula String cache ends before its Continue records"
                                    .into(),
                            })??;
                            if next.kind().get() != CONTINUE {
                                return Err(Error::InvalidRecord {
                                    record_type: next.kind().get(),
                                    message: "formula String cache continuation is not Continue"
                                        .into(),
                                });
                            }
                            continues.push(next.payload().to_vec());
                        },
                    }
                };
                push_entry(
                    &mut entries,
                    Cell {
                        reference: pending.reference,
                        storage: Storage::Formula,
                        style: pending.style,
                        value: Value::FormulaCache(FormulaCache::String(text)),
                    },
                    Some(payload_offset(pending.kind_offset, 6)?),
                    pending.kind_offset,
                    None,
                )?;
                continue;
            }
            if !matches!(record.kind().get(), 0x0221 | 0x0236 | 0x04bc | 0x0091) {
                return Err(Error::InvalidRecord {
                    record_type: record.kind().get(),
                    message: "string-valued Formula is not followed by its String record".into(),
                });
            }
            pending_string_formula = Some(pending);
        }
        if record.kind().get() == EOF {
            found_eof = true;
            break;
        }
        if record.kind().get() == BOF {
            // An embedded chart substream (`[MS-XLS]` 2.1.7.20.6 OBJECTS):
            // its data-cache Number and Label records are the chart's, not
            // cells of this worksheet, as the worksheet parser reads them
            // (records 0293, 0316).
            let mut depth = 1_usize;
            while depth > 0 {
                let nested = records
                    .next()
                    .ok_or(Error::Eof("embedded substream EOF"))??;
                match nested.kind().get() {
                    BOF => depth += 1,
                    EOF => depth -= 1,
                    _ => {},
                }
            }
            continue;
        }
        let kind = record.kind().get();
        let kind_offset = base_offset
            .checked_add(record.offset())
            .ok_or_else(|| Error::InvalidData("cell record offset overflow".into()))?;
        match kind {
            NUMBER => {
                require_payload(record.payload(), NUMBER_PAYLOAD_BYTES)?;
                let reference = parse_reference(record.payload(), kind)?;
                let value = binary::read_f64_le_at(record.payload(), NUMBER_VALUE_OFFSET)?;
                if !valid_xnum(value) {
                    return Err(Error::InvalidRecord {
                        record_type: NUMBER,
                        message: "Number contains an Xnum forbidden by MS-XLS 2.5.342".into(),
                    });
                }
                push_entry(
                    &mut entries,
                    Cell {
                        reference,
                        storage: Storage::Number,
                        style: parse_style(record.payload(), 4)?,
                        value: Value::Number(value),
                    },
                    Some(payload_offset(kind_offset, NUMBER_VALUE_OFFSET)?),
                    kind_offset,
                    None,
                )?;
            },
            RK => {
                require_payload(record.payload(), 10)?;
                let reference = parse_reference(record.payload(), kind)?;
                let value = crate::utils::rk_to_f64(binary::read_u32_le_at(record.payload(), 6)?);
                require_finite_rk(value, kind)?;
                push_entry(
                    &mut entries,
                    Cell {
                        reference,
                        storage: Storage::Rk,
                        style: parse_style(record.payload(), 4)?,
                        value: Value::Number(value),
                    },
                    Some(payload_offset(kind_offset, 6)?),
                    kind_offset,
                    None,
                )?;
            },
            MUL_RK => parse_mul_rk_entries(record.payload(), kind_offset, &mut entries)?,
            MUL_BLANK => parse_mul_blank_entries(record.payload(), kind_offset, &mut entries)?,
            BOOL_ERR => {
                require_payload(record.payload(), 8)?;
                let reference = parse_reference(record.payload(), kind)?;
                let value = match record.payload()[7] {
                    0 if record.payload()[6] <= 1 => Value::Boolean(record.payload()[6] != 0),
                    1 => Value::Error(CellError::new(record.payload()[6])?),
                    _ => {
                        return Err(Error::InvalidRecord {
                            record_type: BOOL_ERR,
                            message: "BoolErr has an invalid Boolean value or error flag".into(),
                        });
                    },
                };
                push_entry(
                    &mut entries,
                    Cell {
                        reference,
                        storage: Storage::BoolErr,
                        style: parse_style(record.payload(), 4)?,
                        value,
                    },
                    Some(payload_offset(kind_offset, 6)?),
                    kind_offset,
                    None,
                )?;
            },
            BLANK => {
                require_payload(record.payload(), 6)?;
                push_entry(
                    &mut entries,
                    Cell {
                        reference: parse_reference(record.payload(), kind)?,
                        storage: Storage::Blank,
                        style: parse_style(record.payload(), 4)?,
                        value: Value::Blank,
                    },
                    None,
                    kind_offset,
                    None,
                )?;
            },
            LABEL_SST => {
                require_payload(record.payload(), 10)?;
                let index = binary::read_u32_le_at(record.payload(), 6)?;
                push_entry(
                    &mut entries,
                    Cell {
                        reference: parse_reference(record.payload(), kind)?,
                        storage: Storage::LabelSst,
                        style: parse_style(record.payload(), 4)?,
                        value: Value::Text(String::new()),
                    },
                    Some(payload_offset(kind_offset, 6)?),
                    kind_offset,
                    Some(index),
                )?;
            },
            FORMULA => {
                pending_string_formula =
                    parse_formula_entry(record.payload(), kind_offset, &mut entries)?;
            },
            _ => {},
        }
    }
    if !found_eof {
        return Err(Error::Eof("worksheet EOF"));
    }
    Ok(entries)
}

struct PendingStringFormula {
    reference: Reference,
    style: StyleIndex,
    kind_offset: usize,
}

fn parse_formula_entry(
    payload: &[u8],
    kind_offset: usize,
    entries: &mut Vec<Entry>,
) -> Result<Option<PendingStringFormula>> {
    if payload.len() < 22 {
        return Err(Error::InvalidLength {
            expected: 22,
            found: payload.len(),
        });
    }
    let (formula_tokens, formula_extra) = crate::formula_metadata::formula_payload_parts(payload)?;
    if !formula_extra.is_empty() {
        crate::formula_metadata::validate_formula_extra(formula_tokens, formula_extra)?;
    }
    let cache = match crate::utils::parse_formula_value(&payload[6..14])? {
        crate::records::FormulaValue::Number(value) => {
            if !valid_xnum(value) {
                return Err(Error::InvalidRecord {
                    record_type: FORMULA,
                    message: "Formula contains a numeric cache forbidden by Xnum".into(),
                });
            }
            FormulaCache::Number(value)
        },
        crate::records::FormulaValue::Bool(value) => FormulaCache::Boolean(value),
        crate::records::FormulaValue::Error(code) => FormulaCache::Error(CellError::new(code)?),
        crate::records::FormulaValue::Empty => FormulaCache::Empty,
        crate::records::FormulaValue::StringPending | crate::records::FormulaValue::String(_) => {
            return Ok(Some(PendingStringFormula {
                reference: parse_reference(payload, FORMULA)?,
                style: parse_style(payload, 4)?,
                kind_offset,
            }));
        },
    };
    push_entry(
        entries,
        Cell {
            reference: parse_reference(payload, FORMULA)?,
            storage: Storage::Formula,
            style: parse_style(payload, 4)?,
            value: Value::FormulaCache(cache),
        },
        Some(payload_offset(kind_offset, 6)?),
        kind_offset,
        None,
    )?;
    Ok(None)
}

fn parse_mul_rk_entries(
    payload: &[u8],
    kind_offset: usize,
    entries: &mut Vec<Entry>,
) -> Result<()> {
    let Some(items_bytes) = payload.len().checked_sub(6) else {
        return Err(Error::InvalidLength {
            expected: 18,
            found: payload.len(),
        });
    };
    if items_bytes % 6 != 0 {
        return Err(Error::InvalidRecord {
            record_type: MUL_RK,
            message: "MulRk payload does not contain whole cells".into(),
        });
    }
    let count = items_bytes / 6;
    if !(2..=256).contains(&count) {
        return Err(Error::InvalidRecord {
            record_type: MUL_RK,
            message: format!("MulRk contains {count} cells; expected 2 through 256"),
        });
    }
    let row = binary::read_u16_le_at(payload, 0)?;
    let first_column = binary::read_u16_le_at(payload, 2)?;
    let last_column = binary::read_u16_le_at(payload, payload.len() - 2)?;
    let count_minus_one = u16::try_from(count - 1)
        .map_err(|_error| Error::InvalidData("MulRk cell count exceeds u16".into()))?;
    if first_column.checked_add(count_minus_one) != Some(last_column) || last_column > 255 {
        return Err(Error::InvalidRecord {
            record_type: MUL_RK,
            message: "MulRk column extent is inconsistent or outside BIFF8".into(),
        });
    }
    for index in 0..count {
        let item_offset = 4_usize
            .checked_add(index.checked_mul(6).ok_or_else(|| {
                Error::InvalidData("MulRk item offset multiplication overflow".into())
            })?)
            .ok_or_else(|| Error::InvalidData("MulRk item offset overflow".into()))?;
        let column_delta = u16::try_from(index)
            .map_err(|_error| Error::InvalidData("MulRk column index exceeds u16".into()))?;
        let encoded_column = first_column
            .checked_add(column_delta)
            .ok_or_else(|| Error::InvalidData("MulRk column overflow".into()))?;
        let column = u8::try_from(encoded_column).map_err(|_error| Error::InvalidRecord {
            record_type: MUL_RK,
            message: "MulRk column is outside the BIFF8 worksheet grid".into(),
        })?;
        let rk_offset = item_offset
            .checked_add(2)
            .ok_or_else(|| Error::InvalidData("MulRk RK offset overflow".into()))?;
        let value = crate::utils::rk_to_f64(binary::read_u32_le_at(payload, rk_offset)?);
        require_finite_rk(value, MUL_RK)?;
        push_entry(
            entries,
            Cell {
                reference: Reference { row, column },
                storage: Storage::MulRk,
                style: parse_style(payload, item_offset)?,
                value: Value::Number(value),
            },
            Some(payload_offset(kind_offset, rk_offset)?),
            kind_offset,
            None,
        )?;
    }
    Ok(())
}

fn parse_mul_blank_entries(
    payload: &[u8],
    kind_offset: usize,
    entries: &mut Vec<Entry>,
) -> Result<()> {
    let Some(items_bytes) = payload.len().checked_sub(6) else {
        return Err(Error::InvalidLength {
            expected: 10,
            found: payload.len(),
        });
    };
    if items_bytes % 2 != 0 {
        return Err(Error::InvalidRecord {
            record_type: MUL_BLANK,
            message: "MulBlank payload does not contain whole cells".into(),
        });
    }
    let count = items_bytes / 2;
    if !(2..=256).contains(&count) {
        return Err(Error::InvalidRecord {
            record_type: MUL_BLANK,
            message: format!("MulBlank contains {count} cells; expected 2 through 256"),
        });
    }
    let row = binary::read_u16_le_at(payload, 0)?;
    let first_column = binary::read_u16_le_at(payload, 2)?;
    let last_column = binary::read_u16_le_at(payload, payload.len() - 2)?;
    let count_minus_one = u16::try_from(count - 1)
        .map_err(|_error| Error::InvalidData("MulBlank cell count exceeds u16".into()))?;
    if first_column.checked_add(count_minus_one) != Some(last_column) || last_column > 255 {
        return Err(Error::InvalidRecord {
            record_type: MUL_BLANK,
            message: "MulBlank column extent is inconsistent or outside BIFF8".into(),
        });
    }
    for index in 0..count {
        let item_offset = 4_usize
            .checked_add(index.checked_mul(2).ok_or_else(|| {
                Error::InvalidData("MulBlank item offset multiplication overflow".into())
            })?)
            .ok_or_else(|| Error::InvalidData("MulBlank item offset overflow".into()))?;
        let column_delta = u16::try_from(index)
            .map_err(|_error| Error::InvalidData("MulBlank column index exceeds u16".into()))?;
        let encoded_column = first_column
            .checked_add(column_delta)
            .ok_or_else(|| Error::InvalidData("MulBlank column overflow".into()))?;
        let column = u8::try_from(encoded_column).map_err(|_error| Error::InvalidRecord {
            record_type: MUL_BLANK,
            message: "MulBlank column is outside the BIFF8 worksheet grid".into(),
        })?;
        push_entry(
            entries,
            Cell {
                reference: Reference { row, column },
                storage: Storage::Blank,
                style: parse_style(payload, item_offset)?,
                value: Value::Blank,
            },
            None,
            kind_offset,
            None,
        )?;
    }
    Ok(())
}

fn parse_reference(payload: &[u8], record_type: u16) -> Result<Reference> {
    if payload.len() < 4 {
        return Err(Error::InvalidLength {
            expected: 4,
            found: payload.len(),
        });
    }
    let row = binary::read_u16_le_at(payload, 0)?;
    let encoded_column = binary::read_u16_le_at(payload, 2)?;
    let column = u8::try_from(encoded_column).map_err(|_error| Error::InvalidRecord {
        record_type,
        message: "cell column is outside the BIFF8 worksheet grid".into(),
    })?;
    Ok(Reference { row, column })
}

fn parse_style(payload: &[u8], offset: usize) -> Result<StyleIndex> {
    Ok(StyleIndex(binary::read_u16_le_at(payload, offset)?))
}

fn payload_offset(kind_offset: usize, relative: usize) -> Result<usize> {
    kind_offset
        .checked_add(4)
        .and_then(|offset| offset.checked_add(relative))
        .ok_or_else(|| Error::InvalidData("cell payload offset overflow".into()))
}

fn push_entry(
    entries: &mut Vec<Entry>,
    cell: Cell,
    value_offset: Option<usize>,
    kind_offset: usize,
    sst_index: Option<u32>,
) -> Result<()> {
    entries
        .try_reserve(1)
        .map_err(|_error| Error::Allocation("indexing editable cell records"))?;
    entries.push(Entry {
        cell,
        value_offset,
        kind_offset,
        sst_index,
    });
    Ok(())
}

fn require_payload(payload: &[u8], expected: usize) -> Result<()> {
    if payload.len() != expected {
        return Err(Error::InvalidLength {
            expected,
            found: payload.len(),
        });
    }
    Ok(())
}

fn require_finite_rk(value: f64, record_type: u16) -> Result<()> {
    if value.is_finite() {
        Ok(())
    } else {
        Err(Error::InvalidRecord {
            record_type,
            message: "RK numeric value is not finite".into(),
        })
    }
}

fn valid_xnum(value: f64) -> bool {
    value.is_normal() || (value == 0.0 && value.is_sign_positive())
}

fn target_storage(
    source: Storage,
    value: &Value,
    text_present: impl Fn(&str) -> bool,
) -> Result<Storage> {
    let target = match (source, value) {
        (Storage::Number, Value::Number(number)) if valid_xnum(*number) => Storage::Number,
        (Storage::Rk | Storage::MulRk, Value::Number(number)) if encode_rk(*number).is_some() => {
            source
        },
        (Storage::LabelSst, Value::Number(number)) if encode_rk(*number).is_some() => Storage::Rk,
        (Storage::BoolErr, Value::Boolean(_) | Value::Error(_)) => Storage::BoolErr,
        (Storage::Blank, Value::Blank) => Storage::Blank,
        (Storage::LabelSst | Storage::Rk, Value::Text(text)) if text_present(text) => {
            Storage::LabelSst
        },
        (Storage::Formula, Value::FormulaCache(cache)) if valid_formula_cache(cache) => {
            Storage::Formula
        },
        _ => {
            return Err(Error::refused(
                RefusalKind::UnsupportedTarget,
                "edit cells",
                "value is not representable in the fixed-width slot",
            ));
        },
    };
    Ok(target)
}

fn storage_for_new_value(value: &Value, text_present: impl Fn(&str) -> bool) -> Result<Storage> {
    match value {
        Value::Number(number) if valid_xnum(*number) => Ok(Storage::Number),
        Value::Boolean(_) | Value::Error(_) => Ok(Storage::BoolErr),
        Value::Blank => Ok(Storage::Blank),
        Value::Text(text) if text_present(text) => Ok(Storage::LabelSst),
        Value::Text(_) => Err(Error::Internal(
            "new BIFF8 text is absent from the effective workbook SST",
        )),
        Value::Number(_) => Err(Error::UnsupportedFeature(
            "BIFF8 Number insertion requires a normal IEEE-754 value or positive zero".into(),
        )),
        Value::FormulaCache(_) => Err(Error::UnsupportedFeature(
            "a cached result cannot create a formula without formula tokens".into(),
        )),
    }
}

fn validate_xf_payload(payload: &[u8]) -> Result<()> {
    if payload.len() != 20 {
        return Err(Error::InvalidData(format!(
            "BIFF8 XF payload must be exactly 20 bytes, found {}",
            payload.len()
        )));
    }
    Ok(())
}

fn validate_rich_text(
    transaction: &Edit,
    text: &str,
    formatting_runs: &[crate::records::SharedStringFormatRun],
) -> Result<()> {
    let character_count = u16::try_from(text.encode_utf16().count()).map_err(|_error| {
        Error::refused(
            RefusalKind::InvalidValue,
            "edit cells",
            "rich shared string exceeds u16 characters",
        )
    })?;
    if text.is_empty() || formatting_runs.is_empty() {
        return Err(Error::InvalidData(
            "rich shared string requires text and formatting runs".into(),
        ));
    }
    u16::try_from(formatting_runs.len()).map_err(|_error| {
        Error::refused(
            RefusalKind::InvalidValue,
            "edit cells",
            "rich shared string has too many formatting runs",
        )
    })?;
    let mut previous = None;
    for run in formatting_runs {
        if run.character_index >= character_count
            || previous.is_some_and(|value| run.character_index <= value)
        {
            return Err(Error::InvalidData(
                "rich shared-string runs must be strictly ordered inside the text".into(),
            ));
        }
        let font_is_effective = transaction
            .source
            .inner
            .xf_records
            .iter()
            .chain(
                transaction
                    .resource_changes
                    .iter()
                    .filter_map(|change| match change {
                        ResourceChange::ExtendedFormat {
                            payload,
                            insert: true,
                            ..
                        } => Some(payload),
                        ResourceChange::SharedString { .. }
                        | ResourceChange::RichSharedString { .. }
                        | ResourceChange::ExtendedFormat { .. }
                        | ResourceChange::FormulaCell { .. } => None,
                    }),
            )
            .any(|payload| binary::read_u16_le_at(payload, 0).ok() == Some(run.font_index));
        if !font_is_effective {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "rich shared-string font is not used by an effective XF",
            ));
        }
        previous = Some(run.character_index);
    }
    Ok(())
}

fn valid_formula_cache(cache: &FormulaCache) -> bool {
    match cache {
        FormulaCache::Number(value) => valid_xnum(*value),
        FormulaCache::Boolean(_) | FormulaCache::Error(_) | FormulaCache::Empty => true,
        FormulaCache::String(text) => {
            let mut units = text.encode_utf16();
            let count = units.clone().count();
            let width = if units.all(|unit| unit <= 0xff) { 1 } else { 2 };
            u16::try_from(count).is_ok()
                && count
                    .checked_mul(width)
                    .and_then(|length| length.checked_add(3))
                    .is_some_and(|length| length <= 8_224)
        },
    }
}

fn encode_rk(value: f64) -> Option<u32> {
    fn upper_word_if_exact(value: f64) -> Option<u32> {
        let bits = value.to_bits();
        if bits.trailing_zeros() >= 34 {
            u32::try_from(bits >> 32).ok()
        } else {
            None
        }
    }
    fn exact_encoding(value: f64, encoded: u32) -> Option<u32> {
        (crate::utils::rk_to_f64(encoded).to_bits() == value.to_bits()).then_some(encoded)
    }

    if !value.is_finite() {
        return None;
    }
    // Keep this ordering and range identical to the BIFF writer's RK
    // encoder.  The integer form stores a signed 30-bit value in bits 2..31;
    // the scaled form stores that same value after multiplying by 100 and
    // sets bit 0.  The upper-word forms remain available for existing RK
    // values whose IEEE-754 low 34 bits are zero (including noncanonical
    // producer encodings), but writer-compatible integer forms are preferred
    // for replacements.
    let int_value = crate::utils::saturating_f64_to_i32(value);
    if f64::from(int_value) == value
        && (-(1_i32 << 29)..(1_i32 << 29)).contains(&int_value)
        && !(value == 0.0 && value.is_sign_negative())
    {
        let encoded = (crate::utils::reinterpret_i32_as_u32(int_value) << 2) | 0x02;
        if let Some(encoded) = exact_encoding(value, encoded) {
            return Some(encoded);
        }
    }

    let scaled = value * 100.0;
    if scaled.is_finite() {
        // Binary floating-point multiplication can land one ulp below or
        // above an intended decimal integer (for example, `0.29 * 100.0`
        // may be `28.999999999999996`).  Try the nearby rounded signed
        // 30-bit candidate, but require a full RK decode to reproduce the
        // requested f64 bit pattern before accepting it.
        let rounded = scaled.round();
        let scaled_int = crate::utils::saturating_f64_to_i32(rounded);
        if f64::from(scaled_int) == rounded
            && (-(1_i32 << 29)..(1_i32 << 29)).contains(&scaled_int)
            && !(rounded == 0.0 && rounded.is_sign_negative())
        {
            let encoded = (crate::utils::reinterpret_i32_as_u32(scaled_int) << 2) | 0x03;
            if let Some(encoded) = exact_encoding(value, encoded) {
                return Some(encoded);
            }
        }
    }

    upper_word_if_exact(value)
        .and_then(|encoded| exact_encoding(value, encoded))
        .or_else(|| {
            scaled
                .is_finite()
                .then(|| {
                    upper_word_if_exact(scaled)
                        .and_then(|encoded| exact_encoding(value, encoded | 1))
                })
                .flatten()
        })
}

fn change_is_effective(snapshot: &Snapshot, change: &Change) -> bool {
    let source = &snapshot.inner.sheets[change.sheet].entries[change.entry].cell;
    source.storage != change.storage || !values_equal(&source.value, &change.value)
}

#[derive(Debug, Clone, Copy)]
struct FixedNumericField {
    start: usize,
    end: usize,
    change: usize,
}

fn changes_are_fixed_numeric(snapshot: &Snapshot, changes: &[Change]) -> bool {
    changes
        .iter()
        .filter(|change| change_is_effective(snapshot, change))
        .all(|change| {
            snapshot
                .inner
                .sheets
                .get(change.sheet)
                .and_then(|sheet| sheet.entries.get(change.entry))
                .is_some_and(|entry| {
                    entry.cell.storage == change.storage
                        && matches!(
                            change.storage,
                            Storage::Number | Storage::Rk | Storage::MulRk
                        )
                        && matches!(change.value, Value::Number(_))
                })
        })
}

fn carry_fixed_numeric_inventory(
    snapshot: &Snapshot,
    workbook: &[u8],
    changes: &[Change],
) -> Result<Vec<SheetData>> {
    if snapshot.inner.workbook_stream.len() != workbook.len() {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "fixed-width numeric edit changed the Workbook stream length",
        ));
    }

    let mut fields = Vec::new();
    fields
        .try_reserve_exact(changes.len())
        .map_err(|_error| Error::Allocation("certifying fixed-width numeric fields"))?;
    for (change_index, change) in changes.iter().enumerate() {
        if !change_is_effective(snapshot, change) {
            continue;
        }
        let entry = snapshot
            .inner
            .sheets
            .get(change.sheet)
            .and_then(|sheet| sheet.entries.get(change.entry))
            .ok_or(Error::PatchConflict {
                operation: "edit cells",
                reason: "fixed-width numeric dependency is stale",
            })?;
        if entry.cell.storage != change.storage {
            return Err(Error::refused(
                RefusalKind::UnsupportedContent,
                "edit cells",
                "fixed-width numeric edit changed its BIFF record family",
            ));
        }
        let width = match (&change.storage, &change.value) {
            (Storage::Number, Value::Number(value)) => {
                require_numeric_field(workbook, entry, &value.to_le_bytes())?;
                8
            },
            (Storage::Rk | Storage::MulRk, Value::Number(value)) => {
                let encoded = encode_rk(*value).ok_or_else(|| {
                    Error::refused(
                        RefusalKind::UnsupportedTarget,
                        "edit cells",
                        "numeric replacement is not exactly representable as RK",
                    )
                })?;
                require_numeric_field(workbook, entry, &encoded.to_le_bytes())?;
                4
            },
            _ => {
                return Err(Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit cells",
                    "inventory reuse requires unchanged Number, RK, or MulRk storage",
                ));
            },
        };
        let start = entry.value_offset.ok_or_else(|| {
            Error::refused(
                RefusalKind::UnsupportedTarget,
                "edit cells",
                "numeric cell has no value field",
            )
        })?;
        let end = start
            .checked_add(width)
            .ok_or_else(|| Error::InvalidData("numeric value range overflow".into()))?;
        fields.push(FixedNumericField {
            start,
            end,
            change: change_index,
        });
    }
    fields.sort_unstable_by_key(|field| field.start);

    let source = snapshot.inner.workbook_stream.as_ref();
    let mut cursor = 0;
    for field in &fields {
        if field.start < cursor {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "fixed-width numeric edit fields overlap",
            ));
        }
        if source.get(cursor..field.start) != workbook.get(cursor..field.start) {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "fixed-width numeric edit changed bytes outside its value fields",
            ));
        }
        cursor = field.end;
    }
    if source.get(cursor..) != workbook.get(cursor..) {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "fixed-width numeric edit changed bytes outside its value fields",
        ));
    }

    let mut sheets = snapshot.inner.sheets.clone();
    for field in fields {
        let change = &changes[field.change];
        let entries = Arc::make_mut(&mut sheets[change.sheet].entries);
        let cell = &mut entries[change.entry].cell;
        cell.storage = change.storage;
        cell.value = change.value.clone();
    }
    Ok(sheets)
}

fn require_numeric_field(workbook: &[u8], entry: &Entry, expected: &[u8]) -> Result<()> {
    let start = entry.value_offset.ok_or_else(|| {
        Error::refused(
            RefusalKind::UnsupportedTarget,
            "edit cells",
            "numeric cell has no value field",
        )
    })?;
    let end = start
        .checked_add(expected.len())
        .ok_or_else(|| Error::InvalidData("numeric value range overflow".into()))?;
    if workbook.get(start..end) != Some(expected) {
        return Err(Error::Internal(
            "fixed-width numeric field failed exact encoded readback",
        ));
    }
    Ok(())
}

fn require_public_worksheet_coverage<R: Read + Seek>(
    workbook: &ValidationWorkbook<R>,
    sheets: &[SheetData],
) -> Result<()> {
    for sheet in sheets {
        if workbook
            .sheet(sheet.workbook_index)
            .and_then(crate::SheetMetadata::parsed_worksheet_index)
            .is_none()
        {
            return Err(Error::refused(
                RefusalKind::UnsupportedContent,
                "edit cells",
                "a worksheet was not published by the complete XLS reader",
            ));
        }
    }
    Ok(())
}

/// The `(tab position, row, column)` of every staged cell: the cells
/// [`verify_public_numeric_readback`] reads, and so the only cells a
/// validation-only open of the target has to decode.
fn numeric_readback_cells(snapshot: &Snapshot, changes: &[Change]) -> Result<KeptCells> {
    let mut cells = Vec::new();
    cells
        .try_reserve_exact(changes.len())
        .map_err(|_error| Error::Allocation("listing numeric readback cells"))?;
    for change in changes {
        let sheet = snapshot
            .inner
            .sheets
            .get(change.sheet)
            .ok_or(Error::PatchConflict {
                operation: "edit cells",
                reason: "edited worksheet is stale",
            })?;
        cells.push((
            sheet.workbook_index,
            change.reference.row(),
            u16::from(change.reference.column()),
        ));
    }
    KeptCells::from_cells(cells)
}

fn verify_public_numeric_readback<R: Read + Seek>(
    workbook: &ValidationWorkbook<R>,
    snapshot: &Snapshot,
    changes: &[Change],
) -> Result<()> {
    for change in changes {
        if !change_is_effective(snapshot, change) {
            continue;
        }
        let source_sheet = &snapshot.inner.sheets[change.sheet];
        let metadata = workbook
            .sheet(source_sheet.workbook_index)
            .ok_or(Error::Internal(
                "edited worksheet disappeared from public readback",
            ))?;
        let worksheet_index = metadata.parsed_worksheet_index().ok_or(Error::Internal(
            "edited tab is not a public worksheet on readback",
        ))?;
        let cell = workbook
            .kept_cell(
                worksheet_index,
                change.reference.row(),
                u16::from(change.reference.column()),
            )?
            .ok_or(Error::Internal(
                "edited cell is absent from public readback",
            ))?;
        let Value::Number(expected) = change.value else {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "inventory reuse received a nonnumeric value",
            ));
        };
        let actual = match cell.value() {
            CellValue::Float(value) | CellValue::DateTime(value) => *value,
            _ => {
                return Err(Error::refused(
                    RefusalKind::UnsupportedOperation,
                    "edit cells",
                    "edited numeric cell has a nonnumeric public value",
                ));
            },
        };
        if actual.to_bits() != expected.to_bits() {
            return Err(Error::Internal(
                "edited numeric cell failed independent public readback",
            ));
        }
    }
    Ok(())
}

fn structural_changes_equal(left: &StructuralChange, right: &StructuralChange) -> bool {
    match (left, right) {
        (
            StructuralChange::Cell {
                sheet: left_sheet,
                reference: left_reference,
                before: left_before,
                after: left_after,
            },
            StructuralChange::Cell {
                sheet: right_sheet,
                reference: right_reference,
                before: right_before,
                after: right_after,
            },
        ) => {
            left_sheet == right_sheet
                && left_reference == right_reference
                && optional_cell_states_equal(left_before.as_ref(), right_before.as_ref())
                && optional_cell_states_equal(left_after.as_ref(), right_after.as_ref())
        },
        (
            StructuralChange::Rows {
                sheet: left_sheet,
                start: left_start,
                count: left_count,
                insert: left_insert,
            },
            StructuralChange::Rows {
                sheet: right_sheet,
                start: right_start,
                count: right_count,
                insert: right_insert,
            },
        ) => {
            (left_sheet, left_start, left_count, left_insert)
                == (right_sheet, right_start, right_count, right_insert)
        },
        (
            StructuralChange::Columns {
                sheet: left_sheet,
                start: left_start,
                count: left_count,
                insert: left_insert,
            },
            StructuralChange::Columns {
                sheet: right_sheet,
                start: right_start,
                count: right_count,
                insert: right_insert,
            },
        ) => {
            (left_sheet, left_start, left_count, left_insert)
                == (right_sheet, right_start, right_count, right_insert)
        },
        (
            StructuralChange::RenameSheet {
                sheet: left_sheet,
                before: left_before,
                after: left_after,
            },
            StructuralChange::RenameSheet {
                sheet: right_sheet,
                before: right_before,
                after: right_after,
            },
        ) => (left_sheet, left_before, left_after) == (right_sheet, right_before, right_after),
        _ => false,
    }
}

fn resource_changes_equal(left: &ResourceChange, right: &ResourceChange) -> bool {
    match (left, right) {
        (
            ResourceChange::SharedString {
                text: left_text,
                insert: left_insert,
            },
            ResourceChange::SharedString {
                text: right_text,
                insert: right_insert,
            },
        ) => (left_text, left_insert) == (right_text, right_insert),
        (
            ResourceChange::RichSharedString {
                text: left_text,
                formatting_runs: left_runs,
                insert: left_insert,
            },
            ResourceChange::RichSharedString {
                text: right_text,
                formatting_runs: right_runs,
                insert: right_insert,
            },
        ) => (left_text, left_runs, left_insert) == (right_text, right_runs, right_insert),
        (
            ResourceChange::ExtendedFormat {
                index: left_index,
                payload: left_payload,
                insert: left_insert,
            },
            ResourceChange::ExtendedFormat {
                index: right_index,
                payload: right_payload,
                insert: right_insert,
            },
        ) => (left_index, left_payload, left_insert) == (right_index, right_payload, right_insert),
        (
            ResourceChange::FormulaCell {
                sheet: left_sheet,
                reference: left_reference,
                style: left_style,
                tokens: left_tokens,
                insert: left_insert,
            },
            ResourceChange::FormulaCell {
                sheet: right_sheet,
                reference: right_reference,
                style: right_style,
                tokens: right_tokens,
                insert: right_insert,
            },
        ) => {
            (
                left_sheet,
                left_reference,
                left_style,
                left_tokens,
                left_insert,
            ) == (
                right_sheet,
                right_reference,
                right_style,
                right_tokens,
                right_insert,
            )
        },
        _ => false,
    }
}

fn resource_target(change: &ResourceChange) -> String {
    match change {
        ResourceChange::SharedString { text, .. } => {
            format!("resource/sst/{}", text_fingerprint(text))
        },
        ResourceChange::RichSharedString {
            text,
            formatting_runs,
            ..
        } => format!(
            "resource/sst-rich/{}",
            rich_text_fingerprint(text, formatting_runs)
        ),
        ResourceChange::ExtendedFormat { index, .. } => {
            format!("resource/xf/{:05}", index.get())
        },
        ResourceChange::FormulaCell {
            sheet, reference, ..
        } => format!(
            "resource/formula/{sheet}/{}/{}",
            reference.row(),
            reference.column()
        ),
    }
}

fn resource_text(change: &ResourceChange) -> Option<&str> {
    match change {
        ResourceChange::SharedString { text, .. }
        | ResourceChange::RichSharedString { text, .. } => Some(text),
        ResourceChange::ExtendedFormat { .. } => None,
        ResourceChange::FormulaCell { .. } => None,
    }
}

fn formula_resource_cell(change: &ResourceChange) -> Option<(usize, Reference)> {
    match change {
        ResourceChange::FormulaCell {
            sheet, reference, ..
        } => Some((*sheet, *reference)),
        ResourceChange::SharedString { .. }
        | ResourceChange::RichSharedString { .. }
        | ResourceChange::ExtendedFormat { .. } => None,
    }
}

fn optional_cell_states_equal(
    left: Option<&(Storage, Value, StyleIndex)>,
    right: Option<&(Storage, Value, StyleIndex)>,
) -> bool {
    match (left, right) {
        (Some(left), Some(right)) => {
            left.0 == right.0 && values_equal(&left.1, &right.1) && left.2 == right.2
        },
        (None, None) => true,
        _ => false,
    }
}

fn operation_sheet_index(change: &StructuralChange) -> usize {
    match change {
        StructuralChange::Cell { sheet, .. }
        | StructuralChange::Rows { sheet, .. }
        | StructuralChange::Columns { sheet, .. }
        | StructuralChange::RenameSheet { sheet, .. } => *sheet,
    }
}

fn structural_target(snapshot: &Snapshot, change: &StructuralChange) -> String {
    let sheet = snapshot.inner.sheets[operation_sheet_index(change)].workbook_index;
    match change {
        StructuralChange::Cell { reference, .. } => semantic_target(sheet, *reference),
        StructuralChange::Rows { start, .. } => format!("sheet/{sheet}/rows/{start}"),
        StructuralChange::Columns { start, .. } => format!("sheet/{sheet}/columns/{start}"),
        StructuralChange::RenameSheet { .. } => format!("sheet/{sheet}/name"),
    }
}

fn values_equal(left: &Value, right: &Value) -> bool {
    match (left, right) {
        (Value::Number(left), Value::Number(right)) => left.to_bits() == right.to_bits(),
        (Value::FormulaCache(left), Value::FormulaCache(right)) => {
            formula_caches_equal(left, right)
        },
        _ => left == right,
    }
}

fn formula_caches_equal(left: &FormulaCache, right: &FormulaCache) -> bool {
    match (left, right) {
        (FormulaCache::Number(left), FormulaCache::Number(right)) => {
            left.to_bits() == right.to_bits()
        },
        _ => left == right,
    }
}

fn storage_record_kind(storage: Storage) -> u16 {
    match storage {
        Storage::Number => NUMBER,
        Storage::Rk => RK,
        Storage::MulRk => MUL_RK,
        Storage::BoolErr => BOOL_ERR,
        Storage::Blank => BLANK,
        Storage::LabelSst => LABEL_SST,
        Storage::Formula => FORMULA,
    }
}

fn write_cell_value(
    workbook: &mut [u8],
    entry: &Entry,
    change: &Change,
    shared_strings: &[String],
) -> Result<()> {
    if change.storage == Storage::Blank {
        return Ok(());
    }
    let offset = entry.value_offset.ok_or_else(|| {
        Error::refused(
            RefusalKind::UnsupportedTarget,
            "edit cells",
            "cell storage has no editable value field",
        )
    })?;
    match (&change.storage, &change.value) {
        (Storage::Number, Value::Number(value)) => {
            write_field(workbook, offset, &value.to_le_bytes())
        },
        (Storage::Rk | Storage::MulRk, Value::Number(value)) => {
            let encoded = encode_rk(*value).ok_or_else(|| {
                Error::refused(
                    RefusalKind::UnsupportedTarget,
                    "edit cells",
                    "numeric replacement is not exactly representable as RK",
                )
            })?;
            write_field(workbook, offset, &encoded.to_le_bytes())
        },
        (Storage::BoolErr, Value::Boolean(value)) => {
            write_field(workbook, offset, &[u8::from(*value), 0])
        },
        (Storage::BoolErr, Value::Error(error)) => {
            write_field(workbook, offset, &[error.code(), 1])
        },
        (Storage::LabelSst, Value::Text(text)) => {
            let index = shared_strings
                .iter()
                .position(|candidate| candidate == text)
                .ok_or(Error::Internal("replacement text is absent from the SST"))?;
            let index = u32::try_from(index)
                .map_err(|_error| Error::InvalidData("SST index exceeds u32".into()))?;
            write_field(workbook, offset, &index.to_le_bytes())
        },
        (Storage::Formula, Value::FormulaCache(cache)) => {
            write_field(workbook, offset, &encode_formula_cache(cache))
        },
        _ => Err(Error::Internal("staged value and target storage disagree")),
    }
}

fn encode_formula_cache(cache: &FormulaCache) -> [u8; 8] {
    match cache {
        FormulaCache::Number(value) => value.to_le_bytes(),
        FormulaCache::Boolean(value) => [1, 0, u8::from(*value), 0, 0, 0, 0xff, 0xff],
        FormulaCache::Error(error) => [2, 0, error.code(), 0, 0, 0, 0xff, 0xff],
        FormulaCache::Empty => [3, 0, 0, 0, 0, 0, 0xff, 0xff],
        FormulaCache::String(_) => [0, 0, 0, 0, 0, 0, 0xff, 0xff],
    }
}

fn write_field(workbook: &mut [u8], offset: usize, bytes: &[u8]) -> Result<()> {
    let end = offset
        .checked_add(bytes.len())
        .ok_or_else(|| Error::InvalidData("cell replacement range overflow".into()))?;
    workbook
        .get_mut(offset..end)
        .ok_or_else(|| {
            Error::InvalidData("cell replacement is outside the Workbook stream".into())
        })?
        .copy_from_slice(bytes);
    Ok(())
}

fn update_sst_total(workbook: &mut [u8], offset: Option<usize>, delta: i64) -> Result<()> {
    let offset = offset.ok_or_else(|| {
        Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit cells",
            "RK/LabelSst conversion requires an existing SST header",
        )
    })?;
    let current = binary::read_u32_le_at(workbook, offset)?;
    let updated = i64::from(current)
        .checked_add(delta)
        .and_then(|value| u32::try_from(value).ok())
        .ok_or_else(|| Error::InvalidData("SST total reference count overflow".into()))?;
    write_field(workbook, offset, &updated.to_le_bytes())
}

/// Retains the shared-string properties the complete open already built.
///
/// The reader keeps exactly this vector behind an `Arc`. When it covers the
/// shared-string table one entry for one, the shared handle is the same value
/// the element-wise clone produced, so the clone is skipped. Any other shape
/// still takes the copying path, which truncates a longer vector and pads a
/// shorter one exactly as before.
fn retained_shared_string_properties<R: Read + Seek>(
    workbook: &ValidationWorkbook<R>,
    len: usize,
) -> Result<Arc<Vec<Option<Box<crate::records::SharedStringProperties>>>>> {
    if u32::try_from(len).is_ok()
        && let Some(properties) = workbook.shared_string_properties_shared()
        && properties.len() == len
    {
        return Ok(properties);
    }
    let mut properties = Vec::new();
    properties
        .try_reserve_exact(len)
        .map_err(|_error| Error::Allocation("retaining shared-string properties"))?;
    for index in 0..len {
        let index = u32::try_from(index)
            .map_err(|_error| Error::InvalidData("SST index exceeds u32".into()))?;
        properties.push(
            workbook
                .shared_string_properties(index)
                .cloned()
                .map(Box::new),
        );
    }
    Ok(Arc::new(properties))
}

fn resolve_shared_strings(sheets: &mut [SheetData], shared_strings: &[String]) -> Result<()> {
    for sheet in sheets {
        for entry in Arc::make_mut(&mut sheet.entries) {
            let Some(index) = entry.sst_index else {
                continue;
            };
            let text = shared_strings
                .get(usize::try_from(index).map_err(|_error| {
                    Error::InvalidData("LabelSst index does not fit usize".into())
                })?)
                .ok_or_else(|| Error::InvalidRecord {
                    record_type: LABEL_SST,
                    message: format!("LabelSst index {index} is outside the SST"),
                })?;
            let mut retained = String::new();
            retained
                .try_reserve_exact(text.len())
                .map_err(|_error| Error::Allocation("retaining LabelSst text"))?;
            retained.push_str(text);
            entry.cell.value = Value::Text(retained);
        }
    }
    Ok(())
}

fn semantic_patch_limits() -> PatchLimits {
    PatchLimits::new(
        BlobLimits::new(0, 0, 0),
        1_048_576,
        MAX_STAGED_CHANGES,
        8,
        65_536,
        4_194_304,
    )
}

fn patch_error(error: litchi_core::patch::PatchError) -> Error {
    Error::InvalidData(format!("durable XLS cell patch: {error}"))
}

fn semantic_target(sheet_position: usize, reference: Reference) -> String {
    format!(
        "sheet/{sheet_position}/cell/{}/{}",
        reference.row(),
        reference.column()
    )
}

fn parse_semantic_target(target: &str) -> Result<(usize, Reference)> {
    let mut parts = target.split('/');
    if parts.next() != Some("sheet") {
        return Err(Error::InvalidData(
            "invalid cell patch target prefix".into(),
        ));
    }
    let sheet = parts
        .next()
        .ok_or_else(|| Error::InvalidData("cell patch target has no sheet position".into()))?
        .parse::<usize>()
        .map_err(|error| {
            Error::InvalidData(format!("invalid cell patch sheet position: {error}"))
        })?;
    if parts.next() != Some("cell") {
        return Err(Error::InvalidData("invalid cell patch cell prefix".into()));
    }
    let row = parts
        .next()
        .ok_or_else(|| Error::InvalidData("cell patch target has no row".into()))?
        .parse::<u32>()
        .map_err(|error| Error::InvalidData(format!("invalid cell patch row: {error}")))?;
    let column = parts
        .next()
        .ok_or_else(|| Error::InvalidData("cell patch target has no column".into()))?
        .parse::<u32>()
        .map_err(|error| Error::InvalidData(format!("invalid cell patch column: {error}")))?;
    if parts.next().is_some() {
        return Err(Error::InvalidData(
            "cell patch target has trailing data".into(),
        ));
    }
    Ok((sheet, Reference::new(row, column)?))
}

fn cell_state(storage: Storage, value: &Value) -> serde_json::Value {
    serde_json::json!({
        "storage": storage_name(storage),
        "value": encode_value(value),
    })
}

fn storage_name(storage: Storage) -> &'static str {
    match storage {
        Storage::Number => "number",
        Storage::Rk => "rk",
        Storage::MulRk => "mul_rk",
        Storage::BoolErr => "bool_err",
        Storage::Blank => "blank",
        Storage::LabelSst => "label_sst",
        Storage::Formula => "formula",
    }
}

fn parse_storage(value: &serde_json::Value) -> Result<Storage> {
    match value.as_str() {
        Some("number") => Ok(Storage::Number),
        Some("rk") => Ok(Storage::Rk),
        Some("mul_rk") => Ok(Storage::MulRk),
        Some("bool_err") => Ok(Storage::BoolErr),
        Some("blank") => Ok(Storage::Blank),
        Some("label_sst") => Ok(Storage::LabelSst),
        Some("formula") => Ok(Storage::Formula),
        _ => Err(Error::InvalidData("invalid cell patch storage".into())),
    }
}

fn encode_value(value: &Value) -> serde_json::Value {
    match value {
        Value::Number(number) => serde_json::json!({
            "kind": "number",
            "bits": format!("{:016x}", number.to_bits()),
        }),
        Value::Boolean(value) => serde_json::json!({"kind": "boolean", "data": value}),
        Value::Error(error) => serde_json::json!({"kind": "error", "code": error.code()}),
        Value::Blank => serde_json::json!({"kind": "blank"}),
        Value::Text(text) => serde_json::json!({"kind": "text", "data": text}),
        Value::FormulaCache(cache) => serde_json::json!({
            "kind": "formula_cache",
            "data": encode_formula_cache_value(cache),
        }),
    }
}

fn encode_formula_cache_value(cache: &FormulaCache) -> serde_json::Value {
    match cache {
        FormulaCache::Number(number) => serde_json::json!({
            "kind": "number",
            "bits": format!("{:016x}", number.to_bits()),
        }),
        FormulaCache::Boolean(value) => serde_json::json!({"kind": "boolean", "data": value}),
        FormulaCache::Error(error) => {
            serde_json::json!({"kind": "error", "code": error.code()})
        },
        FormulaCache::Empty => serde_json::json!({"kind": "empty"}),
        FormulaCache::String(text) => serde_json::json!({"kind": "string", "data": text}),
    }
}

fn parse_cell_state(state: &serde_json::Value) -> Result<(Storage, Value)> {
    let object = state
        .as_object()
        .ok_or_else(|| Error::InvalidData("cell patch state is not an object".into()))?;
    let storage = parse_storage(
        object
            .get("storage")
            .ok_or_else(|| Error::InvalidData("cell patch state has no storage".into()))?,
    )?;
    let value = parse_value(
        object
            .get("value")
            .ok_or_else(|| Error::InvalidData("cell patch state has no value".into()))?,
    )?;
    Ok((storage, value))
}

fn parse_value(value: &serde_json::Value) -> Result<Value> {
    let object = value
        .as_object()
        .ok_or_else(|| Error::InvalidData("cell patch value is not an object".into()))?;
    match object.get("kind").and_then(serde_json::Value::as_str) {
        Some("number") => parse_number_bits(object).map(Value::Number),
        Some("boolean") => object
            .get("data")
            .and_then(serde_json::Value::as_bool)
            .map(Value::Boolean)
            .ok_or_else(|| Error::InvalidData("cell patch Boolean is malformed".into())),
        Some("error") => parse_error_code(object).map(Value::Error),
        Some("blank") => Ok(Value::Blank),
        Some("text") => object
            .get("data")
            .and_then(serde_json::Value::as_str)
            .map(|text| Value::Text(text.to_string()))
            .ok_or_else(|| Error::InvalidData("cell patch text is malformed".into())),
        Some("formula_cache") => object
            .get("data")
            .ok_or_else(|| Error::InvalidData("formula cache has no data".into()))
            .and_then(parse_formula_cache_value)
            .map(Value::FormulaCache),
        _ => Err(Error::InvalidData(
            "cell patch value kind is invalid".into(),
        )),
    }
}

fn parse_formula_cache_value(value: &serde_json::Value) -> Result<FormulaCache> {
    let object = value
        .as_object()
        .ok_or_else(|| Error::InvalidData("formula cache is not an object".into()))?;
    match object.get("kind").and_then(serde_json::Value::as_str) {
        Some("number") => parse_number_bits(object).map(FormulaCache::Number),
        Some("boolean") => object
            .get("data")
            .and_then(serde_json::Value::as_bool)
            .map(FormulaCache::Boolean)
            .ok_or_else(|| Error::InvalidData("formula Boolean cache is malformed".into())),
        Some("error") => parse_error_code(object).map(FormulaCache::Error),
        Some("empty") => Ok(FormulaCache::Empty),
        Some("string") => object
            .get("data")
            .and_then(serde_json::Value::as_str)
            .map(|text| FormulaCache::String(text.to_string()))
            .ok_or_else(|| Error::InvalidData("formula string cache is malformed".into())),
        _ => Err(Error::InvalidData("formula cache kind is invalid".into())),
    }
}

fn parse_number_bits(object: &serde_json::Map<String, serde_json::Value>) -> Result<f64> {
    let bits = object
        .get("bits")
        .and_then(serde_json::Value::as_str)
        .ok_or_else(|| Error::InvalidData("numeric cell patch has no bit string".into()))?;
    if bits.len() != 16
        || bits
            .bytes()
            .any(|byte| !byte.is_ascii_digit() && !(b'a'..=b'f').contains(&byte))
    {
        return Err(Error::InvalidData(
            "numeric cell patch bit string is non-canonical".into(),
        ));
    }
    u64::from_str_radix(bits, 16)
        .map(f64::from_bits)
        .map_err(|error| Error::InvalidData(format!("invalid numeric cell bits: {error}")))
}

fn parse_error_code(object: &serde_json::Map<String, serde_json::Value>) -> Result<CellError> {
    let code = object
        .get("code")
        .and_then(serde_json::Value::as_u64)
        .ok_or_else(|| Error::InvalidData("cell patch error code is malformed".into()))?;
    CellError::new(
        u8::try_from(code)
            .map_err(|_error| Error::InvalidData("cell patch error code exceeds u8".into()))?,
    )
}

fn caseless_eq(left: &str, right: &str) -> bool {
    crate::sheet_name::equivalent(left, right)
}

fn require_bof(payload: &[u8], expected_substream: u16) -> Result<()> {
    if payload.len() < 4 {
        return Err(Error::InvalidLength {
            expected: 4,
            found: payload.len(),
        });
    }
    let version = binary::read_u16_le_at(payload, 0)?;
    let substream = binary::read_u16_le_at(payload, 2)?;
    if version != BIFF8 || substream != expected_substream {
        return Err(Error::InvalidRecord {
            record_type: BOF,
            message: format!(
                "expected BIFF8 substream 0x{expected_substream:04X}, found version 0x{version:04X} and substream 0x{substream:04X}"
            ),
        });
    }
    Ok(())
}

#[cfg(test)]
fn unique_entry(entries: &[Entry], reference: Reference) -> Result<Option<&Entry>> {
    Ok(unique_entry_index(entries, reference)?.map(|index| &entries[index]))
}

#[cfg(test)]
fn unique_entry_index(entries: &[Entry], reference: Reference) -> Result<Option<usize>> {
    let mut found = None;
    for (index, entry) in entries.iter().enumerate() {
        if entry.cell.reference != reference {
            continue;
        }
        if found.replace(index).is_some() {
            return Err(Error::refused(
                RefusalKind::UnsupportedContent,
                "edit cells",
                "a cell has duplicate BIFF8 Number records",
            ));
        }
    }
    Ok(found)
}

fn verify_readback(snapshot: &Snapshot, changes: &[Change]) -> Result<()> {
    for change in changes {
        let sheet = snapshot
            .inner
            .sheets
            .get(change.sheet)
            .ok_or(Error::Internal("edited worksheet disappeared on readback"))?;
        let entry = sheet
            .unique_entry_index(change.reference)?
            .ok_or(Error::Internal("edited cell was not found on readback"))?;
        let cell = &sheet.entries[entry].cell;
        if cell.storage != change.storage || !values_equal(&cell.value, &change.value) {
            return Err(Error::Internal(
                "edited cell value failed semantic readback",
            ));
        }
    }
    Ok(())
}

fn verify_structural_readback(
    snapshot: &Snapshot,
    source: &Snapshot,
    changes: &[StructuralChange],
    resources: &[ResourceChange],
) -> Result<()> {
    for change in changes {
        match change {
            StructuralChange::Cell {
                sheet,
                reference,
                after: None,
                ..
            } if resources
                .iter()
                .any(|resource| formula_resource_cell(resource) == Some((*sheet, *reference))) =>
            {
                // An authored Formula replaced this cell; the resource
                // readback verifies the replacement record.
            },
            StructuralChange::Cell {
                sheet,
                reference,
                after,
                ..
            } => {
                let sheet_data = snapshot
                    .inner
                    .sheets
                    .get(*sheet)
                    .ok_or(Error::Internal("structurally edited worksheet disappeared"))?;
                let actual = sheet_data
                    .unique_entry_index(*reference)?
                    .map(|index| &sheet_data.entries[index])
                    .map(|entry| (entry.cell.storage, &entry.cell.value, entry.cell.style));
                match (after, actual) {
                    (None, None) => {},
                    (
                        Some((storage, value, style)),
                        Some((found_storage, found_value, found_style)),
                    ) if storage == &found_storage
                        && values_equal(value, found_value)
                        && style == &found_style => {},
                    _ => {
                        return Err(Error::Internal("structural cell failed semantic readback"));
                    },
                }
            },
            StructuralChange::Rows {
                sheet,
                start,
                count,
                insert,
            } => verify_shifted_cells(
                snapshot,
                source,
                *sheet,
                Shift::Rows {
                    start: *start,
                    count: *count,
                    insert: *insert,
                },
            )?,
            StructuralChange::Columns {
                sheet,
                start,
                count,
                insert,
            } => verify_shifted_cells(
                snapshot,
                source,
                *sheet,
                Shift::Columns {
                    start: *start,
                    count: *count,
                    insert: *insert,
                },
            )?,
            StructuralChange::RenameSheet { sheet, after, .. } => {
                if snapshot
                    .inner
                    .sheets
                    .get(*sheet)
                    .map(|item| item.name.as_str())
                    != Some(after.as_str())
                {
                    return Err(Error::Internal("worksheet rename failed semantic readback"));
                }
            },
        }
    }
    Ok(())
}

fn verify_resource_readback(snapshot: &Snapshot, changes: &[ResourceChange]) -> Result<()> {
    for change in changes {
        match change {
            ResourceChange::SharedString { text, insert } => {
                let present = snapshot
                    .inner
                    .shared_strings
                    .iter()
                    .any(|candidate| candidate == text);
                if present != *insert {
                    return Err(Error::Internal(
                        "shared-string resource failed semantic readback",
                    ));
                }
            },
            ResourceChange::RichSharedString {
                text,
                formatting_runs,
                insert,
            } => {
                let present =
                    snapshot
                        .inner
                        .shared_strings
                        .iter()
                        .enumerate()
                        .any(|(index, candidate)| {
                            candidate == text
                                && snapshot
                                    .inner
                                    .shared_string_properties
                                    .get(index)
                                    .and_then(Option::as_deref)
                                    .is_some_and(|properties| {
                                        properties.phonetic.is_none()
                                            && properties.formatting_runs.as_slice()
                                                == formatting_runs.as_slice()
                                    })
                        });
                if present != *insert {
                    return Err(Error::Internal(
                        "rich shared-string resource failed semantic readback",
                    ));
                }
            },
            ResourceChange::FormulaCell {
                sheet,
                reference,
                style,
                tokens,
                insert,
            } => {
                let sheet_data = snapshot.inner.sheets.get(*sheet).ok_or(Error::Internal(
                    "Formula readback sheet dependency is stale",
                ))?;
                let current = sheet_data
                    .unique_entry_index(*reference)?
                    .map(|index| &sheet_data.entries[index]);
                let present = match current {
                    Some(entry) => authored_formula_record_matches(
                        snapshot, entry, *reference, *style, tokens,
                    )?,
                    None => false,
                };
                if present != *insert {
                    return Err(Error::Internal("Formula resource failed semantic readback"));
                }
            },
            ResourceChange::ExtendedFormat {
                index,
                payload,
                insert,
            } => {
                let present = snapshot
                    .inner
                    .xf_records
                    .get(usize::from(index.get()))
                    .is_some_and(|candidate| candidate == payload);
                if present != *insert {
                    return Err(Error::Internal("XF resource failed semantic readback"));
                }
            },
        }
    }
    Ok(())
}

#[derive(Clone, Copy)]
enum Shift {
    Rows {
        start: u16,
        count: u16,
        insert: bool,
    },
    Columns {
        start: u8,
        count: u8,
        insert: bool,
    },
}

fn verify_shifted_cells(
    snapshot: &Snapshot,
    source: &Snapshot,
    sheet: usize,
    shift: Shift,
) -> Result<()> {
    let before = &source.inner.sheets[sheet].entries;
    for entry in before.iter() {
        let reference = match shift {
            Shift::Rows {
                start,
                count,
                insert,
            } => {
                if insert && entry.cell.reference.row() >= start {
                    Reference {
                        row: entry.cell.reference.row() + count,
                        column: entry.cell.reference.column(),
                    }
                } else if !insert
                    && entry.cell.reference.row() >= start
                    && entry.cell.reference.row() < start.saturating_add(count)
                {
                    continue;
                } else if !insert && entry.cell.reference.row() >= start.saturating_add(count) {
                    Reference {
                        row: entry.cell.reference.row() - count,
                        column: entry.cell.reference.column(),
                    }
                } else {
                    entry.cell.reference
                }
            },
            Shift::Columns {
                start,
                count,
                insert,
            } => {
                if insert && entry.cell.reference.column() >= start {
                    Reference {
                        row: entry.cell.reference.row(),
                        column: entry.cell.reference.column() + count,
                    }
                } else if !insert
                    && entry.cell.reference.column() >= start
                    && entry.cell.reference.column() < start.saturating_add(count)
                {
                    continue;
                } else if !insert && entry.cell.reference.column() >= start.saturating_add(count) {
                    Reference {
                        row: entry.cell.reference.row(),
                        column: entry.cell.reference.column() - count,
                    }
                } else {
                    entry.cell.reference
                }
            },
        };
        let after_sheet = &snapshot.inner.sheets[sheet];
        let found = after_sheet
            .unique_entry_index(reference)?
            .and_then(|index| after_sheet.entries.get(index))
            .ok_or(Error::Internal("shifted cell disappeared on readback"))?;
        if found.cell.storage != entry.cell.storage
            || found.cell.style != entry.cell.style
            || !values_equal(&found.cell.value, &entry.cell.value)
        {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit cells",
                "shifted cell changed storage, style, or value",
            ));
        }
    }
    Ok(())
}

#[cfg(test)]
impl SemanticPatch {
    /// Unit-test shorthand: replay through the engine.
    pub(crate) fn apply(&self, source: &Snapshot) -> Result<Commit> {
        self.stage(source)?.commit()
    }
}

#[cfg(test)]
impl TransferPlan {
    /// Unit-test shorthand: execute through the engine.
    pub(crate) fn execute(&self, target: &Snapshot) -> Result<Commit> {
        self.stage(target)?.commit()
    }
}

#[cfg(test)]
mod tests;

#[cfg(test)]
mod index_tests;

#[cfg(test)]
mod tab_tests;

#[cfg(test)]
mod formula_tests;

#[cfg(test)]
mod layout_tests;

#[cfg(test)]
mod format_tests;
