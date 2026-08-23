//! Lossless opened-workbook transactions for BIFF8 cell records and sheets.
//!
//! The owner covers `Number` (`[MS-XLS]` 2.4.180), standalone and packed `RK`,
//! `BoolErr`, `Blank`, `LabelSst`, and non-string `Formula` caches. Every
//! fixed-width changes retain their source slots. Structural cell insertion,
//! removal, row/column movement, and sheet rename regenerate the complete
//! affected row-block/`INDEX`/`DBCELL`/dimension closure. Operations that meet
//! an unsupported formula, range, drawing, or packed-cell dependency are
//! refused before bytes are changed. The complete CFB package is reopened
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

mod structural;

use crate::records::{BoundSheetRecord, Encoding, SheetType};
use crate::{Error, Result, SheetKind, Workbook};
use litchi_biff::Records;
use litchi_cfb::consts::STGTY_STORAGE;
use litchi_cfb::{
    ArtifactFingerprint, ComposedOverlaySource, OverlayError, OverlayOperationShape, PublishReport,
    SameLengthStreamSplice, StreamSpliceLimits, ValidatedOverlayPlan,
};
use litchi_core::binary;
pub use litchi_core::patch::HistoryLimits;
use litchi_core::patch::{
    BlobBundle, BlobLimits, DiagnosticFingerprint, Patch as CorePatch, PatchLimits, PatchOperation,
    Reversible, ReversibleOperation,
};
use litchi_core::sheet::{Cell as _, CellValue};
use litchi_core::{ReadAt, SourceVersion};
use litchi_ole_common::object::{Editor as PackageEditor, Limits, Targets};
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

/// Backward-compatible short name for [`CellRange`].
pub type Range = CellRange;

/// A semantic worksheet selector.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Selector<'a> {
    /// Select a worksheet by case-insensitive developer-visible tab name.
    Name(&'a str),
    /// Select a worksheet by checked zero-based workbook tab position.
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
    /// Formatting-only `Blank` record.
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
    /// Returns [`Error::UnsafeEdit`] if the resource does not exist.
    pub fn new(snapshot: &Snapshot, index: u16) -> Result<Self> {
        if usize::from(index) >= snapshot.inner.xf_records.len() {
            return Err(Error::UnsafeEdit(format!(
                "BIFF8 XF index {index} is outside the workbook's {} resources",
                snapshot.inner.xf_records.len()
            )));
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
    fn from_workbook<R: Read + Seek>(workbook: &Workbook<R>, sheets: &[SheetData]) -> Result<Self> {
        require_public_worksheet_coverage(workbook, sheets)?;
        Ok(Self {
            public_worksheet_coverage: true,
            protection: workbook_protection_policy(workbook)?,
            macro_free_workbook: workbook_is_macro_free(workbook),
        })
    }

    fn require(self) -> Result<()> {
        if !self.public_worksheet_coverage {
            return Err(Error::UnsafeEdit(
                "source-backed numeric source lost public worksheet coverage".into(),
            ));
        }
        match self.protection {
            SourceProtectionPolicy::Unprotected => {},
            SourceProtectionPolicy::WorkbookOrShared => {
                return Err(Error::UnsafeEdit(
                    "protected or shared workbooks are not eligible for source-backed numeric edits"
                        .into(),
                ));
            },
            SourceProtectionPolicy::Worksheet => {
                return Err(Error::UnsafeEdit(
                    "protected worksheets are not eligible for source-backed numeric edits".into(),
                ));
            },
        }
        if !self.macro_free_workbook {
            return Err(Error::UnsafeEdit(
                "macro-bearing XLS sources are not eligible for source-backed numeric edits".into(),
            ));
        }
        Ok(())
    }
}

struct Inner {
    bytes: Arc<[u8]>,
    source_version: SourceVersion,
    workbook_path: Vec<String>,
    workbook_stream: Arc<[u8]>,
    shared_strings: Arc<Vec<String>>,
    shared_string_properties: Arc<Vec<Option<Box<crate::records::SharedStringProperties>>>>,
    sst_total_offset: Option<usize>,
    xf_records: Arc<Vec<Vec<u8>>>,
    sheets: Vec<SheetData>,
    source_policy: SourcePolicyFacts,
}

/// Immutable, cheaply cloned snapshot of an editable XLS package.
#[derive(Clone)]
pub struct Snapshot {
    inner: Arc<Inner>,
}

/// Explicit bounded undo/redo history of immutable XLS snapshots.
pub type History = litchi_core::patch::History<Snapshot>;

impl Snapshot {
    /// Opens an unencrypted, unsigned XLS package and captures exact source bytes.
    ///
    /// The ordinary [`Workbook`] reader validates the complete candidate. The
    /// package transaction layer additionally refuses signed, encrypted, and
    /// DRM CFB containers before exposing an edit.
    ///
    /// # Errors
    ///
    /// Returns a typed CFB, BIFF, encryption, allocation, or workbook
    /// validation error before a snapshot is published.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = PackageEditor::open(bytes, Targets::default(), Limits::default())?;
        Self::from_package_editor(package)
    }

    fn from_package_editor(package: PackageEditor) -> Result<Self> {
        let workbook_path = [vec!["Workbook".to_string()], vec!["Book".to_string()]]
            .into_iter()
            .find(|path| package.stream(path).is_some())
            .ok_or_else(|| {
                Error::InvalidData("XLS package has no Workbook or Book stream".into())
            })?;
        let workbook_stream = package
            .stream_shared(&workbook_path)
            .ok_or_else(|| Error::InvalidData("selected XLS Workbook stream disappeared".into()))?;
        let source = package.finish()?;
        let source_version = fresh_snapshot_source_version();

        let (mut sheets, sst_total_offset, xf_records) = parse_workbook_stream(&workbook_stream)?;
        // A full semantic open catches cross-stream and workbook-global
        // dependencies before the narrower source-offset inventory is kept.
        // The legacy reader intentionally skips some malformed optional sheet
        // projections, so this edit owner additionally requires every sheet
        // it can mutate to have survived that complete semantic open.
        let (shared_strings, shared_string_properties, source_policy) = {
            let workbook = Workbook::new(Cursor::new(source.as_slice()))?;
            let source_policy = SourcePolicyFacts::from_workbook(&workbook, &sheets)?;
            let strings = workbook.shared_strings_shared();
            let mut properties = Vec::new();
            properties
                .try_reserve_exact(strings.len())
                .map_err(|_error| Error::Allocation("retaining shared-string properties"))?;
            for index in 0..strings.len() {
                let index = u32::try_from(index)
                    .map_err(|_error| Error::InvalidData("SST index exceeds u32".into()))?;
                properties.push(
                    workbook
                        .shared_string_properties(index)
                        .cloned()
                        .map(Box::new),
                );
            }
            (strings, Arc::new(properties), source_policy)
        };
        resolve_shared_strings(&mut sheets, &shared_strings)?;
        Ok(Self {
            inner: Arc::new(Inner {
                bytes: Arc::from(source),
                source_version,
                workbook_path,
                workbook_stream,
                shared_strings,
                shared_string_properties,
                sst_total_offset,
                xf_records: Arc::new(xf_records),
                sheets,
                source_policy,
            }),
        })
    }

    fn from_fixed_numeric_package_editor(
        package: PackageEditor,
        source_snapshot: &Self,
        changes: &[Change],
    ) -> Result<Self> {
        let workbook_stream = package
            .stream_shared(&source_snapshot.inner.workbook_path)
            .ok_or_else(|| Error::InvalidData("selected XLS Workbook stream disappeared".into()))?;
        let sheets = carry_fixed_numeric_inventory(source_snapshot, &workbook_stream, changes)?;
        let source = package.finish()?;

        // Keep the complete public reader as an independent validation owner.
        // Only the private offset inventory is carried forward after proving
        // that every other Workbook-stream byte is unchanged.
        let workbook = Workbook::new(Cursor::new(source.as_slice()))?;
        let source_policy = SourcePolicyFacts::from_workbook(&workbook, &sheets)?;
        verify_public_numeric_readback(&workbook, source_snapshot, changes)?;

        Ok(Self {
            inner: Arc::new(Inner {
                bytes: Arc::from(source),
                source_version: fresh_snapshot_source_version(),
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
            }),
        })
    }

    fn retag_source_version(self, source_version: SourceVersion) -> Self {
        if self.inner.source_version == source_version {
            return self;
        }
        let inner = &self.inner;
        Self {
            inner: Arc::new(Inner {
                bytes: Arc::clone(&inner.bytes),
                source_version,
                workbook_path: inner.workbook_path.clone(),
                workbook_stream: Arc::clone(&inner.workbook_stream),
                shared_strings: Arc::clone(&inner.shared_strings),
                shared_string_properties: Arc::clone(&inner.shared_string_properties),
                sst_total_offset: inner.sst_total_offset,
                xf_records: Arc::clone(&inner.xf_records),
                sheets: inner.sheets.clone(),
                source_policy: inner.source_policy,
            }),
        }
    }

    /// Returns exact source CFB bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.inner.bytes
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
    pub fn edit(&self) -> Transaction {
        Transaction {
            source: self.clone(),
            changes: Vec::new(),
            structural_changes: Vec::new(),
            resource_changes: Vec::new(),
        }
    }

    /// Starts the unified immutable-workbook transaction.
    #[must_use]
    pub fn transaction(&self) -> Transaction {
        self.edit()
    }

    /// Starts bounded undo/redo history at this immutable snapshot.
    #[must_use]
    pub fn history(&self, limits: HistoryLimits) -> History {
        History::new(self.clone(), limits)
    }

    fn resolve_sheet(&self, selector: Selector<'_>) -> Result<Option<usize>> {
        match selector {
            Selector::Position(position) => Ok(self
                .inner
                .sheets
                .iter()
                .position(|sheet| sheet.workbook_index == position)),
            Selector::Name(name) => {
                let mut matches = self
                    .inner
                    .sheets
                    .iter()
                    .enumerate()
                    .filter(|(_, sheet)| caseless_eq(&sheet.name, name));
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
        unique_entry(&self.data().entries, reference).map(|entry| {
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
        unique_entry(&self.data().entries, reference).map(|entry| entry.map(|item| &item.cell))
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

/// Detached, failure-atomic edits of existing fixed-width BIFF8 cell fields.
#[derive(Clone)]
pub struct Transaction {
    source: Snapshot,
    changes: Vec<Change>,
    structural_changes: Vec<StructuralChange>,
    resource_changes: Vec<ResourceChange>,
}

/// Backward-compatible name for [`Transaction`].
pub type Edit = Transaction;

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
    /// At least one cell has incompatible requested outcomes.
    Conflicts(Box<[CellConflict]>),
    /// Structural operations overlap or request divergent outcomes.
    StructuralConflicts(Box<[OperationConflict]>),
    /// The joined transaction would exceed its finite change bound.
    Limit { observed: usize, limit: usize },
    /// Retaining deterministic conflict or joined state failed.
    Allocation,
}

impl Transaction {
    /// Joins independently prepared changes against the same exact snapshot.
    ///
    /// Disjoint and byte-equivalent cell outcomes compose deterministically.
    /// Conflicts and bounds are checked before this transaction changes.
    ///
    /// # Errors
    ///
    /// Returns the exact-source, structured cell-conflict, allocation, or
    /// finite-bound reason for refusing the incoming transaction.
    pub fn join(&mut self, incoming: Self) -> std::result::Result<&mut Self, JoinError> {
        if self.source.bytes() != incoming.source.bytes() {
            return Err(JoinError::DifferentSource);
        }
        let mut conflicts = Vec::new();
        let mut additions = 0_usize;
        for right in &incoming.changes {
            if !change_is_effective(&incoming.source, right) {
                continue;
            }
            if let Some(left) = self
                .changes
                .iter()
                .filter(|left| change_is_effective(&self.source, left))
                .find(|left| left.sheet == right.sheet && left.entry == right.entry)
            {
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
            } else {
                additions = additions.saturating_add(1);
            }
        }
        if !conflicts.is_empty() {
            return Err(JoinError::Conflicts(conflicts.into_boxed_slice()));
        }
        let mut structural_conflicts = Vec::new();
        for right in &incoming.structural_changes {
            for left in &self.structural_changes {
                if structural_changes_overlap(left, right) && !structural_changes_equal(left, right)
                {
                    structural_conflicts
                        .try_reserve(1)
                        .map_err(|_error| JoinError::Allocation)?;
                    structural_conflicts.push(OperationConflict {
                        target: structural_target(&self.source, right),
                    });
                }
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
                    .structural_changes
                    .iter()
                    .any(|left| structural_changes_equal(left, right))
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
        for change in incoming.changes {
            if !change_is_effective(&incoming.source, &change) {
                continue;
            }
            if self
                .changes
                .iter()
                .filter(|left| change_is_effective(&self.source, left))
                .any(|left| left.sheet == change.sheet && left.entry == change.entry)
            {
                continue;
            }
            self.changes.push(change);
        }
        self.changes.sort_by_key(|change| {
            (
                self.source.inner.sheets[change.sheet].workbook_index,
                change.reference,
            )
        });
        self.structural_changes
            .try_reserve(structural_additions)
            .map_err(|_error| JoinError::Allocation)?;
        for change in incoming.structural_changes {
            if !self
                .structural_changes
                .iter()
                .any(|left| structural_changes_equal(left, &change))
            {
                self.structural_changes.push(change);
            }
        }
        self.structural_changes
            .sort_by_key(|change| structural_target(&self.source, change));
        self.resource_changes
            .try_reserve(resource_additions)
            .map_err(|_error| JoinError::Allocation)?;
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
    /// Returns an error for an invalid Xnum, a missing or ambiguous worksheet,
    /// a duplicate cell, a non-`Number` cell family, or allocation failure.
    pub fn set_number(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        value: f64,
    ) -> Result<()> {
        if !valid_xnum(value) {
            return Err(Error::UnsupportedFeature(
                "BIFF8 Number edits require a normal IEEE-754 value or positive zero".into(),
            ));
        }
        let sheet_index = self
            .source
            .resolve_sheet(selector)?
            .ok_or_else(|| Error::WorksheetNotFound("cell-value edit worksheet selector".into()))?;
        let entries = &self.source.inner.sheets[sheet_index].entries;
        let entry = unique_entry_index(entries, reference)?.ok_or_else(|| {
            Error::UnsupportedFeature(format!(
                "cell ({}, {}) is absent or is not encoded as a BIFF8 Number record",
                reference.row(),
                reference.column()
            ))
        })?;
        if entries[entry].cell.storage != Storage::Number {
            return Err(Error::UnsupportedFeature(format!(
                "cell ({}, {}) is not encoded as a standalone BIFF8 Number record",
                reference.row(),
                reference.column()
            )));
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
    pub fn set_numeric(
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
    pub fn copy_scalar_cells_from<'source, 'target>(
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
            return Err(Error::UnsafeEdit(format!(
                "cross-workbook scalar transfer contains {cell_count} cells; limit is {MAX_SCALAR_TRANSFER_CELLS}"
            )));
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
                return Err(Error::UnsafeEdit(
                    "cross-workbook scalar transfer source inventory exceeds its checked range"
                        .into(),
                ));
            }
            if entry.cell.style != source_style {
                return Err(Error::UnsafeEdit(
                    "cross-workbook scalar transfer refuses source XF remapping".into(),
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
                return Err(Error::UnsafeEdit(format!(
                    "cross-workbook scalar transfer exceeds {MAX_SCALAR_TRANSFER_BYTES} staged bytes"
                )));
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
                return Err(Error::UnsafeEdit(format!(
                    "cell ({}, {}) has duplicate BIFF8 owners",
                    duplicate[0].0.row(),
                    duplicate[0].0.column()
                )));
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
                return Err(Error::UnsafeEdit(
                    "cross-workbook scalar transfer target inventory exceeds its checked range"
                        .into(),
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
                return Err(Error::UnsafeEdit(format!(
                    "cell ({}, {}) has duplicate BIFF8 target owners",
                    duplicate[0].0.row(),
                    duplicate[0].0.column()
                )));
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
                return Err(Error::UnsafeEdit(
                    "cross-workbook scalar transfer target range is occupied by an unrelated cell"
                        .into(),
                ));
            };
            let source_storage = source_cells[source_index].1;
            if !matches!(target_entry.cell.storage, Storage::Number | Storage::Blank)
                || target_entry.cell.storage != source_storage
            {
                return Err(Error::UnsafeEdit(
                    "cross-workbook scalar transfer would change a BIFF cell record width or family"
                        .into(),
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
                    return Err(Error::UnsafeEdit(
                        "cross-workbook scalar transfer source value is not representable in the target slot"
                            .into(),
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
                        Selector::Position(target_position),
                        reference,
                        value,
                        target_style,
                    )?;
                },
                TransferCell::SetNumber { reference, value } => {
                    staged.set_number(Selector::Position(target_position), reference, value)?;
                },
            }
        }
        *self = staged;
        Ok(Some(self))
    }

    /// Alias for [`Self::copy_scalar_cells_from`] using the conventional
    /// worksheet-copy spelling.
    pub fn copy_cells_from<'source, 'target>(
        &mut self,
        donor: &Snapshot,
        source_selector: impl Into<Selector<'source>>,
        source_range: CellRange,
        target_selector: impl Into<Selector<'target>>,
        target_anchor: Reference,
    ) -> Result<Option<&mut Self>> {
        self.copy_scalar_cells_from(
            donor,
            source_selector,
            source_range,
            target_selector,
            target_anchor,
        )
    }

    /// Copies one exact scalar cell from another snapshot.
    pub fn copy_scalar_cell_from<'source, 'target>(
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
            .ok_or_else(|| Error::UnsafeEdit("target worksheet inventory is stale".into()))
    }

    /// Sets an existing cell through a storage-aware, dependency-checked path.
    ///
    /// Numeric `Number`, `RK`, and `MulRk` cells, `BoolErr`, SST
    /// references, and non-string Formula caches are editable in place.
    /// A standalone `RK` and `LabelSst` can be converted into each other
    /// because both have the same ten-byte payload; the SST reference count is
    /// updated atomically. Missing simple text is interned into a bounded SST
    /// tail resource when an absent or canonical adjacent `ExtSST` can be
    /// updated without changing its bucket size.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or duplicate cell, an unrepresentable
    /// storage conversion, an invalid numeric/cache value, or a finite bound.
    pub fn set_value(
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
        let entry_index = unique_entry_index(entries, reference)?.ok_or_else(|| {
            Error::UnsupportedFeature(format!(
                "cell ({}, {}) has no representable source record slot",
                reference.row(),
                reference.column()
            ))
        })?;
        let source_cell = entries[entry_index].cell.clone();
        if source_cell.storage == Storage::Formula
            && (matches!(
                source_cell.value,
                Value::FormulaCache(FormulaCache::String(_))
            ) || matches!(value, Value::FormulaCache(FormulaCache::String(_))))
        {
            let Value::FormulaCache(cache) = &value else {
                return Err(Error::UnsafeEdit(
                    "Formula storage accepts only cached formula results".into(),
                ));
            };
            if !valid_formula_cache(cache) {
                return Err(Error::UnsafeEdit(
                    "formula cache is not representable in BIFF8".into(),
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
        let mut shared_strings = self.effective_shared_strings();
        if let Some(text) = &missing_text {
            shared_strings.push(text.clone());
        }
        let target = target_storage(source_cell.storage, &value, &shared_strings)?;
        if let Some(text) = missing_text {
            self.stage_resource(ResourceChange::SharedString { text, insert: true })?;
        }
        self.stage(sheet_index, entry_index, target, value)
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
    pub fn clear_cell(&mut self, selector: Selector<'_>, reference: Reference) -> Result<()> {
        let sheet_index = self
            .source
            .resolve_sheet(selector)?
            .ok_or_else(|| Error::WorksheetNotFound("cell clear worksheet selector".into()))?;
        let entries = &self.source.inner.sheets[sheet_index].entries;
        let entry_index = unique_entry_index(entries, reference)?.ok_or_else(|| {
            Error::UnsupportedFeature("an absent BIFF8 cell is already clear".into())
        })?;
        if entries[entry_index].cell.storage != Storage::Blank {
            return Err(Error::UnsafeEdit(
                "removing this cell requires a whole row-block rewrite".into(),
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
    pub fn insert_cell(
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
    pub fn insert_rich_text_cell(
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
                return Err(Error::UnsafeEdit(
                    "rich SST reuse requires one exact formatting identity".into(),
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
    pub fn insert_formula(
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
    pub fn insert_formula_with_style(
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
        if unique_entry_index(&self.source.inner.sheets[sheet].entries, reference)?.is_some() {
            return Err(Error::UnsafeEdit(
                "authored Formula target is occupied".into(),
            ));
        }
        let parsed = crate::writer::FormulaTokenizer::new().tokenize(expression)?;
        let tokens = crate::writer::formula::encode_ptg_tokens(&parsed);
        if tokens.is_empty() || tokens.len() > MAX_FORMULA_TOKEN_BYTES {
            return Err(Error::UnsafeEdit(
                "authored Formula token bytes are empty or exceed the BIFF8 record limit".into(),
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

    /// Inserts a previously absent scalar cell with an existing XF resource.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for an occupied location, unsupported value,
    /// stale style resource, dependency closure, or finite transaction bound.
    pub fn insert_cell_with_style(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        value: Value,
        style: StyleIndex,
    ) -> Result<()> {
        self.require_effective_style(style)?;
        let sheet = self.require_sheet(selector, "cell insertion")?;
        if unique_entry_index(&self.source.inner.sheets[sheet].entries, reference)?.is_some() {
            return Err(Error::UnsafeEdit(
                "BIFF8 cell insertion target is occupied".into(),
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
        let shared_strings = self.effective_shared_strings();
        let storage = storage_for_new_value(&value, &shared_strings)?;
        self.stage_structural(StructuralChange::Cell {
            sheet,
            reference,
            before: None,
            after: Some((storage, value, style)),
        })
    }

    /// Physically removes an existing standalone BIFF8 cell record.
    ///
    /// An edge member of `MulRk` can be removed while preserving or reducing
    /// the packed record. Interior packed members and formula ownership groups
    /// are refused because partial deletion would not preserve their semantics.
    ///
    /// # Errors
    ///
    /// Returns a typed absence, ambiguity, dependency, or bound error.
    pub fn remove_cell(&mut self, selector: Selector<'_>, reference: Reference) -> Result<()> {
        let sheet = self.require_sheet(selector, "cell removal")?;
        let entry = unique_entry_index(&self.source.inner.sheets[sheet].entries, reference)?
            .ok_or_else(|| Error::UnsafeEdit("BIFF8 cell removal target is absent".into()))?;
        let cell = &self.source.inner.sheets[sheet].entries[entry].cell;
        if cell.storage == Storage::Formula {
            return Err(Error::UnsafeEdit(
                "formula records require a wider dependency rewrite".into(),
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
    pub fn set_style(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        style: StyleIndex,
    ) -> Result<()> {
        self.require_effective_style(style)?;
        let sheet = self.require_sheet(selector, "cell style edit")?;
        let entry = unique_entry_index(&self.source.inner.sheets[sheet].entries, reference)?
            .ok_or_else(|| Error::UnsafeEdit("BIFF8 style target is absent".into()))?;
        let cell = &self.source.inner.sheets[sheet].entries[entry].cell;
        self.stage_structural(StructuralChange::Cell {
            sheet,
            reference,
            before: Some((cell.storage, cell.value.clone(), cell.style)),
            after: Some((cell.storage, cell.value.clone(), style)),
        })
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
    pub fn duplicate_style(&mut self, source: StyleIndex) -> Result<StyleIndex> {
        self.require_effective_style(source)?;
        let payload = self
            .effective_xf_payload(source)
            .ok_or_else(|| Error::UnsafeEdit("source XF resource is absent".into()))?
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
    pub fn author_style(&mut self, payload: &[u8]) -> Result<StyleIndex> {
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
            .ok_or_else(|| Error::UnsafeEdit("new XF index exceeds u16".into()))?;
        self.stage_resource(ResourceChange::ExtendedFormat {
            index,
            payload: retained,
            insert: true,
        })?;
        Ok(index)
    }

    fn has_shared_string(&self, text: &str) -> bool {
        self.effective_shared_strings()
            .iter()
            .any(|candidate| candidate == text)
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
            .ok_or_else(|| Error::UnsafeEdit(format!("XF index {} is not staged", style.get())))
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
                    return Err(Error::UnsafeEdit(
                        "shared-string resource outcome is already present".into(),
                    ));
                }
                if *insert {
                    structural::certify_sst_insertion(
                        &self.source,
                        &self.resource_changes,
                        &change,
                    )?;
                }
            },
            ResourceChange::RichSharedString {
                text,
                formatting_runs,
                insert,
            } => {
                let present = self.has_rich_shared_string(text, formatting_runs);
                if present == *insert {
                    return Err(Error::UnsafeEdit(
                        "rich shared-string resource outcome is already present".into(),
                    ));
                }
                if *insert {
                    if self.has_shared_string(text) {
                        return Err(Error::UnsafeEdit(
                            "rich SST text is already present in the effective SST".into(),
                        ));
                    }
                    validate_rich_text(self, text, formatting_runs)?;
                    structural::certify_sst_insertion(
                        &self.source,
                        &self.resource_changes,
                        &change,
                    )?;
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
                        return Err(Error::UnsafeEdit(
                            "new XF index is not the next effective resource".into(),
                        ));
                    }
                    validate_xf_payload(payload)?;
                    if !self.has_xf_dependency(payload) {
                        return Err(Error::UnsafeEdit(
                            "new XF resource dependencies are not effective".into(),
                        ));
                    }
                } else if usize::from(index.get()).checked_add(1) != Some(count)
                    || self.effective_xf_payload(*index) != Some(payload.as_slice())
                {
                    return Err(Error::UnsafeEdit(
                        "XF removal is not the exact effective tail resource".into(),
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
                if self.structural_changes.iter().any(|change| {
                    operation_sheet_index(change) == *sheet
                        && !matches!(change, StructuralChange::RenameSheet { .. })
                }) {
                    return Err(Error::UnsafeEdit(
                        "authored Formula cannot overlap another structural worksheet edit".into(),
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
                    return Err(Error::UnsafeEdit(
                        "authored Formula target already has a staged owner".into(),
                    ));
                }
                let sheet_data = self.source.inner.sheets.get(*sheet).ok_or_else(|| {
                    Error::UnsafeEdit("authored Formula sheet dependency is stale".into())
                })?;
                let current = unique_entry(&sheet_data.entries, *reference)?;
                if *insert && current.is_some() {
                    return Err(Error::UnsafeEdit(
                        "authored Formula insertion target is occupied".into(),
                    ));
                }
                if !*insert {
                    let entry = current.ok_or_else(|| {
                        Error::UnsafeEdit("authored Formula inverse target is absent".into())
                    })?;
                    if !authored_formula_record_matches(
                        &self.source,
                        entry,
                        *reference,
                        *style,
                        tokens,
                    )? {
                        return Err(Error::UnsafeEdit(
                            "authored Formula inverse record dependency is stale".into(),
                        ));
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
            return Err(Error::UnsafeEdit(format!(
                "cell transaction exceeds its {MAX_STAGED_CHANGES}-change limit"
            )));
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
    pub fn insert_rows(&mut self, selector: Selector<'_>, start: u16, count: u16) -> Result<()> {
        self.stage_rows(selector, start, count, true)
    }

    /// Deletes row coordinates and their cells, moving later rows upward.
    ///
    /// # Errors
    ///
    /// Returns a typed dependency or reversibility refusal. Deletion is
    /// durable only when the deleted coordinate span contains no cells.
    pub fn delete_rows(&mut self, selector: Selector<'_>, start: u16, count: u16) -> Result<()> {
        self.stage_rows(selector, start, count, false)
    }

    /// Inserts empty column coordinates and moves every retained cell right.
    ///
    /// # Errors
    ///
    /// Returns a typed grid, dependency-closure, or finite-bound error.
    pub fn insert_columns(&mut self, selector: Selector<'_>, start: u8, count: u8) -> Result<()> {
        self.stage_columns(selector, start, count, true)
    }

    /// Deletes column coordinates and their cells, moving later cells left.
    ///
    /// # Errors
    ///
    /// Returns a typed dependency or reversibility refusal. Deletion is
    /// durable only when the deleted coordinate span contains no cells.
    pub fn delete_columns(&mut self, selector: Selector<'_>, start: u8, count: u8) -> Result<()> {
        self.stage_columns(selector, start, count, false)
    }

    /// Renames one worksheet tab while preserving its `BoundSheet` identity.
    ///
    /// # Errors
    ///
    /// Returns a typed name, ambiguity, allocation, or finite-bound error.
    pub fn rename_sheet(&mut self, selector: Selector<'_>, name: &str) -> Result<()> {
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
            return Err(Error::UnsafeEdit(format!(
                "worksheet name {name:?} would be ambiguous"
            )));
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
            return Err(Error::UnsafeEdit(
                "row insertion exceeds the BIFF8 grid".into(),
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
            return Err(Error::UnsafeEdit(
                "column insertion exceeds the BIFF8 grid".into(),
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
        let sheet = operation_sheet_index(&change);
        if !matches!(change, StructuralChange::RenameSheet { .. })
            && self.resource_changes.iter().any(|resource| {
                formula_resource_cell(resource)
                    .is_some_and(|(resource_sheet, _)| resource_sheet == sheet)
            })
        {
            return Err(Error::UnsafeEdit(
                "structural worksheet edit overlaps an authored Formula".into(),
            ));
        }
        if self
            .structural_changes
            .iter()
            .any(|existing| structural_changes_overlap(existing, &change))
        {
            return Err(Error::UnsafeEdit(
                "overlapping structural operations must be prepared and joined separately".into(),
            ));
        }
        let fixed_overlap = self.changes.iter().any(|fixed| {
            fixed.sheet == sheet
                && match &change {
                    StructuralChange::Cell { reference, .. } => fixed.reference == *reference,
                    StructuralChange::Rows { .. } | StructuralChange::Columns { .. } => true,
                    StructuralChange::RenameSheet { .. } => false,
                }
        });
        if fixed_overlap {
            return Err(Error::UnsafeEdit(
                "fixed-width and structural operations overlap in one worksheet".into(),
            ));
        }
        let observed = self
            .changes
            .len()
            .checked_add(self.structural_changes.len())
            .and_then(|value| value.checked_add(self.resource_changes.len()))
            .ok_or_else(|| Error::InvalidData("cell transaction size overflow".into()))?;
        if observed >= MAX_STAGED_CHANGES {
            return Err(Error::UnsafeEdit(format!(
                "cell transaction exceeds its {MAX_STAGED_CHANGES}-change limit"
            )));
        }
        self.structural_changes
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("staging structural XLS changes"))?;
        self.structural_changes.push(change);
        Ok(())
    }

    fn stage(
        &mut self,
        sheet_index: usize,
        entry: usize,
        storage: Storage,
        value: Value,
    ) -> Result<()> {
        if self.structural_changes.iter().any(|change| {
            operation_sheet_index(change) == sheet_index
                && match change {
                    StructuralChange::Cell {
                        reference: structural_reference,
                        ..
                    } => {
                        *structural_reference
                            == self.source.inner.sheets[sheet_index].entries[entry]
                                .cell
                                .reference
                    },
                    StructuralChange::Rows { .. } | StructuralChange::Columns { .. } => true,
                    StructuralChange::RenameSheet { .. } => false,
                }
        }) {
            return Err(Error::UnsafeEdit(
                "fixed-width and structural operations overlap in one worksheet".into(),
            ));
        }
        if let Some(change) = self
            .changes
            .iter_mut()
            .find(|change| change.sheet == sheet_index && change.entry == entry)
        {
            change.storage = storage;
            change.value = value;
        } else {
            if self
                .changes
                .len()
                .saturating_add(self.structural_changes.len())
                .saturating_add(self.resource_changes.len())
                >= MAX_STAGED_CHANGES
            {
                return Err(Error::UnsafeEdit(format!(
                    "cell transaction exceeds its {MAX_STAGED_CHANGES}-change limit"
                )));
            }
            self.changes
                .try_reserve(1)
                .map_err(|_error| Error::Allocation("staging Number cell changes"))?;
            self.changes.push(Change {
                sheet: sheet_index,
                entry,
                reference: self.source.inner.sheets[sheet_index].entries[entry]
                    .cell
                    .reference,
                storage,
                value,
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
    pub fn commit_source_backed(self) -> Result<SourceBackedCommit> {
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
    /// through the public [`Workbook`] reader over a positional source before
    /// this method returns. The overlay plan performs exact source and target
    /// fingerprint checks again whenever it is read or published.
    ///
    /// The result deliberately does not expose the ordinary reversible
    /// [`Patch`], because that contract retains complete before/after CFB
    /// artifacts. Callers that require an in-memory target snapshot or an
    /// artifact inverse can continue to use [`Self::commit_source_backed`].
    /// The plan removes the complete target CFB allocation; the bounded
    /// semantic validation itself may still allocate the Workbook reader's
    /// stream and parsed model while this method is running.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for unsupported staged operations, protected,
    /// signed, encrypted, or macro-bearing sources, stale numeric ranges,
    /// invalid RK values, CFB limits, failed public Workbook validation, or
    /// source freshness/fingerprint changes.
    pub fn commit_source_backed_plan(self) -> Result<SourceBackedPlanCommit> {
        commit_source_backed_numeric_plan(self)
    }

    /// Alias for [`Self::commit_source_backed_plan`] emphasizing that the
    /// returned target is evaluated lazily at publication time.
    pub fn commit_source_backed_lazy(self) -> Result<SourceBackedPlanCommit> {
        self.commit_source_backed_plan()
    }

    /// Publishes a fully reopened snapshot and reversible exact-source patch.
    ///
    /// # Errors
    ///
    /// Returns an error if stream patching, CFB reconstruction, complete XLS
    /// reopen, or typed semantic readback fails.
    pub fn commit(self) -> Result<Commit> {
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
        let semantic = SemanticPatch::from_transaction(
            &self.source,
            &self.changes,
            &self.structural_changes,
            &self.resource_changes,
        )?;
        if changed_cells == 0
            && self.structural_changes.is_empty()
            && self.resource_changes.is_empty()
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
        if sst_delta != 0 {
            update_sst_total(
                &mut workbook_bytes,
                self.source.inner.sst_total_offset,
                sst_delta,
            )?;
        }

        if !self.structural_changes.is_empty() || !self.resource_changes.is_empty() {
            workbook_bytes = structural::apply(
                workbook_bytes,
                &self.source,
                &self.structural_changes,
                &self.resource_changes,
                &shared_strings,
            )?;
        }

        let workbook: Arc<[u8]> = Arc::from(workbook_bytes);
        let mut package = PackageEditor::open(
            self.source.inner.bytes.to_vec(),
            Targets::default(),
            Limits::default(),
        )?;
        package.put_stream_shared(&self.source.inner.workbook_path, workbook)?;
        // The committed package editor has already rendered, reopened, and
        // recaptured the candidate CFB. Reuse it so snapshot construction
        // performs the owner parse and complete Workbook validation once.
        let snapshot = if fixed_cells != 0
            && self.structural_changes.is_empty()
            && self.resource_changes.is_empty()
            && changes_are_fixed_numeric(&self.source, &self.changes)
        {
            Snapshot::from_fixed_numeric_package_editor(package, &self.source, &self.changes)?
        } else {
            Snapshot::from_package_editor(package)?
        };
        verify_readback(&snapshot, &self.changes)?;
        verify_structural_readback(&snapshot, &self.source, &self.structural_changes)?;
        verify_resource_readback(&snapshot, &self.resource_changes)?;
        let patch = Patch::new(
            Arc::clone(&self.source.inner.bytes),
            Arc::clone(&snapshot.inner.bytes),
            semantic,
        );
        Ok(Commit {
            snapshot,
            patch,
            diagnostics: Diagnostics {
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

impl fmt::Debug for Transaction {
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
        return Err(Error::UnsafeEdit(
            "cross-workbook scalar transfer inventory exceeds its checked bound".into(),
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
        return Err(Error::UnsafeEdit(
            "target cell precedes the scalar transfer anchor".into(),
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
        return Err(Error::UnsafeEdit(
            "target cell lies outside the scalar transfer source range".into(),
        ));
    }
    Ok(reference)
}

fn ensure_transfer_has_no_staged_target_overlap(
    transaction: &Transaction,
    sheet: usize,
    target_range: CellRange,
) -> Result<()> {
    if transaction
        .changes
        .iter()
        .any(|change| change.sheet == sheet && target_range.contains(change.reference))
        || transaction.structural_changes.iter().any(|change| {
            operation_sheet_index(change) == sheet
                && !matches!(change, StructuralChange::RenameSheet { .. })
        })
        || transaction.resource_changes.iter().any(|change| {
            formula_resource_cell(change).is_some_and(|(resource_sheet, reference)| {
                resource_sheet == sheet && target_range.contains(reference)
            })
        })
    {
        return Err(Error::UnsafeEdit(
            "cross-workbook scalar transfer overlaps a staged target operation".into(),
        ));
    }
    Ok(())
}

fn ensure_scalar_transfer_container(snapshot: &Snapshot, sheet: usize) -> Result<()> {
    let mut records = Records::new(&snapshot.inner.workbook_stream);
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
                return Err(Error::UnsafeEdit(format!(
                    "cross-workbook scalar transfer refuses protected or dependency BIFF record 0x{kind:04X}"
                )));
            },
            kind if !transfer_known_global_record(kind) => {
                return Err(Error::UnsafeEdit(format!(
                    "cross-workbook scalar transfer refuses unknown global BIFF record 0x{kind:04X}"
                )));
            },
            _ => {},
        }
    }
    let position = bound_sheet_position.ok_or_else(|| {
        Error::UnsafeEdit("selected worksheet has no unambiguous BoundSheet owner".into())
    })?;
    let start = usize::try_from(position)
        .map_err(|_error| Error::InvalidData("BoundSheet position does not fit usize".into()))?;
    let worksheet = snapshot
        .inner
        .workbook_stream
        .get(start..)
        .ok_or_else(|| Error::UnsafeEdit("BoundSheet owner lies outside Workbook".into()))?;
    let mut worksheet_records = Records::new(worksheet);
    let first = worksheet_records
        .next()
        .ok_or(Error::Eof("worksheet BOF"))??;
    require_bof(first.payload(), WORKSHEET)?;
    let mut found_eof = false;
    for record_result in worksheet_records {
        let record = record_result?;
        let kind = record.kind().get();
        if transfer_unsafe_record(kind) || transfer_protected_record(kind, record.payload()) {
            return Err(Error::UnsafeEdit(format!(
                "cross-workbook scalar transfer refuses protected, dependency, drawing, or unsupported BIFF record 0x{kind:04X}"
            )));
        }
        if kind == EOF {
            found_eof = true;
            break;
        }
        if !transfer_known_worksheet_record(kind) {
            return Err(Error::UnsafeEdit(format!(
                "cross-workbook scalar transfer refuses unknown worksheet BIFF record 0x{kind:04X}"
            )));
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
    changed_cells: usize,
    changed_number_fields: usize,
    touched_streams: usize,
}

impl Diagnostics {
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
pub struct Commit {
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
    /// Records this commit in bounded history after checking its exact base.
    ///
    /// # Errors
    ///
    /// Returns a stale-history error or a finite history-weight refusal
    /// without changing the history.
    pub fn record_in(self, history: &mut History) -> Result<Vec<Snapshot>> {
        if history.current().bytes() != self.patch.before() {
            return Err(Error::UnsafeEdit(
                "XLS cell history current snapshot does not match commit base".into(),
            ));
        }
        let weight = u64::try_from(self.patch.before.len())
            .ok()
            .and_then(|before| {
                u64::try_from(self.patch.after.len())
                    .ok()
                    .and_then(|after| before.checked_add(after))
            })
            .ok_or_else(|| Error::InvalidData("XLS history weight overflow".into()))?;
        history
            .record(self.snapshot, weight)
            .map_err(|error| Error::UnsafeEdit(error.to_string()))
    }

    /// Returns the reopened target snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Returns the exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Returns content-free publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Splits this commit into its snapshot, patch, and diagnostics.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch, Diagnostics) {
        (self.snapshot, self.patch, self.diagnostics)
    }
}

/// Content-free evidence for a source-backed fixed-width numeric publication.
///
/// The value retains bounded counts, lengths, opaque CFB fingerprints, and
/// source lineage tokens only. Its operation shape is a low-level contract
/// derived from the validated CFB plan; it reports logical pass scopes rather
/// than runtime counters or allocator/syscall activity. It never copies
/// workbook payloads or semantic cell values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourceBackedDiagnostics {
    changed_cells: usize,
    touched_streams: usize,
    splice_count: usize,
    replacement_bytes: u64,
    changed_spans: usize,
    source_bytes: u64,
    source_workbook_bytes: u64,
    target_workbook_bytes: u64,
    source_version: SourceVersion,
    target_version: SourceVersion,
    source_fingerprint: ArtifactFingerprint,
    target_fingerprint: ArtifactFingerprint,
    operation_shape: OverlayOperationShape,
}

impl SourceBackedDiagnostics {
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

/// Successful source-backed fixed-width numeric publication.
///
/// The reusable [`ValidatedOverlayPlan`] owns only the immutable positional
/// source and changed physical spans. The target snapshot is retained to keep
/// the ordinary XLS semantic and exact reversible-patch contract available to
/// callers that need immediate readback or inverse application. Consequently,
/// this API avoids Workbook-stream reserialization but still materializes one
/// complete target CFB allocation; it is not a bounded-artifact-memory API.
pub struct SourceBackedCommit {
    source: Snapshot,
    snapshot: Snapshot,
    plan: ValidatedOverlayPlan,
    patch: Patch,
    diagnostics: SourceBackedDiagnostics,
}

impl SourceBackedCommit {
    /// Exact immutable source snapshot used by this publication.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Fully reopened target snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Exact-source reversible semantic/artifact patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Reusable checked CFB publication plan.
    #[must_use]
    pub const fn plan(&self) -> &ValidatedOverlayPlan {
        &self.plan
    }

    /// Content-free source/target publication evidence.
    #[must_use]
    pub const fn diagnostics(&self) -> SourceBackedDiagnostics {
        self.diagnostics
    }

    /// Whether every staged numeric replacement was an exact byte no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.plan.is_noop()
    }

    /// Streams the complete validated target to a sequential sink.
    ///
    /// A sink failure may leave a typed prefix in the sink. The returned
    /// [`OverlayError`] retains that partial-output progress.
    pub fn write_to<W: Write>(
        &self,
        writer: &mut W,
    ) -> std::result::Result<PublishReport, OverlayError> {
        self.plan.write_to(writer)
    }

    /// Alias emphasizing the forward-only publication boundary.
    pub fn publish_to_stream<W: Write>(
        &self,
        writer: &mut W,
    ) -> std::result::Result<PublishReport, OverlayError> {
        self.write_to(writer)
    }

    /// Publishes through the common synced sibling-temp/atomic-rename path.
    pub fn save<P: AsRef<std::path::Path>>(
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
/// [`Snapshot`], complete target `Vec<u8>`, or target artifact `Arc<[u8]>`.
/// It keeps the immutable source and the checked CFB overlay plan. The plan
/// owns the compact exact physical replacement spans derived from the numeric
/// splices; logical path and expected-range descriptors are consumed during
/// validation rather than duplicated here. Direct sequential publication
/// retains source/target hashing during its one emission pass and relies on
/// the source's sealed immutable `Arc<[u8]>` ownership instead of generic
/// pre/post scans. Checked composed views retain their complete preflight;
/// atomic saves use the sealed ownership to omit redundant outer scans while
/// retaining complete emission hashes and durability steps.
///
/// The plan is intentionally forward-only. An exact artifact inverse would
/// need a source-bound reverse plan rooted in the composed target; the
/// existing [`Patch`] contract instead retains complete before/after
/// artifacts. Use [`Transaction::commit_source_backed`] when that inverse
/// contract is required.
pub struct SourceBackedPlanCommit {
    source: Snapshot,
    plan: ValidatedOverlayPlan,
    diagnostics: SourceBackedDiagnostics,
}

/// Explicit alias for callers that prefer the numeric-plan terminology.
pub type SourceBackedNumericPlanCommit = SourceBackedPlanCommit;

impl SourceBackedPlanCommit {
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

    /// Whether every staged numeric replacement was an exact byte no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.plan.is_noop()
    }

    /// Returns a checked positional view of the composed target without
    /// materializing the complete artifact.
    ///
    /// The view performs complete source and target fingerprint checks before
    /// it is returned and repeats source freshness checks on each positional
    /// read.
    pub fn composed_source(&self) -> std::result::Result<ComposedOverlaySource, OverlayError> {
        self.plan.composed_source()
    }

    /// Streams the complete validated target to a sequential sink.
    ///
    /// A sink failure may leave a typed prefix in the sink; inspect the
    /// returned [`OverlayError`] before deciding whether to retry.
    pub fn write_to<W: Write>(
        &self,
        writer: &mut W,
    ) -> std::result::Result<PublishReport, OverlayError> {
        self.plan.write_to(writer)
    }

    /// Alias for [`Self::write_to`] emphasizing sequential publication.
    pub fn publish_to_stream<W: Write>(
        &self,
        writer: &mut W,
    ) -> std::result::Result<PublishReport, OverlayError> {
        self.write_to(writer)
    }

    /// Publishes through the common synced sibling-temp and atomic-rename
    /// path.
    pub fn save<P: AsRef<std::path::Path>>(
        &self,
        path: P,
    ) -> std::result::Result<PublishReport, OverlayError> {
        self.plan.save(path)
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
    pub fn len(&self) -> usize {
        self.inner.operations().len()
    }

    /// Whether this patch has no semantic operation.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.inner.operations().is_empty()
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
    pub fn apply(&self, source: &Snapshot) -> Result<Commit> {
        let mut transaction = source.transaction();
        let mut sheet_names = semantic_sheet_names(source);
        for operation in self.inner.operations() {
            apply_semantic_operation(source, &mut transaction, operation, &mut sheet_names)?;
        }
        transaction.commit()
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
    pub fn plan_three_way(base: &Snapshot, left: &Self, right: &Self) -> Result<ThreeWayPlan> {
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
    pub fn plan_transfer(&self, target: &Snapshot) -> TransferPlan {
        let mut refusals = Vec::new();
        let mut transaction = target.transaction();
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
    pub fn execute(&self, target: &Snapshot) -> Result<Commit> {
        if let Some(refusal) = self.refusals.first() {
            return Err(Error::UnsafeEdit(format!(
                "semantic transfer is not executable at {}: {}",
                refusal.target, refusal.reason
            )));
        }
        self.patch.apply(target)
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
                let entry = unique_entry(&sheet_data.entries, *reference)?.ok_or_else(|| {
                    Error::UnsafeEdit("Formula dependency owner is absent".into())
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
            if !insert
                && snapshot.inner.sheets[*sheet].entries.iter().any(|entry| {
                    entry.cell.reference.row() >= *start && entry.cell.reference.row() < end
                })
            {
                return Err(Error::UnsafeEdit(
                    "durable row deletion is refused when the deleted span contains cells".into(),
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
            if !insert
                && snapshot.inner.sheets[*sheet].entries.iter().any(|entry| {
                    entry.cell.reference.column() >= *start && entry.cell.reference.column() < end
                })
            {
                return Err(Error::UnsafeEdit(
                    "durable column deletion is refused when the deleted span contains cells"
                        .into(),
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
            let entry = unique_entry(&sheet_data.entries, *reference)?
                .ok_or_else(|| Error::UnsafeEdit("Formula dependency owner is absent".into()))?;
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
            let sheet_data = snapshot.inner.sheets.get(*sheet).ok_or_else(|| {
                Error::UnsafeEdit("Formula semantic sheet dependency is stale".into())
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
            .ok_or_else(|| Error::UnsafeEdit("Formula semantic sheet dependency is stale".into()))?
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
            .ok_or_else(|| Error::UnsafeEdit("Formula semantic sheet dependency is stale".into()))?
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
    let mut expected = Vec::new();
    crate::writer::biff::write_formula(
        &mut expected,
        u32::from(reference.row()),
        u16::from(reference.column()),
        style.get(),
        tokens,
    )?;
    let start = entry.kind_offset;
    let end = start
        .checked_add(expected.len())
        .ok_or_else(|| Error::InvalidData("authored Formula record range overflow".into()))?;
    Ok(snapshot
        .inner
        .workbook_stream
        .get(start..end)
        .is_some_and(|record| record == expected.as_slice()))
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
        return Err(Error::UnsafeEdit(
            "Formula token dependency is stale".into(),
        ));
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
        .ok_or_else(|| Error::UnsafeEdit("structural cell XF dependency is absent".into()))
}

fn preflight_semantic(source: &Snapshot, patch: &SemanticPatch) -> Result<()> {
    let mut transaction = source.transaction();
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

fn semantic_operations_overlap(left: &PatchOperation, right: &PatchOperation) -> bool {
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
    if operation.op == "sheet.rename" {
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
            | "formula.insert"
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
    transaction: &mut Transaction,
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
                transaction.insert_rows(Selector::Position(sheet), start, count)
            } else {
                transaction.delete_rows(Selector::Position(sheet), start, count)
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
                transaction.insert_columns(Selector::Position(sheet), start, count)
            } else {
                transaction.delete_columns(Selector::Position(sheet), start, count)
            }
        },
        "sheet.rename" => {
            let sheet = parse_sheet_target(&operation.target, "name")?;
            let sheet_index = verify_semantic_sheet(source, sheet_names, sheet, operation)?;
            let name = operation
                .value
                .as_str()
                .ok_or_else(|| Error::InvalidData("sheet rename value is not text".into()))?;
            transaction.rename_sheet(Selector::Position(sheet), name)?;
            sheet_names[sheet_index] = name.to_string();
            Ok(())
        },
        _ => Err(Error::InvalidData(format!(
            "unsupported XLS cell patch operation {:?}",
            operation.op
        ))),
    }
}

fn apply_sst_semantic(transaction: &mut Transaction, operation: &PatchOperation) -> Result<()> {
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
        return Err(Error::UnsafeEdit(
            "SST resource precondition is stale".into(),
        ));
    }
    transaction.stage_resource(ResourceChange::SharedString {
        text: text.to_string(),
        insert: operation.op == "sst.intern",
    })
}

fn apply_rich_sst_semantic(
    transaction: &mut Transaction,
    operation: &PatchOperation,
) -> Result<()> {
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
        return Err(Error::UnsafeEdit(
            "rich SST resource precondition is stale".into(),
        ));
    }
    transaction.stage_resource(ResourceChange::RichSharedString {
        text: text.to_string(),
        formatting_runs,
        insert: operation.op == "sst.intern-rich",
    })
}

fn apply_xf_semantic(transaction: &mut Transaction, operation: &PatchOperation) -> Result<()> {
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
        return Err(Error::UnsafeEdit(
            "XF resource precondition is stale".into(),
        ));
    }
    transaction.stage_resource(ResourceChange::ExtendedFormat {
        index,
        payload,
        insert: matches!(operation.op.as_str(), "xf.author" | "xf.duplicate"),
    })
}

fn apply_formula_semantic(
    source: &Snapshot,
    transaction: &mut Transaction,
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
        .ok_or_else(|| Error::UnsafeEdit("Formula XF resource is absent".into()))?;
    if actual_xf != expected_xf {
        return Err(Error::UnsafeEdit(
            "Formula XF resource dependency is stale".into(),
        ));
    }
    let expected = operation
        .preconditions
        .get("present")
        .and_then(serde_json::Value::as_bool)
        .ok_or_else(|| {
            Error::InvalidData("Formula resource has no presence precondition".into())
        })?;
    let reference = Reference::new(u32::from(row), u32::from(column))?;
    let sheet_data =
        source.inner.sheets.get(sheet).ok_or_else(|| {
            Error::UnsafeEdit("Formula semantic sheet dependency is stale".into())
        })?;
    let current = unique_entry(&sheet_data.entries, reference)?;
    if !expected && current.is_some() {
        return Err(Error::UnsafeEdit(
            "Formula insertion precondition target is occupied".into(),
        ));
    }
    let present = match current {
        Some(entry) => authored_formula_record_matches(source, entry, reference, style, &tokens)?,
        None => false,
    };
    if present != expected {
        return Err(Error::UnsafeEdit(
            "Formula resource precondition is stale".into(),
        ));
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
        .chunks_exact(2)
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
    transaction: &mut Transaction,
    operation: &PatchOperation,
    sheet_names: &[String],
) -> Result<()> {
    let (sheet_position, reference) = parse_semantic_target(&operation.target)?;
    let sheet_index = verify_semantic_sheet(source, sheet_names, sheet_position, operation)?;
    let sheet = &source.inner.sheets[sheet_index];
    let entry_index = unique_entry_index(&sheet.entries, reference)?
        .ok_or_else(|| Error::UnsafeEdit("semantic patch cell is absent".into()))?;
    let cell = &sheet.entries[entry_index].cell;
    let expected = operation
        .preconditions
        .get("state")
        .ok_or_else(|| Error::InvalidData("cell patch has no state precondition".into()))?;
    let (expected_storage, expected_value) = parse_cell_state(expected)?;
    if cell.storage != expected_storage || !values_equal(&cell.value, &expected_value) {
        return Err(Error::UnsafeEdit(
            "semantic patch cell precondition is stale".into(),
        ));
    }
    verify_formula_dependency(source, &sheet.entries[entry_index], operation)?;
    let (storage, value) = parse_cell_state(&operation.value)?;
    let shared_strings = transaction.effective_shared_strings();
    let represented = target_storage(cell.storage, &value, &shared_strings)?;
    if represented != storage {
        return Err(Error::InvalidData(
            "cell patch target storage disagrees with its value".into(),
        ));
    }
    if matches!(cell.value, Value::FormulaCache(FormulaCache::String(_)))
        || matches!(value, Value::FormulaCache(FormulaCache::String(_)))
    {
        return transaction.set_value(Selector::Position(sheet_position), reference, value);
    }
    transaction.stage(sheet_index, entry_index, storage, value)
}

fn apply_structural_cell_semantic(
    source: &Snapshot,
    transaction: &mut Transaction,
    operation: &PatchOperation,
    sheet_names: &[String],
) -> Result<()> {
    let (sheet_position, reference) = parse_semantic_target(&operation.target)?;
    let sheet_index = verify_semantic_sheet(source, sheet_names, sheet_position, operation)?;
    let sheet = &source.inner.sheets[sheet_index];
    let current = unique_entry(&sheet.entries, reference)?.map(|entry| {
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
        return Err(Error::UnsafeEdit(
            "structural cell patch precondition is stale".into(),
        ));
    }
    if let Some(entry) = unique_entry(&sheet.entries, reference)? {
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
            .ok_or_else(|| Error::UnsafeEdit("structural cell XF resource is absent".into()))?;
        if actual_xf != expected_xf {
            return Err(Error::UnsafeEdit(
                "structural cell XF resource dependency is stale".into(),
            ));
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
            let shared_strings = transaction.effective_shared_strings();
            if storage_for_new_value(&value, &shared_strings)? != storage {
                return Err(Error::InvalidData(
                    "structural cell storage disagrees with its value".into(),
                ));
            }
            transaction.insert_cell_with_style(
                Selector::Position(sheet_position),
                reference,
                value,
                style,
            )
        },
        (Some(_), None) => transaction.remove_cell(Selector::Position(sheet_position), reference),
        (Some(before), Some(after)) if before.0 == after.0 && values_equal(&before.1, &after.1) => {
            transaction.set_style(Selector::Position(sheet_position), reference, after.2)
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
        .resolve_sheet(Selector::Position(sheet_position))?
        .ok_or_else(|| Error::UnsafeEdit("semantic patch worksheet is absent".into()))?;
    let expected_name = operation
        .preconditions
        .get("sheet_name")
        .and_then(serde_json::Value::as_str)
        .ok_or_else(|| Error::InvalidData("patch has no sheet-name precondition".into()))?;
    if sheet_names.get(sheet_index).map(String::as_str) != Some(expected_name) {
        return Err(Error::UnsafeEdit(
            "semantic patch worksheet identity is stale".into(),
        ));
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
pub struct Patch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
    semantic: SemanticPatch,
}

impl Patch {
    fn new(before: Arc<[u8]>, after: Arc<[u8]>, semantic: SemanticPatch) -> Self {
        Self {
            before,
            after,
            semantic,
        }
    }

    /// Returns whether the patch is byte-exactly empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before == self.after
    }

    /// Returns the exact required source artifact.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Returns the exact produced artifact.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Durable semantic operation patch for this exact artifact transition.
    #[must_use]
    pub const fn semantic(&self) -> &SemanticPatch {
        &self.semantic
    }

    /// Applies this patch only to its exact immutable source artifact.
    ///
    /// # Errors
    ///
    /// Returns [`Error::UnsafeEdit`] for a stale source, or a typed package
    /// validation error if the retained target can no longer be reopened.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.bytes() != self.before() {
            return Err(Error::UnsafeEdit(
                "XLS cell-value patch source does not match its base snapshot".into(),
            ));
        }
        if self.is_empty() {
            return Ok(source.clone());
        }
        Snapshot::from_bytes(self.after.to_vec())
    }

    /// Returns the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
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
            .field("is_empty", &self.is_empty())
            .finish()
    }
}

#[derive(Clone)]
struct SnapshotSource {
    bytes: Arc<[u8]>,
    version: SourceVersion,
}

impl SnapshotSource {
    fn new(bytes: Arc<[u8]>, version: SourceVersion) -> Self {
        Self { bytes, version }
    }
}

impl ReadAt for SnapshotSource {
    fn len(&self) -> std::io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_error| std::io::Error::other("XLS snapshot length exceeds u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
        let start = match usize::try_from(offset) {
            Ok(start) => start,
            Err(_error) => return Ok(0),
        };
        let Some(source) = self.bytes.get(start..) else {
            return Ok(0);
        };
        let count = source.len().min(output.len());
        output[..count].copy_from_slice(&source[..count]);
        Ok(count)
    }

    fn version(&self) -> std::io::Result<SourceVersion> {
        Ok(self.version)
    }
}

/// A seekable adapter for semantic validation over a lazy composed CFB view.
///
/// `Workbook` is intentionally kept generic over `Read + Seek`, while the
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

fn commit_source_backed_numeric_plan(transaction: Transaction) -> Result<SourceBackedPlanCommit> {
    let Transaction {
        source,
        changes,
        structural_changes,
        resource_changes,
    } = transaction;

    if !structural_changes.is_empty() || !resource_changes.is_empty() {
        return Err(Error::UnsupportedFeature(
            "source-backed numeric plan accepts no structural or resource edits".into(),
        ));
    }
    if !changes_are_fixed_numeric(&source, &changes) {
        return Err(Error::UnsupportedFeature(
            "source-backed numeric plan requires unchanged Number, RK, or MulRk storage".into(),
        ));
    }

    // The complete source Workbook validation already ran when this immutable
    // snapshot was opened. Reuse its private policy facts here; the target is
    // still checked independently below over the composed positional reader.
    source.inner.source_policy.require()?;

    let publisher = SourceBackedOverlayPublisher::open_owned(
        Arc::clone(&source.inner.bytes),
        source.inner.source_version,
    )
    .map_err(source_backed_overlay_error)?;
    let source_version = publisher
        .source_version()
        .map_err(source_backed_overlay_error)?;
    if source_version != source.inner.source_version {
        return Err(Error::UnsafeEdit(
            "source-backed numeric plan lost its source lineage".into(),
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
    let workbook = Workbook::new(ComposedPositionalReader::new(candidate.clone()))?;
    require_public_worksheet_coverage(&workbook, &source.inner.sheets)?;
    require_unprotected_workbook(&workbook)?;
    require_macro_free_workbook(&workbook)?;
    verify_public_numeric_readback(&workbook, source, changes)?;
    drop(workbook);
    candidate.version().map_err(Error::Io)
}

fn commit_source_backed_numeric(transaction: Transaction) -> Result<SourceBackedCommit> {
    let Transaction {
        source,
        changes,
        structural_changes,
        resource_changes,
    } = transaction;

    if !structural_changes.is_empty() || !resource_changes.is_empty() {
        return Err(Error::UnsupportedFeature(
            "source-backed numeric publication accepts no structural or resource edits".into(),
        ));
    }
    if !changes_are_fixed_numeric(&source, &changes) {
        return Err(Error::UnsupportedFeature(
            "source-backed numeric publication requires unchanged Number, RK, or MulRk storage"
                .into(),
        ));
    }

    let semantic = SemanticPatch::from_transaction(&source, &changes, &[], &[])?;
    let source_workbook = Workbook::new(Cursor::new(source.bytes()))?;
    require_public_worksheet_coverage(&source_workbook, &source.inner.sheets)?;
    require_unprotected_workbook(&source_workbook)?;
    require_macro_free_workbook(&source_workbook)?;

    let source_adapter: Arc<dyn ReadAt> = Arc::new(SnapshotSource::new(
        Arc::clone(&source.inner.bytes),
        source.inner.source_version,
    ));
    let publisher =
        SourceBackedOverlayPublisher::open(source_adapter).map_err(source_backed_overlay_error)?;
    let source_version = publisher
        .source_version()
        .map_err(source_backed_overlay_error)?;
    if source_version != source.inner.source_version {
        return Err(Error::UnsafeEdit(
            "source-backed numeric publication lost its source lineage".into(),
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
        let target = Snapshot::from_bytes(target_bytes.clone())?;
        if target.bytes() != target_bytes.as_slice() {
            return Err(Error::UnsafeEdit(
                "source-backed numeric target changed during complete reopen".into(),
            ));
        }
        verify_source_backed_numeric_target(&source, &target, &changes)?;
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
        .ok_or_else(|| Error::UnsafeEdit("source-backed numeric worksheet is stale".into()))?;
    let entry = sheet
        .entries
        .get(change.entry)
        .ok_or_else(|| Error::UnsafeEdit("source-backed numeric cell is stale".into()))?;
    if entry.cell.storage != change.storage
        || !matches!(
            change.storage,
            Storage::Number | Storage::Rk | Storage::MulRk
        )
    {
        return Err(Error::UnsafeEdit(
            "source-backed numeric edit changed its BIFF storage family".into(),
        ));
    }
    if entry.cell.reference != change.reference {
        return Err(Error::UnsafeEdit(
            "source-backed numeric edit changed its BIFF cell reference".into(),
        ));
    }
    let Value::Number(value) = change.value else {
        return Err(Error::UnsafeEdit(
            "source-backed numeric edit has a nonnumeric replacement".into(),
        ));
    };
    let width = match change.storage {
        Storage::Number => 8,
        Storage::Rk | Storage::MulRk => 4,
        Storage::BoolErr | Storage::Blank | Storage::LabelSst | Storage::Formula => {
            return Err(Error::UnsafeEdit(
                "source-backed numeric edit has an unsupported storage family".into(),
            ));
        },
    };
    let kind = binary::read_u16_le_at(&source.inner.workbook_stream, entry.kind_offset)?;
    if kind != storage_record_kind(entry.cell.storage) {
        return Err(Error::UnsafeEdit(
            "source-backed numeric record family is stale".into(),
        ));
    }
    let start = entry
        .value_offset
        .ok_or_else(|| Error::UnsafeEdit("numeric cell has no fixed-width value field".into()))?;
    let end = start
        .checked_add(width)
        .ok_or_else(|| Error::InvalidData("numeric splice range overflows usize".into()))?;
    let source_bytes = source
        .inner
        .workbook_stream
        .get(start..end)
        .ok_or_else(|| Error::InvalidData("numeric splice range is outside Workbook".into()))?;
    let Value::Number(source_value) = entry.cell.value else {
        return Err(Error::UnsafeEdit(
            "source cell does not contain a fixed-width numeric value".into(),
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
            return Err(Error::UnsafeEdit(
                "source cell does not contain a fixed-width numeric value".into(),
            ));
        },
    };
    if decoded_source_value.to_bits() != source_value.to_bits() {
        return Err(Error::UnsafeEdit(
            "source-backed numeric field failed its exact semantic precondition".into(),
        ));
    }
    let expected = source_bytes.to_vec();
    let replacement = match change.storage {
        Storage::Number => {
            if !valid_xnum(value) {
                return Err(Error::UnsupportedFeature(
                    "source-backed Number replacement is not a valid Xnum".into(),
                ));
            }
            value.to_le_bytes().to_vec()
        },
        Storage::Rk | Storage::MulRk => encode_rk(value)
            .ok_or_else(|| {
                Error::UnsupportedFeature(
                    "source-backed RK replacement is not exactly representable".into(),
                )
            })?
            .to_le_bytes()
            .to_vec(),
        Storage::BoolErr | Storage::Blank | Storage::LabelSst | Storage::Formula => {
            return Err(Error::UnsafeEdit(
                "source-backed numeric edit has an unsupported storage family".into(),
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

fn materialize_numeric_plan(plan: &ValidatedOverlayPlan, capacity: usize) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(capacity)
        .map_err(|_error| Error::Allocation("materializing source-backed numeric target"))?;
    plan.write_to(&mut bytes)
        .map_err(source_backed_overlay_error)?;
    if bytes.len() != capacity {
        return Err(Error::UnsafeEdit(
            "source-backed numeric publication changed CFB artifact length".into(),
        ));
    }
    Ok(bytes)
}

fn verify_source_backed_numeric_target(
    source: &Snapshot,
    target: &Snapshot,
    changes: &[Change],
) -> Result<()> {
    if source.inner.workbook_path != target.inner.workbook_path {
        return Err(Error::UnsafeEdit(
            "source-backed numeric publication changed the Workbook owner".into(),
        ));
    }
    let workbook = Workbook::new(Cursor::new(target.bytes()))?;
    require_public_worksheet_coverage(&workbook, &source.inner.sheets)?;
    require_unprotected_workbook(&workbook)?;
    require_macro_free_workbook(&workbook)?;
    let _ = carry_fixed_numeric_inventory(source, &target.inner.workbook_stream, changes)?;
    verify_public_numeric_readback(&workbook, source, changes)
}

fn source_backed_overlay_error(error: OverlayError) -> Error {
    match error {
        OverlayError::Ole(error) => Error::Cfb(error),
        OverlayError::Io(error) => Error::Io(error),
        OverlayError::Allocation { resource, .. } => Error::Allocation(resource),
        other => Error::UnsafeEdit(format!("source-backed numeric overlay refused: {other}")),
    }
}

impl From<OverlayError> for Error {
    fn from(error: OverlayError) -> Self {
        source_backed_overlay_error(error)
    }
}

fn workbook_protection_policy<R: Read + Seek>(
    workbook: &Workbook<R>,
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
        let protection = workbook.xls_worksheet(index)?.protection();
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

fn require_unprotected_workbook<R: Read + Seek>(workbook: &Workbook<R>) -> Result<()> {
    match workbook_protection_policy(workbook)? {
        SourceProtectionPolicy::Unprotected => Ok(()),
        SourceProtectionPolicy::WorkbookOrShared => Err(Error::UnsafeEdit(
            "protected or shared workbooks are not eligible for source-backed numeric edits".into(),
        )),
        SourceProtectionPolicy::Worksheet => Err(Error::UnsafeEdit(
            "protected worksheets are not eligible for source-backed numeric edits".into(),
        )),
    }
}

fn workbook_is_macro_free<R: Read + Seek>(workbook: &Workbook<R>) -> bool {
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

fn require_macro_free_workbook<R: Read + Seek>(workbook: &Workbook<R>) -> Result<()> {
    if workbook_is_macro_free(workbook) {
        Ok(())
    } else {
        Err(Error::UnsafeEdit(
            "macro-bearing XLS sources are not eligible for source-backed numeric edits".into(),
        ))
    }
}

fn require_macro_free_container(publisher: &SourceBackedOverlayPublisher) -> Result<()> {
    if publisher.shared().directory_entries().any(|entry| {
        entry.entry_type == STGTY_STORAGE && entry.name.eq_ignore_ascii_case("_VBA_PROJECT_CUR")
    }) {
        return Err(Error::UnsafeEdit(
            "macro-bearing XLS CFB storage is not eligible for source-backed numeric edits".into(),
        ));
    }
    Ok(())
}

fn parse_workbook_stream(
    source: &Arc<[u8]>,
) -> Result<(Vec<SheetData>, Option<usize>, Vec<Vec<u8>>)> {
    let mut records = Records::new(source);
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
                encoding = Encoding::from_codepage(binary::read_u16_le_at(record.payload(), 0)?)?;
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
        sheets
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("indexing editable worksheets"))?;
        sheets.push(SheetData {
            name: bound.name.clone(),
            workbook_index,
            entries: Arc::new(entries),
        });
    }
    if xf_records.is_empty() {
        return Err(Error::UnsafeEdit(
            "opened-workbook transaction requires at least one XF resource".into(),
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
    let mut records = Records::new(data);
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

fn target_storage(source: Storage, value: &Value, shared_strings: &[String]) -> Result<Storage> {
    let target = match (source, value) {
        (Storage::Number, Value::Number(number)) if valid_xnum(*number) => Storage::Number,
        (Storage::Rk | Storage::MulRk, Value::Number(number)) if encode_rk(*number).is_some() => {
            source
        },
        (Storage::LabelSst, Value::Number(number)) if encode_rk(*number).is_some() => Storage::Rk,
        (Storage::BoolErr, Value::Boolean(_) | Value::Error(_)) => Storage::BoolErr,
        (Storage::Blank, Value::Blank) => Storage::Blank,
        (Storage::LabelSst | Storage::Rk, Value::Text(text))
            if shared_strings.iter().any(|candidate| candidate == text) =>
        {
            Storage::LabelSst
        },
        (Storage::Formula, Value::FormulaCache(cache)) if valid_formula_cache(cache) => {
            Storage::Formula
        },
        _ => {
            return Err(Error::UnsafeEdit(format!(
                "value {value:?} is not representable in the fixed-width {source:?} slot"
            )));
        },
    };
    Ok(target)
}

fn storage_for_new_value(value: &Value, shared_strings: &[String]) -> Result<Storage> {
    match value {
        Value::Number(number) if valid_xnum(*number) => Ok(Storage::Number),
        Value::Boolean(_) | Value::Error(_) => Ok(Storage::BoolErr),
        Value::Blank => Ok(Storage::Blank),
        Value::Text(text) if shared_strings.iter().any(|candidate| candidate == text) => {
            Ok(Storage::LabelSst)
        },
        Value::Text(_) => Err(Error::UnsafeEdit(
            "new BIFF8 text is absent from the effective workbook SST".into(),
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
    transaction: &Transaction,
    text: &str,
    formatting_runs: &[crate::records::SharedStringFormatRun],
) -> Result<()> {
    let character_count = u16::try_from(text.encode_utf16().count())
        .map_err(|_error| Error::UnsafeEdit("rich shared string exceeds u16 characters".into()))?;
    if text.is_empty() || formatting_runs.is_empty() {
        return Err(Error::InvalidData(
            "rich shared string requires text and formatting runs".into(),
        ));
    }
    u16::try_from(formatting_runs.len()).map_err(|_error| {
        Error::UnsafeEdit("rich shared string has too many formatting runs".into())
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
            return Err(Error::UnsafeEdit(format!(
                "rich shared-string font {} is not used by an effective XF",
                run.font_index
            )));
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
        return Err(Error::UnsafeEdit(
            "fixed-width numeric edit changed the Workbook stream length".into(),
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
            .ok_or_else(|| Error::UnsafeEdit("fixed-width numeric dependency is stale".into()))?;
        if entry.cell.storage != change.storage {
            return Err(Error::UnsafeEdit(
                "fixed-width numeric edit changed its BIFF record family".into(),
            ));
        }
        let width = match (&change.storage, &change.value) {
            (Storage::Number, Value::Number(value)) => {
                require_numeric_field(workbook, entry, &value.to_le_bytes())?;
                8
            },
            (Storage::Rk | Storage::MulRk, Value::Number(value)) => {
                let encoded = encode_rk(*value).ok_or_else(|| {
                    Error::UnsafeEdit(
                        "numeric replacement is not exactly representable as RK".into(),
                    )
                })?;
                require_numeric_field(workbook, entry, &encoded.to_le_bytes())?;
                4
            },
            _ => {
                return Err(Error::UnsafeEdit(
                    "inventory reuse requires unchanged Number, RK, or MulRk storage".into(),
                ));
            },
        };
        let start = entry
            .value_offset
            .ok_or_else(|| Error::UnsafeEdit("numeric cell has no value field".into()))?;
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
            return Err(Error::UnsafeEdit(
                "fixed-width numeric edit fields overlap".into(),
            ));
        }
        if source.get(cursor..field.start) != workbook.get(cursor..field.start) {
            return Err(Error::UnsafeEdit(
                "fixed-width numeric edit changed bytes outside its value fields".into(),
            ));
        }
        cursor = field.end;
    }
    if source.get(cursor..) != workbook.get(cursor..) {
        return Err(Error::UnsafeEdit(
            "fixed-width numeric edit changed bytes outside its value fields".into(),
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
    let start = entry
        .value_offset
        .ok_or_else(|| Error::UnsafeEdit("numeric cell has no value field".into()))?;
    let end = start
        .checked_add(expected.len())
        .ok_or_else(|| Error::InvalidData("numeric value range overflow".into()))?;
    if workbook.get(start..end) != Some(expected) {
        return Err(Error::UnsafeEdit(
            "fixed-width numeric field failed exact encoded readback".into(),
        ));
    }
    Ok(())
}

fn require_public_worksheet_coverage<R: Read + Seek>(
    workbook: &Workbook<R>,
    sheets: &[SheetData],
) -> Result<()> {
    for sheet in sheets {
        if workbook
            .sheet(sheet.workbook_index)
            .and_then(crate::SheetMetadata::parsed_worksheet_index)
            .is_none()
        {
            return Err(Error::UnsafeEdit(format!(
                "worksheet at tab position {} was not published by the complete XLS reader",
                sheet.workbook_index
            )));
        }
    }
    Ok(())
}

fn verify_public_numeric_readback<R: Read + Seek>(
    workbook: &Workbook<R>,
    snapshot: &Snapshot,
    changes: &[Change],
) -> Result<()> {
    for change in changes {
        if !change_is_effective(snapshot, change) {
            continue;
        }
        let source_sheet = &snapshot.inner.sheets[change.sheet];
        let metadata = workbook.sheet(source_sheet.workbook_index).ok_or_else(|| {
            Error::UnsafeEdit("edited worksheet disappeared from public readback".into())
        })?;
        let worksheet_index = metadata.parsed_worksheet_index().ok_or_else(|| {
            Error::UnsafeEdit("edited tab is not a public worksheet on readback".into())
        })?;
        let cell = workbook
            .xls_worksheet(worksheet_index)?
            .get_cell(
                u32::from(change.reference.row()),
                u32::from(change.reference.column()),
            )
            .ok_or_else(|| {
                Error::UnsafeEdit("edited cell is absent from public readback".into())
            })?;
        let Value::Number(expected) = change.value else {
            return Err(Error::UnsafeEdit(
                "inventory reuse received a nonnumeric value".into(),
            ));
        };
        let actual = match cell.value() {
            CellValue::Float(value) | CellValue::DateTime(value) => *value,
            _ => {
                return Err(Error::UnsafeEdit(
                    "edited numeric cell has a nonnumeric public value".into(),
                ));
            },
        };
        if actual.to_bits() != expected.to_bits() {
            return Err(Error::UnsafeEdit(
                "edited numeric cell failed independent public readback".into(),
            ));
        }
    }
    Ok(())
}

fn structural_changes_overlap(left: &StructuralChange, right: &StructuralChange) -> bool {
    match (left, right) {
        (
            StructuralChange::RenameSheet { sheet: left, .. },
            StructuralChange::RenameSheet { sheet: right, .. },
        ) => left == right,
        (StructuralChange::RenameSheet { .. }, _) | (_, StructuralChange::RenameSheet { .. }) => {
            false
        },
        (
            StructuralChange::Cell {
                sheet: left_sheet,
                reference: left_reference,
                ..
            },
            StructuralChange::Cell {
                sheet: right_sheet,
                reference: right_reference,
                ..
            },
        ) => left_sheet == right_sheet && left_reference == right_reference,
        (left, right) => operation_sheet_index(left) == operation_sheet_index(right),
    }
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
    let offset = entry
        .value_offset
        .ok_or_else(|| Error::UnsafeEdit("cell storage has no editable value field".into()))?;
    match (&change.storage, &change.value) {
        (Storage::Number, Value::Number(value)) => {
            write_field(workbook, offset, &value.to_le_bytes())
        },
        (Storage::Rk | Storage::MulRk, Value::Number(value)) => {
            let encoded = encode_rk(*value).ok_or_else(|| {
                Error::UnsafeEdit("numeric replacement is not exactly representable as RK".into())
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
                .ok_or_else(|| {
                    Error::UnsafeEdit("replacement text is absent from the SST".into())
                })?;
            let index = u32::try_from(index)
                .map_err(|_error| Error::InvalidData("SST index exceeds u32".into()))?;
            write_field(workbook, offset, &index.to_le_bytes())
        },
        (Storage::Formula, Value::FormulaCache(cache)) => {
            write_field(workbook, offset, &encode_formula_cache(cache))
        },
        _ => Err(Error::UnsafeEdit(
            "staged value and target storage disagree".into(),
        )),
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
        Error::UnsafeEdit("RK/LabelSst conversion requires an existing SST header".into())
    })?;
    let current = binary::read_u32_le_at(workbook, offset)?;
    let updated = i64::from(current)
        .checked_add(delta)
        .and_then(|value| u32::try_from(value).ok())
        .ok_or_else(|| Error::InvalidData("SST total reference count overflow".into()))?;
    write_field(workbook, offset, &updated.to_le_bytes())
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
    left.chars()
        .flat_map(char::to_lowercase)
        .eq(right.chars().flat_map(char::to_lowercase))
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

fn unique_entry(entries: &[Entry], reference: Reference) -> Result<Option<&Entry>> {
    Ok(unique_entry_index(entries, reference)?.map(|index| &entries[index]))
}

fn unique_entry_index(entries: &[Entry], reference: Reference) -> Result<Option<usize>> {
    let mut found = None;
    for (index, entry) in entries.iter().enumerate() {
        if entry.cell.reference != reference {
            continue;
        }
        if found.replace(index).is_some() {
            return Err(Error::UnsafeEdit(format!(
                "cell ({}, {}) has duplicate BIFF8 Number records",
                reference.row(),
                reference.column()
            )));
        }
    }
    Ok(found)
}

fn verify_readback(snapshot: &Snapshot, changes: &[Change]) -> Result<()> {
    for change in changes {
        let entries = &snapshot
            .inner
            .sheets
            .get(change.sheet)
            .ok_or_else(|| Error::UnsafeEdit("edited worksheet disappeared on readback".into()))?
            .entries;
        let cell = &unique_entry(entries, change.reference)?
            .ok_or_else(|| Error::UnsafeEdit("edited cell was not found on readback".into()))?
            .cell;
        if cell.storage != change.storage || !values_equal(&cell.value, &change.value) {
            return Err(Error::UnsafeEdit(
                "edited cell value failed semantic readback".into(),
            ));
        }
    }
    Ok(())
}

fn verify_structural_readback(
    snapshot: &Snapshot,
    source: &Snapshot,
    changes: &[StructuralChange],
) -> Result<()> {
    for change in changes {
        match change {
            StructuralChange::Cell {
                sheet,
                reference,
                after,
                ..
            } => {
                let entries = &snapshot
                    .inner
                    .sheets
                    .get(*sheet)
                    .ok_or_else(|| {
                        Error::UnsafeEdit("structurally edited worksheet disappeared".into())
                    })?
                    .entries;
                let actual = unique_entry(entries, *reference)?
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
                        return Err(Error::UnsafeEdit(
                            "structural cell failed semantic readback".into(),
                        ));
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
                    return Err(Error::UnsafeEdit(
                        "worksheet rename failed semantic readback".into(),
                    ));
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
                    return Err(Error::UnsafeEdit(
                        "shared-string resource failed semantic readback".into(),
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
                    return Err(Error::UnsafeEdit(
                        "rich shared-string resource failed semantic readback".into(),
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
                let sheet_data = snapshot.inner.sheets.get(*sheet).ok_or_else(|| {
                    Error::UnsafeEdit("Formula readback sheet dependency is stale".into())
                })?;
                let current = unique_entry(&sheet_data.entries, *reference)?;
                let present = match current {
                    Some(entry) => authored_formula_record_matches(
                        snapshot, entry, *reference, *style, tokens,
                    )?,
                    None => false,
                };
                if present != *insert {
                    return Err(Error::UnsafeEdit(
                        "Formula resource failed semantic readback".into(),
                    ));
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
                    return Err(Error::UnsafeEdit(
                        "XF resource failed semantic readback".into(),
                    ));
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
    let after = &snapshot.inner.sheets[sheet].entries;
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
        let found = unique_entry(after, reference)?
            .ok_or_else(|| Error::UnsafeEdit("shifted cell disappeared on readback".into()))?;
        if found.cell.storage != entry.cell.storage
            || found.cell.style != entry.cell.style
            || !values_equal(&found.cell.value, &entry.cell.value)
        {
            return Err(Error::UnsafeEdit(
                "shifted cell changed storage, style, or value".into(),
            ));
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests;
