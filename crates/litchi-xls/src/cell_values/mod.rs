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
use crate::{Error, Result, Workbook};
use litchi_biff::Records;
use litchi_core::binary;
pub use litchi_core::patch::HistoryLimits;
use litchi_core::patch::{
    BlobBundle, BlobLimits, DiagnosticFingerprint, Patch as CorePatch, PatchLimits, PatchOperation,
    Reversible, ReversibleOperation,
};
use litchi_ole_common::object::{Editor as PackageEditor, Limits, Targets};
use std::collections::BTreeMap;
use std::fmt;
use std::io::Cursor;
use std::sync::Arc;

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
const BIFF8: u16 = 0x0600;
const WORKBOOK_GLOBALS: u16 = 0x0005;
const WORKSHEET: u16 = 0x0010;
const NUMBER_PAYLOAD_BYTES: usize = 14;
const NUMBER_VALUE_OFFSET: usize = 6;
const MAX_STAGED_CHANGES: usize = 4_096;

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
    entries: Vec<Entry>,
}

struct Inner {
    bytes: Arc<[u8]>,
    workbook_path: Vec<String>,
    workbook_stream: Arc<[u8]>,
    shared_strings: Arc<Vec<String>>,
    sst_total_offset: Option<usize>,
    xf_records: Arc<Vec<Vec<u8>>>,
    sheets: Vec<SheetData>,
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

        let (mut sheets, sst_total_offset, xf_records) = parse_workbook_stream(&workbook_stream)?;
        // A full semantic open catches cross-stream and workbook-global
        // dependencies before the narrower source-offset inventory is kept.
        // The legacy reader intentionally skips some malformed optional sheet
        // projections, so this edit owner additionally requires every sheet
        // it can mutate to have survived that complete semantic open.
        let shared_strings = {
            let workbook = Workbook::new(Cursor::new(source.as_slice()))?;
            for sheet in &sheets {
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
            workbook.shared_strings_shared()
        };
        resolve_shared_strings(&mut sheets, &shared_strings)?;
        Ok(Self {
            inner: Arc::new(Inner {
                bytes: Arc::from(source),
                workbook_path,
                workbook_stream,
                shared_strings,
                sst_total_offset,
                xf_records: Arc::new(xf_records),
                sheets,
            }),
        })
    }

    /// Returns exact source CFB bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.inner.bytes
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
                if resource_target(left) == resource_target(right)
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

    /// Sets an existing cell through a storage-aware, dependency-checked path.
    ///
    /// Numeric `Number`, `RK`, and `MulRk` cells, `BoolErr`, SST
    /// references, and non-string Formula caches are editable in place.
    /// A standalone `RK` and `LabelSst` can be converted into each other
    /// because both have the same ten-byte payload; the SST reference count is
    /// updated atomically. Missing simple text is interned into a bounded SST
    /// tail resource when the workbook has no `ExtSST` offset cache.
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

    fn effective_shared_strings(&self) -> Vec<String> {
        let mut strings = self.source.inner.shared_strings.as_ref().clone();
        for text in self
            .resource_changes
            .iter()
            .filter_map(|change| match change {
                ResourceChange::SharedString {
                    text,
                    insert: false,
                } => Some(text),
                ResourceChange::SharedString { insert: true, .. }
                | ResourceChange::ExtendedFormat { .. } => None,
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
                ResourceChange::SharedString { text, insert: true } => Some(text.clone()),
                ResourceChange::SharedString { insert: false, .. }
                | ResourceChange::ExtendedFormat { .. } => None,
            })
            .collect();
        inserted.sort_by_key(|text| text_fingerprint(text));
        strings.extend(inserted);
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
                        | ResourceChange::ExtendedFormat { .. } => None,
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
        structural::certify_shift(&self.source, sheet)?;
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
        structural::certify_shift(&self.source, sheet)?;
        self.stage_structural(StructuralChange::Columns {
            sheet,
            start,
            count,
            insert,
        })
    }

    fn stage_structural(&mut self, change: StructuralChange) -> Result<()> {
        if self
            .structural_changes
            .iter()
            .any(|existing| structural_changes_overlap(existing, &change))
        {
            return Err(Error::UnsafeEdit(
                "overlapping structural operations must be prepared and joined separately".into(),
            ));
        }
        let sheet = operation_sheet_index(&change);
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
        let changed_cells = fixed_cells.saturating_add(
            self.structural_changes
                .iter()
                .filter(|change| matches!(change, StructuralChange::Cell { .. }))
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
        parse_workbook_stream(&workbook)?;
        let mut package = PackageEditor::open(
            self.source.inner.bytes.to_vec(),
            Targets::default(),
            Limits::default(),
        )?;
        package.put_stream_shared(&self.source.inner.workbook_path, workbook)?;
        let candidate = package.finish()?;
        let snapshot = Snapshot::from_bytes(candidate)?;
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
            append_resource_semantic(change, limits, &mut operations)?;
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
    };
    let mut preconditions = BTreeMap::new();
    preconditions.insert("present".to_string(), serde_json::Value::Bool(present));
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
    let inverse = PatchOperation::new(limits, inverse_op, target, inverse_preconditions, value)
        .map_err(patch_error)?;
    operations.push(ReversibleOperation::new(forward, inverse));
    Ok(())
}

fn text_fingerprint(text: &str) -> String {
    DiagnosticFingerprint::of(text.as_bytes()).as_hex()
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
            ResourceChange::SharedString { .. } | ResourceChange::ExtendedFormat { .. } => None,
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
    left.target == right.target
}

fn semantic_operation_order(operation: &PatchOperation) -> u8 {
    if operation.op == "sheet.rename" {
        3
    } else if matches!(
        operation.op.as_str(),
        "sst.intern" | "sst.remove" | "xf.author" | "xf.duplicate" | "xf.remove"
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
        "xf.author" | "xf.duplicate" | "xf.remove" => apply_xf_semantic(transaction, operation),
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
            entries,
        });
    }
    if xf_records.is_empty() {
        return Err(Error::UnsafeEdit(
            "opened-workbook transaction requires at least one XF resource".into(),
        ));
    }
    if let Some(cell) = sheets
        .iter()
        .flat_map(|sheet| &sheet.entries)
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

    if !value.is_finite() {
        return None;
    }
    upper_word_if_exact(value).or_else(|| {
        let scaled = value * 100.0;
        scaled
            .is_finite()
            .then(|| upper_word_if_exact(scaled).map(|encoded| encoded | 1))
            .flatten()
    })
}

fn change_is_effective(snapshot: &Snapshot, change: &Change) -> bool {
    let source = &snapshot.inner.sheets[change.sheet].entries[change.entry].cell;
    source.storage != change.storage || !values_equal(&source.value, &change.value)
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
        _ => false,
    }
}

fn resource_target(change: &ResourceChange) -> String {
    match change {
        ResourceChange::SharedString { text, .. } => {
            format!("resource/sst/{}", text_fingerprint(text))
        },
        ResourceChange::ExtendedFormat { index, .. } => {
            format!("resource/xf/{:05}", index.get())
        },
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
    for entry in sheets.iter_mut().flat_map(|sheet| &mut sheet.entries) {
        let Some(index) = entry.sst_index else {
            continue;
        };
        let text =
            shared_strings
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
    for entry in before {
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
