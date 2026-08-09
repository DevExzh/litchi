//! Lossless fixed-width edits of existing BIFF8 cell records.
//!
//! The owner covers `Number` (`[MS-XLS]` 2.4.180), standalone and packed `RK`,
//! `BoolErr`, `Blank`, `LabelSst`, and non-string `Formula` caches. Every
//! change is confined to an existing fixed-width field. The only record-family
//! conversion is equal-width standalone `RK` ↔ `LabelSst`, with an existing
//! SST value and atomic reference-count validation. Insertion and physical
//! removal are refused until row-block, `INDEX`/`DBCELL`, dimensions, and
//! formula dependencies can be rebuilt together. The complete CFB package is
//! reopened before publication, and every other captured stream retains its
//! exact payload.

use crate::records::{BoundSheetRecord, Encoding, SheetType};
use crate::{Error, Result, Workbook};
use litchi_biff::Records;
use litchi_core::binary;
pub use litchi_core::patch::HistoryLimits;
use litchi_core::patch::{
    BlobBundle, BlobLimits, Patch as CorePatch, PatchLimits, PatchOperation, Reversible,
    ReversibleOperation,
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
const FORMULA: u16 = 0x0006;
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

        let (mut sheets, sst_total_offset) = parse_workbook_stream(&workbook_stream)?;
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

/// Detached, failure-atomic edits of existing fixed-width BIFF8 cell fields.
#[derive(Clone)]
pub struct Transaction {
    source: Snapshot,
    changes: Vec<Change>,
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
        let observed = self.changes.len().saturating_add(additions);
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
    /// Numeric `Number`, `RK`, and `MulRk` cells, `BoolErr`, existing SST
    /// references, and non-string Formula caches are editable in place.
    /// A standalone `RK` and `LabelSst` can be converted into each other
    /// because both have the same ten-byte payload; the SST reference count is
    /// updated atomically. Text must already exist in the workbook SST.
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
        let entry = &entries[entry_index];
        let target = target_storage(
            entry.cell.storage,
            &value,
            &self.source.inner.shared_strings,
        )?;
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

    fn stage(
        &mut self,
        sheet_index: usize,
        entry: usize,
        storage: Storage,
        value: Value,
    ) -> Result<()> {
        if let Some(change) = self
            .changes
            .iter_mut()
            .find(|change| change.sheet == sheet_index && change.entry == entry)
        {
            change.storage = storage;
            change.value = value;
        } else {
            if self.changes.len() >= MAX_STAGED_CHANGES {
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
        let changed_cells = self
            .changes
            .iter()
            .filter(|change| change_is_effective(&self.source, change))
            .count();
        let semantic = SemanticPatch::from_changes(&self.source, &self.changes)?;
        if changed_cells == 0 {
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
            write_cell_value(
                &mut workbook_bytes,
                entry,
                change,
                &self.source.inner.shared_strings,
            )?;
        }
        if sst_delta != 0 {
            update_sst_total(
                &mut workbook_bytes,
                self.source.inner.sst_total_offset,
                sst_delta,
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
    fn from_changes(snapshot: &Snapshot, changes: &[Change]) -> Result<Self> {
        let limits = semantic_patch_limits();
        let mut operations = Vec::new();
        operations
            .try_reserve(changes.len())
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
        for operation in self.inner.operations() {
            if operation.op != "cell.set" {
                return Err(Error::InvalidData(format!(
                    "unsupported XLS cell patch operation {:?}",
                    operation.op
                )));
            }
            let (sheet_position, reference) = parse_semantic_target(&operation.target)?;
            let sheet_index = source
                .resolve_sheet(Selector::Position(sheet_position))?
                .ok_or_else(|| Error::UnsafeEdit("semantic patch worksheet is absent".into()))?;
            let sheet = &source.inner.sheets[sheet_index];
            let expected_name = operation
                .preconditions
                .get("sheet_name")
                .and_then(serde_json::Value::as_str)
                .ok_or_else(|| {
                    Error::InvalidData("cell patch has no sheet-name precondition".into())
                })?;
            if sheet.name != expected_name {
                return Err(Error::UnsafeEdit(
                    "semantic patch worksheet identity is stale".into(),
                ));
            }
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
            let (storage, value) = parse_cell_state(&operation.value)?;
            let represented = target_storage(cell.storage, &value, &source.inner.shared_strings)?;
            if represented != storage {
                return Err(Error::InvalidData(
                    "cell patch target storage disagrees with its value".into(),
                ));
            }
            transaction.stage(sheet_index, entry_index, storage, value)?;
        }
        transaction.commit()
    }
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

fn parse_workbook_stream(source: &Arc<[u8]>) -> Result<(Vec<SheetData>, Option<usize>)> {
    let mut records = Records::new(source);
    let first = records.next().ok_or(Error::Eof("Workbook globals BOF"))??;
    require_bof(first.payload(), WORKBOOK_GLOBALS)?;

    let mut encoding = Encoding::from_codepage(1252)?;
    let mut bound_payloads = Vec::new();
    let mut sst_total_offset = None;
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
    Ok((sheets, sst_total_offset))
}

fn parse_worksheet(data: &[u8], base_offset: usize) -> Result<Vec<Entry>> {
    let mut records = Records::new(data);
    let first = records.next().ok_or(Error::Eof("worksheet BOF"))??;
    require_bof(first.payload(), WORKSHEET)?;
    let mut entries = Vec::new();
    let mut found_eof = false;
    for record_result in records {
        let record = record_result?;
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
                        value: Value::Text(String::new()),
                    },
                    Some(payload_offset(kind_offset, 6)?),
                    kind_offset,
                    Some(index),
                )?;
            },
            FORMULA => parse_formula_entry(record.payload(), kind_offset, &mut entries)?,
            _ => {},
        }
    }
    if !found_eof {
        return Err(Error::Eof("worksheet EOF"));
    }
    Ok(entries)
}

fn parse_formula_entry(payload: &[u8], kind_offset: usize, entries: &mut Vec<Entry>) -> Result<()> {
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
            return Ok(());
        },
    };
    push_entry(
        entries,
        Cell {
            reference: parse_reference(payload, FORMULA)?,
            storage: Storage::Formula,
            value: Value::FormulaCache(cache),
        },
        Some(payload_offset(kind_offset, 6)?),
        kind_offset,
        None,
    )
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

fn valid_formula_cache(cache: &FormulaCache) -> bool {
    match cache {
        FormulaCache::Number(value) => valid_xnum(*value),
        FormulaCache::Boolean(_) | FormulaCache::Error(_) | FormulaCache::Empty => true,
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

#[cfg(test)]
mod tests;
