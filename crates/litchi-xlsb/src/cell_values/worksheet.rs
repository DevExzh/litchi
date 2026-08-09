//! Source-bound worksheet-cell snapshots and length-stable edits.

use crate::package::error::{Error, Result};
use crate::raw::{Cursor, Header, Limits as RawLimits, Records, Writer, kind};
use litchi_core::binary;
use std::sync::Arc;

const MAX_ROW: u32 = 1_048_575;
const MAX_COLUMN: u32 = 16_383;
const MAX_STYLE_INDEX: u32 = 0x00ff_ffff;
const MAX_CELL_STRING_UNITS: usize = 32_767;

/// Finite resource policy for one worksheet cell-value snapshot.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    source_bytes: usize,
    cells: usize,
    raw: RawLimits,
}

impl Limits {
    /// Safe defaults for an ordinary worksheet part.
    pub const DEFAULT: Self = Self {
        source_bytes: 512 * 1024 * 1024,
        cells: 16_777_216,
        raw: RawLimits::DEFAULT,
    };

    /// Construct an explicit finite resource policy.
    #[must_use]
    pub const fn new(
        source_bytes: usize,
        cells: usize,
        record_bytes: usize,
        string_units: usize,
    ) -> Self {
        Self {
            source_bytes,
            cells,
            raw: RawLimits::new(record_bytes, string_units),
        }
    }

    /// Maximum worksheet-part bytes accepted by this owner.
    #[must_use]
    pub const fn source_bytes(self) -> usize {
        self.source_bytes
    }

    /// Maximum stored cell records retained in one snapshot.
    #[must_use]
    pub const fn cells(self) -> usize {
        self.cells
    }

    /// Per-record raw BIFF12 limits.
    #[must_use]
    pub const fn raw(self) -> RawLimits {
        self.raw
    }

    fn validate(self) -> Result<()> {
        if self.source_bytes == 0
            || self.cells == 0
            || self.raw.payload() == 0
            || self.raw.string_units() == 0
        {
            return Err(Error::InvalidFormat(
                "cell-value limits must all be nonzero".to_string(),
            ));
        }
        if self.raw.payload() > self.source_bytes {
            return Err(Error::InvalidFormat(
                "cell-value record-byte limit cannot exceed the worksheet-source limit".to_string(),
            ));
        }
        Ok(())
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// A checked zero-based worksheet cell reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Reference {
    row: u32,
    column: u32,
}

impl Reference {
    /// Construct a reference in the XLSB worksheet grid.
    ///
    /// # Errors
    ///
    /// Returns an error when either coordinate is outside the XLSB grid.
    pub fn new(row: u32, column: u32) -> Result<Self> {
        if row > MAX_ROW || column > MAX_COLUMN {
            return Err(Error::InvalidCellReference(format!(
                "cell ({row}, {column}) is outside the XLSB worksheet grid"
            )));
        }
        Ok(Self { row, column })
    }

    /// Return the zero-based row.
    #[must_use]
    pub const fn row(self) -> u32 {
        self.row
    }

    /// Return the zero-based column.
    #[must_use]
    pub const fn column(self) -> u32 {
        self.column
    }
}

/// A checked zero-based `BrtXF` index from a cell header.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct StyleIndex(u32);

impl StyleIndex {
    /// Construct a style index representable by the 24-bit `Cell.iStyleRef` field.
    ///
    /// The workbook publication boundary additionally verifies that this index
    /// exists in the current Styles part.
    ///
    /// # Errors
    ///
    /// Returns an error for a value outside the 24-bit wire range.
    pub fn new(value: u32) -> Result<Self> {
        if value > MAX_STYLE_INDEX {
            return Err(Error::UnsupportedFeature(format!(
                "cell style index {value} exceeds the 24-bit BIFF12 range"
            )));
        }
        Ok(Self(value))
    }

    /// Return the zero-based style index.
    #[must_use]
    pub const fn get(self) -> u32 {
        self.0
    }
}

impl TryFrom<u32> for StyleIndex {
    type Error = Error;

    fn try_from(value: u32) -> Result<Self> {
        Self::new(value)
    }
}

impl From<StyleIndex> for u32 {
    fn from(value: StyleIndex) -> Self {
        value.get()
    }
}

/// One of the eight error values permitted by `[MS-XLSB]` section 2.5.98.2.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum CellError {
    /// `#NULL!`.
    Null,
    /// `#DIV/0!`.
    DivisionByZero,
    /// `#VALUE!`.
    Value,
    /// `#REF!`.
    Reference,
    /// `#NAME?`.
    Name,
    /// `#NUM!`.
    Number,
    /// `#N/A`.
    NotAvailable,
    /// `#GETTING_DATA`.
    GettingData,
}

impl CellError {
    /// Decode the normative one-byte `BErr` value.
    ///
    /// # Errors
    ///
    /// Returns an error for a byte not assigned by `[MS-XLSB]` section 2.5.98.2.
    pub fn from_code(code: u8) -> Result<Self> {
        match code {
            0x00 => Ok(Self::Null),
            0x07 => Ok(Self::DivisionByZero),
            0x0f => Ok(Self::Value),
            0x17 => Ok(Self::Reference),
            0x1d => Ok(Self::Name),
            0x24 => Ok(Self::Number),
            0x2a => Ok(Self::NotAvailable),
            0x2b => Ok(Self::GettingData),
            _ => Err(Error::InvalidFormat(format!(
                "invalid BIFF12 BErr value 0x{code:02X}"
            ))),
        }
    }

    /// Encode the normative one-byte `BErr` value.
    #[must_use]
    pub const fn code(self) -> u8 {
        match self {
            Self::Null => 0x00,
            Self::DivisionByZero => 0x07,
            Self::Value => 0x0f,
            Self::Reference => 0x17,
            Self::Name => 0x1d,
            Self::Number => 0x24,
            Self::NotAvailable => 0x2a,
            Self::GettingData => 0x2b,
        }
    }
}

impl TryFrom<u8> for CellError {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self> {
        Self::from_code(value)
    }
}

impl From<CellError> for u8 {
    fn from(value: CellError) -> Self {
        value.code()
    }
}

/// The stored value family of an existing worksheet cell.
///
/// Formula variants represent only the cached result; editing one never
/// changes, evaluates, or resolves the formula token stream.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum Value {
    /// `BrtCellBlank`.
    Blank,
    /// Exactly representable `BrtCellRk` number.
    RkNumber(f64),
    /// `BrtCellError`.
    Error(CellError),
    /// `BrtCellBool`.
    Boolean(bool),
    /// `BrtCellReal`.
    Number(f64),
    /// Inline `BrtCellSt` string.
    InlineString(String),
    /// `BrtCellIsst` shared-string table index.
    SharedStringIndex(u32),
    /// Cached string in `BrtFmlaString`.
    FormulaStringCache(String),
    /// Cached number in `BrtFmlaNum`.
    FormulaNumberCache(f64),
    /// Cached Boolean in `BrtFmlaBool`.
    FormulaBooleanCache(bool),
    /// Cached error in `BrtFmlaError`.
    FormulaErrorCache(CellError),
}

impl Value {
    /// Whether this value is only the cached result attached to formula tokens.
    #[must_use]
    pub const fn is_formula_cache(&self) -> bool {
        matches!(
            self,
            Self::FormulaStringCache(_)
                | Self::FormulaNumberCache(_)
                | Self::FormulaBooleanCache(_)
                | Self::FormulaErrorCache(_)
        )
    }
}

impl PartialEq for Value {
    fn eq(&self, other: &Self) -> bool {
        match (self, other) {
            (Self::Blank, Self::Blank) => true,
            (Self::RkNumber(left), Self::RkNumber(right))
            | (Self::Number(left), Self::Number(right))
            | (Self::FormulaNumberCache(left), Self::FormulaNumberCache(right)) => {
                left.to_bits() == right.to_bits()
            },
            (Self::Error(left), Self::Error(right))
            | (Self::FormulaErrorCache(left), Self::FormulaErrorCache(right)) => left == right,
            (Self::Boolean(left), Self::Boolean(right))
            | (Self::FormulaBooleanCache(left), Self::FormulaBooleanCache(right)) => left == right,
            (Self::InlineString(left), Self::InlineString(right))
            | (Self::FormulaStringCache(left), Self::FormulaStringCache(right)) => left == right,
            (Self::SharedStringIndex(left), Self::SharedStringIndex(right)) => left == right,
            _ => false,
        }
    }
}

impl Eq for Value {}

/// One editable stored cell in worksheet stream order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StoredCell {
    reference: Reference,
    style: StyleIndex,
    show_phonetic: bool,
    value: Value,
}

impl StoredCell {
    /// Cell location selected by this item.
    #[must_use]
    pub const fn reference(&self) -> Reference {
        self.reference
    }

    /// Zero-based cell-XF index retained in the cell header.
    #[must_use]
    pub const fn style(&self) -> StyleIndex {
        self.style
    }

    /// Whether the source cell requests display of phonetic information.
    #[must_use]
    pub const fn show_phonetic(&self) -> bool {
        self.show_phonetic
    }

    /// Stored value or formula-cache family.
    #[must_use]
    pub const fn value(&self) -> &Value {
        &self.value
    }
}

/// An existing `BrtCellReal` value in source order.
#[derive(Debug, Clone, Copy)]
pub struct Number {
    reference: Reference,
    value: f64,
}

impl Number {
    /// Cell location selected by this item.
    #[must_use]
    pub const fn reference(self) -> Reference {
        self.reference
    }

    /// Stored IEEE-754 value, retained exactly on read.
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

/// Immutable, source-bound snapshot of editable stored cells.
#[derive(Debug, Clone)]
pub struct Snapshot {
    source: Arc<[u8]>,
    entries: Vec<Entry>,
    limits: Limits,
}

impl Snapshot {
    /// All supported stored cells in worksheet stream order.
    #[must_use]
    pub fn cells(&self) -> impl ExactSizeIterator<Item = &StoredCell> {
        self.entries.iter().map(|entry| &entry.cell)
    }

    /// Existing direct `BrtCellReal` values in worksheet stream order.
    ///
    /// This compatibility view deliberately excludes RK values and formula
    /// caches; use [`Self::cells`] for the complete typed inventory.
    #[must_use]
    pub fn numbers(&self) -> impl ExactSizeIterator<Item = Number> + '_ {
        Numbers {
            entries: self.entries.iter(),
            remaining: self
                .entries
                .iter()
                .filter(|entry| matches!(&entry.cell.value, Value::Number(_)))
                .count(),
        }
    }

    /// Look up one supported stored cell.
    ///
    /// # Errors
    ///
    /// Returns an error when duplicate source cell records make the coordinate
    /// ambiguous.
    pub fn cell(&self, reference: Reference) -> Result<Option<&StoredCell>> {
        Ok(unique_index(&self.entries, reference)?.map(|index| &self.entries[index].cell))
    }

    /// Look up an existing direct `BrtCellReal` value.
    ///
    /// # Errors
    ///
    /// Returns an error when duplicate source cell records make the coordinate
    /// ambiguous.
    pub fn number(&self, reference: Reference) -> Result<Option<Number>> {
        let Some(cell) = self.cell(reference)? else {
            return Ok(None);
        };
        Ok(match &cell.value {
            Value::Number(value) => Some(Number {
                reference,
                value: *value,
            }),
            Value::Blank
            | Value::RkNumber(_)
            | Value::Error(_)
            | Value::Boolean(_)
            | Value::InlineString(_)
            | Value::SharedStringIndex(_)
            | Value::FormulaStringCache(_)
            | Value::FormulaNumberCache(_)
            | Value::FormulaBooleanCache(_)
            | Value::FormulaErrorCache(_) => None,
        })
    }

    /// Start a detached edit against this exact source stream.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            source: Arc::clone(&self.source),
            entries: self.entries.clone(),
            limits: self.limits,
        }
    }

    /// Exact worksheet bytes guarded by commits and patches.
    #[must_use]
    pub fn source_bytes(&self) -> &[u8] {
        &self.source
    }

    /// Resource policy retained by edits and readback.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.source == other.source && self.limits == other.limits
    }
}

impl Eq for Snapshot {}

struct Numbers<'a> {
    entries: std::slice::Iter<'a, Entry>,
    remaining: usize,
}

impl Iterator for Numbers<'_> {
    type Item = Number;

    fn next(&mut self) -> Option<Self::Item> {
        for entry in self.entries.by_ref() {
            if let Value::Number(value) = &entry.cell.value {
                self.remaining = self.remaining.saturating_sub(1);
                return Some(Number {
                    reference: entry.cell.reference,
                    value: *value,
                });
            }
        }
        self.remaining = 0;
        None
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        (self.remaining, Some(self.remaining))
    }
}

impl ExactSizeIterator for Numbers<'_> {}

/// Detached, length-stable edits of existing cell fields.
#[derive(Debug, Clone)]
pub struct Edit {
    source: Arc<[u8]>,
    entries: Vec<Entry>,
    limits: Limits,
}

impl Edit {
    /// Replace a value while retaining its existing BIFF12 storage family.
    ///
    /// Inline and formula-cache strings must retain their UTF-16 code-unit
    /// count. RK replacements must be exactly representable. Formula variants
    /// alter only cached results, never formula tokens.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent/duplicate cell, a storage-family change,
    /// a non-finite number, a length-changing string, or an inexact RK value.
    pub fn set_value(&mut self, reference: Reference, value: Value) -> Result<()> {
        validate_value(&value)?;
        let index = required_index(&self.entries, reference)?;
        validate_replacement(&self.entries[index].cell.value, &value)?;
        if let Value::RkNumber(number) = &value {
            let _ = encode_rk(*number)?;
        }
        self.entries[index].cell.value = value;
        Ok(())
    }

    /// Set one existing direct numeric cell, including exact RK storage.
    ///
    /// # Errors
    ///
    /// Returns an error for non-finite or inexact values, duplicates, or a
    /// coordinate represented by another cell-record family.
    pub fn set_number(&mut self, reference: Reference, value: f64) -> Result<()> {
        let index = required_index(&self.entries, reference)?;
        let replacement = match &self.entries[index].cell.value {
            Value::Number(_) => Value::Number(value),
            Value::RkNumber(_) => Value::RkNumber(value),
            Value::Blank
            | Value::Error(_)
            | Value::Boolean(_)
            | Value::InlineString(_)
            | Value::SharedStringIndex(_)
            | Value::FormulaStringCache(_)
            | Value::FormulaNumberCache(_)
            | Value::FormulaBooleanCache(_)
            | Value::FormulaErrorCache(_) => {
                return Err(storage_error(reference, "a direct numeric cell"));
            },
        };
        self.set_value(reference, replacement)
    }

    /// Set one existing `BrtCellBool` value.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent/duplicate cell or another storage family.
    pub fn set_boolean(&mut self, reference: Reference, value: bool) -> Result<()> {
        self.set_value(reference, Value::Boolean(value))
    }

    /// Set one existing `BrtCellError` value.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent/duplicate cell or another storage family.
    pub fn set_error(&mut self, reference: Reference, value: CellError) -> Result<()> {
        self.set_value(reference, Value::Error(value))
    }

    /// Set one existing inline `BrtCellSt` string with the same UTF-16 length.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent/duplicate cell, another storage family,
    /// or a replacement with a different UTF-16 code-unit count.
    pub fn set_inline_string(&mut self, reference: Reference, value: String) -> Result<()> {
        self.set_value(reference, Value::InlineString(value))
    }

    /// Point one existing `BrtCellIsst` cell at another shared-string index.
    ///
    /// The workbook publication boundary verifies that the selected index
    /// exists and rejects the commit otherwise.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent/duplicate cell or another storage family.
    pub fn set_shared_string_index(&mut self, reference: Reference, index: u32) -> Result<()> {
        self.set_value(reference, Value::SharedStringIndex(index))
    }

    /// Set the cached numeric result of an existing `BrtFmlaNum` record.
    ///
    /// # Errors
    ///
    /// Returns an error for a non-finite value, absent/duplicate cell, or
    /// another storage family.
    pub fn set_formula_number_cache(&mut self, reference: Reference, value: f64) -> Result<()> {
        self.set_value(reference, Value::FormulaNumberCache(value))
    }

    /// Set the cached Boolean result of an existing `BrtFmlaBool` record.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent/duplicate cell or another storage family.
    pub fn set_formula_boolean_cache(&mut self, reference: Reference, value: bool) -> Result<()> {
        self.set_value(reference, Value::FormulaBooleanCache(value))
    }

    /// Set the cached error result of an existing `BrtFmlaError` record.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent/duplicate cell or another storage family.
    pub fn set_formula_error_cache(
        &mut self,
        reference: Reference,
        value: CellError,
    ) -> Result<()> {
        self.set_value(reference, Value::FormulaErrorCache(value))
    }

    /// Set the cached string result of an existing `BrtFmlaString` record.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent/duplicate cell, another storage family,
    /// or a replacement with a different UTF-16 code-unit count.
    pub fn set_formula_string_cache(&mut self, reference: Reference, value: String) -> Result<()> {
        self.set_value(reference, Value::FormulaStringCache(value))
    }

    /// Change the 24-bit style index of an existing stored cell.
    ///
    /// The workbook publication boundary verifies that the selected style
    /// exists in the current Styles part.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent or duplicate cell.
    pub fn set_style(&mut self, reference: Reference, style: StyleIndex) -> Result<()> {
        let index = required_index(&self.entries, reference)?;
        self.entries[index].cell.style = style;
        Ok(())
    }

    /// Validate and publish a source-checked reversible patch.
    ///
    /// # Errors
    ///
    /// Returns an error if a replacement range cannot be represented or the
    /// generated source stream fails bounded structural readback.
    pub fn commit(self) -> Result<Commit> {
        let changed = self.entries.iter().any(Entry::changed);
        let after = if changed {
            let mut bytes = self.source.to_vec();
            for entry in &self.entries {
                if entry.cell.style != entry.original.style {
                    write_style(&mut bytes, entry.style_offset, entry.cell.style)?;
                }
                if entry.cell.value != entry.original.value {
                    let encoded = encode_value(&entry.cell.value)?;
                    if encoded.len() != entry.value_len {
                        return Err(Error::InvalidFormat(
                            "cell-value replacement changed its bounded wire length".to_string(),
                        ));
                    }
                    let end = entry.value_offset.checked_add(entry.value_len).ok_or(
                        Error::CapacityOverflow {
                            resource: "cell-value replacement range",
                        },
                    )?;
                    let destination = bytes.get_mut(entry.value_offset..end).ok_or_else(|| {
                        Error::InvalidFormat(
                            "cell-value replacement is outside its source worksheet stream"
                                .to_string(),
                        )
                    })?;
                    destination.copy_from_slice(&encoded);
                }
            }
            Arc::from(bytes)
        } else {
            Arc::clone(&self.source)
        };
        let snapshot = read_shared(Arc::clone(&after), self.limits)?;
        Ok(Commit {
            snapshot,
            patch: Patch {
                before: self.source,
                after,
            },
        })
    }
}

/// Successful immutable cell-value commit.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Planned immutable worksheet snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Split this result into the planned snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// Reversible, source-checked worksheet-stream patch.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl Patch {
    /// Whether this patch leaves the source stream untouched.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before == self.after
    }

    /// Exact source required to apply this patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Exact stream produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Apply only to the exact source stream.
    ///
    /// # Errors
    ///
    /// Returns an error when `source` does not equal this patch's before image.
    pub fn apply(&self, source: &[u8]) -> Result<Vec<u8>> {
        if source != self.before.as_ref() {
            return Err(Error::UnsupportedFeature(
                "cell-value patch source snapshot does not match".to_string(),
            ));
        }
        Ok(self.after.to_vec())
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

#[derive(Debug, Clone)]
struct Entry {
    original: StoredCell,
    cell: StoredCell,
    style_offset: usize,
    value_offset: usize,
    value_len: usize,
}

impl Entry {
    fn changed(&self) -> bool {
        self.original != self.cell
    }
}

/// Read one complete worksheet stream with safe default limits.
pub(super) fn read(data: &[u8]) -> Result<Snapshot> {
    read_with_limits(data, Limits::DEFAULT)
}

/// Read one complete worksheet stream with an explicit finite policy.
pub(super) fn read_with_limits(data: &[u8], limits: Limits) -> Result<Snapshot> {
    read_shared(Arc::from(data), limits)
}

fn read_shared(source: Arc<[u8]>, limits: Limits) -> Result<Snapshot> {
    limits.validate()?;
    if source.len() > limits.source_bytes {
        return Err(Error::InvalidLength {
            expected: limits.source_bytes,
            found: source.len(),
        });
    }

    let mut entries = Vec::new();
    let mut in_sheet_data = false;
    let mut current_row = None;
    let mut frt_depth = 0usize;

    for item in Records::with_limits(&source, limits.raw) {
        let record = item?;
        match record.kind() {
            kind::BEGIN_SHEET_DATA => {
                if in_sheet_data {
                    return Err(Error::InvalidFormat(
                        "duplicate BrtBeginSheetData record".to_string(),
                    ));
                }
                in_sheet_data = true;
                current_row = None;
                frt_depth = 0;
            },
            kind::END_SHEET_DATA => {
                if !in_sheet_data {
                    return Err(Error::InvalidFormat(
                        "BrtEndSheetData without BrtBeginSheetData".to_string(),
                    ));
                }
                in_sheet_data = false;
                frt_depth = 0;
            },
            kind::FRT_BEGIN if in_sheet_data => {
                frt_depth = frt_depth.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "worksheet FRT nesting depth",
                })?;
            },
            kind::FRT_END if in_sheet_data => {
                frt_depth = frt_depth.saturating_sub(1);
            },
            kind::ROW_HDR if in_sheet_data && frt_depth == 0 => {
                if record.payload().len() < 17 {
                    return Err(Error::InvalidLength {
                        expected: 17,
                        found: record.payload().len(),
                    });
                }
                let row = binary::read_u32_le_at(record.payload(), 0)?;
                if row > MAX_ROW {
                    return Err(Error::InvalidCellReference(format!(
                        "BrtRowHdr row {row} is outside the XLSB worksheet grid"
                    )));
                }
                current_row = Some(row);
            },
            cell_kind if in_sheet_data && frt_depth == 0 && is_supported_cell(cell_kind) => {
                if entries.len() >= limits.cells {
                    return Err(Error::InvalidLength {
                        expected: limits.cells,
                        found: entries.len().saturating_add(1),
                    });
                }
                let row = current_row.ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "{} appears before a BrtRowHdr",
                        cell_kind_name(cell_kind)
                    ))
                })?;
                let entry = parse_entry(&source, &record, row, limits)?;
                entries
                    .try_reserve(1)
                    .map_err(|allocation_error| Error::Allocation {
                        resource: "cell-value snapshot entries",
                        source: allocation_error,
                    })?;
                entries.push(entry);
            },
            _ => {},
        }
    }
    if in_sheet_data {
        return Err(Error::InvalidFormat(
            "worksheet stream ended before BrtEndSheetData".to_string(),
        ));
    }
    Ok(Snapshot {
        source,
        entries,
        limits,
    })
}

fn parse_entry(
    source: &[u8],
    record: &crate::raw::Record<'_>,
    row: u32,
    limits: Limits,
) -> Result<Entry> {
    let payload = record.payload();
    if payload.len() < 8 {
        return Err(Error::InvalidLength {
            expected: 8,
            found: payload.len(),
        });
    }
    if payload[7] & 0xfe != 0 {
        return Err(Error::InvalidFormat(format!(
            "{} has nonzero reserved Cell flag bits 0x{:02X}",
            cell_kind_name(record.kind()),
            payload[7]
        )));
    }
    let reference = Reference::new(row, binary::read_u32_le_at(payload, 0)?)?;
    let style_raw =
        u32::from(payload[4]) | (u32::from(payload[5]) << 8) | (u32::from(payload[6]) << 16);
    let style = StyleIndex::new(style_raw)?;
    let (value, relative_offset, value_len) = parse_value(record.kind(), payload, limits.raw)?;
    let record_source = source.get(record.offset()..).ok_or_else(|| {
        Error::InvalidFormat(
            "cell record offset is outside its source worksheet stream".to_string(),
        )
    })?;
    let (_, header_len) = Header::parse(record_source, limits.raw)?;
    let payload_offset =
        record
            .offset()
            .checked_add(header_len)
            .ok_or(Error::CapacityOverflow {
                resource: "cell payload offset",
            })?;
    let style_offset = payload_offset
        .checked_add(4)
        .ok_or(Error::CapacityOverflow {
            resource: "cell style offset",
        })?;
    let value_offset =
        payload_offset
            .checked_add(relative_offset)
            .ok_or(Error::CapacityOverflow {
                resource: "cell value offset",
            })?;
    let cell = StoredCell {
        reference,
        style,
        show_phonetic: payload[7] & 1 != 0,
        value,
    };
    Ok(Entry {
        original: cell.clone(),
        cell,
        style_offset,
        value_offset,
        value_len,
    })
}

fn parse_value(
    kind_value: crate::raw::Kind,
    payload: &[u8],
    limits: RawLimits,
) -> Result<(Value, usize, usize)> {
    match kind_value {
        kind::CELL_BLANK => {
            require_exact(payload, 8)?;
            Ok((Value::Blank, 8, 0))
        },
        kind::CELL_RK => {
            require_exact(payload, 12)?;
            let mut cursor = Cursor::new(&payload[8..], "BrtCellRk");
            Ok((Value::RkNumber(cursor.read_rk()?), 8, 4))
        },
        kind::CELL_ERROR => {
            require_exact(payload, 9)?;
            Ok((Value::Error(CellError::from_code(payload[8])?), 8, 1))
        },
        kind::CELL_BOOL => {
            require_exact(payload, 9)?;
            Ok((Value::Boolean(parse_bool(payload[8])?), 8, 1))
        },
        kind::CELL_REAL => {
            require_exact(payload, 16)?;
            let value = binary::read_f64_le_at(payload, 8)?;
            validate_xnum(value, "BrtCellReal")?;
            Ok((Value::Number(value), 8, 8))
        },
        kind::CELL_ST => parse_string_value(payload, limits, false),
        kind::CELL_ISST => {
            require_exact(payload, 12)?;
            Ok((
                Value::SharedStringIndex(binary::read_u32_le_at(payload, 8)?),
                8,
                4,
            ))
        },
        kind::FMLA_STRING => parse_string_value(payload, limits, true),
        kind::FMLA_NUM => {
            require_at_least(payload, 18)?;
            let value = binary::read_f64_le_at(payload, 8)?;
            validate_xnum(value, "BrtFmlaNum cache")?;
            Ok((Value::FormulaNumberCache(value), 8, 8))
        },
        kind::FMLA_BOOL => {
            require_at_least(payload, 11)?;
            Ok((Value::FormulaBooleanCache(parse_bool(payload[8])?), 8, 1))
        },
        kind::FMLA_ERROR => {
            require_at_least(payload, 11)?;
            Ok((
                Value::FormulaErrorCache(CellError::from_code(payload[8])?),
                8,
                1,
            ))
        },
        _ => Err(Error::InvalidRecordType(kind_value.get())),
    }
}

fn parse_string_value(
    payload: &[u8],
    limits: RawLimits,
    formula: bool,
) -> Result<(Value, usize, usize)> {
    require_at_least(payload, if formula { 14 } else { 12 })?;
    let mut cursor = Cursor::with_limits(&payload[8..], "cell string", limits);
    let value = cursor.read_wide_string()?;
    let units = value.encode_utf16().count();
    if units > MAX_CELL_STRING_UNITS {
        return Err(Error::InvalidLength {
            expected: MAX_CELL_STRING_UNITS,
            found: units,
        });
    }
    if !formula && cursor.remaining() != 0 {
        return Err(Error::InvalidFormat(
            "BrtCellSt contains trailing bytes".to_string(),
        ));
    }
    if formula && cursor.remaining() < 2 {
        return Err(Error::InvalidFormat(
            "BrtFmlaString is missing its formula flags and token stream".to_string(),
        ));
    }
    let bytes = units.checked_mul(2).ok_or(Error::CapacityOverflow {
        resource: "cell string bytes",
    })?;
    Ok((
        if formula {
            Value::FormulaStringCache(value)
        } else {
            Value::InlineString(value)
        },
        12,
        bytes,
    ))
}

fn is_supported_cell(kind_value: crate::raw::Kind) -> bool {
    matches!(
        kind_value,
        kind::CELL_BLANK
            | kind::CELL_RK
            | kind::CELL_ERROR
            | kind::CELL_BOOL
            | kind::CELL_REAL
            | kind::CELL_ST
            | kind::CELL_ISST
            | kind::FMLA_STRING
            | kind::FMLA_NUM
            | kind::FMLA_BOOL
            | kind::FMLA_ERROR
    )
}

fn cell_kind_name(kind_value: crate::raw::Kind) -> &'static str {
    match kind_value {
        kind::CELL_BLANK => "BrtCellBlank",
        kind::CELL_RK => "BrtCellRk",
        kind::CELL_ERROR => "BrtCellError",
        kind::CELL_BOOL => "BrtCellBool",
        kind::CELL_REAL => "BrtCellReal",
        kind::CELL_ST => "BrtCellSt",
        kind::CELL_ISST => "BrtCellIsst",
        kind::FMLA_STRING => "BrtFmlaString",
        kind::FMLA_NUM => "BrtFmlaNum",
        kind::FMLA_BOOL => "BrtFmlaBool",
        kind::FMLA_ERROR => "BrtFmlaError",
        _ => "cell record",
    }
}

fn unique_index(entries: &[Entry], reference: Reference) -> Result<Option<usize>> {
    let mut found = None;
    for (index, entry) in entries.iter().enumerate() {
        if entry.cell.reference != reference {
            continue;
        }
        if found.replace(index).is_some() {
            return Err(Error::UnsupportedFeature(format!(
                "cell ({}, {}) has duplicate stored cell records",
                reference.row(),
                reference.column()
            )));
        }
    }
    Ok(found)
}

fn required_index(entries: &[Entry], reference: Reference) -> Result<usize> {
    unique_index(entries, reference)?.ok_or_else(|| {
        Error::UnsupportedFeature(format!(
            "cell ({}, {}) is absent or uses an unsupported cell-record family",
            reference.row(),
            reference.column()
        ))
    })
}

fn validate_replacement(original: &Value, replacement: &Value) -> Result<()> {
    if !same_storage_family(original, replacement) {
        return Err(Error::UnsupportedFeature(
            "cell-value edits cannot change the existing BIFF12 storage family".to_string(),
        ));
    }
    match (original, replacement) {
        (Value::InlineString(before), Value::InlineString(after))
        | (Value::FormulaStringCache(before), Value::FormulaStringCache(after)) => {
            let before_units = before.encode_utf16().count();
            let after_units = after.encode_utf16().count();
            if before_units != after_units {
                return Err(Error::UnsupportedFeature(format!(
                    "cell string replacement changes UTF-16 length from {before_units} to {after_units}"
                )));
            }
        },
        _ => {},
    }
    Ok(())
}

fn same_storage_family(left: &Value, right: &Value) -> bool {
    matches!(
        (left, right),
        (Value::Blank, Value::Blank)
            | (Value::RkNumber(_), Value::RkNumber(_))
            | (Value::Error(_), Value::Error(_))
            | (Value::Boolean(_), Value::Boolean(_))
            | (Value::Number(_), Value::Number(_))
            | (Value::InlineString(_), Value::InlineString(_))
            | (Value::SharedStringIndex(_), Value::SharedStringIndex(_))
            | (Value::FormulaStringCache(_), Value::FormulaStringCache(_))
            | (Value::FormulaNumberCache(_), Value::FormulaNumberCache(_))
            | (Value::FormulaBooleanCache(_), Value::FormulaBooleanCache(_))
            | (Value::FormulaErrorCache(_), Value::FormulaErrorCache(_))
    )
}

fn validate_value(value: &Value) -> Result<()> {
    match value {
        Value::RkNumber(number) => validate_finite(*number, "BrtCellRk"),
        Value::Number(number) => validate_xnum(*number, "BrtCellReal"),
        Value::FormulaNumberCache(number) => validate_xnum(*number, "BrtFmlaNum cache"),
        Value::InlineString(string) | Value::FormulaStringCache(string) => {
            let units = string.encode_utf16().count();
            if units > MAX_CELL_STRING_UNITS {
                return Err(Error::InvalidLength {
                    expected: MAX_CELL_STRING_UNITS,
                    found: units,
                });
            }
            Ok(())
        },
        Value::Blank
        | Value::Error(_)
        | Value::Boolean(_)
        | Value::SharedStringIndex(_)
        | Value::FormulaBooleanCache(_)
        | Value::FormulaErrorCache(_) => Ok(()),
    }
}

fn validate_finite(value: f64, context: &'static str) -> Result<()> {
    if value.is_finite() {
        Ok(())
    } else {
        Err(Error::UnsupportedFeature(format!(
            "{context} edits require a finite IEEE-754 value"
        )))
    }
}

fn validate_xnum(value: f64, context: &'static str) -> Result<()> {
    let positive_zero = value.to_bits() == 0;
    if value.is_normal() || positive_zero {
        Ok(())
    } else {
        Err(Error::UnsupportedFeature(format!(
            "{context} must be a finite normalized IEEE-754 value or positive zero"
        )))
    }
}

fn encode_value(value: &Value) -> Result<Vec<u8>> {
    match value {
        Value::Blank => Ok(Vec::new()),
        Value::RkNumber(number) => encode_rk(*number),
        Value::Error(error) | Value::FormulaErrorCache(error) => Ok(vec![error.code()]),
        Value::Boolean(boolean) | Value::FormulaBooleanCache(boolean) => {
            Ok(vec![u8::from(*boolean)])
        },
        Value::Number(number) | Value::FormulaNumberCache(number) => {
            validate_xnum(*number, "numeric cell")?;
            Ok(number.to_le_bytes().to_vec())
        },
        Value::InlineString(string) | Value::FormulaStringCache(string) => {
            let units = string.encode_utf16().count();
            if units > MAX_CELL_STRING_UNITS {
                return Err(Error::InvalidLength {
                    expected: MAX_CELL_STRING_UNITS,
                    found: units,
                });
            }
            let capacity = units.checked_mul(2).ok_or(Error::CapacityOverflow {
                resource: "encoded cell string",
            })?;
            let mut bytes = Vec::new();
            bytes
                .try_reserve_exact(capacity)
                .map_err(|source| Error::Allocation {
                    resource: "encoded cell string",
                    source,
                })?;
            for unit in string.encode_utf16() {
                bytes.extend_from_slice(&unit.to_le_bytes());
            }
            Ok(bytes)
        },
        Value::SharedStringIndex(index) => Ok(index.to_le_bytes().to_vec()),
    }
}

fn encode_rk(value: f64) -> Result<Vec<u8>> {
    validate_finite(value, "BrtCellRk")?;
    let mut bytes = Vec::new();
    Writer::new(&mut bytes).write_rk(value)?;
    Ok(bytes)
}

fn write_style(bytes: &mut [u8], offset: usize, style: StyleIndex) -> Result<()> {
    let end = offset.checked_add(3).ok_or(Error::CapacityOverflow {
        resource: "cell style replacement range",
    })?;
    let destination = bytes.get_mut(offset..end).ok_or_else(|| {
        Error::InvalidFormat(
            "cell style replacement is outside its source worksheet stream".to_string(),
        )
    })?;
    let encoded = style.get().to_le_bytes();
    destination.copy_from_slice(&encoded[..3]);
    Ok(())
}

fn parse_bool(value: u8) -> Result<bool> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(Error::InvalidFormat(format!(
            "invalid BIFF12 Boolean byte 0x{value:02X}"
        ))),
    }
}

fn require_exact(payload: &[u8], expected: usize) -> Result<()> {
    if payload.len() == expected {
        Ok(())
    } else {
        Err(Error::InvalidLength {
            expected,
            found: payload.len(),
        })
    }
}

fn require_at_least(payload: &[u8], expected: usize) -> Result<()> {
    if payload.len() >= expected {
        Ok(())
    } else {
        Err(Error::InvalidLength {
            expected,
            found: payload.len(),
        })
    }
}

fn storage_error(reference: Reference, expected: &'static str) -> Error {
    Error::UnsupportedFeature(format!(
        "cell ({}, {}) is not encoded as {expected}",
        reference.row(),
        reference.column()
    ))
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    reason = "unit fixtures unwrap only values whose validity is the assertion setup"
)]
mod tests {
    use super::*;
    use crate::raw::{Kind, Writer};

    fn stream() -> Vec<u8> {
        let mut bytes = Vec::new();
        let mut writer = Writer::new(&mut bytes);
        writer
            .write_record(kind::BEGIN_SHEET_DATA, &[])
            .expect("begin");
        let mut row = vec![0; 17];
        row[..4].copy_from_slice(&3_u32.to_le_bytes());
        writer.write_record(kind::ROW_HDR, &row).expect("row");
        writer
            .write_record(Kind::new(0x1234).expect("kind"), &[7, 8, 9])
            .expect("unknown");

        let mut real = vec![0; 16];
        real[..4].copy_from_slice(&2_u32.to_le_bytes());
        real[8..].copy_from_slice(&4.5_f64.to_le_bytes());
        writer.write_record(kind::CELL_REAL, &real).expect("real");

        let mut boolean = vec![0; 9];
        boolean[..4].copy_from_slice(&3_u32.to_le_bytes());
        writer
            .write_record(kind::CELL_BOOL, &boolean)
            .expect("boolean");

        let mut inline = vec![0; 12];
        inline[..4].copy_from_slice(&4_u32.to_le_bytes());
        inline[8..12].copy_from_slice(&2_u32.to_le_bytes());
        inline.extend_from_slice(&[b'A', 0, b'B', 0]);
        writer
            .write_record(kind::CELL_ST, &inline)
            .expect("inline string");

        let mut rk = vec![0; 8];
        rk[..4].copy_from_slice(&5_u32.to_le_bytes());
        Writer::new(&mut rk).write_rk(25.0).expect("rk value");
        writer.write_record(kind::CELL_RK, &rk).expect("rk");

        let mut formula = vec![0; 18];
        formula[..4].copy_from_slice(&6_u32.to_le_bytes());
        formula[8..16].copy_from_slice(&7.5_f64.to_le_bytes());
        writer
            .write_record(kind::FMLA_NUM, &formula)
            .expect("formula number");

        let mut error = vec![0; 9];
        error[..4].copy_from_slice(&7_u32.to_le_bytes());
        error[8] = CellError::DivisionByZero.code();
        writer
            .write_record(kind::CELL_ERROR, &error)
            .expect("error");

        let mut shared = vec![0; 12];
        shared[..4].copy_from_slice(&8_u32.to_le_bytes());
        writer
            .write_record(kind::CELL_ISST, &shared)
            .expect("shared string");

        let mut formula_boolean = vec![0; 11];
        formula_boolean[..4].copy_from_slice(&9_u32.to_le_bytes());
        writer
            .write_record(kind::FMLA_BOOL, &formula_boolean)
            .expect("formula Boolean");

        let mut formula_error = vec![0; 11];
        formula_error[..4].copy_from_slice(&10_u32.to_le_bytes());
        formula_error[8] = CellError::NotAvailable.code();
        writer
            .write_record(kind::FMLA_ERROR, &formula_error)
            .expect("formula error");

        let mut formula_string = vec![0; 12];
        formula_string[..4].copy_from_slice(&11_u32.to_le_bytes());
        formula_string[8..12].copy_from_slice(&2_u32.to_le_bytes());
        formula_string.extend_from_slice(&[b'X', 0, b'Y', 0, 0, 0]);
        writer
            .write_record(kind::FMLA_STRING, &formula_string)
            .expect("formula string");
        writer.write_record(kind::END_SHEET_DATA, &[]).expect("end");
        bytes
    }

    #[test]
    fn edits_all_bounded_scalar_families_and_round_trips_patch() {
        let before = stream();
        let snapshot = read(&before).expect("snapshot");
        assert_eq!(snapshot.cells().len(), 10);

        let real_ref = Reference::new(3, 2).expect("real reference");
        let bool_ref = Reference::new(3, 3).expect("bool reference");
        let string_ref = Reference::new(3, 4).expect("string reference");
        let rk_ref = Reference::new(3, 5).expect("rk reference");
        let formula_ref = Reference::new(3, 6).expect("formula reference");
        let error_ref = Reference::new(3, 7).expect("error reference");
        let shared_ref = Reference::new(3, 8).expect("shared-string reference");
        let formula_bool_ref = Reference::new(3, 9).expect("formula Boolean reference");
        let formula_error_ref = Reference::new(3, 10).expect("formula error reference");
        let formula_string_ref = Reference::new(3, 11).expect("formula string reference");
        let mut edit = snapshot.edit();
        edit.set_number(real_ref, 9.25).expect("set real");
        edit.set_boolean(bool_ref, true).expect("set bool");
        edit.set_inline_string(string_ref, "CD".to_string())
            .expect("set string");
        edit.set_number(rk_ref, 125.0).expect("set RK");
        edit.set_formula_number_cache(formula_ref, 11.0)
            .expect("set cache");
        edit.set_error(error_ref, CellError::Reference)
            .expect("set error");
        edit.set_shared_string_index(shared_ref, 1)
            .expect("set shared-string index");
        edit.set_formula_boolean_cache(formula_bool_ref, true)
            .expect("set formula Boolean");
        edit.set_formula_error_cache(formula_error_ref, CellError::Number)
            .expect("set formula error");
        edit.set_formula_string_cache(formula_string_ref, "ZZ".to_string())
            .expect("set formula string");
        edit.set_style(real_ref, StyleIndex::new(1).expect("style"))
            .expect("set style");
        let commit = edit.commit().expect("commit");
        let after = commit.patch().apply(&before).expect("apply");

        assert_eq!(
            commit.snapshot().number(real_ref).expect("lookup"),
            Some(Number {
                reference: real_ref,
                value: 9.25,
            })
        );
        assert_eq!(
            commit
                .snapshot()
                .cell(real_ref)
                .expect("lookup")
                .expect("real")
                .style(),
            StyleIndex::new(1).expect("style")
        );
        assert_eq!(
            commit
                .snapshot()
                .cell(bool_ref)
                .expect("lookup")
                .expect("bool")
                .value(),
            &Value::Boolean(true)
        );
        assert_eq!(
            commit
                .snapshot()
                .cell(string_ref)
                .expect("lookup")
                .expect("string")
                .value(),
            &Value::InlineString("CD".to_string())
        );
        assert_eq!(
            commit
                .snapshot()
                .cell(rk_ref)
                .expect("lookup")
                .expect("RK")
                .value(),
            &Value::RkNumber(125.0)
        );
        assert_eq!(
            commit
                .snapshot()
                .cell(formula_ref)
                .expect("lookup")
                .expect("formula")
                .value(),
            &Value::FormulaNumberCache(11.0)
        );
        assert_eq!(
            commit
                .snapshot()
                .cell(error_ref)
                .expect("lookup")
                .expect("error")
                .value(),
            &Value::Error(CellError::Reference)
        );
        assert_eq!(
            commit
                .snapshot()
                .cell(shared_ref)
                .expect("lookup")
                .expect("shared string")
                .value(),
            &Value::SharedStringIndex(1)
        );
        assert_eq!(
            commit
                .snapshot()
                .cell(formula_bool_ref)
                .expect("lookup")
                .expect("formula Boolean")
                .value(),
            &Value::FormulaBooleanCache(true)
        );
        assert_eq!(
            commit
                .snapshot()
                .cell(formula_error_ref)
                .expect("lookup")
                .expect("formula error")
                .value(),
            &Value::FormulaErrorCache(CellError::Number)
        );
        assert_eq!(
            commit
                .snapshot()
                .cell(formula_string_ref)
                .expect("lookup")
                .expect("formula string")
                .value(),
            &Value::FormulaStringCache("ZZ".to_string())
        );
        assert_eq!(
            commit.patch().inverse().apply(&after).expect("revert"),
            before
        );
    }

    #[test]
    fn refuses_family_changes_length_changes_inexact_rk_and_stale_sources() {
        let snapshot = read(&stream()).expect("snapshot");
        let string_ref = Reference::new(3, 4).expect("string reference");
        let rk_ref = Reference::new(3, 5).expect("rk reference");
        let mut edit = snapshot.edit();
        assert!(edit.set_boolean(string_ref, true).is_err());
        assert!(
            edit.set_inline_string(string_ref, "longer".to_string())
                .is_err()
        );
        assert!(edit.set_number(rk_ref, 1.0 / 3.0).is_err());

        let real_ref = Reference::new(3, 2).expect("real reference");
        edit.set_number(real_ref, 1.0).expect("set real");
        let commit = edit.commit().expect("commit");
        assert!(commit.patch().apply(b"stale").is_err());
    }

    #[test]
    fn enforces_snapshot_limits_and_style_wire_range() {
        let bytes = stream();
        assert!(read_with_limits(&bytes, Limits::new(bytes.len() - 1, 5, 64, 32)).is_err());
        assert!(read_with_limits(&bytes, Limits::new(bytes.len(), 4, 64, 32)).is_err());
        assert!(StyleIndex::new(MAX_STYLE_INDEX).is_ok());
        assert!(StyleIndex::new(MAX_STYLE_INDEX + 1).is_err());
    }
}
