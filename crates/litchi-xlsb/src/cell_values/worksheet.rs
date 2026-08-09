//! Source-bound worksheet-cell snapshots and structural edits.

use crate::package::error::{Error, Result};
use crate::package::shared_strings::SharedString;
use crate::raw::{Cursor, Header, Limits as RawLimits, Records, Writer, kind};
use litchi_core::binary;
use std::collections::{BTreeMap, BTreeSet};
use std::sync::Arc;

const MAX_ROW: u32 = 1_048_575;
const MAX_COLUMN: u32 = 16_383;
const MAX_STYLE_INDEX: u32 = 0x00ff_ffff;
const MAX_CELL_STRING_UNITS: usize = 32_767;
const MAX_FORMULA_BYTES: usize = crate::formula::MAX_CELL_FORMULA_BYTES;
const PATCH_MAGIC: &[u8; 8] = b"LCXBCLP1";

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

/// Finite policy for durable patch transfer and in-memory history.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TransferLimits {
    bytes: usize,
    changes: usize,
    history_entries: usize,
    history_bytes: usize,
}

impl TransferLimits {
    /// Safe defaults for interactive worksheet editing.
    pub const DEFAULT: Self = Self {
        bytes: 1024 * 1024 * 1024 + 64,
        changes: 1_048_576,
        history_entries: 1024,
        history_bytes: 1024 * 1024 * 1024,
    };

    /// Construct a finite transfer/history policy.
    #[must_use]
    pub const fn new(
        bytes: usize,
        changes: usize,
        history_entries: usize,
        history_bytes: usize,
    ) -> Self {
        Self {
            bytes,
            changes,
            history_entries,
            history_bytes,
        }
    }

    /// Maximum encoded patch bytes.
    #[must_use]
    pub const fn bytes(self) -> usize {
        self.bytes
    }

    /// Maximum semantic changes in one patch.
    #[must_use]
    pub const fn changes(self) -> usize {
        self.changes
    }

    /// Maximum retained undo/redo entries.
    #[must_use]
    pub const fn history_entries(self) -> usize {
        self.history_entries
    }

    /// Maximum aggregate before/after bytes retained by history.
    #[must_use]
    pub const fn history_bytes(self) -> usize {
        self.history_bytes
    }

    fn validate(self) -> Result<()> {
        if self.bytes == 0
            || self.changes == 0
            || self.history_entries == 0
            || self.history_bytes == 0
        {
            return Err(Error::InvalidFormat(
                "cell patch transfer/history limits must all be nonzero".to_string(),
            ));
        }
        Ok(())
    }
}

impl Default for TransferLimits {
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
    /// Inline rich/phonetic `BrtCellRString` value.
    RichString(SharedString),
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
            (Self::RichString(left), Self::RichString(right)) => left == right,
            _ => false,
        }
    }
}

impl Eq for Value {}

/// Inert BIFF12 formula flags, token bytes, and ancillary bytes.
///
/// Construction validates the complete Ptg stream without evaluating it or
/// resolving workbook-scoped names and relationships. Contextual resolution
/// remains part of whole-workbook publication.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellFormula {
    flags: u16,
    tokens: Vec<u8>,
    ancillary: Vec<u8>,
}

impl CellFormula {
    /// Construct one validated cell formula.
    ///
    /// # Errors
    ///
    /// Returns an error for reserved flag bits, an oversized token stream, or
    /// structurally invalid token/ancillary bytes.
    pub fn new(flags: u16, tokens: Vec<u8>, ancillary: Vec<u8>) -> Result<Self> {
        let value = Self::from_source(flags, tokens, ancillary)?;
        let encoded_bytes = value
            .tokens
            .len()
            .checked_add(value.ancillary.len())
            .and_then(|size| size.checked_add(8))
            .ok_or(Error::CapacityOverflow {
                resource: "cell formula bytes",
            })?;
        if encoded_bytes > RawLimits::DEFAULT.payload() {
            return Err(Error::InvalidLength {
                expected: RawLimits::DEFAULT.payload(),
                found: encoded_bytes,
            });
        }
        crate::formula::Parser::with_extra(&value.tokens, &value.ancillary)
            .parse()
            .map_err(|error| Error::InvalidFormula(error.to_string()))?;
        Ok(value)
    }

    fn from_source(flags: u16, tokens: Vec<u8>, ancillary: Vec<u8>) -> Result<Self> {
        if flags & !0x0002 != 0 {
            return Err(Error::InvalidFormula(format!(
                "invalid GrbitFmla flags 0x{flags:04X}"
            )));
        }
        if tokens.len() > MAX_FORMULA_BYTES {
            return Err(Error::InvalidLength {
                expected: MAX_FORMULA_BYTES,
                found: tokens.len(),
            });
        }
        Ok(Self {
            flags,
            tokens,
            ancillary,
        })
    }

    /// Stored `GrbitFmla` flags.
    #[must_use]
    pub const fn flags(&self) -> u16 {
        self.flags
    }

    /// Formula Ptg token stream (`rgce`).
    #[must_use]
    pub fn tokens(&self) -> &[u8] {
        &self.tokens
    }

    /// Formula ancillary stream (`rgbExtra`).
    #[must_use]
    pub fn ancillary(&self) -> &[u8] {
        &self.ancillary
    }
}

/// One editable stored cell in worksheet stream order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StoredCell {
    reference: Reference,
    style: StyleIndex,
    show_phonetic: bool,
    value: Value,
    formula: Option<CellFormula>,
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

    /// Inert formula payload, present exactly for formula-cache record kinds.
    #[must_use]
    pub const fn formula(&self) -> Option<&CellFormula> {
        self.formula.as_ref()
    }
}

/// One semantic cell delta carried by a committed patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Change {
    reference: Reference,
    before: Option<StoredCell>,
    after: Option<StoredCell>,
}

impl Change {
    /// Changed cell coordinate.
    #[must_use]
    pub const fn reference(&self) -> Reference {
        self.reference
    }

    /// Source semantic state; `None` denotes insertion.
    #[must_use]
    pub const fn before(&self) -> Option<&StoredCell> {
        self.before.as_ref()
    }

    /// Destination semantic state; `None` denotes removal.
    #[must_use]
    pub const fn after(&self) -> Option<&StoredCell> {
        self.after.as_ref()
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
            | Value::RichString(_)
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
            inserted: BTreeMap::new(),
            removed: BTreeSet::new(),
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

/// Detached structural edits of worksheet cell records.
#[derive(Debug, Clone)]
pub struct Edit {
    source: Arc<[u8]>,
    entries: Vec<Entry>,
    inserted: BTreeMap<Reference, StoredCell>,
    removed: BTreeSet<Reference>,
    limits: Limits,
}

impl Edit {
    /// Replace a value while retaining its existing BIFF12 storage family.
    ///
    /// Strings may change UTF-16 length. RK replacements must be exactly
    /// representable. Formula variants alter only cached results; use
    /// [`Self::set_formula`] to replace the inert formula payload.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent/duplicate cell, a storage-family change,
    /// a non-finite number, or an inexact RK value.
    pub fn set_value(&mut self, reference: Reference, value: Value) -> Result<()> {
        validate_value_with_limits(&value, self.limits)?;
        if let Value::RkNumber(number) = &value {
            let _ = encode_rk(*number)?;
        }
        if let Some(cell) = self.inserted.get_mut(&reference) {
            validate_replacement(&cell.value, &value)?;
            cell.value = value;
            return Ok(());
        }
        self.removed.remove(&reference);
        let index = required_index(&self.entries, reference)?;
        validate_replacement(&self.entries[index].cell.value, &value)?;
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
            | Value::RichString(_)
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

    /// Set one existing inline `BrtCellSt` string.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent/duplicate cell, another storage family,
    /// or a value outside the XLSB string limit.
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

    /// Replace one existing inline rich/phonetic string.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid run metadata, an absent/duplicate cell, or
    /// another storage family.
    pub fn set_rich_string(&mut self, reference: Reference, value: SharedString) -> Result<()> {
        self.set_value(reference, Value::RichString(value))
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
    /// or a value outside the XLSB string limit.
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
        if let Some(cell) = self.inserted.get_mut(&reference) {
            cell.style = style;
            return Ok(());
        }
        self.removed.remove(&reference);
        let index = required_index(&self.entries, reference)?;
        self.entries[index].cell.style = style;
        Ok(())
    }

    /// Change the phonetic-display bit in an existing cell header.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent or duplicate cell.
    pub fn set_show_phonetic(&mut self, reference: Reference, show: bool) -> Result<()> {
        if let Some(cell) = self.inserted.get_mut(&reference) {
            cell.show_phonetic = show;
            return Ok(());
        }
        self.removed.remove(&reference);
        let index = required_index(&self.entries, reference)?;
        self.entries[index].cell.show_phonetic = show;
        Ok(())
    }

    /// Replace the inert token/ancillary payload of an existing formula cell.
    ///
    /// The cached-result record family and cell header remain unchanged.
    /// Workbook-scoped references are validated at publication.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent/non-formula/duplicate cell.
    pub fn set_formula(&mut self, reference: Reference, formula: CellFormula) -> Result<()> {
        if let Some(cell) = self.inserted.get_mut(&reference) {
            if cell.formula.is_none() {
                return Err(storage_error(reference, "a formula cell"));
            }
            cell.formula = Some(formula);
            return Ok(());
        }
        self.removed.remove(&reference);
        let index = required_index(&self.entries, reference)?;
        if self.entries[index].cell.formula.is_none() {
            return Err(storage_error(reference, "a formula cell"));
        }
        self.entries[index].cell.formula = Some(formula);
        Ok(())
    }

    /// Insert one direct scalar cell and retain the selected BIFF12 family.
    ///
    /// # Errors
    ///
    /// Returns an error for an occupied coordinate, a formula-cache value, or
    /// an invalid value/resource limit.
    pub fn insert(&mut self, reference: Reference, style: StyleIndex, value: Value) -> Result<()> {
        if value.is_formula_cache() {
            return Err(Error::InvalidFormula(
                "formula-cache insertion requires an inert formula payload".to_string(),
            ));
        }
        self.insert_cell(reference, style, value, None)
    }

    /// Insert one formula cell with a typed cached-result record family.
    ///
    /// # Errors
    ///
    /// Returns an error for an occupied coordinate, a direct scalar cache, or
    /// an invalid value/resource limit.
    pub fn insert_formula(
        &mut self,
        reference: Reference,
        style: StyleIndex,
        cache: Value,
        formula: CellFormula,
    ) -> Result<()> {
        if !cache.is_formula_cache() {
            return Err(Error::InvalidFormula(
                "formula insertion requires a formula-cache value family".to_string(),
            ));
        }
        self.insert_cell(reference, style, cache, Some(formula))
    }

    /// Remove one supported stored cell record.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent or duplicate coordinate.
    pub fn remove(&mut self, reference: Reference) -> Result<()> {
        if self.inserted.remove(&reference).is_some() {
            return Ok(());
        }
        let _ = required_index(&self.entries, reference)?;
        self.removed.insert(reference);
        Ok(())
    }

    fn insert_cell(
        &mut self,
        reference: Reference,
        style: StyleIndex,
        value: Value,
        formula: Option<CellFormula>,
    ) -> Result<()> {
        validate_value_with_limits(&value, self.limits)?;
        if let Value::RkNumber(number) = &value {
            let _ = encode_rk(*number)?;
        }
        let formula_family = value.is_formula_cache();
        if formula_family != formula.is_some() {
            return Err(Error::InvalidFormula(
                "cell formula presence does not match its cached-result family".to_string(),
            ));
        }
        if self.inserted.contains_key(&reference) {
            return Err(Error::UnsupportedFeature(format!(
                "cell ({}, {}) already exists",
                reference.row(),
                reference.column()
            )));
        }
        if let Some(index) = unique_index(&self.entries, reference)? {
            if !self.removed.remove(&reference) {
                return Err(Error::UnsupportedFeature(format!(
                    "cell ({}, {}) already exists",
                    reference.row(),
                    reference.column()
                )));
            }
            self.entries[index].cell = StoredCell {
                reference,
                style,
                show_phonetic: false,
                value,
                formula,
            };
            return Ok(());
        }
        if any_cell_at(&self.source, self.limits, reference)? {
            return Err(Error::UnsupportedFeature(format!(
                "cell ({}, {}) already exists in an unsupported record family",
                reference.row(),
                reference.column()
            )));
        }
        if self.inserted.len() >= self.limits.cells {
            return Err(Error::InvalidLength {
                expected: self.limits.cells,
                found: self.inserted.len().saturating_add(1),
            });
        }
        self.inserted.insert(
            reference,
            StoredCell {
                reference,
                style,
                show_phonetic: false,
                value,
                formula,
            },
        );
        Ok(())
    }

    /// Validate and publish a source-checked reversible patch.
    ///
    /// # Errors
    ///
    /// Returns an error if a replacement range cannot be represented or the
    /// generated source stream fails bounded structural readback.
    pub fn commit(self) -> Result<Commit> {
        let retained = self.entries.len().saturating_sub(self.removed.len());
        let final_cells =
            retained
                .checked_add(self.inserted.len())
                .ok_or(Error::CapacityOverflow {
                    resource: "structural cell count",
                })?;
        if final_cells > self.limits.cells {
            return Err(Error::InvalidLength {
                expected: self.limits.cells,
                found: final_cells,
            });
        }
        let changed = self.entries.iter().any(Entry::changed)
            || !self.inserted.is_empty()
            || !self.removed.is_empty();
        let after = if changed {
            Arc::from(rebuild_stream(&self)?)
        } else {
            Arc::clone(&self.source)
        };
        let snapshot = read_shared(Arc::clone(&after), self.limits)?;
        let changes = semantic_changes(&self.entries, &self.inserted, &self.removed, &snapshot)?;
        Ok(Commit {
            snapshot,
            patch: Patch {
                before: self.source,
                after,
                changes,
                limits: self.limits,
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
    changes: Vec<Change>,
    limits: Limits,
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

    /// Deterministic semantic deltas in row/column order.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
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
            changes: self
                .changes
                .iter()
                .map(|change| Change {
                    reference: change.reference,
                    before: change.after.clone(),
                    after: change.before.clone(),
                })
                .collect(),
            limits: self.limits,
        }
    }

    /// Compose two sequential exact-source patches.
    ///
    /// # Errors
    ///
    /// Returns an error unless this patch's destination is the next patch's
    /// exact source or bounded semantic readback fails.
    pub fn compose(&self, next: &Self) -> Result<Self> {
        if self.after != next.before {
            return Err(Error::UnsupportedFeature(
                "cell patch composition source is stale".to_string(),
            ));
        }
        patch_from_images(
            Arc::clone(&self.before),
            Arc::clone(&next.after),
            self.limits,
        )
    }

    /// Merge patches authored from the same base when their semantic cell
    /// selections are disjoint.
    ///
    /// # Errors
    ///
    /// Returns an error for different bases, any overlapping cell selection,
    /// or failed bounded readback.
    pub fn compose_disjoint(&self, other: &Self) -> Result<Self> {
        if self.before != other.before {
            return Err(Error::UnsupportedFeature(
                "disjoint cell patches do not share an exact base".to_string(),
            ));
        }
        let left = self
            .changes
            .iter()
            .map(Change::reference)
            .collect::<BTreeSet<_>>();
        if let Some(reference) = other
            .changes
            .iter()
            .map(Change::reference)
            .find(|reference| left.contains(reference))
        {
            return Err(Error::UnsupportedFeature(format!(
                "disjoint cell patches overlap at ({}, {})",
                reference.row(),
                reference.column()
            )));
        }
        merge_patch_changes(self, other).and_then(|outcome| {
            outcome.patch.ok_or_else(|| {
                Error::InvalidFormat("disjoint patch merge unexpectedly conflicted".to_string())
            })
        })
    }

    /// Perform a three-way semantic merge against the patches' exact common
    /// base. Identical overlapping destinations coalesce; divergent ones are
    /// returned as typed conflicts without publishing a partial patch.
    ///
    /// # Errors
    ///
    /// Returns an error when the patches do not share the same exact base or
    /// reconstruction fails.
    pub fn merge_three_way(&self, other: &Self) -> Result<MergeOutcome> {
        if self.before != other.before {
            return Err(Error::UnsupportedFeature(
                "three-way cell patches do not share an exact base".to_string(),
            ));
        }
        merge_patch_changes(self, other)
    }

    /// Encode a deterministic, versioned exact-source patch envelope.
    ///
    /// Semantic changes are reconstructed and verified on decode, avoiding a
    /// second mutable representation while retaining unknown record bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the patch exceeds the finite transfer policy.
    pub fn to_bytes(&self, limits: TransferLimits) -> Result<Vec<u8>> {
        limits.validate()?;
        if self.changes.len() > limits.changes {
            return Err(Error::InvalidLength {
                expected: limits.changes,
                found: self.changes.len(),
            });
        }
        let header_len = PATCH_MAGIC.len().saturating_add(6 * 8);
        let total = header_len
            .checked_add(self.before.len())
            .and_then(|size| size.checked_add(self.after.len()))
            .ok_or(Error::CapacityOverflow {
                resource: "durable cell patch bytes",
            })?;
        if total > limits.bytes {
            return Err(Error::InvalidLength {
                expected: limits.bytes,
                found: total,
            });
        }
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(total)
            .map_err(|source| Error::Allocation {
                resource: "durable cell patch bytes",
                source,
            })?;
        bytes.extend_from_slice(PATCH_MAGIC);
        for value in [
            self.before.len(),
            self.after.len(),
            self.limits.source_bytes,
            self.limits.cells,
            self.limits.raw.payload(),
            self.limits.raw.string_units(),
        ] {
            let wire = u64::try_from(value).map_err(|_| Error::CapacityOverflow {
                resource: "durable cell patch length",
            })?;
            bytes.extend_from_slice(&wire.to_le_bytes());
        }
        bytes.extend_from_slice(&self.before);
        bytes.extend_from_slice(&self.after);
        Ok(bytes)
    }

    /// Decode and fully validate a deterministic patch envelope.
    ///
    /// # Errors
    ///
    /// Returns an error for an unknown version, truncation, trailing bytes,
    /// exceeded limits, or invalid before/after worksheet images.
    pub fn from_bytes(data: &[u8], limits: TransferLimits) -> Result<Self> {
        limits.validate()?;
        if data.len() > limits.bytes {
            return Err(Error::InvalidLength {
                expected: limits.bytes,
                found: data.len(),
            });
        }
        let header_len = PATCH_MAGIC.len().saturating_add(6 * 8);
        require_at_least(data, header_len)?;
        if data.get(..PATCH_MAGIC.len()) != Some(PATCH_MAGIC.as_slice()) {
            return Err(Error::InvalidFormat(
                "unknown durable cell patch version".to_string(),
            ));
        }
        let mut offset = PATCH_MAGIC.len();
        let mut lengths = [0usize; 6];
        for length in &mut lengths {
            let value = read_u64_at(data, offset)?;
            *length = usize::try_from(value).map_err(|_| Error::CapacityOverflow {
                resource: "durable cell patch length",
            })?;
            offset = offset.checked_add(8).ok_or(Error::CapacityOverflow {
                resource: "durable cell patch header",
            })?;
        }
        let end_before = offset
            .checked_add(lengths[0])
            .ok_or(Error::CapacityOverflow {
                resource: "durable cell patch before image",
            })?;
        let end_after = end_before
            .checked_add(lengths[1])
            .ok_or(Error::CapacityOverflow {
                resource: "durable cell patch after image",
            })?;
        if end_after != data.len() {
            return Err(Error::InvalidLength {
                expected: end_after,
                found: data.len(),
            });
        }
        let worksheet_limits = Limits::new(lengths[2], lengths[3], lengths[4], lengths[5]);
        worksheet_limits.validate()?;
        let before = Arc::from(data[offset..end_before].to_vec());
        let after = Arc::from(data[end_before..end_after].to_vec());
        let patch = patch_from_images(before, after, worksheet_limits)?;
        if patch.changes.len() > limits.changes {
            return Err(Error::InvalidLength {
                expected: limits.changes,
                found: patch.changes.len(),
            });
        }
        Ok(patch)
    }
}

/// One divergent cell selected by a three-way merge.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MergeConflict {
    reference: Reference,
    left: Option<StoredCell>,
    right: Option<StoredCell>,
}

impl MergeConflict {
    /// Conflicting cell coordinate.
    #[must_use]
    pub const fn reference(&self) -> Reference {
        self.reference
    }

    /// Left destination state.
    #[must_use]
    pub const fn left(&self) -> Option<&StoredCell> {
        self.left.as_ref()
    }

    /// Right destination state.
    #[must_use]
    pub const fn right(&self) -> Option<&StoredCell> {
        self.right.as_ref()
    }
}

/// Atomic outcome of a three-way semantic merge.
#[derive(Debug, Clone)]
pub struct MergeOutcome {
    patch: Option<Patch>,
    conflicts: Vec<MergeConflict>,
}

impl MergeOutcome {
    /// Merged patch, present exactly when there are no conflicts.
    #[must_use]
    pub const fn patch(&self) -> Option<&Patch> {
        self.patch.as_ref()
    }

    /// Divergent selections in row/column order.
    #[must_use]
    pub fn conflicts(&self) -> &[MergeConflict] {
        &self.conflicts
    }

    /// Split the atomic outcome.
    #[must_use]
    pub fn into_parts(self) -> (Option<Patch>, Vec<MergeConflict>) {
        (self.patch, self.conflicts)
    }
}

/// Bounded exact-source undo/redo history.
#[derive(Debug, Clone)]
pub struct History {
    entries: Vec<Patch>,
    cursor: usize,
    retained_bytes: usize,
    limits: TransferLimits,
}

impl History {
    /// Construct an empty bounded history.
    ///
    /// # Errors
    ///
    /// Returns an error when any selected bound is zero.
    pub fn new(limits: TransferLimits) -> Result<Self> {
        limits.validate()?;
        Ok(Self {
            entries: Vec::new(),
            cursor: 0,
            retained_bytes: 0,
            limits,
        })
    }

    /// Number of retained entries.
    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Whether no history entries are retained.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Retain a committed patch, discarding the redo branch and then evicting
    /// oldest entries until both configured history bounds hold.
    ///
    /// # Errors
    ///
    /// Returns an error when one patch alone exceeds the byte bound or does
    /// not continue the current exact-source history tip.
    pub fn push(&mut self, patch: Patch) -> Result<()> {
        let patch_bytes = patch.retained_bytes()?;
        if patch_bytes > self.limits.history_bytes {
            return Err(Error::InvalidLength {
                expected: self.limits.history_bytes,
                found: patch_bytes,
            });
        }
        let expected_source = if self.cursor == 0 {
            self.entries.first().map(|entry| &entry.before)
        } else {
            self.entries
                .get(self.cursor.saturating_sub(1))
                .map(|entry| &entry.after)
        };
        if expected_source.is_some_and(|source| source != &patch.before) {
            return Err(Error::UnsupportedFeature(
                "cell history patch does not continue its exact source tip".to_string(),
            ));
        }
        while self.entries.len() > self.cursor {
            if let Some(removed) = self.entries.pop() {
                self.retained_bytes = self
                    .retained_bytes
                    .saturating_sub(removed.retained_bytes()?);
            }
        }
        self.retained_bytes =
            self.retained_bytes
                .checked_add(patch_bytes)
                .ok_or(Error::CapacityOverflow {
                    resource: "cell history bytes",
                })?;
        self.entries.push(patch);
        self.cursor = self.entries.len();
        while self.entries.len() > self.limits.history_entries
            || self.retained_bytes > self.limits.history_bytes
        {
            let removed = self.entries.remove(0);
            self.retained_bytes = self
                .retained_bytes
                .saturating_sub(removed.retained_bytes()?);
            self.cursor = self.cursor.saturating_sub(1);
        }
        Ok(())
    }

    /// Apply one exact inverse at the current history tip.
    ///
    /// # Errors
    ///
    /// Returns an error for an empty undo side or a stale source.
    pub fn undo(&mut self, source: &[u8]) -> Result<Vec<u8>> {
        let index = self.cursor.checked_sub(1).ok_or_else(|| {
            Error::UnsupportedFeature("cell history has no undo entry".to_string())
        })?;
        let result = self.entries[index].inverse().apply(source)?;
        self.cursor = index;
        Ok(result)
    }

    /// Reapply one entry on the current redo side.
    ///
    /// # Errors
    ///
    /// Returns an error for an empty redo side or a stale source.
    pub fn redo(&mut self, source: &[u8]) -> Result<Vec<u8>> {
        let patch = self.entries.get(self.cursor).ok_or_else(|| {
            Error::UnsupportedFeature("cell history has no redo entry".to_string())
        })?;
        let result = patch.apply(source)?;
        self.cursor = self.cursor.saturating_add(1);
        Ok(result)
    }
}

#[derive(Debug, Clone)]
struct Entry {
    original: StoredCell,
    cell: StoredCell,
    record_offset: usize,
}

impl Entry {
    fn changed(&self) -> bool {
        self.original != self.cell
    }
}

/// Read one complete worksheet stream with safe default limits.
#[cfg(test)]
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
                let entry = parse_entry(&record, row, limits)?;
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

fn parse_entry(record: &crate::raw::Record<'_>, row: u32, limits: Limits) -> Result<Entry> {
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
    let (value, formula) = parse_value(record.kind(), payload, limits.raw)?;
    let cell = StoredCell {
        reference,
        style,
        show_phonetic: payload[7] & 1 != 0,
        value,
        formula,
    };
    Ok(Entry {
        original: cell.clone(),
        cell,
        record_offset: record.offset(),
    })
}

fn parse_value(
    kind_value: crate::raw::Kind,
    payload: &[u8],
    limits: RawLimits,
) -> Result<(Value, Option<CellFormula>)> {
    match kind_value {
        kind::CELL_BLANK => {
            require_exact(payload, 8)?;
            Ok((Value::Blank, None))
        },
        kind::CELL_RK => {
            require_exact(payload, 12)?;
            let mut cursor = Cursor::new(&payload[8..], "BrtCellRk");
            Ok((Value::RkNumber(cursor.read_rk()?), None))
        },
        kind::CELL_ERROR => {
            require_exact(payload, 9)?;
            Ok((Value::Error(CellError::from_code(payload[8])?), None))
        },
        kind::CELL_BOOL => {
            require_exact(payload, 9)?;
            Ok((Value::Boolean(parse_bool(payload[8])?), None))
        },
        kind::CELL_REAL => {
            require_exact(payload, 16)?;
            let value = binary::read_f64_le_at(payload, 8)?;
            validate_xnum(value, "BrtCellReal")?;
            Ok((Value::Number(value), None))
        },
        kind::CELL_ST => parse_string_value(payload, limits, false),
        kind::CELL_ISST => {
            require_exact(payload, 12)?;
            Ok((
                Value::SharedStringIndex(binary::read_u32_le_at(payload, 8)?),
                None,
            ))
        },
        kind::CELL_R_STRING => {
            require_at_least(payload, 13)?;
            Ok((Value::RichString(SharedString::parse(&payload[8..])?), None))
        },
        kind::FMLA_STRING => parse_string_value(payload, limits, true),
        kind::FMLA_NUM => {
            require_at_least(payload, 26)?;
            let value = binary::read_f64_le_at(payload, 8)?;
            validate_xnum(value, "BrtFmlaNum cache")?;
            Ok((
                Value::FormulaNumberCache(value),
                Some(parse_formula(payload, 16)?),
            ))
        },
        kind::FMLA_BOOL => {
            require_at_least(payload, 19)?;
            Ok((
                Value::FormulaBooleanCache(parse_bool(payload[8])?),
                Some(parse_formula(payload, 9)?),
            ))
        },
        kind::FMLA_ERROR => {
            require_at_least(payload, 19)?;
            Ok((
                Value::FormulaErrorCache(CellError::from_code(payload[8])?),
                Some(parse_formula(payload, 9)?),
            ))
        },
        _ => Err(Error::InvalidRecordType(kind_value.get())),
    }
}

fn parse_string_value(
    payload: &[u8],
    limits: RawLimits,
    formula: bool,
) -> Result<(Value, Option<CellFormula>)> {
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
    if formula && cursor.remaining() < 10 {
        return Err(Error::InvalidFormat(
            "BrtFmlaString is missing its formula flags and token stream".to_string(),
        ));
    }
    let parsed_formula = if formula {
        let flags_offset = 12usize
            .checked_add(units.checked_mul(2).ok_or(Error::CapacityOverflow {
                resource: "cell string bytes",
            })?)
            .ok_or(Error::CapacityOverflow {
                resource: "formula string offset",
            })?;
        Some(parse_formula(payload, flags_offset)?)
    } else {
        None
    };
    Ok((
        if formula {
            Value::FormulaStringCache(value)
        } else {
            Value::InlineString(value)
        },
        parsed_formula,
    ))
}

fn parse_formula(payload: &[u8], flags_offset: usize) -> Result<CellFormula> {
    let formula_offset = flags_offset.checked_add(2).ok_or(Error::CapacityOverflow {
        resource: "cell formula offset",
    })?;
    require_at_least(payload, formula_offset.saturating_add(8))?;
    let flags = binary::read_u16_le_at(payload, flags_offset)?;
    let (parsed, consumed) = crate::formula::ParsedFormula::parse(&payload[formula_offset..])
        .map_err(|error| Error::InvalidFormula(error.to_string()))?;
    if formula_offset.saturating_add(consumed) != payload.len() {
        return Err(Error::InvalidFormula(format!(
            "formula record has {} trailing bytes",
            payload
                .len()
                .saturating_sub(formula_offset.saturating_add(consumed))
        )));
    }
    CellFormula::from_source(flags, parsed.rgce, parsed.rgcb)
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
            | kind::CELL_R_STRING
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
        kind::CELL_R_STRING => "BrtCellRString",
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
            | (Value::RichString(_), Value::RichString(_))
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
        Value::RichString(string) => {
            let _ = string.encode()?;
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

fn validate_value_with_limits(value: &Value, limits: Limits) -> Result<()> {
    validate_value(value)?;
    if let Value::InlineString(string) | Value::FormulaStringCache(string) = value {
        let units = string.encode_utf16().count();
        if units > limits.raw.string_units() {
            return Err(Error::InvalidLength {
                expected: limits.raw.string_units(),
                found: units,
            });
        }
    }
    Ok(())
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

#[allow(
    clippy::cognitive_complexity,
    clippy::too_many_lines,
    reason = "one record-order state machine keeps structural BIFF12 rewriting and unknown-record copying auditable"
)]
fn rebuild_stream(edit: &Edit) -> Result<Vec<u8>> {
    let final_columns = final_row_columns(edit)?;
    let existing_rows = source_rows(&edit.source, edit.limits)?;
    let source_has_cells = source_contains_cells(&edit.source, edit.limits)?;
    let structural_rows = edit
        .inserted
        .keys()
        .chain(edit.removed.iter())
        .map(|reference| reference.row())
        .collect::<BTreeSet<_>>();
    let mut replacements = BTreeMap::new();
    for entry in &edit.entries {
        replacements.insert(entry.record_offset, entry);
    }
    let mut inserted_rows = BTreeMap::<u32, Vec<&StoredCell>>::new();
    for cell in edit.inserted.values() {
        inserted_rows
            .entry(cell.reference.row())
            .or_default()
            .push(cell);
    }

    let insertion_growth = edit.inserted.values().try_fold(0usize, |total, cell| {
        encoded_cell(cell).and_then(|record| {
            total
                .checked_add(record.len())
                .ok_or(Error::CapacityOverflow {
                    resource: "structural cell insertion bytes",
                })
        })
    })?;
    let replacement_growth = edit
        .entries
        .iter()
        .filter(|entry| entry.changed() && !edit.removed.contains(&entry.cell.reference))
        .try_fold(0usize, |total, entry| {
            let new_len = encoded_cell(&entry.cell)?.len();
            let old_len = record_total_len(&edit.source, entry.record_offset, edit.limits.raw)?;
            total
                .checked_add(new_len.saturating_sub(old_len))
                .ok_or(Error::CapacityOverflow {
                    resource: "structural cell replacement bytes",
                })
        })?;
    let growth =
        insertion_growth
            .checked_add(replacement_growth)
            .ok_or(Error::CapacityOverflow {
                resource: "structural worksheet growth",
            })?;
    let capacity = edit
        .source
        .len()
        .checked_add(growth)
        .ok_or(Error::CapacityOverflow {
            resource: "structural worksheet bytes",
        })?;
    if capacity > edit.limits.source_bytes {
        return Err(Error::InvalidLength {
            expected: edit.limits.source_bytes,
            found: capacity,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve(capacity)
        .map_err(|source| Error::Allocation {
            resource: "structural worksheet bytes",
            source,
        })?;

    let mut current_row = None;
    let mut in_sheet_data = false;
    let mut frt_depth = 0usize;
    let mut emitted_absent_rows = BTreeSet::new();
    let mut pending_insert_index = 0usize;

    for item in Records::with_limits(&edit.source, edit.limits.raw) {
        let record = item?;
        match record.kind() {
            kind::BEGIN_SHEET_DATA => {
                in_sheet_data = true;
                copy_record(&edit.source, &record, edit.limits.raw, &mut output)?;
            },
            kind::FRT_BEGIN if in_sheet_data => {
                frt_depth = frt_depth.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "worksheet FRT nesting depth",
                })?;
                copy_record(&edit.source, &record, edit.limits.raw, &mut output)?;
            },
            kind::FRT_END if in_sheet_data => {
                frt_depth = frt_depth.saturating_sub(1);
                copy_record(&edit.source, &record, edit.limits.raw, &mut output)?;
            },
            kind::ROW_HDR if in_sheet_data && frt_depth == 0 => {
                if let Some(row) = current_row {
                    emit_remaining_insertions(
                        &mut output,
                        inserted_rows.get(&row),
                        &mut pending_insert_index,
                    )?;
                }
                let row = binary::read_u32_le_at(record.payload(), 0)?;
                emit_absent_rows_before(
                    &mut output,
                    row,
                    &inserted_rows,
                    &existing_rows,
                    &mut emitted_absent_rows,
                    &final_columns,
                )?;
                if structural_rows.contains(&row) {
                    let payload = encode_row_header(record.payload(), final_columns.get(&row))?;
                    Writer::with_limits(&mut output, edit.limits.raw)
                        .write_record(kind::ROW_HDR, &payload)?;
                } else {
                    copy_record(&edit.source, &record, edit.limits.raw, &mut output)?;
                }
                current_row = Some(row);
                pending_insert_index = 0;
            },
            kind::END_SHEET_DATA if in_sheet_data => {
                if let Some(row) = current_row {
                    emit_remaining_insertions(
                        &mut output,
                        inserted_rows.get(&row),
                        &mut pending_insert_index,
                    )?;
                }
                emit_absent_rows_before(
                    &mut output,
                    MAX_ROW.saturating_add(1),
                    &inserted_rows,
                    &existing_rows,
                    &mut emitted_absent_rows,
                    &final_columns,
                )?;
                copy_record(&edit.source, &record, edit.limits.raw, &mut output)?;
                in_sheet_data = false;
                current_row = None;
            },
            cell_kind if in_sheet_data && frt_depth == 0 && is_any_cell(cell_kind) => {
                let row = current_row.ok_or_else(|| {
                    Error::InvalidFormat("cell record appears before BrtRowHdr".to_string())
                })?;
                let column = binary::read_u32_le_at(record.payload(), 0)?;
                emit_insertions_before(
                    &mut output,
                    inserted_rows.get(&row),
                    &mut pending_insert_index,
                    column,
                )?;
                if let Some(entry) = replacements.get(&record.offset()) {
                    if !edit.removed.contains(&entry.cell.reference) {
                        if entry.changed() {
                            output.extend_from_slice(&encoded_cell(&entry.cell)?);
                        } else {
                            copy_record(&edit.source, &record, edit.limits.raw, &mut output)?;
                        }
                    }
                } else {
                    copy_record(&edit.source, &record, edit.limits.raw, &mut output)?;
                }
            },
            kind::WS_DIM if !edit.inserted.is_empty() => {
                let payload = expanded_dimensions(
                    record.payload(),
                    edit.inserted.values(),
                    source_has_cells,
                )?;
                Writer::with_limits(&mut output, edit.limits.raw)
                    .write_record(kind::WS_DIM, &payload)?;
            },
            _ => copy_record(&edit.source, &record, edit.limits.raw, &mut output)?,
        }
    }
    if output.len() > edit.limits.source_bytes {
        return Err(Error::InvalidLength {
            expected: edit.limits.source_bytes,
            found: output.len(),
        });
    }
    Ok(output)
}

fn final_row_columns(edit: &Edit) -> Result<BTreeMap<u32, BTreeSet<u32>>> {
    let mut rows = BTreeMap::<u32, BTreeSet<u32>>::new();
    let mut current_row = None;
    let mut in_sheet_data = false;
    let mut frt_depth = 0usize;
    let removed_offsets = edit
        .entries
        .iter()
        .filter(|entry| edit.removed.contains(&entry.cell.reference))
        .map(|entry| entry.record_offset)
        .collect::<BTreeSet<_>>();
    for item in Records::with_limits(&edit.source, edit.limits.raw) {
        let record = item?;
        match record.kind() {
            kind::BEGIN_SHEET_DATA => in_sheet_data = true,
            kind::END_SHEET_DATA => in_sheet_data = false,
            kind::FRT_BEGIN if in_sheet_data => {
                frt_depth = frt_depth.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "worksheet FRT nesting depth",
                })?;
            },
            kind::FRT_END if in_sheet_data => frt_depth = frt_depth.saturating_sub(1),
            kind::ROW_HDR if in_sheet_data && frt_depth == 0 => {
                let row = binary::read_u32_le_at(record.payload(), 0)?;
                current_row = Some(row);
                rows.entry(row).or_default();
            },
            cell_kind
                if in_sheet_data
                    && frt_depth == 0
                    && is_any_cell(cell_kind)
                    && !removed_offsets.contains(&record.offset()) =>
            {
                let row = current_row.ok_or_else(|| {
                    Error::InvalidFormat("cell record appears before BrtRowHdr".to_string())
                })?;
                rows.entry(row)
                    .or_default()
                    .insert(binary::read_u32_le_at(record.payload(), 0)?);
            },
            _ => {},
        }
    }
    for cell in edit.inserted.values() {
        rows.entry(cell.reference.row())
            .or_default()
            .insert(cell.reference.column());
    }
    Ok(rows)
}

fn source_rows(source: &[u8], limits: Limits) -> Result<BTreeSet<u32>> {
    let mut rows = BTreeSet::new();
    let mut in_sheet_data = false;
    let mut frt_depth = 0usize;
    for item in Records::with_limits(source, limits.raw) {
        let record = item?;
        match record.kind() {
            kind::BEGIN_SHEET_DATA => in_sheet_data = true,
            kind::END_SHEET_DATA => in_sheet_data = false,
            kind::FRT_BEGIN if in_sheet_data => {
                frt_depth = frt_depth.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "worksheet FRT nesting depth",
                })?;
            },
            kind::FRT_END if in_sheet_data => frt_depth = frt_depth.saturating_sub(1),
            kind::ROW_HDR if in_sheet_data && frt_depth == 0 => {
                rows.insert(binary::read_u32_le_at(record.payload(), 0)?);
            },
            _ => {},
        }
    }
    Ok(rows)
}

fn source_contains_cells(source: &[u8], limits: Limits) -> Result<bool> {
    let mut in_sheet_data = false;
    let mut frt_depth = 0usize;
    for item in Records::with_limits(source, limits.raw) {
        let record = item?;
        match record.kind() {
            kind::BEGIN_SHEET_DATA => in_sheet_data = true,
            kind::END_SHEET_DATA => in_sheet_data = false,
            kind::FRT_BEGIN if in_sheet_data => {
                frt_depth = frt_depth.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "worksheet FRT nesting depth",
                })?;
            },
            kind::FRT_END if in_sheet_data => frt_depth = frt_depth.saturating_sub(1),
            cell_kind if in_sheet_data && frt_depth == 0 && is_any_cell(cell_kind) => {
                return Ok(true);
            },
            _ => {},
        }
    }
    Ok(false)
}

fn encode_row_header(source: &[u8], columns: Option<&BTreeSet<u32>>) -> Result<Vec<u8>> {
    require_at_least(source, 17)?;
    let mut payload = source[..13].to_vec();
    append_spans(&mut payload, columns)?;
    Ok(payload)
}

fn synthetic_row_header(row: u32, columns: Option<&BTreeSet<u32>>) -> Result<Vec<u8>> {
    let mut payload = Vec::new();
    Writer::new(&mut payload).write_u32(row)?;
    Writer::new(&mut payload).write_u32(0)?;
    Writer::new(&mut payload).write_u16(300)?;
    payload.extend_from_slice(&[0, 0, 0]);
    append_spans(&mut payload, columns)?;
    Ok(payload)
}

fn append_spans(payload: &mut Vec<u8>, columns: Option<&BTreeSet<u32>>) -> Result<()> {
    let mut spans = Vec::<(u32, u32)>::new();
    if let Some(columns) = columns {
        for column in columns {
            let segment = column / 1024;
            if let Some((_, last)) = spans.last_mut()
                && *last / 1024 == segment
            {
                *last = *column;
                continue;
            }
            spans.push((*column, *column));
        }
    }
    let count = u32::try_from(spans.len()).map_err(|_| Error::CapacityOverflow {
        resource: "BrtRowHdr column spans",
    })?;
    payload.extend_from_slice(&count.to_le_bytes());
    for (first, last) in spans {
        payload.extend_from_slice(&first.to_le_bytes());
        payload.extend_from_slice(&last.to_le_bytes());
    }
    Ok(())
}

fn emit_absent_rows_before(
    output: &mut Vec<u8>,
    upper_bound: u32,
    inserted_rows: &BTreeMap<u32, Vec<&StoredCell>>,
    existing_rows: &BTreeSet<u32>,
    emitted: &mut BTreeSet<u32>,
    final_columns: &BTreeMap<u32, BTreeSet<u32>>,
) -> Result<()> {
    for (row, cells) in inserted_rows.range(..upper_bound) {
        if existing_rows.contains(row) || !emitted.insert(*row) {
            continue;
        }
        let payload = synthetic_row_header(*row, final_columns.get(row))?;
        Writer::new(&mut *output).write_record(kind::ROW_HDR, &payload)?;
        for cell in cells {
            output.extend_from_slice(&encoded_cell(cell)?);
        }
    }
    Ok(())
}

fn emit_insertions_before(
    output: &mut Vec<u8>,
    cells: Option<&Vec<&StoredCell>>,
    index: &mut usize,
    upper_column: u32,
) -> Result<()> {
    let Some(cells) = cells else {
        return Ok(());
    };
    while let Some(cell) = cells.get(*index) {
        if cell.reference.column() >= upper_column {
            break;
        }
        output.extend_from_slice(&encoded_cell(cell)?);
        *index = index.saturating_add(1);
    }
    Ok(())
}

fn emit_remaining_insertions(
    output: &mut Vec<u8>,
    cells: Option<&Vec<&StoredCell>>,
    index: &mut usize,
) -> Result<()> {
    let Some(cells) = cells else {
        return Ok(());
    };
    while let Some(cell) = cells.get(*index) {
        output.extend_from_slice(&encoded_cell(cell)?);
        *index = index.saturating_add(1);
    }
    Ok(())
}

fn encoded_cell(cell: &StoredCell) -> Result<Vec<u8>> {
    let mut payload = Vec::new();
    Writer::new(&mut payload).write_u32(cell.reference.column())?;
    let style = cell.style.get().to_le_bytes();
    payload.extend_from_slice(&style[..3]);
    payload.push(u8::from(cell.show_phonetic));
    encode_value_into(&mut payload, &cell.value)?;
    match (&cell.formula, cell.value.is_formula_cache()) {
        (Some(formula), true) => encode_formula_into(&mut payload, formula)?,
        (None, false) => {},
        (Some(_), false) | (None, true) => {
            return Err(Error::InvalidFormula(
                "cell formula presence does not match its cached-result family".to_string(),
            ));
        },
    }
    let mut record = Vec::new();
    Writer::new(&mut record).write_record(value_kind(&cell.value), &payload)?;
    Ok(record)
}

fn encode_value_into(output: &mut Vec<u8>, value: &Value) -> Result<()> {
    match value {
        Value::Blank => {},
        Value::RkNumber(number) => output.extend_from_slice(&encode_rk(*number)?),
        Value::Error(error) | Value::FormulaErrorCache(error) => output.push(error.code()),
        Value::Boolean(boolean) | Value::FormulaBooleanCache(boolean) => {
            output.push(u8::from(*boolean));
        },
        Value::Number(number) | Value::FormulaNumberCache(number) => {
            validate_xnum(*number, "numeric cell")?;
            output.extend_from_slice(&number.to_le_bytes());
        },
        Value::InlineString(string) | Value::FormulaStringCache(string) => {
            Writer::new(output).write_wide_string(string)?;
        },
        Value::SharedStringIndex(index) => output.extend_from_slice(&index.to_le_bytes()),
        Value::RichString(string) => output.extend_from_slice(&string.encode()?),
    }
    Ok(())
}

fn encode_formula_into(output: &mut Vec<u8>, formula: &CellFormula) -> Result<()> {
    output.extend_from_slice(&formula.flags.to_le_bytes());
    let token_len = u32::try_from(formula.tokens.len()).map_err(|_| Error::CapacityOverflow {
        resource: "cell formula token length",
    })?;
    let ancillary_len =
        u32::try_from(formula.ancillary.len()).map_err(|_| Error::CapacityOverflow {
            resource: "cell formula ancillary length",
        })?;
    output.extend_from_slice(&token_len.to_le_bytes());
    output.extend_from_slice(&formula.tokens);
    output.extend_from_slice(&ancillary_len.to_le_bytes());
    output.extend_from_slice(&formula.ancillary);
    Ok(())
}

fn value_kind(value: &Value) -> crate::raw::Kind {
    match value {
        Value::Blank => kind::CELL_BLANK,
        Value::RkNumber(_) => kind::CELL_RK,
        Value::Error(_) => kind::CELL_ERROR,
        Value::Boolean(_) => kind::CELL_BOOL,
        Value::Number(_) => kind::CELL_REAL,
        Value::InlineString(_) => kind::CELL_ST,
        Value::SharedStringIndex(_) => kind::CELL_ISST,
        Value::RichString(_) => kind::CELL_R_STRING,
        Value::FormulaStringCache(_) => kind::FMLA_STRING,
        Value::FormulaNumberCache(_) => kind::FMLA_NUM,
        Value::FormulaBooleanCache(_) => kind::FMLA_BOOL,
        Value::FormulaErrorCache(_) => kind::FMLA_ERROR,
    }
}

fn expanded_dimensions<'a>(
    payload: &[u8],
    inserted: impl Iterator<Item = &'a StoredCell>,
    source_has_cells: bool,
) -> Result<Vec<u8>> {
    require_exact(payload, 16)?;
    let mut first_row = binary::read_u32_le_at(payload, 0)?;
    let mut last_row = binary::read_u32_le_at(payload, 4)?;
    let mut first_column = binary::read_u32_le_at(payload, 8)?;
    let mut last_column = binary::read_u32_le_at(payload, 12)?;
    let was_empty = !source_has_cells;
    let mut saw_insert = false;
    for cell in inserted {
        let reference = cell.reference;
        if was_empty && !saw_insert {
            first_row = reference.row();
            last_row = reference.row();
            first_column = reference.column();
            last_column = reference.column();
        } else {
            first_row = first_row.min(reference.row());
            last_row = last_row.max(reference.row());
            first_column = first_column.min(reference.column());
            last_column = last_column.max(reference.column());
        }
        saw_insert = true;
    }
    let mut result = Vec::with_capacity(16);
    result.extend_from_slice(&first_row.to_le_bytes());
    result.extend_from_slice(&last_row.to_le_bytes());
    result.extend_from_slice(&first_column.to_le_bytes());
    result.extend_from_slice(&last_column.to_le_bytes());
    Ok(result)
}

fn copy_record(
    source: &[u8],
    record: &crate::raw::Record<'_>,
    limits: RawLimits,
    output: &mut Vec<u8>,
) -> Result<()> {
    let record_source = source.get(record.offset()..).ok_or_else(|| {
        Error::InvalidFormat("record offset is outside worksheet source".to_string())
    })?;
    let (_, header_len) = Header::parse(record_source, limits)?;
    let end = record
        .offset()
        .checked_add(header_len)
        .and_then(|offset| offset.checked_add(record.len()))
        .ok_or(Error::CapacityOverflow {
            resource: "worksheet record range",
        })?;
    let bytes = source.get(record.offset()..end).ok_or_else(|| {
        Error::InvalidFormat("record range is outside worksheet source".to_string())
    })?;
    output.extend_from_slice(bytes);
    Ok(())
}

fn record_total_len(source: &[u8], offset: usize, limits: RawLimits) -> Result<usize> {
    let record_source = source.get(offset..).ok_or_else(|| {
        Error::InvalidFormat("record offset is outside worksheet source".to_string())
    })?;
    let (header, header_len) = Header::parse(record_source, limits)?;
    header_len
        .checked_add(header.len())
        .ok_or(Error::CapacityOverflow {
            resource: "worksheet record bytes",
        })
}

fn is_any_cell(kind_value: crate::raw::Kind) -> bool {
    is_supported_cell(kind_value)
}

fn any_cell_at(source: &[u8], limits: Limits, reference: Reference) -> Result<bool> {
    let mut row = None;
    let mut in_sheet_data = false;
    let mut frt_depth = 0usize;
    for item in Records::with_limits(source, limits.raw) {
        let record = item?;
        match record.kind() {
            kind::BEGIN_SHEET_DATA => in_sheet_data = true,
            kind::END_SHEET_DATA => in_sheet_data = false,
            kind::FRT_BEGIN if in_sheet_data => {
                frt_depth = frt_depth.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "worksheet FRT nesting depth",
                })?;
            },
            kind::FRT_END if in_sheet_data => frt_depth = frt_depth.saturating_sub(1),
            kind::ROW_HDR if in_sheet_data && frt_depth == 0 => {
                row = Some(binary::read_u32_le_at(record.payload(), 0)?);
            },
            cell_kind
                if in_sheet_data
                    && frt_depth == 0
                    && is_any_cell(cell_kind)
                    && row == Some(reference.row())
                    && binary::read_u32_le_at(record.payload(), 0)? == reference.column() =>
            {
                return Ok(true);
            },
            _ => {},
        }
    }
    Ok(false)
}

fn semantic_changes(
    entries: &[Entry],
    _inserted: &BTreeMap<Reference, StoredCell>,
    _removed: &BTreeSet<Reference>,
    after: &Snapshot,
) -> Result<Vec<Change>> {
    let mut before_by_reference = BTreeMap::<Reference, StoredCell>::new();
    for entry in entries {
        if before_by_reference
            .insert(entry.original.reference, entry.original.clone())
            .is_some()
        {
            return Err(Error::UnsupportedFeature(format!(
                "cell ({}, {}) has duplicate stored cell records",
                entry.original.reference.row(),
                entry.original.reference.column()
            )));
        }
    }
    let mut after_by_reference = BTreeMap::<Reference, StoredCell>::new();
    for cell in after.cells() {
        if after_by_reference
            .insert(cell.reference, cell.clone())
            .is_some()
        {
            return Err(Error::UnsupportedFeature(format!(
                "cell ({}, {}) has duplicate stored cell records",
                cell.reference.row(),
                cell.reference.column()
            )));
        }
    }
    let references = before_by_reference
        .keys()
        .chain(after_by_reference.keys())
        .copied()
        .collect::<BTreeSet<_>>();
    let mut changes = Vec::new();
    changes
        .try_reserve(references.len())
        .map_err(|source| Error::Allocation {
            resource: "semantic cell patch changes",
            source,
        })?;
    for reference in references {
        let before = before_by_reference.get(&reference);
        let destination = after_by_reference.get(&reference);
        if before != destination {
            changes.push(Change {
                reference,
                before: before.cloned(),
                after: destination.cloned(),
            });
        }
    }
    Ok(changes)
}

fn patch_from_images(before: Arc<[u8]>, after: Arc<[u8]>, limits: Limits) -> Result<Patch> {
    let before_snapshot = read_shared(Arc::clone(&before), limits)?;
    let after_snapshot = read_shared(Arc::clone(&after), limits)?;
    let changes = snapshot_changes(&before_snapshot, &after_snapshot)?;
    Ok(Patch {
        before,
        after,
        changes,
        limits,
    })
}

fn snapshot_changes(before: &Snapshot, after: &Snapshot) -> Result<Vec<Change>> {
    let mut before_by_reference = BTreeMap::<Reference, StoredCell>::new();
    for cell in before.cells() {
        if before_by_reference
            .insert(cell.reference, cell.clone())
            .is_some()
        {
            return Err(Error::UnsupportedFeature(format!(
                "cell ({}, {}) has duplicate stored cell records",
                cell.reference.row(),
                cell.reference.column()
            )));
        }
    }
    let mut after_by_reference = BTreeMap::<Reference, StoredCell>::new();
    for cell in after.cells() {
        if after_by_reference
            .insert(cell.reference, cell.clone())
            .is_some()
        {
            return Err(Error::UnsupportedFeature(format!(
                "cell ({}, {}) has duplicate stored cell records",
                cell.reference.row(),
                cell.reference.column()
            )));
        }
    }
    let references = before_by_reference
        .keys()
        .chain(after_by_reference.keys())
        .copied()
        .collect::<BTreeSet<_>>();
    let mut changes = Vec::new();
    changes
        .try_reserve(references.len())
        .map_err(|source| Error::Allocation {
            resource: "semantic cell patch changes",
            source,
        })?;
    for reference in references {
        let source = before_by_reference.get(&reference);
        let destination = after_by_reference.get(&reference);
        if source != destination {
            changes.push(Change {
                reference,
                before: source.cloned(),
                after: destination.cloned(),
            });
        }
    }
    Ok(changes)
}

fn merge_patch_changes(left: &Patch, right: &Patch) -> Result<MergeOutcome> {
    let left_changes = left
        .changes
        .iter()
        .map(|change| (change.reference, change))
        .collect::<BTreeMap<_, _>>();
    let right_changes = right
        .changes
        .iter()
        .map(|change| (change.reference, change))
        .collect::<BTreeMap<_, _>>();
    let references = left_changes
        .keys()
        .chain(right_changes.keys())
        .copied()
        .collect::<BTreeSet<_>>();
    let mut conflicts = Vec::new();
    let mut destinations = BTreeMap::<Reference, Option<StoredCell>>::new();
    for reference in references {
        match (left_changes.get(&reference), right_changes.get(&reference)) {
            (Some(left_change), Some(right_change)) if left_change.after != right_change.after => {
                conflicts.push(MergeConflict {
                    reference,
                    left: left_change.after.clone(),
                    right: right_change.after.clone(),
                });
            },
            (Some(left_change), Some(_) | None) => {
                destinations.insert(reference, left_change.after.clone());
            },
            (None, Some(right_change)) => {
                destinations.insert(reference, right_change.after.clone());
            },
            (None, None) => {},
        }
    }
    if !conflicts.is_empty() {
        return Ok(MergeOutcome {
            patch: None,
            conflicts,
        });
    }
    let base = read_shared(Arc::clone(&left.before), left.limits)?;
    let mut edit = base.edit();
    for (reference, destination) in destinations {
        match destination {
            Some(cell) => put_stored_cell(&mut edit, cell)?,
            None => edit.remove(reference)?,
        }
    }
    let commit = edit.commit()?;
    Ok(MergeOutcome {
        patch: Some(commit.patch),
        conflicts,
    })
}

fn put_stored_cell(edit: &mut Edit, cell: StoredCell) -> Result<()> {
    validate_value_with_limits(&cell.value, edit.limits)?;
    if cell.value.is_formula_cache() != cell.formula.is_some() {
        return Err(Error::InvalidFormula(
            "merged cell formula presence does not match its value family".to_string(),
        ));
    }
    if let Some(index) = unique_index(&edit.entries, cell.reference)? {
        edit.removed.remove(&cell.reference);
        edit.entries[index].cell = cell;
    } else {
        if any_cell_at(&edit.source, edit.limits, cell.reference)? {
            return Err(Error::UnsupportedFeature(format!(
                "merged cell ({}, {}) collides with an unsupported record family",
                cell.reference.row(),
                cell.reference.column()
            )));
        }
        edit.inserted.insert(cell.reference, cell);
    }
    Ok(())
}

fn read_u64_at(data: &[u8], offset: usize) -> Result<u64> {
    let end = offset.checked_add(8).ok_or(Error::CapacityOverflow {
        resource: "durable cell patch integer",
    })?;
    let bytes = data.get(offset..end).ok_or(Error::InvalidLength {
        expected: end,
        found: data.len(),
    })?;
    let array: [u8; 8] = bytes.try_into().map_err(|_| Error::InvalidLength {
        expected: 8,
        found: bytes.len(),
    })?;
    Ok(u64::from_le_bytes(array))
}

impl Patch {
    fn retained_bytes(&self) -> Result<usize> {
        self.before
            .len()
            .checked_add(self.after.len())
            .ok_or(Error::CapacityOverflow {
                resource: "cell history patch bytes",
            })
    }
}

fn encode_rk(value: f64) -> Result<Vec<u8>> {
    validate_finite(value, "BrtCellRk")?;
    let mut bytes = Vec::new();
    Writer::new(&mut bytes).write_rk(value)?;
    Ok(bytes)
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

    fn append_empty_formula(payload: &mut Vec<u8>) {
        payload.extend_from_slice(&0_u16.to_le_bytes());
        payload.extend_from_slice(&0_u32.to_le_bytes());
        payload.extend_from_slice(&0_u32.to_le_bytes());
    }

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

        let mut formula = vec![0; 16];
        formula[..4].copy_from_slice(&6_u32.to_le_bytes());
        formula[8..16].copy_from_slice(&7.5_f64.to_le_bytes());
        append_empty_formula(&mut formula);
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

        let mut formula_boolean = vec![0; 9];
        formula_boolean[..4].copy_from_slice(&9_u32.to_le_bytes());
        append_empty_formula(&mut formula_boolean);
        writer
            .write_record(kind::FMLA_BOOL, &formula_boolean)
            .expect("formula Boolean");

        let mut formula_error = vec![0; 9];
        formula_error[..4].copy_from_slice(&10_u32.to_le_bytes());
        formula_error[8] = CellError::NotAvailable.code();
        append_empty_formula(&mut formula_error);
        writer
            .write_record(kind::FMLA_ERROR, &formula_error)
            .expect("formula error");

        let mut formula_string = vec![0; 12];
        formula_string[..4].copy_from_slice(&11_u32.to_le_bytes());
        formula_string[8..12].copy_from_slice(&2_u32.to_le_bytes());
        formula_string.extend_from_slice(&[b'X', 0, b'Y', 0]);
        append_empty_formula(&mut formula_string);
        writer
            .write_record(kind::FMLA_STRING, &formula_string)
            .expect("formula string");

        let rich_value = SharedString {
            text: "Rich".to_string(),
            runs: vec![crate::package::SharedStringRun {
                character_index: 0,
                font_id: 0,
            }],
            phonetic: None,
        };
        let mut rich = vec![0; 8];
        rich[..4].copy_from_slice(&12_u32.to_le_bytes());
        rich.extend_from_slice(&rich_value.encode().expect("rich string"));
        writer
            .write_record(kind::CELL_R_STRING, &rich)
            .expect("rich string cell");
        writer.write_record(kind::END_SHEET_DATA, &[]).expect("end");
        bytes
    }

    #[test]
    fn edits_all_bounded_scalar_families_and_round_trips_patch() {
        let before = stream();
        let snapshot = read(&before).expect("snapshot");
        assert_eq!(snapshot.cells().len(), 11);

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
        let rich_string_ref = Reference::new(3, 12).expect("rich string reference");
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
        edit.set_rich_string(
            rich_string_ref,
            SharedString {
                text: "Longer rich text".to_string(),
                runs: vec![crate::package::SharedStringRun {
                    character_index: 0,
                    font_id: 0,
                }],
                phonetic: None,
            },
        )
        .expect("set rich string");
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
            commit
                .snapshot()
                .cell(rich_string_ref)
                .expect("lookup")
                .expect("rich string")
                .value(),
            &Value::RichString(SharedString {
                text: "Longer rich text".to_string(),
                runs: vec![crate::package::SharedStringRun {
                    character_index: 0,
                    font_id: 0,
                }],
                phonetic: None,
            })
        );
        assert_eq!(
            commit.patch().inverse().apply(&after).expect("revert"),
            before
        );
    }

    #[test]
    fn supports_length_changes_and_refuses_family_changes_inexact_rk_and_stale_sources() {
        let snapshot = read(&stream()).expect("snapshot");
        let string_ref = Reference::new(3, 4).expect("string reference");
        let rk_ref = Reference::new(3, 5).expect("rk reference");
        let mut edit = snapshot.edit();
        assert!(edit.set_boolean(string_ref, true).is_err());
        edit.set_inline_string(string_ref, "a much longer string".to_string())
            .expect("length-changing string");
        assert!(edit.set_number(rk_ref, 1.0 / 3.0).is_err());

        let real_ref = Reference::new(3, 2).expect("real reference");
        edit.set_number(real_ref, 1.0).expect("set real");
        let commit = edit.commit().expect("commit");
        assert_eq!(
            commit
                .snapshot()
                .cell(string_ref)
                .expect("lookup")
                .expect("string")
                .value(),
            &Value::InlineString("a much longer string".to_string())
        );
        assert!(commit.patch().apply(b"stale").is_err());
    }

    #[test]
    fn structural_crud_formula_rewrite_and_dependency_metadata_round_trip() {
        let before = stream();
        let snapshot = read(&before).expect("snapshot");
        let removed = Reference::new(3, 3).expect("removed");
        let inserted_same_row = Reference::new(3, 13).expect("same row");
        let inserted_new_row = Reference::new(8, 1).expect("new row");
        let formula_ref = Reference::new(3, 6).expect("formula");
        let mut edit = snapshot.edit();
        edit.remove(removed).expect("remove");
        edit.insert(
            inserted_same_row,
            StyleIndex::new(0).expect("style"),
            Value::InlineString("created".to_string()),
        )
        .expect("insert same row");
        edit.insert(
            inserted_new_row,
            StyleIndex::new(0).expect("style"),
            Value::Number(42.0),
        )
        .expect("insert new row");
        let replacement_formula =
            CellFormula::new(0x0002, vec![0x1E, 0x02, 0x00], vec![]).expect("constant formula");
        edit.set_formula(formula_ref, replacement_formula.clone())
            .expect("formula tokens");
        let commit = edit.commit().expect("commit");
        assert!(commit.snapshot().cell(removed).expect("lookup").is_none());
        assert_eq!(
            commit
                .snapshot()
                .cell(inserted_new_row)
                .expect("lookup")
                .expect("inserted")
                .value(),
            &Value::Number(42.0)
        );
        assert_eq!(
            commit
                .snapshot()
                .cell(formula_ref)
                .expect("lookup")
                .expect("formula")
                .formula(),
            Some(&replacement_formula)
        );
        let after = commit.patch().apply(&before).expect("apply");
        assert_eq!(
            commit.patch().inverse().apply(&after).expect("inverse"),
            before
        );
        assert_eq!(commit.patch().changes().len(), 4);
    }

    #[test]
    fn durable_transfer_merge_conflicts_and_bounded_history_are_exact() {
        let before = stream();
        let base = read(&before).expect("base");
        let left_ref = Reference::new(3, 2).expect("left");
        let right_ref = Reference::new(3, 3).expect("right");

        let mut left_edit = base.edit();
        left_edit.set_number(left_ref, 10.0).expect("left edit");
        let left = left_edit.commit().expect("left commit").patch().clone();
        let encoded = left.to_bytes(TransferLimits::DEFAULT).expect("serialize");
        let decoded = Patch::from_bytes(&encoded, TransferLimits::DEFAULT).expect("decode");
        assert_eq!(decoded.apply(&before).expect("apply"), left.after());

        let mut right_edit = base.edit();
        right_edit.set_boolean(right_ref, true).expect("right edit");
        let right = right_edit.commit().expect("right commit").patch().clone();
        let merged = left.compose_disjoint(&right).expect("disjoint merge");
        let merged_bytes = merged.apply(&before).expect("merged apply");
        let merged_snapshot = read(&merged_bytes).expect("merged readback");
        assert_eq!(
            merged_snapshot
                .number(left_ref)
                .expect("left lookup")
                .expect("left")
                .value(),
            10.0
        );
        assert_eq!(
            merged_snapshot
                .cell(right_ref)
                .expect("right lookup")
                .expect("right")
                .value(),
            &Value::Boolean(true)
        );

        let mut conflict_edit = base.edit();
        conflict_edit.set_number(left_ref, 11.0).expect("conflict");
        let conflict = conflict_edit.commit().expect("commit").patch().clone();
        let outcome = left.merge_three_way(&conflict).expect("merge outcome");
        assert!(outcome.patch().is_none());
        assert_eq!(outcome.conflicts()[0].reference(), left_ref);

        let mut history = History::new(TransferLimits::new(
            encoded.len().saturating_mul(2),
            16,
            2,
            before
                .len()
                .saturating_add(left.after().len())
                .saturating_mul(2),
        ))
        .expect("history");
        history.push(left.clone()).expect("push");
        let after = left.apply(&before).expect("forward");
        let undone = history.undo(&after).expect("undo");
        assert_eq!(undone, before);
        assert_eq!(history.redo(&undone).expect("redo"), after);
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
