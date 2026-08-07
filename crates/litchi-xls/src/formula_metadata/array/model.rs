//! Semantic model for one BIFF8 `Array` record.

use std::sync::Arc;

use crate::{Error, Result};

use super::super::{Cell, Range};

const BIFF8_GRID_CELLS: usize = 65_536 * 256;
const MAX_FORMULA_SOURCE_BYTES: usize = 1_048_576;
const MAX_FORMULA_SCALARS: usize = 524_288;
const MAX_TOKEN_COUNT: usize = 1_800;
const MAX_OPERATOR_DEPTH: usize = 256;
const MAX_STRING_UTF16_UNITS: usize = 255;
const MAX_NESTING_DEPTH: usize = 8;
const MAX_OPERANDS: usize = 40;

/// Resource limits used while compiling, parsing, and staging array formulas.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_formula_bytes: usize,
    max_formula_scalars: usize,
    max_token_bytes: usize,
    max_tokens: usize,
    max_operator_depth: usize,
    max_string_utf16_units: usize,
    max_cells: usize,
    max_nesting_depth: usize,
    max_operands: usize,
    max_record_bytes: usize,
    max_extra_bytes: usize,
}

impl Limits {
    /// Construct checked compilation and ownership limits.
    pub fn new(
        max_formula_bytes: usize,
        max_formula_scalars: usize,
        max_token_bytes: usize,
        max_tokens: usize,
        max_operator_depth: usize,
        max_string_utf16_units: usize,
        max_cells: usize,
    ) -> Result<Self> {
        let limits = Self {
            max_formula_bytes,
            max_formula_scalars,
            max_token_bytes,
            max_tokens,
            max_operator_depth,
            max_string_utf16_units,
            max_cells,
            max_nesting_depth: MAX_NESTING_DEPTH,
            max_operands: MAX_OPERANDS,
            max_record_bytes: super::validation::MAX_RECORD_BYTES,
            max_extra_bytes: super::validation::MAX_RECORD_BYTES
                - super::validation::FIXED_BYTES
                - 1,
        };
        limits.check()?;
        Ok(limits)
    }

    pub const fn max_formula_bytes(self) -> usize {
        self.max_formula_bytes
    }

    pub const fn max_formula_scalars(self) -> usize {
        self.max_formula_scalars
    }

    pub const fn max_token_bytes(self) -> usize {
        self.max_token_bytes
    }

    pub const fn max_tokens(self) -> usize {
        self.max_tokens
    }

    pub const fn max_operator_depth(self) -> usize {
        self.max_operator_depth
    }

    pub const fn max_string_utf16_units(self) -> usize {
        self.max_string_utf16_units
    }

    pub const fn max_cells(self) -> usize {
        self.max_cells
    }

    /// Maximum normative function-call nesting in the parsed RPN tree.
    pub const fn max_nesting_depth(self) -> usize {
        self.max_nesting_depth
    }

    /// Maximum normative RPN operand pressure at the root expression.
    pub const fn max_operands(self) -> usize {
        self.max_operands
    }

    pub const fn max_record_bytes(self) -> usize {
        self.max_record_bytes
    }

    pub const fn max_extra_bytes(self) -> usize {
        self.max_extra_bytes
    }

    pub fn with_max_formula_bytes(mut self, maximum: usize) -> Result<Self> {
        self.max_formula_bytes = maximum;
        self.check()?;
        Ok(self)
    }

    pub fn with_max_formula_scalars(mut self, maximum: usize) -> Result<Self> {
        self.max_formula_scalars = maximum;
        self.check()?;
        Ok(self)
    }

    pub fn with_max_token_bytes(mut self, maximum: usize) -> Result<Self> {
        self.max_token_bytes = maximum;
        self.check()?;
        Ok(self)
    }

    pub fn with_max_tokens(mut self, maximum: usize) -> Result<Self> {
        self.max_tokens = maximum;
        self.check()?;
        Ok(self)
    }

    pub fn with_max_operator_depth(mut self, maximum: usize) -> Result<Self> {
        self.max_operator_depth = maximum;
        self.check()?;
        Ok(self)
    }

    pub fn with_max_string_utf16_units(mut self, maximum: usize) -> Result<Self> {
        self.max_string_utf16_units = maximum;
        self.check()?;
        Ok(self)
    }

    pub fn with_max_cells(mut self, maximum: usize) -> Result<Self> {
        self.max_cells = maximum;
        self.check()?;
        Ok(self)
    }

    pub fn with_max_nesting_depth(mut self, maximum: usize) -> Result<Self> {
        self.max_nesting_depth = maximum;
        self.check()?;
        Ok(self)
    }

    pub fn with_max_operands(mut self, maximum: usize) -> Result<Self> {
        self.max_operands = maximum;
        self.check()?;
        Ok(self)
    }

    pub fn with_max_record_bytes(mut self, maximum: usize) -> Result<Self> {
        self.max_record_bytes = maximum;
        self.check()?;
        Ok(self)
    }

    pub fn with_max_extra_bytes(mut self, maximum: usize) -> Result<Self> {
        self.max_extra_bytes = maximum;
        self.check()?;
        Ok(self)
    }

    fn check(self) -> Result<()> {
        if self.max_formula_bytes == 0
            || self.max_formula_scalars == 0
            || self.max_token_bytes == 0
            || self.max_tokens == 0
            || self.max_operator_depth == 0
            || self.max_string_utf16_units == 0
            || self.max_cells == 0
            || self.max_nesting_depth == 0
            || self.max_operands == 0
            || self.max_record_bytes < super::validation::FIXED_BYTES + 1
        {
            return Err(Error::InvalidData(
                "array-formula limits must be nonzero".to_string(),
            ));
        }
        if self.max_formula_bytes > MAX_FORMULA_SOURCE_BYTES
            || self.max_formula_scalars > MAX_FORMULA_SCALARS
            || self.max_token_bytes > super::validation::MAX_RGCE_BYTES
            || self.max_tokens > MAX_TOKEN_COUNT
            || self.max_operator_depth > MAX_OPERATOR_DEPTH
            || self.max_string_utf16_units > MAX_STRING_UTF16_UNITS
            || self.max_cells > BIFF8_GRID_CELLS
            || self.max_nesting_depth > MAX_NESTING_DEPTH
            || self.max_operands > MAX_OPERANDS
            || self.max_record_bytes > super::validation::MAX_RECORD_BYTES
            || self.max_extra_bytes > self.max_record_bytes - super::validation::FIXED_BYTES - 1
        {
            return Err(Error::InvalidData(
                "array-formula limits exceed BIFF8 structural bounds".to_string(),
            ));
        }
        Ok(())
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_formula_bytes: 65_536,
            max_formula_scalars: 32_768,
            max_token_bytes: super::validation::MAX_RGCE_BYTES,
            max_tokens: MAX_TOKEN_COUNT,
            max_operator_depth: MAX_OPERATOR_DEPTH,
            max_string_utf16_units: 255,
            max_cells: 1_048_576,
            max_nesting_depth: MAX_NESTING_DEPTH,
            max_operands: MAX_OPERANDS,
            max_record_bytes: super::validation::MAX_RECORD_BYTES,
            max_extra_bytes: super::validation::MAX_RECORD_BYTES
                - super::validation::FIXED_BYTES
                - 1,
        }
    }
}

/// Immutable owner of one complete BIFF8 `Array` payload.
///
/// Parsed owners retain ignored source bits and bytes exactly. Authored owners
/// can only be constructed through the canonical constructor, which writes
/// those fields as zero; there is no mutation path that silently canonicalizes
/// a parsed owner.
#[derive(Debug, PartialEq, Eq)]
pub struct Owner {
    range: Range,
    always_calculate: bool,
    reserved: u16,
    unused: [u8; 4],
    tokens: Arc<[u8]>,
    extra: Arc<[u8]>,
    max_cells: usize,
    authored: bool,
}

impl Owner {
    /// Construct a canonical authored owner from already compiled safe tokens.
    pub(crate) fn from_compiled(range: Range, tokens: Vec<u8>) -> Result<Self> {
        Self::from_compiled_with_limits(range, tokens, Limits::default())
    }

    /// Construct a canonical authored owner using caller-supplied limits.
    pub(crate) fn from_compiled_with_limits(
        range: Range,
        tokens: Vec<u8>,
        limits: Limits,
    ) -> Result<Self> {
        let owner = Self {
            range,
            always_calculate: true,
            reserved: 0,
            unused: [0; 4],
            tokens: Arc::from(tokens),
            extra: Arc::from([]),
            max_cells: limits.max_cells(),
            authored: true,
        };
        super::validation::validate(&owner, limits, true)?;
        Ok(owner)
    }

    pub const fn range(&self) -> Range {
        self.range
    }

    pub const fn anchor(&self) -> Cell {
        self.range.first()
    }

    pub const fn always_calculate(&self) -> bool {
        self.always_calculate
    }

    /// Ignored source bits retained for exact no-op serialization.
    pub const fn reserved(&self) -> u16 {
        self.reserved
    }

    /// Ignored source bytes retained for exact no-op serialization.
    pub const fn unused(&self) -> [u8; 4] {
        self.unused
    }

    /// `ArrayParsedFormula.rgce` bytes.
    pub fn tokens(&self) -> &[u8] {
        &self.tokens
    }

    /// `ArrayParsedFormula.rgcb` bytes.
    pub fn extra(&self) -> &[u8] {
        &self.extra
    }

    /// The retained cell-cardinality budget.
    pub const fn max_cells(&self) -> usize {
        self.max_cells
    }

    /// Number of cells in the complete rectangular owner.
    pub fn cell_count(&self) -> usize {
        let rows = usize::from(self.range.last().row() - self.range.first().row()) + 1;
        let cols = usize::from(self.range.last().col() - self.range.first().col()) + 1;
        rows * cols
    }

    pub fn contains(&self, cell: Cell) -> bool {
        self.range.contains(cell)
    }

    /// Iterate the complete owner rectangle lazily in row-major order.
    pub fn cells(&self) -> Cells {
        Cells {
            range: self.range,
            offset: 0,
            count: self.cell_count(),
        }
    }

    /// Exact standalone Formula-record `PtgExp` for every participant.
    pub(crate) const fn anchor_tokens(&self) -> [u8; 5] {
        let row = self.anchor().row().to_le_bytes();
        [0x01, row[0], row[1], self.anchor().col(), 0]
    }

    /// Validate without allocating.
    pub(crate) fn validate(&self) -> Result<()> {
        let limits = Limits::default().with_max_cells(self.max_cells)?;
        super::validation::validate(self, limits, self.authored)
    }

    /// Serialize one complete BIFF8 `Array` payload.
    pub(crate) fn to_payload(&self) -> Result<Vec<u8>> {
        super::codec::write_payload(self)
    }

    pub(super) fn from_wire(
        range: Range,
        always_calculate: bool,
        reserved: u16,
        unused: [u8; 4],
        tokens: Vec<u8>,
        extra: Vec<u8>,
        max_cells: usize,
    ) -> Self {
        Self {
            range,
            always_calculate,
            reserved,
            unused,
            tokens: Arc::from(tokens),
            extra: Arc::from(extra),
            max_cells,
            authored: false,
        }
    }
}

/// Lazy row-major iterator over an array formula's complete rectangle.
#[derive(Debug, Clone)]
pub struct Cells {
    range: Range,
    offset: usize,
    count: usize,
}

impl Iterator for Cells {
    type Item = Cell;

    fn next(&mut self) -> Option<Self::Item> {
        if self.offset == self.count {
            return None;
        }
        let columns = usize::from(self.range.last().col() - self.range.first().col()) + 1;
        let row = usize::from(self.range.first().row()) + self.offset / columns;
        let col = usize::from(self.range.first().col()) + self.offset % columns;
        self.offset += 1;
        Some(Cell::new(row as u16, col as u8))
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.count - self.offset;
        (remaining, Some(remaining))
    }
}

impl ExactSizeIterator for Cells {}
impl std::iter::FusedIterator for Cells {}
