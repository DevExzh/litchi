//! BIFF12 formula values and Parse Tree Generator (Ptg) codecs.
///
/// The wire shapes follow [MS-XLSB] section 2.5.98. Workbook relationship and
/// name resolution intentionally remains in the OOXML host adapter.
///
/// [MS-XLSB] section 2.2.2 defines formulas as an RPN sequence of Ptg
/// structures; this module preserves unknown bytes at the containing formula
/// boundary while validating every modeled token and ancillary payload.
use thiserror::Error;

mod function_table;
use function_table::BUILTIN_FUNCTIONS;

/// Maximum size of an XLSB cell formula token stream.
///
/// [MS-XLSB] 2.5.98.4 requires cce to be less than 16,385 bytes. Excel also
/// emits an empty stream for some cells, so zero remains representable.
pub const MAX_CELL_FORMULA_BYTES: usize = 16_384;

/// Error returned by the standalone formula codec.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// A formula or token violates the BIFF12 formula grammar.
    #[error("invalid formula: {0}")]
    InvalidFormula(String),
    /// A cell or range coordinate is outside the Excel grid.
    #[error("invalid cell reference: {0}")]
    InvalidCellReference(String),
    /// A fixed-width payload is shorter than the required structure.
    #[error("invalid length: expected {expected}, found {found}")]
    InvalidLength { expected: usize, found: usize },
    /// A formula feature is valid but not supported by this codec.
    #[error("unsupported formula feature: {0}")]
    UnsupportedFeature(String),
    /// Text or primitive binary decoding failed.
    #[error("formula encoding: {0}")]
    Encoding(String),
}

/// Result type for standalone formula codecs.
pub type Result<T> = std::result::Result<T, Error>;

fn read_bytes(data: &[u8], offset: usize, length: usize) -> Result<&[u8]> {
    let end = offset
        .checked_add(length)
        .ok_or_else(|| Error::InvalidFormula("formula binary offset overflow".to_string()))?;
    data.get(offset..end).ok_or(Error::InvalidLength {
        expected: end,
        found: data.len(),
    })
}

fn read_u16_le_at(data: &[u8], offset: usize) -> Result<u16> {
    let bytes = read_bytes(data, offset, 2)?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32_le_at(data: &[u8], offset: usize) -> Result<u32> {
    let bytes = read_bytes(data, offset, 4)?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn read_f64_le_at(data: &[u8], offset: usize) -> Result<f64> {
    let bytes = read_bytes(data, offset, 8)?;
    Ok(f64::from_le_bytes([
        bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7],
    ]))
}

fn column_index_to_name(mut column: u32) -> String {
    if column == 0 {
        return String::new();
    }
    let mut name = String::new();
    while column > 0 {
        column -= 1;
        name.insert(0, char::from(b'A' + (column % 26) as u8));
        column /= 26;
    }
    name
}

fn cell_reference(row: u32, column: u32) -> String {
    format!("{}{}", column_index_to_name(column + 1), row + 1)
}

fn parse_cell_reference(value: &str) -> Result<(u32, u32)> {
    let normalized = value.to_ascii_uppercase();
    let mut column = String::new();
    let mut row = String::new();
    let mut digit_seen = false;
    for character in normalized.chars() {
        if character.is_ascii_alphabetic() {
            if digit_seen {
                return Err(Error::InvalidCellReference(normalized));
            }
            column.push(character);
        } else if character.is_ascii_digit() {
            digit_seen = true;
            row.push(character);
        } else {
            return Err(Error::InvalidCellReference(normalized));
        }
    }
    if column.is_empty() || row.is_empty() {
        return Err(Error::InvalidCellReference(normalized));
    }
    let mut column_index = 0_u32;
    for character in column.bytes() {
        column_index = column_index
            .checked_mul(26)
            .and_then(|value| value.checked_add(u32::from(character - b'A' + 1)))
            .ok_or_else(|| Error::InvalidCellReference(normalized.clone()))?;
    }
    let row_index = row
        .parse::<u32>()
        .map_err(|_| Error::InvalidCellReference(normalized.clone()))?;
    if row_index == 0 || column_index == 0 {
        return Err(Error::InvalidCellReference(normalized));
    }
    Ok((row_index - 1, column_index - 1))
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaRange {
    pub row_first: u32,
    pub row_last: u32,
    pub col_first: u32,
    pub col_last: u32,
}

impl FormulaRange {
    pub fn new(row_first: u32, row_last: u32, col_first: u32, col_last: u32) -> Result<Self> {
        let range = Self {
            row_first,
            row_last,
            col_first,
            col_last,
        };
        range.validate()?;
        Ok(range)
    }

    pub fn parse_a1(value: &str) -> Result<Self> {
        let (first, last) = value.split_once(':').unwrap_or((value, value));
        let (row_first, col_first) = parse_cell_reference(first.trim())?;
        let (row_last, col_last) = parse_cell_reference(last.trim())?;
        Self::new(row_first, row_last, col_first, col_last)
    }

    pub fn parse_binary(data: &[u8]) -> Result<Self> {
        if data.len() < 16 {
            return Err(Error::InvalidLength {
                expected: 16,
                found: data.len(),
            });
        }
        Self::new(
            read_u32_le_at(data, 0)?,
            read_u32_le_at(data, 4)?,
            read_u32_le_at(data, 8)?,
            read_u32_le_at(data, 12)?,
        )
    }

    pub fn to_binary(self) -> [u8; 16] {
        let mut bytes = [0_u8; 16];
        bytes[0..4].copy_from_slice(&self.row_first.to_le_bytes());
        bytes[4..8].copy_from_slice(&self.row_last.to_le_bytes());
        bytes[8..12].copy_from_slice(&self.col_first.to_le_bytes());
        bytes[12..16].copy_from_slice(&self.col_last.to_le_bytes());
        bytes
    }

    pub fn contains(self, row: u32, col: u32) -> bool {
        (self.row_first..=self.row_last).contains(&row)
            && (self.col_first..=self.col_last).contains(&col)
    }

    pub fn top_left(self) -> (u32, u32) {
        (self.row_first, self.col_first)
    }

    pub fn to_a1(self) -> String {
        format!(
            "{}:{}",
            cell_reference(self.row_first, self.col_first),
            cell_reference(self.row_last, self.col_last)
        )
    }

    fn validate(self) -> Result<()> {
        if self.row_first > self.row_last
            || self.col_first > self.col_last
            || self.row_last >= 1_048_576
            || self.col_last >= 16_384
        {
            return Err(Error::InvalidCellReference(self.to_a1()));
        }
        Ok(())
    }
}

/// The binary representation of a cell formula (`CellParsedFormula`).
///
/// `rgce` contains the RPN token stream and `rgcb` contains ancillary data for
/// tokens such as arrays. Keeping both buffers allows callers to preserve
/// formulas even when a newer Excel token is not understood by the text
/// converter.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellParsedFormula {
    pub rgce: Vec<u8>,
    pub rgcb: Vec<u8>,
}

/// Scalar value stored in an XLSB `PtgExtraArray`.
#[derive(Debug, Clone, PartialEq)]
pub enum FormulaArrayValue {
    Number(f64),
    String(String),
    Bool(bool),
    Error(u8),
}

/// Kind of non-evaluating memory marker in an XLSB formula token stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormulaMemoryKind {
    Area,
    Error(u8),
    Function,
    NoMemory,
}

/// Row subset selected by an XLSB structured table reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormulaTableRowType {
    Data,
    All,
    Headers,
    DataAlternate,
    DataAndHeaders,
    Totals,
    DataAndTotals,
    Current,
}

/// Column subset selected by a resident structured table reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormulaTableColumns {
    All,
    One(u16),
    Range { first: u16, last: u16 },
}

/// Operand class carried by `PtgList`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormulaTableDataType {
    Reference,
    Value,
    Array,
}

/// Named column subset stored in a nonresident `PtgExtraList`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FormulaTableNamedColumns {
    All,
    One(String),
    Range { first: String, last: String },
}

/// Ancillary table/column names for a structured reference in another workbook.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormulaExternalTableReference {
    pub table: String,
    pub row_type: FormulaTableRowType,
    pub columns: FormulaTableNamedColumns,
}

/// Typed XLSB `PtgList` structured table reference.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormulaTableReference {
    pub sheet_index: u16,
    pub row_type: Option<FormulaTableRowType>,
    pub columns: Option<FormulaTableColumns>,
    pub square_bracket_space: bool,
    pub comma_space: bool,
    pub data_type: FormulaTableDataType,
    pub invalid: bool,
    pub list_index: Option<u32>,
    pub external: Option<FormulaExternalTableReference>,
}

impl CellParsedFormula {
    /// Parse a `CellParsedFormula`, returning the structure and bytes consumed.
    pub fn parse(data: &[u8]) -> Result<(Self, usize)> {
        if data.len() < 4 {
            return Err(Error::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }

        let cce = read_u32_le_at(data, 0)? as usize;
        if cce > MAX_CELL_FORMULA_BYTES {
            return Err(Error::InvalidFormula(format!(
                "cell formula token length {cce} exceeds {MAX_CELL_FORMULA_BYTES}"
            )));
        }
        let cb_offset = 4usize.checked_add(cce).ok_or_else(|| {
            Error::InvalidFormula("cell formula token length overflow".to_string())
        })?;
        if data.len() < cb_offset + 4 {
            return Err(Error::InvalidLength {
                expected: cb_offset + 4,
                found: data.len(),
            });
        }

        let cb = read_u32_le_at(data, cb_offset)? as usize;
        let end = cb_offset
            .checked_add(4)
            .and_then(|offset| offset.checked_add(cb))
            .ok_or_else(|| {
                Error::InvalidFormula("cell formula ancillary length overflow".to_string())
            })?;
        if data.len() < end {
            return Err(Error::InvalidLength {
                expected: end,
                found: data.len(),
            });
        }

        Ok((
            Self {
                rgce: data[4..cb_offset].to_vec(),
                rgcb: data[cb_offset + 4..end].to_vec(),
            },
            end,
        ))
    }

    /// Serialize this formula with its two length prefixes.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        if self.rgce.len() > MAX_CELL_FORMULA_BYTES {
            return Err(Error::InvalidFormula(format!(
                "cell formula token length {} exceeds {MAX_CELL_FORMULA_BYTES}",
                self.rgce.len()
            )));
        }
        let cce = u32::try_from(self.rgce.len())
            .map_err(|_| Error::InvalidFormula("formula is too large".to_string()))?;
        let cb = u32::try_from(self.rgcb.len()).map_err(|_| {
            Error::InvalidFormula("formula ancillary data is too large".to_string())
        })?;
        let mut bytes = Vec::with_capacity(8 + self.rgce.len() + self.rgcb.len());
        bytes.extend_from_slice(&cce.to_le_bytes());
        bytes.extend_from_slice(&self.rgce);
        bytes.extend_from_slice(&cb.to_le_bytes());
        bytes.extend_from_slice(&self.rgcb);
        Ok(bytes)
    }

    /// Create the `PtgExp` placeholder stored in every array/shared formula
    /// cell record.
    pub fn exp(row: u32, col: u32) -> Result<Self> {
        if row >= 1_048_576 || col >= 16_384 {
            return Err(Error::InvalidCellReference(format!(
                "grouped formula cell ({row}, {col})"
            )));
        }
        let mut rgce = Vec::with_capacity(5);
        rgce.push(ptg_types::PTG_EXP);
        rgce.extend_from_slice(&row.to_le_bytes());
        Ok(Self {
            rgce,
            rgcb: col.to_le_bytes().to_vec(),
        })
    }

    /// Return the target cell encoded by a `PtgExp`/`PtgExtraCol` formula.
    pub fn exp_cell(&self) -> Result<Option<(u32, u32)>> {
        if self.rgce.first() != Some(&ptg_types::PTG_EXP) {
            return Ok(None);
        }
        if self.rgce.len() != 5 || self.rgcb.len() != 4 {
            return Err(Error::InvalidFormula(format!(
                "PtgExp requires 5 rgce bytes and 4 rgcb bytes, found {} and {}",
                self.rgce.len(),
                self.rgcb.len()
            )));
        }
        let row = read_u32_le_at(&self.rgce, 1)?;
        let col = read_u32_le_at(&self.rgcb, 0)?;
        if row >= 1_048_576 || col >= 16_384 {
            return Err(Error::InvalidCellReference(format!(
                "PtgExp target ({row}, {col})"
            )));
        }
        Ok(Some((row, col)))
    }
}

/// Kind of formula definition following a `PtgExp` cell record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormulaGroupKind {
    Array,
    Shared,
}

/// Parsed `BrtArrFmla` or `BrtShrFmla` definition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormulaGroup {
    pub kind: FormulaGroupKind,
    pub range: FormulaRange,
    pub formula: CellParsedFormula,
    pub always_calculate: bool,
}

impl FormulaGroup {
    pub fn parse_array(data: &[u8]) -> Result<Self> {
        if data.len() < 17 {
            return Err(Error::InvalidLength {
                expected: 17,
                found: data.len(),
            });
        }
        if data[16] & !1 != 0 {
            return Err(Error::InvalidFormula(format!(
                "BrtArrFmla has reserved flag bits 0x{:02X}",
                data[16] & !1
            )));
        }
        let range = FormulaRange::parse_binary(data)?;
        let (formula, consumed) = CellParsedFormula::parse(&data[17..])?;
        if 17 + consumed != data.len() {
            return Err(Error::InvalidFormula(format!(
                "BrtArrFmla has {} trailing bytes",
                data.len() - 17 - consumed
            )));
        }
        Ok(Self {
            kind: FormulaGroupKind::Array,
            range,
            formula,
            always_calculate: data[16] & 1 != 0,
        })
    }

    pub fn parse_shared(data: &[u8]) -> Result<Self> {
        if data.len() < 16 {
            return Err(Error::InvalidLength {
                expected: 16,
                found: data.len(),
            });
        }
        let range = FormulaRange::parse_binary(data)?;
        let (formula, consumed) = CellParsedFormula::parse(&data[16..])?;
        if 16 + consumed != data.len() {
            return Err(Error::InvalidFormula(format!(
                "BrtShrFmla has {} trailing bytes",
                data.len() - 16 - consumed
            )));
        }
        Ok(Self {
            kind: FormulaGroupKind::Shared,
            range,
            formula,
            always_calculate: false,
        })
    }

    pub fn to_record_data(&self) -> Result<Vec<u8>> {
        self.range.validate()?;
        let formula = self.formula.to_bytes()?;
        let flag_len = usize::from(self.kind == FormulaGroupKind::Array);
        let mut data = Vec::with_capacity(16 + flag_len + formula.len());
        data.extend_from_slice(&self.range.to_binary());
        if self.kind == FormulaGroupKind::Array {
            data.push(u8::from(self.always_calculate));
        }
        data.extend_from_slice(&formula);
        Ok(data)
    }
}

/// Parse Tree Generator (Ptg) token types
///
/// These constants define the various formula token types used in XLSB.
/// Reference: [MS-XLSB] Section 2.5.98.16
#[allow(dead_code)]
pub mod ptg_types {
    // Operands
    pub const PTG_EXP: u8 = 0x01; // Expression
    pub const PTG_TBL: u8 = 0x02; // Table
    pub const PTG_ADD: u8 = 0x03; // Addition
    pub const PTG_SUB: u8 = 0x04; // Subtraction
    pub const PTG_MUL: u8 = 0x05; // Multiplication
    pub const PTG_DIV: u8 = 0x06; // Division
    pub const PTG_POWER: u8 = 0x07; // Exponentiation
    pub const PTG_CONCAT: u8 = 0x08; // Concatenation
    pub const PTG_LT: u8 = 0x09; // Less than
    pub const PTG_LE: u8 = 0x0A; // Less than or equal
    pub const PTG_EQ: u8 = 0x0B; // Equal
    pub const PTG_GE: u8 = 0x0C; // Greater than or equal
    pub const PTG_GT: u8 = 0x0D; // Greater than
    pub const PTG_NE: u8 = 0x0E; // Not equal
    pub const PTG_ISECT: u8 = 0x0F; // Intersection
    pub const PTG_UNION: u8 = 0x10; // Union
    pub const PTG_RANGE: u8 = 0x11; // Range
    pub const PTG_UPLUS: u8 = 0x12; // Unary plus
    pub const PTG_UMINUS: u8 = 0x13; // Unary minus
    pub const PTG_PERCENT: u8 = 0x14; // Percent
    pub const PTG_PAREN: u8 = 0x15; // Parentheses
    pub const PTG_MISSING_ARG: u8 = 0x16; // Missing argument
    pub const PTG_STR: u8 = 0x17; // String constant
    pub const PTG_EXTENDED: u8 = 0x18; // Extended token prefix
    pub const PTG_ATTR: u8 = 0x19; // Attribute
    pub const PTG_SHEET: u8 = 0x1A; // Sheet reference
    pub const PTG_END_SHEET: u8 = 0x1B; // End sheet reference
    pub const PTG_ERR: u8 = 0x1C; // Error value
    pub const PTG_BOOL: u8 = 0x1D; // Boolean constant
    pub const PTG_INT: u8 = 0x1E; // Integer constant
    pub const PTG_NUM: u8 = 0x1F; // Floating point constant

    // References
    pub const PTG_REF: u8 = 0x24; // Cell reference
    pub const PTG_AREA: u8 = 0x25; // Area reference
    pub const PTG_MEM_AREA: u8 = 0x26; // Memory area
    pub const PTG_MEM_ERR: u8 = 0x27; // Memory error
    pub const PTG_MEM_NO_MEM: u8 = 0x28; // Memory no memory
    pub const PTG_MEM_FUNC: u8 = 0x29; // Memory function
    pub const PTG_REF_ERR: u8 = 0x2A; // Reference error
    pub const PTG_AREA_ERR: u8 = 0x2B; // Area error
    pub const PTG_REF_N: u8 = 0x2C; // Cell reference (relative)
    pub const PTG_AREA_N: u8 = 0x2D; // Area reference (relative)

    // Functions
    pub const PTG_NAME_X: u8 = 0x39; // External name
    pub const PTG_REF_3D: u8 = 0x3A; // 3D cell reference
    pub const PTG_AREA_3D: u8 = 0x3B; // 3D area reference
    pub const PTG_REF_ERR_3D: u8 = 0x3C; // 3D reference error
    pub const PTG_AREA_ERR_3D: u8 = 0x3D; // 3D area error

    // Function calls
    pub const PTG_FUNC: u8 = 0x21; // Built-in function with fixed args
    pub const PTG_FUNC_VAR: u8 = 0x22; // Built-in function with variable args

    // Array and name
    pub const PTG_NAME: u8 = 0x23; // Defined name
    pub const PTG_ARRAY: u8 = 0x20; // Array constant
    pub const EPTG_LIST: u8 = 0x19; // Structured table reference
    pub const EPTG_SX_NAME: u8 = 0x1D; // Pivot calculated name
}

/// Formula token representation
///
/// Represents a single token in a formula's RPN sequence.
#[derive(Debug, Clone, PartialEq)]
pub enum FormulaToken {
    /// Number constant
    Number(f64),
    /// String constant
    String(String),
    /// Boolean constant
    Bool(bool),
    /// Error value
    Error(u8),
    /// Integer constant
    Int(u16),
    /// Omitted function argument (`PtgMissArg`).
    MissingArg,
    /// Display parenthesis around the preceding expression (`PtgParen`).
    Parenthesis,
    /// Non-evaluating control/display attribute. The selector is retained for
    /// diagnostics while its payload is consumed by the parser.
    Attribute(u8),
    /// Rectangular constant array stored across `PtgArray` and `RgbExtra`.
    Array {
        rows: u32,
        cols: u32,
        values: Vec<FormulaArrayValue>,
    },
    /// Prefix metadata for a following binary reference expression.
    Memory {
        kind: FormulaMemoryKind,
        expression_bytes: u16,
        /// Cached `UncheckedRfX` values in field order.
        cached_ranges: Vec<[u32; 4]>,
    },
    /// Cell reference (row, col, relative_row, relative_col)
    CellRef {
        row: u32,
        col: u32,
        row_relative: bool,
        col_relative: bool,
    },
    /// Area reference (first_row, last_row, first_col, last_col)
    AreaRef {
        row_first: u32,
        row_last: u32,
        col_first: u32,
        col_last: u32,
        row_first_relative: bool,
        row_last_relative: bool,
        col_first_relative: bool,
        col_last_relative: bool,
    },
    /// Cell reference qualified by an extern-sheet (`Xti`) index.
    CellRef3d {
        sheet_index: u16,
        row: u32,
        col: u32,
        row_relative: bool,
        col_relative: bool,
    },
    /// Area reference qualified by an extern-sheet (`Xti`) index.
    AreaRef3d {
        sheet_index: u16,
        row_first: u32,
        row_last: u32,
        col_first: u32,
        col_last: u32,
        row_first_relative: bool,
        row_last_relative: bool,
        col_first_relative: bool,
        col_last_relative: bool,
    },
    /// Invalid single-cell or area reference. A sheet index is present for
    /// the 3D token forms and retained even though the text form is `#REF!`.
    ReferenceError {
        is_area: bool,
        sheet_index: Option<u16>,
    },
    /// Binary operator
    BinaryOp(BinaryOperator),
    /// Unary operator
    UnaryOp(UnaryOperator),
    /// Function call (function index, argument count, command-table flag).
    Function {
        index: u16,
        arg_count: u8,
        is_command: bool,
    },
    /// Defined name reference
    Name(u32),
    /// Defined name in an external workbook.
    ExternalName { sheet_index: u16, name_index: u32 },
    /// Structured table reference (`PtgList`), including nonresident ancillary names.
    TableReference(FormulaTableReference),
    /// Zero-based calculated pivot field/item name index (`PtgSxName`).
    PivotName(u32),
    /// Unknown/unsupported token
    Unknown(u8),
}

impl FormulaToken {
    /// Encode one of the XLSB extended tokens implemented by this model.
    ///
    /// The first vector is the `Rgce` token and the second is its corresponding
    /// `RgbExtra` payload. Other token families are intentionally rejected.
    pub fn to_extended_binary(&self) -> Result<(Vec<u8>, Vec<u8>)> {
        match self {
            Self::PivotName(index) => {
                let mut token = Vec::with_capacity(6);
                token.extend([ptg_types::PTG_EXTENDED, ptg_types::EPTG_SX_NAME]);
                token.extend_from_slice(&index.to_le_bytes());
                Ok((token, Vec::new()))
            },
            Self::TableReference(reference) => reference.to_extended_binary(),
            _ => Err(Error::InvalidFormula(
                "token is not an extended PtgList/PtgSxName token".to_string(),
            )),
        }
    }
}

impl FormulaTableReference {
    pub fn to_extended_binary(&self) -> Result<(Vec<u8>, Vec<u8>)> {
        let external = self.external.as_ref();
        if self.invalid {
            if self.row_type.is_some()
                || self.columns.is_some()
                || self.list_index.is_some()
                || external.is_some()
            {
                return Err(Error::InvalidFormula(
                    "invalid PtgList cannot carry resident or external table metadata".to_string(),
                ));
            }
        } else if external.is_some() {
            if self.row_type.is_some() || self.columns.is_some() || self.list_index.is_some() {
                return Err(Error::InvalidFormula(
                    "nonresident PtgList cannot carry resident table metadata".to_string(),
                ));
            }
        } else if self.row_type.is_none() || self.columns.is_none() || self.list_index.is_none() {
            return Err(Error::InvalidFormula(
                "resident PtgList is missing table metadata".to_string(),
            ));
        }

        let mut flags = match self.data_type {
            FormulaTableDataType::Reference => 0,
            FormulaTableDataType::Value => 1 << 10,
            FormulaTableDataType::Array => 2 << 10,
        };
        if self.square_bracket_space {
            flags |= 0x0080;
        }
        if self.comma_space {
            flags |= 0x0100;
        }
        if self.invalid {
            flags |= 0x1000;
        }
        if external.is_some() {
            flags |= 0x2000;
        }

        let (list_index, first, last) = if let (Some(row_type), Some(columns), Some(list_index)) =
            (self.row_type, self.columns, self.list_index)
        {
            if list_index == 0 || list_index == u32::MAX {
                return Err(Error::InvalidFormula(format!(
                    "PtgList has invalid table identifier {list_index}"
                )));
            }
            flags |= u16::from(table_row_type_raw(row_type)) << 2;
            match columns {
                FormulaTableColumns::All => (list_index, 0, 0),
                FormulaTableColumns::One(column) => {
                    if column >= 16_384 {
                        return Err(Error::InvalidFormula(
                            "PtgList column is outside worksheet bounds".to_string(),
                        ));
                    }
                    flags |= 1;
                    (list_index, column, column)
                },
                FormulaTableColumns::Range { first, last } => {
                    if first > last || last >= 16_384 {
                        return Err(Error::InvalidFormula(
                            "PtgList column range is invalid".to_string(),
                        ));
                    }
                    flags |= 2;
                    (list_index, first, last)
                },
            }
        } else {
            (0, 0, 0)
        };

        let mut token = Vec::with_capacity(14);
        token.extend([ptg_types::PTG_EXTENDED, ptg_types::EPTG_LIST]);
        token.extend_from_slice(&self.sheet_index.to_le_bytes());
        token.extend_from_slice(&flags.to_le_bytes());
        token.extend_from_slice(&list_index.to_le_bytes());
        token.extend_from_slice(&first.to_le_bytes());
        token.extend_from_slice(&last.to_le_bytes());
        let extra = external
            .map(write_extra_list)
            .transpose()?
            .unwrap_or_default();
        Ok((token, extra))
    }
}

fn write_extra_list(reference: &FormulaExternalTableReference) -> Result<Vec<u8>> {
    let table_units = reference.table.encode_utf16().count();
    if table_units == 0 || table_units >= 256 {
        return Err(Error::InvalidFormula(format!(
            "PtgExtraList table length {table_units} is outside 1..=255"
        )));
    }
    let has_columns = !matches!(reference.columns, FormulaTableNamedColumns::All);
    let mut extra = Vec::new();
    extra.push(u8::from(has_columns));
    extra.extend_from_slice(&u16::from(table_row_type_raw(reference.row_type)).to_le_bytes());
    extra.extend_from_slice(&(table_units as u16).to_le_bytes());
    push_formula_utf16(&mut extra, &reference.table);
    match &reference.columns {
        FormulaTableNamedColumns::All => {},
        FormulaTableNamedColumns::One(name) => {
            extra.extend([0, 0, 1]);
            write_sxos(&mut extra, false, name)?;
        },
        FormulaTableNamedColumns::Range { first, last } => {
            extra.extend([0, 0, 2]);
            write_sxos(&mut extra, true, first)?;
            write_sxos(&mut extra, false, last)?;
        },
    }
    Ok(extra)
}

fn write_sxos(output: &mut Vec<u8>, not_last: bool, name: &str) -> Result<()> {
    let units = name.encode_utf16().count();
    if units == 0 || units > 1_048_576 {
        return Err(Error::InvalidFormula(format!(
            "structured-reference column length {units} is outside 1..=1048576"
        )));
    }
    output.push(u8::from(not_last));
    output.extend_from_slice(&2u16.to_le_bytes());
    output.extend_from_slice(&(units as u32).to_le_bytes());
    push_formula_utf16(output, name);
    Ok(())
}

fn push_formula_utf16(output: &mut Vec<u8>, value: &str) {
    output.extend(value.encode_utf16().flat_map(u16::to_le_bytes));
}

/// Binary operators
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinaryOperator {
    Add,
    Subtract,
    Multiply,
    Divide,
    Power,
    Concat,
    LessThan,
    LessEqual,
    Equal,
    GreaterEqual,
    GreaterThan,
    NotEqual,
    Intersection,
    Union,
    Range,
}

/// Unary operators
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnaryOperator {
    Plus,
    Minus,
    Percent,
}

/// Formula parser
///
/// Parses binary formula bytes into a sequence of tokens.
pub struct FormulaParser<'a> {
    data: &'a [u8],
    offset: usize,
    extra: &'a [u8],
    extra_offset: usize,
    validate_extra: bool,
    memory_expression_ends: Vec<usize>,
    control_flow_targets: Vec<usize>,
    base_cell: Option<(u32, u32)>,
}

impl<'a> FormulaParser<'a> {
    /// Create a new formula parser
    pub fn new(data: &'a [u8]) -> Self {
        FormulaParser {
            data,
            offset: 0,
            extra: &[],
            extra_offset: 0,
            validate_extra: false,
            memory_expression_ends: Vec::new(),
            control_flow_targets: Vec::new(),
            base_cell: None,
        }
    }

    /// Parse a formula together with its corresponding `RgbExtra` stream.
    pub fn with_extra(data: &'a [u8], extra: &'a [u8]) -> Self {
        FormulaParser {
            data,
            offset: 0,
            extra,
            extra_offset: 0,
            validate_extra: true,
            memory_expression_ends: Vec::new(),
            control_flow_targets: Vec::new(),
            base_cell: None,
        }
    }

    /// Parse a shared-formula definition relative to a concrete target cell.
    pub fn with_base_cell(data: &'a [u8], row: u32, col: u32) -> Self {
        FormulaParser {
            data,
            offset: 0,
            extra: &[],
            extra_offset: 0,
            validate_extra: false,
            memory_expression_ends: Vec::new(),
            control_flow_targets: Vec::new(),
            base_cell: Some((row, col)),
        }
    }

    /// Parse a shared formula and ancillary stream at a concrete target cell.
    pub fn with_base_cell_and_extra(data: &'a [u8], extra: &'a [u8], row: u32, col: u32) -> Self {
        FormulaParser {
            data,
            offset: 0,
            extra,
            extra_offset: 0,
            validate_extra: true,
            memory_expression_ends: Vec::new(),
            control_flow_targets: Vec::new(),
            base_cell: Some((row, col)),
        }
    }

    /// Parse the formula into tokens
    ///
    /// Returns a vector of formula tokens in RPN order.
    pub fn parse(&mut self) -> Result<Vec<FormulaToken>> {
        let mut tokens = Vec::new();
        let mut boundaries = Vec::new();

        while self.offset < self.data.len() {
            tokens.push(self.parse_token()?);
            boundaries.push(self.offset);
        }

        if let Some(end) = self
            .memory_expression_ends
            .iter()
            .find(|end| boundaries.binary_search(end).is_err())
        {
            return Err(Error::InvalidFormula(format!(
                "memory expression ends at byte {end}, which is not a token boundary"
            )));
        }

        if let Some(target) = self
            .control_flow_targets
            .iter()
            .find(|target| boundaries.binary_search(target).is_err())
        {
            return Err(Error::InvalidFormula(format!(
                "control-flow target byte {target} is not a token boundary"
            )));
        }

        if self.validate_extra && self.extra_offset != self.extra.len() {
            return Err(Error::InvalidFormula(format!(
                "formula has {} unconsumed ancillary bytes",
                self.extra.len() - self.extra_offset
            )));
        }

        Ok(tokens)
    }

    /// Parse a single token
    fn parse_token(&mut self) -> Result<FormulaToken> {
        if self.offset >= self.data.len() {
            return Err(Error::InvalidFormula(
                "unexpected end of formula token stream".to_string(),
            ));
        }

        let ptg_type = self.data[self.offset];
        self.offset += 1;

        use ptg_types::*;

        match ptg_type {
            PTG_ADD => Ok(FormulaToken::BinaryOp(BinaryOperator::Add)),
            PTG_SUB => Ok(FormulaToken::BinaryOp(BinaryOperator::Subtract)),
            PTG_MUL => Ok(FormulaToken::BinaryOp(BinaryOperator::Multiply)),
            PTG_DIV => Ok(FormulaToken::BinaryOp(BinaryOperator::Divide)),
            PTG_POWER => Ok(FormulaToken::BinaryOp(BinaryOperator::Power)),
            PTG_CONCAT => Ok(FormulaToken::BinaryOp(BinaryOperator::Concat)),
            PTG_LT => Ok(FormulaToken::BinaryOp(BinaryOperator::LessThan)),
            PTG_LE => Ok(FormulaToken::BinaryOp(BinaryOperator::LessEqual)),
            PTG_EQ => Ok(FormulaToken::BinaryOp(BinaryOperator::Equal)),
            PTG_GE => Ok(FormulaToken::BinaryOp(BinaryOperator::GreaterEqual)),
            PTG_GT => Ok(FormulaToken::BinaryOp(BinaryOperator::GreaterThan)),
            PTG_NE => Ok(FormulaToken::BinaryOp(BinaryOperator::NotEqual)),
            PTG_ISECT => Ok(FormulaToken::BinaryOp(BinaryOperator::Intersection)),
            PTG_UNION => Ok(FormulaToken::BinaryOp(BinaryOperator::Union)),
            PTG_RANGE => Ok(FormulaToken::BinaryOp(BinaryOperator::Range)),

            PTG_UPLUS => Ok(FormulaToken::UnaryOp(UnaryOperator::Plus)),
            PTG_UMINUS => Ok(FormulaToken::UnaryOp(UnaryOperator::Minus)),
            PTG_PERCENT => Ok(FormulaToken::UnaryOp(UnaryOperator::Percent)),
            PTG_PAREN => Ok(FormulaToken::Parenthesis),
            PTG_MISSING_ARG => Ok(FormulaToken::MissingArg),
            PTG_EXTENDED => self.parse_extended(),
            PTG_ATTR => self.parse_attr(),

            PTG_INT => self.parse_int(),
            PTG_NUM => self.parse_num(),
            PTG_STR => self.parse_str(),
            PTG_BOOL => self.parse_bool(),
            PTG_ERR => self.parse_err(),

            _ if ptg_type >= 0x20 => match ptg_type & 0x1F {
                0x04 => self.parse_ref(false),
                0x0C => self.parse_ref(true),
                0x05 => self.parse_area(false),
                0x0D => self.parse_area(true),
                0x0A => self.parse_reference_error(false, false),
                0x0B => self.parse_reference_error(true, false),
                0x1C => self.parse_reference_error(false, true),
                0x1D => self.parse_reference_error(true, true),
                0x1A => self.parse_ref_3d(),
                0x1B => self.parse_area_3d(),
                0x00 => self.parse_array(),
                0x06 => self.parse_memory(FormulaMemoryKind::Area),
                0x07 => self.parse_memory(FormulaMemoryKind::Error(0)),
                0x08 => self.parse_memory(FormulaMemoryKind::NoMemory),
                0x09 => self.parse_memory(FormulaMemoryKind::Function),
                0x01 => self.parse_func(),
                0x02 => self.parse_func_var(),
                0x03 => self.parse_name(),
                0x19 => self.parse_name_x(),
                _ => Ok(FormulaToken::Unknown(ptg_type)),
            },

            _ => {
                // Unknown token type
                Ok(FormulaToken::Unknown(ptg_type))
            },
        }
    }

    fn require(&self, len: usize, context: &str) -> Result<()> {
        if self.offset + len <= self.data.len() {
            Ok(())
        } else {
            Err(Error::InvalidFormula(format!(
                "truncated {context} token at byte {}: need {len} bytes, have {}",
                self.offset.saturating_sub(1),
                self.data.len().saturating_sub(self.offset)
            )))
        }
    }

    fn validate_classed_token(&self, context: &str) -> Result<()> {
        let token = self.data[self.offset - 1];
        if token & 0x80 != 0 {
            return Err(Error::InvalidFormula(format!(
                "{context} token 0x{token:02X} has its reserved bit set"
            )));
        }
        Ok(())
    }

    /// Parse integer constant
    fn parse_int(&mut self) -> Result<FormulaToken> {
        self.require(2, "PtgInt")?;

        let value = read_u16_le_at(self.data, self.offset)?;
        self.offset += 2;

        Ok(FormulaToken::Int(value))
    }

    /// Parse floating point constant
    fn parse_num(&mut self) -> Result<FormulaToken> {
        self.require(8, "PtgNum")?;

        let value = read_f64_le_at(self.data, self.offset)?;
        self.offset += 8;
        validate_xnum(value, "PtgNum")?;

        Ok(FormulaToken::Number(value))
    }

    /// Parse string constant
    fn parse_str(&mut self) -> Result<FormulaToken> {
        self.require(2, "PtgStr length")?;
        let len = read_u16_le_at(self.data, self.offset)? as usize;
        self.offset += 2;
        let byte_len = len
            .checked_mul(2)
            .ok_or_else(|| Error::InvalidFormula("PtgStr UTF-16 length overflow".to_string()))?;
        self.require(byte_len, "PtgStr text")?;
        let units = self.data[self.offset..self.offset + byte_len]
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]));
        let string = char::decode_utf16(units)
            .collect::<std::result::Result<String, _>>()
            .map_err(|_| Error::Encoding("invalid UTF-16 in PtgStr".to_string()))?;
        self.offset += byte_len;

        Ok(FormulaToken::String(string))
    }

    /// Parse boolean constant
    fn parse_bool(&mut self) -> Result<FormulaToken> {
        self.require(1, "PtgBool")?;

        let raw = self.data[self.offset];
        self.offset += 1;
        if raw > 1 {
            return Err(Error::InvalidFormula(format!(
                "invalid PtgBool value {raw}"
            )));
        }

        Ok(FormulaToken::Bool(raw != 0))
    }

    /// Parse error constant
    fn parse_err(&mut self) -> Result<FormulaToken> {
        self.require(1, "PtgErr")?;

        let error_code = self.data[self.offset];
        self.offset += 1;
        if !is_formula_error_code(error_code) {
            return Err(Error::InvalidFormula(format!(
                "invalid PtgErr code 0x{error_code:02X}"
            )));
        }

        Ok(FormulaToken::Error(error_code))
    }

    /// Parse the selector-specific payload of the `PtgAttr` token family.
    fn parse_attr(&mut self) -> Result<FormulaToken> {
        self.require(1, "PtgAttr selector")?;
        let selector = self.data[self.offset];
        self.offset += 1;

        if selector == 0x04 {
            // PtgAttrChoose: cOffset is one less than the number of u16
            // offsets that follow it.
            self.require(2, "PtgAttrChoose count")?;
            let count = usize::from(read_u16_le_at(self.data, self.offset)?) + 1;
            self.offset += 2;
            let byte_len = count.checked_mul(2).ok_or_else(|| {
                Error::InvalidFormula("PtgAttrChoose offset count overflow".to_string())
            })?;
            self.require(byte_len, "PtgAttrChoose offsets")?;
            let offsets = self.data[self.offset..self.offset + byte_len]
                .chunks_exact(2)
                .map(|bytes| usize::from(u16::from_le_bytes([bytes[0], bytes[1]])))
                .collect::<Vec<_>>();
            self.offset += byte_len;
            if offsets[0] != byte_len {
                return Err(Error::InvalidFormula(format!(
                    "PtgAttrChoose first offset is {}, expected {byte_len}",
                    offsets[0]
                )));
            }
            self.control_flow_targets.push(self.offset);
            for offset in &offsets[1..] {
                let target = self.offset.checked_add(*offset).ok_or_else(|| {
                    Error::InvalidFormula("PtgAttrChoose target overflow".to_string())
                })?;
                if target > self.data.len() {
                    return Err(Error::InvalidFormula(format!(
                        "PtgAttrChoose target byte {target} exceeds formula length {}",
                        self.data.len()
                    )));
                }
                self.control_flow_targets.push(target);
            }
            return Ok(FormulaToken::Attribute(selector));
        }

        match selector {
            // Semi, If, GoTo, Sum, Baxcel, Space, SpaceSemi, IfError all have
            // a two-byte selector-specific payload after the selector byte.
            0x01 | 0x02 | 0x08 | 0x10 | 0x20 | 0x21 | 0x40 | 0x41 | 0x80 => {
                self.require(2, "PtgAttr payload")?;
                let offset = usize::from(read_u16_le_at(self.data, self.offset)?);
                self.offset += 2;
                if matches!(selector, 0x02 | 0x08 | 0x80) {
                    let adjustment = usize::from(selector == 0x08);
                    let target = self
                        .offset
                        .checked_add(offset)
                        .and_then(|value| value.checked_add(adjustment))
                        .ok_or_else(|| {
                            Error::InvalidFormula("PtgAttr target overflow".to_string())
                        })?;
                    if target > self.data.len() {
                        return Err(Error::InvalidFormula(format!(
                            "PtgAttr target byte {target} exceeds formula length {}",
                            self.data.len()
                        )));
                    }
                    self.control_flow_targets.push(target);
                }
                if selector == 0x10 {
                    // PtgAttrSum is semantically SUM(expression).
                    Ok(FormulaToken::Function {
                        index: 4,
                        arg_count: 1,
                        is_command: false,
                    })
                } else {
                    Ok(FormulaToken::Attribute(selector))
                }
            },
            _ => Err(Error::InvalidFormula(format!(
                "unknown PtgAttr selector 0x{selector:02X}"
            ))),
        }
    }

    fn parse_extended(&mut self) -> Result<FormulaToken> {
        self.require(1, "extended Ptg selector")?;
        let selector = self.data[self.offset];
        self.offset += 1;
        match selector {
            ptg_types::EPTG_LIST => self.parse_list(),
            ptg_types::EPTG_SX_NAME => {
                self.require(4, "PtgSxName")?;
                let index = read_u32_le_at(self.data, self.offset)?;
                self.offset += 4;
                Ok(FormulaToken::PivotName(index))
            },
            _ => Err(Error::InvalidFormula(format!(
                "unknown extended Ptg selector 0x{selector:02X}"
            ))),
        }
    }

    fn parse_list(&mut self) -> Result<FormulaToken> {
        self.require(12, "PtgList")?;
        let sheet_index = read_u16_le_at(self.data, self.offset)?;
        let flags = read_u16_le_at(self.data, self.offset + 2)?;
        let raw_list_index = read_u32_le_at(self.data, self.offset + 4)?;
        let col_first = read_u16_le_at(self.data, self.offset + 8)?;
        let col_last = read_u16_le_at(self.data, self.offset + 10)?;
        self.offset += 12;

        if flags & 0xC000 != 0 {
            return Err(Error::InvalidFormula(format!(
                "PtgList reserved flag bits are nonzero: 0x{:04X}",
                flags & 0xC000
            )));
        }
        let data_type = match (flags >> 10) & 0x03 {
            0 => FormulaTableDataType::Reference,
            1 => FormulaTableDataType::Value,
            2 => FormulaTableDataType::Array,
            _ => {
                return Err(Error::InvalidFormula(
                    "PtgList has reserved data type 3".to_string(),
                ));
            },
        };
        let invalid = flags & 0x1000 != 0;
        let nonresident = !invalid && flags & 0x2000 != 0;
        let (row_type, columns, list_index) = if invalid || nonresident {
            (None, None, None)
        } else {
            if raw_list_index == 0 || raw_list_index == u32::MAX {
                return Err(Error::InvalidFormula(format!(
                    "PtgList has invalid table identifier {raw_list_index}"
                )));
            }
            let row_type = parse_table_row_type(((flags >> 2) & 0x1F) as u8)?;
            let columns = match flags & 0x03 {
                0 => FormulaTableColumns::All,
                1 => {
                    if col_first >= 16_384 {
                        return Err(Error::InvalidFormula(
                            "PtgList first column is outside worksheet bounds".to_string(),
                        ));
                    }
                    FormulaTableColumns::One(col_first)
                },
                2 => {
                    if col_first > col_last || col_last >= 16_384 {
                        return Err(Error::InvalidFormula(
                            "PtgList column range is invalid".to_string(),
                        ));
                    }
                    FormulaTableColumns::Range {
                        first: col_first,
                        last: col_last,
                    }
                },
                _ => {
                    return Err(Error::InvalidFormula(
                        "PtgList has reserved column selector 3".to_string(),
                    ));
                },
            };
            (Some(row_type), Some(columns), Some(raw_list_index))
        };
        let external = if nonresident {
            Some(self.parse_extra_list()?)
        } else {
            None
        };
        Ok(FormulaToken::TableReference(FormulaTableReference {
            sheet_index,
            row_type,
            columns,
            square_bracket_space: flags & 0x0080 != 0,
            comma_space: flags & 0x0100 != 0,
            data_type,
            invalid,
            list_index,
            external,
        }))
    }

    fn parse_extra_list(&mut self) -> Result<FormulaExternalTableReference> {
        self.require_extra(5, "PtgExtraList header")?;
        let has_columns = match self.extra[self.extra_offset] {
            0 => false,
            1 => true,
            value => {
                return Err(Error::InvalidFormula(format!(
                    "invalid PtgExtraList hasColumns {value}"
                )));
            },
        };
        let row_flags = read_u16_le_at(self.extra, self.extra_offset + 1)?;
        if row_flags & !0x001F != 0 {
            return Err(Error::InvalidFormula(format!(
                "PtgExtraList reserved row bits are nonzero: 0x{:04X}",
                row_flags & !0x001F
            )));
        }
        let row_type = parse_table_row_type((row_flags & 0x1F) as u8)?;
        let table_len = usize::from(read_u16_le_at(self.extra, self.extra_offset + 3)?);
        self.extra_offset += 5;
        if table_len == 0 || table_len >= 256 {
            return Err(Error::InvalidFormula(format!(
                "PtgExtraList table length {table_len} is outside 1..=255"
            )));
        }
        let table = self.parse_extra_utf16(table_len, "PtgExtraList table")?;
        let columns = if has_columns {
            self.require_extra(3, "SxSu header")?;
            let reserved = read_u16_le_at(self.extra, self.extra_offset)?;
            let count = self.extra[self.extra_offset + 2];
            self.extra_offset += 3;
            if reserved != 0 {
                return Err(Error::InvalidFormula(
                    "SxSu reserved field is nonzero".to_string(),
                ));
            }
            if !matches!(count, 1 | 2) {
                return Err(Error::InvalidFormula(format!(
                    "invalid SxSu column count {count}"
                )));
            }
            let first = self.parse_sxos(count == 2, "SxSu first column")?;
            if count == 1 {
                FormulaTableNamedColumns::One(first)
            } else {
                let last = self.parse_sxos(false, "SxSu last column")?;
                FormulaTableNamedColumns::Range { first, last }
            }
        } else {
            FormulaTableNamedColumns::All
        };
        Ok(FormulaExternalTableReference {
            table,
            row_type,
            columns,
        })
    }

    fn parse_sxos(&mut self, expected_not_last: bool, context: &str) -> Result<String> {
        self.require_extra(7, context)?;
        let not_last = match self.extra[self.extra_offset] {
            0 => false,
            1 => true,
            value => {
                return Err(Error::InvalidFormula(format!(
                    "invalid {context} notLast {value}"
                )));
            },
        };
        let reserved = read_u16_le_at(self.extra, self.extra_offset + 1)?;
        let length = usize::try_from(read_u32_le_at(self.extra, self.extra_offset + 3)?)
            .map_err(|_| Error::InvalidFormula(format!("{context} length overflow")))?;
        self.extra_offset += 7;
        if not_last != expected_not_last {
            return Err(Error::InvalidFormula(format!(
                "{context} has inconsistent notLast flag"
            )));
        }
        if reserved != 2 {
            return Err(Error::InvalidFormula(format!(
                "{context} reserved field is {reserved}, expected 2"
            )));
        }
        if length == 0 {
            return Err(Error::InvalidFormula(format!("{context} name is empty")));
        }
        self.parse_extra_utf16(length, context)
    }

    fn parse_extra_utf16(&mut self, units: usize, context: &str) -> Result<String> {
        let byte_len = units
            .checked_mul(2)
            .ok_or_else(|| Error::InvalidFormula(format!("{context} length overflow")))?;
        self.require_extra(byte_len, context)?;
        let value = char::decode_utf16(
            self.extra[self.extra_offset..self.extra_offset + byte_len]
                .chunks_exact(2)
                .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]])),
        )
        .collect::<std::result::Result<String, _>>()
        .map_err(|_| Error::Encoding(format!("invalid UTF-16 in {context}")))?;
        self.extra_offset += byte_len;
        Ok(value)
    }

    fn require_extra(&self, len: usize, context: &str) -> Result<()> {
        if self.extra_offset + len <= self.extra.len() {
            Ok(())
        } else {
            Err(Error::InvalidFormula(format!(
                "truncated {context} at ancillary byte {}: need {len} bytes, have {}",
                self.extra_offset,
                self.extra.len().saturating_sub(self.extra_offset)
            )))
        }
    }

    fn parse_array(&mut self) -> Result<FormulaToken> {
        let token = self.data[self.offset - 1];
        if token & 0x80 != 0 || !matches!(token & 0x60, 0x40 | 0x60) {
            return Err(Error::InvalidFormula(format!(
                "PtgArray has invalid data type in token 0x{token:02X}"
            )));
        }
        self.require(14, "PtgArray")?;
        self.offset += 14;
        self.require_extra(8, "PtgExtraArray dimensions")?;
        let rows = read_u32_le_at(self.extra, self.extra_offset)?;
        let cols = read_u32_le_at(self.extra, self.extra_offset + 4)?;
        self.extra_offset += 8;
        if rows == 0 || cols == 0 || rows > 1_048_576 || cols > 16_384 {
            return Err(Error::InvalidFormula(format!(
                "PtgExtraArray dimensions {rows}x{cols} are invalid"
            )));
        }
        let count_u64 = u64::from(rows)
            .checked_mul(u64::from(cols))
            .ok_or_else(|| Error::InvalidFormula("array size overflow".to_string()))?;
        let count = usize::try_from(count_u64)
            .map_err(|_| Error::InvalidFormula("array is too large".to_string()))?;
        // Every SerAr uses at least two bytes. Reject impossible dimensions
        // before allocating based on attacker-controlled counts.
        if count > self.extra.len().saturating_sub(self.extra_offset) / 2 {
            return Err(Error::InvalidFormula(format!(
                "PtgExtraArray declares {count} values beyond its ancillary payload"
            )));
        }
        let mut values = Vec::with_capacity(count);
        for _ in 0..count {
            self.require_extra(1, "SerAr tag")?;
            let tag = self.extra[self.extra_offset];
            self.extra_offset += 1;
            let value = match tag {
                0x00 => {
                    self.require_extra(8, "SerNum")?;
                    let value = read_f64_le_at(self.extra, self.extra_offset)?;
                    self.extra_offset += 8;
                    validate_xnum(value, "SerNum")?;
                    FormulaArrayValue::Number(value)
                },
                0x01 => {
                    self.require_extra(2, "SerStr length")?;
                    let len = usize::from(read_u16_le_at(self.extra, self.extra_offset)?);
                    self.extra_offset += 2;
                    if len >= 256 {
                        return Err(Error::InvalidFormula(format!(
                            "SerStr length {len} exceeds 255 UTF-16 code units"
                        )));
                    }
                    let byte_len = len.checked_mul(2).ok_or_else(|| {
                        Error::InvalidFormula("SerStr length overflow".to_string())
                    })?;
                    self.require_extra(byte_len, "SerStr text")?;
                    let units = self.extra[self.extra_offset..self.extra_offset + byte_len]
                        .chunks_exact(2)
                        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]));
                    let value = char::decode_utf16(units)
                        .collect::<std::result::Result<String, _>>()
                        .map_err(|_| Error::Encoding("invalid UTF-16 in SerStr".to_string()))?;
                    self.extra_offset += byte_len;
                    FormulaArrayValue::String(value)
                },
                0x02 => {
                    self.require_extra(1, "SerBool")?;
                    let value = self.extra[self.extra_offset];
                    self.extra_offset += 1;
                    if value > 1 {
                        return Err(Error::InvalidFormula(format!(
                            "invalid SerBool value {value}"
                        )));
                    }
                    FormulaArrayValue::Bool(value != 0)
                },
                0x04 => {
                    self.require_extra(4, "SerErr")?;
                    let error = self.extra[self.extra_offset];
                    if !matches!(error, 0x00 | 0x07 | 0x0F | 0x17 | 0x1D | 0x24 | 0x2A | 0x2B) {
                        return Err(Error::InvalidFormula(format!(
                            "invalid SerErr code 0x{error:02X}"
                        )));
                    }
                    if self.extra[self.extra_offset + 1..self.extra_offset + 4]
                        .iter()
                        .any(|byte| *byte != 0)
                    {
                        return Err(Error::InvalidFormula(
                            "SerErr reserved bytes are nonzero".to_string(),
                        ));
                    }
                    self.extra_offset += 4;
                    FormulaArrayValue::Error(error)
                },
                _ => {
                    return Err(Error::InvalidFormula(format!(
                        "unknown SerAr tag 0x{tag:02X}"
                    )));
                },
            };
            values.push(value);
        }
        Ok(FormulaToken::Array { rows, cols, values })
    }

    fn parse_memory(&mut self, mut kind: FormulaMemoryKind) -> Result<FormulaToken> {
        let token = self.data[self.offset - 1];
        if token & 0x80 != 0 {
            return Err(Error::InvalidFormula(format!(
                "memory token 0x{token:02X} has its reserved bit set"
            )));
        }
        let (payload_len, cce_offset) = match kind {
            FormulaMemoryKind::Function => (2, 0),
            FormulaMemoryKind::Area | FormulaMemoryKind::NoMemory => (6, 4),
            FormulaMemoryKind::Error(_) => (6, 4),
        };
        self.require(payload_len, "memory token")?;
        if matches!(kind, FormulaMemoryKind::Error(_)) {
            let error = self.data[self.offset];
            if !matches!(error, 0x00 | 0x07 | 0x0F | 0x17 | 0x1D | 0x24 | 0x2A | 0x2B) {
                return Err(Error::InvalidFormula(format!(
                    "invalid PtgMemErr code 0x{error:02X}"
                )));
            }
            kind = FormulaMemoryKind::Error(error);
        }
        let expression_bytes = read_u16_le_at(self.data, self.offset + cce_offset)?;
        self.offset += payload_len;
        if expression_bytes == 0 {
            return Err(Error::InvalidFormula(
                "memory token has an empty reference expression".to_string(),
            ));
        }
        if usize::from(expression_bytes) > self.data.len().saturating_sub(self.offset) {
            return Err(Error::InvalidFormula(format!(
                "memory token declares {expression_bytes} expression bytes, but only {} remain",
                self.data.len().saturating_sub(self.offset)
            )));
        }
        self.memory_expression_ends
            .push(self.offset + usize::from(expression_bytes));

        let mut cached_ranges = Vec::new();
        if kind == FormulaMemoryKind::Area {
            self.require_extra(4, "PtgExtraMem count")?;
            let count = usize::try_from(read_u32_le_at(self.extra, self.extra_offset)?)
                .map_err(|_| Error::InvalidFormula("PtgExtraMem is too large".to_string()))?;
            self.extra_offset += 4;
            if count > self.extra.len().saturating_sub(self.extra_offset) / 16 {
                return Err(Error::InvalidFormula(format!(
                    "PtgExtraMem declares {count} ranges beyond its ancillary payload"
                )));
            }
            cached_ranges.reserve(count);
            for _ in 0..count {
                self.require_extra(16, "PtgExtraMem range")?;
                let range = [
                    read_u32_le_at(self.extra, self.extra_offset)?,
                    read_u32_le_at(self.extra, self.extra_offset + 4)?,
                    read_u32_le_at(self.extra, self.extra_offset + 8)?,
                    read_u32_le_at(self.extra, self.extra_offset + 12)?,
                ];
                self.extra_offset += 16;
                let invalid = range == [1_048_575, 1_048_575, 16_383, 16_383];
                if !invalid {
                    FormulaRange::new(range[0], range[1], range[2], range[3])?;
                }
                cached_ranges.push(range);
            }
        }
        Ok(FormulaToken::Memory {
            kind,
            expression_bytes,
            cached_ranges,
        })
    }

    /// Parse cell reference
    fn parse_ref(&mut self, offset_reference: bool) -> Result<FormulaToken> {
        self.validate_classed_token("PtgRef")?;
        self.require(6, "PtgRef")?;

        let row_data = read_u32_le_at(self.data, self.offset)?;
        let col_data = read_u16_le_at(self.data, self.offset + 4)?;
        self.offset += 6;

        // Extract row and column (with relative flags)
        let col_relative = (col_data & 0x4000) != 0;
        let row_relative = (col_data & 0x8000) != 0;
        let (row, col) = self.resolve_reference(
            row_data,
            col_data & 0x3FFF,
            row_relative,
            col_relative,
            offset_reference,
        )?;

        Ok(FormulaToken::CellRef {
            row,
            col,
            row_relative,
            col_relative,
        })
    }

    /// Parse area reference
    fn parse_area(&mut self, offset_reference: bool) -> Result<FormulaToken> {
        self.validate_classed_token("PtgArea")?;
        self.require(12, "PtgArea")?;

        let row_first_data = read_u32_le_at(self.data, self.offset)?;
        let row_last_data = read_u32_le_at(self.data, self.offset + 4)?;
        let col_first_data = read_u16_le_at(self.data, self.offset + 8)?;
        let col_last_data = read_u16_le_at(self.data, self.offset + 10)?;
        self.offset += 12;

        let col_first_relative = (col_first_data & 0x4000) != 0;
        let row_first_relative = (col_first_data & 0x8000) != 0;
        let col_last_relative = (col_last_data & 0x4000) != 0;
        let row_last_relative = (col_last_data & 0x8000) != 0;
        let (row_first, col_first) = self.resolve_reference(
            row_first_data,
            col_first_data & 0x3FFF,
            row_first_relative,
            col_first_relative,
            offset_reference,
        )?;
        let (row_last, col_last) = self.resolve_reference(
            row_last_data,
            col_last_data & 0x3FFF,
            row_last_relative,
            col_last_relative,
            offset_reference,
        )?;
        FormulaRange::new(row_first, row_last, col_first, col_last)?;

        Ok(FormulaToken::AreaRef {
            row_first,
            row_last,
            col_first,
            col_last,
            row_first_relative,
            row_last_relative,
            col_first_relative,
            col_last_relative,
        })
    }

    fn parse_reference_error(&mut self, is_area: bool, is_3d: bool) -> Result<FormulaToken> {
        let token = self.data[self.offset - 1];
        if token & 0x80 != 0 {
            return Err(Error::InvalidFormula(format!(
                "reference-error token 0x{token:02X} has its reserved bit set"
            )));
        }
        let sheet_index = if is_3d {
            self.require(2, "3D reference-error sheet index")?;
            let index = read_u16_le_at(self.data, self.offset)?;
            self.offset += 2;
            Some(index)
        } else {
            None
        };
        let unused_len = if is_area { 12 } else { 6 };
        self.require(unused_len, "reference-error payload")?;
        self.offset += unused_len;
        Ok(FormulaToken::ReferenceError {
            is_area,
            sheet_index,
        })
    }

    fn parse_ref_3d(&mut self) -> Result<FormulaToken> {
        let token = self.data[self.offset - 1];
        if token & 0x80 != 0 {
            return Err(Error::InvalidFormula(format!(
                "PtgRef3d token 0x{token:02X} has its reserved bit set"
            )));
        }
        self.require(8, "PtgRef3d")?;
        let sheet_index = read_u16_le_at(self.data, self.offset)?;
        let row = read_u32_le_at(self.data, self.offset + 2)?;
        let col_data = read_u16_le_at(self.data, self.offset + 6)?;
        self.offset += 8;
        let col = u32::from(col_data & 0x3FFF);
        if row >= 1_048_576 || col >= 16_384 {
            return Err(Error::InvalidFormula(format!(
                "PtgRef3d coordinate ({row}, {col}) is outside the worksheet"
            )));
        }
        Ok(FormulaToken::CellRef3d {
            sheet_index,
            row,
            col,
            row_relative: col_data & 0x8000 != 0,
            col_relative: col_data & 0x4000 != 0,
        })
    }

    fn parse_area_3d(&mut self) -> Result<FormulaToken> {
        let token = self.data[self.offset - 1];
        if token & 0x80 != 0 {
            return Err(Error::InvalidFormula(format!(
                "PtgArea3d token 0x{token:02X} has its reserved bit set"
            )));
        }
        self.require(14, "PtgArea3d")?;
        let sheet_index = read_u16_le_at(self.data, self.offset)?;
        let row_first = read_u32_le_at(self.data, self.offset + 2)?;
        let row_last = read_u32_le_at(self.data, self.offset + 6)?;
        let col_first_data = read_u16_le_at(self.data, self.offset + 10)?;
        let col_last_data = read_u16_le_at(self.data, self.offset + 12)?;
        self.offset += 14;
        let col_first = u32::from(col_first_data & 0x3FFF);
        let col_last = u32::from(col_last_data & 0x3FFF);
        FormulaRange::new(row_first, row_last, col_first, col_last)?;
        Ok(FormulaToken::AreaRef3d {
            sheet_index,
            row_first,
            row_last,
            col_first,
            col_last,
            row_first_relative: col_first_data & 0x8000 != 0,
            row_last_relative: col_last_data & 0x8000 != 0,
            col_first_relative: col_first_data & 0x4000 != 0,
            col_last_relative: col_last_data & 0x4000 != 0,
        })
    }

    fn resolve_reference(
        &self,
        row_data: u32,
        col_data: u16,
        row_relative: bool,
        col_relative: bool,
        offset_reference: bool,
    ) -> Result<(u32, u32)> {
        if !offset_reference {
            if row_data >= 1_048_576 || col_data >= 16_384 {
                return Err(Error::InvalidFormula(format!(
                    "reference ({row_data}, {col_data}) is outside the worksheet"
                )));
            }
            return Ok((row_data, u32::from(col_data)));
        }
        let (base_row, base_col) = self.base_cell.ok_or_else(|| {
            Error::InvalidFormula(
                "PtgRefN/PtgAreaN requires a target cell for offset resolution".to_string(),
            )
        })?;

        let row = if row_relative {
            add_wrapped_offset(base_row, row_data as i32, 1_048_576)
        } else {
            row_data
        };
        let col = if col_relative {
            let signed = if col_data & 0x2000 != 0 {
                i32::from(col_data) - 0x4000
            } else {
                i32::from(col_data)
            };
            add_wrapped_offset(base_col, signed, 16_384)
        } else {
            u32::from(col_data)
        };
        if row >= 1_048_576 || col >= 16_384 {
            return Err(Error::InvalidFormula(format!(
                "resolved reference ({row}, {col}) is outside the worksheet"
            )));
        }
        Ok((row, col))
    }

    /// Parse function with fixed arguments
    fn parse_func(&mut self) -> Result<FormulaToken> {
        self.validate_classed_token("PtgFunc")?;
        self.require(2, "PtgFunc")?;

        let index = read_u16_le_at(self.data, self.offset)?;
        self.offset += 2;

        let arg_count = Self::get_function_arg_count(index)?;

        Ok(FormulaToken::Function {
            index,
            arg_count,
            is_command: false,
        })
    }

    /// Parse function with variable arguments
    fn parse_func_var(&mut self) -> Result<FormulaToken> {
        self.validate_classed_token("PtgFuncVar")?;
        self.require(3, "PtgFuncVar")?;

        let arg_count = self.data[self.offset];
        let tab = read_u16_le_at(self.data, self.offset + 1)?;
        let index = tab & 0x7FFF;
        let is_command = tab & 0x8000 != 0;
        self.offset += 3;

        if !is_command && let Some(function) = builtin_function_by_index(index) {
            if function.min_args == function.max_args {
                return Err(Error::InvalidFormula(format!(
                    "PtgFuncVar specifies fixed-arity function {}",
                    function.name
                )));
            }
            if !function.accepts_arg_count(arg_count) {
                return Err(Error::InvalidFormula(format!(
                    "{} does not accept {arg_count} arguments",
                    function.name
                )));
            }
        }

        Ok(FormulaToken::Function {
            index,
            arg_count,
            is_command,
        })
    }

    /// Parse defined name reference
    fn parse_name(&mut self) -> Result<FormulaToken> {
        self.validate_classed_token("PtgName")?;
        self.require(4, "PtgName")?;

        let name_index = read_u32_le_at(self.data, self.offset)?;
        self.offset += 4;
        if name_index == 0 {
            return Err(Error::InvalidFormula(
                "PtgName index is one-based and cannot be zero".to_string(),
            ));
        }

        Ok(FormulaToken::Name(name_index))
    }

    fn parse_name_x(&mut self) -> Result<FormulaToken> {
        self.validate_classed_token("PtgNameX")?;
        self.require(6, "PtgNameX")?;
        let sheet_index = read_u16_le_at(self.data, self.offset)?;
        let name_index = read_u32_le_at(self.data, self.offset + 2)?;
        self.offset += 6;
        if name_index == 0 {
            return Err(Error::InvalidFormula(
                "PtgNameX name index is one-based and cannot be zero".to_string(),
            ));
        }
        Ok(FormulaToken::ExternalName {
            sheet_index,
            name_index,
        })
    }

    /// Resolve the parameter count of a fixed-arity `Ftab` function.
    fn get_function_arg_count(index: u16) -> Result<u8> {
        let function = builtin_function_by_index(index).ok_or_else(|| {
            Error::UnsupportedFeature(format!(
                "fixed-arity XLSB function index {index} is not a non-macro Ftab entry"
            ))
        })?;
        if function.min_args != function.max_args {
            return Err(Error::InvalidFormula(format!(
                "PtgFunc specifies variable-arity function {}",
                function.name
            )));
        }
        Ok(function.min_args)
    }
}

/// Context-dependent name and relationship resolution supplied by a host.
///
/// The owner codec never discovers package parts or workbook relationships;
/// it only asks the host for already-resolved formula names and prefixes.
pub trait FormulaResolution {
    fn sheet_prefix(&self, index: u16) -> Result<String>;
    fn defined_name(&self, index: u32) -> Result<String>;
    fn external_name(&self, sheet_index: u16, name_index: u32) -> Result<String>;
    fn table_reference(&self, reference: &FormulaTableReference) -> Result<String>;
    fn pivot_name(&self, index: u32) -> Result<String>;
}

pub struct FormulaConverter;

impl FormulaConverter {
    /// Convert formula tokens to string representation
    ///
    /// Uses RPN to infix conversion with proper operator precedence.
    pub fn tokens_to_string(tokens: &[FormulaToken]) -> String {
        Self::try_tokens_to_string(tokens).unwrap_or_default()
    }

    /// Convert tokens to text, rejecting token streams that cannot be
    /// represented faithfully by this converter.
    pub fn try_tokens_to_string(tokens: &[FormulaToken]) -> Result<String> {
        Self::try_tokens_to_string_with_optional_context(tokens, None)
    }

    /// Convert formula tokens using workbook extern-sheet and name metadata.
    pub fn try_tokens_to_string_with_resolution(
        tokens: &[FormulaToken],
        context: &dyn FormulaResolution,
    ) -> Result<String> {
        Self::try_tokens_to_string_with_optional_context(tokens, Some(context))
    }

    fn try_tokens_to_string_with_optional_context(
        tokens: &[FormulaToken],
        context: Option<&dyn FormulaResolution>,
    ) -> Result<String> {
        let mut stack: Vec<String> = Vec::new();

        for token in tokens {
            match token {
                FormulaToken::Number(n) => stack.push(format!("{}", n)),
                FormulaToken::Int(i) => stack.push(format!("{}", i)),
                FormulaToken::MissingArg => stack.push(String::new()),
                FormulaToken::Parenthesis => {
                    let Some(expression) = stack.pop() else {
                        return Err(Error::InvalidFormula(
                            "PtgParen has no preceding expression".to_string(),
                        ));
                    };
                    stack.push(format!("({expression})"));
                },
                FormulaToken::Attribute(_) => {},
                FormulaToken::Array { rows, cols, values } => {
                    let expected = usize::try_from(u64::from(*rows) * u64::from(*cols))
                        .map_err(|_| Error::InvalidFormula("array is too large".to_string()))?;
                    if values.len() != expected {
                        return Err(Error::InvalidFormula(format!(
                            "array dimensions require {expected} values, found {}",
                            values.len()
                        )));
                    }
                    let mut text = String::from("{");
                    for row in 0..*rows {
                        if row != 0 {
                            text.push(';');
                        }
                        for col in 0..*cols {
                            if col != 0 {
                                text.push(',');
                            }
                            let index =
                                usize::try_from(u64::from(row) * u64::from(*cols) + u64::from(col))
                                    .map_err(|_| {
                                        Error::InvalidFormula("array index overflow".to_string())
                                    })?;
                            match &values[index] {
                                FormulaArrayValue::Number(value) => {
                                    text.push_str(&value.to_string());
                                },
                                FormulaArrayValue::String(value) => {
                                    text.push('"');
                                    text.push_str(&value.replace('"', "\"\""));
                                    text.push('"');
                                },
                                FormulaArrayValue::Bool(value) => {
                                    text.push_str(if *value { "TRUE" } else { "FALSE" });
                                },
                                FormulaArrayValue::Error(error) => {
                                    text.push_str(&Self::error_to_string(*error));
                                },
                            }
                        }
                    }
                    text.push('}');
                    stack.push(text);
                },
                FormulaToken::Memory { .. } => {},
                FormulaToken::String(s) => stack.push(format!("\"{}\"", s.replace('"', "\"\""))),
                FormulaToken::Bool(b) => stack.push(if *b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }),
                FormulaToken::Error(e) => stack.push(Self::error_to_string(*e)),
                FormulaToken::CellRef {
                    row,
                    col,
                    row_relative,
                    col_relative,
                } => {
                    let col_str = column_index_to_name(*col + 1);
                    let row_str = row + 1;
                    let col_prefix = if *col_relative { "" } else { "$" };
                    let row_prefix = if *row_relative { "" } else { "$" };
                    stack.push(format!(
                        "{}{}{}{}",
                        col_prefix, col_str, row_prefix, row_str
                    ));
                },
                FormulaToken::AreaRef {
                    row_first,
                    col_first,
                    row_last,
                    col_last,
                    row_first_relative,
                    row_last_relative,
                    col_first_relative,
                    col_last_relative,
                } => {
                    let first = Self::format_reference(
                        *row_first,
                        *col_first,
                        *row_first_relative,
                        *col_first_relative,
                    );
                    let last = Self::format_reference(
                        *row_last,
                        *col_last,
                        *row_last_relative,
                        *col_last_relative,
                    );
                    stack.push(format!("{}:{}", first, last));
                },
                FormulaToken::CellRef3d {
                    sheet_index,
                    row,
                    col,
                    row_relative,
                    col_relative,
                } => {
                    let context = context.ok_or_else(|| {
                        Error::UnsupportedFeature(
                            "PtgRef3d requires workbook extern-sheet resolution".to_string(),
                        )
                    })?;
                    let prefix = context.sheet_prefix(*sheet_index)?;
                    let reference =
                        Self::format_reference(*row, *col, *row_relative, *col_relative);
                    stack.push(format!("{prefix}!{reference}"));
                },
                FormulaToken::AreaRef3d {
                    sheet_index,
                    row_first,
                    row_last,
                    col_first,
                    col_last,
                    row_first_relative,
                    row_last_relative,
                    col_first_relative,
                    col_last_relative,
                } => {
                    let context = context.ok_or_else(|| {
                        Error::UnsupportedFeature(
                            "PtgArea3d requires workbook extern-sheet resolution".to_string(),
                        )
                    })?;
                    let prefix = context.sheet_prefix(*sheet_index)?;
                    let first = Self::format_reference(
                        *row_first,
                        *col_first,
                        *row_first_relative,
                        *col_first_relative,
                    );
                    let last = Self::format_reference(
                        *row_last,
                        *col_last,
                        *row_last_relative,
                        *col_last_relative,
                    );
                    stack.push(format!("{prefix}!{first}:{last}"));
                },
                FormulaToken::ReferenceError { .. } => stack.push("#REF!".to_string()),
                FormulaToken::BinaryOp(op) => {
                    if stack.len() < 2 {
                        return Err(Error::InvalidFormula(
                            "binary operator has fewer than two operands".to_string(),
                        ));
                    }
                    let right = stack.pop().expect("length checked");
                    let left = stack.pop().expect("length checked");
                    let op_str = Self::binary_op_to_string(*op);
                    stack.push(format!("({}{}{})", left, op_str, right));
                },
                FormulaToken::UnaryOp(op) => {
                    let Some(operand) = stack.pop() else {
                        return Err(Error::InvalidFormula(
                            "unary operator has no operand".to_string(),
                        ));
                    };
                    match op {
                        UnaryOperator::Plus => stack.push(format!("+({})", operand)),
                        UnaryOperator::Minus => stack.push(format!("-({})", operand)),
                        UnaryOperator::Percent => stack.push(format!("({}%)", operand)),
                    }
                },
                FormulaToken::Function {
                    index,
                    arg_count,
                    is_command,
                } => {
                    if *is_command {
                        return Err(Error::UnsupportedFeature(format!(
                            "XLSB command function index {index}"
                        )));
                    }
                    let Some(function) = builtin_function_by_index(*index) else {
                        return Err(Error::UnsupportedFeature(format!(
                            "XLSB built-in function index {index}"
                        )));
                    };
                    let func_name = function.name;
                    if stack.len() < usize::from(*arg_count) {
                        return Err(Error::InvalidFormula(format!(
                            "function {func_name} requires {arg_count} stack operands"
                        )));
                    }
                    let mut args = Vec::new();
                    for _ in 0..*arg_count {
                        if let Some(arg) = stack.pop() {
                            args.insert(0, arg);
                        }
                    }
                    stack.push(format!("{}({})", func_name, args.join(",")));
                },
                FormulaToken::Name(idx) => {
                    let context = context.ok_or_else(|| {
                        Error::UnsupportedFeature(format!(
                            "XLSB defined name index {idx} requires workbook name resolution"
                        ))
                    })?;
                    stack.push(context.defined_name(*idx)?);
                },
                FormulaToken::ExternalName {
                    sheet_index,
                    name_index,
                } => {
                    let context = context.ok_or_else(|| {
                        Error::UnsupportedFeature(
                            "PtgNameX requires workbook external-link resolution".to_string(),
                        )
                    })?;
                    stack.push(context.external_name(*sheet_index, *name_index)?);
                },
                FormulaToken::TableReference(reference) if reference.invalid => {
                    stack.push("#REF!".to_string())
                },
                FormulaToken::TableReference(reference) => {
                    let context = context.ok_or_else(|| {
                        Error::UnsupportedFeature(format!(
                            "structured table reference on Xti {} requires table-definition resolution",
                            reference.sheet_index
                        ))
                    })?;
                    stack.push(context.table_reference(reference)?);
                },
                FormulaToken::PivotName(index) => {
                    let context = context.ok_or_else(|| {
                        Error::InvalidFormula(
                            "PtgSxName requires pivot-cache calculated-name metadata".to_string(),
                        )
                    })?;
                    stack.push(context.pivot_name(*index)?);
                },
                FormulaToken::Unknown(t) => {
                    return Err(Error::UnsupportedFeature(format!(
                        "XLSB formula token 0x{t:02X}"
                    )));
                },
            }
        }

        if stack.len() != 1 {
            return Err(Error::InvalidFormula(format!(
                "formula leaves {} values on the evaluation stack",
                stack.len()
            )));
        }
        Ok(stack.pop().expect("length checked"))
    }

    fn format_reference(row: u32, col: u32, row_relative: bool, col_relative: bool) -> String {
        let col_str = column_index_to_name(col + 1);
        format!(
            "{}{}{}{}",
            if col_relative { "" } else { "$" },
            col_str,
            if row_relative { "" } else { "$" },
            row + 1
        )
    }

    /// Convert binary operator to string
    fn binary_op_to_string(op: BinaryOperator) -> &'static str {
        match op {
            BinaryOperator::Add => "+",
            BinaryOperator::Subtract => "-",
            BinaryOperator::Multiply => "*",
            BinaryOperator::Divide => "/",
            BinaryOperator::Power => "^",
            BinaryOperator::Concat => "&",
            BinaryOperator::LessThan => "<",
            BinaryOperator::LessEqual => "<=",
            BinaryOperator::Equal => "=",
            BinaryOperator::GreaterEqual => ">=",
            BinaryOperator::GreaterThan => ">",
            BinaryOperator::NotEqual => "<>",
            BinaryOperator::Intersection => " ",
            BinaryOperator::Union => ",",
            BinaryOperator::Range => ":",
        }
    }

    /// Convert error code to string
    fn error_to_string(code: u8) -> String {
        match code {
            0x00 => "#NULL!".to_string(),
            0x07 => "#DIV/0!".to_string(),
            0x0F => "#VALUE!".to_string(),
            0x17 => "#REF!".to_string(),
            0x1D => "#NAME?".to_string(),
            0x24 => "#NUM!".to_string(),
            0x2A => "#N/A".to_string(),
            0x2B => "#GETTING_DATA".to_string(),
            _ => format!("#ERR{:02X}!", code),
        }
    }
}

fn parse_table_row_type(value: u8) -> Result<FormulaTableRowType> {
    match value {
        0x00 => Ok(FormulaTableRowType::Data),
        0x01 => Ok(FormulaTableRowType::All),
        0x02 => Ok(FormulaTableRowType::Headers),
        0x04 => Ok(FormulaTableRowType::DataAlternate),
        0x06 => Ok(FormulaTableRowType::DataAndHeaders),
        0x08 => Ok(FormulaTableRowType::Totals),
        0x0C => Ok(FormulaTableRowType::DataAndTotals),
        0x10 => Ok(FormulaTableRowType::Current),
        _ => Err(Error::InvalidFormula(format!(
            "invalid PtgRowType 0x{value:02X}"
        ))),
    }
}

fn table_row_type_raw(value: FormulaTableRowType) -> u8 {
    match value {
        FormulaTableRowType::Data => 0x00,
        FormulaTableRowType::All => 0x01,
        FormulaTableRowType::Headers => 0x02,
        FormulaTableRowType::DataAlternate => 0x04,
        FormulaTableRowType::DataAndHeaders => 0x06,
        FormulaTableRowType::Totals => 0x08,
        FormulaTableRowType::DataAndTotals => 0x0C,
        FormulaTableRowType::Current => 0x10,
    }
}

#[derive(Debug, Clone, Copy)]
struct BuiltinFunction {
    index: u16,
    name: &'static str,
    min_args: u8,
    max_args: u8,
}

impl BuiltinFunction {
    fn accepts_arg_count(self, count: u8) -> bool {
        if count < self.min_args || count > self.max_args {
            return false;
        }
        match self.index {
            // GETPIVOTDATA permits the two mandatory arguments, one optional
            // field, or complete field/item pairs thereafter.
            358 => count <= 3 || count.is_multiple_of(2),
            // COUNTIFS is made solely of range/criteria pairs.
            481 => count.is_multiple_of(2),
            // SUMIFS and AVERAGEIFS have one leading aggregate range followed
            // by range/criteria pairs.
            482 | 484 => !count.is_multiple_of(2),
            _ => true,
        }
    }
}

fn builtin_function_by_index(index: u16) -> Option<BuiltinFunction> {
    let position = BUILTIN_FUNCTIONS
        .binary_search_by_key(&index, |entry| entry.0)
        .ok()?;
    let (index, name, min_args, max_args) = BUILTIN_FUNCTIONS[position];
    Some(BuiltinFunction {
        index,
        name,
        min_args,
        max_args,
    })
}

fn validate_xnum(value: f64, context: &str) -> Result<()> {
    if !value.is_finite()
        || (value == 0.0 && value.is_sign_negative())
        || (value != 0.0 && !value.is_normal())
    {
        return Err(Error::InvalidFormula(format!(
            "{context} contains a non-finite, denormalized, or negative-zero Xnum"
        )));
    }
    Ok(())
}

const FORMULA_ERRORS: &[(&str, u8)] = &[
    ("#GETTING_DATA", 0x2B),
    ("#DIV/0!", 0x07),
    ("#VALUE!", 0x0F),
    ("#NULL!", 0x00),
    ("#NAME?", 0x1D),
    ("#REF!", 0x17),
    ("#NUM!", 0x24),
    ("#N/A", 0x2A),
];

fn is_formula_error_code(value: u8) -> bool {
    FORMULA_ERRORS.iter().any(|(_, code)| *code == value)
}

fn add_wrapped_offset(base: u32, offset: i32, modulus: u32) -> u32 {
    (i64::from(base) + i64::from(offset)).rem_euclid(i64::from(modulus)) as u32
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_and_serializes_cell_formula_lengths() {
        // [MS-XLSB] 2.5.98.4: cce, rgce, cb, and rgbExtra.
        let formula = CellParsedFormula {
            rgce: vec![ptg_types::PTG_INT, 42, 0],
            rgcb: vec![0xAA, 0xBB],
        };
        let encoded = formula.to_bytes().unwrap();
        let (decoded, consumed) = CellParsedFormula::parse(&encoded).unwrap();
        assert_eq!(consumed, encoded.len());
        assert_eq!(decoded, formula);
    }

    #[test]
    fn parses_scalar_ptgs_and_preserves_unknown_tokens() {
        // [MS-XLSB] 2.5.98.34, 2.5.98.63, and the extensible Ptg prefix.
        let mut data = vec![ptg_types::PTG_BOOL, 1, ptg_types::PTG_NUM];
        data.extend_from_slice(&42.5_f64.to_le_bytes());
        data.extend_from_slice(&[0x7F, 0x7E]);
        let mut parser = FormulaParser::new(&data);
        let tokens = parser.parse().unwrap();
        assert!(matches!(tokens[0], FormulaToken::Bool(true)));
        assert!(matches!(tokens[1], FormulaToken::Number(value) if value == 42.5));
        assert!(matches!(tokens[2], FormulaToken::Unknown(0x7F)));
        assert!(matches!(tokens[3], FormulaToken::Unknown(0x7E)));
    }

    #[test]
    fn parses_and_rejects_grouped_formula_records_without_unbounded_allocations() {
        let formula = CellParsedFormula::exp(3, 4).unwrap();
        let group = FormulaGroup {
            kind: FormulaGroupKind::Array,
            range: FormulaRange::new(3, 3, 4, 4).unwrap(),
            formula,
            always_calculate: true,
        };
        let data = group.to_record_data().unwrap();
        assert_eq!(FormulaGroup::parse_array(&data).unwrap(), group);

        let mut oversized = vec![0u8; MAX_CELL_FORMULA_BYTES + 1 + 8];
        oversized[..4].copy_from_slice(
            &u32::try_from(MAX_CELL_FORMULA_BYTES + 1)
                .unwrap()
                .to_le_bytes(),
        );
        assert!(matches!(
            CellParsedFormula::parse(&oversized),
            Err(Error::InvalidFormula(message)) if message.contains("exceeds")
        ));
    }

    #[test]
    fn extended_table_token_round_trips() {
        let token = FormulaToken::TableReference(FormulaTableReference {
            sheet_index: 2,
            row_type: Some(FormulaTableRowType::Data),
            columns: Some(FormulaTableColumns::One(1)),
            square_bracket_space: false,
            comma_space: true,
            data_type: FormulaTableDataType::Reference,
            invalid: false,
            list_index: Some(7),
            external: None,
        });
        let (rgce, rgcb) = token.to_extended_binary().unwrap();
        let mut parser = FormulaParser::with_extra(&rgce, &rgcb);
        assert_eq!(parser.parse().unwrap(), vec![token]);
    }
}
