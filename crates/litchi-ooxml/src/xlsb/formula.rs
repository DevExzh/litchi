//! XLSB formula parsing and generation
//!
//! Excel formulas in XLSB files are stored in a binary format using Reverse Polish Notation (RPN)
//! with Parse Tree Generators (Ptg tokens). This module provides parsing and generation of formulas.
//!
//! # Formula Token Types (Ptgs)
//!
//! Formulas are sequences of tokens that represent operands, operators, and functions:
//! - **Value tokens**: Numbers, strings, booleans, errors
//! - **Operand tokens**: Cell references, ranges, names
//! - **Operator tokens**: Add, subtract, multiply, divide, etc.
//! - **Function tokens**: SUM, IF, VLOOKUP, etc.
//!
//! # Binary Format
//!
//! Each token consists of:
//! 1. Token type byte (identifies the Ptg)
//! 2. Token data (variable length, depends on token type)
//!
//! # Reference
//!
//! - [MS-XLSB] Section 2.5.98 - Formulas
//! - [MS-XLS] Section 2.5.198 - Ptg (for token details, largely compatible)

use crate::xlsb::error::{XlsbError, XlsbResult};
use litchi_core::binary;

/// Maximum size of an XLSB cell formula token stream.
///
/// [MS-XLSB] 2.5.98.4 requires `cce` to be greater than zero and less than
/// 16,385 bytes.
pub const MAX_CELL_FORMULA_BYTES: usize = 16_384;

/// Inclusive worksheet range used by array and shared formulas.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaRange {
    pub row_first: u32,
    pub row_last: u32,
    pub col_first: u32,
    pub col_last: u32,
}

impl FormulaRange {
    pub fn new(row_first: u32, row_last: u32, col_first: u32, col_last: u32) -> XlsbResult<Self> {
        let range = Self {
            row_first,
            row_last,
            col_first,
            col_last,
        };
        range.validate()?;
        Ok(range)
    }

    pub fn parse_a1(value: &str) -> XlsbResult<Self> {
        let (first, last) = value.split_once(':').unwrap_or((value, value));
        let (row_first, col_first) = crate::xlsb::utils::parse_cell_reference(first.trim())?;
        let (row_last, col_last) = crate::xlsb::utils::parse_cell_reference(last.trim())?;
        Self::new(row_first, row_last, col_first, col_last)
    }

    pub fn parse_binary(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 16 {
            return Err(XlsbError::InvalidLength {
                expected: 16,
                found: data.len(),
            });
        }
        Self::new(
            binary::read_u32_le_at(data, 0)?,
            binary::read_u32_le_at(data, 4)?,
            binary::read_u32_le_at(data, 8)?,
            binary::read_u32_le_at(data, 12)?,
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
            crate::xlsb::utils::cell_reference(self.row_first, self.col_first),
            crate::xlsb::utils::cell_reference(self.row_last, self.col_last)
        )
    }

    fn validate(self) -> XlsbResult<()> {
        if self.row_first > self.row_last
            || self.col_first > self.col_last
            || self.row_last >= 1_048_576
            || self.col_last >= 16_384
        {
            return Err(XlsbError::InvalidCellReference(self.to_a1()));
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

impl CellParsedFormula {
    /// Parse a `CellParsedFormula`, returning the structure and bytes consumed.
    pub fn parse(data: &[u8]) -> XlsbResult<(Self, usize)> {
        if data.len() < 4 {
            return Err(XlsbError::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }

        let cce = binary::read_u32_le_at(data, 0)? as usize;
        if cce == 0 || cce > MAX_CELL_FORMULA_BYTES {
            return Err(XlsbError::InvalidFormula(format!(
                "cell formula token length {cce} is outside 1..={MAX_CELL_FORMULA_BYTES}"
            )));
        }
        let cb_offset = 4usize.checked_add(cce).ok_or_else(|| {
            XlsbError::InvalidFormula("cell formula token length overflow".to_string())
        })?;
        if data.len() < cb_offset + 4 {
            return Err(XlsbError::InvalidLength {
                expected: cb_offset + 4,
                found: data.len(),
            });
        }

        let cb = binary::read_u32_le_at(data, cb_offset)? as usize;
        let end = cb_offset
            .checked_add(4)
            .and_then(|offset| offset.checked_add(cb))
            .ok_or_else(|| {
                XlsbError::InvalidFormula("cell formula ancillary length overflow".to_string())
            })?;
        if data.len() < end {
            return Err(XlsbError::InvalidLength {
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
    pub fn to_bytes(&self) -> XlsbResult<Vec<u8>> {
        if self.rgce.is_empty() || self.rgce.len() > MAX_CELL_FORMULA_BYTES {
            return Err(XlsbError::InvalidFormula(format!(
                "cell formula token length {} is outside 1..={MAX_CELL_FORMULA_BYTES}",
                self.rgce.len()
            )));
        }
        let cce = u32::try_from(self.rgce.len())
            .map_err(|_| XlsbError::InvalidFormula("formula is too large".to_string()))?;
        let cb = u32::try_from(self.rgcb.len()).map_err(|_| {
            XlsbError::InvalidFormula("formula ancillary data is too large".to_string())
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
    pub fn exp(row: u32, col: u32) -> XlsbResult<Self> {
        if row >= 1_048_576 || col >= 16_384 {
            return Err(XlsbError::InvalidCellReference(format!(
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
    pub fn exp_cell(&self) -> XlsbResult<Option<(u32, u32)>> {
        if self.rgce.first() != Some(&ptg_types::PTG_EXP) {
            return Ok(None);
        }
        if self.rgce.len() != 5 || self.rgcb.len() != 4 {
            return Err(XlsbError::InvalidFormula(format!(
                "PtgExp requires 5 rgce bytes and 4 rgcb bytes, found {} and {}",
                self.rgce.len(),
                self.rgcb.len()
            )));
        }
        let row = binary::read_u32_le_at(&self.rgce, 1)?;
        let col = binary::read_u32_le_at(&self.rgcb, 0)?;
        if row >= 1_048_576 || col >= 16_384 {
            return Err(XlsbError::InvalidCellReference(format!(
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
    pub fn parse_array(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 17 {
            return Err(XlsbError::InvalidLength {
                expected: 17,
                found: data.len(),
            });
        }
        if data[16] & !1 != 0 {
            return Err(XlsbError::InvalidFormula(format!(
                "BrtArrFmla has reserved flag bits 0x{:02X}",
                data[16] & !1
            )));
        }
        let range = FormulaRange::parse_binary(data)?;
        let (formula, consumed) = CellParsedFormula::parse(&data[17..])?;
        if 17 + consumed != data.len() {
            return Err(XlsbError::InvalidFormula(format!(
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

    pub fn parse_shared(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 16 {
            return Err(XlsbError::InvalidLength {
                expected: 16,
                found: data.len(),
            });
        }
        let range = FormulaRange::parse_binary(data)?;
        let (formula, consumed) = CellParsedFormula::parse(&data[16..])?;
        if 16 + consumed != data.len() {
            return Err(XlsbError::InvalidFormula(format!(
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

    pub fn to_record_data(&self) -> XlsbResult<Vec<u8>> {
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
    /// Unknown/unsupported token
    Unknown(u8),
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
    base_cell: Option<(u32, u32)>,
}

impl<'a> FormulaParser<'a> {
    /// Create a new formula parser
    pub fn new(data: &'a [u8]) -> Self {
        FormulaParser {
            data,
            offset: 0,
            base_cell: None,
        }
    }

    /// Parse a shared-formula definition relative to a concrete target cell.
    pub fn with_base_cell(data: &'a [u8], row: u32, col: u32) -> Self {
        FormulaParser {
            data,
            offset: 0,
            base_cell: Some((row, col)),
        }
    }

    /// Parse the formula into tokens
    ///
    /// Returns a vector of formula tokens in RPN order.
    pub fn parse(&mut self) -> XlsbResult<Vec<FormulaToken>> {
        let mut tokens = Vec::new();

        while self.offset < self.data.len() {
            tokens.push(self.parse_token()?);
        }

        Ok(tokens)
    }

    /// Parse a single token
    fn parse_token(&mut self) -> XlsbResult<FormulaToken> {
        if self.offset >= self.data.len() {
            return Err(XlsbError::InvalidFormula(
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
                0x01 => self.parse_func(),
                0x02 => self.parse_func_var(),
                0x03 => self.parse_name(),
                _ => Ok(FormulaToken::Unknown(ptg_type)),
            },

            _ => {
                // Unknown token type
                Ok(FormulaToken::Unknown(ptg_type))
            },
        }
    }

    fn require(&self, len: usize, context: &str) -> XlsbResult<()> {
        if self.offset + len <= self.data.len() {
            Ok(())
        } else {
            Err(XlsbError::InvalidFormula(format!(
                "truncated {context} token at byte {}: need {len} bytes, have {}",
                self.offset.saturating_sub(1),
                self.data.len().saturating_sub(self.offset)
            )))
        }
    }

    /// Parse integer constant
    fn parse_int(&mut self) -> XlsbResult<FormulaToken> {
        self.require(2, "PtgInt")?;

        let value = binary::read_u16_le_at(self.data, self.offset)?;
        self.offset += 2;

        Ok(FormulaToken::Int(value))
    }

    /// Parse floating point constant
    fn parse_num(&mut self) -> XlsbResult<FormulaToken> {
        self.require(8, "PtgNum")?;

        let value = binary::read_f64_le_at(self.data, self.offset)?;
        self.offset += 8;

        Ok(FormulaToken::Number(value))
    }

    /// Parse string constant
    fn parse_str(&mut self) -> XlsbResult<FormulaToken> {
        self.require(2, "PtgStr length")?;
        let len = binary::read_u16_le_at(self.data, self.offset)? as usize;
        self.offset += 2;
        let byte_len = len.checked_mul(2).ok_or_else(|| {
            XlsbError::InvalidFormula("PtgStr UTF-16 length overflow".to_string())
        })?;
        self.require(byte_len, "PtgStr text")?;
        let units = self.data[self.offset..self.offset + byte_len]
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]));
        let string = char::decode_utf16(units)
            .collect::<Result<String, _>>()
            .map_err(|_| XlsbError::Encoding("invalid UTF-16 in PtgStr".to_string()))?;
        self.offset += byte_len;

        Ok(FormulaToken::String(string))
    }

    /// Parse boolean constant
    fn parse_bool(&mut self) -> XlsbResult<FormulaToken> {
        self.require(1, "PtgBool")?;

        let value = self.data[self.offset] != 0;
        self.offset += 1;

        Ok(FormulaToken::Bool(value))
    }

    /// Parse error constant
    fn parse_err(&mut self) -> XlsbResult<FormulaToken> {
        self.require(1, "PtgErr")?;

        let error_code = self.data[self.offset];
        self.offset += 1;

        Ok(FormulaToken::Error(error_code))
    }

    /// Parse cell reference
    fn parse_ref(&mut self, offset_reference: bool) -> XlsbResult<FormulaToken> {
        self.require(6, "PtgRef")?;

        let row_data = binary::read_u32_le_at(self.data, self.offset)?;
        let col_data = binary::read_u16_le_at(self.data, self.offset + 4)?;
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
    fn parse_area(&mut self, offset_reference: bool) -> XlsbResult<FormulaToken> {
        self.require(12, "PtgArea")?;

        let row_first_data = binary::read_u32_le_at(self.data, self.offset)?;
        let row_last_data = binary::read_u32_le_at(self.data, self.offset + 4)?;
        let col_first_data = binary::read_u16_le_at(self.data, self.offset + 8)?;
        let col_last_data = binary::read_u16_le_at(self.data, self.offset + 10)?;
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

    fn resolve_reference(
        &self,
        row_data: u32,
        col_data: u16,
        row_relative: bool,
        col_relative: bool,
        offset_reference: bool,
    ) -> XlsbResult<(u32, u32)> {
        if !offset_reference {
            return Ok((row_data, u32::from(col_data)));
        }
        let (base_row, base_col) = self.base_cell.ok_or_else(|| {
            XlsbError::InvalidFormula(
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
            return Err(XlsbError::InvalidFormula(format!(
                "resolved reference ({row}, {col}) is outside the worksheet"
            )));
        }
        Ok((row, col))
    }

    /// Parse function with fixed arguments
    fn parse_func(&mut self) -> XlsbResult<FormulaToken> {
        self.require(2, "PtgFunc")?;

        let index = binary::read_u16_le_at(self.data, self.offset)?;
        self.offset += 2;

        // Look up argument count from function table (simplified)
        let arg_count = Self::get_function_arg_count(index);

        Ok(FormulaToken::Function {
            index,
            arg_count,
            is_command: false,
        })
    }

    /// Parse function with variable arguments
    fn parse_func_var(&mut self) -> XlsbResult<FormulaToken> {
        self.require(3, "PtgFuncVar")?;

        let arg_count = self.data[self.offset];
        let tab = binary::read_u16_le_at(self.data, self.offset + 1)?;
        let index = tab & 0x7FFF;
        let is_command = tab & 0x8000 != 0;
        self.offset += 3;

        Ok(FormulaToken::Function {
            index,
            arg_count,
            is_command,
        })
    }

    /// Parse defined name reference
    fn parse_name(&mut self) -> XlsbResult<FormulaToken> {
        self.require(4, "PtgName")?;

        let name_index = binary::read_u32_le_at(self.data, self.offset)?;
        self.offset += 4;

        Ok(FormulaToken::Name(name_index))
    }

    /// Get function argument count by function index
    ///
    /// This is a simplified lookup. In a complete implementation, this would
    /// use a comprehensive table of all Excel functions.
    fn get_function_arg_count(index: u16) -> u8 {
        builtin_function_by_index(index).map_or(0, |function| function.min_args)
    }
}

/// Formula converter - converts tokens to human-readable formula string
///
/// # Note
///
/// This is a simplified converter. A complete implementation would handle
/// all token types and produce fully accurate Excel formula syntax.
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
    pub fn try_tokens_to_string(tokens: &[FormulaToken]) -> XlsbResult<String> {
        let mut stack: Vec<String> = Vec::new();

        for token in tokens {
            match token {
                FormulaToken::Number(n) => stack.push(format!("{}", n)),
                FormulaToken::Int(i) => stack.push(format!("{}", i)),
                FormulaToken::MissingArg => stack.push(String::new()),
                FormulaToken::Parenthesis => {
                    let Some(expression) = stack.pop() else {
                        return Err(XlsbError::InvalidFormula(
                            "PtgParen has no preceding expression".to_string(),
                        ));
                    };
                    stack.push(format!("({expression})"));
                },
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
                    let col_str = crate::xlsb::utils::column_index_to_name(*col + 1);
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
                FormulaToken::BinaryOp(op) => {
                    if stack.len() < 2 {
                        return Err(XlsbError::InvalidFormula(
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
                        return Err(XlsbError::InvalidFormula(
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
                        return Err(XlsbError::UnsupportedFeature(format!(
                            "XLSB command function index {index}"
                        )));
                    }
                    let Some(function) = builtin_function_by_index(*index) else {
                        return Err(XlsbError::UnsupportedFeature(format!(
                            "XLSB built-in function index {index}"
                        )));
                    };
                    let func_name = function.name;
                    if stack.len() < usize::from(*arg_count) {
                        return Err(XlsbError::InvalidFormula(format!(
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
                    return Err(XlsbError::UnsupportedFeature(format!(
                        "XLSB defined name index {idx} requires workbook name resolution"
                    )));
                },
                FormulaToken::Unknown(t) => {
                    return Err(XlsbError::UnsupportedFeature(format!(
                        "XLSB formula token 0x{t:02X}"
                    )));
                },
            }
        }

        if stack.len() != 1 {
            return Err(XlsbError::InvalidFormula(format!(
                "formula leaves {} values on the evaluation stack",
                stack.len()
            )));
        }
        Ok(stack.pop().expect("length checked"))
    }

    fn format_reference(row: u32, col: u32, row_relative: bool, col_relative: bool) -> String {
        let col_str = crate::xlsb::utils::column_index_to_name(col + 1);
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

#[derive(Debug, Clone, Copy)]
struct BuiltinFunction {
    index: u16,
    name: &'static str,
    min_args: u8,
    max_args: u8,
}

fn builtin_function_by_index(index: u16) -> Option<BuiltinFunction> {
    let (name, min_args, max_args) = match index {
        0 => ("COUNT", 0, 30),
        1 => ("IF", 2, 3),
        2 => ("ISNA", 1, 1),
        3 => ("ISERROR", 1, 1),
        4 => ("SUM", 0, 30),
        5 => ("AVERAGE", 1, 30),
        6 => ("MIN", 1, 30),
        7 => ("MAX", 1, 30),
        8 => ("ROW", 0, 1),
        9 => ("COLUMN", 0, 1),
        10 => ("NA", 0, 0),
        11 => ("NPV", 2, 30),
        12 => ("STDEV", 1, 30),
        13 => ("DOLLAR", 1, 2),
        14 => ("FIXED", 2, 3),
        15 => ("SIN", 1, 1),
        16 => ("COS", 1, 1),
        17 => ("TAN", 1, 1),
        18 => ("ATAN", 1, 1),
        19 => ("PI", 0, 0),
        20 => ("SQRT", 1, 1),
        21 => ("EXP", 1, 1),
        22 => ("LN", 1, 1),
        23 => ("LOG10", 1, 1),
        24 => ("ABS", 1, 1),
        25 => ("INT", 1, 1),
        26 => ("SIGN", 1, 1),
        27 => ("ROUND", 2, 2),
        28 => ("LOOKUP", 2, 3),
        29 => ("INDEX", 2, 4),
        30 => ("REPT", 2, 2),
        31 => ("MID", 3, 3),
        32 => ("LEN", 1, 1),
        33 => ("VALUE", 1, 1),
        34 => ("TRUE", 0, 0),
        35 => ("FALSE", 0, 0),
        36 => ("AND", 1, 30),
        37 => ("OR", 1, 30),
        38 => ("NOT", 1, 1),
        39 => ("MOD", 2, 2),
        64 => ("MATCH", 2, 3),
        65 => ("DATE", 3, 3),
        74 => ("NOW", 0, 0),
        82 => ("SEARCH", 2, 3),
        100 => ("CHOOSE", 2, 30),
        101 => ("HLOOKUP", 3, 4),
        102 => ("VLOOKUP", 3, 4),
        115 => ("LEFT", 1, 2),
        116 => ("RIGHT", 1, 2),
        124 => ("FIND", 2, 3),
        221 => ("TODAY", 0, 0),
        336 => ("CONCATENATE", 0, 30),
        345 => ("SUMIF", 2, 3),
        346 => ("COUNTIF", 2, 2),
        _ => return None,
    };
    Some(BuiltinFunction {
        index,
        name,
        min_args,
        max_args,
    })
}

fn builtin_function_by_name(name: &str) -> Option<BuiltinFunction> {
    const INDICES: &[u16] = &[
        0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24,
        25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 64, 65, 74, 82, 100, 101, 102,
        115, 116, 124, 221, 336, 345, 346,
    ];
    INDICES.iter().find_map(|index| {
        let function = builtin_function_by_index(*index)?;
        function.name.eq_ignore_ascii_case(name).then_some(function)
    })
}

/// Compiles a practical, standards-defined subset of Excel formula text to
/// XLSB RPN tokens.
///
/// The compiler supports literals, A1 references and ranges, parentheses,
/// arithmetic/comparison/concatenation operators, percent, and the built-in
/// functions in this module's supported `Ftab` table. Unsupported constructs
/// return an error; they are never replaced by a cached value.
pub struct FormulaCompiler<'a> {
    input: &'a str,
    offset: usize,
}

#[derive(Debug, Clone, Copy)]
enum FormulaEncoding {
    Cell,
    Shared { base_row: u32, base_col: u32 },
}

#[derive(Debug)]
enum CompileExpr {
    Number(f64),
    String(String),
    Bool(bool),
    MissingArg,
    Parenthesized(Box<CompileExpr>),
    Ref(A1Reference),
    Area(A1Reference, A1Reference),
    Unary(UnaryOperator, Box<CompileExpr>),
    Binary(BinaryOperator, Box<CompileExpr>, Box<CompileExpr>),
    Function(BuiltinFunction, Vec<CompileExpr>),
}

#[derive(Debug, Clone, Copy)]
struct A1Reference {
    row: u32,
    col: u32,
    row_relative: bool,
    col_relative: bool,
}

impl<'a> FormulaCompiler<'a> {
    pub fn compile(formula: &'a str) -> XlsbResult<CellParsedFormula> {
        Self::compile_with_encoding(formula, FormulaEncoding::Cell)
    }

    /// Compile a shared formula, encoding relative A1 references as
    /// `PtgRefN`/`PtgAreaN` offsets from the first cell in the shared range.
    pub fn compile_shared(
        formula: &'a str,
        base_row: u32,
        base_col: u32,
    ) -> XlsbResult<CellParsedFormula> {
        if base_row >= 1_048_576 || base_col >= 16_384 {
            return Err(XlsbError::InvalidCellReference(format!(
                "shared formula base ({base_row}, {base_col})"
            )));
        }
        Self::compile_with_encoding(formula, FormulaEncoding::Shared { base_row, base_col })
    }

    fn compile_with_encoding(
        formula: &'a str,
        encoding: FormulaEncoding,
    ) -> XlsbResult<CellParsedFormula> {
        let input = formula.strip_prefix('=').unwrap_or(formula).trim();
        if input.is_empty() {
            return Err(XlsbError::InvalidFormula(
                "formula expression is empty".to_string(),
            ));
        }
        let mut compiler = Self { input, offset: 0 };
        let expression = compiler.parse_comparison()?;
        compiler.skip_spaces();
        if compiler.offset != compiler.input.len() {
            return Err(compiler.error("unexpected trailing input"));
        }

        let mut rgce = Vec::new();
        Self::emit(&expression, &mut rgce, encoding)?;
        if rgce.len() > MAX_CELL_FORMULA_BYTES {
            return Err(XlsbError::InvalidFormula(format!(
                "compiled formula is {} bytes; maximum is {MAX_CELL_FORMULA_BYTES}",
                rgce.len()
            )));
        }
        Ok(CellParsedFormula {
            rgce,
            rgcb: Vec::new(),
        })
    }

    fn parse_comparison(&mut self) -> XlsbResult<CompileExpr> {
        let mut expression = self.parse_concat()?;
        loop {
            let operator = if self.consume("<>") {
                Some(BinaryOperator::NotEqual)
            } else if self.consume("<=") {
                Some(BinaryOperator::LessEqual)
            } else if self.consume(">=") {
                Some(BinaryOperator::GreaterEqual)
            } else if self.consume("=") {
                Some(BinaryOperator::Equal)
            } else if self.consume("<") {
                Some(BinaryOperator::LessThan)
            } else if self.consume(">") {
                Some(BinaryOperator::GreaterThan)
            } else {
                None
            };
            let Some(operator) = operator else { break };
            let right = self.parse_concat()?;
            expression = CompileExpr::Binary(operator, Box::new(expression), Box::new(right));
        }
        Ok(expression)
    }

    fn parse_concat(&mut self) -> XlsbResult<CompileExpr> {
        let mut expression = self.parse_additive()?;
        while self.consume("&") {
            let right = self.parse_additive()?;
            expression = CompileExpr::Binary(
                BinaryOperator::Concat,
                Box::new(expression),
                Box::new(right),
            );
        }
        Ok(expression)
    }

    fn parse_additive(&mut self) -> XlsbResult<CompileExpr> {
        let mut expression = self.parse_multiplicative()?;
        loop {
            let operator = if self.consume("+") {
                Some(BinaryOperator::Add)
            } else if self.consume("-") {
                Some(BinaryOperator::Subtract)
            } else {
                None
            };
            let Some(operator) = operator else { break };
            let right = self.parse_multiplicative()?;
            expression = CompileExpr::Binary(operator, Box::new(expression), Box::new(right));
        }
        Ok(expression)
    }

    fn parse_multiplicative(&mut self) -> XlsbResult<CompileExpr> {
        let mut expression = self.parse_power()?;
        loop {
            let operator = if self.consume("*") {
                Some(BinaryOperator::Multiply)
            } else if self.consume("/") {
                Some(BinaryOperator::Divide)
            } else {
                None
            };
            let Some(operator) = operator else { break };
            let right = self.parse_power()?;
            expression = CompileExpr::Binary(operator, Box::new(expression), Box::new(right));
        }
        Ok(expression)
    }

    fn parse_power(&mut self) -> XlsbResult<CompileExpr> {
        let left = self.parse_unary()?;
        if self.consume("^") {
            let right = self.parse_power()?;
            Ok(CompileExpr::Binary(
                BinaryOperator::Power,
                Box::new(left),
                Box::new(right),
            ))
        } else {
            Ok(left)
        }
    }

    fn parse_unary(&mut self) -> XlsbResult<CompileExpr> {
        if self.consume("+") {
            return Ok(CompileExpr::Unary(
                UnaryOperator::Plus,
                Box::new(self.parse_unary()?),
            ));
        }
        if self.consume("-") {
            return Ok(CompileExpr::Unary(
                UnaryOperator::Minus,
                Box::new(self.parse_unary()?),
            ));
        }
        let mut expression = self.parse_primary()?;
        while self.consume("%") {
            expression = CompileExpr::Unary(UnaryOperator::Percent, Box::new(expression));
        }
        Ok(expression)
    }

    fn parse_primary(&mut self) -> XlsbResult<CompileExpr> {
        self.skip_spaces();
        if self.consume("(") {
            let expression = self.parse_comparison()?;
            if !self.consume(")") {
                return Err(self.error("expected ')'"));
            }
            return Ok(CompileExpr::Parenthesized(Box::new(expression)));
        }
        if self.peek_char() == Some('"') {
            return self.parse_string().map(CompileExpr::String);
        }
        if self
            .peek_char()
            .is_some_and(|ch| ch.is_ascii_digit() || ch == '.')
        {
            return self.parse_number().map(CompileExpr::Number);
        }

        let identifier = self.parse_identifier()?;
        if self.consume("(") {
            let function = builtin_function_by_name(&identifier).ok_or_else(|| {
                XlsbError::UnsupportedFeature(format!(
                    "XLSB formula function {identifier} is not in the supported Ftab set"
                ))
            })?;
            let mut arguments = Vec::new();
            if !self.consume(")") {
                loop {
                    if self.consume(")") {
                        arguments.push(CompileExpr::MissingArg);
                        break;
                    }
                    if self.consume(",") {
                        arguments.push(CompileExpr::MissingArg);
                        continue;
                    }
                    arguments.push(self.parse_comparison()?);
                    if self.consume(")") {
                        break;
                    }
                    if !self.consume(",") {
                        return Err(self.error("expected ',' or ')' in function call"));
                    }
                }
            }
            if arguments.len() < usize::from(function.min_args)
                || arguments.len() > usize::from(function.max_args)
            {
                return Err(XlsbError::InvalidFormula(format!(
                    "{} expects {}..={} arguments, found {}",
                    function.name,
                    function.min_args,
                    function.max_args,
                    arguments.len()
                )));
            }
            return Ok(CompileExpr::Function(function, arguments));
        }
        if identifier.eq_ignore_ascii_case("TRUE") {
            return Ok(CompileExpr::Bool(true));
        }
        if identifier.eq_ignore_ascii_case("FALSE") {
            return Ok(CompileExpr::Bool(false));
        }

        let first = parse_a1_reference(&identifier).ok_or_else(|| {
            self.error("defined names and sheet-qualified references are not yet supported")
        })?;
        if self.consume(":") {
            let second_text = self.parse_identifier()?;
            let second = parse_a1_reference(&second_text)
                .ok_or_else(|| self.error("invalid range end reference"))?;
            Ok(CompileExpr::Area(first, second))
        } else {
            Ok(CompileExpr::Ref(first))
        }
    }

    fn parse_string(&mut self) -> XlsbResult<String> {
        debug_assert_eq!(self.peek_char(), Some('"'));
        self.offset += 1;
        let mut value = String::new();
        loop {
            let Some(ch) = self.peek_char() else {
                return Err(self.error("unterminated string literal"));
            };
            self.offset += ch.len_utf8();
            if ch == '"' {
                if self.peek_char() == Some('"') {
                    self.offset += 1;
                    value.push('"');
                } else {
                    break;
                }
            } else {
                value.push(ch);
            }
        }
        if value.encode_utf16().count() > 255 {
            return Err(XlsbError::InvalidFormula(
                "formula string literal exceeds 255 UTF-16 code units".to_string(),
            ));
        }
        Ok(value)
    }

    fn parse_number(&mut self) -> XlsbResult<f64> {
        self.skip_spaces();
        let start = self.offset;
        let mut seen_exponent = false;
        while let Some(ch) = self.peek_char() {
            if ch.is_ascii_digit() || ch == '.' {
                self.offset += 1;
            } else if matches!(ch, 'e' | 'E') && !seen_exponent {
                seen_exponent = true;
                self.offset += 1;
                if matches!(self.peek_char(), Some('+' | '-')) {
                    self.offset += 1;
                }
            } else {
                break;
            }
        }
        self.input[start..self.offset]
            .parse::<f64>()
            .map_err(|_| self.error("invalid numeric literal"))
    }

    fn parse_identifier(&mut self) -> XlsbResult<String> {
        self.skip_spaces();
        let start = self.offset;
        while let Some(ch) = self.peek_char() {
            if ch.is_ascii_alphanumeric() || matches!(ch, '_' | '.' | '$') {
                self.offset += ch.len_utf8();
            } else {
                break;
            }
        }
        if self.offset == start {
            Err(self.error("expected literal, reference, or function"))
        } else {
            Ok(self.input[start..self.offset].to_string())
        }
    }

    fn consume(&mut self, text: &str) -> bool {
        self.skip_spaces();
        if self.input[self.offset..].starts_with(text) {
            self.offset += text.len();
            true
        } else {
            false
        }
    }

    fn skip_spaces(&mut self) {
        while self.peek_char().is_some_and(char::is_whitespace) {
            self.offset += self.peek_char().expect("checked").len_utf8();
        }
    }

    fn peek_char(&self) -> Option<char> {
        self.input[self.offset..].chars().next()
    }

    fn error(&self, message: &str) -> XlsbError {
        XlsbError::InvalidFormula(format!("{message} at byte {}", self.offset))
    }

    fn emit(
        expression: &CompileExpr,
        output: &mut Vec<u8>,
        encoding: FormulaEncoding,
    ) -> XlsbResult<()> {
        match expression {
            CompileExpr::Number(value) => {
                if value.fract() == 0.0 && *value >= 0.0 && *value <= f64::from(u16::MAX) {
                    output.push(ptg_types::PTG_INT);
                    output.extend_from_slice(&(*value as u16).to_le_bytes());
                } else {
                    output.push(ptg_types::PTG_NUM);
                    output.extend_from_slice(&value.to_le_bytes());
                }
            },
            CompileExpr::String(value) => {
                let utf16: Vec<u16> = value.encode_utf16().collect();
                output.push(ptg_types::PTG_STR);
                output.extend_from_slice(&(utf16.len() as u16).to_le_bytes());
                for unit in utf16 {
                    output.extend_from_slice(&unit.to_le_bytes());
                }
            },
            CompileExpr::Bool(value) => {
                output.push(ptg_types::PTG_BOOL);
                output.push(u8::from(*value));
            },
            CompileExpr::MissingArg => output.push(ptg_types::PTG_MISSING_ARG),
            CompileExpr::Parenthesized(expression) => {
                Self::emit(expression, output, encoding)?;
                output.push(ptg_types::PTG_PAREN);
            },
            CompileExpr::Ref(reference) => match encoding {
                FormulaEncoding::Cell => emit_reference(output, 0x44, *reference),
                FormulaEncoding::Shared { base_row, base_col } => {
                    emit_shared_reference(output, 0x4C, *reference, base_row, base_col)?
                },
            },
            CompileExpr::Area(first, last) => {
                match encoding {
                    FormulaEncoding::Cell => {
                        output.push(0x25); // PtgArea, REFERENCE class
                        output.extend_from_slice(&first.row.to_le_bytes());
                        output.extend_from_slice(&last.row.to_le_bytes());
                        output.extend_from_slice(&reference_column_bits(*first).to_le_bytes());
                        output.extend_from_slice(&reference_column_bits(*last).to_le_bytes());
                    },
                    FormulaEncoding::Shared { base_row, base_col } => {
                        output.push(0x2D); // PtgAreaN, REFERENCE class
                        let (first_row, first_col) =
                            encode_shared_reference(*first, base_row, base_col)?;
                        let (last_row, last_col) =
                            encode_shared_reference(*last, base_row, base_col)?;
                        output.extend_from_slice(&first_row.to_le_bytes());
                        output.extend_from_slice(&last_row.to_le_bytes());
                        output.extend_from_slice(&first_col.to_le_bytes());
                        output.extend_from_slice(&last_col.to_le_bytes());
                    },
                }
            },
            CompileExpr::Unary(operator, operand) => {
                Self::emit(operand, output, encoding)?;
                output.push(match operator {
                    UnaryOperator::Plus => ptg_types::PTG_UPLUS,
                    UnaryOperator::Minus => ptg_types::PTG_UMINUS,
                    UnaryOperator::Percent => ptg_types::PTG_PERCENT,
                });
            },
            CompileExpr::Binary(operator, left, right) => {
                Self::emit(left, output, encoding)?;
                Self::emit(right, output, encoding)?;
                output.push(match operator {
                    BinaryOperator::Add => ptg_types::PTG_ADD,
                    BinaryOperator::Subtract => ptg_types::PTG_SUB,
                    BinaryOperator::Multiply => ptg_types::PTG_MUL,
                    BinaryOperator::Divide => ptg_types::PTG_DIV,
                    BinaryOperator::Power => ptg_types::PTG_POWER,
                    BinaryOperator::Concat => ptg_types::PTG_CONCAT,
                    BinaryOperator::LessThan => ptg_types::PTG_LT,
                    BinaryOperator::LessEqual => ptg_types::PTG_LE,
                    BinaryOperator::Equal => ptg_types::PTG_EQ,
                    BinaryOperator::GreaterEqual => ptg_types::PTG_GE,
                    BinaryOperator::GreaterThan => ptg_types::PTG_GT,
                    BinaryOperator::NotEqual => ptg_types::PTG_NE,
                    BinaryOperator::Intersection => ptg_types::PTG_ISECT,
                    BinaryOperator::Union => ptg_types::PTG_UNION,
                    BinaryOperator::Range => ptg_types::PTG_RANGE,
                });
            },
            CompileExpr::Function(function, arguments) => {
                for argument in arguments {
                    Self::emit(argument, output, encoding)?;
                }
                if function.min_args == function.max_args {
                    output.push(0x41); // PtgFunc, VALUE class
                    output.extend_from_slice(&function.index.to_le_bytes());
                } else {
                    output.push(0x42); // PtgFuncVar, VALUE class
                    output.push(arguments.len() as u8);
                    output.extend_from_slice(&function.index.to_le_bytes());
                }
            },
        }
        Ok(())
    }
}

fn parse_a1_reference(value: &str) -> Option<A1Reference> {
    let bytes = value.as_bytes();
    let mut offset = 0;
    let col_relative = bytes.get(offset) != Some(&b'$');
    if !col_relative {
        offset += 1;
    }
    let col_start = offset;
    while bytes.get(offset).is_some_and(u8::is_ascii_alphabetic) {
        offset += 1;
    }
    if offset == col_start {
        return None;
    }
    let mut col = 0u32;
    for byte in bytes[col_start..offset].iter().map(u8::to_ascii_uppercase) {
        col = col
            .checked_mul(26)?
            .checked_add(u32::from(byte - b'A' + 1))?;
    }
    if col == 0 || col > 16_384 {
        return None;
    }

    let row_relative = bytes.get(offset) != Some(&b'$');
    if !row_relative {
        offset += 1;
    }
    let row_start = offset;
    while bytes.get(offset).is_some_and(u8::is_ascii_digit) {
        offset += 1;
    }
    if offset == row_start || offset != bytes.len() {
        return None;
    }
    let row = value[row_start..offset].parse::<u32>().ok()?;
    if row == 0 || row > 1_048_576 {
        return None;
    }
    Some(A1Reference {
        row: row - 1,
        col: col - 1,
        row_relative,
        col_relative,
    })
}

fn reference_column_bits(reference: A1Reference) -> u16 {
    let mut bits = reference.col as u16;
    if reference.col_relative {
        bits |= 0x4000;
    }
    if reference.row_relative {
        bits |= 0x8000;
    }
    bits
}

fn emit_reference(output: &mut Vec<u8>, token: u8, reference: A1Reference) {
    output.push(token);
    output.extend_from_slice(&reference.row.to_le_bytes());
    output.extend_from_slice(&reference_column_bits(reference).to_le_bytes());
}

fn emit_shared_reference(
    output: &mut Vec<u8>,
    token: u8,
    reference: A1Reference,
    base_row: u32,
    base_col: u32,
) -> XlsbResult<()> {
    let (row, col) = encode_shared_reference(reference, base_row, base_col)?;
    output.push(token);
    output.extend_from_slice(&row.to_le_bytes());
    output.extend_from_slice(&col.to_le_bytes());
    Ok(())
}

fn encode_shared_reference(
    reference: A1Reference,
    base_row: u32,
    base_col: u32,
) -> XlsbResult<(u32, u16)> {
    let row = if reference.row_relative {
        let offset = i64::from(reference.row) - i64::from(base_row);
        i32::try_from(offset)
            .map_err(|_| XlsbError::InvalidFormula("shared row offset overflow".to_string()))?
            as u32
    } else {
        reference.row
    };
    let col_value = if reference.col_relative {
        let offset = i64::from(reference.col) - i64::from(base_col);
        if !(-16_383..=16_383).contains(&offset) {
            return Err(XlsbError::InvalidFormula(format!(
                "shared column offset {offset} is outside the XLSB range"
            )));
        }
        (offset as i32 as u16) & 0x3FFF
    } else {
        reference.col as u16
    };
    let mut col = col_value;
    if reference.col_relative {
        col |= 0x4000;
    }
    if reference.row_relative {
        col |= 0x8000;
    }
    Ok((row, col))
}

fn add_wrapped_offset(base: u32, offset: i32, modulus: u32) -> u32 {
    (i64::from(base) + i64::from(offset)).rem_euclid(i64::from(modulus)) as u32
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_operators() {
        let data = vec![0x03]; // PTG_ADD
        let mut parser = FormulaParser::new(&data);
        let tokens = parser.parse().unwrap();
        assert_eq!(tokens.len(), 1);
        match &tokens[0] {
            FormulaToken::BinaryOp(BinaryOperator::Add) => {},
            _ => panic!("Expected Add operator"),
        }
    }

    #[test]
    fn test_parse_number() {
        let mut data = vec![0x1F]; // PTG_NUM
        data.extend_from_slice(&42.5f64.to_le_bytes());
        let mut parser = FormulaParser::new(&data);
        let tokens = parser.parse().unwrap();
        assert_eq!(tokens.len(), 1);
        match &tokens[0] {
            FormulaToken::Number(n) if (*n - 42.5).abs() < 0.001 => {},
            _ => panic!("Expected number 42.5"),
        }
    }

    #[test]
    fn test_formula_converter() {
        let tokens = vec![
            FormulaToken::Number(1.0),
            FormulaToken::Number(2.0),
            FormulaToken::BinaryOp(BinaryOperator::Add),
        ];
        let formula = FormulaConverter::tokens_to_string(&tokens);
        assert_eq!(formula, "(1+2)");
    }

    #[test]
    fn parses_ms_xlsb_brt_fmla_num_example_formula() {
        // [MS-XLSB] 3.7.37: PtgRef(C13), PtgInt(2), PtgMul.
        let rgce = vec![
            0x44, 0x0C, 0x00, 0x00, 0x00, 0x02, 0xC0, 0x1E, 0x02, 0x00, 0x05,
        ];
        let parsed = CellParsedFormula {
            rgce: rgce.clone(),
            rgcb: Vec::new(),
        };
        let bytes = parsed.to_bytes().unwrap();
        let (roundtrip, consumed) = CellParsedFormula::parse(&bytes).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(roundtrip, parsed);

        let tokens = FormulaParser::new(&rgce).parse().unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
            "(C13*2)"
        );
    }

    #[test]
    fn compiler_matches_ms_xlsb_reference_and_multiply_tokens() {
        let formula = FormulaCompiler::compile("=C13*2").unwrap();
        assert_eq!(
            formula.rgce,
            vec![
                0x44, 0x0C, 0x00, 0x00, 0x00, 0x02, 0xC0, 0x1E, 0x02, 0x00, 0x05,
            ]
        );
    }

    #[test]
    fn compiler_supports_ranges_functions_unicode_and_absolute_refs() {
        let formula = FormulaCompiler::compile("SUM($A$1:B3)+\"荔枝\"").unwrap();
        let tokens = FormulaParser::new(&formula.rgce).parse().unwrap();
        let text = FormulaConverter::try_tokens_to_string(&tokens).unwrap();
        assert_eq!(text, "(SUM($A$1:B3)+\"荔枝\")");
    }

    #[test]
    fn compiler_and_converter_preserve_missing_arguments_and_parentheses() {
        let missing = FormulaCompiler::compile("IF(TRUE,,0)").unwrap();
        assert!(missing.rgce.contains(&ptg_types::PTG_MISSING_ARG));
        let tokens = FormulaParser::new(&missing.rgce).parse().unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
            "IF(TRUE,,0)"
        );

        let parenthesized = FormulaCompiler::compile("(1+2)*3").unwrap();
        assert!(parenthesized.rgce.contains(&ptg_types::PTG_PAREN));
        let tokens = FormulaParser::new(&parenthesized.rgce).parse().unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
            "(((1+2))*3)"
        );
    }

    #[test]
    fn parser_converts_binary_reference_operators() {
        let mut rgce = FormulaCompiler::compile("A1").unwrap().rgce;
        rgce.extend_from_slice(&FormulaCompiler::compile("B2").unwrap().rgce);
        rgce.push(ptg_types::PTG_UNION);
        let tokens = FormulaParser::new(&rgce).parse().unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
            "(A1,B2)"
        );
    }

    #[test]
    fn shared_formula_uses_relative_tokens_and_expands_per_target_cell() {
        // Real shared-formula pattern from POI bug66682.xlsb: the C3:C10
        // formula group references the cell one column earlier.
        let formula = FormulaCompiler::compile_shared("B3", 2, 2).unwrap();
        assert_eq!(formula.rgce, vec![0x4C, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF]);

        let anchor_tokens = FormulaParser::with_base_cell(&formula.rgce, 2, 2)
            .parse()
            .unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&anchor_tokens).unwrap(),
            "B3"
        );
        let follower_tokens = FormulaParser::with_base_cell(&formula.rgce, 3, 2)
            .parse()
            .unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&follower_tokens).unwrap(),
            "B4"
        );
    }

    #[test]
    fn parses_real_poi_shared_formula_definition_losslessly() {
        // BrtShrFmla from POI bug66682.xlsb: C3:C10 refers one column left.
        let bytes = [
            0x02, 0x00, 0x00, 0x00, 0x09, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x02, 0x00,
            0x00, 0x00, 0x07, 0x00, 0x00, 0x00, 0x4C, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0x00,
            0x00, 0x00, 0x00,
        ];
        let group = FormulaGroup::parse_shared(&bytes).unwrap();
        assert_eq!(group.kind, FormulaGroupKind::Shared);
        assert_eq!(group.range.to_a1(), "C3:C10");
        assert_eq!(group.to_record_data().unwrap(), bytes);

        let tokens = FormulaParser::with_base_cell(&group.formula.rgce, 9, 2)
            .parse()
            .unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
            "B10"
        );
    }

    #[test]
    fn parses_real_poi_array_formula_definition_losslessly() {
        // BrtArrFmla from POI bug66682.xlsb. Its PtgName is retained even
        // when the standalone formula converter cannot resolve that name.
        let bytes = [
            0x08, 0x00, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x02, 0x00,
            0x00, 0x00, 0x01, 0x09, 0x00, 0x00, 0x00, 0x23, 0x02, 0x00, 0x00, 0x00, 0x42, 0x01,
            0xFF, 0x00, 0x00, 0x00, 0x00, 0x00,
        ];
        let group = FormulaGroup::parse_array(&bytes).unwrap();
        assert_eq!(group.kind, FormulaGroupKind::Array);
        assert_eq!(group.range.to_a1(), "C9:C9");
        assert!(group.always_calculate);
        assert_eq!(group.to_record_data().unwrap(), bytes);
    }

    #[test]
    fn rejects_malformed_ptg_exp_and_array_flags() {
        let malformed = CellParsedFormula {
            rgce: vec![ptg_types::PTG_EXP, 0, 0],
            rgcb: vec![],
        };
        assert!(matches!(
            malformed.exp_cell(),
            Err(XlsbError::InvalidFormula(_))
        ));

        let mut array = FormulaGroup {
            kind: FormulaGroupKind::Array,
            range: FormulaRange::new(0, 0, 0, 0).unwrap(),
            formula: FormulaCompiler::compile("1+1").unwrap(),
            always_calculate: false,
        }
        .to_record_data()
        .unwrap();
        array[16] = 0x80;
        assert!(matches!(
            FormulaGroup::parse_array(&array),
            Err(XlsbError::InvalidFormula(_))
        ));
    }

    #[test]
    fn shared_formula_preserves_mixed_absolute_references() {
        let formula = FormulaCompiler::compile_shared("$A1+B$2", 4, 3).unwrap();
        let tokens = FormulaParser::with_base_cell(&formula.rgce, 7, 5)
            .parse()
            .unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
            "($A4+D$2)"
        );
    }

    #[test]
    fn cell_parsed_formula_rejects_zero_and_oversized_token_streams() {
        let zero = [0_u8; 8];
        assert!(matches!(
            CellParsedFormula::parse(&zero),
            Err(XlsbError::InvalidFormula(_))
        ));

        let mut oversized = Vec::new();
        oversized.extend_from_slice(&((MAX_CELL_FORMULA_BYTES as u32) + 1).to_le_bytes());
        oversized.extend_from_slice(&[0; 4]);
        assert!(matches!(
            CellParsedFormula::parse(&oversized),
            Err(XlsbError::InvalidFormula(_))
        ));
    }

    #[test]
    fn truncated_token_is_an_error_instead_of_becoming_unknown_bytes() {
        let error = FormulaParser::new(&[0x44, 0x01]).parse().unwrap_err();
        assert!(matches!(error, XlsbError::InvalidFormula(_)));
    }
}
