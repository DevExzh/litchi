//! Semantic XLSB formula and Ptg models.
//!
//! The names in this module are contextual to `formula`.

use super::{Error, Result};

pub(super) fn read_u32_le_at(data: &[u8], offset: usize) -> Result<u32> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| Error::InvalidFormula("formula binary offset overflow".to_string()))?;
    let bytes = data.get(offset..end).ok_or(Error::InvalidLength {
        expected: end,
        found: data.len(),
    })?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

/// Inclusive worksheet range used by grouped and reference formulas.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Range {
    pub row_first: u32,
    pub row_last: u32,
    pub col_first: u32,
    pub col_last: u32,
}

impl Range {
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

    pub(crate) fn validate(self) -> Result<()> {
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

pub(super) fn column_index_to_name(mut column: u32) -> String {
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

/// Binary representation of a cell formula.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParsedFormula {
    pub rgce: Vec<u8>,
    pub rgcb: Vec<u8>,
}

/// Scalar value stored in an XLSB `PtgExtraArray`.
#[derive(Debug, Clone, PartialEq)]
pub enum ArrayValue {
    Number(f64),
    String(String),
    Bool(bool),
    Error(u8),
}

/// Kind of non-evaluating memory marker in a Ptg stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MemoryKind {
    Area,
    Error(u8),
    Function,
    NoMemory,
}

/// Row subset selected by a structured table reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableRowType {
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
pub enum TableColumns {
    All,
    One(u16),
    Range { first: u16, last: u16 },
}

/// Operand class carried by `PtgList`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableDataType {
    Reference,
    Value,
    Array,
}

/// Named column subset stored in a nonresident `PtgExtraList`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TableNamedColumns {
    All,
    One(String),
    Range { first: String, last: String },
}

/// Ancillary table and column names for an external structured reference.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalTableReference {
    pub table: String,
    pub row_type: TableRowType,
    pub columns: TableNamedColumns,
}

/// Typed XLSB `PtgList` structured table reference.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableReference {
    pub sheet_index: u16,
    pub row_type: Option<TableRowType>,
    pub columns: Option<TableColumns>,
    pub square_bracket_space: bool,
    pub comma_space: bool,
    pub data_type: TableDataType,
    pub invalid: bool,
    pub list_index: Option<u32>,
    pub external: Option<ExternalTableReference>,
}

/// Kind of formula definition following a `PtgExp` cell record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum GroupKind {
    Array,
    Shared,
}

/// Parsed `BrtArrFmla` or `BrtShrFmla` definition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Group {
    pub kind: GroupKind,
    pub range: Range,
    pub formula: ParsedFormula,
    pub always_calculate: bool,
}

/// Parse Tree Generator (Ptg) token representation.
#[derive(Debug, Clone, PartialEq)]
pub enum Token {
    Number(f64),
    String(String),
    Bool(bool),
    Error(u8),
    Int(u16),
    MissingArg,
    Parenthesis,
    Attribute(u8),
    Array {
        rows: u32,
        cols: u32,
        values: Vec<ArrayValue>,
    },
    Memory {
        kind: MemoryKind,
        expression_bytes: u16,
        cached_ranges: Vec<[u32; 4]>,
    },
    CellRef {
        row: u32,
        col: u32,
        row_relative: bool,
        col_relative: bool,
    },
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
    CellRef3d {
        sheet_index: u16,
        row: u32,
        col: u32,
        row_relative: bool,
        col_relative: bool,
    },
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
    ReferenceError {
        is_area: bool,
        sheet_index: Option<u16>,
    },
    BinaryOp(BinaryOperator),
    UnaryOp(UnaryOperator),
    Function {
        index: u16,
        arg_count: u8,
        is_command: bool,
    },
    Name(u32),
    ExternalName {
        sheet_index: u16,
        name_index: u32,
    },
    TableReference(TableReference),
    PivotName(u32),
    Unknown(u8),
}

/// Binary operators in a Ptg stream.
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

/// Unary operators in a Ptg stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnaryOperator {
    Plus,
    Minus,
    Percent,
}

/// Parse Tree Generator token constants.
#[allow(dead_code)]
pub mod ptg_types {
    pub const PTG_EXP: u8 = 0x01;
    pub const PTG_TBL: u8 = 0x02;
    pub const PTG_ADD: u8 = 0x03;
    pub const PTG_SUB: u8 = 0x04;
    pub const PTG_MUL: u8 = 0x05;
    pub const PTG_DIV: u8 = 0x06;
    pub const PTG_POWER: u8 = 0x07;
    pub const PTG_CONCAT: u8 = 0x08;
    pub const PTG_LT: u8 = 0x09;
    pub const PTG_LE: u8 = 0x0A;
    pub const PTG_EQ: u8 = 0x0B;
    pub const PTG_GE: u8 = 0x0C;
    pub const PTG_GT: u8 = 0x0D;
    pub const PTG_NE: u8 = 0x0E;
    pub const PTG_ISECT: u8 = 0x0F;
    pub const PTG_UNION: u8 = 0x10;
    pub const PTG_RANGE: u8 = 0x11;
    pub const PTG_UPLUS: u8 = 0x12;
    pub const PTG_UMINUS: u8 = 0x13;
    pub const PTG_PERCENT: u8 = 0x14;
    pub const PTG_PAREN: u8 = 0x15;
    pub const PTG_MISSING_ARG: u8 = 0x16;
    pub const PTG_STR: u8 = 0x17;
    pub const PTG_EXTENDED: u8 = 0x18;
    pub const PTG_ATTR: u8 = 0x19;
    pub const PTG_SHEET: u8 = 0x1A;
    pub const PTG_END_SHEET: u8 = 0x1B;
    pub const PTG_ERR: u8 = 0x1C;
    pub const PTG_BOOL: u8 = 0x1D;
    pub const PTG_INT: u8 = 0x1E;
    pub const PTG_NUM: u8 = 0x1F;
    pub const PTG_REF: u8 = 0x24;
    pub const PTG_AREA: u8 = 0x25;
    pub const PTG_MEM_AREA: u8 = 0x26;
    pub const PTG_MEM_ERR: u8 = 0x27;
    pub const PTG_MEM_NO_MEM: u8 = 0x28;
    pub const PTG_MEM_FUNC: u8 = 0x29;
    pub const PTG_REF_ERR: u8 = 0x2A;
    pub const PTG_AREA_ERR: u8 = 0x2B;
    pub const PTG_REF_N: u8 = 0x2C;
    pub const PTG_AREA_N: u8 = 0x2D;
    pub const PTG_NAME_X: u8 = 0x39;
    pub const PTG_REF_3D: u8 = 0x3A;
    pub const PTG_AREA_3D: u8 = 0x3B;
    pub const PTG_REF_ERR_3D: u8 = 0x3C;
    pub const PTG_AREA_ERR_3D: u8 = 0x3D;
    pub const PTG_FUNC: u8 = 0x21;
    pub const PTG_FUNC_VAR: u8 = 0x22;
    pub const PTG_NAME: u8 = 0x23;
    pub const PTG_ARRAY: u8 = 0x20;
    pub const EPTG_LIST: u8 = 0x19;
    pub const EPTG_SX_NAME: u8 = 0x1D;
}
