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
//! - [MS-XLSB] Section 2.5.97 - Formulas
//! - [MS-XLS] Section 2.5.198 - Ptg (for token details, largely compatible)

use crate::common::binary;
use crate::ooxml::xlsb::error::XlsbResult;

/// Parse Tree Generator (Ptg) token types
///
/// These constants define the various formula token types used in XLSB.
/// Reference: [MS-XLSB] Section 2.5.97.23
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
#[derive(Debug, Clone)]
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
    /// Function call (function index, arg count)
    Function { index: u16, arg_count: u8 },
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
}

impl<'a> FormulaParser<'a> {
    /// Create a new formula parser
    pub fn new(data: &'a [u8]) -> Self {
        FormulaParser { data, offset: 0 }
    }

    /// Parse the formula into tokens
    ///
    /// Returns a vector of formula tokens in RPN order.
    pub fn parse(&mut self) -> XlsbResult<Vec<FormulaToken>> {
        let mut tokens = Vec::new();

        while self.offset < self.data.len() {
            if let Some(token) = self.parse_token()? {
                tokens.push(token);
            } else {
                // Skip unknown tokens
                self.offset += 1;
            }
        }

        Ok(tokens)
    }

    /// Parse a single token
    fn parse_token(&mut self) -> XlsbResult<Option<FormulaToken>> {
        if self.offset >= self.data.len() {
            return Ok(None);
        }

        let ptg_type = self.data[self.offset];
        self.offset += 1;

        use ptg_types::*;

        match ptg_type {
            PTG_ADD => Ok(Some(FormulaToken::BinaryOp(BinaryOperator::Add))),
            PTG_SUB => Ok(Some(FormulaToken::BinaryOp(BinaryOperator::Subtract))),
            PTG_MUL => Ok(Some(FormulaToken::BinaryOp(BinaryOperator::Multiply))),
            PTG_DIV => Ok(Some(FormulaToken::BinaryOp(BinaryOperator::Divide))),
            PTG_POWER => Ok(Some(FormulaToken::BinaryOp(BinaryOperator::Power))),
            PTG_CONCAT => Ok(Some(FormulaToken::BinaryOp(BinaryOperator::Concat))),
            PTG_LT => Ok(Some(FormulaToken::BinaryOp(BinaryOperator::LessThan))),
            PTG_LE => Ok(Some(FormulaToken::BinaryOp(BinaryOperator::LessEqual))),
            PTG_EQ => Ok(Some(FormulaToken::BinaryOp(BinaryOperator::Equal))),
            PTG_GE => Ok(Some(FormulaToken::BinaryOp(BinaryOperator::GreaterEqual))),
            PTG_GT => Ok(Some(FormulaToken::BinaryOp(BinaryOperator::GreaterThan))),
            PTG_NE => Ok(Some(FormulaToken::BinaryOp(BinaryOperator::NotEqual))),

            PTG_UPLUS => Ok(Some(FormulaToken::UnaryOp(UnaryOperator::Plus))),
            PTG_UMINUS => Ok(Some(FormulaToken::UnaryOp(UnaryOperator::Minus))),
            PTG_PERCENT => Ok(Some(FormulaToken::UnaryOp(UnaryOperator::Percent))),

            PTG_INT => self.parse_int(),
            PTG_NUM => self.parse_num(),
            PTG_STR => self.parse_str(),
            PTG_BOOL => self.parse_bool(),
            PTG_ERR => self.parse_err(),

            PTG_REF | PTG_REF_N => self.parse_ref(),
            PTG_AREA | PTG_AREA_N => self.parse_area(),

            PTG_FUNC => self.parse_func(),
            PTG_FUNC_VAR => self.parse_func_var(),

            PTG_NAME => self.parse_name(),

            _ => {
                // Unknown token type
                Ok(Some(FormulaToken::Unknown(ptg_type)))
            },
        }
    }

    /// Parse integer constant
    fn parse_int(&mut self) -> XlsbResult<Option<FormulaToken>> {
        if self.offset + 2 > self.data.len() {
            return Ok(None);
        }

        let value = binary::read_u16_le_at(self.data, self.offset)?;
        self.offset += 2;

        Ok(Some(FormulaToken::Int(value)))
    }

    /// Parse floating point constant
    fn parse_num(&mut self) -> XlsbResult<Option<FormulaToken>> {
        if self.offset + 8 > self.data.len() {
            return Ok(None);
        }

        let value = binary::read_f64_le_at(self.data, self.offset)?;
        self.offset += 8;

        Ok(Some(FormulaToken::Number(value)))
    }

    /// Parse string constant
    fn parse_str(&mut self) -> XlsbResult<Option<FormulaToken>> {
        if self.offset + 1 > self.data.len() {
            return Ok(None);
        }

        let len = self.data[self.offset] as usize;
        self.offset += 1;

        if self.offset + len > self.data.len() {
            return Ok(None);
        }

        // String is stored as UTF-8 (not UTF-16LE like in records)
        let string =
            String::from_utf8_lossy(&self.data[self.offset..self.offset + len]).into_owned();
        self.offset += len;

        Ok(Some(FormulaToken::String(string)))
    }

    /// Parse boolean constant
    fn parse_bool(&mut self) -> XlsbResult<Option<FormulaToken>> {
        if self.offset + 1 > self.data.len() {
            return Ok(None);
        }

        let value = self.data[self.offset] != 0;
        self.offset += 1;

        Ok(Some(FormulaToken::Bool(value)))
    }

    /// Parse error constant
    fn parse_err(&mut self) -> XlsbResult<Option<FormulaToken>> {
        if self.offset + 1 > self.data.len() {
            return Ok(None);
        }

        let error_code = self.data[self.offset];
        self.offset += 1;

        Ok(Some(FormulaToken::Error(error_code)))
    }

    /// Parse cell reference
    fn parse_ref(&mut self) -> XlsbResult<Option<FormulaToken>> {
        if self.offset + 4 > self.data.len() {
            return Ok(None);
        }

        let row_data = binary::read_u16_le_at(self.data, self.offset)?;
        let col_data = binary::read_u16_le_at(self.data, self.offset + 2)?;
        self.offset += 4;

        // Extract row and column (with relative flags)
        let row = (row_data & 0x3FFF) as u32;
        let row_relative = (row_data & 0x8000) != 0;
        let col = (col_data & 0x3FFF) as u32;
        let col_relative = (col_data & 0x8000) != 0;

        Ok(Some(FormulaToken::CellRef {
            row,
            col,
            row_relative,
            col_relative,
        }))
    }

    /// Parse area reference
    fn parse_area(&mut self) -> XlsbResult<Option<FormulaToken>> {
        if self.offset + 8 > self.data.len() {
            return Ok(None);
        }

        let row_first_data = binary::read_u16_le_at(self.data, self.offset)?;
        let row_last_data = binary::read_u16_le_at(self.data, self.offset + 2)?;
        let col_first_data = binary::read_u16_le_at(self.data, self.offset + 4)?;
        let col_last_data = binary::read_u16_le_at(self.data, self.offset + 6)?;
        self.offset += 8;

        let row_first = (row_first_data & 0x3FFF) as u32;
        let row_first_relative = (row_first_data & 0x8000) != 0;
        let row_last = (row_last_data & 0x3FFF) as u32;
        let row_last_relative = (row_last_data & 0x8000) != 0;
        let col_first = (col_first_data & 0x3FFF) as u32;
        let col_first_relative = (col_first_data & 0x8000) != 0;
        let col_last = (col_last_data & 0x3FFF) as u32;
        let col_last_relative = (col_last_data & 0x8000) != 0;

        Ok(Some(FormulaToken::AreaRef {
            row_first,
            row_last,
            col_first,
            col_last,
            row_first_relative,
            row_last_relative,
            col_first_relative,
            col_last_relative,
        }))
    }

    /// Parse function with fixed arguments
    fn parse_func(&mut self) -> XlsbResult<Option<FormulaToken>> {
        if self.offset + 2 > self.data.len() {
            return Ok(None);
        }

        let index = binary::read_u16_le_at(self.data, self.offset)?;
        self.offset += 2;

        // Look up argument count from function table (simplified)
        let arg_count = Self::get_function_arg_count(index);

        Ok(Some(FormulaToken::Function { index, arg_count }))
    }

    /// Parse function with variable arguments
    fn parse_func_var(&mut self) -> XlsbResult<Option<FormulaToken>> {
        if self.offset + 3 > self.data.len() {
            return Ok(None);
        }

        let arg_count = self.data[self.offset];
        let index = binary::read_u16_le_at(self.data, self.offset + 1)?;
        self.offset += 3;

        Ok(Some(FormulaToken::Function { index, arg_count }))
    }

    /// Parse defined name reference
    fn parse_name(&mut self) -> XlsbResult<Option<FormulaToken>> {
        if self.offset + 4 > self.data.len() {
            return Ok(None);
        }

        let name_index = binary::read_u32_le_at(self.data, self.offset)?;
        self.offset += 4;

        Ok(Some(FormulaToken::Name(name_index)))
    }

    /// Get function argument count by function index
    ///
    /// This is a simplified lookup. In a complete implementation, this would
    /// use a comprehensive table of all Excel functions.
    fn get_function_arg_count(index: u16) -> u8 {
        match index {
            0 => 1, // COUNT
            4 => 2, // SUM (variable, but typically 1+)
            1 => 2, // IF (typically 3, but can be 2)
            _ => 1, // Default to 1 for unknown
        }
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
        let mut stack: Vec<String> = Vec::new();

        for token in tokens {
            match token {
                FormulaToken::Number(n) => stack.push(format!("{}", n)),
                FormulaToken::Int(i) => stack.push(format!("{}", i)),
                FormulaToken::String(s) => stack.push(format!("\"{}\"", s)),
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
                    let col_str = crate::ooxml::xlsb::utils::column_index_to_name(*col + 1);
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
                    ..
                } => {
                    let first = crate::ooxml::xlsb::utils::cell_reference(*row_first, *col_first);
                    let last = crate::ooxml::xlsb::utils::cell_reference(*row_last, *col_last);
                    stack.push(format!("{}:{}", first, last));
                },
                FormulaToken::BinaryOp(op) => {
                    if stack.len() >= 2 {
                        let right = stack.pop().unwrap();
                        let left = stack.pop().unwrap();
                        let op_str = Self::binary_op_to_string(*op);
                        stack.push(format!("({}{}{})", left, op_str, right));
                    }
                },
                FormulaToken::UnaryOp(op) => {
                    if !stack.is_empty() {
                        let operand = stack.pop().unwrap();
                        match op {
                            UnaryOperator::Plus => stack.push(format!("+({})", operand)),
                            UnaryOperator::Minus => stack.push(format!("-({})", operand)),
                            UnaryOperator::Percent => stack.push(format!("({}%)", operand)),
                        }
                    }
                },
                FormulaToken::Function { index, arg_count } => {
                    let func_name = Self::function_name(*index);
                    let mut args = Vec::new();
                    for _ in 0..*arg_count {
                        if let Some(arg) = stack.pop() {
                            args.insert(0, arg);
                        }
                    }
                    stack.push(format!("{}({})", func_name, args.join(",")));
                },
                FormulaToken::Name(idx) => stack.push(format!("Name{}", idx)),
                FormulaToken::Unknown(t) => stack.push(format!("?Ptg{:02X}?", t)),
            }
        }

        stack.pop().unwrap_or_default()
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

    /// Get function name by index (simplified)
    fn function_name(index: u16) -> String {
        match index {
            0 => "COUNT".to_string(),
            1 => "IF".to_string(),
            4 => "SUM".to_string(),
            5 => "AVERAGE".to_string(),
            6 => "MIN".to_string(),
            7 => "MAX".to_string(),
            _ => format!("FUNC{}", index),
        }
    }
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
}
