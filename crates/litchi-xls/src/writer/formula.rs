//! XLS formula tokenization (RPN parsing)
//!
//! This module implements Excel's formula tokenization system, converting
//! infix formula notation (e.g., "=A1+B1") to Reverse Polish Notation (RPN)
//! tokens that Excel understands (Ptg - Parse Things).
//!
//! Based on Microsoft's "[MS-XLS]" specification and Apache POI's FormulaParser.
//!
//! # Formula Structure
//!
//! Excel formulas are stored as a sequence of Ptg (Parse Thing) tokens:
//! - **Operand tokens**: References (A1, $B$2), constants (42, "text")
//! - **Operator tokens**: +, -, *, /, ^, &, =, <>, etc.
//! - **Function tokens**: SUM, IF, VLOOKUP, etc.
//!
//! # Example
//!
//! ```text
//! Formula: =A1+B1*2
//! Tokens: [Ref(A1), Ref(B1), Int(2), Mul, Add]
//! ```

use super::super::XlsError;
use std::collections::HashMap;

const MAX_BIFF8_COLUMN: u16 = 255;

/// A checked BIFF8 cell reference.
///
/// Rows and columns are zero-based. The private `u8` column makes references
/// beyond Excel 97-2003 column IV unrepresentable after construction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Ref {
    row: u16,
    col: u8,
    row_rel: bool,
    col_rel: bool,
}

impl Ref {
    /// Construct an absolute reference from zero-based BIFF8 coordinates.
    pub const fn new(row: u16, col: u8) -> Self {
        Self {
            row,
            col,
            row_rel: false,
            col_rel: false,
        }
    }

    /// Convert a wider zero-based column after checking the BIFF8 limit.
    pub fn checked(row: u16, col: u16) -> Result<Self, XlsError> {
        let col = u8::try_from(col).map_err(|_| {
            XlsError::InvalidCellReference(format!(
                "zero-based column {col} exceeds IV ({MAX_BIFF8_COLUMN})"
            ))
        })?;
        Ok(Self::new(row, col))
    }

    /// Parse an A1-style BIFF8 reference.
    pub fn parse(value: &str) -> Result<Self, XlsError> {
        parse_ref(value)
    }

    /// Set whether the row is relative.
    pub const fn rel_row(mut self, relative: bool) -> Self {
        self.row_rel = relative;
        self
    }

    /// Set whether the column is relative.
    pub const fn rel_col(mut self, relative: bool) -> Self {
        self.col_rel = relative;
        self
    }

    /// Return an absolute form of this reference.
    pub const fn abs(mut self) -> Self {
        self.row_rel = false;
        self.col_rel = false;
        self
    }

    /// Zero-based row.
    pub const fn row(self) -> u16 {
        self.row
    }

    /// Zero-based column, bounded to 0..=255.
    pub const fn col(self) -> u8 {
        self.col
    }

    /// Whether the row is relative to the formula cell.
    pub const fn is_row_rel(self) -> bool {
        self.row_rel
    }

    /// Whether the column is relative to the formula cell.
    pub const fn is_col_rel(self) -> bool {
        self.col_rel
    }

    const fn col_flags(self) -> u16 {
        self.col as u16
            | if self.col_rel { 0x4000 } else { 0 }
            | if self.row_rel { 0x8000 } else { 0 }
    }
}

impl std::str::FromStr for Ref {
    type Err = XlsError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        Self::parse(value)
    }
}

/// A checked, ordered BIFF8 cell area.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Area {
    first: Ref,
    last: Ref,
}

impl Area {
    /// Construct an ordered area from its upper-left and lower-right cells.
    pub fn new(first: Ref, last: Ref) -> Result<Self, XlsError> {
        if first.row > last.row || first.col > last.col {
            return Err(XlsError::InvalidCellReference(
                "BIFF8 area endpoints are reversed".to_string(),
            ));
        }
        Ok(Self { first, last })
    }

    /// Upper-left cell.
    pub const fn first(self) -> Ref {
        self.first
    }

    /// Lower-right cell.
    pub const fn last(self) -> Ref {
        self.last
    }
}

/// Ptg (Parse Thing) token types
#[derive(Debug, Clone, PartialEq)]
pub enum Ptg {
    /// Integer constant
    Int(u16),
    /// Number constant
    Num(f64),
    /// String constant
    Str(String),
    /// Boolean constant
    Bool(bool),
    /// Checked cell reference.
    Ref(Ref),
    /// Checked area reference.
    Area(Area),
    /// 3D area reference (external-sheet index and checked area).
    ///
    /// Used by defined names and other structures that require
    /// NameParsedFormula, which MUST use 3D references instead of
    /// plain 2D PtgArea in BIFF8.
    Area3d(u16, Area),
    /// Addition operator
    Add,
    /// Subtraction operator
    Sub,
    /// Multiplication operator
    Mul,
    /// Division operator
    Div,
    /// Power operator
    Power,
    /// Concatenation operator
    Concat,
    /// Less than
    Lt,
    /// Less than or equal
    Le,
    /// Equal
    Eq,
    /// Greater than or equal
    Ge,
    /// Greater than
    Gt,
    /// Not equal
    Ne,
    /// Range operator
    Range,
    /// Unary plus operator
    UnaryPlus,
    /// Unary minus operator
    UnaryMinus,
    /// Percent postfix operator
    Percent,
    /// Function call (function index, arg count)
    Func(u16, u8),
    /// Parentheses
    Paren,
    /// Missing argument
    MissArg,
}

/// Operator precedence
fn get_precedence(op: &str) -> u8 {
    match op {
        ":" => 6,
        "u+" | "u-" => 5,
        "^" => 4,
        "*" | "/" => 3,
        "+" | "-" => 2,
        "&" => 2,
        "=" | "<>" | "<" | "<=" | ">" | ">=" => 1,
        _ => 0,
    }
}

/// Parse a cell reference like "A1" or "$B$2" into a [`Ptg::Ref`] token.
///
/// This is exposed as `pub(crate)` so that other writer components
/// (for example, named range handling) can reuse the same parsing
/// logic and stay consistent with formula tokenization.
pub(crate) fn parse_cell_ref(s: &str) -> Result<Ptg, XlsError> {
    Ref::parse(s).map(Ptg::Ref)
}

fn parse_ref(value: &str) -> Result<Ref, XlsError> {
    let value = value.trim();
    let bytes = value.as_bytes();
    let invalid = || XlsError::InvalidCellReference(value.to_string());
    let mut position = 0;

    let column_relative = if bytes.get(position) == Some(&b'$') {
        position += 1;
        false
    } else {
        true
    };

    let column_start = position;
    let mut column = 0u16;
    while let Some(byte) = bytes.get(position).copied() {
        if !byte.is_ascii_alphabetic() {
            break;
        }
        let digit = u16::from(byte.to_ascii_uppercase() - b'A' + 1);
        column = column
            .checked_mul(26)
            .and_then(|current| current.checked_add(digit))
            .ok_or_else(invalid)?;
        if column > MAX_BIFF8_COLUMN + 1 {
            return Err(invalid());
        }
        position += 1;
    }
    if position == column_start {
        return Err(invalid());
    }

    let row_relative = if bytes.get(position) == Some(&b'$') {
        position += 1;
        false
    } else {
        true
    };

    let row_start = position;
    let mut row = 0u32;
    while let Some(byte) = bytes.get(position).copied() {
        if !byte.is_ascii_digit() {
            return Err(invalid());
        }
        row = row
            .checked_mul(10)
            .and_then(|current| current.checked_add(u32::from(byte - b'0')))
            .ok_or_else(invalid)?;
        if row > 65_536 {
            return Err(invalid());
        }
        position += 1;
    }
    if position == row_start || row == 0 {
        return Err(invalid());
    }

    let column = u8::try_from(column - 1).map_err(|_| invalid())?;
    Ok(Ref::new((row - 1) as u16, column)
        .rel_row(row_relative)
        .rel_col(column_relative))
}

/// Formula tokenizer - converts infix formula to RPN tokens
pub struct FormulaTokenizer {
    /// Built-in function names to indices
    functions: HashMap<String, u16>,
}

impl FormulaTokenizer {
    /// Create a new formula tokenizer
    pub fn new() -> Self {
        let mut functions = HashMap::new();

        // Common Excel functions (index from Excel function table)
        functions.insert("SUM".to_string(), 4);
        functions.insert("IF".to_string(), 1);
        functions.insert("COUNT".to_string(), 0);
        functions.insert("AVERAGE".to_string(), 5);
        functions.insert("MAX".to_string(), 7);
        functions.insert("MIN".to_string(), 6);
        functions.insert("VLOOKUP".to_string(), 102);
        functions.insert("CONCATENATE".to_string(), 336);
        functions.insert("LEFT".to_string(), 115);
        functions.insert("RIGHT".to_string(), 116);
        functions.insert("MID".to_string(), 31);
        functions.insert("LEN".to_string(), 32);
        functions.insert("ROUND".to_string(), 27);
        functions.insert("ABS".to_string(), 24);

        Self { functions }
    }

    /// Tokenize a formula string to RPN tokens
    ///
    /// # Arguments
    ///
    /// * `formula` - Formula string (without leading '=')
    ///
    /// # Returns
    ///
    /// Vector of Ptg tokens in RPN order
    pub fn tokenize(&self, formula: &str) -> Result<Vec<Ptg>, XlsError> {
        let formula = formula.trim();
        if formula.is_empty() {
            return Ok(Vec::new());
        }

        // Simple tokenization using Shunting Yard algorithm
        let mut output = Vec::new();
        let mut operators = Vec::new();
        let mut i = 0;
        let mut expect_operand = true;
        let chars: Vec<char> = formula.chars().collect();

        while i < chars.len() {
            let c = chars[i];

            // Skip whitespace
            if c.is_whitespace() {
                i += 1;
                continue;
            }

            // Number literal
            if c.is_ascii_digit() || c == '.' {
                let start = i;
                while i < chars.len() && (chars[i].is_ascii_digit() || chars[i] == '.') {
                    i += 1;
                }
                let num_str: String = chars[start..i].iter().collect();
                if num_str.contains('.') {
                    let num = num_str.parse::<f64>().map_err(|_| {
                        XlsError::InvalidData(format!("Invalid number: {}", num_str))
                    })?;
                    output.push(Ptg::Num(num));
                } else {
                    let num = num_str.parse::<u16>().map_err(|_| {
                        XlsError::InvalidData(format!("Invalid integer: {}", num_str))
                    })?;
                    output.push(Ptg::Int(num));
                }
                expect_operand = false;
                continue;
            }

            // String literal
            if c == '"' {
                i += 1;
                let mut value = String::new();
                let mut closed = false;
                while i < chars.len() {
                    if chars[i] == '"' {
                        if i + 1 < chars.len() && chars[i + 1] == '"' {
                            value.push('"');
                            i += 2;
                            continue;
                        }
                        i += 1;
                        closed = true;
                        break;
                    }
                    value.push(chars[i]);
                    i += 1;
                }
                if !closed {
                    return Err(XlsError::InvalidFormula(
                        "Unterminated string literal".to_string(),
                    ));
                }
                output.push(Ptg::Str(value));
                expect_operand = false;
                continue;
            }

            // Cell reference or function
            if c.is_ascii_alphabetic() || c == '$' {
                let start = i;
                while i < chars.len() && (chars[i].is_ascii_alphanumeric() || chars[i] == '$') {
                    i += 1;
                }
                let token: String = chars[start..i].iter().collect();

                // Check if it's a function call
                if i < chars.len() && chars[i] == '(' {
                    let func_name = token.to_uppercase();
                    let func_idx = self.functions.get(&func_name).copied().ok_or_else(|| {
                        XlsError::InvalidData(format!("Unknown function: {func_name}"))
                    })?;
                    let mut next = i + 1;
                    while next < chars.len() && chars[next].is_whitespace() {
                        next += 1;
                    }
                    let argc = u8::from(next >= chars.len() || chars[next] != ')');
                    operators.push(("FUNC", func_idx, argc));
                    operators.push(("(", 0, 0));
                    i += 1; // Skip '('
                    expect_operand = true;
                } else {
                    if token.eq_ignore_ascii_case("TRUE") {
                        output.push(Ptg::Bool(true));
                        expect_operand = false;
                        continue;
                    }
                    if token.eq_ignore_ascii_case("FALSE") {
                        output.push(Ptg::Bool(false));
                        expect_operand = false;
                        continue;
                    }
                    output.push(parse_cell_ref(&token)?);
                    expect_operand = false;
                }
                continue;
            }

            // Operators
            if i + 1 < chars.len() {
                let two_char: String = chars[i..i + 2].iter().collect();
                if two_char == "<>" || two_char == "<=" || two_char == ">=" {
                    if expect_operand {
                        return Err(XlsError::InvalidFormula(format!(
                            "Operator {two_char} is missing its left operand"
                        )));
                    }
                    self.handle_operator(&mut output, &mut operators, &two_char)?;
                    expect_operand = true;
                    i += 2;
                    continue;
                }
            }

            let op_str = c.to_string();

            match op_str.as_str() {
                "(" => {
                    operators.push(("(", 0, 0));
                    expect_operand = true;
                },
                ")" => {
                    let zero_arg_function = operators.last().is_some_and(|(op, _, _)| *op == "(")
                        && operators
                            .get(operators.len().saturating_sub(2))
                            .is_some_and(|(op, _, argc)| *op == "FUNC" && *argc == 0);
                    if expect_operand && !zero_arg_function {
                        output.push(Ptg::MissArg);
                    }
                    while let Some((op, func_idx, argc)) = operators.pop() {
                        if op == "(" {
                            break;
                        }
                        if op == "FUNC" {
                            output.push(Ptg::Func(func_idx, argc));
                        } else {
                            self.push_operator(&mut output, op)?;
                        }
                    }
                    if operators.last().is_some_and(|(op, _, _)| *op == "FUNC") {
                        let (_, func_idx, argc) = operators.pop().ok_or_else(|| {
                            XlsError::InvalidFormula(
                                "function operator stack became inconsistent".to_string(),
                            )
                        })?;
                        output.push(Ptg::Func(func_idx, argc));
                    }
                    expect_operand = false;
                },
                "+" | "-" | "*" | "/" | "^" | "&" | "=" | "<" | ">" | ":" => {
                    if expect_operand && (op_str == "+" || op_str == "-") {
                        self.handle_operator(
                            &mut output,
                            &mut operators,
                            if op_str == "+" { "u+" } else { "u-" },
                        )?;
                    } else if expect_operand {
                        return Err(XlsError::InvalidFormula(format!(
                            "Operator {op_str} is missing its left operand"
                        )));
                    } else {
                        self.handle_operator(&mut output, &mut operators, &op_str)?;
                        expect_operand = true;
                    }
                },
                "%" => {
                    if expect_operand {
                        return Err(XlsError::InvalidFormula(
                            "Percent operator is missing its operand".to_string(),
                        ));
                    }
                    output.push(Ptg::Percent);
                },
                "," => {
                    if expect_operand {
                        output.push(Ptg::MissArg);
                    }
                    // Argument separator - pop operators until '('
                    while let Some(&(top_op, _, _)) = operators.last() {
                        if top_op == "(" {
                            break;
                        }
                        let (op, func_idx, argc) = operators.pop().ok_or_else(|| {
                            XlsError::InvalidFormula(
                                "argument operator stack became inconsistent".to_string(),
                            )
                        })?;
                        if op == "FUNC" {
                            output.push(Ptg::Func(func_idx, argc));
                        } else {
                            self.push_operator(&mut output, op)?;
                        }
                    }
                    let open_paren = operators
                        .iter()
                        .rposition(|(op, _, _)| *op == "(")
                        .ok_or_else(|| {
                            XlsError::InvalidData("Argument separator outside function".to_string())
                        })?;
                    let function = open_paren.checked_sub(1).ok_or_else(|| {
                        XlsError::InvalidData("Argument separator outside function".to_string())
                    })?;
                    let (op, _, argc) = operators.get_mut(function).ok_or_else(|| {
                        XlsError::InvalidFormula(
                            "argument operator stack became inconsistent".to_string(),
                        )
                    })?;
                    if *op != "FUNC" {
                        return Err(XlsError::InvalidData(
                            "Argument separator outside function".to_string(),
                        ));
                    }
                    *argc = argc.checked_add(1).ok_or_else(|| {
                        XlsError::InvalidData("Too many function arguments".to_string())
                    })?;
                    expect_operand = true;
                },
                _ => {
                    return Err(XlsError::InvalidData(format!(
                        "Unknown operator: {}",
                        op_str
                    )));
                },
            }

            // CRITICAL: Increment index after processing operator to avoid infinite loop
            i += 1;
        }

        if expect_operand && !output.is_empty() {
            return Err(XlsError::InvalidFormula(
                "Formula ends with an operator".to_string(),
            ));
        }

        // Pop remaining operators
        while let Some((op, func_idx, argc)) = operators.pop() {
            if op == "(" {
                return Err(XlsError::InvalidData("Mismatched parentheses".to_string()));
            }
            if op == "FUNC" {
                output.push(Ptg::Func(func_idx, argc));
            } else {
                self.push_operator(&mut output, op)?;
            }
        }

        Ok(output)
    }

    fn push_operator(&self, output: &mut Vec<Ptg>, op: &str) -> Result<(), XlsError> {
        let ptg = match op {
            "+" => Ptg::Add,
            "-" => Ptg::Sub,
            "*" => Ptg::Mul,
            "/" => Ptg::Div,
            "^" => Ptg::Power,
            "&" => Ptg::Concat,
            "=" => Ptg::Eq,
            "<>" => Ptg::Ne,
            "<" => Ptg::Lt,
            "<=" => Ptg::Le,
            ">" => Ptg::Gt,
            ">=" => Ptg::Ge,
            ":" => Ptg::Range,
            "u+" => Ptg::UnaryPlus,
            "u-" => Ptg::UnaryMinus,
            _ => return Err(XlsError::InvalidData(format!("Unknown operator: {}", op))),
        };
        output.push(ptg);
        Ok(())
    }

    fn handle_operator(
        &self,
        output: &mut Vec<Ptg>,
        operators: &mut Vec<(&'static str, u16, u8)>,
        op: &str,
    ) -> Result<(), XlsError> {
        // Convert string to static str for storage
        let op_static: &'static str = match op {
            "+" => "+",
            "-" => "-",
            "*" => "*",
            "/" => "/",
            "^" => "^",
            "&" => "&",
            "=" => "=",
            "<>" => "<>",
            "<" => "<",
            "<=" => "<=",
            ">" => ">",
            ">=" => ">=",
            ":" => ":",
            "u+" => "u+",
            "u-" => "u-",
            _ => return Err(XlsError::InvalidData(format!("Unknown operator: {}", op))),
        };

        let prec = get_precedence(op);
        while let Some(&(top_op, _, _)) = operators.last() {
            if top_op == "(" {
                break;
            }
            let left_associative = !matches!(op, "u+" | "u-" | "^");
            if get_precedence(top_op) > prec || (left_associative && get_precedence(top_op) == prec)
            {
                let (op, func_idx, argc) = operators.pop().ok_or_else(|| {
                    XlsError::InvalidFormula("operator stack became inconsistent".to_string())
                })?;
                if op == "FUNC" {
                    output.push(Ptg::Func(func_idx, argc));
                } else {
                    self.push_operator(output, op)?;
                }
            } else {
                break;
            }
        }
        operators.push((op_static, 0, 0));
        Ok(())
    }
}

impl Default for FormulaTokenizer {
    fn default() -> Self {
        Self::new()
    }
}

/// Encode Ptg tokens to binary format for BIFF8
pub fn encode_ptg_tokens(tokens: &[Ptg]) -> Vec<u8> {
    let mut bytes = Vec::new();

    for token in tokens {
        match token {
            Ptg::Int(val) => {
                bytes.push(0x1E); // PtgInt
                bytes.extend_from_slice(&val.to_le_bytes());
            },
            Ptg::Num(val) => {
                bytes.push(0x1F); // PtgNum
                bytes.extend_from_slice(&val.to_le_bytes());
            },
            Ptg::Str(s) => {
                bytes.push(0x17); // PtgStr
                let mut utf16: Vec<u16> = s.encode_utf16().take(255).collect();
                if utf16
                    .last()
                    .is_some_and(|unit| (0xd800..=0xdbff).contains(unit))
                {
                    utf16.pop();
                }
                bytes.push(utf16.len() as u8);
                if utf16.iter().all(|unit| *unit <= 0xff) {
                    bytes.push(0); // Compressed Unicode
                    bytes.extend(utf16.iter().map(|unit| *unit as u8));
                } else {
                    bytes.push(1); // UTF-16LE
                    for unit in utf16 {
                        bytes.extend_from_slice(&unit.to_le_bytes());
                    }
                }
            },
            Ptg::Bool(value) => {
                bytes.extend_from_slice(&[0x1d, u8::from(*value)]);
            },
            Ptg::Ref(reference) => {
                bytes.push(0x24); // PtgRef
                bytes.extend_from_slice(&reference.row().to_le_bytes());
                bytes.extend_from_slice(&reference.col_flags().to_le_bytes());
            },
            Ptg::Area(area) => {
                // BIFF8 PtgArea (2D area reference)
                let first = area.first();
                let last = area.last();
                bytes.push(0x25); // PtgArea
                bytes.extend_from_slice(&first.row().to_le_bytes());
                bytes.extend_from_slice(&last.row().to_le_bytes());
                bytes.extend_from_slice(&first.col_flags().to_le_bytes());
                bytes.extend_from_slice(&last.col_flags().to_le_bytes());
            },
            Ptg::Area3d(ixti, area) => {
                // BIFF8 PtgArea3d (3D area reference)
                //
                // Layout: opcode (1 byte) + ixti (2 bytes) + r1 (2) +
                // r2 (2) + c1 (2) + c2 (2).
                let first = area.first();
                let last = area.last();
                bytes.push(0x3B); // PtgArea3d
                bytes.extend_from_slice(&ixti.to_le_bytes());
                bytes.extend_from_slice(&first.row().to_le_bytes());
                bytes.extend_from_slice(&last.row().to_le_bytes());
                bytes.extend_from_slice(&first.col_flags().to_le_bytes());
                bytes.extend_from_slice(&last.col_flags().to_le_bytes());
            },
            Ptg::Add => bytes.push(0x03),
            Ptg::Sub => bytes.push(0x04),
            Ptg::Mul => bytes.push(0x05),
            Ptg::Div => bytes.push(0x06),
            Ptg::Power => bytes.push(0x07),
            Ptg::Concat => bytes.push(0x08),
            Ptg::Lt => bytes.push(0x09),
            Ptg::Le => bytes.push(0x0A),
            Ptg::Eq => bytes.push(0x0B),
            Ptg::Ge => bytes.push(0x0C),
            Ptg::Gt => bytes.push(0x0D),
            Ptg::Ne => bytes.push(0x0E),
            Ptg::Range => bytes.push(0x11),
            Ptg::UnaryPlus => bytes.push(0x12),
            Ptg::UnaryMinus => bytes.push(0x13),
            Ptg::Percent => bytes.push(0x14),
            Ptg::Func(func_idx, argc) => {
                bytes.push(0x42); // PtgFuncVar, value operand class
                bytes.push(*argc);
                bytes.extend_from_slice(&func_idx.to_le_bytes());
            },
            Ptg::Paren => bytes.push(0x15),
            Ptg::MissArg => bytes.push(0x16),
        }
    }

    bytes
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_cell_ref() {
        let ref_a1 = parse_cell_ref("A1").unwrap();
        assert_eq!(ref_a1, Ptg::Ref(Ref::new(0, 0).rel_row(true).rel_col(true)));

        let ref_abs = parse_cell_ref("$B$2").unwrap();
        assert_eq!(ref_abs, Ptg::Ref(Ref::new(1, 1)));
    }

    #[test]
    fn test_parse_cell_ref_row_abs() {
        let ref_row_abs = parse_cell_ref("A$5").unwrap();
        assert_eq!(ref_row_abs, Ptg::Ref(Ref::new(4, 0).rel_col(true)));
    }

    #[test]
    fn test_parse_cell_ref_col_abs() {
        let ref_col_abs = parse_cell_ref("$C10").unwrap();
        assert_eq!(ref_col_abs, Ptg::Ref(Ref::new(9, 2).rel_row(true)));
    }

    #[test]
    fn test_parse_cell_ref_biff8_edges() {
        let ref_aa1 = parse_cell_ref("AA1").unwrap();
        assert_eq!(
            ref_aa1,
            Ptg::Ref(Ref::new(0, 26).rel_row(true).rel_col(true))
        );

        let last = parse_cell_ref("IV65536").unwrap();
        assert_eq!(
            last,
            Ptg::Ref(Ref::new(u16::MAX, u8::MAX).rel_row(true).rel_col(true))
        );
    }

    #[test]
    fn test_parse_cell_ref_invalid() {
        for value in ["", "123", "ABC", "A0", "A65537", "IW1", "ZZZZ1"] {
            assert!(
                matches!(
                    parse_cell_ref(value),
                    Err(XlsError::InvalidCellReference(reference)) if reference == value
                ),
                "unexpected result for {value:?}"
            );
        }

        let oversized = format!("{}1", "Z".repeat(4_096));
        assert!(matches!(
            parse_cell_ref(&oversized),
            Err(XlsError::InvalidCellReference(reference)) if reference == oversized
        ));
    }

    #[test]
    fn checked_reference_and_area_construction_rejects_invalid_bounds() {
        assert_eq!(Ref::checked(0, 255).unwrap(), Ref::new(0, 255));
        assert!(matches!(
            Ref::checked(0, 256),
            Err(XlsError::InvalidCellReference(_))
        ));
        assert!(matches!(
            Area::new(Ref::new(1, 0), Ref::new(0, 0)),
            Err(XlsError::InvalidCellReference(_))
        ));
        assert!(matches!(
            Area::new(Ref::new(0, 1), Ref::new(0, 0)),
            Err(XlsError::InvalidCellReference(_))
        ));
    }

    #[test]
    fn tokenizer_preserves_invalid_reference_errors_without_unwinding() {
        let outcome = std::panic::catch_unwind(|| FormulaTokenizer::new().tokenize("ZZZZ1"));
        assert!(matches!(
            outcome,
            Ok(Err(XlsError::InvalidCellReference(reference))) if reference == "ZZZZ1"
        ));
        assert!(matches!(
            FormulaTokenizer::new().tokenize("IW1"),
            Err(XlsError::InvalidCellReference(reference)) if reference == "IW1"
        ));
    }

    #[test]
    fn test_tokenize_simple() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("A1+B1").unwrap();
        assert_eq!(tokens.len(), 3); // A1, B1, +
    }

    #[test]
    fn test_tokenize_complex() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("A1+B1*2").unwrap();
        // Should be: A1, B1, 2, *, + (RPN)
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn test_tokenize_empty() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("").unwrap();
        assert!(tokens.is_empty());
    }

    #[test]
    fn test_tokenize_whitespace() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("  A1  +  B1  ").unwrap();
        assert_eq!(tokens.len(), 3);
    }

    #[test]
    fn test_tokenize_numbers() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("123+456.78").unwrap();
        assert_eq!(tokens.len(), 3);
        assert!(matches!(tokens[0], Ptg::Int(123)));
        assert!(matches!(tokens[1], Ptg::Num(456.78)));
        assert!(matches!(tokens[2], Ptg::Add));
    }

    #[test]
    fn test_tokenize_string() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("\"Hello World\"").unwrap();
        assert_eq!(tokens.len(), 1);
        assert!(matches!(&tokens[0], Ptg::Str(s) if s == "Hello World"));
    }

    #[test]
    fn test_tokenize_escaped_string_and_booleans() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("IF(TRUE,\"a\"\"b\",FALSE)").unwrap();
        assert!(matches!(tokens[0], Ptg::Bool(true)));
        assert!(matches!(&tokens[1], Ptg::Str(value) if value == "a\"b"));
        assert!(matches!(tokens[2], Ptg::Bool(false)));
        assert!(matches!(tokens[3], Ptg::Func(1, 3)));
        assert!(tokenizer.tokenize("\"unterminated").is_err());
    }

    #[test]
    fn test_tokenize_unary_percent_and_missing_arguments() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("--A1+50%").unwrap();
        assert!(matches!(tokens[0], Ptg::Ref(_)));
        assert!(matches!(tokens[1], Ptg::UnaryMinus));
        assert!(matches!(tokens[2], Ptg::UnaryMinus));
        assert!(matches!(tokens[3], Ptg::Int(50)));
        assert!(matches!(tokens[4], Ptg::Percent));
        assert!(matches!(tokens[5], Ptg::Add));

        let tokens = tokenizer.tokenize("IF(,A1,)").unwrap();
        assert!(matches!(tokens[0], Ptg::MissArg));
        assert!(matches!(tokens[1], Ptg::Ref(_)));
        assert!(matches!(tokens[2], Ptg::MissArg));
        assert!(matches!(tokens[3], Ptg::Func(1, 3)));
        assert!(tokenizer.tokenize("A1+").is_err());
        assert!(tokenizer.tokenize("*A1").is_err());
    }

    #[test]
    fn test_tokenize_subtraction() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("A1-B1").unwrap();
        assert_eq!(tokens.len(), 3);
        assert!(matches!(tokens[2], Ptg::Sub));
    }

    #[test]
    fn test_tokenize_multiplication() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("A1*B1").unwrap();
        assert_eq!(tokens.len(), 3);
        assert!(matches!(tokens[2], Ptg::Mul));
    }

    #[test]
    fn test_tokenize_division() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("A1/B1").unwrap();
        assert_eq!(tokens.len(), 3);
        assert!(matches!(tokens[2], Ptg::Div));
    }

    #[test]
    fn test_tokenize_power() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("A1^2").unwrap();
        assert_eq!(tokens.len(), 3);
        assert!(matches!(tokens[2], Ptg::Power));
    }

    #[test]
    fn test_tokenize_concatenation() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("A1&B1").unwrap();
        assert_eq!(tokens.len(), 3);
        assert!(matches!(tokens[2], Ptg::Concat));
    }

    #[test]
    fn test_tokenize_comparison_operators() {
        let tokenizer = FormulaTokenizer::new();

        let tokens_eq = tokenizer.tokenize("A1=B1").unwrap();
        assert!(matches!(tokens_eq[2], Ptg::Eq));

        let tokens_ne = tokenizer.tokenize("A1<>B1").unwrap();
        assert!(matches!(tokens_ne[2], Ptg::Ne));

        let tokens_lt = tokenizer.tokenize("A1<B1").unwrap();
        assert!(matches!(tokens_lt[2], Ptg::Lt));

        let tokens_le = tokenizer.tokenize("A1<=B1").unwrap();
        assert!(matches!(tokens_le[2], Ptg::Le));

        let tokens_gt = tokenizer.tokenize("A1>B1").unwrap();
        assert!(matches!(tokens_gt[2], Ptg::Gt));

        let tokens_ge = tokenizer.tokenize("A1>=B1").unwrap();
        assert!(matches!(tokens_ge[2], Ptg::Ge));
    }

    #[test]
    fn test_tokenize_parentheses() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("(A1+B1)*C1").unwrap();
        // Should be: A1, B1, +, C1, * (RPN)
        assert_eq!(tokens.len(), 5);
        assert!(matches!(tokens[2], Ptg::Add));
        assert!(matches!(tokens[4], Ptg::Mul));
    }

    #[test]
    fn test_tokenize_function() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("SUM(A1)").unwrap();
        assert!(tokens.iter().any(|t| matches!(t, Ptg::Func(4, 1))));
    }

    #[test]
    fn test_tokenize_function_arguments_and_following_operator() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("ROUND(A1,2)+SUM()").unwrap();
        assert!(matches!(tokens[2], Ptg::Func(27, 2)));
        assert!(matches!(tokens[3], Ptg::Func(4, 0)));
        assert!(matches!(tokens[4], Ptg::Add));
    }

    #[test]
    fn test_tokenize_unknown_function_is_rejected() {
        let tokenizer = FormulaTokenizer::new();
        assert!(tokenizer.tokenize("MADEUP(A1)").is_err());
    }

    #[test]
    fn test_tokenize_precedence() {
        let tokenizer = FormulaTokenizer::new();
        // Multiplication has higher precedence than addition
        let tokens = tokenizer.tokenize("A1+B1*C1").unwrap();
        // Should be: A1, B1, C1, *, + (RPN)
        assert_eq!(tokens.len(), 5);
        assert!(matches!(tokens[3], Ptg::Mul));
        assert!(matches!(tokens[4], Ptg::Add));
    }

    #[test]
    fn test_ptg_enum_variants() {
        // Test that all Ptg variants can be created
        let reference = Ref::new(0, 0);
        let area = Area::new(reference, Ref::new(1, 1)).unwrap();
        let _ = Ptg::Int(42);
        let _ = Ptg::Num(std::f64::consts::PI);
        let _ = Ptg::Str("test".to_string());
        let _ = Ptg::Bool(true);
        let _ = Ptg::Ref(reference);
        let _ = Ptg::Area(area);
        let _ = Ptg::Area3d(0, area);
        let _ = Ptg::Add;
        let _ = Ptg::Sub;
        let _ = Ptg::Mul;
        let _ = Ptg::Div;
        let _ = Ptg::Power;
        let _ = Ptg::Concat;
        let _ = Ptg::Lt;
        let _ = Ptg::Le;
        let _ = Ptg::Eq;
        let _ = Ptg::Ge;
        let _ = Ptg::Gt;
        let _ = Ptg::Ne;
        let _ = Ptg::Range;
        let _ = Ptg::UnaryPlus;
        let _ = Ptg::UnaryMinus;
        let _ = Ptg::Percent;
        let _ = Ptg::Func(0, 1);
        let _ = Ptg::Paren;
        let _ = Ptg::MissArg;
    }

    #[test]
    fn test_ptg_clone() {
        let ptg = Ptg::Ref(Ref::new(1, 2).rel_row(true));
        let cloned = ptg.clone();
        assert_eq!(ptg, cloned);
    }

    #[test]
    fn test_ptg_partial_eq() {
        assert_eq!(Ptg::Add, Ptg::Add);
        assert_eq!(Ptg::Int(42), Ptg::Int(42));
        assert_ne!(Ptg::Int(42), Ptg::Int(43));
    }

    #[test]
    fn test_ptg_debug() {
        let ptg = Ptg::Ref(Ref::new(0, 0).rel_row(true).rel_col(true));
        let debug_str = format!("{:?}", ptg);
        assert!(debug_str.contains("Ref"));
    }

    #[test]
    fn test_tokenizer_default() {
        let tokenizer: FormulaTokenizer = Default::default();
        let tokens = tokenizer.tokenize("A1").unwrap();
        assert_eq!(tokens.len(), 1);
    }

    #[test]
    fn test_encode_ptg_tokens() {
        let tokens = vec![Ptg::Int(42), Ptg::Add, Ptg::Num(std::f64::consts::PI)];
        let bytes = encode_ptg_tokens(&tokens);
        assert!(!bytes.is_empty());
        assert_eq!(bytes[0], 0x1E); // PtgInt opcode
    }

    #[test]
    fn test_encode_ptg_ref() {
        let absolute = Ref::new(5, 3);
        let row_relative = absolute.rel_row(true);
        let column_relative = absolute.rel_col(true);
        let relative = row_relative.rel_col(true);

        assert_eq!(
            encode_ptg_tokens(&[Ptg::Ref(absolute)]),
            [0x24, 0x05, 0x00, 0x03, 0x00]
        );
        assert_eq!(
            encode_ptg_tokens(&[Ptg::Ref(row_relative)]),
            [0x24, 0x05, 0x00, 0x03, 0x80]
        );
        assert_eq!(
            encode_ptg_tokens(&[Ptg::Ref(column_relative)]),
            [0x24, 0x05, 0x00, 0x03, 0x40]
        );
        assert_eq!(
            encode_ptg_tokens(&[Ptg::Ref(relative)]),
            [0x24, 0x05, 0x00, 0x03, 0xc0]
        );
        assert_eq!(
            encode_ptg_tokens(&[Ptg::Ref(Ref::new(u16::MAX, u8::MAX))]),
            [0x24, 0xff, 0xff, 0xff, 0x00]
        );
    }

    #[test]
    fn test_encode_ptg_str() {
        let tokens = vec![Ptg::Str("Test".to_string())];
        let bytes = encode_ptg_tokens(&tokens);
        assert_eq!(bytes[0], 0x17); // PtgStr opcode
        assert_eq!(bytes[1], 4); // String length
    }

    #[test]
    fn test_encode_ptg_str_as_utf16_when_required() {
        let bytes = encode_ptg_tokens(&[Ptg::Str("你好".to_string())]);
        assert_eq!(bytes, [0x17, 2, 1, 0x60, 0x4f, 0x7d, 0x59]);
    }

    #[test]
    fn test_encode_ptg_bool() {
        assert_eq!(encode_ptg_tokens(&[Ptg::Bool(true)]), [0x1d, 1]);
        assert_eq!(encode_ptg_tokens(&[Ptg::Bool(false)]), [0x1d, 0]);
    }

    #[test]
    fn test_encode_ptg_str_does_not_split_surrogate_pair_at_limit() {
        let value = format!("{}😀", "a".repeat(254));
        let bytes = encode_ptg_tokens(&[Ptg::Str(value)]);
        assert_eq!(bytes[1], 254);
        assert_eq!(bytes.len(), 257);
    }

    #[test]
    fn test_encode_ptg_area() {
        let area = Area::new(Ref::new(0, 0), Ref::new(5, 3)).unwrap();
        assert_eq!(
            encode_ptg_tokens(&[Ptg::Area(area)]),
            [0x25, 0x00, 0x00, 0x05, 0x00, 0x00, 0x00, 0x03, 0x00]
        );
    }

    #[test]
    fn test_encode_ptg_area3d() {
        let area = Area::new(Ref::new(0, 0), Ref::new(5, 3)).unwrap();
        assert_eq!(
            encode_ptg_tokens(&[Ptg::Area3d(2, area)]),
            [
                0x3b, 0x02, 0x00, 0x00, 0x00, 0x05, 0x00, 0x00, 0x00, 0x03, 0x00,
            ]
        );
    }

    #[test]
    fn test_encode_ptg_func() {
        let tokens = vec![Ptg::Func(4, 1)]; // SUM with 1 arg
        let bytes = encode_ptg_tokens(&tokens);
        assert_eq!(bytes[0], 0x42); // PtgFuncVar opcode, value operand class
        assert_eq!(bytes[1], 1); // Arg count
    }

    #[test]
    fn test_encode_all_operators() {
        let operators = vec![
            (Ptg::Add, 0x03),
            (Ptg::Sub, 0x04),
            (Ptg::Mul, 0x05),
            (Ptg::Div, 0x06),
            (Ptg::Power, 0x07),
            (Ptg::Concat, 0x08),
            (Ptg::Lt, 0x09),
            (Ptg::Le, 0x0A),
            (Ptg::Eq, 0x0B),
            (Ptg::Ge, 0x0C),
            (Ptg::Gt, 0x0D),
            (Ptg::Ne, 0x0E),
            (Ptg::Range, 0x11),
            (Ptg::UnaryPlus, 0x12),
            (Ptg::UnaryMinus, 0x13),
            (Ptg::Percent, 0x14),
            (Ptg::Paren, 0x15),
            (Ptg::MissArg, 0x16),
        ];

        for (ptg, expected_opcode) in operators {
            let bytes = encode_ptg_tokens(std::slice::from_ref(&ptg));
            assert_eq!(bytes[0], expected_opcode, "Opcode mismatch for {:?}", ptg);
        }
    }

    #[test]
    fn test_tokenize_cell_range() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("SUM(A1:B1)").unwrap();
        assert_eq!(
            tokens[0],
            Ptg::Ref(Ref::new(0, 0).rel_row(true).rel_col(true))
        );
        assert_eq!(
            tokens[1],
            Ptg::Ref(Ref::new(0, 1).rel_row(true).rel_col(true))
        );
        assert!(matches!(tokens[2], Ptg::Range));
        assert!(matches!(tokens[3], Ptg::Func(4, 1)));
    }
}
