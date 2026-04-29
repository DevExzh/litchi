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
//! Tokens: [PtgRef(A1), PtgRef(B1), PtgInt(2), PtgMul, PtgAdd]
//! ```

use super::super::XlsError;
use std::collections::HashMap;

/// Ptg (Parse Thing) token types
#[derive(Debug, Clone, PartialEq)]
pub enum Ptg {
    /// Integer constant
    PtgInt(u16),
    /// Number constant
    PtgNum(f64),
    /// String constant
    PtgStr(String),
    /// Cell reference (row, col, relative flags)
    PtgRef(u16, u16, bool, bool),
    /// Area reference (r1, c1, r2, c2)
    PtgArea(u16, u16, u16, u16),
    /// 3D area reference (ixti, r1, r2, c1, c2)
    ///
    /// Used by defined names and other structures that require
    /// NameParsedFormula, which MUST use 3D references instead of
    /// plain 2D PtgArea in BIFF8.
    PtgArea3d(u16, u16, u16, u16, u16),
    /// Addition operator
    PtgAdd,
    /// Subtraction operator
    PtgSub,
    /// Multiplication operator
    PtgMul,
    /// Division operator
    PtgDiv,
    /// Power operator
    PtgPower,
    /// Concatenation operator
    PtgConcat,
    /// Less than
    PtgLT,
    /// Less than or equal
    PtgLE,
    /// Equal
    PtgEQ,
    /// Greater than or equal
    PtgGE,
    /// Greater than
    PtgGT,
    /// Not equal
    PtgNE,
    /// Function call (function index, arg count)
    PtgFunc(u16, u8),
    /// Parentheses
    PtgParen,
    /// Missing argument
    PtgMissArg,
}

/// Operator precedence
fn get_precedence(op: &str) -> u8 {
    match op {
        "^" => 4,
        "*" | "/" => 3,
        "+" | "-" => 2,
        "&" => 2,
        "=" | "<>" | "<" | "<=" | ">" | ">=" => 1,
        _ => 0,
    }
}

/// Parse a cell reference like "A1" or "$B$2" into a `PtgRef` token.
///
/// This is exposed as `pub(crate)` so that other writer components
/// (for example, named range handling) can reuse the same parsing
/// logic and stay consistent with formula tokenization.
pub(crate) fn parse_cell_ref(s: &str) -> Result<Ptg, XlsError> {
    let s = s.trim();
    let mut col_abs = false;
    let mut row_abs = false;
    let mut chars = s.chars().peekable();

    // Check for absolute column
    if chars.peek() == Some(&'$') {
        col_abs = true;
        chars.next();
    }

    // Parse column (A-Z, AA-ZZ, etc.)
    let mut col_str = String::new();
    while let Some(&c) = chars.peek() {
        if c.is_ascii_alphabetic() {
            col_str.push(chars.next().unwrap());
        } else {
            break;
        }
    }

    if col_str.is_empty() {
        return Err(XlsError::InvalidData(format!(
            "Invalid cell reference: {}",
            s
        )));
    }

    // Convert column letters to number (A=0, B=1, ..., Z=25, AA=26, etc.)
    let mut col = 0u16;
    for c in col_str.chars() {
        col = col * 26 + (c.to_ascii_uppercase() as u16 - 'A' as u16 + 1);
    }
    col -= 1; // Convert to 0-based

    // Check for absolute row
    if chars.peek() == Some(&'$') {
        row_abs = true;
        chars.next();
    }

    // Parse row number
    let row_str: String = chars.collect();
    let row = row_str
        .parse::<u16>()
        .map_err(|_| XlsError::InvalidData(format!("Invalid row number: {}", row_str)))?;

    if row == 0 {
        return Err(XlsError::InvalidData("Row must be >= 1".to_string()));
    }

    Ok(Ptg::PtgRef(row - 1, col, !row_abs, !col_abs))
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
                    output.push(Ptg::PtgNum(num));
                } else {
                    let num = num_str.parse::<u16>().map_err(|_| {
                        XlsError::InvalidData(format!("Invalid integer: {}", num_str))
                    })?;
                    output.push(Ptg::PtgInt(num));
                }
                continue;
            }

            // String literal
            if c == '"' {
                i += 1; // Skip opening quote
                let start = i;
                while i < chars.len() && chars[i] != '"' {
                    i += 1;
                }
                let s: String = chars[start..i].iter().collect();
                output.push(Ptg::PtgStr(s));
                i += 1; // Skip closing quote
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
                    if let Some(&func_idx) = self.functions.get(&func_name) {
                        // For simplicity, assume 1 argument (full implementation would parse args)
                        operators.push(("FUNC", func_idx, 1));
                    }
                    operators.push(("(", 0, 0));
                    i += 1; // Skip '('
                } else {
                    // Try to parse as cell reference
                    match parse_cell_ref(&token) {
                        Ok(ptg) => output.push(ptg),
                        Err(_) => {
                            return Err(XlsError::InvalidData(format!("Unknown token: {}", token)));
                        },
                    }
                }
                continue;
            }

            // Operators
            if i + 1 < chars.len() {
                let two_char: String = chars[i..i + 2].iter().collect();
                if two_char == "<>" || two_char == "<=" || two_char == ">=" {
                    self.handle_operator(&mut output, &mut operators, &two_char)?;
                    i += 2;
                    continue;
                }
            }

            let op_str = c.to_string();

            match op_str.as_str() {
                "(" => operators.push(("(", 0, 0)),
                ")" => {
                    while let Some((op, func_idx, argc)) = operators.pop() {
                        if op == "(" {
                            break;
                        }
                        if op == "FUNC" {
                            output.push(Ptg::PtgFunc(func_idx, argc));
                        } else {
                            self.push_operator(&mut output, op)?;
                        }
                    }
                },
                "+" | "-" | "*" | "/" | "^" | "&" | "=" | "<" | ">" => {
                    self.handle_operator(&mut output, &mut operators, &op_str)?;
                },
                "," => {
                    // Argument separator - pop operators until '('
                    while let Some(&(top_op, _, _)) = operators.last() {
                        if top_op == "(" {
                            break;
                        }
                        let (op, func_idx, argc) = operators.pop().unwrap();
                        if op == "FUNC" {
                            output.push(Ptg::PtgFunc(func_idx, argc));
                        } else {
                            self.push_operator(&mut output, op)?;
                        }
                    }
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

        // Pop remaining operators
        while let Some((op, func_idx, argc)) = operators.pop() {
            if op == "(" {
                return Err(XlsError::InvalidData("Mismatched parentheses".to_string()));
            }
            if op == "FUNC" {
                output.push(Ptg::PtgFunc(func_idx, argc));
            } else {
                self.push_operator(&mut output, op)?;
            }
        }

        Ok(output)
    }

    fn push_operator(&self, output: &mut Vec<Ptg>, op: &str) -> Result<(), XlsError> {
        let ptg = match op {
            "+" => Ptg::PtgAdd,
            "-" => Ptg::PtgSub,
            "*" => Ptg::PtgMul,
            "/" => Ptg::PtgDiv,
            "^" => Ptg::PtgPower,
            "&" => Ptg::PtgConcat,
            "=" => Ptg::PtgEQ,
            "<>" => Ptg::PtgNE,
            "<" => Ptg::PtgLT,
            "<=" => Ptg::PtgLE,
            ">" => Ptg::PtgGT,
            ">=" => Ptg::PtgGE,
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
            _ => return Err(XlsError::InvalidData(format!("Unknown operator: {}", op))),
        };

        let prec = get_precedence(op);
        while let Some(&(top_op, _, _)) = operators.last() {
            if top_op == "(" {
                break;
            }
            if get_precedence(top_op) >= prec {
                let (op, func_idx, argc) = operators.pop().unwrap();
                if op == "FUNC" {
                    output.push(Ptg::PtgFunc(func_idx, argc));
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
            Ptg::PtgInt(val) => {
                bytes.push(0x1E); // PtgInt
                bytes.extend_from_slice(&val.to_le_bytes());
            },
            Ptg::PtgNum(val) => {
                bytes.push(0x1F); // PtgNum
                bytes.extend_from_slice(&val.to_le_bytes());
            },
            Ptg::PtgStr(s) => {
                bytes.push(0x17); // PtgStr
                let s_bytes = s.as_bytes();
                let len = s_bytes.len().min(255) as u8;
                bytes.push(len);
                bytes.push(0); // String flags (uncompressed)
                bytes.extend_from_slice(&s_bytes[..len as usize]);
            },
            Ptg::PtgRef(row, col, row_rel, col_rel) => {
                bytes.push(0x24); // PtgRef
                bytes.extend_from_slice(&row.to_le_bytes());
                let mut col_flags = *col;
                if *col_rel {
                    col_flags |= 0x4000;
                }
                if *row_rel {
                    col_flags |= 0x8000;
                }
                bytes.extend_from_slice(&col_flags.to_le_bytes());
            },
            Ptg::PtgArea(r1, r2, c1, c2) => {
                // BIFF8 PtgArea (2D area reference)
                // Rows are stored as 0-based indices; columns are stored
                // with relative/absolute flags in the upper bits. For the
                // initial implementation we always emit absolute area
                // references, so the flag bits remain clear.
                bytes.push(0x25); // PtgArea
                bytes.extend_from_slice(&r1.to_le_bytes());
                bytes.extend_from_slice(&r2.to_le_bytes());
                bytes.extend_from_slice(&c1.to_le_bytes());
                bytes.extend_from_slice(&c2.to_le_bytes());
            },
            Ptg::PtgArea3d(ixti, r1, r2, c1, c2) => {
                // BIFF8 PtgArea3d (3D area reference)
                //
                // Layout: opcode (1 byte) + ixti (2 bytes) + r1 (2) +
                // r2 (2) + c1 (2) + c2 (2).
                //
                // For now we always emit absolute references, so the
                // relative bits in the column fields remain clear.
                bytes.push(0x3B); // PtgArea3d
                bytes.extend_from_slice(&ixti.to_le_bytes());
                bytes.extend_from_slice(&r1.to_le_bytes());
                bytes.extend_from_slice(&r2.to_le_bytes());
                bytes.extend_from_slice(&c1.to_le_bytes());
                bytes.extend_from_slice(&c2.to_le_bytes());
            },
            Ptg::PtgAdd => bytes.push(0x03),
            Ptg::PtgSub => bytes.push(0x04),
            Ptg::PtgMul => bytes.push(0x05),
            Ptg::PtgDiv => bytes.push(0x06),
            Ptg::PtgPower => bytes.push(0x07),
            Ptg::PtgConcat => bytes.push(0x08),
            Ptg::PtgLT => bytes.push(0x09),
            Ptg::PtgLE => bytes.push(0x0A),
            Ptg::PtgEQ => bytes.push(0x0B),
            Ptg::PtgGE => bytes.push(0x0C),
            Ptg::PtgGT => bytes.push(0x0D),
            Ptg::PtgNE => bytes.push(0x0E),
            Ptg::PtgFunc(func_idx, argc) => {
                bytes.push(0x41); // PtgFuncVar
                bytes.push(*argc);
                bytes.extend_from_slice(&func_idx.to_le_bytes());
            },
            Ptg::PtgParen => bytes.push(0x15),
            Ptg::PtgMissArg => bytes.push(0x16),
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
        assert!(matches!(ref_a1, Ptg::PtgRef(0, 0, true, true)));

        let ref_abs = parse_cell_ref("$B$2").unwrap();
        assert!(matches!(ref_abs, Ptg::PtgRef(1, 1, false, false)));
    }

    #[test]
    fn test_parse_cell_ref_row_abs() {
        let ref_row_abs = parse_cell_ref("A$5").unwrap();
        assert!(matches!(ref_row_abs, Ptg::PtgRef(4, 0, false, true)));
    }

    #[test]
    fn test_parse_cell_ref_col_abs() {
        let ref_col_abs = parse_cell_ref("$C10").unwrap();
        assert!(matches!(ref_col_abs, Ptg::PtgRef(9, 2, true, false)));
    }

    #[test]
    fn test_parse_cell_ref_multiletter_col() {
        let ref_aa1 = parse_cell_ref("AA1").unwrap();
        assert!(matches!(ref_aa1, Ptg::PtgRef(0, 26, true, true)));

        let ref_zz5 = parse_cell_ref("ZZ5").unwrap();
        assert!(matches!(ref_zz5, Ptg::PtgRef(4, 701, true, true)));
    }

    #[test]
    fn test_parse_cell_ref_invalid() {
        assert!(parse_cell_ref("").is_err());
        assert!(parse_cell_ref("123").is_err());
        assert!(parse_cell_ref("ABC").is_err());
        assert!(parse_cell_ref("A0").is_err()); // Row 0 is invalid
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
        assert!(matches!(tokens[0], Ptg::PtgInt(123)));
        assert!(matches!(tokens[1], Ptg::PtgNum(456.78)));
        assert!(matches!(tokens[2], Ptg::PtgAdd));
    }

    #[test]
    fn test_tokenize_string() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("\"Hello World\"").unwrap();
        assert_eq!(tokens.len(), 1);
        assert!(matches!(&tokens[0], Ptg::PtgStr(s) if s == "Hello World"));
    }

    #[test]
    fn test_tokenize_subtraction() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("A1-B1").unwrap();
        assert_eq!(tokens.len(), 3);
        assert!(matches!(tokens[2], Ptg::PtgSub));
    }

    #[test]
    fn test_tokenize_multiplication() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("A1*B1").unwrap();
        assert_eq!(tokens.len(), 3);
        assert!(matches!(tokens[2], Ptg::PtgMul));
    }

    #[test]
    fn test_tokenize_division() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("A1/B1").unwrap();
        assert_eq!(tokens.len(), 3);
        assert!(matches!(tokens[2], Ptg::PtgDiv));
    }

    #[test]
    fn test_tokenize_power() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("A1^2").unwrap();
        assert_eq!(tokens.len(), 3);
        assert!(matches!(tokens[2], Ptg::PtgPower));
    }

    #[test]
    fn test_tokenize_concatenation() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("A1&B1").unwrap();
        assert_eq!(tokens.len(), 3);
        assert!(matches!(tokens[2], Ptg::PtgConcat));
    }

    #[test]
    fn test_tokenize_comparison_operators() {
        let tokenizer = FormulaTokenizer::new();

        let tokens_eq = tokenizer.tokenize("A1=B1").unwrap();
        assert!(matches!(tokens_eq[2], Ptg::PtgEQ));

        let tokens_ne = tokenizer.tokenize("A1<>B1").unwrap();
        assert!(matches!(tokens_ne[2], Ptg::PtgNE));

        let tokens_lt = tokenizer.tokenize("A1<B1").unwrap();
        assert!(matches!(tokens_lt[2], Ptg::PtgLT));

        let tokens_le = tokenizer.tokenize("A1<=B1").unwrap();
        assert!(matches!(tokens_le[2], Ptg::PtgLE));

        let tokens_gt = tokenizer.tokenize("A1>B1").unwrap();
        assert!(matches!(tokens_gt[2], Ptg::PtgGT));

        let tokens_ge = tokenizer.tokenize("A1>=B1").unwrap();
        assert!(matches!(tokens_ge[2], Ptg::PtgGE));
    }

    #[test]
    fn test_tokenize_parentheses() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("(A1+B1)*C1").unwrap();
        // Should be: A1, B1, +, C1, * (RPN)
        assert_eq!(tokens.len(), 5);
        assert!(matches!(tokens[2], Ptg::PtgAdd));
        assert!(matches!(tokens[4], Ptg::PtgMul));
    }

    #[test]
    fn test_tokenize_function() {
        let tokenizer = FormulaTokenizer::new();
        let tokens = tokenizer.tokenize("SUM(A1)").unwrap();
        assert!(tokens.iter().any(|t| matches!(t, Ptg::PtgFunc(4, 1))));
    }

    #[test]
    fn test_tokenize_precedence() {
        let tokenizer = FormulaTokenizer::new();
        // Multiplication has higher precedence than addition
        let tokens = tokenizer.tokenize("A1+B1*C1").unwrap();
        // Should be: A1, B1, C1, *, + (RPN)
        assert_eq!(tokens.len(), 5);
        assert!(matches!(tokens[3], Ptg::PtgMul));
        assert!(matches!(tokens[4], Ptg::PtgAdd));
    }

    #[test]
    fn test_ptg_enum_variants() {
        // Test that all Ptg variants can be created
        let _ = Ptg::PtgInt(42);
        let _ = Ptg::PtgNum(3.14);
        let _ = Ptg::PtgStr("test".to_string());
        let _ = Ptg::PtgRef(0, 0, true, true);
        let _ = Ptg::PtgArea(0, 0, 1, 1);
        let _ = Ptg::PtgArea3d(0, 0, 1, 0, 1);
        let _ = Ptg::PtgAdd;
        let _ = Ptg::PtgSub;
        let _ = Ptg::PtgMul;
        let _ = Ptg::PtgDiv;
        let _ = Ptg::PtgPower;
        let _ = Ptg::PtgConcat;
        let _ = Ptg::PtgLT;
        let _ = Ptg::PtgLE;
        let _ = Ptg::PtgEQ;
        let _ = Ptg::PtgGE;
        let _ = Ptg::PtgGT;
        let _ = Ptg::PtgNE;
        let _ = Ptg::PtgFunc(0, 1);
        let _ = Ptg::PtgParen;
        let _ = Ptg::PtgMissArg;
    }

    #[test]
    fn test_ptg_clone() {
        let ptg = Ptg::PtgRef(1, 2, true, false);
        let cloned = ptg.clone();
        assert_eq!(ptg, cloned);
    }

    #[test]
    fn test_ptg_partial_eq() {
        assert_eq!(Ptg::PtgAdd, Ptg::PtgAdd);
        assert_eq!(Ptg::PtgInt(42), Ptg::PtgInt(42));
        assert_ne!(Ptg::PtgInt(42), Ptg::PtgInt(43));
    }

    #[test]
    fn test_ptg_debug() {
        let ptg = Ptg::PtgRef(0, 0, true, true);
        let debug_str = format!("{:?}", ptg);
        assert!(debug_str.contains("PtgRef"));
    }

    #[test]
    fn test_tokenizer_default() {
        let tokenizer: FormulaTokenizer = Default::default();
        let tokens = tokenizer.tokenize("A1").unwrap();
        assert_eq!(tokens.len(), 1);
    }

    #[test]
    fn test_encode_ptg_tokens() {
        let tokens = vec![Ptg::PtgInt(42), Ptg::PtgAdd, Ptg::PtgNum(3.14)];
        let bytes = encode_ptg_tokens(&tokens);
        assert!(!bytes.is_empty());
        assert_eq!(bytes[0], 0x1E); // PtgInt opcode
    }

    #[test]
    fn test_encode_ptg_ref() {
        let tokens = vec![Ptg::PtgRef(5, 3, true, false)];
        let bytes = encode_ptg_tokens(&tokens);
        assert_eq!(bytes[0], 0x24); // PtgRef opcode
        assert_eq!(bytes.len(), 5); // 1 (opcode) + 2 (row) + 2 (col with flags)
    }

    #[test]
    fn test_encode_ptg_str() {
        let tokens = vec![Ptg::PtgStr("Test".to_string())];
        let bytes = encode_ptg_tokens(&tokens);
        assert_eq!(bytes[0], 0x17); // PtgStr opcode
        assert_eq!(bytes[1], 4); // String length
    }

    #[test]
    fn test_encode_ptg_area() {
        let tokens = vec![Ptg::PtgArea(0, 5, 0, 3)];
        let bytes = encode_ptg_tokens(&tokens);
        assert_eq!(bytes[0], 0x25); // PtgArea opcode
        assert_eq!(bytes.len(), 9); // 1 + 2*4 bytes for coordinates
    }

    #[test]
    fn test_encode_ptg_area3d() {
        let tokens = vec![Ptg::PtgArea3d(0, 0, 5, 0, 3)];
        let bytes = encode_ptg_tokens(&tokens);
        assert_eq!(bytes[0], 0x3B); // PtgArea3d opcode
        assert_eq!(bytes.len(), 11); // 1 + 2 (ixti) + 2*4 bytes for coordinates
    }

    #[test]
    fn test_encode_ptg_func() {
        let tokens = vec![Ptg::PtgFunc(4, 1)]; // SUM with 1 arg
        let bytes = encode_ptg_tokens(&tokens);
        assert_eq!(bytes[0], 0x41); // PtgFuncVar opcode
        assert_eq!(bytes[1], 1); // Arg count
    }

    #[test]
    fn test_encode_all_operators() {
        let operators = vec![
            (Ptg::PtgAdd, 0x03),
            (Ptg::PtgSub, 0x04),
            (Ptg::PtgMul, 0x05),
            (Ptg::PtgDiv, 0x06),
            (Ptg::PtgPower, 0x07),
            (Ptg::PtgConcat, 0x08),
            (Ptg::PtgLT, 0x09),
            (Ptg::PtgLE, 0x0A),
            (Ptg::PtgEQ, 0x0B),
            (Ptg::PtgGE, 0x0C),
            (Ptg::PtgGT, 0x0D),
            (Ptg::PtgNE, 0x0E),
            (Ptg::PtgParen, 0x15),
            (Ptg::PtgMissArg, 0x16),
        ];

        for (ptg, expected_opcode) in operators {
            let bytes = encode_ptg_tokens(&[ptg.clone()]);
            assert_eq!(bytes[0], expected_opcode, "Opcode mismatch for {:?}", ptg);
        }
    }
}
