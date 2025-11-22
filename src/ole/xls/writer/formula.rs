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

/// Parse a cell reference like "A1" or "$B$2"
fn parse_cell_ref(s: &str) -> Result<Ptg, XlsError> {
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
            _ => {}, // Other tokens not yet implemented
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
}
