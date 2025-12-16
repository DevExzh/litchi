//! ODF formula parsing and representation.
//!
//! This module provides support for OpenFormula (ODF 1.2) spreadsheet formulas.
//! It handles parsing, validation, and representation of formulas in ODS files.
//!
//! # Formula Syntax
//!
//! ODF uses OpenFormula syntax (similar to Excel but with some differences):
//! - Cell references: `A1`, `$A$1` (absolute), `Sheet1.A1` (external)
//! - Functions: `SUM(A1:A10)`, `IF(A1>0, "Positive", "Negative")`
//! - Operators: `+`, `-`, `*`, `/`, `^`, `&` (concatenation)
//! - References: `.A1` (relative to current sheet), `[file.ods]Sheet.A1` (external file)
//!
//! # References
//!
//! - OpenFormula 1.2 Specification
//! - odfdo: `3rdparty/odfdo/src/odfdo/utils/formula.py`
use crate::common::{Error, Result};
use phf::{Set, phf_set};
use smallvec::SmallVec;

// ============================================================================
// FORMULA FUNCTION CATALOG
// ============================================================================

/// Standard OpenFormula functions
///
/// This is a compile-time set of valid OpenFormula function names.
/// Using phf for O(1) lookup.
static FORMULA_FUNCTIONS: Set<&'static str> = phf_set! {
    // Mathematical functions
    "ABS", "ACOS", "ACOSH", "ACOT", "ACOTH", "ASIN", "ASINH", "ATAN", "ATAN2", "ATANH",
    "CEILING", "COS", "COSH", "COT", "COTH", "DEGREES", "EXP", "FACT", "FLOOR",
    "INT", "LN", "LOG", "LOG10", "MOD", "PI", "POWER", "PRODUCT", "QUOTIENT",
    "RADIANS", "RAND", "ROUND", "ROUNDDOWN", "ROUNDUP", "SIGN", "SIN", "SINH",
    "SQRT", "SUM", "SUMIF", "SUMIFS", "SUMSQ", "TAN", "TANH", "TRUNC",

    // Statistical functions
    "AVERAGE", "AVERAGEA", "AVERAGEIF", "AVERAGEIFS", "COUNT", "COUNTA", "COUNTBLANK",
    "COUNTIF", "COUNTIFS", "MAX", "MAXA", "MEDIAN", "MIN", "MINA", "MODE",
    "PERCENTILE", "PERCENTRANK", "QUARTILE", "RANK", "STDEV", "STDEVA", "STDEVP",
    "STDEVPA", "VAR", "VARA", "VARP", "VARPA",

    // Logical functions
    "AND", "FALSE", "IF", "IFERROR", "IFNA", "NOT", "OR", "TRUE", "XOR",

    // Text functions
    "CHAR", "CODE", "CONCATENATE", "EXACT", "FIND", "FIXED", "LEFT", "LEN",
    "LOWER", "MID", "PROPER", "REPLACE", "REPT", "RIGHT", "SEARCH", "SUBSTITUTE",
    "T", "TEXT", "TRIM", "UPPER", "VALUE",

    // Date and time functions
    "DATE", "DATEVALUE", "DAY", "DAYS", "DAYS360", "HOUR", "MINUTE", "MONTH",
    "NOW", "SECOND", "TIME", "TIMEVALUE", "TODAY", "WEEKDAY", "YEAR",

    // Lookup and reference functions
    "ADDRESS", "CHOOSE", "COLUMN", "COLUMNS", "HLOOKUP", "INDEX", "INDIRECT",
    "LOOKUP", "MATCH", "OFFSET", "ROW", "ROWS", "VLOOKUP",

    // Information functions
    "CELL", "ERROR.TYPE", "INFO", "ISBLANK", "ISERR", "ISERROR", "ISEVEN",
    "ISLOGICAL", "ISNA", "ISNONTEXT", "ISNUMBER", "ISODD", "ISREF", "ISTEXT",
    "N", "NA", "TYPE",

    // Financial functions
    "DB", "DDB", "FV", "IPMT", "IRR", "MIRR", "NPER", "NPV", "PMT", "PPMT",
    "PV", "RATE", "SLN", "SYD", "VDB",
};

// ============================================================================
// FORMULA COMPONENTS
// ============================================================================

/// A cell reference in a formula
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellRef {
    /// Sheet name (None for current sheet)
    pub sheet: Option<String>,
    /// Column (e.g., "A", "AA")
    pub column: String,
    /// Row number (1-based)
    pub row: u32,
    /// Whether column is absolute ($A)
    pub column_absolute: bool,
    /// Whether row is absolute ($1)
    pub row_absolute: bool,
}

/// A cell range reference (e.g., A1:B10)
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RangeRef {
    /// Starting cell
    pub start: CellRef,
    /// Ending cell
    pub end: CellRef,
}

/// Formula token types
#[derive(Debug, Clone, PartialEq)]
pub enum Token {
    /// Cell reference (e.g., A1, $B$2)
    CellRef(CellRef),
    /// Range reference (e.g., A1:B10)
    RangeRef(RangeRef),
    /// Function call (e.g., SUM)
    Function(String),
    /// Number literal
    Number(f64),
    /// String literal
    String(String),
    /// Boolean literal
    Boolean(bool),
    /// Operator (+, -, *, /, ^, &)
    Operator(char),
    /// Left parenthesis
    LParen,
    /// Right parenthesis
    RParen,
    /// Comma (function argument separator)
    Comma,
    /// Semicolon (array row separator)
    Semicolon,
}

/// Parsed formula structure
#[derive(Debug, Clone)]
pub struct Formula {
    /// Original formula text
    pub text: String,
    /// Parsed tokens
    pub tokens: Vec<Token>,
}

// ============================================================================
// FORMULA PARSER
// ============================================================================

/// Formula parser
pub struct FormulaParser<'a> {
    input: &'a [u8],
    position: usize,
}

impl<'a> FormulaParser<'a> {
    /// Create a new formula parser
    pub fn new(input: &'a str) -> Self {
        Self {
            input: input.as_bytes(),
            position: 0,
        }
    }

    /// Parse the formula
    pub fn parse(mut self) -> Result<Formula> {
        let original = String::from_utf8_lossy(self.input).to_string();
        let mut tokens = Vec::new();

        // Skip leading '=' if present (ODF formulas often start with =)
        if self.peek() == Some(b'=') {
            self.advance();
        }

        while !self.is_at_end() {
            self.skip_whitespace();
            if self.is_at_end() {
                break;
            }

            let token = self.next_token()?;
            tokens.push(token);
        }

        Ok(Formula {
            text: original,
            tokens,
        })
    }

    /// Parse the next token
    fn next_token(&mut self) -> Result<Token> {
        let ch = self
            .peek()
            .ok_or_else(|| Error::InvalidFormat("Unexpected end of formula".to_string()))?;

        match ch {
            b'(' => {
                self.advance();
                Ok(Token::LParen)
            },
            b')' => {
                self.advance();
                Ok(Token::RParen)
            },
            b',' => {
                self.advance();
                Ok(Token::Comma)
            },
            b';' => {
                self.advance();
                Ok(Token::Semicolon)
            },
            b'+' | b'-' | b'*' | b'/' | b'^' | b'&' | b'=' | b'<' | b'>' => {
                self.advance();
                Ok(Token::Operator(ch as char))
            },
            b'"' => self.parse_string(),
            b'0'..=b'9' => self.parse_number(),
            b'.' | b'$' | b'A'..=b'Z' | b'a'..=b'z' | b'[' => {
                // Could be cell reference, range, or function
                self.parse_identifier_or_ref()
            },
            _ => Err(Error::InvalidFormat(format!(
                "Unexpected character in formula: {}",
                ch as char
            ))),
        }
    }

    /// Parse a string literal
    fn parse_string(&mut self) -> Result<Token> {
        self.advance(); // Skip opening quote
        let mut result = String::new();

        while let Some(ch) = self.peek() {
            if ch == b'"' {
                self.advance();
                // Check for escaped quote
                if self.peek() == Some(b'"') {
                    result.push('"');
                    self.advance();
                } else {
                    break;
                }
            } else {
                result.push(ch as char);
                self.advance();
            }
        }

        Ok(Token::String(result))
    }

    /// Parse a number literal
    fn parse_number(&mut self) -> Result<Token> {
        let start = self.position;

        // Integer part
        while let Some(ch) = self.peek() {
            if ch.is_ascii_digit() {
                self.advance();
            } else {
                break;
            }
        }

        // Decimal part
        if self.peek() == Some(b'.') {
            self.advance();
            while let Some(ch) = self.peek() {
                if ch.is_ascii_digit() {
                    self.advance();
                } else {
                    break;
                }
            }
        }

        // Scientific notation
        if let Some(ch) = self.peek()
            && (ch == b'e' || ch == b'E')
        {
            self.advance();
            if let Some(sign) = self.peek()
                && (sign == b'+' || sign == b'-')
            {
                self.advance();
            }
            while let Some(ch) = self.peek() {
                if ch.is_ascii_digit() {
                    self.advance();
                } else {
                    break;
                }
            }
        }

        let num_str = std::str::from_utf8(&self.input[start..self.position])
            .map_err(|_| Error::InvalidFormat("Invalid UTF-8 in number".to_string()))?;

        let num = fast_float2::parse(num_str)
            .map_err(|_| Error::InvalidFormat(format!("Invalid number: {}", num_str)))?;

        Ok(Token::Number(num))
    }

    /// Parse identifier, cell reference, or function
    fn parse_identifier_or_ref(&mut self) -> Result<Token> {
        // Try to parse as cell reference first.
        // IMPORTANT: This parse is speculative; if it fails, we must rewind so that
        // the same input can be parsed as a function/name instead.
        if self.peek() == Some(b'.') || self.peek() == Some(b'$') || self.peek_is_letter() {
            let start_pos = self.position;
            if let Ok(cell_ref) = self.try_parse_cell_ref() {
                // Check if it's a range
                self.skip_whitespace();
                if self.peek() == Some(b':') {
                    self.advance();
                    let end = self.try_parse_cell_ref()?;
                    return Ok(Token::RangeRef(RangeRef {
                        start: cell_ref,
                        end,
                    }));
                }
                return Ok(Token::CellRef(cell_ref));
            }

            // Rewind: not a valid cell ref (e.g., "SUM(" should be a function)
            self.position = start_pos;
        }

        // Try to parse as function or named range
        let start = self.position;
        while let Some(ch) = self.peek() {
            if ch.is_ascii_alphanumeric() || ch == b'_' || ch == b'.' {
                self.advance();
            } else {
                break;
            }
        }

        let ident = std::str::from_utf8(&self.input[start..self.position])
            .map_err(|_| Error::InvalidFormat("Invalid UTF-8 in identifier".to_string()))?
            .to_uppercase();

        // Check if it's a known function
        if FORMULA_FUNCTIONS.contains(ident.as_str()) {
            Ok(Token::Function(ident))
        } else if ident == "TRUE" {
            Ok(Token::Boolean(true))
        } else if ident == "FALSE" {
            Ok(Token::Boolean(false))
        } else {
            // Treat as cell reference or named range
            Err(Error::InvalidFormat(format!(
                "Unknown identifier or invalid cell reference: {}",
                ident
            )))
        }
    }

    /// Try to parse a cell reference
    fn try_parse_cell_ref(&mut self) -> Result<CellRef> {
        let mut sheet = None;

        // Parse sheet name (if present)
        if self.peek() == Some(b'.') {
            self.advance();
            // Current sheet reference
        } else if self.peek_is_letter() {
            // Might have sheet name
            let start = self.position;
            while let Some(ch) = self.peek() {
                if ch == b'.' {
                    let sheet_name = std::str::from_utf8(&self.input[start..self.position])
                        .map_err(|_| Error::InvalidFormat("Invalid sheet name".to_string()))?;
                    sheet = Some(sheet_name.to_string());
                    self.advance(); // Skip dot
                    break;
                }
                if ch.is_ascii_alphanumeric() || ch == b'_' || ch == b' ' {
                    self.advance();
                } else {
                    // Not a sheet name, rewind
                    self.position = start;
                    break;
                }
            }
        }

        // Parse column (absolute or relative)
        let column_absolute = if self.peek() == Some(b'$') {
            self.advance();
            true
        } else {
            false
        };

        // Column letters
        let col_start = self.position;
        while let Some(ch) = self.peek() {
            if ch.is_ascii_uppercase() || ch.is_ascii_lowercase() {
                self.advance();
            } else {
                break;
            }
        }

        if col_start == self.position {
            return Err(Error::InvalidFormat(
                "Expected column in cell reference".to_string(),
            ));
        }

        let column = std::str::from_utf8(&self.input[col_start..self.position])
            .map_err(|_| Error::InvalidFormat("Invalid column".to_string()))?
            .to_uppercase();

        // Parse row (absolute or relative)
        let row_absolute = if self.peek() == Some(b'$') {
            self.advance();
            true
        } else {
            false
        };

        // Row number
        let row_start = self.position;
        while let Some(ch) = self.peek() {
            if ch.is_ascii_digit() {
                self.advance();
            } else {
                break;
            }
        }

        if row_start == self.position {
            return Err(Error::InvalidFormat(
                "Expected row in cell reference".to_string(),
            ));
        }

        let row_str = std::str::from_utf8(&self.input[row_start..self.position])
            .map_err(|_| Error::InvalidFormat("Invalid row".to_string()))?;

        let row = row_str
            .parse::<u32>()
            .map_err(|_| Error::InvalidFormat("Invalid row number".to_string()))?;

        Ok(CellRef {
            sheet,
            column,
            row,
            column_absolute,
            row_absolute,
        })
    }

    /// Peek at current character
    fn peek(&self) -> Option<u8> {
        self.input.get(self.position).copied()
    }

    /// Check if current character is a letter
    fn peek_is_letter(&self) -> bool {
        self.peek().is_some_and(|ch| ch.is_ascii_alphabetic())
    }

    /// Advance position
    fn advance(&mut self) {
        self.position += 1;
    }

    /// Check if at end
    fn is_at_end(&self) -> bool {
        self.position >= self.input.len()
    }

    /// Skip whitespace
    fn skip_whitespace(&mut self) {
        while let Some(ch) = self.peek() {
            if ch.is_ascii_whitespace() {
                self.advance();
            } else {
                break;
            }
        }
    }
}

// ============================================================================
// FORMULA UTILITIES
// ============================================================================

/// Check if a string is a valid OpenFormula function name
#[inline]
#[allow(dead_code)] // Will be used for future enhancements
pub fn is_valid_function(name: &str) -> bool {
    FORMULA_FUNCTIONS.contains(name.to_uppercase().as_str())
}

/// Extract all cell references from a formula
pub fn extract_cell_refs(formula: &Formula) -> SmallVec<[&CellRef; 8]> {
    formula
        .tokens
        .iter()
        .filter_map(|token| match token {
            Token::CellRef(cell_ref) => Some(cell_ref),
            Token::RangeRef(range_ref) => Some(&range_ref.start), // Just start for simplicity
            _ => None,
        })
        .collect()
}

/// Extract all function calls from a formula
pub fn extract_functions(formula: &Formula) -> SmallVec<[&str; 4]> {
    formula
        .tokens
        .iter()
        .filter_map(|token| match token {
            Token::Function(name) => Some(name.as_str()),
            _ => None,
        })
        .collect()
}

// ============================================================================
// TESTS
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_simple_formula() {
        let parser = FormulaParser::new("=A1+B2");
        let formula = parser.parse().unwrap();
        assert_eq!(formula.tokens.len(), 3);
    }

    #[test]
    fn test_parse_function_formula() {
        let parser = FormulaParser::new("=SUM(A1:A10)");
        let formula = parser.parse().unwrap();
        assert!(matches!(formula.tokens[0], Token::Function(_)));
    }

    #[test]
    fn test_parse_absolute_reference() {
        let parser = FormulaParser::new("=$A$1");
        let formula = parser.parse().unwrap();
        match &formula.tokens[0] {
            Token::CellRef(cell_ref) => {
                assert!(cell_ref.column_absolute);
                assert!(cell_ref.row_absolute);
            },
            _ => panic!("Expected cell reference"),
        }
    }

    #[test]
    fn test_is_valid_function() {
        assert!(is_valid_function("SUM"));
        assert!(is_valid_function("AVERAGE"));
        assert!(!is_valid_function("INVALID_FUNCTION"));
    }

    #[test]
    fn test_extract_cell_refs() {
        let parser = FormulaParser::new("=A1+B2+C3");
        let formula = parser.parse().unwrap();
        let refs = extract_cell_refs(&formula);
        assert!(refs.len() >= 2); // At least A1 and B2
    }
}
