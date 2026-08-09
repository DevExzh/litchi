//! ODF formula parsing and representation.
//!
//! This module provides support for `OpenFormula` (ODF 1.2) spreadsheet formulas.
//! It handles parsing, validation, and representation of formulas in ODS files.
//!
//! # Formula Syntax
//!
//! ODF uses `OpenFormula` syntax (similar to Excel but with some differences):
//! - Cell references: `A1`, `$A$1` (absolute), `Sheet1.A1` (external)
//! - Functions: `SUM(A1:A10)`, `IF(A1>0, "Positive", "Negative")`
//! - Operators: `+`, `-`, `*`, `/`, `^`, `&` (concatenation)
//! - References: `.A1` (relative to current sheet), `[file.ods]Sheet.A1` (external file)
//!
//! # References
//!
//! - `OpenFormula` 1.2 Specification
//! - odfdo: `3rdparty/odfdo/src/odfdo/utils/formula.py`
use litchi_core::{Error, Result};
use phf::{Set, phf_set};
use smallvec::SmallVec;

// ============================================================================
// FORMULA FUNCTION CATALOG
// ============================================================================

/// Standard `OpenFormula` functions
///
/// This is a compile-time set of valid `OpenFormula` function names.
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
    #[must_use]
    pub fn new(input: &'a str) -> Self {
        Self {
            input: input.as_bytes(),
            position: 0,
        }
    }

    /// Parse the formula
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn parse(mut self) -> Result<Formula> {
        let original = String::from_utf8_lossy(self.input).to_string();
        let mut tokens = Vec::new();

        // ODF stores formulas as `of:=...`, while the public codec also accepts
        // the shorter `=...` spelling.  Keep the original text in `Formula`,
        // but parse the body directly so stripping the prefix does not require
        // a normalized temporary string.
        let input = std::str::from_utf8(self.input)
            .map_err(|_error| Error::InvalidFormat("Invalid UTF-8 in formula".to_string()))?;
        let body = input.trim();
        let body = body
            .strip_prefix('=')
            .or_else(|| strip_open_formula_prefix(body))
            .unwrap_or(body);
        self.input = body.as_bytes();

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
            b'[' => self.parse_bracket_reference(),
            b'.' | b'$' | b'A'..=b'Z' | b'a'..=b'z' => {
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
            .map_err(|_error| Error::InvalidFormat("Invalid UTF-8 in number".to_string()))?;

        let num = fast_float2::parse(num_str)
            .map_err(|_error| Error::InvalidFormat(format!("Invalid number: {num_str}")))?;

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
            .map_err(|_error| Error::InvalidFormat("Invalid UTF-8 in identifier".to_string()))?
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
                "Unknown identifier or invalid cell reference: {ident}"
            )))
        }
    }

    /// Parse an ODF bracketed cell or range reference, such as `[.A1]` or
    /// `[$Inputs.$A$1:.$B$2]`.
    fn parse_bracket_reference(&mut self) -> Result<Token> {
        self.advance(); // Skip the opening bracket.
        let start = self.position;
        let mut quoted = false;

        while let Some(ch) = self.peek() {
            match ch {
                b'\'' => {
                    quoted = !quoted;
                    self.advance();
                },
                b']' if !quoted => break,
                _ => self.advance(),
            }
        }

        let end = self.position;
        if self.peek() != Some(b']') {
            return Err(Error::InvalidFormat(
                "Unterminated ODF bracketed reference".to_string(),
            ));
        }
        self.advance(); // Skip the closing bracket.

        let reference = std::str::from_utf8(&self.input[start..end]).map_err(|_error| {
            Error::InvalidFormat("Invalid UTF-8 in cell reference".to_string())
        })?;
        parse_open_formula_reference(reference)
    }

    /// Try to parse a cell reference
    fn try_parse_cell_ref(&mut self) -> Result<CellRef> {
        let mut sheet = None;

        // Parse sheet name (if present)
        if self.peek() == Some(b'.') {
            self.advance();
            // Current sheet reference
        } else if self.peek_is_letter() {
            // Might have a sheet-qualified reference like Sheet1.A1.
            // If there is no dot after the identifier chunk, this is a plain
            // cell reference like A1 and we must rewind.
            let start = self.position;
            while let Some(ch) = self.peek() {
                if ch.is_ascii_alphanumeric() || ch == b'_' || ch == b' ' {
                    self.advance();
                } else {
                    break;
                }
            }

            if self.peek() == Some(b'.') {
                let sheet_name = std::str::from_utf8(&self.input[start..self.position])
                    .map_err(|_error| Error::InvalidFormat("Invalid sheet name".to_string()))?;
                sheet = Some(sheet_name.to_string());
                self.advance(); // Skip dot
            } else {
                self.position = start;
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
            .map_err(|_error| Error::InvalidFormat("Invalid column".to_string()))?
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
            .map_err(|_error| Error::InvalidFormat("Invalid row".to_string()))?;

        let row = row_str
            .parse::<u32>()
            .map_err(|_error| Error::InvalidFormat("Invalid row number".to_string()))?;

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

fn strip_open_formula_prefix(value: &str) -> Option<&str> {
    let bytes = value.as_bytes();
    (bytes.len() >= 4 && bytes[..3].eq_ignore_ascii_case(b"of:") && bytes[3] == b'=')
        .then(|| &value[4..])
}

fn parse_open_formula_reference(value: &str) -> Result<Token> {
    let value = value.trim();
    let (start, end) = split_open_formula_range(value)
        .ok_or_else(|| Error::InvalidFormat("Invalid ODF bracketed reference range".to_string()))?;
    let start = parse_open_formula_cell_ref(start)?;
    let Some(end) = end else {
        return Ok(Token::CellRef(start));
    };

    Ok(Token::RangeRef(RangeRef {
        start,
        end: parse_open_formula_cell_ref(end)?,
    }))
}

fn split_open_formula_range(value: &str) -> Option<(&str, Option<&str>)> {
    if value.is_empty() {
        return None;
    }

    let mut quoted = false;
    for (index, character) in value.char_indices() {
        match character {
            '\'' => quoted = !quoted,
            ':' if !quoted => return Some((&value[..index], Some(&value[index + 1..]))),
            _ => {},
        }
    }
    (!quoted).then_some((value, None))
}

fn parse_open_formula_cell_ref(value: &str) -> Result<CellRef> {
    let value = value.trim();
    let (sheet, cell) = if let Some(cell) = value.strip_prefix('.') {
        (None, cell)
    } else {
        let separator = value
            .char_indices()
            .rev()
            .find_map(|(index, character)| (character == '.').then_some(index))
            .ok_or_else(|| {
                Error::InvalidFormat("ODF reference is missing its sheet separator".to_string())
            })?;
        let sheet = parse_open_formula_sheet_name(&value[..separator])?;
        (Some(sheet), &value[separator + 1..])
    };

    let (column, row, column_absolute, row_absolute) = parse_a1_cell_ref(cell)?;
    Ok(CellRef {
        sheet,
        column,
        row,
        column_absolute,
        row_absolute,
    })
}

fn parse_open_formula_sheet_name(value: &str) -> Result<String> {
    let value = value.trim().strip_prefix('$').unwrap_or(value.trim());
    if value.is_empty() {
        return Err(Error::InvalidFormat(
            "ODF reference has an empty sheet name".to_string(),
        ));
    }

    if value.starts_with('\'') || value.ends_with('\'') {
        let value = value
            .strip_prefix('\'')
            .and_then(|value| value.strip_suffix('\''))
            .ok_or_else(|| {
                Error::InvalidFormat("ODF reference has an unterminated sheet name".to_string())
            })?;
        let mut sheet = String::with_capacity(value.len());
        let mut characters = value.chars().peekable();
        while let Some(character) = characters.next() {
            if character == '\'' && characters.next() != Some('\'') {
                return Err(Error::InvalidFormat(
                    "ODF reference has an invalid escaped sheet name".to_string(),
                ));
            }
            sheet.push(character);
        }
        if sheet.is_empty() {
            return Err(Error::InvalidFormat(
                "ODF reference has an empty sheet name".to_string(),
            ));
        }
        Ok(sheet)
    } else if value.contains('\'') {
        Err(Error::InvalidFormat(
            "ODF reference has an invalid sheet name".to_string(),
        ))
    } else {
        Ok(value.to_string())
    }
}

fn parse_a1_cell_ref(value: &str) -> Result<(String, u32, bool, bool)> {
    let value = value.trim();
    let bytes = value.as_bytes();
    let mut position = 0;
    let column_absolute = if bytes.get(position) == Some(&b'$') {
        position += 1;
        true
    } else {
        false
    };
    let column_start = position;
    while bytes.get(position).is_some_and(u8::is_ascii_alphabetic) {
        position += 1;
    }
    if column_start == position {
        return Err(Error::InvalidFormat(
            "ODF reference is missing its column".to_string(),
        ));
    }
    let column = std::str::from_utf8(&bytes[column_start..position])
        .map_err(|_error| Error::InvalidFormat("Invalid UTF-8 in ODF column".to_string()))?
        .to_ascii_uppercase();

    let row_absolute = if bytes.get(position) == Some(&b'$') {
        position += 1;
        true
    } else {
        false
    };
    let row_start = position;
    while bytes.get(position).is_some_and(u8::is_ascii_digit) {
        position += 1;
    }
    if row_start == position || position != bytes.len() {
        return Err(Error::InvalidFormat(
            "ODF reference has an invalid row".to_string(),
        ));
    }
    let row = std::str::from_utf8(&bytes[row_start..position])
        .map_err(|_error| Error::InvalidFormat("Invalid UTF-8 in ODF row".to_string()))?
        .parse::<u32>()
        .map_err(|_error| Error::InvalidFormat("Invalid ODF row number".to_string()))?;
    if row == 0 {
        return Err(Error::InvalidFormat(
            "ODF references must use 1-based rows".to_string(),
        ));
    }

    Ok((column, row, column_absolute, row_absolute))
}

// ============================================================================
// FORMULA UTILITIES
// ============================================================================

/// Check if a string is a valid `OpenFormula` function name
#[inline]
#[allow(dead_code)] // Will be used for future enhancements
#[must_use]
pub fn is_valid_function(name: &str) -> bool {
    FORMULA_FUNCTIONS.contains(name.to_uppercase().as_str())
}

/// Extract all cell references from a formula
#[must_use]
pub fn extract_cell_refs(formula: &Formula) -> SmallVec<[&CellRef; 8]> {
    formula
        .tokens
        .iter()
        .filter_map(|token| match token {
            Token::CellRef(cell_ref) => Some(cell_ref),
            Token::RangeRef(range_ref) => Some(&range_ref.start), // Just start for simplicity
            Token::Function(_)
            | Token::Number(_)
            | Token::String(_)
            | Token::Boolean(_)
            | Token::Operator(_)
            | Token::LParen
            | Token::RParen
            | Token::Comma
            | Token::Semicolon => None,
        })
        .collect()
}

/// Extract all function calls from a formula
#[must_use]
pub fn extract_functions(formula: &Formula) -> SmallVec<[&str; 4]> {
    formula
        .tokens
        .iter()
        .filter_map(|token| match token {
            Token::Function(name) => Some(name.as_str()),
            Token::CellRef(_)
            | Token::RangeRef(_)
            | Token::Number(_)
            | Token::String(_)
            | Token::Boolean(_)
            | Token::Operator(_)
            | Token::LParen
            | Token::RParen
            | Token::Comma
            | Token::Semicolon => None,
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
        let formula = parser
            .parse()
            .expect("test fixture or operation should succeed");
        assert_eq!(formula.tokens.len(), 3);
    }

    #[test]
    fn test_parse_function_formula() {
        let parser = FormulaParser::new("=SUM(A1:A10)");
        let formula = parser
            .parse()
            .expect("test fixture or operation should succeed");
        assert!(matches!(formula.tokens[0], Token::Function(_)));
    }

    #[test]
    fn test_parse_canonical_odf_formula_without_normalizing_the_input() {
        let formula = FormulaParser::new("of:=SUM([$Inputs.$A$1:.$B$2])")
            .parse()
            .expect("test fixture or operation should succeed");

        assert_eq!(formula.text, "of:=SUM([$Inputs.$A$1:.$B$2])");
        assert!(matches!(&formula.tokens[0], Token::Function(name) if name == "SUM"));
        assert!(matches!(
            &formula.tokens[2],
            Token::RangeRef(RangeRef { start, end })
                if start.sheet.as_deref() == Some("Inputs")
                    && start.column == "A"
                    && start.row == 1
                    && start.column_absolute
                    && start.row_absolute
                    && end.sheet.is_none()
                    && end.column == "B"
                    && end.row == 2
                    && end.column_absolute
                    && end.row_absolute
        ));
    }

    #[test]
    fn test_parse_odf_formula_with_quoted_sheet_reference() {
        let formula = FormulaParser::new("OF:=['Bob''s'.$A$1]")
            .parse()
            .expect("test fixture or operation should succeed");

        assert!(matches!(
            &formula.tokens[0],
            Token::CellRef(CellRef { sheet: Some(sheet), column, row, .. })
                if sheet == "Bob's" && column == "A" && *row == 1
        ));
    }

    #[test]
    fn test_parse_absolute_reference() {
        let parser = FormulaParser::new("=$A$1");
        let formula = parser
            .parse()
            .expect("test fixture or operation should succeed");
        match &formula.tokens[0] {
            Token::CellRef(cell_ref) => {
                assert!(cell_ref.column_absolute);
                assert!(cell_ref.row_absolute);
            },
            Token::RangeRef(_)
            | Token::Function(_)
            | Token::Number(_)
            | Token::String(_)
            | Token::Boolean(_)
            | Token::Operator(_)
            | Token::LParen
            | Token::RParen
            | Token::Comma
            | Token::Semicolon => panic!("Expected cell reference"),
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
        let formula = parser
            .parse()
            .expect("test fixture or operation should succeed");
        let refs = extract_cell_refs(&formula);
        assert!(refs.len() >= 2); // At least A1 and B2
    }

    #[test]
    fn test_parse_formula_without_equals() {
        let parser = FormulaParser::new("A1+B1");
        let formula = parser
            .parse()
            .expect("test fixture or operation should succeed");
        assert!(!formula.tokens.is_empty());
    }

    #[test]
    fn test_cell_ref_parsing() {
        let parser = FormulaParser::new("=Sheet1.A1");
        let formula = parser
            .parse()
            .expect("test fixture or operation should succeed");
        match &formula.tokens[0] {
            Token::CellRef(cell_ref) => {
                assert_eq!(cell_ref.sheet, Some("Sheet1".to_string()));
                assert_eq!(cell_ref.column, "A");
                assert_eq!(cell_ref.row, 1);
            },
            Token::RangeRef(_)
            | Token::Function(_)
            | Token::Number(_)
            | Token::String(_)
            | Token::Boolean(_)
            | Token::Operator(_)
            | Token::LParen
            | Token::RParen
            | Token::Comma
            | Token::Semicolon => panic!("Expected cell reference"),
        }
    }

    #[test]
    fn test_range_ref_parsing() {
        let parser = FormulaParser::new("=A1:B10");
        let formula = parser
            .parse()
            .expect("test fixture or operation should succeed");
        match &formula.tokens[0] {
            Token::RangeRef(range_ref) => {
                assert_eq!(range_ref.start.column, "A");
                assert_eq!(range_ref.start.row, 1);
                assert_eq!(range_ref.end.column, "B");
                assert_eq!(range_ref.end.row, 10);
            },
            Token::CellRef(_)
            | Token::Function(_)
            | Token::Number(_)
            | Token::String(_)
            | Token::Boolean(_)
            | Token::Operator(_)
            | Token::LParen
            | Token::RParen
            | Token::Comma
            | Token::Semicolon => panic!("Expected range reference"),
        }
    }

    #[test]
    fn test_number_token() {
        let parser = FormulaParser::new("=42.5");
        let formula = parser
            .parse()
            .expect("test fixture or operation should succeed");
        match &formula.tokens[0] {
            Token::Number(n) => {
                assert!((n - 42.5).abs() < 0.0001);
            },
            Token::CellRef(_)
            | Token::RangeRef(_)
            | Token::Function(_)
            | Token::String(_)
            | Token::Boolean(_)
            | Token::Operator(_)
            | Token::LParen
            | Token::RParen
            | Token::Comma
            | Token::Semicolon => panic!("Expected number token"),
        }
    }

    #[test]
    fn test_string_token() {
        let parser = FormulaParser::new("=\"Hello World\"");
        let formula = parser
            .parse()
            .expect("test fixture or operation should succeed");
        match &formula.tokens[0] {
            Token::String(s) => {
                assert_eq!(s, "Hello World");
            },
            Token::CellRef(_)
            | Token::RangeRef(_)
            | Token::Function(_)
            | Token::Number(_)
            | Token::Boolean(_)
            | Token::Operator(_)
            | Token::LParen
            | Token::RParen
            | Token::Comma
            | Token::Semicolon => panic!("Expected string token: {:?}", formula.tokens),
        }
    }

    #[test]
    fn test_boolean_tokens() {
        let parser = FormulaParser::new("=TRUE()");
        let formula = parser
            .parse()
            .expect("test fixture or operation should succeed");
        assert!(matches!(&formula.tokens[0], Token::Function(f) if f == "TRUE"));

        let parser = FormulaParser::new("=FALSE()");
        let formula = parser
            .parse()
            .expect("test fixture or operation should succeed");
        assert!(matches!(&formula.tokens[0], Token::Function(f) if f == "FALSE"));
    }

    #[test]
    fn test_operators() {
        let parser = FormulaParser::new("=A1+B1-C1*D1/E1^F1");
        let formula = parser
            .parse()
            .expect("test fixture or operation should succeed");
        // Should have cell refs and operators
        let operators: Vec<_> = formula
            .tokens
            .iter()
            .filter(|t| matches!(t, Token::Operator(_)))
            .collect();
        assert!(!operators.is_empty());
    }

    #[test]
    fn test_parentheses() {
        let parser = FormulaParser::new("=(A1+B1)*C1");
        let formula = parser
            .parse()
            .expect("test fixture or operation should succeed");
        let has_lparen = formula.tokens.iter().any(|t| matches!(t, Token::LParen));
        let has_rparen = formula.tokens.iter().any(|t| matches!(t, Token::RParen));
        assert!(has_lparen);
        assert!(has_rparen);
    }

    #[test]
    fn test_function_with_multiple_args() {
        let parser = FormulaParser::new("=IF(A1>0,\"Positive\",\"Negative\")");
        let formula = parser
            .parse()
            .expect("test fixture or operation should succeed");
        assert!(matches!(&formula.tokens[0], Token::Function(f) if f == "IF"));

        // Check for commas
        let commas = formula
            .tokens
            .iter()
            .filter(|t| matches!(t, Token::Comma))
            .count();
        assert_eq!(commas, 2);
    }

    #[test]
    fn test_nested_functions() {
        let parser = FormulaParser::new("=SUM(AVERAGE(A1:A10),MAX(B1:B10))");
        let formula = parser
            .parse()
            .expect("test fixture or operation should succeed");
        let functions: Vec<_> = formula
            .tokens
            .iter()
            .filter_map(|t| match t {
                Token::Function(f) => Some(f.as_str()),
                Token::CellRef(_)
                | Token::RangeRef(_)
                | Token::Number(_)
                | Token::String(_)
                | Token::Boolean(_)
                | Token::Operator(_)
                | Token::LParen
                | Token::RParen
                | Token::Comma
                | Token::Semicolon => None,
            })
            .collect();
        assert!(functions.contains(&"SUM"));
        assert!(functions.contains(&"AVERAGE"));
        assert!(functions.contains(&"MAX"));
    }

    #[test]
    fn test_mixed_references() {
        // Mixed absolute/relative references
        let parser = FormulaParser::new("=$A1+B$1");
        let formula = parser
            .parse()
            .expect("test fixture or operation should succeed");
        match &formula.tokens[0] {
            Token::CellRef(cell_ref) => {
                assert!(cell_ref.column_absolute);
                assert!(!cell_ref.row_absolute);
            },
            Token::RangeRef(_)
            | Token::Function(_)
            | Token::Number(_)
            | Token::String(_)
            | Token::Boolean(_)
            | Token::Operator(_)
            | Token::LParen
            | Token::RParen
            | Token::Comma
            | Token::Semicolon => panic!("Expected cell reference"),
        }
    }

    #[test]
    fn test_formula_struct() {
        let formula = Formula {
            text: "=A1+B1".to_string(),
            tokens: vec![
                Token::CellRef(CellRef {
                    sheet: None,
                    column: "A".to_string(),
                    row: 1,
                    column_absolute: false,
                    row_absolute: false,
                }),
                Token::Operator('+'),
                Token::CellRef(CellRef {
                    sheet: None,
                    column: "B".to_string(),
                    row: 1,
                    column_absolute: false,
                    row_absolute: false,
                }),
            ],
        };
        assert_eq!(formula.text, "=A1+B1");
        assert_eq!(formula.tokens.len(), 3);
    }

    #[test]
    fn test_cell_ref_equality() {
        let ref1 = CellRef {
            sheet: None,
            column: "A".to_string(),
            row: 1,
            column_absolute: false,
            row_absolute: false,
        };
        let ref2 = CellRef {
            sheet: None,
            column: "A".to_string(),
            row: 1,
            column_absolute: false,
            row_absolute: false,
        };
        let ref3 = CellRef {
            sheet: Some("Sheet1".to_string()),
            column: "A".to_string(),
            row: 1,
            column_absolute: false,
            row_absolute: false,
        };
        assert_eq!(ref1, ref2);
        assert_ne!(ref1, ref3);
    }

    #[test]
    fn test_range_ref_equality() {
        let range1 = RangeRef {
            start: CellRef {
                sheet: None,
                column: "A".to_string(),
                row: 1,
                column_absolute: false,
                row_absolute: false,
            },
            end: CellRef {
                sheet: None,
                column: "B".to_string(),
                row: 10,
                column_absolute: false,
                row_absolute: false,
            },
        };
        let range2 = RangeRef {
            start: CellRef {
                sheet: None,
                column: "A".to_string(),
                row: 1,
                column_absolute: false,
                row_absolute: false,
            },
            end: CellRef {
                sheet: None,
                column: "B".to_string(),
                row: 10,
                column_absolute: false,
                row_absolute: false,
            },
        };
        assert_eq!(range1, range2);
    }

    #[test]
    fn test_token_variants() {
        let cell_ref = CellRef {
            sheet: None,
            column: "A".to_string(),
            row: 1,
            column_absolute: false,
            row_absolute: false,
        };
        let token1 = Token::CellRef(cell_ref.clone());
        let token2 = Token::CellRef(cell_ref.clone());
        assert_eq!(token1, token2);

        assert_eq!(Token::Operator('+'), Token::Operator('+'));
        assert_eq!(Token::LParen, Token::LParen);
        assert_eq!(Token::RParen, Token::RParen);
        assert_eq!(Token::Comma, Token::Comma);
        assert_eq!(Token::Semicolon, Token::Semicolon);
    }

    #[test]
    fn test_extract_functions() {
        let parser = FormulaParser::new("=SUM(A1:A10)+AVERAGE(B1:B10)");
        let formula = parser
            .parse()
            .expect("test fixture or operation should succeed");
        let funcs = extract_functions(&formula);
        assert!(funcs.contains(&"SUM"));
        assert!(funcs.contains(&"AVERAGE"));
    }

    #[test]
    fn test_formula_functions_catalog() {
        // Test that common functions are valid
        assert!(is_valid_function("SUM"));
        assert!(is_valid_function("AVERAGE"));
        assert!(is_valid_function("IF"));
        assert!(is_valid_function("VLOOKUP"));
        assert!(is_valid_function("COUNT"));
        assert!(is_valid_function("MAX"));
        assert!(is_valid_function("MIN"));
        assert!(is_valid_function("ABS"));
        assert!(is_valid_function("ROUND"));
        assert!(is_valid_function("TODAY"));
        assert!(is_valid_function("NOW"));

        // Invalid functions
        assert!(!is_valid_function("NOTAFUNCTION"));
        assert!(!is_valid_function(""));
    }

    #[test]
    fn test_whitespace_handling() {
        let parser = FormulaParser::new("=  A1  +  B1  ");
        let formula = parser
            .parse()
            .expect("test fixture or operation should succeed");
        assert!(!formula.tokens.is_empty());
    }

    #[test]
    fn test_complex_formula() {
        let parser = FormulaParser::new("=IF(SUM(A1:A10)>100,AVERAGE(B1:B10),0)");
        let formula = parser
            .parse()
            .expect("test fixture or operation should succeed");
        assert!(!formula.tokens.is_empty());
        // Check all expected tokens are present
        let funcs: Vec<_> = formula
            .tokens
            .iter()
            .filter_map(|t| match t {
                Token::Function(f) => Some(f.as_str()),
                Token::CellRef(_)
                | Token::RangeRef(_)
                | Token::Number(_)
                | Token::String(_)
                | Token::Boolean(_)
                | Token::Operator(_)
                | Token::LParen
                | Token::RParen
                | Token::Comma
                | Token::Semicolon => None,
            })
            .collect();
        assert!(funcs.contains(&"IF"));
        assert!(funcs.contains(&"SUM"));
        assert!(funcs.contains(&"AVERAGE"));
    }
}
