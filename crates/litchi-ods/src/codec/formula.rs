//! ODF formula parsing and representation.
//!
//! This module tokenizes the supported subset of `OpenFormula` expressions and
//! represents formulas stored in ODS files. Function-name recognition covers
//! the complete normative Part 4, chapter 6 catalog; this remains a tokenizer,
//! not a complete expression grammar or evaluator, and it does not validate
//! function arity.
//!
//! # Formula Syntax
//!
//! ODF uses `OpenFormula` syntax (similar to Excel but with some differences):
//! - Cell references: `A1`, `$A$1` (absolute), `Sheet1.A1` (sheet-qualified)
//! - Functions: `SUM(A1:A10)`, `IF(A1>0, "Positive", "Negative")`
//! - Operators: `+`, `-`, `*`, `/`, `^`, `&` (concatenation)
//! - References: `.A1` (relative to current sheet), `[$Inputs.$A$1]` (bracketed)
//!
//! # References
//!
//! - ODF 1.4 Part 4, `OpenFormula` Format
//! - odfdo: `3rdparty/odfdo/src/odfdo/utils/formula.py`
use litchi_core::{Error, Resource, ResourceLimit, Result};
use smallvec::SmallVec;
use std::{borrow::Cow, convert::TryFrom, sync::Arc};

mod functions;
pub mod reference;

use functions::{MAX_STANDARD_FUNCTION_NAME_BYTES, STANDARD_FORMULA_FUNCTIONS};

// ============================================================================
// FORMULA COMPONENTS
// ============================================================================

/// Default maximum UTF-8 byte length of one parsed formula.
pub const DEFAULT_MAX_FORMULA_BYTES: usize = 1024 * 1024;
/// Default maximum number of tokens emitted for one parsed formula.
pub const DEFAULT_MAX_FORMULA_TOKENS: usize = 65_536;

/// Finite parser limits for formulas and their token vectors.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct FormulaLimits {
    max_bytes: usize,
    max_tokens: usize,
    reference: reference::Limits,
}

impl Default for FormulaLimits {
    fn default() -> Self {
        Self {
            max_bytes: DEFAULT_MAX_FORMULA_BYTES,
            max_tokens: DEFAULT_MAX_FORMULA_TOKENS,
            reference: reference::Limits::default(),
        }
    }
}

impl FormulaLimits {
    /// Set the maximum UTF-8 byte length of a formula.
    #[must_use]
    pub const fn with_max_bytes(mut self, value: usize) -> Self {
        self.max_bytes = value;
        self
    }

    /// Set the maximum number of emitted tokens.
    #[must_use]
    pub const fn with_max_tokens(mut self, value: usize) -> Self {
        self.max_tokens = value;
        self
    }

    /// Set the finite limits used for every bracketed reference in this
    /// formula.
    #[must_use]
    pub const fn with_reference_limits(mut self, value: reference::Limits) -> Self {
        self.reference = value;
        self
    }

    /// Return the maximum formula byte length.
    #[must_use]
    pub const fn max_bytes(self) -> usize {
        self.max_bytes
    }

    /// Return the maximum token count.
    #[must_use]
    pub const fn max_tokens(self) -> usize {
        self.max_tokens
    }

    /// Return the finite limits used for bracketed references.
    #[must_use]
    pub const fn reference_limits(self) -> reference::Limits {
        self.reference
    }
}

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
    /// Complete OpenFormula 1.4 bracketed reference metadata.
    ///
    /// The box keeps adding source, subtable, and whole-axis state from
    /// inflating every token in an otherwise ordinary formula.
    Reference(Box<reference::Reference>),
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
    limits: FormulaLimits,
}

impl<'a> FormulaParser<'a> {
    /// Create a new formula parser
    #[must_use]
    pub fn new(input: &'a str) -> Self {
        Self {
            input: input.as_bytes(),
            position: 0,
            limits: FormulaLimits::default(),
        }
    }

    /// Parse the formula
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn parse(self) -> Result<Formula> {
        self.parse_with_limits(&FormulaLimits::default())
    }

    /// Parse the formula with explicit finite byte and token limits.
    pub fn parse_with_limits(mut self, limits: &FormulaLimits) -> Result<Formula> {
        // ODF stores formulas as `of:=...`, while the public codec also accepts
        // the shorter `=...` spelling.  Keep the original text in `Formula`,
        // but parse the body directly so stripping the prefix does not require
        // a normalized temporary string.
        let input = std::str::from_utf8(self.input)
            .map_err(|_error| Error::InvalidFormat("Invalid UTF-8 in formula".to_string()))?;
        if input.len() > limits.max_bytes {
            return Err(formula_limit_error(
                Resource::InputBytes,
                input.len(),
                limits.max_bytes,
            ));
        }
        self.limits = *limits;
        let mut original = String::new();
        original
            .try_reserve_exact(input.len())
            .map_err(|source| Error::Allocation {
                resource: "formula text",
                source,
            })?;
        original.push_str(input);
        let body = input.trim();
        let body = body
            .strip_prefix('=')
            .or_else(|| strip_open_formula_prefix(body))
            .unwrap_or(body);
        self.input = body.as_bytes();
        let mut tokens = Vec::new();

        while !self.is_at_end() {
            self.skip_whitespace();
            if self.is_at_end() {
                break;
            }

            if tokens.len() >= self.limits.max_tokens {
                return Err(formula_limit_error(
                    Resource::Objects,
                    tokens.len().saturating_add(1),
                    self.limits.max_tokens,
                ));
            }
            let token = self.next_token()?;
            if tokens.len() == tokens.capacity() {
                reserve_token_slot(&mut tokens, self.limits.max_tokens)?;
            }
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
        result
            .try_reserve_exact(self.input.len().saturating_sub(self.position))
            .map_err(|source| Error::Allocation {
                resource: "formula string literal",
                source,
            })?;
        let mut segment_start = self.position;

        while let Some(ch) = self.peek() {
            if ch == b'"' {
                // Formula input has already been validated as UTF-8. Copy
                // complete UTF-8 segments instead of treating each byte as a
                // character; this preserves non-ASCII literals verbatim.
                if segment_start < self.position {
                    let segment = std::str::from_utf8(&self.input[segment_start..self.position])
                        .map_err(|_error| {
                            Error::InvalidFormat("Invalid UTF-8 in string literal".to_string())
                        })?;
                    result.push_str(segment);
                }
                self.advance();
                // Check for escaped quote
                if self.peek() == Some(b'"') {
                    result.push('"');
                    self.advance();
                    segment_start = self.position;
                } else {
                    return Ok(Token::String(result));
                }
            } else {
                self.advance();
            }
        }

        Err(Error::InvalidFormat(
            "Unterminated string literal".to_string(),
        ))
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
        // A function name may also be a valid column-plus-row spelling (for
        // example, `BIN2DEC` is column BIN, row 2, followed by `DEC`).  When
        // the complete identifier is followed by an opening parenthesis, the
        // function grammar takes precedence over that speculative cell parse.
        // Keep the whitespace unconsumed so the normal token loop handles it.
        if let Some(function) = self.try_parse_function_call()? {
            return Ok(function);
        }

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
            .trim();

        // Check if it's a known function
        if let Some(function) = lookup_function(ident) {
            Ok(Token::Function(copy_formula_component(
                function,
                "formula function name",
            )?))
        } else if ident == "TRUE" {
            Ok(Token::Boolean(true))
        } else if ident == "FALSE" {
            Ok(Token::Boolean(false))
        } else {
            // Treat as cell reference or named range
            Err(Error::InvalidFormat(format!(
                "Unknown identifier or invalid cell reference: {}",
                ident.to_uppercase()
            )))
        }
    }

    /// Recognize a known function invocation before trying a cell reference.
    fn try_parse_function_call(&mut self) -> Result<Option<Token>> {
        if !self.peek_is_letter() {
            return Ok(None);
        }

        let start = self.position;
        while self
            .input
            .get(self.position)
            .is_some_and(|ch| ch.is_ascii_alphanumeric() || *ch == b'_' || *ch == b'.')
        {
            self.advance();
            // This lexer accepts ASCII identifiers, so a longer spelling
            // cannot match the standard catalog. Avoid scanning a large
            // unknown identifier twice before its ordinary refusal path.
            if self.position - start > MAX_STANDARD_FUNCTION_NAME_BYTES {
                self.position = start;
                return Ok(None);
            }
        }
        let end = self.position;
        if start == end {
            return Ok(None);
        }

        let mut lookahead = end;
        while self
            .input
            .get(lookahead)
            .is_some_and(|ch| ch.is_ascii_whitespace())
        {
            lookahead += 1;
        }
        if self.input.get(lookahead) != Some(&b'(') {
            self.position = start;
            return Ok(None);
        }

        let identifier = match std::str::from_utf8(&self.input[start..end]) {
            Ok(identifier) => identifier,
            Err(_) => {
                self.position = start;
                return Ok(None);
            },
        };
        let Some(function) = lookup_function(identifier) else {
            self.position = start;
            return Ok(None);
        };
        self.position = end;
        Ok(Some(Token::Function(copy_formula_component(
            function,
            "formula function name",
        )?)))
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
                    if quoted && self.input.get(self.position + 1) == Some(&b'\'') {
                        // Apostrophes are escaped by doubling them in ODF
                        // source and sheet names.  Keep a closing bracket in
                        // those quoted spans from ending the reference body.
                        self.position += 2;
                    } else {
                        quoted = !quoted;
                        self.advance();
                    }
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
        parse_open_formula_reference(reference, &self.limits.reference)
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
                sheet = Some(copy_formula_component(sheet_name, "formula sheet name")?);
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
            .map_err(|_error| Error::InvalidFormat("Invalid column".to_string()))?;
        let column = copy_upper_ascii(column, "formula column")?;

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

// The catalog's longest canonical name is 19 ASCII bytes. Four bytes per
// scalar is enough to bound the input that can be normalized without making
// `is_valid_function` perform an unbounded Unicode uppercase allocation.
const MAX_FUNCTION_INPUT_BYTES: usize = MAX_STANDARD_FUNCTION_NAME_BYTES * 4;

fn lookup_function(name: &str) -> Option<&'static str> {
    if name.is_empty() || name.len() > MAX_FUNCTION_INPUT_BYTES {
        return None;
    }

    // Keep the common canonical spelling on the O(1) PHF path.
    if let Some(function) = lookup_canonical_function(name) {
        return Some(function);
    }

    let mut normalized = [0_u8; MAX_STANDARD_FUNCTION_NAME_BYTES];
    let normalized_len = if name.is_ascii() {
        if name.len() > MAX_STANDARD_FUNCTION_NAME_BYTES {
            return None;
        }
        for (index, byte) in name.bytes().enumerate() {
            normalized[index] = byte.to_ascii_uppercase();
        }
        name.len()
    } else {
        let mut normalized_len: usize = 0;
        for character in name.chars() {
            for uppercase in character.to_uppercase() {
                let mut encoded = [0_u8; 4];
                let encoded = uppercase.encode_utf8(&mut encoded).as_bytes();
                let next_len = normalized_len.checked_add(encoded.len())?;
                if next_len > normalized.len() {
                    return None;
                }
                normalized[normalized_len..next_len].copy_from_slice(encoded);
                normalized_len = next_len;
            }
        }
        normalized_len
    };

    let normalized = std::str::from_utf8(&normalized[..normalized_len]).ok()?;
    lookup_canonical_function(normalized)
}

fn lookup_canonical_function(name: &str) -> Option<&'static str> {
    STANDARD_FORMULA_FUNCTIONS.get_key(name).copied()
}

fn strip_open_formula_prefix(value: &str) -> Option<&str> {
    let bytes = value.as_bytes();
    (bytes.len() >= 4 && bytes[..3].eq_ignore_ascii_case(b"of:") && bytes[3] == b'=')
        .then(|| &value[4..])
}

fn reserve_token_slot(tokens: &mut Vec<Token>, maximum: usize) -> Result<()> {
    if tokens.len() >= maximum {
        return Err(formula_limit_error(
            Resource::Objects,
            tokens.len().saturating_add(1),
            maximum,
        ));
    }
    if tokens.len() == tokens.capacity() {
        let next_capacity = if tokens.capacity() == 0 {
            4.min(maximum)
        } else {
            tokens
                .capacity()
                .checked_mul(2)
                .ok_or_else(|| Error::InvalidFormat("formula token capacity overflow".to_string()))?
                .min(maximum)
        };
        let additional = next_capacity
            .checked_sub(tokens.len())
            .ok_or_else(|| Error::InvalidFormat("formula token capacity overflow".to_string()))?;
        tokens
            .try_reserve_exact(additional)
            .map_err(|source| Error::Allocation {
                resource: "formula tokens",
                source,
            })?;
    }
    Ok(())
}

fn formula_limit_error(resource: Resource, actual: usize, maximum: usize) -> Error {
    let Some(observed) = u64::try_from(actual).ok() else {
        return Error::InvalidFormat("formula limit exceeds u64".to_string());
    };
    let Some(limit) = u64::try_from(maximum).ok() else {
        return Error::InvalidFormat("formula limit exceeds u64".to_string());
    };
    Error::ResourceLimit(ResourceLimit {
        resource,
        observed,
        limit,
        scope: Arc::from("ods-formula"),
    })
}

fn copy_formula_component(value: &str, resource: &'static str) -> Result<String> {
    let mut result = String::new();
    result
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    result.push_str(value);
    Ok(result)
}

fn copy_upper_ascii(value: &str, resource: &'static str) -> Result<String> {
    let mut result = copy_formula_component(value, resource)?;
    result.make_ascii_uppercase();
    Ok(result)
}

fn parse_open_formula_reference(value: &str, limits: &reference::Limits) -> Result<Token> {
    let reference = reference::parse_body(value, limits)?;
    Ok(reference.into_token())
}

// ============================================================================
// FORMULA UTILITIES
// ============================================================================

/// Query whether a name matches the case-insensitive ODF 1.4 Part 4 chapter 6
/// function catalog.
///
/// Function names are compared case-insensitively. The lookup uses a bounded
/// stack buffer, so a caller-controlled name cannot trigger an unbounded
/// Unicode-uppercase allocation. Names outside the bounded normalization
/// envelope are rejected. This query does not validate invocation syntax or
/// function arity.
#[inline]
#[must_use]
pub fn is_valid_function(name: &str) -> bool {
    lookup_function(name).is_some()
}

/// Extract legacy cell references from a formula.
///
/// References already represented by legacy tokens are borrowed.  A rich
/// local cell or cell-range reference contributes an owned legacy projection
/// when it has no subtable selector; this keeps quoted and absolute sheet
/// cells visible to existing callers while making the unavoidable loss of
/// those lexical markers explicit in the returned `CellRef`.  External,
/// whole-axis, subtable, inherited, and invalidated references are skipped.
/// Use [`extract_references`] for the complete reference family.
#[must_use]
pub fn extract_cell_refs<'a>(formula: &'a Formula) -> SmallVec<[Cow<'a, CellRef>; 8]> {
    formula
        .tokens
        .iter()
        .filter_map(|token| match token {
            Token::CellRef(cell_ref) => Some(Cow::Borrowed(cell_ref)),
            Token::RangeRef(range_ref) => Some(Cow::Borrowed(&range_ref.start)),
            Token::Reference(reference) => reference.legacy_cell_ref().map(Cow::Owned),
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

/// A borrowed reference occurrence in a parsed formula.
///
/// Legacy tokens remain available for compatible cell and range syntax. Rich
/// bracketed references borrow their boxed [`reference::Reference`] without
/// maintaining a second sidecar allocation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ReferenceView<'a> {
    /// A legacy-compatible cell reference, whether bracketed or unbracketed.
    Cell(&'a CellRef),
    /// A legacy-compatible cell range, whether bracketed or unbracketed.
    Range(&'a RangeRef),
    /// A complete bracketed OpenFormula reference.
    Rich(&'a reference::Reference),
}

/// Extract all reference occurrences without discarding rich OpenFormula
/// metadata.
#[must_use]
pub fn extract_references<'a>(formula: &'a Formula) -> SmallVec<[ReferenceView<'a>; 8]> {
    formula
        .tokens
        .iter()
        .filter_map(|token| match token {
            Token::CellRef(reference) => Some(ReferenceView::Cell(reference)),
            Token::RangeRef(reference) => Some(ReferenceView::Range(reference)),
            Token::Reference(reference) => Some(ReferenceView::Rich(reference)),
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
            | Token::Semicolon
            | Token::Reference(_) => None,
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
    fn test_function_calls_take_precedence_over_cell_reference_shape() {
        let formula = FormulaParser::new("=BIN2DEC (\"101\")")
            .parse()
            .expect("test fixture or operation should succeed");
        assert!(matches!(&formula.tokens[0], Token::Function(name) if name == "BIN2DEC"));
        assert!(matches!(formula.tokens[1], Token::LParen));

        let formula = FormulaParser::new("=BINOM.DIST.RANGE(A1;1;2)")
            .parse()
            .expect("test fixture or operation should succeed");
        assert!(matches!(&formula.tokens[0], Token::Function(name) if name == "BINOM.DIST.RANGE"));

        // Without an invocation parenthesis, the same spelling remains a
        // valid cell reference (column LOG, row 10).
        let formula = FormulaParser::new("=LOG10")
            .parse()
            .expect("test fixture or operation should succeed");
        assert!(matches!(
            &formula.tokens[0],
            Token::CellRef(CellRef { column, row, .. }) if column == "LOG" && *row == 10
        ));
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
            Token::Reference(reference)
                if matches!(
                    reference.as_ref(),
                    reference::Reference::Local(reference::Address::Cells(start, end))
                        if matches!(
                            &start.sheet,
                            reference::SheetSelector::Explicit(locator)
                                if locator.sheet.name == "Inputs" && locator.sheet.absolute
                        )
                            && matches!(
                                &start.value,
                                reference::EndpointValue::Cell(cell)
                                    if cell.column.label == "A"
                                        && cell.row.number == 1
                                        && cell.column.absolute
                                        && cell.row.absolute
                            )
                            && matches!(end.sheet, reference::SheetSelector::Inherited)
                            && matches!(
                                &end.value,
                                reference::EndpointValue::Cell(cell)
                                    if cell.column.label == "B"
                                        && cell.row.number == 2
                                        && cell.column.absolute
                                        && cell.row.absolute
                            )
                )
        ));
    }

    #[test]
    fn test_parse_odf_formula_with_quoted_sheet_reference() {
        let formula = FormulaParser::new("OF:=['Bob''s'.$A$1]")
            .parse()
            .expect("test fixture or operation should succeed");

        assert!(matches!(
            &formula.tokens[0],
            Token::Reference(reference)
                if matches!(
                    reference.as_ref(),
                    reference::Reference::Local(reference::Address::Cell(endpoint))
                        if matches!(
                            &endpoint.sheet,
                            reference::SheetSelector::Explicit(locator)
                                if locator.sheet.name == "Bob's" && locator.sheet.quoted
                        )
                            && matches!(
                                &endpoint.value,
                                reference::EndpointValue::Cell(cell)
                                    if cell.column.label == "A" && cell.row.number == 1
                            )
                )
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
            | Token::Semicolon
            | Token::Reference(_) => panic!("Expected cell reference"),
        }
    }

    #[test]
    fn test_is_valid_function() {
        assert!(is_valid_function("SUM"));
        assert!(is_valid_function("AVERAGE"));
        assert!(!is_valid_function("INVALID_FUNCTION"));
    }

    #[test]
    fn test_function_lookup_is_bounded_and_case_insensitive_without_ascii_only_regression() {
        assert!(is_valid_function("sum"));
        assert!(is_valid_function("ſUM"));
        assert!(is_valid_function("ıF"));
        assert!(!is_valid_function(
            &"A".repeat(MAX_FUNCTION_INPUT_BYTES + 1)
        ));
        assert!(!is_valid_function("not-a-function"));
    }

    #[test]
    fn test_string_literals_preserve_utf8_and_require_a_closing_quote() {
        let formula = FormulaParser::new("=UNICODE(\"α🌟\")")
            .parse()
            .expect("test fixture or operation should succeed");
        assert!(matches!(
            &formula.tokens[2],
            Token::String(value) if value == "α🌟"
        ));

        let formula = FormulaParser::new("=\"a\"\"b\"")
            .parse()
            .expect("test fixture or operation should succeed");
        assert!(matches!(&formula.tokens[0], Token::String(value) if value == "a\"b"));

        assert!(FormulaParser::new("=\"unterminated").parse().is_err());
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
    fn test_extract_cell_refs_projects_legacy_compatible_rich_cells() {
        let formula = FormulaParser::new("=SUM(['Bob''s'.$A$1:.B2])")
            .parse()
            .expect("quoted rich cell range should parse");
        let refs = extract_cell_refs(&formula);
        assert!(matches!(
            refs.as_slice(),
            [Cow::Owned(CellRef {
                sheet: Some(sheet),
                column,
                row: 1,
                column_absolute: true,
                row_absolute: true,
            })] if sheet == "Bob's" && column == "A"
        ));

        let external = FormulaParser::new("=['file:///book.ods'#.A1]")
            .parse()
            .expect("external rich cell should parse");
        assert!(extract_cell_refs(&external).is_empty());
    }

    #[test]
    fn test_formula_limits_forward_reference_limits() {
        let limits = FormulaLimits::default()
            .with_reference_limits(reference::Limits::default().with_max_name_bytes(3));
        let error = FormulaParser::new("of:=[Sheet.A1]")
            .parse_with_limits(&limits)
            .expect_err("formula reference name limit should be enforced");
        assert!(matches!(error, Error::ResourceLimit(_)));
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
            | Token::Semicolon
            | Token::Reference(_) => panic!("Expected cell reference"),
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
            | Token::Semicolon
            | Token::Reference(_) => panic!("Expected range reference"),
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
            | Token::Semicolon
            | Token::Reference(_) => panic!("Expected number token"),
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
            | Token::Semicolon
            | Token::Reference(_) => panic!("Expected string token: {:?}", formula.tokens),
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
                | Token::Semicolon
                | Token::Reference(_) => None,
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
            | Token::Semicolon
            | Token::Reference(_) => panic!("Expected cell reference"),
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
    fn test_every_part4_catalog_name_is_tokenized_as_a_function_call() {
        assert_eq!(STANDARD_FORMULA_FUNCTIONS.len(), 393);
        for name in STANDARD_FORMULA_FUNCTIONS.iter() {
            let formula = FormulaParser::new(&format!("={name}()"))
                .parse()
                .expect("catalog function should parse as an invocation");
            assert!(matches!(&formula.tokens[0], Token::Function(found) if found == *name));
        }
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
                | Token::Semicolon
                | Token::Reference(_) => None,
            })
            .collect();
        assert!(funcs.contains(&"IF"));
        assert!(funcs.contains(&"SUM"));
        assert!(funcs.contains(&"AVERAGE"));
    }
}
