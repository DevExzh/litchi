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

/// Strict, inert OpenFormula 1.4 expression grammar and flat syntax tree.
pub mod expression;
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
    /// Semicolon (function-parameter or inline-array-column separator)
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
    // The legacy sheet grammar scans `[A-Za-z0-9_ ]+` before deciding whether
    // a dot follows. Keep the last immutable-input run so compact misses and
    // adjacent cell-shaped names do not rescan the same suffix quadratically.
    legacy_sheet_scan: Option<(usize, usize)>,
    #[cfg(test)]
    legacy_sheet_scan_work: usize,
}

impl<'a> FormulaParser<'a> {
    /// Create a new formula parser
    #[must_use]
    pub fn new(input: &'a str) -> Self {
        Self {
            input: input.as_bytes(),
            position: 0,
            limits: FormulaLimits::default(),
            legacy_sheet_scan: None,
            #[cfg(test)]
            legacy_sheet_scan_work: 0,
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

    /// Parse an OpenFormula string literal (ODF 1.4 Part 4, §5.4).
    fn parse_string(&mut self) -> Result<Token> {
        self.advance(); // Skip opening quote
        let content_start = self.position;
        let mut cursor = content_start;
        let mut escaped_quote_pairs = 0_usize;

        // First pass: find the real closing quote and count doubled-quote
        // pairs. Formula input has already passed the outer UTF-8 validation,
        // so no temporary decoded string is needed for this pass. Each pair
        // removes one byte from the source span's decoded length.
        let closing = loop {
            let Some(ch) = self.input.get(cursor).copied() else {
                return Err(Error::InvalidFormat(
                    "Unterminated string literal".to_string(),
                ));
            };
            if ch != b'"' {
                cursor += 1;
                continue;
            }

            if self.input.get(cursor + 1) == Some(&b'"') {
                escaped_quote_pairs = escaped_quote_pairs.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Formula string literal length overflow".to_string())
                })?;
                cursor += 2;
                continue;
            }
            break cursor;
        };
        // Part 4 §5.4 excludes U+0000 from the string grammar. Check the
        // admitted content in one optimized slice operation, after the scan
        // has established a complete literal and before allocating its value.
        let content = &self.input[content_start..closing];
        if content.contains(&0) {
            return Err(Error::InvalidFormat(
                "NUL is not allowed in string literal".to_string(),
            ));
        }
        let decoded_bytes = (closing - content_start)
            .checked_sub(escaped_quote_pairs)
            .ok_or_else(|| {
                Error::InvalidFormat("Formula string literal length overflow".to_string())
            })?;

        // Reserve only the decoded payload after the complete syntax has been
        // admitted. Empty literals keep String::new's zero-allocation state.
        let mut result = String::new();
        if decoded_bytes != 0 {
            result
                .try_reserve_exact(decoded_bytes)
                .map_err(|source| Error::Allocation {
                    resource: "formula string literal",
                    source,
                })?;
        }

        // With no doubled quotes, the decoded and source spans have the same
        // byte length. Reuse the first pass's result to avoid rescanning every
        // byte of the common plain-literal case.
        if escaped_quote_pairs == 0 {
            let segment =
                std::str::from_utf8(&self.input[content_start..closing]).map_err(|_error| {
                    Error::InvalidFormat("Invalid UTF-8 in string literal".to_string())
                })?;
            result.push_str(segment);
            self.position = closing + 1;
            return Ok(Token::String(result));
        }

        // Second pass: copy complete UTF-8 spans and collapse doubled quotes.
        // The exact reservation above means these pushes cannot grow the
        // string; they only populate the admitted buffer.
        let mut segment_start = content_start;
        cursor = content_start;
        while cursor < closing {
            if self.input[cursor] == b'"' {
                if segment_start < cursor {
                    let segment = std::str::from_utf8(&self.input[segment_start..cursor]).map_err(
                        |_error| {
                            Error::InvalidFormat("Invalid UTF-8 in string literal".to_string())
                        },
                    )?;
                    result.push_str(segment);
                }
                result.push('"');
                cursor += 2;
                segment_start = cursor;
            } else {
                cursor += 1;
            }
        }
        if segment_start < closing {
            let segment =
                std::str::from_utf8(&self.input[segment_start..closing]).map_err(|_error| {
                    Error::InvalidFormat("Invalid UTF-8 in string literal".to_string())
                })?;
            result.push_str(segment);
        }

        self.position = closing + 1;
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
        // Scan the ordinary ASCII candidate once. A function name may also be
        // a valid column-plus-row spelling (for example, `BIN2DEC` is column
        // BIN, row 2, followed by `DEC`). The complete invocation therefore
        // takes precedence over a cell parse, while a bare spelling still
        // follows the legacy cell/name rules.
        if let Some(token) = self.try_parse_compact_identifier_or_ref()? {
            return match token {
                Token::CellRef(cell_ref) => self.finish_cell_ref(cell_ref),
                token => Ok(token),
            };
        }

        // Try to parse as cell reference first.
        // IMPORTANT: This parse is speculative; if it fails, we must rewind so that
        // the same input can be parsed as a function/name instead.
        if self.peek() == Some(b'.') || self.peek() == Some(b'$') || self.peek_is_letter() {
            let start_pos = self.position;
            match self.try_parse_cell_ref() {
                Ok(cell_ref) => return self.finish_cell_ref(cell_ref),
                // Syntax failure is the expected speculative miss for names
                // such as `SUM`. Allocation and typed resource failures are
                // real failures and must not be hidden by the fallback name
                // parser.
                Err(error) if matches!(&error, Error::InvalidFormat(_)) => {
                    self.position = start_pos;
                },
                Err(error) => return Err(error),
            }
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

    /// Finish a successfully parsed legacy cell and consume a possible range
    /// endpoint. The first endpoint has already been admitted and copied, so
    /// this keeps the common cell path free of a second identifier scan.
    fn finish_cell_ref(&mut self, cell_ref: CellRef) -> Result<Token> {
        self.skip_whitespace();
        if self.peek() == Some(b':') {
            self.advance();
            let end = self.try_parse_cell_ref()?;
            Ok(Token::RangeRef(RangeRef {
                start: cell_ref,
                end,
            }))
        } else {
            Ok(Token::CellRef(cell_ref))
        }
    }

    /// Scan an ASCII identifier/cell candidate once and classify the compact
    /// forms that make up the common formula path.
    ///
    /// Spaced sheet names and malformed suffixes deliberately return `None` so
    /// the established parser below can preserve their exact behavior. This
    /// method restores `position` on a syntax miss, and defers fallible
    /// component copies until after the row has been parsed.
    fn try_parse_compact_identifier_or_ref(&mut self) -> Result<Option<Token>> {
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
            // Keep the previous bounded-name fast rejection. Long inputs are
            // handed to the established parser, which may still recognize a
            // legacy cell with a large column or row component without asking
            // the function catalog to normalize the whole spelling.
            if self.position - start > MAX_STANDARD_FUNCTION_NAME_BYTES {
                self.position = start;
                return Ok(None);
            }
        }
        let end = self.position;
        if start == end {
            return Ok(None);
        }

        // Function invocation takes precedence over the cell shape. Keep
        // whitespace between the name and `(` unconsumed, as the old function
        // lookahead did.
        let mut lookahead = end;
        while self
            .input
            .get(lookahead)
            .is_some_and(|ch| ch.is_ascii_whitespace())
        {
            lookahead += 1;
        }
        if self.input.get(lookahead) == Some(&b'(') {
            // The parser input was validated as UTF-8 before tokenization;
            // this slice is ASCII by construction. Keep the defensive branch
            // so this helper remains total if its caller is reused later.
            let identifier = match std::str::from_utf8(&self.input[start..end]) {
                Ok(identifier) => identifier,
                Err(_) => {
                    self.position = start;
                    return Ok(None);
                },
            };
            if let Some(function) = lookup_function(identifier) {
                self.position = end;
                return Ok(Some(Token::Function(copy_formula_component(
                    function,
                    "formula function name",
                )?)));
            }
        }

        let Some(parts) = parse_compact_cell_parts(&self.input[start..end]) else {
            self.position = start;
            return Ok(None);
        };

        // The legacy sheet parser permits literal spaces inside an unquoted
        // sheet name. For example, `A1 Sheet.B2` is one sheet-qualified cell,
        // while `A1 + Sheet.B2` is two cells. Hand the former back to that
        // parser before copying the compact cell components. The check is
        // limited to the legacy space character; tabs and other formula
        // whitespace were never part of that sheet-name spelling.
        if parts.sheet.is_none() && self.has_spaced_sheet_qualifier(end) {
            self.position = start;
            return Ok(None);
        }

        // All syntax, including checked row parsing, succeeded before these
        // fallible copies. This prevents a failed speculative column/sheet
        // allocation from being mistaken for an identifier fallback.
        let sheet = parts
            .sheet
            .map(|(sheet_start, sheet_end)| {
                let value =
                    std::str::from_utf8(&self.input[start + sheet_start..start + sheet_end])
                        .map_err(|_error| Error::InvalidFormat("Invalid sheet name".to_string()))?;
                copy_formula_component(value, "formula sheet name")
            })
            .transpose()?;
        let column =
            std::str::from_utf8(&self.input[start + parts.column.0..start + parts.column.1])
                .map_err(|_error| Error::InvalidFormat("Invalid column".to_string()))?;
        let column = copy_upper_ascii(column, "formula column")?;

        self.position = end;
        Ok(Some(Token::CellRef(CellRef {
            sheet,
            column,
            row: parts.row,
            column_absolute: false,
            row_absolute: false,
        })))
    }

    /// Check whether a compact cell is followed by a legacy space-bearing
    /// sheet locator. This is only called after a compact coordinate matched,
    /// so it does no work on names, functions, or ordinary operator spacing.
    fn has_spaced_sheet_qualifier(&mut self, end: usize) -> bool {
        if self.input.get(end) != Some(&b' ') {
            return false;
        }
        let position = self.legacy_sheet_scan_end(end);
        self.input.get(position) == Some(&b'.')
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
        let mut sheet_range = None;

        // Parse sheet name (if present)
        if self.peek() == Some(b'.') {
            self.advance();
            // Current sheet reference
        } else if self.peek_is_letter() {
            // Might have a sheet-qualified reference like Sheet1.A1.
            // If there is no dot after the identifier chunk, this is a plain
            // cell reference like A1 and we must rewind.
            let start = self.position;
            let sheet_end = self.legacy_sheet_scan_end(start);
            self.position = sheet_end;

            if self.peek() == Some(b'.') {
                sheet_range = Some((start, sheet_end));
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

        let col_end = self.position;

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

        // Defer component copies until all of the coordinate syntax, including
        // checked row parsing, has succeeded. A speculative name parse can
        // therefore backtrack only syntax errors and never hide an allocation
        // failure from its caller.
        let sheet = sheet_range
            .map(|(start, end)| {
                let sheet_name = std::str::from_utf8(&self.input[start..end])
                    .map_err(|_error| Error::InvalidFormat("Invalid sheet name".to_string()))?;
                copy_formula_component(sheet_name, "formula sheet name")
            })
            .transpose()?;
        let column = std::str::from_utf8(&self.input[col_start..col_end])
            .map_err(|_error| Error::InvalidFormat("Invalid column".to_string()))?;
        let column = copy_upper_ascii(column, "formula column")?;

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

    /// Return the end of the legacy unquoted sheet-name run beginning at
    /// `start`. The input is immutable for a parser, so a cached run also
    /// describes every suffix that starts inside it.
    fn legacy_sheet_scan_end(&mut self, start: usize) -> usize {
        if let Some((cached_start, cached_end)) = self.legacy_sheet_scan
            && start >= cached_start
            && start < cached_end
        {
            return cached_end;
        }

        let mut end = start;
        while self
            .input
            .get(end)
            .is_some_and(|byte| byte.is_ascii_alphanumeric() || *byte == b'_' || *byte == b' ')
        {
            #[cfg(test)]
            {
                self.legacy_sheet_scan_work += 1;
            }
            end += 1;
        }
        self.legacy_sheet_scan = Some((start, end));
        end
    }
}

#[derive(Clone, Copy)]
struct CompactCellParts {
    /// Offset range for an optional compact sheet name.
    sheet: Option<(usize, usize)>,
    /// Offset range for the column label.
    column: (usize, usize),
    /// Checked decimal row value.
    row: u32,
}

/// Parse the compact, ASCII cell spelling from an already scanned candidate.
///
/// The returned ranges are offsets into `input`; no owned component is made
/// until the caller has received a complete, checked coordinate. Returning
/// `None` means a syntax miss, so the caller can preserve the legacy fallback
/// for names, spaced sheet names, and malformed suffixes.
fn parse_compact_cell_parts(input: &[u8]) -> Option<CompactCellParts> {
    let coordinate_start = if let Some(dot) = input.iter().position(|byte| *byte == b'.') {
        if dot == 0
            || input[..dot]
                .iter()
                .any(|byte| !byte.is_ascii_alphanumeric() && *byte != b'_')
        {
            return None;
        }
        Some(dot + 1)
    } else {
        None
    };
    let coordinate_start = coordinate_start.unwrap_or(0);
    let sheet_range = coordinate_start.checked_sub(1).map(|dot| (0, dot));

    let mut position = coordinate_start;
    let column_start = position;
    while input
        .get(position)
        .is_some_and(|byte| byte.is_ascii_alphabetic())
    {
        position += 1;
    }
    if column_start == position {
        return None;
    }
    let column_end = position;

    let row_start = position;
    while input
        .get(position)
        .is_some_and(|byte| byte.is_ascii_digit())
    {
        position += 1;
    }
    if row_start == position || position != input.len() {
        return None;
    }

    let mut row = 0_u32;
    for byte in &input[row_start..position] {
        row = row.checked_mul(10)?.checked_add(u32::from(byte - b'0'))?;
    }

    Some(CompactCellParts {
        sheet: sheet_range,
        column: (column_start, column_end),
        row,
    })
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

#[inline]
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

        let formula = FormulaParser::new("=\"\"+A1")
            .parse()
            .expect("empty string literals should retain token boundaries");
        assert!(matches!(
            &formula.tokens[0],
            Token::String(value) if value.is_empty() && value.capacity() == 0
        ));
        assert!(matches!(formula.tokens[1], Token::Operator('+')));
        assert!(matches!(formula.tokens[2], Token::CellRef(_)));

        let error = FormulaParser::new("=\"a\0b\"").parse();
        assert!(matches!(
            error,
            Err(Error::InvalidFormat(message)) if message.contains("NUL")
        ));

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
    fn test_compact_cell_handoff_preserves_space_bearing_sheet_names() {
        let formula = FormulaParser::new("=A1 Sheet.B2+A1 .C3")
            .parse()
            .expect("legacy space-bearing sheet names should parse");
        assert!(matches!(
            &formula.tokens[0],
            Token::CellRef(CellRef {
                sheet: Some(sheet),
                column,
                row: 2,
                ..
            }) if sheet == "A1 Sheet" && column == "B"
        ));
        assert!(matches!(formula.tokens[1], Token::Operator('+')));
        assert!(matches!(
            &formula.tokens[2],
            Token::CellRef(CellRef {
                sheet: Some(sheet),
                column,
                row: 3,
                ..
            }) if sheet == "A1 " && column == "C"
        ));
    }

    #[test]
    fn test_legacy_sheet_scan_cache_keeps_scan_work_linear() {
        fn tokenize_and_measure(source: &str) -> (usize, usize) {
            let mut parser = FormulaParser::new(source);
            // Keep the original formula prefix out of this private tokenizer
            // loop; parse_with_limits makes the same body transition.
            parser.position = 1;
            let mut token_count = 0;
            while !parser.is_at_end() {
                parser.skip_whitespace();
                if parser.is_at_end() {
                    break;
                }
                parser
                    .next_token()
                    .expect("the generated legacy sequence should tokenize");
                token_count += 1;
            }
            (parser.legacy_sheet_scan_work, token_count)
        }

        const CELLS: usize = 1_024;

        let mut spaced = String::from("=");
        for index in 0..CELLS {
            if index != 0 {
                spaced.push(' ');
            }
            spaced.push_str("A1");
        }
        let (spaced_work, spaced_tokens) = tokenize_and_measure(&spaced);
        assert_eq!(spaced_tokens, CELLS);
        assert!(
            spaced_work <= spaced.len(),
            "spaced sequence rescanned too much input: {spaced_work} > {}",
            spaced.len()
        );

        let contiguous = format!("={}", "A1".repeat(CELLS));
        let (contiguous_work, contiguous_tokens) = tokenize_and_measure(&contiguous);
        assert_eq!(contiguous_tokens, CELLS);
        assert!(
            contiguous_work <= contiguous.len(),
            "contiguous compact misses rescanned too much input: {contiguous_work} > {}",
            contiguous.len()
        );

        let mut ranges = String::from("=");
        for index in 0..CELLS {
            if index != 0 {
                ranges.push(' ');
            }
            ranges.push_str("A1:A1");
        }
        let (range_work, range_tokens) = tokenize_and_measure(&ranges);
        assert_eq!(range_tokens, CELLS);
        assert!(
            range_work <= ranges.len(),
            "range endpoint scans rescanned too much input: {range_work} > {}",
            ranges.len()
        );

        let absolute = format!("={}", "$A$1 ".repeat(CELLS));
        let (absolute_work, absolute_tokens) = tokenize_and_measure(&absolute);
        assert_eq!(absolute_tokens, CELLS);
        assert_eq!(
            absolute_work, 0,
            "absolute cells should not enter sheet scans"
        );

        // The first compact A1 is rejected as a possible space-bearing sheet
        // reference, then the legacy parser starts earlier than that cached
        // suffix. It must rescan only once before later suffixes reuse the
        // new cache endpoint.
        let mut backward = String::from("=A1 ");
        for index in 0..CELLS {
            if index != 0 {
                backward.push(' ');
            }
            backward.push_str("Alias");
        }
        backward.push_str(".B2");
        let (backward_work, backward_tokens) = tokenize_and_measure(&backward);
        assert_eq!(backward_tokens, 1);
        assert!(
            backward_work > backward.len(),
            "backward compact-to-legacy handoff did not exercise its second bounded scan"
        );
        assert!(
            backward_work <= backward.len() * 2,
            "backward handoff exceeded two linear scans: {backward_work} > {}",
            backward.len() * 2
        );
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
