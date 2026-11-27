//! Delimited text: CSV, TSV, PRN and any other single-byte delimiter.
//!
//! [`Tokenizer`] is the one delimited-text tokenizer of the workspace: the
//! text [`Workbook`](crate::Workbook), [`read`] and the unified
//! `File::import` (`litchi-ops`) all split records with it.
//!
//! | Aspect | Behaviour |
//! |---|---|
//! | Records | `\r\n`, `\n` and `\r` end a record. A blank line is a record with one empty field. Records may have different lengths. |
//! | Quotes | A field that *starts* with the quote is quoted: it may hold delimiters and line breaks, a doubled quote is one quote, and text after the closing quote up to the delimiter is kept verbatim (`"a"b` reads `ab`). A quote inside an unquoted field is literal. An unterminated quote is [`Error::UnterminatedQuote`]. [`Dialect::quote`] with `None` reads every field verbatim. |
//! | Escapes | With [`Dialect::escapes`], a backslash inside a quoted field escapes `n`, `r`, `t`, `\` and the quote; another escaped byte keeps its backslash. |
//! | Comments | With [`Dialect::comment`], a record whose first byte is the marker is skipped to the end of its line. |
//! | Characters | With [`Dialect::refuse_controls`], C0 controls other than tab and line breaks, and U+FFFE/U+FFFF, are [`Error::InvalidCharacter`]. |
//!
//! ```
//! use litchi_core::sheet::CellValue;
//! use litchi_sheet_text::delimited::{self, Options};
//!
//! let rows = delimited::read(&b"name,score\n\"Lovelace, Ada\",36\n"[..], &Options::csv())?;
//! assert_eq!(rows[1], [CellValue::String("Lovelace, Ada".into()), CellValue::Int(36)]);
//!
//! let mut out = Vec::new();
//! delimited::write(&rows, &mut out, &Options::tsv())?;
//! assert_eq!(out, b"name\tscore\nLovelace, Ada\t36");
//! # Ok::<(), litchi_sheet_text::Error>(())
//! ```

use std::io::{Read, Write};

use litchi_core::BomKind;
use litchi_core::sheet::CellValue;

use crate::common::emit;
use crate::encoding::Encoding;
use crate::error::{Error, OptionKind, Result};
use crate::limits::{CellBudget, SemanticLimits, SourceLimits};

/// The record grammar a [`Tokenizer`] reads.
///
/// The default is RFC 4180-style CSV: a comma, the `"` quote, no comments,
/// no backslash escapes, and every character accepted.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Dialect {
    delimiter: u8,
    quote: Option<u8>,
    comment: Option<u8>,
    escapes: bool,
    refuse_controls: bool,
}

impl Default for Dialect {
    fn default() -> Self {
        Self::new(b',')
    }
}

impl Dialect {
    /// A dialect with this delimiter, the `"` quote, no comments, no
    /// escapes and every character accepted.
    #[must_use]
    pub const fn new(delimiter: u8) -> Self {
        Self {
            delimiter,
            quote: Some(b'"'),
            comment: None,
            escapes: false,
            refuse_controls: false,
        }
    }

    /// Use this delimiter.
    #[must_use]
    pub const fn delimiter(mut self, delimiter: u8) -> Self {
        self.delimiter = delimiter;
        self
    }

    /// Use this quote, or `None` to read fields verbatim.
    #[must_use]
    pub const fn quote(mut self, quote: Option<u8>) -> Self {
        self.quote = quote;
        self
    }

    /// Skip records starting with this marker, or `None` for no comments.
    #[must_use]
    pub const fn comment(mut self, comment: Option<u8>) -> Self {
        self.comment = comment;
        self
    }

    /// Whether a backslash escapes inside quoted fields.
    #[must_use]
    pub const fn escapes(mut self, escapes: bool) -> Self {
        self.escapes = escapes;
        self
    }

    /// Whether characters no spreadsheet cell can store are refused.
    #[must_use]
    pub const fn refuse_controls(mut self, refuse: bool) -> Self {
        self.refuse_controls = refuse;
        self
    }

    /// The delimiter byte.
    #[must_use]
    pub const fn delimiter_byte(&self) -> u8 {
        self.delimiter
    }

    /// The quote byte, if any.
    #[must_use]
    pub const fn quote_byte(&self) -> Option<u8> {
        self.quote
    }

    /// The comment marker, if any.
    #[must_use]
    pub const fn comment_byte(&self) -> Option<u8> {
        self.comment
    }

    /// Whether backslash escapes are read.
    #[must_use]
    pub const fn has_escapes(&self) -> bool {
        self.escapes
    }

    /// Check that the bytes can delimit fields unambiguously.
    ///
    /// # Errors
    ///
    /// [`Error::InvalidOption`] naming the first unusable byte.
    pub fn validate(&self) -> Result<()> {
        let usable = |byte: u8| byte.is_ascii() && !matches!(byte, b'\r' | b'\n' | 0);
        if let Some(quote) = self.quote
            && !usable(quote)
        {
            return Err(Error::InvalidOption {
                option: OptionKind::Quote,
            });
        }
        if !usable(self.delimiter) || Some(self.delimiter) == self.quote {
            return Err(Error::InvalidOption {
                option: OptionKind::Delimiter,
            });
        }
        if let Some(comment) = self.comment
            && (!usable(comment) || comment == self.delimiter || Some(comment) == self.quote)
        {
            return Err(Error::InvalidOption {
                option: OptionKind::Comment,
            });
        }
        Ok(())
    }
}

/// An Excel `sep=X` first line: the declared ASCII delimiter and the text
/// after the line, or `None` and the whole text.
///
/// ```
/// use litchi_sheet_text::delimited::declared_delimiter;
///
/// assert_eq!(declared_delimiter("sep=;\na;b"), (Some(b';'), "a;b"));
/// assert_eq!(declared_delimiter("a,b"), (None, "a,b"));
/// ```
#[must_use]
pub fn declared_delimiter(text: &str) -> (Option<u8>, &str) {
    let Some(rest) = text.strip_prefix("sep=") else {
        return (None, text);
    };
    let bytes = rest.as_bytes();
    let Some(&delimiter) = bytes.first() else {
        return (None, text);
    };
    if !delimiter.is_ascii() || matches!(delimiter, b'\r' | b'\n' | 0) {
        return (None, text);
    }
    let after = &rest[1..];
    let body = if let Some(body) = after.strip_prefix("\r\n") {
        body
    } else if let Some(body) = after
        .strip_prefix('\n')
        .or_else(|| after.strip_prefix('\r'))
    {
        body
    } else if after.is_empty() {
        after
    } else {
        return (None, text);
    };
    (Some(delimiter), body)
}

/// A borrowing record and field tokenizer over decoded text.
///
/// ```
/// use litchi_sheet_text::delimited::{Dialect, Tokenizer};
///
/// let mut tokenizer = Tokenizer::new("a,\"b,c\"\n# note\nd", Dialect::new(b',').comment(Some(b'#')))?;
/// let mut records = Vec::new();
/// while tokenizer.next_record() {
///     let mut fields = Vec::new();
///     loop {
///         fields.push(tokenizer.field()?.to_owned());
///         if !tokenizer.next_field() {
///             break;
///         }
///     }
///     tokenizer.end_record();
///     records.push(fields);
/// }
/// assert_eq!(records, [vec!["a", "b,c"], vec!["d"]]);
/// # Ok::<(), litchi_sheet_text::Error>(())
/// ```
#[derive(Debug)]
pub struct Tokenizer<'a> {
    input: &'a [u8],
    text: &'a str,
    position: usize,
    line: usize,
    dialect: Dialect,
    scratch: String,
    /// The class of each byte inside an unquoted field.
    unquoted: Classes,
    /// The class of each byte inside a quoted field.
    quoted: Classes,
}

/// A byte class table: `0` for an ordinary byte, else [`STOP`], [`CHECK`]
/// or [`SPECIAL`]. One lookup per byte replaces the comparisons against
/// the delimiter, line breaks, the quote and the refused characters.
#[derive(Clone)]
struct Classes([u8; 256]);

impl std::fmt::Debug for Classes {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("Classes")
    }
}

/// Ends an unquoted field (the delimiter or a line break).
const STOP: u8 = 1;
/// May start a refused character (a C0 control or the lead byte of
/// U+FFFE/U+FFFF).
const CHECK: u8 = 2;
/// Needs handling inside a quoted field: the quote, a line break or an
/// escaping backslash.
const SPECIAL: u8 = 4;

impl Classes {
    fn new(dialect: &Dialect, quoted: bool) -> Self {
        let mut table = [0u8; 256];
        if dialect.refuse_controls {
            for byte in 0..0x20u8 {
                if !matches!(byte, b'\t' | b'\r' | b'\n') {
                    table[usize::from(byte)] = CHECK;
                }
            }
            table[0xEF] = CHECK;
        }
        if quoted {
            for byte in [b'\r', b'\n'] {
                table[usize::from(byte)] = SPECIAL;
            }
            if let Some(quote) = dialect.quote {
                table[usize::from(quote)] = SPECIAL;
            }
            if dialect.escapes {
                table[usize::from(b'\\')] = SPECIAL;
            }
        } else {
            for byte in [dialect.delimiter, b'\r', b'\n'] {
                table[usize::from(byte)] = STOP;
            }
        }
        Self(table)
    }

    #[inline]
    fn of(&self, byte: u8) -> u8 {
        self.0[usize::from(byte)]
    }
}

impl<'a> Tokenizer<'a> {
    /// Tokenize `text` under `dialect`, counting lines from 1.
    ///
    /// # Errors
    ///
    /// [`Error::InvalidOption`] when [`Dialect::validate`] refuses it.
    pub fn new(text: &'a str, dialect: Dialect) -> Result<Self> {
        dialect.validate()?;
        Ok(Self {
            input: text.as_bytes(),
            text,
            position: 0,
            line: 1,
            dialect,
            scratch: String::new(),
            unquoted: Classes::new(&dialect, false),
            quoted: Classes::new(&dialect, true),
        })
    }

    /// Number the first line `line` (for text after a `sep=` line).
    #[must_use]
    pub const fn first_line(mut self, line: usize) -> Self {
        self.line = line;
        self
    }

    /// The current one-based line.
    #[must_use]
    pub const fn line(&self) -> usize {
        self.line
    }

    /// Whether the whole input has been read.
    #[inline]
    #[must_use]
    pub fn at_end(&self) -> bool {
        self.position >= self.input.len()
    }

    /// Start the next record, skipping comment lines; `false` at the end.
    #[inline]
    pub fn next_record(&mut self) -> bool {
        if let Some(comment) = self.dialect.comment {
            while self.input.get(self.position) == Some(&comment) {
                while let Some(&byte) = self.input.get(self.position) {
                    if byte == b'\r' || byte == b'\n' {
                        break;
                    }
                    self.position += 1;
                }
                self.end_record();
            }
        }
        !self.at_end()
    }

    /// Read one field, leaving the position on its delimiter, line break or
    /// the end. The text borrows the input, or a scratch buffer for a quoted
    /// field.
    ///
    /// # Errors
    ///
    /// [`Error::UnterminatedQuote`], and [`Error::InvalidCharacter`] under
    /// [`Dialect::refuse_controls`].
    #[inline]
    pub fn field(&mut self) -> Result<&str> {
        if self
            .dialect
            .quote
            .is_some_and(|quote| self.input.get(self.position) == Some(&quote))
        {
            return self.quoted();
        }
        let start = self.position;
        self.unquoted_run()?;
        Ok(&self.text[start..self.position])
    }

    /// Advance over unquoted bytes up to a delimiter, line break or the end.
    #[inline]
    fn unquoted_run(&mut self) -> Result<()> {
        while let Some(&byte) = self.input.get(self.position) {
            match self.unquoted.of(byte) {
                0 => {},
                STOP => break,
                _ => self.check(byte)?,
            }
            self.position += 1;
        }
        Ok(())
    }

    fn quoted(&mut self) -> Result<&str> {
        let quote = self.dialect.quote.unwrap_or(b'"');
        let line = self.line;
        self.position += 1;
        self.scratch.clear();
        let mut run = self.position;
        loop {
            let Some(&byte) = self.input.get(self.position) else {
                return Err(Error::UnterminatedQuote { line });
            };
            match self.quoted.of(byte) {
                0 => {},
                CHECK => self.check(byte)?,
                _ if byte == quote => {
                    self.scratch.push_str(&self.text[run..self.position]);
                    if self.input.get(self.position + 1) == Some(&quote) {
                        self.scratch.push(char::from(quote));
                        self.position += 2;
                        run = self.position;
                        continue;
                    }
                    self.position += 1;
                    break;
                },
                _ if byte == b'\\' => {
                    let escaped = match self.input.get(self.position + 1) {
                        Some(b'n') => Some('\n'),
                        Some(b'r') => Some('\r'),
                        Some(b't') => Some('\t'),
                        Some(b'\\') => Some('\\'),
                        Some(&next) if next == quote => Some(char::from(quote)),
                        _ => None,
                    };
                    if let Some(escaped) = escaped {
                        self.scratch.push_str(&self.text[run..self.position]);
                        self.scratch.push(escaped);
                        self.position += 2;
                        run = self.position;
                        continue;
                    }
                },
                // A line break.
                _ => {
                    if byte == b'\n' || self.input.get(self.position + 1) != Some(&b'\n') {
                        self.line += 1;
                    }
                },
            }
            self.position += 1;
        }
        // Text after the closing quote up to the delimiter is kept verbatim,
        // as spreadsheet applications read `"a"b`.
        let tail = self.position;
        self.unquoted_run()?;
        self.scratch.push_str(&self.text[tail..self.position]);
        Ok(&self.scratch)
    }

    /// Refuse the byte at the current position when it starts a character
    /// no cell can store: a C0 control other than tab and line breaks, or
    /// U+FFFE/U+FFFF (`EF BF BE`/`EF BF BF`).
    #[inline]
    fn check(&self, byte: u8) -> Result<()> {
        let invalid = (byte < 0x20 && !matches!(byte, b'\t' | b'\r' | b'\n'))
            || (byte == 0xEF
                && self.input.get(self.position + 1) == Some(&0xBF)
                && matches!(self.input.get(self.position + 2), Some(0xBE | 0xBF)));
        if invalid {
            return Err(Error::InvalidCharacter { line: self.line });
        }
        Ok(())
    }

    /// Consume a delimiter after a field; `false` at a line break or the
    /// end, where the record ends.
    #[inline]
    pub fn next_field(&mut self) -> bool {
        if self.input.get(self.position) == Some(&self.dialect.delimiter) {
            self.position += 1;
            return true;
        }
        false
    }

    /// Consume the line break that ends a record.
    #[inline]
    pub fn end_record(&mut self) {
        match self.input.get(self.position) {
            Some(b'\r') => {
                self.position += 1;
                if self.input.get(self.position) == Some(&b'\n') {
                    self.position += 1;
                }
                self.line += 1;
            },
            Some(b'\n') => {
                self.position += 1;
                self.line += 1;
            },
            _ => {},
        }
    }
}

/// Options for delimited text: the [`Dialect`], field trimming, the source
/// encoding, the byte-order mark to write, and the read limits.
///
/// The default ([`Options::csv`]) reads a comma, the `"` quote, `#` comment
/// lines and backslash escapes inside quotes, keeps whitespace, decodes by
/// byte-order mark ([`Encoding::Auto`]) and writes no byte-order mark. The
/// setters are infallible; an unusable byte is refused by the read as
/// [`Error::InvalidOption`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Options {
    dialect: Dialect,
    trim: bool,
    encoding: Encoding,
    write_bom: Option<BomKind>,
    source_limits: SourceLimits,
    semantic_limits: SemanticLimits,
}

impl Default for Options {
    fn default() -> Self {
        Self::csv()
    }
}

impl Options {
    /// Comma-separated values.
    #[must_use]
    pub fn csv() -> Self {
        Self {
            dialect: Dialect::new(b',').comment(Some(b'#')).escapes(true),
            trim: false,
            encoding: Encoding::Auto,
            write_bom: None,
            source_limits: SourceLimits::default(),
            semantic_limits: SemanticLimits::default(),
        }
    }

    /// Tab-separated values.
    #[must_use]
    pub fn tsv() -> Self {
        Self::csv().delimiter(b'\t')
    }

    /// Semicolon-separated values (PRN).
    #[must_use]
    pub fn prn() -> Self {
        Self::csv().delimiter(b';')
    }

    /// Pipe-separated values.
    #[must_use]
    pub fn pipe() -> Self {
        Self::csv().delimiter(b'|')
    }

    /// Use this field delimiter.
    #[must_use]
    pub const fn delimiter(mut self, delimiter: u8) -> Self {
        self.dialect = self.dialect.delimiter(delimiter);
        self
    }

    /// Use this quote, or `None` to read and write fields verbatim.
    #[must_use]
    pub const fn quote(mut self, quote: Option<u8>) -> Self {
        self.dialect = self.dialect.quote(quote);
        self
    }

    /// Skip lines starting with this marker, or `None` for no comments.
    #[must_use]
    pub const fn comment(mut self, comment: Option<u8>) -> Self {
        self.dialect = self.dialect.comment(comment);
        self
    }

    /// Whether a backslash escapes inside quoted fields.
    #[must_use]
    pub const fn escapes(mut self, escapes: bool) -> Self {
        self.dialect = self.dialect.escapes(escapes);
        self
    }

    /// Whether ASCII whitespace is trimmed from both ends of each field.
    #[must_use]
    pub const fn trim(mut self, trim: bool) -> Self {
        self.trim = trim;
        self
    }

    /// Decode the source with this encoding.
    #[must_use]
    pub const fn encoding(mut self, encoding: Encoding) -> Self {
        self.encoding = encoding;
        self
    }

    /// Write this byte-order mark before the output, or none.
    #[must_use]
    pub const fn write_bom(mut self, bom: Option<BomKind>) -> Self {
        self.write_bom = bom;
        self
    }

    /// Bound the bytes read.
    #[must_use]
    pub const fn source_limits(mut self, limits: SourceLimits) -> Self {
        self.source_limits = limits;
        self
    }

    /// Bound the cells materialized.
    #[must_use]
    pub const fn semantic_limits(mut self, limits: SemanticLimits) -> Self {
        self.semantic_limits = limits;
        self
    }

    /// The record grammar.
    #[must_use]
    pub const fn dialect(&self) -> Dialect {
        self.dialect
    }
}

/// Read every record of a delimited source, inferring each field's
/// [`CellValue`] with [`CellValue::infer_from_str`].
///
/// # Errors
///
/// [`Error::Io`], [`Error::Encoding`], [`Error::InvalidOption`],
/// [`Error::UnterminatedQuote`] and [`Error::Limit`].
pub fn read(reader: impl Read, options: &Options) -> Result<Vec<Vec<CellValue>>> {
    Ok(read_with_bom(reader, options)?.0)
}

/// [`read`], also reporting the byte-order mark found.
pub(crate) fn read_with_bom(
    reader: impl Read,
    options: &Options,
) -> Result<(Vec<Vec<CellValue>>, Option<BomKind>)> {
    options.dialect.validate()?;
    let bytes = options.source_limits.read(reader)?;
    let (text, bom) = options.encoding.decode(&bytes)?;
    Ok((parse(&text, options)?, bom))
}

fn parse(text: &str, options: &Options) -> Result<Vec<Vec<CellValue>>> {
    let mut tokenizer = Tokenizer::new(text, options.dialect)?;
    let mut budget = CellBudget::new(options.semantic_limits);
    let mut rows = Vec::new();
    while tokenizer.next_record() {
        let mut row = Vec::new();
        loop {
            budget.charge(1)?;
            let field = tokenizer.field()?;
            let field = if options.trim {
                field.trim_matches(|character: char| character.is_ascii_whitespace())
            } else {
                field
            };
            row.try_reserve(1)?;
            row.push(CellValue::infer_from_str(field));
            if !tokenizer.next_field() {
                break;
            }
        }
        tokenizer.end_record();
        rows.try_reserve(1)?;
        rows.push(row);
    }
    Ok(rows)
}

/// Write `rows` as delimited text: rows separated by `\n` (none after the
/// last), and a field quoted when it holds the delimiter, the quote, a line
/// break, a leading comment marker, or (with escapes) a backslash, and
/// always for booleans, dates, errors and formulas (`=` prefixed).
///
/// # Errors
///
/// [`Error::Io`] and [`Error::InvalidOption`].
pub fn write(rows: &[Vec<CellValue>], writer: impl Write, options: &Options) -> Result<()> {
    options.dialect.validate()?;
    emit(writer, options.write_bom, |writer| {
        body(rows, writer, &options.dialect)
    })
}

fn body(rows: &[Vec<CellValue>], writer: &mut dyn Write, dialect: &Dialect) -> Result<()> {
    let delimiter = dialect.delimiter;
    for (row_index, row) in rows.iter().enumerate() {
        for (column, cell) in row.iter().enumerate() {
            if column > 0 {
                writer.write_all(&[delimiter])?;
            }
            let (field, always) = match cell {
                CellValue::Empty => (String::new(), false),
                CellValue::Bool(value) => (if *value { "TRUE" } else { "FALSE" }.to_owned(), true),
                CellValue::Int(value) => (value.to_string(), false),
                CellValue::Float(value) => (value.to_string(), false),
                CellValue::DateTime(value) => (value.to_string(), true),
                CellValue::String(value) => (value.clone(), false),
                CellValue::Error(value) => (value.clone(), true),
                CellValue::Formula { formula, .. } => (format!("={formula}"), true),
            };
            write_field(writer, &field, always, column == 0, dialect)?;
        }
        if row_index + 1 < rows.len() {
            writer.write_all(b"\n")?;
        }
    }
    Ok(())
}

fn write_field(
    writer: &mut dyn Write,
    field: &str,
    always: bool,
    first: bool,
    dialect: &Dialect,
) -> Result<()> {
    let Some(quote) = dialect.quote else {
        writer.write_all(field.as_bytes())?;
        return Ok(());
    };
    let bytes = field.as_bytes();
    let needs = always
        || bytes.iter().any(|&byte| {
            byte == dialect.delimiter || byte == quote || byte == b'\n' || byte == b'\r'
        })
        || (first && dialect.comment.is_some() && bytes.first() == dialect.comment.as_ref())
        || (dialect.escapes && bytes.contains(&b'\\'));
    if !needs {
        writer.write_all(bytes)?;
        return Ok(());
    }
    let mut quoted = Vec::with_capacity(bytes.len() + 2);
    quoted.push(quote);
    for &byte in bytes {
        if byte == quote {
            quoted.push(quote);
        } else if byte == b'\\' && dialect.escapes {
            quoted.push(b'\\');
        }
        quoted.push(byte);
    }
    quoted.push(quote);
    writer.write_all(&quoted)?;
    Ok(())
}

#[cfg(test)]
mod tests;
