//! CSV and TSV import into a worksheet (design §13 P0, record 0266).
//!
//! [`File::import`](crate::File::import) decodes delimited text, infers each
//! field's type, and writes the fields into a worksheet of an XLSX, XLSB,
//! XLS or ODS workbook in **one** atomic native commit (XLS: chained chunks
//! of one mutation), starting at the target cell
//! (`/Sheet1/B2`) or at `A1` of the target sheet (`/Sheet1`). The original
//! handle is never changed; empty input is an exact no-op.
//!
//! ```
//! # #[cfg(feature = "xlsx")]
//! # {
//! use litchi_core::FileFormat;
//! use litchi_ops::import::Options;
//!
//! let file = litchi_ops::create(FileFormat::Xlsx)?;
//! let commit = file.import("/Sheet1", b"name,score\nAda,36\n", &Options::csv().header(true))?;
//! assert_eq!(commit.value().cells(), 4);
//! assert_eq!(commit.value().numbers(), 1);
//! let score = commit.file().get("/Sheet1/B2")?;
//! assert_eq!(score.view().text()?.as_deref(), Some("36"));
//! # }
//! # Ok::<(), litchi_ops::Error>(())
//! ```
//!
//! # Text
//!
//! | Aspect | Behaviour |
//! |---|---|
//! | Encoding | [`Encoding::Auto`] (the default) reads a UTF-8, UTF-16 LE or UTF-16 BE byte-order mark and otherwise assumes UTF-8; an explicit encoding strips its own mark. Malformed input is [`Code::InvalidInput`](crate::Code::InvalidInput). |
//! | Delimiter | One ASCII character; by default the file's own `sep=X` first line, else `,` ([`Options::tsv`] uses a tab). The `sep=` line is never data. |
//! | Quote | `"` by default (doubled to escape); [`Options::quote`] with `None` reads fields verbatim. Quoted fields may contain delimiters and line breaks; an unterminated quote is [`Code::InvalidInput`](crate::Code::InvalidInput). |
//! | Records | `\r\n`, `\n` and `\r` end a record. A blank line is a record with one empty field, so later rows keep their line positions. Records may have different lengths. |
//! | Header | With [`Options::header`] the first record is written as text, never inferred. |
//! | Characters | C0 controls other than tab, line feed and carriage return, and the non-characters U+FFFE and U+FFFF, are [`Code::InvalidInput`](crate::Code::InvalidInput): no spreadsheet format stores them. |
//! | Cell length | A field longer than 32,767 characters (Excel's cell limit) is [`Code::InvalidValue`](crate::Code::InvalidValue). |
//!
//! # Types
//!
//! Inference is on by default and can be switched off as a whole
//! ([`Options::infer`]) or per kind. In order:
//!
//! 1. a number: an optional sign, digits with an optional fraction, and an
//!    optional exponent (`-12.5e3`); no grouping separators or currency;
//!    finite values only;
//! 2. an ISO 8601 date (`2026-10-13`) or date-time (`2026-10-13T08:30:00`,
//!    with an optional fraction and `Z`, or a space for `T`), stored as a
//!    serial number in the workbook's date system with a `yyyy-mm-dd` or
//!    `yyyy-mm-dd hh:mm:ss` number format. Dates before the date system's
//!    epoch stay text. XLS has no unified number-format write yet, so XLS
//!    keeps dates as text ([`Summary::dates_as_text`]); ODS stores them as
//!    native `date` cells (record 0295);
//! 3. `TRUE` or `FALSE`, in any case;
//! 4. text.
//!
//! A field starting with `=` is text unless [`Options::formulas`] is on, in
//! which case it is written as a formula. Formulas are off by default so an
//! untrusted file cannot plant formulas.
//!
//! An empty field writes nothing on an absent cell and clears the value of an
//! existing one, keeping its formatting.
//!
//! # Limits
//!
//! The transaction's [`SemanticLimits`]
//! bound the import: the encoded input and the staged payload (each field's
//! text plus a fixed per-cell charge) are bounded by `max_staged_bytes`, the
//! number of fields by `max_expanded_targets`, and the retained native
//! snapshots by `max_snapshot_bytes`. The data must also fit the format's
//! grid from the origin cell ([`Code::OutOfRange`](crate::Code::OutOfRange)).

#![cfg_attr(
    not(any(feature = "xlsx", feature = "xlsb", feature = "xls", feature = "ods")),
    allow(
        dead_code,
        reason = "only the workbook import adapters consume the parsed table"
    )
)]

use std::fmt;
use std::ops::Range;

use crate::error::{Error as UnifiedError, Result};
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use litchi_sheet_text::delimited::{Dialect, Tokenizer, declared_delimiter};

#[cfg(test)]
mod tests;

/// Excel's maximum number of characters in one cell.
pub(crate) const MAX_CELL_CHARS: usize = 32_767;

/// Refuse a rewritten cell text longer than [`MAX_CELL_CHARS`] (find and
/// merge, record 0295).
#[cfg(any(feature = "xlsb", feature = "xls", feature = "ods"))]
pub(crate) fn check_cell_length(text: &str, format: litchi_core::FileFormat) -> Result<()> {
    if text.len() > MAX_CELL_CHARS && text.chars().count() > MAX_CELL_CHARS {
        return Err(crate::Error::new(
            crate::Code::InvalidValue,
            "a rewritten cell would be longer than 32,767 characters",
        )
        .with_format(format));
    }
    Ok(())
}

/// The fixed staging charge of one imported field, beside its text.
const FIELD_CHARGE: usize = 64;

/// How the source bytes are decoded: the encoding of the shared
/// `litchi-sheet-text` decoder, which import uses.
pub use litchi_sheet_text::Encoding;

/// Import options: CSV with type inference and no formulas by default.
///
/// The setters are infallible; an invalid delimiter or quote is reported by
/// the import itself as [`Code::InvalidArgument`](crate::Code::InvalidArgument).
///
/// ```
/// use litchi_ops::import::{Encoding, Options};
///
/// let tsv = Options::tsv().header(true).encoding(Encoding::Utf16Le).infer_dates(false);
/// let semicolons = Options::csv().delimiter(';').infer(false);
/// assert_ne!(tsv, semicolons);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Options {
    delimiter: Option<char>,
    quote: Option<char>,
    header: bool,
    encoding: Encoding,
    numbers: bool,
    booleans: bool,
    dates: bool,
    formulas: bool,
}

impl Default for Options {
    fn default() -> Self {
        Self::new()
    }
}

impl Options {
    /// CSV options: the delimiter from a `sep=` line, else `,`.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            delimiter: None,
            quote: Some('"'),
            header: false,
            encoding: Encoding::Auto,
            numbers: true,
            booleans: true,
            dates: true,
            formulas: false,
        }
    }

    /// The same as [`Self::new`].
    #[must_use]
    pub const fn csv() -> Self {
        Self::new()
    }

    /// TSV options: a tab delimiter.
    #[must_use]
    pub const fn tsv() -> Self {
        Self::new().with_delimiter('\t')
    }

    const fn with_delimiter(mut self, delimiter: char) -> Self {
        self.delimiter = Some(delimiter);
        self
    }

    /// Use this field delimiter: one ASCII character other than the quote,
    /// a line break or NUL. It overrides a `sep=` line.
    #[must_use]
    pub const fn delimiter(self, delimiter: char) -> Self {
        self.with_delimiter(delimiter)
    }

    /// Set the quote character, or `None` to read fields verbatim.
    #[must_use]
    pub const fn quote(mut self, quote: Option<char>) -> Self {
        self.quote = quote;
        self
    }

    /// Whether the first record is a header row, written as text.
    #[must_use]
    pub const fn header(mut self, header: bool) -> Self {
        self.header = header;
        self
    }

    /// Set the source encoding.
    #[must_use]
    pub const fn encoding(mut self, encoding: Encoding) -> Self {
        self.encoding = encoding;
        self
    }

    /// Turn number, date and boolean inference on or off together.
    #[must_use]
    pub const fn infer(mut self, infer: bool) -> Self {
        self.numbers = infer;
        self.booleans = infer;
        self.dates = infer;
        self
    }

    /// Turn number inference on or off.
    #[must_use]
    pub const fn infer_numbers(mut self, infer: bool) -> Self {
        self.numbers = infer;
        self
    }

    /// Turn boolean (`TRUE`/`FALSE`) inference on or off.
    #[must_use]
    pub const fn infer_booleans(mut self, infer: bool) -> Self {
        self.booleans = infer;
        self
    }

    /// Turn ISO 8601 date inference on or off.
    #[must_use]
    pub const fn infer_dates(mut self, infer: bool) -> Self {
        self.dates = infer;
        self
    }

    /// Whether a field starting with `=` is written as a formula. Off by
    /// default: such fields are text.
    #[must_use]
    pub const fn formulas(mut self, formulas: bool) -> Self {
        self.formulas = formulas;
        self
    }
}

/// Which option an [`Error::InvalidOption`] names.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum OptionKind {
    /// The delimiter is not ASCII, is a line break or NUL, or equals the
    /// quote.
    Delimiter,
    /// The quote is not ASCII or is a line break or NUL.
    Quote,
}

/// A content-free import failure. Positions are one-based; `line` counts
/// physical lines of the decoded text, so a quoted line break advances it.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// An option cannot be used.
    InvalidOption {
        /// The option.
        option: OptionKind,
    },
    /// The bytes are not valid in the selected encoding.
    Encoding {
        /// Byte offset of the first invalid sequence.
        offset: usize,
    },
    /// A quoted field is not closed before the end of the input.
    UnterminatedQuote {
        /// The line where the field starts.
        line: usize,
    },
    /// A C0 control character other than tab and line breaks, or the
    /// non-character U+FFFE or U+FFFF: no spreadsheet cell can store them.
    InvalidCharacter {
        /// The line where it occurs.
        line: usize,
    },
    /// A field is longer than 32,767 characters.
    FieldTooLong {
        /// The one-based record.
        record: usize,
        /// The one-based field within the record.
        field: usize,
    },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidOption {
                option: OptionKind::Delimiter,
            } => formatter.write_str(
                "the import delimiter must be one ASCII character other than the quote, a line break or NUL",
            ),
            Self::InvalidOption {
                option: OptionKind::Quote,
            } => formatter
                .write_str("the import quote must be one ASCII character other than a line break or NUL"),
            Self::Encoding { offset } => write!(
                formatter,
                "the import source is not valid in its encoding at byte {offset}"
            ),
            Self::UnterminatedQuote { line } => write!(
                formatter,
                "the quoted field starting on line {line} is not closed"
            ),
            Self::InvalidCharacter { line } => write!(
                formatter,
                "line {line} contains a character no spreadsheet cell can store"
            ),
            Self::FieldTooLong { record, field } => write!(
                formatter,
                "field {field} of record {record} is longer than {MAX_CELL_CHARS} characters"
            ),
        }
    }
}

impl std::error::Error for Error {}

/// Content-free facts about one import, the closure value of the returned
/// [`Commit`](crate::transaction::Commit).
///
/// Coordinates are zero-based; `rows` and `columns` describe the imported
/// rectangle from the origin, including empty fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Summary {
    sheet: usize,
    row: u32,
    column: u32,
    rows: u32,
    columns: u32,
    cells: usize,
    texts: usize,
    numbers: usize,
    booleans: usize,
    dates: usize,
    dates_as_text: usize,
    formulas: usize,
    empty: usize,
}

impl Summary {
    /// The worksheet's zero-based tab position.
    #[must_use]
    pub const fn sheet(&self) -> usize {
        self.sheet
    }

    /// The zero-based origin row.
    #[must_use]
    pub const fn row(&self) -> u32 {
        self.row
    }

    /// The zero-based origin column.
    #[must_use]
    pub const fn column(&self) -> u32 {
        self.column
    }

    /// Records imported, the header included.
    #[must_use]
    pub const fn rows(&self) -> u32 {
        self.rows
    }

    /// Fields in the longest record.
    #[must_use]
    pub const fn columns(&self) -> u32 {
        self.columns
    }

    /// Non-empty fields written.
    #[must_use]
    pub const fn cells(&self) -> usize {
        self.cells
    }

    /// Fields written as text, header fields and dates kept as text
    /// included.
    #[must_use]
    pub const fn texts(&self) -> usize {
        self.texts
    }

    /// Fields written as numbers.
    #[must_use]
    pub const fn numbers(&self) -> usize {
        self.numbers
    }

    /// Fields written as booleans.
    #[must_use]
    pub const fn booleans(&self) -> usize {
        self.booleans
    }

    /// Fields written as dates (serial numbers with a date format).
    #[must_use]
    pub const fn dates(&self) -> usize {
        self.dates
    }

    /// Date fields kept as text: before the date system's epoch, or in a
    /// format without a unified number-format write (XLS).
    #[must_use]
    pub const fn dates_as_text(&self) -> usize {
        self.dates_as_text
    }

    /// Fields written as formulas.
    #[must_use]
    pub const fn formulas(&self) -> usize {
        self.formulas
    }

    /// Empty fields (each clears an existing cell's value).
    #[must_use]
    pub const fn empty(&self) -> usize {
        self.empty
    }

    pub(crate) fn at(origin: Origin, table: &Table) -> Self {
        Self {
            sheet: origin.sheet,
            row: origin.row,
            column: origin.column,
            rows: table.rows,
            columns: table.columns,
            ..Self::default()
        }
    }

    /// Count one written field.
    pub(crate) fn count(&mut self, written: Written) {
        match written {
            Written::Text => self.texts += 1,
            Written::Number => self.numbers += 1,
            Written::Boolean => self.booleans += 1,
            Written::Date => self.dates += 1,
            Written::DateAsText => {
                self.dates_as_text += 1;
                self.texts += 1;
            },
            Written::Formula => self.formulas += 1,
            Written::Empty => {
                self.empty += 1;
                return;
            },
        }
        self.cells += 1;
    }
}

/// How one field was written, for [`Summary::count`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Written {
    Text,
    Number,
    Boolean,
    #[cfg_attr(
        not(any(feature = "xlsx", feature = "xlsb", feature = "ods")),
        allow(dead_code, reason = "XLS keeps imported dates as text")
    )]
    Date,
    #[cfg_attr(
        not(any(feature = "xlsx", feature = "xlsb", feature = "xls")),
        allow(dead_code, reason = "ODS imports dates natively")
    )]
    DateAsText,
    #[cfg_attr(
        not(any(feature = "xlsx", feature = "xlsb", feature = "xls")),
        allow(dead_code, reason = "ODS import refuses formulas")
    )]
    Formula,
    Empty,
}

/// A resolved import (or evaluation) target: the zero-based worksheet
/// position and origin cell, and whether the path named one cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Origin {
    pub(crate) sheet: usize,
    pub(crate) row: u32,
    pub(crate) column: u32,
    pub(crate) cell: bool,
}

/// An inferred field value. Text and formula payloads index
/// [`Table::text`].
#[derive(Debug, Clone, PartialEq)]
pub(crate) enum Value {
    Empty,
    Text(Range<usize>),
    /// A finite number; `lexical` is a canonical decimal spelling of it.
    Number {
        value: f64,
        lexical: Range<usize>,
    },
    Bool(bool),
    /// A proleptic Gregorian date: days from 1970-01-01 and the seconds of
    /// the day (with fraction). `text` keeps the original spelling for a
    /// format that stores dates as text.
    Date {
        days: i64,
        seconds: f64,
        time: bool,
        text: Range<usize>,
    },
    /// A formula without its leading `=`.
    Formula(Range<usize>),
}

/// One field at its zero-based record and field index.
#[derive(Debug, Clone, PartialEq)]
pub(crate) struct Field {
    pub(crate) row: u32,
    pub(crate) column: u32,
    pub(crate) value: Value,
}

/// A parsed, inferred import in record order.
#[derive(Debug, Default)]
pub(crate) struct Table {
    text: String,
    fields: Vec<Field>,
    rows: u32,
    columns: u32,
}

impl Table {
    pub(crate) fn fields(&self) -> &[Field] {
        &self.fields
    }

    pub(crate) fn text(&self, range: &Range<usize>) -> &str {
        &self.text[range.clone()]
    }

    /// Whether the import writes nothing at all.
    pub(crate) fn is_empty(&self) -> bool {
        self.fields.is_empty()
    }

    /// Check that the rectangle from `(row, column)` fits a grid of
    /// `max_rows` by `max_columns`.
    pub(crate) fn check_grid(
        &self,
        row: u32,
        column: u32,
        max_rows: u32,
        max_columns: u32,
    ) -> Result<()> {
        let last_row = u64::from(row) + u64::from(self.rows.max(1)) - 1;
        let last_column = u64::from(column) + u64::from(self.columns.max(1)) - 1;
        if last_row >= u64::from(max_rows) || last_column >= u64::from(max_columns) {
            return Err(UnifiedError::new(
                crate::Code::OutOfRange,
                "the imported data does not fit the worksheet grid from its origin cell",
            )
            .with_suggestion("import into an earlier origin cell or split the source"));
        }
        Ok(())
    }
}

/// Serial number of a date in the 1900 (`false`) or 1904 (`true`) date
/// system, or `None` before the system's epoch.
#[cfg_attr(
    not(any(feature = "xlsx", feature = "xlsb")),
    allow(dead_code, reason = "XLS keeps imported dates as text")
)]
pub(crate) fn serial(days: i64, seconds: f64, date1904: bool) -> Option<f64> {
    // 1899-12-30 and 1904-01-01 as days from 1970-01-01.
    const EXCEL_1900_BASE: i64 = -25_569;
    const EXCEL_1904_BASE: i64 = -24_107;
    // 1900-01-01 and 1900-03-01 as days from 1970-01-01.
    const FIRST_1900: i64 = -25_567;
    const LEAP_BUG: i64 = -25_508;
    let whole = if date1904 {
        let serial = days - EXCEL_1904_BASE;
        if serial < 0 {
            return None;
        }
        serial
    } else if days < FIRST_1900 {
        return None;
    } else if days < LEAP_BUG {
        // Excel counts the non-existent 1900-02-29, so January and
        // February 1900 are one day lower than the true day count.
        days - EXCEL_1900_BASE - 1
    } else {
        days - EXCEL_1900_BASE
    };
    #[allow(clippy::cast_precision_loss, reason = "serials stay far below 2^52")]
    Some(whole as f64 + seconds / 86_400.0)
}

/// Decode, parse and infer `source` under `options`, charging `limits`.
pub(crate) fn parse(source: &[u8], options: &Options, limits: &SemanticLimits) -> Result<Table> {
    let quote = match options.quote {
        Some(quote) => Some(ascii(quote, OptionKind::Quote)?),
        None => None,
    };
    let explicit = match options.delimiter {
        Some(delimiter) => {
            let delimiter = ascii(delimiter, OptionKind::Delimiter)?;
            if Some(delimiter) == quote {
                return Err(Error::InvalidOption {
                    option: OptionKind::Delimiter,
                }
                .into());
            }
            Some(delimiter)
        },
        None => None,
    };
    if source.len() > limits.max_staged_bytes() {
        return Err(limit(
            LimitKind::StagedBytes,
            source.len(),
            limits.max_staged_bytes(),
        ));
    }
    let (text, _) = options.encoding.decode(source).map_err(leaf)?;
    let (declared, body) = declared_delimiter(&text);
    let delimiter = match (explicit, declared) {
        (Some(delimiter), _) => delimiter,
        (None, Some(declared)) => {
            if Some(declared) == quote {
                return Err(Error::InvalidOption {
                    option: OptionKind::Delimiter,
                }
                .into());
            }
            declared
        },
        (None, None) => b',',
    };
    let dialect = Dialect::new(delimiter).quote(quote).refuse_controls(true);
    let mut parser = Tokenizer::new(body, dialect)
        .map_err(leaf)?
        .first_line(1 + usize::from(declared.is_some()));
    let mut table = Table::default();
    let mut staged = 0usize;
    let mut record = 0u32;
    while parser.next_record() {
        let mut column = 0u32;
        loop {
            let field = parser.field().map_err(leaf)?;
            let value = infer(field, record, column, options, &mut table.text)?;
            if table.fields.len() >= limits.max_expanded_targets() {
                return Err(limit(
                    LimitKind::ExpandedTargets,
                    table.fields.len().saturating_add(1),
                    limits.max_expanded_targets(),
                ));
            }
            staged = staged
                .saturating_add(FIELD_CHARGE)
                .saturating_add(field.len());
            if staged > limits.max_staged_bytes() {
                return Err(limit(
                    LimitKind::StagedBytes,
                    staged,
                    limits.max_staged_bytes(),
                ));
            }
            table
                .fields
                .try_reserve(1)
                .map_err(|source| TransactionError::Allocation { source })?;
            table.fields.push(Field {
                row: record,
                column,
                value,
            });
            column = column.saturating_add(1);
            if !parser.next_field() {
                break;
            }
        }
        parser.end_record();
        table.columns = table.columns.max(column);
        record = record.saturating_add(1);
    }
    table.rows = record;
    // Empty fields alone write nothing on a fresh sheet but may clear
    // existing values, so they are retained; a table with no fields at all
    // (empty input, or a `sep=` line alone) is the exact no-op.
    Ok(table)
}

fn limit(kind: LimitKind, observed: usize, maximum: usize) -> UnifiedError {
    TransactionError::Limit {
        kind,
        observed,
        maximum,
    }
    .into()
}

fn ascii(character: char, option: OptionKind) -> Result<u8> {
    match u8::try_from(character) {
        Ok(byte) if byte.is_ascii() && !matches!(byte, b'\r' | b'\n' | 0) => Ok(byte),
        _ => Err(Error::InvalidOption { option }.into()),
    }
}

/// Map a tokenizer or decoder failure onto the import's own errors.
fn leaf(error: litchi_sheet_text::Error) -> UnifiedError {
    use litchi_sheet_text::Error as Leaf;
    match error {
        Leaf::Encoding { offset } => Error::Encoding { offset }.into(),
        Leaf::UnterminatedQuote { line } => Error::UnterminatedQuote { line }.into(),
        Leaf::InvalidCharacter { line } => Error::InvalidCharacter { line }.into(),
        Leaf::Allocation { source } => TransactionError::Allocation { source }.into(),
        Leaf::InvalidOption {
            option: litchi_sheet_text::OptionKind::Quote,
        } => Error::InvalidOption {
            option: OptionKind::Quote,
        }
        .into(),
        // The options were checked before tokenizing; the tokenizer reports
        // nothing else for import's dialect.
        _ => Error::InvalidOption {
            option: OptionKind::Delimiter,
        }
        .into(),
    }
}

/// Infer one field's value, copying any text it needs into `arena`.
fn infer(
    field: &str,
    record: u32,
    column: u32,
    options: &Options,
    arena: &mut String,
) -> Result<Value> {
    if field.is_empty() {
        return Ok(Value::Empty);
    }
    if field.len() > MAX_CELL_CHARS && field.chars().count() > MAX_CELL_CHARS {
        return Err(Error::FieldTooLong {
            record: record as usize + 1,
            field: column as usize + 1,
        }
        .into());
    }
    let mut store = |text: &str| -> Result<Range<usize>> {
        arena
            .try_reserve(text.len())
            .map_err(|source| TransactionError::Allocation { source })?;
        let start = arena.len();
        arena.push_str(text);
        Ok(start..arena.len())
    };
    if options.header && record == 0 {
        return Ok(Value::Text(store(field)?));
    }
    if options.formulas
        && let Some(formula) = field.strip_prefix('=')
        && !formula.is_empty()
    {
        return Ok(Value::Formula(store(formula)?));
    }
    if options.numbers
        && let Some(value) = number(field)
    {
        let lexical = if canonical_number(field) {
            store(field)?
        } else {
            store(&value.to_string())?
        };
        return Ok(Value::Number { value, lexical });
    }
    if options.dates
        && let Some((days, seconds, time)) = date(field)
    {
        return Ok(Value::Date {
            days,
            seconds,
            time,
            text: store(field)?,
        });
    }
    if options.booleans {
        if field.eq_ignore_ascii_case("true") {
            return Ok(Value::Bool(true));
        }
        if field.eq_ignore_ascii_case("false") {
            return Ok(Value::Bool(false));
        }
    }
    Ok(Value::Text(store(field)?))
}

/// The number grammar: `[+-]? (digits ('.' digits*)? | '.' digits)
/// ([eE] [+-]? digits)?`, finite, and normal or zero (a subnormal value has
/// no exact BIFF form and stays text). Negative zero reads as zero.
fn number(field: &str) -> Option<f64> {
    let bytes = field.as_bytes();
    let mut index = usize::from(matches!(bytes.first(), Some(b'+' | b'-')));
    let integer = digits(bytes, &mut index);
    let mut fraction = 0;
    if bytes.get(index) == Some(&b'.') {
        index += 1;
        fraction = digits(bytes, &mut index);
    }
    if integer == 0 && fraction == 0 {
        return None;
    }
    if matches!(bytes.get(index), Some(b'e' | b'E')) {
        index += 1;
        if matches!(bytes.get(index), Some(b'+' | b'-')) {
            index += 1;
        }
        if digits(bytes, &mut index) == 0 {
            return None;
        }
    }
    if index != bytes.len() {
        return None;
    }
    let value: f64 = field.parse().ok()?;
    if !value.is_finite() || !(value.is_normal() || value == 0.0) {
        return None;
    }
    Some(if value == 0.0 { 0.0 } else { value })
}

fn digits(bytes: &[u8], index: &mut usize) -> usize {
    let start = *index;
    while bytes.get(*index).is_some_and(u8::is_ascii_digit) {
        *index += 1;
    }
    *index - start
}

/// Whether a number field is already the canonical `xsd:double` spelling a
/// workbook stores verbatim: no `+`, no superfluous leading zero, digits on
/// both sides of a point, and not negative zero.
fn canonical_number(field: &str) -> bool {
    let unsigned = field.strip_prefix('-').unwrap_or(field);
    let mantissa = unsigned
        .split_once(['e', 'E'])
        .map_or(unsigned, |(mantissa, _)| mantissa);
    let (integer, fraction) = mantissa.split_once('.').unwrap_or((mantissa, "1"));
    if field.starts_with('+') || integer.is_empty() || fraction.is_empty() {
        return false;
    }
    if integer != "0" && integer.starts_with('0') {
        return false;
    }
    // Negative zero is not canonical.
    let zero = mantissa.bytes().all(|byte| matches!(byte, b'0' | b'.'));
    !(field.starts_with('-') && zero)
}

/// Parse an ISO 8601 date or date-time: days from 1970-01-01, seconds of
/// the day, and whether a time was present.
fn date(field: &str) -> Option<(i64, f64, bool)> {
    let bytes = field.as_bytes();
    if bytes.len() < 10 || bytes[4] != b'-' || bytes[7] != b'-' {
        return None;
    }
    let year = number_at(bytes, 0, 4)?;
    let month = number_at(bytes, 5, 2)?;
    let day = number_at(bytes, 8, 2)?;
    if !(1..=12).contains(&month) || day == 0 || day > days_in_month(year, month) {
        return None;
    }
    let days = days_from_civil(i64::from(year), month, day);
    if bytes.len() == 10 {
        return Some((days, 0.0, false));
    }
    let mut rest = &bytes[10..];
    if !matches!(rest.first(), Some(b'T' | b' ')) {
        return None;
    }
    rest = &rest[1..];
    if rest.last() == Some(&b'Z') {
        rest = &rest[..rest.len() - 1];
    }
    if rest.len() < 8 || rest[2] != b':' || rest[5] != b':' {
        return None;
    }
    let hour = number_at(rest, 0, 2)?;
    let minute = number_at(rest, 3, 2)?;
    let second = number_at(rest, 6, 2)?;
    if hour > 23 || minute > 59 || second > 59 {
        return None;
    }
    let mut seconds = f64::from(hour * 3600 + minute * 60 + second);
    let fraction = &rest[8..];
    if !fraction.is_empty() {
        let digits = fraction.strip_prefix(b".")?;
        if digits.is_empty() || digits.len() > 9 || !digits.iter().all(u8::is_ascii_digit) {
            return None;
        }
        let text = std::str::from_utf8(digits).ok()?;
        let value: f64 = text.parse().ok()?;
        #[allow(
            clippy::cast_possible_truncation,
            clippy::cast_possible_wrap,
            reason = "at most nine digits"
        )]
        let scale = 10f64.powi(digits.len() as i32);
        seconds += value / scale;
    }
    Some((days, seconds, true))
}

fn number_at(bytes: &[u8], start: usize, len: usize) -> Option<u32> {
    let slice = bytes.get(start..start + len)?;
    slice.iter().try_fold(0u32, |value, byte| {
        byte.is_ascii_digit()
            .then(|| value * 10 + u32::from(byte - b'0'))
    })
}

const fn days_in_month(year: u32, month: u32) -> u32 {
    match month {
        2 if year % 4 == 0 && (year % 100 != 0 || year % 400 == 0) => 29,
        2 => 28,
        4 | 6 | 9 | 11 => 30,
        _ => 31,
    }
}

/// Days from 1970-01-01 in the proleptic Gregorian calendar (Howard
/// Hinnant's `days_from_civil`).
const fn days_from_civil(year: i64, month: u32, day: u32) -> i64 {
    let year = if month <= 2 { year - 1 } else { year };
    let era = if year >= 0 { year } else { year - 399 } / 400;
    let year_of_era = year - era * 400;
    let month = month as i64;
    let shifted = if month > 2 { month - 3 } else { month + 9 };
    let day_of_year = (153 * shifted + 2) / 5 + day as i64 - 1;
    let day_of_era = year_of_era * 365 + year_of_era / 4 - year_of_era / 100 + day_of_year;
    era * 146_097 + day_of_era - 719_468
}
