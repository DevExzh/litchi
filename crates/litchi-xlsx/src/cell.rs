//! Exact cell states and sparse borrowed traversal.

use std::fmt;
use std::hash::{Hash, Hasher};
use std::mem::{size_of, size_of_val};
use std::ops::Deref;
use std::sync::Arc;

use chrono::{DateTime, NaiveDate, NaiveDateTime};
use litchi_sheet::{Cell as Address, Column as ColumnIndex, Rect, Row as RowIndex};

use crate::column;
use crate::error::{Result, allocation, invalid};
use crate::formula::{Formula, Kind};
use crate::layout::Defaults;
use crate::leniency::{CellDefect, Location, ToleranceReport};
use crate::merge;
use crate::row;

pub(crate) const MAX_CELL_CHARACTERS: usize = 32_767;

/// Workbook-local lineage for opaque shared-string identities.
#[derive(Debug)]
pub(crate) struct SharedStringLineage;

/// Opaque shared-string identity retained by semantic patch states.
///
/// The physical `SpreadsheetML` table index remains private. Keys are only
/// comparable inside the workbook lineage that produced them.
#[derive(Clone)]
pub struct SharedStringKey {
    raw: usize,
    lineage: Arc<SharedStringLineage>,
}

impl SharedStringKey {
    pub(crate) fn new(raw: usize, lineage: Arc<SharedStringLineage>) -> Self {
        Self { raw, lineage }
    }

    pub(crate) const fn raw(&self) -> usize {
        self.raw
    }

    pub(crate) fn rebind(&mut self, lineage: &Arc<SharedStringLineage>) {
        self.lineage = Arc::clone(lineage);
    }
}

impl PartialEq for SharedStringKey {
    fn eq(&self, other: &Self) -> bool {
        self.raw == other.raw && Arc::ptr_eq(&self.lineage, &other.lineage)
    }
}

impl Eq for SharedStringKey {}

impl Hash for SharedStringKey {
    fn hash<H: Hasher>(&self, state: &mut H) {
        self.raw.hash(state);
        Arc::as_ptr(&self.lineage).hash(state);
    }
}

impl fmt::Debug for SharedStringKey {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("SharedStringKey(..)")
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct SharedFormulaStorage {
    pub(crate) index: u32,
    pub(crate) range: Rect,
    pub(crate) reference: Box<str>,
    pub(crate) master: bool,
}

/// The stored state of one cell record.
///
/// Absence is represented by `Option<Cell>` at lookup sites, so it cannot be
/// confused with an explicitly stored empty cell.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Cell {
    /// A `<c>` record exists but has no primary payload.
    Empty,
    /// An exact non-formula value.
    Value(Value),
    /// A formula and its optional separately qualified cache.
    Formula(Formula),
    /// A cell representation not yet understood by this facade. Core inert
    /// fields are available on [`Unknown`]; the snapshot retains original part
    /// bytes for lossless future saves.
    Unknown(Unknown),
}

impl Cell {
    /// Parse a checked A1 reference into one-based `(column, row)` numbers.
    ///
    /// This small lexical helper is shared by worksheet package codecs whose
    /// references are not yet materialized as a [`litchi_sheet::Cell`].
    pub fn reference_to_coords(reference: &str) -> Result<(u32, u32)> {
        const MAX_COLUMN: u32 = 16_384;
        const MAX_ROW: u32 = 1_048_576;

        let bytes = reference.as_bytes();
        let column_end = bytes
            .iter()
            .position(u8::is_ascii_digit)
            .ok_or_else(|| invalid(format!("invalid cell reference '{reference}'")))?;
        if column_end == 0 || column_end == bytes.len() {
            return Err(invalid(format!("invalid cell reference '{reference}'")));
        }

        let mut column = 0u32;
        for byte in &bytes[..column_end] {
            if !byte.is_ascii_alphabetic() {
                return Err(invalid(format!(
                    "invalid column in cell reference '{reference}'"
                )));
            }
            let digit = u32::from(byte.to_ascii_uppercase() - b'A' + 1);
            column = column
                .checked_mul(26)
                .and_then(|value| value.checked_add(digit))
                .ok_or_else(|| {
                    invalid(format!("column overflows in cell reference '{reference}'"))
                })?;
        }
        if column > MAX_COLUMN {
            return Err(invalid(format!(
                "column exceeds Excel limits in cell reference '{reference}'"
            )));
        }

        let row = std::str::from_utf8(&bytes[column_end..])
            .ok()
            .and_then(|value| value.parse::<u32>().ok())
            .filter(|value| *value != 0 && *value <= MAX_ROW)
            .ok_or_else(|| {
                invalid(format!(
                    "row exceeds Excel limits in cell reference '{reference}'"
                ))
            })?;
        Ok((column, row))
    }
}

/// Exact semantic state at one logical worksheet coordinate.
///
/// A covered coordinate is reported before any producer-stored follower cell,
/// because the merge anchor owns its visible content. Sparse stored traversal
/// remains available through [`crate::Worksheet::cells`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum View<'a> {
    /// No cell record or covering merge exists at this coordinate.
    Missing,
    /// The coordinate is covered by a merge whose anchor is `range.start()`.
    Covered(Rect),
    /// One physical cell record is stored at this coordinate.
    Stored(&'a Cell),
}

impl<'a> View<'a> {
    /// Borrow the physical cell state when this coordinate owns one.
    #[must_use]
    pub const fn stored(self) -> Option<&'a Cell> {
        match self {
            Self::Stored(cell) => Some(cell),
            Self::Missing | Self::Covered(_) => None,
        }
    }

    /// Covering merged range, if this is a non-anchor coordinate.
    #[must_use]
    pub const fn merge(self) -> Option<Rect> {
        match self {
            Self::Covered(range) => Some(range),
            Self::Missing | Self::Stored(_) => None,
        }
    }

    /// Whether this coordinate has neither a record nor merge coverage.
    #[must_use]
    pub const fn is_missing(self) -> bool {
        matches!(self, Self::Missing)
    }
}

/// Exact value stored by `SpreadsheetML`.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Value {
    Bool(bool),
    /// Numeric lexical form, retained without format-based coercion.
    Number(Number),
    Text(Text),
    /// Checked ISO 8601 lexical form from a `t="d"` cell.
    Date(Date),
    Error(ErrorValue),
}

impl Value {
    /// Construct inert plain text.
    pub fn text(value: impl Into<Text>) -> Self {
        Self::Text(value.into())
    }

    /// Construct an explicitly typed, checked ISO 8601 date lexical value.
    pub fn date(value: impl Into<Text>) -> Result<Self> {
        Date::new(value).map(Self::Date)
    }

    pub(crate) fn validate_for_write(&self) -> Result<()> {
        let text = match self {
            Self::Error(ErrorValue::Unknown(_)) => {
                return Err(invalid(
                    "unrecognized worksheet error values cannot be authored",
                ));
            },
            Self::Text(text) => Some(text),
            Self::Date(date) => Some(&date.0),
            Self::Bool(_) | Self::Error(_) => None,
            Self::Number(number) => return number.validate_for_write(),
        };
        if text.is_some_and(|text| text.chars().count() > MAX_CELL_CHARACTERS) {
            return Err(invalid(format!(
                "cell text exceeds {MAX_CELL_CHARACTERS} characters"
            )));
        }
        Ok(())
    }

    pub(crate) fn retained_bytes(&self) -> usize {
        size_of::<Self>().saturating_add(match self {
            Self::Number(number) => number.as_str().len(),
            Self::Text(text) | Self::Date(Date(text)) => text.as_str().len(),
            Self::Error(error) => error.as_str().len(),
            Self::Bool(_) => 0,
        })
    }
}

/// A checked ISO 8601 lexical value for a `SpreadsheetML` date cell.
///
/// The original lexical form is retained so a read/write cycle does not
/// silently normalize producer data.
#[derive(Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Date(Text);

impl Date {
    /// Validate and retain an ISO 8601 date or date-time lexical form.
    pub fn new(value: impl Into<Text>) -> Result<Self> {
        let value = value.into();
        let lexical = value.as_str();
        let valid = NaiveDate::parse_from_str(lexical, "%Y-%m-%d").is_ok()
            || NaiveDateTime::parse_from_str(lexical, "%Y-%m-%dT%H:%M:%S%.f").is_ok()
            || DateTime::parse_from_rfc3339(lexical).is_ok();
        if !valid {
            return Err(invalid(format!(
                "invalid ISO 8601 worksheet date '{lexical}'"
            )));
        }
        Ok(Self(value))
    }

    /// Exact stored lexical form.
    #[must_use]
    pub fn as_str(&self) -> &str {
        self.0.as_str()
    }
}

impl Deref for Date {
    type Target = str;

    fn deref(&self) -> &Self::Target {
        self.as_str()
    }
}

impl AsRef<str> for Date {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl fmt::Debug for Date {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_tuple("Date").field(&self.0).finish()
    }
}

impl fmt::Display for Date {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl From<bool> for Value {
    fn from(value: bool) -> Self {
        Self::Bool(value)
    }
}

impl From<Number> for Value {
    fn from(value: Number) -> Self {
        Self::Number(value)
    }
}

impl From<Text> for Value {
    fn from(value: Text) -> Self {
        Self::Text(value)
    }
}

impl From<String> for Value {
    fn from(value: String) -> Self {
        Self::Text(value.into())
    }
}

impl From<&str> for Value {
    fn from(value: &str) -> Self {
        Self::Text(value.into())
    }
}

macro_rules! exact_integer_value {
    ($($integer:ty),+ $(,)?) => {
        $(
            impl From<$integer> for Value {
                fn from(value: $integer) -> Self {
                    Self::Number(Number(value.to_string().into_boxed_str()))
                }
            }
        )+
    };
}

exact_integer_value!(i8, i16, i32, u8, u16, u32);

/// Primary payload accepted by [`crate::workbook::edit::Sheet::set`].
///
/// Plain strings are always inert text. Formula interpretation requires an
/// explicit checked [`Formula`].
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Content {
    Value(Value),
    Formula(Formula),
}

impl Content {
    pub(crate) fn validate_for_write(&self) -> Result<()> {
        match self {
            Self::Value(value) => value.validate_for_write(),
            Self::Formula(formula)
                if matches!(formula.kind(), Kind::Scalar) && formula.cached().is_none() =>
            {
                Ok(())
            },
            Self::Formula(formula) if formula.cached().is_some() => Err(invalid(
                "writing a stored formula cache requires an explicit cache policy",
            )),
            Self::Formula(_) => Err(invalid(
                "array and data-table formulas require a range-scoped editor",
            )),
        }
    }

    /// Whether `cell` is exactly `self.as_cell()`, without building it.
    pub(crate) fn same_as_cell(&self, cell: &Cell) -> bool {
        match (self, cell) {
            (Self::Value(value), Cell::Value(other)) => value == other,
            (Self::Formula(formula), Cell::Formula(other)) => formula == other,
            _ => false,
        }
    }

    pub(crate) fn as_cell(&self) -> Cell {
        match self {
            Self::Value(value) => Cell::Value(value.clone()),
            Self::Formula(formula) => Cell::Formula(formula.clone()),
        }
    }

    pub(crate) fn retained_bytes(&self) -> usize {
        size_of::<Self>().saturating_add(match self {
            Self::Value(value) => value.retained_bytes(),
            Self::Formula(formula) => formula.text().len().saturating_add(
                formula
                    .cached()
                    .map_or(0, |cache| cache.value().retained_bytes()),
            ),
        })
    }
}

impl From<Formula> for Content {
    fn from(value: Formula) -> Self {
        Self::Formula(value)
    }
}

impl<T> From<T> for Content
where
    Value: From<T>,
{
    fn from(value: T) -> Self {
        Self::Value(Value::from(value))
    }
}

/// An exact `SpreadsheetML` number.
#[derive(Clone, PartialEq, Eq, Hash)]
pub struct Number(Box<str>);

impl Number {
    /// Validate and retain a numeric lexical form without normalizing it.
    pub fn new(value: impl Into<Box<str>>) -> Result<Self> {
        let value = value.into();
        Self::validate_lexical(&value)?;
        Ok(Self(value))
    }

    /// Validate a worksheet number without taking ownership of its lexical
    /// form. This is used by streaming readers that must validate an
    /// unselected value while retaining ownership only for selected cells.
    pub(crate) fn validate_lexical(value: &str) -> Result<()> {
        if is_short_integer(value) {
            // Exactly the lexemes `f64::from_str` parses to a finite value
            // without rounding into the infinities, so the full float parse
            // below would accept them too.
            return Ok(());
        }
        let parsed = value
            .trim()
            .parse::<f64>()
            .map_err(|_source| invalid(format!("invalid worksheet number '{value}'")))?;
        if !parsed.is_finite() {
            return Err(invalid(format!("non-finite worksheet number '{value}'")));
        }
        Ok(())
    }

    /// Exact stored lexical form.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Convert the value to IEEE-754 binary64.
    ///
    /// `None` keeps this accessor safe if a future lossless reader accepts a
    /// numeric lexical form outside Rust's binary64 parser.
    #[must_use]
    pub fn as_f64(&self) -> Option<f64> {
        self.0
            .trim()
            .parse()
            .ok()
            .filter(|value: &f64| value.is_finite())
    }

    pub(crate) fn validate_for_write(&self) -> Result<()> {
        if self.0.chars().count() > MAX_CELL_CHARACTERS {
            return Err(invalid(format!(
                "cell number exceeds {MAX_CELL_CHARACTERS} characters"
            )));
        }
        Ok(())
    }
}

/// Whether `value` is an optionally signed run of at most 18 ASCII digits.
///
/// Every such lexeme is a finite binary64 number, so it passes the number
/// validation without the general float parser.
fn is_short_integer(value: &str) -> bool {
    let digits = value.strip_prefix(['-', '+']).unwrap_or(value).as_bytes();
    !digits.is_empty() && digits.len() <= 18 && digits.iter().all(u8::is_ascii_digit)
}

impl TryFrom<f64> for Number {
    type Error = crate::Error;

    fn try_from(value: f64) -> Result<Self> {
        if !value.is_finite() {
            return Err(invalid("spreadsheet numbers must be finite"));
        }
        Self::new(value.to_string())
    }
}

impl TryFrom<f32> for Number {
    type Error = crate::Error;

    fn try_from(value: f32) -> Result<Self> {
        Self::try_from(f64::from(value))
    }
}

impl fmt::Debug for Number {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_tuple("Number").field(&self.0).finish()
    }
}

impl fmt::Display for Number {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}

/// The rich-text runs of one cell (record 0295): each run's text and its
/// run properties (`rPr`) exactly as the source item spells them, so a run
/// text can change while every run keeps its formatting.
///
/// Read with [`crate::Worksheet::rich_text`], changed with
/// [`RichText::set_run_text`], written with
/// [`crate::workbook::edit::Sheet::set_rich_text`] as the cell's own inline rich
/// string.
///
/// ```
/// # fn main() -> litchi_xlsx::Result<()> {
/// use litchi_xlsx::{Package, RichText};
///
/// let mut edit = Package::new()?.edit()?;
/// let mut rich = RichText::plain("Hello");
/// rich.set_run_text(0, "Bye")?;
/// edit.sheet("Sheet1")?.expect("worksheet").set_rich_text("A1", rich)?;
/// let workbook = edit.commit()?.into_package().workbook();
/// let sheet = workbook.sheet("Sheet1")?.expect("worksheet");
/// assert!(sheet.has_rich_text("A1")?);
/// # Ok(())
/// # }
/// ```
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RichText {
    runs: Vec<RichRun>,
}

/// One run of a [`RichText`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RichRun {
    /// The `rPr` children as `(local name, raw attributes)`; `None` for a
    /// run without `rPr`.
    pub(crate) properties: Option<Box<[(Box<str>, Box<str>)]>>,
    text: String,
}

impl RichText {
    /// One unformatted run holding `text`.
    #[must_use]
    pub fn plain(text: impl Into<String>) -> Self {
        Self {
            runs: vec![RichRun {
                properties: None,
                text: text.into(),
            }],
        }
    }

    /// The runs `runs`, in order (record 0338). Writing needs at least one.
    ///
    /// ```
    /// use litchi_xlsx::{FontChange, RichRun, RichText};
    ///
    /// let mut bold = RichRun::new("Bold");
    /// bold.apply_font(&FontChange::new().bold(Some(true)))?;
    /// let rich = RichText::new(vec![bold, RichRun::new(" plain")]);
    /// assert_eq!(rich.text(), "Bold plain");
    /// assert!(rich.runs()[0].font()?.bold);
    /// # Ok::<(), litchi_xlsx::Error>(())
    /// ```
    #[must_use]
    pub fn new(runs: Vec<RichRun>) -> Self {
        Self { runs }
    }

    pub(crate) fn from_runs(runs: Vec<RichRun>) -> Self {
        Self { runs }
    }

    /// The runs, in order.
    #[must_use]
    pub fn runs(&self) -> &[RichRun] {
        &self.runs
    }

    /// The run at `index`, to change it in place (record 0338).
    pub fn run_mut(&mut self, index: usize) -> Option<&mut RichRun> {
        self.runs.get_mut(index)
    }

    /// Insert `run` before position `index` (`runs().len()` appends).
    ///
    /// # Errors
    ///
    /// An invalid-input error for an index past the end.
    pub fn insert_run(&mut self, index: usize, run: RichRun) -> Result<()> {
        if index > self.runs.len() {
            return Err(invalid("rich-text run index is out of range"));
        }
        self.runs
            .try_reserve(1)
            .map_err(|source| allocation("rich-text runs", source))?;
        self.runs.insert(index, run);
        Ok(())
    }

    /// Remove and return the run at `index`.
    ///
    /// # Errors
    ///
    /// An invalid-input error for an index past the last run.
    pub fn remove_run(&mut self, index: usize) -> Result<RichRun> {
        if index >= self.runs.len() {
            return Err(invalid("rich-text run index is out of range"));
        }
        Ok(self.runs.remove(index))
    }

    /// The visible text: every run's text, joined.
    #[must_use]
    pub fn text(&self) -> String {
        self.runs.iter().map(|run| run.text.as_str()).collect()
    }

    /// Replace the text of run `index`, keeping its properties.
    ///
    /// # Errors
    ///
    /// An invalid-input error for an index past the last run.
    pub fn set_run_text(&mut self, index: usize, text: impl Into<String>) -> Result<()> {
        let run = self
            .runs
            .get_mut(index)
            .ok_or_else(|| invalid("rich-text run index is out of range"))?;
        run.text = text.into();
        Ok(())
    }

    pub(crate) fn validate_for_write(&self) -> Result<()> {
        if self.runs.is_empty() {
            return Err(invalid("rich text needs at least one run"));
        }
        let count: usize = self.runs.iter().map(|run| run.text.chars().count()).sum();
        if count > MAX_CELL_CHARACTERS {
            return Err(invalid(format!(
                "cell text exceeds {MAX_CELL_CHARACTERS} characters"
            )));
        }
        Ok(())
    }

    /// Conservative logical bytes retained by this authored rich string.
    ///
    /// The transaction capture budget accounts for the run text and the
    /// lexical run properties that are copied into the pending cell action;
    /// allocator capacity is deliberately not folded into the package/cache
    /// census.
    #[doc(hidden)]
    pub fn retained_bytes(&self) -> usize {
        self.runs.iter().fold(size_of::<Self>(), |bytes, run| {
            let properties = run.properties.as_deref().map_or(0, |items| {
                items
                    .iter()
                    .fold(size_of_val(items), |bytes, (name, value)| {
                        bytes.saturating_add(name.len()).saturating_add(value.len())
                    })
            });
            bytes
                .saturating_add(size_of::<RichRun>())
                .saturating_add(run.text.len())
                .saturating_add(properties)
        })
    }
}

impl RichRun {
    /// An unformatted run holding `text` (record 0338).
    #[must_use]
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            properties: None,
            text: text.into(),
        }
    }

    pub(crate) fn with_properties(
        properties: Option<Box<[(Box<str>, Box<str>)]>>,
        text: String,
    ) -> Self {
        Self { properties, text }
    }

    /// The run's text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Replace the run's text, keeping its properties (record 0338).
    pub fn set_text(&mut self, text: impl Into<String>) {
        self.text = text.into();
    }

    /// Whether the run carries run properties (`rPr`).
    #[must_use]
    pub fn is_formatted(&self) -> bool {
        self.properties.is_some()
    }

    /// The run's typed font, read from its run properties (record 0338):
    /// `rFont` is the font name, and an absent property is the default (no
    /// flag, no value). Properties the [`Font`] model has no field for
    /// (`outline`, `shadow`, `condense`, `extend`) are not read.
    ///
    /// [`Font`]: crate::style::stylesheet::font::Font
    ///
    /// # Errors
    ///
    /// An invalid-input error for a property value that is not its
    /// `SpreadsheetML` lexical form.
    pub fn font(&self) -> Result<crate::style::stylesheet::font::Font> {
        rich_font::read(self.properties.as_deref().unwrap_or_default())
    }

    /// Apply `change` to the run's properties (record 0338): each property
    /// the change mentions is replaced (`Some`) or removed (`None`); the
    /// others, and properties the change has no setter for, keep their
    /// markup. A new property is placed in Excel's run-property order.
    ///
    /// # Errors
    ///
    /// An invalid-input error for a value the change cannot spell.
    pub fn apply_font(&mut self, change: &crate::style::FontChange) -> Result<()> {
        let mut properties = self.properties.take().map(Vec::from).unwrap_or_default();
        let result = rich_font::apply(&mut properties, change);
        self.properties = (!properties.is_empty()).then(|| properties.into_boxed_slice());
        result
    }
}

/// Typed reads and writes of a rich-text run's `rPr` children (record
/// 0338), kept as `(local name, raw attributes)` pairs.
mod rich_font {
    use litchi_core::xml::escape_xml;
    use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
    use quick_xml::events::Event;

    use crate::error::{Result, invalid};
    use crate::style::FontChange;
    use crate::style::stylesheet::border::Tint;
    use crate::style::stylesheet::font::{
        Font, FontColor, FontColorKind, Scheme, Script, Underline,
    };

    /// Excel's order of run properties, used to place a new one.
    const ORDER: [&str; 15] = [
        "b",
        "i",
        "strike",
        "condense",
        "extend",
        "outline",
        "shadow",
        "u",
        "vertAlign",
        "sz",
        "color",
        "rFont",
        "family",
        "charset",
        "scheme",
    ];

    type Property = (Box<str>, Box<str>);

    /// The attributes of one property, unescaped.
    fn attributes(raw: &str) -> Result<Vec<(String, String)>> {
        let xml = format!("<x {raw}/>");
        let mut reader = quick_xml::Reader::from_str(&xml);
        let Ok(Event::Empty(element)) = reader.read_event() else {
            return Err(invalid("malformed rich-text run property attributes"));
        };
        let mut out = Vec::new();
        for attribute in element.checked_attributes() {
            let attribute = attribute
                .map_err(|error| invalid(format!("rich-text run property attribute: {error}")))?;
            let name = String::from_utf8_lossy(attribute.key.local_name().as_ref()).into_owned();
            let value = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| invalid(format!("rich-text run property value: {error}")))?
                .into_owned();
            out.push((name, value));
        }
        Ok(out)
    }

    fn value<'a>(attributes: &'a [(String, String)], name: &str) -> Option<&'a str> {
        attributes
            .iter()
            .find(|(key, _)| key == name)
            .map(|(_, value)| value.as_str())
    }

    fn flag(attributes: &[(String, String)], name: &str) -> Result<bool> {
        match value(attributes, "val") {
            None | Some("1" | "true") => Ok(true),
            Some("0" | "false") => Ok(false),
            Some(other) => Err(invalid(format!(
                "invalid run property {name} value '{other}'"
            ))),
        }
    }

    fn number(attributes: &[(String, String)], name: &str) -> Result<Option<u32>> {
        value(attributes, "val")
            .map(|text| {
                text.parse()
                    .map_err(|_| invalid(format!("invalid run property {name} value '{text}'")))
            })
            .transpose()
    }

    pub(super) fn read(properties: &[Property]) -> Result<Font> {
        let mut font = Font::new();
        for (name, raw) in properties {
            let attributes = attributes(raw)?;
            match &**name {
                "b" => font.bold = flag(&attributes, name)?,
                "i" => font.italic = flag(&attributes, name)?,
                "strike" => font.strike = flag(&attributes, name)?,
                "u" => {
                    let token = value(&attributes, "val").unwrap_or("single");
                    font.underline = Some(
                        token
                            .parse::<Underline>()
                            .map_err(|error| invalid(error.to_string()))?,
                    );
                },
                "vertAlign" => {
                    font.script = match value(&attributes, "val") {
                        Some("baseline") => Some(Script::Baseline),
                        Some("superscript") => Some(Script::Superscript),
                        Some("subscript") => Some(Script::Subscript),
                        other => {
                            return Err(invalid(format!(
                                "invalid run vertAlign value '{}'",
                                other.unwrap_or_default()
                            )));
                        },
                    };
                },
                "sz" => {
                    let text = value(&attributes, "val").unwrap_or_default();
                    let size: f64 = text
                        .parse()
                        .map_err(|_| invalid(format!("invalid run font size '{text}'")))?;
                    if !size.is_finite() || size <= 0.0 {
                        return Err(invalid(format!("invalid run font size '{text}'")));
                    }
                    font.size = Some(size);
                },
                "color" => font.color = Some(color(&attributes)?),
                "rFont" => font.name = value(&attributes, "val").map(str::to_owned),
                "family" => font.family = number(&attributes, name)?,
                "charset" => font.charset = number(&attributes, name)?,
                "scheme" => {
                    font.scheme = value(&attributes, "val")
                        .map(|token| {
                            token
                                .parse::<Scheme>()
                                .map_err(|error| invalid(error.to_string()))
                        })
                        .transpose()?;
                },
                _ => {},
            }
        }
        Ok(font)
    }

    fn color(attributes: &[(String, String)]) -> Result<FontColor> {
        let tint = value(attributes, "tint")
            .map(|text| {
                text.parse::<f64>()
                    .ok()
                    .and_then(|tint| Tint::new(tint).ok())
                    .ok_or_else(|| invalid(format!("invalid run color tint '{text}'")))
            })
            .transpose()?;
        let mut color = if let Some(rgb) = value(attributes, "rgb") {
            FontColor::rgb(
                rgb.parse()
                    .map_err(|error| invalid(format!("invalid run RGB color '{rgb}': {error}")))?,
            )
        } else if let Some(theme) = value(attributes, "theme") {
            FontColor::theme(
                theme
                    .parse()
                    .map_err(|_| invalid(format!("invalid run theme color '{theme}'")))?,
            )
        } else if let Some(indexed) = value(attributes, "indexed") {
            FontColor::indexed(
                indexed
                    .parse()
                    .map_err(|_| invalid(format!("invalid run indexed color '{indexed}'")))?,
            )
        } else if matches!(value(attributes, "auto"), Some("1" | "true")) {
            FontColor::automatic()
        } else {
            FontColor::default_base()
        };
        if let Some(tint) = tint {
            color = color.with_tint(tint);
        }
        Ok(color)
    }

    fn val(text: &str) -> Box<str> {
        format!("val=\"{}\"", escape_xml(text)).into_boxed_str()
    }

    fn flag_value(value: bool) -> Box<str> {
        if value {
            Box::default()
        } else {
            "val=\"0\"".into()
        }
    }

    fn color_value(color: &FontColor) -> Result<Box<str>> {
        let mut xml = String::new();
        crate::style::stylesheet::writer::write_font_color(&mut xml, "color", color)?;
        let inner = xml
            .strip_prefix("<color")
            .and_then(|rest| rest.strip_suffix("/>"))
            .unwrap_or_default()
            .trim();
        if matches!(color.kind(), FontColorKind::Default) && inner.is_empty() {
            return Ok(Box::default());
        }
        Ok(inner.into())
    }

    /// Replace (`Some`) or remove (`None`) the property `name`.
    fn put(properties: &mut Vec<Property>, name: &str, value: Option<Box<str>>) {
        let existing = properties.iter().position(|(held, _)| &**held == name);
        properties.retain(|(held, _)| &**held != name);
        let Some(value) = value else {
            return;
        };
        let at = existing.unwrap_or_else(|| {
            let rank = ORDER
                .iter()
                .position(|item| *item == name)
                .unwrap_or(ORDER.len());
            properties
                .iter()
                .position(|(held, _)| {
                    ORDER
                        .iter()
                        .position(|item| *item == &**held)
                        .is_some_and(|other| other > rank)
                })
                .unwrap_or(properties.len())
        });
        let at = at.min(properties.len());
        properties.insert(at, (name.into(), value));
    }

    pub(super) fn apply(properties: &mut Vec<Property>, change: &FontChange) -> Result<()> {
        for (name, value) in [
            ("b", change.bold_change()),
            ("i", change.italic_change()),
            ("strike", change.strike_change()),
            ("condense", change.condense_change()),
            ("extend", change.extend_change()),
            ("outline", change.outline_change()),
            ("shadow", change.shadow_change()),
        ] {
            if let Some(value) = value {
                put(properties, name, value.map(flag_value));
            }
        }
        if let Some(value) = change.underline_change() {
            put(
                properties,
                "u",
                value.map(|underline| match underline {
                    Underline::Single => Box::default(),
                    other => val(other.as_str()),
                }),
            );
        }
        if let Some(value) = change.script_change() {
            put(
                properties,
                "vertAlign",
                value.map(|script| val(script.as_str())),
            );
        }
        if let Some(value) = change.size_change() {
            put(properties, "sz", value.map(|size| val(&size.to_string())));
        }
        if let Some(value) = change.color_change() {
            let value = value.map(color_value).transpose()?;
            put(properties, "color", value);
        }
        if let Some(value) = change.name_change() {
            put(properties, "rFont", value.map(val));
        }
        if let Some(value) = change.family_change() {
            put(
                properties,
                "family",
                value.map(|family| val(&family.to_string())),
            );
        }
        if let Some(value) = change.charset_change() {
            put(
                properties,
                "charset",
                value.map(|charset| val(&charset.to_string())),
            );
        }
        if let Some(value) = change.scheme_change() {
            put(
                properties,
                "scheme",
                value.map(|scheme| val(scheme.as_str())),
            );
        }
        Ok(())
    }
}

/// Cheaply cloned immutable text, including resolved shared strings.
#[derive(Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Text(Arc<str>);

impl Text {
    /// Move or borrow text into an immutable value.
    pub fn new(value: impl Into<Arc<str>>) -> Self {
        Self(value.into())
    }

    /// Borrow the text.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl Deref for Text {
    type Target = str;

    fn deref(&self) -> &Self::Target {
        self.as_str()
    }
}

impl AsRef<str> for Text {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl From<String> for Text {
    fn from(value: String) -> Self {
        Self(value.into())
    }
}

impl From<&str> for Text {
    fn from(value: &str) -> Self {
        Self(value.into())
    }
}

impl fmt::Debug for Text {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.0.fmt(formatter)
    }
}

impl fmt::Display for Text {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}

/// Spreadsheet error value.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ErrorValue {
    Null,
    DivZero,
    Value,
    Ref,
    Name,
    Num,
    NotAvailable,
    GettingData,
    Spill,
    Calc,
    Field,
    Blocked,
    Connect,
    Unknown(Text),
}

impl ErrorValue {
    pub(crate) fn parse(value: &str) -> Self {
        match value.trim() {
            "#NULL!" => Self::Null,
            "#DIV/0!" => Self::DivZero,
            "#VALUE!" => Self::Value,
            "#REF!" => Self::Ref,
            "#NAME?" => Self::Name,
            "#NUM!" => Self::Num,
            "#N/A" => Self::NotAvailable,
            "#GETTING_DATA" => Self::GettingData,
            "#SPILL!" => Self::Spill,
            "#CALC!" => Self::Calc,
            "#FIELD!" => Self::Field,
            "#BLOCKED!" => Self::Blocked,
            "#CONNECT!" => Self::Connect,
            other => Self::Unknown(other.into()),
        }
    }

    /// Spreadsheet lexical form.
    #[must_use]
    pub fn as_str(&self) -> &str {
        match self {
            Self::Null => "#NULL!",
            Self::DivZero => "#DIV/0!",
            Self::Value => "#VALUE!",
            Self::Ref => "#REF!",
            Self::Name => "#NAME?",
            Self::Num => "#NUM!",
            Self::NotAvailable => "#N/A",
            Self::GettingData => "#GETTING_DATA",
            Self::Spill => "#SPILL!",
            Self::Calc => "#CALC!",
            Self::Field => "#FIELD!",
            Self::Blocked => "#BLOCKED!",
            Self::Connect => "#CONNECT!",
            Self::Unknown(value) => value,
        }
    }
}

/// Bounded diagnostic for a cell encoding not yet modeled semantically.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Unknown {
    kind: Text,
    value: Option<Text>,
    formula: Option<Text>,
}

impl Unknown {
    pub(crate) fn new(
        kind: impl Into<Text>,
        value: Option<String>,
        formula: Option<String>,
    ) -> Self {
        Self {
            kind: kind.into(),
            value: value.map(Text::from),
            formula: formula.map(Text::from),
        }
    }

    /// Producer cell type or formula form that was not recognized.
    #[must_use]
    pub fn kind(&self) -> &str {
        &self.kind
    }

    /// Uninterpreted value text, when present.
    #[must_use]
    pub fn value(&self) -> Option<&str> {
        self.value.as_deref()
    }

    /// Uninterpreted formula text, when present.
    #[must_use]
    pub fn formula(&self) -> Option<&str> {
        self.formula.as_deref()
    }
}

#[derive(Debug, Clone)]
pub(crate) struct Stored {
    pub(crate) address: Address,
    pub(crate) cell: Cell,
    // Retained for the shared-style facade. Native indexes never escape this
    // migration boundary.
    pub(crate) style: Option<u32>,
    // Retained so same-workbook transfers can preserve the exact shared-string
    // item, including rich-text runs, without exposing its physical index.
    pub(crate) shared_string: Option<usize>,
    // Inline strings are flattened into `Cell::Value(Text)` for the ordinary
    // facade. Retain whether the source carried rich runs so dependency-aware
    // transfers can refuse formatting-bearing inline text as a plain scalar.
    pub(crate) inline_rich: bool,
    // Array/data-table formula ranges and expanded shared-formula ownership
    // are retained even when the ordinary formula facade exposes a scalar
    // member. Range-aware transfers use this provenance to refuse copying
    // through an owner whose anchor is outside the selected rectangle.
    pub(crate) formula_range: Option<Rect>,
    pub(crate) shared_formula: Option<SharedFormulaStorage>,
    #[allow(
        dead_code,
        reason = "the cached column supports internal sparse-cell indexing"
    )]
    pub(crate) cell_metadata: Option<u32>,
    #[allow(
        dead_code,
        reason = "the cached row supports internal sparse-cell indexing"
    )]
    pub(crate) value_metadata: Option<u32>,
}

#[derive(Debug, Default)]
pub(crate) struct Store {
    cells: Box<[Stored]>,
    cell_rows: Box<[CellRowStart]>,
    rows: Box<[row::Stored]>,
    columns: Box<[column::Stored]>,
    defaults: Option<Defaults>,
    merges: merge::Index,
    extents: Extents,
    /// Repairs made while this store was built; see `crate::leniency`.
    tolerance: ToleranceReport,
}

#[derive(Debug)]
struct CellRowStart {
    row: RowIndex,
    start: usize,
}

/// Distinct worksheet cell-bound summaries.
///
/// Except for the producer-declared hint, these ranges describe stored cell
/// records only. Row/column defaults, drawings, merges, and other sheet objects
/// are intentionally not folded into the semantic cell extents.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Extents {
    declared: Option<Rect>,
    stored: Option<Rect>,
    content: Option<Rect>,
    styled: Option<Rect>,
}

impl Extents {
    /// Producer-declared worksheet `dimension`, when present.
    #[must_use]
    pub const fn declared(&self) -> Option<Rect> {
        self.declared
    }

    /// Bounds of every explicit cell record, including empty metadata cells.
    #[must_use]
    pub const fn stored(&self) -> Option<Rect> {
        self.stored
    }

    /// Bounds of cells with a value, formula, or unknown primary payload.
    #[must_use]
    pub const fn content(&self) -> Option<Rect> {
        self.content
    }

    /// Bounds of cells with an explicit local shared-style reference.
    #[must_use]
    pub const fn styled(&self) -> Option<Rect> {
        self.styled
    }

    /// Bounds of cells with content or direct local formatting.
    ///
    /// This does not include formatting inherited from row/column defaults.
    #[must_use]
    pub const fn used(&self) -> Option<Rect> {
        match (self.content, self.styled) {
            (Some(content), Some(styled)) => Some(content.union(styled)),
            (Some(content), None) => Some(content),
            (None, Some(styled)) => Some(styled),
            (None, None) => None,
        }
    }
}

impl Store {
    pub(crate) const fn stored_cell_count(&self) -> usize {
        self.cells.len()
    }

    pub(crate) fn from_unsorted(
        mut cells: Vec<Stored>,
        mut rows: Vec<row::Stored>,
        columns: Box<[column::Stored]>,
        defaults: Option<Defaults>,
        merges: Vec<Rect>,
        declared: Option<Rect>,
    ) -> Result<Self> {
        cells.sort_unstable_by_key(|entry| entry.address);
        if let Some(pair) = cells
            .windows(2)
            .find(|pair| pair[0].address == pair[1].address)
        {
            return Err(invalid(format!(
                "duplicate worksheet cell at {:?}",
                pair[0].address
            )));
        }

        rows.sort_unstable_by_key(|entry| entry.index);
        if let Some(pair) = rows.windows(2).find(|pair| pair[0].index == pair[1].index) {
            return Err(invalid(format!(
                "duplicate worksheet row {}",
                pair[0].index.get() + 1
            )));
        }

        let mut stored = Bounds::default();
        let mut content = Bounds::default();
        let mut styled = Bounds::default();
        let mut cell_rows = Vec::new();
        for (index, entry) in cells.iter().enumerate() {
            if index == 0 || cells[index - 1].address.row() != entry.address.row() {
                cell_rows
                    .try_reserve(1)
                    .map_err(|source| allocation("worksheet cell-row index", source))?;
                cell_rows.push(CellRowStart {
                    row: entry.address.row(),
                    start: index,
                });
            }
            stored.push(entry.address);
            if !matches!(entry.cell, Cell::Empty) {
                content.push(entry.address);
            }
            if entry.style.is_some() {
                styled.push(entry.address);
            }
        }
        let merges = merge::Index::new(merges)?;
        Ok(Self {
            cells: cells.into_boxed_slice(),
            cell_rows: cell_rows.into_boxed_slice(),
            rows: rows.into_boxed_slice(),
            columns,
            defaults,
            merges,
            extents: Extents {
                declared,
                stored: stored.finish()?,
                content: content.finish()?,
                styled: styled.finish()?,
            },
            tolerance: ToleranceReport::default(),
        })
    }

    /// Repairs made while this store was built, in discovery order.
    pub(crate) const fn tolerance_report(&self) -> &ToleranceReport {
        &self.tolerance
    }

    /// Attach the repairs the parse that built this store made.
    pub(crate) fn set_tolerance_report(&mut self, report: ToleranceReport) {
        self.tolerance = report;
    }

    /// Whether any cell, row, or column record names a shared style.
    pub(crate) fn references_styles(&self) -> bool {
        self.cells.iter().any(|entry| entry.style.is_some())
            || self
                .rows
                .iter()
                .any(|entry| entry.properties.style.is_some())
            || self
                .columns
                .iter()
                .any(|entry| entry.properties.style.is_some())
    }

    /// Repair every cell, row, and column style index at or beyond the
    /// workbook's `cellXfs` count (audit finding 4.2) and record each repair.
    ///
    /// Excel substitutes the default format for a stale index. The stale
    /// index becomes the mandatory default format, `cellXfs` entry 0, so the
    /// record stays explicitly styled; when the workbook has no cell formats
    /// at all, the index is dropped and the styled extent is recomputed.
    pub(crate) fn repair_style_indexes(&mut self, len: u32) -> Result<()> {
        let substitute = (len > 0).then_some(0);
        let mut dropped = false;
        for entry in &mut self.cells {
            if entry.style.is_some_and(|key| key >= len) {
                entry.style = substitute;
                dropped |= substitute.is_none();
                self.tolerance.record(
                    CellDefect::StyleIndexOutOfRange,
                    Location::Cell(entry.address),
                );
            }
        }
        for entry in &mut self.rows {
            if entry.properties.style.is_some_and(|key| key >= len) {
                entry.properties.style = substitute;
                self.tolerance
                    .record(CellDefect::StyleIndexOutOfRange, Location::Row(entry.index));
            }
        }
        for entry in &mut self.columns {
            if entry.properties.style.is_some_and(|key| key >= len) {
                entry.properties.style = substitute;
                self.tolerance.record(
                    CellDefect::StyleIndexOutOfRange,
                    Location::Columns {
                        first: entry.first,
                        last: entry.last,
                    },
                );
            }
        }
        if dropped {
            let mut styled = Bounds::default();
            for entry in self.cells.iter().filter(|entry| entry.style.is_some()) {
                styled.push(entry.address);
            }
            self.extents.styled = styled.finish()?;
        }
        Ok(())
    }

    /// Merge source cell records named by reduced-readback provenance into a
    /// parsed candidate. The parsed store remains authoritative for worksheet
    /// structure; indexes and extents are rebuilt for the combined cells.
    pub(crate) fn merge_omitted_cells(
        parsed: Self,
        source: &Self,
        omitted: &[Rect],
    ) -> Result<Option<Self>> {
        if omitted.is_empty() || !omitted_ranges_are_ordered(omitted) {
            return Ok(None);
        }
        let source_count = omitted_entries(&source.cells, omitted).count();
        if source_count == 0 || omitted_entries(&parsed.cells, omitted).next().is_some() {
            return Ok(None);
        }
        let total = parsed
            .cells
            .len()
            .checked_add(source_count)
            .ok_or_else(|| invalid("merged worksheet cell count overflows usize"))?;
        let mut cells = Vec::new();
        cells
            .try_reserve_exact(total)
            .map_err(|source| allocation("merged worksheet cells", source))?;
        let Store {
            cells: parsed_cells,
            rows,
            columns,
            defaults,
            merges,
            extents,
            tolerance,
            ..
        } = parsed;
        cells.extend(parsed_cells.into_vec());
        for entry in omitted_entries(&source.cells, omitted) {
            cells.push(entry.clone());
        }
        let mut merge_ranges = Vec::new();
        merge_ranges
            .try_reserve_exact(merges.as_slice().len())
            .map_err(|source| allocation("merged worksheet ranges", source))?;
        merge_ranges.extend_from_slice(merges.as_slice());
        Self::from_unsorted(
            cells,
            rows.into_vec(),
            columns,
            defaults,
            merge_ranges,
            extents.declared,
        )
        .map(|mut merged| {
            // The omitted records were copied from `source` as already
            // repaired, so their repairs travel with them.
            let mut tolerance = tolerance;
            let mut carried = ToleranceReport::default();
            for tolerated in source.tolerance.iter() {
                if let Location::Cell(address) = tolerated.location
                    && omitted.iter().any(|range| range.contains(address))
                {
                    carried.record(tolerated.defect, tolerated.location.clone());
                }
            }
            tolerance.append(&carried);
            merged.tolerance = tolerance;
            Some(merged)
        })
    }

    /// Whether `omitted` are ordered single-row ranges none of which holds a
    /// cell of this store.
    ///
    /// This is the collision refusal of [`Self::merge_omitted_cells`] for a
    /// readback that only verifies: a parsed record inside a range whose
    /// records were omitted would be a second record at an omitted address,
    /// which the complete parse refuses as a duplicate.
    pub(crate) fn avoids_omitted_cells(&self, omitted: &[Rect]) -> bool {
        omitted_ranges_are_ordered(omitted)
            && omitted_entries(&self.cells, omitted).next().is_none()
    }

    pub(crate) fn view(&self, address: Address) -> View<'_> {
        if let Some(range) = self.merges.containing(address)
            && range.start() != address
        {
            return View::Covered(range);
        }
        self.entry(address)
            .map_or(View::Missing, |entry| View::Stored(&entry.cell))
    }

    #[cfg(test)]
    pub(crate) fn get(&self, address: Address) -> Option<&Cell> {
        self.entry(address).map(|entry| &entry.cell)
    }

    pub(crate) fn entry(&self, address: Address) -> Option<&Stored> {
        self.cells
            .binary_search_by_key(&address, |entry| entry.address)
            .ok()
            .and_then(|index| self.cells.get(index))
    }

    pub(crate) fn entries(&self) -> &[Stored] {
        &self.cells
    }

    pub(crate) fn row(&self, index: RowIndex) -> row::Row<'_> {
        row::Row::new(index, self.row_entry(index))
    }

    pub(crate) fn row_entry(&self, index: RowIndex) -> Option<&row::Stored> {
        self.rows
            .binary_search_by_key(&index, |entry| entry.index)
            .ok()
            .and_then(|position| self.rows.get(position))
    }

    pub(crate) fn row_entries(&self) -> &[row::Stored] {
        &self.rows
    }

    pub(crate) fn rows(&self) -> row::Rows<'_> {
        row::Rows::new(&self.rows)
    }

    pub(crate) fn column(&self, index: ColumnIndex) -> column::Column<'_> {
        column::Column::new(index, self.column_entry(index))
    }

    pub(crate) fn column_entry(&self, index: ColumnIndex) -> Option<&column::Stored> {
        column::entry(&self.columns, index)
    }

    pub(crate) fn column_entries(&self) -> &[column::Stored] {
        &self.columns
    }

    pub(crate) fn columns(&self) -> column::Columns<'_> {
        column::Columns::new(&self.columns)
    }

    pub(crate) const fn defaults(&self) -> Option<&Defaults> {
        self.defaults.as_ref()
    }

    pub(crate) fn merges(&self) -> merge::Merges<'_> {
        self.merges.iter()
    }

    pub(crate) fn merge_ranges(&self) -> &[Rect] {
        self.merges.as_slice()
    }

    pub(crate) fn cells(&self, range: Rect) -> Cells<'_> {
        Cells {
            entries: self.entries_in(range),
        }
    }

    /// The stored record at `index` of the row-major record order, in O(1).
    pub(crate) fn record(&self, index: RecordIndex) -> Option<(Address, &Cell)> {
        self.cells
            .get(index.0)
            .map(|entry| (entry.address, &entry.cell))
    }

    /// Stored records inside `range`, in row-major order, with their
    /// record-level metadata.
    pub(crate) fn entries_in(&self, range: Rect) -> Entries<'_> {
        let start = self
            .cell_rows
            .partition_point(|entry| entry.row < range.start().row());
        Entries {
            cells: &self.cells,
            rows: &self.cell_rows[start..],
            current: &[],
            position: 0,
            range,
        }
    }

    pub(crate) const fn extents(&self) -> &Extents {
        &self.extents
    }

    /// Logical retained weight of this parsed store, in bytes.
    ///
    /// The weight counts the store's record arrays, its merge index, its
    /// repair report, and the payload text each record owns. Shared-string
    /// text belongs to the workbook's shared-string table and is not counted
    /// again. Allocator rounding is ignored, so this is a cache weight, not a
    /// resident-memory measurement.
    pub(crate) fn weight_bytes(&self) -> usize {
        let arrays = [
            (self.cells.len(), size_of::<Stored>()),
            (self.cell_rows.len(), size_of::<CellRowStart>()),
            (self.rows.len(), size_of::<row::Stored>()),
            (self.columns.len(), size_of::<column::Stored>()),
            (
                self.tolerance.defects().len(),
                size_of::<crate::leniency::ToleratedDefect>(),
            ),
        ];
        let mut weight = size_of::<Self>().saturating_add(self.merges.weight_bytes());
        for (count, size) in arrays {
            weight = weight.saturating_add(count.saturating_mul(size));
        }
        self.cells.iter().fold(weight, |weight, entry| {
            weight.saturating_add(entry.owned_bytes())
        })
    }
}

/// Modeled header of one shared text allocation (its strong and weak counts).
const SHARED_TEXT_HEADER_BYTES: usize = 2 * size_of::<usize>();

impl Stored {
    /// Payload bytes this record owns outside its inline size.
    fn owned_bytes(&self) -> usize {
        let cell = match &self.cell {
            Cell::Empty => 0,
            Cell::Value(value) => value_owned_bytes(value, self.shared_string.is_some()),
            Cell::Formula(formula) => formula_owned_bytes(formula),
            Cell::Unknown(unknown) => text_owned_bytes(&unknown.kind)
                .saturating_add(unknown.value.as_ref().map_or(0, text_owned_bytes))
                .saturating_add(unknown.formula.as_ref().map_or(0, text_owned_bytes)),
        };
        cell.saturating_add(
            self.shared_formula
                .as_ref()
                .map_or(0, |shared| shared.reference.len()),
        )
    }
}

fn text_owned_bytes(text: &Text) -> usize {
    text.len().saturating_add(SHARED_TEXT_HEADER_BYTES)
}

/// `shared` text is owned by the shared-string table, not by the record.
fn value_owned_bytes(value: &Value, shared: bool) -> usize {
    match value {
        Value::Bool(_) => 0,
        Value::Number(number) => number.as_str().len(),
        Value::Text(_) if shared => 0,
        Value::Text(text) => text_owned_bytes(text),
        Value::Date(date) => text_owned_bytes(&date.0),
        Value::Error(ErrorValue::Unknown(text)) => text_owned_bytes(text),
        Value::Error(_) => 0,
    }
}

fn formula_owned_bytes(formula: &Formula) -> usize {
    let kind = match formula.kind() {
        Kind::Scalar => 0,
        Kind::Array { range } | Kind::DataTable { range } => {
            range.as_ref().map_or(0, text_owned_bytes)
        },
        Kind::Unknown(text) => text_owned_bytes(text),
    };
    formula.text().len().saturating_add(kind).saturating_add(
        formula
            .cached()
            .map_or(0, |cache| value_owned_bytes(cache.value(), false)),
    )
}

fn omitted_ranges_are_ordered(ranges: &[Rect]) -> bool {
    let mut previous_end = None::<(u32, u32)>;
    for range in ranges {
        let start = range.start();
        let row = start.row().get();
        if row.checked_add(1) != Some(range.end().0) {
            return false;
        }
        let start = (row, start.column().get());
        if previous_end.is_some_and(|end| start < end) {
            return false;
        }
        // Rect::end() is a two-dimensional exclusive bound. A single-row
        // omission ends at this column in its own row, not in the next row.
        previous_end = Some((row, range.end().1));
    }
    true
}

fn omitted_entries<'a>(
    entries: &'a [Stored],
    ranges: &'a [Rect],
) -> impl Iterator<Item = &'a Stored> {
    let mut range_index = 0usize;
    entries.iter().filter(move |entry| {
        let row = entry.address.row().get();
        let column = entry.address.column().get();
        while let Some(range) = ranges.get(range_index) {
            let range_row = range.start().row().get();
            if row > range_row || (row == range_row && column >= range.end().1) {
                range_index += 1;
                continue;
            }
            return range.contains(entry.address);
        }
        false
    })
}

#[derive(Debug, Default)]
struct Bounds {
    value: Option<(u32, u32, u32, u32)>,
}

impl Bounds {
    fn push(&mut self, address: Address) {
        let row = address.row().get();
        let column = address.column().get();
        self.value = Some(self.value.map_or(
            (row, column, row, column),
            |(min_row, min_column, max_row, max_column)| {
                (
                    min_row.min(row),
                    min_column.min(column),
                    max_row.max(row),
                    max_column.max(column),
                )
            },
        ));
    }

    fn finish(self) -> Result<Option<Rect>> {
        let Some((min_row, min_column, max_row, max_column)) = self.value else {
            return Ok(None);
        };
        let start = Address::at(min_row, min_column)?;
        Rect::new(start, max_row + 1, max_column + 1)
            .map(Some)
            .map_err(|error| invalid(error.to_string()))
    }
}

/// Borrowed sparse cells inside a half-open range.
#[derive(Debug)]
pub struct Cells<'a> {
    entries: Entries<'a>,
}

impl<'a> Cells<'a> {
    /// The same traversal, with each record's position in its grid's
    /// row-major record order, which reads the record again in O(1)
    /// ([`crate::workbook::source::Grid::record`], [`crate::Worksheet::record`]).
    ///
    /// ```
    /// # fn demo(grid: &litchi_xlsx::workbook::source::Grid) -> litchi_xlsx::Result<()> {
    /// for (index, at, cell) in grid.cells("A1:C10")?.indexed() {
    ///     assert_eq!(grid.record(index), Some((at, cell)));
    /// }
    /// # Ok(())
    /// # }
    /// ```
    #[must_use]
    pub fn indexed(self) -> IndexedCells<'a> {
        IndexedCells {
            entries: self.entries,
        }
    }
}

impl<'a> Iterator for Cells<'a> {
    type Item = (Address, &'a Cell);

    #[inline]
    fn next(&mut self) -> Option<Self::Item> {
        self.entries
            .next()
            .map(|entry| (entry.address, &entry.cell))
    }
}

/// The position of a stored record in its grid's row-major record order.
///
/// It reads the record again in O(1) from the grid that produced it
/// ([`Cells::indexed`]); a position from another grid reads another record
/// or none.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct RecordIndex(usize);

impl RecordIndex {
    /// The zero-based position.
    #[must_use]
    pub const fn get(self) -> usize {
        self.0
    }

    /// The record position `index`.
    #[must_use]
    pub const fn new(index: usize) -> Self {
        Self(index)
    }
}

/// Borrowed sparse cells inside a half-open range, each with its
/// [`RecordIndex`] ([`Cells::indexed`]).
#[derive(Debug)]
pub struct IndexedCells<'a> {
    entries: Entries<'a>,
}

impl<'a> Iterator for IndexedCells<'a> {
    type Item = (RecordIndex, Address, &'a Cell);

    #[inline]
    fn next(&mut self) -> Option<Self::Item> {
        let (index, entry) = self.entries.next_indexed()?;
        Some((RecordIndex(index), entry.address, &entry.cell))
    }
}

/// Borrowed stored records inside a half-open range, in row-major order.
#[derive(Debug)]
pub(crate) struct Entries<'a> {
    cells: &'a [Stored],
    rows: &'a [CellRowStart],
    current: &'a [Stored],
    /// The record position of `current`'s first entry.
    position: usize,
    range: Rect,
}

impl<'a> Iterator for Entries<'a> {
    type Item = &'a Stored;

    #[inline]
    fn next(&mut self) -> Option<Self::Item> {
        self.next_indexed().map(|(_, entry)| entry)
    }
}

impl<'a> Entries<'a> {
    /// The next record with its position in the record order.
    fn next_indexed(&mut self) -> Option<(usize, &'a Stored)> {
        loop {
            if let Some((entry, remaining)) = self.current.split_first() {
                self.current = remaining;
                let index = self.position;
                self.position += 1;
                return Some((index, entry));
            }

            let (row, remaining_rows) = self.rows.split_first()?;
            self.rows = remaining_rows;
            if row.row.get() >= self.range.end().0 {
                self.rows = &[];
                return None;
            }

            let start_column = self.range.start().column().get();
            let end = remaining_rows
                .first()
                .map_or(self.cells.len(), |next| next.start);
            let row_cells = &self.cells[row.start..end];
            let start = if start_column == ColumnIndex::FIRST.get() {
                0
            } else {
                row_cells.partition_point(|entry| entry.address.column().get() < start_column)
            };
            let selected = &row_cells[start..];
            let selected_len = if self.range.end().1 == ColumnIndex::LAST.get() + 1 {
                selected.len()
            } else if self.range.end().1 == start_column + 1 {
                usize::from(
                    selected
                        .first()
                        .is_some_and(|entry| entry.address.column().get() == start_column),
                )
            } else {
                selected.partition_point(|entry| entry.address.column().get() < self.range.end().1)
            };
            self.current = &selected[..selected_len];
            self.position = row.start + start;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn indexed_cells_read_their_records_again() {
        let addresses = [(0, 0), (0, 3), (0, 7), (2, 1), (2, 2), (5, 0), (5, 9)];
        let store = empty_store(&addresses);
        let whole = Rect::new(Address::at(0, 0).expect("start"), 10, 10).expect("range");
        let indexed: Vec<_> = store.cells(whole).indexed().collect();
        assert_eq!(indexed.len(), addresses.len());
        for (position, (index, at, cell)) in indexed.iter().enumerate() {
            assert_eq!(index.get(), position, "the whole grid lists every record");
            assert_eq!(store.record(*index), Some((*at, *cell)));
        }
        // A range that starts and ends mid-row lists the records it covers,
        // each at its position in the whole order.
        let inner = Rect::new(Address::at(0, 1).expect("start"), 6, 8).expect("range");
        let plain: Vec<_> = store.cells(inner).collect();
        let indexed: Vec<_> = store.cells(inner).indexed().collect();
        assert_eq!(plain.len(), indexed.len());
        for ((at, cell), (index, indexed_at, indexed_cell)) in plain.iter().zip(&indexed) {
            assert_eq!((at, cell), (indexed_at, indexed_cell));
            assert_eq!(store.record(*index), Some((*at, *cell)));
        }
        assert_eq!(
            indexed
                .iter()
                .map(|(index, ..)| index.get())
                .collect::<Vec<_>>(),
            [1, 2, 3, 4]
        );
        assert_eq!(store.record(RecordIndex::new(addresses.len())), None);
    }

    fn empty_store(addresses: &[(u32, u32)]) -> Store {
        let cells = addresses
            .iter()
            .map(|&(row, column)| Stored {
                address: Address::at(row, column).expect("bounded test address"),
                cell: Cell::Empty,
                style: None,
                shared_string: None,
                inline_rich: false,
                formula_range: None,
                shared_formula: None,
                cell_metadata: None,
                value_metadata: None,
            })
            .collect();
        Store::from_unsorted(cells, Vec::new(), Box::new([]), None, Vec::new(), None)
            .expect("valid cell store")
    }

    #[test]
    fn omitted_range_collisions_are_detected() {
        let store = empty_store(&[(0, 1), (2, 3)]);
        let range = |row: u32, first: u32, last: u32| {
            Rect::new(Address::at(row, first).expect("start"), row + 1, last + 1)
                .expect("single-row range")
        };
        assert!(store.avoids_omitted_cells(&[]));
        assert!(store.avoids_omitted_cells(&[range(0, 2, 5), range(1, 0, 9), range(2, 0, 2)]));
        // A stored cell inside an omitted range collides.
        assert!(!store.avoids_omitted_cells(&[range(0, 0, 1)]));
        assert!(!store.avoids_omitted_cells(&[range(0, 2, 5), range(2, 3, 3)]));
        // Unordered or multi-row omissions are refused like the merge refuses them.
        assert!(!store.avoids_omitted_cells(&[range(1, 0, 0), range(0, 5, 6)]));
        let tall = Rect::new(Address::at(0, 2).expect("start"), 2, 3).expect("two rows");
        assert!(!store.avoids_omitted_cells(&[tall]));
    }

    #[test]
    fn range_cells_skip_columns_outside_each_selected_row() {
        let store = empty_store(&[
            (0, 0),
            (0, 1),
            (0, 2),
            (0, 3),
            (1, 0),
            (1, 2),
            (1, 3),
            (2, 0),
            (2, 1),
            (2, 2),
            (2, 3),
            (3, 1),
        ]);
        let range = Rect::at(0, 1, 3, 3).expect("bounded test range");

        let addresses = store
            .cells(range)
            .map(|(address, _cell)| address)
            .collect::<Vec<_>>();

        assert_eq!(
            addresses,
            vec![
                Address::at(0, 1).unwrap(),
                Address::at(0, 2).unwrap(),
                Address::at(1, 2).unwrap(),
                Address::at(2, 1).unwrap(),
                Address::at(2, 2).unwrap(),
            ]
        );
    }

    #[test]
    fn range_cells_handle_sparse_rows_and_grid_edges() {
        let store = empty_store(&[(0, 0), (2, 0), (2, 16_383), (3, 0), (1_048_575, 16_383)]);

        let middle = Rect::at(1, 0, 4, 1).expect("bounded middle range");
        assert_eq!(
            store
                .cells(middle)
                .map(|(address, _cell)| address)
                .collect::<Vec<_>>(),
            vec![Address::at(2, 0).unwrap(), Address::at(3, 0).unwrap()]
        );

        let last = Rect::single(Address::at(1_048_575, 16_383).unwrap());
        assert_eq!(
            store
                .cells(last)
                .map(|(address, _cell)| address)
                .collect::<Vec<_>>(),
            vec![Address::at(1_048_575, 16_383).unwrap()]
        );
    }

    #[test]
    fn numbers_preserve_lexemes_and_convert_explicitly() {
        let number = Number::new("  -0.000  ").expect("valid number");
        assert_eq!(number.as_str(), "  -0.000  ");
        assert_eq!(number.as_f64(), Some(-0.0));
        assert!(Number::new("NaN").is_err());
        assert!(Number::new("not a number").is_err());
        assert!(Number::try_from(f64::INFINITY).is_err());
    }

    #[test]
    fn borrowed_number_validation_matches_owned_constructor_errors() {
        for (value, expected) in [
            ("not a number", "invalid worksheet number 'not a number'"),
            ("NaN", "non-finite worksheet number 'NaN'"),
            ("1e999", "non-finite worksheet number '1e999'"),
        ] {
            let borrowed = Number::validate_lexical(value)
                .expect_err("invalid borrowed number")
                .to_string();
            let owned = Number::new(value)
                .expect_err("invalid owned number")
                .to_string();
            assert_eq!(borrowed, owned);
            assert!(owned.contains(expected));
        }
        for value in ["  -0.000  ", "6.02E+23"] {
            Number::validate_lexical(value).expect("valid borrowed number");
            Number::new(value).expect("valid owned number");
        }
    }

    #[test]
    fn short_integer_fast_path_matches_the_float_validation() {
        fn reference(value: &str) -> bool {
            value
                .trim()
                .parse::<f64>()
                .is_ok_and(|parsed| parsed.is_finite())
        }
        let mut cases = vec![
            "0".to_owned(),
            "-1".to_owned(),
            "+1".to_owned(),
            "007".to_owned(),
            "-0".to_owned(),
            "123456789012345678".to_owned(),
            "-123456789012345678".to_owned(),
            "1234567890123456789".to_owned(),
            "9".repeat(309),
            String::new(),
            "-".to_owned(),
            "+".to_owned(),
            "--1".to_owned(),
            "+-1".to_owned(),
            "1-".to_owned(),
            " 1".to_owned(),
            "1 ".to_owned(),
            "1.".to_owned(),
            "1e5".to_owned(),
            "0x1".to_owned(),
            "\u{663}".to_owned(),
            "１".to_owned(),
            "inf".to_owned(),
            "NaN".to_owned(),
        ];
        for length in 1..=20 {
            cases.push("7".repeat(length));
            cases.push(format!("-{}", "3".repeat(length)));
        }
        for value in cases {
            if is_short_integer(&value) {
                assert!(reference(&value), "{value}");
            }
            assert_eq!(
                Number::validate_lexical(&value).is_ok(),
                reference(&value),
                "{value}"
            );
        }
    }

    #[test]
    fn edit_content_keeps_text_inert_and_formulas_explicit() {
        assert!(matches!(
            Content::from("=SUM(A1:A3)"),
            Content::Value(Value::Text(text)) if text.as_str() == "=SUM(A1:A3)"
        ));
        assert!(matches!(
            Content::from(42_i32),
            Content::Value(Value::Number(number)) if number.as_str() == "42"
        ));
        assert!(matches!(
            Content::from(Formula::new("SUM(A1:A3)").expect("formula")),
            Content::Formula(_)
        ));
    }

    #[test]
    fn dates_are_checked_and_keep_their_lexical_form() {
        let date = Date::new("2026-07-31T12:34:56.250-07:00").expect("date");
        assert_eq!(date.as_str(), "2026-07-31T12:34:56.250-07:00");
        assert!(Date::new("2026-02-29").is_err());
        assert!(Value::date("not a date").is_err());
    }

    #[test]
    fn producer_unknown_error_values_are_read_only() {
        let content = Content::Value(Value::Error(ErrorValue::Unknown("#VENDOR!".into())));
        assert!(content.validate_for_write().is_err());
    }
}
