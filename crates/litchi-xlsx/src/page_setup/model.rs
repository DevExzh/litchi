//! Typed SpreadsheetML worksheet page-setup models.

use std::borrow::Cow;
use std::error::Error as StdError;
use std::fmt::{Display, Formatter};
use std::str::FromStr;

// These bounds are shared with the page-setup XML codec but are not part of
// the public facade.
pub(super) const MAX_MEASURE_BYTES: usize = 128;
pub(super) const MAX_RELATIONSHIP_ID_BYTES: usize = 4096;

/// Printed page orientation.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Orientation {
    /// Let the printer or application choose its default orientation.
    #[default]
    Default,
    /// Print with the shorter edge horizontal.
    Portrait,
    /// Print with the longer edge horizontal.
    Landscape,
}

impl Orientation {
    /// Return the exact SpreadsheetML token.
    #[inline]
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Default => "default",
            Self::Portrait => "portrait",
            Self::Landscape => "landscape",
        }
    }
}

impl FromStr for Orientation {
    type Err = LexicalError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "default" => Ok(Self::Default),
            "portrait" => Ok(Self::Portrait),
            "landscape" => Ok(Self::Landscape),
            _ => Err(LexicalError::new("page orientation")),
        }
    }
}

impl Display for Orientation {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RangeKind {
    Paper,
    Scale,
    Fit,
    FirstPage,
    Copies,
    Dpi,
}

/// Error returned when a numeric page-setup value is outside its Office domain.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RangeError {
    kind: RangeKind,
    value: i64,
}

impl RangeError {
    const fn new(kind: RangeKind, value: i64) -> Self {
        Self { kind, value }
    }

    /// Return the rejected value.
    #[inline]
    #[must_use]
    pub const fn value(self) -> i64 {
        self.value
    }

    /// Return the page-setup field whose range was violated.
    #[inline]
    #[must_use]
    pub const fn field(self) -> &'static str {
        match self.kind {
            RangeKind::Paper => "paper size",
            RangeKind::Scale => "scale",
            RangeKind::Fit => "fit",
            RangeKind::FirstPage => "first page",
            RangeKind::Copies => "copies",
            RangeKind::Dpi => "DPI",
        }
    }
}

impl Display for RangeError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self.kind {
            RangeKind::Paper => write!(formatter, "paper size {} is reserved", self.value),
            RangeKind::Scale => write!(
                formatter,
                "page scale {} is neither automatic value 0 nor in 10..=400",
                self.value
            ),
            RangeKind::Fit => write!(formatter, "page fit {} is outside 0..=32,767", self.value),
            RangeKind::FirstPage => write!(
                formatter,
                "first page {} is neither in -32,767..=-1 nor 1..=32,767",
                self.value
            ),
            RangeKind::Copies => {
                write!(formatter, "copy count {} is outside 1..=32,767", self.value)
            },
            RangeKind::Dpi => write!(formatter, "printer DPI {} must be at least 1", self.value),
        }
    }
}

impl StdError for RangeError {}

/// Checked printer paper-size code.
///
/// SpreadsheetML defines this as `unsignedInt`: named values occupy `1..=118`
/// except reserved values `48` and `49`, `119..=255` is reserved, and custom
/// printer codes occupy `256..=u32::MAX`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Paper(u32);

impl Paper {
    /// US Letter (`8.5 × 11` inches).
    pub const LETTER: Self = Self(1);
    /// ISO A4 (`210 × 297` millimeters).
    pub const A4: Self = Self(9);
    /// Largest value permitted by SpreadsheetML's `unsignedInt` domain.
    pub const MAX: u32 = u32::MAX;

    /// Construct a checked paper-size code.
    pub const fn new(value: u32) -> Result<Self, RangeError> {
        if matches!(value, 1..=47 | 50..=118) || value >= 256 {
            Ok(Self(value))
        } else {
            Err(RangeError::new(RangeKind::Paper, value as i64))
        }
    }

    /// Return the numeric SpreadsheetML paper-size code.
    #[inline]
    #[must_use]
    pub const fn get(self) -> u32 {
        self.0
    }

    /// Whether this is one of Office's named paper sizes.
    #[inline]
    #[must_use]
    pub const fn is_named(self) -> bool {
        matches!(self.0, 1..=47 | 50..=118)
    }

    /// Whether printer settings define the paper dimensions.
    #[inline]
    #[must_use]
    pub const fn is_custom(self) -> bool {
        self.0 >= 256
    }
}

impl Default for Paper {
    fn default() -> Self {
        Self::LETTER
    }
}

impl TryFrom<u32> for Paper {
    type Error = RangeError;

    fn try_from(value: u32) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

impl From<Paper> for u32 {
    fn from(value: Paper) -> Self {
        value.get()
    }
}

impl Display for Paper {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        self.0.fmt(formatter)
    }
}

/// Checked worksheet print scale.
///
/// Value `0` requests automatic scaling; explicit percentages are `10..=400`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Scale(u16);

impl Scale {
    /// Let printer settings choose the scale.
    pub const AUTO: Self = Self(0);
    /// The schema default of 100 percent.
    pub const DEFAULT: Self = Self(100);

    /// Construct automatic scaling (`0`) or an explicit `10..=400` percentage.
    pub const fn new(value: u16) -> Result<Self, RangeError> {
        if value == 0 || (value >= 10 && value <= 400) {
            Ok(Self(value))
        } else {
            Err(RangeError::new(RangeKind::Scale, value as i64))
        }
    }

    /// Return the numeric SpreadsheetML scale.
    #[inline]
    #[must_use]
    pub const fn get(self) -> u16 {
        self.0
    }

    /// Whether printer settings should choose the scale automatically.
    #[inline]
    #[must_use]
    pub const fn is_auto(self) -> bool {
        self.0 == 0
    }
}

impl Default for Scale {
    fn default() -> Self {
        Self::DEFAULT
    }
}

impl TryFrom<u32> for Scale {
    type Error = RangeError;

    fn try_from(value: u32) -> Result<Self, Self::Error> {
        if value == 0 || (10..=400).contains(&value) {
            Ok(Self(value as u16))
        } else {
            Err(RangeError::new(RangeKind::Scale, value as i64))
        }
    }
}

impl From<Scale> for u16 {
    fn from(value: Scale) -> Self {
        value.get()
    }
}

impl Display for Scale {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        self.0.fmt(formatter)
    }
}

/// Checked number of pages used by a fit-to-width or fit-to-height setting.
///
/// Value `0` means the output is not fit to a specific count in that axis.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Fit(u16);

impl Fit {
    /// Do not constrain this axis to a specific page count.
    pub const NONE: Self = Self(0);
    /// Fit this axis to one page.
    pub const ONE: Self = Self(1);
    /// Largest page count supported by Office.
    pub const MAX: u16 = 32_767;

    /// Construct a page-fit count in `0..=32,767`.
    pub const fn new(value: u16) -> Result<Self, RangeError> {
        if value <= Self::MAX {
            Ok(Self(value))
        } else {
            Err(RangeError::new(RangeKind::Fit, value as i64))
        }
    }

    /// Return the numeric SpreadsheetML page-fit count.
    #[inline]
    #[must_use]
    pub const fn get(self) -> u16 {
        self.0
    }

    /// Whether no specific page count is requested for this axis.
    #[inline]
    #[must_use]
    pub const fn is_unbounded(self) -> bool {
        self.0 == 0
    }
}

impl Default for Fit {
    fn default() -> Self {
        Self::ONE
    }
}

impl TryFrom<u32> for Fit {
    type Error = RangeError;

    fn try_from(value: u32) -> Result<Self, Self::Error> {
        if value <= u32::from(Self::MAX) {
            Ok(Self(value as u16))
        } else {
            Err(RangeError::new(RangeKind::Fit, value as i64))
        }
    }
}

impl From<Fit> for u16 {
    fn from(value: Fit) -> Self {
        value.get()
    }
}

impl Display for Fit {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        self.0.fmt(formatter)
    }
}

/// Order in which worksheet pages are printed.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Order {
    /// Print top-to-bottom before moving right.
    #[default]
    DownThenOver,
    /// Print left-to-right before moving down.
    OverThenDown,
}

impl Order {
    /// Return the exact SpreadsheetML token.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::DownThenOver => "downThenOver",
            Self::OverThenDown => "overThenDown",
        }
    }
}

impl FromStr for Order {
    type Err = LexicalError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "downThenOver" => Ok(Self::DownThenOver),
            "overThenDown" => Ok(Self::OverThenDown),
            _ => Err(LexicalError::new("page order")),
        }
    }
}

impl Display for Order {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// How cell comments are printed.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Comments {
    /// Do not print cell comments.
    #[default]
    None,
    /// Print comments as displayed on the worksheet.
    AsDisplayed,
    /// Print comments after the worksheet.
    AtEnd,
}

impl Comments {
    /// Return the exact SpreadsheetML token.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::AsDisplayed => "asDisplayed",
            Self::AtEnd => "atEnd",
        }
    }
}

impl FromStr for Comments {
    type Err = LexicalError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "none" => Ok(Self::None),
            "asDisplayed" => Ok(Self::AsDisplayed),
            "atEnd" => Ok(Self::AtEnd),
            _ => Err(LexicalError::new("cell-comments mode")),
        }
    }
}

impl Display for Comments {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// How cell errors are printed.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum ErrorMode {
    /// Print the displayed error value.
    #[default]
    Displayed,
    /// Print a blank cell.
    Blank,
    /// Print a dash.
    Dash,
    /// Print `#N/A`.
    NotAvailable,
}

impl ErrorMode {
    /// Return the exact SpreadsheetML token.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Displayed => "displayed",
            Self::Blank => "blank",
            Self::Dash => "dash",
            Self::NotAvailable => "NA",
        }
    }
}

impl FromStr for ErrorMode {
    type Err = LexicalError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "displayed" => Ok(Self::Displayed),
            "blank" => Ok(Self::Blank),
            "dash" => Ok(Self::Dash),
            "NA" => Ok(Self::NotAvailable),
            _ => Err(LexicalError::new("print-errors mode")),
        }
    }
}

impl Display for ErrorMode {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Unit identifier from `ST_PositiveUniversalMeasure`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Unit {
    Millimeter,
    Centimeter,
    Inch,
    Point,
    Pica,
    PicaAlternative,
}

impl Unit {
    /// Return the exact `ST_PositiveUniversalMeasure` suffix.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Millimeter => "mm",
            Self::Centimeter => "cm",
            Self::Inch => "in",
            Self::Point => "pt",
            Self::Pica => "pc",
            Self::PicaAlternative => "pi",
        }
    }
}

impl FromStr for Unit {
    type Err = LexicalError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "mm" => Ok(Self::Millimeter),
            "cm" => Ok(Self::Centimeter),
            "in" => Ok(Self::Inch),
            "pt" => Ok(Self::Point),
            "pc" => Ok(Self::Pica),
            "pi" => Ok(Self::PicaAlternative),
            _ => Err(LexicalError::new("universal-measure unit")),
        }
    }
}

impl Display for Unit {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Error returned for an invalid exact page-setup lexical value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LexicalError {
    field: &'static str,
}

impl LexicalError {
    const fn new(field: &'static str) -> Self {
        Self { field }
    }

    /// Return the invalid field name.
    #[must_use]
    pub const fn field(self) -> &'static str {
        self.field
    }
}

impl Display for LexicalError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        write!(formatter, "invalid {}", self.field)
    }
}

impl StdError for LexicalError {}

/// Exact `ST_PositiveUniversalMeasure` decimal and unit.
///
/// The decimal is retained verbatim after validating the schema grammar
/// `[0-9]+(\.[0-9]+)?`; floating-point rounding and exponent notation are
/// therefore impossible.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Measure {
    decimal: Box<str>,
    unit: Unit,
}

impl Measure {
    /// Construct an exact non-negative decimal measure.
    ///
    /// A borrowed `&str` is length-checked before allocation; an owned
    /// [`String`] is consumed and reused for storage.
    pub fn new<'a>(decimal: impl Into<Cow<'a, str>>, unit: Unit) -> Result<Self, LexicalError> {
        let decimal = decimal.into();
        if decimal
            .len()
            .checked_add(unit.as_str().len())
            .is_some_and(|length| length <= MAX_MEASURE_BYTES)
            && valid_decimal(&decimal)
        {
            let decimal = match decimal {
                Cow::Borrowed(value) => Box::<str>::from(value),
                Cow::Owned(value) => value.into_boxed_str(),
            };
            Ok(Self { decimal, unit })
        } else {
            Err(LexicalError::new("measure decimal"))
        }
    }

    /// Construct an integer measure without lexical validation failure.
    #[must_use]
    pub fn whole(value: u32, unit: Unit) -> Self {
        Self {
            decimal: value.to_string().into_boxed_str(),
            unit,
        }
    }

    /// Return the exact decimal without its unit suffix.
    #[must_use]
    pub fn decimal(&self) -> &str {
        &self.decimal
    }

    /// Return the unit.
    #[must_use]
    pub const fn unit(&self) -> Unit {
        self.unit
    }
}

impl Display for Measure {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        write!(formatter, "{}{}", self.decimal, self.unit.as_str())
    }
}

/// Signed first printed page number.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct FirstPage(i16);

impl FirstPage {
    pub const MIN: i16 = -32_767;
    pub const MAX: i16 = 32_767;
    pub const DEFAULT: Self = Self(1);

    /// Construct a first-page number in Office's interoperable range.
    pub const fn new(value: i32) -> Result<Self, RangeError> {
        if (value >= Self::MIN as i32 && value < 0) || (value >= 1 && value <= Self::MAX as i32) {
            Ok(Self(value as i16))
        } else {
            Err(RangeError::new(RangeKind::FirstPage, value as i64))
        }
    }

    /// Decode SpreadsheetML's unsigned two's-complement representation.
    pub const fn from_wire(value: u32) -> Result<Self, RangeError> {
        if value <= Self::MAX as u32 {
            Self::new(value as i32)
        } else {
            let signed = value as i32;
            if signed >= Self::MIN as i32 && signed < 0 {
                Self::new(signed)
            } else {
                Err(RangeError::new(RangeKind::FirstPage, value as i64))
            }
        }
    }

    /// Return the semantic signed page number.
    #[must_use]
    pub const fn get(self) -> i16 {
        self.0
    }

    /// Return the SpreadsheetML `unsignedInt` representation.
    #[must_use]
    pub const fn wire(self) -> u32 {
        (self.0 as i32) as u32
    }
}

impl Default for FirstPage {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Checked number of copies to print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Copies(u16);

impl Copies {
    pub const ONE: Self = Self(1);
    pub const MAX: u16 = 32_767;

    /// Construct a copy count in `1..=32,767`.
    pub const fn new(value: u16) -> Result<Self, RangeError> {
        if value >= 1 && value <= Self::MAX {
            Ok(Self(value))
        } else {
            Err(RangeError::new(RangeKind::Copies, value as i64))
        }
    }

    /// Return the copy count.
    #[must_use]
    pub const fn get(self) -> u16 {
        self.0
    }
}

impl Default for Copies {
    fn default() -> Self {
        Self::ONE
    }
}

impl TryFrom<u32> for Copies {
    type Error = RangeError;

    fn try_from(value: u32) -> Result<Self, Self::Error> {
        if value >= 1 && value <= u32::from(Self::MAX) {
            Ok(Self(value as u16))
        } else {
            Err(RangeError::new(RangeKind::Copies, value as i64))
        }
    }
}

/// Printer resolution in dots per inch.
///
/// Office requires an explicit resolution to be at least one dot per inch.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Dpi(u32);

impl Dpi {
    pub const DEFAULT: Self = Self(600);

    /// Construct an explicit positive printer resolution.
    pub const fn new(value: u32) -> Result<Self, RangeError> {
        if value >= 1 {
            Ok(Self(value))
        } else {
            Err(RangeError::new(RangeKind::Dpi, value as i64))
        }
    }

    /// Return the wire value.
    #[must_use]
    pub const fn get(self) -> u32 {
        self.0
    }
}

impl Default for Dpi {
    fn default() -> Self {
        Self::DEFAULT
    }
}

impl TryFrom<u32> for Dpi {
    type Error = RangeError;

    fn try_from(value: u32) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

/// Validated internal OPC relationship identifier.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct RelId(Box<str>);

impl RelId {
    /// Construct an XML `NCName` relationship identifier.
    ///
    /// A borrowed `&str` is length-checked before allocation; an owned
    /// [`String`] is consumed and reused for storage.
    pub fn new<'a>(value: impl Into<Cow<'a, str>>) -> Result<Self, LexicalError> {
        let value = value.into();
        if value.len() <= MAX_RELATIONSHIP_ID_BYTES && litchi_ooxml_common::xml::is_ncname(&value) {
            let value = match value {
                Cow::Borrowed(value) => Box::<str>::from(value),
                Cow::Owned(value) => value.into_boxed_str(),
            };
            Ok(Self(value))
        } else {
            Err(LexicalError::new("printer-settings relationship ID"))
        }
    }

    /// Borrow the relationship identifier.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Move the identifier into a string for the focused relationship layer.
    pub fn into_string(self) -> String {
        self.0.into()
    }
}

impl Display for RelId {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        self.0.fmt(formatter)
    }
}

/// One lossless typed worksheet `pageSetup` settings value.
///
/// Every core setting is optional so XML attribute absence survives read,
/// update, and publication. The printer-settings relationship is deliberately
/// owned by [`crate::xlsx::printer_settings`], which prevents ordinary page
/// authoring from emitting a dangling raw relationship ID. Public fields keep
/// struct-update syntax concise while private representations prevent invalid
/// fixed-domain values.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct Setup {
    /// Paper code; absent has schema default [`Paper::LETTER`].
    pub paper: Option<Paper>,
    /// Exact custom paper width.
    pub paper_width: Option<Measure>,
    /// Exact custom paper height.
    pub paper_height: Option<Measure>,
    /// Print scale; absent has schema default [`Scale::DEFAULT`].
    pub scale: Option<Scale>,
    /// First page; absent has schema default [`FirstPage::DEFAULT`].
    pub first_page: Option<FirstPage>,
    /// Horizontal fit; absent has schema default [`Fit::ONE`].
    pub fit_to_width: Option<Fit>,
    /// Vertical fit; absent has schema default [`Fit::ONE`].
    pub fit_to_height: Option<Fit>,
    /// Page order; absent has [`Order::DownThenOver`].
    pub order: Option<Order>,
    /// Orientation; absent has [`Orientation::Default`].
    pub orientation: Option<Orientation>,
    /// Printer-default policy; absent is `true` per the schema.
    pub use_printer_defaults: Option<bool>,
    /// Black-and-white output; absent is `false`.
    pub black_and_white: Option<bool>,
    /// Draft output; absent is `false`.
    pub draft: Option<bool>,
    /// Comment printing; absent has [`Comments::None`].
    pub comments: Option<Comments>,
    /// Whether to use `first_page`; absent is `false`.
    pub use_first_page: Option<bool>,
    /// Error printing; absent has [`ErrorMode::Displayed`].
    pub errors: Option<ErrorMode>,
    /// Horizontal resolution; absent has [`Dpi::DEFAULT`].
    pub horizontal_dpi: Option<Dpi>,
    /// Vertical resolution; absent has [`Dpi::DEFAULT`].
    pub vertical_dpi: Option<Dpi>,
    /// Copy count; absent has [`Copies::ONE`].
    pub copies: Option<Copies>,
}

impl Setup {
    /// Create a paper-specific setup with every other attribute absent.
    #[must_use]
    pub const fn new(paper: Paper) -> Self {
        Self {
            paper: Some(paper),
            paper_width: None,
            paper_height: None,
            scale: None,
            first_page: None,
            fit_to_width: None,
            fit_to_height: None,
            order: None,
            orientation: None,
            use_printer_defaults: None,
            black_and_white: None,
            draft: None,
            comments: None,
            use_first_page: None,
            errors: None,
            horizontal_dpi: None,
            vertical_dpi: None,
            copies: None,
        }
    }

    /// Return the effective schema printer-default policy.
    #[must_use]
    pub const fn uses_printer_defaults(&self) -> bool {
        match self.use_printer_defaults {
            Some(value) => value,
            None => true,
        }
    }
}

fn valid_decimal(value: &str) -> bool {
    let mut parts = value.split('.');
    let Some(whole) = parts.next() else {
        return false;
    };
    if whole.is_empty() || !whole.bytes().all(|byte| byte.is_ascii_digit()) {
        return false;
    }
    if let Some(fraction) = parts.next()
        && (fraction.is_empty() || !fraction.bytes().all(|byte| byte.is_ascii_digit()))
    {
        return false;
    }
    parts.next().is_none()
}
