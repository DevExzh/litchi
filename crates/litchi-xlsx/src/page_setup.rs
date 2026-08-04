//! Typed worksheet page-setup metadata and authoring.

use std::borrow::Cow;
use std::error::Error as StdError;
use std::fmt::{Display, Formatter};
use std::str::FromStr;

use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::{Error, Result};
use litchi_ooxml_common::{MceCapabilities, MceLimits, process_markup_compatibility};

const CORE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const REL: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_MEASURE_BYTES: usize = 128;
const MAX_RELATIONSHIP_ID_BYTES: usize = 4096;

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

    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
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
    pub const fn new(value: u32) -> std::result::Result<Self, RangeError> {
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

    fn try_from(value: u32) -> std::result::Result<Self, Self::Error> {
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
    pub const fn new(value: u16) -> std::result::Result<Self, RangeError> {
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

    fn try_from(value: u32) -> std::result::Result<Self, Self::Error> {
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
    pub const fn new(value: u16) -> std::result::Result<Self, RangeError> {
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

    fn try_from(value: u32) -> std::result::Result<Self, Self::Error> {
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

    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
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

    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
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

    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
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

    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
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
    pub fn new<'a>(
        decimal: impl Into<Cow<'a, str>>,
        unit: Unit,
    ) -> std::result::Result<Self, LexicalError> {
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
    pub const fn new(value: i32) -> std::result::Result<Self, RangeError> {
        if (value >= Self::MIN as i32 && value < 0) || (value >= 1 && value <= Self::MAX as i32) {
            Ok(Self(value as i16))
        } else {
            Err(RangeError::new(RangeKind::FirstPage, value as i64))
        }
    }

    /// Decode SpreadsheetML's unsigned two's-complement representation.
    pub const fn from_wire(value: u32) -> std::result::Result<Self, RangeError> {
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
    pub const fn new(value: u16) -> std::result::Result<Self, RangeError> {
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

    fn try_from(value: u32) -> std::result::Result<Self, Self::Error> {
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
    pub const fn new(value: u32) -> std::result::Result<Self, RangeError> {
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

    fn try_from(value: u32) -> std::result::Result<Self, Self::Error> {
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
    pub fn new<'a>(value: impl Into<Cow<'a, str>>) -> std::result::Result<Self, LexicalError> {
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

/// Parse a worksheet's optional typed core `pageSetup` element.
pub fn parse_worksheet_page_setup(xml: &[u8]) -> Result<Option<Setup>> {
    Ok(parse_worksheet_page_setup_parts(xml, Projection::Settings)?.map(|parsed| parsed.setup))
}

/// Parse only the relationship edge owned by the printer-settings layer.
pub fn parse_worksheet_page_setup_relationship_id(xml: &[u8]) -> Result<Option<RelId>> {
    Ok(
        parse_worksheet_page_setup_parts(xml, Projection::Relationship)?
            .and_then(|parsed| parsed.printer_settings),
    )
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Projection {
    Settings,
    Relationship,
}

struct ParsedSetup {
    setup: Setup,
    printer_settings: Option<RelId>,
}

fn parse_worksheet_page_setup_parts(
    xml: &[u8],
    projection: Projection,
) -> Result<Option<ParsedSetup>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("worksheet XML is too large"));
    }
    let validated =
        process_markup_compatibility(xml, &MceCapabilities::default(), &MceLimits::default())?;
    let selected = if validated.report.alternate_content_count == 0 {
        xml
    } else {
        validated.xml.as_ref()
    };
    parse_selected(selected, projection)
}

fn parse_selected(xml: &[u8], projection: Projection) -> Result<Option<ParsedSetup>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut worksheet_namespace = None;
    let mut root_closed = false;
    let mut result = None;
    let mut open: Option<(usize, ParsedSetup)> = None;

    loop {
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("worksheet XML nesting overflow"))?;
                if depth > MAX_DEPTH {
                    return Err(invalid("worksheet XML nesting is too deep"));
                }
                if depth == 1 {
                    if worksheet_namespace.is_some()
                        || element.local_name().as_ref() != b"worksheet"
                    {
                        return Err(invalid("page-setup parser requires a worksheet root"));
                    }
                    worksheet_namespace =
                        Some(spreadsheet_namespace(&namespace).ok_or_else(|| {
                            invalid("page-setup parser requires a worksheet root")
                        })?);
                } else if depth == 2 && element.local_name().as_ref() == b"pageSetup" {
                    let expected = worksheet_namespace
                        .ok_or_else(|| invalid("missing worksheet namespace"))?;
                    if spreadsheet(&namespace) && !exact(&namespace, expected) {
                        return Err(invalid(
                            "pageSetup namespace does not match worksheet conformance",
                        ));
                    }
                    if exact(&namespace, expected) {
                        if result.is_some() || open.is_some() {
                            return Err(invalid("duplicate worksheet pageSetup element"));
                        }
                        open = Some((
                            depth,
                            parse_setup(
                                &element,
                                decoder,
                                resolver,
                                relationship_namespace(expected)?,
                                projection,
                            )?,
                        ));
                    }
                } else if open.is_some() {
                    return Err(invalid("pageSetup must not contain child elements"));
                }
            },
            Event::Empty(element) => {
                if depth == 1 && element.local_name().as_ref() == b"pageSetup" {
                    let expected = worksheet_namespace
                        .ok_or_else(|| invalid("missing worksheet namespace"))?;
                    if spreadsheet(&namespace) && !exact(&namespace, expected) {
                        return Err(invalid(
                            "pageSetup namespace does not match worksheet conformance",
                        ));
                    }
                    if exact(&namespace, expected) {
                        if result.is_some() || open.is_some() {
                            return Err(invalid("duplicate worksheet pageSetup element"));
                        }
                        result = Some(parse_setup(
                            &element,
                            decoder,
                            resolver,
                            relationship_namespace(expected)?,
                            projection,
                        )?);
                    }
                } else if open.is_some() {
                    return Err(invalid("pageSetup must not contain child elements"));
                }
            },
            Event::Text(text) => {
                if open.is_some() && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("pageSetup must not contain text"));
                }
            },
            Event::CData(_) if open.is_some() => {
                return Err(invalid("pageSetup must not contain CDATA"));
            },
            Event::End(element) => {
                let closes_page_setup = open
                    .as_ref()
                    .is_some_and(|(element_depth, _)| *element_depth == depth);
                if closes_page_setup && let Some((_, setup)) = open.take() {
                    result = Some(setup);
                }
                if depth == 1 {
                    let expected = worksheet_namespace
                        .ok_or_else(|| invalid("missing worksheet namespace"))?;
                    if !exact(&namespace, expected) || element.local_name().as_ref() != b"worksheet"
                    {
                        return Err(invalid("invalid worksheet closing element"));
                    }
                    root_closed = true;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unexpected XML end element"))?;
            },
            Event::GeneralRef(reference) => {
                if reference.resolve_char_ref().map_err(xml_error)?.is_none()
                    && !matches!(
                        reference.decode().map_err(xml_error)?.as_ref(),
                        "amp" | "lt" | "gt" | "apos" | "quot"
                    )
                {
                    return Err(invalid("custom XML entities are rejected"));
                }
                if open.is_some() {
                    return Err(invalid("pageSetup must not contain entity text"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) | Event::CData(_) => {},
            Event::Eof => break,
        }
    }
    if worksheet_namespace.is_none() || !root_closed || depth != 0 || open.is_some() {
        return Err(invalid("incomplete worksheet page-setup XML"));
    }
    Ok(result)
}

fn parse_setup(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    relationship_namespace: &[u8],
    projection: Projection,
) -> Result<ParsedSetup> {
    let mut setup = Setup::default();
    let mut printer_settings = None;
    let mut seen = [false; 18];
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let qualified_name = attribute.key.as_ref();
        if qualified_name == b"xmlns" || qualified_name.starts_with(b"xmlns:") {
            continue;
        }
        let local_name = attribute.key.local_name();
        match resolver.resolve_attribute(attribute.key).0 {
            ResolveResult::Unbound => {},
            ResolveResult::Bound(Namespace(namespace))
                if namespace == relationship_namespace && local_name.as_ref() == b"id" =>
            {
                if printer_settings.is_some() {
                    return Err(invalid("duplicate pageSetup relationship ID"));
                }
                let value = attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(xml_error)?;
                printer_settings = Some(
                    RelId::new(value.into_owned()).map_err(|error| invalid(error.to_string()))?,
                );
                continue;
            },
            ResolveResult::Bound(_) => {
                return Err(invalid(format!(
                    "unexpected qualified pageSetup attribute '{}'",
                    String::from_utf8_lossy(qualified_name)
                )));
            },
            ResolveResult::Unknown(prefix) => {
                return Err(invalid(format!(
                    "undeclared pageSetup attribute prefix '{}'",
                    String::from_utf8_lossy(prefix.as_ref())
                )));
            },
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        if projection == Projection::Relationship {
            if is_setup_attribute(local_name.as_ref()) {
                continue;
            }
            return Err(invalid(format!(
                "unknown pageSetup attribute '{}'",
                String::from_utf8_lossy(local_name.as_ref())
            )));
        }
        let slot = match local_name.as_ref() {
            b"paperSize" => {
                setup.paper = Some(
                    Paper::try_from(parse_u32(&value, "paperSize")?)
                        .map_err(|error| invalid(error.to_string()))?,
                );
                0
            },
            b"paperWidth" => {
                setup.paper_width = Some(parse_measure(&value, "paperWidth")?);
                1
            },
            b"paperHeight" => {
                setup.paper_height = Some(parse_measure(&value, "paperHeight")?);
                2
            },
            b"scale" => {
                setup.scale = Some(
                    Scale::try_from(parse_u32(&value, "scale")?)
                        .map_err(|error| invalid(error.to_string()))?,
                );
                3
            },
            b"firstPageNumber" => {
                setup.first_page = Some(
                    FirstPage::from_wire(parse_u32(&value, "firstPageNumber")?)
                        .map_err(|error| invalid(error.to_string()))?,
                );
                4
            },
            b"fitToWidth" => {
                setup.fit_to_width = Some(
                    Fit::try_from(parse_u32(&value, "fitToWidth")?)
                        .map_err(|error| invalid(error.to_string()))?,
                );
                5
            },
            b"fitToHeight" => {
                setup.fit_to_height = Some(
                    Fit::try_from(parse_u32(&value, "fitToHeight")?)
                        .map_err(|error| invalid(error.to_string()))?,
                );
                6
            },
            b"pageOrder" => {
                setup.order = Some(parse_order(&value)?);
                7
            },
            b"orientation" => {
                setup.orientation = Some(parse_orientation(&value)?);
                8
            },
            b"usePrinterDefaults" => {
                setup.use_printer_defaults = Some(parse_bool(&value, "usePrinterDefaults")?);
                9
            },
            b"blackAndWhite" => {
                setup.black_and_white = Some(parse_bool(&value, "blackAndWhite")?);
                10
            },
            b"draft" => {
                setup.draft = Some(parse_bool(&value, "draft")?);
                11
            },
            b"cellComments" => {
                setup.comments = Some(parse_comments(&value)?);
                12
            },
            b"useFirstPageNumber" => {
                setup.use_first_page = Some(parse_bool(&value, "useFirstPageNumber")?);
                13
            },
            b"errors" => {
                setup.errors = Some(parse_errors(&value)?);
                14
            },
            b"horizontalDpi" => {
                setup.horizontal_dpi = Some(
                    Dpi::new(parse_u32(&value, "horizontalDpi")?)
                        .map_err(|error| invalid(error.to_string()))?,
                );
                15
            },
            b"verticalDpi" => {
                setup.vertical_dpi = Some(
                    Dpi::new(parse_u32(&value, "verticalDpi")?)
                        .map_err(|error| invalid(error.to_string()))?,
                );
                16
            },
            b"copies" => {
                setup.copies = Some(
                    Copies::try_from(parse_u32(&value, "copies")?)
                        .map_err(|error| invalid(error.to_string()))?,
                );
                17
            },
            name => {
                return Err(invalid(format!(
                    "unknown pageSetup attribute '{}'",
                    String::from_utf8_lossy(name)
                )));
            },
        };
        if seen[slot] {
            return Err(invalid("duplicate pageSetup attribute"));
        }
        seen[slot] = true;
    }
    Ok(ParsedSetup {
        setup,
        printer_settings,
    })
}

fn is_setup_attribute(name: &[u8]) -> bool {
    matches!(
        name,
        b"paperSize"
            | b"paperWidth"
            | b"paperHeight"
            | b"scale"
            | b"firstPageNumber"
            | b"fitToWidth"
            | b"fitToHeight"
            | b"pageOrder"
            | b"orientation"
            | b"usePrinterDefaults"
            | b"blackAndWhite"
            | b"draft"
            | b"cellComments"
            | b"useFirstPageNumber"
            | b"errors"
            | b"horizontalDpi"
            | b"verticalDpi"
            | b"copies"
    )
}

fn parse_measure(raw: &str, field: &str) -> Result<Measure> {
    if raw.len() > MAX_MEASURE_BYTES {
        return Err(invalid(format!("{field} measure is too long")));
    }
    let boundary = raw
        .len()
        .checked_sub(2)
        .ok_or_else(|| invalid(format!("invalid {field} unit")))?;
    let number = raw
        .get(..boundary)
        .ok_or_else(|| invalid(format!("invalid {field} unit")))?;
    let suffix = raw
        .get(boundary..)
        .ok_or_else(|| invalid(format!("invalid {field} unit")))?;
    let unit = suffix
        .parse()
        .map_err(|_| invalid(format!("invalid {field} unit")))?;
    Measure::new(number, unit).map_err(|_| invalid(format!("invalid {field} measure")))
}

fn parse_u32(raw: &str, field: &str) -> Result<u32> {
    raw.parse()
        .map_err(|_| invalid(format!("invalid pageSetup {field}")))
}
fn parse_bool(raw: &str, field: &str) -> Result<bool> {
    match raw {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!("invalid pageSetup {field} boolean"))),
    }
}
fn parse_orientation(raw: &str) -> Result<Orientation> {
    raw.parse()
        .map_err(|_| invalid("invalid pageSetup orientation"))
}
fn parse_order(raw: &str) -> Result<Order> {
    raw.parse()
        .map_err(|_| invalid("invalid pageSetup pageOrder"))
}
fn parse_comments(raw: &str) -> Result<Comments> {
    raw.parse()
        .map_err(|_| invalid("invalid pageSetup cellComments"))
}
fn parse_errors(raw: &str) -> Result<ErrorMode> {
    raw.parse().map_err(|_| invalid("invalid pageSetup errors"))
}

fn spreadsheet(namespace: &ResolveResult<'_>) -> bool {
    spreadsheet_namespace(namespace).is_some()
}
fn spreadsheet_namespace(namespace: &ResolveResult<'_>) -> Option<&'static [u8]> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == CORE => Some(CORE),
        ResolveResult::Bound(Namespace(value)) if *value == STRICT => Some(STRICT),
        _ => None,
    }
}
fn relationship_namespace(namespace: &[u8]) -> Result<&'static [u8]> {
    match namespace {
        CORE => Ok(REL),
        STRICT => Ok(STRICT_REL),
        _ => Err(invalid(
            "pageSetup namespace does not select a relationship dialect",
        )),
    }
}
fn exact(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == expected)
}
fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{OpcPackage, PackURI};

    const START: &str = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#;

    fn parse(body: &str) -> Result<Option<Setup>> {
        parse_worksheet_page_setup(format!("{START}{body}</worksheet>").as_bytes())
    }

    fn parse_fixture(path: &str) -> Result<Setup> {
        let package = OpcPackage::open(path).unwrap();
        let uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
        parse_worksheet_page_setup(package.get_part(&uri).unwrap().blob())?
            .ok_or_else(|| invalid("fixture has no pageSetup"))
    }

    #[test]
    fn closed_tokens_have_one_from_str_and_display_mapping() {
        for (token, value) in [
            ("default", Orientation::Default),
            ("portrait", Orientation::Portrait),
            ("landscape", Orientation::Landscape),
        ] {
            assert_eq!(token.parse(), Ok(value));
            assert_eq!(value.to_string(), token);
        }
        for (token, value) in [
            ("downThenOver", Order::DownThenOver),
            ("overThenDown", Order::OverThenDown),
        ] {
            assert_eq!(token.parse(), Ok(value));
            assert_eq!(value.to_string(), token);
        }
        for (token, value) in [
            ("none", Comments::None),
            ("asDisplayed", Comments::AsDisplayed),
            ("atEnd", Comments::AtEnd),
        ] {
            assert_eq!(token.parse(), Ok(value));
            assert_eq!(value.to_string(), token);
        }
        for (token, value) in [
            ("displayed", ErrorMode::Displayed),
            ("blank", ErrorMode::Blank),
            ("dash", ErrorMode::Dash),
            ("NA", ErrorMode::NotAvailable),
        ] {
            assert_eq!(token.parse(), Ok(value));
            assert_eq!(value.to_string(), token);
        }
        for (token, value) in [
            ("mm", Unit::Millimeter),
            ("cm", Unit::Centimeter),
            ("in", Unit::Inch),
            ("pt", Unit::Point),
            ("pc", Unit::Pica),
            ("pi", Unit::PicaAlternative),
        ] {
            assert_eq!(token.parse(), Ok(value));
            assert_eq!(value.to_string(), token);
        }

        assert!("Landscape".parse::<Orientation>().is_err());
        assert!("na".parse::<ErrorMode>().is_err());
        assert!("px".parse::<Unit>().is_err());
    }

    #[test]
    fn parses_every_attribute_without_erasing_lexical_values() {
        let setup = parse(r#"<pageSetup paperSize="9" paperWidth="21cm" paperHeight="297mm" scale="125" firstPageNumber="3" fitToWidth="2" fitToHeight="4" pageOrder="overThenDown" orientation="landscape" usePrinterDefaults="1" blackAndWhite="true" draft="1" cellComments="atEnd" useFirstPageNumber="true" errors="NA" horizontalDpi="1200" verticalDpi="600" copies="2" r:id="rId7"/>"#).unwrap().unwrap();
        assert_eq!(
            setup,
            Setup {
                paper: Some(Paper::A4),
                paper_width: Some(Measure::new("21", Unit::Centimeter).unwrap()),
                paper_height: Some(Measure::new("297", Unit::Millimeter).unwrap()),
                scale: Some(Scale::new(125).unwrap()),
                first_page: Some(FirstPage::new(3).unwrap()),
                fit_to_width: Some(Fit::new(2).unwrap()),
                fit_to_height: Some(Fit::new(4).unwrap()),
                order: Some(Order::OverThenDown),
                orientation: Some(Orientation::Landscape),
                use_printer_defaults: Some(true),
                black_and_white: Some(true),
                draft: Some(true),
                comments: Some(Comments::AtEnd),
                use_first_page: Some(true),
                errors: Some(ErrorMode::NotAvailable),
                horizontal_dpi: Some(Dpi::new(1_200).unwrap()),
                vertical_dpi: Some(Dpi::new(600).unwrap()),
                copies: Some(Copies::new(2).unwrap()),
            }
        );
        let document = format!(r#"{START}<pageSetup r:id="rId7"/></worksheet>"#);
        assert_eq!(
            parse_worksheet_page_setup_relationship_id(document.as_bytes())
                .unwrap()
                .as_ref()
                .map(RelId::as_str),
            Some("rId7")
        );
    }

    #[test]
    fn preserves_absence_and_resolves_the_printer_default_only_on_request() {
        let absent = parse("<pageSetup/>").unwrap().unwrap();
        assert_eq!(absent, Setup::default());
        assert!(absent.uses_printer_defaults());

        let explicit_false = parse(r#"<pageSetup usePrinterDefaults="0"/>"#)
            .unwrap()
            .unwrap();
        assert_eq!(explicit_false.use_printer_defaults, Some(false));
        assert!(!explicit_false.uses_printer_defaults());
        assert_ne!(absent, explicit_false);
        assert!(parse("").unwrap().is_none());
    }

    #[test]
    fn checked_numeric_types_cover_excel_boundaries() {
        assert_eq!(Scale::new(0), Ok(Scale::AUTO));
        assert_eq!(Scale::new(10).unwrap().get(), 10);
        assert_eq!(Scale::new(400).unwrap().get(), 400);
        assert!(Scale::new(9).is_err());
        assert!(Scale::new(401).is_err());
        assert_eq!(Fit::new(0), Ok(Fit::NONE));
        assert_eq!(Fit::new(Fit::MAX).unwrap().get(), Fit::MAX);
        assert!(Fit::new(Fit::MAX + 1).is_err());

        for value in [1, 47, 50, 118, 256, 2_147_483_647, Paper::MAX] {
            assert_eq!(Paper::new(value).unwrap().get(), value, "paper {value}");
        }
        assert_eq!(Paper::try_from(9), Ok(Paper::A4));
        assert!(Paper::try_from(256).unwrap().is_custom());
        assert!(!Paper::A4.is_custom());

        for value in [1, 9, 401, u32::MAX] {
            assert!(Scale::try_from(value).is_err(), "scale {value}");
        }
        for value in [32_768, u32::MAX] {
            assert!(Fit::try_from(value).is_err(), "fit {value}");
        }
        for value in [0, 48, 49, 119, 255] {
            assert!(Paper::try_from(value).is_err(), "paper {value}");
        }

        assert_eq!(Copies::new(1), Ok(Copies::ONE));
        assert_eq!(Copies::new(Copies::MAX).unwrap().get(), Copies::MAX);
        assert!(Copies::new(0).is_err());
        assert!(Copies::try_from(32_768).is_err());

        assert_eq!(Dpi::new(1).unwrap().get(), 1);
        assert_eq!(Dpi::new(u32::MAX).unwrap().get(), u32::MAX);
        assert!(Dpi::new(0).is_err());
    }

    #[test]
    fn signed_first_page_round_trips_through_the_unsigned_wire_domain() {
        let minimum = FirstPage::new(-32_767).unwrap();
        assert_eq!(minimum.wire(), 4_294_934_529);
        assert_eq!(FirstPage::from_wire(minimum.wire()), Ok(minimum));
        assert_eq!(FirstPage::new(-1).unwrap().wire(), u32::MAX);
        assert_eq!(FirstPage::from_wire(u32::MAX).unwrap().get(), -1);
        assert!(FirstPage::new(0).is_err());
        assert!(FirstPage::new(-32_768).is_err());
        assert!(FirstPage::new(32_768).is_err());
        assert!(FirstPage::from_wire(0).is_err());
        assert!(FirstPage::from_wire(32_768).is_err());
        assert!(FirstPage::from_wire(4_294_934_528).is_err());
        assert!(parse(r#"<pageSetup firstPageNumber="0"/>"#).is_err());

        let parsed = parse(r#"<pageSetup firstPageNumber="4294934529"/>"#)
            .unwrap()
            .unwrap();
        assert_eq!(parsed.first_page, Some(minimum));
    }

    #[test]
    fn exact_measures_reject_non_schema_numbers_and_oversized_lexicals() {
        let exact = Measure::new("00.50", Unit::Centimeter).unwrap();
        assert_eq!(exact.decimal(), "00.50");
        assert_eq!(exact.unit(), Unit::Centimeter);
        assert_eq!(exact.to_string(), "00.50cm");
        assert_eq!(
            Measure::new("0", Unit::Millimeter).unwrap().to_string(),
            "0mm"
        );

        for decimal in ["+1", "1e2", ".5", "1.", "-1", ""] {
            assert!(
                Measure::new(decimal, Unit::Millimeter).is_err(),
                "decimal {decimal}"
            );
        }

        let maximum = "1".repeat(MAX_MEASURE_BYTES - Unit::Millimeter.as_str().len());
        assert!(Measure::new(maximum, Unit::Millimeter).is_ok());
        let oversized = "1".repeat(MAX_MEASURE_BYTES - Unit::Millimeter.as_str().len() + 1);
        assert!(Measure::new(oversized.as_str(), Unit::Millimeter).is_err());
        assert!(Measure::new(oversized.clone(), Unit::Millimeter).is_err());
        assert!(parse(&format!(r#"<pageSetup paperWidth="{oversized}mm"/>"#)).is_err());

        let parsed = parse(r#"<pageSetup paperWidth="00.50cm" paperHeight="0mm"/>"#)
            .unwrap()
            .unwrap();
        assert_eq!(parsed.paper_width, Some(exact));
        assert_eq!(
            parsed.paper_height,
            Some(Measure::new("0", Unit::Millimeter).unwrap())
        );
    }

    #[test]
    fn relationship_ids_validate_before_borrowed_or_owned_storage() {
        assert_eq!(RelId::new("rId7").unwrap().as_str(), "rId7");
        assert_eq!(
            RelId::new(String::from("printerSettings1"))
                .unwrap()
                .as_str(),
            "printerSettings1"
        );

        let oversized = "r".repeat(MAX_RELATIONSHIP_ID_BYTES + 1);
        assert!(RelId::new(oversized.as_str()).is_err());
        assert!(RelId::new(oversized).is_err());
    }

    #[test]
    fn relationship_projection_is_independent_from_nonconforming_settings() {
        let xml = format!(r#"{START}<pageSetup horizontalDpi="0" r:id="rIdPrinter"/></worksheet>"#);
        assert!(parse_worksheet_page_setup(xml.as_bytes()).is_err());
        assert_eq!(
            parse_worksheet_page_setup_relationship_id(xml.as_bytes())
                .unwrap()
                .unwrap()
                .as_str(),
            "rIdPrinter"
        );
    }

    #[test]
    fn parser_rejects_every_out_of_domain_numeric_family() {
        let automatic = parse(r#"<pageSetup scale="0" fitToWidth="0" fitToHeight="32767"/>"#)
            .unwrap()
            .unwrap();
        assert!(automatic.scale.unwrap().is_auto());
        assert!(automatic.fit_to_width.unwrap().is_unbounded());
        assert_eq!(automatic.fit_to_height.unwrap().get(), 32_767);

        assert!(parse(r#"<pageSetup scale="9"/>"#).is_err());
        assert!(parse(r#"<pageSetup scale="401"/>"#).is_err());
        assert!(parse(r#"<pageSetup fitToWidth="32768"/>"#).is_err());
        assert!(parse(r#"<pageSetup paperSize="0"/>"#).is_err());
        assert!(parse(r#"<pageSetup paperSize="48"/>"#).is_err());
        assert!(parse(r#"<pageSetup paperSize="119"/>"#).is_err());
        assert!(parse(r#"<pageSetup horizontalDpi="0"/>"#).is_err());
        assert!(parse(r#"<pageSetup verticalDpi="0"/>"#).is_err());
        assert!(parse(r#"<pageSetup copies="0"/>"#).is_err());
        assert!(parse(r#"<pageSetup copies="32768"/>"#).is_err());
    }

    #[test]
    fn rejects_bad_enums_measures_and_content() {
        assert!(parse(r#"<pageSetup orientation="sideways"/>"#).is_err());
        assert!(parse(r#"<pageSetup paperWidth="+1mm"/>"#).is_err());
        assert!(parse(r#"<pageSetup paperWidth="1e2mm"/>"#).is_err());
        assert!(parse(r#"<pageSetup paperWidth=".5mm"/>"#).is_err());
        assert!(parse(r#"<pageSetup paperWidth="1.mm"/>"#).is_err());
        assert!(parse(r#"<pageSetup errors="na"/>"#).is_err());
        assert!(parse(r#"<pageSetup r:id="1bad"/>"#).is_err());
        assert!(parse(r#"<pageSetup><x/></pageSetup>"#).is_err());
    }

    #[test]
    fn page_setup_attributes_use_expanded_names_strictly() {
        let local_default = parse(&format!(
            r#"<pageSetup xmlns="{}" xmlns:unused="urn:unused" orientation="landscape"/>"#,
            String::from_utf8_lossy(CORE)
        ))
        .unwrap()
        .unwrap();
        assert_eq!(local_default.orientation, Some(Orientation::Landscape));

        let aliased_relationship = format!(
            r#"<worksheet xmlns="{}"><pageSetup xmlns:rels="{}" rels:id="rIdAlias"/></worksheet>"#,
            String::from_utf8_lossy(CORE),
            String::from_utf8_lossy(REL)
        );
        assert_eq!(
            parse_worksheet_page_setup_relationship_id(aliased_relationship.as_bytes())
                .unwrap()
                .unwrap()
                .as_str(),
            "rIdAlias"
        );
        assert!(
            parse(r#"<pageSetup xmlns:rels="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" rels:id="rId2"/>"#)
                .is_err()
        );

        let undeclared_relationship = format!(
            r#"<worksheet xmlns="{}"><pageSetup r:id="rId1"/></worksheet>"#,
            String::from_utf8_lossy(CORE)
        );
        assert!(parse_worksheet_page_setup(undeclared_relationship.as_bytes()).is_err());
        assert!(parse(r#"<pageSetup xmlns:v="urn:vendor" v:mode="x"/>"#).is_err());
        assert!(parse(r#"<pageSetup id="rId1"/>"#).is_err());
        assert!(parse(r#"<pageSetup r:target="rId1"/>"#).is_err());
        assert!(
            parse(&format!(
                r#"<pageSetup xmlns:s="{}" s:orientation="landscape"/>"#,
                String::from_utf8_lossy(CORE)
            ))
            .is_err()
        );

        let mismatched_relationship = format!(
            r#"<worksheet xmlns="{}" xmlns:r="{}"><pageSetup r:id="rId1"/></worksheet>"#,
            String::from_utf8_lossy(CORE),
            String::from_utf8_lossy(STRICT_REL)
        );
        assert!(
            parse_worksheet_page_setup_relationship_id(mismatched_relationship.as_bytes()).is_err()
        );

        let mismatched_element = format!(
            r#"<worksheet xmlns="{}" xmlns:s="{}" xmlns:r="{}"><s:pageSetup r:id="rId1"/></worksheet>"#,
            String::from_utf8_lossy(CORE),
            String::from_utf8_lossy(STRICT),
            String::from_utf8_lossy(STRICT_REL)
        );
        assert!(parse_worksheet_page_setup(mismatched_element.as_bytes()).is_err());

        let strict = format!(
            r#"<s:worksheet xmlns:s="{}" xmlns:r="{}"><s:pageSetup r:id="rId9"/></s:worksheet>"#,
            String::from_utf8_lossy(STRICT),
            String::from_utf8_lossy(STRICT_REL)
        );
        assert_eq!(
            parse_worksheet_page_setup_relationship_id(strict.as_bytes())
                .unwrap()
                .unwrap()
                .as_str(),
            "rId9"
        );

        let strict_with_transitional_relationship = format!(
            r#"<s:worksheet xmlns:s="{}" xmlns:r="{}"><s:pageSetup r:id="rId9"/></s:worksheet>"#,
            String::from_utf8_lossy(STRICT),
            String::from_utf8_lossy(REL)
        );
        assert!(
            parse_worksheet_page_setup_relationship_id(
                strict_with_transitional_relationship.as_bytes()
            )
            .is_err()
        );
    }

    #[test]
    fn loads_poi_resolution_and_relationship_fixture() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/poi/test-data/spreadsheet/45540_classic_Header.xlsx"
        );
        let setup = parse_fixture(path).unwrap();
        assert_eq!(setup.orientation, Some(Orientation::Portrait));
        assert_eq!(setup.horizontal_dpi, Some(Dpi::new(1_200).unwrap()));
        assert_eq!(setup.vertical_dpi, Some(Dpi::new(1_200).unwrap()));
        let package = OpcPackage::open(path).unwrap();
        let uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
        assert_eq!(
            parse_worksheet_page_setup_relationship_id(package.get_part(&uri).unwrap().blob())
                .unwrap()
                .as_ref()
                .map(RelId::as_str),
            Some("rId1")
        );
    }

    #[test]
    fn rejects_a_libreoffice_fixture_with_nonconforming_zero_dpi() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf136721_letter_sized_paper.xlsx"
        );
        let error = parse_fixture(path).unwrap_err();
        assert!(error.to_string().contains("DPI"));
    }
}
