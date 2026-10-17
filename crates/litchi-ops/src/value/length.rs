//! Lengths in English Metric Units, read and printed with OfficeCLI's unit
//! rules (`OC/Core/EmuConverter.cs`, `ParseHelpers.ParseFontSize`).

use std::fmt::{self, Write as _};
use std::str::FromStr;

use super::error::{ErrorKind, ValueError, What};
use super::number::{Decimal, Styles, parse_f64, parse_i64};
use super::text::{StackStr, strip_suffix_ci};

const EMU_PER_INCH: i64 = 914_400;
const EMU_PER_CM: i64 = 360_000;
const EMU_PER_POINT: i64 = 12_700;

/// A length in English Metric Units (EMU), the OOXML drawing unit:
/// 914 400 per inch, 360 000 per centimetre, 36 000 per millimetre, 12 700
/// per point, 9 525 per pixel at 96 DPI.
///
/// [`Length::parse_in`] reads OfficeCLI's unit vocabulary (`emu`, `mm`,
/// `cm`, `in`, `pc`, `pt`, `px`, `Q`) under a [`LengthContext`];
/// [`Length::display`] prints it the way OfficeCLI prints the property
/// (a [`LengthStyle`]). [`Display`](fmt::Display) uses
/// [`LengthStyle::Auto`].
///
/// ```
/// use litchi_ops::value::{Length, LengthContext, LengthStyle};
///
/// let x = Length::parse_in("2.54cm", LengthContext::Emu)?;
/// assert_eq!(x, Length::from_emu(914_400));
/// assert_eq!(x.to_string(), "2.54cm");
/// assert_eq!(Length::from_emu(12_192_000).to_string(), "960pt");
///
/// // A bare number is EMU, except for line widths, where it is points.
/// assert_eq!(Length::parse_in("12700", LengthContext::Emu)?.emu(), 12_700);
/// assert_eq!(Length::parse_in("1", LengthContext::LineWidth)?.emu(), 12_700);
/// assert_eq!(Length::from_points(11).display(LengthStyle::Points).to_string(), "11pt");
/// # Ok::<(), litchi_ops::Error>(())
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub struct Length(i64);

/// How a bare number and a value's range are read: the parse half of a
/// property's length type.
///
/// ```
/// use litchi_ops::value::{Length, LengthContext};
///
/// assert_eq!(Length::parse_in("2", LengthContext::Emu)?.emu(), 2);
/// assert_eq!(Length::parse_in("2", LengthContext::LineWidth)?.emu(), 25_400);
/// assert_eq!(Length::parse_in("2", LengthContext::FontSize)?.emu(), 25_400);
/// assert!(Length::parse_in("-2", LengthContext::Extent).is_err());
/// # Ok::<(), litchi_ops::Error>(())
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum LengthContext {
    /// Positions and offsets (OfficeCLI `EmuConverter.ParseEmu`): a bare
    /// whole number (or `1e6`-style scientific number) is EMU; any unit;
    /// negative values allowed.
    #[default]
    Emu,
    /// Extents such as `width` and `height` (`ParseEmuAsInt`): as
    /// [`LengthContext::Emu`], but non-negative and at most 2 147 483 647
    /// EMU.
    Extent,
    /// Line widths (`ParseLineWidth`): a bare number, decimals included, is
    /// points; non-negative; at most 2 147 483 647 EMU.
    LineWidth,
    /// Font sizes (`ParseFontSize`): points, with an optional `pt` suffix and
    /// no other unit; from 0.5 pt to 4000 pt.
    FontSize,
    /// Nonnegative Word paragraph spacing: bare values are twips; pt/cm/in
    /// are converted and rounded ties-to-even to unsigned 32-bit twips.
    /// Group separators are currently refused (record 0047).
    WordSpacing,
}

/// How a property prints its length: the display half of a property's
/// length type. Each style is one of OfficeCLI's formatters.
///
/// ```
/// use litchi_ops::value::{Length, LengthStyle};
///
/// let inch = Length::from_emu(914_400);
/// assert_eq!(inch.display(LengthStyle::Auto).to_string(), "2.54cm");
/// assert_eq!(inch.display(LengthStyle::PointsFirst).to_string(), "72pt");
/// assert_eq!(Length::from_emu(180_000).display(LengthStyle::LineWidth).to_string(), "14.17pt");
/// assert_eq!(Length::from_emu(2_640_000).display(LengthStyle::Lossy).to_string(), "7.33cm");
/// assert_eq!(Length::from_half_points(21).display(LengthStyle::Points).to_string(), "10.5pt");
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum LengthStyle {
    /// `EmuConverter.FormatEmu`: the first of centimetres, points and inches
    /// whose two-decimal form re-reads as the exact EMU value, else
    /// `<n>emu`; zero is `0cm`. Positions and extents.
    #[default]
    Auto,
    /// `FormatEmuPtFirst`: points, then centimetres, then inches, else
    /// `<n>emu`; zero is `0pt`. Slide and page sizes.
    PointsFirst,
    /// `FormatEmuLossy`: as [`LengthStyle::Auto`], but falling back to
    /// centimetres rounded to two decimals. Derived widths.
    Lossy,
    /// `FormatLineWidth`: points rounded to two decimals, or `<n>emu` when
    /// that rounds to zero; zero is `0pt`.
    LineWidth,
    /// Points rounded to two decimals (`11pt`, `10.5pt`): font sizes and
    /// paragraph spacing.
    Points,
}

/// A property's length type: how it is read and how it is printed.
///
/// ```
/// use litchi_ops::value::{LengthContext, LengthSpec, LengthStyle};
///
/// assert_eq!(LengthSpec::LINE_WIDTH, LengthSpec::new(LengthContext::LineWidth, LengthStyle::LineWidth));
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct LengthSpec {
    /// How text is read.
    pub context: LengthContext,
    /// How the value is printed.
    pub style: LengthStyle,
}

impl LengthSpec {
    /// Positions and offsets (`x`, `y`): [`LengthContext::Emu`],
    /// [`LengthStyle::Auto`].
    pub const POSITION: Self = Self::new(LengthContext::Emu, LengthStyle::Auto);
    /// Extents (`width`, `height`): [`LengthContext::Extent`],
    /// [`LengthStyle::Auto`].
    pub const EXTENT: Self = Self::new(LengthContext::Extent, LengthStyle::Auto);
    /// Line widths: [`LengthContext::LineWidth`], [`LengthStyle::LineWidth`].
    pub const LINE_WIDTH: Self = Self::new(LengthContext::LineWidth, LengthStyle::LineWidth);
    /// Font sizes: [`LengthContext::FontSize`], [`LengthStyle::Points`].
    pub const FONT_SIZE: Self = Self::new(LengthContext::FontSize, LengthStyle::Points);

    /// A length type that reads in `context` and prints in `style`.
    #[must_use]
    pub const fn new(context: LengthContext, style: LengthStyle) -> Self {
        Self { context, style }
    }
}

impl Length {
    /// Zero.
    pub const ZERO: Self = Self(0);

    /// A length of `emu` English Metric Units.
    #[must_use]
    pub const fn from_emu(emu: i64) -> Self {
        Self(emu)
    }

    /// A length of whole points (12 700 EMU each).
    #[must_use]
    pub const fn from_points(points: i32) -> Self {
        Self(points as i64 * EMU_PER_POINT)
    }

    /// A length of hundredths of a point (127 EMU each), the unit of
    /// DrawingML font sizes and spacing.
    #[must_use]
    pub const fn from_centipoints(centipoints: i32) -> Self {
        Self(centipoints as i64 * 127)
    }

    /// A length of half-points (6 350 EMU each), the unit of
    /// WordprocessingML font sizes.
    #[must_use]
    pub const fn from_half_points(half_points: i32) -> Self {
        Self(half_points as i64 * 6_350)
    }

    /// A length of twips, twentieths of a point (635 EMU each), the unit of
    /// WordprocessingML spacing and indents.
    #[must_use]
    pub const fn from_twips(twips: i32) -> Self {
        Self(twips as i64 * 635)
    }

    /// Convert this length to a checked signed 32-bit twip value.
    ///
    /// Native Word paragraph properties use twips while the unified value
    /// layer stores lengths as EMU.  Conversion rounds half away from zero,
    /// matching the existing Word spacing setter, and refuses values outside
    /// the native signed range instead of truncating them.
    #[cfg(feature = "rtf")]
    pub(crate) fn checked_twips(self) -> Option<i32> {
        let adjusted = if self.0 >= 0 {
            self.0.checked_add(317)?
        } else {
            self.0.checked_sub(317)?
        };
        i32::try_from(adjusted.checked_div(635)?).ok()
    }

    /// The length in EMU.
    #[must_use]
    pub const fn emu(self) -> i64 {
        self.0
    }

    /// The length in points, as a float.
    #[must_use]
    #[allow(
        clippy::cast_precision_loss,
        reason = "EMU values beyond 2^53 lose precision, as OfficeCLI's double arithmetic does"
    )]
    pub fn points(self) -> f64 {
        self.0 as f64 / EMU_PER_POINT as f64
    }

    /// Reads `text` as a position or offset ([`LengthContext::Emu`]).
    ///
    /// # Errors
    ///
    /// As [`Length::parse_in`].
    pub fn parse(text: &str) -> Result<Self, ValueError<'_>> {
        Self::parse_in(text, LengthContext::Emu)
    }

    /// Reads `text` under `context`.
    ///
    /// Units are matched case-insensitively at the end of the trimmed text,
    /// in OfficeCLI's order: `mm`, `cm`, `in`, `pc` (12 pt), `pt`, `px`
    /// (9 525 EMU), `Q` (a quarter millimetre), `emu`. The number before a
    /// unit is a .NET `NumberStyles.Float` number and is rounded to the
    /// nearest EMU, ties to even. Without a unit, the context decides.
    ///
    /// # Errors
    ///
    /// Returns an `invalid_value` [`ValueError`] when the text is empty, has
    /// no number before its unit, uses a font- or viewport-relative unit
    /// (`em`, `rem`, `ex`, `vw`, `vh`), is a bare decimal in an EMU context
    /// (`1.5`), is not finite, exceeds the 64-bit EMU range, or breaks the
    /// context's range (negative or above 32 bits for extents and line
    /// widths; outside 0.5 pt to 4000 pt for font sizes).
    pub fn parse_in(text: &str, context: LengthContext) -> Result<Self, ValueError<'_>> {
        match context {
            LengthContext::Emu => parse_emu(text, What::Length).map(Self),
            LengthContext::Extent => parse_extent(text, What::Length).map(Self),
            LengthContext::LineWidth => parse_line_width(text).map(Self),
            LengthContext::FontSize => parse_font_size(text).map(Self),
            LengthContext::WordSpacing => parse_word_spacing(text).map(Self),
        }
    }

    /// The length a typed value coerces to under `context` (record 0326):
    /// the value [`Length::parse_in`] reads from this length's
    /// [`LengthStyle::Auto`] text, given without writing and reading the
    /// text when the read is certain to succeed with this exact value.
    /// `None` when the text decides: the read would refuse it (out of the
    /// context's range, or a unit the context does not read) or round it
    /// (word spacing is kept in twips).
    pub(crate) const fn coerced_in(self, context: LengthContext) -> Option<Self> {
        // Below 2^50 EMU the unit text reads back exactly (see `exact`).
        const EXACT: i64 = 1 << 50;
        let emu = self.0;
        let certain = match context {
            LengthContext::Emu => -EXACT < emu && emu < EXACT,
            LengthContext::Extent | LengthContext::LineWidth => 0 <= emu && emu <= i32::MAX as i64,
            // Only the `pt` text is a font size: a whole number of
            // hundredths of a point that is not one of centimetres (which
            // `Auto` prints first), from 0.5 pt to 4000 pt.
            LengthContext::FontSize => {
                emu % (EMU_PER_POINT / 100) == 0
                    && emu % (EMU_PER_CM / 100) != 0
                    && EMU_PER_POINT / 2 <= emu
                    && emu <= 4000 * EMU_PER_POINT
            },
            LengthContext::WordSpacing => false,
        };
        if certain { Some(self) } else { None }
    }

    /// The canonical text of this length in `style`. Allocation-free.
    #[must_use]
    pub fn display(self, style: LengthStyle) -> LengthDisplay {
        let emu = self.0;
        LengthDisplay(match style {
            LengthStyle::Auto if emu == 0 => Form::Unit(Decimal::normalized(false, 0, 0), Unit::Cm),
            LengthStyle::Auto => exact(emu, Unit::Cm)
                .or_else(|| exact(emu, Unit::Pt))
                .or_else(|| exact(emu, Unit::In))
                .unwrap_or(Form::Emu(emu)),
            LengthStyle::PointsFirst if emu == 0 => {
                Form::Unit(Decimal::normalized(false, 0, 0), Unit::Pt)
            },
            LengthStyle::PointsFirst => exact(emu, Unit::Pt)
                .or_else(|| exact(emu, Unit::Cm))
                .or_else(|| exact(emu, Unit::In))
                .unwrap_or(Form::Emu(emu)),
            LengthStyle::Lossy if emu == 0 => {
                Form::Unit(Decimal::normalized(false, 0, 0), Unit::Cm)
            },
            LengthStyle::Lossy => exact(emu, Unit::Cm)
                .or_else(|| exact(emu, Unit::Pt))
                .or_else(|| exact(emu, Unit::In))
                .unwrap_or_else(|| Form::Unit(rounded(emu, Unit::Cm), Unit::Cm)),
            LengthStyle::LineWidth if emu == 0 => {
                Form::Unit(Decimal::normalized(false, 0, 0), Unit::Pt)
            },
            LengthStyle::LineWidth => {
                let points = rounded(emu, Unit::Pt);
                if points.is_zero() {
                    Form::Emu(emu)
                } else {
                    Form::Unit(points, Unit::Pt)
                }
            },
            LengthStyle::Points => Form::Unit(rounded(emu, Unit::Pt), Unit::Pt),
        })
    }

    /// Prints two lengths that belong together (a slide's width and height)
    /// in one unit when one unit is exact for both, trying points, then
    /// centimetres, then inches; otherwise each as [`LengthStyle::Auto`]
    /// (OfficeCLI `FormatEmuPaired`).
    ///
    /// ```
    /// use litchi_ops::value::Length;
    ///
    /// let (width, height) = Length::display_pair(Length::from_emu(12_192_000), Length::from_emu(6_858_000));
    /// assert_eq!((width.to_string(), height.to_string()), ("960pt".to_owned(), "540pt".to_owned()));
    /// ```
    #[must_use]
    pub fn display_pair(first: Self, second: Self) -> (LengthDisplay, LengthDisplay) {
        for unit in [Unit::Pt, Unit::Cm, Unit::In] {
            if let (Some(a), Some(b)) = (exact(first.0, unit), exact(second.0, unit)) {
                return (LengthDisplay(a), LengthDisplay(b));
            }
        }
        (
            first.display(LengthStyle::Auto),
            second.display(LengthStyle::Auto),
        )
    }
}

impl fmt::Display for Length {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.display(LengthStyle::Auto).fmt(formatter)
    }
}

impl FromStr for Length {
    type Err = crate::Error;

    /// Reads a position or offset ([`LengthContext::Emu`]).
    fn from_str(text: &str) -> Result<Self, Self::Err> {
        Self::parse(text).map_err(crate::Error::from)
    }
}

/// The canonical text of a [`Length`] in one [`LengthStyle`], computed when
/// built and written by [`Display`](fmt::Display) without allocating.
///
/// ```
/// use litchi_ops::value::{Length, LengthStyle};
///
/// let text = Length::from_emu(1_800).display(LengthStyle::Auto);
/// assert_eq!(text.to_string(), "1800emu");
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct LengthDisplay(Form);

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum Form {
    Unit(Decimal, Unit),
    Emu(i64),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum Unit {
    Cm,
    Pt,
    In,
}

impl Unit {
    const fn per_unit(self) -> i64 {
        match self {
            Self::Cm => EMU_PER_CM,
            Self::Pt => EMU_PER_POINT,
            Self::In => EMU_PER_INCH,
        }
    }

    const fn suffix(self) -> &'static str {
        match self {
            Self::Cm => "cm",
            Self::Pt => "pt",
            Self::In => "in",
        }
    }
}

/// OfficeCLI's `TryFormatUnit`: the two-decimal form in `unit` when it re-reads
/// as exactly `emu`. That needs `emu` to be a whole number of hundredths of
/// the unit, and, as OfficeCLI checks by re-parsing its text as a double, the
/// floating-point read to land on `emu`, which can fail above 2^53 EMU. Zero
/// has no exact form here; the styles print it themselves.
#[allow(
    clippy::cast_precision_loss,
    reason = "the floating-point re-read is the check being replicated"
)]
fn exact(emu: i64, unit: Unit) -> Option<Form> {
    let hundredth = unit.per_unit() / 100;
    if emu == 0 || emu % hundredth != 0 {
        return None;
    }
    // `rounded` without its division: the value is a whole number of
    // hundredths, so rounding changes nothing.
    let value = Decimal::normalized(emu < 0, u128::from((emu / hundredth).unsigned_abs()), 2);
    // Below 2^50 EMU the double re-read is off by less than half an EMU, so
    // it always lands on `emu`; only larger values need the check.
    if emu.unsigned_abs() < 1 << 50 {
        return Some(Form::Unit(value, unit));
    }
    let mut text = StackStr::<48>::new();
    write!(text, "{value}").ok()?;
    let reread = text.as_str().parse::<f64>().ok()? * unit.per_unit() as f64;
    (round_emu(reread) == Some(emu)).then_some(Form::Unit(value, unit))
}

/// `emu` in `unit`, rounded half away from zero to two decimals, as .NET's
/// `0.##` format rounds.
fn rounded(emu: i64, unit: Unit) -> Decimal {
    Decimal::from_ratio(emu, unit.per_unit(), 2).unwrap_or(Decimal::normalized(false, 0, 0))
}

impl fmt::Display for LengthDisplay {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self.0 {
            Form::Unit(value, unit) => {
                value.fmt(formatter)?;
                formatter.write_str(unit.suffix())
            },
            Form::Emu(emu) => write!(formatter, "{emu}emu"),
        }
    }
}

/// What query comparison reads from a length's canonical text, known without
/// writing the text (`value::compare`, record 0262): OfficeCLI's
/// `ExtractNumber`, `ExtractUnit` and `TryParseEmu` of the text.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) struct DisplayFacts {
    /// The number before the unit suffix; `None` for the `<n>emu` form,
    /// which `ExtractNumber` cannot read.
    pub(super) number: Option<f64>,
    /// The unit suffix (`cm`, `pt`, `in`); `None` for `<n>emu`, which
    /// `ExtractUnit` does not recognize.
    pub(super) unit: Option<DisplayUnit>,
    /// The EMU value `ParseEmu` reads the text back as.
    pub(super) emu: Option<i64>,
}

/// The unit suffix of a length's canonical text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum DisplayUnit {
    Cm,
    Pt,
    In,
}

impl Length {
    /// The comparison facts of this length's canonical text in `style`
    /// ([`LengthDisplay::facts`]). A length below 2^50 EMU that is a whole
    /// number of hundredths of the style's first such unit reads back as
    /// itself, so its facts need no rounding: the number is those
    /// hundredths over 100 (one correctly rounded division) and the EMU
    /// value is the length.
    pub(super) fn display_facts(self, style: LengthStyle) -> Option<DisplayFacts> {
        let emu = self.0;
        let exact = |unit: DisplayUnit, hundredth: i64| {
            #[allow(
                clippy::cast_precision_loss,
                reason = "below 2^50 EMU the hundredths are below 2^53, exact as a double"
            )]
            let number = (emu / hundredth) as f64 / 100.0;
            (emu % hundredth == 0).then_some(DisplayFacts {
                number: Some(number),
                unit: Some(unit),
                emu: Some(emu),
            })
        };
        if emu != 0 && emu.unsigned_abs() < 1 << 50 {
            let cm = || exact(DisplayUnit::Cm, EMU_PER_CM / 100);
            let pt = || exact(DisplayUnit::Pt, EMU_PER_POINT / 100);
            let inch = || exact(DisplayUnit::In, EMU_PER_INCH / 100);
            let found = match style {
                LengthStyle::Auto | LengthStyle::Lossy => cm().or_else(pt).or_else(inch),
                LengthStyle::PointsFirst => pt().or_else(cm).or_else(inch),
                _ => None,
            };
            match (found, style) {
                (Some(facts), _) => return Some(facts),
                (None, LengthStyle::Auto | LengthStyle::PointsFirst) => {
                    return Some(DisplayFacts {
                        number: None,
                        unit: None,
                        emu: Some(emu),
                    });
                },
                _ => {},
            }
        }
        self.display(style).facts()
    }
}

impl LengthDisplay {
    /// The comparison facts of this text, or `None` when they are only
    /// known by writing the text and reading it back (a number of more
    /// than 2^53 hundredths of a unit).
    pub(super) fn facts(self) -> Option<DisplayFacts> {
        match self.0 {
            Form::Emu(emu) => Some(DisplayFacts {
                number: None,
                unit: None,
                emu: Some(emu),
            }),
            Form::Unit(value, unit) => {
                // `ParseEmu` reads the number with the same correctly rounded
                // parse `ExtractNumber` does, then scales it by the unit's
                // size (the `UNITS` factors) and rounds to even.
                let number = value.to_f64_exact()?;
                #[allow(
                    clippy::cast_precision_loss,
                    reason = "the unit sizes are small integers, exact as doubles"
                )]
                let factor = unit.per_unit() as f64;
                Some(DisplayFacts {
                    number: Some(number),
                    unit: Some(match unit {
                        Unit::Cm => DisplayUnit::Cm,
                        Unit::Pt => DisplayUnit::Pt,
                        Unit::In => DisplayUnit::In,
                    }),
                    emu: round_emu(number * factor),
                })
            },
        }
    }
}

/// The units OfficeCLI's `ParseEmu` accepts, in its matching order, with
/// their size in EMU.
const UNITS: [(&str, &str, f64); 7] = [
    ("mm", "mm", 36_000.0),
    ("cm", "cm", 360_000.0),
    ("in", "in", 914_400.0),
    ("pc", "pc", 152_400.0),
    ("pt", "pt", 12_700.0),
    ("px", "px", 9_525.0),
    ("q", "Q", 9_000.0),
];

/// Units that need font or viewport context, longest first so `rem` is not
/// read as `em`.
const UNSUPPORTED_UNITS: [&str; 5] = ["rem", "em", "ex", "vw", "vh"];

/// 2^63 as a float: the first value outside the 64-bit range.
const TWO_POW_63: f64 = 9_223_372_036_854_775_808.0;

/// Rounds an EMU value to the nearest integer, ties to even (.NET
/// `Math.Round`), refusing values outside the 64-bit range.
#[allow(
    clippy::cast_possible_truncation,
    reason = "the value is rounded and checked against the 64-bit range first"
)]
fn round_emu(value: f64) -> Option<i64> {
    let rounded = value.round_ties_even();
    (-TWO_POW_63..TWO_POW_63)
        .contains(&rounded)
        .then_some(rounded as i64)
}

/// OfficeCLI `EmuConverter.ParseEmu`.
fn parse_emu(text: &str, what: What) -> Result<i64, ValueError<'_>> {
    let error = |kind| ValueError::new(what, kind, text);
    let value = text.trim();
    if value.is_empty() {
        return Err(error(ErrorKind::Empty));
    }
    for (suffix, unit, factor) in UNITS {
        if let Some(number) = strip_suffix_ci(value, suffix) {
            if number.trim().is_empty() {
                return Err(error(ErrorKind::MissingNumber { unit }));
            }
            let parsed = parse_f64(number, Styles::FLOAT)
                .filter(|parsed| parsed.is_finite())
                .ok_or_else(|| error(ErrorKind::BadNumber { unit }).at(number))?;
            return round_emu(parsed * factor).ok_or_else(|| error(ErrorKind::Overflow));
        }
    }
    if let Some(number) = strip_suffix_ci(value, "emu") {
        if number.trim().is_empty() {
            return Err(error(ErrorKind::MissingNumber { unit: "emu" }));
        }
        return parse_i64(number, Styles::INTEGER)
            .ok_or_else(|| error(ErrorKind::BadNumber { unit: "emu" }).at(number));
    }
    if let Some(unit) = UNSUPPORTED_UNITS
        .into_iter()
        .find(|unit| strip_suffix_ci(value, unit).is_some())
    {
        return Err(error(ErrorKind::UnsupportedUnit { unit }));
    }
    if let Some(emu) = parse_i64(value, Styles::INTEGER) {
        return Ok(emu);
    }
    // Bare scientific notation ("1e6") is EMU too; a plain decimal is not.
    if value.contains(['e', 'E']) {
        if let Some(parsed) = parse_f64(value, Styles::FLOAT) {
            if !parsed.is_finite() {
                return Err(error(ErrorKind::NotFinite));
            }
            return round_emu(parsed).ok_or_else(|| error(ErrorKind::Overflow));
        }
    }
    Err(error(ErrorKind::Syntax))
}

/// OfficeCLI `EmuConverter.ParseEmuAsInt`.
fn parse_extent(text: &str, what: What) -> Result<i64, ValueError<'_>> {
    let emu = parse_emu(text, what)?;
    check_extent(emu).map_err(|kind| ValueError::new(what, kind, text))
}

fn check_extent(emu: i64) -> Result<i64, ErrorKind> {
    if emu < 0 {
        Err(ErrorKind::Negative)
    } else if emu > i64::from(i32::MAX) {
        Err(ErrorKind::TooLarge)
    } else {
        Ok(emu)
    }
}

/// OfficeCLI `EmuConverter.ParseLineWidth`: a bare number is points.
fn parse_line_width(text: &str) -> Result<i64, ValueError<'_>> {
    let error = |kind| ValueError::new(What::LineWidth, kind, text);
    let value = text.trim();
    if value.is_empty() {
        return Err(error(ErrorKind::Empty));
    }
    let emu = match parse_f64(value, Styles::FLOAT) {
        Some(points) if !points.is_finite() => return Err(error(ErrorKind::NotFinite)),
        Some(points) => {
            round_emu(points * EMU_PER_POINT as f64).ok_or_else(|| error(ErrorKind::Overflow))?
        },
        None => parse_emu(text, What::LineWidth)?,
    };
    check_extent(emu).map_err(error)
}

/// OfficeCLI `ParseHelpers.ParseFontSize`, stored as EMU.
fn parse_font_size(text: &str) -> Result<i64, ValueError<'_>> {
    let error = |kind| ValueError::new(What::FontSize, kind, text);
    let mut value = text.trim();
    if let Some(number) = strip_suffix_ci(value, "pt") {
        value = number.trim();
    }
    if value.contains(',') {
        return Err(error(ErrorKind::Comma));
    }
    let points = match parse_f64(value, Styles::FLOAT_THOUSANDS) {
        Some(points) if points.is_finite() => points,
        Some(_) => return Err(error(ErrorKind::NotFinite)),
        None if value.is_empty() => return Err(error(ErrorKind::Empty)),
        None => return Err(error(ErrorKind::Syntax)),
    };
    if points <= 0.0 {
        return Err(error(ErrorKind::Negative));
    }
    if points < 0.5 {
        return Err(error(ErrorKind::TooSmall));
    }
    if points > 4000.0 {
        return Err(error(ErrorKind::TooLarge));
    }
    round_emu(points * EMU_PER_POINT as f64).ok_or_else(|| error(ErrorKind::Overflow))
}

/// OfficeCLI `EmuConverter.TryParseEmu`: the EMU value of `text`, or `None`.
/// Allocation-free on every path.
pub(super) fn try_emu(text: &str) -> Option<i64> {
    parse_emu(text, What::Length).ok()
}

fn parse_word_spacing(text: &str) -> Result<i64, ValueError<'_>> {
    let error = |kind| ValueError::new(What::WordSpacing, kind, text);
    let input = text.trim();
    let (number, factor) = if let Some(number) = strip_suffix_ci(input, "pt") {
        (number, 20.0)
    } else if let Some(number) = strip_suffix_ci(input, "cm") {
        (number, 1440.0 / 2.54)
    } else if let Some(number) = strip_suffix_ci(input, "in") {
        (number, 1440.0)
    } else {
        (input, 1.0)
    };
    let value = parse_f64(number, Styles::FLOAT).ok_or_else(|| error(ErrorKind::Syntax))?;
    if !value.is_finite() {
        return Err(error(ErrorKind::NotFinite));
    }
    if value < 0.0 {
        return Err(error(ErrorKind::Negative));
    }
    let twips = (value * factor).round_ties_even();
    if twips > f64::from(u32::MAX) {
        return Err(error(ErrorKind::OutOfRange));
    }
    round_emu(twips * 635.0).ok_or_else(|| error(ErrorKind::Overflow))
}
