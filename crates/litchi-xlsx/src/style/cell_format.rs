//! Readable local formatting of shared cell formats.
//!
//! A worksheet cell references one shared cell format, a `cellXfs` record of
//! the workbook's `styles.xml`. [`Format`] is what that record states: its
//! number format, the font, fill and border records it points at, its
//! alignment and its protection, each as typed values (booleans, closed
//! enumerations, lengths and colours) rather than strings.
//!
//! The values are **local**. They are exactly what the record and the
//! records it references say. Nothing is inherited from the parent cell style
//! (`cellStyleXfs`), the `apply*` flags are not interpreted, and theme and
//! indexed colours stay references: a [`Color::Theme`] keeps its index and
//! tint and is never collapsed to RGB. An optional property that the record
//! does not author is `None`, and an explicitly authored `false` is
//! `Some(false)`.
//!
//! # Cost
//!
//! The first [`Format`] requested from a workbook parses `styles.xml` once
//! into an immutable format table. The table is retained by a bounded,
//! single-entry, weighted clean-value cache owned by the workbook (ADR 0005):
//! it is keyed by the identity of the styles payload, so a changed payload
//! never reuses an old table; it is evicted when another payload generation
//! is parsed, and a table whose weight exceeds the cache's ceiling is
//! returned without being retained. Every later [`Format`] of the same
//! payload is a table lookup. A [`Format`] pins the table it was read from.
//!
//! # `OfficeCLI` property names
//!
//! The typed accessors cover every cell-formatting property `OfficeCLI`
//! reads back for an Excel cell (`schemas/help/xlsx/cell.json`), so a node
//! view can serve those keys without parsing again. Native `numFmtId` is
//! available only through the explicit raw metadata function; ordinary
//! unified views must omit that identifier.
//!
//!
//! | `OfficeCLI` key | Typed source |
//! |---|---|
//! | `numFmtId`, `numberformat` | [`Format::number_format`] ([`raw::number_format_id`], [`NumberFormat::code`]) |
//! | `font.bold`, `font.italic`, `strike` | [`Font::bold`], [`Font::italic`], [`Font::strike`] |
//! | `font.name`, `font.size`, `font.color` | [`Font::name`], [`Font::size`] (points), [`Font::color`] |
//! | `underline`, `superscript`, `subscript` | [`Font::underline`], [`Font::script`] |
//! | `font.charset`, `font.family`, `font.scheme` | [`Font::charset`], [`Font::family`], [`Font::scheme`] |
//! | `font.outline`, `font.shadow`, `font.condense`, `font.extend` | [`Font::outline`], [`Font::shadow`], [`Font::condense`], [`Font::extend`] |
//! | `fill`, `fillPattern`, `fillBg` | [`Fill::Pattern`] ([`PatternFill`]) or [`Fill::Gradient`] ([`GradientFill`]) |
//! | `border.{left,right,top,bottom,diagonal}` and their `.color` | [`Border`] sides ([`Side::line`], [`Side::color`]) |
//! | `border.diagonalUp`, `border.diagonalDown` | [`Diagonal::dir`] |
//! | `alignment.*` | [`Alignment`] |
//! | `protection.locked`, `protection.hidden` | [`Protection::locked`], [`Protection::hidden`] |
//! | `quotePrefix` | [`Format::quote_prefix`] |
//!
//! `OfficeCLI` reports nothing for a cell whose format is the base format,
//! and no font, fill or border keys when the format uses the workbook's
//! default record of that kind. [`Format::is_base`],
//! [`Format::has_default_font`], [`Format::has_default_fill`] and
//! [`Format::has_default_border`] answer those questions without exposing
//! physical record indexes.

use std::fmt;
use std::hash::{Hash, Hasher};
use std::sync::Arc;

use litchi_opc::SourceLimits;

pub use super::stylesheet::alignment::{
    Alignment, Horizontal, Indent, Reading, Rotation, Vertical,
};
pub use super::stylesheet::border::{Border, Color, Diagonal, Dir, Line, Side, Tint};
pub use super::stylesheet::font::{Scheme, Script, Underline};
pub use crate::color::Rgb;

use super::stylesheet::{builtin_format_code, is_date_format};
use crate::error::{Result, invalid};

mod cache;
mod parser;

/// Maximum number of records accepted by the native styles collections and by
/// checked gradient authoring. The staged edit byte budget remains the
/// authoritative limit for an actual rewrite.
pub(crate) const MAX_RECORDS: usize = 1_000_000;

#[cfg(test)]
mod tests;

pub(crate) use cache::{FormatCache, FormatKey};

/// Explicit access to native SpreadsheetML formatting identifiers.
pub mod raw {
    /// The stored number-format identifier, including custom workbook-local IDs.
    #[must_use]
    pub const fn number_format_id(format: &super::NumberFormat<'_>) -> u32 {
        format.id()
    }
}

/// The number format a shared cell format applies.
///
/// The identifier is the record's `numFmtId`, `0` (General) when the record
/// does not author one. The code is the workbook's own `numFmt` declaration
/// for that identifier when it has one, otherwise the ECMA-376 §18.8.30 code
/// of a built-in format. Built-in identifiers whose code depends on the
/// locale (for example 5–8 and 27–36) have no code.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct NumberFormat<'a> {
    id: u32,
    code: Option<&'a str>,
    declared: bool,
}

impl<'a> NumberFormat<'a> {
    /// The number-format identifier (`numFmtId`).
    #[must_use]
    const fn id(&self) -> u32 {
        self.id
    }

    /// The format code, when the workbook declares it or it is a
    /// locale-independent built-in format.
    #[must_use]
    pub const fn code(&self) -> Option<&'a str> {
        self.code
    }

    /// Whether the code comes from the workbook's own `numFmt` declaration
    /// rather than the built-in table.
    #[must_use]
    pub const fn is_declared(&self) -> bool {
        self.declared
    }

    /// Whether this is the General format (identifier 0).
    #[must_use]
    pub const fn is_general(&self) -> bool {
        self.id == 0
    }

    /// Whether the code formats a date or time, by the same lexical rule as
    /// [`crate::style::stylesheet::is_date_format`]. A format without a code
    /// is not reported as a date.
    #[must_use]
    pub fn is_date_time(&self) -> bool {
        self.code.is_some_and(is_date_format)
    }
}

/// A font record referenced by a shared cell format.
///
/// Every property is exactly what the record authors: `None` means the
/// record does not state it. A boolean property authored as `<b/>` or
/// `<b val="1"/>` is `Some(true)`; `<b val="0"/>` is `Some(false)`.
#[derive(Clone, PartialEq, Default)]
pub struct Font {
    name: Option<Box<str>>,
    size: Option<f64>,
    bold: Option<bool>,
    italic: Option<bool>,
    strike: Option<bool>,
    outline: Option<bool>,
    shadow: Option<bool>,
    condense: Option<bool>,
    extend: Option<bool>,
    underline: Option<Underline>,
    script: Option<Script>,
    color: Option<Color>,
    family: Option<u32>,
    charset: Option<u32>,
    scheme: Option<Scheme>,
}

impl Font {
    /// Typeface name (`name/@val`).
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Font size in points (`sz/@val`), finite and positive.
    #[must_use]
    pub const fn size(&self) -> Option<f64> {
        self.size
    }

    /// Bold (`b`).
    #[must_use]
    pub const fn bold(&self) -> Option<bool> {
        self.bold
    }

    /// Italic (`i`).
    #[must_use]
    pub const fn italic(&self) -> Option<bool> {
        self.italic
    }

    /// Single strikethrough (`strike`).
    #[must_use]
    pub const fn strike(&self) -> Option<bool> {
        self.strike
    }

    /// Outline glyphs (`outline`).
    #[must_use]
    pub const fn outline(&self) -> Option<bool> {
        self.outline
    }

    /// Shadowed glyphs (`shadow`).
    #[must_use]
    pub const fn shadow(&self) -> Option<bool> {
        self.shadow
    }

    /// Condensed spacing (`condense`).
    #[must_use]
    pub const fn condense(&self) -> Option<bool> {
        self.condense
    }

    /// Extended spacing (`extend`).
    #[must_use]
    pub const fn extend(&self) -> Option<bool> {
        self.extend
    }

    /// Underline style (`u`); an explicitly authored `none` is
    /// `Some(Underline::None)`.
    #[must_use]
    pub const fn underline(&self) -> Option<Underline> {
        self.underline
    }

    /// Baseline placement (`vertAlign`): superscript, subscript or baseline.
    #[must_use]
    pub const fn script(&self) -> Option<Script> {
        self.script
    }

    /// Font colour (`color`), with its theme, palette or automatic reference
    /// and tint kept as authored.
    #[must_use]
    pub const fn color(&self) -> Option<Color> {
        self.color
    }

    /// Font family class (`family/@val`).
    #[must_use]
    pub const fn family(&self) -> Option<u32> {
        self.family
    }

    /// Character set (`charset/@val`).
    #[must_use]
    pub const fn charset(&self) -> Option<u32> {
        self.charset
    }

    /// Theme font scheme (`scheme/@val`).
    #[must_use]
    pub const fn scheme(&self) -> Option<Scheme> {
        self.scheme
    }
}

/// A fill record referenced by a shared cell format.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Fill {
    /// A pattern fill (`patternFill`). A `fill` element without a child is a
    /// pattern fill that authors nothing.
    Pattern(PatternFill),
    /// A gradient fill (`gradientFill`).
    Gradient(GradientFill),
}

impl Fill {
    /// The pattern fill, when this is one.
    #[must_use]
    pub const fn pattern(&self) -> Option<&PatternFill> {
        match self {
            Self::Pattern(fill) => Some(fill),
            Self::Gradient(_) => None,
        }
    }

    /// The gradient fill, when this is one.
    #[must_use]
    pub const fn gradient(&self) -> Option<&GradientFill> {
        match self {
            Self::Gradient(fill) => Some(fill),
            Self::Pattern(_) => None,
        }
    }

    /// Validate this fill against the checked authoring profile.
    ///
    /// Source parsing intentionally accepts a broader finite geometry domain
    /// so an existing workbook can be inspected. Callers that pass a
    /// read-created fill to an authoring operation can use this method before
    /// writing it.
    pub fn validate(&self) -> Result<()> {
        match self {
            Self::Pattern(_) => Ok(()),
            Self::Gradient(fill) => fill.validate(),
        }
    }
}

/// A pattern fill: pattern, foreground and background colours.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct PatternFill {
    pattern: Option<Pattern>,
    foreground: Option<Color>,
    background: Option<Color>,
}

impl PatternFill {
    /// Construct a pattern fill while preserving each optional field exactly.
    #[must_use]
    pub const fn new(
        pattern: Option<Pattern>,
        foreground: Option<Color>,
        background: Option<Color>,
    ) -> Self {
        Self {
            pattern,
            foreground,
            background,
        }
    }

    /// The authored `patternType`; the schema gives it no default.
    #[must_use]
    pub const fn pattern(&self) -> Option<Pattern> {
        self.pattern
    }

    /// Foreground colour (`fgColor`), the colour of a solid fill.
    #[must_use]
    pub const fn foreground(&self) -> Option<Color> {
        self.foreground
    }

    /// Background colour (`bgColor`).
    #[must_use]
    pub const fn background(&self) -> Option<Color> {
        self.background
    }
}

/// A fill pattern (`ST_PatternType`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Pattern {
    None,
    Solid,
    MediumGray,
    DarkGray,
    LightGray,
    DarkHorizontal,
    DarkVertical,
    DarkDown,
    DarkUp,
    DarkGrid,
    DarkTrellis,
    LightHorizontal,
    LightVertical,
    LightDown,
    LightUp,
    LightGrid,
    LightTrellis,
    Gray125,
    Gray0625,
}

impl Pattern {
    /// Every pattern, in `ST_PatternType` order.
    pub const ALL: [Self; 19] = [
        Self::None,
        Self::Solid,
        Self::MediumGray,
        Self::DarkGray,
        Self::LightGray,
        Self::DarkHorizontal,
        Self::DarkVertical,
        Self::DarkDown,
        Self::DarkUp,
        Self::DarkGrid,
        Self::DarkTrellis,
        Self::LightHorizontal,
        Self::LightVertical,
        Self::LightDown,
        Self::LightUp,
        Self::LightGrid,
        Self::LightTrellis,
        Self::Gray125,
        Self::Gray0625,
    ];

    /// The exact `SpreadsheetML` token.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Solid => "solid",
            Self::MediumGray => "mediumGray",
            Self::DarkGray => "darkGray",
            Self::LightGray => "lightGray",
            Self::DarkHorizontal => "darkHorizontal",
            Self::DarkVertical => "darkVertical",
            Self::DarkDown => "darkDown",
            Self::DarkUp => "darkUp",
            Self::DarkGrid => "darkGrid",
            Self::DarkTrellis => "darkTrellis",
            Self::LightHorizontal => "lightHorizontal",
            Self::LightVertical => "lightVertical",
            Self::LightDown => "lightDown",
            Self::LightUp => "lightUp",
            Self::LightGrid => "lightGrid",
            Self::LightTrellis => "lightTrellis",
            Self::Gray125 => "gray125",
            Self::Gray0625 => "gray0625",
        }
    }

    /// Parse an exact `SpreadsheetML` token, without allocating.
    #[must_use]
    pub fn from_token(token: &str) -> Option<Self> {
        Self::ALL
            .iter()
            .copied()
            .find(|pattern| pattern.as_str() == token)
    }
}

impl fmt::Display for Pattern {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// A gradient fill (`gradientFill`).
///
/// Geometry attributes the record does not author take their schema
/// defaults: a linear gradient at 0 degrees with all four path edges at 0.
#[derive(Debug, Clone, PartialEq)]
pub struct GradientFill {
    kind: GradientKind,
    degree: f64,
    left: f64,
    right: f64,
    top: f64,
    bottom: f64,
    stops: Box<[GradientStop]>,
}

impl GradientFill {
    /// Check the native gradient stop-count ceiling before allocating a stop array.
    ///
    /// Edit-specific retained-byte budgets still apply when the value is staged.
    pub fn validate_stop_count(count: usize) -> Result<()> {
        validate_stop_count(count)
    }

    /// Construct a checked linear gradient.
    ///
    /// The degree is inclusive in `0..=360`. Stop positions are inclusive in
    /// `0..=1`; their order is retained exactly. Zero-valued authored numbers
    /// are canonicalized to positive zero so equivalent values hash alike.
    pub fn linear(degree: f64, mut stops: Box<[GradientStop]>) -> Result<Self> {
        let degree = checked_degree(degree)?;
        checked_stops(&mut stops)?;
        Ok(Self {
            kind: GradientKind::Linear,
            degree,
            left: 0.0,
            right: 0.0,
            top: 0.0,
            bottom: 0.0,
            stops,
        })
    }

    /// Construct a checked path gradient.
    ///
    /// Path edges and stop positions are inclusive in `0..=1`. A path
    /// gradient uses the schema's default degree of zero; stop order is
    /// retained exactly. Zero-valued authored numbers are canonicalized to
    /// positive zero so equivalent values hash alike.
    pub fn path(
        left: f64,
        right: f64,
        top: f64,
        bottom: f64,
        mut stops: Box<[GradientStop]>,
    ) -> Result<Self> {
        let left = checked_unit_interval(left, "gradient left edge")?;
        let right = checked_unit_interval(right, "gradient right edge")?;
        let top = checked_unit_interval(top, "gradient top edge")?;
        let bottom = checked_unit_interval(bottom, "gradient bottom edge")?;
        checked_stops(&mut stops)?;
        Ok(Self {
            kind: GradientKind::Path,
            degree: 0.0,
            left,
            right,
            top,
            bottom,
            stops,
        })
    }

    /// Validate this gradient against the checked authoring profile without
    /// changing its values.
    ///
    /// This is useful for a gradient obtained from the permissive native
    /// reader before it is supplied to a whole-fill authoring operation.
    pub fn validate(&self) -> Result<()> {
        checked_degree(self.degree)?;
        checked_unit_interval(self.left, "gradient left edge")?;
        checked_unit_interval(self.right, "gradient right edge")?;
        checked_unit_interval(self.top, "gradient top edge")?;
        checked_unit_interval(self.bottom, "gradient bottom edge")?;
        validate_stops(&self.stops)
    }

    /// Linear or path gradient (`type`).
    #[must_use]
    pub const fn kind(&self) -> GradientKind {
        self.kind
    }

    /// Angle of a linear gradient in degrees (`degree`).
    #[must_use]
    pub const fn degree(&self) -> f64 {
        self.degree
    }

    /// Left edge of a path gradient's inner rectangle (`left`), 0 to 1.
    #[must_use]
    pub const fn left(&self) -> f64 {
        self.left
    }

    /// Right edge of a path gradient's inner rectangle (`right`), 0 to 1.
    #[must_use]
    pub const fn right(&self) -> f64 {
        self.right
    }

    /// Top edge of a path gradient's inner rectangle (`top`), 0 to 1.
    #[must_use]
    pub const fn top(&self) -> f64 {
        self.top
    }

    /// Bottom edge of a path gradient's inner rectangle (`bottom`), 0 to 1.
    #[must_use]
    pub const fn bottom(&self) -> f64 {
        self.bottom
    }

    /// Colour stops in document order.
    #[must_use]
    pub fn stops(&self) -> &[GradientStop] {
        &self.stops
    }
}

/// The kind of a gradient fill (`ST_GradientType`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum GradientKind {
    /// A linear gradient, the schema default.
    #[default]
    Linear,
    /// A path gradient.
    Path,
}

impl GradientKind {
    /// The exact `SpreadsheetML` token.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Linear => "linear",
            Self::Path => "path",
        }
    }

    /// Parse an exact `SpreadsheetML` token.
    #[must_use]
    pub fn from_token(token: &str) -> Option<Self> {
        match token {
            "linear" => Some(Self::Linear),
            "path" => Some(Self::Path),
            _ => None,
        }
    }
}

/// One colour stop of a gradient fill.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct GradientStop {
    position: f64,
    color: Color,
}

impl GradientStop {
    /// Construct a checked colour stop with a position in `0..=1`.
    ///
    /// A signed zero position is stored as positive zero so authored values
    /// have one canonical representation.
    pub fn new(position: f64, color: Color) -> Result<Self> {
        let position = checked_unit_interval(position, "gradient-stop position")?;
        Ok(Self { position, color })
    }

    /// Position along the gradient, from 0 to 1.
    #[must_use]
    pub const fn position(&self) -> f64 {
        self.position
    }

    /// Colour at this stop.
    #[must_use]
    pub const fn color(&self) -> Color {
        self.color
    }
}

impl Eq for GradientStop {}

impl Hash for GradientStop {
    fn hash<H: Hasher>(&self, state: &mut H) {
        canonical_float_bits(self.position).hash(state);
        self.color.hash(state);
    }
}

impl Eq for GradientFill {}

impl Hash for GradientFill {
    fn hash<H: Hasher>(&self, state: &mut H) {
        self.kind.hash(state);
        canonical_float_bits(self.degree).hash(state);
        canonical_float_bits(self.left).hash(state);
        canonical_float_bits(self.right).hash(state);
        canonical_float_bits(self.top).hash(state);
        canonical_float_bits(self.bottom).hash(state);
        self.stops.hash(state);
    }
}

fn checked_degree(value: f64) -> Result<f64> {
    if !value.is_finite() || !(0.0..=360.0).contains(&value) {
        return Err(invalid(
            "gradient degree must be finite and between 0 and 360",
        ));
    }
    Ok(canonical_zero(value))
}

fn checked_unit_interval(value: f64, description: &'static str) -> Result<f64> {
    if !value.is_finite() || !(0.0..=1.0).contains(&value) {
        return Err(invalid(format!(
            "{description} must be finite and between 0 and 1"
        )));
    }
    Ok(canonical_zero(value))
}

fn checked_stops(stops: &mut [GradientStop]) -> Result<()> {
    validate_stops(stops)?;
    for stop in stops {
        stop.position = canonical_zero(stop.position);
    }
    Ok(())
}

fn validate_stops(stops: &[GradientStop]) -> Result<()> {
    validate_stop_count(stops.len())?;
    for stop in stops {
        checked_unit_interval(stop.position, "gradient-stop position")?;
    }
    Ok(())
}

fn validate_stop_count(count: usize) -> Result<()> {
    if count > MAX_RECORDS {
        return Err(crate::Error::ResourceLimit(litchi_core::ResourceLimit {
            resource: litchi_core::Resource::Objects,
            observed: u64::try_from(count).unwrap_or(u64::MAX),
            limit: u64::try_from(MAX_RECORDS).unwrap_or(u64::MAX),
            scope: Arc::from("gradient fill stops"),
        }));
    }
    Ok(())
}

fn canonical_zero(value: f64) -> f64 {
    if value == 0.0 { 0.0 } else { value }
}

fn canonical_float_bits(value: f64) -> u64 {
    canonical_zero(value).to_bits()
}

/// Cell protection authored by a shared cell format (`protection`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct Protection {
    locked: Option<bool>,
    hidden: Option<bool>,
}

impl Protection {
    /// Whether the cell is locked when its sheet is protected (`locked`).
    #[must_use]
    pub const fn locked(&self) -> Option<bool> {
        self.locked
    }

    /// Whether the cell's formula is hidden when its sheet is protected
    /// (`hidden`).
    #[must_use]
    pub const fn hidden(&self) -> Option<bool> {
        self.hidden
    }
}

/// One `cellXfs` record as the format table keeps it.
#[derive(Debug, Clone, PartialEq, Default)]
struct CellXf {
    number_format: Option<u32>,
    font: Option<u32>,
    fill: Option<u32>,
    border: Option<u32>,
    alignment: Option<Alignment>,
    protection: Option<Protection>,
    quote_prefix: Option<bool>,
}

/// The parsed, immutable format table of one styles payload.
#[derive(Debug, PartialEq)]
pub(crate) struct FormatTable {
    /// Declared number formats, sorted by identifier.
    number_formats: Box<[(u32, Box<str>)]>,
    fonts: Box<[Font]>,
    fills: Box<[Fill]>,
    borders: Box<[Border]>,
    cell_formats: Box<[CellXf]>,
    weight: usize,
}

impl FormatTable {
    /// Parse the format-bearing collections of a `styles.xml` payload.
    #[cfg(test)]
    pub(crate) fn parse(content: &[u8]) -> Result<Self> {
        Self::parse_with_limits(content, SourceLimits::default())
    }

    /// Parse a `styles.xml` payload under the caller's physical XML policy.
    pub(crate) fn parse_with_limits(content: &[u8], limits: SourceLimits) -> Result<Self> {
        crate::source_limits::validate_xml_limits(content, limits)?;
        parser::parse(content)
    }

    /// Assemble a table from parsed collections; `number_formats` is sorted
    /// by identifier, without duplicates.
    fn from_parts(
        number_formats: Vec<(u32, Box<str>)>,
        fonts: Vec<Font>,
        fills: Vec<Fill>,
        borders: Vec<Border>,
        cell_formats: Vec<CellXf>,
    ) -> Self {
        debug_assert!(
            number_formats.windows(2).all(|pair| pair[0].0 < pair[1].0),
            "number formats are sorted and unique"
        );
        let mut table = Self {
            number_formats: number_formats.into_boxed_slice(),
            fonts: fonts.into_boxed_slice(),
            fills: fills.into_boxed_slice(),
            borders: borders.into_boxed_slice(),
            cell_formats: cell_formats.into_boxed_slice(),
            weight: 0,
        };
        table.weight = table.measure();
        table
    }

    /// Logical retained weight in bytes: the table, its slices and their
    /// string and gradient-stop payloads. It is the cache's admission
    /// measure, not an allocator or RSS measurement.
    fn measure(&self) -> usize {
        use std::mem::size_of;

        let codes = self
            .number_formats
            .iter()
            .map(|(_, code)| code.len())
            .sum::<usize>();
        let names = self
            .fonts
            .iter()
            .filter_map(|font| font.name.as_deref())
            .map(str::len)
            .sum::<usize>();
        let stops = self
            .fills
            .iter()
            .filter_map(Fill::gradient)
            .map(|gradient| gradient.stops.len())
            .sum::<usize>();
        size_of::<Self>()
            .saturating_add(
                self.number_formats
                    .len()
                    .saturating_mul(size_of::<(u32, Box<str>)>()),
            )
            .saturating_add(codes)
            .saturating_add(self.fonts.len().saturating_mul(size_of::<Font>()))
            .saturating_add(names)
            .saturating_add(self.fills.len().saturating_mul(size_of::<Fill>()))
            .saturating_add(stops.saturating_mul(size_of::<GradientStop>()))
            .saturating_add(self.borders.len().saturating_mul(size_of::<Border>()))
            .saturating_add(self.cell_formats.len().saturating_mul(size_of::<CellXf>()))
    }

    /// Logical retained weight in bytes.
    pub(crate) const fn weight(&self) -> usize {
        self.weight
    }

    /// Number of shared cell formats (`cellXfs` records).
    pub(crate) const fn len(&self) -> usize {
        self.cell_formats.len()
    }

    fn number_format(&self, id: u32) -> NumberFormat<'_> {
        match self
            .number_formats
            .binary_search_by_key(&id, |(key, _)| *key)
        {
            Ok(position) => NumberFormat {
                id,
                code: Some(&self.number_formats[position].1),
                declared: true,
            },
            Err(_) => NumberFormat {
                id,
                code: builtin_format_code(id),
                declared: false,
            },
        }
    }
}

/// The readable local formatting of one shared cell format.
///
/// A `Format` is a cheap handle: it shares the workbook's parsed format
/// table, which it keeps alive, and is `Send + Sync`. Read one with
/// [`crate::Style::format`] on an in-memory workbook, or with
/// [`crate::SourceWorksheet::format`] on a source-backed one.
///
/// ```no_run
/// use litchi_xlsx::Package;
///
/// let package = Package::open("report.xlsx")?;
/// let sheet = package.workbook().sheet("Summary")?.expect("sheet");
/// if let Some(style) = sheet.style("B2")? {
///     let format = style.format()?;
///     let bold = format.font().and_then(|font| font.bold()) == Some(true);
///     println!("bold: {bold}, code: {:?}", format.number_format().code());
/// }
/// # Ok::<(), litchi_xlsx::Error>(())
/// ```
#[derive(Clone)]
pub struct Format {
    table: Arc<FormatTable>,
    index: u32,
}

// Format handles are shared across threads by the unified operations layer.
const _: () = {
    const fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Format>();
    assert_send_sync::<FormatTable>();
};

impl Format {
    /// Select one record of a parsed table.
    pub(crate) fn new(table: Arc<FormatTable>, index: u32) -> Result<Self> {
        let count = table.len();
        let inside = usize::try_from(index).is_ok_and(|position| position < count);
        if !inside {
            return Err(invalid(format!(
                "shared style {index} is outside the workbook's {count} cell formats"
            )));
        }
        Ok(Self { table, index })
    }

    fn record(&self) -> &CellXf {
        // `new` proved the index is inside the immutable table.
        &self.table.cell_formats[self.index as usize]
    }

    /// Whether this is the workbook's base format, the one a cell without
    /// a local style uses (`cellXfs` record 0).
    #[must_use]
    pub const fn is_base(&self) -> bool {
        self.index == 0
    }

    /// The number format.
    #[must_use]
    pub fn number_format(&self) -> NumberFormat<'_> {
        self.table
            .number_format(self.record().number_format.unwrap_or(0))
    }

    /// The font record, or `None` when the reference does not resolve (the
    /// workbook has no font records, or the record's `fontId` is past their
    /// end).
    #[must_use]
    pub fn font(&self) -> Option<&Font> {
        component(&self.table.fonts, self.record().font)
    }

    /// The fill record, or `None` when the reference does not resolve.
    #[must_use]
    pub fn fill(&self) -> Option<&Fill> {
        component(&self.table.fills, self.record().fill)
    }

    /// The border record, or `None` when the reference does not resolve.
    #[must_use]
    pub fn border(&self) -> Option<&Border> {
        component(&self.table.borders, self.record().border)
    }

    /// The record's own alignment, when it authors one.
    #[must_use]
    pub fn alignment(&self) -> Option<&Alignment> {
        self.record().alignment.as_ref()
    }

    /// The record's own protection, when it authors one.
    #[must_use]
    pub fn protection(&self) -> Option<Protection> {
        self.record().protection
    }

    /// Whether text entered in the cell is kept as text (`quotePrefix`).
    #[must_use]
    pub fn quote_prefix(&self) -> bool {
        self.record().quote_prefix == Some(true)
    }

    /// Whether the font is the workbook's default font record (the record
    /// authors no `fontId`, or `fontId="0"`).
    #[must_use]
    pub fn has_default_font(&self) -> bool {
        self.record().font.unwrap_or(0) == 0
    }

    /// Whether the fill is the workbook's default fill record (no `fillId`,
    /// or `fillId="0"`).
    #[must_use]
    pub fn has_default_fill(&self) -> bool {
        self.record().fill.unwrap_or(0) == 0
    }

    /// Whether the border is the workbook's default border record (no
    /// `borderId`, or `borderId="0"`).
    #[must_use]
    pub fn has_default_border(&self) -> bool {
        self.record().border.unwrap_or(0) == 0
    }

    /// Whether two handles name the same record of the same parsed table.
    #[must_use]
    pub fn same(&self, other: &Self) -> bool {
        self.index == other.index && Arc::ptr_eq(&self.table, &other.table)
    }
}

fn component<T>(values: &[T], reference: Option<u32>) -> Option<&T> {
    values.get(usize::try_from(reference.unwrap_or(0)).ok()?)
}

impl fmt::Debug for Format {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Format")
            .field("is_base", &self.is_base())
            .finish_non_exhaustive()
    }
}

impl fmt::Debug for Font {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Font")
            .field("has_name", &self.name.is_some())
            .finish_non_exhaustive()
    }
}

impl fmt::Debug for NumberFormat<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("NumberFormat")
            .field("declared", &self.declared)
            .field("has_code", &self.code.is_some())
            .finish_non_exhaustive()
    }
}
