//! Canonical semantic model for ODF font-face declarations.

use super::{
    MAX_AGGREGATE_TEXT_BYTES, MAX_FONT_FACES, MAX_FORMATS_PER_SOURCE, MAX_SOURCES_PER_FACE,
    MAX_VALUE_BYTES, invalid,
};
use litchi_core::{Error, Result};
use std::collections::HashSet;

/// Generic CSS font family used by `style:font-family-generic`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Family {
    Roman,
    Swiss,
    Modern,
    Decorative,
    Script,
    System,
}

impl Family {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "roman" => Ok(Self::Roman),
            "swiss" => Ok(Self::Swiss),
            "modern" => Ok(Self::Modern),
            "decorative" => Ok(Self::Decorative),
            "script" => Ok(Self::Script),
            "system" => Ok(Self::System),
            _ => invalid(format!("unsupported style:font-family-generic '{value}'")),
        }
    }

    pub(super) fn as_str(self) -> &'static str {
        match self {
            Self::Roman => "roman",
            Self::Swiss => "swiss",
            Self::Modern => "modern",
            Self::Decorative => "decorative",
            Self::Script => "script",
            Self::System => "system",
        }
    }
}

/// Font pitch stored by `style:font-pitch`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Pitch {
    Fixed,
    Variable,
}

impl Pitch {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "fixed" => Ok(Self::Fixed),
            "variable" => Ok(Self::Variable),
            _ => invalid(format!("unsupported style:font-pitch '{value}'")),
        }
    }

    pub(super) fn as_str(self) -> &'static str {
        match self {
            Self::Fixed => "fixed",
            Self::Variable => "variable",
        }
    }
}

/// SVG font style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Style {
    Normal,
    Italic,
    Oblique,
}

impl Style {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "normal" => Ok(Self::Normal),
            "italic" => Ok(Self::Italic),
            "oblique" => Ok(Self::Oblique),
            _ => invalid(format!("unsupported svg:font-style '{value}'")),
        }
    }

    pub(super) fn as_str(self) -> &'static str {
        match self {
            Self::Normal => "normal",
            Self::Italic => "italic",
            Self::Oblique => "oblique",
        }
    }
}

/// SVG font variant.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Variant {
    Normal,
    SmallCaps,
}

impl Variant {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "normal" => Ok(Self::Normal),
            "small-caps" => Ok(Self::SmallCaps),
            _ => invalid(format!("unsupported svg:font-variant '{value}'")),
        }
    }

    pub(super) fn as_str(self) -> &'static str {
        match self {
            Self::Normal => "normal",
            Self::SmallCaps => "small-caps",
        }
    }
}

/// SVG font weight, including every standard numeric weight.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Weight {
    Normal,
    Bold,
    Weight100,
    Weight200,
    Weight300,
    Weight400,
    Weight500,
    Weight600,
    Weight700,
    Weight800,
    Weight900,
}

impl Weight {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "normal" => Ok(Self::Normal),
            "bold" => Ok(Self::Bold),
            "100" => Ok(Self::Weight100),
            "200" => Ok(Self::Weight200),
            "300" => Ok(Self::Weight300),
            "400" => Ok(Self::Weight400),
            "500" => Ok(Self::Weight500),
            "600" => Ok(Self::Weight600),
            "700" => Ok(Self::Weight700),
            "800" => Ok(Self::Weight800),
            "900" => Ok(Self::Weight900),
            _ => invalid(format!("unsupported svg:font-weight '{value}'")),
        }
    }

    pub(super) fn as_str(self) -> &'static str {
        match self {
            Self::Normal => "normal",
            Self::Bold => "bold",
            Self::Weight100 => "100",
            Self::Weight200 => "200",
            Self::Weight300 => "300",
            Self::Weight400 => "400",
            Self::Weight500 => "500",
            Self::Weight600 => "600",
            Self::Weight700 => "700",
            Self::Weight800 => "800",
            Self::Weight900 => "900",
        }
    }
}

/// SVG font stretch classification.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Stretch {
    Normal,
    UltraCondensed,
    ExtraCondensed,
    Condensed,
    SemiCondensed,
    SemiExpanded,
    Expanded,
    ExtraExpanded,
    UltraExpanded,
}

impl Stretch {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "normal" => Ok(Self::Normal),
            "ultra-condensed" => Ok(Self::UltraCondensed),
            "extra-condensed" => Ok(Self::ExtraCondensed),
            "condensed" => Ok(Self::Condensed),
            "semi-condensed" => Ok(Self::SemiCondensed),
            "semi-expanded" => Ok(Self::SemiExpanded),
            "expanded" => Ok(Self::Expanded),
            "extra-expanded" => Ok(Self::ExtraExpanded),
            "ultra-expanded" => Ok(Self::UltraExpanded),
            _ => invalid(format!("unsupported svg:font-stretch '{value}'")),
        }
    }

    pub(super) fn as_str(self) -> &'static str {
        match self {
            Self::Normal => "normal",
            Self::UltraCondensed => "ultra-condensed",
            Self::ExtraCondensed => "extra-condensed",
            Self::Condensed => "condensed",
            Self::SemiCondensed => "semi-condensed",
            Self::SemiExpanded => "semi-expanded",
            Self::Expanded => "expanded",
            Self::ExtraExpanded => "extra-expanded",
            Self::UltraExpanded => "ultra-expanded",
        }
    }
}

/// A validated positive ODF length used by `svg:font-size`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Length(String);

impl Length {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_positive_length(&value)?;
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// One numeric SVG font metric.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum MetricKind {
    UnitsPerEm,
    StemV,
    StemH,
    Slope,
    CapHeight,
    XHeight,
    AccentHeight,
    Ascent,
    Descent,
    Ideographic,
    Alphabetic,
    Mathematical,
    Hanging,
    VerticalIdeographic,
    VerticalAlphabetic,
    VerticalMathematical,
    VerticalHanging,
    UnderlinePosition,
    UnderlineThickness,
    StrikethroughPosition,
    StrikethroughThickness,
    OverlinePosition,
    OverlineThickness,
}

impl MetricKind {
    pub(super) fn from_local(local: &[u8]) -> Option<Self> {
        match local {
            b"units-per-em" => Some(Self::UnitsPerEm),
            b"stemv" => Some(Self::StemV),
            b"stemh" => Some(Self::StemH),
            b"slope" => Some(Self::Slope),
            b"cap-height" => Some(Self::CapHeight),
            b"x-height" => Some(Self::XHeight),
            b"accent-height" => Some(Self::AccentHeight),
            b"ascent" => Some(Self::Ascent),
            b"descent" => Some(Self::Descent),
            b"ideographic" => Some(Self::Ideographic),
            b"alphabetic" => Some(Self::Alphabetic),
            b"mathematical" => Some(Self::Mathematical),
            b"hanging" => Some(Self::Hanging),
            b"v-ideographic" => Some(Self::VerticalIdeographic),
            b"v-alphabetic" => Some(Self::VerticalAlphabetic),
            b"v-mathematical" => Some(Self::VerticalMathematical),
            b"v-hanging" => Some(Self::VerticalHanging),
            b"underline-position" => Some(Self::UnderlinePosition),
            b"underline-thickness" => Some(Self::UnderlineThickness),
            b"strikethrough-position" => Some(Self::StrikethroughPosition),
            b"strikethrough-thickness" => Some(Self::StrikethroughThickness),
            b"overline-position" => Some(Self::OverlinePosition),
            b"overline-thickness" => Some(Self::OverlineThickness),
            _ => None,
        }
    }

    pub(super) fn as_str(self) -> &'static str {
        match self {
            Self::UnitsPerEm => "units-per-em",
            Self::StemV => "stemv",
            Self::StemH => "stemh",
            Self::Slope => "slope",
            Self::CapHeight => "cap-height",
            Self::XHeight => "x-height",
            Self::AccentHeight => "accent-height",
            Self::Ascent => "ascent",
            Self::Descent => "descent",
            Self::Ideographic => "ideographic",
            Self::Alphabetic => "alphabetic",
            Self::Mathematical => "mathematical",
            Self::Hanging => "hanging",
            Self::VerticalIdeographic => "v-ideographic",
            Self::VerticalAlphabetic => "v-alphabetic",
            Self::VerticalMathematical => "v-mathematical",
            Self::VerticalHanging => "v-hanging",
            Self::UnderlinePosition => "underline-position",
            Self::UnderlineThickness => "underline-thickness",
            Self::StrikethroughPosition => "strikethrough-position",
            Self::StrikethroughThickness => "strikethrough-thickness",
            Self::OverlinePosition => "overline-position",
            Self::OverlineThickness => "overline-thickness",
        }
    }

    pub(super) fn order(self) -> u8 {
        match self {
            Self::UnitsPerEm => 0,
            Self::StemV => 1,
            Self::StemH => 2,
            Self::Slope => 3,
            Self::CapHeight => 4,
            Self::XHeight => 5,
            Self::AccentHeight => 6,
            Self::Ascent => 7,
            Self::Descent => 8,
            Self::Ideographic => 9,
            Self::Alphabetic => 10,
            Self::Mathematical => 11,
            Self::Hanging => 12,
            Self::VerticalIdeographic => 13,
            Self::VerticalAlphabetic => 14,
            Self::VerticalMathematical => 15,
            Self::VerticalHanging => 16,
            Self::UnderlinePosition => 17,
            Self::UnderlineThickness => 18,
            Self::StrikethroughPosition => 19,
            Self::StrikethroughThickness => 20,
            Self::OverlinePosition => 21,
            Self::OverlineThickness => 22,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Metric {
    pub kind: MetricKind,
    pub value: i64,
}

/// An inert simple XLink used by font resources.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Link {
    pub href: String,
    pub actuate_on_request: bool,
}

/// One item inside `svg:font-face-src`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Source {
    Uri {
        link: Link,
        /// Ordered optional `svg:string` format hints.
        formats: Vec<Option<String>>,
    },
    LocalName(Option<String>),
}

/// One complete standard `style:font-face` declaration.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Face {
    pub name: String,
    pub font_adornments: Option<String>,
    pub generic_family: Option<Family>,
    pub pitch: Option<Pitch>,
    pub charset: Option<String>,
    pub family: Option<String>,
    pub style: Option<Style>,
    pub variant: Option<Variant>,
    pub weight: Option<Weight>,
    pub stretch: Option<Stretch>,
    pub size: Option<Length>,
    pub unicode_range: Option<String>,
    pub panose_1: Option<String>,
    pub widths: Option<String>,
    pub bounding_box: Option<String>,
    pub metrics: Vec<Metric>,
    pub sources: Vec<Source>,
    pub definition_source: Option<Link>,
}

/// Semantic contents of one optional `office:font-face-decls` element.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Declarations {
    pub faces: Vec<Face>,
}

impl Declarations {
    pub fn face(&self, name: &str) -> Option<&Face> {
        self.faces.iter().find(|face| face.name == name)
    }

    pub fn validate(&self) -> Result<()> {
        if self.faces.len() > MAX_FONT_FACES {
            return invalid(format!(
                "font-face declarations exceed the {MAX_FONT_FACES} face limit"
            ));
        }
        let mut names = HashSet::with_capacity(self.faces.len());
        let mut text_bytes = 0usize;
        for face in &self.faces {
            validate_value(&face.name, "style:name", false)?;
            if !names.insert(face.name.as_str()) {
                return invalid(format!("duplicate style:font-face name '{}'", face.name));
            }
            text_bytes = add_text_bytes(text_bytes, face.name.len())?;
            for (value, name) in [
                (face.font_adornments.as_deref(), "style:font-adornments"),
                (face.charset.as_deref(), "style:font-charset"),
                (face.family.as_deref(), "svg:font-family"),
                (face.unicode_range.as_deref(), "svg:unicode-range"),
                (face.panose_1.as_deref(), "svg:panose-1"),
                (face.widths.as_deref(), "svg:widths"),
                (face.bounding_box.as_deref(), "svg:bbox"),
            ] {
                if let Some(value) = value {
                    validate_value(value, name, true)?;
                    text_bytes = add_text_bytes(text_bytes, value.len())?;
                }
            }
            if let Some(charset) = &face.charset {
                validate_text_encoding(charset)?;
            }
            if let Some(size) = &face.size {
                validate_positive_length(size.as_str())?;
                text_bytes = add_text_bytes(text_bytes, size.as_str().len())?;
            }
            let mut metric_kinds = HashSet::with_capacity(face.metrics.len());
            for metric in &face.metrics {
                if !metric_kinds.insert(metric.kind) {
                    return invalid(format!(
                        "font face '{}' contains duplicate svg:{}",
                        face.name,
                        metric.kind.as_str()
                    ));
                }
            }
            if face.sources.len() > MAX_SOURCES_PER_FACE {
                return invalid(format!(
                    "font face '{}' exceeds the {MAX_SOURCES_PER_FACE} source limit",
                    face.name
                ));
            }
            for source in &face.sources {
                match source {
                    Source::Uri { link, formats } => {
                        validate_link(link)?;
                        text_bytes = add_text_bytes(text_bytes, link.href.len())?;
                        if formats.len() > MAX_FORMATS_PER_SOURCE {
                            return invalid(format!(
                                "font source exceeds the {MAX_FORMATS_PER_SOURCE} format limit"
                            ));
                        }
                        for format in formats.iter().flatten() {
                            validate_value(format, "svg:string", true)?;
                            text_bytes = add_text_bytes(text_bytes, format.len())?;
                        }
                    },
                    Source::LocalName(name) => {
                        if let Some(name) = name {
                            validate_value(name, "svg:name", true)?;
                            text_bytes = add_text_bytes(text_bytes, name.len())?;
                        }
                    },
                }
            }
            if let Some(link) = &face.definition_source {
                validate_link(link)?;
                text_bytes = add_text_bytes(text_bytes, link.href.len())?;
            }
        }
        Ok(())
    }

    pub fn to_xml(&self) -> Result<String> {
        super::codec::write_declarations(self)
    }
}

fn validate_positive_length(value: &str) -> Result<()> {
    let Some(number) = ["cm", "mm", "in", "pt", "pc", "px"]
        .iter()
        .find_map(|unit| value.strip_suffix(unit))
    else {
        return invalid(format!("invalid positive ODF length '{value}'"));
    };
    if number.is_empty() || number.len() > MAX_VALUE_BYTES {
        return invalid(format!("invalid positive ODF length '{value}'"));
    }
    let mut dots = 0usize;
    let mut digits = 0usize;
    let mut nonzero = false;
    for byte in number.bytes() {
        match byte {
            b'.' => dots += 1,
            b'0'..=b'9' => {
                digits += 1;
                nonzero |= byte != b'0';
            },
            _ => return invalid(format!("invalid positive ODF length '{value}'")),
        }
    }
    if dots > 1 || digits == 0 || !nonzero || number == "." {
        return invalid(format!("invalid positive ODF length '{value}'"));
    }
    Ok(())
}

pub(super) fn validate_text_encoding(value: &str) -> Result<()> {
    validate_value(value, "style:font-charset", false)?;
    let mut bytes = value.bytes();
    if !bytes.next().is_some_and(|byte| byte.is_ascii_alphabetic())
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'_' | b'-'))
    {
        return invalid(format!("invalid style:font-charset '{value}'"));
    }
    Ok(())
}

pub(super) fn validate_link(link: &Link) -> Result<()> {
    validate_value(&link.href, "xlink:href", false)
}

pub(super) fn validate_value(value: &str, name: &str, allow_empty: bool) -> Result<()> {
    if !allow_empty && value.is_empty() {
        return invalid(format!("{name} must not be empty"));
    }
    if value.len() > MAX_VALUE_BYTES {
        return invalid(format!("{name} exceeds the {MAX_VALUE_BYTES} byte limit"));
    }
    Ok(())
}

pub(super) fn add_text_bytes(current: usize, additional: usize) -> Result<usize> {
    let value = current
        .checked_add(additional)
        .ok_or_else(|| Error::InvalidFormat("font metadata size overflow".to_string()))?;
    if value > MAX_AGGREGATE_TEXT_BYTES {
        invalid(format!(
            "font metadata exceeds the {MAX_AGGREGATE_TEXT_BYTES} aggregate byte limit"
        ))
    } else {
        Ok(value)
    }
}
