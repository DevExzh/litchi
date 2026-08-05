//! Complete typed ODF `style:table-cell-properties` support.
//!
//! ODF 1.3 section 17.4 covers the full attribute set (alignment, direction, writing
//! mode, shadow, background color and image, borders and diagonals, border line widths,
//! padding, wrapping, rotation, cell protection, printing and shrinking) plus the single
//! optional `style:background-image` child. Unknown attributes, children, or text are
//! rejected.

use super::row::{
    BackgroundColor, BackgroundImage, BackgroundPosition, BackgroundSource,
    HorizontalBackgroundPosition, Opacity, Repeat, VerticalBackgroundPosition,
};
use super::table::{Shadow, WritingMode};
use crate::{FlatDocument, Package};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{QName, ResolveResult},
    reader::NsReader,
};
use std::collections::HashSet;

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FO: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";
const STYLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const FO_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const DRAW_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const XLINK_NS: &str = "http://www.w3.org/1999/xlink";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_STYLES: usize = 65_536;
const MAX_VALUE: usize = 4096;
const MAX_TOTAL: usize = 16 * 1024 * 1024;
const MAX_BINARY: usize = 8 * 1024 * 1024;
const MAX_ATTRIBUTES: usize = 48;
/// Upper bound for `style:decimal-places`; far beyond any spreadsheet implementation.
const MAX_DECIMAL_PLACES: u32 = 100;
/// `style:rotation-angle` is an integer number of degrees below a full turn.
const MAX_ROTATION_DEGREES: u16 = 359;

fn bad(x: impl Into<String>) -> Error {
    Error::InvalidFormat(x.into())
}
fn safe(x: &str, name: &str, empty: bool) -> Result<()> {
    if (!empty && x.is_empty()) || x.len() > MAX_VALUE || x.chars().any(char::is_control) {
        return Err(bad(format!("invalid {name}")));
    }
    Ok(())
}

/// A positive or non-negative ODF physical length used by cell paddings and border widths.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Length(String);
impl Length {
    pub fn positive(x: impl Into<String>) -> Result<Self> {
        Self::new(x.into(), false)
    }
    pub fn non_negative(x: impl Into<String>) -> Result<Self> {
        Self::new(x.into(), true)
    }
    fn new(x: String, zero: bool) -> Result<Self> {
        if x.len() > MAX_VALUE {
            return Err(bad("table-cell length is too large"));
        }
        let Some(number) = ["cm", "mm", "in", "pt", "pc", "px"]
            .iter()
            .find_map(|unit| x.strip_suffix(unit))
        else {
            return Err(bad("table-cell length must use an ODF physical unit"));
        };
        if number.starts_with(['+', '-']) {
            return Err(bad("table-cell length cannot be signed"));
        }
        let mut parts = number.split('.');
        let whole = parts.next().unwrap_or_default();
        let fraction = parts.next();
        if parts.next().is_some()
            || whole.is_empty()
            || !whole.bytes().all(|c| c.is_ascii_digit())
            || fraction
                .is_some_and(|part| part.is_empty() || !part.bytes().all(|c| c.is_ascii_digit()))
        {
            return Err(bad("invalid table-cell length"));
        }
        if !zero && !number.bytes().any(|c| c.is_ascii_digit() && c != b'0') {
            return Err(bad("table-cell length must be positive"));
        }
        Ok(Self(x))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// An `fo:border*` or `style:diagonal-*` border description.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Border(String);
impl Border {
    pub fn new(x: impl Into<String>) -> Result<Self> {
        let x = x.into();
        safe(&x, "table-cell border", false)?;
        Ok(Self(x))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// A three-part border line width: inner line, space between lines, outer line.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BorderWidths {
    pub inner_width: Length,
    pub space: Length,
    pub outer_width: Length,
}
impl BorderWidths {
    fn parse(x: &str) -> Result<Self> {
        let words: Vec<_> = x.split_ascii_whitespace().collect();
        let [inner, space, outer] = words.as_slice() else {
            return Err(bad("border line width needs exactly three lengths"));
        };
        Ok(Self {
            inner_width: Length::positive(*inner)?,
            space: Length::positive(*space)?,
            outer_width: Length::positive(*outer)?,
        })
    }
    fn xml(&self) -> String {
        format!(
            "{} {} {}",
            self.inner_width.as_str(),
            self.space.as_str(),
            self.outer_width.as_str()
        )
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalAlign {
    Top,
    Middle,
    Bottom,
    Automatic,
}
impl VerticalAlign {
    fn parse(x: &str) -> Result<Self> {
        match x {
            "top" => Ok(Self::Top),
            "middle" => Ok(Self::Middle),
            "bottom" => Ok(Self::Bottom),
            "automatic" => Ok(Self::Automatic),
            _ => Err(bad("invalid style:vertical-align")),
        }
    }
    fn xml(self) -> &'static str {
        match self {
            Self::Top => "top",
            Self::Middle => "middle",
            Self::Bottom => "bottom",
            Self::Automatic => "automatic",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextAlignSource {
    Fix,
    ValueType,
}
impl TextAlignSource {
    fn parse(x: &str) -> Result<Self> {
        match x {
            "fix" => Ok(Self::Fix),
            "value-type" => Ok(Self::ValueType),
            _ => Err(bad("invalid style:text-align-source")),
        }
    }
    fn xml(self) -> &'static str {
        match self {
            Self::Fix => "fix",
            Self::ValueType => "value-type",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Direction {
    Ltr,
    Ttb,
}
impl Direction {
    fn parse(x: &str) -> Result<Self> {
        match x {
            "ltr" => Ok(Self::Ltr),
            "ttb" => Ok(Self::Ttb),
            _ => Err(bad("invalid style:direction")),
        }
    }
    fn xml(self) -> &'static str {
        match self {
            Self::Ltr => "ltr",
            Self::Ttb => "ttb",
        }
    }
}

/// `style:glyph-orientation-vertical`: `auto` or a zero angle in any unit.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GlyphOrientation(String);
impl GlyphOrientation {
    pub fn new(x: impl Into<String>) -> Result<Self> {
        let x = x.into();
        match x.as_str() {
            "auto" | "0" | "0deg" | "0rad" | "0grad" => Ok(Self(x)),
            _ => Err(bad("invalid style:glyph-orientation-vertical")),
        }
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Wrap {
    NoWrap,
    Wrap,
}
impl Wrap {
    fn parse(x: &str) -> Result<Self> {
        match x {
            "no-wrap" => Ok(Self::NoWrap),
            "wrap" => Ok(Self::Wrap),
            _ => Err(bad("invalid fo:wrap-option")),
        }
    }
    fn xml(self) -> &'static str {
        match self {
            Self::NoWrap => "no-wrap",
            Self::Wrap => "wrap",
        }
    }
}

/// `style:rotation-angle` as a whole number of degrees below a full turn.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RotationAngle(u16);
impl RotationAngle {
    pub fn new(degrees: u16) -> Result<Self> {
        if degrees > MAX_ROTATION_DEGREES {
            return Err(bad("style:rotation-angle is out of range"));
        }
        Ok(Self(degrees))
    }
    fn parse(x: &str) -> Result<Self> {
        let degrees: u16 = x.parse().map_err(|_| bad("invalid style:rotation-angle"))?;
        Self::new(degrees)
    }
    pub fn degrees(self) -> u16 {
        self.0
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RotationAlign {
    None,
    Bottom,
    Top,
    Center,
}
impl RotationAlign {
    fn parse(x: &str) -> Result<Self> {
        match x {
            "none" => Ok(Self::None),
            "bottom" => Ok(Self::Bottom),
            "top" => Ok(Self::Top),
            "center" => Ok(Self::Center),
            _ => Err(bad("invalid style:rotation-align")),
        }
    }
    fn xml(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Bottom => "bottom",
            Self::Top => "top",
            Self::Center => "center",
        }
    }
}

/// `style:cell-protect`: `none`, `hidden-and-protected`, or a combination of the
/// `protected` and `formula-hidden` flags.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Protection {
    None,
    HiddenAndProtected,
    Protected,
    FormulaHidden,
    ProtectedFormulaHidden,
}
impl Protection {
    fn parse(x: &str) -> Result<Self> {
        match x {
            "none" => return Ok(Self::None),
            "hidden-and-protected" => return Ok(Self::HiddenAndProtected),
            _ => {},
        }
        let mut protected = false;
        let mut formula_hidden = false;
        for word in x.split_ascii_whitespace() {
            match word {
                "protected" if !protected => protected = true,
                "formula-hidden" if !formula_hidden => formula_hidden = true,
                _ => return Err(bad("invalid style:cell-protect")),
            }
        }
        match (protected, formula_hidden) {
            (true, false) => Ok(Self::Protected),
            (false, true) => Ok(Self::FormulaHidden),
            (true, true) => Ok(Self::ProtectedFormulaHidden),
            (false, false) => Err(bad("invalid style:cell-protect")),
        }
    }
    fn xml(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::HiddenAndProtected => "hidden-and-protected",
            Self::Protected => "protected",
            Self::FormulaHidden => "formula-hidden",
            Self::ProtectedFormulaHidden => "protected formula-hidden",
        }
    }
}

/// Complete `style:table-cell-properties` value.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Properties {
    pub vertical_align: Option<VerticalAlign>,
    pub text_align_source: Option<TextAlignSource>,
    pub direction: Option<Direction>,
    pub glyph_orientation_vertical: Option<GlyphOrientation>,
    pub writing_mode: Option<WritingMode>,
    pub shadow: Option<Shadow>,
    pub background_color: Option<BackgroundColor>,
    pub border: Option<Border>,
    pub border_top: Option<Border>,
    pub border_bottom: Option<Border>,
    pub border_left: Option<Border>,
    pub border_right: Option<Border>,
    pub diagonal_tl_br: Option<Border>,
    pub diagonal_tl_br_widths: Option<BorderWidths>,
    pub diagonal_bl_tr: Option<Border>,
    pub diagonal_bl_tr_widths: Option<BorderWidths>,
    pub border_line_width: Option<BorderWidths>,
    pub border_line_width_top: Option<BorderWidths>,
    pub border_line_width_bottom: Option<BorderWidths>,
    pub border_line_width_left: Option<BorderWidths>,
    pub border_line_width_right: Option<BorderWidths>,
    pub padding: Option<Length>,
    pub padding_top: Option<Length>,
    pub padding_bottom: Option<Length>,
    pub padding_left: Option<Length>,
    pub padding_right: Option<Length>,
    pub wrap_option: Option<Wrap>,
    pub rotation_angle: Option<RotationAngle>,
    pub rotation_align: Option<RotationAlign>,
    pub cell_protect: Option<Protection>,
    pub print_content: Option<bool>,
    pub decimal_places: Option<u32>,
    pub repeat_content: Option<bool>,
    pub shrink_to_fit: Option<bool>,
    pub background_image: Option<BackgroundImage>,
}
impl Properties {
    pub fn validate(&self) -> Result<()> {
        if let Some(value) = self.decimal_places
            && value > MAX_DECIMAL_PLACES
        {
            return Err(bad("style:decimal-places is out of range"));
        }
        if let Some(image) = &self.background_image {
            image.validate()?;
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = format!(
            r#"<style:table-cell-properties xmlns:style="{STYLE_NS}" xmlns:office="{OFFICE_NS}" xmlns:fo="{FO_NS}" xmlns:draw="{DRAW_NS}" xmlns:xlink="{XLINK_NS}""#
        );
        if let Some(value) = self.vertical_align {
            xml.push_str(&format!(r#" style:vertical-align="{}""#, value.xml()));
        }
        if let Some(value) = self.text_align_source {
            xml.push_str(&format!(r#" style:text-align-source="{}""#, value.xml()));
        }
        if let Some(value) = self.direction {
            xml.push_str(&format!(r#" style:direction="{}""#, value.xml()));
        }
        if let Some(value) = &self.glyph_orientation_vertical {
            xml.push_str(&format!(
                r#" style:glyph-orientation-vertical="{}""#,
                value.as_str()
            ));
        }
        if let Some(value) = self.writing_mode {
            xml.push_str(&format!(r#" style:writing-mode="{}""#, writing_xml(value)));
        }
        if let Some(value) = &self.shadow {
            xml.push_str(&format!(
                r#" style:shadow="{}""#,
                escape_xml(value.as_str())
            ));
        }
        if let Some(value) = &self.background_color {
            xml.push_str(&format!(r#" fo:background-color="{}""#, value.as_str()));
        }
        for (name, value) in [
            ("fo:border", &self.border),
            ("fo:border-top", &self.border_top),
            ("fo:border-bottom", &self.border_bottom),
            ("fo:border-left", &self.border_left),
            ("fo:border-right", &self.border_right),
            ("style:diagonal-tl-br", &self.diagonal_tl_br),
            ("style:diagonal-bl-tr", &self.diagonal_bl_tr),
        ] {
            if let Some(value) = value {
                xml.push_str(&format!(r#" {name}="{}""#, escape_xml(value.as_str())));
            }
        }
        for (name, value) in [
            ("style:diagonal-tl-br-widths", &self.diagonal_tl_br_widths),
            ("style:diagonal-bl-tr-widths", &self.diagonal_bl_tr_widths),
            ("style:border-line-width", &self.border_line_width),
            ("style:border-line-width-top", &self.border_line_width_top),
            (
                "style:border-line-width-bottom",
                &self.border_line_width_bottom,
            ),
            ("style:border-line-width-left", &self.border_line_width_left),
            (
                "style:border-line-width-right",
                &self.border_line_width_right,
            ),
        ] {
            if let Some(value) = value {
                xml.push_str(&format!(r#" {name}="{}""#, value.xml()));
            }
        }
        for (name, value) in [
            ("fo:padding", &self.padding),
            ("fo:padding-top", &self.padding_top),
            ("fo:padding-bottom", &self.padding_bottom),
            ("fo:padding-left", &self.padding_left),
            ("fo:padding-right", &self.padding_right),
        ] {
            if let Some(value) = value {
                xml.push_str(&format!(r#" {name}="{}""#, value.as_str()));
            }
        }
        if let Some(value) = self.wrap_option {
            xml.push_str(&format!(r#" fo:wrap-option="{}""#, value.xml()));
        }
        if let Some(value) = self.rotation_angle {
            xml.push_str(&format!(r#" style:rotation-angle="{}""#, value.degrees()));
        }
        if let Some(value) = self.rotation_align {
            xml.push_str(&format!(r#" style:rotation-align="{}""#, value.xml()));
        }
        if let Some(value) = self.cell_protect {
            xml.push_str(&format!(r#" style:cell-protect="{}""#, value.xml()));
        }
        if let Some(value) = self.print_content {
            xml.push_str(&format!(r#" style:print-content="{value}""#));
        }
        if let Some(value) = self.decimal_places {
            xml.push_str(&format!(r#" style:decimal-places="{value}""#));
        }
        if let Some(value) = self.repeat_content {
            xml.push_str(&format!(r#" style:repeat-content="{value}""#));
        }
        if let Some(value) = self.shrink_to_fit {
            xml.push_str(&format!(r#" style:shrink-to-fit="{value}""#));
        }
        if let Some(image) = &self.background_image {
            xml.push('>');
            xml.push_str(&image.to_xml_fragment()?);
            xml.push_str("</style:table-cell-properties>");
        } else {
            xml.push_str("/>");
        }
        Ok(xml)
    }
}
fn parse_writing(x: &str) -> Result<WritingMode> {
    match x {
        "lr-tb" => Ok(WritingMode::LrTb),
        "rl-tb" => Ok(WritingMode::RlTb),
        "tb-rl" => Ok(WritingMode::TbRl),
        "tb-lr" => Ok(WritingMode::TbLr),
        "lr" => Ok(WritingMode::Lr),
        "rl" => Ok(WritingMode::Rl),
        "tb" => Ok(WritingMode::Tb),
        "page" => Ok(WritingMode::Page),
        _ => Err(bad("invalid style:writing-mode")),
    }
}
fn writing_xml(x: WritingMode) -> &'static str {
    match x {
        WritingMode::LrTb => "lr-tb",
        WritingMode::RlTb => "rl-tb",
        WritingMode::TbRl => "tb-rl",
        WritingMode::TbLr => "tb-lr",
        WritingMode::Lr => "lr",
        WritingMode::Rl => "rl",
        WritingMode::Tb => "tb",
        WritingMode::Page => "page",
    }
}

/// A named or default table-cell style declaration carrying typed cell properties.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Style {
    pub name: Option<String>,
    pub parent_style_name: Option<String>,
    pub is_default_style: bool,
    pub properties: Option<Properties>,
}
impl Style {
    pub fn named(name: impl Into<String>, properties: Option<Properties>) -> Result<Self> {
        let value = Self {
            name: Some(name.into()),
            parent_style_name: None,
            is_default_style: false,
            properties,
        };
        value.validate()?;
        Ok(value)
    }
    pub fn default_style(properties: Option<Properties>) -> Self {
        Self {
            name: None,
            parent_style_name: None,
            is_default_style: true,
            properties,
        }
    }
    pub fn validate(&self) -> Result<()> {
        match (&self.name, self.is_default_style) {
            (Some(value), false) => safe(value, "table-cell style name", false)?,
            (None, true) => {},
            _ => return Err(bad("invalid table-cell style identity")),
        }
        if let Some(value) = &self.parent_style_name {
            if self.is_default_style {
                return Err(bad("default table-cell style cannot have a parent"));
            }
            safe(value, "parent table-cell style name", false)?;
        }
        if let Some(value) = &self.properties {
            value.validate()?;
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let tag = if self.is_default_style {
            "default-style"
        } else {
            "style"
        };
        let mut xml = format!(r#"<style:{tag} xmlns:style="{STYLE_NS}" style:family="table-cell""#);
        if let Some(value) = &self.name {
            xml.push_str(&format!(r#" style:name="{}""#, escape_xml(value)));
        }
        if let Some(value) = &self.parent_style_name {
            xml.push_str(&format!(
                r#" style:parent-style-name="{}""#,
                escape_xml(value)
            ));
        }
        if let Some(value) = &self.properties {
            xml.push('>');
            xml.push_str(&value.to_xml_fragment()?);
            xml.push_str(&format!("</style:{tag}>"));
        } else {
            xml.push_str("/>");
        }
        Ok(xml)
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Styles {
    pub styles: Vec<Style>,
}
impl Styles {
    pub fn get(&self, name: &str) -> Option<&Style> {
        self.styles
            .iter()
            .find(|style| style.name.as_deref() == Some(name))
    }
    pub fn default_style(&self) -> Option<&Style> {
        self.styles.iter().find(|style| style.is_default_style)
    }
}

#[derive(Clone, Copy, PartialEq, Eq, Hash)]
enum Ns {
    Office,
    Style,
    Fo,
    Draw,
    Xlink,
    Other,
}
fn ns(value: ResolveResult<'_>) -> Ns {
    match value {
        ResolveResult::Bound(x) if x.as_ref() == OFFICE => Ns::Office,
        ResolveResult::Bound(x) if x.as_ref() == STYLE => Ns::Style,
        ResolveResult::Bound(x) if x.as_ref() == FO => Ns::Fo,
        ResolveResult::Bound(x) if x.as_ref() == DRAW => Ns::Draw,
        ResolveResult::Bound(x) if x.as_ref() == XLINK => Ns::Xlink,
        _ => Ns::Other,
    }
}
fn element(reader: &NsReader<&[u8]>, name: QName<'_>) -> (Ns, Vec<u8>) {
    let (namespace, local) = reader.resolver().resolve_element(name);
    (ns(namespace), local.as_ref().to_vec())
}
fn attributes(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<Vec<(Ns, Vec<u8>, String)>> {
    let mut result = Vec::new();
    let mut seen = HashSet::new();
    for attribute in start.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| bad(format!("invalid table-cell attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        if result.len() >= MAX_ATTRIBUTES {
            return Err(bad("too many table-cell attributes"));
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let key = (ns(namespace), local.as_ref().to_vec());
        if !seen.insert(key.clone()) {
            return Err(bad("duplicate table-cell attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| bad(format!("invalid table-cell value: {error}")))?
            .into_owned();
        if value.len() > MAX_VALUE {
            return Err(bad("table-cell attribute value is too large"));
        }
        result.push((key.0, key.1, value));
    }
    Ok(result)
}
fn take(attrs: &mut Vec<(Ns, Vec<u8>, String)>, namespace: Ns, local: &[u8]) -> Option<String> {
    attrs
        .iter()
        .position(|x| x.0 == namespace && x.1 == local)
        .map(|at| attrs.remove(at).2)
}
fn boolean(value: &str) -> Result<bool> {
    match value {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(bad("ODF boolean must be true or false")),
    }
}
fn style_header(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
    default: bool,
) -> Result<Option<Style>> {
    let mut attrs = attributes(reader, version, start)?;
    if take(&mut attrs, Ns::Style, b"family").as_deref() != Some("table-cell") {
        return Ok(None);
    }
    let value = Style {
        name: take(&mut attrs, Ns::Style, b"name"),
        parent_style_name: take(&mut attrs, Ns::Style, b"parent-style-name"),
        is_default_style: default,
        properties: None,
    };
    value.validate()?;
    Ok(Some(value))
}
fn cell_properties(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<Properties> {
    let mut attrs = attributes(reader, version, start)?;
    let value = Properties {
        vertical_align: take(&mut attrs, Ns::Style, b"vertical-align")
            .map(|x| VerticalAlign::parse(&x))
            .transpose()?,
        text_align_source: take(&mut attrs, Ns::Style, b"text-align-source")
            .map(|x| TextAlignSource::parse(&x))
            .transpose()?,
        direction: take(&mut attrs, Ns::Style, b"direction")
            .map(|x| Direction::parse(&x))
            .transpose()?,
        glyph_orientation_vertical: take(&mut attrs, Ns::Style, b"glyph-orientation-vertical")
            .map(GlyphOrientation::new)
            .transpose()?,
        writing_mode: take(&mut attrs, Ns::Style, b"writing-mode")
            .map(|x| parse_writing(&x))
            .transpose()?,
        shadow: take(&mut attrs, Ns::Style, b"shadow")
            .map(Shadow::new)
            .transpose()?,
        background_color: take(&mut attrs, Ns::Fo, b"background-color")
            .map(BackgroundColor::new)
            .transpose()?,
        border: take(&mut attrs, Ns::Fo, b"border")
            .map(Border::new)
            .transpose()?,
        border_top: take(&mut attrs, Ns::Fo, b"border-top")
            .map(Border::new)
            .transpose()?,
        border_bottom: take(&mut attrs, Ns::Fo, b"border-bottom")
            .map(Border::new)
            .transpose()?,
        border_left: take(&mut attrs, Ns::Fo, b"border-left")
            .map(Border::new)
            .transpose()?,
        border_right: take(&mut attrs, Ns::Fo, b"border-right")
            .map(Border::new)
            .transpose()?,
        diagonal_tl_br: take(&mut attrs, Ns::Style, b"diagonal-tl-br")
            .map(Border::new)
            .transpose()?,
        diagonal_tl_br_widths: take(&mut attrs, Ns::Style, b"diagonal-tl-br-widths")
            .map(|x| BorderWidths::parse(&x))
            .transpose()?,
        diagonal_bl_tr: take(&mut attrs, Ns::Style, b"diagonal-bl-tr")
            .map(Border::new)
            .transpose()?,
        diagonal_bl_tr_widths: take(&mut attrs, Ns::Style, b"diagonal-bl-tr-widths")
            .map(|x| BorderWidths::parse(&x))
            .transpose()?,
        border_line_width: take(&mut attrs, Ns::Style, b"border-line-width")
            .map(|x| BorderWidths::parse(&x))
            .transpose()?,
        border_line_width_top: take(&mut attrs, Ns::Style, b"border-line-width-top")
            .map(|x| BorderWidths::parse(&x))
            .transpose()?,
        border_line_width_bottom: take(&mut attrs, Ns::Style, b"border-line-width-bottom")
            .map(|x| BorderWidths::parse(&x))
            .transpose()?,
        border_line_width_left: take(&mut attrs, Ns::Style, b"border-line-width-left")
            .map(|x| BorderWidths::parse(&x))
            .transpose()?,
        border_line_width_right: take(&mut attrs, Ns::Style, b"border-line-width-right")
            .map(|x| BorderWidths::parse(&x))
            .transpose()?,
        padding: take(&mut attrs, Ns::Fo, b"padding")
            .map(Length::non_negative)
            .transpose()?,
        padding_top: take(&mut attrs, Ns::Fo, b"padding-top")
            .map(Length::non_negative)
            .transpose()?,
        padding_bottom: take(&mut attrs, Ns::Fo, b"padding-bottom")
            .map(Length::non_negative)
            .transpose()?,
        padding_left: take(&mut attrs, Ns::Fo, b"padding-left")
            .map(Length::non_negative)
            .transpose()?,
        padding_right: take(&mut attrs, Ns::Fo, b"padding-right")
            .map(Length::non_negative)
            .transpose()?,
        wrap_option: take(&mut attrs, Ns::Fo, b"wrap-option")
            .map(|x| Wrap::parse(&x))
            .transpose()?,
        rotation_angle: take(&mut attrs, Ns::Style, b"rotation-angle")
            .map(|x| RotationAngle::parse(&x))
            .transpose()?,
        rotation_align: take(&mut attrs, Ns::Style, b"rotation-align")
            .map(|x| RotationAlign::parse(&x))
            .transpose()?,
        cell_protect: take(&mut attrs, Ns::Style, b"cell-protect")
            .map(|x| Protection::parse(&x))
            .transpose()?,
        print_content: take(&mut attrs, Ns::Style, b"print-content")
            .map(|x| boolean(&x))
            .transpose()?,
        decimal_places: take(&mut attrs, Ns::Style, b"decimal-places")
            .map(|x| {
                x.parse::<u32>()
                    .map_err(|_| bad("invalid style:decimal-places"))
            })
            .transpose()?,
        repeat_content: take(&mut attrs, Ns::Style, b"repeat-content")
            .map(|x| boolean(&x))
            .transpose()?,
        shrink_to_fit: take(&mut attrs, Ns::Style, b"shrink-to-fit")
            .map(|x| boolean(&x))
            .transpose()?,
        background_image: None,
    };
    if !attrs.is_empty() {
        return Err(bad("unknown style:table-cell-properties attribute"));
    }
    value.validate()?;
    Ok(value)
}
fn position(x: &str) -> Result<BackgroundPosition> {
    let words: Vec<_> = x.split_ascii_whitespace().collect();
    let horizontal = |x| match x {
        "left" => Some(HorizontalBackgroundPosition::Left),
        "center" => Some(HorizontalBackgroundPosition::Center),
        "right" => Some(HorizontalBackgroundPosition::Right),
        _ => None,
    };
    let vertical = |x| match x {
        "top" => Some(VerticalBackgroundPosition::Top),
        "center" => Some(VerticalBackgroundPosition::Center),
        "bottom" => Some(VerticalBackgroundPosition::Bottom),
        _ => None,
    };
    match words.as_slice() {
        ["left"] => Ok(BackgroundPosition::Left),
        ["center"] => Ok(BackgroundPosition::Center),
        ["right"] => Ok(BackgroundPosition::Right),
        ["top"] => Ok(BackgroundPosition::Top),
        ["bottom"] => Ok(BackgroundPosition::Bottom),
        [a, b] => horizontal(a)
            .zip(vertical(b))
            .or_else(|| horizontal(b).zip(vertical(a)))
            .map(|(h, v)| BackgroundPosition::Pair(h, v))
            .ok_or_else(|| bad("invalid background position")),
        _ => Err(bad("invalid background position")),
    }
}
struct ParsedImage {
    image: BackgroundImage,
    linked: bool,
}
fn image_attributes(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<ParsedImage> {
    let mut attrs = attributes(reader, version, start)?;
    let repeat = take(&mut attrs, Ns::Style, b"repeat")
        .map(|x| match x.as_str() {
            "no-repeat" => Ok(Repeat::NoRepeat),
            "repeat" => Ok(Repeat::Repeat),
            "stretch" => Ok(Repeat::Stretch),
            _ => Err(bad("invalid background repeat")),
        })
        .transpose()?;
    let pos = take(&mut attrs, Ns::Style, b"position")
        .map(|x| position(&x))
        .transpose()?;
    let filter_name = take(&mut attrs, Ns::Style, b"filter-name");
    if let Some(x) = &filter_name {
        safe(x, "style:filter-name", true)?;
    }
    let opacity = take(&mut attrs, Ns::Draw, b"opacity")
        .map(Opacity::new)
        .transpose()?;
    let kind = take(&mut attrs, Ns::Xlink, b"type");
    let href = take(&mut attrs, Ns::Xlink, b"href");
    let show = take(&mut attrs, Ns::Xlink, b"show");
    let actuate = take(&mut attrs, Ns::Xlink, b"actuate");
    if !attrs.is_empty() {
        return Err(bad("unknown style:background-image attribute"));
    }
    let linked = kind.is_some() || href.is_some() || show.is_some() || actuate.is_some();
    let source = if linked {
        if kind.as_deref() != Some("simple")
            || href.is_none()
            || show.as_deref().is_some_and(|x| x != "embed")
            || actuate.as_deref().is_some_and(|x| x != "onLoad")
        {
            return Err(bad("invalid background-image xlink group"));
        }
        BackgroundSource::Link {
            href: href.unwrap(),
            show_embed: show.is_some(),
            actuate_on_load: actuate.is_some(),
        }
    } else {
        BackgroundSource::Empty
    };
    let image = BackgroundImage {
        repeat,
        position: pos,
        filter_name,
        opacity,
        source,
    };
    image.validate()?;
    Ok(ParsedImage { image, linked })
}

struct Active {
    depth: usize,
    style: Style,
    seen_properties: bool,
    properties_depth: Option<usize>,
    image_depth: Option<usize>,
    binary_depth: Option<usize>,
    binary: String,
    image_linked: bool,
}
fn push_style(out: &mut Vec<Style>, style: Style, total: &mut usize) -> Result<()> {
    if out.len() >= MAX_STYLES
        || out
            .iter()
            .any(|x| x.name == style.name && x.is_default_style == style.is_default_style)
    {
        return Err(bad("duplicate or excessive table-cell style"));
    }
    *total += style.to_xml_fragment()?.len();
    if *total > MAX_TOTAL {
        return Err(bad("table-cell style data is too large"));
    }
    out.push(style);
    Ok(())
}

/// Parse direct table-cell styles in `office:styles` and `office:automatic-styles`.
pub fn parse(xml: &str) -> Result<Styles> {
    if xml.len() > MAX_XML {
        return Err(bad("styles XML is too large"));
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut active: Option<Active> = None;
    let mut out = Vec::new();
    let mut total = 0;
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(bad("styles XML nesting is too deep"));
                }
                let current = element(&reader, start.name());
                let parent = stack.last();
                let direct = parent.is_some_and(|x| {
                    x.0 == Ns::Office && matches!(x.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                stack.push(current.clone());
                let depth = stack.len();
                if direct {
                    if let Some(style) =
                        style_header(&reader, version, &start, current.1 == b"default-style")?
                    {
                        active = Some(Active {
                            depth,
                            style,
                            seen_properties: false,
                            properties_depth: None,
                            image_depth: None,
                            binary_depth: None,
                            binary: String::new(),
                            image_linked: false,
                        });
                    }
                    continue;
                }
                if let Some(state) = active.as_mut() {
                    if depth == state.depth + 1
                        && current.0 == Ns::Style
                        && current.1 == b"table-cell-properties"
                    {
                        if state.seen_properties {
                            return Err(bad("duplicate style:table-cell-properties"));
                        }
                        state.seen_properties = true;
                        state.style.properties = Some(cell_properties(&reader, version, &start)?);
                        state.properties_depth = Some(depth);
                    } else if current.1 == b"table-cell-properties" {
                        return Err(bad(
                            "style:table-cell-properties has invalid namespace or parent",
                        ));
                    } else if state.properties_depth.is_some()
                        && depth == state.properties_depth.unwrap() + 1
                        && current.0 == Ns::Style
                        && current.1 == b"background-image"
                    {
                        if state.image_depth.is_some()
                            || state
                                .style
                                .properties
                                .as_ref()
                                .unwrap()
                                .background_image
                                .is_some()
                        {
                            return Err(bad("duplicate style:background-image"));
                        }
                        let parsed = image_attributes(&reader, version, &start)?;
                        state.style.properties.as_mut().unwrap().background_image =
                            Some(parsed.image);
                        state.image_linked = parsed.linked;
                        state.image_depth = Some(depth);
                    } else if current.1 == b"background-image" {
                        return Err(bad(
                            "style:background-image has invalid namespace or parent",
                        ));
                    } else if state.image_depth.is_some()
                        && depth == state.image_depth.unwrap() + 1
                        && current.0 == Ns::Office
                        && current.1 == b"binary-data"
                    {
                        if state.image_linked || state.binary_depth.is_some() {
                            return Err(bad("invalid office:binary-data in background image"));
                        }
                        state.binary_depth = Some(depth);
                        state.binary.clear();
                    } else if state.properties_depth.is_some()
                        && depth > state.properties_depth.unwrap()
                    {
                        return Err(bad("unexpected table-cell property child"));
                    }
                }
            },
            Ok(Event::Empty(start)) => {
                let current = element(&reader, start.name());
                let parent = stack.last();
                let depth = stack.len() + 1;
                let direct = parent.is_some_and(|x| {
                    x.0 == Ns::Office && matches!(x.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                if direct {
                    if let Some(style) =
                        style_header(&reader, version, &start, current.1 == b"default-style")?
                    {
                        push_style(&mut out, style, &mut total)?;
                    }
                    continue;
                }
                if let Some(state) = active.as_mut() {
                    if depth == state.depth + 1
                        && current.0 == Ns::Style
                        && current.1 == b"table-cell-properties"
                    {
                        if state.seen_properties {
                            return Err(bad("duplicate style:table-cell-properties"));
                        }
                        state.seen_properties = true;
                        state.style.properties = Some(cell_properties(&reader, version, &start)?);
                    } else if current.1 == b"table-cell-properties" {
                        return Err(bad(
                            "style:table-cell-properties has invalid namespace or parent",
                        ));
                    } else if state.properties_depth.is_some()
                        && depth == state.properties_depth.unwrap() + 1
                        && current.0 == Ns::Style
                        && current.1 == b"background-image"
                    {
                        if state
                            .style
                            .properties
                            .as_ref()
                            .unwrap()
                            .background_image
                            .is_some()
                        {
                            return Err(bad("duplicate style:background-image"));
                        }
                        state.style.properties.as_mut().unwrap().background_image =
                            Some(image_attributes(&reader, version, &start)?.image);
                    } else if current.1 == b"background-image" {
                        return Err(bad(
                            "style:background-image has invalid namespace or parent",
                        ));
                    } else if state.image_depth.is_some()
                        && depth == state.image_depth.unwrap() + 1
                        && current.0 == Ns::Office
                        && current.1 == b"binary-data"
                    {
                        if state.image_linked {
                            return Err(bad("linked background image cannot contain binary data"));
                        }
                        state
                            .style
                            .properties
                            .as_mut()
                            .unwrap()
                            .background_image
                            .as_mut()
                            .unwrap()
                            .source = BackgroundSource::Embedded(Vec::new());
                    } else if state.properties_depth.is_some()
                        && depth > state.properties_depth.unwrap()
                    {
                        return Err(bad("unexpected table-cell property child"));
                    }
                }
            },
            Ok(Event::Text(text)) => {
                let bytes: &[u8] = text.as_ref();
                if let Some(state) = active.as_mut() {
                    if state.binary_depth.is_some() {
                        if state.binary.len() + bytes.len() > MAX_BINARY * 2 {
                            return Err(bad("encoded office:binary-data is too large"));
                        }
                        state.binary.push_str(&String::from_utf8_lossy(bytes));
                    } else if state.properties_depth.is_some()
                        && !bytes.iter().all(u8::is_ascii_whitespace)
                    {
                        return Err(bad("unexpected text in table-cell properties"));
                    }
                }
            },
            Ok(Event::CData(text)) => {
                let bytes: &[u8] = text.as_ref();
                if let Some(state) = active.as_mut() {
                    if state.binary_depth.is_some() {
                        if state.binary.len() + bytes.len() > MAX_BINARY * 2 {
                            return Err(bad("encoded office:binary-data is too large"));
                        }
                        state.binary.push_str(&String::from_utf8_lossy(bytes));
                    } else if state.properties_depth.is_some()
                        && !bytes.iter().all(u8::is_ascii_whitespace)
                    {
                        return Err(bad("unexpected text in table-cell properties"));
                    }
                }
            },
            Ok(Event::End(_)) => {
                let depth = stack.len();
                if let Some(state) = active.as_mut() {
                    if state.binary_depth == Some(depth) {
                        let data = base64_decode(&state.binary)?;
                        state
                            .style
                            .properties
                            .as_mut()
                            .unwrap()
                            .background_image
                            .as_mut()
                            .unwrap()
                            .source = BackgroundSource::Embedded(data);
                        state.binary_depth = None;
                    }
                    if state.image_depth == Some(depth) {
                        state.image_depth = None;
                    }
                    if state.properties_depth == Some(depth) {
                        state.properties_depth = None;
                    }
                }
                if active.as_ref().is_some_and(|x| x.depth == depth) {
                    push_style(&mut out, active.take().unwrap().style, &mut total)?;
                }
                stack.pop();
            },
            Ok(Event::Decl(decl)) => {
                version = decl
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(bad(format!("invalid styles XML: {error}"))),
        }
    }
    if !stack.is_empty() || active.is_some() {
        return Err(bad("truncated styles XML"));
    }
    Ok(Styles { styles: out })
}

#[derive(Default)]
struct Span {
    start: usize,
    end: usize,
    end_start: usize,
    qname: String,
    empty: bool,
}
#[derive(Default)]
struct TargetSpans {
    style: Span,
    properties: Option<Span>,
}
fn boundary(xml: &str, end: usize) -> Result<usize> {
    xml[..end]
        .rfind('<')
        .ok_or_else(|| bad("invalid XML event boundary"))
}
fn replace_span(xml: &str, span: &Span, value: &str) -> String {
    format!("{}{}{}", &xml[..span.start], value, &xml[span.end..])
}
fn expand_span(xml: &str, span: &Span, value: &str) -> Result<String> {
    let raw = &xml[span.start..span.end];
    let slash = raw
        .rfind("/>")
        .ok_or_else(|| bad("invalid empty element"))?;
    Ok(replace_span(
        xml,
        span,
        &format!("{}>{value}</{}>", &raw[..slash], span.qname),
    ))
}

/// Losslessly replace, insert, or remove one existing cell style's property element.
pub fn set_xml(xml: &str, requested: &Style) -> Result<String> {
    requested.validate()?;
    if xml.len() > MAX_XML {
        return Err(bad("styles XML is too large"));
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut depth_target = None;
    let mut active: Option<TargetSpans> = None;
    let mut found = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let direct = parent.is_some_and(|x| {
                    x.0 == Ns::Office && matches!(x.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                stack.push(current.clone());
                let depth = stack.len();
                if direct {
                    if let Some(style) =
                        style_header(&reader, version, &start, current.1 == b"default-style")?
                        && style.name == requested.name
                        && style.is_default_style == requested.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target table-cell style"));
                        }
                        depth_target = Some(depth);
                        active = Some(TargetSpans {
                            style: Span {
                                start: begin,
                                qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                                ..Default::default()
                            },
                            ..Default::default()
                        });
                    }
                } else if depth_target.is_some_and(|d| depth == d + 1)
                    && current.0 == Ns::Style
                    && current.1 == b"table-cell-properties"
                {
                    let span = Span {
                        start: begin,
                        qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                        ..Default::default()
                    };
                    if active.as_mut().unwrap().properties.replace(span).is_some() {
                        return Err(bad("duplicate style:table-cell-properties"));
                    }
                }
            },
            Ok(Event::Empty(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let depth = stack.len() + 1;
                let direct = parent.is_some_and(|x| {
                    x.0 == Ns::Office && matches!(x.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                let span = Span {
                    start: begin,
                    end,
                    end_start: begin,
                    qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                    empty: true,
                };
                if direct {
                    if let Some(style) =
                        style_header(&reader, version, &start, current.1 == b"default-style")?
                        && style.name == requested.name
                        && style.is_default_style == requested.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target table-cell style"));
                        }
                        found = Some(TargetSpans {
                            style: span,
                            ..Default::default()
                        });
                    }
                } else if depth_target.is_some_and(|d| depth == d + 1)
                    && current.0 == Ns::Style
                    && current.1 == b"table-cell-properties"
                    && active.as_mut().unwrap().properties.replace(span).is_some()
                {
                    return Err(bad("duplicate style:table-cell-properties"));
                }
            },
            Ok(Event::End(_)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let depth = stack.len();
                if let Some(spans) = active.as_mut() {
                    if spans.properties.as_ref().is_some_and(|s| s.end == 0)
                        && depth_target.is_some_and(|d| depth == d + 1)
                    {
                        let s = spans.properties.as_mut().unwrap();
                        s.end_start = begin;
                        s.end = end;
                    }
                    if depth_target == Some(depth) {
                        spans.style.end_start = begin;
                        spans.style.end = end;
                        found = active.take();
                        depth_target = None;
                    }
                }
                stack.pop();
            },
            Ok(Event::Decl(decl)) => {
                version = decl
                    .xml_version()
                    .map_err(|e| bad(format!("unsupported XML version: {e}")))?
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(e) => return Err(bad(format!("invalid styles XML: {e}"))),
        }
    }
    let spans = found.ok_or_else(|| bad("target table-cell style does not exist"))?;
    let replacement = requested
        .properties
        .as_ref()
        .map(Properties::to_xml_fragment)
        .transpose()?;
    if let Some(properties) = &spans.properties {
        return Ok(replace_span(
            xml,
            properties,
            replacement.as_deref().unwrap_or(""),
        ));
    }
    let Some(replacement) = replacement else {
        return Ok(xml.to_owned());
    };
    if spans.style.empty {
        return expand_span(xml, &spans.style, &replacement);
    }
    let mut out = xml.to_owned();
    out.insert_str(spans.style.end_start, &replacement);
    Ok(out)
}

fn base64_decode(value: &str) -> Result<Vec<u8>> {
    let clean: Vec<u8> = value.bytes().filter(|x| !x.is_ascii_whitespace()).collect();
    if !clean.len().is_multiple_of(4) {
        return Err(bad("invalid office:binary-data base64"));
    }
    let val = |x| match x {
        b'A'..=b'Z' => Some(x - b'A'),
        b'a'..=b'z' => Some(x - b'a' + 26),
        b'0'..=b'9' => Some(x - b'0' + 52),
        b'+' => Some(62),
        b'/' => Some(63),
        _ => None,
    };
    let mut out = Vec::with_capacity(clean.len() / 4 * 3);
    for (index, chunk) in clean.chunks(4).enumerate() {
        let last = index + 1 == clean.len() / 4;
        let pad = (chunk[2] == b'=') as usize + (chunk[3] == b'=') as usize;
        if !last && pad != 0 || chunk[2] == b'=' && chunk[3] != b'=' || pad > 2 {
            return Err(bad("invalid office:binary-data padding"));
        }
        let a = val(chunk[0]).ok_or_else(|| bad("invalid office:binary-data base64"))? as u32;
        let b = val(chunk[1]).ok_or_else(|| bad("invalid office:binary-data base64"))? as u32;
        let c = if chunk[2] == b'=' {
            0
        } else {
            val(chunk[2]).ok_or_else(|| bad("invalid office:binary-data base64"))? as u32
        };
        let d = if chunk[3] == b'=' {
            0
        } else {
            val(chunk[3]).ok_or_else(|| bad("invalid office:binary-data base64"))? as u32
        };
        let n = a << 18 | b << 12 | c << 6 | d;
        out.push((n >> 16) as u8);
        if pad < 2 {
            out.push((n >> 8) as u8)
        }
        if pad < 1 {
            out.push(n as u8)
        }
        if out.len() > MAX_BINARY {
            return Err(bad("office:binary-data is too large"));
        }
    }
    Ok(out)
}

impl Package {
    pub fn cell_style_properties(&self) -> Result<Styles> {
        self.styles_xml()?
            .map_or_else(|| Ok(Default::default()), |xml| parse(&xml))
    }
}
impl FlatDocument {
    pub fn cell_style_properties(&self) -> Result<Styles> {
        parse(self.xml())
    }
}
