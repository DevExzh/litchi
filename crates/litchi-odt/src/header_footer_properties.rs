//! Typed ODF page header/footer layout properties.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::ResolveResult,
    reader::NsReader,
};
use std::collections::HashSet;

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FO: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const SVG: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_VALUE: usize = 4096;
const MAX_TOTAL: usize = 16 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_LAYOUTS: usize = 65_536;
const MAX_ATTRIBUTES: usize = 64;

fn bad(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('"', "&quot;")
        .replace('>', "&gt;")
}
fn ns(result: ResolveResult<'_>) -> Vec<u8> {
    match result {
        ResolveResult::Bound(value) => value.as_ref().to_vec(),
        _ => Vec::new(),
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Length(String);
impl Length {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if value.is_empty() || value.len() > MAX_VALUE || value.chars().any(char::is_control) {
            return Err(bad("invalid header/footer length"));
        }
        if value == "0" || value == "+0" || value == "-0" {
            return Ok(Self(value));
        }
        let unit = ["cm", "mm", "in", "pt", "pc", "px"]
            .into_iter()
            .find(|unit| value.ends_with(unit))
            .ok_or_else(|| bad("header/footer length requires a physical unit"))?;
        let number = &value[..value.len() - unit.len()];
        if number.is_empty()
            || number.contains(['e', 'E'])
            || number
                .parse::<f64>()
                .map_or(true, |number| !number.is_finite())
        {
            return Err(bad("invalid header/footer length number"));
        }
        Ok(Self(value))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
    fn nonnegative(value: String, name: &str) -> Result<Self> {
        let parsed = Self::new(value)?;
        if parsed.0.starts_with('-') && parsed.0 != "-0" {
            return Err(bad(format!("{name} must be nonnegative")));
        }
        Ok(parsed)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Color {
    Transparent,
    Rgb(u8, u8, u8),
}
impl Color {
    fn parse(value: &str, transparent: bool) -> Result<Self> {
        if transparent && value == "transparent" {
            return Ok(Self::Transparent);
        }
        if value.len() != 7 || !value.starts_with('#') {
            return Err(bad("color must be #RRGGBB"));
        }
        let rgb = u32::from_str_radix(&value[1..], 16).map_err(|_| bad("invalid color"))?;
        Ok(Self::Rgb((rgb >> 16) as u8, (rgb >> 8) as u8, rgb as u8))
    }
    fn xml(self) -> String {
        match self {
            Self::Transparent => "transparent".into(),
            Self::Rgb(r, g, b) => format!("#{r:02X}{g:02X}{b:02X}"),
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BorderStyle {
    Hidden,
    Dotted,
    Dashed,
    Solid,
    Double,
    Groove,
    Ridge,
    Inset,
    Outset,
}
impl BorderStyle {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "hidden" => Ok(Self::Hidden),
            "dotted" => Ok(Self::Dotted),
            "dashed" => Ok(Self::Dashed),
            "solid" => Ok(Self::Solid),
            "double" => Ok(Self::Double),
            "groove" => Ok(Self::Groove),
            "ridge" => Ok(Self::Ridge),
            "inset" => Ok(Self::Inset),
            "outset" => Ok(Self::Outset),
            _ => Err(bad("invalid header/footer border style")),
        }
    }
    fn xml(self) -> &'static str {
        match self {
            Self::Hidden => "hidden",
            Self::Dotted => "dotted",
            Self::Dashed => "dashed",
            Self::Solid => "solid",
            Self::Double => "double",
            Self::Groove => "groove",
            Self::Ridge => "ridge",
            Self::Inset => "inset",
            Self::Outset => "outset",
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Border {
    None,
    Line {
        width: Length,
        style: BorderStyle,
        color: Color,
    },
}
impl Border {
    fn parse(value: &str) -> Result<Self> {
        if value == "none" {
            return Ok(Self::None);
        }
        let parts: Vec<_> = value.split_ascii_whitespace().collect();
        if parts.len() != 3 {
            return Err(bad("border must be none or width style color"));
        }
        Ok(Self::Line {
            width: Length::nonnegative(parts[0].into(), "border width")?,
            style: BorderStyle::parse(parts[1])?,
            color: Color::parse(parts[2], false)?,
        })
    }
    fn xml(&self) -> String {
        match self {
            Self::None => "none".into(),
            Self::Line {
                width,
                style,
                color,
            } => format!("{} {} {}", width.as_str(), style.xml(), color.xml()),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct BorderLineWidth {
    pub inner: Length,
    pub spacing: Length,
    pub outer: Length,
}
impl BorderLineWidth {
    fn parse(value: &str) -> Result<Self> {
        let parts: Vec<_> = value.split_ascii_whitespace().collect();
        if parts.len() != 3 {
            return Err(bad("border-line-width requires three lengths"));
        }
        Ok(Self {
            inner: Length::nonnegative(parts[0].into(), "inner border width")?,
            spacing: Length::nonnegative(parts[1].into(), "border spacing")?,
            outer: Length::nonnegative(parts[2].into(), "outer border width")?,
        })
    }
    fn xml(&self) -> String {
        format!(
            "{} {} {}",
            self.inner.as_str(),
            self.spacing.as_str(),
            self.outer.as_str()
        )
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Shadow {
    None,
    Drop {
        color: Color,
        offset_x: Length,
        offset_y: Length,
    },
}
impl Shadow {
    fn parse(value: &str) -> Result<Self> {
        if value == "none" {
            return Ok(Self::None);
        }
        let parts: Vec<_> = value.split_ascii_whitespace().collect();
        if parts.len() != 3 {
            return Err(bad("shadow must be none or color x-offset y-offset"));
        }
        Ok(Self::Drop {
            color: Color::parse(parts[0], false)?,
            offset_x: Length::new(parts[1])?,
            offset_y: Length::new(parts[2])?,
        })
    }
    fn xml(&self) -> String {
        match self {
            Self::None => "none".into(),
            Self::Drop {
                color,
                offset_x,
                offset_y,
            } => format!(
                "{} {} {}",
                color.xml(),
                offset_x.as_str(),
                offset_y.as_str()
            ),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Edges<T> {
    pub all: Option<T>,
    pub left: Option<T>,
    pub right: Option<T>,
    pub top: Option<T>,
    pub bottom: Option<T>,
}
impl<T> Default for Edges<T> {
    fn default() -> Self {
        Self {
            all: None,
            left: None,
            right: None,
            top: None,
            bottom: None,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Region {
    Header,
    Footer,
}
impl Region {
    fn wrapper(self) -> &'static str {
        match self {
            Self::Header => "header-style",
            Self::Footer => "footer-style",
        }
    }
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct StyleProperties {
    pub height: Option<Length>,
    pub min_height: Option<Length>,
    pub margins: Edges<Length>,
    pub borders: Edges<Border>,
    pub border_line_widths: Edges<BorderLineWidth>,
    pub padding: Edges<Length>,
    pub background_color: Option<Color>,
    pub shadow: Option<Shadow>,
    pub dynamic_spacing: Option<bool>,
    pub background_image: Option<crate::SectionBackgroundImage>,
}

impl StyleProperties {
    pub fn validate(&self) -> Result<()> {
        if let Some(image) = &self.background_image {
            image.validate()?;
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = format!(
            "<style:header-footer-properties xmlns:style=\"{}\" xmlns:fo=\"{}\" xmlns:svg=\"{}\" xmlns:draw=\"{}\" xmlns:xlink=\"{}\"",
            String::from_utf8_lossy(STYLE),
            String::from_utf8_lossy(FO),
            String::from_utf8_lossy(SVG),
            String::from_utf8_lossy(DRAW),
            String::from_utf8_lossy(XLINK)
        );
        if let Some(v) = &self.height {
            attr(&mut xml, "svg:height", v.as_str());
        }
        if let Some(v) = &self.min_height {
            attr(&mut xml, "fo:min-height", v.as_str());
        }
        write_edges(&mut xml, "fo:margin", &self.margins, |v| {
            v.as_str().to_string()
        });
        write_edges(&mut xml, "fo:border", &self.borders, Border::xml);
        write_edges(
            &mut xml,
            "style:border-line-width",
            &self.border_line_widths,
            BorderLineWidth::xml,
        );
        write_edges(&mut xml, "fo:padding", &self.padding, |v| {
            v.as_str().to_string()
        });
        if let Some(v) = self.background_color {
            attr(&mut xml, "fo:background-color", &v.xml());
        }
        if let Some(v) = &self.shadow {
            attr(&mut xml, "style:shadow", &v.xml());
        }
        if let Some(v) = self.dynamic_spacing {
            attr(
                &mut xml,
                "style:dynamic-spacing",
                if v { "true" } else { "false" },
            );
        }
        if let Some(image) = &self.background_image {
            xml.push('>');
            write_background(&mut xml, image)?;
            xml.push_str("</style:header-footer-properties>");
        } else {
            xml.push_str("/>");
        }
        Ok(xml)
    }
    pub fn to_region_fragment(&self, region: Region) -> Result<String> {
        Ok(format!(
            "<style:{} xmlns:style=\"{}\">{}</style:{}>",
            region.wrapper(),
            String::from_utf8_lossy(STYLE),
            self.to_xml_fragment()?,
            region.wrapper()
        ))
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Properties {
    pub page_layout_name: String,
    pub region: Region,
    pub properties: StyleProperties,
}

fn attr(xml: &mut String, name: &str, value: &str) {
    xml.push(' ');
    xml.push_str(name);
    xml.push_str("=\"");
    xml.push_str(&escape(value));
    xml.push('"');
}
fn write_edges<T>(xml: &mut String, base: &str, edges: &Edges<T>, lexical: impl Fn(&T) -> String) {
    if let Some(v) = &edges.all {
        attr(xml, base, &lexical(v));
    }
    for (suffix, value) in [
        ("left", &edges.left),
        ("right", &edges.right),
        ("top", &edges.top),
        ("bottom", &edges.bottom),
    ] {
        if let Some(v) = value {
            attr(xml, &format!("{base}-{suffix}"), &lexical(v));
        }
    }
}
fn write_background(xml: &mut String, image: &crate::SectionBackgroundImage) -> Result<()> {
    image.validate()?;
    xml.push_str("<style:background-image");
    if let Some(v) = &image.href {
        attr(xml, "xlink:href", v);
    }
    if let Some(v) = image.repeat {
        attr(
            xml,
            "style:repeat",
            match v {
                crate::BackgroundRepeat::Repeat => "repeat",
                crate::BackgroundRepeat::Stretch => "stretch",
                crate::BackgroundRepeat::NoRepeat => "no-repeat",
            },
        );
    }
    if let Some(v) = &image.position {
        attr(xml, "style:position", v);
    }
    if let Some(v) = &image.filter_name {
        attr(xml, "style:filter-name", v);
    }
    if let Some(v) = image.opacity_percent {
        attr(xml, "draw:opacity", &format!("{v}%"));
    }
    if let Some(v) = &image.xlink_type {
        attr(xml, "xlink:type", v);
    }
    if let Some(v) = &image.show {
        attr(xml, "xlink:show", v);
    }
    if let Some(v) = &image.actuate {
        attr(xml, "xlink:actuate", v);
    }
    xml.push_str("/>");
    Ok(())
}

#[allow(clippy::type_complexity)]
fn attributes(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    element: &BytesStart<'_>,
    total: &mut usize,
) -> Result<Vec<(Vec<u8>, Vec<u8>, String)>> {
    if element.attributes().count() > MAX_ATTRIBUTES {
        return Err(bad("header/footer properties exceeds attribute cap"));
    }
    let mut out = Vec::new();
    let mut seen = HashSet::new();
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(|e| bad(format!("invalid header/footer attribute: {e}")))?;
        if item.key.as_ref() == b"xmlns" || item.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (n, l) = reader.resolver().resolve_attribute(item.key);
        let key = (ns(n), l.as_ref().to_vec());
        if !seen.insert(key.clone()) {
            return Err(bad("duplicate expanded header/footer attribute"));
        }
        let value = item
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|e| bad(format!("invalid header/footer value: {e}")))?
            .into_owned();
        if value.len() > MAX_VALUE {
            return Err(bad("header/footer value exceeds cap"));
        }
        *total = total
            .checked_add(value.len())
            .ok_or_else(|| bad("header/footer aggregate overflow"))?;
        if *total > MAX_TOTAL {
            return Err(bad("header/footer aggregate exceeds cap"));
        }
        out.push((key.0, key.1, value));
    }
    Ok(out)
}
fn take(a: &mut Vec<(Vec<u8>, Vec<u8>, String)>, n: &[u8], l: &[u8]) -> Option<String> {
    a.iter()
        .position(|x| x.0 == n && x.1 == l)
        .map(|i| a.remove(i).2)
}
fn parse_edges<T>(
    a: &mut Vec<(Vec<u8>, Vec<u8>, String)>,
    namespace: &[u8],
    base: &[u8],
    parse: impl Fn(String) -> Result<T>,
) -> Result<Edges<T>> {
    let mut name = base.to_vec();
    let all = take(a, namespace, &name).map(&parse).transpose()?;
    let mut side = |suffix: &[u8]| {
        name = base.to_vec();
        name.push(b'-');
        name.extend_from_slice(suffix);
        take(a, namespace, &name).map(&parse).transpose()
    };
    Ok(Edges {
        all,
        left: side(b"left")?,
        right: side(b"right")?,
        top: side(b"top")?,
        bottom: side(b"bottom")?,
    })
}
fn parse_properties(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    element: &BytesStart<'_>,
    total: &mut usize,
) -> Result<StyleProperties> {
    let mut a = attributes(reader, version, element, total)?;
    let p = StyleProperties {
        height: take(&mut a, SVG, b"height")
            .map(|v| Length::nonnegative(v, "svg:height"))
            .transpose()?,
        min_height: take(&mut a, FO, b"min-height")
            .map(|v| Length::nonnegative(v, "fo:min-height"))
            .transpose()?,
        margins: parse_edges(&mut a, FO, b"margin", Length::new)?,
        borders: parse_edges(&mut a, FO, b"border", |v| Border::parse(&v))?,
        border_line_widths: parse_edges(&mut a, STYLE, b"border-line-width", |v| {
            BorderLineWidth::parse(&v)
        })?,
        padding: parse_edges(&mut a, FO, b"padding", |v| {
            Length::nonnegative(v, "padding")
        })?,
        background_color: take(&mut a, FO, b"background-color")
            .map(|v| Color::parse(&v, true))
            .transpose()?,
        shadow: take(&mut a, STYLE, b"shadow")
            .map(|v| Shadow::parse(&v))
            .transpose()?,
        dynamic_spacing: take(&mut a, STYLE, b"dynamic-spacing")
            .map(|v| match v.as_str() {
                "true" => Ok(true),
                "false" => Ok(false),
                _ => Err(bad("style:dynamic-spacing must be true or false")),
            })
            .transpose()?,
        background_image: None,
    };
    if !a.is_empty() {
        return Err(bad(
            "unsupported or wrongly namespaced header/footer property",
        ));
    }
    Ok(p)
}
fn parse_background(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    element: &BytesStart<'_>,
    total: &mut usize,
) -> Result<crate::SectionBackgroundImage> {
    let mut a = attributes(reader, version, element, total)?;
    let image = crate::SectionBackgroundImage {
        href: take(&mut a, XLINK, b"href"),
        position: take(&mut a, STYLE, b"position"),
        filter_name: take(&mut a, STYLE, b"filter-name"),
        xlink_type: take(&mut a, XLINK, b"type"),
        show: take(&mut a, XLINK, b"show"),
        actuate: take(&mut a, XLINK, b"actuate"),
        repeat: take(&mut a, STYLE, b"repeat")
            .map(|v| match v.as_str() {
                "repeat" => Ok(crate::BackgroundRepeat::Repeat),
                "stretch" => Ok(crate::BackgroundRepeat::Stretch),
                "no-repeat" => Ok(crate::BackgroundRepeat::NoRepeat),
                _ => Err(bad("invalid style:repeat")),
            })
            .transpose()?,
        opacity_percent: take(&mut a, DRAW, b"opacity")
            .map(|v| {
                let n = v
                    .strip_suffix('%')
                    .ok_or_else(|| bad("draw:opacity requires percent"))?
                    .parse::<u8>()
                    .map_err(|_| bad("invalid draw:opacity"))?;
                if n > 100 {
                    return Err(bad("draw:opacity exceeds 100%"));
                }
                Ok(n)
            })
            .transpose()?,
    };
    if !a.is_empty() {
        return Err(bad(
            "unsupported or wrongly namespaced background-image attribute",
        ));
    }
    image.validate()?;
    Ok(image)
}
fn element(reader: &NsReader<&[u8]>, name: quick_xml::name::QName<'_>) -> (Vec<u8>, Vec<u8>) {
    let (n, l) = reader.resolver().resolve_element(name);
    (ns(n), l.as_ref().to_vec())
}

pub fn parse_page_layout_header_footer_properties(xml: &str) -> Result<Vec<Properties>> {
    if xml.len() > MAX_XML {
        return Err(bad("header/footer XML exceeds size cap"));
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Vec<u8>, Vec<u8>)> = Vec::new();
    let mut layout: Option<(usize, String)> = None;
    let mut region: Option<(usize, Region)> = None;
    let mut active: Option<(usize, StyleProperties, bool)> = None;
    let mut total = 0usize;
    let mut out = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Decl(d)) => {
                version = d
                    .xml_version()
                    .map_err(|e| bad(format!("unsupported XML version: {e}")))?
            },
            Ok(Event::Start(e)) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(bad("header/footer XML exceeds depth cap"));
                }
                let c = element(&reader, e.name());
                let parent = stack.last();
                let depth = stack.len() + 1;
                if parent.is_some_and(|p| p.0 == OFFICE && p.1 == b"automatic-styles")
                    && c.0 == STYLE
                    && c.1 == b"page-layout"
                {
                    let mut a = attributes(&reader, version, &e, &mut total)?;
                    let name = take(&mut a, STYLE, b"name")
                        .ok_or_else(|| bad("page-layout lacks style:name"))?;
                    layout = Some((depth, name));
                } else if layout.as_ref().is_some_and(|x| depth == x.0 + 1)
                    && c.0 == STYLE
                    && matches!(c.1.as_slice(), b"header-style" | b"footer-style")
                {
                    region = Some((
                        depth,
                        if c.1 == b"header-style" {
                            Region::Header
                        } else {
                            Region::Footer
                        },
                    ));
                } else if region.as_ref().is_some_and(|x| depth == x.0 + 1)
                    && c.0 == STYLE
                    && c.1 == b"header-footer-properties"
                {
                    if active.is_some() {
                        return Err(bad("duplicate header-footer-properties"));
                    }
                    active = Some((
                        depth,
                        parse_properties(&reader, version, &e, &mut total)?,
                        false,
                    ));
                } else if active.as_ref().is_some_and(|x| depth == x.0 + 1) {
                    if c.0 != STYLE || c.1 != b"background-image" || active.as_ref().unwrap().2 {
                        return Err(bad("invalid header-footer-properties child"));
                    }
                    let image = parse_background(&reader, version, &e, &mut total)?;
                    active.as_mut().unwrap().1.background_image = Some(image);
                    active.as_mut().unwrap().2 = true;
                }
                stack.push(c);
            },
            Ok(Event::Empty(e)) => {
                let c = element(&reader, e.name());
                let depth = stack.len() + 1;
                if region.as_ref().is_some_and(|x| depth == x.0 + 1)
                    && c.0 == STYLE
                    && c.1 == b"header-footer-properties"
                {
                    let properties = parse_properties(&reader, version, &e, &mut total)?;
                    push_entry(&mut out, &layout, &region, properties)?;
                } else if active.as_ref().is_some_and(|x| depth == x.0 + 1) {
                    if c.0 != STYLE || c.1 != b"background-image" || active.as_ref().unwrap().2 {
                        return Err(bad("invalid header-footer-properties child"));
                    }
                    let image = parse_background(&reader, version, &e, &mut total)?;
                    active.as_mut().unwrap().1.background_image = Some(image);
                    active.as_mut().unwrap().2 = true;
                }
            },
            Ok(Event::End(_)) => {
                let depth = stack.len();
                if active.as_ref().is_some_and(|x| x.0 == depth) {
                    let properties = active.take().unwrap().1;
                    push_entry(&mut out, &layout, &region, properties)?;
                }
                if region.as_ref().is_some_and(|x| x.0 == depth) {
                    region = None;
                }
                if layout.as_ref().is_some_and(|x| x.0 == depth) {
                    layout = None;
                }
                stack.pop();
            },
            Ok(Event::Text(t)) => {
                if active.is_some() {
                    let value = t
                        .decode()
                        .map_err(|e| bad(format!("invalid property text: {e}")))?;
                    if !value.trim().is_empty() {
                        return Err(bad("unexpected header/footer property text"));
                    }
                }
            },
            Ok(Event::CData(d)) => {
                if active.is_some() && !d.is_empty() {
                    return Err(bad("CDATA is not allowed in header/footer properties"));
                }
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are forbidden"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(e) => return Err(bad(format!("invalid header/footer XML: {e}"))),
        }
    }
    Ok(out)
}
fn push_entry(
    out: &mut Vec<Properties>,
    layout: &Option<(usize, String)>,
    region: &Option<(usize, Region)>,
    properties: StyleProperties,
) -> Result<()> {
    let name = layout
        .as_ref()
        .ok_or_else(|| bad("properties outside page layout"))?
        .1
        .clone();
    let region = region
        .ok_or_else(|| bad("properties outside header/footer style"))?
        .1;
    if out.len() >= MAX_LAYOUTS {
        return Err(bad("header/footer layout count exceeds cap"));
    }
    if out
        .iter()
        .any(|x| x.page_layout_name == name && x.region == region)
    {
        return Err(bad("duplicate header/footer properties for page layout"));
    }
    out.push(Properties {
        page_layout_name: name,
        region,
        properties,
    });
    Ok(())
}

pub(crate) fn parse_region_properties(xml: &str) -> Result<Option<StyleProperties>> {
    let wrapped = format!(
        "<office:document-styles xmlns:office=\"{}\" xmlns:style=\"{}\"><office:automatic-styles><style:page-layout style:name=\"x\">{xml}</style:page-layout></office:automatic-styles></office:document-styles>",
        String::from_utf8_lossy(OFFICE),
        String::from_utf8_lossy(STYLE)
    );
    Ok(parse_page_layout_header_footer_properties(&wrapped)?
        .into_iter()
        .next()
        .map(|x| x.properties))
}

pub(crate) fn replace_page_layout_region_properties(
    layout: &crate::PageLayout,
    region: Region,
    properties: &StyleProperties,
) -> Result<String> {
    let property = properties.to_xml_fragment()?;
    let wrapper = match region {
        Region::Header => layout.header_style_xml.as_deref(),
        Region::Footer => layout.footer_style_xml.as_deref(),
    };
    let mut replacement = layout.xml.clone();
    if let Some(wrapper) = wrapper {
        let start = replacement
            .find(wrapper)
            .ok_or_else(|| bad("header/footer wrapper not found in page layout"))?;
        let updated = replace_in_wrapper(wrapper, &property)?;
        replacement.replace_range(start..start + wrapper.len(), &updated);
    } else {
        let fragment = properties.to_region_fragment(region)?;
        let insertion = if region == Region::Header {
            layout
                .footer_style_xml
                .as_ref()
                .and_then(|footer| replacement.find(footer))
                .unwrap_or_else(|| replacement.rfind("</").unwrap_or(replacement.len()))
        } else {
            replacement.rfind("</").unwrap_or(replacement.len())
        };
        replacement.insert_str(insertion, &fragment);
    }
    Ok(replacement)
}
fn replace_in_wrapper(wrapper: &str, property: &str) -> Result<String> {
    let mut reader = NsReader::from_reader(wrapper.as_bytes());
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut start: Option<usize> = None;
    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|e| bad(format!("invalid header/footer wrapper: {e}")))?;
        let style_element = matches!(namespace,ResolveResult::Bound(n)if n.as_ref()==STYLE);
        let event = event.into_owned();
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(e) => {
                depth += 1;
                if depth == 2
                    && style_element
                    && e.local_name().as_ref() == b"header-footer-properties"
                {
                    start = Some(event_start);
                }
            },
            Event::Empty(e)
                if depth == 1
                    && style_element
                    && e.local_name().as_ref() == b"header-footer-properties" =>
            {
                let mut out = wrapper.to_string();
                out.replace_range(event_start..event_end, property);
                return Ok(out);
            },
            Event::End(_) => {
                if let Some(begin) = start
                    && depth == 2
                {
                    let mut out = wrapper.to_string();
                    out.replace_range(begin..event_end, property);
                    return Ok(out);
                }
                depth = depth.saturating_sub(1)
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear()
    }
    if wrapper.trim_end().ends_with("/>") {
        let close = wrapper.rfind("/>").unwrap();
        let open = &wrapper[1..]
            .split(|c: char| c.is_ascii_whitespace() || c == '/')
            .next()
            .ok_or_else(|| bad("invalid header/footer wrapper"))?;
        let mut out = String::new();
        out.push_str(&wrapper[..close]);
        out.push('>');
        out.push_str(property);
        out.push_str("</");
        out.push_str(open);
        out.push('>');
        Ok(out)
    } else {
        let close = wrapper
            .rfind("</")
            .ok_or_else(|| bad("unterminated header/footer wrapper"))?;
        let mut out = wrapper.to_string();
        out.insert_str(close, property);
        Ok(out)
    }
}

impl crate::OpenDocumentPackage {
    pub fn page_layout_header_footer_properties(&self) -> Result<Vec<Properties>> {
        self.styles_xml()?.map_or_else(
            || Ok(Vec::new()),
            |xml| parse_page_layout_header_footer_properties(&xml),
        )
    }
}
impl crate::FlatOpenDocument {
    pub fn page_layout_header_footer_properties(&self) -> Result<Vec<Properties>> {
        parse_page_layout_header_footer_properties(self.xml())
    }
}
