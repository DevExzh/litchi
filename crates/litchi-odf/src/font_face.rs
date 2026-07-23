//! Semantic ODF font-face declarations.
//!
//! Linked font resources are exposed as inert metadata. This module never loads
//! a URI, installs a font, or interprets embedded font data.

use std::collections::HashSet;

use crate::{FlatOpenDocument, OpenDocumentPackage};
use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::ResolveResult,
    reader::NsReader,
};

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const SVG_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
const MAX_DOCUMENT_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_FONT_FACES: usize = 4_096;
const MAX_SOURCES_PER_FACE: usize = 1_024;
const MAX_FORMATS_PER_SOURCE: usize = 64;
const MAX_VALUE_BYTES: usize = 64 * 1024;
const MAX_AGGREGATE_TEXT_BYTES: usize = 8 * 1024 * 1024;
const MAX_XML_DEPTH: usize = 128;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum NamespaceKind {
    Office,
    Style,
    Svg,
    Xlink,
    Other,
}

/// Generic CSS font family used by `style:font-family-generic`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OdfGenericFontFamily {
    Roman,
    Swiss,
    Modern,
    Decorative,
    Script,
    System,
}

impl OdfGenericFontFamily {
    fn parse(value: &str) -> Result<Self> {
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

    fn as_str(self) -> &'static str {
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
pub enum OdfFontPitch {
    Fixed,
    Variable,
}

impl OdfFontPitch {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "fixed" => Ok(Self::Fixed),
            "variable" => Ok(Self::Variable),
            _ => invalid(format!("unsupported style:font-pitch '{value}'")),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Fixed => "fixed",
            Self::Variable => "variable",
        }
    }
}

/// SVG font style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OdfFontStyle {
    Normal,
    Italic,
    Oblique,
}

impl OdfFontStyle {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "normal" => Ok(Self::Normal),
            "italic" => Ok(Self::Italic),
            "oblique" => Ok(Self::Oblique),
            _ => invalid(format!("unsupported svg:font-style '{value}'")),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Normal => "normal",
            Self::Italic => "italic",
            Self::Oblique => "oblique",
        }
    }
}

/// SVG font variant.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OdfFontVariant {
    Normal,
    SmallCaps,
}

impl OdfFontVariant {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "normal" => Ok(Self::Normal),
            "small-caps" => Ok(Self::SmallCaps),
            _ => invalid(format!("unsupported svg:font-variant '{value}'")),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Normal => "normal",
            Self::SmallCaps => "small-caps",
        }
    }
}

/// SVG font weight, including every standard numeric weight.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OdfFontWeight {
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

impl OdfFontWeight {
    fn parse(value: &str) -> Result<Self> {
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

    fn as_str(self) -> &'static str {
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
pub enum OdfFontStretch {
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

impl OdfFontStretch {
    fn parse(value: &str) -> Result<Self> {
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

    fn as_str(self) -> &'static str {
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
pub struct OdfPositiveLength(String);

impl OdfPositiveLength {
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
pub enum OdfFontMetricKind {
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

impl OdfFontMetricKind {
    fn from_local(local: &[u8]) -> Option<Self> {
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

    fn as_str(self) -> &'static str {
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

    fn order(self) -> u8 {
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
pub struct OdfFontMetric {
    pub kind: OdfFontMetricKind,
    pub value: i64,
}

/// An inert simple XLink used by font resources.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfFontFaceLink {
    pub href: String,
    pub actuate_on_request: bool,
}

/// One item inside `svg:font-face-src`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OdfFontFaceSource {
    Uri {
        link: OdfFontFaceLink,
        /// Ordered optional `svg:string` format hints.
        formats: Vec<Option<String>>,
    },
    LocalName(Option<String>),
}

/// One complete standard `style:font-face` declaration.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct OdfFontFace {
    pub name: String,
    pub font_adornments: Option<String>,
    pub generic_family: Option<OdfGenericFontFamily>,
    pub pitch: Option<OdfFontPitch>,
    pub charset: Option<String>,
    pub family: Option<String>,
    pub style: Option<OdfFontStyle>,
    pub variant: Option<OdfFontVariant>,
    pub weight: Option<OdfFontWeight>,
    pub stretch: Option<OdfFontStretch>,
    pub size: Option<OdfPositiveLength>,
    pub unicode_range: Option<String>,
    pub panose_1: Option<String>,
    pub widths: Option<String>,
    pub bounding_box: Option<String>,
    pub metrics: Vec<OdfFontMetric>,
    pub sources: Vec<OdfFontFaceSource>,
    pub definition_source: Option<OdfFontFaceLink>,
}

/// Semantic contents of one optional `office:font-face-decls` element.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct OdfFontFaceDeclarations {
    pub faces: Vec<OdfFontFace>,
}

impl OdfFontFaceDeclarations {
    pub fn face(&self, name: &str) -> Option<&OdfFontFace> {
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
                    OdfFontFaceSource::Uri { link, formats } => {
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
                    OdfFontFaceSource::LocalName(name) => {
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
        self.validate()?;
        let mut output = String::with_capacity(256 + self.faces.len() * 128);
        output.push_str("<office:font-face-decls xmlns:office=\"");
        output.push_str(std::str::from_utf8(OFFICE_NAMESPACE).expect("namespace is UTF-8"));
        output.push_str("\" xmlns:style=\"");
        output.push_str(std::str::from_utf8(STYLE_NAMESPACE).expect("namespace is UTF-8"));
        output.push_str("\" xmlns:svg=\"");
        output.push_str(std::str::from_utf8(SVG_NAMESPACE).expect("namespace is UTF-8"));
        output.push_str("\" xmlns:xlink=\"");
        output.push_str(std::str::from_utf8(XLINK_NAMESPACE).expect("namespace is UTF-8"));
        output.push_str("\">");
        for face in &self.faces {
            output.push_str("<style:font-face");
            write_attr(&mut output, "style:name", Some(&face.name));
            write_attr(
                &mut output,
                "style:font-adornments",
                face.font_adornments.as_deref(),
            );
            write_attr(
                &mut output,
                "style:font-family-generic",
                face.generic_family.map(OdfGenericFontFamily::as_str),
            );
            write_attr(
                &mut output,
                "style:font-pitch",
                face.pitch.map(OdfFontPitch::as_str),
            );
            write_attr(&mut output, "style:font-charset", face.charset.as_deref());
            write_attr(&mut output, "svg:font-family", face.family.as_deref());
            write_attr(
                &mut output,
                "svg:font-style",
                face.style.map(OdfFontStyle::as_str),
            );
            write_attr(
                &mut output,
                "svg:font-variant",
                face.variant.map(OdfFontVariant::as_str),
            );
            write_attr(
                &mut output,
                "svg:font-weight",
                face.weight.map(OdfFontWeight::as_str),
            );
            write_attr(
                &mut output,
                "svg:font-stretch",
                face.stretch.map(OdfFontStretch::as_str),
            );
            write_attr(
                &mut output,
                "svg:font-size",
                face.size.as_ref().map(OdfPositiveLength::as_str),
            );
            write_attr(
                &mut output,
                "svg:unicode-range",
                face.unicode_range.as_deref(),
            );
            let mut metrics: Vec<_> = face.metrics.iter().collect();
            metrics.sort_unstable_by_key(|metric| metric.kind.order());
            for metric in metrics {
                output.push_str(" svg:");
                output.push_str(metric.kind.as_str());
                output.push_str("=\"");
                output.push_str(&metric.value.to_string());
                output.push('"');
            }
            write_attr(&mut output, "svg:panose-1", face.panose_1.as_deref());
            write_attr(&mut output, "svg:widths", face.widths.as_deref());
            write_attr(&mut output, "svg:bbox", face.bounding_box.as_deref());

            if face.sources.is_empty() && face.definition_source.is_none() {
                output.push_str("/>");
                continue;
            }
            output.push('>');
            if !face.sources.is_empty() {
                output.push_str("<svg:font-face-src>");
                for source in &face.sources {
                    match source {
                        OdfFontFaceSource::Uri { link, formats } => {
                            output.push_str("<svg:font-face-uri");
                            write_link_attrs(&mut output, link);
                            if formats.is_empty() {
                                output.push_str("/>");
                            } else {
                                output.push('>');
                                for format in formats {
                                    output.push_str("<svg:font-face-format");
                                    write_attr(&mut output, "svg:string", format.as_deref());
                                    output.push_str("/>");
                                }
                                output.push_str("</svg:font-face-uri>");
                            }
                        },
                        OdfFontFaceSource::LocalName(name) => {
                            output.push_str("<svg:font-face-name");
                            write_attr(&mut output, "svg:name", name.as_deref());
                            output.push_str("/>");
                        },
                    }
                }
                output.push_str("</svg:font-face-src>");
            }
            if let Some(link) = &face.definition_source {
                output.push_str("<svg:definition-src");
                write_link_attrs(&mut output, link);
                output.push_str("/>");
            }
            output.push_str("</style:font-face>");
        }
        output.push_str("</office:font-face-decls>");
        Ok(output)
    }
}

/// Parse an optional direct `office:font-face-decls` child from ODF XML.
pub fn parse_font_face_declarations(xml: &str) -> Result<Option<OdfFontFaceDeclarations>> {
    if xml.len() > MAX_DOCUMENT_XML_BYTES {
        return invalid(format!(
            "ODF XML exceeds the {MAX_DOCUMENT_XML_BYTES} byte font-face limit"
        ));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut result = None;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Start(element)
                if namespace == NamespaceKind::Office
                    && element.local_name().as_ref() == b"font-face-decls" =>
            {
                if depth != 1 {
                    return invalid("office:font-face-decls must be a direct document child");
                }
                if result.is_some() {
                    return invalid("ODF XML contains duplicate office:font-face-decls");
                }
                result = Some(parse_declarations(&mut reader)?);
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Office
                    && element.local_name().as_ref() == b"font-face-decls" =>
            {
                if depth != 1 {
                    return invalid("office:font-face-decls must be a direct document child");
                }
                if result.replace(OdfFontFaceDeclarations::default()).is_some() {
                    return invalid("ODF XML contains duplicate office:font-face-decls");
                }
            },
            Event::Start(_) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("ODF XML depth overflow".to_string()))?;
                if depth > MAX_XML_DEPTH {
                    return invalid(format!("ODF XML exceeds the {MAX_XML_DEPTH} depth limit"));
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid ODF XML depth".to_string()))?;
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODF font metadata"),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok(result)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FontFaceDeclarationsPart {
    Content,
    Styles,
    Flat,
}

impl FontFaceDeclarationsPart {
    fn root_local(self) -> &'static [u8] {
        match self {
            Self::Content => b"document-content",
            Self::Styles => b"document-styles",
            Self::Flat => b"document",
        }
    }

    fn root_name(self) -> &'static str {
        match self {
            Self::Content => "office:document-content",
            Self::Styles => "office:document-styles",
            Self::Flat => "office:document",
        }
    }
}

#[derive(Clone)]
struct XmlSpan {
    start: usize,
    end: usize,
}

struct FontFaceDeclarationsLocation {
    target: Option<XmlSpan>,
    insertion: usize,
}

fn parse_font_face_declarations_in_part(
    xml: &str,
    part: FontFaceDeclarationsPart,
) -> Result<Option<OdfFontFaceDeclarations>> {
    Ok(locate_font_face_declarations(xml, part)?.0)
}

pub(crate) fn parse_content_font_face_declarations(
    xml: &str,
) -> Result<Option<OdfFontFaceDeclarations>> {
    parse_font_face_declarations_in_part(xml, FontFaceDeclarationsPart::Content)
}

pub(crate) fn parse_styles_font_face_declarations(
    xml: &str,
) -> Result<Option<OdfFontFaceDeclarations>> {
    parse_font_face_declarations_in_part(xml, FontFaceDeclarationsPart::Styles)
}

/// Insert or replace content-part font-face declarations without rewriting
/// unrelated content XML.
pub(crate) fn set_content_font_face_declarations_xml(
    xml: &str,
    declarations: &OdfFontFaceDeclarations,
) -> Result<(String, Option<OdfFontFaceDeclarations>)> {
    set_font_face_declarations_xml(xml, declarations, FontFaceDeclarationsPart::Content)
}

/// Insert or replace styles-part font-face declarations without rewriting
/// unrelated styles XML.
pub(crate) fn set_styles_font_face_declarations_xml(
    xml: &str,
    declarations: &OdfFontFaceDeclarations,
) -> Result<(String, Option<OdfFontFaceDeclarations>)> {
    set_font_face_declarations_xml(xml, declarations, FontFaceDeclarationsPart::Styles)
}

/// Remove content-part font-face declarations without rewriting unrelated
/// content XML.
pub(crate) fn remove_content_font_face_declarations_xml(
    xml: &str,
) -> Result<(String, Option<OdfFontFaceDeclarations>)> {
    remove_font_face_declarations_xml(xml, FontFaceDeclarationsPart::Content)
}

/// Remove styles-part font-face declarations without rewriting unrelated
/// styles XML.
pub(crate) fn remove_styles_font_face_declarations_xml(
    xml: &str,
) -> Result<(String, Option<OdfFontFaceDeclarations>)> {
    remove_font_face_declarations_xml(xml, FontFaceDeclarationsPart::Styles)
}

fn set_font_face_declarations_xml(
    xml: &str,
    declarations: &OdfFontFaceDeclarations,
    part: FontFaceDeclarationsPart,
) -> Result<(String, Option<OdfFontFaceDeclarations>)> {
    declarations.validate()?;
    let (old, location) = locate_font_face_declarations(xml, part)?;
    let fragment = declarations.to_xml()?;
    let updated = if let Some(target) = location.target {
        replace_span(xml, &target, &fragment)
    } else {
        insert_at(xml, location.insertion, &fragment)
    };
    parse_font_face_declarations_in_part(&updated, part)?;
    Ok((updated, old))
}

fn remove_font_face_declarations_xml(
    xml: &str,
    part: FontFaceDeclarationsPart,
) -> Result<(String, Option<OdfFontFaceDeclarations>)> {
    let (old, location) = locate_font_face_declarations(xml, part)?;
    let Some(target) = location.target else {
        return Ok((xml.to_owned(), old));
    };
    let updated = replace_span(xml, &target, "");
    parse_font_face_declarations_in_part(&updated, part)?;
    Ok((updated, old))
}

fn locate_font_face_declarations(
    xml: &str,
    part: FontFaceDeclarationsPart,
) -> Result<(
    Option<OdfFontFaceDeclarations>,
    FontFaceDeclarationsLocation,
)> {
    let declarations = parse_font_face_declarations(xml)?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::<(NamespaceKind, Vec<u8>)>::new();
    let mut root_open_end = None;
    let mut root_closed = false;
    let mut target = None;
    let mut open_target = None::<(usize, usize)>;
    let mut scripts_end = None;
    let mut open_scripts = None::<usize>;

    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&resolved);
        match event {
            Event::Start(element) => {
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                let local = element.local_name().as_ref().to_vec();
                let depth = stack
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("ODF XML depth overflow".to_string()))?;
                if depth > MAX_XML_DEPTH {
                    return invalid(format!("ODF XML exceeds the {MAX_XML_DEPTH} depth limit"));
                }
                if depth == 1 {
                    if namespace != NamespaceKind::Office || local != part.root_local() {
                        return invalid(format!(
                            "font-face declarations require a {} root",
                            part.root_name()
                        ));
                    }
                    root_open_end = Some(end);
                } else if depth == 2 {
                    if namespace == NamespaceKind::Office && local == b"font-face-decls" {
                        if target.is_some() || open_target.is_some() {
                            return invalid("ODF XML contains duplicate office:font-face-decls");
                        }
                        open_target = Some((depth, start));
                    } else if part == FontFaceDeclarationsPart::Content
                        && namespace == NamespaceKind::Office
                        && local == b"scripts"
                    {
                        open_scripts = Some(depth);
                    }
                }
                stack.push((namespace, local));
            },
            Event::Empty(element) => {
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                let local = element.local_name().as_ref().to_vec();
                let depth = stack
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("ODF XML depth overflow".to_string()))?;
                if depth == 1 {
                    return invalid(format!(
                        "font-face declarations require a non-empty {} root",
                        part.root_name()
                    ));
                }
                if depth == 2 && namespace == NamespaceKind::Office && local == b"font-face-decls" {
                    if target.is_some() || open_target.is_some() {
                        return invalid("ODF XML contains duplicate office:font-face-decls");
                    }
                    target = Some(XmlSpan { start, end });
                } else if depth == 2
                    && part == FontFaceDeclarationsPart::Content
                    && namespace == NamespaceKind::Office
                    && local == b"scripts"
                {
                    scripts_end = Some(end);
                }
            },
            Event::End(_) => {
                let end = reader.buffer_position() as usize;
                let depth = stack.len();
                if open_target.is_some_and(|(target_depth, _)| target_depth == depth) {
                    let (_, start) = open_target.take().expect("target depth was checked");
                    target = Some(XmlSpan { start, end });
                }
                if open_scripts.is_some_and(|scripts_depth| scripts_depth == depth) {
                    open_scripts = None;
                    scripts_end = Some(end);
                }
                if depth == 1 {
                    root_closed = true;
                }
                stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("invalid ODF XML font-face element depth".to_string())
                })?;
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if root_open_end.is_none() || !root_closed || !stack.is_empty() || open_target.is_some() {
        return invalid(format!(
            "unterminated {} while locating font-face declarations",
            part.root_name()
        ));
    }
    let insertion = match part {
        FontFaceDeclarationsPart::Content => scripts_end.unwrap_or_else(|| {
            root_open_end.expect("non-empty document root has an opening event")
        }),
        FontFaceDeclarationsPart::Styles | FontFaceDeclarationsPart::Flat => {
            root_open_end.expect("non-empty document root has an opening event")
        },
    };
    Ok((
        declarations,
        FontFaceDeclarationsLocation { target, insertion },
    ))
}

fn event_start(xml: &str, end: usize) -> Result<usize> {
    xml[..end]
        .rfind('<')
        .ok_or_else(|| Error::InvalidFormat("invalid ODF font-face XML event boundary".to_string()))
}

fn replace_span(xml: &str, span: &XmlSpan, replacement: &str) -> String {
    let mut output = String::with_capacity(xml.len() - (span.end - span.start) + replacement.len());
    output.push_str(&xml[..span.start]);
    output.push_str(replacement);
    output.push_str(&xml[span.end..]);
    output
}

fn insert_at(xml: &str, insertion: usize, fragment: &str) -> String {
    let mut output = String::with_capacity(xml.len() + fragment.len());
    output.push_str(&xml[..insertion]);
    output.push_str(fragment);
    output.push_str(&xml[insertion..]);
    output
}

impl OpenDocumentPackage {
    /// Return content-part font-face declarations.
    ///
    /// Font resource links are retained as inert metadata only. This method
    /// does not fetch a URI, load a font, or inspect embedded font data.
    pub fn content_font_face_declarations(&self) -> Result<Option<OdfFontFaceDeclarations>> {
        let xml = self.content_xml()?;
        parse_content_font_face_declarations(&xml)
    }

    /// Return styles-part font-face declarations.
    ///
    /// Font resource links are retained as inert metadata only. This method
    /// does not fetch a URI, load a font, or inspect embedded font data.
    pub fn styles_font_face_declarations(&self) -> Result<Option<OdfFontFaceDeclarations>> {
        self.styles_xml()?
            .map_or_else(|| Ok(None), |xml| parse_styles_font_face_declarations(&xml))
    }
}

impl FlatOpenDocument {
    /// Return the flat document's font-face declarations.
    ///
    /// Font resource links are retained as inert metadata only. This method
    /// does not fetch a URI, load a font, or inspect embedded font data.
    pub fn font_face_declarations(&self) -> Result<Option<OdfFontFaceDeclarations>> {
        parse_font_face_declarations_in_part(self.xml(), FontFaceDeclarationsPart::Flat)
    }
}

fn parse_declarations(reader: &mut NsReader<&[u8]>) -> Result<OdfFontFaceDeclarations> {
    let mut faces = Vec::new();
    let mut names = HashSet::new();
    let mut text_bytes = 0usize;
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Start(element)
                if namespace == NamespaceKind::Style
                    && element.local_name().as_ref() == b"font-face" =>
            {
                ensure_face_capacity(faces.len())?;
                let mut face = parse_face_attributes(reader, &element, &mut text_bytes)?;
                parse_face_children(reader, &mut face, &mut text_bytes)?;
                if !names.insert(face.name.clone()) {
                    return invalid(format!("duplicate style:font-face name '{}'", face.name));
                }
                faces.push(face);
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Style
                    && element.local_name().as_ref() == b"font-face" =>
            {
                ensure_face_capacity(faces.len())?;
                let face = parse_face_attributes(reader, &element, &mut text_bytes)?;
                if !names.insert(face.name.clone()) {
                    return invalid(format!("duplicate style:font-face name '{}'", face.name));
                }
                faces.push(face);
            },
            Event::End(element)
                if namespace == NamespaceKind::Office
                    && element.local_name().as_ref() == b"font-face-decls" =>
            {
                break;
            },
            Event::Text(text) => {
                require_whitespace(&text.decode().map_err(xml_error)?, "font-face-decls")?
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::DocType(_) => {
                return invalid("DOCTYPE is not allowed in font-face declarations");
            },
            Event::Eof => return invalid("unterminated office:font-face-decls"),
            _ => return invalid("unsupported child in office:font-face-decls"),
        }
        buffer.clear();
    }
    Ok(OdfFontFaceDeclarations { faces })
}

fn parse_face_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    text_bytes: &mut usize,
) -> Result<OdfFontFace> {
    let mut face = OdfFontFace::default();
    let mut name_seen = false;
    let mut metric_kinds = HashSet::new();
    let mut seen = HashSet::new();
    for (namespace, local, value) in attributes(reader, element)? {
        if !seen.insert((namespace, local.clone())) {
            return invalid("duplicate style:font-face attribute");
        }
        *text_bytes = add_text_bytes(*text_bytes, value.len())?;
        match (namespace, local.as_slice()) {
            (NamespaceKind::Style, b"name") => {
                validate_value(&value, "style:name", false)?;
                face.name = value;
                name_seen = true;
            },
            (NamespaceKind::Style, b"font-adornments") => face.font_adornments = Some(value),
            (NamespaceKind::Style, b"font-family-generic") => {
                face.generic_family = Some(OdfGenericFontFamily::parse(&value)?)
            },
            (NamespaceKind::Style, b"font-pitch") => {
                face.pitch = Some(OdfFontPitch::parse(&value)?)
            },
            (NamespaceKind::Style, b"font-charset") => {
                validate_text_encoding(&value)?;
                face.charset = Some(value);
            },
            (NamespaceKind::Svg, b"font-family") => face.family = Some(value),
            (NamespaceKind::Svg, b"font-style") => face.style = Some(OdfFontStyle::parse(&value)?),
            (NamespaceKind::Svg, b"font-variant") => {
                face.variant = Some(OdfFontVariant::parse(&value)?)
            },
            (NamespaceKind::Svg, b"font-weight") => {
                face.weight = Some(OdfFontWeight::parse(&value)?)
            },
            (NamespaceKind::Svg, b"font-stretch") => {
                face.stretch = Some(OdfFontStretch::parse(&value)?)
            },
            (NamespaceKind::Svg, b"font-size") => face.size = Some(OdfPositiveLength::new(value)?),
            (NamespaceKind::Svg, b"unicode-range") => face.unicode_range = Some(value),
            (NamespaceKind::Svg, b"panose-1") => face.panose_1 = Some(value),
            (NamespaceKind::Svg, b"widths") => face.widths = Some(value),
            (NamespaceKind::Svg, b"bbox") => face.bounding_box = Some(value),
            (NamespaceKind::Svg, local) => {
                let Some(kind) = OdfFontMetricKind::from_local(local) else {
                    return invalid("unsupported SVG style:font-face attribute");
                };
                if !metric_kinds.insert(kind) {
                    return invalid(format!("duplicate svg:{} metric", kind.as_str()));
                }
                let value = value.parse::<i64>().map_err(|_| {
                    Error::InvalidFormat(format!("invalid svg:{} integer", kind.as_str()))
                })?;
                face.metrics.push(OdfFontMetric { kind, value });
            },
            _ => return invalid("style:font-face attribute has an unsupported namespace"),
        }
    }
    if !name_seen {
        return invalid("style:font-face requires style:name");
    }
    Ok(face)
}

fn parse_face_children(
    reader: &mut NsReader<&[u8]>,
    face: &mut OdfFontFace,
    text_bytes: &mut usize,
) -> Result<()> {
    let mut source_seen = false;
    let mut definition_seen = false;
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Start(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"font-face-src" =>
            {
                if source_seen || definition_seen {
                    return invalid("svg:font-face-src is duplicate or out of order");
                }
                source_seen = true;
                face.sources = parse_sources(reader, text_bytes)?;
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"font-face-src" =>
            {
                return invalid("svg:font-face-src requires at least one source");
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"definition-src" =>
            {
                if definition_seen {
                    return invalid("duplicate svg:definition-src");
                }
                definition_seen = true;
                let link = parse_link(reader, &element)?;
                *text_bytes = add_text_bytes(*text_bytes, link.href.len())?;
                face.definition_source = Some(link);
            },
            Event::Start(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"definition-src" =>
            {
                if definition_seen {
                    return invalid("duplicate svg:definition-src");
                }
                definition_seen = true;
                let link = parse_link(reader, &element)?;
                require_empty(reader, NamespaceKind::Svg, b"definition-src")?;
                *text_bytes = add_text_bytes(*text_bytes, link.href.len())?;
                face.definition_source = Some(link);
            },
            Event::End(element)
                if namespace == NamespaceKind::Style
                    && element.local_name().as_ref() == b"font-face" =>
            {
                break;
            },
            Event::Text(text) => {
                require_whitespace(&text.decode().map_err(xml_error)?, "font-face")?
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => return invalid("unterminated style:font-face"),
            _ => return invalid("unsupported child in style:font-face"),
        }
        buffer.clear();
    }
    Ok(())
}

fn parse_sources(
    reader: &mut NsReader<&[u8]>,
    text_bytes: &mut usize,
) -> Result<Vec<OdfFontFaceSource>> {
    let mut sources = Vec::new();
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Start(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"font-face-uri" =>
            {
                ensure_source_capacity(sources.len())?;
                let link = parse_link(reader, &element)?;
                *text_bytes = add_text_bytes(*text_bytes, link.href.len())?;
                let formats = parse_formats(reader, text_bytes)?;
                sources.push(OdfFontFaceSource::Uri { link, formats });
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"font-face-uri" =>
            {
                ensure_source_capacity(sources.len())?;
                let link = parse_link(reader, &element)?;
                *text_bytes = add_text_bytes(*text_bytes, link.href.len())?;
                sources.push(OdfFontFaceSource::Uri {
                    link,
                    formats: Vec::new(),
                });
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"font-face-name" =>
            {
                ensure_source_capacity(sources.len())?;
                let name = optional_single_svg_attribute(reader, &element, b"name")?;
                if let Some(name) = &name {
                    *text_bytes = add_text_bytes(*text_bytes, name.len())?;
                }
                sources.push(OdfFontFaceSource::LocalName(name));
            },
            Event::Start(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"font-face-name" =>
            {
                ensure_source_capacity(sources.len())?;
                let name = optional_single_svg_attribute(reader, &element, b"name")?;
                require_empty(reader, NamespaceKind::Svg, b"font-face-name")?;
                if let Some(name) = &name {
                    *text_bytes = add_text_bytes(*text_bytes, name.len())?;
                }
                sources.push(OdfFontFaceSource::LocalName(name));
            },
            Event::End(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"font-face-src" =>
            {
                break;
            },
            Event::Text(text) => {
                require_whitespace(&text.decode().map_err(xml_error)?, "font-face-src")?
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => return invalid("unterminated svg:font-face-src"),
            _ => return invalid("unsupported child in svg:font-face-src"),
        }
        buffer.clear();
    }
    if sources.is_empty() {
        return invalid("svg:font-face-src requires at least one source");
    }
    Ok(sources)
}

fn parse_formats(
    reader: &mut NsReader<&[u8]>,
    text_bytes: &mut usize,
) -> Result<Vec<Option<String>>> {
    let mut formats = Vec::new();
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Empty(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"font-face-format" =>
            {
                if formats.len() >= MAX_FORMATS_PER_SOURCE {
                    return invalid(format!(
                        "font source exceeds the {MAX_FORMATS_PER_SOURCE} format limit"
                    ));
                }
                let value = optional_single_svg_attribute(reader, &element, b"string")?;
                if let Some(value) = &value {
                    *text_bytes = add_text_bytes(*text_bytes, value.len())?;
                }
                formats.push(value);
            },
            Event::Start(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"font-face-format" =>
            {
                if formats.len() >= MAX_FORMATS_PER_SOURCE {
                    return invalid(format!(
                        "font source exceeds the {MAX_FORMATS_PER_SOURCE} format limit"
                    ));
                }
                let value = optional_single_svg_attribute(reader, &element, b"string")?;
                require_empty(reader, NamespaceKind::Svg, b"font-face-format")?;
                if let Some(value) = &value {
                    *text_bytes = add_text_bytes(*text_bytes, value.len())?;
                }
                formats.push(value);
            },
            Event::End(element)
                if namespace == NamespaceKind::Svg
                    && element.local_name().as_ref() == b"font-face-uri" =>
            {
                break;
            },
            Event::Text(text) => {
                require_whitespace(&text.decode().map_err(xml_error)?, "font-face-uri")?
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => return invalid("unterminated svg:font-face-uri"),
            _ => return invalid("unsupported child in svg:font-face-uri"),
        }
        buffer.clear();
    }
    Ok(formats)
}

fn parse_link(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<OdfFontFaceLink> {
    let mut kind = None;
    let mut href = None;
    let mut actuate = None;
    for (namespace, local, value) in attributes(reader, element)? {
        if namespace != NamespaceKind::Xlink {
            return invalid("font source link attribute has an unsupported namespace");
        }
        let slot = match local.as_slice() {
            b"type" => &mut kind,
            b"href" => &mut href,
            b"actuate" => &mut actuate,
            _ => return invalid("unsupported XLink font source attribute"),
        };
        if slot.replace(value).is_some() {
            return invalid("duplicate XLink font source attribute");
        }
    }
    if kind.as_deref() != Some("simple") {
        return invalid("font source requires xlink:type='simple'");
    }
    if actuate.as_deref().is_some_and(|value| value != "onRequest") {
        return invalid("font source xlink:actuate must be 'onRequest'");
    }
    let href =
        href.ok_or_else(|| Error::InvalidFormat("font source requires xlink:href".to_string()))?;
    validate_value(&href, "xlink:href", false)?;
    Ok(OdfFontFaceLink {
        href,
        actuate_on_request: actuate.is_some(),
    })
}

fn optional_single_svg_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected: &[u8],
) -> Result<Option<String>> {
    let attributes = attributes(reader, element)?;
    if attributes.is_empty() {
        return Ok(None);
    }
    if attributes.len() != 1
        || attributes[0].0 != NamespaceKind::Svg
        || attributes[0].1.as_slice() != expected
    {
        return invalid("font source element contains an unsupported attribute");
    }
    validate_value(&attributes[0].2, "SVG font source attribute", true)?;
    Ok(Some(
        attributes.into_iter().next().expect("one attribute").2,
    ))
}

fn require_empty(
    reader: &mut NsReader<&[u8]>,
    expected_namespace: NamespaceKind,
    expected_local: &[u8],
) -> Result<()> {
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::End(element)
                if namespace == expected_namespace
                    && element.local_name().as_ref() == expected_local =>
            {
                return Ok(());
            },
            Event::Text(text) => {
                require_whitespace(&text.decode().map_err(xml_error)?, "empty font source")?
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => return invalid("unterminated empty font source element"),
            _ => return invalid("font source element must be empty"),
        }
        buffer.clear();
    }
}

fn attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Vec<(NamespaceKind, Vec<u8>, String)>> {
    let mut output = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_kind(&namespace);
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?
            .into_owned();
        output.push((namespace, local.as_ref().to_vec(), value));
    }
    Ok(output)
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(value) if value.as_ref() == OFFICE_NAMESPACE => NamespaceKind::Office,
        ResolveResult::Bound(value) if value.as_ref() == STYLE_NAMESPACE => NamespaceKind::Style,
        ResolveResult::Bound(value) if value.as_ref() == SVG_NAMESPACE => NamespaceKind::Svg,
        ResolveResult::Bound(value) if value.as_ref() == XLINK_NAMESPACE => NamespaceKind::Xlink,
        _ => NamespaceKind::Other,
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

fn validate_text_encoding(value: &str) -> Result<()> {
    validate_value(value, "style:font-charset", false)?;
    let mut bytes = value.bytes();
    if !bytes.next().is_some_and(|byte| byte.is_ascii_alphabetic())
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'_' | b'-'))
    {
        return invalid(format!("invalid style:font-charset '{value}'"));
    }
    Ok(())
}

fn validate_link(link: &OdfFontFaceLink) -> Result<()> {
    validate_value(&link.href, "xlink:href", false)
}

fn validate_value(value: &str, name: &str, allow_empty: bool) -> Result<()> {
    if !allow_empty && value.is_empty() {
        return invalid(format!("{name} must not be empty"));
    }
    if value.len() > MAX_VALUE_BYTES {
        return invalid(format!("{name} exceeds the {MAX_VALUE_BYTES} byte limit"));
    }
    Ok(())
}

fn add_text_bytes(current: usize, additional: usize) -> Result<usize> {
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

fn ensure_face_capacity(count: usize) -> Result<()> {
    if count >= MAX_FONT_FACES {
        invalid(format!(
            "font-face declarations exceed the {MAX_FONT_FACES} face limit"
        ))
    } else {
        Ok(())
    }
}

fn ensure_source_capacity(count: usize) -> Result<()> {
    if count >= MAX_SOURCES_PER_FACE {
        invalid(format!(
            "font face exceeds the {MAX_SOURCES_PER_FACE} source limit"
        ))
    } else {
        Ok(())
    }
}

fn require_whitespace(value: &str, context: &str) -> Result<()> {
    if value.trim().is_empty() {
        Ok(())
    } else {
        invalid(format!("{context} cannot contain text"))
    }
}

fn write_link_attrs(output: &mut String, link: &OdfFontFaceLink) {
    output.push_str(" xlink:type=\"simple\"");
    write_attr(output, "xlink:href", Some(&link.href));
    if link.actuate_on_request {
        output.push_str(" xlink:actuate=\"onRequest\"");
    }
}

fn write_attr(output: &mut String, name: &str, value: Option<&str>) {
    let Some(value) = value else { return };
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    escape_attribute(output, value);
    output.push('"');
}

fn escape_attribute(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '"' => output.push_str("&quot;"),
            '\r' => output.push_str("&#13;"),
            '\n' => output.push_str("&#10;"),
            '\t' => output.push_str("&#9;"),
            _ => output.push(character),
        }
    }
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::InvalidFormat(format!("invalid ODF font-face XML: {error}"))
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    const SVG: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
    const XLINK: &str = "http://www.w3.org/1999/xlink";

    fn document(body: &str) -> String {
        format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:s="{STYLE}" xmlns:v="{SVG}" xmlns:x="{XLINK}">{body}<o:body/></o:document-content>"#
        )
    }

    #[test]
    fn parses_and_round_trips_complete_font_face_metadata() {
        // ODF 1.2/1.3 style:font-face grammar; values mirror the declarations
        // emitted by LibreOffice and modeled by odfpy's style/svg constructors.
        let xml = document(
            r#"<o:font-face-decls><s:font-face s:name="Body &amp; Text" s:font-adornments="Regular" s:font-family-generic="swiss" s:font-pitch="variable" s:font-charset="UTF-8" v:font-family="'Liberation Sans'" v:font-style="italic" v:font-variant="small-caps" v:font-weight="700" v:font-stretch="semi-expanded" v:font-size="10.5pt" v:unicode-range="U+0-10FFFF" v:units-per-em="2048" v:ascent="1854" v:descent="-434" v:panose-1="2 11 6 4" v:widths="1 2" v:bbox="0 -434 2000 1854"><v:font-face-src><v:font-face-uri x:type="simple" x:href="Fonts/body.ttf" x:actuate="onRequest"><v:font-face-format v:string="truetype"/><v:font-face-format/></v:font-face-uri><v:font-face-name v:name="Liberation Sans"/></v:font-face-src><v:definition-src x:type="simple" x:href="Fonts/body.svg"/></s:font-face><s:font-face s:name="Empty"/></o:font-face-decls>"#,
        );
        let declarations = parse_font_face_declarations(&xml).unwrap().unwrap();
        assert_eq!(declarations.faces.len(), 2);
        let face = declarations.face("Body & Text").unwrap();
        assert_eq!(face.generic_family, Some(OdfGenericFontFamily::Swiss));
        assert_eq!(face.weight, Some(OdfFontWeight::Weight700));
        assert_eq!(face.size.as_ref().unwrap().as_str(), "10.5pt");
        assert_eq!(face.metrics.len(), 3);
        assert_eq!(face.sources.len(), 2);
        assert!(matches!(
            &face.sources[0],
            OdfFontFaceSource::Uri { formats, .. } if formats == &[Some("truetype".to_string()), None]
        ));

        let serialized = declarations.to_xml().unwrap();
        let reparsed = parse_font_face_declarations(&format!(
            r#"<office:document-content xmlns:office="{OFFICE}">{serialized}<office:body/></office:document-content>"#
        ))
        .unwrap()
        .unwrap();
        assert_eq!(reparsed, declarations);
    }

    #[test]
    fn rejects_malformed_font_face_grammar() {
        for body in [
            r#"<o:font-face-decls><s:font-face/></o:font-face-decls>"#,
            r#"<o:font-face-decls><s:font-face s:name="A"/><s:font-face s:name="A"/></o:font-face-decls>"#,
            r#"<o:font-face-decls><s:font-face s:name="A" v:font-weight="950"/></o:font-face-decls>"#,
            r#"<o:font-face-decls><s:font-face s:name="A" v:font-size="0pt"/></o:font-face-decls>"#,
            r#"<o:font-face-decls><s:font-face s:name="A"><v:font-face-src/></s:font-face></o:font-face-decls>"#,
            r#"<o:font-face-decls><s:font-face s:name="A"><v:font-face-src><v:font-face-uri x:type="extended" x:href="x"/></v:font-face-src></s:font-face></o:font-face-decls>"#,
            r#"<o:font-face-decls><s:font-face s:name="A">active text</s:font-face></o:font-face-decls>"#,
        ] {
            assert!(
                parse_font_face_declarations(&document(body)).is_err(),
                "{body}"
            );
        }
    }

    #[test]
    fn rejects_misplaced_or_duplicate_containers() {
        assert!(
            parse_font_face_declarations(&document(r#"<o:body><o:font-face-decls/></o:body>"#))
                .is_err()
        );
        assert!(
            parse_font_face_declarations(&document(r#"<o:font-face-decls/><o:font-face-decls/>"#))
                .is_err()
        );
    }
}
