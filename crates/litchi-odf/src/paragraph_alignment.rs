//! Typed ODF paragraph alignment properties.
//!
//! Models the alignment attribute group of `style:paragraph-properties`
//! (`fo:text-align`, `style:vertical-align`). `fo:text-align` accepts the
//! `start`, `end`, `left`, `right`, `center`, and `justify` tokens and
//! `style:vertical-align` accepts `top`, `middle`, `bottom`, `auto`, and
//! `baseline`. The sibling `style:justify-single-word` attribute is owned by
//! the line-spacing module; all other sibling-owned attributes are ignored.
//! Duplicates and malformed owned values are rejected.

use crate::{FlatOpenDocument, OpenDocumentPackage, paragraph_margin::rewrite_start_tag};
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
const STYLE_STR: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FO_STR: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_STYLES: usize = 65_536;
const MAX_VALUE: usize = 4_096;
const MAX_TOTAL: usize = 16 * 1024 * 1024;
const MAX_ATTRIBUTES: usize = 64;

/// Whether this module owns the attribute with the given expanded name.
fn owned_attribute(namespace: Ns, local: &[u8]) -> bool {
    matches!(
        (namespace, local),
        (Ns::Fo, b"text-align") | (Ns::Style, b"vertical-align")
    )
}

fn bad(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
fn name_ok(value: &str, field: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_VALUE || value.chars().any(char::is_control) {
        return Err(bad(format!("invalid {field}")));
    }
    Ok(())
}

/// The `fo:text-align` value of a paragraph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParagraphTextAlign {
    Start,
    End,
    Left,
    Right,
    Center,
    Justify,
}
impl ParagraphTextAlign {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "start" => Ok(Self::Start),
            "end" => Ok(Self::End),
            "left" => Ok(Self::Left),
            "right" => Ok(Self::Right),
            "center" => Ok(Self::Center),
            "justify" => Ok(Self::Justify),
            _ => Err(bad("invalid fo:text-align")),
        }
    }
    fn xml(self) -> &'static str {
        match self {
            Self::Start => "start",
            Self::End => "end",
            Self::Left => "left",
            Self::Right => "right",
            Self::Center => "center",
            Self::Justify => "justify",
        }
    }
}

/// The `style:vertical-align` value of a paragraph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParagraphVerticalAlign {
    Top,
    Middle,
    Bottom,
    Auto,
    Baseline,
}
impl ParagraphVerticalAlign {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "top" => Ok(Self::Top),
            "middle" => Ok(Self::Middle),
            "bottom" => Ok(Self::Bottom),
            "auto" => Ok(Self::Auto),
            "baseline" => Ok(Self::Baseline),
            _ => Err(bad("invalid style:vertical-align")),
        }
    }
    fn xml(self) -> &'static str {
        match self {
            Self::Top => "top",
            Self::Middle => "middle",
            Self::Bottom => "bottom",
            Self::Auto => "auto",
            Self::Baseline => "baseline",
        }
    }
}

/// The alignment attribute group of one `style:paragraph-properties` element.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ParagraphAlignment {
    pub text_align: Option<ParagraphTextAlign>,
    pub vertical_align: Option<ParagraphVerticalAlign>,
}
impl ParagraphAlignment {
    pub fn new() -> Self {
        Self::default()
    }
    pub fn validate(&self) -> Result<()> {
        Ok(())
    }
    /// Serialized owned attributes, each prefixed with one space.
    fn attributes_xml(&self) -> String {
        let mut xml = String::new();
        if let Some(value) = self.text_align {
            xml.push_str(&format!(r#" fo:text-align="{}""#, value.xml()));
        }
        if let Some(value) = self.vertical_align {
            xml.push_str(&format!(r#" style:vertical-align="{}""#, value.xml()));
        }
        xml
    }
    /// Emit the properties as a `style:paragraph-properties` fragment.
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml =
            format!(r#"<style:paragraph-properties xmlns:style="{STYLE_STR}" xmlns:fo="{FO_STR}""#);
        xml.push_str(&self.attributes_xml());
        xml.push_str("/>");
        Ok(xml)
    }
}

/// A named or default paragraph style and its alignment properties.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParagraphStyleAlignment {
    pub name: Option<String>,
    pub parent_style_name: Option<String>,
    pub is_default_style: bool,
    pub properties: Option<ParagraphAlignment>,
}
impl ParagraphStyleAlignment {
    pub fn named(name: impl Into<String>, properties: Option<ParagraphAlignment>) -> Result<Self> {
        let result = Self {
            name: Some(name.into()),
            parent_style_name: None,
            is_default_style: false,
            properties,
        };
        result.validate()?;
        Ok(result)
    }
    pub fn default_style(properties: Option<ParagraphAlignment>) -> Self {
        Self {
            name: None,
            parent_style_name: None,
            is_default_style: true,
            properties,
        }
    }
    pub fn validate(&self) -> Result<()> {
        match (&self.name, self.is_default_style) {
            (Some(name), false) => name_ok(name, "paragraph style name")?,
            (None, true) => {},
            _ => return Err(bad("paragraph style identity is inconsistent")),
        }
        if let Some(parent) = &self.parent_style_name {
            if self.is_default_style {
                return Err(bad("default paragraph style cannot have a parent"));
            }
            name_ok(parent, "parent paragraph style name")?;
        }
        if let Some(properties) = &self.properties {
            properties.validate()?;
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
        let mut xml = format!(r#"<style:{tag} xmlns:style="{STYLE_STR}" style:family="paragraph""#);
        if let Some(name) = &self.name {
            xml.push_str(&format!(r#" style:name="{}""#, escape_xml(name)));
        }
        if let Some(parent) = &self.parent_style_name {
            xml.push_str(&format!(
                r#" style:parent-style-name="{}""#,
                escape_xml(parent)
            ));
        }
        if let Some(properties) = &self.properties {
            xml.push('>');
            xml.push_str(&properties.to_xml_fragment()?);
            xml.push_str(&format!("</style:{tag}>"));
        } else {
            xml.push_str("/>");
        }
        Ok(xml)
    }
}

/// All paragraph styles of a styles part that carry alignment properties.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ParagraphStyleAlignmentSet {
    pub styles: Vec<ParagraphStyleAlignment>,
}
impl ParagraphStyleAlignmentSet {
    pub fn get(&self, name: &str) -> Option<&ParagraphStyleAlignment> {
        self.styles
            .iter()
            .find(|style| style.name.as_deref() == Some(name))
    }
    pub fn default_style(&self) -> Option<&ParagraphStyleAlignment> {
        self.styles.iter().find(|style| style.is_default_style)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum Ns {
    Office,
    Style,
    Fo,
    Other,
}
fn known(resolve: ResolveResult<'_>) -> Ns {
    match resolve {
        ResolveResult::Bound(value) if value.as_ref() == OFFICE => Ns::Office,
        ResolveResult::Bound(value) if value.as_ref() == STYLE => Ns::Style,
        ResolveResult::Bound(value) if value.as_ref() == FO => Ns::Fo,
        _ => Ns::Other,
    }
}
fn element(reader: &NsReader<&[u8]>, name: QName<'_>) -> (Ns, Vec<u8>) {
    let (namespace, local) = reader.resolver().resolve_element(name);
    (known(namespace), local.as_ref().to_vec())
}
fn attribute_value(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    attribute: &quick_xml::events::attributes::Attribute<'_>,
) -> Result<String> {
    attribute
        .decoded_and_normalized_value(version, reader.decoder())
        .map(|value| value.into_owned())
        .map_err(|error| bad(format!("invalid attribute value: {error}")))
}

fn style_attributes(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<Option<ParagraphStyleAlignment>> {
    let mut name = None;
    let mut parent = None;
    let mut family = None;
    let mut seen = HashSet::new();
    for attribute in start.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| bad(format!("invalid style attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let key = (known(namespace), local.as_ref().to_vec());
        if !seen.insert(key.clone()) {
            return Err(bad("duplicate paragraph style attribute"));
        }
        let value = attribute_value(reader, version, &attribute)?;
        if value.len() > MAX_VALUE {
            return Err(bad("paragraph style attribute is too large"));
        }
        if key.0 == Ns::Style {
            match local.as_ref() {
                b"name" => name = Some(value),
                b"parent-style-name" => parent = Some(value),
                b"family" => family = Some(value),
                _ => {},
            }
        }
    }
    if family.as_deref() != Some("paragraph") {
        return Ok(None);
    }
    let result = ParagraphStyleAlignment {
        name,
        parent_style_name: parent,
        is_default_style: start.local_name().as_ref() == b"default-style",
        properties: None,
    };
    result.validate()?;
    Ok(Some(result))
}

fn alignment_attributes(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<ParagraphAlignment> {
    let mut properties = ParagraphAlignment::new();
    let mut seen = HashSet::new();
    let mut count = 0;
    for attribute in start.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| bad(format!("invalid alignment attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        count += 1;
        if count > MAX_ATTRIBUTES {
            return Err(bad("too many paragraph-properties attributes"));
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = known(namespace);
        let key = (namespace, local.as_ref().to_vec());
        if !seen.insert(key) {
            return Err(bad("duplicate paragraph-properties attribute"));
        }
        let value = attribute_value(reader, version, &attribute)?;
        if value.len() > MAX_VALUE {
            return Err(bad("paragraph-properties attribute is too large"));
        }
        match (namespace, local.as_ref()) {
            (Ns::Fo, b"text-align") => {
                properties.text_align = Some(ParagraphTextAlign::parse(&value)?);
            },
            (Ns::Style, b"vertical-align") => {
                properties.vertical_align = Some(ParagraphVerticalAlign::parse(&value)?);
            },
            // Other paragraph-properties attributes are owned by sibling modules.
            _ => {},
        }
    }
    properties.validate()?;
    Ok(properties)
}

fn push_style(
    styles: &mut Vec<ParagraphStyleAlignment>,
    style: ParagraphStyleAlignment,
    total: &mut usize,
) -> Result<()> {
    if styles.len() >= MAX_STYLES {
        return Err(bad("too many paragraph styles"));
    }
    if styles
        .iter()
        .any(|old| old.name == style.name && old.is_default_style == style.is_default_style)
    {
        return Err(bad("duplicate paragraph style identity"));
    }
    *total += style.name.as_deref().map_or(0, str::len)
        + style.parent_style_name.as_deref().map_or(0, str::len);
    if *total > MAX_TOTAL {
        return Err(bad("paragraph alignment data is too large"));
    }
    styles.push(style);
    Ok(())
}

fn is_paragraph_style(current: &(Ns, Vec<u8>), parent: Option<&(Ns, Vec<u8>)>) -> bool {
    parent.is_some_and(|(namespace, local)| {
        *namespace == Ns::Office && matches!(local.as_slice(), b"styles" | b"automatic-styles")
    }) && current.0 == Ns::Style
        && matches!(current.1.as_slice(), b"style" | b"default-style")
}

struct Active {
    depth: usize,
    style: ParagraphStyleAlignment,
    seen_properties: bool,
}

/// Parse paragraph styles and their alignment properties from a styles part.
pub fn parse_paragraph_style_alignments(xml: &str) -> Result<ParagraphStyleAlignmentSet> {
    if xml.len() > MAX_XML {
        return Err(bad("styles XML is too large"));
    }
    if !xml.contains("paragraph-properties") {
        return Ok(ParagraphStyleAlignmentSet::default());
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut active: Option<Active> = None;
    let mut styles = Vec::new();
    let mut total = 0;
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(bad("styles XML nesting is too deep"));
                }
                let current = element(&reader, start.name());
                let direct = is_paragraph_style(&current, stack.last());
                stack.push(current.clone());
                let depth = stack.len();
                if direct {
                    if let Some(style) = style_attributes(&reader, version, &start)? {
                        active = Some(Active {
                            depth,
                            style,
                            seen_properties: false,
                        });
                    }
                    continue;
                }
                if let Some(state) = active.as_mut()
                    && depth == state.depth + 1
                    && current.0 == Ns::Style
                    && current.1 == b"paragraph-properties"
                {
                    if state.seen_properties {
                        return Err(bad("duplicate style:paragraph-properties"));
                    }
                    state.seen_properties = true;
                    state.style.properties = Some(alignment_attributes(&reader, version, &start)?);
                }
            },
            Ok(Event::Empty(start)) => {
                let current = element(&reader, start.name());
                let direct = is_paragraph_style(&current, stack.last());
                let depth = stack.len() + 1;
                if direct {
                    if let Some(style) = style_attributes(&reader, version, &start)? {
                        push_style(&mut styles, style, &mut total)?;
                    }
                    continue;
                }
                if let Some(state) = active.as_mut()
                    && depth == state.depth + 1
                    && current.0 == Ns::Style
                    && current.1 == b"paragraph-properties"
                {
                    if state.seen_properties {
                        return Err(bad("duplicate style:paragraph-properties"));
                    }
                    state.seen_properties = true;
                    state.style.properties = Some(alignment_attributes(&reader, version, &start)?);
                }
            },
            Ok(Event::End(_)) => {
                let depth = stack.len();
                if active.as_ref().is_some_and(|state| state.depth == depth) {
                    push_style(&mut styles, active.take().unwrap().style, &mut total)?;
                }
                stack.pop();
            },
            Ok(Event::Decl(decl)) => {
                version = decl
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?;
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
    Ok(ParagraphStyleAlignmentSet { styles })
}

#[derive(Default)]
struct Span {
    start: usize,
    end: usize,
    end_start: usize,
    qname: String,
    empty: bool,
    owned: Vec<String>,
    missing_ns: (bool, bool),
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
    let slash = raw.rfind("/>").ok_or_else(|| bad("invalid empty element"))?;
    Ok(replace_span(
        xml,
        span,
        &format!("{}>{value}</{}>", &raw[..slash], span.qname),
    ))
}

/// Qualified names of the owned attributes as they literally appear on one
/// start tag, so lossless mutation works under arbitrary prefix aliases.
fn owned_qnames(reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Vec<String>> {
    let mut owned = Vec::new();
    for attribute in start.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| bad(format!("invalid alignment attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if owned_attribute(known(namespace), local.as_ref()) {
            owned.push(String::from_utf8_lossy(attribute.key.as_ref()).into_owned());
        }
    }
    Ok(owned)
}

/// Whether the `fo:` and `style:` prefixes are not bound to their ODF
/// namespaces in the reader's current scope, so inserted attributes need
/// local namespace declarations.
fn missing_ns_decls(reader: &NsReader<&[u8]>) -> (bool, bool) {
    let unbound = |probe: &[u8], uri: &[u8]| {
        !matches!(
            reader.resolver().resolve_attribute(QName(probe)),
            (ResolveResult::Bound(namespace), _) if namespace.as_ref() == uri
        )
    };
    (unbound(b"fo:x", FO), unbound(b"style:x", STYLE))
}

/// Prepend local namespace declarations to the serialized attributes for any
/// prefix that is unbound in the target scope.
fn qualify_insert(insert: &str, missing: (bool, bool)) -> String {
    let mut qualified = String::new();
    if missing.0 && insert.contains(" fo:") {
        qualified.push_str(&format!(r#" xmlns:fo="{FO_STR}""#));
    }
    if missing.1 && insert.contains(" style:") {
        qualified.push_str(&format!(r#" xmlns:style="{STYLE_STR}""#));
    }
    qualified.push_str(insert);
    qualified
}

/// Losslessly replace, insert, or remove this module's alignment attributes on
/// one existing paragraph style's `style:paragraph-properties` element.
/// Attributes owned by sibling modules and child elements are preserved.
pub fn set_paragraph_style_alignment_xml(
    xml: &str,
    requested: &ParagraphStyleAlignment,
) -> Result<String> {
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
                let direct = is_paragraph_style(&current, stack.last());
                stack.push(current.clone());
                let depth = stack.len();
                if direct {
                    if let Some(style) = style_attributes(&reader, version, &start)?
                        && style.name == requested.name
                        && style.is_default_style == requested.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target paragraph style"));
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
                } else if depth_target.is_some_and(|target| depth == target + 1)
                    && current.0 == Ns::Style
                    && current.1 == b"paragraph-properties"
                {
                    let span = Span {
                        start: begin,
                        end,
                        qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                        owned: owned_qnames(&reader, &start)?,
                        missing_ns: missing_ns_decls(&reader),
                        ..Default::default()
                    };
                    if active.as_mut().unwrap().properties.replace(span).is_some() {
                        return Err(bad("duplicate style:paragraph-properties"));
                    }
                }
            },
            Ok(Event::Empty(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let direct = is_paragraph_style(&current, stack.last());
                let span = Span {
                    start: begin,
                    end,
                    end_start: begin,
                    qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                    empty: true,
                    ..Default::default()
                };
                if direct {
                    if let Some(style) = style_attributes(&reader, version, &start)?
                        && style.name == requested.name
                        && style.is_default_style == requested.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target paragraph style"));
                        }
                        found = Some(TargetSpans {
                            style: span,
                            ..Default::default()
                        });
                    }
                } else if depth_target.is_some_and(|target| stack.len() + 1 == target + 1)
                    && current.0 == Ns::Style
                    && current.1 == b"paragraph-properties"
                {
                    let span = Span {
                        owned: owned_qnames(&reader, &start)?,
                        missing_ns: missing_ns_decls(&reader),
                        ..span
                    };
                    if active.as_mut().unwrap().properties.replace(span).is_some() {
                        return Err(bad("duplicate style:paragraph-properties"));
                    }
                }
            },
            Ok(Event::End(_)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let depth = stack.len();
                if let Some(spans) = active.as_mut() {
                    if spans
                        .properties
                        .as_ref()
                        .is_some_and(|span| span.end_start == 0)
                        && depth_target.is_some_and(|target| depth == target + 1)
                    {
                        let span = spans.properties.as_mut().unwrap();
                        span.end_start = begin;
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
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?;
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(bad(format!("invalid styles XML: {error}"))),
        }
    }
    let spans = found.ok_or_else(|| bad("target paragraph style does not exist"))?;
    let insert = requested
        .properties
        .as_ref()
        .map(ParagraphAlignment::attributes_xml)
        .unwrap_or_default();
    if let Some(properties) = &spans.properties {
        let raw = &xml[properties.start..properties.end];
        let insert = qualify_insert(&insert, properties.missing_ns);
        let rewritten = rewrite_start_tag(raw, &properties.owned, &insert)?;
        if properties.empty {
            return Ok(replace_span(xml, properties, &rewritten));
        }
        return Ok(format!(
            "{}{}{}",
            &xml[..properties.start],
            rewritten,
            &xml[properties.end..]
        ));
    }
    if requested.properties.is_none() {
        return Ok(xml.to_owned());
    }
    let fragment = requested.properties.as_ref().unwrap().to_xml_fragment()?;
    if spans.style.empty {
        return expand_span(xml, &spans.style, &fragment);
    }
    let mut out = xml.to_owned();
    out.insert_str(spans.style.end_start, &fragment);
    Ok(out)
}

impl OpenDocumentPackage {
    pub fn paragraph_style_alignments(&self) -> Result<ParagraphStyleAlignmentSet> {
        self.styles_xml()?.map_or_else(
            || Ok(ParagraphStyleAlignmentSet::default()),
            |xml| parse_paragraph_style_alignments(&xml),
        )
    }
}
impl FlatOpenDocument {
    pub fn paragraph_style_alignments(&self) -> Result<ParagraphStyleAlignmentSet> {
        parse_paragraph_style_alignments(self.xml())
    }
}
