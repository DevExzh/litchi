//! Bounded `DrawingML` theme XML codecs.

use std::{borrow::Cow, collections::BTreeSet, ops::Range};

use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::mce::process_ooxml;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::{NsReader, Reader};

use crate::{Error, LimitKind, Result};

use super::model::{
    Color, Face, FontSet, MAX_SCRIPT_COUNT, Override, Palette, Slot, System, Theme, validate_fonts,
    validate_name, validate_palette, validate_xml_text,
};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

pub const NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub const STRICT_NAMESPACE: &str = "http://purl.oclc.org/ooxml/drawingml/main";
pub const MAX_XML_BYTES: usize = 8 * 1024 * 1024;
const MAX_NODES: usize = 100_000;
const MAX_DEPTH: usize = 128;
const MAX_NAMESPACE_DECLARATIONS: usize = 256;
const MAX_ATTRIBUTES_PER_ELEMENT: usize = 256;
const MAX_ATTRIBUTE_VALUE_BYTES: usize = 1 << 20;
const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";

/// Caller-selected bounds for a raw DrawingML theme codec operation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[must_use]
pub struct Limits {
    /// Maximum complete theme XML bytes accepted or emitted.
    pub max_xml_bytes: usize,
    /// Maximum element nodes scanned.
    pub max_nodes: usize,
    /// Maximum element nesting depth.
    pub max_depth: usize,
    /// Maximum namespace declarations on one element.
    pub max_namespace_declarations: usize,
    /// Maximum attributes on one element.
    pub max_attributes_per_element: usize,
    /// Maximum decoded attribute value bytes.
    pub max_attribute_value_bytes: usize,
}

impl Limits {
    /// Historical bounded theme policy.
    pub const DEFAULT: Self = Self {
        max_xml_bytes: MAX_XML_BYTES,
        max_nodes: MAX_NODES,
        max_depth: MAX_DEPTH,
        max_namespace_declarations: MAX_NAMESPACE_DECLARATIONS,
        max_attributes_per_element: MAX_ATTRIBUTES_PER_ELEMENT,
        max_attribute_value_bytes: MAX_ATTRIBUTE_VALUE_BYTES,
    };

    /// Construct a policy with a caller-selected complete-theme byte cap.
    #[must_use]
    pub const fn new(max_xml_bytes: usize) -> Self {
        Self {
            max_xml_bytes,
            ..Self::DEFAULT
        }
    }

    fn validate(self) -> Result<Self> {
        check_limit(LimitKind::XmlBytes, self.max_xml_bytes, MAX_XML_BYTES)?;
        check_limit(LimitKind::XmlNodes, self.max_nodes, MAX_NODES)?;
        check_limit(LimitKind::XmlDepth, self.max_depth, MAX_DEPTH)?;
        check_limit(
            LimitKind::NamespaceDeclarations,
            self.max_namespace_declarations,
            MAX_NAMESPACE_DECLARATIONS,
        )?;
        check_limit(
            LimitKind::XmlAttributes,
            self.max_attributes_per_element,
            MAX_ATTRIBUTES_PER_ELEMENT,
        )?;
        check_limit(
            LimitKind::AttributeValueBytes,
            self.max_attribute_value_bytes,
            MAX_ATTRIBUTE_VALUE_BYTES,
        )?;
        Ok(self)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

// `CT_StyleMatrix` requires at least three entries in every style list. Keep
// the authored default deliberately small while still valid in both the
// Transitional and Strict DrawingML schemas.
const FORMAT_SCHEME: &str = concat!(
    "<a:fmtScheme name=\"Office\">",
    "<a:fillStyleLst>",
    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>",
    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>",
    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>",
    "</a:fillStyleLst>",
    "<a:lnStyleLst>",
    "<a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">",
    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>",
    "<a:prstDash val=\"solid\"/><a:miter lim=\"80%\"/></a:ln>",
    "<a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">",
    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>",
    "<a:prstDash val=\"solid\"/><a:miter lim=\"80%\"/></a:ln>",
    "<a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">",
    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>",
    "<a:prstDash val=\"solid\"/><a:miter lim=\"80%\"/></a:ln>",
    "</a:lnStyleLst>",
    "<a:effectStyleLst>",
    "<a:effectStyle><a:effectLst/></a:effectStyle>",
    "<a:effectStyle><a:effectLst/></a:effectStyle>",
    "<a:effectStyle><a:effectLst/></a:effectStyle>",
    "</a:effectStyleLst>",
    "<a:bgFillStyleLst>",
    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>",
    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>",
    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>",
    "</a:bgFillStyleLst>",
    "</a:fmtScheme>"
);

/// Encode a complete theme part.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn encode_part(
    name: &str,
    colors: &Palette,
    fonts: &FontSet,
    limits: Limits,
) -> Result<Vec<u8>> {
    let limits = limits.validate()?;
    validate_name("theme", name)?;
    validate_palette(colors)?;
    validate_fonts(fonts)?;
    preflight_part(name, colors, fonts, limits)?;
    let output_len = serialized_part_len(name, colors, fonts)?;
    let mut xml = output_string(output_len, limits)?;
    write_part(&mut xml, name, colors, fonts);
    Ok(xml.into_bytes())
}

fn write_part<T: XmlSink>(xml: &mut T, name: &str, colors: &Palette, fonts: &FontSet) {
    xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    xml.push_str("<a:theme xmlns:a=\"");
    xml.push_str(NAMESPACE);
    xml.push_str("\" name=\"");
    escape(xml, name);
    xml.push_str("\"><a:themeElements>");
    push_palette(xml, colors, false);
    push_fonts(xml, fonts, false);
    xml.push_str(FORMAT_SCHEME);
    xml.push_str("</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>");
}

/// Encode a theme override part.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn encode_override(value: &Override, limits: Limits) -> Result<Vec<u8>> {
    let limits = limits.validate()?;
    if value.colors.is_none() && value.fonts.is_none() {
        return Err(invalid("theme override requires at least one scheme"));
    }
    if let Some(colors) = &value.colors {
        validate_palette(colors)?;
    }
    if let Some(fonts) = &value.fonts {
        validate_fonts(fonts)?;
    }
    preflight_override(value, limits)?;
    let output_len = serialized_override_len(value)?;
    let mut xml = output_string(output_len, limits)?;
    write_override(&mut xml, value);
    Ok(xml.into_bytes())
}

fn write_override<T: XmlSink>(xml: &mut T, value: &Override) {
    xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    xml.push_str("<a:themeOverride xmlns:a=\"");
    xml.push_str(NAMESPACE);
    xml.push_str("\">");
    if let Some(colors) = &value.colors {
        push_palette(xml, colors, false);
    }
    if let Some(fonts) = &value.fonts {
        push_fonts(xml, fonts, false);
    }
    xml.push_str("</a:themeOverride>");
}

/// Parse a complete theme part.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn read(xml: &[u8], limits: Limits) -> Result<Theme> {
    let parsed = parse(xml, "theme", limits)?;
    let colors = parsed
        .colors
        .ok_or_else(|| invalid("theme has no color palette"))?;
    let fonts = parsed
        .fonts
        .ok_or_else(|| invalid("theme has no font set"))?;
    Ok(Theme {
        name: parsed.name.unwrap_or_default(),
        colors,
        fonts,
    })
}

/// Parse a theme override, retaining only its typed color and font schemes.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn read_override(xml: &[u8], limits: Limits) -> Result<Override> {
    let parsed = parse(xml, "themeOverride", limits)?;
    Ok(Override {
        colors: parsed.colors,
        fonts: parsed.fonts,
    })
}

/// Replace a direct `clrScheme` or `fontScheme` child of `themeElements`.
///
/// Replacement is refused when the source scheme contains namespace-qualified,
/// unknown, or otherwise unmodeled attributes or children. This keeps a typed
/// replacement from silently dropping source data that the shared model cannot
/// represent.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn replace_scheme(
    xml: &[u8],
    local: &[u8],
    replacement: &[u8],
    limits: Limits,
) -> Result<Vec<u8>> {
    let limits = limits.validate()?;
    let range = scheme_replacement_range(xml, local, limits)?;
    let source_namespace = theme_root_namespace(xml, limits)?;
    let replacement_stats =
        validate_replacement_fragment(replacement, local, source_namespace, limits)?;
    let source_stats = source_stats(xml, &range, limits)?;
    let result_nodes = source_stats
        .nodes
        .checked_sub(source_stats.target_nodes)
        .and_then(|nodes| nodes.checked_add(replacement_stats.nodes))
        .ok_or_else(|| invalid("patched theme XML node count overflows"))?;
    if result_nodes > limits.max_nodes {
        return Err(limit(LimitKind::XmlNodes, result_nodes, limits.max_nodes));
    }
    let result_depth = source_stats.outside_max_depth.max(
        source_stats
            .target_parent_depth
            .saturating_add(replacement_stats.max_depth),
    );
    if result_depth > limits.max_depth {
        return Err(limit(LimitKind::XmlDepth, result_depth, limits.max_depth));
    }
    let adjusted_len = adjusted_replacement_len(replacement, source_namespace)?;
    let output_len = range
        .start
        .checked_add(adjusted_len)
        .and_then(|length| {
            xml.len()
                .checked_sub(range.end)
                .and_then(|suffix| length.checked_add(suffix))
        })
        .ok_or_else(|| invalid("patched theme XML length overflows"))?;
    if output_len > limits.max_xml_bytes {
        return Err(limit(LimitKind::XmlBytes, output_len, limits.max_xml_bytes));
    }
    let adjusted_replacement = adjust_replacement_namespace(replacement, source_namespace)?;
    let replacement = adjusted_replacement.as_deref().unwrap_or(replacement);
    debug_assert_eq!(replacement.len(), adjusted_len);
    let mut output = Vec::with_capacity(output_len);
    output.extend_from_slice(&xml[..range.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&xml[range.end..]);
    bounded(output, limits)
}

/// Find one directly owned, losslessly replaceable scheme in a complete Theme
/// part without allocating a patched copy.
///
/// The returned byte range covers the complete source element. Namespace,
/// parent, duplicate-target, depth, size, and unsupported-content checks are
/// identical to [`replace_scheme`], so callers can stage several ranges before
/// publishing one checked output.
///
/// # Errors
///
/// Returns an error when the source is malformed, the target is absent or
/// ambiguous, or replacing it would discard source data outside the typed
/// theme model.
pub fn scheme_replacement_range(xml: &[u8], local: &[u8], limits: Limits) -> Result<Range<usize>> {
    direct_scheme_range(xml, local, limits)?
        .ok_or_else(|| invalid("theme scheme is missing from themeElements"))
}

/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn encode_palette_fragment(value: &Palette, limits: Limits) -> Result<Vec<u8>> {
    let limits = limits.validate()?;
    validate_palette(value)?;
    preflight_palette(value, true, limits)?;
    let output_len = serialized_palette_len(value, true)?;
    let mut output = output_string(output_len, limits)?;
    push_palette(&mut output, value, true);
    Ok(output.into_bytes())
}

/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn encode_fonts_fragment(value: &FontSet, limits: Limits) -> Result<Vec<u8>> {
    let limits = limits.validate()?;
    validate_fonts(value)?;
    preflight_fonts(value, true, limits)?;
    let output_len = serialized_fonts_len(value, true)?;
    let mut output = output_string(output_len, limits)?;
    push_fonts(&mut output, value, true);
    Ok(output.into_bytes())
}

fn push_palette<T: XmlSink>(xml: &mut T, value: &Palette, declare_namespace: bool) {
    xml.push_str("<a:clrScheme");
    if declare_namespace {
        xml.push_str(" xmlns:a=\"");
        xml.push_str(NAMESPACE);
        xml.push('"');
    }
    xml.push_str(" name=\"");
    escape(xml, value.name());
    xml.push_str("\">");
    for slot in Slot::ALL {
        xml.push_str("<a:");
        xml.push_str(slot.token());
        xml.push('>');
        let Some(color) = value.color(slot) else {
            continue;
        };
        match color {
            Color::Rgb(value) => {
                xml.push_str("<a:srgbClr val=\"");
                xml.push_str(value);
                xml.push_str("\"/>");
            },
            Color::System { kind, last } => {
                xml.push_str("<a:sysClr val=\"");
                xml.push_str(kind.token());
                if let Some(last) = last {
                    xml.push_str("\" lastClr=\"");
                    xml.push_str(last);
                }
                xml.push_str("\"/>");
            },
        }
        xml.push_str("</a:");
        xml.push_str(slot.token());
        xml.push('>');
    }
    xml.push_str("</a:clrScheme>");
}

fn push_fonts<T: XmlSink>(xml: &mut T, value: &FontSet, declare_namespace: bool) {
    xml.push_str("<a:fontScheme");
    if declare_namespace {
        xml.push_str(" xmlns:a=\"");
        xml.push_str(NAMESPACE);
        xml.push('"');
    }
    xml.push_str(" name=\"");
    escape(xml, value.name());
    xml.push_str("\">");
    push_face(xml, "majorFont", value.major());
    push_face(xml, "minorFont", value.minor());
    xml.push_str("</a:fontScheme>");
}

fn push_face<T: XmlSink>(xml: &mut T, name: &str, face: &Face) {
    xml.push_str("<a:");
    xml.push_str(name);
    xml.push_str("><a:latin typeface=\"");
    escape(xml, &face.latin);
    xml.push_str("\"/><a:ea typeface=\"");
    escape(xml, &face.east_asian);
    xml.push_str("\"/><a:cs typeface=\"");
    escape(xml, &face.complex_script);
    xml.push_str("\"/>");
    for script in &face.scripts {
        xml.push_str("<a:font script=\"");
        escape(xml, &script.code);
        xml.push_str("\" typeface=\"");
        escape(xml, &script.typeface);
        xml.push_str("\"/>");
    }
    xml.push_str("</a:");
    xml.push_str(name);
    xml.push('>');
}

#[derive(Default)]
struct LengthSink {
    length: u128,
}

trait XmlSink {
    fn push_str(&mut self, value: &str);
    fn push(&mut self, value: char);
}

impl XmlSink for String {
    fn push_str(&mut self, value: &str) {
        String::push_str(self, value);
    }

    fn push(&mut self, value: char) {
        String::push(self, value);
    }
}

impl XmlSink for LengthSink {
    fn push_str(&mut self, value: &str) {
        self.length += value.len() as u128;
    }

    fn push(&mut self, value: char) {
        self.length += value.len_utf8() as u128;
    }
}

/// Admission facts for authored theme output.  The writer computes these
/// counts from the actual model values before reserving its output string, so
/// narrow node/depth/attribute policies cannot be bypassed by a large
/// `with_capacity` followed by a late byte check.
#[derive(Clone, Copy)]
struct OutputBudget {
    limits: Limits,
    nodes: usize,
    namespaces: usize,
}

impl OutputBudget {
    fn new(limits: Limits) -> Self {
        Self {
            limits,
            nodes: 0,
            namespaces: 0,
        }
    }

    fn element(&mut self, depth: usize, attributes: usize) -> Result<()> {
        self.nodes = self
            .nodes
            .checked_add(1)
            .ok_or_else(|| invalid("theme authored node count overflows"))?;
        if self.nodes > self.limits.max_nodes {
            return Err(limit(
                LimitKind::XmlNodes,
                self.nodes,
                self.limits.max_nodes,
            ));
        }
        if depth > self.limits.max_depth {
            return Err(limit(LimitKind::XmlDepth, depth, self.limits.max_depth));
        }
        if attributes > self.limits.max_attributes_per_element {
            return Err(limit(
                LimitKind::XmlAttributes,
                attributes,
                self.limits.max_attributes_per_element,
            ));
        }
        Ok(())
    }

    fn namespace(&mut self, value: &str) -> Result<()> {
        self.namespaces = self
            .namespaces
            .checked_add(1)
            .ok_or_else(|| invalid("theme namespace declaration count overflows"))?;
        if self.namespaces > self.limits.max_namespace_declarations {
            return Err(limit(
                LimitKind::NamespaceDeclarations,
                self.namespaces,
                self.limits.max_namespace_declarations,
            ));
        }
        self.attribute(value)
    }

    fn attribute(&self, value: &str) -> Result<()> {
        if value.len() > self.limits.max_attribute_value_bytes {
            return Err(limit(
                LimitKind::AttributeValueBytes,
                value.len(),
                self.limits.max_attribute_value_bytes,
            ));
        }
        Ok(())
    }
}

fn preflight_part(name: &str, colors: &Palette, fonts: &FontSet, limits: Limits) -> Result<()> {
    let mut budget = OutputBudget::new(limits);
    budget.element(1, 2)?;
    budget.namespace(NAMESPACE)?;
    budget.attribute(name)?;
    budget.element(2, 0)?;
    preflight_palette_with_depth(colors, false, 3, &mut budget)?;
    preflight_fonts_with_depth(fonts, false, 3, &mut budget)?;
    preflight_fixed_scheme(2, &mut budget)?;
    budget.element(3, 0)?;
    budget.element(3, 0)?;
    Ok(())
}

fn preflight_override(value: &Override, limits: Limits) -> Result<()> {
    let mut budget = OutputBudget::new(limits);
    budget.element(1, 1)?;
    budget.namespace(NAMESPACE)?;
    if let Some(colors) = &value.colors {
        preflight_palette_with_depth(colors, false, 2, &mut budget)?;
    }
    if let Some(fonts) = &value.fonts {
        preflight_fonts_with_depth(fonts, false, 2, &mut budget)?;
    }
    Ok(())
}

fn preflight_palette(value: &Palette, declare_namespace: bool, limits: Limits) -> Result<()> {
    let mut budget = OutputBudget::new(limits);
    preflight_palette_with_depth(value, declare_namespace, 1, &mut budget)
}

fn preflight_palette_with_depth(
    value: &Palette,
    declare_namespace: bool,
    depth: usize,
    budget: &mut OutputBudget,
) -> Result<()> {
    budget.element(depth, usize::from(declare_namespace) + 1)?;
    if declare_namespace {
        budget.namespace(NAMESPACE)?;
    }
    budget.attribute(value.name())?;
    for slot in Slot::ALL {
        budget.element(depth + 1, 0)?;
        let Some(color) = value.color(slot) else {
            continue;
        };
        match color {
            Color::Rgb(value) => {
                budget.element(depth + 2, 1)?;
                budget.attribute(value)?;
            },
            Color::System { kind, last } => {
                budget.element(depth + 2, usize::from(last.is_some()) + 1)?;
                budget.attribute(kind.token())?;
                if let Some(last) = last {
                    budget.attribute(last)?;
                }
            },
        }
    }
    Ok(())
}

fn preflight_fonts(value: &FontSet, declare_namespace: bool, limits: Limits) -> Result<()> {
    let mut budget = OutputBudget::new(limits);
    preflight_fonts_with_depth(value, declare_namespace, 1, &mut budget)
}

fn preflight_fonts_with_depth(
    value: &FontSet,
    declare_namespace: bool,
    depth: usize,
    budget: &mut OutputBudget,
) -> Result<()> {
    budget.element(depth, usize::from(declare_namespace) + 1)?;
    if declare_namespace {
        budget.namespace(NAMESPACE)?;
    }
    budget.attribute(value.name())?;
    preflight_face(value.major(), depth + 1, budget)?;
    preflight_face(value.minor(), depth + 1, budget)?;
    Ok(())
}

fn preflight_face(face: &Face, depth: usize, budget: &mut OutputBudget) -> Result<()> {
    budget.element(depth, 0)?;
    budget.element(depth + 1, 1)?;
    budget.attribute(&face.latin)?;
    budget.element(depth + 1, 1)?;
    budget.attribute(&face.east_asian)?;
    budget.element(depth + 1, 1)?;
    budget.attribute(&face.complex_script)?;
    for script in &face.scripts {
        budget.element(depth + 1, 2)?;
        budget.attribute(&script.code)?;
        budget.attribute(&script.typeface)?;
    }
    Ok(())
}

fn preflight_fixed_scheme(base_depth: usize, budget: &mut OutputBudget) -> Result<()> {
    let mut reader = Reader::from_str(FORMAT_SCHEME);
    reader.config_mut().trim_text(false);
    let mut stack = 0usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                let attributes = preflight_fixed_attributes(&element, reader.decoder(), budget)?;
                budget.element(base_depth + stack + 1, attributes)?;
                stack += 1;
            },
            Event::Empty(element) => {
                let attributes = preflight_fixed_attributes(&element, reader.decoder(), budget)?;
                budget.element(base_depth + stack + 1, attributes)?;
            },
            Event::End(_) => {
                stack = stack
                    .checked_sub(1)
                    .ok_or_else(|| invalid("theme fixed scheme stack underflow"))?;
            },
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::Eof => break,
            _ => return Err(invalid("theme fixed scheme contains unsupported markup")),
        }
    }
    if stack != 0 {
        return Err(invalid("theme fixed scheme is unterminated"));
    }
    Ok(())
}

fn preflight_fixed_attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    budget: &mut OutputBudget,
) -> Result<usize> {
    let mut count = 0usize;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("theme fixed attribute count overflows"))?;
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        budget.attribute(&value)?;
    }
    Ok(count)
}

fn serialized_part_len(name: &str, colors: &Palette, fonts: &FontSet) -> Result<usize> {
    let mut sink = LengthSink::default();
    write_part(&mut sink, name, colors, fonts);
    sink.length
        .try_into()
        .map_err(|_| invalid("theme output length overflows usize"))
}

fn serialized_override_len(value: &Override) -> Result<usize> {
    let mut sink = LengthSink::default();
    write_override(&mut sink, value);
    sink.length
        .try_into()
        .map_err(|_| invalid("theme override output length overflows usize"))
}

fn serialized_palette_len(value: &Palette, declare_namespace: bool) -> Result<usize> {
    let mut sink = LengthSink::default();
    push_palette(&mut sink, value, declare_namespace);
    sink.length
        .try_into()
        .map_err(|_| invalid("theme palette output length overflows usize"))
}

fn serialized_fonts_len(value: &FontSet, declare_namespace: bool) -> Result<usize> {
    let mut sink = LengthSink::default();
    push_fonts(&mut sink, value, declare_namespace);
    sink.length
        .try_into()
        .map_err(|_| invalid("theme fonts output length overflows usize"))
}

fn output_string(length: usize, limits: Limits) -> Result<String> {
    if length > limits.max_xml_bytes {
        return Err(limit(LimitKind::XmlBytes, length, limits.max_xml_bytes));
    }
    let mut output = String::new();
    output
        .try_reserve_exact(length)
        .map_err(|_| invalid("theme output allocation failed"))?;
    Ok(output)
}

#[derive(Default)]
struct Parsed {
    name: Option<String>,
    colors: Option<Palette>,
    fonts: Option<FontSet>,
}

struct Frame {
    local: Vec<u8>,
    drawingml: bool,
}

struct FontState {
    name: String,
    major: Option<Face>,
    minor: Option<Face>,
    target: Option<bool>,
}

fn is_drawingml_namespace(namespace: &ResolveResult<'_>) -> bool {
    drawingml_namespace(namespace).is_some()
}

fn drawingml_namespace(namespace: &ResolveResult<'_>) -> Option<&'static [u8]> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == NAMESPACE.as_bytes() => {
            Some(NAMESPACE.as_bytes())
        },
        ResolveResult::Bound(Namespace(value)) if *value == STRICT_NAMESPACE.as_bytes() => {
            Some(STRICT_NAMESPACE.as_bytes())
        },
        _ => None,
    }
}

fn reject_unknown_namespace(namespace: &ResolveResult<'_>) -> Result<()> {
    if matches!(namespace, ResolveResult::Unknown(_)) {
        return Err(invalid("theme XML uses an undeclared namespace prefix"));
    }
    Ok(())
}

fn is_parent(stack: &[Frame], local: &[u8]) -> bool {
    stack
        .last()
        .is_some_and(|frame| frame.drawingml && frame.local.as_slice() == local)
}

fn is_color_slot_parent(stack: &[Frame]) -> bool {
    stack.last().is_some_and(|frame| {
        frame.drawingml
            && std::str::from_utf8(&frame.local)
                .ok()
                .and_then(Slot::from_token)
                .is_some()
    })
}

fn is_font_face_parent(stack: &[Frame]) -> bool {
    stack.last().is_some_and(|frame| {
        frame.drawingml && matches!(frame.local.as_slice(), b"majorFont" | b"minorFont")
    })
}

fn parse(xml: &[u8], root_name: &str, limits: Limits) -> Result<Parsed> {
    let limits = limits.validate()?;
    if xml.len() > limits.max_xml_bytes {
        return Err(limit(LimitKind::XmlBytes, xml.len(), limits.max_xml_bytes));
    }
    let processed = process_ooxml(xml)?;
    if processed.len() > limits.max_xml_bytes {
        return Err(limit(
            LimitKind::XmlBytes,
            processed.len(),
            limits.max_xml_bytes,
        ));
    }
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().check_end_names = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(limits.max_namespace_declarations);
    let mut stack: Vec<Frame> = Vec::new();
    let mut parsed = Parsed::default();
    let mut palette_name = String::new();
    let mut palette_values: Vec<(Slot, Color)> = Vec::new();
    let mut current_slot = None;
    let mut fonts: Option<FontState> = None;
    let mut nodes = 0usize;
    let scheme_depth = usize::from(root_name == "theme") + 1;
    let scheme_parent = if root_name == "theme" {
        b"themeElements".as_slice()
    } else {
        root_name.as_bytes()
    };
    let mut root_seen = false;

    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        reject_unknown_namespace(&namespace)?;
        match event {
            Event::Start(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("theme node count overflow"))?;
                if nodes > limits.max_nodes {
                    return Err(limit(LimitKind::XmlNodes, nodes, limits.max_nodes));
                }
                let local = element.local_name().as_ref().to_vec();
                let depth = stack.len();
                let drawingml = is_drawingml_namespace(&namespace);
                if root_seen {
                    if stack.is_empty() {
                        return Err(invalid("theme XML has multiple root elements"));
                    }
                    open(
                        &local,
                        drawingml,
                        &stack,
                        depth,
                        scheme_depth,
                        scheme_parent,
                        &element,
                        reader.decoder(),
                        &mut palette_name,
                        &mut palette_values,
                        &mut current_slot,
                        &mut fonts,
                        &mut parsed,
                        limits,
                    )?;
                } else {
                    if local != root_name.as_bytes() || !drawingml {
                        return Err(invalid("theme XML has an unexpected root"));
                    }
                    root_seen = true;
                    parsed.name = attr(&element, b"name", reader.decoder(), limits)?;
                }
                if stack.len() >= limits.max_depth {
                    return Err(limit(
                        LimitKind::XmlDepth,
                        stack.len() + 1,
                        limits.max_depth,
                    ));
                }
                stack.push(Frame { local, drawingml });
            },
            Event::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("theme node count overflow"))?;
                if nodes > limits.max_nodes {
                    return Err(limit(LimitKind::XmlNodes, nodes, limits.max_nodes));
                }
                let local = element.local_name().as_ref().to_vec();
                let depth = stack.len();
                let drawingml = is_drawingml_namespace(&namespace);
                if root_seen {
                    if stack.is_empty() {
                        return Err(invalid("theme XML has multiple root elements"));
                    }
                    open(
                        &local,
                        drawingml,
                        &stack,
                        depth,
                        scheme_depth,
                        scheme_parent,
                        &element,
                        reader.decoder(),
                        &mut palette_name,
                        &mut palette_values,
                        &mut current_slot,
                        &mut fonts,
                        &mut parsed,
                        limits,
                    )?;
                    close(
                        &local,
                        drawingml,
                        &stack,
                        depth,
                        scheme_depth,
                        scheme_parent,
                        &mut palette_name,
                        &mut palette_values,
                        &mut current_slot,
                        &mut fonts,
                        &mut parsed,
                    )?;
                } else {
                    if local != root_name.as_bytes() || !drawingml {
                        return Err(invalid("theme XML has an unexpected root"));
                    }
                    root_seen = true;
                    parsed.name = attr(&element, b"name", reader.decoder(), limits)?;
                }
            },
            Event::End(element) => {
                let local = element.local_name();
                let drawingml = is_drawingml_namespace(&namespace);
                let Some(frame) = stack.pop() else {
                    return Err(invalid("theme XML has an unexpected closing element"));
                };
                if frame.local.as_slice() != local.as_ref() || frame.drawingml != drawingml {
                    return Err(invalid("theme XML closing element does not match"));
                }
                let depth = stack.len();
                close(
                    local.as_ref(),
                    drawingml,
                    &stack,
                    depth,
                    scheme_depth,
                    scheme_parent,
                    &mut palette_name,
                    &mut palette_values,
                    &mut current_slot,
                    &mut fonts,
                    &mut parsed,
                )?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("theme XML contains forbidden markup"));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if !root_seen || !stack.is_empty() {
        return Err(invalid("theme XML is unterminated"));
    }
    Ok(parsed)
}

#[allow(
    clippy::too_many_arguments,
    reason = "one theme start event updates the complete bounded schema parser state"
)]
fn open(
    local: &[u8],
    drawingml: bool,
    stack: &[Frame],
    depth: usize,
    scheme_depth: usize,
    scheme_parent: &[u8],
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    palette_name: &mut String,
    palette_values: &mut Vec<(Slot, Color)>,
    current_slot: &mut Option<Slot>,
    fonts: &mut Option<FontState>,
    parsed: &mut Parsed,
    limits: Limits,
) -> Result<()> {
    if drawingml
        && depth == scheme_depth
        && local == b"clrScheme"
        && is_parent(stack, scheme_parent)
    {
        if parsed.colors.is_some() {
            return Err(invalid("theme contains multiple color palettes"));
        }
        palette_name.clear();
        palette_name.push_str(
            attr(element, b"name", decoder, limits)?
                .as_deref()
                .unwrap_or_default(),
        );
    } else if drawingml
        && depth == scheme_depth
        && local == b"fontScheme"
        && is_parent(stack, scheme_parent)
    {
        if fonts.is_some() || parsed.fonts.is_some() {
            return Err(invalid("theme contains multiple font sets"));
        }
        *fonts = Some(FontState {
            name: attr(element, b"name", decoder, limits)?.unwrap_or_default(),
            major: None,
            minor: None,
            target: None,
        });
    } else if drawingml
        && parsed.colors.is_none()
        && depth == scheme_depth + 1
        && is_parent(stack, b"clrScheme")
    {
        if let Some(slot) = std::str::from_utf8(local).ok().and_then(Slot::from_token)
            && current_slot.replace(slot).is_some()
        {
            return Err(invalid("theme has nested color slots"));
        }
    } else if drawingml
        && parsed.colors.is_none()
        && depth == scheme_depth + 2
        && current_slot.is_some()
        && is_color_slot_parent(stack)
    {
        let color = match local {
            b"srgbClr" => Color::rgb(
                &attr(element, b"val", decoder, limits)?
                    .ok_or_else(|| invalid("srgbClr lacks val"))?,
            )?,
            b"sysClr" => {
                let kind = attr(element, b"val", decoder, limits)?
                    .and_then(|value| System::from_token(&value))
                    .ok_or_else(|| invalid("sysClr has an unknown system color"))?;
                Color::system(kind, attr(element, b"lastClr", decoder, limits)?.as_deref())?
            },
            _ => return Ok(()),
        };
        let slot = current_slot.ok_or_else(|| invalid("theme color slot is missing"))?;
        if palette_values.iter().any(|(existing, _)| *existing == slot) {
            return Err(invalid("theme contains a duplicate color slot"));
        }
        palette_values.push((slot, color));
    } else if drawingml && let Some(fonts) = fonts.as_mut() {
        if depth == scheme_depth + 1 && local == b"majorFont" && is_parent(stack, b"fontScheme") {
            if fonts.major.is_some() {
                return Err(invalid("font set has multiple major faces"));
            }
            fonts.major = Some(Face::new(""));
            fonts.target = Some(false);
        } else if depth == scheme_depth + 1
            && local == b"minorFont"
            && is_parent(stack, b"fontScheme")
        {
            if fonts.minor.is_some() {
                return Err(invalid("font set has multiple minor faces"));
            }
            fonts.minor = Some(Face::new(""));
            fonts.target = Some(true);
        } else if depth == scheme_depth + 2 && is_font_face_parent(stack) {
            let Some(target) = fonts.target else {
                return Ok(());
            };
            let face = if target {
                fonts.minor.as_mut()
            } else {
                fonts.major.as_mut()
            }
            .ok_or_else(|| invalid("font face is missing"))?;
            match local {
                b"latin" => {
                    face.latin = attr(element, b"typeface", decoder, limits)?
                        .ok_or_else(|| invalid("latin face lacks typeface"))?;
                },
                b"ea" => {
                    face.east_asian =
                        attr(element, b"typeface", decoder, limits)?.unwrap_or_default()
                },
                b"cs" => {
                    face.complex_script =
                        attr(element, b"typeface", decoder, limits)?.unwrap_or_default();
                },
                b"font" => face.scripts.push(super::model::Script {
                    code: attr(element, b"script", decoder, limits)?.unwrap_or_default(),
                    typeface: attr(element, b"typeface", decoder, limits)?
                        .ok_or_else(|| invalid("script face lacks typeface"))?,
                }),
                _ => {},
            }
        }
    }
    Ok(())
}

fn close(
    local: &[u8],
    drawingml: bool,
    stack: &[Frame],
    depth: usize,
    scheme_depth: usize,
    scheme_parent: &[u8],
    palette_name: &mut String,
    palette_values: &mut Vec<(Slot, Color)>,
    current_slot: &mut Option<Slot>,
    fonts: &mut Option<FontState>,
    parsed: &mut Parsed,
) -> Result<()> {
    if drawingml
        && local == b"clrScheme"
        && depth == scheme_depth
        && is_parent(stack, scheme_parent)
    {
        let palette = Palette::new(std::mem::take(palette_name));
        let palette = palette_values
            .drain(..)
            .fold(palette, |palette, (slot, color)| palette.with(slot, color));
        validate_palette(&palette)?;
        parsed.colors = Some(palette);
    } else if drawingml
        && local == b"fontScheme"
        && depth == scheme_depth
        && is_parent(stack, scheme_parent)
    {
        let fonts = fonts
            .take()
            .ok_or_else(|| invalid("font set parser state is missing"))?;
        let major = fonts
            .major
            .ok_or_else(|| invalid("font set lacks a major face"))?;
        let minor = fonts
            .minor
            .ok_or_else(|| invalid("font set lacks a minor face"))?;
        let value = FontSet::new(fonts.name, major, minor);
        validate_fonts(&value)?;
        parsed.fonts = Some(value);
    } else if drawingml
        && current_slot.is_some()
        && depth == scheme_depth + 1
        && is_parent(stack, b"clrScheme")
        && Slot::from_token(std::str::from_utf8(local).unwrap_or_default()).is_some()
    {
        *current_slot = None;
    } else if drawingml
        && let Some(fonts) = fonts.as_mut()
        && matches!(local, b"majorFont" | b"minorFont")
        && depth == scheme_depth + 1
        && is_parent(stack, b"fontScheme")
    {
        fonts.target = None;
    }
    Ok(())
}

fn attr(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
    limits: Limits,
) -> Result<Option<String>> {
    let mut count = 0usize;
    let mut value = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("theme attribute count overflow"))?;
        if count > limits.max_attributes_per_element {
            return Err(limit(
                LimitKind::XmlAttributes,
                count,
                limits.max_attributes_per_element,
            ));
        }
        if attribute.key.prefix().is_none() && attribute.key.local_name().as_ref() == name {
            if value.is_some() {
                return Err(invalid(format!(
                    "duplicate theme XML attribute '{}'",
                    String::from_utf8_lossy(name)
                )));
            }
            crate::preflight_attribute_value_bytes(
                attribute.value.as_ref(),
                limits.max_attribute_value_bytes,
            )?;
            value = Some(
                attribute
                    .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .into_owned(),
            );
        }
    }
    Ok(value)
}

fn theme_root_namespace(xml: &[u8], limits: Limits) -> Result<&'static [u8]> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(limits.max_namespace_declarations);
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if element.local_name().as_ref() != b"theme" {
                    return Err(invalid("theme XML has an unexpected root"));
                }
                return match namespace {
                    ResolveResult::Bound(Namespace(value)) if value == NAMESPACE.as_bytes() => {
                        Ok(NAMESPACE.as_bytes())
                    },
                    ResolveResult::Bound(Namespace(value))
                        if value == STRICT_NAMESPACE.as_bytes() =>
                    {
                        Ok(STRICT_NAMESPACE.as_bytes())
                    },
                    _ => Err(invalid("theme XML has an unsupported root namespace")),
                };
            },
            Event::DocType(_) | Event::PI(_) | Event::End(_) => {
                return Err(invalid("theme XML has forbidden markup before its root"));
            },
            Event::Eof => return Err(invalid("theme XML has no root")),
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

fn adjust_replacement_namespace(
    replacement: &[u8],
    source_namespace: &[u8],
) -> Result<Option<Vec<u8>>> {
    let Some(range) = replacement_root_namespace_range(replacement, source_namespace)? else {
        return Ok(None);
    };
    let current = &replacement[range.clone()];
    if current == source_namespace {
        return Ok(None);
    }
    if current != NAMESPACE.as_bytes() && current != STRICT_NAMESPACE.as_bytes() {
        return Err(invalid(
            "theme replacement has an unsupported DrawingML namespace",
        ));
    }
    let adjusted_len = adjusted_replacement_len(replacement, source_namespace)?;
    let mut adjusted = Vec::with_capacity(adjusted_len);
    adjusted.extend_from_slice(&replacement[..range.start]);
    adjusted.extend_from_slice(source_namespace);
    adjusted.extend_from_slice(&replacement[range.end..]);
    Ok(Some(adjusted))
}

fn adjusted_replacement_len(replacement: &[u8], source_namespace: &[u8]) -> Result<usize> {
    let Some(range) = replacement_root_namespace_range(replacement, source_namespace)? else {
        return Ok(replacement.len());
    };
    let current = &replacement[range];
    if current == source_namespace {
        return Ok(replacement.len());
    }
    if current != NAMESPACE.as_bytes() && current != STRICT_NAMESPACE.as_bytes() {
        return Err(invalid(
            "theme replacement has an unsupported DrawingML namespace",
        ));
    }
    replacement
        .len()
        .checked_sub(current.len())
        .and_then(|length| length.checked_add(source_namespace.len()))
        .ok_or_else(|| invalid("theme replacement length overflows"))
}

fn replacement_root_namespace_range(
    xml: &[u8],
    source_namespace: &[u8],
) -> Result<Option<Range<usize>>> {
    let mut reader = Reader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("theme replacement offset exceeds usize"))?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("theme replacement offset exceeds usize"))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let prefix = element.name().prefix().map(|prefix| prefix.into_inner());
                let raw = xml
                    .get(start..end)
                    .ok_or_else(|| invalid("theme replacement root range is invalid"))?;
                let qname_relative = if prefix == Some(b"a".as_slice()) {
                    optional_attribute_value_range(raw, b"xmlns:a")?
                } else if prefix.is_none() {
                    optional_attribute_value_range(raw, b"xmlns")?
                } else {
                    let prefix = prefix.expect("checked arbitrary root prefix");
                    let mut wanted = b"xmlns:".to_vec();
                    wanted.extend_from_slice(prefix);
                    optional_attribute_value_range(raw, &wanted)?
                };
                let counterpart = if source_namespace == NAMESPACE.as_bytes() {
                    STRICT_NAMESPACE.as_bytes()
                } else if source_namespace == STRICT_NAMESPACE.as_bytes() {
                    NAMESPACE.as_bytes()
                } else {
                    return Err(invalid("theme source has an unsupported namespace"));
                };
                let default_relative = prefix
                    .is_some()
                    .then(|| optional_attribute_value_range(raw, b"xmlns"));
                let default_relative = match default_relative {
                    Some(result) => result?,
                    None => None,
                };
                let default_range = default_relative
                    .as_ref()
                    .map(|relative| (start + relative.start)..(start + relative.end));
                let default_current = default_range
                    .as_ref()
                    .map(|range| {
                        xml.get(range.clone())
                            .ok_or_else(|| invalid("theme replacement namespace range is invalid"))
                    })
                    .transpose()?;

                let Some(relative) = qname_relative else {
                    if default_current == Some(counterpart) {
                        return Ok(default_range);
                    }
                    return Ok(None);
                };
                let range = (start + relative.start)..(start + relative.end);
                let current = xml
                    .get(range.clone())
                    .ok_or_else(|| invalid("theme replacement namespace range is invalid"))?;
                if current == source_namespace {
                    if default_current == Some(counterpart) {
                        return Ok(default_range);
                    }
                    return Ok(None);
                }
                if current == counterpart {
                    if default_current == Some(counterpart) {
                        return Err(invalid(
                            "theme replacement has multiple cross-conformance namespace bindings",
                        ));
                    }
                    return Ok(Some(range));
                }
                return Err(invalid(
                    "theme replacement has an unsupported DrawingML namespace",
                ));
            },
            Event::DocType(_) | Event::PI(_) | Event::End(_) => {
                return Err(invalid(
                    "theme replacement has forbidden markup before its root",
                ));
            },
            Event::Eof => return Ok(None),
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

/// Validate a replacement fragment before namespace adjustment or patched
/// output allocation. The source scheme checker validates the existing
/// modeled scheme; this pass applies the caller's generic XML/resource
/// policy to arbitrary replacement markup while retaining opaque children.
#[derive(Clone, Copy, Debug, Default)]
struct XmlStats {
    nodes: usize,
    max_depth: usize,
}

#[derive(Clone, Copy, Debug, Default)]
struct SourceStats {
    nodes: usize,
    target_nodes: usize,
    outside_max_depth: usize,
    target_parent_depth: usize,
}

#[derive(Debug)]
struct ReplacementFrame {
    local: Vec<u8>,
    opaque: bool,
    prefix: Option<Vec<u8>>,
}

#[derive(Debug, Default)]
struct ReplacementSemantics {
    palette_slots: u16,
    palette_values: [u8; Slot::ALL.len()],
    font_faces: [bool; 2],
    font_face_fields: [[bool; 3]; 2],
    font_script_counts: [usize; 2],
}

fn replacement_slot_index(local: &[u8]) -> Option<usize> {
    let token = std::str::from_utf8(local).ok()?;
    let slot = Slot::from_token(token)?;
    Slot::ALL.iter().position(|candidate| *candidate == slot)
}

fn replacement_semantic_open(
    semantics: &mut ReplacementSemantics,
    target: &[u8],
    stack: &[ReplacementFrame],
    local: &[u8],
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    limits: Limits,
) -> Result<()> {
    if stack.is_empty() {
        let name = attr(element, b"name", decoder, limits)?
            .ok_or_else(|| invalid("theme replacement scheme lacks name"))?;
        let kind = if target == b"clrScheme" {
            "color palette"
        } else {
            "font set"
        };
        validate_name(kind, &name)?;
        return Ok(());
    }
    let parent = stack.last().expect("checked non-empty replacement stack");
    if parent.opaque {
        return Ok(());
    }
    if target == b"clrScheme" {
        if parent.local.as_slice() == b"clrScheme" {
            let Some(index) = replacement_slot_index(local) else {
                return Ok(());
            };
            let bit = 1_u16 << index;
            if semantics.palette_slots & bit != 0 {
                return Err(invalid("color scheme replacement has a duplicate slot"));
            }
            semantics.palette_slots |= bit;
            return Ok(());
        }
        let Some(index) = replacement_slot_index(&parent.local) else {
            return Ok(());
        };
        if semantics.palette_values[index] != 0 {
            return Err(invalid("color scheme replacement slot has multiple values"));
        }
        match local {
            b"srgbClr" => {
                let value = attr(element, b"val", decoder, limits)?
                    .ok_or_else(|| invalid("replacement srgbClr lacks val"))?;
                validate_replacement_hex(&value)?;
            },
            b"sysClr" => {
                let kind = attr(element, b"val", decoder, limits)?
                    .and_then(|value| System::from_token(&value))
                    .ok_or_else(|| invalid("replacement sysClr has an unknown val"))?;
                let last = attr(element, b"lastClr", decoder, limits)?;
                if let Some(last) = last {
                    validate_replacement_hex(&last)?;
                }
                let _ = kind;
            },
            _ => return Ok(()),
        }
        semantics.palette_values[index] = 1;
    } else if target == b"fontScheme" {
        if parent.local.as_slice() == b"fontScheme" {
            let index = match local {
                b"majorFont" => 0,
                b"minorFont" => 1,
                _ => return Ok(()),
            };
            if semantics.font_faces[index] {
                return Err(invalid("font scheme replacement has a duplicate face"));
            }
            semantics.font_faces[index] = true;
            return Ok(());
        }
        let Some(face_index) = (match parent.local.as_slice() {
            b"majorFont" => Some(0),
            b"minorFont" => Some(1),
            _ => None,
        }) else {
            return Ok(());
        };
        let field_index = match local {
            b"latin" => 0,
            b"ea" => 1,
            b"cs" => 2,
            b"font" => {
                if semantics.font_script_counts[face_index] >= MAX_SCRIPT_COUNT {
                    return Err(invalid("font face replacement has too many script entries"));
                }
                let code = attr(element, b"script", decoder, limits)?
                    .ok_or_else(|| invalid("replacement script face lacks script"))?;
                if code.is_empty()
                    || code.chars().count() > 16
                    || !code.bytes().all(|byte| byte.is_ascii_alphanumeric())
                {
                    return Err(invalid("replacement script code is invalid"));
                }
                let typeface = attr(element, b"typeface", decoder, limits)?
                    .ok_or_else(|| invalid("replacement script face lacks typeface"))?;
                validate_name("replacement script typeface", &typeface)?;
                semantics.font_script_counts[face_index] += 1;
                return Ok(());
            },
            _ => return Ok(()),
        };
        if semantics.font_face_fields[face_index][field_index] {
            return Err(invalid("font face replacement has a duplicate field"));
        }
        semantics.font_face_fields[face_index][field_index] = true;
        let typeface = attr(element, b"typeface", decoder, limits)?;
        if field_index == 0 {
            let typeface =
                typeface.ok_or_else(|| invalid("replacement latin face lacks typeface"))?;
            validate_name("replacement latin typeface", &typeface)?;
        } else if let Some(typeface) = typeface {
            validate_xml_text("replacement typeface", &typeface)?;
            if typeface.chars().count() > 256 {
                return Err(invalid("replacement typeface exceeds 256 characters"));
            }
        }
    }
    Ok(())
}

fn validate_replacement_hex(value: &str) -> Result<()> {
    if value.len() != 6 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(invalid(
            "replacement color must contain six hexadecimal digits",
        ));
    }
    Ok(())
}

fn replacement_semantic_close(
    semantics: &ReplacementSemantics,
    target: &[u8],
    local: &[u8],
    parent: Option<&ReplacementFrame>,
) -> Result<()> {
    if target == b"clrScheme" {
        if local == b"clrScheme" && parent.is_none() {
            let expected = (1_u16 << Slot::ALL.len()) - 1;
            if semantics.palette_slots != expected
                || semantics.palette_values.iter().any(|count| *count != 1)
            {
                return Err(invalid(
                    "color scheme replacement must contain every slot exactly once",
                ));
            }
        } else if parent
            .is_some_and(|frame| !frame.opaque && frame.local.as_slice() == b"clrScheme")
            && replacement_slot_index(local).is_some()
        {
            let index = replacement_slot_index(local).expect("checked slot");
            if semantics.palette_values[index] != 1 {
                return Err(invalid(
                    "color scheme replacement slot must contain one color value",
                ));
            }
        }
    } else if target == b"fontScheme" {
        if local == b"fontScheme" && parent.is_none() {
            if semantics.font_faces != [true, true]
                || semantics.font_face_fields.iter().any(|fields| !fields[0])
            {
                return Err(invalid(
                    "font scheme replacement must contain major and minor latin faces",
                ));
            }
        } else if let Some(parent) = parent
            && !parent.opaque
            && parent.local.as_slice() == b"fontScheme"
            && matches!(local, b"majorFont" | b"minorFont")
        {
            let index = usize::from(local == b"minorFont");
            if !semantics.font_face_fields[index][0] {
                return Err(invalid("font face replacement lacks a latin face"));
            }
        }
    }
    Ok(())
}

fn validate_replacement_fragment(
    replacement: &[u8],
    target: &[u8],
    source_namespace: &[u8],
    limits: Limits,
) -> Result<XmlStats> {
    if replacement.len() > limits.max_xml_bytes {
        return Err(limit(
            LimitKind::XmlBytes,
            replacement.len(),
            limits.max_xml_bytes,
        ));
    }
    let mut reader = NsReader::from_reader(replacement);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(limits.max_namespace_declarations);
    let mut stack: Vec<ReplacementFrame> = Vec::new();
    let mut semantics = ReplacementSemantics::default();
    let mut root_seen = false;
    let mut root_closed = false;
    let mut nodes = 0usize;
    let mut max_depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let namespace = match &event {
            Event::Start(element) | Event::Empty(element) => {
                reader.resolver().resolve_element(element.name()).0
            },
            Event::End(element) => reader.resolver().resolve_element(element.name()).0,
            _ => ResolveResult::Unbound,
        };
        match event {
            Event::Start(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("theme replacement node count overflows"))?;
                if nodes > limits.max_nodes {
                    return Err(limit(LimitKind::XmlNodes, nodes, limits.max_nodes));
                }
                if root_closed || (root_seen && stack.is_empty()) {
                    return Err(invalid("theme replacement has more than one root"));
                }
                validate_replacement_namespace(
                    &namespace,
                    &element,
                    !root_seen,
                    target,
                    source_namespace,
                    stack
                        .first()
                        .is_some_and(|frame| frame.prefix.as_deref() == Some(b"a".as_slice()))
                        || (!root_seen
                            && element
                                .name()
                                .prefix()
                                .is_some_and(|prefix| prefix.as_ref() == b"a")),
                    stack.last().is_some_and(|frame| frame.opaque),
                )?;
                validate_replacement_attributes(&element, &reader, limits)?;
                validate_modeled_namespace_binding(
                    &element,
                    !root_seen,
                    target,
                    &namespace,
                    source_namespace,
                    reader.decoder(),
                    stack.last().is_some_and(|frame| frame.opaque),
                )?;
                let local = element.local_name().as_ref().to_vec();
                if !root_seen && local.as_slice() != target {
                    return Err(invalid("theme replacement root does not match target"));
                }
                let parent_opaque = stack.last().is_some_and(|frame| frame.opaque);
                let opaque = parent_opaque
                    || !replacement_element_is_drawingml(
                        &namespace,
                        &element,
                        source_namespace,
                        stack
                            .first()
                            .is_some_and(|frame| frame.prefix.as_deref() == Some(b"a".as_slice()))
                            || (!root_seen
                                && element
                                    .name()
                                    .prefix()
                                    .is_some_and(|prefix| prefix.as_ref() == b"a")),
                    );
                validate_replacement_scheme_element(
                    target,
                    &stack,
                    &namespace,
                    &element,
                    source_namespace,
                    opaque,
                    stack
                        .first()
                        .is_some_and(|frame| frame.prefix.as_deref() == Some(b"a".as_slice()))
                        || (!root_seen
                            && element
                                .name()
                                .prefix()
                                .is_some_and(|prefix| prefix.as_ref() == b"a")),
                    limits,
                )?;
                replacement_semantic_open(
                    &mut semantics,
                    target,
                    &stack,
                    element.local_name().as_ref(),
                    &element,
                    reader.decoder(),
                    limits,
                )?;
                if stack.len() >= limits.max_depth {
                    return Err(limit(
                        LimitKind::XmlDepth,
                        stack.len() + 1,
                        limits.max_depth,
                    ));
                }
                max_depth = max_depth.max(stack.len() + 1);
                root_seen = true;
                stack.push(ReplacementFrame {
                    local,
                    opaque,
                    prefix: element
                        .name()
                        .prefix()
                        .map(|prefix| prefix.into_inner().to_vec()),
                });
            },
            Event::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("theme replacement node count overflows"))?;
                if nodes > limits.max_nodes {
                    return Err(limit(LimitKind::XmlNodes, nodes, limits.max_nodes));
                }
                if root_closed || (root_seen && stack.is_empty()) {
                    return Err(invalid("theme replacement has more than one root"));
                }
                validate_replacement_namespace(
                    &namespace,
                    &element,
                    !root_seen,
                    target,
                    source_namespace,
                    stack
                        .first()
                        .is_some_and(|frame| frame.prefix.as_deref() == Some(b"a".as_slice()))
                        || (!root_seen
                            && element
                                .name()
                                .prefix()
                                .is_some_and(|prefix| prefix.as_ref() == b"a")),
                    stack.last().is_some_and(|frame| frame.opaque),
                )?;
                validate_replacement_attributes(&element, &reader, limits)?;
                validate_modeled_namespace_binding(
                    &element,
                    !root_seen,
                    target,
                    &namespace,
                    source_namespace,
                    reader.decoder(),
                    stack.last().is_some_and(|frame| frame.opaque),
                )?;
                if !root_seen && element.local_name().as_ref() != target {
                    return Err(invalid("theme replacement root does not match target"));
                }
                let parent_opaque = stack.last().is_some_and(|frame| frame.opaque);
                let opaque = parent_opaque
                    || !replacement_element_is_drawingml(
                        &namespace,
                        &element,
                        source_namespace,
                        stack
                            .first()
                            .is_some_and(|frame| frame.prefix.as_deref() == Some(b"a".as_slice()))
                            || (!root_seen
                                && element
                                    .name()
                                    .prefix()
                                    .is_some_and(|prefix| prefix.as_ref() == b"a")),
                    );
                validate_replacement_scheme_element(
                    target,
                    &stack,
                    &namespace,
                    &element,
                    source_namespace,
                    opaque,
                    stack
                        .first()
                        .is_some_and(|frame| frame.prefix.as_deref() == Some(b"a".as_slice()))
                        || (!root_seen
                            && element
                                .name()
                                .prefix()
                                .is_some_and(|prefix| prefix.as_ref() == b"a")),
                    limits,
                )?;
                replacement_semantic_open(
                    &mut semantics,
                    target,
                    &stack,
                    element.local_name().as_ref(),
                    &element,
                    reader.decoder(),
                    limits,
                )?;
                replacement_semantic_close(
                    &semantics,
                    target,
                    element.local_name().as_ref(),
                    stack.last(),
                )?;
                if 1 > limits.max_depth {
                    return Err(limit(LimitKind::XmlDepth, 1, limits.max_depth));
                }
                max_depth = max_depth.max(stack.len() + 1);
                if !root_seen {
                    root_seen = true;
                    root_closed = true;
                }
            },
            Event::End(element) => {
                let Some(local) = stack.pop() else {
                    return Err(invalid(
                        "theme replacement has an unexpected closing element",
                    ));
                };
                if local.local.as_slice() != element.local_name().as_ref() {
                    return Err(invalid("theme replacement closing element does not match"));
                }
                replacement_semantic_close(&semantics, target, &local.local, stack.last())?;
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                let value = std::str::from_utf8(text.as_ref()).map_err(xml_error)?;
                validate_xml_text("theme replacement text", value)?;
                if !value.chars().all(char::is_whitespace) {
                    return Err(invalid("theme replacement contains unsupported text"));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) => {
                return Err(invalid("theme replacement contains unsupported content"));
            },
            Event::Comment(comment) => {
                validate_xml_text(
                    "theme replacement comment",
                    std::str::from_utf8(comment.as_ref()).map_err(xml_error)?,
                )?;
            },
            Event::Decl(_) | Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("theme replacement contains forbidden markup"));
            },
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("theme replacement is unterminated"));
    }
    Ok(XmlStats { nodes, max_depth })
}

fn source_stats(xml: &[u8], target: &Range<usize>, limits: Limits) -> Result<SourceStats> {
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().check_end_names = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(limits.max_namespace_declarations);
    let mut stack = Vec::<Vec<u8>>::new();
    let mut stats = SourceStats::default();
    let mut root_seen = false;
    let mut root_closed = false;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("theme XML offset exceeds usize"))?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("theme XML offset exceeds usize"))?;
        let (namespace, event) = reader.resolver().resolve_event(event);
        reject_unknown_namespace(&namespace)?;
        match event {
            Event::Start(element) => {
                if root_closed || (root_seen && stack.is_empty()) {
                    return Err(invalid("theme XML has multiple root elements"));
                }
                validate_replacement_attributes(&element, &reader, limits)?;
                stats.nodes = stats
                    .nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("theme node count overflows"))?;
                if stats.nodes > limits.max_nodes {
                    return Err(limit(LimitKind::XmlNodes, stats.nodes, limits.max_nodes));
                }
                let depth = stack.len() + 1;
                if depth > limits.max_depth {
                    return Err(limit(LimitKind::XmlDepth, depth, limits.max_depth));
                }
                let in_target = start >= target.start && start < target.end;
                if in_target {
                    if stats.target_nodes == 0 {
                        stats.target_parent_depth = stack.len();
                    }
                    stats.target_nodes += 1;
                } else {
                    stats.outside_max_depth = stats.outside_max_depth.max(depth);
                }
                root_seen = true;
                stack.push(element.local_name().as_ref().to_vec());
            },
            Event::Empty(element) => {
                if root_closed || (root_seen && stack.is_empty()) {
                    return Err(invalid("theme XML has multiple root elements"));
                }
                validate_replacement_attributes(&element, &reader, limits)?;
                stats.nodes = stats
                    .nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("theme node count overflows"))?;
                if stats.nodes > limits.max_nodes {
                    return Err(limit(LimitKind::XmlNodes, stats.nodes, limits.max_nodes));
                }
                let depth = stack.len() + 1;
                if depth > limits.max_depth {
                    return Err(limit(LimitKind::XmlDepth, depth, limits.max_depth));
                }
                let in_target = start >= target.start && start < target.end;
                if in_target {
                    if stats.target_nodes == 0 {
                        stats.target_parent_depth = stack.len();
                    }
                    stats.target_nodes += 1;
                } else {
                    stats.outside_max_depth = stats.outside_max_depth.max(depth);
                }
                if !root_seen {
                    root_seen = true;
                    root_closed = true;
                }
            },
            Event::End(element) => {
                let Some(local) = stack.pop() else {
                    return Err(invalid("theme XML nesting underflow"));
                };
                if local.as_slice() != element.local_name().as_ref() {
                    return Err(invalid("theme XML closing element does not match"));
                }
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("theme XML contains forbidden markup"));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
        let _ = end;
    }
    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("theme XML is unterminated"));
    }
    if stats.target_nodes == 0 {
        return Err(invalid("theme replacement range has no source nodes"));
    }
    Ok(stats)
}

fn validate_replacement_namespace(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    root: bool,
    target: &[u8],
    source_namespace: &[u8],
    allow_unknown_a: bool,
    opaque_parent: bool,
) -> Result<()> {
    if opaque_parent && !root {
        return match namespace {
            ResolveResult::Unknown(_) => Err(invalid(
                "theme replacement uses an undeclared namespace prefix",
            )),
            _ => Ok(()),
        };
    }
    let local = element.local_name();
    let known = is_known_replacement_local(local.as_ref(), target);
    match namespace {
        ResolveResult::Bound(Namespace(value))
            if root
                && (*value == source_namespace
                    || (*value == NAMESPACE.as_bytes()
                        && source_namespace == STRICT_NAMESPACE.as_bytes())
                    || (*value == STRICT_NAMESPACE.as_bytes()
                        && source_namespace == NAMESPACE.as_bytes())) =>
        {
            Ok(())
        },
        ResolveResult::Bound(Namespace(value)) if root => {
            Err(invalid("theme replacement root has a foreign namespace"))
        },
        ResolveResult::Bound(Namespace(value))
            if known
                && (*value == source_namespace
                    || (*value == NAMESPACE.as_bytes()
                        && source_namespace == STRICT_NAMESPACE.as_bytes())
                    || (*value == STRICT_NAMESPACE.as_bytes()
                        && source_namespace == NAMESPACE.as_bytes())) =>
        {
            Ok(())
        },
        ResolveResult::Bound(Namespace(_)) if known => Err(invalid(
            "theme replacement modeled child has a cross-family namespace",
        )),
        ResolveResult::Bound(_) => Ok(()),
        ResolveResult::Unknown(prefix) if allow_unknown_a && prefix.as_slice() == b"a" => Ok(()),
        ResolveResult::Unbound if !root && element.name().prefix().is_none() => Ok(()),
        ResolveResult::Unbound if root && element.name().prefix().is_none() => {
            Err(invalid("theme replacement root has no DrawingML namespace"))
        },
        ResolveResult::Unknown(_) => Err(invalid(
            "theme replacement uses an undeclared namespace prefix",
        )),
        ResolveResult::Unbound => Err(invalid("theme replacement has an unbound namespace")),
    }
}

fn replacement_element_is_drawingml(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    source_namespace: &[u8],
    allow_unknown_a: bool,
) -> bool {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            *value == source_namespace
                || (*value == NAMESPACE.as_bytes()
                    && source_namespace == STRICT_NAMESPACE.as_bytes())
                || (*value == STRICT_NAMESPACE.as_bytes()
                    && source_namespace == NAMESPACE.as_bytes())
        },
        ResolveResult::Unknown(prefix) if allow_unknown_a && prefix.as_slice() == b"a" => true,
        ResolveResult::Unbound if element.name().prefix().is_none() => false,
        _ => false,
    }
}

fn validate_replacement_scheme_element(
    target: &[u8],
    stack: &[ReplacementFrame],
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    source_namespace: &[u8],
    opaque: bool,
    allow_unknown_a: bool,
    limits: Limits,
) -> Result<()> {
    let local = element.local_name();
    let drawingml =
        replacement_element_is_drawingml(namespace, element, source_namespace, allow_unknown_a);
    if opaque {
        if !drawingml
            && !stack.last().is_some_and(|frame| frame.opaque)
            && is_known_replacement_local(local.as_ref(), target)
        {
            return Err(invalid(
                "theme replacement modeled child has no DrawingML namespace",
            ));
        }
        return Ok(());
    }
    if !drawingml {
        return Err(invalid(
            "theme replacement modeled child has no DrawingML namespace",
        ));
    }
    let parent = stack
        .last()
        .map(|frame| frame.local.as_slice())
        .unwrap_or_default();
    let allowed: &[&[u8]] = if stack.is_empty() {
        if local.as_ref() != target {
            return Err(invalid("theme replacement root does not match target"));
        }
        &[b"name"]
    } else if target == b"clrScheme" {
        if parent == b"clrScheme" {
            if Slot::from_token(std::str::from_utf8(local.as_ref()).unwrap_or_default()).is_none() {
                return Err(invalid(
                    "color scheme contains an unsupported replacement child",
                ));
            }
            &[]
        } else if is_color_replacement_value(local.as_ref(), parent) {
            match local.as_ref() {
                b"srgbClr" => &[b"val"],
                b"sysClr" => &[b"val", b"lastClr"],
                _ => &[b"val"],
            }
        } else {
            return Err(invalid(
                "color scheme contains an unsupported replacement child",
            ));
        }
    } else if target == b"fontScheme" {
        if parent == b"fontScheme" {
            if !matches!(local.as_ref(), b"majorFont" | b"minorFont") {
                return Err(invalid(
                    "font scheme contains an unsupported replacement child",
                ));
            }
            &[]
        } else if matches!(parent, b"majorFont" | b"minorFont") {
            match local.as_ref() {
                b"latin" | b"ea" | b"cs" => &[b"typeface"],
                b"font" => &[b"script", b"typeface"],
                _ => {
                    return Err(invalid(
                        "font face contains an unsupported replacement child",
                    ));
                },
            }
        } else {
            return Err(invalid(
                "font scheme contains an unsupported replacement child",
            ));
        }
    } else {
        return Err(invalid("unsupported theme scheme for replacement"));
    };
    ensure_replacement_attributes(element, allowed, limits)
}

fn is_color_replacement_value(local: &[u8], parent: &[u8]) -> bool {
    is_color_replacement_slot(parent) && is_color_replacement_value_name(local)
}

fn is_color_replacement_slot(local: &[u8]) -> bool {
    Slot::from_token(std::str::from_utf8(local).unwrap_or_default()).is_some()
}

fn is_color_replacement_value_name(local: &[u8]) -> bool {
    matches!(local, b"srgbClr" | b"sysClr")
}

fn ensure_replacement_attributes(
    element: &BytesStart<'_>,
    allowed: &[&[u8]],
    limits: Limits,
) -> Result<()> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        if attribute.key.prefix().is_some() || !allowed.contains(&key) {
            return Err(invalid(
                "theme replacement contains an unsupported attribute",
            ));
        }
        if allowed.contains(&key) {
            crate::preflight_attribute_value_bytes(
                attribute.value.as_ref(),
                limits.max_attribute_value_bytes,
            )?;
        }
    }
    Ok(())
}

fn validate_theme_namespace_binding(prefix: &[u8], value: &str) -> Result<()> {
    if prefix == b"xmlns" {
        return Err(invalid(
            "theme replacement rebinds the reserved xmlns prefix",
        ));
    }
    if value == XMLNS_NAMESPACE {
        return Err(invalid(
            "theme replacement binds the reserved XMLNS namespace",
        ));
    }
    if value == XML_NAMESPACE && prefix != b"xml" {
        return Err(invalid(
            "theme replacement binds the reserved XML namespace to another prefix",
        ));
    }
    if !prefix.is_empty() && value.is_empty() {
        return Err(invalid(
            "theme replacement uses an empty prefixed namespace binding",
        ));
    }
    if prefix == b"xml" && value != XML_NAMESPACE {
        return Err(invalid("theme replacement rebinds the xml namespace"));
    }
    Ok(())
}

/// Reject a local namespace declaration that leaves a modeled child in the
/// opposite DrawingML conformance namespace. The root binding is normalized
/// before output, but a nested declaration would survive that rewrite.
fn validate_modeled_namespace_binding(
    element: &BytesStart<'_>,
    root: bool,
    target: &[u8],
    namespace: &ResolveResult<'_>,
    source_namespace: &[u8],
    decoder: quick_xml::encoding::Decoder,
    opaque_parent: bool,
) -> Result<()> {
    if root || opaque_parent || !is_known_replacement_local(element.local_name().as_ref(), target) {
        return Ok(());
    }
    let counterpart = if source_namespace == NAMESPACE.as_bytes() {
        STRICT_NAMESPACE.as_bytes()
    } else if source_namespace == STRICT_NAMESPACE.as_bytes() {
        NAMESPACE.as_bytes()
    } else {
        return Err(invalid("theme source has an unsupported namespace"));
    };
    if !matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == counterpart) {
        return Ok(());
    }
    let element_prefix = element.name().prefix().map(|prefix| prefix.into_inner());
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        let key = attribute.key.as_ref();
        let binding_prefix = if key == b"xmlns" {
            Some(&[][..])
        } else {
            key.strip_prefix(b"xmlns:")
        };
        let Some(binding_prefix) = binding_prefix else {
            continue;
        };
        let applies_to_element = match element_prefix {
            Some(prefix) => binding_prefix == prefix,
            None => binding_prefix.is_empty(),
        };
        if !applies_to_element {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?;
        if value.as_bytes() == counterpart {
            return Err(invalid(
                "theme replacement modeled child locally rebinds a cross-conformance namespace",
            ));
        }
    }
    Ok(())
}

fn is_known_replacement_local(local: &[u8], target: &[u8]) -> bool {
    if local == target {
        return true;
    }
    matches!(
        local,
        b"dk1"
            | b"lt1"
            | b"dk2"
            | b"lt2"
            | b"accent1"
            | b"accent2"
            | b"accent3"
            | b"accent4"
            | b"accent5"
            | b"accent6"
            | b"hlink"
            | b"folHlink"
            | b"srgbClr"
            | b"sysClr"
            | b"schemeClr"
            | b"scrgbClr"
            | b"prstClr"
            | b"tint"
            | b"shade"
            | b"alpha"
            | b"red"
            | b"green"
            | b"blue"
            | b"hue"
            | b"sat"
            | b"lum"
            | b"gamma"
            | b"invGamma"
            | b"fontScheme"
            | b"majorFont"
            | b"minorFont"
            | b"latin"
            | b"ea"
            | b"cs"
            | b"font"
            | b"clrScheme"
    )
}

fn validate_replacement_attributes(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    limits: Limits,
) -> Result<()> {
    let mut count = 0usize;
    let mut expanded = BTreeSet::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("theme replacement attribute count overflows"))?;
        if count > limits.max_attributes_per_element {
            return Err(limit(
                LimitKind::XmlAttributes,
                count,
                limits.max_attributes_per_element,
            ));
        }
        if attribute.value.contains(&b'<') {
            return Err(invalid("theme replacement attribute contains raw '<'"));
        }
        let is_namespace = attribute.key.as_ref() == b"xmlns"
            || attribute.key.as_ref().strip_prefix(b"xmlns:").is_some();
        // Namespace bindings have their own structural budget.  Do not charge
        // the selected scalar/attribute-value budget for a binding URI: a
        // narrow value policy applies to modeled attribute values, while the
        // namespace URI remains bounded by the crate-wide XML hard cap and
        // the namespace-count checks performed by the surrounding visitor.
        let raw_value_limit = if is_namespace {
            MAX_ATTRIBUTE_VALUE_BYTES
        } else {
            limits.max_attribute_value_bytes
        };
        crate::preflight_attribute_value_bytes(attribute.value.as_ref(), raw_value_limit)?;
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, reader.decoder())
            .map_err(xml_error)?;
        validate_xml_text("theme replacement attribute", &value)?;
        if is_namespace {
            let prefix = attribute
                .key
                .as_ref()
                .strip_prefix(b"xmlns:")
                .unwrap_or(&[]);
            validate_theme_namespace_binding(prefix, &value)?;
        } else {
            let (namespace, _local) = reader.resolver().resolve_attribute(attribute.key);
            if matches!(namespace, ResolveResult::Unknown(_)) {
                return Err(invalid(
                    "theme replacement attribute uses an undeclared namespace prefix",
                ));
            }
            let namespace = match namespace {
                ResolveResult::Bound(Namespace(value)) => {
                    Some(normalized_namespace_name(value, reader.decoder())?)
                },
                ResolveResult::Unbound => None,
                ResolveResult::Unknown(_) => unreachable!("unknown namespace was rejected"),
            };
            if !expanded.insert(ExpandedAttributeName {
                namespace,
                local: attribute.key.local_name().into_inner(),
            }) {
                return Err(invalid(
                    "theme replacement has duplicate expanded attributes",
                ));
            }
        }
        if !is_namespace && value.len() > limits.max_attribute_value_bytes {
            return Err(limit(
                LimitKind::AttributeValueBytes,
                value.len(),
                limits.max_attribute_value_bytes,
            ));
        }
    }
    Ok(())
}

#[derive(Debug, Eq, Ord, PartialEq, PartialOrd)]
struct ExpandedAttributeName<'a> {
    namespace: Option<Cow<'a, [u8]>>,
    local: &'a [u8],
}

fn normalized_namespace_name<'a>(
    value: &'a [u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<Cow<'a, [u8]>> {
    let decoded = decoder.decode(value).map_err(xml_error)?;
    let unescaped = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
    if unescaped.as_ref().as_bytes() == value {
        Ok(Cow::Borrowed(value))
    } else {
        Ok(Cow::Owned(unescaped.into_owned().into_bytes()))
    }
}

fn optional_attribute_value_range(raw: &[u8], wanted: &[u8]) -> Result<Option<Range<usize>>> {
    let mut cursor = 1usize;
    while cursor < raw.len()
        && !raw[cursor].is_ascii_whitespace()
        && !matches!(raw[cursor], b'>' | b'/')
    {
        cursor += 1;
    }
    let mut found = None;
    while cursor < raw.len() {
        while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= raw.len() || matches!(raw[cursor], b'>' | b'/') {
            break;
        }
        let key_start = cursor;
        while cursor < raw.len()
            && !raw[cursor].is_ascii_whitespace()
            && !matches!(raw[cursor], b'=' | b'>' | b'/')
        {
            cursor += 1;
        }
        let key_end = cursor;
        while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= raw.len() || raw[cursor] != b'=' {
            return Err(invalid("theme replacement attribute has no value"));
        }
        cursor += 1;
        while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *raw
            .get(cursor)
            .ok_or_else(|| invalid("theme replacement attribute value is missing"))?;
        if !matches!(quote, b'"' | b'\'') {
            return Err(invalid("theme replacement attribute value is not quoted"));
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < raw.len() && raw[cursor] != quote {
            cursor += 1;
        }
        let value_end = cursor;
        if cursor >= raw.len() {
            return Err(invalid("theme replacement attribute value is unterminated"));
        }
        if &raw[key_start..key_end] == wanted {
            if found.is_some() {
                return Err(invalid(
                    "theme replacement has duplicate namespace attributes",
                ));
            }
            found = Some(value_start..value_end);
        }
        cursor += 1;
    }
    Ok(found)
}

fn direct_scheme_range(xml: &[u8], target: &[u8], limits: Limits) -> Result<Option<Range<usize>>> {
    let limits = limits.validate()?;
    if xml.len() > limits.max_xml_bytes {
        return Err(limit(LimitKind::XmlBytes, xml.len(), limits.max_xml_bytes));
    }
    if !matches!(target, b"clrScheme" | b"fontScheme") {
        return Err(invalid("unsupported theme scheme for replacement"));
    }
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().check_end_names = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(limits.max_namespace_declarations);
    let mut stack: Vec<(usize, Frame)> = Vec::new();
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut found = None;
    let mut target_namespace = None;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("theme XML offset exceeds usize"))?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("theme XML offset exceeds usize"))?;
        let (namespace, event) = reader.resolver().resolve_event(event);
        reject_unknown_namespace(&namespace)?;
        match event {
            Event::Start(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("theme node count overflow"))?;
                if nodes > limits.max_nodes {
                    return Err(limit(LimitKind::XmlNodes, nodes, limits.max_nodes));
                }
                validate_replacement_attributes(&element, &reader, limits)?;
                let local = element.local_name().as_ref().to_vec();
                let current_namespace = drawingml_namespace(&namespace);
                let drawingml = current_namespace.is_some();
                if !root_seen {
                    if local != b"theme" || !drawingml {
                        return Err(invalid("theme XML has an unexpected root"));
                    }
                    root_seen = true;
                } else if stack.is_empty() {
                    return Err(invalid("theme XML has multiple root elements"));
                }
                let direct =
                    is_direct_theme_parent(&stack) && drawingml && local.as_slice() == target;
                if direct {
                    if found.is_some() {
                        return Err(invalid("theme contains multiple replacement schemes"));
                    }
                    let source_namespace = current_namespace
                        .ok_or_else(|| invalid("theme scheme contains an unsupported namespace"))?;
                    validate_scheme_element(
                        target,
                        &stack,
                        &local,
                        current_namespace,
                        source_namespace,
                        &element,
                        reader.decoder(),
                        limits,
                    )?;
                    target_namespace = Some(source_namespace);
                } else if is_target_stack(&stack, target) {
                    let source_namespace = target_namespace
                        .ok_or_else(|| invalid("theme scheme namespace state is missing"))?;
                    validate_scheme_element(
                        target,
                        &stack,
                        &local,
                        current_namespace,
                        source_namespace,
                        &element,
                        reader.decoder(),
                        limits,
                    )?;
                }
                if stack.len() >= limits.max_depth {
                    return Err(limit(
                        LimitKind::XmlDepth,
                        stack.len() + 1,
                        limits.max_depth,
                    ));
                }
                stack.push((start, Frame { local, drawingml }));
            },
            Event::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("theme node count overflow"))?;
                if nodes > limits.max_nodes {
                    return Err(limit(LimitKind::XmlNodes, nodes, limits.max_nodes));
                }
                validate_replacement_attributes(&element, &reader, limits)?;
                let local = element.local_name().as_ref().to_vec();
                let current_namespace = drawingml_namespace(&namespace);
                let drawingml = current_namespace.is_some();
                if !root_seen {
                    if local != b"theme" || !drawingml {
                        return Err(invalid("theme XML has an unexpected root"));
                    }
                    root_seen = true;
                } else if stack.is_empty() {
                    return Err(invalid("theme XML has multiple root elements"));
                }
                let direct =
                    is_direct_theme_parent(&stack) && drawingml && local.as_slice() == target;
                if direct || is_target_stack(&stack, target) {
                    let source_namespace = if direct {
                        current_namespace.ok_or_else(|| {
                            invalid("theme scheme contains an unsupported namespace")
                        })?
                    } else {
                        target_namespace
                            .ok_or_else(|| invalid("theme scheme namespace state is missing"))?
                    };
                    validate_scheme_element(
                        target,
                        &stack,
                        &local,
                        current_namespace,
                        source_namespace,
                        &element,
                        reader.decoder(),
                        limits,
                    )?;
                }
                if direct {
                    if found.is_some() {
                        return Err(invalid("theme contains multiple replacement schemes"));
                    }
                    found = Some(start..end);
                }
                if stack.len() >= limits.max_depth {
                    return Err(limit(
                        LimitKind::XmlDepth,
                        stack.len() + 1,
                        limits.max_depth,
                    ));
                }
            },
            Event::End(element) => {
                let current_namespace = drawingml_namespace(&namespace);
                let drawingml = current_namespace.is_some();
                if is_target_stack(&stack, target) {
                    let source_namespace = target_namespace
                        .ok_or_else(|| invalid("theme scheme namespace state is missing"))?;
                    if current_namespace != Some(source_namespace) {
                        return Err(invalid("theme scheme contains a cross-family namespace"));
                    }
                }
                let Some((open, frame)) = stack.pop() else {
                    return Err(invalid("theme XML nesting underflow"));
                };
                if frame.local.as_slice() != element.local_name().as_ref()
                    || frame.drawingml != drawingml
                {
                    return Err(invalid("theme XML closing element does not match"));
                }
                if stack.len() == 2
                    && is_direct_theme_parent(&stack)
                    && frame.drawingml
                    && frame.local.as_slice() == target
                {
                    if found.is_some() {
                        return Err(invalid("theme contains multiple replacement schemes"));
                    }
                    found = Some(open..end);
                    target_namespace = None;
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("theme XML contains forbidden markup"));
            },
            Event::Eof => break,
            Event::Text(text) if is_target_stack(&stack, target) => {
                if !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("theme scheme contains unsupported text"));
                }
            },
            Event::CData(_) | Event::Comment(_) | Event::GeneralRef(_)
                if is_target_stack(&stack, target) =>
            {
                return Err(invalid("theme scheme contains unsupported content"));
            },
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if !root_seen || !stack.is_empty() {
        return Err(invalid("theme XML is unterminated"));
    }
    Ok(found)
}

fn is_direct_theme_parent(stack: &[(usize, Frame)]) -> bool {
    stack.len() == 2
        && stack[0].1.drawingml
        && stack[0].1.local.as_slice() == b"theme"
        && stack[1].1.drawingml
        && stack[1].1.local.as_slice() == b"themeElements"
}

fn is_target_stack(stack: &[(usize, Frame)], target: &[u8]) -> bool {
    stack.len() >= 3
        && is_direct_theme_parent(&stack[..2])
        && stack[2].1.drawingml
        && stack[2].1.local.as_slice() == target
}

fn validate_scheme_element(
    target: &[u8],
    stack: &[(usize, Frame)],
    local: &[u8],
    current_namespace: Option<&[u8]>,
    source_namespace: &[u8],
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    limits: Limits,
) -> Result<()> {
    let current_namespace = current_namespace
        .ok_or_else(|| invalid("theme scheme contains an unsupported namespace"))?;
    if current_namespace != source_namespace {
        return Err(invalid("theme scheme contains an unsupported namespace"));
    }
    if stack.len() == 2 {
        if local != target {
            return Err(invalid("theme scheme contains an unsupported child"));
        }
        ensure_supported_attributes(element, decoder, &[b"name"], source_namespace, limits)?;
        return Ok(());
    }
    let parent = stack
        .last()
        .map(|(_, frame)| frame.local.as_slice())
        .unwrap_or_default();
    if target == b"clrScheme" {
        if parent == b"clrScheme" {
            if Slot::from_token(std::str::from_utf8(local).unwrap_or_default()).is_none() {
                return Err(invalid("color scheme contains an unsupported child"));
            }
            ensure_supported_attributes(element, decoder, &[], source_namespace, limits)?;
        } else if is_color_value_parent(stack) {
            match local {
                b"srgbClr" => ensure_supported_attributes(
                    element,
                    decoder,
                    &[b"val"],
                    source_namespace,
                    limits,
                )?,
                b"sysClr" => {
                    ensure_supported_attributes(
                        element,
                        decoder,
                        &[b"val", b"lastClr"],
                        source_namespace,
                        limits,
                    )?;
                },
                _ => return Err(invalid("color slot contains an unsupported child")),
            }
        } else {
            return Err(invalid("color scheme contains an unsupported child"));
        }
    } else if target == b"fontScheme" {
        if parent == b"fontScheme" {
            if !matches!(local, b"majorFont" | b"minorFont") {
                return Err(invalid("font scheme contains an unsupported child"));
            }
            ensure_supported_attributes(element, decoder, &[], source_namespace, limits)?;
        } else if is_font_face_parent_pair(stack) {
            let allowed = match local {
                b"latin" | b"ea" | b"cs" => &[&b"typeface"[..]][..],
                b"font" => &[&b"script"[..], &b"typeface"[..]][..],
                _ => return Err(invalid("font face contains an unsupported child")),
            };
            ensure_supported_attributes(element, decoder, allowed, source_namespace, limits)?;
        } else {
            return Err(invalid("font scheme contains an unsupported child"));
        }
    }
    Ok(())
}

fn is_color_value_parent(stack: &[(usize, Frame)]) -> bool {
    stack.last().is_some_and(|(_, frame)| {
        frame.drawingml
            && Slot::from_token(std::str::from_utf8(&frame.local).unwrap_or_default()).is_some()
    })
}

fn is_font_face_parent_pair(stack: &[(usize, Frame)]) -> bool {
    stack.last().is_some_and(|(_, frame)| {
        frame.drawingml && matches!(frame.local.as_slice(), b"majorFont" | b"minorFont")
    })
}

fn ensure_supported_attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    allowed: &[&[u8]],
    source_namespace: &[u8],
    limits: Limits,
) -> Result<()> {
    let mut count = 0usize;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("theme attribute count overflow"))?;
        if count > limits.max_attributes_per_element {
            return Err(limit(
                LimitKind::XmlAttributes,
                count,
                limits.max_attributes_per_element,
            ));
        }
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            crate::preflight_attribute_value_bytes(
                attribute.value.as_ref(),
                limits.max_attribute_value_bytes,
            )?;
            let value = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?;
            if value.len() > limits.max_attribute_value_bytes {
                return Err(limit(
                    LimitKind::AttributeValueBytes,
                    value.len(),
                    limits.max_attribute_value_bytes,
                ));
            }
            if value.as_ref().as_bytes() != source_namespace {
                return Err(invalid(
                    "theme scheme declares a foreign or cross-family namespace",
                ));
            }
            continue;
        }
        if attribute.key.prefix().is_some() || !allowed.contains(&key) {
            return Err(invalid("theme scheme contains an unsupported attribute"));
        }
        crate::preflight_attribute_value_bytes(
            attribute.value.as_ref(),
            limits.max_attribute_value_bytes,
        )?;
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        if value.len() > limits.max_attribute_value_bytes {
            return Err(limit(
                LimitKind::AttributeValueBytes,
                value.len(),
                limits.max_attribute_value_bytes,
            ));
        }
    }
    Ok(())
}

fn bounded(value: Vec<u8>, limits: Limits) -> Result<Vec<u8>> {
    if value.len() > limits.max_xml_bytes {
        return Err(limit(
            LimitKind::XmlBytes,
            value.len(),
            limits.max_xml_bytes,
        ));
    }
    Ok(value)
}

fn escape<T: XmlSink>(output: &mut T, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            '\t' => output.push_str("&#x9;"),
            '\n' => output.push_str("&#xA;"),
            '\r' => output.push_str("&#xD;"),
            character => output.push(character),
        }
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::limit(kind, observed, maximum)
}

fn check_limit(kind: LimitKind, value: usize, maximum: usize) -> Result<()> {
    if value == 0 || value > maximum {
        return Err(Error::invalid_limit(kind, value, maximum));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::theme::model::{Color, Face, FontSet, Palette, Slot, System};

    fn encode_part(name: &str, colors: &Palette, fonts: &FontSet) -> Result<Vec<u8>> {
        super::encode_part(name, colors, fonts, Limits::DEFAULT)
    }

    fn encode_override(value: &Override) -> Result<Vec<u8>> {
        super::encode_override(value, Limits::DEFAULT)
    }

    fn read(xml: &[u8]) -> Result<Theme> {
        super::read(xml, Limits::DEFAULT)
    }

    fn read_override(xml: &[u8]) -> Result<Override> {
        super::read_override(xml, Limits::DEFAULT)
    }

    fn replace_scheme(xml: &[u8], local: &[u8], replacement: &[u8]) -> Result<Vec<u8>> {
        super::replace_scheme(xml, local, replacement, Limits::DEFAULT)
    }

    fn scheme_replacement_range(xml: &[u8], local: &[u8]) -> Result<Range<usize>> {
        super::scheme_replacement_range(xml, local, Limits::DEFAULT)
    }

    fn encode_palette_fragment(value: &Palette) -> Result<Vec<u8>> {
        super::encode_palette_fragment(value, Limits::DEFAULT)
    }

    fn valid_palette_fragment() -> Vec<u8> {
        encode_palette_fragment(&palette()).unwrap()
    }

    fn valid_fonts_fragment() -> Vec<u8> {
        encode_fonts_fragment(
            &FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos")),
            Limits::DEFAULT,
        )
        .unwrap()
    }

    fn palette() -> Palette {
        Slot::ALL
            .into_iter()
            .fold(Palette::new("Office"), |palette, slot| {
                let color = if slot == Slot::Dark1 {
                    Color::system(System::WindowText, Some("000000")).unwrap()
                } else {
                    Color::rgb("4F81BD").unwrap()
                };
                palette.with(slot, color)
            })
    }

    #[test]
    fn complete_theme_round_trips() {
        let fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
        let xml = encode_part("Office Theme", &palette(), &fonts).unwrap();
        let parsed = read(&xml).unwrap();
        assert_eq!(parsed.name, "Office Theme");
        assert_eq!(parsed.colors, palette());
        assert_eq!(parsed.fonts, fonts);
    }

    #[test]
    fn authored_outputs_preflight_serialized_size_before_allocation() {
        let colors = palette();
        let fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
        let complete = super::encode_part("Office Theme", &colors, &fonts, Limits::DEFAULT)
            .expect("complete theme");
        assert_eq!(
            super::encode_part("Office Theme", &colors, &fonts, Limits::new(complete.len()),)
                .expect("exact complete-theme cap"),
            complete
        );
        assert!(matches!(
            super::encode_part(
                "Office Theme",
                &colors,
                &fonts,
                Limits::new(complete.len() - 1),
            ),
            Err(Error::Limit {
                kind: LimitKind::XmlBytes,
                ..
            })
        ));

        let override_value = Override::new().colors(colors.clone()).fonts(fonts.clone());
        let override_xml =
            super::encode_override(&override_value, Limits::DEFAULT).expect("theme override");
        assert!(matches!(
            super::encode_override(&override_value, Limits::new(override_xml.len() - 1)),
            Err(Error::Limit {
                kind: LimitKind::XmlBytes,
                ..
            })
        ));

        let palette_xml =
            super::encode_palette_fragment(&colors, Limits::DEFAULT).expect("palette fragment");
        assert!(matches!(
            super::encode_palette_fragment(&colors, Limits::new(palette_xml.len() - 1)),
            Err(Error::Limit {
                kind: LimitKind::XmlBytes,
                ..
            })
        ));

        let fonts_xml = encode_fonts_fragment(&fonts, Limits::DEFAULT).expect("fonts fragment");
        assert!(matches!(
            encode_fonts_fragment(&fonts, Limits::new(fonts_xml.len() - 1)),
            Err(Error::Limit {
                kind: LimitKind::XmlBytes,
                ..
            })
        ));
    }

    #[test]
    fn authored_outputs_preflight_all_selected_structure_dimensions() {
        let colors = palette();
        let fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
        let mut nodes = Limits::DEFAULT;
        nodes.max_nodes = 1;
        assert!(matches!(
            super::encode_part("Office", &colors, &fonts, nodes),
            Err(Error::Limit {
                kind: LimitKind::XmlNodes,
                ..
            })
        ));

        let mut depth = Limits::DEFAULT;
        depth.max_depth = 1;
        assert!(matches!(
            super::encode_part("Office", &colors, &fonts, depth),
            Err(Error::Limit {
                kind: LimitKind::XmlDepth,
                ..
            })
        ));

        let mut attributes = Limits::DEFAULT;
        attributes.max_attributes_per_element = 1;
        assert!(matches!(
            super::encode_part("Office", &colors, &fonts, attributes),
            Err(Error::Limit {
                kind: LimitKind::XmlAttributes,
                ..
            })
        ));

        let mut values = Limits::DEFAULT;
        values.max_attribute_value_bytes = 1;
        assert!(matches!(
            super::encode_part("Office", &colors, &fonts, values),
            Err(Error::Limit {
                kind: LimitKind::AttributeValueBytes,
                ..
            })
        ));

        let mut namespaces = Limits::DEFAULT;
        namespaces.max_namespace_declarations = 1;
        assert!(super::encode_part("Office", &colors, &fonts, namespaces).is_ok());
    }

    #[test]
    fn all_system_color_tokens_round_trip() {
        assert_eq!(System::ALL.len(), 30);
        for kind in System::ALL {
            assert_eq!(System::from_token(kind.token()), Some(kind));
            assert!(Color::system(kind, Some("000000")).is_ok());
        }
        let colors = Slot::ALL
            .into_iter()
            .fold(Palette::new("Office"), |palette, slot| {
                palette.with(
                    slot,
                    Color::system(System::ThreeDDarkShadow, Some("000000")).unwrap(),
                )
            });
        let fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
        let parsed = read(&encode_part("Office", &colors, &fonts).unwrap()).unwrap();
        assert_eq!(parsed.colors, colors);

        let invalid = String::from_utf8(encode_part("Office", &palette(), &fonts).unwrap())
            .unwrap()
            .replacen(
                "<a:srgbClr val=\"4F81BD\"/>",
                "<a:sysClr val=\"futureColor\" lastClr=\"000000\"/>",
                1,
            );
        assert!(read(invalid.as_bytes()).is_err());
    }

    #[test]
    fn override_round_trips_and_scheme_patching_is_local() {
        let colors = palette();
        let fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
        let override_value = Override::new().colors(colors.clone()).fonts(fonts.clone());
        assert_eq!(
            read_override(&encode_override(&override_value).unwrap()).unwrap(),
            override_value
        );
        let original = encode_part("Office", &colors, &fonts).unwrap();
        let color_range = scheme_replacement_range(&original, b"clrScheme").unwrap();
        assert!(color_range.start < color_range.end);
        let changed = Palette::new("Changed");
        let changed = Slot::ALL.into_iter().fold(changed, |palette, slot| {
            palette.with(slot, Color::rgb("FFFFFF").unwrap())
        });
        let fragment = encode_palette_fragment(&changed).unwrap();
        let patched = replace_scheme(&original, b"clrScheme", &fragment).unwrap();
        assert_eq!(read(&patched).unwrap().colors, changed);
        assert!(
            patched
                .windows(FORMAT_SCHEME.len())
                .any(|window| window == FORMAT_SCHEME.as_bytes())
        );
    }

    #[test]
    fn authored_format_scheme_has_schema_required_entries() {
        let fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
        let xml = encode_part("Office", &palette(), &fonts).unwrap();
        let text = String::from_utf8(xml).unwrap();
        assert_eq!(text.matches("<a:solidFill>").count(), 9);
        assert_eq!(text.matches("<a:ln ").count(), 3);
        assert_eq!(text.matches("<a:effectStyle>").count(), 3);
        assert_eq!(text.matches("<a:effectLst/>").count(), 3);
    }

    #[test]
    fn typed_read_requires_drawingml_namespaces_and_direct_scheme_parents() {
        let xml = br#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:foreign" name="Foreign"><a:themeElements><x:clrScheme name="Foreign"><x:dk1><x:srgbClr val="000000"/></x:dk1></x:clrScheme><a:fontScheme name="Foreign"><a:majorFont><a:latin typeface="Aptos"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#;
        assert!(read(xml).is_err());

        let nested = br#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:wrapper><a:clrScheme name="Office"/></a:wrapper></a:themeElements></a:theme>"#;
        assert!(replace_scheme(nested, b"clrScheme", b"<a:clrScheme/>").is_err());

        let undeclared = br#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><x:clrScheme/></a:themeElements></a:theme>"#;
        assert!(read(undeclared).is_err());
    }

    #[test]
    fn replacement_rejects_foreign_roots_but_keeps_opaque_extensions() {
        let fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
        let original = encode_part("Office", &palette(), &fonts).unwrap();
        assert!(
            replace_scheme(
                &original,
                b"clrScheme",
                br#"<x:clrScheme xmlns:x="urn:foreign"/>"#
            )
            .is_err()
        );
        let mut with_opaque = String::from_utf8(valid_palette_fragment()).unwrap();
        with_opaque = with_opaque.replacen(
            "</a:clrScheme>",
            "<x:ext xmlns:x=\"urn:foreign\"><x:item/></x:ext></a:clrScheme>",
            1,
        );
        let patched = replace_scheme(&original, b"clrScheme", with_opaque.as_bytes()).unwrap();
        assert_eq!(read(&patched).unwrap().colors, palette());
        let nested_known_opaque = with_opaque.replace(
            "<x:item/>",
            &format!("<dk1 xmlns=\"{STRICT_NAMESPACE}\"><a:srgbClr val=\"4F81BD\"/></dk1>"),
        );
        let patched_nested_opaque =
            replace_scheme(&original, b"clrScheme", nested_known_opaque.as_bytes()).unwrap();
        let scheme_range = scheme_replacement_range(&original, b"clrScheme").unwrap();
        let mut expected_nested_opaque = original.clone();
        expected_nested_opaque.splice(scheme_range, nested_known_opaque.as_bytes().iter().copied());
        assert_eq!(patched_nested_opaque, expected_nested_opaque);
        assert_eq!(read(&patched_nested_opaque).unwrap().colors, palette());
        let undeclared_attribute = with_opaque.replace("<x:item/>", "<x:item y:flag=\"1\"/>");
        assert!(replace_scheme(&original, b"clrScheme", undeclared_attribute.as_bytes(),).is_err());
        let duplicate_expanded = with_opaque.replace(
            "<x:ext xmlns:x=\"urn:foreign\">",
            "<x:ext xmlns:x=\"urn:foreign\" xmlns:y=\"urn&#x3A;foreign\" x:flag=\"1\" y:flag=\"2\">",
        );
        assert!(replace_scheme(&original, b"clrScheme", duplicate_expanded.as_bytes(),).is_err());
        assert!(
            replace_scheme(
                &original,
                b"clrScheme",
                br#"<d:clrScheme xmlns:d="http://purl.oclc.org/ooxml/drawingml/main"/>"#
            )
            .is_err()
        );
        assert!(
            replace_scheme(
                &original,
                b"clrScheme",
                br#"<clrScheme xmlns="http://purl.oclc.org/ooxml/drawingml/main"/>"#
            )
            .is_err()
        );
        assert!(
            replace_scheme(
                &original,
                b"clrScheme",
                br#"<a:clrScheme><x:dk1 xmlns:x="urn:foreign"/></a:clrScheme>"#
            )
            .is_err()
        );
        assert!(
            replace_scheme(
                &original,
                b"clrScheme",
                br#"<a:clrScheme><a:fontScheme/></a:clrScheme>"#
            )
            .is_err()
        );
        assert!(replace_scheme(&original, b"clrScheme", br#"<a:clrScheme future="1"/>"#).is_err());
        assert!(
            replace_scheme(
                &original,
                b"clrScheme",
                br#"<a:clrScheme><a:future/></a:clrScheme>"#
            )
            .is_err()
        );
        for binding in [
            br#"xmlns:xml="urn:foreign""#.as_slice(),
            br#"xmlns:xmlns="urn:foreign""#.as_slice(),
            br#"xmlns:x="http://www.w3.org/XML/1998/namespace""#.as_slice(),
            br#"xmlns:x="http://www.w3.org/2000/xmlns/""#.as_slice(),
            br#"xmlns:x="""#.as_slice(),
        ] {
            let mut replacement = String::from_utf8(valid_palette_fragment()).unwrap();
            replacement.insert_str(
                replacement.find(" name=\"").unwrap(),
                &format!("{} ", String::from_utf8_lossy(binding)),
            );
            assert!(replace_scheme(&original, b"clrScheme", replacement.as_bytes()).is_err());
        }
    }

    #[test]
    fn scheme_replacement_refuses_ignored_source_content() {
        let fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
        let original =
            String::from_utf8(encode_part("Office", &palette(), &fonts).unwrap()).unwrap();
        let replacement = valid_palette_fragment();
        let with_unknown_child =
            original.replacen("</a:clrScheme>", "<a:extLst/></a:clrScheme>", 1);
        assert!(replace_scheme(with_unknown_child.as_bytes(), b"clrScheme", &replacement).is_err());

        let with_unknown_attribute = original.replacen(
            "<a:clrScheme name=\"Office\">",
            "<a:clrScheme name=\"Office\" future=\"1\">",
            1,
        );
        assert!(
            replace_scheme(
                with_unknown_attribute.as_bytes(),
                b"clrScheme",
                &replacement
            )
            .is_err()
        );

        let with_foreign_namespace = original.replacen(
            "<a:clrScheme name=\"Office\">",
            "<a:clrScheme xmlns:x=\"urn:foreign\" name=\"Office\">",
            1,
        );
        assert!(
            replace_scheme(
                with_foreign_namespace.as_bytes(),
                b"clrScheme",
                &replacement
            )
            .is_err()
        );

        let with_cross_family_namespace = original.replacen(
            "<a:clrScheme name=\"Office\">",
            "<a:clrScheme xmlns:x=\"http://purl.oclc.org/ooxml/drawingml/main\" name=\"Office\">",
            1,
        );
        assert!(
            replace_scheme(
                with_cross_family_namespace.as_bytes(),
                b"clrScheme",
                &replacement
            )
            .is_err()
        );

        let with_matching_namespace = original.replacen(
            "<a:clrScheme name=\"Office\">",
            "<a:clrScheme xmlns:x=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office\">",
            1,
        );
        assert!(
            replace_scheme(
                with_matching_namespace.as_bytes(),
                b"clrScheme",
                &replacement
            )
            .is_ok()
        );
    }

    #[test]
    fn scheme_replacement_rebinds_generated_fragment_to_strict_source() {
        let fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
        let original = encode_part("Office", &palette(), &fonts).unwrap();
        let strict = original
            .windows(NAMESPACE.len())
            .enumerate()
            .find_map(|(index, window)| (window == NAMESPACE.as_bytes()).then_some(index))
            .map(|index| {
                let mut bytes = original.clone();
                bytes.splice(
                    index..index + NAMESPACE.len(),
                    STRICT_NAMESPACE.as_bytes().iter().copied(),
                );
                bytes
            })
            .unwrap();
        let strict_name = Slot::ALL
            .into_iter()
            .fold(Palette::new(NAMESPACE), |palette, slot| {
                palette.with(slot, Color::rgb("FFFFFF").unwrap())
            });
        let patched = replace_scheme(
            &strict,
            b"clrScheme",
            &encode_palette_fragment(&strict_name).unwrap(),
        )
        .unwrap();
        assert!(
            patched
                .windows(STRICT_NAMESPACE.len())
                .any(|window| { window == STRICT_NAMESPACE.as_bytes() })
        );
        assert!(
            patched
                .windows(NAMESPACE.len())
                .any(|window| { window == NAMESPACE.as_bytes() })
        );
        assert!(
            patched
                .windows(b"xmlns:a=\"http://purl.oclc.org/ooxml/drawingml/main\"".len())
                .any(|window| {
                    window == b"xmlns:a=\"http://purl.oclc.org/ooxml/drawingml/main\""
                })
        );
        assert_eq!(read(&patched).unwrap().colors.name(), NAMESPACE);
    }

    #[test]
    fn authored_text_rejects_xml_forbidden_chars_and_preserves_whitespace() {
        let colors = Slot::ALL
            .into_iter()
            .fold(Palette::new("Office\n"), |palette, slot| {
                palette.with(slot, Color::rgb("4472C4").unwrap())
            });
        let fonts = FontSet::new(
            "Office\t",
            Face::new("Aptos\r")
                .east_asian("Noto\n")
                .complex_script("Arabic\t"),
            Face::new("Aptos").script("Latn", "Noto Sans"),
        );
        let xml = encode_part("Theme\n", &colors, &fonts).unwrap();
        assert!(
            xml.windows(b"name=\"Theme&#xA;\"".len())
                .any(|window| window == b"name=\"Theme&#xA;\"")
        );
        assert!(
            xml.windows(b"name=\"Office&#x9;\"".len())
                .any(|window| window == b"name=\"Office&#x9;\"")
        );
        assert!(
            xml.windows(b"typeface=\"Aptos&#xD;\"".len())
                .any(|window| window == b"typeface=\"Aptos&#xD;\"")
        );
        let parsed = read(&xml).unwrap();
        assert_eq!(parsed.name, "Theme\n");
        assert_eq!(parsed.fonts.name(), "Office\t");
        assert_eq!(parsed.fonts.major().latin, "Aptos\r");
        assert_eq!(parsed.fonts.major().east_asian, "Noto\n");
        assert_eq!(parsed.fonts.major().complex_script, "Arabic\t");

        let bad_palette = Slot::ALL
            .into_iter()
            .fold(Palette::new("bad\u{1}"), |palette, slot| {
                palette.with(slot, Color::rgb("4472C4").unwrap())
            });
        let plain_fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
        assert!(encode_part("Office", &bad_palette, &plain_fonts).is_err());
        assert!(
            encode_part(
                "Office",
                &palette(),
                &FontSet::new("Office", Face::new("Aptos\u{B}"), Face::new("Aptos"))
            )
            .is_err()
        );
        assert!(
            encode_part(
                "Office",
                &palette(),
                &FontSet::new(
                    "Office",
                    Face::new("Aptos").script("Latn", "Noto\u{1}"),
                    Face::new("Aptos")
                )
            )
            .is_err()
        );
    }

    #[test]
    fn scheme_replacement_preflights_structure_and_decoded_attributes() {
        let fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
        let original = encode_part("Office", &palette(), &fonts).unwrap();
        let mut node_limits = Limits::DEFAULT;
        node_limits.max_nodes = 2;
        assert!(matches!(
            super::replace_scheme(
                &original,
                b"clrScheme",
                b"<a:clrScheme><a:dk1/><a:lt1/></a:clrScheme>",
                node_limits,
            ),
            Err(Error::Limit {
                kind: LimitKind::XmlNodes,
                observed,
                maximum: 2,
            }) if observed > 2
        ));
        assert!(
            super::replace_scheme(
                &original,
                b"clrScheme",
                b"<a:clrScheme><a:future/></a:wrong>",
                Limits::DEFAULT,
            )
            .is_err()
        );

        let mut value_limits = Limits::DEFAULT;
        value_limits.max_attribute_value_bytes = 16;
        let short_replacement = String::from_utf8(valid_palette_fragment())
            .unwrap()
            .replacen("name=\"Office\"", "name=\"a&amp;bcde\"", 1);
        assert!(
            super::replace_scheme(
                &original,
                b"clrScheme",
                short_replacement.as_bytes(),
                value_limits,
            )
            .is_ok()
        );
        let long_replacement =
            short_replacement.replace("name=\"a&amp;bcde\"", "name=\"a&amp;bcdefghijklmnopq\"");
        assert!(matches!(
            super::replace_scheme(
                &original,
                b"clrScheme",
                long_replacement.as_bytes(),
                value_limits,
            ),
            Err(Error::Limit {
                kind: LimitKind::AttributeValueBytes,
                observed,
                maximum: 16,
            }) if observed > 16
        ));
    }

    #[test]
    fn scheme_replacement_requires_complete_supported_state() {
        let fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
        let original = encode_part("Office", &palette(), &fonts).unwrap();
        let complete = String::from_utf8(valid_palette_fragment()).unwrap();
        let missing_slot = complete.replacen(
            "<a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1>",
            "",
            1,
        );
        assert!(replace_scheme(&original, b"clrScheme", missing_slot.as_bytes()).is_err());
        let duplicate_slot = complete.replacen(
            "</a:clrScheme>",
            "<a:dk1><a:srgbClr val=\"FFFFFF\"/></a:dk1></a:clrScheme>",
            1,
        );
        assert!(replace_scheme(&original, b"clrScheme", duplicate_slot.as_bytes()).is_err());
        let empty_color =
            complete.replacen("<a:srgbClr val=\"4F81BD\"/>", "<a:srgbClr val=\"\"/>", 1);
        assert!(replace_scheme(&original, b"clrScheme", empty_color.as_bytes()).is_err());
        let unsupported_color = complete.replacen(
            "<a:srgbClr val=\"4F81BD\"/>",
            "<a:schemeClr val=\"accent1\"/>",
            1,
        );
        assert!(replace_scheme(&original, b"clrScheme", unsupported_color.as_bytes(),).is_err());

        let complete_fonts = String::from_utf8(valid_fonts_fragment()).unwrap();
        let incomplete_fonts = complete_fonts.replacen("<a:latin typeface=\"Aptos\"/>", "", 1);
        assert!(replace_scheme(&original, b"fontScheme", incomplete_fonts.as_bytes(),).is_err());
        let scripts = (0..=MAX_SCRIPT_COUNT)
            .map(|index| format!("<a:font script=\"S{index}\" typeface=\"Aptos\"/>"))
            .collect::<String>();
        let too_many_scripts =
            complete_fonts.replacen("</a:majorFont>", &format!("{scripts}</a:majorFont>"), 1);
        assert!(replace_scheme(&original, b"fontScheme", too_many_scripts.as_bytes(),).is_err());
        let empty_font_name = complete_fonts.replacen("name=\"Office\"", "name=\"\"", 1);
        assert!(replace_scheme(&original, b"fontScheme", empty_font_name.as_bytes(),).is_err());
    }

    #[test]
    fn scheme_replacement_normalizes_default_and_arbitrary_prefixes() {
        let fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
        let original = encode_part("Office", &palette(), &fonts).unwrap();
        let strict =
            String::from_utf8(original.clone())
                .unwrap()
                .replacen(NAMESPACE, STRICT_NAMESPACE, 1);
        let generated = String::from_utf8(valid_palette_fragment()).unwrap();
        let default_prefix = generated.replace("<a:", "<").replace("</a:", "</").replace(
            "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"",
            "xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/main\"",
        );
        let patched_default =
            replace_scheme(strict.as_bytes(), b"clrScheme", default_prefix.as_bytes()).unwrap();
        assert_eq!(read(&patched_default).unwrap().colors, palette());

        let arbitrary_prefix = generated
            .replace("<a:", "<d:")
            .replace("</a:", "</d:")
            .replace(
                "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"",
                "xmlns:d=\"http://schemas.openxmlformats.org/drawingml/2006/main\"",
            );
        let patched_arbitrary =
            replace_scheme(strict.as_bytes(), b"clrScheme", arbitrary_prefix.as_bytes()).unwrap();
        assert_eq!(read(&patched_arbitrary).unwrap().colors, palette());

        let mut inherited_prefix = generated.replace(
            "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"",
            "xmlns=\"http://purl.oclc.org/ooxml/drawingml/main\"",
        );
        for slot in Slot::ALL {
            let token = slot.token();
            inherited_prefix = inherited_prefix
                .replace(&format!("<a:{token}>"), &format!("<{token}>"))
                .replace(&format!("</a:{token}>"), &format!("</{token}>"));
        }
        let patched_inherited =
            replace_scheme(&original, b"clrScheme", inherited_prefix.as_bytes()).unwrap();
        assert_eq!(read(&patched_inherited).unwrap().colors, palette());
        assert!(
            !patched_inherited
                .windows(b"xmlns=\"http://purl.oclc.org/ooxml/drawingml/main\"".len())
                .any(|window| { window == b"xmlns=\"http://purl.oclc.org/ooxml/drawingml/main\"" })
        );

        let nested_default = generated
            .replacen(
                "<a:dk1>",
                "<dk1 xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/main\">",
                1,
            )
            .replacen("</a:dk1>", "</dk1>", 1);
        assert!(matches!(
            replace_scheme(strict.as_bytes(), b"clrScheme", nested_default.as_bytes()),
            Err(Error::Invalid(message))
                if message.contains("locally rebinds a cross-conformance namespace")
        ));

        let nested_alias = generated
            .replacen(
                "<a:dk1>",
                "<d:dk1 xmlns:d=\"http://schemas.openxmlformats.org/drawingml/2006/main\">",
                1,
            )
            .replacen("</a:dk1>", "</d:dk1>", 1);
        assert!(matches!(
            replace_scheme(strict.as_bytes(), b"clrScheme", nested_alias.as_bytes()),
            Err(Error::Invalid(message))
                if message.contains("locally rebinds a cross-conformance namespace")
        ));
    }
}
