//! Source-preserving support for `[MS-ODRAWXML]` §2.44 `formatcode2`.
//!
//! The specification defines `formatcode2` as an `ST_Xstring` element and as
//! an attribute in the `drawing/2015/06/chart` namespace. It does not define
//! a containing chart type or a unique host placement. This module therefore
//! owns one complete `formatcode2` element and a selected host start-tag
//! attribute. Host packages may validate placement and use the corresponding
//! read/write helpers at their own grammar boundary; this shared owner does
//! not invent a parent or attach either form to an arbitrary chart node.

use std::{collections::HashSet, ops::Range, sync::Arc};

use litchi_core::xml::ReaderOrigin;
use litchi_core::xml::escape_xml;
use litchi_ooxml_common::xml_name::{is_ncname, is_qualified_name};
use quick_xml::{
    XmlVersion,
    escape::{resolve_xml_entity, unescape, unescape_with},
    events::{BytesDecl, BytesEnd, BytesPI, BytesRef, BytesStart, BytesText, Event},
    name::ResolveResult,
    reader::{NsReader, Reader},
};

use crate::{Error, LimitKind, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

/// `[MS-ODRAWXML]` §2.44 target namespace.
pub const NAMESPACE: &str = "http://schemas.microsoft.com/office/drawing/2015/06/chart";
/// Maximum complete source fragment bytes retained by this owner.
pub const MAX_XML_BYTES: usize = 1 << 20;
/// Maximum decoded string bytes retained by the typed projection.
pub const MAX_VALUE_BYTES: usize = 64 * 1024;
/// Maximum lexical bytes retained while decoding the SpreadsheetML
/// `ST_Xstring` escape form. A legal escape can expand to one UTF-8 scalar
/// from seven ASCII bytes, so this is charged independently of the decoded
/// semantic value bound.
const MAX_LEXICAL_VALUE_BYTES: usize = MAX_VALUE_BYTES * 7;
/// Maximum attributes inspected on the root element.
pub const MAX_ATTRIBUTES: usize = 64;
/// Maximum element nodes accepted by the fragment scanner.
pub const MAX_NODES: usize = 16;
/// Maximum nesting depth accepted by the fragment scanner.
pub const MAX_DEPTH: usize = 8;
/// Maximum namespace prefix or URI bytes inspected on one declaration.
pub const MAX_NAMESPACE_BYTES: usize = 4 * 1024;

/// Caller-selected bounds for raw `formatcode2` XML operations.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
#[must_use]
pub struct Limits {
    pub max_xml_bytes: usize,
    pub max_value_bytes: usize,
    pub max_attributes: usize,
    pub max_nodes: usize,
    pub max_depth: usize,
    pub max_namespace_bytes: usize,
    pub max_output_bytes: usize,
}

impl Limits {
    pub const DEFAULT: Self = Self {
        max_xml_bytes: MAX_XML_BYTES,
        max_value_bytes: MAX_VALUE_BYTES,
        max_attributes: MAX_ATTRIBUTES,
        max_nodes: MAX_NODES,
        max_depth: MAX_DEPTH,
        max_namespace_bytes: MAX_NAMESPACE_BYTES,
        max_output_bytes: MAX_XML_BYTES,
    };

    #[must_use]
    pub const fn new(
        max_xml_bytes: usize,
        max_value_bytes: usize,
        max_attributes: usize,
        max_nodes: usize,
        max_depth: usize,
        max_namespace_bytes: usize,
        max_output_bytes: usize,
    ) -> Self {
        Self {
            max_xml_bytes,
            max_value_bytes,
            max_attributes,
            max_nodes,
            max_depth,
            max_namespace_bytes,
            max_output_bytes,
        }
    }

    fn validate(self) -> Result<Self> {
        for (kind, value, maximum) in [
            (LimitKind::XmlBytes, self.max_xml_bytes, MAX_XML_BYTES),
            (LimitKind::TextBytes, self.max_value_bytes, MAX_VALUE_BYTES),
            (
                LimitKind::XmlAttributes,
                self.max_attributes,
                MAX_ATTRIBUTES,
            ),
            (LimitKind::XmlNodes, self.max_nodes, MAX_NODES),
            (LimitKind::XmlDepth, self.max_depth, MAX_DEPTH),
            (
                LimitKind::NamespaceBytes,
                self.max_namespace_bytes,
                MAX_NAMESPACE_BYTES,
            ),
            (LimitKind::OutputBytes, self.max_output_bytes, MAX_XML_BYTES),
        ] {
            if value == 0 || value > maximum {
                return Err(Error::InvalidLimit {
                    kind,
                    value,
                    maximum,
                });
            }
        }
        Ok(self)
    }

    fn max_lexical_value_bytes(self) -> Result<usize> {
        self.max_value_bytes
            .checked_mul(7)
            .map(|value| value.min(MAX_LEXICAL_VALUE_BYTES))
            .ok_or_else(|| {
                invalid("formatcode2 lexical value limit overflows the selected value budget")
            })
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE: &[u8] = b"http://www.w3.org/2000/xmlns/";

/// A checked semantic `ST_Xstring` value shared by the element and attribute
/// owners in this module.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct Value {
    value: Arc<str>,
}

impl Value {
    /// Construct a bounded semantic value. SpreadsheetML escape sequences are
    /// applied only when this value is serialized.
    pub fn new(value: impl AsRef<str>) -> Result<Self> {
        Ok(Self {
            value: checked_value(value.as_ref())?,
        })
    }

    /// Borrow the decoded semantic string.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.value
    }

    /// Replace the semantic string after applying the value bound.
    pub fn set(&mut self, value: impl AsRef<str>) -> Result<&mut Self> {
        self.value = checked_value(value.as_ref())?;
        Ok(self)
    }
}

/// A checked `formatcode2` element with an optional exact source backing.
#[derive(Debug, Clone)]
#[must_use]
pub struct Element {
    value: Value,
    source: Option<Arc<Source>>,
}

/// A source-backed `formatcode2` attribute on one complete host start tag.
///
/// The shared specification does not identify a containing chart type, so the
/// helper validates the qualified attribute and its start-tag XML only. A
/// package host must validate the parent element and placement before calling
/// [`read_attribute`].
#[derive(Debug, Clone)]
#[must_use]
pub struct Attribute {
    value: Value,
    source: Option<Arc<AttributeSource>>,
}

impl PartialEq for Attribute {
    fn eq(&self, other: &Self) -> bool {
        self.value == other.value
    }
}

impl Eq for Attribute {}

impl Attribute {
    /// Borrow the decoded semantic string.
    #[must_use]
    pub fn value(&self) -> &str {
        self.value.as_str()
    }

    /// Return the exact retained host start tag, if parsed from XML.
    #[must_use]
    pub fn source(&self) -> Option<&[u8]> {
        self.source.as_deref().map(|source| source.xml.as_ref())
    }

    /// Replace the semantic string while retaining the host tag source.
    pub fn set_value(&mut self, value: impl AsRef<str>) -> Result<&mut Self> {
        self.value.set(value)?;
        Ok(self)
    }

    fn source_state(&self) -> Option<&AttributeSource> {
        self.source.as_deref()
    }
}

#[derive(Debug)]
struct AttributeSource {
    xml: Arc<[u8]>,
    value_range: Range<usize>,
    value: Value,
}

impl PartialEq for Element {
    fn eq(&self, other: &Self) -> bool {
        self.value == other.value
    }
}

impl Eq for Element {}

impl Element {
    /// Construct a detached bounded `formatcode2` value.
    ///
    /// The string follows the XML `ST_Xstring` domain. XML-forbidden control
    /// characters are encoded as `_xHHHH_` on output, as required by the
    /// SpreadsheetML string type.
    ///
    /// # Errors
    ///
    /// Returns an error when the decoded value exceeds [`MAX_VALUE_BYTES`].
    pub fn new(value: impl AsRef<str>) -> Result<Self> {
        let value = Value::new(value)?;
        Ok(Self {
            value,
            source: None,
        })
    }

    /// Borrow the decoded format string.
    #[must_use]
    pub fn value(&self) -> &str {
        self.value.as_str()
    }

    /// Return the exact retained source, if this value was parsed from XML.
    #[must_use]
    pub fn source(&self) -> Option<&[u8]> {
        self.source.as_deref().map(|source| source.xml.as_ref())
    }

    /// Replace the decoded format string while retaining source markup for a
    /// subsequent scalar write.
    pub fn set_value(&mut self, value: impl AsRef<str>) -> Result<&mut Self> {
        self.value.set(value)?;
        Ok(self)
    }

    fn source_state(&self) -> Option<&Source> {
        self.source.as_deref()
    }
}

#[derive(Debug)]
struct Source {
    xml: Arc<[u8]>,
    root_start: usize,
    root_end: usize,
    root_qname: Vec<u8>,
    content_start: Option<usize>,
    content_end: Option<usize>,
    segments: Vec<Segment>,
    value: Value,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SegmentKind {
    Text,
    CData,
    GeneralRef,
}

#[derive(Debug, Clone)]
struct Segment {
    range: Range<usize>,
    kind: SegmentKind,
}

#[derive(Debug)]
struct Parsed {
    root_start: usize,
    root_end: usize,
    root_qname: Vec<u8>,
    content_start: Option<usize>,
    content_end: Option<usize>,
    segments: Vec<Segment>,
    value: Value,
}

/// Read one complete `formatcode2` element while retaining its source bytes.
///
/// XML declarations and comments around the root are accepted for standalone
/// fragment use. A host embedding a source fragment must apply its own
/// element-only placement rule; [`write()`] preserves the accepted source.
pub fn read(xml: &[u8], limits: &Limits) -> Result<Element> {
    let limits = limits.validate()?;
    if xml.len() > limits.max_xml_bytes {
        return Err(limit(LimitKind::XmlBytes, xml.len(), limits.max_xml_bytes));
    }
    let parsed = scan(xml, limits)?;
    if parsed.value.as_str().len() > limits.max_value_bytes {
        return Err(limit(
            LimitKind::TextBytes,
            parsed.value.as_str().len(),
            limits.max_value_bytes,
        ));
    }
    let source = Arc::from(xml);
    Ok(Element {
        value: parsed.value.clone(),
        source: Some(Arc::new(Source {
            xml: source,
            root_start: parsed.root_start,
            root_end: parsed.root_end,
            root_qname: parsed.root_qname,
            content_start: parsed.content_start,
            content_end: parsed.content_end,
            segments: parsed.segments,
            value: parsed.value,
        })),
    })
}

/// Read a complete fragment while retaining an existing immutable source
/// allocation.
pub fn read_shared(xml: Arc<[u8]>, limits: &Limits) -> Result<Element> {
    let limits = limits.validate()?;
    if xml.len() > limits.max_xml_bytes {
        return Err(limit(LimitKind::XmlBytes, xml.len(), limits.max_xml_bytes));
    }
    let parsed = scan(xml.as_ref(), limits)?;
    if parsed.value.as_str().len() > limits.max_value_bytes {
        return Err(limit(
            LimitKind::TextBytes,
            parsed.value.as_str().len(),
            limits.max_value_bytes,
        ));
    }
    Ok(Element {
        value: parsed.value.clone(),
        source: Some(Arc::new(Source {
            xml,
            root_start: parsed.root_start,
            root_end: parsed.root_end,
            root_qname: parsed.root_qname,
            content_start: parsed.content_start,
            content_end: parsed.content_end,
            segments: parsed.segments,
            value: parsed.value,
        })),
    })
}

/// Read one complete XML start tag containing the qualified `formatcode2`
/// attribute. The tag may be self-closing or an opening tag whose matching
/// end tag is supplied by the host outside this fragment.
pub fn read_attribute(xml: &[u8], limits: &Limits) -> Result<Attribute> {
    let limits = limits.validate()?;
    if xml.len() > limits.max_xml_bytes {
        return Err(limit(LimitKind::XmlBytes, xml.len(), limits.max_xml_bytes));
    }
    let parsed = scan_attribute(xml, limits)?;
    if parsed.value.as_str().len() > limits.max_value_bytes {
        return Err(limit(
            LimitKind::TextBytes,
            parsed.value.as_str().len(),
            limits.max_value_bytes,
        ));
    }
    let source = Arc::from(xml);
    Ok(Attribute {
        value: parsed.value.clone(),
        source: Some(Arc::new(AttributeSource {
            xml: source,
            value_range: parsed.value_range,
            value: parsed.value,
        })),
    })
}

/// Read a qualified `formatcode2` attribute whose element or attribute prefix
/// is declared by the host outside this selected start tag.
///
/// Each pair is `(prefix, namespace URI)`. The helper uses these declarations
/// only to resolve this one tag; it does not infer a host element, placement,
/// or namespace scope. Local declarations on the selected tag take precedence
/// over a supplied inherited binding. The retained source remains the exact
/// caller-provided tag, without synthetic declarations. The temporary
/// namespace-complete view must also fit [`MAX_XML_BYTES`] and
/// [`MAX_ATTRIBUTES`]; inherited declarations consume that temporary headroom.
pub fn read_attribute_with_bindings(
    xml: &[u8],
    bindings: &[(&str, &str)],
    limits: &Limits,
) -> Result<Attribute> {
    let limits = limits.validate()?;
    if xml.len() > limits.max_xml_bytes {
        return Err(limit(LimitKind::XmlBytes, xml.len(), limits.max_xml_bytes));
    }
    let parsed = scan_attribute_with_bindings(xml, bindings, limits)?;
    if parsed.value.as_str().len() > limits.max_value_bytes {
        return Err(limit(
            LimitKind::TextBytes,
            parsed.value.as_str().len(),
            limits.max_value_bytes,
        ));
    }
    Ok(Attribute {
        value: parsed.value.clone(),
        source: Some(Arc::new(AttributeSource {
            xml: Arc::from(xml),
            value_range: parsed.value_range,
            value: parsed.value,
        })),
    })
}

/// Read an attribute start tag while retaining an existing immutable source
/// allocation.
pub fn read_attribute_shared(xml: Arc<[u8]>, limits: &Limits) -> Result<Attribute> {
    let limits = limits.validate()?;
    if xml.len() > limits.max_xml_bytes {
        return Err(limit(LimitKind::XmlBytes, xml.len(), limits.max_xml_bytes));
    }
    let parsed = scan_attribute(xml.as_ref(), limits)?;
    if parsed.value.as_str().len() > limits.max_value_bytes {
        return Err(limit(
            LimitKind::TextBytes,
            parsed.value.as_str().len(),
            limits.max_value_bytes,
        ));
    }
    Ok(Attribute {
        value: parsed.value.clone(),
        source: Some(Arc::new(AttributeSource {
            xml,
            value_range: parsed.value_range,
            value: parsed.value,
        })),
    })
}

/// Shared-source variant of [`read_attribute_with_bindings`].
pub fn read_attribute_shared_with_bindings(
    xml: Arc<[u8]>,
    bindings: &[(&str, &str)],
    limits: &Limits,
) -> Result<Attribute> {
    let limits = limits.validate()?;
    if xml.len() > limits.max_xml_bytes {
        return Err(limit(LimitKind::XmlBytes, xml.len(), limits.max_xml_bytes));
    }
    let parsed = scan_attribute_with_bindings(xml.as_ref(), bindings, limits)?;
    if parsed.value.as_str().len() > limits.max_value_bytes {
        return Err(limit(
            LimitKind::TextBytes,
            parsed.value.as_str().len(),
            limits.max_value_bytes,
        ));
    }
    Ok(Attribute {
        value: parsed.value.clone(),
        source: Some(Arc::new(AttributeSource {
            xml,
            value_range: parsed.value_range,
            value: parsed.value,
        })),
    })
}

/// Serialize only the XML attribute value lexical form for a host start tag.
pub fn write_attribute_value(value: &Value, limits: &Limits) -> Result<Vec<u8>> {
    let limits = limits.validate()?;
    if value.as_str().len() > limits.max_value_bytes {
        return Err(limit(
            LimitKind::TextBytes,
            value.as_str().len(),
            limits.max_value_bytes,
        ));
    }
    let encoded_len = escaped_xstring_xml_len(value.as_str(), XStringContext::Attribute)?;
    ensure_output_limit_with(encoded_len, limits.max_output_bytes)?;
    let encoded = encode_xstring_xml(value.as_str(), XStringContext::Attribute)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(encoded.len())
        .map_err(|_| invalid("formatcode2 attribute output allocation failed"))?;
    output.extend_from_slice(encoded.as_bytes());
    Ok(output)
}

/// Serialize a source-backed `formatcode2` attribute start tag.
pub fn write_attribute(value: &Attribute, limits: &Limits) -> Result<Vec<u8>> {
    let limits = limits.validate()?;
    validate_attribute_value_with(value, limits.max_value_bytes)?;
    if let Some(source) = value.source_state() {
        if source.value == value.value {
            return clone_bounded_with(source.xml.as_ref(), limits.max_output_bytes);
        }
        return rewrite_attribute(source, value.value.as_str(), limits.max_output_bytes);
    }
    Err(invalid(
        "formatcode2 attribute has no source backing; use write_attribute_value",
    ))
}

/// Serialize a source-backed `formatcode2` attribute to a caller-provided
/// sink.
pub fn write_attribute_to<W: std::io::Write>(
    writer: &mut W,
    value: &Attribute,
    limits: &Limits,
) -> Result<()> {
    let limits = limits.validate()?;
    validate_attribute_value_with(value, limits.max_value_bytes)?;
    if let Some(source) = value.source_state()
        && source.value == value.value
    {
        ensure_output_limit_with(source.xml.len(), limits.max_output_bytes)?;
        writer.write_all(source.xml.as_ref())?;
        return Ok(());
    }
    write_attribute_to_bounded(writer, value, limits.max_output_bytes)
}

#[derive(Debug)]
struct ParsedAttribute {
    value_range: Range<usize>,
    value: Value,
}

/// Inspect raw namespace declaration bytes before `NsReader` stores them in
/// its resolver. The resolver's per-element declaration cap protects its own
/// buffer; this pass also bounds prefixes and raw URI values before that
/// buffer can be populated.
fn preflight_namespace_limits(xml: &[u8], limits: Limits) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    let mut nodes = 0usize;
    let mut depth = 0usize;
    loop {
        let event = reader.read_event().map_err(xml_event_error)?;
        match event {
            Event::Start(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("formatcode2 node count overflow"))?;
                if nodes > limits.max_nodes {
                    return Err(limit(LimitKind::XmlNodes, nodes, limits.max_nodes));
                }
                if depth >= limits.max_depth {
                    return Err(limit(LimitKind::XmlDepth, depth + 1, limits.max_depth));
                }
                depth += 1;
                validate_namespace_declarations(&element, limits)?;
            },
            Event::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("formatcode2 node count overflow"))?;
                if nodes > limits.max_nodes {
                    return Err(limit(LimitKind::XmlNodes, nodes, limits.max_nodes));
                }
                validate_namespace_declarations(&element, limits)?;
            },
            Event::End(end) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("formatcode2 fragment has an unmatched close"))?;
                validate_qname_prefix(end.name().into_inner(), limits.max_namespace_bytes)?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(())
}

fn validate_namespace_declarations(element: &BytesStart<'_>, limits: Limits) -> Result<()> {
    validate_qname_prefix(element.name().into_inner(), limits.max_namespace_bytes)?;
    let mut declarations = 0usize;
    for raw_attribute in element.checked_attributes() {
        let attribute = raw_attribute.map_err(xml_error)?;
        validate_qname_prefix(attribute.key.into_inner(), limits.max_namespace_bytes)?;
        if !is_namespace_attribute(attribute.key) {
            continue;
        }
        declarations = declarations
            .checked_add(1)
            .ok_or_else(|| invalid("formatcode2 namespace declaration count overflow"))?;
        if declarations > limits.max_attributes {
            return Err(limit(
                LimitKind::XmlAttributes,
                declarations,
                limits.max_attributes,
            ));
        }
        let prefix = namespace_attribute_prefix(attribute.key)?;
        if prefix.len() > limits.max_namespace_bytes {
            return Err(limit(
                LimitKind::NamespacePrefixBytes,
                prefix.len(),
                limits.max_namespace_bytes,
            ));
        }
        if attribute.value.as_ref().len() > limits.max_namespace_bytes {
            return Err(limit(
                LimitKind::NamespaceBytes,
                attribute.value.as_ref().len(),
                limits.max_namespace_bytes,
            ));
        }
        let _ = normalize_namespace_uri(attribute.value.as_ref(), limits)?;
    }
    Ok(())
}

fn scan_attribute(xml: &[u8], limits: Limits) -> Result<ParsedAttribute> {
    preflight_namespace_limits(xml, limits)?;
    preflight_attribute_value_budget(xml, limits)?;
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(limits.max_attributes);
    let mut parsed = None;
    loop {
        let event_start = position(&reader, origin)?;
        let (resolved, event) = reader.read_resolved_event().map_err(xml_event_error)?;
        let namespace_result = resolved_namespace(&resolved, limits);
        let event = event.into_owned();
        let event_end = position(&reader, origin)?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                // Only a leading byte-order mark may precede the tag.
                if event_start != origin.skipped() || parsed.is_some() {
                    return Err(invalid(
                        "formatcode2 attribute fragment must contain one start tag",
                    ));
                }
                validate_element_name(&element, limits.max_namespace_bytes)?;
                validate_attributes(&element, &reader, limits)?;
                namespace_result?;
                parsed = Some(attribute_projection(
                    &element,
                    xml,
                    event_start,
                    event_end,
                    &reader,
                    limits,
                )?);
            },
            Event::Eof => break,
            _ => {
                return Err(invalid(
                    "formatcode2 attribute fragment contains non-tag markup",
                ));
            },
        }
    }
    parsed.ok_or_else(|| invalid("formatcode2 attribute fragment has no start tag"))
}

fn scan_attribute_with_bindings(
    xml: &[u8],
    bindings: &[(&str, &str)],
    limits: Limits,
) -> Result<ParsedAttribute> {
    if bindings.is_empty() {
        return scan_attribute(xml, limits);
    }
    preflight_namespace_limits(xml, limits)?;
    preflight_attribute_value_budget(xml, limits)?;
    let local_prefixes = local_namespace_prefixes(xml, limits)?;
    let additions = checked_inherited_bindings(bindings, &local_prefixes, limits)?;
    if additions.is_empty() {
        return scan_attribute(xml, limits);
    }
    let decorated = decorate_attribute_start_tag(xml, &additions, limits)?;
    scan_attribute(&decorated, limits)
}

fn local_namespace_prefixes(xml: &[u8], limits: Limits) -> Result<Vec<Vec<u8>>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(limits.max_attributes);
    let mut prefixes = Vec::new();
    let event = reader.read_event().map_err(xml_event_error)?;
    match event {
        Event::Start(element) | Event::Empty(element) => {
            for raw_attribute in element.checked_attributes() {
                let attribute = raw_attribute.map_err(xml_error)?;
                if !is_namespace_attribute(attribute.key) {
                    continue;
                }
                let prefix = namespace_attribute_prefix(attribute.key)?;
                if prefix.len() > limits.max_namespace_bytes {
                    return Err(limit(
                        LimitKind::NamespacePrefixBytes,
                        prefix.len(),
                        limits.max_namespace_bytes,
                    ));
                }
                if prefixes.len() >= limits.max_attributes {
                    return Err(limit(
                        LimitKind::NamespaceDeclarations,
                        prefixes.len() + 1,
                        limits.max_attributes,
                    ));
                }
                prefixes.push(prefix.to_vec());
            }
            Ok(prefixes)
        },
        Event::Eof => Err(invalid("formatcode2 attribute fragment has no start tag")),
        _ => Err(invalid(
            "formatcode2 attribute fragment must begin with a start tag",
        )),
    }
}

fn checked_inherited_bindings<'a>(
    bindings: &'a [(&'a str, &'a str)],
    local_prefixes: &[Vec<u8>],
    limits: Limits,
) -> Result<Vec<(&'a [u8], &'a [u8])>> {
    if bindings.len() > limits.max_attributes {
        return Err(limit(
            LimitKind::NamespaceDeclarations,
            bindings.len(),
            limits.max_attributes,
        ));
    }
    let mut additions = Vec::new();
    for (prefix, namespace) in bindings {
        let prefix = prefix.as_bytes();
        let namespace = namespace.as_bytes();
        if prefix.len() > limits.max_namespace_bytes {
            return Err(limit(
                LimitKind::NamespacePrefixBytes,
                prefix.len(),
                limits.max_namespace_bytes,
            ));
        }
        if namespace.len() > limits.max_namespace_bytes {
            return Err(limit(
                LimitKind::NamespaceBytes,
                namespace.len(),
                limits.max_namespace_bytes,
            ));
        }
        validate_xml_text(
            std::str::from_utf8(namespace).map_err(xml_error)?,
            "formatcode2 inherited namespace URI",
        )?;
        validate_namespace_binding(prefix, namespace)?;
        if local_prefixes
            .iter()
            .any(|local| local.as_slice() == prefix)
        {
            continue;
        }
        if let Some((_, existing_namespace)) = additions
            .iter()
            .find(|(existing_prefix, _)| *existing_prefix == prefix)
        {
            if *existing_namespace != namespace {
                return Err(invalid(
                    "formatcode2 inherited namespace prefix has conflicting bindings",
                ));
            }
            continue;
        }
        additions.push((prefix, namespace));
    }
    Ok(additions)
}

fn decorate_attribute_start_tag(
    xml: &[u8],
    additions: &[(&[u8], &[u8])],
    limits: Limits,
) -> Result<Vec<u8>> {
    let close = start_tag_close(xml)?;
    let insertion = if close > 0 && xml[close - 1] == b'/' {
        close - 1
    } else {
        close
    };
    let mut addition_len = 0usize;
    for (prefix, namespace) in additions {
        let escaped_len = escaped_xml_bytes_len(namespace)?;
        let declaration_len = if prefix.is_empty() {
            7
        } else {
            8 + prefix.len()
        };
        addition_len = addition_len
            .checked_add(1)
            .and_then(|length| length.checked_add(declaration_len))
            .and_then(|length| length.checked_add(escaped_len))
            .and_then(|length| length.checked_add(1))
            .ok_or_else(|| invalid("formatcode2 attribute XML length overflows"))?;
    }
    let output_len = xml
        .len()
        .checked_add(addition_len)
        .ok_or_else(|| invalid("formatcode2 attribute XML length overflows"))?;
    if output_len > limits.max_xml_bytes {
        return Err(limit(LimitKind::XmlBytes, output_len, limits.max_xml_bytes));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| invalid("formatcode2 attribute output allocation failed"))?;
    output.extend_from_slice(&xml[..insertion]);
    for (prefix, namespace) in additions {
        output.push(b' ');
        output.extend_from_slice(b"xmlns");
        if !prefix.is_empty() {
            output.push(b':');
            output.extend_from_slice(prefix);
        }
        output.extend_from_slice(b"=\"");
        let escaped = escape_xml(std::str::from_utf8(namespace).map_err(xml_error)?);
        output.extend_from_slice(escaped.as_bytes());
        output.push(b'"');
    }
    output.extend_from_slice(&xml[insertion..]);
    Ok(output)
}

fn start_tag_close(xml: &[u8]) -> Result<usize> {
    let mut quote = None;
    for (index, byte) in xml.iter().copied().enumerate() {
        match quote {
            Some(current) if byte == current => quote = None,
            Some(_) => {},
            None if matches!(byte, b'\'' | b'"') => quote = Some(byte),
            None if byte == b'>' => return Ok(index),
            None => {},
        }
    }
    Err(invalid(
        "formatcode2 attribute start tag has no close marker",
    ))
}

fn escaped_xml_bytes_len(value: &[u8]) -> Result<usize> {
    let text = std::str::from_utf8(value).map_err(xml_error)?;
    text.chars().try_fold(0usize, |length, character| {
        length
            .checked_add(escaped_xml_character_len(character))
            .ok_or_else(|| invalid("formatcode2 attribute XML length overflows"))
    })
}

fn attribute_projection(
    element: &BytesStart<'_>,
    xml: &[u8],
    event_start: usize,
    event_end: usize,
    reader: &NsReader<&[u8]>,
    limits: Limits,
) -> Result<ParsedAttribute> {
    let mut selected = None;
    for raw_attribute in element.checked_attributes() {
        let attribute = raw_attribute.map_err(xml_error)?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace_matches = match namespace {
            ResolveResult::Bound(namespace) => {
                normalize_namespace_uri(namespace.into_inner(), limits)? == NAMESPACE.as_bytes()
            },
            ResolveResult::Unbound => false,
            ResolveResult::Unknown(_) => {
                return Err(invalid("formatcode2 attribute has an undeclared prefix"));
            },
        };
        if local.as_ref() != b"formatcode2" || !namespace_matches {
            continue;
        }
        if selected.is_some() {
            return Err(invalid("formatcode2 attribute occurs more than once"));
        }
        let lexical = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?;
        let max_lexical_value_bytes = limits.max_lexical_value_bytes()?;
        preflight_normalized_value(attribute.value.as_ref(), max_lexical_value_bytes)?;
        validate_xml_text(lexical.as_ref(), "formatcode2 attribute value")?;
        let value = decode_xstring(lexical.as_ref(), limits.max_value_bytes)?;
        let value_range = source_value_range(
            xml,
            event_start,
            event_end,
            attribute.key.as_ref(),
            attribute.value.as_ref(),
        )?;
        selected = Some(ParsedAttribute { value_range, value });
    }
    selected.ok_or_else(|| invalid("formatcode2 attribute is absent"))
}

fn rewrite_attribute(
    source: &AttributeSource,
    value: &str,
    max_output_bytes: usize,
) -> Result<Vec<u8>> {
    let replacement_len = escaped_xstring_xml_len(value, XStringContext::Attribute)?;
    let output_len = source
        .xml
        .len()
        .checked_sub(source.value_range.end - source.value_range.start)
        .and_then(|length| length.checked_add(replacement_len))
        .ok_or_else(|| invalid("formatcode2 attribute XML length overflows"))?;
    ensure_output_limit_with(output_len, max_output_bytes)?;
    let replacement = encode_xstring_xml(value, XStringContext::Attribute)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| invalid("formatcode2 attribute output allocation failed"))?;
    output.extend_from_slice(&source.xml[..source.value_range.start]);
    output.extend_from_slice(replacement.as_bytes());
    output.extend_from_slice(&source.xml[source.value_range.end..]);
    Ok(output)
}

fn write_attribute_to_bounded<W: std::io::Write>(
    writer: &mut W,
    value: &Attribute,
    max_output_bytes: usize,
) -> Result<()> {
    let source = value.source_state().ok_or_else(|| {
        invalid("formatcode2 attribute has no source backing; use write_attribute_value")
    })?;
    let replacement_len = escaped_xstring_xml_len(value.value.as_str(), XStringContext::Attribute)?;
    let output_len = source
        .xml
        .len()
        .checked_sub(source.value_range.end - source.value_range.start)
        .and_then(|length| length.checked_add(replacement_len))
        .ok_or_else(|| invalid("formatcode2 attribute XML length overflows"))?;
    ensure_output_limit_with(output_len, max_output_bytes)?;
    let replacement = encode_xstring_xml(value.value.as_str(), XStringContext::Attribute)?;
    writer.write_all(&source.xml[..source.value_range.start])?;
    writer.write_all(replacement.as_bytes())?;
    writer.write_all(&source.xml[source.value_range.end..])?;
    Ok(())
}

fn validate_attribute_value_with(value: &Attribute, maximum: usize) -> Result<()> {
    if value.value.as_str().len() > maximum {
        return Err(limit(
            LimitKind::TextBytes,
            value.value.as_str().len(),
            maximum,
        ));
    }
    Ok(())
}

fn source_value_range(
    xml: &[u8],
    event_start: usize,
    event_end: usize,
    key: &[u8],
    raw: &[u8],
) -> Result<Range<usize>> {
    let tag = xml
        .get(event_start..event_end)
        .ok_or_else(|| invalid("formatcode2 attribute source range is invalid"))?;
    let mut cursor = 1usize;
    while cursor < tag.len()
        && !is_xml_whitespace_byte(tag[cursor])
        && !matches!(tag[cursor], b'>' | b'/')
    {
        cursor += 1;
    }
    while cursor < tag.len() {
        while cursor < tag.len() && is_xml_whitespace_byte(tag[cursor]) {
            cursor += 1;
        }
        if cursor >= tag.len() || matches!(tag[cursor], b'>' | b'/') {
            break;
        }
        let key_start = cursor;
        while cursor < tag.len()
            && !is_xml_whitespace_byte(tag[cursor])
            && !matches!(tag[cursor], b'=' | b'>' | b'/')
        {
            cursor += 1;
        }
        let key_end = cursor;
        while cursor < tag.len() && is_xml_whitespace_byte(tag[cursor]) {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            return Err(invalid("formatcode2 attribute source range is invalid"));
        }
        cursor += 1;
        while cursor < tag.len() && is_xml_whitespace_byte(tag[cursor]) {
            cursor += 1;
        }
        let quote = match tag.get(cursor).copied() {
            Some(quote @ (b'\'' | b'"')) => quote,
            _ => return Err(invalid("formatcode2 attribute source range is invalid")),
        };
        cursor += 1;
        let value_start = cursor;
        while cursor < tag.len() && tag[cursor] != quote {
            cursor += 1;
        }
        let value_end = cursor;
        if tag.get(value_start..value_end) == Some(raw) && tag.get(key_start..key_end) == Some(key)
        {
            let start = event_start
                .checked_add(value_start)
                .ok_or_else(|| invalid("formatcode2 attribute source range overflows"))?;
            let end = event_start
                .checked_add(value_end)
                .ok_or_else(|| invalid("formatcode2 attribute source range overflows"))?;
            return Ok(start..end);
        }
        if cursor < tag.len() {
            cursor += 1;
        }
    }
    Err(invalid("formatcode2 attribute source range is absent"))
}

fn is_xml_whitespace_byte(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\r' | b'\n')
}

/// Serialize a `formatcode2` element.
///
/// An unchanged parsed value returns the exact source. A changed parsed value
/// patches only text-bearing event ranges and retains comments, processing
/// instructions, namespace spelling, and surrounding lexical bytes.
pub fn write(value: &Element, limits: &Limits) -> Result<Vec<u8>> {
    let limits = limits.validate()?;
    validate_value_with(value, limits.max_value_bytes)?;
    if let Some(source) = value.source_state() {
        if source.value == value.value {
            return clone_bounded_with(source.xml.as_ref(), limits.max_output_bytes);
        }
        return rewrite_source(source, value.value.as_str(), limits.max_output_bytes);
    }
    write_detached(value.value.as_str(), limits.max_output_bytes)
}

/// Serialize a `formatcode2` element to a caller-provided sink.
pub fn write_to<W: std::io::Write>(writer: &mut W, value: &Element, limits: &Limits) -> Result<()> {
    let limits = limits.validate()?;
    validate_value_with(value, limits.max_value_bytes)?;
    if let Some(source) = value.source_state()
        && source.value == value.value
    {
        ensure_output_limit_with(source.xml.len(), limits.max_output_bytes)?;
        writer.write_all(source.xml.as_ref())?;
        return Ok(());
    }
    write_to_bounded(writer, value, limits.max_output_bytes)
}

fn checked_value(value: &str) -> Result<Arc<str>> {
    if value.len() > MAX_VALUE_BYTES {
        return Err(limit(LimitKind::TextBytes, value.len(), MAX_VALUE_BYTES));
    }
    Ok(Arc::from(value))
}

fn validate_value_with(value: &Element, maximum: usize) -> Result<()> {
    if value.value.as_str().len() > maximum {
        return Err(limit(
            LimitKind::TextBytes,
            value.value.as_str().len(),
            maximum,
        ));
    }
    Ok(())
}

fn write_detached(value: &str, max_output_bytes: usize) -> Result<Vec<u8>> {
    let content_len = escaped_xstring_xml_len(value, XStringContext::Element)?;
    let start = b"<c16r2:formatcode2 xmlns:c16r2=\"";
    let output_len = start
        .len()
        .checked_add(NAMESPACE.len())
        .and_then(|length| length.checked_add(2))
        .and_then(|length| length.checked_add(content_len))
        .and_then(|length| length.checked_add(b"</c16r2:formatcode2>".len()))
        .ok_or_else(|| invalid("formatcode2 XML length overflows"))?;
    ensure_output_limit_with(output_len, max_output_bytes)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| invalid("formatcode2 output allocation failed"))?;
    let content = encode_xstring_xml(value, XStringContext::Element)?;
    output.extend_from_slice(start);
    output.extend_from_slice(NAMESPACE.as_bytes());
    output.extend_from_slice(b"\">");
    output.extend_from_slice(content.as_bytes());
    output.extend_from_slice(b"</c16r2:formatcode2>");
    Ok(output)
}

fn clone_bounded(xml: &[u8]) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(xml.len())
        .map_err(|_| invalid("formatcode2 output allocation failed"))?;
    output.extend_from_slice(xml);
    Ok(output)
}

fn clone_bounded_with(xml: &[u8], maximum: usize) -> Result<Vec<u8>> {
    if xml.len() > maximum {
        return Err(limit(output_limit_kind(maximum), xml.len(), maximum));
    }
    clone_bounded(xml)
}

fn rewrite_source(source: &Source, value: &str, max_output_bytes: usize) -> Result<Vec<u8>> {
    let mut replacements = Vec::<(Range<usize>, Vec<u8>)>::new();
    let mut writable = None;
    for segment in &source.segments {
        if matches!(
            segment.kind,
            SegmentKind::Text | SegmentKind::CData | SegmentKind::GeneralRef
        ) && writable.is_none()
        {
            writable = Some(segment.clone());
        }
    }
    if let Some(segment) = writable {
        let replacement_len = if segment.kind == SegmentKind::CData && !value.contains("]]>") {
            12usize
                .checked_add(encoded_xstring_len(value, XStringContext::Element)?)
                .ok_or_else(|| invalid("formatcode2 XML length overflows"))?
        } else {
            escaped_xstring_xml_len(value, XStringContext::Element)?
        };
        let mut output_len = source.xml.len();
        output_len = output_len
            .checked_sub(segment.range.end - segment.range.start)
            .and_then(|length| length.checked_add(replacement_len))
            .ok_or_else(|| invalid("formatcode2 XML length overflows"))?;
        for other in &source.segments {
            if other.range != segment.range {
                output_len = output_len
                    .checked_sub(other.range.end - other.range.start)
                    .ok_or_else(|| invalid("formatcode2 XML length underflows"))?;
            }
        }
        ensure_output_limit_with(output_len, max_output_bytes)?;
        let replacement = if segment.kind == SegmentKind::CData && !value.contains("]]>") {
            let encoded = encode_xstring(value, XStringContext::Element)?;
            let mut bytes = Vec::new();
            bytes.extend_from_slice(b"<![CDATA[");
            bytes.extend_from_slice(encoded.as_bytes());
            bytes.extend_from_slice(b"]]>");
            bytes
        } else {
            encode_xstring_xml(value, XStringContext::Element)?.into_bytes()
        };
        replacements.push((segment.range, replacement));
        for segment in &source.segments {
            if segment.range != replacements[0].0 {
                replacements.push((segment.range.clone(), Vec::new()));
            }
        }
    } else if let (Some(start), Some(end)) = (source.content_start, source.content_end) {
        let escaped_len = escaped_xstring_xml_len(value, XStringContext::Element)?;
        let output_len = source
            .xml
            .len()
            .checked_add(escaped_len)
            .ok_or_else(|| invalid("formatcode2 XML length overflows"))?;
        ensure_output_limit_with(output_len, max_output_bytes)?;
        let encoded = encode_xstring_xml(value, XStringContext::Element)?;
        replacements.push((start..start, encoded.into_bytes()));
        let _ = end;
    } else {
        return rewrite_empty_root(source, value, max_output_bytes);
    }
    apply_replacements(source.xml.as_ref(), replacements, max_output_bytes)
}

fn write_to_bounded<W: std::io::Write>(
    writer: &mut W,
    value: &Element,
    max_output_bytes: usize,
) -> Result<()> {
    let Some(source) = value.source_state() else {
        return write_detached_to(writer, value.value.as_str(), max_output_bytes);
    };
    write_source_to(writer, source, value.value.as_str(), max_output_bytes)
}

fn write_source_to<W: std::io::Write>(
    writer: &mut W,
    source: &Source,
    value: &str,
    max_output_bytes: usize,
) -> Result<()> {
    let mut writable = None;
    for segment in &source.segments {
        if matches!(
            segment.kind,
            SegmentKind::Text | SegmentKind::CData | SegmentKind::GeneralRef
        ) && writable.is_none()
        {
            writable = Some(segment.clone());
        }
    }
    if let Some(segment) = writable {
        let segment_range = segment.range.clone();
        let replacement_len = if segment.kind == SegmentKind::CData && !value.contains("]]>") {
            12usize
                .checked_add(encoded_xstring_len(value, XStringContext::Element)?)
                .ok_or_else(|| invalid("formatcode2 XML length overflows"))?
        } else {
            escaped_xstring_xml_len(value, XStringContext::Element)?
        };
        let mut output_len = source.xml.len();
        output_len = output_len
            .checked_sub(segment.range.end - segment.range.start)
            .and_then(|length| length.checked_add(replacement_len))
            .ok_or_else(|| invalid("formatcode2 XML length underflows"))?;
        for segment in &source.segments {
            if segment.range != segment_range {
                output_len = output_len
                    .checked_sub(segment.range.end - segment.range.start)
                    .ok_or_else(|| invalid("formatcode2 XML length underflows"))?;
            }
        }
        ensure_output_limit_with(output_len, max_output_bytes)?;
        let mut replacements = Vec::with_capacity(source.segments.len());
        let replacement = if segment.kind == SegmentKind::CData && !value.contains("]]>") {
            let encoded = encode_xstring(value, XStringContext::Element)?;
            let mut bytes = Vec::new();
            bytes.extend_from_slice(b"<![CDATA[");
            bytes.extend_from_slice(encoded.as_bytes());
            bytes.extend_from_slice(b"]]>");
            bytes
        } else {
            encode_xstring_xml(value, XStringContext::Element)?.into_bytes()
        };
        replacements.push((segment.range, replacement));
        for segment in &source.segments {
            if segment.range != replacements[0].0 {
                replacements.push((segment.range.clone(), Vec::new()));
            }
        }
        return apply_replacements_to(writer, source.xml.as_ref(), replacements, max_output_bytes);
    }
    if let (Some(start), Some(end)) = (source.content_start, source.content_end) {
        let escaped_len = escaped_xstring_xml_len(value, XStringContext::Element)?;
        let output_len = source
            .xml
            .len()
            .checked_add(escaped_len)
            .ok_or_else(|| invalid("formatcode2 XML length overflows"))?;
        ensure_output_limit_with(output_len, max_output_bytes)?;
        let encoded = encode_xstring_xml(value, XStringContext::Element)?;
        writer.write_all(&source.xml[..start])?;
        writer.write_all(encoded.as_bytes())?;
        writer.write_all(&source.xml[start..])?;
        let _ = end;
        return Ok(());
    }
    write_empty_root_to(writer, source, value, max_output_bytes)
}

fn write_detached_to<W: std::io::Write>(
    writer: &mut W,
    value: &str,
    max_output_bytes: usize,
) -> Result<()> {
    let content_len = escaped_xstring_xml_len(value, XStringContext::Element)?;
    let start = b"<c16r2:formatcode2 xmlns:c16r2=\"";
    let output_len = start
        .len()
        .checked_add(NAMESPACE.len())
        .and_then(|length| length.checked_add(2))
        .and_then(|length| length.checked_add(content_len))
        .and_then(|length| length.checked_add(b"</c16r2:formatcode2>".len()))
        .ok_or_else(|| invalid("formatcode2 XML length overflows"))?;
    ensure_output_limit_with(output_len, max_output_bytes)?;
    let content = encode_xstring_xml(value, XStringContext::Element)?;
    writer.write_all(start)?;
    writer.write_all(NAMESPACE.as_bytes())?;
    writer.write_all(b"\">")?;
    writer.write_all(content.as_bytes())?;
    writer.write_all(b"</c16r2:formatcode2>")?;
    Ok(())
}

fn rewrite_empty_root(source: &Source, value: &str, max_output_bytes: usize) -> Result<Vec<u8>> {
    let raw = &source.xml[source.root_start..source.root_end];
    let Some(slash) = raw.iter().rposition(|byte| *byte == b'/') else {
        return Err(invalid("formatcode2 empty root has no close marker"));
    };
    if raw.get(slash + 1) != Some(&b'>') {
        return Err(invalid("formatcode2 empty root has no close marker"));
    }
    let escaped_len = escaped_xstring_xml_len(value, XStringContext::Element)?;
    let replacement_len = raw[..slash]
        .len()
        .checked_add(1)
        .and_then(|length| length.checked_add(escaped_len))
        .and_then(|length| length.checked_add(2 + source.root_qname.len() + 1))
        .ok_or_else(|| invalid("formatcode2 XML length overflows"))?;
    let output_len = source
        .xml
        .len()
        .checked_sub(raw.len())
        .and_then(|length| length.checked_add(replacement_len))
        .ok_or_else(|| invalid("formatcode2 XML length overflows"))?;
    ensure_output_limit_with(output_len, max_output_bytes)?;
    let mut replacement = Vec::new();
    replacement
        .try_reserve_exact(replacement_len)
        .map_err(|_| invalid("formatcode2 output allocation failed"))?;
    replacement.extend_from_slice(&raw[..slash]);
    replacement.push(b'>');
    let encoded = encode_xstring_xml(value, XStringContext::Element)?;
    replacement.extend_from_slice(encoded.as_bytes());
    replacement.extend_from_slice(b"</");
    replacement.extend_from_slice(&source.root_qname);
    replacement.push(b'>');
    apply_replacements(
        source.xml.as_ref(),
        vec![(source.root_start..source.root_end, replacement)],
        max_output_bytes,
    )
}

fn write_empty_root_to<W: std::io::Write>(
    writer: &mut W,
    source: &Source,
    value: &str,
    max_output_bytes: usize,
) -> Result<()> {
    let raw = &source.xml[source.root_start..source.root_end];
    let Some(slash) = raw.iter().rposition(|byte| *byte == b'/') else {
        return Err(invalid("formatcode2 empty root has no close marker"));
    };
    if raw.get(slash + 1) != Some(&b'>') {
        return Err(invalid("formatcode2 empty root has no close marker"));
    }
    let escaped_len = escaped_xstring_xml_len(value, XStringContext::Element)?;
    let replacement_len = raw[..slash]
        .len()
        .checked_add(1)
        .and_then(|length| length.checked_add(escaped_len))
        .and_then(|length| length.checked_add(2 + source.root_qname.len() + 1))
        .ok_or_else(|| invalid("formatcode2 XML length overflows"))?;
    let output_len = source
        .xml
        .len()
        .checked_sub(raw.len())
        .and_then(|length| length.checked_add(replacement_len))
        .ok_or_else(|| invalid("formatcode2 XML length overflows"))?;
    ensure_output_limit_with(output_len, max_output_bytes)?;
    let encoded = encode_xstring_xml(value, XStringContext::Element)?;
    writer.write_all(&source.xml[..source.root_start])?;
    writer.write_all(&raw[..slash])?;
    writer.write_all(b">")?;
    writer.write_all(encoded.as_bytes())?;
    writer.write_all(b"</")?;
    writer.write_all(&source.root_qname)?;
    writer.write_all(b">")?;
    writer.write_all(&source.xml[source.root_end..])?;
    Ok(())
}

fn apply_replacements(
    xml: &[u8],
    mut replacements: Vec<(Range<usize>, Vec<u8>)>,
    max_output_bytes: usize,
) -> Result<Vec<u8>> {
    replacements.sort_by_key(|(range, _)| range.start);
    for pair in replacements.windows(2) {
        if pair[0].0.end > pair[1].0.start {
            return Err(invalid("formatcode2 source ranges overlap"));
        }
    }
    let mut output_len = xml.len();
    for (range, bytes) in &replacements {
        output_len = output_len
            .checked_sub(range.end - range.start)
            .and_then(|length| length.checked_add(bytes.len()))
            .ok_or_else(|| invalid("formatcode2 XML length overflows"))?;
    }
    ensure_output_limit_with(output_len, max_output_bytes)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| invalid("formatcode2 output allocation failed"))?;
    let mut cursor = 0;
    for (range, bytes) in replacements {
        output.extend_from_slice(&xml[cursor..range.start]);
        output.extend_from_slice(&bytes);
        cursor = range.end;
    }
    output.extend_from_slice(&xml[cursor..]);
    Ok(output)
}

fn apply_replacements_to<W: std::io::Write>(
    writer: &mut W,
    xml: &[u8],
    mut replacements: Vec<(Range<usize>, Vec<u8>)>,
    max_output_bytes: usize,
) -> Result<()> {
    replacements.sort_by_key(|(range, _)| range.start);
    for pair in replacements.windows(2) {
        if pair[0].0.end > pair[1].0.start {
            return Err(invalid("formatcode2 source ranges overlap"));
        }
    }
    let mut output_len = xml.len();
    for (range, bytes) in &replacements {
        output_len = output_len
            .checked_sub(range.end - range.start)
            .and_then(|length| length.checked_add(bytes.len()))
            .ok_or_else(|| invalid("formatcode2 XML length overflows"))?;
    }
    ensure_output_limit_with(output_len, max_output_bytes)?;
    let mut cursor = 0;
    for (range, bytes) in replacements {
        writer.write_all(&xml[cursor..range.start])?;
        writer.write_all(&bytes)?;
        cursor = range.end;
    }
    writer.write_all(&xml[cursor..])?;
    Ok(())
}

fn scan(xml: &[u8], limits: Limits) -> Result<Parsed> {
    preflight_namespace_limits(xml, limits)?;
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(limits.max_attributes);
    let mut stack = Vec::<Vec<u8>>::new();
    let mut root_seen = false;
    let mut root_closed = false;
    let mut root_start = None;
    let mut root_end = None;
    let mut root_qname = None;
    let mut content_start = None;
    let mut content_end = None;
    let mut segments = Vec::new();
    let mut lexical_value = String::new();
    let max_lexical_value_bytes = limits.max_lexical_value_bytes()?;
    let mut nodes = 0usize;
    let mut declaration_seen = false;
    let mut pre_root_event_seen = false;
    let mut root_empty = false;

    loop {
        let event_start = position(&reader, origin)?;
        let (resolved, event) = reader.read_resolved_event().map_err(xml_event_error)?;
        let namespace = resolved_namespace(&resolved, limits)?;
        if let Event::Start(element) | Event::Empty(element) = &event {
            preflight_attribute_value_event(element, max_lexical_value_bytes)?;
        }
        let event = event.into_owned();
        let event_end = position(&reader, origin)?;
        if !root_seen && !matches!(&event, Event::Decl(_) | Event::Eof) {
            pre_root_event_seen = true;
        }
        match event {
            Event::Decl(declaration) => {
                if root_seen || declaration_seen || pre_root_event_seen {
                    return Err(invalid("formatcode2 XML declaration is misplaced"));
                }
                validate_declaration(&declaration)?;
                declaration_seen = true;
            },
            Event::Start(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("formatcode2 node count overflow"))?;
                if nodes > limits.max_nodes {
                    return Err(limit(LimitKind::XmlNodes, nodes, limits.max_nodes));
                }
                if stack.len() >= limits.max_depth {
                    return Err(limit(
                        LimitKind::XmlDepth,
                        stack.len() + 1,
                        limits.max_depth,
                    ));
                }
                validate_element_name(&element, limits.max_namespace_bytes)?;
                validate_attributes(&element, &reader, limits)?;
                if root_closed {
                    return Err(invalid(
                        "formatcode2 fragment has more than one root element",
                    ));
                }
                let local_name = element.local_name();
                let local = local_name.as_ref();
                if !root_seen {
                    if local != b"formatcode2" || namespace != NAMESPACE.as_bytes() {
                        return Err(invalid("formatcode2 fragment has the wrong root"));
                    }
                    if !only_namespace_attributes(&element)? {
                        return Err(invalid("formatcode2 root has unexpected attributes"));
                    }
                    root_seen = true;
                    root_start = Some(event_start);
                    root_qname = Some(element.name().as_ref().to_vec());
                    content_start = Some(event_end);
                } else if stack.len() == 1 {
                    return Err(invalid("formatcode2 value contains a child element"));
                }
                stack.push(element.name().as_ref().to_vec());
            },
            Event::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("formatcode2 node count overflow"))?;
                if nodes > limits.max_nodes {
                    return Err(limit(LimitKind::XmlNodes, nodes, limits.max_nodes));
                }
                validate_element_name(&element, limits.max_namespace_bytes)?;
                validate_attributes(&element, &reader, limits)?;
                if root_closed {
                    return Err(invalid(
                        "formatcode2 fragment has more than one root element",
                    ));
                }
                if !root_seen {
                    if element.local_name().as_ref() != b"formatcode2"
                        || namespace != NAMESPACE.as_bytes()
                    {
                        return Err(invalid("formatcode2 fragment has the wrong root"));
                    }
                    if !only_namespace_attributes(&element)? {
                        return Err(invalid("formatcode2 root has unexpected attributes"));
                    }
                    root_seen = true;
                    root_closed = true;
                    root_empty = true;
                    root_start = Some(event_start);
                    root_end = Some(event_end);
                    root_qname = Some(element.name().as_ref().to_vec());
                } else {
                    return Err(invalid("formatcode2 value contains a child element"));
                }
            },
            Event::End(end) => {
                validate_end_name(&end, limits.max_namespace_bytes)?;
                if stack.pop().is_none() {
                    return Err(invalid(
                        "formatcode2 fragment has an unmatched closing element",
                    ));
                }
                if stack.is_empty() {
                    root_closed = true;
                    root_end = Some(event_end);
                    content_end = Some(event_start);
                }
            },
            Event::Text(text) => {
                if !root_seen || stack.is_empty() {
                    if !is_xml_whitespace(text.as_ref())? {
                        return Err(invalid(
                            "formatcode2 has non-whitespace text outside its root",
                        ));
                    }
                } else {
                    validate_raw_text(text.as_ref())?;
                    preflight_unescaped_value(text.as_ref(), max_lexical_value_bytes)?;
                    let decoded = text.decode().map_err(xml_error)?;
                    let decoded = unescape(decoded.as_ref()).map_err(xml_error)?;
                    validate_xml_text(decoded.as_ref(), "formatcode2 text")?;
                    append_lexical(
                        &mut lexical_value,
                        decoded.as_ref(),
                        max_lexical_value_bytes,
                    )?;
                    segments.push(Segment {
                        range: event_start..event_end,
                        kind: SegmentKind::Text,
                    });
                }
            },
            Event::CData(data) => {
                if !root_seen || stack.is_empty() {
                    return Err(invalid("formatcode2 contains data outside its root"));
                }
                preflight_lexical_bytes(data.as_ref(), max_lexical_value_bytes)?;
                let decoded = data.decode().map_err(xml_error)?;
                validate_xml_text(decoded.as_ref(), "formatcode2 CDATA")?;
                append_lexical(
                    &mut lexical_value,
                    decoded.as_ref(),
                    max_lexical_value_bytes,
                )?;
                segments.push(Segment {
                    range: event_start..event_end,
                    kind: SegmentKind::CData,
                });
            },
            Event::GeneralRef(reference) => {
                if !root_seen || stack.is_empty() {
                    return Err(invalid("formatcode2 contains a reference outside its root"));
                }
                let character = validate_general_ref(&reference, limits.max_value_bytes)?;
                if let Some(character) = character {
                    append_lexical_char(&mut lexical_value, character, max_lexical_value_bytes)?;
                }
                segments.push(Segment {
                    range: event_start..event_end,
                    kind: SegmentKind::GeneralRef,
                });
            },
            Event::Comment(comment) => validate_comment(&comment)?,
            Event::PI(instruction) => validate_processing_instruction(&instruction)?,
            Event::DocType(_) => return Err(invalid("formatcode2 cannot contain a document type")),
            Event::Eof => break,
        }
    }
    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("formatcode2 fragment has no complete root"));
    }
    if root_empty {
        content_start = None;
        content_end = None;
    }
    let value = decode_xstring(&lexical_value, limits.max_value_bytes)?;
    Ok(Parsed {
        root_start: root_start.ok_or_else(|| invalid("formatcode2 root start is missing"))?,
        root_end: root_end.ok_or_else(|| invalid("formatcode2 root end is missing"))?,
        root_qname: root_qname.ok_or_else(|| invalid("formatcode2 root name is missing"))?,
        content_start,
        content_end,
        segments,
        value,
    })
}

fn only_namespace_attributes(element: &BytesStart<'_>) -> Result<bool> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if !is_namespace_attribute(attribute.key) {
            return Ok(false);
        }
    }
    Ok(true)
}

fn validate_element_name(element: &BytesStart<'_>, max_namespace_bytes: usize) -> Result<()> {
    let qualified_name = element.name();
    let name = std::str::from_utf8(qualified_name.as_ref()).map_err(xml_error)?;
    if !is_qualified_name(name) {
        return Err(invalid("formatcode2 element name is invalid"));
    }
    validate_qname_prefix(qualified_name.as_ref(), max_namespace_bytes)?;
    Ok(())
}

fn validate_end_name(element: &BytesEnd<'_>, max_namespace_bytes: usize) -> Result<()> {
    let qualified_name = element.name();
    let name = std::str::from_utf8(qualified_name.as_ref()).map_err(xml_error)?;
    if !is_qualified_name(name) {
        return Err(invalid("formatcode2 closing element name is invalid"));
    }
    validate_qname_prefix(qualified_name.as_ref(), max_namespace_bytes)?;
    Ok(())
}

fn validate_qname_prefix(name: &[u8], max_namespace_bytes: usize) -> Result<()> {
    if let Some(colon) = name.iter().position(|byte| *byte == b':')
        && colon > max_namespace_bytes
    {
        return Err(limit(
            LimitKind::NamespacePrefixBytes,
            colon,
            max_namespace_bytes,
        ));
    }
    Ok(())
}

fn validate_attributes(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    limits: Limits,
) -> Result<()> {
    let mut count = 0usize;
    let mut seen = HashSet::<&[u8]>::new();
    let mut expanded = HashSet::<(Vec<u8>, Vec<u8>)>::new();
    for attribute in element.checked_attributes() {
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("formatcode2 attribute count overflow"))?;
        if count > limits.max_attributes {
            return Err(limit(
                LimitKind::XmlAttributes,
                count,
                limits.max_attributes,
            ));
        }
        let attribute = attribute.map_err(xml_error)?;
        let key_bytes = attribute.key.into_inner();
        let key = std::str::from_utf8(key_bytes).map_err(xml_error)?;
        if !is_qualified_name(key) && !is_namespace_attribute(attribute.key) {
            return Err(invalid("formatcode2 attribute name is invalid"));
        }
        // Validate the prefix before retaining the key in any duplicate set.
        // This keeps a hostile, oversized QName from becoming an unbounded
        // owned allocation merely to report a namespace error.
        validate_qname_prefix(key_bytes, limits.max_namespace_bytes)?;
        if !seen.insert(key_bytes) {
            return Err(invalid("formatcode2 element has duplicate attributes"));
        }
        validate_raw_attribute_value(attribute.value.as_ref())?;
        if attribute.key.local_name().as_ref() == b"formatcode2" {
            preflight_normalized_value(
                attribute.value.as_ref(),
                limits.max_lexical_value_bytes()?,
            )?;
        }
        let value = decode_attribute_value(attribute.value.as_ref(), reader)?;
        validate_xml_text(value.as_str(), "formatcode2 attribute value")?;
        if is_namespace_attribute(attribute.key) {
            let prefix = namespace_attribute_prefix(attribute.key)?;
            if prefix.len() > limits.max_namespace_bytes {
                return Err(limit(
                    LimitKind::NamespacePrefixBytes,
                    prefix.len(),
                    limits.max_namespace_bytes,
                ));
            }
            if value.len() > limits.max_namespace_bytes {
                return Err(limit(
                    LimitKind::NamespaceBytes,
                    value.len(),
                    limits.max_namespace_bytes,
                ));
            }
            validate_namespace_binding(prefix, value.as_bytes())?;
        } else {
            let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
            let namespace = match resolved {
                ResolveResult::Bound(namespace) => {
                    normalize_namespace_uri(namespace.into_inner(), limits)?
                },
                ResolveResult::Unbound => Vec::new(),
                ResolveResult::Unknown(prefix) => {
                    let prefix_bytes: &[u8] = prefix.as_ref();
                    if prefix_bytes.len() > limits.max_namespace_bytes {
                        return Err(limit(
                            LimitKind::NamespacePrefixBytes,
                            prefix_bytes.len(),
                            limits.max_namespace_bytes,
                        ));
                    }
                    return Err(invalid("formatcode2 attribute has an undeclared prefix"));
                },
            };
            if namespace.len() > limits.max_namespace_bytes {
                return Err(limit(
                    LimitKind::NamespaceBytes,
                    namespace.len(),
                    limits.max_namespace_bytes,
                ));
            }
            if !expanded.insert((namespace, local.into_inner().to_vec())) {
                return Err(invalid(
                    "formatcode2 element has duplicate expanded attributes",
                ));
            }
        }
    }
    Ok(())
}

fn decode_attribute_value(raw: &[u8], reader: &NsReader<&[u8]>) -> Result<String> {
    let decoded = reader.decoder().decode(raw).map_err(xml_error)?;
    unescape_with(decoded.as_ref(), resolve_xml_entity)
        .map(|value| value.into_owned())
        .map_err(xml_error)
}

/// Count the XML-normalized attribute value without materializing it. This is
/// the same XML 1.0 normalization used by `quick_xml` for
/// `decoded_and_normalized_value`: ordinary characters retain their UTF-8
/// width, whitespace becomes one ASCII space, and character/predefined
/// references contribute their decoded scalar width. The count is performed
/// before the projection allocates its normalized lexical string.
fn preflight_normalized_value(raw: &[u8], maximum: usize) -> Result<()> {
    let value = std::str::from_utf8(raw).map_err(xml_error)?;
    let bytes = value.as_bytes();
    let mut length = 0usize;
    let mut index = 0usize;
    while index < bytes.len() {
        let addition = match bytes[index] {
            b'&' => {
                let end = bytes
                    .get(index + 1..)
                    .ok_or_else(|| {
                        invalid("formatcode2 attribute contains an unterminated entity")
                    })?
                    .iter()
                    .position(|byte| *byte == b';')
                    .map(|offset| index + 1 + offset)
                    .ok_or_else(|| {
                        invalid("formatcode2 attribute contains an unterminated entity")
                    })?;
                let entity = &value[index + 1..end];
                index = end + 1;
                decoded_entity_len(entity)?
            },
            b'\r' => {
                index += 1;
                if bytes.get(index) == Some(&b'\n') {
                    index += 1;
                }
                1
            },
            b'\n' | b'\t' => {
                index += 1;
                1
            },
            _ => {
                let character = value[index..]
                    .chars()
                    .next()
                    .ok_or_else(|| invalid("formatcode2 attribute contains invalid UTF-8"))?;
                index += character.len_utf8();
                character.len_utf8()
            },
        };
        length = length
            .checked_add(addition)
            .ok_or_else(|| invalid("formatcode2 lexical value length overflows"))?;
        if length > maximum {
            return Err(limit(LimitKind::LexicalValueBytes, length, maximum));
        }
    }
    Ok(())
}

fn preflight_attribute_value_budget(xml: &[u8], limits: Limits) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    let event = reader.read_event().map_err(xml_event_error)?;
    let element = match event {
        Event::Start(element) | Event::Empty(element) => element,
        Event::Eof => return Err(invalid("formatcode2 attribute fragment has no start tag")),
        _ => {
            return Err(invalid(
                "formatcode2 attribute fragment must begin with a start tag",
            ));
        },
    };
    preflight_attribute_value_event(&element, limits.max_lexical_value_bytes()?)
}

fn preflight_attribute_value_event(element: &BytesStart<'_>, maximum: usize) -> Result<()> {
    for raw_attribute in element.checked_attributes() {
        let attribute = raw_attribute.map_err(xml_error)?;
        if attribute.key.local_name().as_ref() == b"formatcode2" {
            preflight_normalized_value(attribute.value.as_ref(), maximum)?;
        }
    }
    Ok(())
}

fn preflight_unescaped_value(raw: &[u8], maximum: usize) -> Result<()> {
    let value = std::str::from_utf8(raw).map_err(xml_error)?;
    let bytes = value.as_bytes();
    let mut length = 0usize;
    let mut index = 0usize;
    while index < bytes.len() {
        let addition = if bytes[index] == b'&' {
            let end = bytes
                .get(index + 1..)
                .ok_or_else(|| invalid("formatcode2 text contains an unterminated entity"))?
                .iter()
                .position(|byte| *byte == b';')
                .map(|offset| index + 1 + offset)
                .ok_or_else(|| invalid("formatcode2 text contains an unterminated entity"))?;
            let entity = &value[index + 1..end];
            index = end + 1;
            decoded_entity_len(entity)?
        } else {
            let character = value[index..]
                .chars()
                .next()
                .ok_or_else(|| invalid("formatcode2 text contains invalid UTF-8"))?;
            index += character.len_utf8();
            character.len_utf8()
        };
        length = length
            .checked_add(addition)
            .ok_or_else(|| invalid("formatcode2 lexical value length overflows"))?;
        if length > maximum {
            return Err(limit(LimitKind::LexicalValueBytes, length, maximum));
        }
    }
    Ok(())
}

fn preflight_lexical_bytes(raw: &[u8], maximum: usize) -> Result<()> {
    let length = std::str::from_utf8(raw).map_err(xml_error)?.len();
    if length > maximum {
        return Err(limit(LimitKind::LexicalValueBytes, length, maximum));
    }
    Ok(())
}

fn decoded_entity_len(entity: &str) -> Result<usize> {
    if let Some(value) = resolve_xml_entity(entity) {
        return Ok(value.len());
    }
    let number = entity
        .strip_prefix("#x")
        .or_else(|| entity.strip_prefix("#X"))
        .map(|digits| u32::from_str_radix(digits, 16))
        .or_else(|| entity.strip_prefix('#').map(|digits| digits.parse::<u32>()))
        .transpose()
        .map_err(|error| {
            invalid(format!(
                "formatcode2 character reference is invalid: {error}"
            ))
        })?;
    let Some(number) = number else {
        return Err(invalid("formatcode2 attribute contains an unknown entity"));
    };
    let character = char::from_u32(number)
        .ok_or_else(|| invalid("formatcode2 character reference is invalid"))?;
    validate_xml_text(
        &character.to_string(),
        "formatcode2 attribute character reference",
    )?;
    Ok(character.len_utf8())
}

fn namespace_attribute_prefix<'a>(key: quick_xml::name::QName<'a>) -> Result<&'a [u8]> {
    let key = key.into_inner();
    if key == b"xmlns" {
        return Ok(&[]);
    }
    key.strip_prefix(b"xmlns:")
        .filter(|prefix| !prefix.is_empty())
        .ok_or_else(|| invalid("formatcode2 namespace declaration is invalid"))
}

fn validate_namespace_binding(prefix: &[u8], value: &[u8]) -> Result<()> {
    if !prefix.is_empty() && !is_ncname(std::str::from_utf8(prefix).map_err(xml_error)?) {
        return Err(invalid(
            "formatcode2 namespace declaration prefix is invalid",
        ));
    }
    if value == XMLNS_NAMESPACE {
        return Err(invalid(
            "formatcode2 declaration binds the reserved XMLNS namespace",
        ));
    }
    if value == XML_NAMESPACE && prefix != b"xml" {
        return Err(invalid(
            "formatcode2 declaration binds the XML namespace to a non-xml prefix",
        ));
    }
    if prefix == b"xml" && value != XML_NAMESPACE {
        return Err(invalid("formatcode2 xml prefix has the wrong namespace"));
    }
    if !prefix.is_empty() && value.is_empty() {
        return Err(invalid("formatcode2 prefixed namespace binding is empty"));
    }
    Ok(())
}

fn is_namespace_attribute(key: quick_xml::name::QName<'_>) -> bool {
    let key = key.as_ref();
    key == b"xmlns" || key.starts_with(b"xmlns:")
}

fn validate_raw_attribute_value(value: &[u8]) -> Result<()> {
    if value.contains(&b'<') {
        return Err(invalid(
            "formatcode2 attribute contains an unescaped '<' delimiter",
        ));
    }
    Ok(())
}

fn validate_raw_text(value: &[u8]) -> Result<()> {
    if value.contains(&b'<') {
        return Err(invalid(
            "formatcode2 text contains an unescaped '<' delimiter",
        ));
    }
    if value.windows(3).any(|window| window == b"]]>") {
        return Err(invalid(
            "formatcode2 text contains the forbidden ]]> delimiter",
        ));
    }
    Ok(())
}

fn validate_general_ref(reference: &BytesRef<'_>, max_value_bytes: usize) -> Result<Option<char>> {
    if reference.len() > max_value_bytes {
        return Err(limit(
            LimitKind::ScalarBytes,
            reference.len(),
            max_value_bytes,
        ));
    }
    if reference.is_char_ref() {
        let character = reference
            .resolve_char_ref()
            .map_err(xml_error)?
            .ok_or_else(|| invalid("formatcode2 character reference is invalid"))?;
        let mut encoded = [0u8; 4];
        validate_xml_text(character.encode_utf8(&mut encoded), "formatcode2 reference")?;
        return Ok(Some(character));
    }
    let entity = reference.decode().map_err(xml_error)?;
    let replacement = resolve_xml_entity(entity.as_ref())
        .ok_or_else(|| invalid("formatcode2 contains an unknown general entity"))?;
    let mut characters = replacement.chars();
    let character = characters.next();
    if characters.next().is_some() {
        return Err(invalid("formatcode2 entity is not scalar text"));
    }
    Ok(character)
}

fn validate_declaration(declaration: &BytesDecl<'_>) -> Result<()> {
    let content = declaration.as_ref();
    let content = std::str::from_utf8(content).map_err(xml_error)?;
    let start = BytesStart::from_content(content, 3);
    let mut seen_version = false;
    let mut seen_encoding = false;
    let mut seen_standalone = false;
    let mut phase = 0u8;
    for raw_attribute in start.checked_attributes() {
        let attribute = raw_attribute.map_err(xml_error)?;
        match attribute.key.as_ref() {
            b"version" => {
                if seen_version || phase != 0 {
                    return Err(invalid(
                        "formatcode2 XML declaration has invalid attribute order",
                    ));
                }
                seen_version = true;
                phase = 1;
                if attribute.value.as_ref() != b"1.0" {
                    return Err(invalid("formatcode2 XML declaration must use version 1.0"));
                }
            },
            b"encoding" => {
                if !seen_version || seen_encoding || phase > 2 {
                    return Err(invalid(
                        "formatcode2 XML declaration has invalid attribute order",
                    ));
                }
                seen_encoding = true;
                phase = 2;
                if !attribute.value.eq_ignore_ascii_case(b"utf-8") {
                    return Err(invalid(
                        "formatcode2 XML declaration must use UTF-8 encoding",
                    ));
                }
            },
            b"standalone" => {
                if !seen_version || seen_standalone || phase > 2 {
                    return Err(invalid(
                        "formatcode2 XML declaration has invalid attribute order",
                    ));
                }
                seen_standalone = true;
                phase = 3;
                if !matches!(attribute.value.as_ref(), b"yes" | b"no") {
                    return Err(invalid(
                        "formatcode2 XML declaration standalone must be yes or no",
                    ));
                }
            },
            _ => {
                return Err(invalid(
                    "formatcode2 XML declaration has an unknown attribute",
                ));
            },
        }
    }
    if !seen_version {
        return Err(invalid("formatcode2 XML declaration must declare version"));
    }
    Ok(())
}

fn resolved_namespace(resolved: &ResolveResult<'_>, limits: Limits) -> Result<Vec<u8>> {
    match resolved {
        ResolveResult::Bound(namespace) => normalize_namespace_uri(namespace.into_inner(), limits),
        ResolveResult::Unbound => Ok(Vec::new()),
        ResolveResult::Unknown(prefix) => {
            let prefix_bytes: &[u8] = prefix.as_ref();
            if prefix_bytes.len() > limits.max_namespace_bytes {
                return Err(limit(
                    LimitKind::NamespacePrefixBytes,
                    prefix_bytes.len(),
                    limits.max_namespace_bytes,
                ));
            }
            Err(invalid(format!(
                "formatcode2 element has an undeclared prefix '{}',",
                String::from_utf8_lossy(prefix_bytes)
            )))
        },
    }
}

fn normalize_namespace_uri(raw: &[u8], limits: Limits) -> Result<Vec<u8>> {
    if raw.len() > limits.max_namespace_bytes {
        return Err(limit(
            LimitKind::NamespaceBytes,
            raw.len(),
            limits.max_namespace_bytes,
        ));
    }
    let raw = std::str::from_utf8(raw).map_err(xml_error)?;
    let decoded = unescape_with(raw, resolve_xml_entity)
        .map_err(xml_error)?
        .into_owned();
    validate_xml_text(&decoded, "formatcode2 namespace URI")?;
    if decoded.len() > limits.max_namespace_bytes {
        return Err(limit(
            LimitKind::NamespaceBytes,
            decoded.len(),
            limits.max_namespace_bytes,
        ));
    }
    Ok(decoded.into_bytes())
}

fn append_lexical(value: &mut String, addition: &str, maximum: usize) -> Result<()> {
    let length = value
        .len()
        .checked_add(addition.len())
        .ok_or_else(|| invalid("formatcode2 lexical value length overflows"))?;
    if length > maximum {
        return Err(limit(LimitKind::LexicalValueBytes, length, maximum));
    }
    value.push_str(addition);
    Ok(())
}

fn append_lexical_char(value: &mut String, character: char, maximum: usize) -> Result<()> {
    let length = value
        .len()
        .checked_add(character.len_utf8())
        .ok_or_else(|| invalid("formatcode2 lexical value length overflows"))?;
    if length > maximum {
        return Err(limit(LimitKind::LexicalValueBytes, length, maximum));
    }
    value.push(character);
    Ok(())
}

fn decode_xstring(value: &str, maximum: usize) -> Result<Value> {
    let bytes = value.as_bytes();
    let mut decoded = String::new();
    decoded
        .try_reserve(value.len().min(maximum))
        .map_err(|_| invalid("formatcode2 value allocation failed"))?;
    let mut copied_until = 0usize;
    let mut index = 0usize;
    while index + 7 <= bytes.len() {
        let Some((unit, end)) = xstring_escape_at(bytes, index) else {
            index += 1;
            continue;
        };
        append_decoded(&mut decoded, &value[copied_until..index], maximum)?;
        if (0xD800..=0xDBFF).contains(&unit) {
            let Some((low, pair_end)) = xstring_escape_at(bytes, end) else {
                return Err(invalid(
                    "formatcode2 ST_Xstring has an unpaired high surrogate",
                ));
            };
            if !(0xDC00..=0xDFFF).contains(&low) {
                return Err(invalid(
                    "formatcode2 ST_Xstring has an unpaired high surrogate",
                ));
            }
            let scalar = 0x1_0000 + ((u32::from(unit) - 0xD800) << 10) + (u32::from(low) - 0xDC00);
            let character = char::from_u32(scalar)
                .ok_or_else(|| invalid("formatcode2 ST_Xstring surrogate pair is invalid"))?;
            append_decoded_char(&mut decoded, character, maximum)?;
            index = pair_end;
            copied_until = pair_end;
        } else if (0xDC00..=0xDFFF).contains(&unit) {
            return Err(invalid(
                "formatcode2 ST_Xstring has an unpaired low surrogate",
            ));
        } else {
            let character = char::from_u32(u32::from(unit))
                .ok_or_else(|| invalid("formatcode2 ST_Xstring code unit is invalid"))?;
            append_decoded_char(&mut decoded, character, maximum)?;
            index = end;
            copied_until = end;
        }
    }
    append_decoded(&mut decoded, &value[copied_until..], maximum)?;
    Ok(Value {
        value: Arc::from(decoded),
    })
}

fn append_decoded(value: &mut String, addition: &str, maximum: usize) -> Result<()> {
    let length = value
        .len()
        .checked_add(addition.len())
        .ok_or_else(|| invalid("formatcode2 value length overflows"))?;
    if length > maximum {
        return Err(limit(LimitKind::TextBytes, length, maximum));
    }
    value.push_str(addition);
    Ok(())
}

fn append_decoded_char(value: &mut String, character: char, maximum: usize) -> Result<()> {
    let length = value
        .len()
        .checked_add(character.len_utf8())
        .ok_or_else(|| invalid("formatcode2 value length overflows"))?;
    if length > maximum {
        return Err(limit(LimitKind::TextBytes, length, maximum));
    }
    value.push(character);
    Ok(())
}

#[derive(Debug, Clone, Copy)]
enum XStringContext {
    Element,
    Attribute,
}

fn encoded_xstring_len(value: &str, context: XStringContext) -> Result<usize> {
    value
        .char_indices()
        .try_fold(0usize, |length, (index, character)| {
            let addition = encoded_character_len(value.as_bytes(), index, character, context);
            length
                .checked_add(addition)
                .ok_or_else(|| invalid("formatcode2 encoded value length overflows"))
        })
}

fn escaped_xstring_xml_len(value: &str, context: XStringContext) -> Result<usize> {
    value
        .char_indices()
        .try_fold(0usize, |length, (index, character)| {
            let addition =
                if character == '_' && xstring_escape_at(value.as_bytes(), index).is_some() {
                    7
                } else if matches!(context, XStringContext::Attribute)
                    && matches!(character, '\t' | '\n' | '\r')
                {
                    5
                } else if matches!(character, '\u{fffe}' | '\u{ffff}')
                    || (character < '\u{20}' && !matches!(character, '\t' | '\n' | '\r'))
                    || matches!(character, '\r')
                {
                    7 * character.len_utf16()
                } else {
                    escaped_xml_character_len(character)
                };
            length
                .checked_add(addition)
                .ok_or_else(|| invalid("formatcode2 XML length overflows"))
        })
}

fn encoded_character_len(
    bytes: &[u8],
    index: usize,
    character: char,
    context: XStringContext,
) -> usize {
    if character == '_' && xstring_escape_at(bytes, index).is_some() {
        7
    } else if matches!(character, '\u{fffe}' | '\u{ffff}')
        || (character < '\u{20}' && !matches!(character, '\t' | '\n' | '\r'))
        || matches!(character, '\r')
        || (matches!(context, XStringContext::Attribute) && matches!(character, '\t' | '\n'))
    {
        7 * character.len_utf16()
    } else {
        character.len_utf8()
    }
}

fn escaped_xml_character_len(character: char) -> usize {
    match character {
        '&' => 5,
        '<' | '>' => 4,
        '"' | '\'' => 6,
        _ => character.len_utf8(),
    }
}

fn encode_xstring(value: &str, context: XStringContext) -> Result<String> {
    let encoded_len = encoded_xstring_len(value, context)?;
    let mut encoded = String::new();
    encoded
        .try_reserve_exact(encoded_len)
        .map_err(|_| invalid("formatcode2 encoded value allocation failed"))?;
    for (index, character) in value.char_indices() {
        if character == '_' && xstring_escape_at(value.as_bytes(), index).is_some() {
            encoded.push_str("_x005F_");
        } else if matches!(character, '\u{fffe}' | '\u{ffff}')
            || (character < '\u{20}' && !matches!(character, '\t' | '\n' | '\r'))
            || matches!(character, '\r')
            || (matches!(context, XStringContext::Attribute) && matches!(character, '\t' | '\n'))
        {
            let mut units = [0u16; 2];
            for unit in character.encode_utf16(&mut units) {
                push_xstring_escape(&mut encoded, *unit);
            }
        } else {
            encoded.push(character);
        }
    }
    Ok(encoded)
}

fn encode_xstring_xml(value: &str, context: XStringContext) -> Result<String> {
    let encoded_len = escaped_xstring_xml_len(value, context)?;
    let mut encoded = String::new();
    encoded
        .try_reserve_exact(encoded_len)
        .map_err(|_| invalid("formatcode2 encoded value allocation failed"))?;
    for (index, character) in value.char_indices() {
        if matches!(context, XStringContext::Attribute) && matches!(character, '\t' | '\n' | '\r') {
            match character {
                '\t' => encoded.push_str("&#x9;"),
                '\n' => encoded.push_str("&#xA;"),
                '\r' => encoded.push_str("&#xD;"),
                _ => unreachable!(),
            }
        } else if character == '_' && xstring_escape_at(value.as_bytes(), index).is_some() {
            encoded.push_str("_x005F_");
        } else if matches!(character, '\u{fffe}' | '\u{ffff}')
            || (character < '\u{20}' && !matches!(character, '\t' | '\n' | '\r'))
            || matches!(character, '\r')
        {
            let mut units = [0u16; 2];
            for unit in character.encode_utf16(&mut units) {
                push_xstring_escape(&mut encoded, *unit);
            }
        } else {
            encode_xstring_xml_character(&mut encoded, character);
        }
    }
    Ok(encoded)
}

fn encode_xstring_xml_character(output: &mut String, character: char) {
    match character {
        '&' => output.push_str("&amp;"),
        '<' => output.push_str("&lt;"),
        '>' => output.push_str("&gt;"),
        '"' => output.push_str("&quot;"),
        '\'' => output.push_str("&apos;"),
        _ => output.push(character),
    }
}

fn push_xstring_escape(output: &mut String, unit: u16) {
    const HEX: &[u8; 16] = b"0123456789ABCDEF";
    output.push_str("_x");
    for shift in [12, 8, 4, 0] {
        output.push(char::from(HEX[usize::from((unit >> shift) & 0xF)]));
    }
    output.push('_');
}

fn xstring_escape_at(bytes: &[u8], index: usize) -> Option<(u16, usize)> {
    let escape = bytes.get(index..index.checked_add(7)?)?;
    if escape[0] != b'_' || escape[1] != b'x' || escape[6] != b'_' {
        return None;
    }
    let mut value = 0u16;
    for byte in &escape[2..6] {
        value = value.checked_mul(16)?;
        value = value.checked_add(hex(*byte)?)?;
    }
    Some((value, index + 7))
}

fn hex(byte: u8) -> Option<u16> {
    match byte {
        b'0'..=b'9' => Some(u16::from(byte - b'0')),
        b'a'..=b'f' => Some(u16::from(byte - b'a' + 10)),
        b'A'..=b'F' => Some(u16::from(byte - b'A' + 10)),
        _ => None,
    }
}

fn ensure_output_limit_with(output_len: usize, maximum: usize) -> Result<()> {
    if output_len > maximum {
        return Err(limit(output_limit_kind(maximum), output_len, maximum));
    }
    Ok(())
}

fn output_limit_kind(maximum: usize) -> LimitKind {
    if maximum == MAX_XML_BYTES {
        LimitKind::XmlBytes
    } else {
        LimitKind::OutputBytes
    }
}

fn is_xml_whitespace(bytes: &[u8]) -> Result<bool> {
    let text = std::str::from_utf8(bytes).map_err(xml_error)?;
    Ok(text
        .chars()
        .all(|character| matches!(character, ' ' | '\t' | '\r' | '\n')))
}

fn validate_xml_text(value: &str, description: &'static str) -> Result<()> {
    for character in value.chars() {
        let valid = character == '\t'
            || character == '\n'
            || character == '\r'
            || (' '..='\u{D7FF}').contains(&character)
            || ('\u{E000}'..='\u{FFFD}').contains(&character)
            || ('\u{10000}'..='\u{10FFFF}').contains(&character);
        if !valid {
            return Err(invalid(format!(
                "{description} contains an invalid XML character"
            )));
        }
    }
    Ok(())
}

fn validate_comment(comment: &BytesText<'_>) -> Result<()> {
    let text = std::str::from_utf8(comment.as_ref()).map_err(xml_error)?;
    validate_xml_text(text, "formatcode2 comment")
}

fn validate_processing_instruction(instruction: &BytesPI<'_>) -> Result<()> {
    let target = std::str::from_utf8(instruction.target()).map_err(xml_error)?;
    if !is_ncname(target) {
        return Err(invalid(
            "formatcode2 processing-instruction target is invalid",
        ));
    }
    if target.eq_ignore_ascii_case("xml") {
        return Err(invalid(
            "formatcode2 processing-instruction target is reserved",
        ));
    }
    let content = std::str::from_utf8(instruction.content()).map_err(xml_error)?;
    validate_xml_text(content, "formatcode2 processing-instruction")
}

fn position(reader: &NsReader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| invalid("formatcode2 XML position exceeds usize"))
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

fn xml_event_error(error: impl std::fmt::Display) -> Error {
    let message = error.to_string();
    if message.contains("invalid character reference")
        || message.contains("character is not permitted in XML")
    {
        return invalid("formatcode2 reference contains an invalid XML character");
    }
    if message.contains("cannot be bound to 'http://www.w3.org/2000/xmlns/'") {
        return invalid("formatcode2 declaration binds the reserved XMLNS namespace");
    }
    if message.contains("namespace prefix") && message.contains("xmlns") {
        return invalid("formatcode2 namespace declaration prefix is invalid");
    }
    Error::Xml(message)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::limit(kind, observed, maximum)
}
