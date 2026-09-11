//! Source-backed support for the `[MS-ODRAWXML]` SVG blip extension.
//!
//! `asvg:svgBlip` is deliberately small in the standard: it is a
//! relationship-bearing element with the shared `AG_Blob` attributes.  The
//! public value therefore keeps the relationship metadata typed while
//! retaining unknown attributes and children for forward compatibility.  A
//! parsed value keeps its source bytes, so an untouched fragment can be
//! written without changing prefixes, attribute order, or lexical forms.

use std::{fmt, io::Write, sync::Arc};

use litchi_ooxml_common::{relationships, xml::is_ncname};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace as XmlNamespace, Prefix, QName, ResolveResult},
    reader::{NsReader, Reader},
};
use thiserror::Error as ThisError;

use crate::{Error, Result};

/// Transitional `[MS-ODRAWXML]` SVG namespace.
pub const NAMESPACE: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
/// The relationship namespace used by transitional and strict OOXML.
pub const RELATIONSHIP_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/// The strict OOXML relationship namespace.
pub const RELATIONSHIP_NAMESPACE_STRICT: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships";
/// The XML namespace reserved for the `xml` prefix.
pub const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
/// The namespace reserved for namespace declaration machinery.
pub const XMLNS_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";

/// Maximum source fragment size accepted by this bounded codec.
pub const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
/// Maximum relationship identifier size accepted by this codec.
pub const MAX_RELATIONSHIP_ID_BYTES: usize = 256;
/// Maximum namespace URI or prefix text retained by one fragment.
pub const MAX_NAMESPACE_BYTES: usize = 4_096;
/// Maximum decoded unknown attribute value retained by one fragment.
pub const MAX_ATTRIBUTE_VALUE_BYTES: usize = 1_048_576;
/// Maximum unknown qualified attribute name retained by one fragment.
pub const MAX_ATTRIBUTE_NAME_BYTES: usize = 4_096;
/// Maximum retained namespace declarations on one element.
pub const MAX_NAMESPACE_DECLARATIONS: usize = 256;
/// Maximum retained unknown attributes on one element.
pub const MAX_ATTRIBUTES: usize = 512;
/// Maximum retained unknown child elements on one element.
pub const MAX_CHILDREN: usize = 512;
/// Maximum nested XML depth while validating unknown children.
pub const MAX_DEPTH: usize = 128;
/// Maximum element count while validating unknown children.
pub const MAX_NODES: usize = 100_000;

/// A checked `ST_RelationshipId` value.
#[derive(Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[must_use]
pub struct RelationshipId(Box<str>);

impl RelationshipId {
    /// Construct a bounded XML `NCName` relationship identifier.
    ///
    /// # Errors
    ///
    /// Returns [`ValueError::RelationshipId`] when the value is empty, too
    /// long, or is not an XML `NCName`.
    pub fn new(value: impl AsRef<str>) -> std::result::Result<Self, ValueError> {
        let value = value.as_ref();
        if value.len() > MAX_RELATIONSHIP_ID_BYTES {
            return Err(ValueError::TooLong {
                field: "relationship ID",
                limit: MAX_RELATIONSHIP_ID_BYTES,
            });
        }
        if value.is_empty() || !is_ncname(value) {
            return Err(ValueError::RelationshipId {
                value: value.to_owned(),
            });
        }
        Ok(Self(value.into()))
    }

    /// Borrow the exact lexical relationship identifier.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for RelationshipId {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl fmt::Display for RelationshipId {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl TryFrom<&str> for RelationshipId {
    type Error = ValueError;

    fn try_from(value: &str) -> std::result::Result<Self, Self::Error> {
        Self::new(value)
    }
}

impl TryFrom<String> for RelationshipId {
    type Error = ValueError;

    fn try_from(value: String) -> std::result::Result<Self, Self::Error> {
        Self::new(value)
    }
}

impl From<RelationshipId> for String {
    fn from(value: RelationshipId) -> Self {
        value.0.into()
    }
}

/// The two independent alternatives of the DrawingML `AG_Blob` attribute
/// group.  Most Office files use `embedded`; linked SVG resources are retained
/// as a first-class value because dropping them would change document meaning.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
#[must_use]
pub struct Reference {
    /// Package-local SVG relationship (`r:embed`).
    pub embedded: Option<RelationshipId>,
    /// External SVG relationship (`r:link`).
    pub linked: Option<RelationshipId>,
}

impl Reference {
    /// Construct an empty relationship reference.
    #[must_use]
    pub const fn none() -> Self {
        Self {
            embedded: None,
            linked: None,
        }
    }

    /// Construct an embedded relationship reference.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is not a bounded XML relationship ID.
    pub fn embedded(value: impl AsRef<str>) -> std::result::Result<Self, ValueError> {
        Ok(Self {
            embedded: Some(RelationshipId::new(value)?),
            linked: None,
        })
    }

    /// Construct a linked relationship reference.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is not a bounded XML relationship ID.
    pub fn linked(value: impl AsRef<str>) -> std::result::Result<Self, ValueError> {
        Ok(Self {
            embedded: None,
            linked: Some(RelationshipId::new(value)?),
        })
    }

    /// Return whether neither relationship attribute is present.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.embedded.is_none() && self.linked.is_none()
    }
}

/// A namespace declaration retained from a parsed `svgBlip` element.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[must_use]
pub struct Namespace {
    prefix: Option<Box<str>>,
    uri: Box<str>,
}

impl Namespace {
    /// Construct a namespace declaration; `None` denotes the default prefix.
    /// An empty URI clears the default namespace, as in `xmlns=""`.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid prefix, an empty prefixed URI, or an
    /// overlong URI.
    pub fn new(
        prefix: Option<&str>,
        uri: impl AsRef<str>,
    ) -> std::result::Result<Self, ValueError> {
        if let Some(prefix) = prefix
            && (!prefix.is_empty() && (prefix.len() > MAX_NAMESPACE_BYTES || !is_ncname(prefix)))
        {
            if prefix.len() > MAX_NAMESPACE_BYTES {
                return Err(ValueError::TooLong {
                    field: "namespace prefix",
                    limit: MAX_NAMESPACE_BYTES,
                });
            }
            return Err(ValueError::NamespacePrefix {
                value: prefix.to_owned(),
            });
        }
        if prefix == Some("xmlns") {
            return Err(ValueError::NamespaceBinding);
        }
        let uri = uri.as_ref();
        if uri.is_empty() && prefix.is_some_and(|value| !value.is_empty()) {
            return Err(ValueError::EmptyNamespace);
        }
        if uri.len() > MAX_NAMESPACE_BYTES {
            return Err(ValueError::TooLong {
                field: "namespace URI",
                limit: MAX_NAMESPACE_BYTES,
            });
        }
        if uri == XMLNS_NAMESPACE
            || (prefix == Some("xml") && uri != XML_NAMESPACE)
            || (prefix != Some("xml") && uri == XML_NAMESPACE)
        {
            return Err(ValueError::NamespaceBinding);
        }
        Ok(Self {
            prefix: prefix.filter(|value| !value.is_empty()).map(Into::into),
            uri: uri.into(),
        })
    }

    /// Borrow the prefix; `None` denotes the default namespace.
    #[must_use]
    pub fn prefix(&self) -> Option<&str> {
        self.prefix.as_deref()
    }

    /// Borrow the namespace URI.
    #[must_use]
    pub fn uri(&self) -> &str {
        &self.uri
    }
}

/// An unknown attribute retained for a future extension vocabulary.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct Attribute {
    name: Box<str>,
    value: Box<str>,
}

impl Attribute {
    /// Construct a retained XML attribute.
    ///
    /// # Errors
    ///
    /// Returns an error when the name is not a qualified XML name.
    pub fn new(
        name: impl AsRef<str>,
        value: impl AsRef<str>,
    ) -> std::result::Result<Self, ValueError> {
        let name = name.as_ref();
        if name.len() > MAX_ATTRIBUTE_NAME_BYTES {
            return Err(ValueError::TooLong {
                field: "attribute name",
                limit: MAX_ATTRIBUTE_NAME_BYTES,
            });
        }
        if !litchi_ooxml_common::xml_name::is_qualified_name(name)
            || name == "xmlns"
            || name.starts_with("xmlns:")
        {
            return Err(ValueError::AttributeName {
                value: name.to_owned(),
            });
        }
        if value.as_ref().len() > MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(ValueError::TooLong {
                field: "attribute value",
                limit: MAX_ATTRIBUTE_VALUE_BYTES,
            });
        }
        Ok(Self {
            name: name.into(),
            value: value.as_ref().into(),
        })
    }

    /// Borrow the qualified attribute name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Borrow the decoded attribute value.
    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }
}

/// An unmodeled child content item retained in order for forward compatibility.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct Child {
    xml: Box<[u8]>,
}

impl Child {
    /// Borrow the exact child content bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.xml
    }
}

/// Failure to construct a bounded SVG blip scalar.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum ValueError {
    /// A relationship identifier is invalid or exceeds the bound.
    #[error("invalid SVG relationship ID '{value}'")]
    RelationshipId { value: String },
    /// A namespace prefix is invalid.
    #[error("invalid SVG namespace prefix '{value}'")]
    NamespacePrefix { value: String },
    /// A prefixed namespace URI is empty.
    #[error("prefixed SVG namespace URI is empty")]
    EmptyNamespace,
    /// An unknown attribute name is not an XML qualified name.
    #[error("invalid SVG attribute name '{value}'")]
    AttributeName { value: String },
    /// A namespace declaration uses a reserved XML namespace binding.
    #[error("invalid reserved SVG namespace binding")]
    NamespaceBinding,
    /// A retained scalar exceeds its explicit memory bound.
    #[error("SVG {field} exceeds the limit of {limit} bytes")]
    TooLong { field: &'static str, limit: usize },
}

/// Typed and lossless metadata for one `asvg:svgBlip` element.
#[derive(Debug, Clone)]
#[must_use]
pub struct SvgBlip {
    reference: Reference,
    namespaces: Vec<Namespace>,
    attributes: Vec<Attribute>,
    children: Vec<Child>,
    prefix: Option<Box<str>>,
    source: Option<Arc<[u8]>>,
}

impl PartialEq for SvgBlip {
    fn eq(&self, other: &Self) -> bool {
        self.reference == other.reference
            && self.namespaces == other.namespaces
            && self.attributes == other.attributes
            && self.children == other.children
            && self.prefix == other.prefix
            && self.source.as_deref() == other.source.as_deref()
    }
}

impl Eq for SvgBlip {}

impl SvgBlip {
    /// Construct a new empty SVG blip with the supplied relationship metadata.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::Invalid`] if the reference is not valid.
    pub fn new(reference: Reference) -> Result<Self> {
        validate_reference(&reference)?;
        Ok(Self {
            reference,
            namespaces: Vec::new(),
            attributes: Vec::new(),
            children: Vec::new(),
            prefix: Some("asvg".into()),
            source: None,
        })
    }

    /// Return the typed relationship metadata.
    #[must_use]
    pub const fn reference(&self) -> &Reference {
        &self.reference
    }

    /// Return the embedded SVG relationship, when present.
    #[must_use]
    pub fn embedded(&self) -> Option<&RelationshipId> {
        self.reference.embedded.as_ref()
    }

    /// Return the linked SVG relationship, when present.
    #[must_use]
    pub fn linked(&self) -> Option<&RelationshipId> {
        self.reference.linked.as_ref()
    }

    /// Replace relationship metadata and invalidate the source fast path.
    ///
    /// # Errors
    ///
    /// Returns an error when the reference reuses one relationship ID for both
    /// alternatives.
    pub fn set_reference(&mut self, reference: Reference) -> Result<()> {
        validate_reference(&reference)?;
        self.reference = reference;
        self.source = None;
        Ok(())
    }

    /// Borrow retained namespace declarations.
    #[must_use]
    pub fn namespaces(&self) -> &[Namespace] {
        &self.namespaces
    }

    /// Borrow retained unknown attributes.
    #[must_use]
    pub fn attributes(&self) -> &[Attribute] {
        &self.attributes
    }

    /// Borrow retained unknown children.
    #[must_use]
    pub fn children(&self) -> &[Child] {
        &self.children
    }

    /// Borrow the exact source fragment when this value came from XML.
    #[must_use]
    pub fn source(&self) -> Option<&[u8]> {
        self.source.as_deref()
    }

    pub(crate) fn from_wire(
        reference: Reference,
        namespaces: Vec<Namespace>,
        attributes: Vec<Attribute>,
        children: Vec<Child>,
        prefix: Option<Box<str>>,
        source: Arc<[u8]>,
    ) -> Result<Self> {
        validate_reference(&reference)?;
        Ok(Self {
            reference,
            namespaces,
            attributes,
            children,
            prefix,
            source: Some(source),
        })
    }
}

/// Read one complete SVG blip fragment.
///
/// # Errors
///
/// Returns an error for malformed XML, a wrong root namespace/name, invalid
/// relationship IDs, or exhausted resource bounds.
pub fn read(xml: &[u8]) -> Result<SvgBlip> {
    codec::read(xml)
}

/// Serialize one SVG blip fragment.
///
/// Parsed, unchanged values return their source bytes.  Modified or newly
/// constructed values use deterministic XML with retained unknown content.
///
/// # Errors
///
/// Returns an error when validation fails or the bounded output is too large.
pub fn write(value: &SvgBlip) -> Result<Vec<u8>> {
    codec::write(value)
}

/// Serialize one SVG blip fragment to a caller-provided sink.
///
/// # Errors
///
/// Returns an error when validation or the sink write fails.
pub fn write_to<W: Write>(writer: &mut W, value: &SvgBlip) -> Result<()> {
    codec::write_to(writer, value)
}

/// XML codec for [`SvgBlip`].
pub mod codec {
    use super::*;

    /// Read one complete `svgBlip` element.
    pub fn read(xml: &[u8]) -> Result<SvgBlip> {
        if xml.len() > MAX_XML_BYTES {
            return Err(limit("SVG blip XML bytes", MAX_XML_BYTES));
        }
        let mut reader = NsReader::from_reader(xml);
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        reader.config_mut().check_comments = true;
        let mut buffer = Vec::new();
        let mut root_seen = false;
        let mut root_closed = false;
        let mut value = None;

        loop {
            let event_start = position(&reader, "SVG blip")?;
            let (resolved, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(xml_error)?;
            let resolved = resolved_namespace(&resolved)?;
            let event = event.into_owned();
            let event_end = position(&reader, "SVG blip")?;
            match event {
                Event::Decl(_) if !root_seen => {},
                Event::Start(element) if !root_seen => {
                    let (local, namespace) = resolved_name(&resolved, &element.name())?;
                    require_root(&local, &namespace, element.name().prefix())?;
                    root_seen = true;
                    let root_start = event_start;
                    let child_end = capture_element(&mut reader, &mut buffer)?;
                    value = Some(parse_root(
                        &element,
                        &reader,
                        xml,
                        root_start,
                        child_end,
                        local_prefix(&element.name())?,
                        false,
                    )?);
                    root_closed = true;
                },
                Event::Empty(element) if !root_seen => {
                    let (local, namespace) = resolved_name(&resolved, &element.name())?;
                    require_root(&local, &namespace, element.name().prefix())?;
                    root_seen = true;
                    root_closed = true;
                    value = Some(parse_root(
                        &element,
                        &reader,
                        xml,
                        event_start,
                        event_end,
                        local_prefix(&element.name())?,
                        true,
                    )?);
                },
                Event::Start(_) | Event::Empty(_) if root_seen => {
                    return Err(invalid("SVG blip fragment has more than one root element"));
                },
                Event::Start(_) | Event::Empty(_) => {
                    return Err(invalid("SVG blip fragment root is invalid"));
                },
                Event::Text(text)
                    if (!root_seen || root_closed)
                        && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
                {
                    return Err(invalid("SVG blip fragment has text outside its root"));
                },
                Event::CData(_) | Event::GeneralRef(_) if !root_seen || root_closed => {
                    return Err(invalid("SVG blip fragment has markup outside its root"));
                },
                Event::DocType(_) | Event::PI(_) => {
                    return Err(invalid(
                        "SVG blip fragment contains forbidden document markup",
                    ));
                },
                Event::Eof => break,
                Event::End(_) | Event::Decl(_) if !root_seen || root_closed => {
                    return Err(invalid("SVG blip fragment has markup outside its root"));
                },
                Event::End(_)
                | Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::GeneralRef(_) => {},
                Event::Decl(_) => {
                    return Err(invalid("SVG blip fragment has a late declaration"));
                },
            }
            buffer.clear();
        }

        let value = value.ok_or_else(|| invalid("SVG blip fragment has no root element"))?;
        if !root_closed {
            return Err(invalid("SVG blip fragment root is not closed"));
        }
        validate(&value)?;
        Ok(value)
    }

    /// Serialize one SVG blip fragment.
    pub fn write(value: &SvgBlip) -> Result<Vec<u8>> {
        validate(value)?;
        if let Some(source) = &value.source {
            return Ok(source.to_vec());
        }
        let plan = output_plan(value)?;
        let size = serialized_size(value, &plan)?;
        let mut output = Vec::new();
        output.reserve_exact(size);
        write_inner(&mut output, value, &plan);
        debug_assert_eq!(output.len(), size);
        Ok(output)
    }

    /// Serialize one SVG blip fragment to a caller-owned sink.
    pub fn write_to<W: Write>(writer: &mut W, value: &SvgBlip) -> Result<()> {
        writer.write_all(&write(value)?)?;
        Ok(())
    }

    fn output_plan(value: &SvgBlip) -> Result<OutputPlan> {
        let root_prefix = choose_root_prefix(value)?;
        let relationship_prefix = if value.reference.is_empty() {
            None
        } else {
            Some(choose_relationship_prefix(value)?.into_boxed_str())
        };
        Ok(OutputPlan {
            root_prefix,
            relationship_prefix,
        })
    }

    fn choose_root_prefix(value: &SvgBlip) -> Result<Option<Box<str>>> {
        let requested = value.prefix.as_deref();
        if !has_namespace_binding(value, requested, NAMESPACE)
            && value
                .namespaces
                .iter()
                .any(|namespace| namespace.prefix() == requested)
        {
            return Ok(Some(
                next_prefix(value, requested.unwrap_or("asvg"))?.into(),
            ));
        }
        Ok(requested.map(Into::into))
    }

    fn choose_relationship_prefix(value: &SvgBlip) -> Result<String> {
        if let Some(namespace) = value.namespaces.iter().find(|namespace| {
            namespace.prefix() == Some("r") && is_relationship_uri(namespace.uri())
        }) {
            return Ok(namespace.prefix().unwrap_or("r").to_owned());
        }
        if let Some(namespace) = value
            .namespaces
            .iter()
            .find(|namespace| is_relationship_uri(namespace.uri()) && namespace.prefix().is_some())
        {
            return Ok(namespace.prefix().unwrap_or("r").to_owned());
        }
        if !value
            .namespaces
            .iter()
            .any(|namespace| namespace.prefix() == Some("r"))
        {
            return Ok("r".to_owned());
        }
        next_prefix(value, "r")
    }

    fn next_prefix(value: &SvgBlip, base: &str) -> Result<String> {
        let base = if base.is_empty() { "asvg" } else { base };
        for suffix in 2..=MAX_NAMESPACE_DECLARATIONS + 2 {
            let candidate = format!("{base}{suffix}");
            if candidate.len() > MAX_NAMESPACE_BYTES {
                return Err(limit("SVG namespace prefix", MAX_NAMESPACE_BYTES));
            }
            if !value
                .namespaces
                .iter()
                .any(|namespace| namespace.prefix() == Some(candidate.as_str()))
            {
                return Ok(candidate);
            }
        }
        Err(limit("SVG namespace prefixes", MAX_NAMESPACE_DECLARATIONS))
    }

    fn has_namespace_binding(value: &SvgBlip, prefix: Option<&str>, uri: &str) -> bool {
        value
            .namespaces
            .iter()
            .any(|namespace| namespace.prefix() == prefix && namespace.uri() == uri)
    }

    fn has_relationship_binding(value: &SvgBlip, prefix: &str) -> bool {
        value.namespaces.iter().any(|namespace| {
            namespace.prefix() == Some(prefix) && is_relationship_uri(namespace.uri())
        })
    }

    fn is_relationship_uri(uri: &str) -> bool {
        uri == RELATIONSHIP_NAMESPACE || uri == RELATIONSHIP_NAMESPACE_STRICT
    }

    fn serialized_size(value: &SvgBlip, plan: &OutputPlan) -> Result<usize> {
        let mut size = 1usize;
        if let Some(prefix) = plan.root_prefix.as_deref() {
            add_size(&mut size, prefix.len() + 1)?;
        }
        add_size(&mut size, b"svgBlip".len())?;
        for namespace in &value.namespaces {
            namespace_size(&mut size, namespace.prefix(), namespace.uri())?;
        }
        if !has_namespace_binding(value, plan.root_prefix.as_deref(), NAMESPACE) {
            namespace_size(&mut size, plan.root_prefix.as_deref(), NAMESPACE)?;
        }
        if let Some(prefix) = plan.relationship_prefix.as_deref()
            && !has_relationship_binding(value, prefix)
        {
            namespace_size(&mut size, Some(prefix), RELATIONSHIP_NAMESPACE)?;
        }
        for attribute in &value.attributes {
            add_size(&mut size, 4 + attribute.name().len())?;
            add_size(&mut size, escaped_size(attribute.value())?)?;
        }
        if let Some(prefix) = plan.relationship_prefix.as_deref() {
            if let Some(id) = &value.reference.embedded {
                relationship_size(&mut size, prefix, "embed", id.as_str())?;
            }
            if let Some(id) = &value.reference.linked {
                relationship_size(&mut size, prefix, "link", id.as_str())?;
            }
        }
        if value.children.is_empty() {
            add_size(&mut size, 2)?;
        } else {
            add_size(&mut size, 1)?;
            for child in &value.children {
                add_size(&mut size, child.as_bytes().len())?;
            }
            add_size(&mut size, 2)?;
            if let Some(prefix) = plan.root_prefix.as_deref() {
                add_size(&mut size, prefix.len() + 1)?;
            }
            add_size(&mut size, b"svgBlip>".len())?;
        }
        Ok(size)
    }

    fn namespace_size(size: &mut usize, prefix: Option<&str>, uri: &str) -> Result<()> {
        add_size(size, 6 + prefix.map_or(0, |prefix| prefix.len() + 1) + 3)?;
        add_size(size, escaped_size(uri)?)
    }

    fn relationship_size(size: &mut usize, prefix: &str, name: &str, value: &str) -> Result<()> {
        add_size(size, 5 + prefix.len() + name.len())?;
        add_size(size, escaped_size(value)?)
    }

    fn escaped_size(value: &str) -> Result<usize> {
        let mut size = 0usize;
        for byte in value.bytes() {
            let escaped = match byte {
                b'<' | b'>' => 4,
                b'&' => 5,
                b'\'' | b'"' => 6,
                _ => 1,
            };
            add_size(&mut size, escaped)?;
        }
        Ok(size)
    }

    fn add_size(size: &mut usize, addition: usize) -> Result<()> {
        *size = size
            .checked_add(addition)
            .ok_or_else(|| limit("SVG blip output bytes", MAX_XML_BYTES))?;
        if *size > MAX_XML_BYTES {
            return Err(limit("SVG blip output bytes", MAX_XML_BYTES));
        }
        Ok(())
    }

    fn parse_root<R: std::io::BufRead>(
        element: &BytesStart<'_>,
        reader: &NsReader<R>,
        xml: &[u8],
        start: usize,
        end: usize,
        prefix: Option<Box<str>>,
        empty: bool,
    ) -> Result<SvgBlip> {
        let decoder = reader.decoder();
        let namespaces = declarations(element, decoder)?;
        if namespaces.len() > MAX_NAMESPACE_DECLARATIONS {
            return Err(limit(
                "SVG blip namespace declarations",
                MAX_NAMESPACE_DECLARATIONS,
            ));
        }
        let reference = reference(element, reader)?;
        let mut attributes = Vec::new();
        for attribute in element.attributes() {
            let attribute = attribute.map_err(xml_error)?;
            let raw_name = attribute.key.as_ref();
            if raw_name == b"xmlns"
                || raw_name.starts_with(b"xmlns:")
                || is_relationship_attribute(attribute.key, reader.resolver())
            {
                continue;
            }
            validate_attribute_namespace(attribute.key, reader.resolver(), prefix.as_deref())?;
            if attribute.value.len() > MAX_ATTRIBUTE_VALUE_BYTES {
                return Err(limit(
                    "SVG blip attribute value bytes",
                    MAX_ATTRIBUTE_VALUE_BYTES,
                ));
            }
            let name = std::str::from_utf8(raw_name).map_err(xml_error)?;
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(xml_error)?
                .into_owned();
            attributes.push(Attribute::new(name, value).map_err(value_error)?);
            if attributes.len() > MAX_ATTRIBUTES {
                return Err(limit("SVG blip attributes", MAX_ATTRIBUTES));
            }
        }
        let mut children = Vec::new();
        if !empty {
            let raw = xml
                .get(start..end)
                .ok_or_else(|| invalid("SVG blip child range is outside input"))?;
            collect_children(raw, &mut children)?;
        }
        SvgBlip::from_wire(
            reference,
            namespaces,
            attributes,
            children,
            prefix,
            Arc::from(
                xml.get(start..end)
                    .ok_or_else(|| invalid("SVG blip source range is outside input"))?
                    .to_vec()
                    .into_boxed_slice(),
            ),
        )
    }

    fn collect_children(xml: &[u8], children: &mut Vec<Child>) -> Result<()> {
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        let mut buffer = Vec::new();
        next_element(&mut reader, &mut buffer)?;
        loop {
            let start = position(&reader, "SVG blip child")?;
            let event = reader
                .read_event_into(&mut buffer)
                .map_err(xml_error)?
                .into_owned();
            let end = position(&reader, "SVG blip child")?;
            match event {
                Event::Start(_) => {
                    let end = capture_plain_element(&mut reader, &mut buffer)?;
                    retain_child(xml, start, end, children)?;
                },
                Event::Empty(_) => {
                    retain_child(xml, start, end, children)?;
                },
                Event::End(_) => break,
                Event::DocType(_) | Event::PI(_) | Event::Decl(_) => {
                    return Err(invalid("SVG blip contains forbidden child markup"));
                },
                Event::Eof => return Err(invalid("SVG blip is unterminated")),
                Event::Text(_) | Event::CData(_) | Event::Comment(_) | Event::GeneralRef(_) => {
                    retain_child(xml, start, end, children)?;
                },
            }
            buffer.clear();
        }
        Ok(())
    }

    fn retain_child(xml: &[u8], start: usize, end: usize, children: &mut Vec<Child>) -> Result<()> {
        if children.len() >= MAX_CHILDREN {
            return Err(limit("SVG blip children", MAX_CHILDREN));
        }
        let raw = xml
            .get(start..end)
            .ok_or_else(|| invalid("SVG blip child range is outside input"))?;
        children.push(Child {
            xml: raw.to_vec().into_boxed_slice(),
        });
        Ok(())
    }

    struct OutputPlan {
        root_prefix: Option<Box<str>>,
        relationship_prefix: Option<Box<str>>,
    }

    fn write_inner(output: &mut Vec<u8>, value: &SvgBlip, plan: &OutputPlan) {
        output.extend_from_slice(b"<");
        if let Some(prefix) = plan.root_prefix.as_deref() {
            output.extend_from_slice(prefix.as_bytes());
            output.push(b':');
        }
        output.extend_from_slice(b"svgBlip");
        for namespace in &value.namespaces {
            write_namespace(output, namespace);
        }
        if !has_namespace_binding(value, plan.root_prefix.as_deref(), NAMESPACE) {
            write_namespace_value(output, plan.root_prefix.as_deref(), NAMESPACE);
        }
        if let Some(prefix) = plan.relationship_prefix.as_deref()
            && !has_relationship_binding(value, prefix)
        {
            write_namespace_value(output, Some(prefix), RELATIONSHIP_NAMESPACE);
        }
        for attribute in &value.attributes {
            output.extend_from_slice(b" ");
            output.extend_from_slice(attribute.name().as_bytes());
            output.extend_from_slice(b"=\"");
            push_escaped(output, attribute.value());
            output.push(b'"');
        }
        write_reference(
            output,
            &value.reference,
            plan.relationship_prefix.as_deref(),
        );
        if value.children.is_empty() {
            output.extend_from_slice(b"/>");
        } else {
            output.push(b'>');
            for child in &value.children {
                output.extend_from_slice(child.as_bytes());
            }
            output.extend_from_slice(b"</");
            if let Some(prefix) = plan.root_prefix.as_deref() {
                output.extend_from_slice(prefix.as_bytes());
                output.push(b':');
            }
            output.extend_from_slice(b"svgBlip>");
        }
    }

    fn write_reference(output: &mut Vec<u8>, reference: &Reference, prefix: Option<&str>) {
        let Some(prefix) = prefix else {
            return;
        };
        if let Some(id) = &reference.embedded {
            output.extend_from_slice(b" ");
            output.extend_from_slice(prefix.as_bytes());
            output.extend_from_slice(b":embed=\"");
            push_escaped(output, id.as_str());
            output.push(b'"');
        }
        if let Some(id) = &reference.linked {
            output.extend_from_slice(b" ");
            output.extend_from_slice(prefix.as_bytes());
            output.extend_from_slice(b":link=\"");
            push_escaped(output, id.as_str());
            output.push(b'"');
        }
    }

    fn write_namespace(output: &mut Vec<u8>, namespace: &Namespace) {
        write_namespace_value(output, namespace.prefix(), namespace.uri());
    }

    fn write_namespace_value(output: &mut Vec<u8>, prefix: Option<&str>, uri: &str) {
        output.extend_from_slice(b" xmlns");
        if let Some(prefix) = prefix {
            output.push(b':');
            output.extend_from_slice(prefix.as_bytes());
        }
        output.extend_from_slice(b"=\"");
        push_escaped(output, uri);
        output.push(b'"');
    }

    fn push_escaped(output: &mut Vec<u8>, value: &str) {
        output.extend_from_slice(quick_xml::escape::escape(value).as_bytes());
    }

    fn reference<R: std::io::BufRead>(
        element: &BytesStart<'_>,
        reader: &NsReader<R>,
    ) -> Result<Reference> {
        for attribute in element.attributes() {
            let attribute = attribute.map_err(xml_error)?;
            if (attribute.key.local_name().as_ref() == b"embed"
                || attribute.key.local_name().as_ref() == b"link")
                && attribute.value.len() > MAX_RELATIONSHIP_ID_BYTES * 4
            {
                return Err(limit(
                    "SVG relationship attribute bytes",
                    MAX_RELATIONSHIP_ID_BYTES * 4,
                ));
            }
        }
        let embedded =
            relationships::attribute_value(element, b"embed", reader.decoder(), reader.resolver())?;
        let linked =
            relationships::attribute_value(element, b"link", reader.decoder(), reader.resolver())?;
        let embedded = embedded
            .map(RelationshipId::new)
            .transpose()
            .map_err(value_error)?;
        let linked = linked
            .map(RelationshipId::new)
            .transpose()
            .map_err(value_error)?;
        Ok(Reference { embedded, linked })
    }

    fn declarations(
        element: &BytesStart<'_>,
        decoder: quick_xml::encoding::Decoder,
    ) -> Result<Vec<Namespace>> {
        let mut result = Vec::new();
        for attribute in element.attributes() {
            let attribute = attribute.map_err(xml_error)?;
            let raw = attribute.key.as_ref();
            let prefix = if raw == b"xmlns" {
                None
            } else if let Some(prefix) = raw.strip_prefix(b"xmlns:") {
                Some(std::str::from_utf8(prefix).map_err(xml_error)?)
            } else {
                continue;
            };
            if result
                .iter()
                .any(|item: &Namespace| item.prefix() == prefix)
            {
                return Err(invalid("SVG blip has duplicate namespace declarations"));
            }
            attribute
                .value
                .len()
                .le(&MAX_NAMESPACE_BYTES)
                .then_some(())
                .ok_or_else(|| limit("SVG blip namespace URI bytes", MAX_NAMESPACE_BYTES))?;
            let uri = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(xml_error)?
                .into_owned();
            result.push(Namespace::new(prefix, uri).map_err(value_error)?);
        }
        Ok(result)
    }

    fn is_relationship_attribute(
        key: QName<'_>,
        resolver: &quick_xml::name::NamespaceResolver,
    ) -> bool {
        let local = key.local_name();
        if local.as_ref() != b"embed" && local.as_ref() != b"link" {
            return false;
        }
        let (namespace, _) = resolver.resolve_attribute(key);
        matches!(
            namespace,
            ResolveResult::Bound(XmlNamespace(value))
                if value == relationships::TRANSITIONAL_NAMESPACE
                    || value == relationships::STRICT_NAMESPACE
        ) || matches!(namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"r")
    }

    fn validate_attribute_namespace(
        key: QName<'_>,
        resolver: &quick_xml::name::NamespaceResolver,
        fallback_prefix: Option<&str>,
    ) -> Result<()> {
        let Some(prefix) = key.prefix() else {
            return Ok(());
        };
        if prefix.as_ref() == b"xml" {
            return Ok(());
        }
        let (namespace, _) = resolver.resolve_attribute(key);
        match namespace {
            ResolveResult::Bound(_) => Ok(()),
            ResolveResult::Unknown(unresolved)
                if fallback_prefix
                    .is_some_and(|fallback| fallback.as_bytes() == unresolved.as_slice()) =>
            {
                Ok(())
            },
            _ => Err(invalid("SVG blip attribute prefix is not bound")),
        }
    }

    fn require_root(local: &str, namespace: &str, prefix: Option<Prefix<'_>>) -> Result<()> {
        if local != "svgBlip"
            || (namespace != NAMESPACE
                && !(namespace.is_empty()
                    && prefix.is_some_and(|prefix| prefix.as_ref() == b"asvg")))
        {
            return Err(invalid("SVG blip root name or namespace is invalid"));
        }
        Ok(())
    }

    fn local_prefix(name: &QName<'_>) -> Result<Option<Box<str>>> {
        name.prefix()
            .map(|prefix| {
                std::str::from_utf8(prefix.as_ref())
                    .map(Into::into)
                    .map_err(xml_error)
            })
            .transpose()
    }

    fn resolved_name(resolved: &str, name: &QName<'_>) -> Result<(String, String)> {
        let local = std::str::from_utf8(name.local_name().as_ref())
            .map(str::to_owned)
            .map_err(xml_error)?;
        Ok((local, resolved.to_owned()))
    }

    fn resolved_namespace(resolved: &ResolveResult<'_>) -> Result<String> {
        match resolved {
            ResolveResult::Bound(namespace) => std::str::from_utf8(namespace.as_ref())
                .map(str::to_owned)
                .map_err(xml_error),
            ResolveResult::Unknown(_) | ResolveResult::Unbound => Ok(String::new()),
        }
    }

    fn next_element(reader: &mut Reader<&[u8]>, buffer: &mut Vec<u8>) -> Result<()> {
        loop {
            let event = reader
                .read_event_into(buffer)
                .map_err(xml_error)?
                .into_owned();
            match event {
                Event::Start(_) | Event::Empty(_) => return Ok(()),
                Event::Text(text) if !text.as_ref().iter().all(u8::is_ascii_whitespace) => {
                    return Err(invalid("SVG blip has text outside its root"));
                },
                Event::Decl(_) => {},
                Event::DocType(_) | Event::PI(_) => {
                    return Err(invalid("SVG blip contains forbidden document markup"));
                },
                Event::Eof => return Err(invalid("SVG blip has no root")),
                Event::End(_)
                | Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::GeneralRef(_) => {},
            }
            buffer.clear();
        }
    }

    fn capture_element<R: std::io::BufRead>(
        reader: &mut NsReader<R>,
        buffer: &mut Vec<u8>,
    ) -> Result<usize> {
        let mut depth = 1usize;
        let mut nodes = 0usize;
        loop {
            buffer.clear();
            let (_, event) = reader.read_resolved_event_into(buffer).map_err(xml_error)?;
            let event = event.into_owned();
            let end = position(reader, "SVG blip")?;
            match event {
                Event::Start(_) => {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("SVG blip XML nesting overflow"))?;
                    if depth > MAX_DEPTH {
                        return Err(limit("SVG blip XML depth", MAX_DEPTH));
                    }
                    nodes = nodes.saturating_add(1);
                    if nodes > MAX_NODES {
                        return Err(limit("SVG blip XML nodes", MAX_NODES));
                    }
                },
                Event::Empty(_) => {
                    nodes = nodes.saturating_add(1);
                    if nodes > MAX_NODES {
                        return Err(limit("SVG blip XML nodes", MAX_NODES));
                    }
                },
                Event::End(_) => {
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("SVG blip XML nesting underflow"))?;
                    if depth == 0 {
                        return Ok(end);
                    }
                },
                Event::DocType(_) | Event::PI(_) => {
                    return Err(invalid("SVG blip contains forbidden document markup"));
                },
                Event::Eof => return Err(invalid("SVG blip root is unterminated")),
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::GeneralRef(_) => {},
            }
        }
    }

    fn capture_plain_element<R: std::io::BufRead>(
        reader: &mut Reader<R>,
        buffer: &mut Vec<u8>,
    ) -> Result<usize> {
        let mut depth = 1usize;
        loop {
            buffer.clear();
            let event = reader
                .read_event_into(buffer)
                .map_err(xml_error)?
                .into_owned();
            let end = position(reader, "SVG blip child")?;
            match event {
                Event::Start(_) => {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("SVG blip child nesting overflow"))?;
                    if depth > MAX_DEPTH {
                        return Err(limit("SVG blip child depth", MAX_DEPTH));
                    }
                },
                Event::End(_) => {
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("SVG blip child nesting underflow"))?;
                    if depth == 0 {
                        return Ok(end);
                    }
                },
                Event::DocType(_) | Event::PI(_) => {
                    return Err(invalid("SVG blip child contains forbidden markup"));
                },
                Event::Eof => return Err(invalid("SVG blip child is unterminated")),
                Event::Empty(_)
                | Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::GeneralRef(_) => {},
            }
        }
    }

    fn position<R: std::io::BufRead>(reader: &Reader<R>, what: &str) -> Result<usize> {
        usize::try_from(reader.buffer_position())
            .map_err(|_| invalid(format!("{what} offset exceeds usize")))
    }

    fn validate(value: &SvgBlip) -> Result<()> {
        validate_reference(&value.reference)?;
        if value.namespaces.len() > MAX_NAMESPACE_DECLARATIONS {
            return Err(limit(
                "SVG blip namespace declarations",
                MAX_NAMESPACE_DECLARATIONS,
            ));
        }
        if value.attributes.len() > MAX_ATTRIBUTES {
            return Err(limit("SVG blip attributes", MAX_ATTRIBUTES));
        }
        if value.children.len() > MAX_CHILDREN {
            return Err(limit("SVG blip children", MAX_CHILDREN));
        }
        for namespace in &value.namespaces {
            let _ = Namespace::new(namespace.prefix(), namespace.uri()).map_err(value_error)?;
        }
        for attribute in &value.attributes {
            let _ = Attribute::new(attribute.name(), attribute.value()).map_err(value_error)?;
        }
        for child in &value.children {
            validate_fragment(child.as_bytes())?;
        }
        if let Some(source) = &value.source {
            if source.len() > MAX_XML_BYTES {
                return Err(limit("SVG blip XML bytes", MAX_XML_BYTES));
            }
        }
        Ok(())
    }

    fn validate_fragment(xml: &[u8]) -> Result<()> {
        if xml.is_empty() || xml.len() > MAX_XML_BYTES {
            return Err(limit("SVG blip child bytes", MAX_XML_BYTES));
        }
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        let mut buffer = Vec::new();
        let mut depth = 0usize;
        let mut top_level_items = 0usize;
        let mut nodes = 0usize;
        loop {
            let event = reader
                .read_event_into(&mut buffer)
                .map_err(xml_error)?
                .into_owned();
            match event {
                Event::Start(_) => {
                    if depth == 0 && top_level_items > 0 {
                        return Err(invalid("SVG blip child has multiple roots"));
                    }
                    if depth == 0 {
                        top_level_items += 1;
                    }
                    depth += 1;
                    nodes += 1;
                    if depth > MAX_DEPTH || nodes > MAX_NODES {
                        return Err(limit("SVG blip child nodes", MAX_NODES));
                    }
                },
                Event::Empty(_) => {
                    if depth == 0 && top_level_items > 0 {
                        return Err(invalid("SVG blip child has multiple roots"));
                    }
                    if depth == 0 {
                        top_level_items += 1;
                    }
                    nodes += 1;
                    if nodes > MAX_NODES {
                        return Err(limit("SVG blip child nodes", MAX_NODES));
                    }
                },
                Event::End(_) => {
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("SVG blip child has unexpected end"))?;
                },
                Event::Text(_) | Event::CData(_) | Event::Comment(_) | Event::GeneralRef(_)
                    if depth == 0 =>
                {
                    if top_level_items > 0 {
                        return Err(invalid("SVG blip child has multiple roots"));
                    }
                    top_level_items = 1;
                },
                Event::DocType(_) | Event::Decl(_) | Event::PI(_) => {
                    return Err(invalid("SVG blip child contains document markup"));
                },
                Event::Eof => break,
                Event::Text(_) | Event::CData(_) | Event::Comment(_) | Event::GeneralRef(_) => {},
            }
            buffer.clear();
        }
        if depth != 0 || top_level_items != 1 {
            return Err(invalid("SVG blip child is not one complete XML item"));
        }
        Ok(())
    }
}

fn validate_reference(reference: &Reference) -> Result<()> {
    if let (Some(embedded), Some(linked)) = (&reference.embedded, &reference.linked)
        && embedded == linked
    {
        return Err(invalid(
            "SVG blip reuses one relationship ID for embed and link",
        ));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}

fn xml_error(error: impl fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

fn value_error(error: impl fmt::Display) -> Error {
    Error::Invalid(error.to_string())
}
