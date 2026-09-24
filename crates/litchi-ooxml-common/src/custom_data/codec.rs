//! Bounded, source-preserving X14 `datastoreItem` XML codec.
//!
//! The typed projection intentionally contains only the required UID and one
//! direct extension list.  Every other byte remains in the source image.  In
//! particular, extension descendants are never rebuilt from a lossy tree.

use super::model::{ExtensionList, Properties};
use crate::{Error, Result, XmlError};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::QName;
use quick_xml::reader::NsReader;
use std::collections::HashSet;
use std::ops::Range;

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";

const MAX_PROPERTIES_XML_BYTES: usize = 4 * 1024 * 1024;
const MAX_EXTENSION_XML_BYTES: usize = 2 * 1024 * 1024;
const MAX_STRING_BYTES: usize = 1024 * 1024;
const MAX_NODES: usize = 100_000;
const MAX_EVENTS: usize = 1_000_000;
const MAX_DEPTH: usize = 128;
const MAX_NAMESPACE_BYTES: usize = MAX_PROPERTIES_XML_BYTES;
const MAX_ATTRIBUTES: usize = MAX_NODES;
const MAX_UID_UNITS: usize = 65_535;

/// Resource ceilings for the bounded Custom Data XML codec.
///
/// The public convenience functions use [`Self::standard`].  Package owners
/// should pass a lower profile through the `_with_limits` functions before
/// parsing or constructing any temporary XML state.  Every field is capped by
/// an immutable format-wide maximum; callers may tighten, but never widen,
/// those ceilings.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum bytes in one complete `datastoreItem` XML document or output.
    pub properties_xml_bytes: usize,
    /// Maximum bytes in one retained `extLst` fragment.
    pub extension_xml_bytes: usize,
    /// Maximum decoded bytes in one XML string value.
    pub string_bytes: usize,
    /// Maximum XML element nodes in one document or fragment.
    pub nodes: usize,
    /// Maximum parser events in one document or fragment.
    pub events: usize,
    /// Maximum XML element nesting depth.
    pub depth: usize,
    /// Maximum bytes retained in one namespace environment.
    pub namespace_bytes: usize,
    /// Maximum attributes or namespace declarations on one element.
    pub attributes: usize,
    /// Maximum UTF-16 code units in the required storage UID.
    pub uid_units: usize,
}

impl Limits {
    /// Conservative default resource profile used by the compatibility API.
    pub const DEFAULT: Self = Self {
        properties_xml_bytes: MAX_PROPERTIES_XML_BYTES,
        extension_xml_bytes: MAX_EXTENSION_XML_BYTES,
        string_bytes: MAX_STRING_BYTES,
        nodes: MAX_NODES,
        events: MAX_EVENTS,
        depth: MAX_DEPTH,
        namespace_bytes: MAX_NAMESPACE_BYTES,
        attributes: MAX_ATTRIBUTES,
        uid_units: MAX_UID_UNITS,
    };

    /// Absolute maximum bytes in one complete Properties XML document.
    pub const MAX_PROPERTIES_XML_BYTES: usize = MAX_PROPERTIES_XML_BYTES;
    /// Absolute maximum bytes in one retained extension fragment.
    pub const MAX_EXTENSION_XML_BYTES: usize = MAX_EXTENSION_XML_BYTES;
    /// Absolute maximum decoded bytes in one XML string value.
    pub const MAX_STRING_BYTES: usize = MAX_STRING_BYTES;
    /// Absolute maximum XML element nodes in one document.
    pub const MAX_NODES: usize = MAX_NODES;
    /// Absolute maximum parser events in one document.
    pub const MAX_EVENTS: usize = MAX_EVENTS;
    /// Absolute maximum XML nesting depth.
    pub const MAX_DEPTH: usize = MAX_DEPTH;
    /// Absolute maximum bytes in one namespace environment.
    pub const MAX_NAMESPACE_BYTES: usize = MAX_NAMESPACE_BYTES;
    /// Absolute maximum attributes on one element.
    pub const MAX_ATTRIBUTES: usize = MAX_ATTRIBUTES;
    /// Absolute maximum UTF-16 code units in a storage UID.
    pub const MAX_UID_UNITS: usize = MAX_UID_UNITS;

    /// Return the conservative default profile used by the compatibility API.
    #[must_use]
    pub const fn standard() -> Self {
        Self::DEFAULT
    }

    /// Validate this profile against immutable codec ceilings.
    pub const fn validate(self) -> Result<Self> {
        if self.properties_xml_bytes > MAX_PROPERTIES_XML_BYTES {
            return Err(limit_for(
                "properties XML bytes",
                self.properties_xml_bytes,
                MAX_PROPERTIES_XML_BYTES,
            ));
        }
        if self.extension_xml_bytes > MAX_EXTENSION_XML_BYTES {
            return Err(limit_for(
                "extension XML bytes",
                self.extension_xml_bytes,
                MAX_EXTENSION_XML_BYTES,
            ));
        }
        if self.string_bytes > MAX_STRING_BYTES {
            return Err(limit_for(
                "XML string bytes",
                self.string_bytes,
                MAX_STRING_BYTES,
            ));
        }
        if self.nodes > MAX_NODES {
            return Err(limit_for("XML node count", self.nodes, MAX_NODES));
        }
        if self.events > MAX_EVENTS {
            return Err(limit_for("XML event count", self.events, MAX_EVENTS));
        }
        if self.depth > MAX_DEPTH {
            return Err(limit_for("XML depth", self.depth, MAX_DEPTH));
        }
        if self.namespace_bytes > MAX_NAMESPACE_BYTES {
            return Err(limit_for(
                "Custom Data XML namespace bytes",
                self.namespace_bytes,
                MAX_NAMESPACE_BYTES,
            ));
        }
        if self.attributes > MAX_ATTRIBUTES {
            return Err(limit_for(
                "XML attribute count",
                self.attributes,
                MAX_ATTRIBUTES,
            ));
        }
        if self.uid_units > MAX_UID_UNITS {
            return Err(limit_for(
                "Custom Data UID UTF-16 units",
                self.uid_units,
                MAX_UID_UNITS,
            ));
        }
        Ok(self)
    }

    /// Set the complete Properties XML byte ceiling.
    #[must_use]
    pub const fn with_properties_xml_bytes(mut self, value: usize) -> Self {
        self.properties_xml_bytes = value;
        self
    }

    /// Set the retained extension-fragment byte ceiling.
    #[must_use]
    pub const fn with_extension_xml_bytes(mut self, value: usize) -> Self {
        self.extension_xml_bytes = value;
        self
    }

    /// Set the decoded XML-string byte ceiling.
    #[must_use]
    pub const fn with_string_bytes(mut self, value: usize) -> Self {
        self.string_bytes = value;
        self
    }

    /// Set the XML node ceiling.
    #[must_use]
    pub const fn with_nodes(mut self, value: usize) -> Self {
        self.nodes = value;
        self
    }

    /// Set the parser-event ceiling.
    #[must_use]
    pub const fn with_events(mut self, value: usize) -> Self {
        self.events = value;
        self
    }

    /// Set the XML nesting-depth ceiling.
    #[must_use]
    pub const fn with_depth(mut self, value: usize) -> Self {
        self.depth = value;
        self
    }

    /// Set the namespace-environment byte ceiling.
    #[must_use]
    pub const fn with_namespace_bytes(mut self, value: usize) -> Self {
        self.namespace_bytes = value;
        self
    }

    /// Set the per-element attribute count ceiling.
    #[must_use]
    pub const fn with_attributes(mut self, value: usize) -> Self {
        self.attributes = value;
        self
    }

    /// Set the storage UID UTF-16 code-unit ceiling.
    #[must_use]
    pub const fn with_uid_units(mut self, value: usize) -> Self {
        self.uid_units = value;
        self
    }

    /// Alias for [`Self::with_uid_units`].
    #[must_use]
    pub const fn with_max_uid_units(self, value: usize) -> Self {
        self.with_uid_units(value)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::standard()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ExpectedRoot {
    Properties,
    ExtensionList,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FrameKind {
    Properties,
    ExtensionList,
    Extension,
    Opaque,
}

#[derive(Debug)]
struct Frame {
    namespace: String,
    name: String,
    start: usize,
    kind: FrameKind,
    direct_child_count: usize,
}

#[derive(Debug)]
struct Attribute {
    namespace: String,
    name: String,
    value: String,
    value_range: Range<usize>,
    qualified: bool,
}

#[derive(Debug)]
struct SourceInfo {
    id: Option<String>,
    id_range: Option<Range<usize>>,
    root_open: Range<usize>,
    root_context: Vec<(String, String)>,
    extension: Option<Range<usize>>,
    extension_open: Option<Range<usize>>,
    extension_context: Option<Vec<(String, String)>>,
}

#[derive(Debug, Default)]
struct NamespaceState {
    bindings: Vec<NamespaceBinding>,
    scopes: Vec<(usize, usize)>,
    bytes: usize,
}

#[derive(Debug)]
struct NamespaceBinding {
    prefix: String,
    uri: String,
}

impl NamespaceState {
    fn seed(&mut self, bindings: &[(String, String)], limits: &Limits) -> Result<()> {
        check_limit(
            limits,
            "XML namespace binding count",
            bindings.len(),
            limits.attributes,
        )?;
        self.bindings
            .try_reserve(bindings.len())
            .map_err(|source| allocation("Custom Data XML inherited namespaces", source))?;
        let mut bytes = 0usize;
        for (prefix, uri) in bindings {
            let binding_bytes = prefix.len().checked_add(uri.len()).ok_or_else(|| {
                limit_for(
                    "Custom Data XML namespace bytes",
                    usize::MAX,
                    limits.namespace_bytes,
                )
            })?;
            bytes = bytes.checked_add(binding_bytes).ok_or_else(|| {
                limit_for(
                    "Custom Data XML namespace bytes",
                    usize::MAX,
                    limits.namespace_bytes,
                )
            })?;
            check_limit(
                limits,
                "Custom Data XML namespace bytes",
                bytes,
                limits.namespace_bytes,
            )?;
            self.bindings.push(NamespaceBinding {
                prefix: copy_string(prefix, "Custom Data XML namespace prefix", limits)?,
                uri: copy_string(uri, "Custom Data XML namespace URI", limits)?,
            });
        }
        self.bytes = bytes;
        Ok(())
    }

    fn snapshot(&self, limits: &Limits) -> Result<Vec<(String, String)>> {
        check_limit(
            limits,
            "XML namespace binding count",
            self.bindings.len(),
            limits.attributes,
        )?;
        let mut snapshot = Vec::new();
        snapshot
            .try_reserve(self.bindings.len())
            .map_err(|source| allocation("Custom Data XML namespace snapshot", source))?;
        for binding in &self.bindings {
            let prefix = copy_string(&binding.prefix, "Custom Data XML namespace prefix", limits)?;
            let uri = copy_string(&binding.uri, "Custom Data XML namespace URI", limits)?;
            snapshot.push((prefix, uri));
        }
        Ok(snapshot)
    }

    fn enter(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        limits: &Limits,
    ) -> Result<()> {
        let attribute_count = element.attributes().count();
        check_limit(
            limits,
            "XML attribute count",
            attribute_count,
            limits.attributes,
        )?;
        self.scopes
            .try_reserve(1)
            .map_err(|source| allocation("Custom Data XML namespace scopes", source))?;
        self.scopes.push((self.bindings.len(), self.bytes));
        for item in element.attributes().with_checks(true) {
            let item = item.map_err(xml_error)?;
            let key = item.key.as_ref();
            let Some(prefix) = namespace_declaration_prefix(key)? else {
                continue;
            };
            check_limit(
                limits,
                "XML string bytes",
                item.value.as_ref().len(),
                limits.string_bytes,
            )?;
            let value = item
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(xml_error)?
                .into_owned();
            validate_xml_characters(value.as_bytes())?;
            validate_namespace_declaration(key, &value)?;
            if self.bindings[self.scopes.last().map_or(0, |scope| scope.0)..]
                .iter()
                .any(|binding| binding.prefix == prefix)
            {
                return Err(invalid("duplicate XML namespace declaration"));
            }
            let binding_bytes = prefix.len().checked_add(value.len()).ok_or_else(|| {
                limit_for(
                    "Custom Data XML namespace bytes",
                    usize::MAX,
                    limits.namespace_bytes,
                )
            })?;
            let total_bytes = self.bytes.checked_add(binding_bytes).ok_or_else(|| {
                limit_for(
                    "Custom Data XML namespace bytes",
                    usize::MAX,
                    limits.namespace_bytes,
                )
            })?;
            check_limit(
                limits,
                "Custom Data XML namespace bytes",
                total_bytes,
                limits.namespace_bytes,
            )?;
            self.bindings
                .try_reserve(1)
                .map_err(|source| allocation("Custom Data XML namespace bindings", source))?;
            self.bindings.push(NamespaceBinding { prefix, uri: value });
            self.bytes = total_bytes;
        }
        Ok(())
    }

    fn leave(&mut self) {
        if let Some((mark, bytes)) = self.scopes.pop() {
            self.bindings.truncate(mark);
            self.bytes = bytes;
        }
    }

    fn resolve(&self, name: QName<'_>, allow_unknown: bool, attribute: bool) -> Result<String> {
        let (_, prefix) = name.decompose();
        let namespace = match prefix {
            None if attribute => String::new(),
            None => self.lookup("").map_or_else(String::new, ToOwned::to_owned),
            Some(prefix) if prefix.as_ref() == b"xml" => XML_NAMESPACE.to_owned(),
            Some(prefix) => {
                let prefix = std::str::from_utf8(prefix.as_ref()).map_err(xml_error)?;
                if let Some(uri) = self.lookup(prefix) {
                    uri.to_owned()
                } else if allow_unknown {
                    String::new()
                } else {
                    return Err(invalid(format!("unbound XML prefix '{prefix}'")));
                }
            },
        };
        Ok(namespace)
    }

    fn lookup(&self, prefix: &str) -> Option<&str> {
        self.bindings
            .iter()
            .rev()
            .find(|binding| binding.prefix == prefix)
            .map(|binding| binding.uri.as_str())
    }
}

fn decorate_extension_fragment(
    xml: &[u8],
    inherited_namespaces: &[(String, String)],
    limits: &Limits,
) -> Result<Vec<u8>> {
    let (opening, declared) = extension_root_opening(xml, limits)?;
    let mut additions = Vec::<(&str, &str)>::new();
    additions
        .try_reserve(inherited_namespaces.len())
        .map_err(|source| allocation("Custom Data namespace decoration", source))?;
    for (prefix, uri) in inherited_namespaces {
        if prefix == "xml" || declared.iter().any(|item| item == prefix) {
            continue;
        }
        additions.push((prefix.as_str(), uri.as_str()));
    }
    if additions.is_empty() {
        return copy_bounded(
            xml,
            limits.extension_xml_bytes,
            "extension XML bytes",
            limits,
        );
    }
    let mut added = 0usize;
    for (prefix, uri) in &additions {
        let prefix_len = if prefix.is_empty() {
            b" xmlns=\"".len()
        } else {
            b" xmlns:"
                .len()
                .checked_add(prefix.len())
                .and_then(|length| length.checked_add(b"=\"".len()))
                .ok_or_else(|| {
                    limit_for(
                        "extension XML bytes",
                        usize::MAX,
                        limits.extension_xml_bytes,
                    )
                })?
        };
        let uri_len = escaped_attribute_len(uri)?;
        added = added
            .checked_add(prefix_len)
            .and_then(|length| length.checked_add(uri_len))
            .and_then(|length| length.checked_add(1))
            .ok_or_else(|| {
                limit_for(
                    "extension XML bytes",
                    usize::MAX,
                    limits.extension_xml_bytes,
                )
            })?;
    }
    let total = xml.len().checked_add(added).ok_or_else(|| {
        limit_for(
            "extension XML bytes",
            usize::MAX,
            limits.extension_xml_bytes,
        )
    })?;
    check_limit(
        limits,
        "extension XML bytes",
        total,
        limits.extension_xml_bytes,
    )?;
    let insertion = if xml
        .get(opening.end.saturating_sub(2)..opening.end)
        .is_some_and(|bytes| bytes == b"/>")
    {
        opening.end.saturating_sub(2)
    } else {
        opening.end.saturating_sub(1)
    };
    if insertion < opening.start
        || xml.get(insertion) != Some(&b'>') && xml.get(insertion) != Some(&b'/')
    {
        return Err(invalid("invalid extension root opening range"));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(total)
        .map_err(|source| allocation("Custom Data decorated extension XML", source))?;
    output.extend_from_slice(&xml[..insertion]);
    for (prefix, uri) in additions {
        if prefix.is_empty() {
            output.extend_from_slice(b" xmlns=\"");
        } else {
            output.extend_from_slice(b" xmlns:");
            output.extend_from_slice(prefix.as_bytes());
            output.extend_from_slice(b"=\"");
        }
        append_escaped_attribute(&mut output, uri);
        output.push(b'"');
    }
    output.extend_from_slice(&xml[insertion..]);
    debug_assert_eq!(output.len(), total);
    Ok(output)
}

fn decorated_extension_matches(
    original: &[u8],
    candidate: &[u8],
    inherited_namespaces: &[(String, String)],
    limits: &Limits,
) -> Result<bool> {
    let decorated = decorate_extension_fragment(original, inherited_namespaces, limits)?;
    Ok(decorated == candidate)
}

fn extension_root_opening(xml: &[u8], limits: &Limits) -> Result<(Range<usize>, Vec<String>)> {
    let mut reader = NsReader::from_reader(xml);
    let mut events = 0usize;
    loop {
        let start = position(&reader)?;
        let event = reader.read_event().map_err(xml_error)?;
        count_event(&mut events, limits)?;
        let end = position(&reader)?;
        let element = match &event {
            Event::Start(element) | Event::Empty(element) => element,
            Event::Text(text) => {
                let decoded = text.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                if is_xml_space_only(&decoded) {
                    continue;
                }
                return Err(invalid("invalid extension XML root opening"));
            },
            Event::Comment(_) | Event::PI(_) => continue,
            Event::Eof => return Err(invalid("missing extension XML root")),
            _ => return Err(invalid("invalid extension XML root opening")),
        };
        let attribute_count = element.attributes().count();
        check_limit(
            limits,
            "XML attribute count",
            attribute_count,
            limits.attributes,
        )?;
        let mut declared = Vec::new();
        declared
            .try_reserve(attribute_count)
            .map_err(|source| allocation("Custom Data extension declarations", source))?;
        for item in element.attributes().with_checks(true) {
            let item = item.map_err(xml_error)?;
            if let Some(prefix) = namespace_declaration_prefix(item.key.as_ref())? {
                declared.push(prefix);
            }
        }
        return Ok((start..end, declared));
    }
}

fn is_xml_space_only(value: &str) -> bool {
    value
        .chars()
        .all(|character| matches!(character, '\u{9}' | '\u{a}' | '\u{d}' | '\u{20}'))
}

fn count_event(events: &mut usize, limits: &Limits) -> Result<()> {
    *events = events
        .checked_add(1)
        .ok_or_else(|| limit_for("XML event count", usize::MAX, limits.events))?;
    check_limit(limits, "XML event count", *events, limits.events)
}

fn escaped_attribute_len(value: &str) -> Result<usize> {
    let mut length = 0usize;
    for character in value.chars() {
        let added = match character {
            '&' => 5,
            '<' => 4,
            '>' => 4,
            '"' | '\'' => 6,
            _ => character.len_utf8(),
        };
        length = length
            .checked_add(added)
            .ok_or_else(|| limit_for("escaped XML attribute length", usize::MAX, usize::MAX))?;
    }
    Ok(length)
}

fn append_escaped_attribute(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\'' => output.extend_from_slice(b"&apos;"),
            _ => {
                let mut encoded = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut encoded).as_bytes());
            },
        }
    }
}

/// Parse one Custom Data Properties part.
///
/// The returned extension bytes retain the exact source span of the direct
/// `extLst`, including prefixes, declarations on descendants, comments,
/// processing instructions, CDATA, entities, whitespace, and attribute order.
/// When the fragment relies on declarations from the datastore root, the
/// returned standalone fragment receives equivalent declarations on its own
/// root.  The package host retains the original document separately for
/// source-bound no-op preservation.
pub fn parse_properties(xml: &[u8]) -> Result<Properties> {
    parse_properties_with_limits(xml, &Limits::standard())
}

/// Parse one Custom Data Properties part under an explicit resource profile.
pub fn parse_properties_with_limits(xml: &[u8], limits: &Limits) -> Result<Properties> {
    let limits = limits.validate()?;
    let info = parse_document(xml, ExpectedRoot::Properties, false, &limits)?;
    let id = info
        .id
        .ok_or_else(|| invalid("Custom Data Properties root is missing its required id"))?;
    let extension_list = if let Some(range) = info.extension {
        let extension = copy_bounded(
            &xml[range],
            limits.extension_xml_bytes,
            "extension XML bytes",
            &limits,
        )?;
        let extension = if let Some(context) = info.extension_context.as_deref() {
            if validate_extension_fragment_with_mode(&extension, false, &limits).is_err() {
                decorate_extension_fragment(&extension, context, &limits)?
            } else {
                extension
            }
        } else {
            extension
        };
        Some(ExtensionList { xml: extension })
    } else {
        None
    };
    let value = Properties { id, extension_list };
    validate_properties(&value, false, &limits)?;
    Ok(value)
}

/// Deterministically serialize a Custom Data Properties part.
///
/// The root wrapper is compact and stable.  A supplied extension list is
/// validated and then copied byte-for-byte into the wrapper.
pub fn write_properties(value: &Properties) -> Result<Vec<u8>> {
    write_properties_with_limits(value, &Limits::standard())
}

/// Deterministically serialize Custom Data Properties under an explicit
/// resource profile.
pub fn write_properties_with_limits(value: &Properties, limits: &Limits) -> Result<Vec<u8>> {
    let limits = limits.validate()?;
    validate_properties(value, false, &limits)?;
    let id_len = escaped_xstring_len(&value.id, &limits)?;
    let root_name = b"<x14:datastoreItem xmlns:x14=\"".as_slice();
    let root_namespace = X14.as_bytes();
    let root_id = b"\" id=\"";
    let (root_suffix, root_end) = if value.extension_list.is_some() {
        (b"\">".as_slice(), b"</".as_slice())
    } else {
        (b"\"/>".as_slice(), b"".as_slice())
    };
    let closing_prefix = b"x14:datastoreItem>".as_slice();
    let total = root_name
        .len()
        .checked_add(root_namespace.len())
        .and_then(|n| n.checked_add(root_id.len()))
        .and_then(|n| n.checked_add(id_len))
        .and_then(|n| n.checked_add(root_suffix.len()))
        .and_then(|n| {
            n.checked_add(
                value
                    .extension_list
                    .as_ref()
                    .map_or(0, |extension| extension.xml.len()),
            )
        })
        .and_then(|n| {
            n.checked_add(root_end.len()).and_then(|n| {
                n.checked_add(if value.extension_list.is_some() {
                    closing_prefix.len()
                } else {
                    0
                })
            })
        })
        .ok_or_else(|| {
            limit_for(
                "serialized properties XML bytes",
                usize::MAX,
                limits.properties_xml_bytes,
            )
        })?;
    check_limit(
        &limits,
        "serialized properties XML bytes",
        total,
        limits.properties_xml_bytes,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(total)
        .map_err(|source| allocation("custom-data properties XML", source))?;
    output.extend_from_slice(root_name);
    output.extend_from_slice(root_namespace);
    output.extend_from_slice(root_id);
    append_escaped_xstring(&mut output, &value.id);
    output.extend_from_slice(root_suffix);
    if let Some(extension) = &value.extension_list {
        output.extend_from_slice(&extension.xml);
    }
    if value.extension_list.is_some() {
        output.extend_from_slice(root_end);
        output.extend_from_slice(closing_prefix);
    }
    debug_assert_eq!(output.len(), total);
    Ok(output)
}

/// Replace only the required root `id` value in a validated source part.
#[doc(hidden)]
pub fn rewrite_id(source: &litchi_opc::OwnedXmlPart, id: &str) -> Result<litchi_opc::OwnedXmlPart> {
    rewrite_id_with_limits(source, id, &Limits::standard())
}

/// Rewrite one source-bound UID under an explicit resource profile.
#[doc(hidden)]
pub fn rewrite_id_with_limits(
    source: &litchi_opc::OwnedXmlPart,
    id: &str,
    limits: &Limits,
) -> Result<litchi_opc::OwnedXmlPart> {
    rewrite_id_with_limits_and_output(source, id, limits, limits.properties_xml_bytes)
}

/// Rewrite a retained UID with independent source and replacement-output limits.
/// An exact no-op shares the source token without constructing output.
#[doc(hidden)]
pub fn rewrite_id_with_limits_and_output(
    source: &litchi_opc::OwnedXmlPart,
    id: &str,
    limits: &Limits,
    max_output_bytes: usize,
) -> Result<litchi_opc::OwnedXmlPart> {
    let limits = limits.validate()?;
    let maximum = max_output_bytes.min(limits.properties_xml_bytes);
    let info = parse_document(source.bytes(), ExpectedRoot::Properties, false, &limits)?;
    validate_id(id, &limits)?;
    if info.id.as_deref() == Some(id) {
        // A semantic no-op must retain the source token, including entity
        // spelling, quote style, and surrounding whitespace.
        return Ok(source.clone());
    }
    let range = info
        .id_range
        .ok_or_else(|| invalid("Custom Data storage ID attribute is absent"))?;
    let replacement_len = escaped_xstring_len(id, &limits)?;
    let output_len = source
        .bytes()
        .len()
        .checked_sub(range.len())
        .and_then(|length| length.checked_add(replacement_len))
        .ok_or_else(|| limit_for("properties XML output", usize::MAX, maximum))?;
    check_limit(&limits, "properties XML output", output_len, maximum)?;
    let replacement = try_escaped_xstring(id, &limits)?;
    let result = source.replace_attributes(&[(range, replacement)])?;
    check_limit(
        &limits,
        "properties XML output",
        result.bytes().len(),
        maximum,
    )?;
    Ok(result)
}

/// Replace or remove only the direct `extLst` child of a retained source root.
#[doc(hidden)]
pub fn rewrite_extension_list(
    source: &litchi_opc::OwnedXmlPart,
    extension: Option<&ExtensionList>,
) -> Result<litchi_opc::OwnedXmlPart> {
    rewrite_extension_list_with_limits(source, extension, &Limits::standard())
}

/// Replace or remove one source-bound extension list under an explicit
/// resource profile.
#[doc(hidden)]
pub fn rewrite_extension_list_with_limits(
    source: &litchi_opc::OwnedXmlPart,
    extension: Option<&ExtensionList>,
    limits: &Limits,
) -> Result<litchi_opc::OwnedXmlPart> {
    rewrite_extension_list_with_limits_and_output(
        source,
        extension,
        limits,
        limits.properties_xml_bytes,
    )
}

/// Rewrite a retained extension with independent source and replacement-output
/// limits. Exact no-ops retain their original source token.
#[doc(hidden)]
pub fn rewrite_extension_list_with_limits_and_output(
    source: &litchi_opc::OwnedXmlPart,
    extension: Option<&ExtensionList>,
    limits: &Limits,
    max_output_bytes: usize,
) -> Result<litchi_opc::OwnedXmlPart> {
    let limits = limits.validate()?;
    let maximum = max_output_bytes.min(limits.properties_xml_bytes);
    let info = parse_document(source.bytes(), ExpectedRoot::Properties, false, &limits)?;
    let Some(extension) = extension else {
        return if let Some(range) = info.extension_open {
            let removed = info.extension.as_ref().map_or(0, Range::len);
            let output_len = source
                .bytes()
                .len()
                .checked_sub(removed)
                .ok_or_else(|| invalid("Custom Data extension range exceeds source"))?;
            check_limit(&limits, "properties XML output", output_len, maximum)?;
            let result = source.remove_element(range)?;
            check_limit(
                &limits,
                "properties XML output",
                result.bytes().len(),
                maximum,
            )?;
            Ok(result)
        } else {
            Ok(source.clone())
        };
    };
    let inherited = info
        .extension_context
        .as_deref()
        .unwrap_or(info.root_context.as_slice());
    validate_extension_fragment_with_context(&extension.xml, inherited, &limits)?;
    if let Some(range) = info.extension.as_ref() {
        let original = &source.bytes()[range.clone()];
        if original == extension.xml.as_slice()
            || decorated_extension_matches(original, &extension.xml, inherited, &limits)?
        {
            return Ok(source.clone());
        }
    }
    // Expanding an empty root replaces its slash with a full closing tag.
    let root_tag = &source.bytes()[info.root_open.clone()];
    let expansion = if info.extension.is_none() && root_tag.ends_with(b"/>") {
        root_tag
            .iter()
            .skip(1)
            .take_while(|byte| !byte.is_ascii_whitespace() && **byte != b'/' && **byte != b'>')
            .count()
            .checked_add(2)
            .ok_or_else(|| invalid("Custom Data root expansion overflow"))?
    } else {
        0
    };
    let output_len = source
        .bytes()
        .len()
        .checked_sub(info.extension.as_ref().map_or(0, Range::len))
        .and_then(|n| n.checked_add(extension.xml.len()))
        .and_then(|n| n.checked_add(expansion))
        .ok_or_else(|| limit_for("properties XML output", usize::MAX, maximum))?;
    check_limit(&limits, "properties XML output", output_len, maximum)?;
    match info.extension_open {
        Some(range) => {
            let result = source.replace_element(range, &extension.xml)?;
            check_limit(
                &limits,
                "properties XML output",
                result.bytes().len(),
                maximum,
            )?;
            Ok(result)
        },
        None => {
            let result = source.append_element(info.root_open, &extension.xml)?;
            check_limit(
                &limits,
                "properties XML output",
                result.bytes().len(),
                maximum,
            )?;
            Ok(result)
        },
    }
}

/// Validate and retain an extension projection for semantic comparison.
///
/// No canonicalization is performed: lexical source bytes are part of the
/// preservation contract.
#[doc(hidden)]
pub fn canonical_extension(extension: Option<&ExtensionList>) -> Result<Option<Vec<u8>>> {
    canonical_extension_with_limits(extension, &Limits::standard())
}

/// Validate and retain an extension projection under an explicit profile.
#[doc(hidden)]
pub fn canonical_extension_with_limits(
    extension: Option<&ExtensionList>,
    limits: &Limits,
) -> Result<Option<Vec<u8>>> {
    let limits = limits.validate()?;
    extension
        .map(|extension| {
            validate_extension_fragment_with_mode(&extension.xml, false, &limits)?;
            copy_bounded(
                &extension.xml,
                limits.extension_xml_bytes,
                "extension XML bytes",
                &limits,
            )
        })
        .transpose()
}

/// Validate a parsed/staged properties value for a source-bound host.
///
/// The value passed through this API is the standalone-ready projection.  A
/// parsed source fragment carries any bindings it needs on its own root; a
/// staged fragment that relies on a host root must be validated through the
/// source-context rewrite path instead.
#[doc(hidden)]
pub fn validate_source_properties(value: &Properties) -> Result<()> {
    validate_source_properties_with_limits(value, &Limits::standard())
}

/// Validate source-bound properties under an explicit resource profile.
#[doc(hidden)]
pub fn validate_source_properties_with_limits(value: &Properties, limits: &Limits) -> Result<()> {
    let limits = limits.validate()?;
    validate_properties(value, false, &limits)
}

/// Validate that a source belongs to a SpreadsheetML workbook part.
pub fn validate_workbook_root(xml: &[u8]) -> Result<()> {
    validate_workbook_root_with_limits(xml, &Limits::standard())
}

/// Validate a workbook root under an explicit resource profile.
pub fn validate_workbook_root_with_limits(xml: &[u8], limits: &Limits) -> Result<()> {
    let limits = limits.validate()?;
    let (namespace, name) = scan_generic_document(xml, &limits)?;
    if name == "workbook" && matches!(namespace.as_str(), SML | STRICT_SML) {
        Ok(())
    } else {
        Err(invalid(
            "Custom Data Properties source must be a workbook part",
        ))
    }
}

fn parse_document(
    xml: &[u8],
    expected: ExpectedRoot,
    allow_inherited_namespaces: bool,
    limits: &Limits,
) -> Result<SourceInfo> {
    parse_document_with_context(xml, expected, allow_inherited_namespaces, &[], limits)
}

fn parse_document_with_context(
    xml: &[u8],
    expected: ExpectedRoot,
    allow_inherited_namespaces: bool,
    inherited_namespaces: &[(String, String)],
    limits: &Limits,
) -> Result<SourceInfo> {
    check_limit(
        limits,
        "properties XML bytes",
        xml.len(),
        limits.properties_xml_bytes,
    )?;
    validate_xml_characters(xml)?;
    let mut reader = NsReader::from_reader(xml);
    let mut namespaces = NamespaceState::default();
    namespaces.seed(inherited_namespaces, limits)?;
    let mut stack = Vec::<Frame>::new();
    let mut root_seen = false;
    let mut root_closed = false;
    let mut nodes = 0usize;
    let mut events = 0usize;
    let mut extension_count = 0usize;
    let mut id = None;
    let mut id_range = None;
    let mut root_open = None;
    let mut root_context = None;
    let mut extension = None;
    let mut extension_open = None;
    let mut extension_context = None;
    let mut saw_declaration = false;
    let mut saw_prolog_content = false;

    loop {
        let start = position(&reader)?;
        let event = reader.read_event().map_err(xml_error)?;
        count_event(&mut events, limits)?;
        let end = position(&reader)?;
        match event {
            Event::Start(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| limit_for("XML node count", usize::MAX, limits.nodes))?;
                check_structure(nodes, stack.len(), limits)?;
                let depth = stack.len();
                let insertion_context = if expected == ExpectedRoot::Properties && depth == 1 {
                    Some(namespaces.snapshot(limits)?)
                } else {
                    None
                };
                namespaces.enter(&reader, &element, limits)?;
                let attributes = element_attributes(
                    &reader,
                    &element,
                    xml,
                    &namespaces,
                    allow_inherited_namespaces,
                    limits,
                )?;
                let (namespace, name) = element_name(
                    &namespaces,
                    element.name(),
                    allow_inherited_namespaces,
                    limits,
                )?;
                let kind = if depth == 0 {
                    admit_root(
                        expected,
                        &namespace,
                        &name,
                        &attributes,
                        allow_inherited_namespaces,
                        limits,
                        &mut id,
                        &mut id_range,
                    )?;
                    if root_seen || root_closed {
                        return Err(invalid("multiple XML roots"));
                    }
                    root_seen = true;
                    root_open = Some(start..end);
                    if expected == ExpectedRoot::Properties {
                        root_context = Some(namespaces.snapshot(limits)?);
                    }
                    if expected == ExpectedRoot::Properties {
                        FrameKind::Properties
                    } else {
                        FrameKind::ExtensionList
                    }
                } else if root_closed {
                    return Err(invalid("element appears after the XML root"));
                } else if depth == 1 {
                    match expected {
                        ExpectedRoot::Properties => {
                            if namespace == X14 && name == "extLst" {
                                validate_ext_list_attributes(&attributes)?;
                                extension_count = extension_count
                                    .checked_add(1)
                                    .ok_or_else(|| invalid("direct extLst count overflow"))?;
                                if extension_count > 1 {
                                    return Err(invalid(
                                        "datastoreItem permits at most one direct extLst",
                                    ));
                                }
                                extension_open = Some(start..end);
                                extension_context = insertion_context;
                                FrameKind::ExtensionList
                            } else {
                                return Err(invalid(
                                    "datastoreItem permits only one direct X14 extLst child",
                                ));
                            }
                        },
                        ExpectedRoot::ExtensionList => {
                            if is_extension_element(&namespace, &name) {
                                validate_extension_attributes(&attributes, limits)?;
                                FrameKind::Extension
                            } else {
                                return Err(invalid(
                                    "extLst permits only direct core SpreadsheetML ext children",
                                ));
                            }
                        },
                    }
                } else if depth == 2
                    && stack
                        .last()
                        .is_some_and(|frame| frame.kind == FrameKind::ExtensionList)
                {
                    if is_extension_element(&namespace, &name) {
                        validate_extension_attributes(&attributes, limits)?;
                        FrameKind::Extension
                    } else {
                        return Err(invalid(
                            "extLst permits only direct core SpreadsheetML ext children",
                        ));
                    }
                } else if stack
                    .last()
                    .is_some_and(|frame| frame.kind == FrameKind::Extension)
                {
                    let parent = stack
                        .last_mut()
                        .ok_or_else(|| invalid("missing direct extension parent"))?;
                    parent.direct_child_count = parent
                        .direct_child_count
                        .checked_add(1)
                        .ok_or_else(|| invalid("direct extension child count overflow"))?;
                    if parent.direct_child_count > 1 {
                        return Err(invalid(
                            "core SpreadsheetML ext permits exactly one wildcard child",
                        ));
                    }
                    FrameKind::Opaque
                } else {
                    FrameKind::Opaque
                };
                stack
                    .try_reserve(1)
                    .map_err(|source| allocation("Custom Data XML frames", source))?;
                stack.push(Frame {
                    namespace,
                    name,
                    start,
                    kind,
                    direct_child_count: 0,
                });
            },
            Event::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| limit_for("XML node count", usize::MAX, limits.nodes))?;
                check_structure(nodes, stack.len(), limits)?;
                let depth = stack.len();
                let insertion_context = if expected == ExpectedRoot::Properties && depth == 1 {
                    Some(namespaces.snapshot(limits)?)
                } else {
                    None
                };
                namespaces.enter(&reader, &element, limits)?;
                let attributes = element_attributes(
                    &reader,
                    &element,
                    xml,
                    &namespaces,
                    allow_inherited_namespaces,
                    limits,
                )?;
                let (namespace, name) = element_name(
                    &namespaces,
                    element.name(),
                    allow_inherited_namespaces,
                    limits,
                )?;
                if depth == 0 {
                    admit_root(
                        expected,
                        &namespace,
                        &name,
                        &attributes,
                        allow_inherited_namespaces,
                        limits,
                        &mut id,
                        &mut id_range,
                    )?;
                    if root_seen || root_closed {
                        return Err(invalid("multiple XML roots"));
                    }
                    root_seen = true;
                    root_closed = true;
                    root_open = Some(start..end);
                    if expected == ExpectedRoot::Properties {
                        root_context = Some(namespaces.snapshot(limits)?);
                    }
                } else if root_closed {
                    return Err(invalid("element appears after the XML root"));
                } else if depth == 1 {
                    match expected {
                        ExpectedRoot::Properties if namespace == X14 && name == "extLst" => {
                            validate_ext_list_attributes(&attributes)?;
                            extension_count = extension_count
                                .checked_add(1)
                                .ok_or_else(|| invalid("direct extLst count overflow"))?;
                            if extension_count > 1 {
                                return Err(invalid(
                                    "datastoreItem permits at most one direct extLst",
                                ));
                            }
                            extension_open = Some(start..end);
                            extension = Some(start..end);
                            check_limit(
                                limits,
                                "extension XML bytes",
                                end.checked_sub(start)
                                    .ok_or_else(|| invalid("invalid extension XML source range"))?,
                                limits.extension_xml_bytes,
                            )?;
                            extension_context = insertion_context;
                        },
                        ExpectedRoot::ExtensionList if is_extension_element(&namespace, &name) => {
                            validate_extension_attributes(&attributes, limits)?;
                            return Err(invalid(
                                "core SpreadsheetML ext requires exactly one wildcard child",
                            ));
                        },
                        ExpectedRoot::Properties => {
                            return Err(invalid(
                                "datastoreItem permits only one direct X14 extLst child",
                            ));
                        },
                        ExpectedRoot::ExtensionList => {
                            return Err(invalid(
                                "extLst permits only direct core SpreadsheetML ext children",
                            ));
                        },
                    }
                } else if depth == 2
                    && stack
                        .last()
                        .is_some_and(|frame| frame.kind == FrameKind::ExtensionList)
                {
                    if is_extension_element(&namespace, &name) {
                        validate_extension_attributes(&attributes, limits)?;
                        return Err(invalid(
                            "core SpreadsheetML ext requires exactly one wildcard child",
                        ));
                    } else {
                        return Err(invalid(
                            "extLst permits only direct core SpreadsheetML ext children",
                        ));
                    }
                } else if stack
                    .last()
                    .is_some_and(|frame| frame.kind == FrameKind::Extension)
                {
                    let parent = stack
                        .last_mut()
                        .ok_or_else(|| invalid("missing direct extension parent"))?;
                    parent.direct_child_count = parent
                        .direct_child_count
                        .checked_add(1)
                        .ok_or_else(|| invalid("direct extension child count overflow"))?;
                    if parent.direct_child_count > 1 {
                        return Err(invalid(
                            "core SpreadsheetML ext permits exactly one wildcard child",
                        ));
                    }
                }
                namespaces.leave();
            },
            Event::End(element) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML closing element"))?;
                let (end_namespace, end_name) = element_name(
                    &namespaces,
                    element.name(),
                    allow_inherited_namespaces,
                    limits,
                )?;
                if frame.namespace != end_namespace || frame.name != end_name {
                    return Err(invalid("XML closing element does not match its opener"));
                }
                if frame.kind == FrameKind::Extension && frame.direct_child_count != 1 {
                    return Err(invalid(
                        "core SpreadsheetML ext requires exactly one wildcard child",
                    ));
                }
                if frame.kind == FrameKind::ExtensionList && expected == ExpectedRoot::Properties {
                    check_limit(
                        limits,
                        "extension XML bytes",
                        end.checked_sub(frame.start)
                            .ok_or_else(|| invalid("invalid extension XML source range"))?,
                        limits.extension_xml_bytes,
                    )?;
                    extension = Some(frame.start..end);
                }
                if frame.kind == FrameKind::Properties
                    || (frame.kind == FrameKind::ExtensionList
                        && expected == ExpectedRoot::ExtensionList)
                {
                    root_closed = true;
                }
                namespaces.leave();
            },
            Event::Text(text) => {
                check_limit(
                    limits,
                    "XML string bytes",
                    text.as_ref().len(),
                    limits.string_bytes,
                )?;
                if text.as_ref().windows(3).any(|window| window == b"]]>") {
                    return Err(invalid("raw ]]> is not permitted in XML text"));
                }
                let decoded = text.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                if stack.last().is_none() {
                    saw_prolog_content = true;
                    if !is_xml_space_only(&decoded) {
                        return Err(invalid("text outside XML root"));
                    }
                } else if stack.last().is_some_and(|frame| {
                    matches!(
                        frame.kind,
                        FrameKind::Properties | FrameKind::ExtensionList | FrameKind::Extension
                    )
                }) && !is_xml_space_only(&decoded)
                {
                    return Err(invalid(
                        "recognized Custom Data containers and direct ext elements permit only XML whitespace text",
                    ));
                }
            },
            Event::GeneralRef(reference) => {
                if stack.last().is_some_and(|frame| {
                    matches!(
                        frame.kind,
                        FrameKind::Properties | FrameKind::ExtensionList | FrameKind::Extension
                    )
                }) {
                    return Err(invalid(
                        "recognized Custom Data containers and direct ext elements do not permit direct entity text",
                    ));
                }
                if stack.is_empty() {
                    return Err(invalid("entity outside XML root"));
                }
                let _ = crate::xml::decode_xml_reference(&reference)?;
            },
            Event::CData(_) => {
                if stack.is_empty() {
                    return Err(invalid("CDATA outside XML root"));
                }
                if stack.last().is_some_and(|frame| {
                    matches!(
                        frame.kind,
                        FrameKind::Properties | FrameKind::ExtensionList | FrameKind::Extension
                    )
                }) {
                    return Err(invalid(
                        "recognized Custom Data containers and direct ext elements do not permit direct CDATA",
                    ));
                }
            },
            Event::DocType(_) => return Err(invalid("DTDs are rejected")),
            Event::Decl(_) => {
                if expected == ExpectedRoot::ExtensionList
                    || saw_declaration
                    || root_seen
                    || !stack.is_empty()
                    || saw_prolog_content
                {
                    return Err(invalid("XML declaration is not at the document prolog"));
                }
                saw_declaration = true;
            },
            Event::Comment(_) | Event::PI(_) => {
                if !root_seen {
                    saw_prolog_content = true;
                }
            },
            Event::Eof => break,
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated Custom Data Properties XML"));
    }
    if !root_seen || !root_closed {
        return Err(invalid(
            "missing or unterminated Custom Data Properties XML root",
        ));
    }
    if expected == ExpectedRoot::Properties && id.is_none() {
        return Err(invalid(
            "Custom Data Properties root is missing its required id",
        ));
    }
    Ok(SourceInfo {
        id,
        id_range,
        root_open: root_open.ok_or_else(|| invalid("missing XML root range"))?,
        root_context: root_context.unwrap_or_default(),
        extension,
        extension_open,
        extension_context,
    })
}

fn scan_generic_document(xml: &[u8], limits: &Limits) -> Result<(String, String)> {
    check_limit(
        limits,
        "properties XML bytes",
        xml.len(),
        limits.properties_xml_bytes,
    )?;
    validate_xml_characters(xml)?;
    let mut reader = NsReader::from_reader(xml);
    let mut namespaces = NamespaceState::default();
    let mut stack = Vec::<(String, String)>::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut events = 0usize;
    let mut declaration = false;
    let mut prolog_content = false;
    loop {
        let event = reader.read_event().map_err(xml_error)?;
        count_event(&mut events, limits)?;
        match event {
            Event::Start(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| limit_for("XML node count", usize::MAX, limits.nodes))?;
                check_structure(nodes, stack.len(), limits)?;
                namespaces.enter(&reader, &element, limits)?;
                let _ = element_attributes(&reader, &element, xml, &namespaces, false, limits)?;
                let name = element_name(&namespaces, element.name(), false, limits)?;
                if stack.is_empty() {
                    if root.is_some() {
                        return Err(invalid("multiple XML roots"));
                    }
                    root = Some(name.clone());
                }
                stack
                    .try_reserve(1)
                    .map_err(|source| allocation("Custom Data workbook XML frames", source))?;
                stack.push(name);
            },
            Event::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| limit_for("XML node count", usize::MAX, limits.nodes))?;
                check_structure(nodes, stack.len(), limits)?;
                namespaces.enter(&reader, &element, limits)?;
                let _ = element_attributes(&reader, &element, xml, &namespaces, false, limits)?;
                if stack.is_empty() {
                    if root.is_some() {
                        return Err(invalid("multiple XML roots"));
                    }
                    root = Some(element_name(&namespaces, element.name(), false, limits)?);
                }
                namespaces.leave();
            },
            Event::End(element) => {
                let frame = stack
                    .last()
                    .ok_or_else(|| invalid("unexpected XML closing element"))?;
                let (namespace, name) = element_name(&namespaces, element.name(), false, limits)?;
                if frame.0 != namespace || frame.1 != name {
                    return Err(invalid("XML closing element does not match its opener"));
                }
                stack.pop();
                namespaces.leave();
            },
            Event::Text(text) => {
                check_limit(
                    limits,
                    "XML string bytes",
                    text.as_ref().len(),
                    limits.string_bytes,
                )?;
                let value = text.decode().map_err(xml_error)?;
                let value = quick_xml::escape::unescape(&value).map_err(xml_error)?;
                if stack.is_empty() {
                    prolog_content = true;
                    if !is_xml_space_only(&value) {
                        return Err(invalid("text outside XML root"));
                    }
                }
            },
            Event::GeneralRef(reference) => {
                if stack.is_empty() {
                    return Err(invalid("entity outside XML root"));
                }
                let _ = crate::xml::decode_xml_reference(&reference)?;
            },
            Event::DocType(_) => return Err(invalid("DTDs are rejected")),
            Event::Decl(_) => {
                if declaration || root.is_some() || !stack.is_empty() || prolog_content {
                    return Err(invalid("XML declaration is not at the document prolog"));
                }
                declaration = true;
            },
            Event::CData(_) => {
                if stack.is_empty() {
                    return Err(invalid("CDATA outside XML root"));
                }
            },
            Event::Comment(_) | Event::PI(_) => {
                if root.is_none() {
                    prolog_content = true;
                }
            },
            Event::Eof => break,
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated XML root"));
    }
    root.ok_or_else(|| invalid("missing XML root"))
}

fn admit_root(
    expected: ExpectedRoot,
    namespace: &str,
    name: &str,
    attributes: &[Attribute],
    allow_inherited_namespaces: bool,
    limits: &Limits,
    id: &mut Option<String>,
    id_range: &mut Option<Range<usize>>,
) -> Result<()> {
    let expected_name = match expected {
        ExpectedRoot::Properties => "datastoreItem",
        ExpectedRoot::ExtensionList => "extLst",
    };
    if (namespace != X14 && !(allow_inherited_namespaces && namespace.is_empty()))
        || name != expected_name
    {
        return Err(invalid(format!(
            "expected {{{X14}}}{expected_name}, got {{{namespace}}}{name}"
        )));
    }
    match expected {
        ExpectedRoot::Properties => {
            let mut found = None;
            for attribute in attributes {
                if !attribute.qualified && attribute.namespace.is_empty() && attribute.name == "id"
                {
                    if found.is_some() {
                        return Err(invalid("duplicate datastoreItem id attribute"));
                    }
                    let decoded = decode_spreadsheet_text(&attribute.value, limits)?;
                    validate_id(&decoded, limits)?;
                    found = Some(decoded);
                    *id_range = Some(attribute.value_range.clone());
                } else {
                    return Err(invalid("unexpected datastoreItem attribute"));
                }
            }
            *id = found;
            if id.is_none() {
                return Err(invalid("datastoreItem id attribute is required"));
            }
        },
        ExpectedRoot::ExtensionList => validate_ext_list_attributes(attributes)?,
    }
    Ok(())
}

fn validate_ext_list_attributes(attributes: &[Attribute]) -> Result<()> {
    if attributes.is_empty() {
        return Ok(());
    }
    Err(invalid(
        "x14 extLst does not permit non-namespace attributes",
    ))
}

fn validate_extension_attributes(attributes: &[Attribute], limits: &Limits) -> Result<()> {
    let mut uri = false;
    for attribute in attributes {
        if !attribute.qualified && attribute.namespace.is_empty() && attribute.name == "uri" {
            if uri {
                return Err(invalid("duplicate core SpreadsheetML ext uri attribute"));
            }
            bounded_string(&attribute.value, "extension URI", limits)?;
            uri = true;
        } else {
            return Err(invalid(
                "core SpreadsheetML ext has an unexpected attribute",
            ));
        }
    }
    Ok(())
}

fn is_extension_element(namespace: &str, name: &str) -> bool {
    name == "ext" && namespace == SML
}

fn element_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    xml: &[u8],
    namespaces: &NamespaceState,
    allow_inherited_namespaces: bool,
    limits: &Limits,
) -> Result<Vec<Attribute>> {
    let attribute_count = element.attributes().count();
    check_limit(
        limits,
        "XML attribute count",
        attribute_count,
        limits.attributes,
    )?;
    let mut attributes = Vec::new();
    attributes
        .try_reserve(attribute_count)
        .map_err(|source| allocation("Custom Data XML attributes", source))?;
    let mut seen = HashSet::<(String, String)>::new();
    seen.try_reserve(element.attributes().count())
        .map_err(|source| allocation("Custom Data XML attribute identities", source))?;
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let key = item.key.as_ref();
        check_limit(
            limits,
            "XML string bytes",
            item.value.as_ref().len(),
            limits.string_bytes,
        )?;
        if item.value.as_ref().contains(&b'<') {
            return Err(invalid(
                "raw '<' is not permitted in an XML attribute value",
            ));
        }
        let value = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?
            .into_owned();
        validate_xml_characters(value.as_bytes())?;
        if namespace_declaration_prefix(key)?.is_some() {
            validate_namespace_declaration(key, &value)?;
            continue;
        }
        check_limit(limits, "XML string bytes", key.len(), limits.string_bytes)?;
        let namespace = namespaces.resolve(item.key, allow_inherited_namespaces, true)?;
        let name = std::str::from_utf8(item.key.local_name().as_ref())
            .map_err(xml_error)?
            .to_owned();
        if !seen.insert((namespace.clone(), name.clone())) {
            return Err(invalid("duplicate expanded XML attribute"));
        }
        let raw = item.value.as_ref();
        let value_range = value_span(xml, raw)?;
        attributes.push(Attribute {
            namespace,
            name,
            value,
            value_range,
            qualified: item.key.prefix().is_some(),
        });
    }
    Ok(attributes)
}

fn validate_namespace_declaration(key: &[u8], value: &str) -> Result<()> {
    let prefix = if key == b"xmlns" {
        None
    } else {
        Some(
            std::str::from_utf8(key.strip_prefix(b"xmlns:").unwrap_or_default())
                .map_err(xml_error)?,
        )
    };
    if prefix == Some("xmlns") || value == XMLNS_NAMESPACE {
        return Err(invalid("reserved XMLNS namespace binding"));
    }
    if prefix == Some("xml") && value != XML_NAMESPACE {
        return Err(invalid("xml prefix must bind to its reserved namespace"));
    }
    if value == XML_NAMESPACE && prefix != Some("xml") {
        return Err(invalid("only xml may bind the reserved XML namespace"));
    }
    if prefix.is_some_and(|prefix| prefix != "xml") && value.is_empty() {
        return Err(invalid("prefixed XML namespace binding may not be empty"));
    }
    Ok(())
}

fn namespace_declaration_prefix(key: &[u8]) -> Result<Option<String>> {
    if key == b"xmlns" {
        return Ok(Some(String::new()));
    }
    let Some(prefix) = key.strip_prefix(b"xmlns:") else {
        return Ok(None);
    };
    if prefix.is_empty() {
        return Err(invalid("XML namespace declaration prefix is empty"));
    }
    let prefix = std::str::from_utf8(prefix).map_err(xml_error)?;
    if !crate::xml_name::is_ncname(prefix) {
        return Err(invalid("XML namespace declaration prefix is not an NCName"));
    }
    Ok(Some(prefix.to_owned()))
}

fn element_name(
    namespaces: &NamespaceState,
    name: QName<'_>,
    allow_inherited_namespaces: bool,
    limits: &Limits,
) -> Result<(String, String)> {
    let namespace = namespaces.resolve(name, allow_inherited_namespaces, false)?;
    let local_name = name.local_name();
    check_limit(
        limits,
        "XML string bytes",
        local_name.as_ref().len(),
        limits.string_bytes,
    )?;
    let local = std::str::from_utf8(local_name.as_ref()).map_err(xml_error)?;
    Ok((namespace, local.to_owned()))
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|error| invalid(format!("XML source offset does not fit usize: {error}")))
}

fn check_structure(nodes: usize, depth: usize, limits: &Limits) -> Result<()> {
    if nodes > limits.nodes {
        return Err(limit_for("XML node count", nodes, limits.nodes));
    }
    if depth >= limits.depth {
        return Err(limit_for(
            "XML depth",
            depth.saturating_add(1),
            limits.depth,
        ));
    }
    Ok(())
}

fn validate_properties(
    value: &Properties,
    allow_inherited_namespaces: bool,
    limits: &Limits,
) -> Result<()> {
    validate_id(&value.id, limits)?;
    if let Some(extension) = &value.extension_list {
        check_limit(
            limits,
            "extension XML bytes",
            extension.xml.len(),
            limits.extension_xml_bytes,
        )?;
        validate_extension_fragment_with_mode(&extension.xml, allow_inherited_namespaces, limits)?;
    }
    Ok(())
}

fn validate_id(value: &str, limits: &Limits) -> Result<()> {
    let units = value.encode_utf16().count();
    check_limit(
        limits,
        "Custom Data storage id UTF-16 units",
        units,
        limits.uid_units,
    )?;
    bounded_string(value, "storage id bytes", limits)
}

#[cfg(test)]
fn validate_extension_fragment(xml: &[u8]) -> Result<()> {
    validate_extension_fragment_with_mode(xml, false, &Limits::standard())
}

fn validate_extension_fragment_with_mode(
    xml: &[u8],
    allow_inherited_namespaces: bool,
    limits: &Limits,
) -> Result<()> {
    validate_extension_fragment_with_context_and_mode(xml, &[], allow_inherited_namespaces, limits)
}

fn validate_extension_fragment_with_context(
    xml: &[u8],
    inherited_namespaces: &[(String, String)],
    limits: &Limits,
) -> Result<()> {
    validate_extension_fragment_with_context_and_mode(xml, inherited_namespaces, false, limits)
}

fn validate_extension_fragment_with_context_and_mode(
    xml: &[u8],
    inherited_namespaces: &[(String, String)],
    allow_inherited_namespaces: bool,
    limits: &Limits,
) -> Result<()> {
    check_limit(
        limits,
        "extension XML bytes",
        xml.len(),
        limits.extension_xml_bytes,
    )?;
    if xml.starts_with(&[0xEF, 0xBB, 0xBF]) {
        return Err(invalid(
            "embedded extension XML may not contain a leading BOM",
        ));
    }
    let _ = parse_document_with_context(
        xml,
        ExpectedRoot::ExtensionList,
        allow_inherited_namespaces,
        inherited_namespaces,
        limits,
    )?;
    Ok(())
}

fn bounded_string(value: &str, name: &'static str, limits: &Limits) -> Result<()> {
    check_limit(limits, name, value.len(), limits.string_bytes)
}

fn copy_bounded(value: &[u8], max: usize, name: &'static str, limits: &Limits) -> Result<Vec<u8>> {
    check_limit(limits, name, value.len(), max)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|source| allocation(name, source))?;
    output.extend_from_slice(value);
    Ok(output)
}

fn copy_string(value: &str, name: &'static str, limits: &Limits) -> Result<String> {
    check_limit(limits, name, value.len(), limits.string_bytes)?;
    let mut output = String::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|source| allocation(name, source))?;
    output.push_str(value);
    Ok(output)
}

fn check_limit(_limits: &Limits, resource: &'static str, actual: usize, max: usize) -> Result<()> {
    if actual > max {
        Err(limit_for(resource, actual, max))
    } else {
        Ok(())
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

const fn limit_for(resource: &'static str, actual: usize, max: usize) -> Error {
    Error::Limit {
        resource,
        max,
        actual,
    }
}

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    Error::Allocation { resource, source }
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Decode(XmlError::Malformed(error.to_string()))
}

fn value_span(xml: &[u8], raw: &[u8]) -> Result<Range<usize>> {
    let start = (raw.as_ptr() as usize)
        .checked_sub(xml.as_ptr() as usize)
        .ok_or_else(|| invalid("XML attribute value is not source-backed"))?;
    let end = start
        .checked_add(raw.len())
        .ok_or_else(|| invalid("XML attribute source range overflow"))?;
    let quote = start.checked_sub(1).and_then(|at| xml.get(at));
    if xml.get(start..end) != Some(raw)
        || !matches!(quote, Some(b'\'' | b'"'))
        || xml.get(end) != quote
    {
        return Err(invalid("invalid XML attribute source range"));
    }
    Ok(start..end)
}

fn validate_xml_characters(xml: &[u8]) -> Result<()> {
    let xml = xml.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(xml);
    let text = std::str::from_utf8(xml).map_err(xml_error)?;
    if text.chars().any(|character| {
        !matches!(
            character,
            '\t' | '\n' | '\r' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}'
        )
    }) {
        return Err(invalid("invalid XML character"));
    }
    Ok(())
}

fn escaped_xstring_len(value: &str, limits: &Limits) -> Result<usize> {
    let bytes = value.as_bytes();
    let mut length = 0usize;
    for (at, character) in value.char_indices() {
        let added = if character == '_'
            && bytes
                .get(at..at.saturating_add(7))
                .is_some_and(|slice| spreadsheet_escape_at(slice, 0).is_some())
        {
            7
        } else if matches!(character, '\u{9}' | '\u{A}' | '\u{D}') || character >= '\u{20}' {
            match character {
                '&' => 5,
                '<' => 4,
                '"' | '\'' => 6,
                '\t' | '\n' | '\r' => 5,
                '\u{fffe}' | '\u{ffff}' => 7,
                _ => character.len_utf8(),
            }
        } else {
            let mut units = [0; 2];
            character.encode_utf16(&mut units).len() * 7
        };
        length = length.checked_add(added).ok_or_else(|| {
            limit_for(
                "escaped XML attribute length",
                usize::MAX,
                limits.string_bytes,
            )
        })?;
    }
    check_limit(limits, "storage id bytes", length, limits.string_bytes)?;
    Ok(length)
}

fn try_escaped_xstring(value: &str, limits: &Limits) -> Result<Vec<u8>> {
    let length = escaped_xstring_len(value, limits)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|source| allocation("escaped XML attribute", source))?;
    append_escaped_xstring(&mut output, value);
    debug_assert_eq!(output.len(), length);
    Ok(output)
}

fn append_escaped_xstring(output: &mut Vec<u8>, value: &str) {
    let bytes = value.as_bytes();
    for (at, character) in value.char_indices() {
        if character == '_'
            && bytes
                .get(at..at.saturating_add(7))
                .is_some_and(|slice| spreadsheet_escape_at(slice, 0).is_some())
        {
            output.extend_from_slice(b"_x005F_");
            continue;
        }
        if matches!(character, '\u{9}' | '\u{A}' | '\u{D}') || character >= '\u{20}' {
            match character {
                '&' => output.extend_from_slice(b"&amp;"),
                '<' => output.extend_from_slice(b"&lt;"),
                '"' => output.extend_from_slice(b"&quot;"),
                '\'' => output.extend_from_slice(b"&apos;"),
                '\t' => output.extend_from_slice(b"&#x9;"),
                '\n' => output.extend_from_slice(b"&#xA;"),
                '\r' => output.extend_from_slice(b"&#xD;"),
                '\u{fffe}' => output.extend_from_slice(b"_xFFFE_"),
                '\u{ffff}' => output.extend_from_slice(b"_xFFFF_"),
                _ => {
                    let mut encoded = [0; 4];
                    output.extend_from_slice(character.encode_utf8(&mut encoded).as_bytes());
                },
            }
            continue;
        }
        let mut units = [0; 2];
        for unit in character.encode_utf16(&mut units) {
            let mut escape = [0u8; 7];
            escape[0] = b'_';
            escape[1] = b'x';
            let digits = b"0123456789ABCDEF";
            escape[2] = digits[usize::from(*unit >> 12)];
            escape[3] = digits[usize::from((*unit >> 8) & 0xF)];
            escape[4] = digits[usize::from((*unit >> 4) & 0xF)];
            escape[5] = digits[usize::from(*unit & 0xF)];
            escape[6] = b'_';
            output.extend_from_slice(&escape);
        }
    }
}

fn spreadsheet_escape_at(bytes: &[u8], index: usize) -> Option<(u16, usize)> {
    let escape = bytes.get(index..index.checked_add(7)?)?;
    if escape[0] != b'_'
        || escape[1] != b'x'
        || escape[6] != b'_'
        || !escape[2..6].iter().all(|byte| byte.is_ascii_hexdigit())
    {
        return None;
    }
    let mut value = 0u16;
    for byte in &escape[2..6] {
        value = value.checked_mul(16)?;
        value = value.checked_add(u16::from(hex(*byte)?))?;
    }
    Some((value, index + 7))
}

fn hex(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn decode_spreadsheet_text(value: &str, limits: &Limits) -> Result<String> {
    check_limit(limits, "storage id bytes", value.len(), limits.string_bytes)?;
    let units = spreadsheet_text_utf16_units(value)?;
    check_limit(
        limits,
        "Custom Data storage id UTF-16 units",
        units,
        limits.uid_units,
    )?;
    let bytes = value.as_bytes();
    let mut decoded = String::new();
    decoded
        .try_reserve(value.len())
        .map_err(|source| allocation("storage id", source))?;
    let mut copied_until = 0;
    let mut index = 0;
    while index + 7 <= bytes.len() {
        let Some((unit, end)) = spreadsheet_escape_at(bytes, index) else {
            index += 1;
            continue;
        };
        decoded.push_str(&value[copied_until..index]);
        if (0xD800..=0xDBFF).contains(&unit) {
            let Some((low, pair_end)) = spreadsheet_escape_at(bytes, end) else {
                return Err(invalid(format!(
                    "unpaired high surrogate in SpreadsheetML escape at byte {index}"
                )));
            };
            if !(0xDC00..=0xDFFF).contains(&low) {
                return Err(invalid(format!(
                    "unpaired high surrogate in SpreadsheetML escape at byte {index}"
                )));
            }
            let scalar = 0x1_0000 + ((u32::from(unit) - 0xD800) << 10) + (u32::from(low) - 0xDC00);
            let character = char::from_u32(scalar)
                .ok_or_else(|| invalid("invalid surrogate pair in SpreadsheetML escape"))?;
            decoded.push(character);
            index = pair_end;
            copied_until = pair_end;
        } else if (0xDC00..=0xDFFF).contains(&unit) {
            return Err(invalid(format!(
                "unpaired low surrogate in SpreadsheetML escape at byte {index}"
            )));
        } else {
            let character = char::from_u32(u32::from(unit))
                .ok_or_else(|| invalid("invalid code unit in SpreadsheetML escape"))?;
            decoded.push(character);
            index = end;
            copied_until = end;
        }
    }
    decoded.push_str(&value[copied_until..]);
    check_limit(
        limits,
        "storage id bytes",
        decoded.len(),
        limits.string_bytes,
    )?;
    Ok(decoded)
}

fn spreadsheet_text_utf16_units(value: &str) -> Result<usize> {
    let bytes = value.as_bytes();
    let mut units = 0usize;
    let mut index = 0usize;
    while index < bytes.len() {
        if let Some((unit, end)) = bytes
            .get(index..)
            .and_then(|slice| spreadsheet_escape_at(slice, 0))
        {
            if (0xD800..=0xDBFF).contains(&unit) {
                let Some(next) = index.checked_add(end).and_then(|at| bytes.get(at..)) else {
                    return Err(invalid("SpreadsheetML escape offset overflow"));
                };
                let Some((low, low_len)) = spreadsheet_escape_at(next, 0) else {
                    return Err(invalid(format!(
                        "unpaired high surrogate in SpreadsheetML escape at byte {index}"
                    )));
                };
                if !(0xDC00..=0xDFFF).contains(&low) {
                    return Err(invalid(format!(
                        "unpaired high surrogate in SpreadsheetML escape at byte {index}"
                    )));
                }
                units = units.checked_add(2).ok_or_else(|| {
                    limit_for(
                        "Custom Data storage id UTF-16 units",
                        usize::MAX,
                        usize::MAX,
                    )
                })?;
                index = index
                    .checked_add(end)
                    .and_then(|at| at.checked_add(low_len))
                    .ok_or_else(|| invalid("SpreadsheetML escape offset overflow"))?;
                continue;
            }
            if (0xDC00..=0xDFFF).contains(&unit) {
                return Err(invalid(format!(
                    "unpaired low surrogate in SpreadsheetML escape at byte {index}"
                )));
            }
            units = units.checked_add(1).ok_or_else(|| {
                limit_for(
                    "Custom Data storage id UTF-16 units",
                    usize::MAX,
                    usize::MAX,
                )
            })?;
            index = index
                .checked_add(end)
                .ok_or_else(|| invalid("SpreadsheetML escape offset overflow"))?;
            continue;
        }
        let character = value[index..]
            .chars()
            .next()
            .ok_or_else(|| invalid("invalid UTF-8 character boundary"))?;
        let mut encoded = [0; 2];
        units = units
            .checked_add(character.encode_utf16(&mut encoded).len())
            .ok_or_else(|| {
                limit_for(
                    "Custom Data storage id UTF-16 units",
                    usize::MAX,
                    usize::MAX,
                )
            })?;
        index = index
            .checked_add(character.len_utf8())
            .ok_or_else(|| invalid("UTF-8 character offset overflow"))?;
    }
    Ok(units)
}

#[cfg(test)]
mod tests {
    use super::*;

    const XML: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

    #[test]
    fn preserves_source_extension_markup() {
        let source = format!(
            "<?xml version=\"1.0\"?>\n<!--before--><p:datastoreItem xmlns:p=\"{XML}\" xmlns:s=\"{SML}\" id=\"Storage&#45;1\"> <?inside?><p:extLst>\n<s:ext uri=\"urn:test\"><v:opaque xmlns:v=\"urn:v\"><![CDATA[<? and ]]><?pi?></v:opaque><!--tail--></s:ext>\n</p:extLst><!--after--></p:datastoreItem>"
        );
        let properties = parse_properties(source.as_bytes()).unwrap();
        assert_eq!(properties.id, "Storage-1");
        let extension = properties.extension_list.unwrap().xml;
        assert!(String::from_utf8_lossy(&extension).contains("<![CDATA[<? and ]]>"));
        assert!(String::from_utf8_lossy(&extension).contains("<?pi?>"));
        assert!(String::from_utf8_lossy(&extension).contains("xmlns:v=\"urn:v\""));
    }

    #[test]
    fn allows_empty_id_and_xstring_round_trip() {
        for id in [
            "",
            "A & 'B'",
            "_x0041_",
            "\u{0}\u{1}\u{fffe}\u{ffff}😀",
            "\t\n\r",
        ] {
            let value = Properties {
                id: id.into(),
                extension_list: None,
            };
            assert_eq!(
                parse_properties(&write_properties(&value).unwrap()).unwrap(),
                value
            );
        }
    }

    #[test]
    fn rejects_embedded_prolog_and_direct_non_whitespace() {
        assert!(validate_extension_fragment(
            b"<?xml version=\"1.0\"?><x14:extLst xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"/>"
        )
        .is_err());
        assert!(
            parse_properties(
                format!("<x14:datastoreItem xmlns:x14=\"{XML}\" id=\"x\">text</x14:datastoreItem>")
                    .as_bytes()
            )
            .is_err()
        );
    }

    #[test]
    fn accepts_document_bom_and_declaration_but_rejects_fragment_prologs() {
        let body = format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><x14:datastoreItem xmlns:x14=\"{XML}\" id=\"x\"/>"
        );
        let mut source = vec![0xEF, 0xBB, 0xBF];
        source.extend_from_slice(body.as_bytes());
        let parsed = parse_properties(&source).unwrap();
        assert_eq!(parsed.id, "x");
        assert!(validate_extension_fragment(
            b"\xEF\xBB\xBF<x14:extLst xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"/>"
        )
        .is_err());
    }

    #[test]
    fn arbitrary_prefix_and_entity_namespace_are_admitted() {
        let xml = format!(
            "<d:datastoreItem xmlns:d=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9&#x2F;main\" xmlns:s=\"{SML}\" id=\"x\"><d:extLst><s:ext uri=\"urn:test\"><q:opaque xmlns:q=\"urn:q\"/></s:ext></d:extLst></d:datastoreItem>"
        );
        assert!(parse_properties(xml.as_bytes()).is_ok());
    }

    #[test]
    fn extension_uri_entity_lexeme_is_retained() {
        let xml = format!(
            "<x14:datastoreItem xmlns:x14=\"{XML}\" xmlns:s=\"{SML}\" id=\"x\"><x14:extLst><s:ext uri=\"urn&#x3A;vendor\"><v:opaque xmlns:v=\"urn:v\"/></s:ext></x14:extLst></x14:datastoreItem>"
        );
        let properties = parse_properties(xml.as_bytes()).unwrap();
        assert_eq!(
            properties.extension_list.unwrap().xml,
            format!(
                "<x14:extLst xmlns:x14=\"{XML}\" xmlns:s=\"{SML}\"><s:ext uri=\"urn&#x3A;vendor\"><v:opaque xmlns:v=\"urn:v\"/></s:ext></x14:extLst>"
            )
            .into_bytes()
        );
    }

    #[test]
    fn rejects_cdata_outside_the_document_root() {
        let xml = format!("<![CDATA[before]]><x14:datastoreItem xmlns:x14=\"{XML}\" id=\"x\"/>");
        assert!(parse_properties(xml.as_bytes()).is_err());
        assert!(
            validate_workbook_root(
                format!("<![CDATA[before]]><workbook xmlns=\"{SML}\"/>").as_bytes()
            )
            .is_err()
        );
    }

    #[test]
    fn rejects_raw_cdata_delimiter_in_opaque_text() {
        let xml = format!(
            "<x14:datastoreItem xmlns:x14=\"{XML}\" xmlns:s=\"{SML}\" id=\"x\"><x14:extLst><s:ext uri=\"u\"><opaque>literal]]></opaque></s:ext></x14:extLst></x14:datastoreItem>"
        );
        assert!(parse_properties(xml.as_bytes()).is_err());
    }

    #[test]
    fn inherited_extension_prefixes_are_admitted_for_source_projection() {
        let source = format!(
            "<p:datastoreItem xmlns:p=\"{XML}\" xmlns:s=\"{SML}\" xmlns:q=\"urn:opaque\" id=\"x\"><p:extLst><s:ext uri=\"u\"><q:opaque><q:child/></q:opaque></s:ext></p:extLst></p:datastoreItem>"
        );
        let properties = parse_properties(source.as_bytes()).unwrap();
        let extension = properties.extension_list.as_ref().unwrap();
        assert_eq!(
            extension.xml,
            format!(
                "<p:extLst xmlns:p=\"{XML}\" xmlns:s=\"{SML}\" xmlns:q=\"urn:opaque\"><s:ext uri=\"u\"><q:opaque><q:child/></q:opaque></s:ext></p:extLst>"
            )
            .into_bytes()
        );
        // The source-bound host keeps the exact raw span separately while the
        // public projection remains standalone-ready.
        assert_eq!(properties.id, "x");
    }

    #[test]
    fn rejects_strict_sml_extension_in_x14_extension_list() {
        let source = format!(
            "<x14:datastoreItem xmlns:x14=\"{XML}\" xmlns:s=\"{STRICT_SML}\" id=\"x\"><x14:extLst><s:ext><opaque/></s:ext></x14:extLst></x14:datastoreItem>"
        );
        assert!(parse_properties(source.as_bytes()).is_err());
    }
}
