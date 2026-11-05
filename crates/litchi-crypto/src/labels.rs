//! Typed MS-OFFCRYPTO sensitivity-label metadata.
//!
//! Unknown `extLst` entries are retained as inert XML fragments so a caller can
//! round-trip future extensions without interpreting or executing them.

use std::collections::HashSet;
use std::fmt;

use bitflags::bitflags;
use litchi_core::xml::ReaderOrigin;
use litchi_ole_common::xml_attributes::BytesStartExt as _;
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};

pub const NAMESPACE: &str = "http://schemas.microsoft.com/office/2020/mipLabelMetadata";
const MAX_LABEL_INFO_BYTES: usize = 16 * 1024 * 1024;
const MAX_LABELS: usize = 65_536;
const MAX_EXTENSIONS: usize = 65_536;
const MAX_XML_DEPTH: usize = 256;
const MAX_ATTRIBUTE_BYTES: usize = 1_048_576;
const MAX_XML_NODES: usize = 65_536;
const MAX_OUTPUT_BYTES: usize = 16 * 1024 * 1024;

/// Resource policy for one sensitivity-label parse or serialization.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum bytes in the LabelInfo XML stream.
    pub max_xml_bytes: usize,
    /// Maximum typed labels.
    pub max_labels: usize,
    /// Maximum opaque extension entries.
    pub max_extensions: usize,
    /// Maximum nested extension XML depth.
    pub max_xml_depth: usize,
    /// Maximum XML parser events.
    pub max_xml_nodes: usize,
    /// Maximum decoded attribute or extension URI bytes.
    pub max_attribute_bytes: usize,
    /// Maximum serialized LabelInfo bytes.
    pub max_output_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_xml_bytes: MAX_LABEL_INFO_BYTES,
            max_labels: MAX_LABELS,
            max_extensions: MAX_EXTENSIONS,
            max_xml_depth: MAX_XML_DEPTH,
            max_xml_nodes: MAX_XML_NODES,
            max_attribute_bytes: MAX_ATTRIBUTE_BYTES,
            max_output_bytes: MAX_OUTPUT_BYTES,
        }
    }
}

impl Limits {
    fn validate(self) -> Result<Self, Error> {
        if self.max_xml_bytes == 0 || self.max_xml_bytes > MAX_LABEL_INFO_BYTES {
            return Err(invalid_limit(
                LimitKind::XmlBytes,
                self.max_xml_bytes,
                MAX_LABEL_INFO_BYTES,
            ));
        }
        if self.max_labels == 0 || self.max_labels > MAX_LABELS {
            return Err(invalid_limit(
                LimitKind::Labels,
                self.max_labels,
                MAX_LABELS,
            ));
        }
        if self.max_extensions == 0 || self.max_extensions > MAX_EXTENSIONS {
            return Err(invalid_limit(
                LimitKind::Extensions,
                self.max_extensions,
                MAX_EXTENSIONS,
            ));
        }
        if self.max_xml_depth == 0 || self.max_xml_depth > MAX_XML_DEPTH {
            return Err(invalid_limit(
                LimitKind::XmlDepth,
                self.max_xml_depth,
                MAX_XML_DEPTH,
            ));
        }
        if self.max_xml_nodes == 0 || self.max_xml_nodes > MAX_XML_NODES {
            return Err(invalid_limit(
                LimitKind::XmlNodes,
                self.max_xml_nodes,
                MAX_XML_NODES,
            ));
        }
        if self.max_attribute_bytes == 0 || self.max_attribute_bytes > MAX_ATTRIBUTE_BYTES {
            return Err(invalid_limit(
                LimitKind::AttributeBytes,
                self.max_attribute_bytes,
                MAX_ATTRIBUTE_BYTES,
            ));
        }
        if self.max_output_bytes == 0 || self.max_output_bytes > MAX_OUTPUT_BYTES {
            return Err(invalid_limit(
                LimitKind::OutputBytes,
                self.max_output_bytes,
                MAX_OUTPUT_BYTES,
            ));
        }
        Ok(self)
    }

    /// Derive the label profile retained by a DataSpaces inspection.
    pub(crate) fn from_spaces(limits: &crate::spaces::Limits) -> Self {
        Self {
            max_xml_bytes: limits.max_stream_bytes.min(MAX_LABEL_INFO_BYTES),
            max_labels: limits.max_entries.min(MAX_LABELS),
            max_extensions: limits.max_entries.min(MAX_EXTENSIONS),
            max_xml_depth: limits.max_xml_depth.min(MAX_XML_DEPTH),
            max_xml_nodes: limits.max_xml_nodes.min(MAX_XML_NODES),
            max_attribute_bytes: limits.max_string_bytes.min(MAX_ATTRIBUTE_BYTES),
            max_output_bytes: limits.max_output_bytes.min(MAX_OUTPUT_BYTES),
        }
    }
}

/// Resource dimensions governed by a sensitivity-label profile.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum LimitKind {
    /// LabelInfo XML bytes.
    XmlBytes,
    /// Typed labels.
    Labels,
    /// Opaque extension entries.
    Extensions,
    /// Nested extension XML depth.
    XmlDepth,
    /// XML parser events.
    XmlNodes,
    /// Decoded attribute or extension URI bytes.
    AttributeBytes,
    /// Serialized LabelInfo bytes.
    OutputBytes,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let name = match self {
            Self::XmlBytes => "XML bytes",
            Self::Labels => "labels",
            Self::Extensions => "extensions",
            Self::XmlDepth => "XML depth",
            Self::XmlNodes => "XML nodes",
            Self::AttributeBytes => "attribute bytes",
            Self::OutputBytes => "output bytes",
        };
        formatter.write_str(name)
    }
}

/// Canonical lowercase classification GUID used by the `siteId` attribute.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Guid([u8; 16]);

impl Guid {
    #[must_use]
    pub const fn from_bytes(bytes: [u8; 16]) -> Self {
        Self(bytes)
    }

    #[must_use]
    pub const fn as_bytes(&self) -> &[u8; 16] {
        &self.0
    }
}

impl fmt::Display for Guid {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        for (index, byte) in self.0.iter().enumerate() {
            if matches!(index, 4 | 6 | 8 | 10) {
                formatter.write_str("-")?;
            }
            write!(formatter, "{byte:02x}")?;
        }
        Ok(())
    }
}

impl std::str::FromStr for Guid {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        if value.len() != 36
            || value.bytes().enumerate().any(|(index, byte)| match index {
                8 | 13 | 18 | 23 => byte != b'-',
                _ => !byte.is_ascii_digit() && !(b'a'..=b'f').contains(&byte),
            })
        {
            return Err(invalid("siteId is not a canonical lowercase GUID"));
        }
        let mut output = [0; 16];
        for (nibble, byte) in value.bytes().filter(|byte| *byte != b'-').enumerate() {
            let digit = match byte {
                b'0'..=b'9' => byte - b'0',
                b'a'..=b'f' => byte - b'a' + 10,
                _ => return Err(invalid("siteId contains a non-hexadecimal digit")),
            };
            if nibble % 2 == 0 {
                output[nibble / 2] = digit << 4;
            } else {
                output[nibble / 2] |= digit;
            }
        }
        Ok(Self(output))
    }
}

/// Assignment method retained according to the open string-valued schema.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Method {
    Standard,
    Privileged,
    /// Empty method required for a removed-label tombstone.
    Removed,
    /// A future case-sensitive method value.
    Other(String),
}

impl Method {
    #[must_use]
    pub fn as_str(&self) -> &str {
        match self {
            Self::Standard => "Standard",
            Self::Privileged => "Privileged",
            Self::Removed => "",
            Self::Other(value) => value,
        }
    }
}

bitflags! {
    /// Content-marking DWORD. Unknown future bits are preserved on read.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct Content: u32 {
        const HEADER = 1;
        const FOOTER = 2;
        const WATERMARK = 4;
        const ENCRYPTION = 8;
    }
}

impl Content {
    #[must_use]
    pub const fn unknown_bits(self) -> u32 {
        self.bits() & !Self::all().bits()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Label {
    pub id: String,
    pub enabled: bool,
    pub method: Method,
    pub site_id: Guid,
    pub content_bits: Option<Content>,
    pub removed: bool,
}

/// Opaque, validated complete `<ext>` fragment with its non-empty URI.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Extension {
    pub uri: String,
    pub xml: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct List {
    pub labels: Vec<Label>,
    pub extensions: Vec<Extension>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    InvalidLimit {
        kind: LimitKind,
        value: usize,
        maximum: usize,
    },
    LimitExceeded {
        kind: LimitKind,
        observed: usize,
        maximum: usize,
    },
    Allocation {
        resource: &'static str,
    },
    Invalid(String),
    Xml(String),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidLimit {
                kind,
                value,
                maximum,
            } => write!(
                formatter,
                "invalid LabelInfo {kind} limit {value}; maximum is {maximum}"
            ),
            Self::LimitExceeded {
                kind,
                observed,
                maximum,
            } => write!(
                formatter,
                "LabelInfo {kind} limit exceeded: observed {observed}, maximum {maximum}"
            ),
            Self::Allocation { resource } => {
                write!(formatter, "LabelInfo allocation failed for {resource}")
            },
            Self::Invalid(message) => write!(formatter, "invalid LabelInfo: {message}"),
            Self::Xml(message) => write!(formatter, "invalid LabelInfo XML: {message}"),
        }
    }
}

impl std::error::Error for Error {}

struct Budget {
    limits: Limits,
    nodes: usize,
}

impl Budget {
    fn new(limits: Limits) -> Self {
        Self { limits, nodes: 0 }
    }

    fn node(&mut self) -> Result<(), Error> {
        self.nodes = self.nodes.checked_add(1).ok_or_else(|| {
            limit_exceeded(LimitKind::XmlNodes, usize::MAX, self.limits.max_xml_nodes)
        })?;
        if self.nodes > self.limits.max_xml_nodes {
            return Err(limit_exceeded(
                LimitKind::XmlNodes,
                self.nodes,
                self.limits.max_xml_nodes,
            ));
        }
        Ok(())
    }
}

/// Parse and validate a complete `LabelInfo` XML stream.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when the stream violates the `LabelInfo`
/// grammar, [`Error::LimitExceeded`] when the selected profile is crossed,
/// and [`Error::Xml`] when the underlying XML is malformed.
pub fn parse(xml: &[u8], limits: &Limits) -> Result<List, Error> {
    let limits = limits.validate()?;
    if xml.len() > limits.max_xml_bytes {
        return Err(limit_exceeded(
            LimitKind::XmlBytes,
            xml.len(),
            limits.max_xml_bytes,
        ));
    }
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    // Offsets address `xml`, whose leading byte-order mark precedes reader
    // position zero.
    let origin = ReaderOrigin::of(xml);
    let mut buffer = Vec::new();
    let mut budget = Budget::new(limits);

    let declaration = read_event(&mut reader, &mut buffer, &mut budget)?;
    if !matches!(declaration, Event::Decl(_)) || reader.buffer_position() == 0 {
        return Err(invalid("XML declaration must be the initial construct"));
    }
    buffer.clear();
    let root = match read_event(&mut reader, &mut buffer, &mut budget)? {
        Event::Start(root) => root.into_owned(),
        Event::Empty(root) => {
            validate_root(&root, reader.decoder(), &limits)?;
            ensure_only_trailing_misc(&mut reader, &mut buffer, &mut budget)?;
            return Ok(List::default());
        },
        Event::End(_)
        | Event::Text(_)
        | Event::CData(_)
        | Event::Comment(_)
        | Event::Decl(_)
        | Event::PI(_)
        | Event::DocType(_)
        | Event::Eof
        | Event::GeneralRef(_) => return Err(invalid("labelList must be the one root element")),
    };
    let prefixes = validate_root(&root, reader.decoder(), &limits)?;
    let mut result = List::default();
    let mut seen_sites = HashSet::new();
    let mut extension_list_seen = false;

    loop {
        buffer.clear();
        let event_start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("XML position overflows usize"))?;
        match read_event(&mut reader, &mut buffer, &mut budget)? {
            Event::Empty(element) if has_name(&element, &prefixes, b"label") => {
                if extension_list_seen {
                    return Err(invalid("label follows extLst"));
                }
                push_label(
                    &mut result,
                    &mut seen_sites,
                    parse_label(&element, reader.decoder(), &prefixes, &limits)?,
                    &limits,
                )?;
            },
            Event::Start(element) if has_name(&element, &prefixes, b"label") => {
                if extension_list_seen {
                    return Err(invalid("label follows extLst"));
                }
                let label = parse_label(&element, reader.decoder(), &prefixes, &limits)?;
                expect_empty_element_body(
                    &mut reader,
                    &mut buffer,
                    &prefixes,
                    b"label",
                    &mut budget,
                )?;
                push_label(&mut result, &mut seen_sites, label, &limits)?;
            },
            Event::Start(element) if has_name(&element, &prefixes, b"extLst") => {
                if extension_list_seen {
                    return Err(invalid("labelList contains multiple extLst elements"));
                }
                validate_inherited_namespace(&element, &prefixes, reader.decoder(), &limits)?;
                reject_non_namespace_attributes(&element)?;
                extension_list_seen = true;
                parse_extension_list(
                    &mut reader,
                    &mut buffer,
                    xml,
                    &prefixes,
                    &mut result,
                    &mut budget,
                )?;
            },
            Event::Empty(element) if has_name(&element, &prefixes, b"extLst") => {
                if extension_list_seen {
                    return Err(invalid("labelList contains multiple extLst elements"));
                }
                validate_inherited_namespace(&element, &prefixes, reader.decoder(), &limits)?;
                reject_non_namespace_attributes(&element)?;
                extension_list_seen = true;
            },
            Event::End(element) if has_end_name(&element, &prefixes, b"labelList") => break,
            Event::Text(text) if whitespace(text.as_ref()) => {},
            Event::Comment(_) | Event::PI(_) => {},
            Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(invalid("DTD and general entity references are forbidden"));
            },
            Event::Eof => return Err(invalid("labelList is not closed")),
            event @ (Event::Start(_)
            | Event::Empty(_)
            | Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Decl(_)) => {
                return Err(invalid(format!(
                    "unexpected construct at byte {event_start}: {}",
                    event_kind(&event)
                )));
            },
        }
    }
    ensure_only_trailing_misc(&mut reader, &mut buffer, &mut budget)?;
    validate_list(&result, false, &limits)?;
    Ok(result)
}

/// Serialize `LabelInfo` deterministically, including the normative label
/// attribute order. Unknown extension fragments remain byte-for-byte intact.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when the list violates validation rules or
/// [`Error::LimitExceeded`] when the serialized stream exceeds the caller's
/// output limit.
pub fn write(value: &List, limits: &Limits) -> Result<Vec<u8>, Error> {
    let limits = limits.validate()?;
    validate_list(value, true, &limits)?;
    let mut output = Vec::new();
    append_output(
        &mut output,
        b"<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?>",
        &limits,
    )?;
    append_output(&mut output, b"<clbl:labelList xmlns:clbl=\"", &limits)?;
    append_output(&mut output, NAMESPACE.as_bytes(), &limits)?;
    append_output(&mut output, b"\" xmlns=\"", &limits)?;
    append_output(&mut output, NAMESPACE.as_bytes(), &limits)?;
    append_output(&mut output, b"\"", &limits)?;
    let mut declared_prefixes = HashSet::new();
    declared_prefixes.insert(b"clbl".to_vec());
    for extension in &value.extensions {
        let prefix = extension_prefix(&extension.xml)?;
        if !prefix.is_empty() && declared_prefixes.insert(prefix.to_vec()) {
            append_output(&mut output, b" xmlns:", &limits)?;
            append_output(&mut output, prefix, &limits)?;
            append_output(&mut output, b"=\"", &limits)?;
            append_output(&mut output, NAMESPACE.as_bytes(), &limits)?;
            append_output(&mut output, b"\"", &limits)?;
        }
    }
    append_output(&mut output, b">", &limits)?;
    for label in &value.labels {
        append_output(&mut output, b"<clbl:label id=\"", &limits)?;
        push_escaped_attribute(&mut output, &label.id, &limits)?;
        append_output(&mut output, b"\" enabled=\"", &limits)?;
        append_output(
            &mut output,
            &[if label.enabled { b'1' } else { b'0' }],
            &limits,
        )?;
        append_output(&mut output, b"\" method=\"", &limits)?;
        push_escaped_attribute(&mut output, label.method.as_str(), &limits)?;
        append_output(&mut output, b"\" siteId=\"", &limits)?;
        append_output(&mut output, label.site_id.to_string().as_bytes(), &limits)?;
        if let Some(bits) = label.content_bits {
            append_output(&mut output, b"\" contentBits=\"", &limits)?;
            append_output(&mut output, bits.bits().to_string().as_bytes(), &limits)?;
        }
        append_output(&mut output, b"\" removed=\"", &limits)?;
        append_output(
            &mut output,
            &[if label.removed { b'1' } else { b'0' }],
            &limits,
        )?;
        append_output(&mut output, b"\"/>", &limits)?;
    }
    if !value.extensions.is_empty() {
        append_output(&mut output, b"<clbl:extLst>", &limits)?;
        for extension in &value.extensions {
            append_output(&mut output, &extension.xml, &limits)?;
        }
        append_output(&mut output, b"</clbl:extLst>", &limits)?;
    }
    append_output(&mut output, b"</clbl:labelList>", &limits)?;
    Ok(output)
}

fn parse_extension_list(
    reader: &mut Reader<&[u8]>,
    buffer: &mut Vec<u8>,
    xml: &[u8],
    prefixes: &HashSet<Vec<u8>>,
    result: &mut List,
    budget: &mut Budget,
) -> Result<(), Error> {
    // `xml` is the reader's input; extension bytes are sliced from it.
    let origin = ReaderOrigin::of(xml);
    loop {
        buffer.clear();
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("XML position overflows usize"))?;
        match read_event(reader, buffer, budget)? {
            Event::Start(element) if has_name(&element, prefixes, b"ext") => {
                validate_inherited_namespace(&element, prefixes, reader.decoder(), &budget.limits)?;
                let uri = required_attribute(&element, b"uri", reader.decoder())?;
                if uri.is_empty() || uri.len() > budget.limits.max_attribute_bytes {
                    if uri.is_empty() {
                        return Err(invalid("extension URI is empty"));
                    }
                    return Err(limit_exceeded(
                        LimitKind::AttributeBytes,
                        uri.len(),
                        budget.limits.max_attribute_bytes,
                    ));
                }
                if result.extensions.len() >= budget.limits.max_extensions {
                    return Err(limit_exceeded(
                        LimitKind::Extensions,
                        result.extensions.len().saturating_add(1),
                        budget.limits.max_extensions,
                    ));
                }
                consume_extension(reader, buffer, prefixes, budget)?;
                let end = origin
                    .offset(reader.buffer_position())
                    .ok_or_else(|| invalid("XML position overflows usize"))?;
                result
                    .extensions
                    .try_reserve(1)
                    .map_err(|_error| Error::Allocation {
                        resource: "extensions",
                    })?;
                result.extensions.push(Extension {
                    uri,
                    xml: xml
                        .get(start..end)
                        .ok_or_else(|| invalid("extension XML range is invalid"))?
                        .to_vec(),
                });
            },
            Event::Empty(element) if has_name(&element, prefixes, b"ext") => {
                return Err(invalid(format!(
                    "extension '{}' has no required child element",
                    required_attribute(&element, b"uri", reader.decoder())?
                )));
            },
            Event::End(element) if has_end_name(&element, prefixes, b"extLst") => return Ok(()),
            Event::Text(text) if whitespace(text.as_ref()) => {},
            Event::Comment(_) | Event::PI(_) => {},
            Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(invalid("DTD and general entity references are forbidden"));
            },
            Event::Eof => return Err(invalid("extLst is not closed")),
            Event::Start(_)
            | Event::Empty(_)
            | Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Decl(_) => return Err(invalid("extLst contains an unexpected construct")),
        }
    }
}

fn consume_extension(
    reader: &mut Reader<&[u8]>,
    buffer: &mut Vec<u8>,
    prefixes: &HashSet<Vec<u8>>,
    budget: &mut Budget,
) -> Result<(), Error> {
    let mut depth = 1usize;
    let mut direct_children = 0usize;
    loop {
        buffer.clear();
        match read_event(reader, buffer, budget)? {
            Event::Start(_) => {
                if depth == 1 {
                    direct_children += 1;
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    limit_exceeded(LimitKind::XmlDepth, usize::MAX, budget.limits.max_xml_depth)
                })?;
                if depth > budget.limits.max_xml_depth {
                    return Err(limit_exceeded(
                        LimitKind::XmlDepth,
                        depth,
                        budget.limits.max_xml_depth,
                    ));
                }
            },
            Event::Empty(_) => {
                if depth == 1 {
                    direct_children += 1;
                }
            },
            Event::End(element) => {
                depth -= 1;
                if depth == 0 {
                    if !has_end_name(&element, prefixes, b"ext") {
                        return Err(invalid("extension has a mismatched closing element"));
                    }
                    if direct_children != 1 {
                        return Err(invalid("extension must contain exactly one child element"));
                    }
                    return Ok(());
                }
            },
            Event::Text(text) if depth > 1 || whitespace(text.as_ref()) => {},
            Event::CData(_) if depth > 1 => {},
            Event::Comment(_) | Event::PI(_) => {},
            Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(invalid("DTD and general entity references are forbidden"));
            },
            Event::Eof => return Err(invalid("extension is not closed")),
            Event::Text(_) | Event::CData(_) | Event::Decl(_) => {
                return Err(invalid("extension has text outside its child element"));
            },
        }
    }
}

fn parse_label(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    prefixes: &HashSet<Vec<u8>>,
    limits: &Limits,
) -> Result<Label, Error> {
    validate_inherited_namespace(element, prefixes, decoder, limits)?;
    let mut id = None;
    let mut enabled = None;
    let mut method = None;
    let mut site_id = None;
    let mut content_bits = None;
    let mut removed = None;
    for raw_attribute in element.checked_attributes() {
        let attribute = raw_attribute.map_err(|error| xml_error(error.to_string()))?;
        if is_namespace_attribute(attribute.key.as_ref()) {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| xml_error(error.to_string()))?
            .into_owned();
        if value.len() > limits.max_attribute_bytes {
            return Err(limit_exceeded(
                LimitKind::AttributeBytes,
                value.len(),
                limits.max_attribute_bytes,
            ));
        }
        let slot_was_set = match attribute.key.as_ref() {
            b"id" => id.replace(value).is_some(),
            b"enabled" => enabled.replace(parse_bool(&value)?).is_some(),
            b"method" => method.replace(parse_method(value)).is_some(),
            b"siteId" => site_id.replace(value.parse()?).is_some(),
            b"contentBits" => content_bits
                .replace(Content::from_bits_retain(value.parse().map_err(
                    |_err| invalid("contentBits is not an unsigned decimal DWORD"),
                )?))
                .is_some(),
            b"removed" => removed.replace(parse_bool(&value)?).is_some(),
            name => {
                return Err(invalid(format!(
                    "label has unexpected attribute '{}'",
                    String::from_utf8_lossy(name)
                )));
            },
        };
        if slot_was_set {
            return Err(invalid("label has a duplicate attribute"));
        }
    }
    Ok(Label {
        id: id.ok_or_else(|| invalid("label lacks id"))?,
        enabled: enabled.ok_or_else(|| invalid("label lacks enabled"))?,
        method: method.ok_or_else(|| invalid("label lacks method"))?,
        site_id: site_id.ok_or_else(|| invalid("label lacks siteId"))?,
        content_bits,
        removed: removed.ok_or_else(|| invalid("label lacks removed"))?,
    })
}

fn validate_list(value: &List, writing: bool, limits: &Limits) -> Result<(), Error> {
    if value.labels.len() > limits.max_labels {
        return Err(limit_exceeded(
            LimitKind::Labels,
            value.labels.len(),
            limits.max_labels,
        ));
    }
    if value.extensions.len() > limits.max_extensions {
        return Err(limit_exceeded(
            LimitKind::Extensions,
            value.extensions.len(),
            limits.max_extensions,
        ));
    }
    let mut sites = HashSet::with_capacity(value.labels.len());
    for label in &value.labels {
        validate_label(label, writing, limits)?;
        if !sites.insert(label.site_id) {
            return Err(invalid("multiple labels use the same siteId"));
        }
    }
    for extension in &value.extensions {
        if extension.xml.len() > limits.max_xml_bytes {
            return Err(limit_exceeded(
                LimitKind::XmlBytes,
                extension.xml.len(),
                limits.max_xml_bytes,
            ));
        }
        if extension.uri.is_empty() || extension.uri.len() > limits.max_attribute_bytes {
            if extension.uri.is_empty() {
                return Err(invalid("extension URI is empty or too long"));
            }
            return Err(limit_exceeded(
                LimitKind::AttributeBytes,
                extension.uri.len(),
                limits.max_attribute_bytes,
            ));
        }
        let parsed = parse_extension_fragment(&extension.xml, limits)?;
        if parsed != extension.uri {
            return Err(invalid("extension URI disagrees with its XML fragment"));
        }
    }
    Ok(())
}

fn validate_label(label: &Label, writing: bool, limits: &Limits) -> Result<(), Error> {
    if label.id.is_empty() {
        return Err(invalid("label id is empty or too long"));
    }
    if label.id.len() > limits.max_attribute_bytes {
        return Err(limit_exceeded(
            LimitKind::AttributeBytes,
            label.id.len(),
            limits.max_attribute_bytes,
        ));
    }
    if label.id.bytes().any(|byte| byte.is_ascii_uppercase()) {
        return Err(invalid("label id is not lowercase"));
    }
    if label.method.as_str().len() > limits.max_attribute_bytes {
        return Err(limit_exceeded(
            LimitKind::AttributeBytes,
            label.method.as_str().len(),
            limits.max_attribute_bytes,
        ));
    }
    if label.removed {
        if !matches!(label.method, Method::Removed) {
            return Err(invalid("removed label method must be empty"));
        }
        if writing && label.content_bits.is_some() {
            return Err(invalid("writer omits contentBits for removed labels"));
        }
    } else if matches!(label.method, Method::Removed) {
        return Err(invalid("active label method cannot be empty"));
    }
    if writing
        && label
            .content_bits
            .is_some_and(|bits| bits.unknown_bits() != 0)
    {
        return Err(invalid("writer cannot set reserved contentBits"));
    }
    Ok(())
}

fn parse_extension_fragment(xml: &[u8], limits: &Limits) -> Result<String, Error> {
    if xml.len() > limits.max_xml_bytes {
        return Err(limit_exceeded(
            LimitKind::XmlBytes,
            xml.len(),
            limits.max_xml_bytes,
        ));
    }
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut budget = Budget::new(*limits);
    match read_event(&mut reader, &mut buffer, &mut budget)? {
        Event::Start(element) if local_name(element.name().as_ref()) == b"ext" => {
            let uri = required_attribute(&element, b"uri", reader.decoder())?;
            let prefix = element
                .name()
                .as_ref()
                .iter()
                .position(|byte| *byte == b':')
                .map_or_else(Vec::new, |index| element.name().as_ref()[..index].to_vec());
            let prefixes = HashSet::from([prefix]);
            consume_extension(&mut reader, &mut buffer, &prefixes, &mut budget)?;
            ensure_only_trailing_misc(&mut reader, &mut buffer, &mut budget)?;
            Ok(uri)
        },
        Event::Start(_)
        | Event::End(_)
        | Event::Empty(_)
        | Event::Text(_)
        | Event::CData(_)
        | Event::Comment(_)
        | Event::Decl(_)
        | Event::PI(_)
        | Event::DocType(_)
        | Event::Eof
        | Event::GeneralRef(_) => Err(invalid("extension XML is not one complete ext element")),
    }
}

fn extension_prefix(xml: &[u8]) -> Result<&[u8], Error> {
    let name_start = xml
        .strip_prefix(b"<")
        .ok_or_else(|| invalid("extension XML lacks a start tag"))?;
    let name_end = name_start
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || matches!(byte, b'/' | b'>'))
        .ok_or_else(|| invalid("extension XML has an unterminated start tag"))?;
    let name = &name_start[..name_end];
    if local_name(name) != b"ext" {
        return Err(invalid("extension XML is not an ext element"));
    }
    Ok(name
        .iter()
        .position(|byte| *byte == b':')
        .map_or(&[][..], |index| &name[..index]))
}

fn validate_root(
    root: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    limits: &Limits,
) -> Result<HashSet<Vec<u8>>, Error> {
    let name = root.name();
    let bytes = name.as_ref();
    if local_name(bytes) != b"labelList" {
        return Err(invalid("root element is not labelList"));
    }
    let root_prefix = bytes
        .iter()
        .position(|byte| *byte == b':')
        .map_or_else(Vec::new, |index| bytes[..index].to_vec());
    let namespace_key = if root_prefix.is_empty() {
        b"xmlns".to_vec()
    } else {
        let mut key = b"xmlns:".to_vec();
        key.extend_from_slice(&root_prefix);
        key
    };
    let namespace = optional_attribute(root, &namespace_key, decoder)?
        .ok_or_else(|| invalid("labelList namespace is not declared"))?;
    if namespace != NAMESPACE {
        return Err(invalid("labelList uses the wrong namespace"));
    }
    let mut prefixes = HashSet::new();
    for raw_attribute in root.checked_attributes() {
        let attribute = raw_attribute.map_err(|error| xml_error(error.to_string()))?;
        if !is_namespace_attribute(attribute.key.as_ref()) {
            return Err(invalid("labelList has an unexpected attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| xml_error(error.to_string()))?;
        if value.len() > limits.max_attribute_bytes {
            return Err(limit_exceeded(
                LimitKind::AttributeBytes,
                value.len(),
                limits.max_attribute_bytes,
            ));
        }
        if value == NAMESPACE {
            prefixes.insert(
                attribute
                    .key
                    .as_ref()
                    .strip_prefix(b"xmlns:")
                    .unwrap_or_default()
                    .to_vec(),
            );
        }
    }
    if !prefixes.contains(&root_prefix) {
        return Err(invalid("root prefix is not bound to the label namespace"));
    }
    Ok(prefixes)
}

fn validate_inherited_namespace(
    element: &BytesStart<'_>,
    prefixes: &HashSet<Vec<u8>>,
    decoder: quick_xml::encoding::Decoder,
    limits: &Limits,
) -> Result<(), Error> {
    let name = element.name();
    let prefix = name
        .as_ref()
        .iter()
        .position(|byte| *byte == b':')
        .map_or(&[][..], |index| &name.as_ref()[..index]);
    if !prefixes.contains(prefix) {
        return Err(invalid(
            "element prefix is not bound to the label namespace",
        ));
    }
    let key = if prefix.is_empty() {
        b"xmlns".to_vec()
    } else {
        let mut key = b"xmlns:".to_vec();
        key.extend_from_slice(prefix);
        key
    };
    if let Some(namespace) = optional_attribute(element, &key, decoder)? {
        if namespace.len() > limits.max_attribute_bytes {
            return Err(limit_exceeded(
                LimitKind::AttributeBytes,
                namespace.len(),
                limits.max_attribute_bytes,
            ));
        }
        if namespace != NAMESPACE {
            return Err(invalid("element rebinds the label namespace prefix"));
        }
    }
    Ok(())
}

fn push_label(
    result: &mut List,
    sites: &mut HashSet<Guid>,
    label: Label,
    limits: &Limits,
) -> Result<(), Error> {
    if result.labels.len() >= limits.max_labels {
        return Err(limit_exceeded(
            LimitKind::Labels,
            result.labels.len().saturating_add(1),
            limits.max_labels,
        ));
    }
    validate_label(&label, false, limits)?;
    if !sites.insert(label.site_id) {
        return Err(invalid("multiple labels use the same siteId"));
    }
    result
        .labels
        .try_reserve(1)
        .map_err(|_error| Error::Allocation { resource: "labels" })?;
    result.labels.push(label);
    Ok(())
}

fn expect_empty_element_body(
    reader: &mut Reader<&[u8]>,
    buffer: &mut Vec<u8>,
    prefixes: &HashSet<Vec<u8>>,
    local: &[u8],
    budget: &mut Budget,
) -> Result<(), Error> {
    loop {
        buffer.clear();
        match read_event(reader, buffer, budget)? {
            Event::End(end) if has_end_name(&end, prefixes, local) => return Ok(()),
            Event::Text(text) if whitespace(text.as_ref()) => {},
            Event::Comment(_) | Event::PI(_) => {},
            Event::Start(_)
            | Event::Empty(_)
            | Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Decl(_)
            | Event::DocType(_)
            | Event::Eof
            | Event::GeneralRef(_) => {
                return Err(invalid("label element must not contain content"));
            },
        }
    }
}

fn ensure_only_trailing_misc(
    reader: &mut Reader<&[u8]>,
    buffer: &mut Vec<u8>,
    budget: &mut Budget,
) -> Result<(), Error> {
    loop {
        buffer.clear();
        match read_event(reader, buffer, budget)? {
            Event::Eof => return Ok(()),
            Event::Text(text) if whitespace(text.as_ref()) => {},
            Event::Comment(_) | Event::PI(_) => {},
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Decl(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => return Err(invalid("XML has content after labelList")),
        }
    }
}

fn required_attribute(
    element: &BytesStart<'_>,
    key: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<String, Error> {
    optional_attribute(element, key, decoder)?
        .ok_or_else(|| invalid(format!("element lacks {}", String::from_utf8_lossy(key))))
}

fn optional_attribute(
    element: &BytesStart<'_>,
    key: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<Option<String>, Error> {
    let mut value = None;
    for raw_attribute in element.checked_attributes() {
        let attribute = raw_attribute.map_err(|error| xml_error(error.to_string()))?;
        if attribute.key.as_ref() == key {
            let normalized = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| xml_error(error.to_string()))?
                .into_owned();
            if value.replace(normalized).is_some() {
                return Err(invalid("duplicate XML attribute"));
            }
        }
    }
    Ok(value)
}

fn reject_non_namespace_attributes(element: &BytesStart<'_>) -> Result<(), Error> {
    for raw_attribute in element.checked_attributes() {
        let attribute = raw_attribute.map_err(|error| xml_error(error.to_string()))?;
        if !is_namespace_attribute(attribute.key.as_ref()) {
            return Err(invalid("element has an unexpected attribute"));
        }
    }
    Ok(())
}

fn parse_method(value: String) -> Method {
    match value.as_str() {
        "Standard" => Method::Standard,
        "Privileged" => Method::Privileged,
        "" => Method::Removed,
        _ => Method::Other(value),
    }
}

fn parse_bool(value: &str) -> Result<bool, Error> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid("boolean attribute is not an XML Schema boolean")),
    }
}

fn has_name(element: &BytesStart<'_>, prefixes: &HashSet<Vec<u8>>, local: &[u8]) -> bool {
    qualified_name_matches(element.name().as_ref(), prefixes, local)
}

fn has_end_name(
    element: &quick_xml::events::BytesEnd<'_>,
    prefixes: &HashSet<Vec<u8>>,
    local: &[u8],
) -> bool {
    qualified_name_matches(element.name().as_ref(), prefixes, local)
}

fn qualified_name_matches(name: &[u8], prefixes: &HashSet<Vec<u8>>, local: &[u8]) -> bool {
    if local_name(name) != local {
        return false;
    }
    let prefix = name
        .iter()
        .position(|byte| *byte == b':')
        .map_or(&[][..], |index| &name[..index]);
    prefixes.contains(prefix)
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn is_namespace_attribute(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn whitespace(bytes: &[u8]) -> bool {
    bytes.iter().all(u8::is_ascii_whitespace)
}

fn read_event<'a>(
    reader: &mut Reader<&[u8]>,
    buffer: &'a mut Vec<u8>,
    budget: &mut Budget,
) -> Result<Event<'a>, Error> {
    let event = reader
        .read_event_into(buffer)
        .map_err(|error| xml_error(error.to_string()))?;
    if !matches!(&event, Event::Eof) {
        budget.node()?;
    }
    Ok(event)
}

fn event_kind(event: &Event<'_>) -> &'static str {
    match event {
        Event::Start(_) => "start element",
        Event::End(_) => "end element",
        Event::Empty(_) => "empty element",
        Event::Text(_) => "text",
        Event::CData(_) => "CDATA",
        Event::Comment(_) => "comment",
        Event::Decl(_) => "XML declaration",
        Event::PI(_) => "processing instruction",
        Event::DocType(_) => "DTD",
        Event::Eof => "end of file",
        Event::GeneralRef(_) => "entity reference",
    }
}

fn push_escaped_attribute(output: &mut Vec<u8>, value: &str, limits: &Limits) -> Result<(), Error> {
    for byte in value.bytes() {
        match byte {
            b'&' => append_output(output, b"&amp;", limits)?,
            b'<' => append_output(output, b"&lt;", limits)?,
            b'"' => append_output(output, b"&quot;", limits)?,
            b'\t' => append_output(output, b"&#x9;", limits)?,
            b'\n' => append_output(output, b"&#xA;", limits)?,
            b'\r' => append_output(output, b"&#xD;", limits)?,
            _ => append_output(output, &[byte], limits)?,
        }
    }
    Ok(())
}

fn append_output(output: &mut Vec<u8>, bytes: &[u8], limits: &Limits) -> Result<(), Error> {
    let end = output.len().checked_add(bytes.len()).ok_or_else(|| {
        limit_exceeded(LimitKind::OutputBytes, usize::MAX, limits.max_output_bytes)
    })?;
    if end > limits.max_output_bytes {
        return Err(limit_exceeded(
            LimitKind::OutputBytes,
            end,
            limits.max_output_bytes,
        ));
    }
    output
        .try_reserve(bytes.len())
        .map_err(|_error| Error::Allocation { resource: "output" })?;
    output.extend_from_slice(bytes);
    Ok(())
}

fn invalid_limit(kind: LimitKind, value: usize, maximum: usize) -> Error {
    Error::InvalidLimit {
        kind,
        value,
        maximum,
    }
}

fn limit_exceeded(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::LimitExceeded {
        kind,
        observed,
        maximum,
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn xml_error(message: impl Into<String>) -> Error {
    Error::Xml(message.into())
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        reason = "test code panics on failure; unwrap keeps assertions concise"
    )]
    use super::*;

    const SITE: &str = "12345678-1234-5678-90ab-1234567890ab";

    fn parse(xml: &[u8]) -> Result<List, Error> {
        super::parse(xml, &Limits::default())
    }

    fn write(value: &List) -> Result<Vec<u8>, Error> {
        super::write(value, &Limits::default())
    }

    fn active_label() -> Label {
        Label {
            id: "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee".to_string(),
            enabled: true,
            method: Method::Privileged,
            site_id: SITE.parse().unwrap(),
            content_bits: Some(Content::HEADER | Content::ENCRYPTION),
            removed: false,
        }
    }

    #[test]
    fn typed_labels_round_trip_in_normative_attribute_order() {
        let value = List {
            labels: vec![active_label()],
            extensions: Vec::new(),
        };
        let xml = write(&value).unwrap();
        let text = std::str::from_utf8(&xml).unwrap();
        assert!(text.contains(
            "id=\"aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee\" enabled=\"1\" method=\"Privileged\" siteId=\"12345678-1234-5678-90ab-1234567890ab\" contentBits=\"9\" removed=\"0\""
        ));
        assert_eq!(parse(&xml).unwrap(), value);
    }

    #[test]
    fn unknown_extension_is_preserved_exactly() {
        let xml = format!(
            "<?xml version=\"1.0\" encoding=\"utf-8\"?><x:labelList xmlns:x=\"{NAMESPACE}\"><x:extLst><x:ext uri=\"urn:test\"><future:data xmlns:future=\"urn:future\" value=\"1\"/></x:ext></x:extLst></x:labelList>"
        );
        let parsed = parse(xml.as_bytes()).unwrap();
        assert_eq!(parsed.extensions[0].uri, "urn:test");
        assert_eq!(
            parsed.extensions[0].xml,
            b"<x:ext uri=\"urn:test\"><future:data xmlns:future=\"urn:future\" value=\"1\"/></x:ext>"
        );
        let round_trip = write(&parsed).unwrap();
        assert!(
            std::str::from_utf8(&round_trip)
                .unwrap()
                .contains(&format!("xmlns:x=\"{NAMESPACE}\""))
        );
        assert_eq!(parse(&round_trip).unwrap(), parsed);
    }

    /// Change 0765: a stream that begins with a UTF-8 byte-order mark keeps
    /// its extension bytes exact and reports errors at byte offsets of the
    /// stream, three bytes after the reader's uncounted positions.
    #[test]
    fn a_byte_order_marked_stream_parses_and_reports_like_an_unmarked_one() {
        let plain = format!(
            "<?xml version=\"1.0\" encoding=\"utf-8\"?><x:labelList xmlns:x=\"{NAMESPACE}\"><x:extLst><x:ext uri=\"urn:test\"><future:data xmlns:future=\"urn:future\" value=\"1\"/></x:ext></x:extLst></x:labelList>"
        );
        let marked = format!("\u{feff}{plain}");
        assert_eq!(
            parse(marked.as_bytes()).unwrap(),
            parse(plain.as_bytes()).unwrap()
        );

        let plain = format!(
            "<?xml version=\"1.0\" encoding=\"utf-8\"?><x:labelList xmlns:x=\"{NAMESPACE}\"><x:unknown/></x:labelList>"
        );
        let marked = format!("\u{feff}{plain}");
        let offset = |xml: &str| {
            let message = parse(xml.as_bytes()).unwrap_err().to_string();
            let start = message.find("at byte ").unwrap() + "at byte ".len();
            message[start..]
                .chars()
                .take_while(char::is_ascii_digit)
                .collect::<String>()
                .parse::<usize>()
                .unwrap()
        };
        let plain_offset = offset(&plain);
        assert_eq!(
            &plain.as_bytes()[plain_offset..plain_offset + 10],
            b"<x:unknown"
        );
        assert_eq!(offset(&marked), plain_offset + 3);
    }

    #[test]
    fn removed_label_is_a_typed_tombstone() {
        let value = List {
            labels: vec![Label {
                id: SITE.to_string(),
                enabled: false,
                method: Method::Removed,
                site_id: SITE.parse().unwrap(),
                content_bits: None,
                removed: true,
            }],
            extensions: Vec::new(),
        };
        assert_eq!(parse(&write(&value).unwrap()).unwrap(), value);
    }

    #[test]
    fn malformed_and_conflicting_metadata_is_rejected() {
        assert!(parse(b"<labelList/>").is_err());
        let duplicate = format!(
            "<?xml version=\"1.0\"?><labelList xmlns=\"{NAMESPACE}\"><label id=\"a\" enabled=\"1\" method=\"Standard\" siteId=\"{SITE}\" removed=\"0\"/><label id=\"b\" enabled=\"1\" method=\"Standard\" siteId=\"{SITE}\" removed=\"0\"/></labelList>"
        );
        assert!(parse(duplicate.as_bytes()).is_err());
        let rebound = format!(
            "<?xml version=\"1.0\"?><x:labelList xmlns:x=\"{NAMESPACE}\"><x:label xmlns:x=\"urn:wrong\" id=\"a\" enabled=\"1\" method=\"Standard\" siteId=\"{SITE}\" removed=\"0\"/></x:labelList>"
        );
        assert!(parse(rebound.as_bytes()).is_err());

        let mut removed = active_label();
        removed.removed = true;
        assert!(
            write(&List {
                labels: vec![removed],
                extensions: Vec::new(),
            })
            .is_err()
        );
    }

    #[test]
    fn dtd_and_reserved_writer_bits_are_rejected() {
        let xml = format!("<?xml version=\"1.0\"?><!DOCTYPE x><labelList xmlns=\"{NAMESPACE}\"/>");
        assert!(parse(xml.as_bytes()).is_err());
        let mut label = active_label();
        label.content_bits = Some(Content::from_bits_retain(0x10));
        assert!(
            write(&List {
                labels: vec![label],
                extensions: Vec::new(),
            })
            .is_err()
        );
    }

    #[test]
    fn caller_limits_cover_exact_size_count_depth_and_output() {
        let invalid_profile = Limits {
            max_output_bytes: 0,
            ..Limits::default()
        };
        assert!(matches!(
            super::write(&List::default(), &invalid_profile),
            Err(Error::InvalidLimit {
                kind: LimitKind::OutputBytes,
                ..
            })
        ));

        let value = List {
            labels: vec![active_label()],
            extensions: Vec::new(),
        };
        let xml = write(&value).unwrap();
        let exact = Limits {
            max_xml_bytes: xml.len(),
            max_output_bytes: xml.len(),
            ..Limits::default()
        };
        assert_eq!(super::parse(&xml, &exact).unwrap(), value);
        assert_eq!(super::write(&value, &exact).unwrap(), xml);

        let too_small_output = Limits {
            max_output_bytes: xml.len() - 1,
            ..Limits::default()
        };
        assert!(matches!(
            super::write(&value, &too_small_output),
            Err(Error::LimitExceeded {
                kind: LimitKind::OutputBytes,
                ..
            })
        ));

        let too_few_labels = Limits {
            max_labels: 1,
            ..Limits::default()
        };
        let two = List {
            labels: vec![
                active_label(),
                Label {
                    id: "bbbbbbbb-bbbb-cccc-dddd-eeeeeeeeeeee".to_string(),
                    site_id: "87654321-4321-8765-ba09-ba9876543210".parse().unwrap(),
                    ..active_label()
                },
            ],
            extensions: Vec::new(),
        };
        let two_xml = super::write(&two, &Limits::default()).unwrap();
        assert!(matches!(
            super::parse(&two_xml, &too_few_labels),
            Err(Error::LimitExceeded {
                kind: LimitKind::Labels,
                ..
            })
        ));

        let nested = format!(
            "<?xml version=\"1.0\"?><x:labelList xmlns:x=\"{NAMESPACE}\"><x:extLst><x:ext uri=\"urn:test\"><x:a><x:b/></x:a></x:ext></x:extLst></x:labelList>"
        );
        let shallow = Limits {
            max_xml_depth: 1,
            ..Limits::default()
        };
        assert!(matches!(
            super::parse(nested.as_bytes(), &shallow),
            Err(Error::LimitExceeded {
                kind: LimitKind::XmlDepth,
                ..
            })
        ));
    }
}
