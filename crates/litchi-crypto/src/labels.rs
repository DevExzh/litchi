//! Typed MS-OFFCRYPTO sensitivity-label metadata.
//!
//! Unknown `extLst` entries are retained as inert XML fragments so a caller can
//! round-trip future extensions without interpreting or executing them.

use std::collections::HashSet;
use std::fmt;

use bitflags::bitflags;
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};

pub const NAMESPACE: &str = "http://schemas.microsoft.com/office/2020/mipLabelMetadata";
const MAX_LABEL_INFO_BYTES: usize = 16 * 1024 * 1024;
const MAX_LABELS: usize = 65_536;
const MAX_EXTENSIONS: usize = 65_536;
const MAX_XML_DEPTH: usize = 256;
const MAX_ATTRIBUTE_BYTES: usize = 1_048_576;

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
pub enum Error {
    Invalid(String),
    Xml(String),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Invalid(message) => write!(formatter, "invalid LabelInfo: {message}"),
            Self::Xml(message) => write!(formatter, "invalid LabelInfo XML: {message}"),
        }
    }
}

impl std::error::Error for Error {}

/// Parse and validate a complete `LabelInfo` XML stream.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when the stream violates the `LabelInfo`
/// grammar or a safety limit, and [`Error::Xml`] when the underlying XML is
/// malformed.
pub fn parse(xml: &[u8]) -> Result<List, Error> {
    if xml.len() > MAX_LABEL_INFO_BYTES {
        return Err(invalid("stream exceeds the 16 MiB safety limit"));
    }
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();

    let declaration = read_event(&mut reader, &mut buffer)?;
    if !matches!(declaration, Event::Decl(_)) || reader.buffer_position() == 0 {
        return Err(invalid("XML declaration must be the initial construct"));
    }
    buffer.clear();
    let root = match read_event(&mut reader, &mut buffer)? {
        Event::Start(root) => root.into_owned(),
        Event::Empty(root) => {
            validate_root(&root, reader.decoder())?;
            ensure_only_trailing_misc(&mut reader, &mut buffer)?;
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
    let prefixes = validate_root(&root, reader.decoder())?;
    let mut result = List::default();
    let mut seen_sites = HashSet::new();
    let mut extension_list_seen = false;

    loop {
        buffer.clear();
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_err| invalid("XML position overflows usize"))?;
        match read_event(&mut reader, &mut buffer)? {
            Event::Empty(element) if has_name(&element, &prefixes, b"label") => {
                if extension_list_seen {
                    return Err(invalid("label follows extLst"));
                }
                push_label(
                    &mut result,
                    &mut seen_sites,
                    parse_label(&element, reader.decoder(), &prefixes)?,
                )?;
            },
            Event::Start(element) if has_name(&element, &prefixes, b"label") => {
                if extension_list_seen {
                    return Err(invalid("label follows extLst"));
                }
                let label = parse_label(&element, reader.decoder(), &prefixes)?;
                expect_empty_element_body(&mut reader, &mut buffer, &prefixes, b"label")?;
                push_label(&mut result, &mut seen_sites, label)?;
            },
            Event::Start(element) if has_name(&element, &prefixes, b"extLst") => {
                if extension_list_seen {
                    return Err(invalid("labelList contains multiple extLst elements"));
                }
                validate_inherited_namespace(&element, &prefixes, reader.decoder())?;
                reject_non_namespace_attributes(&element)?;
                extension_list_seen = true;
                parse_extension_list(&mut reader, &mut buffer, xml, &prefixes, &mut result)?;
            },
            Event::Empty(element) if has_name(&element, &prefixes, b"extLst") => {
                if extension_list_seen {
                    return Err(invalid("labelList contains multiple extLst elements"));
                }
                validate_inherited_namespace(&element, &prefixes, reader.decoder())?;
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
    ensure_only_trailing_misc(&mut reader, &mut buffer)?;
    validate_list(&result, false)?;
    Ok(result)
}

/// Serialize `LabelInfo` deterministically, including the normative label
/// attribute order. Unknown extension fragments remain byte-for-byte intact.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when the list violates validation rules or the
/// serialized stream exceeds the 16 MiB safety limit.
pub fn write(value: &List) -> Result<Vec<u8>, Error> {
    validate_list(value, true)?;
    let mut output = Vec::new();
    output.extend_from_slice(b"<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?>");
    output.extend_from_slice(b"<clbl:labelList xmlns:clbl=\"");
    output.extend_from_slice(NAMESPACE.as_bytes());
    output.extend_from_slice(b"\" xmlns=\"");
    output.extend_from_slice(NAMESPACE.as_bytes());
    output.extend_from_slice(b"\"");
    let mut declared_prefixes = HashSet::new();
    declared_prefixes.insert(b"clbl".to_vec());
    for extension in &value.extensions {
        let prefix = extension_prefix(&extension.xml)?;
        if !prefix.is_empty() && declared_prefixes.insert(prefix.to_vec()) {
            output.extend_from_slice(b" xmlns:");
            output.extend_from_slice(prefix);
            output.extend_from_slice(b"=\"");
            output.extend_from_slice(NAMESPACE.as_bytes());
            output.extend_from_slice(b"\"");
        }
    }
    output.extend_from_slice(b">");
    for label in &value.labels {
        output.extend_from_slice(b"<clbl:label id=\"");
        push_escaped_attribute(&mut output, &label.id);
        output.extend_from_slice(b"\" enabled=\"");
        output.push(if label.enabled { b'1' } else { b'0' });
        output.extend_from_slice(b"\" method=\"");
        push_escaped_attribute(&mut output, label.method.as_str());
        output.extend_from_slice(b"\" siteId=\"");
        output.extend_from_slice(label.site_id.to_string().as_bytes());
        if let Some(bits) = label.content_bits {
            output.extend_from_slice(b"\" contentBits=\"");
            output.extend_from_slice(bits.bits().to_string().as_bytes());
        }
        output.extend_from_slice(b"\" removed=\"");
        output.push(if label.removed { b'1' } else { b'0' });
        output.extend_from_slice(b"\"/>");
    }
    if !value.extensions.is_empty() {
        output.extend_from_slice(b"<clbl:extLst>");
        for extension in &value.extensions {
            output.extend_from_slice(&extension.xml);
        }
        output.extend_from_slice(b"</clbl:extLst>");
    }
    output.extend_from_slice(b"</clbl:labelList>");
    if output.len() > MAX_LABEL_INFO_BYTES {
        return Err(invalid("serialized stream exceeds the 16 MiB safety limit"));
    }
    Ok(output)
}

fn parse_extension_list(
    reader: &mut Reader<&[u8]>,
    buffer: &mut Vec<u8>,
    xml: &[u8],
    prefixes: &HashSet<Vec<u8>>,
    result: &mut List,
) -> Result<(), Error> {
    loop {
        buffer.clear();
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_err| invalid("XML position overflows usize"))?;
        match read_event(reader, buffer)? {
            Event::Start(element) if has_name(&element, prefixes, b"ext") => {
                validate_inherited_namespace(&element, prefixes, reader.decoder())?;
                let uri = required_attribute(&element, b"uri", reader.decoder())?;
                if uri.is_empty() {
                    return Err(invalid("extension URI is empty"));
                }
                if result.extensions.len() >= MAX_EXTENSIONS {
                    return Err(invalid("too many extension entries"));
                }
                consume_extension(reader, buffer, prefixes)?;
                let end = usize::try_from(reader.buffer_position())
                    .map_err(|_err| invalid("XML position overflows usize"))?;
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
) -> Result<(), Error> {
    let mut depth = 1usize;
    let mut direct_children = 0usize;
    loop {
        buffer.clear();
        match read_event(reader, buffer)? {
            Event::Start(_) => {
                if depth == 1 {
                    direct_children += 1;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("XML depth overflow"))?;
                if depth > MAX_XML_DEPTH {
                    return Err(invalid("extension XML is too deeply nested"));
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
) -> Result<Label, Error> {
    validate_inherited_namespace(element, prefixes, decoder)?;
    let mut id = None;
    let mut enabled = None;
    let mut method = None;
    let mut site_id = None;
    let mut content_bits = None;
    let mut removed = None;
    for raw_attribute in element.attributes().with_checks(true) {
        let attribute = raw_attribute.map_err(|error| xml_error(error.to_string()))?;
        if is_namespace_attribute(attribute.key.as_ref()) {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| xml_error(error.to_string()))?
            .into_owned();
        if value.len() > MAX_ATTRIBUTE_BYTES {
            return Err(invalid("label attribute exceeds the 1 MiB safety limit"));
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

fn validate_list(value: &List, writing: bool) -> Result<(), Error> {
    if value.labels.len() > MAX_LABELS {
        return Err(invalid("too many labels"));
    }
    if value.extensions.len() > MAX_EXTENSIONS {
        return Err(invalid("too many extensions"));
    }
    let mut sites = HashSet::with_capacity(value.labels.len());
    for label in &value.labels {
        validate_label(label, writing)?;
        if !sites.insert(label.site_id) {
            return Err(invalid("multiple labels use the same siteId"));
        }
    }
    for extension in &value.extensions {
        if extension.uri.is_empty() || extension.uri.len() > MAX_ATTRIBUTE_BYTES {
            return Err(invalid("extension URI is empty or too long"));
        }
        let parsed = parse_extension_fragment(&extension.xml)?;
        if parsed != extension.uri {
            return Err(invalid("extension URI disagrees with its XML fragment"));
        }
    }
    Ok(())
}

fn validate_label(label: &Label, writing: bool) -> Result<(), Error> {
    if label.id.is_empty() || label.id.len() > MAX_ATTRIBUTE_BYTES {
        return Err(invalid("label id is empty or too long"));
    }
    if label.id.bytes().any(|byte| byte.is_ascii_uppercase()) {
        return Err(invalid("label id is not lowercase"));
    }
    if label.method.as_str().len() > MAX_ATTRIBUTE_BYTES {
        return Err(invalid("label method is too long"));
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

fn parse_extension_fragment(xml: &[u8]) -> Result<String, Error> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    match read_event(&mut reader, &mut buffer)? {
        Event::Start(element) if local_name(element.name().as_ref()) == b"ext" => {
            let uri = required_attribute(&element, b"uri", reader.decoder())?;
            let prefix = element
                .name()
                .as_ref()
                .iter()
                .position(|byte| *byte == b':')
                .map_or_else(Vec::new, |index| element.name().as_ref()[..index].to_vec());
            let prefixes = HashSet::from([prefix]);
            consume_extension(&mut reader, &mut buffer, &prefixes)?;
            ensure_only_trailing_misc(&mut reader, &mut buffer)?;
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
    for raw_attribute in root.attributes().with_checks(true) {
        let attribute = raw_attribute.map_err(|error| xml_error(error.to_string()))?;
        if !is_namespace_attribute(attribute.key.as_ref()) {
            return Err(invalid("labelList has an unexpected attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| xml_error(error.to_string()))?;
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
    if optional_attribute(element, &key, decoder)?.is_some_and(|namespace| namespace != NAMESPACE) {
        return Err(invalid("element rebinds the label namespace prefix"));
    }
    Ok(())
}

fn push_label(result: &mut List, sites: &mut HashSet<Guid>, label: Label) -> Result<(), Error> {
    if result.labels.len() >= MAX_LABELS {
        return Err(invalid("too many labels"));
    }
    validate_label(&label, false)?;
    if !sites.insert(label.site_id) {
        return Err(invalid("multiple labels use the same siteId"));
    }
    result.labels.push(label);
    Ok(())
}

fn expect_empty_element_body(
    reader: &mut Reader<&[u8]>,
    buffer: &mut Vec<u8>,
    prefixes: &HashSet<Vec<u8>>,
    local: &[u8],
) -> Result<(), Error> {
    loop {
        buffer.clear();
        match read_event(reader, buffer)? {
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
) -> Result<(), Error> {
    loop {
        buffer.clear();
        match read_event(reader, buffer)? {
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
    for raw_attribute in element.attributes().with_checks(true) {
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
    for raw_attribute in element.attributes().with_checks(true) {
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

fn read_event<'a>(reader: &mut Reader<&[u8]>, buffer: &'a mut Vec<u8>) -> Result<Event<'a>, Error> {
    reader
        .read_event_into(buffer)
        .map_err(|error| xml_error(error.to_string()))
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

fn push_escaped_attribute(output: &mut Vec<u8>, value: &str) {
    for byte in value.bytes() {
        match byte {
            b'&' => output.extend_from_slice(b"&amp;"),
            b'<' => output.extend_from_slice(b"&lt;"),
            b'"' => output.extend_from_slice(b"&quot;"),
            b'\t' => output.extend_from_slice(b"&#x9;"),
            b'\n' => output.extend_from_slice(b"&#xA;"),
            b'\r' => output.extend_from_slice(b"&#xD;"),
            _ => output.push(byte),
        }
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
}
