//! Bounded XML validation for neutral Ribbon customUI documents.

use crate::xml_name;
use crate::{Error, Result};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{OpcPackage, Part};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesDecl, BytesPI, BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;
use std::borrow::Cow;
use std::collections::HashSet;

use super::model::{Family, Limits, UI2_NAMESPACE, V2007_NAMESPACE, V2010_NAMESPACE, Version};

pub(super) fn validate_xml(xml: &[u8], family: Family, limits: &Limits) -> Result<Version> {
    if xml.len() > limits.xml_bytes {
        return Err(Error::Limit {
            resource: "Ribbon XML bytes",
            max: limits.xml_bytes,
            actual: xml.len(),
        });
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_comments = true;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    let mut first_event = true;
    let mut version = None;

    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader.read_resolved_event()?;
        let was_first = first_event;
        first_event = false;
        match event {
            Event::Decl(declaration) => {
                count_node(&mut nodes, limits)?;
                if !was_first || declaration_seen || root_seen {
                    return Err(Error::Invalid(
                        "Ribbon XML declaration must be the first event and occur once".into(),
                    ));
                }
                validate_declaration(&declaration)?;
                declaration_seen = true;
            },
            Event::DocType(_) => {
                return Err(Error::Invalid("DTD is forbidden in Ribbon XML".into()));
            },
            Event::PI(instruction) => {
                count_node(&mut nodes, limits)?;
                validate_instruction(&reader, &instruction)?;
            },
            Event::Start(element) => {
                count_node(&mut nodes, limits)?;
                validate_qname(element.name().as_ref(), "element")?;
                let root_version = if depth == 0 {
                    if root_seen || root_closed {
                        return Err(Error::Invalid(
                            "Ribbon XML must contain exactly one root".into(),
                        ));
                    }
                    Some(validate_root(&namespace, &element, decoder, family)?)
                } else {
                    validate_element_namespace(&namespace)?;
                    None
                };
                validate_attributes(&reader, &element, decoder, &mut nodes, limits)?;
                if let Some(root_version) = root_version {
                    version = Some(root_version);
                    root_seen = true;
                }
                depth = depth.checked_add(1).ok_or(Error::Limit {
                    resource: "Ribbon XML depth",
                    max: limits.depth,
                    actual: usize::MAX,
                })?;
                if depth > limits.depth {
                    return Err(Error::Limit {
                        resource: "Ribbon XML depth",
                        max: limits.depth,
                        actual: depth,
                    });
                }
            },
            Event::Empty(element) => {
                count_node(&mut nodes, limits)?;
                validate_qname(element.name().as_ref(), "element")?;
                let child_depth = depth.checked_add(1).ok_or(Error::Limit {
                    resource: "Ribbon XML depth",
                    max: limits.depth,
                    actual: usize::MAX,
                })?;
                if child_depth > limits.depth {
                    return Err(Error::Limit {
                        resource: "Ribbon XML depth",
                        max: limits.depth,
                        actual: child_depth,
                    });
                }
                let root_version = if depth == 0 {
                    if root_seen || root_closed {
                        return Err(Error::Invalid(
                            "Ribbon XML must contain exactly one root".into(),
                        ));
                    }
                    Some(validate_root(&namespace, &element, decoder, family)?)
                } else {
                    validate_element_namespace(&namespace)?;
                    None
                };
                validate_attributes(&reader, &element, decoder, &mut nodes, limits)?;
                if let Some(root_version) = root_version {
                    version = Some(root_version);
                    root_seen = true;
                    root_closed = true;
                }
            },
            Event::End(_) => {
                count_node(&mut nodes, limits)?;
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Invalid("Ribbon XML has an unexpected end element".into())
                })?;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                count_node(&mut nodes, limits)?;
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(format!("invalid Ribbon XML text: {error}")))?;
                validate_xml_chars(&value)?;
                if depth == 0 && !value.trim().is_empty() {
                    return Err(Error::Invalid(
                        "Ribbon XML contains text outside its root".into(),
                    ));
                }
            },
            Event::CData(text) => {
                count_node(&mut nodes, limits)?;
                if depth == 0 {
                    return Err(Error::Invalid(
                        "Ribbon XML contains CDATA outside its root".into(),
                    ));
                }
                let value = text
                    .decode()
                    .map_err(|error| Error::Xml(format!("invalid Ribbon CDATA: {error}")))?;
                validate_xml_chars(&value)?;
            },
            Event::GeneralRef(reference) => {
                count_node(&mut nodes, limits)?;
                if depth == 0 {
                    return Err(Error::Invalid(
                        "Ribbon XML contains an entity reference outside its root".into(),
                    ));
                }
                validate_reference(&reference)?;
            },
            Event::Comment(comment) => {
                count_node(&mut nodes, limits)?;
                let value = comment
                    .decode()
                    .map_err(|error| Error::Xml(format!("invalid Ribbon comment: {error}")))?;
                validate_xml_chars(&value)?;
            },
            Event::Eof => break,
        }
    }

    if !root_seen || !root_closed || depth != 0 {
        return Err(Error::Invalid(
            "Ribbon XML must contain one complete customUI root".into(),
        ));
    }
    version.ok_or_else(|| Error::Invalid("Ribbon XML has no customUI root".into()))
}

fn validate_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    nodes: &mut usize,
    limits: &Limits,
) -> Result<()> {
    let mut expanded = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        count_node(nodes, limits)?;
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(format!("invalid Ribbon XML attribute: {error}")))?;
        validate_xml_chars(&value)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            validate_namespace_declaration(attribute.key.as_ref(), &value)?;
            continue;
        }
        validate_qname(attribute.key.as_ref(), "attribute")?;
        let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = normalized_namespace(&namespace, decoder, "attribute")?;
        let QName(raw_name) = attribute.key;
        let local_name = raw_name
            .rsplit(|byte| *byte == b':')
            .next()
            .unwrap_or(raw_name);
        if !expanded.insert((namespace, local_name)) {
            return Err(Error::Invalid(format!(
                "duplicate expanded Ribbon attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            )));
        }
    }
    Ok(())
}

fn validate_element_namespace(namespace: &ResolveResult<'_>) -> Result<()> {
    if let ResolveResult::Unknown(prefix) = namespace {
        return Err(Error::Invalid(format!(
            "unbound Ribbon element namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )));
    }
    Ok(())
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn validate_namespace_declaration(name: &[u8], value: &str) -> Result<()> {
    let prefix = name.strip_prefix(b"xmlns:");
    if let Some(prefix) = prefix {
        let prefix = std::str::from_utf8(prefix)
            .map_err(|error| Error::Xml(format!("invalid namespace prefix: {error}")))?;
        if !xml_name::is_ncname(prefix) || prefix == "xmlns" {
            return Err(Error::Invalid(format!(
                "invalid Ribbon namespace prefix '{prefix}'"
            )));
        }
        if value.is_empty() {
            return Err(Error::Invalid(format!(
                "Ribbon namespace prefix '{prefix}' cannot be undeclared in XML 1.0"
            )));
        }
        if (prefix == "xml") != (value == "http://www.w3.org/XML/1998/namespace") {
            return Err(Error::Invalid(
                "the XML namespace URI may be bound only to the 'xml' prefix".into(),
            ));
        }
    } else if value == "http://www.w3.org/XML/1998/namespace" {
        return Err(Error::Invalid(
            "the XML namespace URI may be bound only to the 'xml' prefix".into(),
        ));
    }
    if value == "http://www.w3.org/2000/xmlns/" {
        return Err(Error::Invalid(
            "the xmlns namespace URI cannot be rebound".into(),
        ));
    }
    if value
        .bytes()
        .any(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
    {
        return Err(Error::Invalid(
            "Ribbon namespace URI cannot contain XML whitespace".into(),
        ));
    }
    Ok(())
}

fn normalized_namespace<'a>(
    namespace: &ResolveResult<'a>,
    decoder: quick_xml::encoding::Decoder,
    kind: &str,
) -> Result<Cow<'a, str>> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => normalize_namespace_value(value, decoder),
        ResolveResult::Unbound => Ok(Cow::Borrowed("")),
        ResolveResult::Unknown(prefix) => Err(Error::Invalid(format!(
            "unbound Ribbon {kind} namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn normalize_namespace_value<'a>(
    value: &'a [u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<Cow<'a, str>> {
    let decoded = decoder
        .decode(value)
        .map_err(|error| Error::Xml(format!("invalid Ribbon namespace URI: {error}")))?;
    let normalized = match decoded {
        Cow::Borrowed(value) => quick_xml::escape::unescape(value)
            .map_err(|error| Error::Xml(format!("invalid Ribbon namespace URI: {error}")))?,
        Cow::Owned(value) => Cow::Owned(
            quick_xml::escape::unescape(&value)
                .map_err(|error| Error::Xml(format!("invalid Ribbon namespace URI: {error}")))?
                .into_owned(),
        ),
    };
    validate_xml_chars(&normalized)?;
    Ok(normalized)
}

fn validate_reference(reference: &quick_xml::events::BytesRef<'_>) -> Result<()> {
    if let Some(character) = reference
        .resolve_char_ref()
        .map_err(|error| Error::Xml(format!("invalid Ribbon character reference: {error}")))?
    {
        return if is_xml_char(character) {
            Ok(())
        } else {
            Err(Error::Xml(format!(
                "Ribbon character reference U+{:04X} is forbidden by XML 1.0",
                u32::from(character)
            )))
        };
    }
    let name = reference
        .decode()
        .map_err(|error| Error::Xml(format!("invalid Ribbon entity reference: {error}")))?;
    if matches!(name.as_ref(), "amp" | "lt" | "gt" | "apos" | "quot") {
        Ok(())
    } else {
        Err(Error::Invalid(format!(
            "unsupported Ribbon entity reference '&{name};'"
        )))
    }
}

fn validate_declaration(declaration: &BytesDecl<'_>) -> Result<()> {
    let version = declaration.xml_version()?;
    if version != XmlVersion::Explicit1_0 {
        return Err(Error::Invalid(
            "Ribbon XML declaration must use version 1.0".into(),
        ));
    }
    let declaration_text = std::str::from_utf8(declaration.as_ref())
        .map_err(|error| Error::Xml(format!("invalid Ribbon XML declaration: {error}")))?;
    let raw = BytesStart::from_content(declaration_text, 3);
    let mut state = 0u8;
    for attribute in raw.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.prefix().is_some() {
            return Err(Error::Invalid(format!(
                "unexpected Ribbon XML declaration attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            )));
        }
        state = match (state, attribute.key.as_ref()) {
            (0, b"version") => 1,
            (1, b"encoding") => 2,
            (1 | 2, b"standalone") => 3,
            _ => {
                return Err(Error::Invalid(format!(
                    "unexpected or out-of-order Ribbon XML declaration attribute '{}'",
                    String::from_utf8_lossy(attribute.key.as_ref())
                )));
            },
        };
        std::str::from_utf8(attribute.value.as_ref())
            .map_err(|error| Error::Xml(format!("invalid Ribbon XML declaration: {error}")))?;
    }
    if let Some(encoding) = declaration.encoding() {
        let encoding = encoding.map_err(|error| Error::Xml(error.to_string()))?;
        let encoding = std::str::from_utf8(&encoding)
            .map_err(|error| Error::Xml(format!("invalid Ribbon XML encoding: {error}")))?;
        if !valid_encoding_name(encoding) {
            return Err(Error::Invalid(format!(
                "Ribbon XML encoding '{encoding}' is not an EncName"
            )));
        }
        if !encoding.eq_ignore_ascii_case("UTF-8") {
            return Err(Error::Invalid(format!(
                "Ribbon XML encoding '{encoding}' is unsupported; Ribbon XML must be UTF-8"
            )));
        }
    }
    if let Some(standalone) = declaration.standalone() {
        let standalone = standalone.map_err(|error| Error::Xml(error.to_string()))?;
        if !matches!(standalone.as_ref(), b"yes" | b"no") {
            return Err(Error::Invalid(
                "Ribbon XML standalone must be 'yes' or 'no'".into(),
            ));
        }
    }
    Ok(())
}

fn validate_instruction(reader: &NsReader<&[u8]>, instruction: &BytesPI<'_>) -> Result<()> {
    let target = reader
        .decoder()
        .decode(instruction.target())
        .map_err(|error| Error::Xml(format!("invalid Ribbon instruction target: {error}")))?;
    if !xml_name::is_xml_name(&target) || target.eq_ignore_ascii_case("xml") {
        return Err(Error::Invalid(format!(
            "invalid Ribbon processing-instruction target '{target}'"
        )));
    }
    let content = reader
        .decoder()
        .decode(instruction.content())
        .map_err(|error| Error::Xml(format!("invalid Ribbon instruction content: {error}")))?;
    validate_xml_chars(&content)
}

fn validate_qname(value: &[u8], kind: &str) -> Result<()> {
    let value = std::str::from_utf8(value)
        .map_err(|error| Error::Xml(format!("invalid Ribbon {kind} name: {error}")))?;
    if xml_name::is_qualified_name(value) {
        Ok(())
    } else {
        Err(Error::Invalid(format!(
            "invalid Ribbon {kind} QName '{value}'"
        )))
    }
}

fn validate_xml_chars(value: &str) -> Result<()> {
    if value.chars().all(is_xml_char) {
        Ok(())
    } else {
        Err(Error::Xml(
            "Ribbon XML contains a character forbidden by XML 1.0".into(),
        ))
    }
}

fn is_xml_char(character: char) -> bool {
    matches!(
        character,
        '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
    )
}

fn valid_encoding_name(value: &str) -> bool {
    let mut bytes = value.bytes();
    bytes.next().is_some_and(|byte| byte.is_ascii_alphabetic())
        && bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'_' | b'-'))
}

fn validate_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    family: Family,
) -> Result<Version> {
    if element.local_name().as_ref() != b"customUI" {
        return Err(Error::Invalid(
            "Ribbon XML root element must be customUI".into(),
        ));
    }
    let namespace = normalized_namespace(namespace, decoder, "root element")?;
    match (family, namespace.as_ref()) {
        (Family::Legacy, V2007_NAMESPACE) => Ok(Version::V2007),
        (Family::Modern, V2010_NAMESPACE) => Ok(Version::V2010),
        (Family::Modern, UI2_NAMESPACE) => Ok(Version::Ui2),
        _ => Err(Error::Invalid(
            "Ribbon XML root namespace does not match its package relationship".into(),
        )),
    }
}

fn count_node(nodes: &mut usize, limits: &Limits) -> Result<()> {
    *nodes = nodes.checked_add(1).ok_or(Error::Limit {
        resource: "Ribbon XML nodes",
        max: limits.nodes,
        actual: usize::MAX,
    })?;
    if *nodes > limits.nodes {
        return Err(Error::Limit {
            resource: "Ribbon XML nodes",
            max: limits.nodes,
            actual: *nodes,
        });
    }
    Ok(())
}

pub(super) fn validate_images(
    package: &OpcPackage,
    ribbon: &dyn Part,
    count: &mut usize,
    limits: &Limits,
) -> Result<()> {
    for relationship in ribbon.rels().iter() {
        if !matches!(relationship.reltype(), rt::IMAGE | rt::STRICT_IMAGE) {
            return Err(Error::Relationship(format!(
                "Ribbon part '{}' may relate only to image parts; '{}' has type '{}'",
                ribbon.partname().as_str(),
                relationship.r_id(),
                relationship.reltype()
            )));
        }
        *count = count.checked_add(1).ok_or(Error::Limit {
            resource: "Ribbon image relationships",
            max: limits.images,
            actual: usize::MAX,
        })?;
        if *count > limits.images {
            return Err(Error::Limit {
                resource: "Ribbon image relationships",
                max: limits.images,
                actual: *count,
            });
        }
        require_internal_target(relationship, "Ribbon image")?;
        let target = relationship.target_partname().map_err(|error| {
            Error::Relationship(format!("invalid Ribbon image target: {error}"))
        })?;
        let image = package.get_part(&target).map_err(|error| {
            Error::Missing(format!(
                "Ribbon image part '{}' does not exist: {error}",
                target.as_str()
            ))
        })?;
        if !is_image_content_type(image.content_type()) {
            return Err(Error::ContentType {
                expected: "image/*".to_owned(),
                actual: image.content_type().to_owned(),
            });
        }
    }
    Ok(())
}

pub(super) fn require_internal_target(
    relationship: &litchi_opc::Relationship,
    context: &str,
) -> Result<()> {
    if relationship.is_external() {
        return Err(Error::Relationship(format!(
            "{context} relationship '{}' must be internal",
            relationship.r_id()
        )));
    }
    if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
        return Err(Error::Relationship(format!(
            "{context} relationship '{}' target cannot contain a query or fragment",
            relationship.r_id()
        )));
    }
    Ok(())
}

pub(super) fn require_content_type(part: &dyn Part, expected: &str) -> Result<()> {
    if part.content_type() == expected {
        Ok(())
    } else {
        Err(Error::ContentType {
            expected: expected.to_owned(),
            actual: part.content_type().to_owned(),
        })
    }
}

fn is_image_content_type(value: &str) -> bool {
    if !valid_content_type(value) {
        return false;
    }
    let essence = value.split(';').next().unwrap_or(value);
    let Some((kind, subtype)) = essence.split_once('/') else {
        return false;
    };
    kind.eq_ignore_ascii_case("image") && is_mime_token(subtype)
}

fn valid_content_type(value: &str) -> bool {
    let mut components = value.split(';');
    let Some((kind, subtype)) = components.next().unwrap_or_default().split_once('/') else {
        return false;
    };
    if !is_mime_token(kind) || !is_mime_token(subtype) {
        return false;
    }
    components.all(|parameter| {
        let Some((name, raw_value)) = parameter.split_once('=') else {
            return false;
        };
        if !is_mime_token(name) {
            return false;
        }
        let value = raw_value
            .strip_prefix('"')
            .and_then(|value| value.strip_suffix('"'))
            .unwrap_or(raw_value);
        is_mime_token(value)
            && (!raw_value.contains('"')
                || (raw_value.starts_with('"') && raw_value.ends_with('"')))
    })
}

fn is_mime_token(value: &str) -> bool {
    !value.is_empty()
        && value.bytes().all(|byte| {
            byte.is_ascii_alphanumeric()
                || matches!(
                    byte,
                    b'!' | b'#'
                        | b'$'
                        | b'%'
                        | b'&'
                        | b'\''
                        | b'*'
                        | b'+'
                        | b'-'
                        | b'.'
                        | b'^'
                        | b'_'
                        | b'`'
                        | b'|'
                        | b'~'
                )
        })
}
