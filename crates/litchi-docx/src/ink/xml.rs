//! Borrowed, bounded XML token validation for the DOCX Ink story scanner.

use std::str;

use litchi_ooxml_common::xml_name::is_qualified_name;
use quick_xml::{
    XmlVersion,
    events::{BytesDecl, BytesStart},
    name::{Namespace, NamespaceResolver, QName, ResolveResult},
};

use crate::{Error, Result};

const MAX_ATTRIBUTES: usize = 256;
const MAX_ATTRIBUTE_BYTES: usize = 1024 * 1024;
const MAX_DECLARATION_BYTES: usize = 4096;
const MAX_QNAME_BYTES: usize = 4096;
const MAX_REFERENCE_BYTES: usize = 4096;
const MAX_TEXT_BYTES: usize = 1024 * 1024;

/// Validate borrowed XML text or CDATA content.
pub(super) fn text(bytes: &[u8]) -> Result<&str> {
    if bytes.len() > MAX_TEXT_BYTES {
        return Err(limit("XML text", bytes.len(), MAX_TEXT_BYTES));
    }
    let value = str::from_utf8(bytes).map_err(|error| xml_error("XML text", error))?;
    validate_xml_characters(value, "XML text")?;
    Ok(value)
}

/// Validate one borrowed general reference name from a quick-xml event.
pub(super) fn reference(bytes: &[u8]) -> Result<()> {
    if bytes.len() > MAX_REFERENCE_BYTES {
        return Err(limit("XML reference", bytes.len(), MAX_REFERENCE_BYTES));
    }
    let value = str::from_utf8(bytes).map_err(|error| xml_error("XML reference", error))?;
    match value {
        "amp" | "lt" | "gt" | "apos" | "quot" => Ok(()),
        value if let Some(digits) = value.strip_prefix("#x") => {
            if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                return Err(invalid("XML hexadecimal character reference is invalid"));
            }
            let codepoint = u32::from_str_radix(digits, 16)
                .map_err(|_| invalid("XML hexadecimal character reference is invalid"))?;
            validate_reference_character(codepoint)
        },
        value if let Some(digits) = value.strip_prefix('#') => {
            if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
                return Err(invalid("XML decimal character reference is invalid"));
            }
            let codepoint = digits
                .parse::<u32>()
                .map_err(|_| invalid("XML decimal character reference is invalid"))?;
            validate_reference_character(codepoint)
        },
        _ => Err(invalid("unsupported XML general entity reference")),
    }
}

/// Validate an XML 1.0 declaration with a bounded, borrowed grammar scan.
pub(super) fn declaration(declaration: &BytesDecl<'_>) -> Result<()> {
    let raw = declaration.as_ref();
    if raw.len() > MAX_DECLARATION_BYTES {
        return Err(limit("XML declaration", raw.len(), MAX_DECLARATION_BYTES));
    }
    str::from_utf8(raw).map_err(|error| xml_error("XML declaration", error))?;

    let mut cursor = 0usize;
    skip_whitespace(raw, &mut cursor);
    if !consume(raw, &mut cursor, b"xml") || !has_whitespace(raw.get(cursor).copied()) {
        return Err(invalid("XML declaration is malformed"));
    }
    skip_whitespace(raw, &mut cursor);

    let (name, value) = parse_declaration_attribute(raw, &mut cursor)?;
    if name != b"version" || value != b"1.0" {
        return Err(invalid("XML declaration must start with version 1.0"));
    }
    let mut previous = b"version".as_slice();
    while cursor < raw.len() {
        if !has_whitespace(raw.get(cursor).copied()) {
            return Err(invalid("XML declaration attributes must be separated"));
        }
        skip_whitespace(raw, &mut cursor);
        if cursor == raw.len() {
            break;
        }
        let (name, value) = parse_declaration_attribute(raw, &mut cursor)?;
        let valid = match (previous, name) {
            (b"version", b"encoding") => {
                valid_encoding_name(value) && value.eq_ignore_ascii_case(b"UTF-8")
            },
            (b"version" | b"encoding", b"standalone") => matches!(value, b"yes" | b"no"),
            _ => false,
        };
        if !valid {
            return Err(invalid(
                "XML declaration has an invalid, duplicate, or out-of-order attribute",
            ));
        }
        previous = name;
    }
    Ok(())
}

/// Validate one borrowed XML start tag and its in-scope namespace bindings.
pub(super) fn element(element: &BytesStart<'_>, resolver: &NamespaceResolver) -> Result<()> {
    validate_qname(element.name().as_ref(), "XML element QName")?;
    if matches!(
        resolver.resolve_element(element.name()).0,
        ResolveResult::Unknown(_)
    ) {
        return Err(invalid("XML element uses an undeclared namespace prefix"));
    }

    let mut count = 0usize;
    let mut names = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        count = count
            .checked_add(1)
            .ok_or_else(|| limit("XML attributes", usize::MAX, MAX_ATTRIBUTES))?;
        if count > MAX_ATTRIBUTES {
            return Err(limit("XML attributes", count, MAX_ATTRIBUTES));
        }
        validate_qname(attribute.key.as_ref(), "XML attribute QName")?;
        if attribute.value.len() > MAX_ATTRIBUTE_BYTES {
            return Err(limit(
                "XML attribute",
                attribute.value.len(),
                MAX_ATTRIBUTE_BYTES,
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, element.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?;
        if value.len() > MAX_ATTRIBUTE_BYTES {
            return Err(limit("XML attribute", value.len(), MAX_ATTRIBUTE_BYTES));
        }
        validate_xml_characters(&value, "XML attribute")?;

        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        let expanded = expanded_attribute_name(attribute.key, resolver)?;
        if names.iter().any(|seen| same_expanded_name(*seen, expanded)) {
            return Err(invalid("XML element has duplicate expanded attributes"));
        }
        names.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "DOCX XML expanded attributes",
            source,
        })?;
        names.push(expanded);
    }
    Ok(())
}

#[derive(Clone, Copy)]
struct ExpandedAttributeName<'namespace, 'local> {
    namespace: Option<&'namespace [u8]>,
    local: &'local [u8],
}

fn expanded_attribute_name<'namespace, 'local>(
    key: QName<'local>,
    resolver: &'namespace NamespaceResolver,
) -> Result<ExpandedAttributeName<'namespace, 'local>> {
    let (namespace, local) = resolver.resolve_attribute(key);
    let namespace = match namespace {
        ResolveResult::Unbound => None,
        ResolveResult::Bound(Namespace(value)) => Some(value),
        ResolveResult::Unknown(_) => {
            return Err(invalid("XML attribute uses an undeclared namespace prefix"));
        },
    };
    Ok(ExpandedAttributeName {
        namespace,
        local: local.into_inner(),
    })
}

fn same_expanded_name(
    left: ExpandedAttributeName<'_, '_>,
    right: ExpandedAttributeName<'_, '_>,
) -> bool {
    left.local == right.local && left.namespace == right.namespace
}

fn validate_qname(name: &[u8], resource: &'static str) -> Result<()> {
    if name.is_empty() || name.len() > MAX_QNAME_BYTES {
        return Err(limit(resource, name.len(), MAX_QNAME_BYTES));
    }
    let value = str::from_utf8(name).map_err(|error| xml_error(resource, error))?;
    if is_qualified_name(value) {
        Ok(())
    } else {
        Err(invalid("XML contains an invalid QName"))
    }
}

fn validate_xml_characters(value: &str, resource: &'static str) -> Result<()> {
    if value.chars().all(is_xml10_character) {
        Ok(())
    } else {
        Err(invalid_with_resource(
            resource,
            "contains a character forbidden by XML 1.0",
        ))
    }
}

const fn is_xml10_character(value: char) -> bool {
    matches!(
        value,
        '\u{9}'
            | '\u{A}'
            | '\u{D}'
            | '\u{20}'..='\u{D7FF}'
            | '\u{E000}'..='\u{FFFD}'
            | '\u{10000}'..='\u{10FFFF}'
    )
}

fn validate_reference_character(codepoint: u32) -> Result<()> {
    let character = char::from_u32(codepoint)
        .ok_or_else(|| invalid("XML character reference is not a Unicode scalar"))?;
    if is_xml10_character(character) {
        Ok(())
    } else {
        Err(invalid(
            "XML character reference is not an XML 1.0 character",
        ))
    }
}

fn parse_declaration_attribute<'a>(
    raw: &'a [u8],
    cursor: &mut usize,
) -> Result<(&'a [u8], &'a [u8])> {
    let name_start = *cursor;
    while raw
        .get(*cursor)
        .copied()
        .is_some_and(|byte| byte.is_ascii_alphabetic())
    {
        *cursor += 1;
    }
    if *cursor == name_start {
        return Err(invalid("XML declaration attribute name is missing"));
    }
    let name = &raw[name_start..*cursor];
    skip_whitespace(raw, cursor);
    if raw.get(*cursor) != Some(&b'=') {
        return Err(invalid("XML declaration attribute equals sign is missing"));
    }
    *cursor += 1;
    skip_whitespace(raw, cursor);
    let quote = raw
        .get(*cursor)
        .copied()
        .filter(|value| matches!(value, b'\'' | b'"'))
        .ok_or_else(|| invalid("XML declaration attribute quote is missing"))?;
    *cursor += 1;
    let value_start = *cursor;
    while raw
        .get(*cursor)
        .copied()
        .is_some_and(|value| value != quote)
    {
        *cursor += 1;
    }
    if raw.get(*cursor) != Some(&quote) {
        return Err(invalid("XML declaration attribute is unterminated"));
    }
    let value = &raw[value_start..*cursor];
    *cursor += 1;
    Ok((name, value))
}

fn valid_encoding_name(value: &[u8]) -> bool {
    value.first().is_some_and(u8::is_ascii_alphabetic)
        && value[1..]
            .iter()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'_' | b'-'))
}

const fn has_whitespace(value: Option<u8>) -> bool {
    matches!(value, Some(b' ' | b'\t' | b'\r' | b'\n'))
}

fn skip_whitespace(raw: &[u8], cursor: &mut usize) {
    while has_whitespace(raw.get(*cursor).copied()) {
        *cursor += 1;
    }
}

fn consume(raw: &[u8], cursor: &mut usize, token: &[u8]) -> bool {
    raw.get(*cursor..)
        .is_some_and(|remaining| remaining.starts_with(token))
        .then(|| *cursor += token.len())
        .is_some()
}

fn xml_error(resource: &'static str, error: impl std::fmt::Display) -> Error {
    Error::Xml(format!("invalid {resource}: {error}"))
}

fn invalid(message: &'static str) -> Error {
    Error::Invalid(message.into())
}

fn invalid_with_resource(resource: &'static str, message: &'static str) -> Error {
    Error::Invalid(format!("{resource} {message}"))
}

fn limit(resource: &'static str, actual: usize, maximum: usize) -> Error {
    Error::InkLimit {
        resource,
        actual,
        maximum,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn start(content: &str, name_len: usize) -> (BytesStart<'_>, NamespaceResolver) {
        let element = BytesStart::from_content(content, name_len);
        let mut resolver = NamespaceResolver::default();
        resolver.push(&element).expect("test namespace bindings");
        (element, resolver)
    }

    #[test]
    fn text_rejects_invalid_utf8_and_xml_controls() {
        assert_eq!(text("hello".as_bytes()).expect("valid text"), "hello");
        assert!(text(b"bad\0text").is_err());
        assert!(text(&[0xff]).is_err());
    }

    #[test]
    fn references_require_predefined_names_or_unsigned_numeric_digits() {
        for value in [b"amp".as_slice(), b"#65".as_slice(), b"#x41".as_slice()] {
            assert!(reference(value).is_ok(), "{value:?}");
        }
        for value in [
            b"future".as_slice(),
            b"#".as_slice(),
            b"#x".as_slice(),
            b"#+65".as_slice(),
            b"#x+41".as_slice(),
            b"#0".as_slice(),
            b"#x110000".as_slice(),
        ] {
            assert!(reference(value).is_err(), "{value:?}");
        }
    }

    #[test]
    fn declaration_checks_version_order_duplicates_encoding_and_standalone() {
        let valid = BytesDecl::from_start(BytesStart::from_content(
            "xml version='1.0' encoding=\"UTF-8\" standalone='yes'",
            3,
        ));
        assert!(declaration(&valid).is_ok());

        for raw in [
            "xml version='1.1'",
            "xml encoding='UTF-8' version='1.0'",
            "xml version='1.0' version='1.0'",
            "xml version='1.0' encoding='UTF-16'",
            "xml version='1.0' standalone='maybe'",
            "xml version='1.0' foo='bar'",
            "xml version='1.0'encoding='UTF-8'",
        ] {
            let declaration_event = BytesDecl::from_start(BytesStart::from_content(raw, 3));
            assert!(declaration(&declaration_event).is_err(), "{raw}");
        }
    }

    #[test]
    fn element_checks_decoded_values_expanded_duplicates_and_unknown_prefixes() {
        let (start_tag, resolver) = start(
            "root xmlns:a='urn:x' xmlns:b='urn:x' a:id='one' b:id='two'",
            4,
        );
        assert!(element(&start_tag, &resolver).is_err());

        let (start_tag, resolver) = start("root a:id='one'", 4);
        assert!(element(&start_tag, &resolver).is_err());

        let (start_tag, resolver) = start("root value='a&amp;&#65;'", 4);
        assert!(element(&start_tag, &resolver).is_ok());
    }

    #[test]
    fn element_rejects_invalid_attribute_values_and_attribute_quota() {
        let (start_tag, resolver) = start("root value='bad\0value'", 4);
        assert!(element(&start_tag, &resolver).is_err());

        let mut exact_content = String::from("root");
        for index in 0..MAX_ATTRIBUTES {
            exact_content.push_str(&format!(" a{index}='x'"));
        }
        let exact_start = BytesStart::from_content(&exact_content, 4);
        let exact_resolver = NamespaceResolver::default();
        assert!(element(&exact_start, &exact_resolver).is_ok());

        let mut over_content = String::from("root");
        for index in 0..=MAX_ATTRIBUTES {
            over_content.push_str(&format!(" a{index}='x'"));
        }
        let over_start = BytesStart::from_content(&over_content, 4);
        let over_resolver = NamespaceResolver::default();
        assert!(element(&over_start, &over_resolver).is_err());
    }
}
