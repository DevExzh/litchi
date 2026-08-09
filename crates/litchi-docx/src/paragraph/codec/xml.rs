#![expect(
    clippy::option_option,
    reason = "nested options distinguish omitted, present-empty, and present-valued XML"
)]
#![expect(
    clippy::ref_option,
    reason = "the public API shape is retained for compatibility"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::shadow_unrelated,
    reason = "local parser names mirror the OOXML role currently being decoded"
)]
//! Shared `WordprocessingML` namespace and attribute primitives.

use crate::error::{Error, Result};
use crate::namespace::is_wordprocessing_namespace;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesEnd, BytesStart};
use quick_xml::name::{NamespaceResolver, QName, ResolveResult};

pub(crate) fn is_fragment_word_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local_name: &[u8],
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> bool {
    if name.local_name().as_ref() != local_name {
        return false;
    }
    if is_wordprocessing_namespace(namespace) {
        return true;
    }
    match namespace {
        ResolveResult::Unknown(prefix) => {
            fragment_prefix
                .as_ref()
                .and_then(|prefix| prefix.as_deref())
                == Some(prefix.as_slice())
        },
        ResolveResult::Unbound => fragment_prefix == &Some(None),
        ResolveResult::Bound(_) => false,
    }
}

pub(super) fn element_prefix(element: &BytesStart<'_>) -> Vec<u8> {
    let name = element.name();
    let raw = name.as_ref();
    raw.iter()
        .position(|byte| *byte == b':')
        .map_or_else(Vec::new, |index| raw[..index].to_vec())
}

pub(super) fn same_word_prefix(element: &BytesStart<'_>, prefix: Option<&[u8]>) -> bool {
    prefix.is_some_and(|prefix| element_prefix(element) == prefix)
}

pub(super) fn same_word_prefix_end(element: &BytesEnd<'_>, prefix: Option<&[u8]>) -> bool {
    let name = element.name();
    let raw = name.as_ref();
    let end = raw.iter().position(|byte| *byte == b':').unwrap_or(0);
    prefix.is_some_and(|prefix| {
        if end == 0 {
            prefix.is_empty()
        } else {
            &raw[..end] == prefix
        }
    })
}

pub(super) fn paragraph_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> Result<String> {
    for attribute in element.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() == name {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map(std::borrow::Cow::into_owned)
                .map_err(|error| Error::Xml(error.to_string()));
        }
    }
    Err(Error::InvalidFormat(format!(
        "paragraph property is missing '{}'",
        String::from_utf8_lossy(name)
    )))
}

pub(super) fn word_attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_fragment_word_name(&namespace, attribute.key, name, fragment_prefix) {
            continue;
        }
        if value.is_some() {
            return Err(Error::InvalidFormat(format!(
                "duplicate Word attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}

pub(super) fn set_paragraph_property<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.is_some() {
        return Err(Error::InvalidFormat(format!(
            "paragraph has duplicate {name}"
        )));
    }
    *slot = Some(value);
    Ok(())
}

#[inline]
pub(super) fn is_on(value: &[u8]) -> bool {
    matches!(value, b"true" | b"1" | b"on")
}
