#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Checked Word 2010 drawing-extension values.

use super::model::AnchorId;
use crate::error::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use std::fmt::Write as FmtWrite;

/// Namespace used by `[MS-DOCX]` 2.2.6 `anchorId` attributes.
pub(crate) const WORD_2010_WORDML_NAMESPACE: &[u8] =
    b"http://schemas.microsoft.com/office/word/2010/wordml";

/// Parse the lexical `ST_LongHexNumber`/`ST_EditId` form used by the checked
/// Word drawing identifiers.
pub(crate) fn parse_anchor_id_text(text: &str) -> Result<AnchorId> {
    if text.len() != 8 || !text.bytes().all(is_ascii_hex_digit) {
        return Err(Error::InvalidFormat(format!(
            "invalid anchorId '{text}'; expected eight hexadecimal digits"
        )));
    }
    let value = u32::from_str_radix(text, 16)
        .map_err(|error| Error::InvalidFormat(format!("invalid anchorId '{text}': {error}")))?;
    AnchorId::new(value).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "anchorId '{text}' is outside the nonzero, below-0x80000000 range"
        ))
    })
}

/// Parse a namespaced Word 2010 `anchorId` attribute from a legacy object or
/// picture root.
pub(crate) fn parse_word2010_anchor_id(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: Decoder,
) -> Result<Option<AnchorId>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"anchorId" {
            continue;
        }

        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_word2010_namespace(&namespace) {
            return Err(Error::InvalidFormat(
                "legacy anchorId is not in the Word 2010 wordml namespace".to_string(),
            ));
        }
        if value.is_some() {
            return Err(Error::InvalidFormat(
                "legacy drawing has duplicate anchorId attributes".to_string(),
            ));
        }

        let text = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        value = Some(parse_anchor_id_text(&text)?);
    }
    Ok(value)
}

/// Append a valid Word 2010 `anchorId` and the required ignorable namespace
/// declaration to a generated `w:object` or `w:pict` root.
pub(crate) fn append_word2010_anchor_id(
    xml: &mut String,
    anchor_id: Option<AnchorId>,
) -> Result<()> {
    let Some(anchor_id) = anchor_id else {
        return Ok(());
    };
    write!(
        xml,
        r#" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14" w14:anchorId="{:08x}""#,
        anchor_id.get()
    )
    .map_err(|error| Error::Xml(error.to_string()))?;
    Ok(())
}

#[inline]
fn is_ascii_hex_digit(value: u8) -> bool {
    value.is_ascii_digit() || matches!(value, b'a'..=b'f' | b'A'..=b'F')
}

#[inline]
fn is_word2010_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value)) if *value == WORD_2010_WORDML_NAMESPACE
    )
}
