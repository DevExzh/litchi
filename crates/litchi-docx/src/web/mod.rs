//! Layered WordprocessingML web-settings models, codec, and OPC owner.
//!
//! The historical `litchi_docx::web` module remains the public facade.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use codec::{parse, read, write};
pub use model::{
    Border, Borders, Child, Color, Conformance, Div, Frame, Frameset, Id, Key, Layout, Screen,
    Scrollbar, Settings, SplitBar, Twips,
};
pub use package::{load, put, remove};

pub(super) use crate::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};

pub(super) const WORD_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub(super) const STRICT_WORD_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
pub(super) const CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml";
pub(super) const STRICT_OFFICE_DOCUMENT_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
pub(super) const STRICT_WEB_SETTINGS_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/webSettings";
pub(super) const MAX_XML_BYTES: usize = 8 * 1024 * 1024;
pub(super) const MAX_XML_EVENTS: usize = 262_144;
pub(super) const MAX_TEXT_BYTES: usize = 64 * 1024;
pub(super) const MAX_FRAMESET_NESTING: usize = 128;

#[derive(Default)]
pub(super) struct ParseBudget {
    events: usize,
}

impl ParseBudget {
    pub(super) fn event(&mut self) -> Result<()> {
        self.events = self
            .events
            .checked_add(1)
            .ok_or_else(|| invalid("web-settings event count overflow"))?;
        if self.events > MAX_XML_EVENTS {
            return Err(invalid(format!(
                "web-settings XML exceeds {MAX_XML_EVENTS} events"
            )));
        }
        Ok(())
    }
}

pub(super) fn is_web_settings_relationship(value: &str) -> bool {
    value == litchi_opc::constants::relationship_type::WEB_SETTINGS
        || value == STRICT_WEB_SETTINGS_RELATIONSHIP
}

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(super) fn reserve_one<T>(values: &mut Vec<T>, resource: &'static str) -> Result<()> {
    values
        .try_reserve(1)
        .map_err(|source| Error::Allocation { resource, source })
}

pub(super) fn is_wordprocessing_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == WORD_NAMESPACE || *value == STRICT_WORD_NAMESPACE
    )
}

pub(super) fn word_attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let is_word_attribute = is_wordprocessing_namespace(&namespace)
            || matches!(namespace, ResolveResult::Unbound)
            || matches!(namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"w");
        if !is_word_attribute {
            continue;
        }
        if value.is_some() {
            let name = std::str::from_utf8(name)
                .map_err(|_| invalid("Word attribute name is not UTF-8"))?;
            return Err(invalid(format!("duplicate Word attribute '{name}'")));
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
