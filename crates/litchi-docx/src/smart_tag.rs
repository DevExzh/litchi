#![expect(
    clippy::module_name_repetitions,
    reason = "public names retain established OOXML facade terminology"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! `WordprocessingML` smart-tag metadata.

use crate::error::{Error, Result};
use crate::namespace::{NamespaceBindings, is_wordprocessing_namespace};
use crate::paragraph::{extract_word_text_with_bindings, is_fragment_word_name};
use litchi_core::XmlSlice;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, PrefixDeclaration, ResolveResult};
use quick_xml::reader::NsReader;

/// One custom attribute attached to a smart tag.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SmartTagAttribute {
    /// Optional namespace URI for the attribute.
    pub uri: Option<String>,
    /// Attribute name.
    pub name: String,
    /// Attribute value.
    pub value: String,
}

/// A run-level Word smart tag and its custom metadata.
#[derive(Debug, Clone)]
pub struct SmartTag {
    /// Optional namespace URI identifying the smart-tag vocabulary.
    pub uri: Option<String>,
    /// Required smart-tag element name.
    pub element: String,
    /// Custom attributes from the optional `smartTagPr` container.
    pub attributes: Vec<SmartTagAttribute>,
    xml: XmlSlice,
    namespaces: NamespaceBindings,
}

impl SmartTag {
    pub(crate) fn parse(xml: XmlSlice, namespaces: NamespaceBindings) -> Result<Self> {
        let mut reader = NsReader::from_reader(xml.as_bytes());
        for (prefix, namespace) in namespaces.iter() {
            let prefix = prefix
                .as_deref()
                .map_or(PrefixDeclaration::Default, PrefixDeclaration::Named);
            reader
                .resolver_mut()
                .add(prefix, Namespace(namespace))
                .map_err(|_| Error::Xml("invalid inherited smart-tag namespace".into()))?;
        }
        let mut fragment_prefix = None;
        let mut uri = None;
        let mut element = None;
        let mut attributes = Vec::new();
        let mut saw_root = false;
        let mut properties_depth = None;
        let mut depth = 0usize;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            let resolver = reader.resolver();
            let (namespace, event) = resolver.resolve_event(event);
            if !saw_root
                && depth == 0
                && let Event::Start(start) | Event::Empty(start) = &event
                && !matches!(namespace, ResolveResult::Bound(_))
            {
                fragment_prefix = Some(start.name().prefix().map(|p| p.into_inner().to_vec()));
            }
            let is_word = |start: &BytesStart<'_>, name: &[u8]| {
                is_fragment_word_name(&namespace, start.name(), name, &fragment_prefix)
            };
            match event {
                Event::Start(start) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("smart-tag XML nesting is too deep".into())
                    })?;
                    match start.local_name().as_ref() {
                        b"smartTag" if depth == 1 && !saw_root && is_word(&start, b"smartTag") => {
                            saw_root = true;
                            uri = optional_attribute(
                                &start,
                                b"uri",
                                decoder,
                                resolver,
                                &fragment_prefix,
                            )?;
                            element = optional_attribute(
                                &start,
                                b"element",
                                decoder,
                                resolver,
                                &fragment_prefix,
                            )?;
                        },
                        b"smartTagPr"
                            if depth == 2
                                && saw_root
                                && properties_depth.is_none()
                                && is_word(&start, b"smartTagPr") =>
                        {
                            properties_depth = Some(depth);
                        },
                        b"attr"
                            if properties_depth.is_some_and(|value| depth == value + 1)
                                && is_word(&start, b"attr") =>
                        {
                            attributes.push(parse_attribute(
                                &start,
                                decoder,
                                resolver,
                                &fragment_prefix,
                            )?);
                        },
                        _ => {},
                    }
                },
                Event::Empty(start)
                    if is_word(&start, b"attr")
                        && properties_depth.is_some_and(|value| depth == value) =>
                {
                    attributes.push(parse_attribute(
                        &start,
                        decoder,
                        resolver,
                        &fragment_prefix,
                    )?);
                },
                Event::Empty(start) if depth == 0 && is_word(&start, b"smartTag") && !saw_root => {
                    saw_root = true;
                    uri = optional_attribute(&start, b"uri", decoder, resolver, &fragment_prefix)?;
                    element = optional_attribute(
                        &start,
                        b"element",
                        decoder,
                        resolver,
                        &fragment_prefix,
                    )?;
                },
                Event::End(end) => {
                    if properties_depth == Some(depth) && end.local_name().as_ref() == b"smartTagPr"
                    {
                        properties_depth = None;
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("invalid smart-tag XML nesting".into())
                    })?;
                },
                Event::Eof if depth != 0 => {
                    return Err(Error::InvalidFormat("unterminated smart-tag XML".into()));
                },
                Event::Eof => break,
                Event::Empty(_)
                | Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_) => {},
            }
        }

        if !saw_root {
            return Err(Error::InvalidFormat(
                "smart-tag XML has no smartTag root".into(),
            ));
        }
        Ok(Self {
            uri,
            element: element.ok_or_else(|| {
                Error::InvalidFormat("Word smart tag is missing its element attribute".into())
            })?,
            attributes,
            xml,
            namespaces,
        })
    }

    /// Extract the visible text contained by this smart tag.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn text(&self) -> Result<String> {
        extract_word_text_with_bindings(self.xml.as_bytes(), &self.namespaces)
    }

    /// Return the original smart-tag XML fragment.
    #[inline]
    #[must_use]
    pub fn raw_xml(&self) -> &[u8] {
        self.xml.as_bytes()
    }
}

impl PartialEq for SmartTag {
    fn eq(&self, other: &Self) -> bool {
        self.uri == other.uri
            && self.element == other.element
            && self.attributes == other.attributes
            && self.raw_xml() == other.raw_xml()
    }
}

impl Eq for SmartTag {}

fn parse_attribute(
    start: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> Result<SmartTagAttribute> {
    Ok(SmartTagAttribute {
        uri: optional_attribute(start, b"uri", decoder, resolver, fragment_prefix)?,
        name: optional_attribute(start, b"name", decoder, resolver, fragment_prefix)?.ok_or_else(
            || Error::InvalidFormat("Word smart-tag property is missing its name".into()),
        )?,
        value: optional_attribute(start, b"val", decoder, resolver, fragment_prefix)?.ok_or_else(
            || Error::InvalidFormat("Word smart-tag property is missing its value".into()),
        )?,
    })
}

fn optional_attribute(
    start: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in start.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        // Keep the existing unqualified-attribute compatibility of Word
        // fragments, while explicit foreign bindings never become metadata.
        if !is_wordprocessing_namespace(&namespace)
            && !matches!(namespace, ResolveResult::Unbound)
            && !is_fragment_word_name(&namespace, attribute.key, name, fragment_prefix)
        {
            continue;
        }
        if value.is_some() {
            return Err(Error::InvalidFormat(format!(
                "Word smart tag contains duplicate '{}' attributes",
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
