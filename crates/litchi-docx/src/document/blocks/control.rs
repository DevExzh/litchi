//! The identity of a block content control, read from its `w:sdtPr`.

use crate::error::{Error, Result};
use crate::namespace::{NamespaceBindings, is_wordprocessing_namespace};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, PrefixDeclaration};
use quick_xml::reader::NsReader;

/// The `w:id`, `w:tag` and `w:alias` of one `w:sdt`.
#[derive(Clone, Default)]
pub(super) struct Identity {
    pub(super) id: Option<i32>,
    pub(super) tag: Option<String>,
    pub(super) alias: Option<String>,
}

fn set_once<T>(target: &mut Option<T>, value: T, description: &str) -> Result<()> {
    if target.is_some() {
        return Err(Error::InvalidFormat(format!("duplicate {description}")));
    }
    *target = Some(value);
    Ok(())
}

fn value(element: &BytesStart<'_>, reader: &NsReader<&[u8]>, description: &str) -> Result<String> {
    let mut value = None;
    for attribute in element.checked_attributes() {
        let attribute =
            attribute.map_err(|_| Error::Xml("malformed content-control attribute".into()))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if local.as_ref() != b"val" || !is_wordprocessing_namespace(&namespace) {
            continue;
        }
        if value.is_some() {
            return Err(Error::InvalidFormat(
                "duplicate content-control value".into(),
            ));
        }
        let decoded = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|_| Error::Xml("invalid content-control value".into()))?;
        let mut owned = String::new();
        owned
            .try_reserve_exact(decoded.len())
            .map_err(|source| Error::Allocation {
                resource: "content-control value",
                source,
            })?;
        owned.push_str(&decoded);
        value = Some(owned);
    }
    value.ok_or_else(|| Error::InvalidFormat(format!("{description} lacks w:val")))
}

fn read_property(
    identity: &mut Identity,
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    is_word: bool,
) -> Result<()> {
    if !is_word {
        return Ok(());
    }
    match element.local_name().as_ref() {
        b"id" => {
            let text = value(element, reader, "content-control ID")?;
            let id = text
                .trim()
                .parse::<i32>()
                .map_err(|_| Error::InvalidFormat("invalid content-control ID value".into()))?;
            set_once(&mut identity.id, id, "content-control ID")
        },
        b"tag" => {
            let tag = value(element, reader, "content-control tag")?;
            set_once(&mut identity.tag, tag, "content-control tag")
        },
        b"alias" => {
            let alias = value(element, reader, "content-control alias")?;
            set_once(&mut identity.alias, alias, "content-control alias")
        },
        _ => Ok(()),
    }
}

/// Read the identity of the `w:sdt` element `fragment`, whose inherited
/// namespace bindings are `inherited`.
///
/// Validate the complete fragment and read only direct properties of its
/// root-level `w:sdtPr`.
pub(super) fn parse(fragment: &[u8], inherited: &NamespaceBindings) -> Result<Identity> {
    let mut reader = NsReader::from_reader(fragment);
    for (prefix, namespace) in inherited.iter() {
        let prefix = prefix
            .as_deref()
            .map_or(PrefixDeclaration::Default, PrefixDeclaration::Named);
        reader
            .resolver_mut()
            .add(prefix, Namespace(namespace))
            .map_err(|_| Error::Xml("invalid content-control namespace".into()))?;
    }
    let mut identity = Identity::default();
    let (mut depth, mut nodes) = (0usize, 0usize);
    let (mut root_seen, mut properties_seen, mut content_seen, mut in_properties) =
        (false, false, false, false);
    loop {
        let event = reader
            .read_event()
            .map_err(|_| Error::Xml("malformed content-control XML".into()))?;
        let (empty, element) = match event {
            Event::Start(element) => (false, element),
            Event::Empty(element) => (true, element),
            Event::End(_) => {
                if depth == 2 {
                    in_properties = false;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::Xml("invalid content-control nesting".into()))?;
                continue;
            },
            Event::Eof => {
                if root_seen && depth == 0 {
                    return Ok(identity);
                }
                return Err(Error::Xml("incomplete content-control XML".into()));
            },
            Event::DocType(_) | Event::PI(_) | Event::Decl(_) => {
                return Err(Error::Xml("prohibited content-control declaration".into()));
            },
            Event::Text(text)
                if depth == 0
                    && !crate::parts::document_part::is_xml_outer_whitespace(text.as_ref()) =>
            {
                return Err(Error::Xml("text outside content-control root".into()));
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(Error::Xml("text outside content-control root".into()));
            },
            Event::Text(text) => {
                super::scan::validate_text(
                    &text
                        .xml10_content()
                        .map_err(|_| Error::Xml("invalid control text".into()))?,
                )?;
                continue;
            },
            Event::CData(text) => {
                super::scan::validate_text(
                    &text
                        .decode()
                        .map_err(|_| Error::Xml("invalid control CDATA".into()))?,
                )?;
                continue;
            },
            Event::GeneralRef(reference) => {
                super::scan::validate_text(
                    &litchi_ooxml_common::xml::decode_xml_reference(&reference)
                        .map_err(|_| Error::Xml("invalid content-control reference".into()))?,
                )?;
                continue;
            },
            _ => continue,
        };
        nodes += 1;
        if nodes > 1_000_000 || depth >= 256 {
            return Err(Error::BlockReadLimit {
                resource: "control traversal",
                actual: if depth >= 256 { depth + 1 } else { nodes },
                maximum: if depth >= 256 { 256 } else { 1_000_000 },
            });
        }
        super::scan::validate_element(&reader, &element)?;
        let (namespace, local) = reader.resolver().resolve_element(element.name());
        let is_word = is_wordprocessing_namespace(&namespace);
        match depth {
            0 => {
                if root_seen || !is_word || local.as_ref() != b"sdt" {
                    return Err(Error::Xml("unexpected content-control root".into()));
                }
                root_seen = true;
            },
            1 if is_word && local.as_ref() == b"sdtPr" => {
                if properties_seen || content_seen {
                    return Err(Error::Xml(
                        "duplicate or out-of-order content-control properties".into(),
                    ));
                }
                properties_seen = true;
                in_properties = !empty;
            },
            1 if is_word && local.as_ref() == b"sdtContent" => {
                if content_seen {
                    return Err(Error::Xml("duplicate content-control content".into()));
                }
                content_seen = true;
            },
            2 if in_properties => read_property(&mut identity, &element, &reader, is_word)?,
            _ => {},
        }
        if !empty {
            depth += 1;
        }
    }
}

impl std::fmt::Debug for Identity {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str("Identity { .. }")
    }
}
