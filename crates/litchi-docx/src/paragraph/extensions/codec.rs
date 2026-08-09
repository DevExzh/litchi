#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Namespace-aware XML codec for paragraph and table-row extensions.

use crate::error::{Error, Result};
use crate::namespace::is_wordprocessing_namespace;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::fmt::Write as FmtWrite;

use super::model::{Extensions, Id, Ids, WORD_2010_NAMESPACE};
use super::validation::{parse_id, parse_on_off};

const MC_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/// Parse paragraph-level extension attributes from one complete `w:p`.
pub(crate) fn parse_paragraph(xml: &[u8]) -> Result<Extensions> {
    let (ids, no_spell_err) = parse_root(xml, b"p", true)?;
    let mut value = Extensions::new();
    value.set_ids(ids)?;
    value.set_no_spell_err(no_spell_err);
    Ok(value)
}

/// Parse row-level extension attributes from one complete `w:tr`.
pub(crate) fn parse_row(xml: &[u8]) -> Result<Ids> {
    parse_root(xml, b"tr", false).map(|(ids, _)| ids)
}

fn parse_root(
    xml: &[u8],
    root_name: &[u8],
    allow_no_spell_err: bool,
) -> Result<(Ids, Option<bool>)> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut root_closed = false;
    let mut para_id = None;
    let mut text_id = None;
    let mut no_spell_err = None;

    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let decoder = reader.decoder();
        let (namespace, event) = resolver.resolve_event(event);

        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if saw_root || root_closed || !is_word_root(&namespace, &element, root_name) {
                        return Err(Error::InvalidFormat(format!(
                            "Word extension XML must have one {root_name:?} root"
                        )));
                    }
                    parse_attributes(
                        &element,
                        &resolver,
                        decoder,
                        allow_no_spell_err,
                        &mut para_id,
                        &mut text_id,
                        &mut no_spell_err,
                    )?;
                    saw_root = true;
                    depth = 1;
                } else {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("Word extension XML nesting is too deep".into())
                    })?;
                }
            },
            Event::Empty(element) => {
                if depth == 0 {
                    if saw_root || root_closed || !is_word_root(&namespace, &element, root_name) {
                        return Err(Error::InvalidFormat(format!(
                            "Word extension XML must have one {root_name:?} root"
                        )));
                    }
                    parse_attributes(
                        &element,
                        &resolver,
                        decoder,
                        allow_no_spell_err,
                        &mut para_id,
                        &mut text_id,
                        &mut no_spell_err,
                    )?;
                    saw_root = true;
                    root_closed = true;
                }
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("Word extension XML has an unexpected end".into())
                })?;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(text) if depth == 0 => {
                let value = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                if !value.trim().is_empty() {
                    return Err(Error::InvalidFormat(
                        "Word extension XML has text outside its root".into(),
                    ));
                }
            },
            Event::CData(text) if depth == 0 => {
                if !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(
                        "Word extension XML has content outside its root".into(),
                    ));
                }
            },
            Event::Decl(_) | Event::Comment(_) if !saw_root => {},
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::InvalidFormat(
                    "Word extension XML cannot contain a DTD or processing instruction".into(),
                ));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }

    if !saw_root || depth != 0 || !root_closed {
        return Err(Error::InvalidFormat(
            "Word extension XML is missing a complete root".into(),
        ));
    }
    let ids = Ids::from_parts(para_id, text_id)?;
    Ok((ids, no_spell_err))
}

fn is_word_root(namespace: &ResolveResult<'_>, element: &BytesStart<'_>, expected: &[u8]) -> bool {
    is_wordprocessing_namespace(namespace) && element.local_name().as_ref() == expected
}

fn parse_attributes(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: Decoder,
    allow_no_spell_err: bool,
    para_id: &mut Option<Id>,
    text_id: &mut Option<Id>,
    no_spell_err: &mut Option<bool>,
) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let local = attribute.key.local_name();
        let known = matches!(local.as_ref(), b"paraId" | b"textId" | b"noSpellErr");
        if !known {
            continue;
        }

        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_word2010_namespace(&namespace) {
            return Err(Error::InvalidFormat(format!(
                "Word extension attribute '{}' is not in the Word 2010 wordml namespace",
                String::from_utf8_lossy(local.as_ref())
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        match local.as_ref() {
            b"paraId" => {
                if para_id.is_some() {
                    return Err(Error::InvalidFormat(
                        "paragraph has duplicate Word 2010 paraId attributes".into(),
                    ));
                }
                *para_id = Some(parse_id(&value, "paraId")?);
            },
            b"textId" => {
                if text_id.is_some() {
                    return Err(Error::InvalidFormat(
                        "paragraph has duplicate Word 2010 textId attributes".into(),
                    ));
                }
                *text_id = Some(parse_id(&value, "textId")?);
            },
            b"noSpellErr" => {
                if !allow_no_spell_err {
                    return Err(Error::InvalidFormat(
                        "Word 2010 noSpellErr is only valid on a paragraph".into(),
                    ));
                }
                if no_spell_err.is_some() {
                    return Err(Error::InvalidFormat(
                        "paragraph has duplicate Word 2010 noSpellErr attributes".into(),
                    ));
                }
                *no_spell_err = Some(parse_on_off(&value)?);
            },
            _ => unreachable!(),
        }
    }
    Ok(())
}

fn is_word2010_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value)) if *value == WORD_2010_NAMESPACE.as_bytes()
    )
}

/// Append paragraph extension attributes to a generated `w:p` start tag.
pub(crate) fn append_paragraph_attributes(
    value: &Extensions,
    requires_w14: bool,
    xml: &mut String,
) -> Result<()> {
    value.validate()?;
    if value.is_empty() && !requires_w14 {
        return Ok(());
    }
    append_namespace_attributes(xml)?;
    if let Some(id) = value.ids().para_id() {
        write!(xml, " w14:paraId=\"{id}\"")?;
    }
    if let Some(id) = value.ids().text_id() {
        write!(xml, " w14:textId=\"{id}\"")?;
    }
    if let Some(value) = value.no_spell_err() {
        write!(xml, " w14:noSpellErr=\"{}\"", i32::from(value))?;
    }
    Ok(())
}

/// Append table-row identifier attributes to a generated `w:tr` start tag.
pub(crate) fn append_row_attributes(value: &Ids, xml: &mut String) -> Result<()> {
    value.validate()?;
    if value.para_id().is_none() && value.text_id().is_none() {
        return Ok(());
    }
    append_namespace_attributes(xml)?;
    if let Some(id) = value.para_id() {
        write!(xml, " w14:paraId=\"{id}\"")?;
    }
    if let Some(id) = value.text_id() {
        write!(xml, " w14:textId=\"{id}\"")?;
    }
    Ok(())
}

fn append_namespace_attributes(xml: &mut String) -> Result<()> {
    write!(
        xml,
        " xmlns:w14=\"{WORD_2010_NAMESPACE}\" xmlns:mc=\"{MC_NAMESPACE}\" mc:Ignorable=\"w14\""
    )?;
    Ok(())
}
