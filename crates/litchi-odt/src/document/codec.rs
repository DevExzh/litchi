//! Namespace-aware XML codecs for ODT document references.

use crate::elements::xml::{
    DRAW_NAMESPACE, TEXT_NAMESPACE, XLINK_NAMESPACE, append_checked, append_text_control,
    decode_reference, is_bound, namespaced_attribute,
};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;
use std::collections::HashSet;

use super::validation;

struct ActiveHyperlink {
    href: Option<String>,
    text: String,
    depth: usize,
}

pub(super) fn parse_hyperlinks(xml: &str) -> Result<Vec<(String, String)>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut document_depth = 0usize;
    let mut active: Option<ActiveHyperlink> = None;
    let mut links = Vec::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid hyperlink XML: {error}")))?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                document_depth = validation::checked_reference_depth(document_depth)?;
                if let Some(link) = active.as_mut() {
                    if text_element && element.local_name().as_ref() == b"a" {
                        return Err(Error::InvalidFormat(
                            "nested text:a hyperlinks are not allowed".to_string(),
                        ));
                    }
                    link.depth += 1;
                    if text_element {
                        append_text_control(&reader, element, &mut link.text)?;
                    }
                } else if text_element && element.local_name().as_ref() == b"a" {
                    active = Some(ActiveHyperlink {
                        href: namespaced_attribute(
                            &reader,
                            element,
                            XLINK_NAMESPACE,
                            b"href",
                            "text:a",
                        )?,
                        text: String::new(),
                        depth: 1,
                    });
                }
            },
            Event::Empty(ref element) => {
                if let Some(link) = active.as_mut() {
                    if text_element && element.local_name().as_ref() == b"a" {
                        return Err(Error::InvalidFormat(
                            "nested text:a hyperlinks are not allowed".to_string(),
                        ));
                    }
                    if text_element {
                        append_text_control(&reader, element, &mut link.text)?;
                    }
                } else if text_element
                    && element.local_name().as_ref() == b"a"
                    && let Some(href) =
                        namespaced_attribute(&reader, element, XLINK_NAMESPACE, b"href", "text:a")?
                {
                    validation::ensure_reference_capacity(links.len(), "hyperlinks")?;
                    links.push((String::new(), href));
                }
            },
            Event::Text(ref value) if active.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid hyperlink text: {error}"))
                    })?;
                append_checked(
                    &mut active.as_mut().expect("checked hyperlink").text,
                    &value,
                )?;
            },
            Event::CData(ref value) if active.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid hyperlink CDATA: {error}"))
                    })?;
                append_checked(
                    &mut active.as_mut().expect("checked hyperlink").text,
                    &value,
                )?;
            },
            Event::GeneralRef(ref reference) if active.is_some() => {
                let value = decode_reference(reference, "hyperlink")?;
                append_checked(
                    &mut active.as_mut().expect("checked hyperlink").text,
                    &value,
                )?;
            },
            Event::End(_) => {
                document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("hyperlink XML stack underflow".to_string())
                })?;
                if let Some(link) = active.as_mut() {
                    link.depth = link.depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("hyperlink element stack underflow".to_string())
                    })?;
                    if link.depth == 0 {
                        let link = active.take().expect("checked hyperlink");
                        if let Some(href) = link.href {
                            validation::ensure_reference_capacity(links.len(), "hyperlinks")?;
                            links.push((link.text, href));
                        }
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if document_depth != 0 || active.is_some() {
        return Err(Error::InvalidFormat(
            "incomplete hyperlink XML structure".to_string(),
        ));
    }
    Ok(links)
}

pub(super) fn parse_bookmark_names(xml: &str) -> Result<Vec<String>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut names = Vec::new();
    let mut unique_names = HashSet::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid bookmark XML: {error}")))?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                depth = validation::checked_reference_depth(depth)?;
                collect_bookmark_name(
                    &reader,
                    text_element,
                    element,
                    &mut names,
                    &mut unique_names,
                )?;
            },
            Event::Empty(ref element) => {
                collect_bookmark_name(
                    &reader,
                    text_element,
                    element,
                    &mut names,
                    &mut unique_names,
                )?;
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("bookmark XML stack underflow".to_string())
                })?;
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 {
        return Err(Error::InvalidFormat(
            "incomplete bookmark XML structure".to_string(),
        ));
    }
    Ok(names)
}

fn collect_bookmark_name(
    reader: &NsReader<&[u8]>,
    text_element: bool,
    element: &quick_xml::events::BytesStart<'_>,
    names: &mut Vec<String>,
    unique_names: &mut HashSet<String>,
) -> Result<()> {
    if text_element
        && matches!(
            element.local_name().as_ref(),
            b"bookmark" | b"bookmark-start" | b"bookmark-end"
        )
        && let Some(name) =
            namespaced_attribute(reader, element, TEXT_NAMESPACE, b"name", "bookmark")?
        && unique_names.insert(name.clone())
    {
        validation::ensure_reference_capacity(names.len(), "bookmark names")?;
        names.push(name);
    }
    Ok(())
}

pub(super) fn parse_image_references(xml: &str) -> Result<Vec<String>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut references = Vec::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid image XML: {error}")))?;
        let draw_element = is_bound(&namespace, DRAW_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                depth = validation::checked_reference_depth(depth)?;
                collect_image_reference(&reader, draw_element, element, &mut references)?;
            },
            Event::Empty(ref element) => {
                collect_image_reference(&reader, draw_element, element, &mut references)?;
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("image XML stack underflow".to_string()))?;
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 {
        return Err(Error::InvalidFormat(
            "incomplete image XML structure".to_string(),
        ));
    }
    Ok(references)
}

fn collect_image_reference(
    reader: &NsReader<&[u8]>,
    draw_element: bool,
    element: &quick_xml::events::BytesStart<'_>,
    references: &mut Vec<String>,
) -> Result<()> {
    if draw_element
        && element.local_name().as_ref() == b"image"
        && let Some(href) =
            namespaced_attribute(reader, element, XLINK_NAMESPACE, b"href", "draw:image")?
    {
        validation::ensure_reference_capacity(references.len(), "image references")?;
        references.push(href);
    }
    Ok(())
}
