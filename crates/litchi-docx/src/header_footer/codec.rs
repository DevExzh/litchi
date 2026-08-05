//! Bounded WordprocessingML and markup-compatibility codec for stories.

use crate::error::{Error, Result};
use litchi_ooxml_common::mce::process_ooxml;
use quick_xml::events::Event;
use quick_xml::reader::Reader;
use std::sync::Arc;

use super::model::{MAX_XML_BYTES, MAX_XML_DEPTH, MAX_XML_NODES, Role};

/// Validate a story and return its typed root role.
pub(crate) fn validate(xml: &[u8]) -> Result<Role> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "header/footer XML exceeds {MAX_XML_BYTES} bytes"
        )));
    }

    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut role = None;
    let mut closed = false;

    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("header/footer XML element counter overflow".into())
                })?;
                if nodes > MAX_XML_NODES {
                    return Err(Error::InvalidFormat(format!(
                        "header/footer XML exceeds {MAX_XML_NODES} elements"
                    )));
                }
                if depth == 0 {
                    if closed {
                        return Err(Error::InvalidFormat(
                            "header/footer XML contains multiple roots".into(),
                        ));
                    }
                    role = Some(root_role(element.name().local_name().as_ref())?);
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("header/footer XML nesting overflow".into())
                })?;
                if depth > MAX_XML_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "header/footer XML nesting exceeds {MAX_XML_DEPTH}"
                    )));
                }
            },
            Event::Empty(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("header/footer XML element counter overflow".into())
                })?;
                if nodes > MAX_XML_NODES {
                    return Err(Error::InvalidFormat(format!(
                        "header/footer XML exceeds {MAX_XML_NODES} elements"
                    )));
                }
                if depth == 0 {
                    if closed {
                        return Err(Error::InvalidFormat(
                            "header/footer XML contains multiple roots".into(),
                        ));
                    }
                    role = Some(root_role(element.name().local_name().as_ref())?);
                    closed = true;
                }
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("header/footer XML has an unexpected end tag".into())
                })?;
                if depth == 0 {
                    closed = true;
                }
            },
            Event::Text(text) if depth == 0 => {
                if !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(
                        "header/footer XML has non-whitespace text outside its root".into(),
                    ));
                }
            },
            Event::CData(data) if depth == 0 => {
                if !data.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(
                        "header/footer XML has CDATA outside its root".into(),
                    ));
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if depth != 0 {
        return Err(Error::InvalidFormat(
            "header/footer XML contains an unterminated element".into(),
        ));
    }
    role.ok_or_else(|| Error::InvalidFormat("header/footer XML has no root".into()))
}

/// Apply the shared bounded MCE visibility projection for semantic traversal.
pub(crate) fn semantic_xml(xml: &Arc<Vec<u8>>) -> Result<Arc<Vec<u8>>> {
    let processed = process_ooxml(xml.as_slice())?;
    let output = match processed {
        std::borrow::Cow::Borrowed(_) => Arc::clone(xml),
        std::borrow::Cow::Owned(value) => Arc::new(value),
    };
    if output.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "processed header/footer XML exceeds {MAX_XML_BYTES} bytes"
        )));
    }
    Ok(output)
}

fn root_role(name: &[u8]) -> Result<Role> {
    match name {
        b"hdr" => Ok(Role::Header),
        b"ftr" => Ok(Role::Footer),
        _ => Err(Error::InvalidFormat(format!(
            "expected w:hdr or w:ftr root, found '{}'",
            String::from_utf8_lossy(name)
        ))),
    }
}
