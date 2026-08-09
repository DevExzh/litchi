//! Namespace-aware ODB content validation.

use litchi_core::{Error, Result};
use quick_xml::{
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DATABASE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:database:1.0";
const MAX_CONTENT_BYTES: usize = 256 * 1024 * 1024;
const MAX_DEPTH: usize = 512;
const MAX_EVENTS: usize = 1_000_000;

#[derive(Clone, Copy, PartialEq, Eq)]
enum Element {
    Document,
    Body,
    Database,
    DataSource,
    Other,
}

/// Validate a UTF-8 content part before authoring it into a package.
pub(crate) fn validate(xml: &str) -> Result<()> {
    if xml.len() > MAX_CONTENT_BYTES {
        return Err(invalid("ODB content.xml exceeds the family limit"));
    }

    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Element>::new();
    let mut events = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut database_seen = false;
    let mut data_source_count = 0usize;

    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("ODB content.xml event count overflow"))?;
        if events > MAX_EVENTS {
            return Err(invalid("ODB content.xml exceeds the event limit"));
        }

        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(&format!("invalid ODB content.xml: {error}")))?;
        match event {
            Event::Start(element) => {
                let classified = classify(
                    stack.last().copied(),
                    &namespace,
                    Some(element.local_name().as_ref()),
                );
                if stack.is_empty() {
                    if root_seen || root_closed || classified != Element::Document {
                        return Err(invalid(
                            "ODB content.xml has no office:document-content root",
                        ));
                    }
                    root_seen = true;
                }
                if stack.len() >= MAX_DEPTH {
                    return Err(invalid("ODB content.xml exceeds the nesting limit"));
                }
                record(classified, &mut database_seen, &mut data_source_count)?;
                stack.push(classified);
            },
            Event::Empty(element) => {
                let classified = classify(
                    stack.last().copied(),
                    &namespace,
                    Some(element.local_name().as_ref()),
                );
                if stack.is_empty() {
                    return Err(invalid("ODB content.xml root cannot be empty"));
                }
                if stack.len() >= MAX_DEPTH {
                    return Err(invalid("ODB content.xml exceeds the nesting limit"));
                }
                record(classified, &mut database_seen, &mut data_source_count)?;
            },
            Event::End(_) => {
                if stack.pop().is_none() {
                    return Err(invalid("ODB content.xml has an unmatched closing tag"));
                }
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::DocType(_) => {
                return Err(invalid("DOCTYPE is not permitted in ODB content.xml"));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }

    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("ODB content.xml is incomplete"));
    }
    if !database_seen {
        return Err(invalid("ODB content.xml has no office:database body"));
    }
    if data_source_count != 1 {
        return Err(invalid(
            "ODB office:database must contain exactly one db:data-source",
        ));
    }
    Ok(())
}

fn classify(
    parent: Option<Element>,
    namespace: &ResolveResult<'_>,
    local: Option<&[u8]>,
) -> Element {
    match (parent, namespace, local) {
        (None, ResolveResult::Bound(Namespace(uri)), Some(b"document-content"))
            if *uri == OFFICE_NAMESPACE =>
        {
            Element::Document
        },
        (Some(Element::Document), ResolveResult::Bound(Namespace(uri)), Some(b"body"))
            if *uri == OFFICE_NAMESPACE =>
        {
            Element::Body
        },
        (Some(Element::Body), ResolveResult::Bound(Namespace(uri)), Some(b"database"))
            if *uri == OFFICE_NAMESPACE =>
        {
            Element::Database
        },
        (Some(Element::Database), ResolveResult::Bound(Namespace(uri)), Some(b"data-source"))
            if *uri == DATABASE_NAMESPACE =>
        {
            Element::DataSource
        },
        _ => Element::Other,
    }
}

fn record(element: Element, database_seen: &mut bool, data_source_count: &mut usize) -> Result<()> {
    match element {
        Element::Database => {
            if *database_seen {
                return Err(invalid(
                    "ODB content.xml has multiple office:database bodies",
                ));
            }
            *database_seen = true;
        },
        Element::DataSource => {
            *data_source_count = data_source_count
                .checked_add(1)
                .ok_or_else(|| invalid("ODB data-source count overflow"))?;
        },
        Element::Document | Element::Body | Element::Other => {},
    }
    Ok(())
}

fn invalid(message: &str) -> Error {
    Error::InvalidFormat(message.to_owned())
}

#[cfg(test)]
mod tests {
    use super::validate;

    #[test]
    fn requires_family_body() {
        assert!(validate(
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0"><o:body><o:database><d:data-source/></o:database></o:body></o:document-content>"#,
        )
        .is_ok());
        assert!(validate("<office:database/>").is_err());
    }
}
