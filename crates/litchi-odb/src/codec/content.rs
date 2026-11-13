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
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><o:body><o:database><d:data-source><d:connection-data><d:connection-resource xlink:href="" xlink:type="simple"/></d:connection-data></d:data-source></o:database></o:body></o:document-content>"#,
        )
        .is_ok());
        assert!(validate("<office:database/>").is_err());
    }

    /// Runs `run` over the XML envelope orderings built around one root element
    /// and compares each `label => outcome` line with the frozen expectation.
    fn assert_envelope_matrix(
        open: &str,
        close: &str,
        empty: &str,
        expected: &[&str],
        run: impl Fn(&str) -> String,
    ) {
        const DECL: &str = r#"<?xml version="1.0" encoding="UTF-8"?>"#;
        let doc = format!("{open}{close}");
        let rows = [
            ("empty input", String::new()),
            ("whitespace only", " \n".to_owned()),
            ("declaration only", DECL.to_owned()),
            ("empty-element root", empty.to_owned()),
            ("start and end root", doc.clone()),
            ("declaration then root", format!("{DECL}{doc}")),
            ("declaration, whitespace, root", format!("{DECL}\n{doc}")),
            ("root then whitespace", format!("{doc}\n ")),
            ("root then comment", format!("{doc}<!--c-->")),
            ("root then PI", format!("{doc}<?pi x?>")),
            ("root then text", format!("{doc}text")),
            ("root then second root", format!("{doc}{doc}")),
            (
                "root then foreign start",
                format!("{doc}<foreign></foreign>"),
            ),
            ("root then foreign empty", format!("{doc}<foreign/>")),
            (
                "empty-element root then second root",
                format!("{empty}{doc}"),
            ),
            ("empty-element root then text", format!("{empty}text")),
            ("declaration after root", format!("{doc}{DECL}")),
            ("two declarations", format!("{DECL}{DECL}{doc}")),
            ("comment before declaration", format!("<!--c-->{DECL}{doc}")),
            ("comment before root", format!("<!--c-->{doc}")),
            ("PI before root", format!("<?pi x?>{doc}")),
            ("text before root", format!("text{doc}")),
            ("whitespace before declaration", format!(" {DECL}{doc}")),
            ("unterminated root", open.to_owned()),
            ("stray end tag", format!("{doc}</stray>")),
            ("DOCTYPE", format!("{DECL}<!DOCTYPE r>{doc}")),
            ("DOCTYPE after root", format!("{doc}<!DOCTYPE r>")),
            ("top-level CDATA", format!("<![CDATA[x]]>{doc}")),
            ("CDATA after root", format!("{doc}<![CDATA[x]]>")),
            ("entity reference before root", format!("&amp;{doc}")),
            ("entity reference after root", format!("{doc}&amp;")),
        ];
        let actual: Vec<String> = rows
            .iter()
            .map(|(label, xml)| format!("{label} => {}", run(xml)))
            .collect();
        let listing: Vec<String> = actual.iter().map(|line| format!("{line:?},")).collect();
        assert!(
            actual
                .iter()
                .map(String::as_str)
                .eq(expected.iter().copied()),
            "actual envelope matrix:\n{}",
            listing.join("\n")
        );
    }

    #[test]
    fn content_validation_envelope_matrix() {
        let root = "o:document-content xmlns:o=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:d=\"urn:oasis:names:tc:opendocument:xmlns:database:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\"";
        assert_envelope_matrix(
            &format!(
                "<{root}><o:body><o:database><d:data-source><d:connection-data><d:connection-resource xlink:href=\"\" xlink:type=\"simple\"/></d:connection-data></d:data-source></o:database></o:body>"
            ),
            "</o:document-content>",
            &format!("<{root}/>"),
            &[
                "empty input => InvalidFormat(\"ODB content.xml is incomplete\")",
                "whitespace only => InvalidFormat(\"ODB content.xml is incomplete\")",
                "declaration only => InvalidFormat(\"ODB content.xml is incomplete\")",
                "empty-element root => InvalidFormat(\"ODB content.xml root cannot be empty\")",
                "start and end root => Ok(())",
                "declaration then root => Ok(())",
                "declaration, whitespace, root => Ok(())",
                "root then whitespace => Ok(())",
                "root then comment => Ok(())",
                "root then PI => Ok(())",
                "root then text => Ok(())",
                "root then second root => InvalidFormat(\"ODB content.xml has no office:document-content root\")",
                "root then foreign start => InvalidFormat(\"ODB content.xml has no office:document-content root\")",
                "root then foreign empty => InvalidFormat(\"ODB content.xml root cannot be empty\")",
                "empty-element root then second root => InvalidFormat(\"ODB content.xml root cannot be empty\")",
                "empty-element root then text => InvalidFormat(\"ODB content.xml root cannot be empty\")",
                "declaration after root => Ok(())",
                "two declarations => Ok(())",
                "comment before declaration => Ok(())",
                "comment before root => Ok(())",
                "PI before root => Ok(())",
                "text before root => Ok(())",
                "whitespace before declaration => Ok(())",
                "unterminated root => InvalidFormat(\"ODB content.xml is incomplete\")",
                "stray end tag => InvalidFormat(\"invalid ODB content.xml: ill-formed document: close tag `</stray>` does not match any open tag\")",
                "DOCTYPE => InvalidFormat(\"DOCTYPE is not permitted in ODB content.xml\")",
                "DOCTYPE after root => InvalidFormat(\"DOCTYPE is not permitted in ODB content.xml\")",
                "top-level CDATA => Ok(())",
                "CDATA after root => Ok(())",
                "entity reference before root => Ok(())",
                "entity reference after root => Ok(())",
            ],
            |xml| match validate(xml) {
                Ok(value) => format!("Ok({value:?})"),
                Err(error) => format!("{error:?}"),
            },
        );
    }
}
