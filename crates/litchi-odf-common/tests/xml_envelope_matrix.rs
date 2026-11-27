//! Root-envelope orderings accepted and refused by the public XML readers.

use litchi_odf_common::core::private::scan_xml;
use litchi_odf_common::core::validate_content_document_part;
use litchi_odf_common::{ResolvedReader, XmlStreamLimits, chart, compact_xml};
use quick_xml::events::Event;
use std::io::Cursor;

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const CHART: &str = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";

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
        ("text before root", format!("text{doc}")),
        ("unterminated root", open.to_owned()),
        ("stray end tag", format!("{doc}</stray>")),
        ("DOCTYPE", format!("{DECL}<!DOCTYPE r>{doc}")),
        ("top-level CDATA", format!("<![CDATA[x]]>{doc}")),
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
fn compact_xml_envelope_matrix() {
    assert_envelope_matrix(
        "<r><c/>",
        "</r>",
        "<r/>",
        &[
            "empty input => Error { kind: MalformedXml, offset: 0 }",
            "whitespace only => Error { kind: FormattingWhitespace, offset: 0 }",
            "declaration only => Error { kind: MalformedXml, offset: 38 }",
            "empty-element root => Ok",
            "start and end root => Ok",
            "declaration then root => Ok",
            "declaration, whitespace, root => Error { kind: FormattingWhitespace, offset: 38 }",
            "root then whitespace => Error { kind: FormattingWhitespace, offset: 11 }",
            "root then comment => Ok",
            "root then PI => Ok",
            "root then text => Error { kind: FormattingWhitespace, offset: 11 }",
            "root then second root => Error { kind: MalformedXml, offset: 11 }",
            "root then foreign start => Error { kind: MalformedXml, offset: 11 }",
            "root then foreign empty => Error { kind: MalformedXml, offset: 11 }",
            "empty-element root then second root => Error { kind: MalformedXml, offset: 4 }",
            "empty-element root then text => Error { kind: FormattingWhitespace, offset: 4 }",
            "declaration after root => Ok",
            "two declarations => Ok",
            "comment before declaration => Ok",
            "text before root => Error { kind: FormattingWhitespace, offset: 0 }",
            "unterminated root => Error { kind: MalformedXml, offset: 7 }",
            "stray end tag => Error { kind: MalformedXml, offset: 11 }",
            "DOCTYPE => Error { kind: DocumentType, offset: 38 }",
            "top-level CDATA => Error { kind: MalformedXml, offset: 0 }",
        ],
        |xml| match compact_xml::validate(xml.as_bytes(), compact_xml::Limits::default()) {
            Ok(()) => "Ok".to_owned(),
            Err(error) => format!("{error:?}"),
        },
    );
}

#[test]
fn chart_reader_envelope_matrix() {
    assert_envelope_matrix(
        &format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:c="{CHART}"><o:body><o:chart><c:chart c:class="c:bar"><c:plot-area/></c:chart></o:chart></o:body>"#
        ),
        "</o:document-content>",
        &format!(r#"<o:document-content xmlns:o="{OFFICE}"/>"#),
        &[
            "empty input => Core(InvalidFormat(\"incomplete standalone chart structure\"))",
            "whitespace only => Core(InvalidFormat(\"incomplete standalone chart structure\"))",
            "declaration only => Core(InvalidFormat(\"incomplete standalone chart structure\"))",
            "empty-element root => Core(InvalidFormat(\"chart content root cannot be empty\"))",
            "start and end root => Ok(children=1)",
            "declaration then root => Ok(children=1)",
            "declaration, whitespace, root => Ok(children=1)",
            "root then whitespace => Ok(children=1)",
            "root then comment => Ok(children=1)",
            "root then PI => Ok(children=1)",
            "root then text => Core(InvalidFormat(\"text is not allowed outside the chart content root\"))",
            "root then second root => Core(InvalidFormat(\"chart content must have one office:document-content root\"))",
            "root then foreign start => Core(InvalidFormat(\"chart content must have one office:document-content root\"))",
            "root then foreign empty => Core(InvalidFormat(\"chart content root cannot be empty\"))",
            "empty-element root then second root => Core(InvalidFormat(\"chart content root cannot be empty\"))",
            "empty-element root then text => Core(InvalidFormat(\"chart content root cannot be empty\"))",
            "declaration after root => Ok(children=1)",
            "two declarations => Ok(children=1)",
            "comment before declaration => Ok(children=1)",
            "text before root => Core(InvalidFormat(\"text is not allowed outside the chart content root\"))",
            "unterminated root => Core(InvalidFormat(\"incomplete standalone chart structure\"))",
            "stray end tag => Core(InvalidFormat(\"XML parsing error: ill-formed document: close tag `</stray>` does not match any open tag\"))",
            "DOCTYPE => Ok(children=1)",
            "top-level CDATA => Core(InvalidFormat(\"content is not allowed outside the chart content root\"))",
        ],
        |xml| match chart::read(xml, XmlStreamLimits::default()) {
            Ok(element) => format!("Ok(children={})", element.children().len()),
            Err(error) => format!("{error:?}"),
        },
    );
}

#[test]
fn resolved_reader_document_envelope_matrix() {
    assert_envelope_matrix(
        "<r><c/>",
        "</r>",
        "<r/>",
        &[
            "empty input => Core(InvalidFormat(\"malformed XML at byte 0: XML document has no single complete root element\"))",
            "whitespace only => Core(InvalidFormat(\"malformed XML at byte 2: XML document has no single complete root element\"))",
            "declaration only => Core(InvalidFormat(\"malformed XML at byte 38: XML document has no single complete root element\"))",
            "empty-element root => Ok(events=1)",
            "start and end root => Ok(events=3)",
            "declaration then root => Ok(events=4)",
            "declaration, whitespace, root => Ok(events=5)",
            "root then whitespace => Ok(events=4)",
            "root then comment => Ok(events=4)",
            "root then PI => Ok(events=4)",
            "root then text => Core(InvalidFormat(\"malformed XML at byte 11: non-whitespace text appears outside the XML root\"))",
            "root then second root => Core(InvalidFormat(\"malformed XML at byte 11: content appears after the XML root\"))",
            "root then foreign start => Core(InvalidFormat(\"malformed XML at byte 11: content appears after the XML root\"))",
            "root then foreign empty => Core(InvalidFormat(\"malformed XML at byte 11: content appears after the XML root\"))",
            "empty-element root then second root => Core(InvalidFormat(\"malformed XML at byte 4: content appears after the XML root\"))",
            "empty-element root then text => Core(InvalidFormat(\"malformed XML at byte 4: non-whitespace text appears outside the XML root\"))",
            "declaration after root => Core(InvalidFormat(\"malformed XML at byte 11: XML declaration is not the first document event\"))",
            "two declarations => Core(InvalidFormat(\"malformed XML at byte 38: XML declaration is not the first document event\"))",
            "comment before declaration => Core(InvalidFormat(\"malformed XML at byte 8: XML declaration is not the first document event\"))",
            "text before root => Core(InvalidFormat(\"malformed XML at byte 0: non-whitespace text appears outside the XML root\"))",
            "unterminated root => Core(InvalidFormat(\"malformed XML at byte 7: XML document has no single complete root element\"))",
            "stray end tag => Core(InvalidFormat(\"XML parsing error: ill-formed document: close tag `</stray>` does not match any open tag\"))",
            "DOCTYPE => Core(InvalidFormat(\"malformed XML at byte 38: DTD and DOCTYPE declarations are prohibited\"))",
            "top-level CDATA => Core(InvalidFormat(\"malformed XML at byte 0: CDATA appears outside the XML root\"))",
        ],
        |xml| {
            let mut reader =
                match ResolvedReader::from_xml_document(xml, XmlStreamLimits::default()) {
                    Ok(reader) => reader,
                    Err(error) => return format!("{error:?}"),
                };
            let mut events = 0usize;
            loop {
                match reader.read_resolved_event() {
                    Ok((_, Event::Eof)) => return format!("Ok(events={events})"),
                    Ok(_) => events += 1,
                    Err(error) => return format!("{error:?}"),
                }
            }
        },
    );
}

#[test]
fn stream_scanner_envelope_matrix() {
    assert_envelope_matrix(
        "<r><c/>",
        "</r>",
        "<r/>",
        &[
            "empty input => InvalidFormat(\"malformed XML at byte 0: XML document has no single complete root element\")",
            "whitespace only => InvalidFormat(\"malformed XML at byte 2: XML document has no single complete root element\")",
            "declaration only => InvalidFormat(\"malformed XML at byte 38: XML document has no single complete root element\")",
            "empty-element root => Ok(events=2)",
            "start and end root => Ok(events=4)",
            "declaration then root => Ok(events=5)",
            "declaration, whitespace, root => Ok(events=6)",
            "root then whitespace => Ok(events=5)",
            "root then comment => Ok(events=5)",
            "root then PI => Ok(events=5)",
            "root then text => InvalidFormat(\"malformed XML at byte 11: non-whitespace text appears outside the XML root\")",
            "root then second root => InvalidFormat(\"malformed XML at byte 11: content appears after the XML root\")",
            "root then foreign start => InvalidFormat(\"malformed XML at byte 11: content appears after the XML root\")",
            "root then foreign empty => InvalidFormat(\"malformed XML at byte 11: content appears after the XML root\")",
            "empty-element root then second root => InvalidFormat(\"malformed XML at byte 4: content appears after the XML root\")",
            "empty-element root then text => InvalidFormat(\"malformed XML at byte 4: non-whitespace text appears outside the XML root\")",
            "declaration after root => InvalidFormat(\"malformed XML at byte 11: XML declaration is not the first document event\")",
            "two declarations => InvalidFormat(\"malformed XML at byte 38: XML declaration is not the first document event\")",
            "comment before declaration => InvalidFormat(\"malformed XML at byte 8: XML declaration is not the first document event\")",
            "text before root => InvalidFormat(\"malformed XML at byte 0: non-whitespace text appears outside the XML root\")",
            "unterminated root => InvalidFormat(\"malformed XML at byte 7: XML document has no single complete root element\")",
            "stray end tag => InvalidFormat(\"malformed XML at byte 11: ill-formed document: close tag `</stray>` does not match any open tag\")",
            "DOCTYPE => InvalidFormat(\"malformed XML at byte 38: DTD and DOCTYPE declarations are prohibited\")",
            "top-level CDATA => InvalidFormat(\"malformed XML at byte 0: CDATA appears outside the XML root\")",
        ],
        |xml| match scan_xml(
            Cursor::new(xml.as_bytes()),
            XmlStreamLimits::default(),
            |_event, _tracker| Ok(()),
        ) {
            Ok(report) => format!("Ok(events={})", report.events()),
            Err(error) => format!("{error:?}"),
        },
    );
}

#[test]
fn content_document_envelope_matrix() {
    assert_envelope_matrix(
        &format!(
            r#"<office:document-content xmlns:office="{OFFICE}"><office:body><office:text></office:text></office:body>"#
        ),
        "</office:document-content>",
        &format!(r#"<office:document-content xmlns:office="{OFFICE}"/>"#),
        &[
            "empty input => InvalidFormat(\"ODT content.xml has no complete expected body\")",
            "whitespace only => InvalidFormat(\"ODT content.xml has no complete expected body\")",
            "declaration only => InvalidFormat(\"ODT content.xml has no complete expected body\")",
            "empty-element root => InvalidFormat(\"ODT content.xml has an invalid empty root\")",
            "start and end root => Ok",
            "declaration then root => Ok",
            "declaration, whitespace, root => Ok",
            "root then whitespace => Ok",
            "root then comment => Ok",
            "root then PI => Ok",
            "root then text => InvalidFormat(\"ODT content.xml has unexpected text outside its root\")",
            "root then second root => InvalidFormat(\"ODT content.xml has content after its root\")",
            "root then foreign start => InvalidFormat(\"ODT content.xml has content after its root\")",
            "root then foreign empty => InvalidFormat(\"ODT content.xml has an invalid empty root\")",
            "empty-element root then second root => InvalidFormat(\"ODT content.xml has an invalid empty root\")",
            "empty-element root then text => InvalidFormat(\"ODT content.xml has an invalid empty root\")",
            "declaration after root => InvalidFormat(\"ODT content.xml has an XML declaration outside its prologue\")",
            "two declarations => InvalidFormat(\"ODT content.xml has an XML declaration outside its prologue\")",
            "comment before declaration => InvalidFormat(\"ODT content.xml has an XML declaration outside its prologue\")",
            "text before root => InvalidFormat(\"ODT content.xml has unexpected text outside its root\")",
            "unterminated root => InvalidFormat(\"ODT content.xml has no complete expected body\")",
            "stray end tag => InvalidFormat(\"invalid ODT content.xml: ill-formed document: close tag `</stray>` does not match any open tag\")",
            "DOCTYPE => InvalidFormat(\"ODT content.xml must not contain a doctype\")",
            "top-level CDATA => InvalidFormat(\"ODT content.xml has content outside its root\")",
        ],
        |xml| match validate_content_document_part(xml, "<office:text", "ODT") {
            Ok(()) => "Ok".to_owned(),
            Err(error) => format!("{error:?}"),
        },
    );
}
