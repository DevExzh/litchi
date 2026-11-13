use super::generated_xml::{GeneratedXmlEnvelope, GeneratedXmlLimits};
use litchi_core::Error;
use std::io::{Read, Write};

fn limits() -> GeneratedXmlLimits {
    GeneratedXmlLimits::default()
}

fn read_all<F>(
    envelope: GeneratedXmlEnvelope,
    maximum: usize,
    produce: F,
) -> litchi_core::Result<(Vec<u8>, super::generated_xml::GeneratedXmlReport)>
where
    F: FnMut(&mut dyn Write) -> litchi_core::Result<bool>,
{
    let mut reader = envelope.prepare(limits(), maximum, produce)?;
    reader.prefetch_first()?;
    let mut bytes = Vec::new();
    reader.read_to_end(&mut bytes).map_err(Error::Io)?;
    Ok((bytes, reader.report()))
}

#[test]
fn zero_fragments_publish_the_direct_shell_and_exact_base_report() {
    let prefix = b"<?xml version=\"1.0\" encoding=\"UTF-8\"?><root><body>";
    let suffix = b"</body></root>";
    let envelope = GeneratedXmlEnvelope::try_new(prefix, suffix).unwrap();
    let (bytes, report) = read_all(envelope, 256, |_output| Ok(false)).unwrap();
    let mut expected = prefix.to_vec();
    expected.extend_from_slice(suffix);
    assert_eq!(bytes, expected);

    let base = xml_minifier::audit::verify_authored(&expected, limits()).unwrap();
    assert_eq!(report.bytes(), base.bytes());
    assert_eq!(report.attributes(), base.attributes());
    assert_eq!(report.text_bytes(), base.text_bytes());
    assert_eq!(report.events(), base.events());
    assert_eq!(report.max_depth(), base.max_depth());
    assert_eq!(report.fragments(), 0);
}

#[test]
fn exact_depth_limit_does_not_count_a_fake_placeholder() {
    let limits = GeneratedXmlLimits::builder().depth(1).unwrap().build();
    let envelope = GeneratedXmlEnvelope::try_new(b"<root>", b"</root>").unwrap();
    let mut reader = envelope.prepare(limits, 64, |_output| Ok(false)).unwrap();
    reader.prefetch_first().unwrap();
    let mut bytes = Vec::new();
    reader.read_to_end(&mut bytes).unwrap();
    assert_eq!(bytes, b"<root></root>");
    assert_eq!(reader.report().max_depth(), 1);
}

#[test]
fn multiple_fragments_match_direct_audit_accounting() {
    let prefix = b"<root><body>";
    let suffix = b"</body></root>";
    let rows: [&[u8]; 2] = [b"<row>hello&amp;&#13;</row>", b"<row><cell/></row>"];
    let mut index = 0usize;
    let envelope = GeneratedXmlEnvelope::try_new(prefix, suffix).unwrap();
    let (bytes, report) = read_all(envelope, 256, |output| {
        if index == rows.len() {
            return Ok(false);
        }
        output.write_all(rows[index]).map_err(Error::Io)?;
        index += 1;
        Ok(true)
    })
    .unwrap();

    let mut expected = prefix.to_vec();
    expected.extend_from_slice(rows[0]);
    expected.extend_from_slice(rows[1]);
    expected.extend_from_slice(suffix);
    let audited = xml_minifier::audit::verify_authored(&expected, limits()).unwrap();
    assert_eq!(bytes, expected);
    assert_eq!(report.bytes(), audited.bytes());
    assert_eq!(report.attributes(), audited.attributes());
    assert_eq!(report.text_bytes(), audited.text_bytes());
    assert_eq!(report.events(), audited.events());
    assert_eq!(report.max_depth(), audited.max_depth());
    assert_eq!(report.fragments(), 2);
}

#[test]
fn predefined_and_numeric_references_inside_fragments_are_allowed() {
    let envelope = GeneratedXmlEnvelope::try_new(b"<root>", b"</root>").unwrap();
    let mut done = false;
    let (_bytes, report) = read_all(envelope, 128, |output| {
        if done {
            return Ok(false);
        }
        done = true;
        output
            .write_all(b"<row>&amp;&lt;&gt;&quot;&apos;&#13;&#x0D;</row>")
            .map_err(Error::Io)?;
        Ok(true)
    })
    .unwrap();
    assert_eq!(report.fragments(), 1);
}

#[test]
fn uppercase_hex_reference_is_rejected() {
    let envelope = GeneratedXmlEnvelope::try_new(b"<root>", b"</root>").unwrap();
    let mut reader = envelope
        .prepare(limits(), 128, |output| {
            output.write_all(b"<row>&#X0D;</row>").map_err(Error::Io)?;
            Ok(true)
        })
        .unwrap();
    assert!(reader.prefetch_first().is_err());
}

#[test]
fn outside_reference_is_rejected() {
    let envelope = GeneratedXmlEnvelope::try_new(b"<root>", b"</root>").unwrap();
    let mut reader = envelope
        .prepare(limits(), 128, |output| {
            output.write_all(b"&amp;").map_err(Error::Io)?;
            Ok(true)
        })
        .unwrap();
    assert!(reader.prefetch_first().is_err());
}

#[test]
fn producer_result_contract_and_ignored_overflow_are_latched() {
    let envelope = GeneratedXmlEnvelope::try_new(b"<root>", b"</root>").unwrap();
    let mut empty_true = envelope.prepare(limits(), 128, |_output| Ok(true)).unwrap();
    assert!(empty_true.prefetch_first().is_err());

    let envelope = GeneratedXmlEnvelope::try_new(b"<root>", b"</root>").unwrap();
    let mut nonempty_false = envelope
        .prepare(limits(), 128, |output| {
            output.write_all(b"<row/>").map_err(Error::Io)?;
            Ok(false)
        })
        .unwrap();
    assert!(nonempty_false.prefetch_first().is_err());

    let envelope = GeneratedXmlEnvelope::try_new(b"<root>", b"</root>").unwrap();
    let mut ignored_overflow = envelope
        .prepare(limits(), 4, |output| {
            let _ = output.write_all(b"<row/>");
            Ok(true)
        })
        .unwrap();
    assert!(ignored_overflow.prefetch_first().is_err());
}

#[test]
fn failed_first_prefetch_cannot_publish_the_shell_if_ignored() {
    let envelope = GeneratedXmlEnvelope::try_new(b"<root>", b"</root>").unwrap();
    let mut reader = envelope.prepare(limits(), 128, |_output| Ok(true)).unwrap();
    assert!(reader.prefetch_first().is_err());
    let mut bytes = Vec::new();
    assert!(reader.read_to_end(&mut bytes).is_err());
    assert!(bytes.is_empty());
}

fn rejects_second_fragment_at_aggregate_limit(
    xml_limits: GeneratedXmlLimits,
    fragment: &'static [u8],
) {
    let envelope = GeneratedXmlEnvelope::try_new(b"<root>", b"</root>").unwrap();
    let mut calls = 0usize;
    let mut reader = envelope
        .prepare(xml_limits, fragment.len(), |output| {
            if calls == 2 {
                return Ok(false);
            }
            output.write_all(fragment).map_err(Error::Io)?;
            calls += 1;
            Ok(true)
        })
        .unwrap();
    reader.prefetch_first().unwrap();
    let mut bytes = Vec::new();
    assert!(reader.read_to_end(&mut bytes).is_err());
    assert_eq!(calls, 2);
}

#[test]
fn aggregate_byte_limit_rejects_individually_valid_fragments() {
    let shell = xml_minifier::audit::verify_authored(b"<root></root>", limits()).unwrap();
    let fragment = b"<row>x</row>";
    let one = xml_minifier::audit::verify_authored(fragment, limits()).unwrap();
    let xml_limits = GeneratedXmlLimits::builder()
        .bytes(shell.bytes() + one.bytes())
        .unwrap()
        .build();
    rejects_second_fragment_at_aggregate_limit(xml_limits, fragment);
}

#[test]
fn aggregate_attribute_limit_rejects_individually_valid_fragments() {
    let shell = xml_minifier::audit::verify_authored(b"<root></root>", limits()).unwrap();
    let fragment = b"<row kind=\"x\"/>";
    let one = xml_minifier::audit::verify_authored(fragment, limits()).unwrap();
    let xml_limits = GeneratedXmlLimits::builder()
        .attributes(shell.attributes() + one.attributes())
        .unwrap()
        .build();
    rejects_second_fragment_at_aggregate_limit(xml_limits, fragment);
}

#[test]
fn aggregate_text_limit_rejects_individually_valid_fragments() {
    let shell = xml_minifier::audit::verify_authored(b"<root></root>", limits()).unwrap();
    let fragment = b"<row>x</row>";
    let one = xml_minifier::audit::verify_authored(fragment, limits()).unwrap();
    let xml_limits = GeneratedXmlLimits::builder()
        .text_bytes(shell.text_bytes() + one.text_bytes())
        .unwrap()
        .build();
    rejects_second_fragment_at_aggregate_limit(xml_limits, fragment);
}

#[test]
fn aggregate_event_limit_rejects_individually_valid_fragments() {
    let shell = xml_minifier::audit::verify_authored(b"<root></root>", limits()).unwrap();
    let fragment = b"<row/>";
    let one = xml_minifier::audit::verify_authored(fragment, limits()).unwrap();
    let first_composed_events = shell
        .events()
        .checked_add(one.events().checked_sub(1).unwrap())
        .unwrap();
    let xml_limits = GeneratedXmlLimits::builder()
        .events(first_composed_events)
        .unwrap()
        .build();
    rejects_second_fragment_at_aggregate_limit(xml_limits, fragment);
}

#[test]
fn envelope_shape_requires_direct_boundary_and_rejects_inherited_space() {
    assert!(GeneratedXmlEnvelope::try_new(b"<root><body>", b"</other></root>").is_err());
    assert!(GeneratedXmlEnvelope::try_new(b"<root xml:space=\"preserve\">", b"</root>").is_err());
    assert!(GeneratedXmlEnvelope::try_new(b"<root>text", b"</root>").is_err());
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
fn generated_fragment_envelope_matrix() {
    assert_envelope_matrix(
        "<r><c/>",
        "</r>",
        "<r/>",
        &[
            "empty input => InvalidFormat(\"generated XML: producer returned a non-EOF result without a fragment\")",
            "whitespace only => InvalidFormat(\"generated XML: XML audit failed: noncompact XML FormattingWhitespace at byte 0\")",
            "declaration only => InvalidFormat(\"generated XML: XML audit failed: malformed XML at byte 38: XML must contain exactly one document element\")",
            "empty-element root => Ok(bytes=17, fragments=1)",
            "start and end root => Ok(bytes=24, fragments=1)",
            "declaration then root => InvalidFormat(\"generated XML: fragment contains a declaration, doctype, comment, or processing instruction\")",
            "declaration, whitespace, root => InvalidFormat(\"generated XML: XML audit failed: noncompact XML FormattingWhitespace at byte 38\")",
            "root then whitespace => InvalidFormat(\"generated XML: XML audit failed: noncompact XML FormattingWhitespace at byte 11\")",
            "root then comment => InvalidFormat(\"generated XML: fragment contains a declaration, doctype, comment, or processing instruction\")",
            "root then PI => InvalidFormat(\"generated XML: fragment contains a declaration, doctype, comment, or processing instruction\")",
            "root then text => InvalidFormat(\"generated XML: XML audit failed: malformed XML at byte 11: character data outside the document element\")",
            "root then second root => InvalidFormat(\"generated XML: XML audit failed: malformed XML at byte 11: multiple document elements\")",
            "root then foreign start => InvalidFormat(\"generated XML: XML audit failed: malformed XML at byte 11: multiple document elements\")",
            "root then foreign empty => InvalidFormat(\"generated XML: XML audit failed: malformed XML at byte 11: multiple document elements\")",
            "empty-element root then second root => InvalidFormat(\"generated XML: XML audit failed: malformed XML at byte 4: multiple document elements\")",
            "empty-element root then text => InvalidFormat(\"generated XML: XML audit failed: malformed XML at byte 4: character data outside the document element\")",
            "declaration after root => InvalidFormat(\"generated XML: fragment contains a declaration, doctype, comment, or processing instruction\")",
            "two declarations => InvalidFormat(\"generated XML: fragment contains a declaration, doctype, comment, or processing instruction\")",
            "comment before declaration => InvalidFormat(\"generated XML: fragment contains a declaration, doctype, comment, or processing instruction\")",
            "text before root => InvalidFormat(\"generated XML: XML audit failed: malformed XML at byte 0: character data outside the document element\")",
            "unterminated root => InvalidFormat(\"generated XML: XML audit failed: malformed XML at byte 7: unclosed document element\")",
            "stray end tag => InvalidFormat(\"generated XML: XML audit failed: malformed XML at byte 19: ill-formed document: close tag `</stray>` does not match any open tag\")",
            "DOCTYPE => InvalidFormat(\"generated XML: XML audit failed: DTD and DOCTYPE are not allowed at byte 38\")",
            "top-level CDATA => InvalidFormat(\"generated XML: XML audit failed: malformed XML at byte 0: CDATA outside the document element\")",
        ],
        |xml| {
            let envelope = match GeneratedXmlEnvelope::try_new(b"<root>", b"</root>") {
                Ok(envelope) => envelope,
                Err(error) => return format!("{error:?}"),
            };
            let mut pending = Some(xml.as_bytes());
            match read_all(envelope, 4096, |output| match pending.take() {
                Some(fragment) => {
                    output.write_all(fragment).map_err(Error::Io)?;
                    Ok(true)
                },
                None => Ok(false),
            }) {
                Ok((bytes, report)) => {
                    format!(
                        "Ok(bytes={}, fragments={})",
                        bytes.len(),
                        report.fragments()
                    )
                },
                Err(error) => format!("{error:?}"),
            }
        },
    );
}
