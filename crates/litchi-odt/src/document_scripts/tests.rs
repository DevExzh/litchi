//! Regression tests for the document script model and XML codec.

use super::*;
use litchi_odf_common::XmlStreamLimits;

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const SCRIPT: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const PRESENTATION: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";

#[test]
fn parses_and_round_trips_inert_document_scripts() {
    // ODF 1.2/1.3 office:scripts grammar; LibreOffice xmlscripti.cxx consumes
    // the same office:script and office:event-listeners sequence.
    let xml = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:s="{SCRIPT}" xmlns:p="{PRESENTATION}" xmlns:x="http://www.w3.org/1999/xlink" xmlns:ooo="http://openoffice.org/2004/office"><o:scripts><o:script s:language="ooo:Basic"><ooo:libraries><ooo:library name="A&amp;B"/></ooo:libraries></o:script><o:script s:language="Python"/><o:event-listeners><s:event-listener s:event-name="dom:load" s:language="ooo:script" s:macro-name="Standard.Module.Main"/><s:event-listener s:event-name="dom:unload" s:language="javascript" x:type="simple" x:href="Scripts/close.js" x:actuate="onRequest"/><p:event-listener s:event-name="dom:click" p:action="next-page"/></o:event-listeners></o:scripts><o:body/></o:document-content>"#
    );
    let scripts = parse_scripts(&xml, &XmlStreamLimits::owner_default_value())
        .unwrap()
        .unwrap();
    assert_eq!(scripts.scripts.len(), 2);
    assert_eq!(scripts.scripts[0].language, "ooo:Basic");
    assert!(scripts.scripts[0].content_xml.contains("ooo:library"));
    assert_eq!(scripts.event_listeners.len(), 3);
    assert!(matches!(
        &scripts.event_listeners[1],
        EventListener::Script(ScriptEventListener {
            binding: ScriptBinding::Linked { href },
            ..
        }) if href == "Scripts/close.js"
    ));

    let serialized = scripts
        .to_xml(&XmlStreamLimits::owner_default_value())
        .unwrap();
    let reparsed = parse_scripts(&format!(
            r#"<office:document-content xmlns:office="{OFFICE}">{serialized}<office:body/></office:document-content>"#
        ), &XmlStreamLimits::owner_default_value())
        .unwrap()
        .unwrap();
    assert_eq!(reparsed, scripts);
}

#[test]
fn rejects_active_or_malformed_script_metadata() {
    let wrap = |body: &str| {
        format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:s="{SCRIPT}" xmlns:x="http://www.w3.org/1999/xlink"><o:scripts>{body}</o:scripts><o:body/></o:document-content>"#
        )
    };
    for body in [
        r"<o:script/>",
        r#"<o:event-listeners/><o:script s:language="Python"/>"#,
        r#"<o:event-listeners><s:event-listener s:event-name="load" s:language="x"/></o:event-listeners>"#,
        r#"<o:event-listeners><s:event-listener s:event-name="load" s:language="x" s:macro-name="M" x:type="simple" x:href="S"/></o:event-listeners>"#,
        r#"<o:event-listeners><s:event-listener s:event-name="load" s:language="x" x:type="extended" x:href="S"/></o:event-listeners>"#,
        r#"<o:script s:language="Python"><!DOCTYPE x></o:script>"#,
    ] {
        assert!(
            parse_scripts(&wrap(body), &XmlStreamLimits::owner_default_value()).is_err(),
            "{body}"
        );
    }
}

#[test]
fn rejects_nested_and_duplicate_script_containers() {
    let nested = format!(
        r#"<o:document-content xmlns:o="{OFFICE}"><o:body><o:scripts/></o:body></o:document-content>"#
    );
    assert!(parse_scripts(&nested, &XmlStreamLimits::owner_default_value()).is_err());
    let duplicate = format!(
        r#"<o:document-content xmlns:o="{OFFICE}"><o:scripts/><o:scripts/><o:body/></o:document-content>"#
    );
    assert!(parse_scripts(&duplicate, &XmlStreamLimits::owner_default_value()).is_err());
}

#[test]
fn script_writer_honors_exact_and_tiny_output_budgets() {
    let nested = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:s="{SCRIPT}"><o:scripts><o:script s:language="Python"><x:node xmlns:x="urn:test"/></o:script></o:scripts><o:body/></o:document-content>"#
    );
    let shallow =
        XmlStreamLimits::new(nested.len() as u64 + 1, 2, 100_000, 100_000, 64 * 1024).unwrap();
    assert!(parse_scripts(&nested, &shallow).is_err());

    let scripts = Scripts::default();
    let expected = scripts
        .to_xml(&XmlStreamLimits::owner_default_value())
        .unwrap();
    let exact =
        XmlStreamLimits::new(expected.len() as u64, 256, 100_000, 100_000, 64 * 1024).unwrap();
    assert_eq!(scripts.to_xml(&exact).unwrap(), expected);
    let tiny = XmlStreamLimits::new(1, 256, 100_000, 100_000, 64 * 1024).unwrap();
    assert!(scripts.to_xml(&tiny).is_err());
}

#[test]
fn authored_script_fragments_obey_aggregate_stream_limits() {
    let scripts = Scripts {
        scripts: vec![EmbeddedScript {
            language: "Python".to_string(),
            content_xml: "<a x=\"1\"/><b y=\"2\"/>".to_string(),
        }],
        ..Scripts::default()
    };
    let limits = XmlStreamLimits::new(64 * 1024, 256, 100, 1, 64 * 1024).unwrap();
    assert!(scripts.to_xml(&limits).is_err());

    let event_limited = XmlStreamLimits::new(64 * 1024, 256, 3, 100, 64 * 1024).unwrap();
    assert!(scripts.to_xml(&event_limited).is_err());
}

#[test]
fn authored_fragments_share_the_emitted_stream_budget() {
    let scripts = Scripts {
        scripts: vec![
            EmbeddedScript {
                language: "Python".to_string(),
                content_xml: r#"<first x="1"/>"#.to_string(),
            },
            EmbeddedScript {
                language: "Python".to_string(),
                content_xml: r#"<second y="2"/>"#.to_string(),
            },
        ],
        ..Scripts::default()
    };

    let aggregate = XmlStreamLimits::new(64 * 1024, 256, 8, 7, 64 * 1024).unwrap();
    assert!(scripts.to_xml(&aggregate).is_err());

    let expected = scripts
        .to_xml(&XmlStreamLimits::owner_default_value())
        .unwrap();
    let exact = XmlStreamLimits::new(expected.len() as u64, 256, 9, 8, 64 * 1024).unwrap();
    assert_eq!(scripts.to_xml(&exact).unwrap(), expected);
}

#[test]
fn emitted_script_envelopes_use_actual_namespace_limits() {
    let scripts = Scripts::default();
    let expected = scripts
        .to_xml(&XmlStreamLimits::owner_default_value())
        .unwrap();
    let exact = XmlStreamLimits::new(expected.len() as u64, 256, 100_000, 100_000, 64 * 1024)
        .unwrap()
        .namespace_limits(4, 4, 4, 4096, 16, 128)
        .unwrap();
    assert_eq!(scripts.to_xml(&exact).unwrap(), expected);

    for limits in [
        XmlStreamLimits::owner_default_value()
            .namespace_limits(1024, 3, 256 * 1024, 16 * 1024 * 1024, 64 * 1024, 64 * 1024)
            .unwrap(),
        XmlStreamLimits::owner_default_value()
            .namespace_limits(1024, 256, 3, 16 * 1024 * 1024, 64 * 1024, 64 * 1024)
            .unwrap(),
        XmlStreamLimits::owner_default_value()
            .namespace_limits(1024, 256, 256 * 1024, 1, 64 * 1024, 64 * 1024)
            .unwrap(),
        XmlStreamLimits::owner_default_value()
            .namespace_limits(1024, 256, 256 * 1024, 16 * 1024 * 1024, 3, 64 * 1024)
            .unwrap(),
        XmlStreamLimits::owner_default_value()
            .namespace_limits(1024, 256, 256 * 1024, 16 * 1024 * 1024, 64 * 1024, 8)
            .unwrap(),
    ] {
        assert!(scripts.to_xml(&limits).is_err());
    }
}

#[test]
fn emitted_script_fragments_use_actual_depth_and_namespace_scope() {
    let scripts = Scripts {
        scripts: vec![EmbeddedScript {
            language: "Python".to_string(),
            content_xml: "<x><y/></x>".to_string(),
        }],
        ..Scripts::default()
    };
    let expected = scripts
        .to_xml(&XmlStreamLimits::owner_default_value())
        .unwrap();
    let exact = XmlStreamLimits::new(expected.len() as u64, 4, 100_000, 100_000, 64 * 1024)
        .unwrap()
        .namespace_limits(4, 256, 256 * 1024, 16 * 1024 * 1024, 64 * 1024, 64 * 1024)
        .unwrap();
    assert_eq!(scripts.to_xml(&exact).unwrap(), expected);

    let depth_under = XmlStreamLimits::new(expected.len() as u64, 3, 100_000, 100_000, 64 * 1024)
        .unwrap()
        .namespace_limits(4, 256, 256 * 1024, 16 * 1024 * 1024, 64 * 1024, 64 * 1024)
        .unwrap();
    assert!(scripts.to_xml(&depth_under).is_err());

    let scope_under = XmlStreamLimits::new(expected.len() as u64, 4, 100_000, 100_000, 64 * 1024)
        .unwrap()
        .namespace_limits(3, 256, 256 * 1024, 16 * 1024 * 1024, 64 * 1024, 64 * 1024)
        .unwrap();
    assert!(scripts.to_xml(&scope_under).is_err());
}
