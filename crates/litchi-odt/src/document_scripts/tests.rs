//! Regression tests for the document script model and XML codec.

use super::*;

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
    let scripts = parse_scripts(&xml).unwrap().unwrap();
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

    let serialized = scripts.to_xml().unwrap();
    let reparsed = parse_scripts(&format!(
            r#"<office:document-content xmlns:office="{OFFICE}">{serialized}<office:body/></office:document-content>"#
        ))
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
        assert!(parse_scripts(&wrap(body)).is_err(), "{body}");
    }
}

#[test]
fn rejects_nested_and_duplicate_script_containers() {
    let nested = format!(
        r#"<o:document-content xmlns:o="{OFFICE}"><o:body><o:scripts/></o:body></o:document-content>"#
    );
    assert!(parse_scripts(&nested).is_err());
    let duplicate = format!(
        r#"<o:document-content xmlns:o="{OFFICE}"><o:scripts/><o:scripts/><o:body/></o:document-content>"#
    );
    assert!(parse_scripts(&duplicate).is_err());
}
