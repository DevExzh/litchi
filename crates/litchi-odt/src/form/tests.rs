//! Regression tests for the layered form codec.

use super::*;

const PREFIX: &str = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:f="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:x="http://www.w3.org/1999/xlink"><o:body><o:text><o:forms><f:form f:name="Main"><f:button xml:id="submit" f:name="Submit">"#;
const SUFFIX: &str = "</f:button></f:form></o:forms></o:text></o:body></o:document-content>";

#[test]
fn retains_inert_form_event_listener_metadata_and_owner() {
    let xml = format!(
        r#"{PREFIX}<o:event-listeners><s:event-listener s:event-name="dom:click" s:language="ooo:script" s:macro-name="vnd.sun.star.script:Module.Run" x:type="simple" x:href="Scripts/Main" x:actuate="onRequest"/></o:event-listeners>{SUFFIX}"#
    );
    let forms = parse_form_parts(&[(&xml, Part::Content)]).unwrap();
    assert!(forms.has_event_listeners);
    assert_eq!(forms.event_listeners.len(), 1);
    let listener = &forms.event_listeners[0];
    assert_eq!(listener.event_name.as_deref(), Some("dom:click"));
    assert_eq!(listener.language.as_deref(), Some("ooo:script"));
    assert_eq!(listener.href.as_deref(), Some("Scripts/Main"));
    assert_eq!(listener.actuate, Some(Actuate::OnRequest));
    assert!(listener.simple_link);
    assert!(matches!(
        &listener.target,
        Target::Control { name: Some(name), .. } if name == "Submit"
    ));
}

#[test]
fn accepts_expanded_empty_listener_and_rejects_active_or_malformed_content() {
    let expanded = format!(
        r#"{PREFIX}<o:event-listeners><s:event-listener s:event-name="change" s:language="none"></s:event-listener></o:event-listeners>{SUFFIX}"#
    );
    assert_eq!(
        parse_form_parts(&[(&expanded, Part::Content)])
            .unwrap()
            .event_listeners
            .len(),
        1
    );
    for body in [
        r#"<o:event-listeners><s:event-listener s:event-name="" s:language="none"/></o:event-listeners>"#,
        r#"<o:event-listeners><s:event-listener s:event-name="x" s:language="none" x:type="extended"/></o:event-listeners>"#,
        r#"<o:event-listeners><s:event-listener s:event-name="x" s:language="none"><f:property/></s:event-listener></o:event-listeners>"#,
        r#"<o:event-listeners>active text</o:event-listeners>"#,
    ] {
        let xml = format!("{PREFIX}{body}{SUFFIX}");
        assert!(
            parse_form_parts(&[(&xml, Part::Content)]).is_err(),
            "accepted {body}"
        );
    }
}
