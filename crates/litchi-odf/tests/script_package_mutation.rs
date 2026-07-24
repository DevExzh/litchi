use litchi_odf::{
    Document, DocumentBuilder, OdfDocumentEventListener, OdfEmbeddedScript, OdfScriptBinding,
    OdfScriptEventListener, OdfScriptResourceKind, OdfScriptResourceSpec, Presentation,
    Spreadsheet,
};

fn document() -> Document {
    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("preserved body").unwrap();
    Document::from_bytes(builder.build().unwrap()).unwrap()
}

fn script(language: &str, value: &str) -> OdfEmbeddedScript {
    OdfEmbeddedScript {
        language: language.to_string(),
        content_xml: format!(r#"<payload xmlns="urn:litchi:test">{value}</payload>"#),
    }
}

fn linked(href: &str) -> OdfDocumentEventListener {
    OdfDocumentEventListener::Script(OdfScriptEventListener {
        event_name: "dom:load".to_string(),
        language: "ooo:Basic".to_string(),
        binding: OdfScriptBinding::Linked { href: href.to_string() },
    })
}

fn module(bytes: &[u8]) -> OdfScriptResourceSpec {
    OdfScriptResourceSpec {
        kind: OdfScriptResourceKind::BasicModule,
        preferred_path: None,
        media_type: "text/xml".to_string(),
        bytes: bytes.to_vec(),
    }
}

#[test]
fn generated_document_script_and_listener_crud_is_inert_and_ordered() {
    let mut document = document();
    assert_eq!(document.add_document_script(&script("ooo:Basic", "one")).unwrap(), 0);
    assert_eq!(document.add_document_script(&script("text/python", "two")).unwrap(), 1);
    document.move_document_script(1, 0).unwrap();
    document.replace_document_script(1, &script("ooo:Basic", "replaced")).unwrap();

    let macro_listener = OdfDocumentEventListener::Script(OdfScriptEventListener {
        event_name: "dom:load".to_string(),
        language: "ooo:Basic".to_string(),
        binding: OdfScriptBinding::MacroName("application:Standard.Module1.Main".to_string()),
    });
    document.add_document_event_listener(&macro_listener).unwrap();
    document.add_document_event_listener(&linked("https://example.invalid/inert.js")).unwrap();
    document.move_document_event_listener(1, 0).unwrap();

    let parsed = document.document_scripts().unwrap().unwrap();
    assert_eq!(parsed.scripts[0].language, "text/python");
    assert!(parsed.scripts[1].content_xml.contains("replaced"));
    assert_eq!(parsed.event_listeners.len(), 2);
    assert!(document.text().unwrap().contains("preserved body"));

    document.remove_document_event_listener(1).unwrap();
    document.remove_document_script(0).unwrap();
    assert_eq!(document.document_scripts().unwrap().unwrap().scripts.len(), 1);
}

#[test]
fn resources_allocate_collisions_update_and_preserve_reachability() {
    let mut document = document();
    let first = document.add_script_resource(&module(b"<module xmlns=\"urn:test\">one</module>")).unwrap();
    let second = document.add_script_resource(&module(b"<module xmlns=\"urn:test\">two</module>")).unwrap();
    assert_ne!(first, second);
    assert_eq!(document.find_script_resource(&first).unwrap().unwrap().media_type, "text/xml");

    document.add_document_event_listener(&linked(&first)).unwrap();
    assert!(document.remove_script_resource(&first).is_err());
    assert!(document.find_script_resource(&first).unwrap().is_some());

    document.remove_document_event_listener(0).unwrap();
    let replacement = module(b"<module xmlns=\"urn:test\">updated</module>");
    document.update_script_resource(&first, &replacement).unwrap();
    assert!(String::from_utf8(document.find_script_resource(&first).unwrap().unwrap().bytes).unwrap().contains("updated"));
    document.remove_script_resource(&first).unwrap();
    assert!(document.find_script_resource(&first).unwrap().is_none());

    let opaque = OdfScriptResourceSpec {
        kind: OdfScriptResourceKind::Opaque,
        preferred_path: Some("Scripts/payload.bin".to_string()),
        media_type: "application/octet-stream".to_string(),
        bytes: vec![0, 255, 1, 2],
    };
    let path = document.add_script_resource(&opaque).unwrap();
    assert_eq!(document.find_script_resource(&path).unwrap().unwrap().bytes, opaque.bytes);
}

#[test]
fn malformed_paths_entities_and_executable_uris_roll_back() {
    let mut document = document();
    let traversal = OdfScriptResourceSpec {
        kind: OdfScriptResourceKind::Opaque,
        preferred_path: Some("Scripts/../content.xml".to_string()),
        media_type: "application/octet-stream".to_string(),
        bytes: vec![1],
    };
    assert!(document.add_script_resource(&traversal).is_err());

    let dtd = OdfScriptResourceSpec {
        kind: OdfScriptResourceKind::Dialog,
        preferred_path: Some("Dialogs/bad.xml".to_string()),
        media_type: "text/xml".to_string(),
        bytes: b"<!DOCTYPE x [<!ENTITY e SYSTEM 'file:///etc/passwd'>]><x>&e;</x>".to_vec(),
    };
    assert!(document.add_script_resource(&dtd).is_err());
    assert!(document.add_document_event_listener(&linked("javascript:alert(1)")).is_err());
    let scripts = document.document_scripts().unwrap().unwrap();
    assert!(scripts.scripts.is_empty());
    assert!(scripts.event_listeners.is_empty());
    assert!(document.script_resources().unwrap().is_empty());
}

#[test]
fn ods_and_odp_facades_expose_the_same_inert_api() {
    let _ods: fn(&Spreadsheet) -> litchi_core::Result<Option<litchi_odf::OdfDocumentScripts>> = Spreadsheet::document_scripts;
    let _odp: fn(&Presentation) -> litchi_core::Result<Option<litchi_odf::OdfDocumentScripts>> = Presentation::document_scripts;
}

#[test]
fn bundled_libreoffice_macro_fixture_is_only_read_as_bytes() {
    let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/libreoffice-core/xmlsecurity/qa/unit/signing/data/macro.odt");
    if !fixture.exists() { return; }
    let document = Document::open(fixture).unwrap();
    let resources = document.script_resources().unwrap();
    assert!(resources.iter().any(|resource| resource.path == "Basic/Standard/Module1.xml"));
    assert!(resources.iter().any(|resource| !resource.bytes.is_empty()));
}
