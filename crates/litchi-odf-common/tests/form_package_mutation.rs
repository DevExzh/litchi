use litchi_odf_common::{
    constants,
    core::{OwnedPackage, PackageWriter},
    package::edit::{rebuild_package, splice},
};

const LISTENER: &str = r#"<office:event-listeners><script:event-listener script:event-name="dom:click" script:language="ooo:script" script:macro-name="Standard.Module1.Main" xlink:type="simple" xlink:href="vnd.sun.star.script:Standard.Module1.Main"/></office:event-listeners>"#;
const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:text><!--unknown-preserved--><office:forms><form:form form:name="Main">LISTENER<form:text form:name="Text" xml:id="control_id"/></form:form></office:forms><text:p>body</text:p></office:text></office:body></office:document-content>"#;

fn package() -> Vec<u8> {
    let content = CONTENT.replace("LISTENER", LISTENER);
    let mut writer = PackageWriter::new();
    writer.set_mimetype(constants::ODF_TEXT).unwrap();
    writer
        .add_file(constants::ODF_CONTENT, content.as_bytes())
        .unwrap();
    writer
        .add_file_with_media_type(
            "META-INF/documentsignatures.xml",
            b"<signatures/>",
            "text/xml",
        )
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

#[test]
fn neutral_rebuild_preserves_inert_forms_and_drops_stale_signatures() {
    let source = OwnedPackage::from_bytes(package()).unwrap();
    let content = String::from_utf8(source.get_file(constants::ODF_CONTENT).unwrap()).unwrap();
    let control = r#"<form:button form:name="SafeButton" xml:id="safe_button"/>"#;
    let insertion = content.find("</form:form>").unwrap();
    let updated = splice(&content, insertion, insertion, control).unwrap();
    let rebuilt = rebuild_package(
        &source,
        &updated,
        Vec::new(),
        Vec::new(),
        Vec::new(),
        Vec::new(),
    )
    .unwrap();
    let rebuilt = OwnedPackage::from_bytes(rebuilt).unwrap();
    let rebuilt_content =
        String::from_utf8(rebuilt.get_file(constants::ODF_CONTENT).unwrap()).unwrap();
    assert!(rebuilt_content.contains(LISTENER));
    assert!(rebuilt_content.contains(control));
    assert!(rebuilt_content.contains("unknown-preserved"));
    assert!(!rebuilt_content.contains('\n'));
    assert!(!rebuilt.has_file("META-INF/documentsignatures.xml").unwrap());
}

#[test]
fn invalid_form_splice_does_not_change_the_source_snapshot() {
    let source = OwnedPackage::from_bytes(package()).unwrap();
    let content = String::from_utf8(source.get_file(constants::ODF_CONTENT).unwrap()).unwrap();
    let before = source.as_bytes().to_vec();
    assert!(splice(&content, content.len() + 1, content.len() + 1, "").is_err());
    assert_eq!(source.as_bytes(), before);
    assert!(content.contains(LISTENER));
}
