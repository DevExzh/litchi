use super::*;

#[test]
fn body_edits_preserve_settings_part_byte_for_byte() {
    let mut package = Package::new().unwrap();
    let settings_uri = PackURI::new("/word/settings.xml").unwrap();
    let before = package.opc.get_part(&settings_uri).unwrap().blob().to_vec();

    package
        .document_mut()
        .unwrap()
        .add_paragraph_with_text("body-only edit");
    package.to_stream(Cursor::new(Vec::new())).unwrap();

    assert_eq!(package.opc.get_part(&settings_uri).unwrap().blob(), before);
}

#[test]
fn new_package_uses_leaf_owned_web_settings_bytes() {
    use crate::web::{Conformance, Settings, write};

    let package = Package::new().unwrap();
    let uri = PackURI::new("/word/webSettings.xml").unwrap();
    let expected = write(&Settings::default(), Conformance::Transitional).unwrap();

    assert_eq!(package.opc.get_part(&uri).unwrap().blob(), expected);
}

#[test]
fn body_edits_preserve_web_settings_part_byte_for_byte() {
    let mut package = Package::new().unwrap();
    let web_settings_uri = PackURI::new("/word/webSettings.xml").unwrap();
    let before = package
        .opc
        .get_part(&web_settings_uri)
        .unwrap()
        .blob()
        .to_vec();

    package
        .document_mut()
        .unwrap()
        .add_paragraph_with_text("body-only edit");
    package.to_stream(Cursor::new(Vec::new())).unwrap();

    assert_eq!(
        package.opc.get_part(&web_settings_uri).unwrap().blob(),
        before
    );
}

#[test]
fn edits_web_settings_without_rewriting_document_content() {
    use crate::web::{Conformance, Screen};

    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    let (mut settings, conformance) = package.web().unwrap().unwrap();
    settings
        .set_allow_png(false)
        .set_optimize_for_browser(true)
        .set_target_screen_size(Screen::Pixels1600x1200);
    assert_eq!(conformance, Conformance::Transitional);
    assert!(package.put_web(settings, conformance).unwrap());
    package.save(file.path()).unwrap();

    let reopened = Package::open(file.path()).unwrap();
    let (settings, conformance) = reopened.document().unwrap().web().unwrap().unwrap();
    assert_eq!(conformance, Conformance::Transitional);
    assert_eq!(settings.allow_png(), Some(false));
    assert_eq!(settings.optimize_for_browser(), Some(true));
    assert_eq!(settings.target_screen_size(), Some(Screen::Pixels1600x1200));
}

#[test]
fn web_settings_updates_preserve_frame_relationship_ids() {
    use crate::web::{Frameset, Layout};

    const FRAME_RELATIONSHIP: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/frame";
    let mut package = Package::new().unwrap();
    let web_settings_uri = PackURI::new("/word/webSettings.xml").unwrap();
    let relationship_id = package
        .opc
        .get_part_mut(&web_settings_uri)
        .unwrap()
        .relate_to("frame1.html", FRAME_RELATIONSHIP);

    let mut frameset = Frameset::default();
    frameset.set_layout(Layout::Rows);
    frameset
        .add_frame()
        .unwrap()
        .set_name("main")
        .unwrap()
        .set_rel(&relationship_id)
        .unwrap();
    let (mut settings, conformance) = package.web().unwrap().unwrap();
    settings.set_frameset(frameset);
    assert!(package.put_web(settings, conformance).unwrap());
    package.to_stream(Cursor::new(Vec::new())).unwrap();

    let part = package.opc.get_part(&web_settings_uri).unwrap();
    let relationship = part.rels().get(&relationship_id).unwrap();
    assert_eq!(relationship.reltype(), FRAME_RELATIONSHIP);
    assert_eq!(relationship.target_ref(), "frame1.html");
    assert!(package.document().unwrap().web().unwrap().is_some());
}

#[test]
fn creates_a_web_settings_relationship_when_missing() {
    use crate::web::{Conformance, Settings};
    use litchi_opc::constants::relationship_type as rt;

    let mut package = Package::new().unwrap();
    let doc_uri = PackURI::new("/word/document.xml").unwrap();
    assert!(package.remove_web().unwrap());
    let mut settings = Settings::default();
    settings.set_encoding("utf-8").unwrap();
    assert!(
        package
            .put_web(settings, Conformance::Transitional)
            .unwrap()
    );
    package.to_stream(Cursor::new(Vec::new())).unwrap();

    let relationship = package
        .opc
        .get_part(&doc_uri)
        .unwrap()
        .rels()
        .part_with_reltype(rt::WEB_SETTINGS)
        .unwrap();
    assert_eq!(relationship.target_ref(), "webSettings.xml");
    assert_eq!(
        package
            .document()
            .unwrap()
            .web()
            .unwrap()
            .unwrap()
            .0
            .encoding(),
        Some("utf-8")
    );
}

#[test]
fn reads_and_updates_strict_web_settings_relationships() {
    use crate::web::{Conformance, Settings};
    use litchi_opc::constants::relationship_type as rt;

    let mut package = Package::new().unwrap();
    assert!(package.remove_web().unwrap());
    let (relationship_id, target_ref) = {
        let relationship = package
            .opc
            .rels()
            .part_with_reltype(rt::OFFICE_DOCUMENT)
            .unwrap();
        (
            relationship.r_id().to_owned(),
            relationship.target_ref().to_owned(),
        )
    };
    package.opc.rels_mut().remove(&relationship_id);
    package.opc.rels_mut().add_relationship(
        rt::STRICT_OFFICE_DOCUMENT.to_owned(),
        target_ref,
        relationship_id.clone(),
        false,
    );

    let mut settings = Settings::default();
    settings.set_save_smart_tags_as_xml(true);
    assert!(package.put_web(settings, Conformance::Strict).unwrap());
    package.to_stream(Cursor::new(Vec::new())).unwrap();

    let (_, conformance) = package.document().unwrap().web().unwrap().unwrap();
    assert_eq!(conformance, Conformance::Strict);
    assert_eq!(
        package
            .document()
            .unwrap()
            .web()
            .unwrap()
            .unwrap()
            .0
            .save_smart_tags_as_xml(),
        Some(true)
    );
}

#[test]
fn rejects_ambiguous_or_external_web_settings_relationships() {
    use crate::web::{Conformance, Settings};
    use litchi_opc::constants::relationship_type as rt;

    let mut duplicate = Package::new().unwrap();
    let doc_uri = PackURI::new("/word/document.xml").unwrap();
    duplicate
        .opc
        .get_part_mut(&doc_uri)
        .unwrap()
        .rels_mut()
        .add_relationship(
            Conformance::Strict.relationship().to_owned(),
            "webSettings.xml".to_owned(),
            "rIdDuplicateWebSettings".to_owned(),
            false,
        );
    assert!(duplicate.document().unwrap().web().is_err());
    assert!(duplicate.web().is_err());
    assert!(
        duplicate
            .put_web(Settings::default(), Conformance::Transitional)
            .is_err()
    );

    let mut external = Package::new().unwrap();
    let relationship_id = external
        .opc
        .get_part(&doc_uri)
        .unwrap()
        .rels()
        .part_with_reltype(rt::WEB_SETTINGS)
        .unwrap()
        .r_id()
        .to_owned();
    let document_part = external.opc.get_part_mut(&doc_uri).unwrap();
    document_part.rels_mut().remove(&relationship_id);
    document_part.rels_mut().add_relationship(
        rt::WEB_SETTINGS.to_owned(),
        "https://example.invalid/webSettings.xml".to_owned(),
        relationship_id,
        true,
    );
    assert!(external.document().unwrap().web().is_err());
    assert!(external.web().is_err());
}

fn settings_state(package: &Package) -> (Vec<u8>, String) {
    let settings_uri = PackURI::new("/word/settings.xml").unwrap();
    let part = package.opc.get_part(&settings_uri).unwrap();
    (part.blob().to_vec(), part.rels().to_xml())
}

#[test]
fn adds_replaces_removes_and_reopens_attached_template() {
    use crate::settings::is_attached_template_relationship;

    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    package
        .set_attached_template_uri("file:///templates/Corporate.dotx")
        .unwrap();
    let attached = package.attached_template().unwrap().unwrap();
    let relationship_id = attached.relationship_id().to_owned();
    assert_eq!(attached.target_uri(), "file:///templates/Corporate.dotx");

    package
        .set_attached_template_uri("https://example.test/New.dotx?a=1&b=2")
        .unwrap();
    let replacement = package.attached_template().unwrap().unwrap();
    assert_eq!(replacement.relationship_id(), relationship_id);
    assert_eq!(
        replacement.target_uri(),
        "https://example.test/New.dotx?a=1&b=2"
    );
    let settings_uri = PackURI::new("/word/settings.xml").unwrap();
    let settings_part = package.opc.get_part(&settings_uri).unwrap();
    assert_eq!(
        settings_part
            .rels()
            .iter()
            .filter(|relationship| is_attached_template_relationship(relationship.reltype()))
            .count(),
        1
    );

    package.save(file.path()).unwrap();
    let mut reopened = Package::open(file.path()).unwrap();
    assert_eq!(
        reopened.attached_template().unwrap().unwrap().target_uri(),
        "https://example.test/New.dotx?a=1&b=2"
    );
    let removed = reopened.remove_attached_template().unwrap().unwrap();
    assert_eq!(removed.relationship_id(), relationship_id);
    assert!(reopened.attached_template().unwrap().is_none());
    let part = reopened.opc.get_part(&settings_uri).unwrap();
    assert!(!String::from_utf8_lossy(part.blob()).contains("attachedTemplate"));
    assert!(
        !part
            .rels()
            .iter()
            .any(|relationship| is_attached_template_relationship(relationship.reltype()))
    );
}

#[test]
fn attached_template_mutation_preserves_unrelated_xml_and_relationships() {
    let mut package = Package::new().unwrap();
    let settings_uri = PackURI::new("/word/settings.xml").unwrap();
    let part = package.opc.get_part_mut(&settings_uri).unwrap();
    part.set_blob(br#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:opaque"><!--keep--><q:zoom q:percent="137"/><x:opaque><![CDATA[a < b]]></x:opaque></q:settings>"#.to_vec());
    part.rels_mut().add_relationship(
        "urn:unrelated".to_owned(),
        "https://example.test/keep?a=1&b=2".to_owned(),
        "customRelationship".to_owned(),
        true,
    );

    package
        .set_attached_template_uri("file:///templates/Keep.dotx")
        .unwrap();
    let part = package.opc.get_part(&settings_uri).unwrap();
    let xml = String::from_utf8_lossy(part.blob());
    assert!(
        xml.contains(
            r#"<!--keep--><q:zoom q:percent="137"/><x:opaque><![CDATA[a < b]]></x:opaque>"#
        )
    );
    let unrelated = part.rels().get("customRelationship").unwrap();
    assert_eq!(unrelated.reltype(), "urn:unrelated");
    assert_eq!(unrelated.target_ref(), "https://example.test/keep?a=1&b=2");
}

#[test]
fn reads_strict_prefixed_attached_template_and_rewrites_word_compatible_type() {
    use crate::settings::STRICT_ATTACHED_TEMPLATE_RELATIONSHIP;

    let mut package = Package::new().unwrap();
    let settings_uri = PackURI::new("/word/settings.xml").unwrap();
    let part = package.opc.get_part_mut(&settings_uri).unwrap();
    part.set_blob(br#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:rel="http://purl.oclc.org/ooxml/officeDocument/relationships"><s:attachedTemplate rel:id="arbitrary-id"/></s:settings>"#.to_vec());
    part.rels_mut().add_relationship(
        STRICT_ATTACHED_TEMPLATE_RELATIONSHIP.to_owned(),
        "file:///strict.dotx".to_owned(),
        "arbitrary-id".to_owned(),
        true,
    );
    assert_eq!(
        package.attached_template().unwrap().unwrap().target_uri(),
        "file:///strict.dotx"
    );

    package
        .set_attached_template_uri("file:///compatible.dotx")
        .unwrap();
    let part = package.opc.get_part(&settings_uri).unwrap();
    let relationship = part.rels().get("arbitrary-id").unwrap();
    assert_eq!(relationship.reltype(), ATTACHED_TEMPLATE_RELATIONSHIP);
    assert!(relationship.is_external());
    assert!(
        String::from_utf8_lossy(part.blob())
            .contains(r#"<s:attachedTemplate rel:id="arbitrary-id"/>"#)
    );
}

#[test]
fn attached_template_failures_are_atomic() {
    let mut invalid_target = Package::new().unwrap();
    let before = settings_state(&invalid_target);
    assert!(
        invalid_target
            .set_attached_template_uri("file:///bad path.dotx")
            .is_err()
    );
    assert_eq!(settings_state(&invalid_target), before);

    let mut malformed = Package::new().unwrap();
    malformed
        .set_attached_template_uri("file:///valid.dotx")
        .unwrap();
    let settings_uri = PackURI::new("/word/settings.xml").unwrap();
    malformed
        .opc
        .get_part_mut(&settings_uri)
        .unwrap()
        .rels_mut()
        .add_relationship(
            ATTACHED_TEMPLATE_RELATIONSHIP.to_owned(),
            "file:///duplicate.dotx".to_owned(),
            "duplicate-id".to_owned(),
            true,
        );
    let before = settings_state(&malformed);
    assert!(
        malformed
            .set_attached_template_uri("file:///replacement.dotx")
            .is_err()
    );
    assert_eq!(settings_state(&malformed), before);
    assert!(malformed.remove_attached_template().is_err());
    assert_eq!(settings_state(&malformed), before);
}

#[test]
fn protection_rewrite_preserves_attached_template_relationship() {
    use crate::settings::ProtectionType;

    let mut package = Package::new().unwrap();
    package
        .set_attached_template_uri("file:///templates/Protected.dotx")
        .unwrap();
    let relationship_id = package
        .attached_template()
        .unwrap()
        .unwrap()
        .relationship_id()
        .to_owned();
    package
        .document_mut()
        .unwrap()
        .set_protection(ProtectionType::ReadOnly);
    package.to_stream(Cursor::new(Vec::new())).unwrap();

    let attached = package.attached_template().unwrap().unwrap();
    assert_eq!(attached.relationship_id(), relationship_id);
    assert_eq!(attached.target_uri(), "file:///templates/Protected.dotx");
    let settings_uri = PackURI::new("/word/settings.xml").unwrap();
    assert!(
        package
            .opc
            .get_part(&settings_uri)
            .unwrap()
            .rels()
            .get(&relationship_id)
            .is_some()
    );
}

#[test]
fn document_variable_package_lifecycle_is_deterministic_and_reopens() {
    let file = NamedTempFile::with_suffix(".docx").unwrap();
    let mut package = Package::new().unwrap();
    assert_eq!(
        package
            .set_document_variable("Company & Team", "A < B")
            .unwrap(),
        None
    );
    package.set_document_variable("second", "two").unwrap();
    assert_eq!(
        package
            .set_document_variable("Company & Team", "updated")
            .unwrap(),
        Some("A < B".into())
    );
    assert_eq!(
        package.document_variables().unwrap().unwrap().names(),
        vec!["Company & Team", "second"]
    );
    package.save(file.path()).unwrap();

    let mut reopened = Package::open(file.path()).unwrap();
    let variables = reopened
        .document()
        .unwrap()
        .document_variables()
        .unwrap()
        .unwrap();
    assert_eq!(variables.get("Company & Team"), Some("updated"));
    assert_eq!(variables.get("second"), Some("two"));
    assert_eq!(
        reopened.remove_document_variable("Company & Team").unwrap(),
        Some("updated".into())
    );
    assert_eq!(reopened.clear_document_variables().unwrap(), 1);
    assert!(reopened.document_variables().unwrap().unwrap().is_empty());
    let settings_uri = PackURI::new("/word/settings.xml").unwrap();
    assert!(
        !String::from_utf8_lossy(reopened.opc.get_part(&settings_uri).unwrap().blob())
            .contains("docVars")
    );
}

#[test]
fn document_variable_mutation_preserves_xml_relationships_and_protection() {
    use crate::settings::ProtectionType;

    let mut package = Package::new().unwrap();
    let settings_uri = PackURI::new("/word/settings.xml").unwrap();
    package.opc.get_part_mut(&settings_uri).unwrap().set_blob(
            br#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:opaque"><!--keep--><q:zoom q:percent="133"/><x:opaque><![CDATA[a < b]]></x:opaque><q:rsids/></q:settings>"#.to_vec(),
        );
    package
        .opc
        .get_part_mut(&settings_uri)
        .unwrap()
        .rels_mut()
        .add_relationship(
            "urn:unrelated".to_owned(),
            "https://example.test/keep".to_owned(),
            "keep-id".to_owned(),
            true,
        );
    package
        .set_attached_template_uri("file:///templates/Variables.dotx")
        .unwrap();
    let attached_id = package
        .attached_template()
        .unwrap()
        .unwrap()
        .relationship_id()
        .to_owned();
    package.set_document_variable("project", "Litchi").unwrap();
    let part = package.opc.get_part(&settings_uri).unwrap();
    let xml = String::from_utf8_lossy(part.blob());
    assert!(
        xml.contains(
            r#"<!--keep--><q:zoom q:percent="133"/><x:opaque><![CDATA[a < b]]></x:opaque>"#
        )
    );
    assert!(part.rels().get("keep-id").is_some());
    assert!(part.rels().get(&attached_id).is_some());

    package
        .document_mut()
        .unwrap()
        .set_protection(ProtectionType::ReadOnly);
    package.to_stream(Cursor::new(Vec::new())).unwrap();
    assert_eq!(
        package
            .document_variables()
            .unwrap()
            .unwrap()
            .get("project"),
        Some("Litchi")
    );
    let part = package.opc.get_part(&settings_uri).unwrap();
    assert!(part.rels().get("keep-id").is_some());
    assert!(part.rels().get(&attached_id).is_some());
}

#[test]
fn document_variable_mutation_failures_are_atomic() {
    let mut package = Package::new().unwrap();
    let before = settings_state(&package);
    assert!(package.set_document_variable("", "invalid").is_err());
    assert_eq!(settings_state(&package), before);
    assert!(
        package
            .set_document_variable("too-long", "x".repeat(65_281))
            .is_err()
    );
    assert_eq!(settings_state(&package), before);

    let settings_uri = PackURI::new("/word/settings.xml").unwrap();
    package.opc.get_part_mut(&settings_uri).unwrap().set_blob(
            br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docVars><w:docVar w:name="duplicate" w:val="one"/><w:docVar w:name="duplicate" w:val="two"/></w:docVars></w:settings>"#.to_vec(),
        );
    let malformed = settings_state(&package);
    assert!(package.set_document_variable("new", "value").is_err());
    assert_eq!(settings_state(&package), malformed);
    assert!(package.remove_document_variable("duplicate").is_err());
    assert_eq!(settings_state(&package), malformed);
    assert!(package.clear_document_variables().is_err());
    assert_eq!(settings_state(&package), malformed);
}

#[test]
fn reads_strict_settings_relationship_and_mce_fallback_variables() {
    use litchi_opc::constants::relationship_type as rt;

    const STRICT_SETTINGS_RELATIONSHIP: &str =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";
    let mut package = Package::new().unwrap();
    let doc_uri = PackURI::new("/word/document.xml").unwrap();
    let relationship_id = package
        .opc
        .get_part(&doc_uri)
        .unwrap()
        .rels()
        .part_with_reltype(rt::SETTINGS)
        .unwrap()
        .r_id()
        .to_owned();
    let document = package.opc.get_part_mut(&doc_uri).unwrap();
    let target = document
        .rels_mut()
        .remove(&relationship_id)
        .unwrap()
        .target_ref()
        .to_owned();
    document.rels_mut().add_relationship(
        STRICT_SETTINGS_RELATIONSHIP.to_owned(),
        target,
        relationship_id,
        false,
    );
    let settings_uri = PackURI::new("/word/settings.xml").unwrap();
    package.opc.get_part_mut(&settings_uri).unwrap().set_blob(
            br#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:unsupported" mc:Ignorable="u"><mc:AlternateContent><mc:Choice Requires="u"><s:docVars><s:docVar s:name="choice" s:val="ignored"/></s:docVars></mc:Choice><mc:Fallback><s:docVars><s:docVar s:name="fallback" s:val="selected"/></s:docVars></mc:Fallback></mc:AlternateContent></s:settings>"#.to_vec(),
        );

    let package_variables = package.document_variables().unwrap().unwrap();
    assert_eq!(package_variables.get("fallback"), Some("selected"));
    assert!(!package_variables.contains("choice"));
    let document_variables = package
        .document()
        .unwrap()
        .document_variables()
        .unwrap()
        .unwrap();
    assert_eq!(document_variables.get("fallback"), Some("selected"));
}
