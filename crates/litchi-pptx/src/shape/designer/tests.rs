#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::{EXTENSION_URI, NAMESPACE, codec, model::Snapshot, package, transaction};
use crate::tag::Conformance;
use litchi_opc::{OpcPackage, PackURI, XmlPart};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const FUTURE: &str = "urn:future-designer";

fn shape_with_design(value: &str) -> Vec<u8> {
    format!(
        r#"<p:sp xmlns:p="{PML}" xmlns:p15="{NAMESPACE}"><p:nvSpPr><p:cNvPr id="1" name="Box"/><p:cNvSpPr/><p:nvPr><p:extLst><p:ext uri="{EXTENSION_URI}"><p15:designElem val="{value}"/></p:ext></p:extLst></p:nvPr></p:nvSpPr><p:spPr/></p:sp>"#
    )
    .into_bytes()
}

fn slide_with_shape(shape: &str) -> Vec<u8> {
    format!(
        r#"<p:sld xmlns:p="{PML}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>{shape}</p:spTree></p:cSld></p:sld>"#
    )
    .into_bytes()
}

fn owner_package(xml: Vec<u8>) -> (OpcPackage, PackURI) {
    let owner = PackURI::new("/ppt/slides/slide1.xml").expect("owner URI");
    let mut package = OpcPackage::new();
    package.add_part(Box::new(XmlPart::new(
        owner.clone(),
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
        xml,
    )));
    (package, owner)
}

#[test]
fn reads_schema_boolean_forms_and_replaces_only_the_typed_range() {
    let original = shape_with_design("1");
    let source = codec::read(&original).expect("source");
    assert_eq!(
        source.snapshot.as_ref().and_then(Snapshot::value),
        Some(true)
    );

    let updated = transaction::set(&original, &source, false).expect("set");
    let parsed = codec::read(&updated).expect("updated source");
    assert_eq!(
        parsed.snapshot.as_ref().and_then(Snapshot::value),
        Some(false)
    );
    assert!(
        updated
            .windows(b"val=\"false\"".len())
            .any(|window| { window == b"val=\"false\"" })
    );
    assert!(contains(&updated, b"<p:spPr/>"));
}

#[test]
fn unknown_extension_entries_and_known_extension_children_are_lossless() {
    let xml = format!(
        r#"<p:sp xmlns:p="{PML}" xmlns:p15="{NAMESPACE}" xmlns:f="{FUTURE}"><p:nvSpPr><p:cNvPr id="1" name="Box"/><p:cNvSpPr/><p:nvPr><p:extLst><p:ext uri="urn:before"><f:payload answer="42"/></p:ext><p:ext uri="{EXTENSION_URI}"><f:future keep="yes"/><p15:designElem val="true"/></p:ext><p:ext uri="urn:after"/></p:extLst></p:nvPr></p:nvSpPr><p:spPr/></p:sp>"#
    )
    .into_bytes();
    let source = codec::read(&xml).expect("source");
    assert_eq!(
        source.snapshot.as_ref().and_then(Snapshot::value),
        Some(true)
    );
    assert_eq!(
        source
            .snapshot
            .as_ref()
            .expect("snapshot")
            .unknown_extensions()
            .len(),
        2
    );

    let updated = transaction::set(&xml, &source, false).expect("set");
    assert!(contains(&updated, b"<f:payload answer=\"42\"/>"));
    assert!(contains(&updated, b"<f:future keep=\"yes\"/>"));
    let removed = transaction::remove(&updated, &codec::read(&updated).expect("updated source"))
        .expect("remove")
        .expect("removed value");
    assert_eq!(
        codec::read(&removed).expect("removed source").snapshot,
        None
    );
    assert!(contains(&removed, b"<f:future keep=\"yes\"/>"));
    assert!(contains(&removed, b"uri=\"urn:before\""));
}

#[test]
fn creates_and_removes_extension_lists_without_touching_the_shape_tail() {
    let xml = br#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:nvSpPr><p:cNvPr id="1" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>"#;
    let source = codec::read(xml).expect("source");
    let updated = transaction::set(xml, &source, true).expect("set");
    let parsed = codec::read(&updated).expect("updated source");
    assert_eq!(
        parsed.snapshot.as_ref().and_then(Snapshot::value),
        Some(true)
    );
    assert!(contains(&updated, b"<p:spPr/>"));

    let source = codec::read(&updated).expect("updated source");
    let removed = transaction::remove(&updated, &source)
        .expect("remove")
        .expect("removed value");
    assert_eq!(
        codec::read(&removed).expect("removed source").snapshot,
        None
    );
    assert!(contains(&removed, b"<p:nvPr/>"));
}

#[test]
fn detached_edits_are_atomic_and_preserve_the_snapshot_contract() {
    let snapshot = Snapshot::new(true);
    let mut editor = snapshot.edit();
    editor.set(false).expect("set");
    assert!(editor.is_changed());
    assert_eq!(editor.snapshot().value(), Some(false));
    let committed = editor.commit().expect("commit");
    assert_eq!(snapshot.value(), Some(true));
    assert_eq!(committed.value(), Some(false));

    let mut editor = committed.edit();
    editor.clear();
    assert_eq!(editor.commit().expect("clear").value(), None);
}

#[test]
fn package_operations_are_selector_first_and_atomic() {
    let shape = shape_with_design("true");
    let (mut package, owner) = owner_package(slide_with_shape(
        std::str::from_utf8(&shape).expect("shape UTF-8"),
    ));

    assert_eq!(
        package::load(&package, &owner, "Box")
            .expect("load")
            .and_then(|snapshot| snapshot.value()),
        Some(true)
    );
    let previous = package::put(&mut package, &owner, 0_usize, false)
        .expect("put")
        .expect("previous");
    assert_eq!(previous.value(), Some(true));
    assert_eq!(
        package::load(&package, &owner, "Box")
            .expect("load")
            .and_then(|snapshot| snapshot.value()),
        Some(false)
    );
    let removed = package::remove(&mut package, &owner, "Box")
        .expect("remove")
        .expect("removed");
    assert_eq!(removed.value(), Some(false));
    assert!(
        package::load(&package, &owner, 0_usize)
            .expect("load after remove")
            .is_none()
    );
}

#[test]
fn public_package_and_slide_facades_use_the_same_shape_selector() {
    let mut authored = crate::package::Package::new().expect("package");
    {
        let presentation = authored.presentation_mut().expect("mutable presentation");
        let slide = presentation.add_slide().expect("slide");
        slide.add_text_box("Designer", 914_400, 914_400, 2_743_200, 914_400);
    }
    let bytes = authored.to_bytes().expect("serialize");
    let mut package = crate::package::Package::from_bytes(&bytes).expect("reopen");

    assert!(
        package
            .shape_design_element("Slide 256", "TextBox")
            .expect("read")
            .is_none()
    );
    assert!(
        package
            .put_shape_design_element("Slide 256", "TextBox", true)
            .expect("put")
            .is_none()
    );
    let slide = package
        .presentation()
        .expect("presentation")
        .slide(0)
        .expect("slide lookup")
        .expect("slide");
    assert_eq!(
        slide
            .shape_design_element("TextBox")
            .expect("slide read")
            .and_then(|snapshot| snapshot.value()),
        Some(true)
    );
    assert_eq!(
        package
            .remove_shape_design_element("Slide 256", 0_usize)
            .expect("remove")
            .and_then(|snapshot| snapshot.value()),
        Some(true)
    );
}

#[test]
fn validation_rejects_invalid_values_and_duplicate_known_extensions() {
    let invalid = shape_with_design("maybe");
    assert!(codec::read(&invalid).is_err());

    let duplicate = format!(
        r#"<p:sp xmlns:p="{PML}" xmlns:p15="{NAMESPACE}"><p:nvSpPr><p:cNvPr id="1" name="Box"/><p:cNvSpPr/><p:nvPr><p:extLst><p:ext uri="{EXTENSION_URI}"><p15:designElem val="true"/></p:ext><p:ext uri="{EXTENSION_URI}"><p15:designElem val="false"/></p:ext></p:extLst></p:nvPr><p:spPr/></p:sp>"#
    )
    .into_bytes();
    assert!(codec::read(&duplicate).is_err());
}

#[test]
fn strict_profile_uses_the_strict_presentation_namespace() {
    let xml = br#"<p:sp xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"><p:nvSpPr><p:cNvPr id="1" name="Box"/><p:cNvSpPr/><p:nvPr><p:extLst/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>"#;
    let source = codec::read(xml).expect("source");
    let updated = transaction::set(xml, &source, true).expect("set");
    assert!(contains(
        &updated,
        b"http://purl.oclc.org/ooxml/presentationml/main"
    ));
    assert_eq!(
        codec::read(&updated)
            .expect("updated source")
            .layout
            .conformance,
        Conformance::Strict
    );
}

fn contains(haystack: &[u8], needle: &[u8]) -> bool {
    haystack
        .windows(needle.len())
        .any(|window| window == needle)
}
