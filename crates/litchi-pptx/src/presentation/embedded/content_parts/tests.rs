use super::{Limits, load_slide};
use crate::Error;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";

#[test]
fn loads_opaque_payload_and_relationship_metadata_losslessly() {
    let payload = b"<?xml version=\"1.0\"?><opaque:root xmlns:opaque=\"urn:opaque\"><opaque:value>raw &amp; bytes</opaque:value></opaque:root>";
    let (package, slide_name) = package_with_internal_payload(payload);
    let slide = package.get_part(&slide_name).expect("slide");
    let mut limits = Limits::default();
    let parts = load_slide(&package, 3, slide, &mut limits).expect("content part");

    assert_eq!(parts.len(), 1);
    let content_part = &parts[0];
    assert_eq!(content_part.slide_index(), 3);
    assert_eq!(content_part.index(), 0);
    assert_eq!(content_part.relationship_id(), "rIdOpaque");
    assert_eq!(
        content_part.anchor().xml(),
        b"<p:contentPart r:id=\"rIdOpaque\"><p14:nvContentPartPr/></p:contentPart>"
    );
    assert_eq!(content_part.relationship().id(), "rIdOpaque");
    assert_eq!(
        content_part.relationship().relationship_type(),
        rt::CUSTOM_XML
    );
    assert_eq!(
        content_part.relationship().target_ref(),
        "../custom/opaque.xml"
    );

    let payload_view = content_part.payload().expect("internal payload");
    assert_eq!(payload_view.part_name().as_str(), "/ppt/custom/opaque.xml");
    assert_eq!(payload_view.content_type(), "application/xml");
    assert_eq!(payload_view.bytes(), payload);
    assert_eq!(payload_view.relationships().len(), 1);
    assert_eq!(payload_view.relationships()[0].id(), "rIdPayloadLink");
    assert_eq!(
        payload_view.relationships()[0].target_ref(),
        "https://example.invalid/opaque"
    );
}

#[test]
fn discovery_is_a_source_preserving_noop() {
    let (package, slide_name) = package_with_internal_payload(b"opaque");
    let before = package
        .get_part(&slide_name)
        .expect("slide")
        .blob()
        .to_vec();
    let slide = package.get_part(&slide_name).expect("slide");
    let _ = load_slide(&package, 0, slide, &mut Limits::default()).expect("content part");
    assert_eq!(package.get_part(&slide_name).unwrap().blob(), before);
}

#[test]
fn retains_external_target_mode_without_following_it() {
    let (package, slide_name) = package_with_external_payload();
    let slide = package.get_part(&slide_name).expect("slide");
    let parts = load_slide(&package, 0, slide, &mut Limits::default()).expect("content part");
    let relationship = parts[0].relationship();

    assert_eq!(relationship.target_mode(), litchi_opc::TargetMode::External);
    assert_eq!(relationship.target_ref(), "https://example.invalid/content");
    assert!(relationship.payload().is_none());
    assert_eq!(
        relationship.target().external_ref(),
        Some("https://example.invalid/content")
    );
}

#[test]
fn rejects_missing_anchor_relationship() {
    let package = package_with_anchor(b"<p:contentPart r:id=\"rIdMissing\"/>", None, None);
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.get_part(&slide_name).expect("slide");
    let error = load_slide(&package, 0, slide, &mut Limits::default()).unwrap_err();
    assert!(matches!(error, Error::Relationship(message) if message.contains("rIdMissing")));
}

#[test]
fn rejects_anchor_without_relationship_id() {
    let package = package_with_anchor(b"<p:contentPart/>", None, None);
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.get_part(&slide_name).expect("slide");
    let error = load_slide(&package, 0, slide, &mut Limits::default()).unwrap_err();
    assert!(matches!(error, Error::Invalid(message) if message.contains("missing r:id")));
}

#[test]
fn enforces_content_part_count_limit() {
    let package = package_with_anchor(
        b"<p:contentPart r:id=\"rId1\"/><p:contentPart r:id=\"rId2\"/>",
        Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rId1", false)),
        Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rId2", false)),
    );
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.get_part(&slide_name).expect("slide");
    let mut limits = Limits::new(1, 1024, 4096, 8).unwrap();
    let error = load_slide(&package, 0, slide, &mut limits).unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            resource: "content-part count",
            ..
        }
    ));
}

fn package_with_internal_payload(payload: &[u8]) -> (OpcPackage, PackURI) {
    let relationship = Some((rt::CUSTOM_XML, "../custom/opaque.xml", "rIdOpaque", false));
    let mut package = package_with_anchor(
        b"<p:contentPart r:id=\"rIdOpaque\"><p14:nvContentPartPr/></p:contentPart>",
        relationship,
        None,
    );
    let payload_name = PackURI::new("/ppt/custom/opaque.xml").unwrap();
    package
        .get_part_mut(&payload_name)
        .expect("payload")
        .set_blob(payload.to_vec());
    (package, PackURI::new("/ppt/slides/slide1.xml").unwrap())
}

fn package_with_external_payload() -> (OpcPackage, PackURI) {
    let package = package_with_anchor(
        b"<p:contentPart r:id=\"rIdExternal\"/>",
        Some((
            rt::CUSTOM_XML,
            "https://example.invalid/content",
            "rIdExternal",
            true,
        )),
        None,
    );
    (package, PackURI::new("/ppt/slides/slide1.xml").unwrap())
}

fn package_with_anchor(
    anchors: &[u8],
    relationship: Option<(&str, &str, &str, bool)>,
    relationship_2: Option<(&str, &str, &str, bool)>,
) -> OpcPackage {
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let xml = format!(
        "<p:sld xmlns:p=\"{PML}\" xmlns:r=\"{REL}\" xmlns:p14=\"{P14}\"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>{}</p:spTree></p:cSld></p:sld>",
        String::from_utf8(anchors.to_vec()).unwrap()
    );
    let mut slide = BlobPart::new(slide_name.clone(), ct::PML_SLIDE.into(), xml.into_bytes());
    if let Some((relationship_type, target, id, external)) = relationship {
        slide.rels_mut().add_relationship(
            relationship_type.to_owned(),
            target.to_owned(),
            id.to_owned(),
            external,
        );
    }
    if let Some((relationship_type, target, id, external)) = relationship_2 {
        slide.rels_mut().add_relationship(
            relationship_type.to_owned(),
            target.to_owned(),
            id.to_owned(),
            external,
        );
    }
    let mut package = OpcPackage::new();
    package.add_part(Box::new(slide));
    let payload_name = PackURI::new("/ppt/custom/opaque.xml").unwrap();
    let mut payload = BlobPart::new(
        payload_name,
        "application/xml".to_owned(),
        b"opaque payload".to_vec(),
    );
    payload.rels_mut().add_relationship(
        rt::HYPERLINK.to_owned(),
        "https://example.invalid/opaque".to_owned(),
        "rIdPayloadLink".to_owned(),
        true,
    );
    package.add_part(Box::new(payload));
    package
}
