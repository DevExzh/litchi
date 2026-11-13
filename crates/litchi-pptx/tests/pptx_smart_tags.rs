use litchi_opc::PackURI;
use litchi_opc::constants::content_type as ct;
use litchi_opc::part::{BlobPart, Part};
use litchi_pptx::parts::PresentationPart;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/smart-tags/presentation.xml");

#[test]
fn presentation_smart_tags_relationship_is_retained_by_the_root_owner() {
    // SmartTags is an inert, producer-defined root relationship. The
    // standalone root owner intentionally preserves it without exposing a
    // format-specific convenience method for a feature with no semantic CRUD.
    let mut part = BlobPart::new(
        PackURI::new("/ppt/presentation.xml").unwrap(),
        ct::PML_PRESENTATION_MAIN.to_owned(),
        PRESENTATION_XML.to_vec(),
    );
    part.rels_mut().add_relationship(
        "urn:litchi:smart-tags".to_owned(),
        "smartTags.xml".to_owned(),
        "rIdSmartTags".to_owned(),
        false,
    );

    let presentation = PresentationPart::from_part(&part).unwrap();
    let relationship = presentation.part().rels().get("rIdSmartTags").unwrap();
    assert_eq!(relationship.r_id(), "rIdSmartTags");
    assert_eq!(relationship.target_ref(), "smartTags.xml");
    assert!(
        std::str::from_utf8(presentation.part().blob())
            .unwrap()
            .contains(r#"<p:smartTags r:id="rIdSmartTags"/>"#)
    );
}
