use litchi_ooxml::pptx::Package;
use litchi_ooxml::{OoxmlError, PackURI};
use litchi_opc::part::BlobPart;
use tempfile::NamedTempFile;

const TAG_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags";
const TAG_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.tags+xml";
const LOCAL_PRIMARY_TAGS: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/tags/basic_tags.xml");
const LOCAL_SECONDARY_TAGS: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/tags/secondary_tags.xml");

#[test]
fn package_inventory_reports_local_tag_lists() {
    let package = package_with_local_tag_lists();

    let tag_lists = package.tag_lists().unwrap();
    assert_eq!(tag_lists.len(), 2);

    let primary = &tag_lists[0];
    assert_eq!(primary.slide_index(), 0);
    assert_eq!(primary.tag_list_index(), 0);
    assert_eq!(primary.relationship_id(), "rIdTags1");
    assert_eq!(primary.part_name(), "/ppt/tags/tag1.xml");
    assert_eq!(primary.tag_list().tags().len(), 2);
    assert_eq!(primary.tag_list().tags()[0].name(), "OWNER");
    assert_eq!(primary.tag_list().tags()[0].value(), "Alice");
    assert_eq!(primary.tag_list().tags()[1].value(), "<not-a-command/>");
    assert_eq!(
        primary.tag_list().extension_attributes()[0].qualified_name(),
        "ext:origin"
    );
    assert_eq!(
        primary.tag_list().extension_attributes()[0].value(),
        "local"
    );

    let secondary = &tag_lists[1];
    assert_eq!(secondary.slide_index(), 1);
    assert_eq!(secondary.tag_list_index(), 0);
    assert_eq!(secondary.relationship_id(), "rIdTags2");
    assert_eq!(secondary.part_name(), "/ppt/tags/tag2.xml");
    assert_eq!(secondary.tag_list().tags()[0].name(), "STATUS");
    assert_eq!(secondary.tag_list().tags()[0].value(), "Review");

    assert_eq!(
        package.presentation().unwrap().tag_lists().unwrap(),
        tag_lists
    );
}

#[test]
fn package_inventory_rejects_external_tag_relationships() {
    let mut package = package_with_local_tag_lists();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.opc_package_mut().get_part_mut(&slide_name).unwrap();
    slide.rels_mut().remove("rIdTags1");
    slide.rels_mut().add_relationship(
        TAG_RELATIONSHIP_TYPE.to_string(),
        "https://example.invalid/tags.xml".to_string(),
        "rIdExternalTags".to_string(),
        true,
    );

    assert!(matches!(
        package.tag_lists(),
        Err(OoxmlError::InvalidFormat(message)) if message.contains("cannot be external")
    ));
}

#[test]
fn package_inventory_rejects_wrong_tag_content_type() {
    let mut package = package_with_local_tag_lists();
    let part_name = PackURI::new("/ppt/tags/tag1.xml").unwrap();
    assert!(package.opc_package_mut().remove_part(&part_name));
    package.opc_package_mut().add_part(Box::new(BlobPart::new(
        part_name,
        "application/xml".to_string(),
        LOCAL_PRIMARY_TAGS.to_vec(),
    )));

    assert!(matches!(
        package.tag_lists(),
        Err(OoxmlError::InvalidContentType { expected, got })
            if expected == TAG_CONTENT_TYPE && got == "application/xml"
    ));
}

fn package_with_local_tag_lists() -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    install_local_tag_lists(&mut package);
    package
}

fn install_local_tag_lists(package: &mut Package) {
    for (slide_name, target, relationship_id, part_name, xml) in [
        (
            "/ppt/slides/slide1.xml",
            "../tags/tag1.xml",
            "rIdTags1",
            "/ppt/tags/tag1.xml",
            LOCAL_PRIMARY_TAGS,
        ),
        (
            "/ppt/slides/slide2.xml",
            "../tags/tag2.xml",
            "rIdTags2",
            "/ppt/tags/tag2.xml",
            LOCAL_SECONDARY_TAGS,
        ),
    ] {
        let slide_name = PackURI::new(slide_name).unwrap();
        let part_name = PackURI::new(part_name).unwrap();
        package.opc_package_mut().add_part(Box::new(BlobPart::new(
            part_name,
            TAG_CONTENT_TYPE.to_string(),
            xml.to_vec(),
        )));
        package
            .opc_package_mut()
            .get_part_mut(&slide_name)
            .unwrap()
            .rels_mut()
            .add_relationship(
                TAG_RELATIONSHIP_TYPE.to_string(),
                target.to_string(),
                relationship_id.to_string(),
                false,
            );
    }
}
