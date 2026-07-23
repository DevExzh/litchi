use litchi_ooxml::OoxmlError;
use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::Package;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/handout-master/presentation.xml");
const HANDOUT_MASTER_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/handout-master/handout-master.xml");
const WRONG_ROOT_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/handout-master/wrong-root.xml");
const STRICT_HANDOUT_MASTER_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/handoutMaster";

#[test]
fn presentation_handout_master_is_resolved() {
    let package = package_with_handout_master();
    let presentation = package.presentation().unwrap();

    assert_eq!(
        presentation.handout_master_relationship_id().unwrap(),
        Some("rIdHandout".to_string())
    );
    let handout_master = presentation.handout_master().unwrap().unwrap();
    assert!(handout_master.header_footer.show_header);
    assert!(handout_master.header_footer.show_footer);
    assert!(handout_master.header_footer.show_slide_number);
    assert!(handout_master.header_footer.show_date_time);
    assert_eq!(handout_master.background_color.as_deref(), Some("112233"));
}

#[test]
fn handout_master_relationship_is_validated() {
    let mut package = package_with_handout_master();
    replace_handout_relationship(
        &mut package,
        rt::THEME,
        "handoutMasters/handoutMaster1.xml",
        false,
    );
    assert!(matches!(
        package.presentation().unwrap().handout_master(),
        Err(OoxmlError::InvalidRelationship(message))
            if message.contains("is not a handout-master relationship")
    ));

    let mut package = package_with_handout_master();
    replace_handout_relationship(
        &mut package,
        rt::HANDOUT_MASTER,
        "https://example.invalid/handoutMaster.xml",
        true,
    );
    assert!(matches!(
        package.presentation().unwrap().handout_master(),
        Err(OoxmlError::InvalidRelationship(message)) if message.contains("must be internal")
    ));

    let mut package = package_with_handout_master();
    replace_handout_relationship(
        &mut package,
        rt::HANDOUT_MASTER,
        "notesMasters/notesMaster1.xml",
        false,
    );
    assert!(matches!(
        package.presentation().unwrap().handout_master(),
        Err(OoxmlError::InvalidContentType { expected, got })
            if expected == ct::PML_HANDOUT_MASTER && got == ct::PML_NOTES_MASTER
    ));
}

#[test]
fn strict_handout_master_relationship_is_supported() {
    let mut package = package_with_handout_master();
    replace_handout_relationship(
        &mut package,
        STRICT_HANDOUT_MASTER_RELATIONSHIP_TYPE,
        "handoutMasters/handoutMaster1.xml",
        false,
    );

    assert!(
        package
            .presentation()
            .unwrap()
            .handout_master()
            .unwrap()
            .is_some()
    );
}

#[test]
fn handout_master_root_is_validated() {
    let mut package = package_with_handout_master();
    let handout_name = PackURI::new("/ppt/handoutMasters/handoutMaster1.xml").unwrap();
    package
        .opc_package_mut()
        .get_part_mut(&handout_name)
        .unwrap()
        .set_blob(WRONG_ROOT_XML.to_vec());

    assert!(matches!(
        package.presentation().unwrap().handout_master(),
        Err(OoxmlError::InvalidFormat(message)) if message.contains("handoutMaster root")
    ));
}

fn package_with_handout_master() -> Package {
    let mut package = Package::new().unwrap();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let handout_name = PackURI::new("/ppt/handoutMasters/handoutMaster1.xml").unwrap();

    {
        let presentation = package
            .opc_package_mut()
            .get_part_mut(&presentation_name)
            .unwrap();
        presentation.set_blob(PRESENTATION_XML.to_vec());
        presentation.rels_mut().add_relationship(
            rt::HANDOUT_MASTER.to_string(),
            "handoutMasters/handoutMaster1.xml".to_string(),
            "rIdHandout".to_string(),
            false,
        );
    }
    package.opc_package_mut().add_part(Box::new(BlobPart::new(
        handout_name,
        ct::PML_HANDOUT_MASTER.to_string(),
        HANDOUT_MASTER_XML.to_vec(),
    )));
    package
}

fn replace_handout_relationship(
    package: &mut Package,
    relationship_type: &str,
    target: &str,
    is_external: bool,
) {
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let presentation = package
        .opc_package_mut()
        .get_part_mut(&presentation_name)
        .unwrap();
    presentation.rels_mut().remove("rIdHandout");
    presentation.rels_mut().add_relationship(
        relationship_type.to_string(),
        target.to_string(),
        "rIdHandout".to_string(),
        is_external,
    );
}
