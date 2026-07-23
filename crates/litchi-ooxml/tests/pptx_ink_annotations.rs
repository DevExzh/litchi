use litchi_ooxml::pptx::{INK_CONTENT_TYPE, Package};
use litchi_ooxml::{OoxmlError, PackURI};
use litchi_opc::XmlPart;
use litchi_opc::constants::relationship_type::CUSTOM_XML;
use tempfile::NamedTempFile;

const LOCAL_INK: &[u8] = include_bytes!("../../../test-data/ooxml/pptx/ink/basic_ink.xml");

#[test]
fn package_inventory_reports_local_ink_content_parts() {
    let package = package_with_ink();

    let annotations = package.ink_annotations().unwrap();
    assert_eq!(annotations.len(), 1);

    let annotation = &annotations[0];
    assert_eq!(annotation.slide_index(), 0);
    assert_eq!(annotation.content_part_index(), 0);
    assert_eq!(annotation.relationship_id(), "rIdInk");
    assert_eq!(annotation.part_name().as_str(), "/ppt/ink/ink1.xml");
    assert_eq!(annotation.trace_count(), 2);
    assert_eq!(annotation.trace_group_count(), 1);

    assert_eq!(
        package.presentation().unwrap().ink_annotations().unwrap(),
        annotations
    );
}

#[test]
fn package_inventory_rejects_missing_ink_targets() {
    let mut package = package_with_ink();
    let part_name = PackURI::new("/ppt/ink/ink1.xml").unwrap();
    assert!(package.opc_package_mut().remove_part(&part_name));

    let error = package.ink_annotations().unwrap_err();
    assert!(matches!(
        error,
        OoxmlError::PartNotFound(message) if message.contains("/ppt/ink/ink1.xml")
    ));
}

#[test]
fn package_inventory_rejects_malformed_inkml() {
    let mut package = package_with_ink();
    let part_name = PackURI::new("/ppt/ink/ink1.xml").unwrap();
    package
        .opc_package_mut()
        .get_part_mut(&part_name)
        .unwrap()
        .set_blob(b"<ink/>".to_vec());

    let error = package.ink_annotations().unwrap_err();
    assert!(matches!(error, OoxmlError::InvalidFormat(_)));
}

fn package_with_ink() -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    install_local_ink(&mut package);
    package
}

fn install_local_ink(package: &mut Package) {
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    {
        let slide = package.opc_package_mut().get_part_mut(&slide_name).unwrap();
        let xml = std::str::from_utf8(slide.blob()).unwrap();
        let updated = xml.replacen(
            "</p:spTree>",
            "<p:contentPart r:id=\"rIdInk\"/></p:spTree>",
            1,
        );
        assert_ne!(updated, xml);
        slide.set_blob(updated.into_bytes());
        slide.rels_mut().add_relationship(
            CUSTOM_XML.to_string(),
            "../ink/ink1.xml".to_string(),
            "rIdInk".to_string(),
            false,
        );
    }

    package.opc_package_mut().add_part(Box::new(XmlPart::new(
        PackURI::new("/ppt/ink/ink1.xml").unwrap(),
        INK_CONTENT_TYPE.to_string(),
        LOCAL_INK.to_vec(),
    )));
}
