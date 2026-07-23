use litchi_ooxml::pptx::{Package, PptxOleObjectMode, PptxOleObjectTarget, PptxOlePayloadKind};
use litchi_ooxml::{OoxmlError, PackURI};
use litchi_opc::constants::{
    content_type::{OFC_OLE_OBJECT, OFC_PACKAGE},
    relationship_type::{OLE_OBJECT, PACKAGE},
};
use litchi_opc::part::BlobPart;
use tempfile::NamedTempFile;

const LOCAL_OLE_OBJECTS: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/ole/basic_ole_objects.xml");

#[test]
fn package_inventory_reports_local_ole_objects() {
    let package = package_with_local_ole_objects();

    let objects = package.ole_objects().unwrap();
    assert_eq!(objects.len(), 3);

    let workbook = &objects[0];
    assert_eq!(workbook.slide_index(), 0);
    assert_eq!(workbook.object_index(), 0);
    assert_eq!(workbook.shape_id(), Some(101));
    assert_eq!(workbook.shape_name(), Some("Embedded workbook"));
    assert_eq!(workbook.legacy_shape_id(), Some("_x0000_s1025"));
    assert_eq!(workbook.name(), Some("Workbook"));
    assert_eq!(workbook.program_id(), Some("Excel.Sheet.12"));
    assert_eq!(workbook.show_as_icon(), Some(true));
    assert_eq!(workbook.preview_width(), Some(914_400));
    assert_eq!(workbook.preview_height(), Some(457_200));
    assert_eq!(workbook.mode(), PptxOleObjectMode::Embedded);
    assert_eq!(workbook.relationship_id(), Some("rIdOle"));
    assert_eq!(workbook.payload_kind(), Some(PptxOlePayloadKind::OleObject));
    assert_eq!(workbook.preview_relationship_id(), Some("rIdPreview"));
    assert!(matches!(
        workbook.target(),
        Some(PptxOleObjectTarget::Internal {
            part_name,
            content_type,
            relationship_type,
        }) if part_name.as_str() == "/ppt/embeddings/oleObject1.bin"
            && content_type == OFC_OLE_OBJECT
            && relationship_type == OLE_OBJECT
    ));

    let package_object = &objects[1];
    assert_eq!(package_object.object_index(), 1);
    assert_eq!(package_object.mode(), PptxOleObjectMode::Embedded);
    assert_eq!(
        package_object.payload_kind(),
        Some(PptxOlePayloadKind::Package)
    );
    assert!(matches!(
        package_object.target(),
        Some(PptxOleObjectTarget::Internal {
            part_name,
            content_type,
            relationship_type,
        }) if part_name.as_str() == "/ppt/embeddings/package1.bin"
            && content_type == OFC_PACKAGE
            && relationship_type == PACKAGE
    ));

    let linked = &objects[2];
    assert_eq!(linked.object_index(), 2);
    assert_eq!(linked.mode(), PptxOleObjectMode::Linked);
    assert_eq!(linked.payload_kind(), Some(PptxOlePayloadKind::OleObject));
    assert!(matches!(
        linked.target(),
        Some(PptxOleObjectTarget::External {
            target,
            relationship_type,
        }) if target == "https://example.invalid/linked-document"
            && relationship_type == OLE_OBJECT
    ));

    assert_eq!(
        package.presentation().unwrap().ole_objects().unwrap(),
        objects
    );
}

#[test]
fn package_inventory_rejects_missing_ole_relationships() {
    let mut package = package_with_local_ole_objects();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    package
        .opc_package_mut()
        .get_part_mut(&slide_name)
        .unwrap()
        .rels_mut()
        .remove("rIdOle");

    assert!(matches!(
        package.ole_objects(),
        Err(OoxmlError::InvalidRelationship(message)) if message.contains("rIdOle")
    ));
}

#[test]
fn package_inventory_rejects_wrong_ole_payload_content_type() {
    let mut package = package_with_local_ole_objects();
    let payload_name = PackURI::new("/ppt/embeddings/oleObject1.bin").unwrap();
    assert!(package.opc_package_mut().remove_part(&payload_name));
    package.opc_package_mut().add_part(Box::new(BlobPart::new(
        payload_name,
        OFC_PACKAGE.to_string(),
        b"inert package payload with an OLE relationship".to_vec(),
    )));

    assert!(matches!(
        package.ole_objects(),
        Err(OoxmlError::InvalidContentType { expected, got })
            if expected == OFC_OLE_OBJECT && got == OFC_PACKAGE
    ));
}

#[test]
fn package_inventory_ignores_non_ole_graphic_data() {
    let mut package = package_with_local_ole_objects();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.opc_package_mut().get_part_mut(&slide_name).unwrap();
    let xml = std::str::from_utf8(slide.blob()).unwrap();
    let updated = xml.replace(
        "http://schemas.openxmlformats.org/presentationml/2006/ole",
        "urn:example:not-ole",
    );
    assert_ne!(updated, xml);
    slide.set_blob(updated.into_bytes());

    assert!(package.ole_objects().unwrap().is_empty());
}

fn package_with_local_ole_objects() -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    install_local_ole_objects(&mut package);
    package
}

fn install_local_ole_objects(package: &mut Package) {
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    {
        let slide = package.opc_package_mut().get_part_mut(&slide_name).unwrap();
        let xml = std::str::from_utf8(slide.blob()).unwrap();
        let updated = xml.replacen(
            "</p:spTree>",
            &format!(
                "{}{}",
                std::str::from_utf8(LOCAL_OLE_OBJECTS).unwrap(),
                "</p:spTree>"
            ),
            1,
        );
        assert_ne!(updated, xml);
        slide.set_blob(updated.into_bytes());
        slide.rels_mut().add_relationship(
            OLE_OBJECT.to_string(),
            "../embeddings/oleObject1.bin".to_string(),
            "rIdOle".to_string(),
            false,
        );
        slide.rels_mut().add_relationship(
            PACKAGE.to_string(),
            "../embeddings/package1.bin".to_string(),
            "rIdPackage".to_string(),
            false,
        );
        slide.rels_mut().add_relationship(
            OLE_OBJECT.to_string(),
            "https://example.invalid/linked-document".to_string(),
            "rIdLinked".to_string(),
            true,
        );
    }

    package.opc_package_mut().add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/embeddings/oleObject1.bin").unwrap(),
        OFC_OLE_OBJECT.to_string(),
        b"inert OLE payload".to_vec(),
    )));
    package.opc_package_mut().add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/embeddings/package1.bin").unwrap(),
        OFC_PACKAGE.to_string(),
        b"inert package payload".to_vec(),
    )));
}
