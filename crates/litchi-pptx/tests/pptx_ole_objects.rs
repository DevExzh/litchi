use litchi_opc::constants::{
    content_type::{OFC_OLE_OBJECT, OFC_PACKAGE},
    relationship_type::{OLE_OBJECT, PACKAGE},
};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::presentation::embedded::ole::{self, Kind, Mode, Target};
use litchi_pptx::{Error, Package};
use tempfile::NamedTempFile;

const LOCAL_OLE_OBJECTS: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/ole/basic_ole_objects.xml");

#[test]
fn package_inventory_reports_local_ole_objects() {
    let package = package_with_local_ole_objects();

    let inventory = objects(&package);
    assert_eq!(inventory.len(), 3);

    let workbook = &inventory[0];
    assert_eq!(workbook.slide_index(), 0);
    assert_eq!(workbook.index(), 0);
    assert_eq!(workbook.shape_id(), Some(101));
    assert_eq!(workbook.shape_name(), Some("Embedded workbook"));
    assert_eq!(workbook.legacy_shape_id(), Some("_x0000_s1025"));
    assert_eq!(workbook.name(), Some("Workbook"));
    assert_eq!(workbook.program_id(), Some("Excel.Sheet.12"));
    assert_eq!(workbook.show_as_icon(), Some(true));
    assert_eq!(workbook.preview_width(), Some(914_400));
    assert_eq!(workbook.preview_height(), Some(457_200));
    assert_eq!(workbook.mode(), Mode::Embedded);
    assert_eq!(workbook.relationship_id(), Some("rIdOle"));
    assert_eq!(workbook.kind(), Some(Kind::OleObject));
    assert_eq!(workbook.preview_relationship_id(), Some("rIdPreview"));
    assert!(matches!(
        workbook.target(),
        Some(Target::Internal {
            part_name,
            content_type,
            relationship_type,
        }) if part_name.as_str() == "/ppt/embeddings/oleObject1.bin"
            && content_type == OFC_OLE_OBJECT
            && relationship_type == OLE_OBJECT
    ));

    let package_object = &inventory[1];
    assert_eq!(package_object.index(), 1);
    assert_eq!(package_object.mode(), Mode::Embedded);
    assert_eq!(
        package_object.kind(),
        Some(Kind::Package)
    );
    assert!(matches!(
        package_object.target(),
        Some(Target::Internal {
            part_name,
            content_type,
            relationship_type,
        }) if part_name.as_str() == "/ppt/embeddings/package1.bin"
            && content_type == OFC_PACKAGE
            && relationship_type == PACKAGE
    ));

    let linked = &inventory[2];
    assert_eq!(linked.index(), 2);
    assert_eq!(linked.mode(), Mode::Linked);
    assert_eq!(linked.kind(), Some(Kind::OleObject));
    assert!(matches!(
        linked.target(),
        Some(Target::External {
            target,
            relationship_type,
        }) if target == "https://example.invalid/linked-document"
            && relationship_type == OLE_OBJECT
    ));

    assert_eq!(
        objects(&package),
        inventory
    );
}

#[test]
fn package_inventory_rejects_missing_ole_relationships() {
    let mut package = package_with_local_ole_objects();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    package = edit_package(package, |opc| {
        opc.get_part_mut(&slide_name).unwrap().rels_mut().remove("rIdOle");
    });

    assert!(matches!(
        objects_result(&package),
        Err(Error::Relationship(message)) if message.contains("rIdOle")
    ));
}

#[test]
fn package_inventory_rejects_wrong_ole_payload_content_type() {
    let mut package = package_with_local_ole_objects();
    let payload_name = PackURI::new("/ppt/embeddings/oleObject1.bin").unwrap();
    package = edit_package(package, |opc| {
            assert!(opc.remove_part(&payload_name));
            opc.add_part(Box::new(BlobPart::new(
                payload_name,
                OFC_PACKAGE.to_string(),
                b"inert package payload with an OLE relationship".to_vec(),
            )));
    });

    assert!(matches!(
        objects_result(&package),
        Err(Error::ContentType { expected, actual })
            if expected == OFC_OLE_OBJECT && actual == OFC_PACKAGE
    ));
}

#[test]
fn package_inventory_ignores_non_ole_graphic_data() {
    let mut package = package_with_local_ole_objects();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    package = edit_package(package, |opc| {
            let slide = opc.get_part_mut(&slide_name).unwrap();
            let xml = std::str::from_utf8(slide.blob()).unwrap();
            let updated = xml.replace(
                "http://schemas.openxmlformats.org/presentationml/2006/ole",
                "urn:example:not-ole",
            );
            assert_ne!(updated, xml);
            slide.set_blob(updated.into_bytes());
    });

    assert!(objects(&package).is_empty());
}

fn package_with_local_ole_objects() -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let package = Package::open(output.path()).unwrap();
    edit_package(package, install_local_ole_objects)
}

fn install_local_ole_objects(package: &mut OpcPackage) {
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.get_part_mut(&slide_name).unwrap();
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
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/embeddings/oleObject1.bin").unwrap(),
        OFC_OLE_OBJECT.to_string(),
        b"inert OLE payload".to_vec(),
    )));
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/embeddings/package1.bin").unwrap(),
        OFC_PACKAGE.to_string(),
        b"inert package payload".to_vec(),
    )));
}

fn objects(package: &Package) -> Vec<ole::Object> {
    objects_result(package).unwrap()
}

fn objects_result(package: &Package) -> litchi_pptx::Result<Vec<ole::Object>> {
    let opc = package.opc()?;
    let slides = package.presentation()?.slides()?;
    let mut limits = ole::Limits::default();
    slides
        .iter()
        .enumerate()
        .map(|(index, slide)| ole::load_slide(opc, index, slide.part().part(), &mut limits))
        .collect::<litchi_pptx::Result<Vec<_>>>()
        .map(|groups| groups.into_iter().flatten().collect())
}

fn edit_package(mut package: Package, edit: impl FnOnce(&mut OpcPackage)) -> Package {
    let bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&bytes).unwrap();
    edit(&mut opc);
    Package::from_opc_package(opc).unwrap()
}
