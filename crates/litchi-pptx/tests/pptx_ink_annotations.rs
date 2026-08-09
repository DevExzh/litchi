#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::XmlPart;
use litchi_opc::constants::relationship_type::CUSTOM_XML;
use litchi_opc::{OpcError, OpcPackage, PackURI};
use litchi_pptx::presentation::embedded::ink::{self, CONTENT_TYPE};
use litchi_pptx::{Error, Package};
use tempfile::NamedTempFile;

// The negative cases republish this positive InkML fixture before corrupting
// it, so keep it inside the production writer's compact XML contract.
const LOCAL_INK: &[u8] = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"><inkml:traceGroup>"#,
    r#"<inkml:trace>0 0, 10 10</inkml:trace>"#,
    r#"<inkml:trace>10 10, 20 20</inkml:trace>"#,
    r#"</inkml:traceGroup></inkml:ink>"#,
)
.as_bytes();

#[test]
fn package_inventory_reports_local_ink_content_parts() {
    let package = package_with_ink();

    let inventory = annotations(&package);
    assert_eq!(inventory.len(), 1);

    let annotation = &inventory[0];
    assert_eq!(annotation.slide_index(), 0);
    assert_eq!(annotation.index(), 0);
    assert_eq!(annotation.relationship_id(), "rIdInk");
    assert_eq!(annotation.part_name().as_str(), "/ppt/ink/ink1.xml");
    assert_eq!(annotation.trace_count(), 2);
    assert_eq!(annotation.trace_group_count(), 1);

    assert_eq!(annotations(&package), inventory);
}

#[test]
fn package_inventory_rejects_missing_ink_targets() {
    let mut package = package_with_ink();
    let part_name = PackURI::new("/ppt/ink/ink1.xml").unwrap();
    package = edit_package(package, |opc| {
        assert!(opc.remove_part(&part_name));
    });

    let error = match annotations_result(&package) {
        Ok(_) => panic!("a removed InkML target must not be discovered"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::Opc(OpcError::PartNotFound(message)) if message.contains("/ppt/ink/ink1.xml")
    ));
}

#[test]
fn package_inventory_rejects_malformed_inkml() {
    let mut package = package_with_ink();
    let part_name = PackURI::new("/ppt/ink/ink1.xml").unwrap();
    package = edit_package(package, |opc| {
        opc.get_part_mut(&part_name)
            .unwrap()
            .set_blob(b"<ink/>".to_vec());
    });

    let error = match annotations_result(&package) {
        Ok(_) => panic!("malformed InkML must be rejected"),
        Err(error) => error,
    };
    assert!(matches!(error, Error::Invalid(_)));
}

fn package_with_ink() -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let package = Package::open(output.path()).unwrap();
    edit_package(package, install_local_ink)
}

fn install_local_ink(package: &mut OpcPackage) {
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.get_part_mut(&slide_name).unwrap();
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
    package.add_part(Box::new(XmlPart::new(
        PackURI::new("/ppt/ink/ink1.xml").unwrap(),
        CONTENT_TYPE.to_string(),
        LOCAL_INK.to_vec(),
    )));
}

fn annotations(package: &Package) -> Vec<ink::Annotation> {
    annotations_result(package).unwrap()
}

fn annotations_result(package: &Package) -> litchi_pptx::Result<Vec<ink::Annotation>> {
    let opc = package.opc()?;
    let slides = package.presentation()?.slides()?;
    let mut limits = ink::Limits::default();
    slides
        .iter()
        .enumerate()
        .map(|(index, slide)| ink::load_slide(opc, index, slide.part().part(), &mut limits))
        .collect::<litchi_pptx::Result<Vec<_>>>()
        .map(|groups| groups.into_iter().flatten().collect())
}

fn edit_package(mut package: Package, edit: impl FnOnce(&mut OpcPackage)) -> Package {
    let bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&bytes).unwrap();
    edit(&mut opc);
    Package::from_opc_package(opc).unwrap()
}
