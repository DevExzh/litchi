#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_ooxml_common::Error as CommonError;
use litchi_ooxml_common::ribbon::{Family, Version, load, put, remove};
use litchi_opc::OpcPackage;

const OFFICE_2007_XML: &[u8] = include_bytes!("../../../test-data/ooxml/ribbonx/office2007.xml");
const OFFICE_2010_XML: &[u8] = include_bytes!("../../../test-data/ooxml/ribbonx/office2010.xml");

#[test]
fn ribbon_storage_exposes_fixed_slots_and_effective_precedence() {
    // Ribbon customization is format-neutral OPC graph logic, so this test
    // follows its common owner directly instead of routing through a PPTX
    // package alias.
    let mut package = OpcPackage::new();
    put(&mut package, Version::V2007, OFFICE_2007_XML.to_vec()).unwrap();

    let modern_xml = OFFICE_2010_XML.to_vec();
    let moved_allocation = modern_xml.as_ptr();
    put(&mut package, Version::V2010, modern_xml).unwrap();

    let ribbons = load(&package).unwrap();
    let legacy = ribbons.legacy().unwrap();
    let modern = ribbons.modern().unwrap();

    assert_eq!(legacy.version(), Version::V2007);
    assert_eq!(legacy.xml(), OFFICE_2007_XML);
    assert_eq!(modern.version(), Version::V2010);
    assert_eq!(modern.xml(), OFFICE_2010_XML);
    assert_eq!(modern.xml().as_ptr(), moved_allocation);
    assert_eq!(ribbons.effective(), Some(modern));
    assert_eq!(ribbons.iter().count(), 2);
}

#[test]
fn common_owner_removes_ribbon_families_independently() {
    let mut package = OpcPackage::new();
    put(&mut package, Version::V2007, OFFICE_2007_XML.to_vec()).unwrap();
    put(&mut package, Version::V2010, OFFICE_2010_XML.to_vec()).unwrap();

    assert!(remove(&mut package, Family::Modern).unwrap());
    let ribbons = load(&package).unwrap();
    assert!(ribbons.modern().is_none());
    assert_eq!(ribbons.effective(), ribbons.legacy());

    assert!(remove(&mut package, Family::Legacy).unwrap());
    assert!(load(&package).unwrap().effective().is_none());
    assert!(!remove(&mut package, Family::Legacy).unwrap());
}

#[test]
fn common_ribbon_loader_rejects_external_relationships() {
    let mut package = OpcPackage::new();
    put(&mut package, Version::V2010, OFFICE_2010_XML.to_vec()).unwrap();
    let relationship_id = load(&package).unwrap().modern().unwrap().id().to_owned();

    let relationships = package.relationships_mut();
    relationships.remove(&relationship_id);
    relationships.add_relationship(
        Version::V2010.relationship().to_owned(),
        "https://example.invalid/customUI.xml".to_owned(),
        relationship_id,
        true,
    );

    assert!(matches!(
        load(&package),
        Err(CommonError::Relationship(message)) if message.contains("internal")
    ));
}
