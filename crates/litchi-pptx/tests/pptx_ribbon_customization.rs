#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_ooxml_common::Error as CommonError;
use litchi_ooxml_common::ribbon::{Edit, EditOptions, Family, Limits, Version, load};
use litchi_opc::OpcPackage;

const OFFICE_2007_XML: &[u8] = include_bytes!("../../../test-data/ooxml/ribbonx/office2007.xml");
const OFFICE_2010_XML: &[u8] = include_bytes!("../../../test-data/ooxml/ribbonx/office2010.xml");

fn put(
    package: &mut OpcPackage,
    version: Version,
    xml: Vec<u8>,
    limits: &Limits,
) -> litchi_ooxml_common::Result<()> {
    let mut edit = Edit::new_with_options(package, EditOptions { limits: *limits })?;
    edit.put(version, xml)?;
    *package = edit.commit()?.into_package();
    Ok(())
}

fn remove(package: &mut OpcPackage, family: Family) -> litchi_ooxml_common::Result<bool> {
    let mut edit = Edit::new(package)?;
    let changed = edit.remove(family)?;
    *package = edit.commit()?.into_package();
    Ok(changed)
}

#[test]
fn ribbon_storage_exposes_fixed_slots_and_effective_precedence() {
    // Ribbon customization is format-neutral OPC graph logic, so this test
    // follows its common owner directly instead of routing through a PPTX
    // package alias.
    let mut package = OpcPackage::new();
    put(
        &mut package,
        Version::V2007,
        OFFICE_2007_XML.to_vec(),
        &Limits::standard(),
    )
    .unwrap();

    let modern_xml = OFFICE_2010_XML.to_vec();
    let moved_allocation = modern_xml.as_ptr();
    put(
        &mut package,
        Version::V2010,
        modern_xml,
        &Limits::standard(),
    )
    .unwrap();

    let ribbons = load(&package, &Limits::standard()).unwrap();
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
    put(
        &mut package,
        Version::V2007,
        OFFICE_2007_XML.to_vec(),
        &Limits::standard(),
    )
    .unwrap();
    put(
        &mut package,
        Version::V2010,
        OFFICE_2010_XML.to_vec(),
        &Limits::standard(),
    )
    .unwrap();

    assert!(remove(&mut package, Family::Modern).unwrap());
    let ribbons = load(&package, &Limits::standard()).unwrap();
    assert!(ribbons.modern().is_none());
    assert_eq!(ribbons.effective(), ribbons.legacy());

    assert!(remove(&mut package, Family::Legacy).unwrap());
    assert!(
        load(&package, &Limits::standard())
            .unwrap()
            .effective()
            .is_none()
    );
    assert!(!remove(&mut package, Family::Legacy).unwrap());
}

#[test]
fn common_ribbon_loader_rejects_external_relationships() {
    let mut package = OpcPackage::new();
    put(
        &mut package,
        Version::V2010,
        OFFICE_2010_XML.to_vec(),
        &Limits::standard(),
    )
    .unwrap();
    let relationship_id = load(&package, &Limits::standard())
        .unwrap()
        .modern()
        .unwrap()
        .id()
        .to_owned();

    let relationships = package.relationships_mut();
    relationships.remove(&relationship_id);
    relationships.add_relationship(
        Version::V2010.relationship().to_owned(),
        "https://example.invalid/customUI.xml".to_owned(),
        relationship_id,
        true,
    );

    assert!(matches!(
        load(&package, &Limits::standard()),
        Err(CommonError::Relationship(message)) if message.contains("internal")
    ));
}
