use litchi_ooxml::OoxmlError;
use litchi_ooxml::pptx::Package;
use litchi_ooxml_common::Error as CommonError;
use litchi_ooxml_common::ribbon::{Family, Version};

const OFFICE_2007_XML: &[u8] = include_bytes!("../../../test-data/ooxml/ribbonx/office2007.xml");
const OFFICE_2010_XML: &[u8] = include_bytes!("../../../test-data/ooxml/ribbonx/office2010.xml");

#[test]
fn presentation_exposes_fixed_slots_and_effective_precedence() {
    let mut package = Package::new().unwrap();
    package
        .put_ribbon(Version::V2007, OFFICE_2007_XML.to_vec())
        .unwrap();

    let modern_xml = OFFICE_2010_XML.to_vec();
    let moved_allocation = modern_xml.as_ptr();
    package.put_ribbon(Version::V2010, modern_xml).unwrap();

    let presentation = package.presentation().unwrap();
    let ribbons = presentation.ribbon().unwrap();
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
fn package_removes_ribbon_families_independently() {
    let mut package = Package::new().unwrap();
    package
        .put_ribbon(Version::V2007, OFFICE_2007_XML.to_vec())
        .unwrap()
        .put_ribbon(Version::V2010, OFFICE_2010_XML.to_vec())
        .unwrap();

    assert!(package.remove_ribbon(Family::Modern).unwrap());
    let ribbons = package.ribbon().unwrap();
    assert!(ribbons.modern().is_none());
    assert_eq!(ribbons.effective(), ribbons.legacy());

    assert!(package.remove_ribbon(Family::Legacy).unwrap());
    assert!(package.ribbon().unwrap().effective().is_none());
    assert!(!package.remove_ribbon(Family::Legacy).unwrap());
}

#[test]
fn presentation_rejects_external_ribbon_relationships() {
    let mut package = Package::new().unwrap();
    package
        .put_ribbon(Version::V2010, OFFICE_2010_XML.to_vec())
        .unwrap();
    let relationship_id = package.ribbon().unwrap().modern().unwrap().id().to_owned();

    package
        .edit_opc(|opc| {
            let relationships = opc.relationships_mut();
            relationships.remove(&relationship_id);
            relationships.add_relationship(
                Version::V2010.relationship().to_owned(),
                "https://example.invalid/customUI.xml".to_owned(),
                relationship_id,
                true,
            );
            Ok(())
        })
        .unwrap();

    assert!(matches!(
        package.presentation().unwrap().ribbon(),
        Err(OoxmlError::Common(CommonError::Relationship(message)))
            if message.contains("internal")
    ));
}
