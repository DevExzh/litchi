use litchi_ooxml::OoxmlError;
use litchi_ooxml::pptx::Package;
use litchi_ooxml::ribbonx::{RIBBONX_2010_RELATIONSHIP_TYPE, RibbonCustomizationVersion};

const RIBBON_XML: &[u8] = include_bytes!("../../../test-data/ooxml/ribbonx/office2010.xml");

#[test]
fn presentation_loads_ribbon_customizations() {
    let mut package = Package::new().unwrap();
    package
        .set_ribbon_customization(RibbonCustomizationVersion::Office2010, RIBBON_XML)
        .unwrap();

    let presentation = package.presentation().unwrap();
    let customizations = presentation.ribbon_customizations().unwrap();
    assert_eq!(customizations.len(), 1);
    assert_eq!(
        customizations[0].version(),
        RibbonCustomizationVersion::Office2010
    );
    assert_eq!(customizations[0].xml(), RIBBON_XML);
    assert_eq!(
        presentation.ribbon_customization().unwrap().unwrap(),
        customizations[0]
    );
}

#[test]
fn presentation_rejects_external_ribbon_customization_relationships() {
    let mut package = Package::new().unwrap();
    let customization = package
        .set_ribbon_customization(RibbonCustomizationVersion::Office2010, RIBBON_XML)
        .unwrap();
    let relationships = package.opc_package_mut().relationships_mut();
    relationships.remove(customization.relationship_id());
    relationships.add_relationship(
        RIBBONX_2010_RELATIONSHIP_TYPE.to_string(),
        "https://example.invalid/customUI.xml".to_string(),
        customization.relationship_id().to_string(),
        true,
    );

    assert!(matches!(
        package.presentation().unwrap().ribbon_customizations(),
        Err(OoxmlError::InvalidFormat(message)) if message.contains("must be internal")
    ));
}
