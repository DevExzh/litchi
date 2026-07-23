use litchi_ooxml::ribbonx::{
    RIBBONX_2007_RELATIONSHIP_TYPE, RIBBONX_2010_RELATIONSHIP_TYPE, RIBBONX_CONTENT_TYPE,
    RibbonCustomizationVersion, load_ribbon_customization, load_ribbon_customizations,
    store_ribbon_customization,
};
use litchi_ooxml::{OpcPackage, PackURI};
use litchi_opc::{PackageWriter, XmlPart};

const OFFICE_2007_XML: &[u8] = include_bytes!("../../../test-data/ooxml/ribbonx/office2007.xml");
const OFFICE_2010_XML: &[u8] = include_bytes!("../../../test-data/ooxml/ribbonx/office2010.xml");
const CUSTOM_UI2_XML: &[u8] = include_bytes!("../../../test-data/ooxml/ribbonx/customui2.xml");

#[test]
fn stores_and_loads_local_ribbonx_xml() {
    let mut package = OpcPackage::new();

    let stored = store_ribbon_customization(
        &mut package,
        RibbonCustomizationVersion::Office2007,
        OFFICE_2007_XML,
    )
    .unwrap();

    assert_eq!(stored.part_name().as_str(), "/customUI/customUI.xml");
    assert_eq!(stored.version(), RibbonCustomizationVersion::Office2007);
    assert_eq!(stored.xml(), OFFICE_2007_XML);

    let loaded = load_ribbon_customization(&package).unwrap().unwrap();
    assert_eq!(loaded, stored);

    let bytes = PackageWriter::to_bytes(&package).unwrap();
    let reopened = OpcPackage::from_bytes(&bytes).unwrap();
    assert_eq!(
        load_ribbon_customization(&reopened).unwrap().unwrap(),
        stored
    );
}

#[test]
fn updates_the_newer_relationship_family_in_place() {
    let mut package = OpcPackage::new();
    let first = store_ribbon_customization(
        &mut package,
        RibbonCustomizationVersion::CustomUi2,
        CUSTOM_UI2_XML,
    )
    .unwrap();
    assert_eq!(load_ribbon_customization(&package).unwrap().unwrap(), first);

    let updated = store_ribbon_customization(
        &mut package,
        RibbonCustomizationVersion::Office2010,
        OFFICE_2010_XML,
    )
    .unwrap();

    assert_eq!(updated.part_name(), first.part_name());
    assert_eq!(updated.version(), RibbonCustomizationVersion::Office2010);
    assert_eq!(updated.xml(), OFFICE_2010_XML);
    assert_eq!(
        package
            .rels()
            .get(updated.relationship_id())
            .unwrap()
            .reltype(),
        RIBBONX_2010_RELATIONSHIP_TYPE
    );
}

#[test]
fn retains_legacy_and_newer_customizations_with_newer_precedence() {
    let mut package = OpcPackage::new();
    let legacy = store_ribbon_customization(
        &mut package,
        RibbonCustomizationVersion::Office2007,
        OFFICE_2007_XML,
    )
    .unwrap();
    let newer = store_ribbon_customization(
        &mut package,
        RibbonCustomizationVersion::Office2010,
        OFFICE_2010_XML,
    )
    .unwrap();

    assert_eq!(
        load_ribbon_customizations(&package).unwrap(),
        vec![legacy, newer.clone()]
    );
    assert_eq!(load_ribbon_customization(&package).unwrap().unwrap(), newer);
}

#[test]
fn rejects_a_customization_root_that_does_not_match_its_relationship_version() {
    let mut package = OpcPackage::new();
    let part_name = PackURI::new("/customUI/customUI.xml").unwrap();
    package.add_part(Box::new(XmlPart::new(
        part_name,
        RIBBONX_CONTENT_TYPE.to_string(),
        OFFICE_2010_XML.to_vec(),
    )));
    package.relate_to("customUI/customUI.xml", RIBBONX_2007_RELATIONSHIP_TYPE);

    assert!(load_ribbon_customization(&package).is_err());
}

#[test]
fn rejects_external_or_duplicate_customization_relationships() {
    let mut external = OpcPackage::new();
    external.relate_to_external(
        "https://example.invalid/customUI.xml",
        RIBBONX_2007_RELATIONSHIP_TYPE,
    );
    assert!(load_ribbon_customization(&external).is_err());

    let mut duplicate = OpcPackage::new();
    for (name, xml) in [
        ("/customUI/customUI2.xml", CUSTOM_UI2_XML),
        ("/customUI/customUI14.xml", OFFICE_2010_XML),
    ] {
        duplicate.add_part(Box::new(XmlPart::new(
            PackURI::new(name).unwrap(),
            RIBBONX_CONTENT_TYPE.to_string(),
            xml.to_vec(),
        )));
    }
    duplicate.relate_to("customUI/customUI2.xml", RIBBONX_2010_RELATIONSHIP_TYPE);
    duplicate.relate_to("customUI/customUI14.xml", RIBBONX_2010_RELATIONSHIP_TYPE);
    assert!(load_ribbon_customization(&duplicate).is_err());
}

#[test]
fn document_package_wrappers_store_inert_ribbonx_xml() {
    let mut docx = litchi_ooxml::docx::Package::new().unwrap();
    let docx_customization = docx
        .set_ribbon_customization(RibbonCustomizationVersion::Office2007, OFFICE_2007_XML)
        .unwrap();
    assert_eq!(
        docx.ribbon_customizations().unwrap(),
        vec![docx_customization.clone()]
    );
    assert_eq!(
        docx.ribbon_customization().unwrap().unwrap(),
        docx_customization
    );

    let mut xlsx = litchi_ooxml::xlsx::Workbook::create().unwrap();
    let xlsx_customization = xlsx
        .set_ribbon_customization(RibbonCustomizationVersion::Office2010, OFFICE_2010_XML)
        .unwrap();
    assert_eq!(
        xlsx.ribbon_customization().unwrap().unwrap(),
        xlsx_customization
    );

    let mut pptx = litchi_ooxml::pptx::Package::new().unwrap();
    let pptx_customization = pptx
        .set_ribbon_customization(RibbonCustomizationVersion::Office2010, OFFICE_2010_XML)
        .unwrap();
    assert_eq!(
        pptx.ribbon_customization().unwrap().unwrap(),
        pptx_customization
    );
}
