use litchi_ooxml::{OpcPackage, PackURI};
use litchi_ooxml_common::ribbon::{self, Family, Version};
use litchi_opc::{PackageWriter, XmlPart};

const OFFICE_2007_XML: &[u8] = include_bytes!("../../../test-data/ooxml/ribbonx/office2007.xml");
const OFFICE_2010_XML: &[u8] = include_bytes!("../../../test-data/ooxml/ribbonx/office2010.xml");
const CUSTOM_UI2_XML: &[u8] = include_bytes!("../../../test-data/ooxml/ribbonx/customui2.xml");

#[test]
fn stores_borrows_and_roundtrips_local_ribbon_xml() {
    let mut package = OpcPackage::new();
    ribbon::put(&mut package, Version::V2007, OFFICE_2007_XML.to_vec()).unwrap();

    let ribbons = ribbon::load(&package).unwrap();
    let stored = ribbons.effective().unwrap();
    assert_eq!(stored.part().as_str(), "/customUI/customUI.xml");
    assert_eq!(stored.version(), Version::V2007);
    assert_eq!(stored.xml(), OFFICE_2007_XML);
    assert_eq!(
        stored.xml().as_ptr(),
        package.get_part(stored.part()).unwrap().blob().as_ptr()
    );

    let bytes = PackageWriter::to_bytes(&package).unwrap();
    let reopened = OpcPackage::from_bytes(&bytes).unwrap();
    let reopened = ribbon::load(&reopened).unwrap();
    let reopened = reopened.effective().unwrap();
    assert_eq!(reopened.part().as_str(), "/customUI/customUI.xml");
    assert_eq!(reopened.version(), Version::V2007);
    assert_eq!(reopened.xml(), OFFICE_2007_XML);
}

#[test]
fn updates_the_modern_family_in_place() {
    let mut package = OpcPackage::new();
    ribbon::put(&mut package, Version::Ui2, CUSTOM_UI2_XML.to_vec()).unwrap();
    let first_part = ribbon::load(&package)
        .unwrap()
        .modern()
        .unwrap()
        .part()
        .clone();

    ribbon::put(&mut package, Version::V2010, OFFICE_2010_XML.to_vec()).unwrap();

    let updated = ribbon::load(&package).unwrap();
    let updated = updated.modern().unwrap();
    assert_eq!(updated.part(), &first_part);
    assert_eq!(updated.version(), Version::V2010);
    assert_eq!(updated.xml(), OFFICE_2010_XML);
    assert_eq!(
        package.rels().get(updated.id()).unwrap().reltype(),
        Version::V2010.relationship()
    );
}

#[test]
fn keeps_fixed_family_slots_with_modern_precedence() {
    let mut package = OpcPackage::new();
    ribbon::put(&mut package, Version::V2007, OFFICE_2007_XML.to_vec()).unwrap();
    ribbon::put(&mut package, Version::V2010, OFFICE_2010_XML.to_vec()).unwrap();

    let ribbons = ribbon::load(&package).unwrap();
    assert_eq!(ribbons.legacy().unwrap().version(), Version::V2007);
    assert_eq!(ribbons.modern().unwrap().version(), Version::V2010);
    assert_eq!(ribbons.effective(), ribbons.modern());
    assert_eq!(ribbons.iter().count(), 2);
}

#[test]
fn rejects_a_root_that_does_not_match_its_relationship_family() {
    let mut package = OpcPackage::new();
    let part_name = PackURI::new("/customUI/customUI.xml").unwrap();
    package.add_part(Box::new(XmlPart::new(
        part_name,
        "application/xml".to_string(),
        OFFICE_2010_XML.to_vec(),
    )));
    package.relate_to("customUI/customUI.xml", Version::V2007.relationship());

    assert!(ribbon::load(&package).is_err());
}

#[test]
fn document_facades_offer_concise_owned_crud() {
    let mut docx = litchi_ooxml::docx::Package::new().unwrap();
    docx.put_ribbon(Version::V2007, OFFICE_2007_XML.to_vec())
        .unwrap();
    assert_eq!(
        docx.ribbon().unwrap().effective().unwrap().xml(),
        OFFICE_2007_XML
    );
    assert!(docx.remove_ribbon(Family::Legacy).unwrap());
    assert!(docx.ribbon().unwrap().effective().is_none());

    let mut xlsx = litchi_ooxml::xlsx::Workbook::create().unwrap();
    xlsx.put_ribbon(Version::V2010, OFFICE_2010_XML.to_vec())
        .unwrap();
    assert_eq!(
        xlsx.ribbon().unwrap().effective().unwrap().version(),
        Version::V2010
    );

    let mut pptx = litchi_ooxml::pptx::Package::new().unwrap();
    pptx.put_ribbon(Version::V2010, OFFICE_2010_XML.to_vec())
        .unwrap();
    assert_eq!(
        pptx.ribbon().unwrap().effective().unwrap().xml(),
        OFFICE_2010_XML
    );
}
