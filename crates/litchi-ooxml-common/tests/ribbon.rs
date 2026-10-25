use litchi_ooxml_common::ribbon::{self, Version};
use litchi_opc::{OpcPackage, PackURI, PackageWriter, XmlPart};

const OFFICE_2007_XML: &[u8] = include_bytes!("../../../test-data/ooxml/ribbonx/office2007.xml");
const OFFICE_2010_XML: &[u8] = include_bytes!("../../../test-data/ooxml/ribbonx/office2010.xml");
const CUSTOM_UI2_XML: &[u8] = include_bytes!("../../../test-data/ooxml/ribbonx/customui2.xml");

#[test]
fn stores_borrows_and_roundtrips_local_ribbon_xml() {
    let mut package = OpcPackage::new();
    package = ribbon::put(
        &package,
        Version::V2007,
        OFFICE_2007_XML.to_vec(),
        &ribbon::Limits::standard(),
    )
    .unwrap()
    .into_package();

    let ribbons = ribbon::load(&package, &ribbon::Limits::standard()).unwrap();
    let stored = ribbons.effective().unwrap();
    assert_eq!(stored.part().as_str(), "/customUI/customUI.xml");
    assert_eq!(stored.version(), Version::V2007);
    assert_eq!(stored.xml(), OFFICE_2007_XML);
    assert_eq!(
        stored.xml().as_ptr(),
        package.get_part(stored.part()).unwrap().blob().as_ptr()
    );

    let bytes = PackageWriter::to_bytes(&package).unwrap();
    let reopened = OpcPackage::from_bytes(&bytes, litchi_opc::SourceLimits::default()).unwrap();
    let reopened = ribbon::load(&reopened, &ribbon::Limits::standard()).unwrap();
    let reopened = reopened.effective().unwrap();
    assert_eq!(reopened.part().as_str(), "/customUI/customUI.xml");
    assert_eq!(reopened.version(), Version::V2007);
    assert_eq!(reopened.xml(), OFFICE_2007_XML);
}

#[test]
fn updates_the_modern_family_in_place() {
    let mut package = OpcPackage::new();
    package = ribbon::put(
        &package,
        Version::Ui2,
        CUSTOM_UI2_XML.to_vec(),
        &ribbon::Limits::standard(),
    )
    .unwrap()
    .into_package();
    let first_part = ribbon::load(&package, &ribbon::Limits::standard())
        .unwrap()
        .modern()
        .unwrap()
        .part()
        .clone();

    package = ribbon::put(
        &package,
        Version::V2010,
        OFFICE_2010_XML.to_vec(),
        &ribbon::Limits::standard(),
    )
    .unwrap()
    .into_package();

    let updated = ribbon::load(&package, &ribbon::Limits::standard()).unwrap();
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
    package = ribbon::put(
        &package,
        Version::V2007,
        OFFICE_2007_XML.to_vec(),
        &ribbon::Limits::standard(),
    )
    .unwrap()
    .into_package();
    package = ribbon::put(
        &package,
        Version::V2010,
        OFFICE_2010_XML.to_vec(),
        &ribbon::Limits::standard(),
    )
    .unwrap()
    .into_package();

    let ribbons = ribbon::load(&package, &ribbon::Limits::standard()).unwrap();
    assert_eq!(ribbons.legacy().unwrap().version(), Version::V2007);
    assert_eq!(ribbons.modern().unwrap().version(), Version::V2010);
    assert_eq!(ribbons.effective(), ribbons.modern());
    assert_eq!(ribbons.iter().count(), 2);
}

#[test]
fn staged_edit_publishes_both_families_and_replays_without_mutating_source() {
    let source = OpcPackage::new();
    let mut edit = ribbon::Edit::new(&source).unwrap();
    assert!(edit.put(Version::V2007, OFFICE_2007_XML.to_vec(),).unwrap());
    assert!(edit.put(Version::V2010, OFFICE_2010_XML.to_vec(),).unwrap());

    let commit = edit.commit().unwrap();
    assert!(
        ribbon::load(&source, &ribbon::Limits::standard())
            .unwrap()
            .iter()
            .next()
            .is_none()
    );
    let published = commit.package().clone();
    let replayed = commit.patch().apply(&source).unwrap();
    assert_eq!(
        ribbon::load(&published, &ribbon::Limits::standard()).unwrap(),
        ribbon::load(replayed.package(), &ribbon::Limits::standard()).unwrap()
    );
    let restored = commit.patch().inverse().apply(&published).unwrap();
    assert!(
        ribbon::load(restored.package(), &ribbon::Limits::standard())
            .unwrap()
            .iter()
            .next()
            .is_none()
    );
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

    assert!(ribbon::load(&package, &ribbon::Limits::standard()).is_err());
}
