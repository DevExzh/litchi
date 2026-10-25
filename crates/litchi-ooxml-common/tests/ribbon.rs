use litchi_ooxml_common::ribbon::{self, Family, Version};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, XmlPart};

const OFFICE_2007_XML: &[u8] = include_bytes!("../../../test-data/ooxml/ribbonx/office2007.xml");
const OFFICE_2010_XML: &[u8] = include_bytes!("../../../test-data/ooxml/ribbonx/office2010.xml");
const CUSTOM_UI2_XML: &[u8] = include_bytes!("../../../test-data/ooxml/ribbonx/customui2.xml");

fn put(package: &mut OpcPackage, version: Version, xml: Vec<u8>) {
    let mut edit = ribbon::Edit::new(package).unwrap();
    edit.put(version, xml).unwrap();
    *package = edit.commit().unwrap().into_package();
}

#[test]
fn stores_borrows_and_roundtrips_local_ribbon_xml() {
    let mut package = OpcPackage::new();
    put(&mut package, Version::V2007, OFFICE_2007_XML.to_vec());

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
    put(&mut package, Version::Ui2, CUSTOM_UI2_XML.to_vec());
    let first_part = ribbon::load(&package, &ribbon::Limits::standard())
        .unwrap()
        .modern()
        .unwrap()
        .part()
        .clone();

    put(&mut package, Version::V2010, OFFICE_2010_XML.to_vec());

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
    put(&mut package, Version::V2007, OFFICE_2007_XML.to_vec());
    put(&mut package, Version::V2010, OFFICE_2010_XML.to_vec());

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
fn ribbon_patch_checks_root_content_types_and_opaque_closure() {
    let mut source = OpcPackage::new();
    put(&mut source, Version::V2007, OFFICE_2007_XML.to_vec());
    let ribbon_part = ribbon::load(&source, &ribbon::Limits::standard())
        .unwrap()
        .legacy()
        .unwrap()
        .part()
        .clone();
    let image = PackURI::new("/customUI/icon.png").unwrap();
    source.add_part(Box::new(BlobPart::new(
        image.clone(),
        "image/png".into(),
        vec![1, 2, 3],
    )));
    source
        .get_part_mut(&ribbon_part)
        .unwrap()
        .relate_to("icon.png", rt::IMAGE);

    let mut edit = ribbon::Edit::new(&source).unwrap();
    edit.put(
        Version::V2007,
        br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><ribbon/></customUI>"#
            .to_vec(),
    )
    .unwrap();
    let commit = edit.commit().unwrap();

    let independent = commit.patch().apply(&source).unwrap();
    assert_eq!(
        ribbon::load(independent.package(), &ribbon::Limits::standard()).unwrap(),
        ribbon::load(commit.package(), &ribbon::Limits::standard()).unwrap()
    );
    assert!(independent.package().get_part(&image).is_ok());

    let mut root_conflict = source.clone();
    root_conflict.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    assert!(commit.patch().apply(&root_conflict).is_err());

    let mut content_types_conflict = source.clone();
    content_types_conflict.add_part(Box::new(BlobPart::new(
        PackURI::new("/opaque.bin").unwrap(),
        "application/octet-stream".into(),
        vec![9],
    )));
    assert!(commit.patch().apply(&content_types_conflict).is_err());
}

#[test]
fn ribbon_remove_inverse_restores_unshared_binary_image() {
    let mut source = OpcPackage::new();
    put(&mut source, Version::V2007, OFFICE_2007_XML.to_vec());
    let ribbon_part = ribbon::load(&source, &ribbon::Limits::standard())
        .unwrap()
        .legacy()
        .unwrap()
        .part()
        .clone();
    let image = PackURI::new("/customUI/icon.png").unwrap();
    source.add_part(Box::new(BlobPart::new(
        image.clone(),
        "image/png".into(),
        vec![1, 2, 3, 4],
    )));
    source
        .get_part_mut(&ribbon_part)
        .unwrap()
        .relate_to("icon.png", rt::IMAGE);

    let mut edit = ribbon::Edit::new(&source).unwrap();
    assert!(edit.remove(Family::Legacy).unwrap());
    let removed = edit.commit().unwrap();
    assert!(removed.package().get_part(&image).is_err());

    let restored = removed.patch().inverse().apply(removed.package()).unwrap();
    assert_eq!(
        restored.package().get_part(&image).unwrap().blob(),
        &[1, 2, 3, 4]
    );
    assert_eq!(
        ribbon::load(restored.package(), &ribbon::Limits::standard())
            .unwrap()
            .legacy()
            .unwrap()
            .xml(),
        OFFICE_2007_XML
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
