use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::parts::PresentationPart;
use litchi_pptx::{Error, Package, Presentation};

const STRICT_SLIDE_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/slide";
const STRICT_SLIDE_MASTER_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/slideMaster";

#[test]
fn root_slide_and_master_relationships_reject_external_targets() {
    let mut package = package_with_slide();
    replace_presentation_relationship(
        &mut package,
        rt::SLIDE,
        rt::SLIDE,
        "https://example.invalid/slide.xml",
        true,
    );
    assert!(matches!(
        presentation(&package).slides(),
        Err(Error::Relationship(message)) if message.contains("must be internal")
    ));

    let mut package = package_without_slides();
    replace_presentation_relationship(
        &mut package,
        rt::SLIDE_MASTER,
        rt::SLIDE_MASTER,
        "https://example.invalid/slide-master.xml",
        true,
    );
    assert!(matches!(
        presentation(&package).slide_masters(),
        Err(Error::Relationship(message)) if message.contains("must be internal")
    ));
}

#[test]
fn root_slide_and_master_relationships_reject_wrong_types() {
    let mut package = package_with_slide();
    replace_presentation_relationship(
        &mut package,
        rt::SLIDE,
        rt::THEME,
        "slides/slide1.xml",
        false,
    );
    assert!(matches!(
        presentation(&package).slides(),
        Err(Error::Relationship(message)) if message.contains("unexpected type")
    ));

    let mut package = package_without_slides();
    replace_presentation_relationship(
        &mut package,
        rt::SLIDE_MASTER,
        rt::THEME,
        "slideMasters/slideMaster1.xml",
        false,
    );
    assert!(matches!(
        presentation(&package).slide_masters(),
        Err(Error::Relationship(message)) if message.contains("is not a slide-master relationship")
    ));
}

#[test]
fn root_slide_and_master_relationships_reject_wrong_content_types() {
    let mut package = package_with_slide();
    replace_presentation_relationship(
        &mut package,
        rt::SLIDE,
        rt::SLIDE,
        "slideMasters/slideMaster1.xml",
        false,
    );
    assert!(matches!(
        presentation(&package).slides(),
        Err(Error::ContentType { expected, actual })
            if expected == ct::PML_SLIDE && actual == ct::PML_SLIDE_MASTER
    ));

    let mut package = package_without_slides();
    replace_presentation_relationship(
        &mut package,
        rt::SLIDE_MASTER,
        rt::SLIDE_MASTER,
        "slideLayouts/slideLayout1.xml",
        false,
    );
    assert!(matches!(
        presentation(&package).slide_masters(),
        Err(Error::ContentType { expected, actual })
            if expected == ct::PML_SLIDE_MASTER && actual == ct::PML_SLIDE_LAYOUT
    ));
}

#[test]
fn strict_root_slide_and_master_relationships_are_supported() {
    let mut package = package_with_slide();
    let (slide_relationship_id, slide_target) = presentation_relationship(&package, rt::SLIDE);
    let (master_relationship_id, master_target) =
        presentation_relationship(&package, rt::SLIDE_MASTER);
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let presentation_part = package.get_part_mut(&presentation_name).unwrap();
    presentation_part.rels_mut().remove(&slide_relationship_id);
    presentation_part.rels_mut().add_relationship(
        STRICT_SLIDE_RELATIONSHIP_TYPE.to_owned(),
        slide_target,
        slide_relationship_id,
        false,
    );
    presentation_part.rels_mut().remove(&master_relationship_id);
    presentation_part.rels_mut().add_relationship(
        STRICT_SLIDE_MASTER_RELATIONSHIP_TYPE.to_owned(),
        master_target,
        master_relationship_id,
        false,
    );

    let presentation = presentation(&package);
    assert_eq!(presentation.slides().unwrap().len(), 1);
    assert_eq!(presentation.slide_masters().unwrap().len(), 1);
}

fn package_with_slide() -> OpcPackage {
    let mut authored = Package::new().unwrap();
    authored.presentation_mut().unwrap().add_slide().unwrap();
    OpcPackage::from_bytes(&authored.to_bytes().unwrap()).unwrap()
}

fn package_without_slides() -> OpcPackage {
    let mut authored = Package::new().unwrap();
    OpcPackage::from_bytes(&authored.to_bytes().unwrap()).unwrap()
}

fn presentation(package: &OpcPackage) -> Presentation<'_> {
    let part = PresentationPart::from_package(package).unwrap();
    Presentation::new(part, package)
}

fn replace_presentation_relationship(
    package: &mut OpcPackage,
    old_type: &str,
    new_type: &str,
    target: &str,
    external: bool,
) {
    let (relationship_id, _) = presentation_relationship(package, old_type);
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let presentation = package.get_part_mut(&presentation_name).unwrap();
    presentation.rels_mut().remove(&relationship_id);
    presentation.rels_mut().add_relationship(
        new_type.to_owned(),
        target.to_owned(),
        relationship_id,
        external,
    );
}

fn presentation_relationship(package: &OpcPackage, relationship_type: &str) -> (String, String) {
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    package
        .get_part(&presentation_name)
        .unwrap()
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == relationship_type)
        .map(|relationship| {
            (
                relationship.r_id().to_owned(),
                relationship.target_ref().to_owned(),
            )
        })
        .unwrap()
}
