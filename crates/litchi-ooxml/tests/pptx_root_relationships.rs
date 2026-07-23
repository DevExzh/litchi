use litchi_ooxml::OoxmlError;
use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::Package;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use tempfile::NamedTempFile;

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
    let presentation = package.presentation().unwrap();
    assert!(matches!(
        presentation.slides(),
        Err(OoxmlError::InvalidRelationship(message)) if message.contains("must be internal")
    ));

    let mut package = Package::new().unwrap();
    replace_presentation_relationship(
        &mut package,
        rt::SLIDE_MASTER,
        rt::SLIDE_MASTER,
        "https://example.invalid/slide-master.xml",
        true,
    );
    let presentation = package.presentation().unwrap();
    assert!(matches!(
        presentation.slide_masters(),
        Err(OoxmlError::InvalidRelationship(message)) if message.contains("must be internal")
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
    let presentation = package.presentation().unwrap();
    assert!(matches!(
        presentation.slides(),
        Err(OoxmlError::InvalidRelationship(message)) if message.contains("is not a slide relationship")
    ));

    let mut package = Package::new().unwrap();
    replace_presentation_relationship(
        &mut package,
        rt::SLIDE_MASTER,
        rt::THEME,
        "slideMasters/slideMaster1.xml",
        false,
    );
    let presentation = package.presentation().unwrap();
    assert!(matches!(
        presentation.slide_masters(),
        Err(OoxmlError::InvalidRelationship(message))
            if message.contains("is not a slide-master relationship")
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
    let presentation = package.presentation().unwrap();
    assert!(matches!(
        presentation.slides(),
        Err(OoxmlError::InvalidContentType { expected, got })
            if expected == ct::PML_SLIDE && got == ct::PML_SLIDE_MASTER
    ));

    let mut package = Package::new().unwrap();
    replace_presentation_relationship(
        &mut package,
        rt::SLIDE_MASTER,
        rt::SLIDE_MASTER,
        "slideLayouts/slideLayout1.xml",
        false,
    );
    let presentation = package.presentation().unwrap();
    assert!(matches!(
        presentation.slide_masters(),
        Err(OoxmlError::InvalidContentType { expected, got })
            if expected == ct::PML_SLIDE_MASTER && got == ct::PML_SLIDE_LAYOUT
    ));
}

#[test]
fn strict_root_slide_and_master_relationships_are_supported() {
    let mut package = package_with_slide();
    let (slide_relationship_id, slide_target) = presentation_relationship(&package, rt::SLIDE);
    let (master_relationship_id, master_target) =
        presentation_relationship(&package, rt::SLIDE_MASTER);
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let presentation = package
        .opc_package_mut()
        .get_part_mut(&presentation_name)
        .unwrap();
    presentation.rels_mut().remove(&slide_relationship_id);
    presentation.rels_mut().add_relationship(
        STRICT_SLIDE_RELATIONSHIP_TYPE.to_string(),
        slide_target,
        slide_relationship_id,
        false,
    );
    presentation.rels_mut().remove(&master_relationship_id);
    presentation.rels_mut().add_relationship(
        STRICT_SLIDE_MASTER_RELATIONSHIP_TYPE.to_string(),
        master_target,
        master_relationship_id,
        false,
    );

    let presentation = package.presentation().unwrap();
    assert_eq!(presentation.slides().unwrap().len(), 1);
    assert_eq!(presentation.slide_masters().unwrap().len(), 1);
}

fn package_with_slide() -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();
    Package::open(output.path()).unwrap()
}

fn replace_presentation_relationship(
    package: &mut Package,
    old_type: &str,
    new_type: &str,
    target: &str,
    external: bool,
) {
    let (relationship_id, _) = presentation_relationship(package, old_type);
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let presentation = package
        .opc_package_mut()
        .get_part_mut(&presentation_name)
        .unwrap();
    presentation.rels_mut().remove(&relationship_id);
    presentation.rels_mut().add_relationship(
        new_type.to_string(),
        target.to_string(),
        relationship_id,
        external,
    );
}

fn presentation_relationship(package: &Package, relationship_type: &str) -> (String, String) {
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    package
        .opc_package()
        .get_part(&presentation_name)
        .unwrap()
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == relationship_type)
        .map(|relationship| {
            (
                relationship.r_id().to_string(),
                relationship.target_ref().to_string(),
            )
        })
        .unwrap()
}
