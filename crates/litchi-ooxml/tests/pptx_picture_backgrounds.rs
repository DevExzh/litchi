use litchi_ooxml::OoxmlError;
use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::{ImageFormat, Package, PictureStyle, SlideBackground};
use litchi_opc::constants::relationship_type as rt;
use tempfile::NamedTempFile;

const LOCAL_PNG: &[u8] = include_bytes!("../../../test-data/images/png/lena.png");

#[test]
fn picture_background_round_trips_with_its_image_resource() {
    let package = package_with_picture_background();
    let expected = expected_background();

    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    assert_eq!(slides[0].background().unwrap(), Some(expected.clone()));
    assert_eq!(slides[0].effective_background().unwrap(), Some(expected));
}

#[test]
fn picture_background_rejects_external_image_relationships() {
    let mut package = package_with_picture_background();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let relationship_id = package
        .opc_package()
        .get_part(&slide_name)
        .unwrap()
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::IMAGE)
        .unwrap()
        .r_id()
        .to_string();
    let slide = package.opc_package_mut().get_part_mut(&slide_name).unwrap();
    slide.rels_mut().remove(&relationship_id);
    slide.rels_mut().add_relationship(
        rt::IMAGE.to_string(),
        "https://example.invalid/background.png".to_string(),
        relationship_id,
        true,
    );

    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    assert!(matches!(
        slides[0].background(),
        Err(OoxmlError::InvalidRelationship(message)) if message.contains("must be internal")
    ));
}

fn expected_background() -> SlideBackground {
    SlideBackground::Picture {
        image_data: LOCAL_PNG.to_vec(),
        format: ImageFormat::Png,
        style: PictureStyle::Tile,
    }
}

fn package_with_picture_background() -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package
        .presentation_mut()
        .unwrap()
        .add_slide()
        .unwrap()
        .set_background(expected_background());
    package.save(output.path()).unwrap();
    Package::open(output.path()).unwrap()
}
