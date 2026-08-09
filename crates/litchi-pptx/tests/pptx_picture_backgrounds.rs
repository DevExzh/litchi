#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::PackURI;
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::part::BlobPart;
use litchi_pptx::{ImageFormat, Package, PictureStyle, SlideBackground};
use tempfile::NamedTempFile;

const LOCAL_PNG: &[u8] = include_bytes!("../../../test-data/images/png/lena.png");

#[test]
fn picture_background_serializes_with_its_image_relationship_slot() {
    let package = package_with_picture_background();
    let expected = expected_background();
    let fragment = expected.to_xml(Some("rIdBackground")).unwrap();
    assert!(fragment.contains(r#"<a:blip r:embed="rIdBackground"/>"#));
    assert!(fragment.contains("<a:tile/>"));

    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.opc().unwrap().get_part(&slide_name).unwrap();
    let xml = std::str::from_utf8(slide.blob()).unwrap();
    assert!(xml.contains(r#"<a:blip r:embed="rIdBackground"/>"#));
    assert!(xml.contains("<a:tile/>"));
    let relationship = slide
        .rels()
        .get("rIdBackground")
        .expect("picture background relationship");
    assert_eq!(relationship.reltype(), rt::IMAGE);
    assert_eq!(relationship.target_ref(), "../media/background.png");
    assert_eq!(
        package
            .opc()
            .unwrap()
            .get_part(&PackURI::new("/ppt/media/background.png").unwrap())
            .unwrap()
            .blob(),
        LOCAL_PNG
    );
}

#[test]
fn picture_background_rejects_external_image_relationships() {
    let mut package = package_with_picture_background();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let relationship_id = package
        .opc()
        .unwrap()
        .get_part(&slide_name)
        .unwrap()
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::IMAGE)
        .unwrap()
        .r_id()
        .to_string();
    package
        .edit_opc(|opc| {
            let slide = opc.get_part_mut(&slide_name)?;
            slide.rels_mut().remove(&relationship_id);
            slide.rels_mut().add_relationship(
                rt::IMAGE.to_string(),
                "https://example.invalid/background.png".to_string(),
                relationship_id.clone(),
                true,
            );
            Ok(())
        })
        .unwrap();

    let relationship = package
        .opc()
        .unwrap()
        .get_part(&slide_name)
        .unwrap()
        .rels()
        .get(&relationship_id)
        .unwrap();
    assert!(relationship.is_external());
    assert_eq!(
        relationship.target_ref(),
        "https://example.invalid/background.png"
    );
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
    let mut package = Package::open(output.path()).unwrap();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let image_name = PackURI::new("/ppt/media/background.png").unwrap();
    package
        .edit_opc(|opc| {
            let slide = opc.get_part_mut(&slide_name)?;
            let xml = std::str::from_utf8(slide.blob()).unwrap();
            let updated = xml.replace("<a:blip/>", r#"<a:blip r:embed="rIdBackground"/>"#);
            assert_ne!(updated, xml, "generated slide must contain a picture fill");
            slide.set_blob(updated.into_bytes());
            slide.rels_mut().add_relationship(
                rt::IMAGE.to_string(),
                "../media/background.png".to_string(),
                "rIdBackground".to_string(),
                false,
            );
            opc.add_part(Box::new(BlobPart::new(
                image_name,
                "image/png".to_string(),
                LOCAL_PNG.to_vec(),
            )));
            Ok(())
        })
        .unwrap();
    package
}
