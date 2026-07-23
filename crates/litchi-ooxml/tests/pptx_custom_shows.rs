use litchi_ooxml::pptx::Package;
use litchi_ooxml::{OoxmlError, PackURI};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/custom-shows/presentation.xml");
const SLIDE_XML: &[u8] = include_bytes!("../../../test-data/ooxml/pptx/custom-shows/slide.xml");

#[test]
fn presentation_custom_shows_resolve_slide_membership() {
    let package = package_with_custom_shows();
    let custom_shows = package.presentation().unwrap().custom_shows().unwrap();

    assert_eq!(custom_shows.shows.len(), 2);
    assert_eq!(custom_shows.shows[0].id, 7);
    assert_eq!(custom_shows.shows[0].name, "Opening");
    assert_eq!(custom_shows.shows[0].slide_ids, [256, 258]);
    assert_eq!(custom_shows.shows[1].id, 8);
    assert_eq!(custom_shows.shows[1].name, "Recap");
    assert_eq!(custom_shows.shows[1].slide_ids, [257]);
}

#[test]
fn presentation_custom_shows_validate_slide_relationships() {
    let mut package = package_with_custom_shows();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let presentation = package
        .opc_package_mut()
        .get_part_mut(&presentation_name)
        .unwrap();
    presentation.rels_mut().remove("rIdSlideOne");
    presentation.rels_mut().add_relationship(
        rt::THEME.to_string(),
        "slides/slide1.xml".to_string(),
        "rIdSlideOne".to_string(),
        false,
    );

    assert!(matches!(
        package.presentation().unwrap().custom_shows(),
        Err(OoxmlError::InvalidFormat(message))
            if message.contains("is not an internal slide relationship")
    ));
}

fn package_with_custom_shows() -> Package {
    let mut package = Package::new().unwrap();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();

    {
        let presentation = package
            .opc_package_mut()
            .get_part_mut(&presentation_name)
            .unwrap();
        presentation.set_blob(PRESENTATION_XML.to_vec());
        for (relationship_id, target) in [
            ("rIdSlideOne", "slides/slide1.xml"),
            ("rIdSlideTwo", "slides/slide2.xml"),
            ("rIdSlideThree", "slides/slide3.xml"),
        ] {
            presentation.rels_mut().add_relationship(
                rt::SLIDE.to_string(),
                target.to_string(),
                relationship_id.to_string(),
                false,
            );
        }
    }

    for name in [
        "/ppt/slides/slide1.xml",
        "/ppt/slides/slide2.xml",
        "/ppt/slides/slide3.xml",
    ] {
        package.opc_package_mut().add_part(Box::new(BlobPart::new(
            PackURI::new(name).unwrap(),
            ct::PML_SLIDE.to_string(),
            SLIDE_XML.to_vec(),
        )));
    }
    package
}
