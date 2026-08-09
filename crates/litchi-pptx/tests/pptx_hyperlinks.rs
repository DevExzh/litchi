#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::PackURI;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_pptx::{Error, Hyperlink, Package};

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/hyperlinks/presentation.xml");
const SLIDE_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/hyperlinks/slide-with-links.xml");
const MALFORMED_SLIDE_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/hyperlinks/malformed-slide.xml");

#[test]
fn presentation_hyperlinks_include_strict_relationships_and_inline_actions() {
    let package = package_with_hyperlinks();
    let hyperlinks = package.presentation().unwrap().hyperlinks().unwrap();

    assert_eq!(hyperlinks.len(), 2);
    assert_eq!(hyperlinks[0].0, 0);
    assert!(matches!(
        &hyperlinks[0].1,
        Hyperlink::External { url, tooltip: None } if url == "https://example.invalid/"
    ));
    assert_eq!(hyperlinks[1].0, 0);
    assert!(matches!(
        &hyperlinks[1].1,
        Hyperlink::Slide {
            slide_number: 2,
            tooltip: Some(tooltip),
        } if tooltip == "Next slide"
    ));
}

#[test]
fn presentation_hyperlinks_reject_malformed_inline_xml() {
    let mut package = package_with_hyperlinks();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    package
        .edit_opc(|opc| {
            opc.get_part_mut(&slide_name)
                .unwrap()
                .set_blob(MALFORMED_SLIDE_XML.to_vec());
            Ok(())
        })
        .unwrap();

    assert!(matches!(
        package.presentation().unwrap().hyperlinks(),
        Err(Error::Xml(_))
    ));
}

fn package_with_hyperlinks() -> Package {
    let mut package = Package::new().unwrap();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();

    package
        .edit_opc(|opc| {
            let presentation = opc.get_part_mut(&presentation_name).unwrap();
            presentation.set_blob(PRESENTATION_XML.to_vec());
            presentation.rels_mut().add_relationship(
                rt::SLIDE.to_string(),
                "slides/slide1.xml".to_string(),
                "rIdSlideOne".to_string(),
                false,
            );
            opc.add_part(Box::new(BlobPart::new(
                slide_name.clone(),
                ct::PML_SLIDE.to_string(),
                SLIDE_XML.to_vec(),
            )));
            opc.get_part_mut(&slide_name)
                .unwrap()
                .rels_mut()
                .add_relationship(
                    rt::STRICT_HYPERLINK.to_string(),
                    "https://example.invalid/".to_string(),
                    "rIdExternal".to_string(),
                    true,
                );
            Ok(())
        })
        .unwrap();
    package
}
