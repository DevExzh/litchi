#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_pptx::Error;
use litchi_pptx::presentation_properties::metadata::guides::{ColorKind, Guides, Orientation};

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/extended-guides/presentation.xml");
const INVALID_PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/extended-guides/invalid-presentation.xml");

#[test]
fn presentation_extended_guides_are_available_at_the_metadata_owner() {
    let guides = Guides::from_xml(PRESENTATION_XML).unwrap();

    let slide = guides.slide.as_ref().unwrap();
    assert_eq!(slide.guides.len(), 1);
    assert_eq!(slide.guides[0].id, 1);
    assert_eq!(slide.guides[0].name.as_deref(), Some("Middle"));
    assert_eq!(slide.guides[0].orientation, Some(Orientation::Horizontal));
    assert_eq!(slide.guides[0].position, Some(2160));
    assert_eq!(slide.guides[0].user_drawn, Some(true));
    assert_eq!(slide.guides[0].color.kind, ColorKind::Srgb);
    assert!(guides.notes.as_ref().unwrap().guides.is_empty());
}

#[test]
fn presentation_extended_guides_preserve_validation_errors() {
    let error = Guides::from_xml(INVALID_PRESENTATION_XML).unwrap_err();
    assert!(matches!(
        error,
        Error::Invalid(_) | Error::Xml(_) | Error::MarkupCompatibility(_) | Error::Decode(_)
    ));
}
