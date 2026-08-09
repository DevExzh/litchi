#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_pptx::animations::{Effect, Sequence};

const LAYOUT_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/master-animations/layout.xml");
const MASTER_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/master-animations/master.xml");

#[test]
fn layout_and_master_timing_metadata_is_exposed_by_animation_owner() {
    // Presentation::SlideLayout and SlideMaster deliberately remain focused
    // graph views. Their timing parts are parsed by the standalone animation
    // owner directly, preserving the same typed effect assertions.
    let layout_animations = Sequence::parse_slide_xml(LAYOUT_XML).unwrap();
    assert_eq!(layout_animations.animations.len(), 1);
    assert_eq!(layout_animations.animations[0].shape_id, 3);
    assert_eq!(layout_animations.animations[0].effect, Effect::Fade);

    let master_animations = Sequence::parse_slide_xml(MASTER_XML).unwrap();
    assert_eq!(master_animations.animations.len(), 1);
    assert_eq!(master_animations.animations[0].shape_id, 4);
    assert_eq!(master_animations.animations[0].effect, Effect::Fade);
}
