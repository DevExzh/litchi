//! Create a Keynote presentation with a native editable custom action timing curve.

use litchi_iwa::keynote::{
    KeynoteBuildSettings, KeynoteBuildTimingCurve, KeynoteDocumentBuilder, KeynoteMotionPathPoint,
    KeynoteRotationDirection,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_keynote_custom_timing_curve <output.key>")?;

    let mut keynote = KeynoteDocumentBuilder::new()
        .title("Custom timing curve")
        .subtitle("Created entirely by litchi-iwa")
        .build()?;
    let drawable = keynote
        .slide_drawables(0)?
        .into_iter()
        .next()
        .ok_or("the initial slide has no drawable")?;
    let settings = KeynoteBuildSettings::rotate_action(720.0, KeynoteRotationDirection::Clockwise)
        .with_custom_timing_curve(KeynoteBuildTimingCurve::cubic(
            KeynoteMotionPathPoint::new(0.18, 0.04),
            KeynoteMotionPathPoint::new(0.82, 0.96),
        ))?;
    keynote.add_slide_build(0, drawable.id.get(), settings)?;
    keynote.save(output)?;
    Ok(())
}
