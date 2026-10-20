//! Create a Keynote presentation with a native editable custom action timing curve.

use litchi_iwa::keynote::{self};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_keynote_custom_timing_curve <output.key>")?;

    let mut keynote = keynote::Builder::new()
        .title("Custom timing curve")
        .subtitle("Created entirely by litchi-iwa")
        .build()?;
    let drawable = keynote
        .slide_drawables(0)?
        .into_iter()
        .next()
        .ok_or("the initial slide has no drawable")?;
    let settings =
        keynote::BuildSettings::rotate_action(720.0, keynote::RotationDirection::Clockwise)
            .with_custom_timing_curve(keynote::BuildTimingCurve::cubic(
                keynote::MotionPathPoint::new(0.18, 0.04),
                keynote::MotionPathPoint::new(0.82, 0.96),
            ))?;
    keynote.add_slide_build(0, drawable.id.get(), settings)?;
    keynote.save(output)?;
    Ok(())
}
