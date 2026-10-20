//! Create a Keynote presentation with a native editable Dissolve Build In.

use litchi_iwa::keynote::{self};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_keynote_dissolve_build <output.key>")?;

    let mut keynote = keynote::Builder::new()
        .title("Dissolve build")
        .subtitle("Created entirely by litchi-iwa")
        .build()?;
    let drawable = keynote
        .slide_drawables(0)?
        .into_iter()
        .next()
        .ok_or("the initial slide has no drawable")?;
    keynote.add_slide_build(0, drawable.id.get(), keynote::BuildSettings::dissolve_in())?;
    keynote.save(output)?;
    Ok(())
}
