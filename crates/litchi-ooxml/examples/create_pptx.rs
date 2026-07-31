//! Create a minimal presentation for native-Office interoperability checks.

use std::{env, error::Error, io};

use litchi_ooxml::pptx::Package;

fn main() -> Result<(), Box<dyn Error + Send + Sync>> {
    let destination = env::args_os().nth(1).ok_or_else(|| {
        io::Error::new(
            io::ErrorKind::InvalidInput,
            "usage: create_pptx <destination.pptx>",
        )
    })?;

    let mut package = Package::new()?;
    let slide = package.presentation_mut()?.add_slide()?;
    slide.set_title("Litchi Office verification");
    slide.add_text_box(
        "Created without a PowerPoint repair prompt",
        914_400,
        1_828_800,
        7_315_200,
        914_400,
    );
    package.save(destination)?;
    Ok(())
}
