//! Create a minimal presentation for native-Office interoperability checks.

use std::{env, error::Error, io};

use litchi_ooxml::pptx::Package;
use litchi_ooxml::pptx::tag::{List, Tag};

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
    package.save(&destination)?;

    // The legacy presentation writer and the lossless package editor are
    // deliberately separate. Reopening establishes an immutable package
    // snapshot before attaching a move-owned tag list.
    let mut package = Package::open(&destination)?;
    let mut tags = List::new();
    tags.add(Tag::new("LITCHI_VERIFY", "pptx-tag-crud-v1")?)?;
    let _ = package.put_tags(0, tags)?;
    package.save(destination)?;
    Ok(())
}
