//! Create a minimal presentation for native-Office interoperability checks.

use std::{env, error::Error, io};

use litchi_pptx::tag::{List, Tag};

fn main() -> Result<(), Box<dyn Error + Send + Sync>> {
    let destination = env::args_os().nth(1).ok_or_else(|| {
        io::Error::new(
            io::ErrorKind::InvalidInput,
            "usage: create_pptx <destination.pptx>",
        )
    })?;

    let mut builder = litchi_pptx::builder::Builder::new()?;
    let slide = builder.presentation()?.add_slide()?;
    slide.set_title("Litchi Office verification");
    slide.add_text_box(
        "Created without a PowerPoint repair prompt",
        914_400,
        1_828_800,
        7_315_200,
        914_400,
    );
    let package = builder.build()?;
    let mut tags = List::new();
    tags.add(Tag::new("LITCHI_VERIFY", "pptx-tag-crud-v1")?)?;
    let mut edit = package.edit_tags()?;
    let _ = edit.set(0_usize, tags)?;
    let package = edit.commit()?.into_package();
    package.save(destination)?;
    Ok(())
}
