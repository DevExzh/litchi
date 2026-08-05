//! Create a minimal presentation for native-Office interoperability checks.

use std::{env, error::Error, io};

use litchi_pptx::Package;
use litchi_pptx::tag::{List, Tag};

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

    // Reopening establishes an immutable package snapshot before attaching a
    // move-owned tag list through the standalone package service.
    let package = Package::open(&destination)?;
    let owner = package
        .presentation()?
        .slide(0)?
        .ok_or_else(|| io::Error::other("created presentation has no slide"))?
        .part()
        .part()
        .partname()
        .clone();
    let mut graph = package.opc()?.clone();
    let mut tags = List::new();
    tags.add(Tag::new("LITCHI_VERIFY", "pptx-tag-crud-v1")?)?;
    let _ = litchi_pptx::tag::put(&mut graph, &owner, tags)?;
    let mut package = Package::from_opc_package(graph)?;
    package.save(destination)?;
    Ok(())
}
