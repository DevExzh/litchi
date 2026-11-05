//! Duplicate one Keynote slide, or add a fresh slide from a theme layout,
//! through the selector-first slide-creation transaction.
//!
//! ```text
//! duplicate_slide <input.key> <output.key> <index:N|name:NAME> <destination>
//! duplicate_slide <input.key> <output.key> layout:<NAME> <destination>
//! duplicate_slide <input.key> --layouts
//! ```

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::fs::OpenOptions;
use std::path::{Path, PathBuf};

use litchi_keynote::{Package, Position, SlideSelector};

const USAGE: &str = "usage: duplicate_slide <input.key> <output.key> \
                     <index:N|name:NAME|layout:NAME> <destination> \
                     | duplicate_slide <input.key> --layouts";

fn main() -> Result<(), Box<dyn Error>> {
    let arguments = std::env::args().skip(1).collect::<Vec<_>>();
    match arguments.as_slice() {
        [input, flag] if flag == "--layouts" => list_layouts(Path::new(input)),
        [input, output, source, destination] => create(
            Path::new(input),
            &PathBuf::from(output),
            source,
            destination.parse()?,
        ),
        _ => Err(USAGE.into()),
    }
}

fn list_layouts(input: &Path) -> Result<(), Box<dyn Error>> {
    let package = Package::open(input)?;
    let default = package.default_slide_layout()?;
    for layout in package.slide_layouts()?.iter() {
        let marker = if default.as_ref() == Some(layout) {
            " (default)"
        } else {
            ""
        };
        println!(
            "{}: {}{marker}",
            layout.position().get(),
            layout.name().unwrap_or("<unnamed>")
        );
    }
    Ok(())
}

fn create(
    input: &Path,
    output: &Path,
    source: &str,
    destination: usize,
) -> Result<(), Box<dyn Error>> {
    if input == output {
        return Err("input and output paths must differ".into());
    }
    let package = Package::open(input)?;
    let mut edit = package.edit_slide_creation();
    let destination = Position::new(destination);
    if let Some(layout) = source.strip_prefix("layout:") {
        edit.add_slide(layout, destination)?;
    } else if let Some(index) = source.strip_prefix("index:") {
        edit.duplicate_slide(SlideSelector::index(index.parse()?), destination)?;
    } else if let Some(name) = source.strip_prefix("name:") {
        edit.duplicate_slide(SlideSelector::name(name), destination)?;
    } else {
        return Err(USAGE.into());
    }
    let commit = edit.commit()?;

    let restored = commit
        .package()
        .apply_slide_creation(&commit.patch().inverse())?;
    if restored.package().slides()?.len() != package.slides()?.len() {
        return Err("inverse patch did not restore the input slide count".into());
    }

    let mut file = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(output)?;
    commit.package().write_to(&mut file)?;
    file.sync_all()?;
    println!(
        "slides: {} -> {}, created={}, touched_components={}",
        package.slides()?.len(),
        commit.package().slides()?.len(),
        commit.diagnostics().created_slides(),
        commit.diagnostics().touched_components(),
    );
    Ok(())
}
