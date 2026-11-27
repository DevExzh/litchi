//! Move one Keynote slide through the selector-first immutable transaction.

use std::env;
use std::fs::OpenOptions;

use litchi_keynote::{Package, Position, SlideSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: move_slide <input.key> <output.key> <source-index-or-name> <destination-index> [inverse-output.key]",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let source = arguments.next().ok_or("missing source selector")?;
    let destination = arguments
        .next()
        .ok_or("missing destination index")?
        .parse::<usize>()?;
    let inverse_output = arguments.next();
    if arguments.next().is_some() {
        return Err("unexpected trailing arguments".into());
    }

    let package = Package::open(input)?;
    let mut edit = package.edit_slide_order();
    match source.parse::<usize>() {
        Ok(index) => {
            edit.move_slide(SlideSelector::index(index), Position::new(destination))?;
        },
        Err(_error) => {
            edit.move_slide(SlideSelector::name(&source), Position::new(destination))?;
        },
    }
    let commit = edit.commit()?;
    write_new(&output, commit.package())?;
    if let Some(inverse_path) = inverse_output {
        let inverse = commit.patch().inverse();
        let restored = commit.package().apply_slide_order(&inverse)?;
        write_new(&inverse_path, restored.package())?;
    }
    Ok(())
}

fn write_new(path: &str, package: &Package) -> Result<(), Box<dyn std::error::Error>> {
    let mut destination_file = OpenOptions::new().write(true).create_new(true).open(path)?;
    package.write_to(&mut destination_file)?;
    destination_file.sync_all()?;
    Ok(())
}
