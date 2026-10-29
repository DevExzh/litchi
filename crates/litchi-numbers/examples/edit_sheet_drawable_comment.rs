//! Set or clear one Numbers sheet-drawable comment through selector-first APIs.

use std::env;
use std::path::PathBuf;

use litchi_numbers::sheet::drawable::DrawableSelector;
use litchi_numbers::{Package, SheetSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = PathBuf::from(
        arguments
            .next()
            .ok_or("usage: edit_sheet_drawable_comment <input.numbers> <output.numbers> <sheet-index> <drawable-index> <text|--clear>")?,
    );
    let output = PathBuf::from(arguments.next().ok_or("missing output path")?);
    let sheet_index: usize = arguments.next().ok_or("missing sheet index")?.parse()?;
    let drawable_index: usize = arguments.next().ok_or("missing drawable index")?.parse()?;
    let replacement = arguments
        .next()
        .ok_or("missing replacement text or --clear")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let package = Package::open(&input)?;
    let sheet = SheetSelector::index(sheet_index);
    let drawable = DrawableSelector::index(drawable_index);
    let old = package.sheet_drawable_comment(sheet, drawable)?;
    let mut edit = package.edit_sheet_drawable_comment(sheet, drawable)?;
    if replacement == "--clear" {
        edit.clear()?;
    } else {
        edit.set(&replacement)?;
    }
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_sheet_drawable_comment(&commit.patch().inverse())?;
    if exact_bytes(restored.package())? != exact_bytes(&package)? {
        return Err("inverse patch did not restore the exact source package".into());
    }
    commit.package().save(&output)?;

    let reopened = Package::open(&output)?;
    let new = reopened.sheet_drawable_comment(sheet, drawable)?;
    println!(
        "sheet={sheet_index} drawable_index={drawable_index} old={:?} new={:?}",
        old.as_ref().map(|comment| comment.body()),
        new.as_ref().map(|comment| comment.body())
    );
    Ok(())
}

fn exact_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}
