//! Set or clear one Pages section name through a semantic selector.

use std::error::Error;
use std::fs::OpenOptions;
use std::io::Write as _;
use std::path::{Path, PathBuf};

use litchi_pages::{Package, SectionSelector};

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(arguments.next().ok_or(
        "usage: edit_section_name <input.pages> <output.pages> <section-index> <name|--clear> [inverse-output.pages]",
    )?);
    let output = PathBuf::from(arguments.next().ok_or("missing output Pages path")?);
    let position = arguments
        .next()
        .ok_or("missing semantic section index")?
        .into_string()
        .map_err(|_value| "section index is not valid UTF-8")?
        .parse::<usize>()?;
    let name = arguments
        .next()
        .ok_or("missing section name or --clear")?
        .into_string()
        .map_err(|_value| "section name is not valid UTF-8")?;
    let inverse_output = arguments.next().map(PathBuf::from);
    if arguments.next().is_some() {
        return Err("unexpected trailing arguments".into());
    }

    let package = Package::open(input)?;
    let mut edit = package.edit_section_name(SectionSelector::index(position))?;
    if name == "--clear" {
        edit.clear_name();
    } else {
        edit.set_name(Some(&name))?;
    }
    let commit = edit.commit()?;
    write_new(&output, commit.package().source_bytes())?;

    if let Some(inverse_path) = inverse_output {
        let inverse = commit.patch().inverse();
        let restored = commit.package().apply_section_name(&inverse)?;
        write_new(&inverse_path, restored.package().source_bytes())?;
    }
    Ok(())
}

fn write_new(path: &Path, bytes: &[u8]) -> Result<(), Box<dyn Error>> {
    let mut destination = OpenOptions::new().write(true).create_new(true).open(path)?;
    destination.write_all(bytes)?;
    destination.sync_all()?;
    Ok(())
}
