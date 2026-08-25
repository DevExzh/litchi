//! Replace the text owned by one Pages section through the focused package API.

use std::env;

use litchi_pages::{Package, SectionSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = env::args().skip(1);
    let input = args.next().ok_or(
        "usage: edit_pages_section_text <input.pages> <output.pages> <section-index> <text>",
    )?;
    let output = args.next().ok_or(
        "usage: edit_pages_section_text <input.pages> <output.pages> <section-index> <text>",
    )?;
    let section_index = args
        .next()
        .ok_or(
            "usage: edit_pages_section_text <input.pages> <output.pages> <section-index> <text>",
        )?
        .parse::<usize>()?;
    let replacement = args.next().ok_or(
        "usage: edit_pages_section_text <input.pages> <output.pages> <section-index> <text>",
    )?;
    if args.next().is_some() {
        return Err("unexpected trailing arguments".into());
    }

    let package = Package::open(input)?;
    let selector = SectionSelector::index(section_index);
    let previous = package.section_text(selector)?.to_owned();
    let commit = package.set_section_text(selector, &replacement)?;
    let mut output_file = std::fs::File::create(output)?;
    commit.package().write_to(&mut output_file)?;

    println!(
        "updated section {}: {:?} -> {:?}",
        section_index, previous, replacement
    );
    Ok(())
}
