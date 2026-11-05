//! Add a Pages section after another one, or remove a section (record 0312).
//!
//! ```text
//! cargo run -p litchi-pages --example edit_section_lifecycle -- \
//!     add <input.pages> <output.pages> <after-index> [name]
//! cargo run -p litchi-pages --example edit_section_lifecycle -- \
//!     remove <input.pages> <output.pages> <section-index>
//! ```

use std::error::Error;

use litchi_pages::{Package, SectionSelector};

fn main() -> Result<(), Box<dyn Error>> {
    let usage = "usage: edit_section_lifecycle add|remove <input.pages> <output.pages> <section-index> [name]";
    let mut arguments = std::env::args().skip(1);
    let verb = arguments.next().ok_or(usage)?;
    let input = arguments.next().ok_or(usage)?;
    let output = arguments.next().ok_or(usage)?;
    let position = arguments.next().ok_or(usage)?.parse::<usize>()?;
    let name = arguments.next();
    let package = Package::open(input)?;
    let edit = package.edit_section_lifecycle();
    let edit = match verb.as_str() {
        "add" => edit.add(SectionSelector::index(position), name.as_deref())?,
        "remove" => edit.remove(SectionSelector::index(position))?,
        _ => return Err(usage.into()),
    };
    let commit = edit.commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    std::fs::write(output, bytes)?;
    println!(
        "sections={} objects={}",
        commit.package().sections().len(),
        commit.diagnostics().objects()
    );
    Ok(())
}
