//! Replace one reachable Pages header/footer text slot.

use std::{env, fs::File};

use litchi_pages::Position;
use litchi_pages::header_footer::{HeaderFooterSelector, Kind, Template};
use litchi_pages::{Package, SectionSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or("usage: edit_pages_header_footer <input.pages> <output.pages> <section-index> <first|even|odd> <header|footer> <slot> <text>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let section_index = arguments
        .next()
        .ok_or("missing section index")?
        .parse::<usize>()?;
    let template = match arguments.next().ok_or("missing template")? {
        value if value.eq_ignore_ascii_case("first") => Template::First,
        value if value.eq_ignore_ascii_case("even") => Template::Even,
        value if value.eq_ignore_ascii_case("odd") => Template::Odd,
        _ => return Err("template must be first, even, or odd".into()),
    };
    let kind = match arguments.next().ok_or("missing header/footer kind")? {
        value if value.eq_ignore_ascii_case("header") => Kind::Header,
        value if value.eq_ignore_ascii_case("footer") => Kind::Footer,
        _ => return Err("kind must be header or footer".into()),
    };
    let slot = arguments.next().ok_or("missing slot")?.parse::<usize>()?;
    let replacement = arguments.next().ok_or("missing replacement text")?;

    let package = Package::open(input)?;
    let selector = HeaderFooterSelector::new(
        SectionSelector::index(section_index),
        template,
        kind,
        Position::new(slot),
    );
    let mut edit = package.edit_header_footer_text(selector)?;
    edit.set(&replacement)?;
    let commit = edit.commit()?;
    let mut output = File::create(output)?;
    commit.package().write_to(&mut output)?;
    Ok(())
}
