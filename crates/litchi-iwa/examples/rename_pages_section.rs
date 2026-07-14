//! Rename one reachable Pages section.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: rename_pages_section <input.pages> <output.pages> <section-id> <name>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let section_id = arguments
        .next()
        .ok_or("missing section ID")?
        .parse::<u64>()?;
    let name = arguments.next().ok_or("missing section name")?;

    let mut editor = PagesEditor::open(input)?;
    editor.set_section_name(section_id, Some(&name))?;
    editor.save(output)?;
    Ok(())
}
