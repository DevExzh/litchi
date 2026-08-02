//! Remove a non-initial Pages section without deleting body text.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: remove_pages_section <input.pages> <output.pages> <section-id>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let section_id = arguments
        .next()
        .ok_or("missing section ID")?
        .parse::<u64>()?;

    let mut editor = PagesEditor::open(input)?;
    let removed = editor.remove_section(section_id)?;
    editor.save(output)?;
    println!(
        "removed section={} utf16_index={} name={:?}",
        removed.object_id, removed.character_index, removed.name
    );
    Ok(())
}
