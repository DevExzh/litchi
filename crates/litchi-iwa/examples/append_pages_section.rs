//! Append a Pages section that inherits an existing section's layout.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: append_pages_section <input.pages> <output.pages> <source-id> <name>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let source_id = arguments
        .next()
        .ok_or("missing source section ID")?
        .parse::<u64>()?;
    let name = arguments.next().ok_or("missing section name")?;

    let mut editor = PagesEditor::open(input)?;
    let created = editor.append_section(source_id, &name)?;
    editor.save(output)?;
    println!(
        "created section={} utf16_index={} name={:?}",
        created.object_id, created.character_index, created.name
    );
    Ok(())
}
