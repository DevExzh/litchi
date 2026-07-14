//! Insert a native Pages section break at a UTF-16 body position.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(usage())?;
    let output = arguments.next().ok_or(usage())?;
    let source_section_id = arguments.next().ok_or(usage())?.parse::<u64>()?;
    let character_index = arguments.next().ok_or(usage())?.parse::<usize>()?;
    let name = arguments.next().ok_or(usage())?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = PagesEditor::open(input)?;
    let created = editor.insert_section(source_section_id, character_index, &name)?;
    editor.save(output)?;
    println!(
        "inserted section={} utf16_index={} name={:?}",
        created.object_id, created.character_index, created.name
    );
    Ok(())
}

fn usage() -> &'static str {
    "usage: insert_pages_section <input.pages> <output.pages> \
     <source-section-id> <utf16-character-index> <name>"
}
