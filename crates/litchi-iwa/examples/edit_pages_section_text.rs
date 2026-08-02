use std::env;

use litchi_iwa::pages::PagesEditor;

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

    let mut editor = PagesEditor::open(input)?;
    let section = editor
        .sections()
        .get(section_index)
        .cloned()
        .ok_or_else(|| format!("section index {section_index} is out of range"))?;
    let previous = editor.section_text(section.object_id)?;
    editor.set_section_text(section.object_id, &replacement)?;
    editor.save(output)?;

    println!(
        "updated section {} ({:?}): {:?} -> {:?}",
        section_index, section.name, previous, replacement
    );
    Ok(())
}
