//! Create an empty Keynote slide from a presentation theme layout.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: add_keynote_slide <input.key> <output.key> [layout-index]")?;
    let output = arguments.next().ok_or("missing output path")?;
    let requested: Option<usize> = arguments.next().map(|value| value.parse()).transpose()?;

    let mut editor = KeynoteEditor::open(input)?;
    let layouts = editor.slide_layouts()?;
    let layout = match requested {
        Some(index) => layouts.get(index).ok_or("layout index is out of range")?,
        None => layouts
            .iter()
            .find(|layout| layout.is_default)
            .ok_or("presentation theme has no default layout")?,
    };
    let layout_name = layout.name.clone();
    let slide = editor.add_slide(layout.id)?;
    editor.save(output)?;
    println!(
        "created slide {} (node {}, object {}) from layout {layout_name:?}",
        slide.index, slide.node_id, slide.slide_id
    );
    Ok(())
}
