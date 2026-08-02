//! Change Pages page dimensions while preserving the remaining layout fields.

use std::env;

use litchi_iwa::pages::{PagesEditor, PagesPageOrientation};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_pages_layout <input.pages> <output.pages> <width> <height> \
             <portrait|landscape>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let width = arguments
        .next()
        .ok_or("missing page width")?
        .parse::<f32>()?;
    let height = arguments
        .next()
        .ok_or("missing page height")?
        .parse::<f32>()?;
    let orientation = match arguments.next().ok_or("missing page orientation")?.as_str() {
        "portrait" => PagesPageOrientation::Portrait,
        "landscape" => PagesPageOrientation::Landscape,
        _ => return Err("page orientation must be portrait or landscape".into()),
    };
    if arguments.next().is_some() {
        return Err("unexpected extra argument".into());
    }

    let mut editor = PagesEditor::open(input)?;
    let mut layout = editor.page_layout()?;
    layout.page_width = Some(width);
    layout.page_height = Some(height);
    layout.orientation = Some(orientation);
    editor.set_page_layout(layout)?;
    editor.save(output)?;
    Ok(())
}
