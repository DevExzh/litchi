use std::env;

use litchi_iwa::keynote::{KeynoteEditor, KeynoteSlideTextPlaceholder};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = env::args().skip(1);
    let input = args.next().ok_or(
        "usage: set_keynote_placeholder_visibility <input.key> <output.key> \
         <slide-index> <title|body> <true|false>",
    )?;
    let output = args.next().ok_or("missing output path")?;
    let slide_index = args.next().ok_or("missing slide index")?.parse()?;
    let placeholder = match args.next().as_deref() {
        Some("title") => KeynoteSlideTextPlaceholder::Title,
        Some("body") => KeynoteSlideTextPlaceholder::Body,
        Some(value) => {
            return Err(format!("unknown placeholder {value:?}; use title or body").into());
        },
        None => return Err("missing placeholder; use title or body".into()),
    };
    let visible = args.next().ok_or("missing visibility")?.parse()?;
    if args.next().is_some() {
        return Err("too many arguments".into());
    }

    let mut editor = KeynoteEditor::open(input)?;
    editor.set_slide_text_placeholder_visible(slide_index, placeholder, visible)?;
    editor.save(output)?;
    Ok(())
}
