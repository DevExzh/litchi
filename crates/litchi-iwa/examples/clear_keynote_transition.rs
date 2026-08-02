use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: clear_keynote_transition <input.key> <output.key> <slide-index>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index = arguments
        .next()
        .ok_or("missing slide index")?
        .parse::<usize>()?;
    if arguments.next().is_some() {
        return Err("unexpected extra argument".into());
    }

    let mut editor = KeynoteEditor::open(input)?;
    editor.clear_slide_transition(slide_index)?;
    editor.save(output)?;
    Ok(())
}
