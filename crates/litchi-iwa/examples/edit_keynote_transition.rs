//! Change a Keynote slide transition's timing while preserving its effect.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_keynote_transition <input.key> <output.key> <slide-index> <duration> <delay> <automatic>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index = arguments
        .next()
        .ok_or("missing slide index")?
        .parse::<usize>()?;
    let duration = arguments.next().ok_or("missing duration")?.parse::<f64>()?;
    let delay = arguments.next().ok_or("missing delay")?.parse::<f64>()?;
    let automatic = arguments
        .next()
        .ok_or("missing automatic flag")?
        .parse::<bool>()?;

    let mut editor = KeynoteEditor::open(input)?;
    let mut transition = editor
        .slides()?
        .get(slide_index)
        .and_then(|slide| slide.transition.clone())
        .ok_or("slide has no modern transition attributes")?;
    transition.duration = Some(duration);
    transition.delay = Some(delay);
    transition.is_automatic = Some(automatic);
    editor.set_slide_transition(slide_index, transition)?;
    editor.save(output)?;
    Ok(())
}
