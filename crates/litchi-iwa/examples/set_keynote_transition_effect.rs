use std::env;

use litchi_iwa::keynote::{Effect, KeynoteEditor};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: set_keynote_transition_effect <input.key> <output.key> <slide-index> <none|dissolve|magic-move|native-identifier>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index = arguments
        .next()
        .ok_or("missing slide index")?
        .parse::<usize>()?;
    let effect = match arguments.next().ok_or("missing effect")?.as_str() {
        "none" => None,
        "dissolve" => Some(Effect::Dissolve),
        "magic-move" => Some(Effect::MagicMove),
        identifier => Some(Effect::from_identifier(identifier)),
    };
    if arguments.next().is_some() {
        return Err("unexpected extra argument".into());
    }

    let mut editor = KeynoteEditor::open(input)?;
    if effect.is_none() {
        editor.clear_slide_transition(slide_index)?;
        editor.save(output)?;
        return Ok(());
    }
    let mut settings = editor
        .slide_transition(slide_index)?
        .ok_or("slide has no modern transition attributes")?;
    settings.effect = effect;
    editor.set_slide_transition(slide_index, settings)?;
    editor.save(output)?;
    Ok(())
}
