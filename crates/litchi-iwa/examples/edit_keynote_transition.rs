//! Change a Keynote slide transition's timing and typed custom behavior.

use std::env;

use litchi_iwa::keynote::{
    KeynoteEditor, KeynoteTransitionAcceleration, KeynoteTransitionTextDelivery,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_keynote_transition <input.key> <output.key> <slide-index> <duration> <delay> <automatic> <unchanged|linear|ease-in|ease-out|ease-in-out|custom> <unchanged|object|word|character|line>",
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
    let acceleration = match arguments.next().ok_or("missing acceleration")?.as_str() {
        "unchanged" => None,
        "linear" => Some(KeynoteTransitionAcceleration::Linear),
        "ease-in" => Some(KeynoteTransitionAcceleration::EaseIn),
        "ease-out" => Some(KeynoteTransitionAcceleration::EaseOut),
        "ease-in-out" => Some(KeynoteTransitionAcceleration::EaseInOut),
        "custom" => Some(KeynoteTransitionAcceleration::Custom),
        _ => {
            return Err(
                "acceleration must be unchanged, linear, ease-in, ease-out, ease-in-out, or custom"
                    .into(),
            );
        },
    };
    let text_delivery = match arguments.next().ok_or("missing text delivery")?.as_str() {
        "unchanged" => None,
        "object" => Some(KeynoteTransitionTextDelivery::ByObject),
        "word" => Some(KeynoteTransitionTextDelivery::ByWord),
        "character" => Some(KeynoteTransitionTextDelivery::ByCharacter),
        "line" => Some(KeynoteTransitionTextDelivery::ByLine),
        _ => {
            return Err("text delivery must be unchanged, object, word, character, or line".into());
        },
    };

    let mut editor = KeynoteEditor::open(input)?;
    let mut transition = editor
        .slides()?
        .get(slide_index)
        .and_then(|slide| slide.transition.clone())
        .ok_or("slide has no modern transition attributes")?;
    transition.duration = Some(duration);
    transition.delay = Some(delay);
    transition.is_automatic = Some(automatic);
    if let Some(acceleration) = acceleration {
        transition.custom_parameters.acceleration = Some(acceleration);
    }
    if let Some(text_delivery) = text_delivery {
        transition.custom_parameters.text_delivery = Some(text_delivery);
    }
    editor.set_slide_transition(slide_index, transition)?;
    editor.save(output)?;
    Ok(())
}
