//! Change a Keynote slide transition's timing and typed custom behavior.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;
use litchi_keynote::transition::{Acceleration, TextDelivery};

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
        "linear" => Some(Acceleration::Linear),
        "ease-in" => Some(Acceleration::EaseIn),
        "ease-out" => Some(Acceleration::EaseOut),
        "ease-in-out" => Some(Acceleration::EaseInOut),
        "custom" => Some(Acceleration::Custom),
        _ => {
            return Err(
                "acceleration must be unchanged, linear, ease-in, ease-out, ease-in-out, or custom"
                    .into(),
            );
        },
    };
    let text_delivery = match arguments.next().ok_or("missing text delivery")?.as_str() {
        "unchanged" => None,
        "object" => Some(TextDelivery::ByObject),
        "word" => Some(TextDelivery::ByWord),
        "character" => Some(TextDelivery::ByCharacter),
        "line" => Some(TextDelivery::ByLine),
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
    transition.set_duration(Some(duration))?;
    transition.set_delay(Some(delay))?;
    transition.set_is_automatic(Some(automatic));
    let mut custom_parameters = transition.custom_parameters().clone();
    if let Some(acceleration) = acceleration {
        custom_parameters.set_acceleration(Some(acceleration));
    }
    if let Some(text_delivery) = text_delivery {
        custom_parameters.set_text_delivery(Some(text_delivery));
    }
    transition.set_custom_parameters(custom_parameters)?;
    editor.set_slide_transition(slide_index, transition)?;
    editor.save(output)?;
    Ok(())
}
