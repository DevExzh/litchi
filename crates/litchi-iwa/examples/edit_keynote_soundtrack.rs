//! Change a Keynote presentation soundtrack's playback mode and volume.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;
use litchi_keynote::soundtrack::{Mode, Settings};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_keynote_soundtrack <input.key> <output.key> <play-once|loop|do-not-play> <volume>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let mode = match arguments.next().ok_or("missing soundtrack mode")?.as_str() {
        "play-once" => Mode::PlayOnce,
        "loop" => Mode::Loop,
        "do-not-play" => Mode::DoNotPlay,
        _ => return Err("soundtrack mode must be play-once, loop, or do-not-play".into()),
    };
    let volume = arguments.next().ok_or("missing volume")?.parse::<f64>()?;

    let mut editor = KeynoteEditor::open(input)?;
    let _current = editor
        .soundtrack_settings()?
        .ok_or("presentation has no soundtrack object")?;
    editor.set_soundtrack_settings(Settings::new(Some(volume), Some(mode))?)?;
    editor.save(output)?;
    Ok(())
}
