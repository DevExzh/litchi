//! Change Keynote slide dimensions and playback flags.

use std::env;

use litchi_iwa::keynote::{KeynoteEditor, Mode};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_keynote_show <input.key> <output.key> <normal|self-playing|links-only> <width> <height> <loop> <autoplay>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let mode = match arguments
        .next()
        .ok_or("missing presentation mode")?
        .as_str()
    {
        "normal" => Mode::Normal,
        "self-playing" => Mode::SelfPlaying,
        "links-only" => Mode::LinksOnly,
        _ => return Err("presentation mode must be normal, self-playing, or links-only".into()),
    };
    let width = arguments.next().ok_or("missing width")?.parse::<f32>()?;
    let height = arguments.next().ok_or("missing height")?.parse::<f32>()?;
    let loop_presentation = arguments
        .next()
        .ok_or("missing loop flag")?
        .parse::<bool>()?;
    let autoplay = arguments
        .next()
        .ok_or("missing autoplay flag")?
        .parse::<bool>()?;

    let mut editor = KeynoteEditor::open(input)?;
    let mut settings = editor.show_settings()?;
    settings.set_size(width, height)?;
    settings.set_mode(Some(mode))?;
    settings.set_loop_presentation(Some(loop_presentation));
    settings.set_automatically_plays_upon_open(Some(autoplay));
    editor.set_show_settings(settings)?;
    editor.save(output)?;
    Ok(())
}
