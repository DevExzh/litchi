//! Edit Keynote dimensions and playback behavior through semantic settings.

use std::error::Error;
use std::fs::OpenOptions;
use std::io::Write as _;
use std::path::{Path, PathBuf};

use litchi_keynote::{Mode, Package, Size};

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(arguments.next().ok_or(
        "usage: edit_show_settings <input.key> <output.key> <normal|self-playing|links-only> <width> <height> <loop> <autoplay> [inverse-output.key]",
    )?);
    let output = PathBuf::from(arguments.next().ok_or("missing output Keynote path")?);
    let mode_argument = arguments
        .next()
        .ok_or("missing semantic presentation mode")?
        .into_string()
        .map_err(|_value| "presentation mode is not valid UTF-8")?;
    let mode = match mode_argument.as_str() {
        "normal" => Mode::Normal,
        "self-playing" => Mode::SelfPlaying,
        "links-only" => Mode::LinksOnly,
        _ => return Err("presentation mode must be normal, self-playing, or links-only".into()),
    };
    let width = parse_argument::<f32>(&mut arguments, "missing presentation width")?;
    let height = parse_argument::<f32>(&mut arguments, "missing presentation height")?;
    let loop_presentation = parse_argument::<bool>(&mut arguments, "missing loop flag")?;
    let autoplay = parse_argument::<bool>(&mut arguments, "missing autoplay flag")?;
    let inverse_output = arguments.next().map(PathBuf::from);
    if arguments.next().is_some() {
        return Err("unexpected trailing arguments".into());
    }

    let package = Package::open(input)?;
    let mut edit = package.edit_show_settings()?;
    let settings = edit.settings_mut();
    settings.set_size(Size::new(width, height)?);
    settings.set_mode(Some(mode))?;
    settings.set_loop_presentation(Some(loop_presentation));
    settings.set_automatically_plays_upon_open(Some(autoplay));
    let commit = edit.commit()?;
    write_new(&output, commit.package().source_bytes())?;
    if let Some(inverse_path) = inverse_output {
        let inverse = commit.patch().inverse();
        let restored = commit.package().apply_show_settings(&inverse)?;
        write_new(&inverse_path, restored.package().source_bytes())?;
    }
    Ok(())
}

fn parse_argument<T>(
    arguments: &mut impl Iterator<Item = std::ffi::OsString>,
    missing: &'static str,
) -> Result<T, Box<dyn Error>>
where
    T: std::str::FromStr,
    T::Err: Error + 'static,
{
    Ok(arguments
        .next()
        .ok_or(missing)?
        .into_string()
        .map_err(|_value| "argument is not valid UTF-8")?
        .parse::<T>()?)
}

fn write_new(path: &Path, bytes: &[u8]) -> Result<(), Box<dyn Error>> {
    let mut destination = OpenOptions::new().write(true).create_new(true).open(path)?;
    destination.write_all(bytes)?;
    destination.sync_all()?;
    Ok(())
}
