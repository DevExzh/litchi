//! Edit Keynote dimensions and playback behavior through semantic settings.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::{OsStr, OsString};
use std::io;
use std::path::{Path, PathBuf};

use litchi_keynote::{
    Package,
    show::{Mode, Settings, Size},
};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: edit_show_settings <input.key> <output.key> \\
                     <normal|self-playing|links-only> <width> <height> <true|false:loop> \\
                     <true|false:autoplay> [--inverse PATH] \\
                     [--slide-numbers true|false|unset]";

#[derive(Default)]
struct OutputOptions {
    inverse: Option<PathBuf>,
    slide_numbers: Option<Option<bool>>,
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let mode = parse_mode(required_text(&mut arguments, "missing presentation mode")?)?;
    let width = parse_number(
        required_text(&mut arguments, "missing presentation width")?,
        "width",
    )?;
    let height = parse_number(
        required_text(&mut arguments, "missing presentation height")?,
        "height",
    )?;
    let loop_presentation =
        parse_bool(required_text(&mut arguments, "missing loop flag")?, "loop")?;
    let autoplay = parse_bool(
        required_text(&mut arguments, "missing autoplay flag")?,
        "autoplay",
    )?;
    let output_options = parse_output_options(&mut arguments)?;

    if input == output {
        return Err(invalid_input("input and output paths must differ"));
    }
    if output_options
        .inverse
        .as_deref()
        .is_some_and(|path| path == input || path == output)
    {
        return Err(invalid_input(
            "inverse path must differ from input and output paths",
        ));
    }

    let package = Package::open(&input)?;
    let before = package.show_settings()?;
    let edit = package.edit_show_settings()?;
    let mut settings: Settings = edit.settings();
    settings.set_size(Size::new(width, height)?);
    settings.set_mode(Some(mode))?;
    settings.set_loop_presentation(Some(loop_presentation));
    settings.set_automatically_plays_upon_open(Some(autoplay));
    if let Some(slide_numbers) = output_options.slide_numbers {
        settings.set_slide_numbers_visible(slide_numbers);
    }
    let commit = edit.set(settings).commit()?;

    let inverse = output_options
        .inverse
        .as_ref()
        .map(|_| {
            commit
                .package()
                .apply_show_settings(&commit.patch().inverse())
        })
        .transpose()?;
    if let Some(restored) = inverse.as_ref() {
        if restored.package().show_settings()? != before {
            return Err(invalid_input(
                "inverse patch did not restore original show settings",
            ));
        }
    }

    save_new(&output, commit.package())?;
    if let (Some(path), Some(restored)) = (output_options.inverse, inverse) {
        save_new(&path, restored.package())?;
    }

    println!(
        "show settings: changed={}, touched_components={}, deleted_previews={}",
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().deleted_previews(),
    );
    Ok(())
}

fn parse_mode(value: String) -> Result<Mode, Box<dyn Error>> {
    match value.as_str() {
        "normal" => Ok(Mode::Normal),
        "self-playing" => Ok(Mode::SelfPlaying),
        "links-only" => Ok(Mode::LinksOnly),
        _ => Err(invalid_input(
            "presentation mode must be normal, self-playing, or links-only",
        )),
    }
}

fn parse_number(value: String, name: &str) -> Result<f32, Box<dyn Error>> {
    value
        .parse()
        .map_err(|_| invalid_input(format!("presentation {name} must be a finite number")))
}

fn parse_bool(value: String, name: &str) -> Result<bool, Box<dyn Error>> {
    match value.as_str() {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(invalid_input(format!("{name} must be true or false"))),
    }
}

fn required_argument(
    arguments: &mut impl Iterator<Item = OsString>,
    message: &'static str,
) -> Result<OsString, Box<dyn Error>> {
    arguments.next().ok_or_else(|| invalid_input(message))
}

fn required_text(
    arguments: &mut impl Iterator<Item = OsString>,
    message: &'static str,
) -> Result<String, Box<dyn Error>> {
    required_argument(arguments, message)?
        .into_string()
        .map_err(|_| invalid_input("settings arguments must be valid UTF-8"))
}

fn parse_output_options(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<OutputOptions, Box<dyn Error>> {
    let mut options = OutputOptions::default();
    while let Some(flag) = arguments.next() {
        if flag == OsStr::new("--inverse") {
            if options.inverse.is_some() {
                return Err(invalid_input("--inverse may be specified only once"));
            }
            options.inverse = Some(PathBuf::from(required_argument(
                arguments,
                "missing --inverse path",
            )?));
        } else if flag == OsStr::new("--slide-numbers") {
            if options.slide_numbers.is_some() {
                return Err(invalid_input("--slide-numbers may be specified only once"));
            }
            let value = required_text(arguments, "missing --slide-numbers value")?;
            options.slide_numbers = Some(match value.as_str() {
                "unset" => None,
                "true" => Some(true),
                "false" => Some(false),
                _ => {
                    return Err(invalid_input(
                        "--slide-numbers must be true, false, or unset",
                    ));
                },
            });
        } else {
            return Err(invalid_input(
                "unexpected trailing argument; expected --inverse or --slide-numbers",
            ));
        }
    }
    Ok(options)
}

/// Publishes through a sibling temporary file without overwriting an existing target.
fn save_new(path: &Path, package: &Package) -> Result<(), Box<dyn Error>> {
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let mut temporary = NamedTempFile::new_in(parent)?;
    package.write_to(temporary.as_file_mut())?;
    temporary.as_file().sync_all()?;
    temporary
        .persist_noclobber(path)
        .map_err(|error| Box::new(error.error))?;
    Ok(())
}

fn invalid_input(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!("{}\n\n{USAGE}", message.into()),
    ))
}
