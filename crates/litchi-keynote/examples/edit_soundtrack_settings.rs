//! Edit an existing Keynote soundtrack's playback mode and volume.

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
    soundtrack::{Mode, Settings},
};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: edit_soundtrack_settings <input.key> <output.key> \\
                     [--mode play-once|loop|do-not-play|unset] \\
                     [--volume NUMBER|unset] [--inverse PATH]";

#[derive(Default)]
struct Options {
    mode: Option<Option<Mode>>,
    volume: Option<Option<f64>>,
    inverse: Option<PathBuf>,
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let options = parse_options(&mut arguments)?;

    if options.mode.is_none() && options.volume.is_none() {
        return Err(invalid_input("specify --mode, --volume, or both"));
    }
    if input == output {
        return Err(invalid_input("input and output paths must differ"));
    }
    if options
        .inverse
        .as_deref()
        .is_some_and(|path| path == input || path == output)
    {
        return Err(invalid_input(
            "inverse path must differ from input and output paths",
        ));
    }

    let package = Package::open(&input)?;
    let before = package
        .soundtrack_settings()?
        .ok_or_else(|| invalid_input("presentation has no soundtrack"))?;
    let edit = package.edit_soundtrack_settings()?;
    let mut settings: Settings = edit.settings();
    if let Some(mode) = options.mode {
        settings.set_mode(mode)?;
    }
    if let Some(volume) = options.volume {
        settings.set_volume(volume)?;
    }
    let commit = edit.set(settings).commit()?;
    if commit.package().soundtrack_settings()? != Some(settings) {
        return Err(invalid_input(
            "committed soundtrack settings did not match the requested state",
        ));
    }

    let inverse = options
        .inverse
        .as_ref()
        .map(|_| {
            commit
                .package()
                .apply_soundtrack_settings(&commit.patch().inverse())
        })
        .transpose()?;
    if let Some(restored) = inverse.as_ref() {
        if exact_bytes(restored.package())? != exact_bytes(&package)? {
            return Err(invalid_input(
                "inverse patch did not restore the exact input package",
            ));
        }
        if restored.package().soundtrack_settings()? != Some(before) {
            return Err(invalid_input(
                "inverse patch did not restore the original soundtrack settings",
            ));
        }
    }

    save_new(&output, commit.package())?;
    if let (Some(path), Some(restored)) = (options.inverse, inverse) {
        save_new(&path, restored.package())?;
    }

    println!(
        "soundtrack settings: changed={}, touched_components={}, full_reparse_performed={}",
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().full_reparse_performed(),
    );
    Ok(())
}

fn parse_options(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<Options, Box<dyn Error>> {
    let mut options = Options::default();
    while let Some(flag) = arguments.next() {
        if flag == OsStr::new("--mode") {
            if options.mode.is_some() {
                return Err(invalid_input("--mode may be specified only once"));
            }
            options.mode = Some(parse_mode(required_text(
                arguments,
                "missing --mode value",
            )?)?);
        } else if flag == OsStr::new("--volume") {
            if options.volume.is_some() {
                return Err(invalid_input("--volume may be specified only once"));
            }
            options.volume = Some(parse_volume(required_text(
                arguments,
                "missing --volume value",
            )?)?);
        } else if flag == OsStr::new("--inverse") {
            if options.inverse.is_some() {
                return Err(invalid_input("--inverse may be specified only once"));
            }
            options.inverse = Some(PathBuf::from(required_argument(
                arguments,
                "missing --inverse path",
            )?));
        } else {
            return Err(invalid_input(
                "unexpected trailing argument; expected --mode, --volume, or --inverse",
            ));
        }
    }
    Ok(options)
}

fn parse_mode(value: String) -> Result<Option<Mode>, Box<dyn Error>> {
    match value.as_str() {
        "play-once" => Ok(Some(Mode::PlayOnce)),
        "loop" => Ok(Some(Mode::Loop)),
        "do-not-play" => Ok(Some(Mode::DoNotPlay)),
        "unset" => Ok(None),
        _ => Err(invalid_input(
            "--mode must be play-once, loop, do-not-play, or unset",
        )),
    }
}

fn parse_volume(value: String) -> Result<Option<f64>, Box<dyn Error>> {
    if value == "unset" {
        return Ok(None);
    }
    value.parse().map(Some).map_err(|_| {
        invalid_input("--volume must be a finite number between zero and one, or unset")
    })
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
        .map_err(|_| invalid_input("soundtrack settings arguments must be valid UTF-8"))
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

fn exact_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn invalid_input(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!("{}\n\n{USAGE}", message.into()),
    ))
}
