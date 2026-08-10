//! Edit one Keynote slide transition without exposing native identities.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example intentionally reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::{OsStr, OsString};
use std::io;
use std::path::{Path, PathBuf};

use litchi_keynote::{
    Package, SlideSelector,
    transition::{Acceleration, Effect, TextDelivery},
};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: edit_slide_transition <input.key> <output.key> \\
                     <index:N|name:NAME> <clear|set> \\
                     [effect duration-seconds delay-seconds automatic \\
                     unchanged|linear|ease-in|ease-out|ease-in-out|custom \\
                     unchanged|object|word|character|line] [--inverse PATH]";

enum SelectedSlide {
    Index(usize),
    Name(String),
}

impl SelectedSlide {
    fn selector(&self) -> SlideSelector<'_> {
        match self {
            Self::Index(index) => SlideSelector::index(*index),
            Self::Name(name) => SlideSelector::name(name),
        }
    }
}

enum Operation {
    Clear,
    Set {
        effect: Effect,
        duration: f64,
        delay: f64,
        automatic: bool,
        acceleration: Option<Acceleration>,
        text_delivery: Option<TextDelivery>,
    },
}

struct OutputOptions {
    inverse: Option<PathBuf>,
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let slide = parse_selector(required_text(&mut arguments, "missing slide selector")?)?;
    let operation = parse_operation(&mut arguments)?;
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
    let edit = package.edit_slide_transition(slide.selector())?;
    let edit = match operation {
        Operation::Clear => edit.clear()?,
        Operation::Set {
            effect,
            duration,
            delay,
            automatic,
            acceleration,
            text_delivery,
        } => {
            let mut settings = edit
                .settings()
                .cloned()
                .ok_or_else(|| invalid_input("selected slide has no editable modern transition"))?;
            settings.set_effect(Some(effect))?;
            settings.set_duration(Some(duration))?;
            settings.set_delay(Some(delay))?;
            settings.set_is_automatic(Some(automatic));
            if acceleration.is_some() || text_delivery.is_some() {
                let mut custom_parameters = *settings.custom_parameters();
                if let Some(value) = acceleration {
                    custom_parameters.set_acceleration(Some(value));
                }
                if let Some(value) = text_delivery {
                    custom_parameters.set_text_delivery(Some(value));
                }
                settings.set_custom_parameters(custom_parameters)?;
            }
            edit.set(settings)?
        },
    };
    let commit = edit.commit()?;

    let inverse = output_options
        .inverse
        .as_ref()
        .map(|_| {
            commit
                .package()
                .apply_slide_transition(&commit.patch().inverse())
        })
        .transpose()?;
    if let Some(restored) = inverse.as_ref() {
        if exact_bytes(restored.package())? != exact_bytes(&package)? {
            return Err(invalid_input(
                "inverse patch did not restore the exact input package",
            ));
        }
    }

    save_new(&output, commit.package())?;
    if let (Some(path), Some(restored)) = (output_options.inverse, inverse) {
        save_new(&path, restored.package())?;
    }
    println!(
        "slide transition: changed={}, touched_components={}",
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
    );
    Ok(())
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
        .map_err(|_| invalid_input("transition arguments must be valid UTF-8"))
}

fn parse_selector(value: String) -> Result<SelectedSlide, Box<dyn Error>> {
    if let Some(index) = value.strip_prefix("index:") {
        return index
            .parse()
            .map(SelectedSlide::Index)
            .map_err(|_| invalid_input("slide index must be a non-negative integer"));
    }
    if let Some(name) = value.strip_prefix("name:") {
        return Ok(SelectedSlide::Name(name.to_owned()));
    }
    Err(invalid_input("selector must start with index: or name:"))
}

fn parse_operation(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<Operation, Box<dyn Error>> {
    match required_text(arguments, "missing transition operation")?.as_str() {
        "clear" => Ok(Operation::Clear),
        "set" => {
            let effect = parse_effect(required_text(arguments, "missing transition effect")?)?;
            let duration = parse_number(
                required_text(arguments, "missing transition duration")?,
                "duration",
            )?;
            let delay = parse_number(
                required_text(arguments, "missing transition delay")?,
                "delay",
            )?;
            let automatic = parse_bool(
                required_text(arguments, "missing automatic transition flag")?,
                "automatic transition flag",
            )?;
            let acceleration =
                parse_acceleration(required_text(arguments, "missing custom acceleration")?)?;
            let text_delivery =
                parse_text_delivery(required_text(arguments, "missing text delivery")?)?;
            Ok(Operation::Set {
                effect,
                duration,
                delay,
                automatic,
                acceleration,
                text_delivery,
            })
        },
        _ => Err(invalid_input("transition operation must be clear or set")),
    }
}

fn parse_effect(value: String) -> Result<Effect, Box<dyn Error>> {
    match value.as_str() {
        "none" => Ok(Effect::None),
        "dissolve" => Ok(Effect::Dissolve),
        "magic-move" => Ok(Effect::MagicMove),
        identifier => Ok(Effect::unknown(identifier)?),
    }
}

fn parse_number(value: String, label: &str) -> Result<f64, Box<dyn Error>> {
    value
        .parse()
        .map_err(|_| invalid_input(format!("transition {label} must be a finite number")))
}

fn parse_bool(value: String, label: &str) -> Result<bool, Box<dyn Error>> {
    match value.as_str() {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(invalid_input(format!("{label} must be true or false"))),
    }
}

fn parse_acceleration(value: String) -> Result<Option<Acceleration>, Box<dyn Error>> {
    match value.as_str() {
        "unchanged" => Ok(None),
        "linear" => Ok(Some(Acceleration::Linear)),
        "ease-in" => Ok(Some(Acceleration::EaseIn)),
        "ease-out" => Ok(Some(Acceleration::EaseOut)),
        "ease-in-out" => Ok(Some(Acceleration::EaseInOut)),
        "custom" => Ok(Some(Acceleration::Custom)),
        _ => Err(invalid_input(
            "acceleration must be unchanged, linear, ease-in, ease-out, ease-in-out, or custom",
        )),
    }
}

fn parse_text_delivery(value: String) -> Result<Option<TextDelivery>, Box<dyn Error>> {
    match value.as_str() {
        "unchanged" => Ok(None),
        "object" => Ok(Some(TextDelivery::ByObject)),
        "word" => Ok(Some(TextDelivery::ByWord)),
        "character" => Ok(Some(TextDelivery::ByCharacter)),
        "line" => Ok(Some(TextDelivery::ByLine)),
        _ => Err(invalid_input(
            "text delivery must be unchanged, object, word, character, or line",
        )),
    }
}

fn parse_output_options(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<OutputOptions, Box<dyn Error>> {
    let Some(flag) = arguments.next() else {
        return Ok(OutputOptions { inverse: None });
    };
    if flag != OsStr::new("--inverse") {
        return Err(invalid_input(
            "unexpected trailing argument; expected --inverse PATH",
        ));
    }
    let inverse = PathBuf::from(required_argument(arguments, "missing --inverse path")?);
    if arguments.next().is_some() {
        return Err(invalid_input("unexpected trailing arguments"));
    }
    Ok(OutputOptions {
        inverse: Some(inverse),
    })
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
