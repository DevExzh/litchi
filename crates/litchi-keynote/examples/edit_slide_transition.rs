//! Edit one Keynote slide transition without exposing native identities.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example intentionally reports its committed semantic change"
)]

use std::env;
use std::fs::OpenOptions;

use litchi_keynote::{Effect, Package, SlideSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_slide_transition <input.key> <output.key> <index:N|name:NAME> <clear|set> [effect duration-seconds delay-seconds automatic] [inverse-output.key]",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let selector_argument = arguments.next().ok_or("missing slide selector")?;
    let selector = parse_selector(&selector_argument)?;
    let operation = arguments.next().ok_or("missing transition operation")?;

    let package = Package::open(input)?;
    let mut edit = package.edit_slide_transition(selector)?;
    let inverse_output = match operation.as_str() {
        "clear" => {
            edit.clear()?;
            arguments.next()
        },
        "set" => {
            let effect = parse_effect(&arguments.next().ok_or("missing transition effect")?)?;
            let duration = arguments
                .next()
                .ok_or("missing transition duration")?
                .parse()?;
            let delay = arguments
                .next()
                .ok_or("missing transition delay")?
                .parse()?;
            let automatic = arguments
                .next()
                .ok_or("missing automatic transition flag")?
                .parse()?;
            let mut settings = edit
                .settings()
                .cloned()
                .ok_or("selected slide has no editable modern transition")?;
            settings.set_effect(Some(effect))?;
            settings.set_duration(Some(duration))?;
            settings.set_delay(Some(delay))?;
            settings.set_is_automatic(Some(automatic));
            edit.set_transition(settings)?;
            arguments.next()
        },
        _ => return Err("transition operation must be clear or set".into()),
    };
    if arguments.next().is_some() {
        return Err("unexpected trailing arguments".into());
    }
    let commit = edit.commit()?;
    write_new(&output, commit.package())?;
    if let Some(inverse_path) = inverse_output {
        let restored = commit
            .package()
            .apply_slide_transition(&commit.patch().inverse())?;
        write_new(&inverse_path, restored.package())?;
    }
    println!(
        "slide transition: changed={}, touched_components={}",
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
    );
    Ok(())
}

fn parse_selector(value: &str) -> Result<SlideSelector<'_>, Box<dyn std::error::Error>> {
    if let Some(index) = value.strip_prefix("index:") {
        return Ok(SlideSelector::index(index.parse()?));
    }
    if let Some(name) = value.strip_prefix("name:") {
        return Ok(SlideSelector::name(name));
    }
    Err("selector must start with index: or name:".into())
}

fn parse_effect(value: &str) -> Result<Effect, Box<dyn std::error::Error>> {
    match value {
        "none" => Ok(Effect::None),
        "dissolve" => Ok(Effect::Dissolve),
        "magic-move" => Ok(Effect::MagicMove),
        identifier => Ok(Effect::unknown(identifier)?),
    }
}

fn write_new(path: &str, package: &Package) -> Result<(), Box<dyn std::error::Error>> {
    let mut destination = OpenOptions::new().write(true).create_new(true).open(path)?;
    package.write_to(&mut destination)?;
    destination.sync_all()?;
    Ok(())
}
