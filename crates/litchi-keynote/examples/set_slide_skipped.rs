//! Set one Keynote slide's playback skip state without exposing native IDs.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example intentionally reports its committed semantic change"
)]

use std::error::Error;
use std::fs::OpenOptions;
use std::path::PathBuf;

use litchi_keynote::{Package, SlideSelector};

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(arguments.next().ok_or(
        "usage: set_slide_skipped <input.key> <output.key> <index:N|name:NAME> <true|false>",
    )?);
    let output = PathBuf::from(arguments.next().ok_or("missing output Keynote path")?);
    let selector = arguments
        .next()
        .ok_or("missing slide selector")?
        .into_string()
        .map_err(|_value| "slide selector is not valid UTF-8")?;
    let skipped = arguments
        .next()
        .ok_or("missing skip state")?
        .into_string()
        .map_err(|_value| "skip state is not valid UTF-8")?
        .parse::<bool>()?;
    if arguments.next().is_some() {
        return Err("too many arguments".into());
    }

    let package = Package::open(input)?;
    let mut edit = package.edit();
    if let Some(index) = selector.strip_prefix("index:") {
        edit.set_slide_skipped(SlideSelector::index(index.parse()?), skipped)?;
    } else if let Some(name) = selector.strip_prefix("name:") {
        edit.set_slide_skipped(SlideSelector::name(name), skipped)?;
    } else {
        return Err("selector must start with index: or name:".into());
    }

    let commit = edit.commit()?;
    let mut destination = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(output)?;
    commit.package().write_to(&mut destination)?;
    destination.sync_all()?;
    println!(
        "slide {}: skipped {} -> {}; changed={}, touched_components={}",
        commit.patch().position().get(),
        commit.patch().before(),
        commit.patch().after(),
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
    );
    Ok(())
}
