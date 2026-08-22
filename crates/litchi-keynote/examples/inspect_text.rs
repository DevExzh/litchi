//! Inspect reachable Keynote text without exposing native object identities.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example intentionally prints semantic inspection results"
)]

use std::env;
use std::io;

use litchi_keynote::{Package, SlideSelector};

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

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args_os().skip(1);
    let path = arguments
        .next()
        .ok_or("usage: inspect_text <presentation.key> [index:N|name:NAME]")?;
    let selected = arguments
        .next()
        .map(|value| {
            let value = value
                .into_string()
                .map_err(|_value| "slide selector must be valid UTF-8")?;
            parse_selector(&value)
        })
        .transpose()?;
    if arguments.next().is_some() {
        return Err("unexpected trailing argument".into());
    }

    let package = Package::open(path)?;
    package.validate()?;

    let stats = package.stats()?;
    println!(
        "slides={} objects={}",
        stats.slide_count, stats.total_objects
    );

    if let Some(selected) = selected {
        print_slide(&package, selected.selector())?;
    } else {
        for slide in package.slides()? {
            print_slide(&package, slide.position_selector())?;
        }
    }

    // Keep the aggregate view as a compatibility aid while the role-specific
    // values above make title/body/notes ownership explicit at the package
    // boundary.
    println!("{}", package.text()?);
    Ok(())
}

fn parse_selector(value: &str) -> Result<SelectedSlide, Box<dyn std::error::Error>> {
    if let Some(index) = value.strip_prefix("index:") {
        return index.parse().map(SelectedSlide::Index).map_err(|_error| {
            io::Error::new(io::ErrorKind::InvalidInput, "invalid slide index").into()
        });
    }
    if let Some(name) = value.strip_prefix("name:") {
        if name.is_empty() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "slide name must not be empty",
            )
            .into());
        }
        return Ok(SelectedSlide::Name(name.to_owned()));
    }
    Err(io::Error::new(
        io::ErrorKind::InvalidInput,
        "slide selector must start with index: or name:",
    )
    .into())
}

fn print_slide(
    package: &Package,
    selector: SlideSelector<'_>,
) -> Result<(), Box<dyn std::error::Error>> {
    let slide = package.show()?.select_slide(selector)?.ok_or_else(|| {
        io::Error::new(io::ErrorKind::NotFound, "slide selector matched no slide")
    })?;
    let title = package.slide_title(selector)?;
    let body = package.slide_body(selector)?;
    let notes = package.slide_notes(selector)?;

    println!(
        "slide={} name={:?}\ntitle={title:?}\nbody={body:?}\nnotes={notes:?}",
        slide.index(),
        slide.name(),
    );
    Ok(())
}
