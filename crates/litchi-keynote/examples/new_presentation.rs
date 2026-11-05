//! Author a fresh one-slide Keynote presentation.
//!
//! With only an output path, writes the deterministic blank presentation
//! (`Package::new`). With a title (and optional subtitle and notes), builds an
//! independent document with a freshly generated identity.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports what it wrote"
)]

use std::error::Error;
use std::fs::OpenOptions;
use std::path::PathBuf;

use litchi_keynote::{Builder, Identity, Package};

const USAGE: &str = "usage: new_presentation <output.key> [TITLE [SUBTITLE [NOTES]]]";

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args().skip(1);
    let output = PathBuf::from(arguments.next().ok_or(USAGE)?);
    let title = arguments.next();
    let subtitle = arguments.next();
    let notes = arguments.next();
    if arguments.next().is_some() {
        return Err(USAGE.into());
    }

    let package = match title {
        None => Package::new(),
        Some(title) => {
            let mut builder = Builder::new().identity(Identity::generate()).title(title);
            if let Some(subtitle) = subtitle {
                builder = builder.subtitle(subtitle);
            }
            if let Some(notes) = notes {
                builder = builder.notes(notes);
            }
            builder.build()?
        },
    };

    let mut destination = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&output)?;
    package.write_to(&mut destination)?;
    destination.sync_all()?;
    println!(
        "wrote {} slide(s), identity {}",
        package.slides()?.len(),
        package
            .identity()?
            .map_or_else(|| "absent".to_owned(), |identity| identity.to_string())
    );
    Ok(())
}
