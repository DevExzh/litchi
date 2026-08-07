//! Inspect reachable Keynote text without exposing native object identities.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example intentionally prints semantic inspection results"
)]

use std::env;

use litchi_keynote::Package;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args_os()
        .nth(1)
        .ok_or("usage: inspect_text <presentation.key>")?;
    let package = Package::open(path)?;
    package.validate()?;

    let stats = package.stats()?;
    println!(
        "slides={} objects={}",
        stats.slide_count, stats.total_objects
    );
    println!("{}", package.text()?);
    Ok(())
}
