//! Generate a PPTX table-style artifact for native Office checks.
//!
//! Run with:
//!
//! ```sh
//! cargo run -p litchi-pptx --example owner_native_smoke -- \
//!     target/office-owner-smoke
//! ```

use std::io;
use std::path::PathBuf;

use litchi_pptx::Package;
use litchi_pptx::table::style::{self, Def, Id, List, Parts};

type Result<T> = std::result::Result<T, Box<dyn std::error::Error + Send + Sync>>;

const TABLE_STYLE: &str = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";

fn main() -> Result<()> {
    let directory = std::env::args_os()
        .nth(1)
        .map_or_else(|| PathBuf::from("target/office-owner-smoke"), PathBuf::from);
    std::fs::create_dir_all(&directory)?;

    let destination = directory.join("table-style-owner.pptx");
    let mut package = Package::new()?;
    let slide = package.presentation_mut()?.add_slide()?;
    slide.set_title("Litchi PPTX table-style owner");
    slide.add_text_box(
        "Typed style catalog",
        914_400,
        4_000_000,
        7_315_200,
        914_400,
    );
    package.save(&destination)?;

    let package = Package::open(&destination)?;
    let mut graph = package.opc()?.clone();
    let conformance = style::conformance(&graph)?;
    let id = Id::parse(TABLE_STYLE)?;
    let mut styles = style::load(&graph)?.unwrap_or_else(|| List::new(conformance, id));
    let mut definition = Def::new(id, "Litchi native smoke")?;
    let expected = Parts::BACKGROUND | Parts::WHOLE | Parts::FIRST_ROW;
    let _ = definition.reset_parts(expected);
    styles.add(definition)?;
    let _ = style::put(&mut graph, styles)?;
    let mut package = Package::from_opc_package(graph)?;
    package.save(&destination)?;

    let reopened = Package::open(&destination)?;
    let styles = style::load(reopened.opc()?)?
        .ok_or_else(|| missing("round-tripped PPTX table-style catalog"))?;
    if !styles
        .get(id)
        .is_some_and(|style| style.parts() == expected)
    {
        return Err(missing("round-tripped PPTX table-style definition").into());
    }

    println!("{}", directory.canonicalize()?.display());
    Ok(())
}

fn missing(value: &'static str) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidData, format!("missing {value}"))
}
