//! Read Pages, Numbers, or Keynote through the format-neutral iWork facade.
//!
//! The example snapshots one regular package file or app-authored package
//! directory under explicit physical and semantic limits, detects its
//! application from the retained state, and prints only archive-free values.
//! Native object identifiers and lower-level IWA crates never enter the
//! application-facing code.
//!
//! ```text
//! cargo run -p litchi --example read_iwork --features iwork -- document.pages
//! ```

#![allow(
    clippy::print_stderr,
    clippy::print_stdout,
    reason = "this command-line example intentionally renders semantic document content"
)]

#[cfg(feature = "iwork")]
use std::{env, error::Error, path::PathBuf};

#[cfg(feature = "iwork")]
use litchi::iwork::{Document, Options, SnapshotLimits, SourceLimits};

#[cfg(feature = "iwork")]
const MEBIBYTE: u64 = 1024 * 1024;
#[cfg(feature = "iwork")]
const MEBIBYTE_USIZE: usize = 1024 * 1024;
#[cfg(feature = "iwork")]
const MAX_INPUT_BYTES: u64 = 512 * MEBIBYTE;

#[cfg(feature = "iwork")]
fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = env::args_os().skip(1);
    let path = PathBuf::from(arguments.next().ok_or(
        "usage: read_iwork <document.pages|document.numbers|document.key|package-directory>",
    )?);
    if arguments.next().is_some() {
        return Err("unexpected extra argument".into());
    }

    let options = options()?;
    let document = Document::open_with_options(&path, options)?;
    let snapshot = document.snapshot();

    println!("{} document: {}", document.format(), snapshot.summary());

    for table in snapshot.tables() {
        println!(
            "table {} {:?}: {}x{}, {} materialized cell(s)",
            table.position() + 1,
            table.name(),
            table.row_count(),
            table.column_count(),
            table.cell_count(),
        );
    }
    for slide in snapshot.slides() {
        println!(
            "slide {}: name={:?}, title={:?}, skipped={}, builds={}, transition={}",
            slide.position() + 1,
            slide.name(),
            slide.title(),
            slide.is_skipped(),
            slide.build_count(),
            slide.has_transition(),
        );
    }
    for section in snapshot.sections() {
        println!(
            "section {}: kind={:?}, name={:?}, heading={:?}, pages={:?}",
            section.position() + 1,
            section.kind(),
            section.name(),
            section.heading(),
            section.page_count(),
        );
    }

    for text in snapshot.iter_text() {
        println!("text {:?}: {}", text.role(), text.value());
    }

    Ok(())
}

#[cfg(feature = "iwork")]
fn options() -> Result<Options, litchi::iwork::Error> {
    let source = SourceLimits::new(
        MAX_INPUT_BYTES,
        20_000,
        256 * MEBIBYTE,
        1024 * MEBIBYTE,
        256 * MEBIBYTE_USIZE,
    )?;
    let snapshot = SnapshotLimits::new(4_096, 4_096, 4_096, 32 * MEBIBYTE_USIZE)?;
    Ok(Options::new(source, snapshot))
}

#[cfg(not(feature = "iwork"))]
fn main() {
    eprintln!("enable the `iwork` feature to run this example");
}
