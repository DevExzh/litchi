//! Read Pages, Numbers, or Keynote through the unified `File`.
//!
//! The example opens one regular package file or app-authored package
//! directory under explicit physical and semantic limits (the iWork leaf
//! options of `OpenOptions`), then prints its outline and text. Native
//! object identifiers and lower-level IWA crates never enter the
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
const MAX_INPUT_BYTES: u64 = 512 * 1024 * 1024;

#[cfg(feature = "iwork")]
fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = env::args_os().skip(1);
    let path = PathBuf::from(arguments.next().ok_or(
        "usage: read_iwork <document.pages|document.numbers|document.key|package-directory>",
    )?);
    if arguments.next().is_some() {
        return Err("unexpected extra argument".into());
    }

    let options = litchi::OpenOptions::new()
        .max_input_bytes(MAX_INPUT_BYTES)
        .pages_options(
            litchi::pages::OpenOptions::new().source_limits(
                litchi::pages::SourceLimits::builder()
                    .max_input_bytes(MAX_INPUT_BYTES)?
                    .build()?,
            ),
        )
        .keynote_options(
            litchi::keynote::OpenOptions::new().source_limits(
                litchi::keynote::SourceLimits::builder()
                    .max_input_bytes(MAX_INPUT_BYTES)?
                    .build()?,
            ),
        )
        .numbers_options(
            litchi::numbers::OpenOptions::new().source_limits(
                litchi::numbers::SourceLimits::builder()
                    .max_input_bytes(MAX_INPUT_BYTES)?
                    .build()?,
            ),
        );
    let file = litchi::open_with(&path, &options)?;

    println!("{:?} document", file.format());
    println!("{}", file.view(litchi::view::Mode::Outline)?);
    println!("{}", file.view(litchi::view::Mode::Stats)?);
    println!("{}", file.text()?);
    Ok(())
}

#[cfg(not(feature = "iwork"))]
fn main() {
    eprintln!("enable the `iwork` feature to run this example");
}
