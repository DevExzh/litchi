//! Copy an iWork package without rebuilding its ZIP envelope.
//!
//! The example is intentionally a no-op: it opens and writes the package
//! through the focused physical archive owner, then verifies that the output
//! is byte-for-byte identical to the input. Unknown IWA payloads, unsupported
//! ZIP members, and legacy package envelopes remain untouched.

#![allow(
    clippy::print_stdout,
    reason = "This command-line example intentionally reports its verification result."
)]

mod support;

use std::env;
use std::fs::{self, File};
use std::path::Path;
use std::sync::Arc;

use litchi_iwa_archive::Limits;
use litchi_iwa_archive::package::Catalog;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: preserve_iwork <input> <output>")?;
    let output = arguments
        .next()
        .ok_or("usage: preserve_iwork <input> <output>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let input_bytes: Arc<[u8]> = support::read_package(Path::new(&input), Limits::default())?;
    let package = Catalog::from_shared_bytes(Arc::clone(&input_bytes))?;
    package.write_to(File::create(&output)?)?;

    let output_bytes = fs::read(&output)?;
    if input_bytes.as_ref() != output_bytes.as_slice() {
        return Err(format!(
            "preserve-mode output differs from input ({} vs {} bytes)",
            input_bytes.len(),
            output_bytes.len()
        )
        .into());
    }

    println!("preserved {} bytes exactly", input_bytes.len());
    Ok(())
}
