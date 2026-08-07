//! Copy an iWork package without rebuilding its ZIP envelope.
//!
//! The example is intentionally a no-op: it opens and saves the package, then
//! verifies that the output is byte-for-byte identical to the input. This is
//! useful for validating preserve-mode ingress and egress on real iWork files.

use std::env;
use std::fs;

use litchi_iwa::raw::package::IWorkPackage;

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

    let input_bytes = fs::read(&input)?;
    let package = IWorkPackage::open(&input)?;
    package.save(&output)?;
    let output_bytes = fs::read(&output)?;
    if input_bytes != output_bytes {
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
