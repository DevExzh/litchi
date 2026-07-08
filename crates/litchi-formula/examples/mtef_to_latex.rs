//! Convert MTEF (MathType Equation Format) binary data to LaTeX.
//!
//! Run with:
//!
//! ```bash
//! # use the bundled inline sample bytes
//! cargo run -p litchi-formula --example mtef_to_latex --all-features
//!
//! # or pass a path to a file containing raw MTEF bytes
//! cargo run -p litchi-formula --example mtef_to_latex --all-features -- /path/to/equation.mtef
//! ```
//!
//! When invoked without arguments, this example feeds a minimal valid MTEF
//! header (copied from the `litchi-formula` test suite) through
//! [`MtefParser`] to demonstrate version-info extraction and end-to-end
//! conversion via [`mtef_to_latex`]. When a path is provided, the bytes at
//! that path are read and converted instead.

use std::fs;
use std::path::PathBuf;

use litchi_formula::{Formula, MtefParser, mtef_to_latex};

/// A minimal but structurally valid MTEF byte sequence.
///
/// Lifted from `crates/litchi-formula/src/mtef/mod.rs::tests::test_mtef_parser_with_valid_header`.
/// It contains the 28-byte OLE wrapper, an MTEF v5 header identifying
/// MathType on Windows, and just enough body (a SIZE record followed by
/// an END record) to terminate parsing cleanly.
const SAMPLE_MTEF: &[u8] = &[
    // OLE header (28 bytes)
    0x1C, 0x00, // cb_hdr = 28
    0x00, 0x00, 0x02, 0x00, // version = 0x00020000 (little endian)
    0xD3, 0xC2, // format = 0xC2D3
    0x0B, 0x00, 0x00, 0x00, // size = 11 (MTEF header + minimal content)
    0x00, 0x00, 0x00, 0x00, // reserved[0]
    0x00, 0x00, 0x00, 0x00, // reserved[1]
    0x00, 0x00, 0x00, 0x00, // reserved[2]
    0x00, 0x00, 0x00, 0x00, // reserved[3]
    // MTEF header with signature
    0x28, 0x04, 0x6D, 0x74, // signature "(\x04mt"
    0x05, // version = 5
    0x01, // platform = 1 (Windows)
    0x01, // product = 1 (MathType)
    0x01, // version = 1
    0x00, // version_sub = 0
    0x00, // application_key (empty null-terminated string)
    0x00, // inline = 0
    // Minimal MTEF content (SIZE + END tags)
    0x09, // SIZE tag
    0x00, // END tag
];

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let arg = std::env::args().nth(1).map(PathBuf::from);

    let (source_label, bytes): (String, Vec<u8>) = match arg {
        Some(path) => {
            let bytes = fs::read(&path)?;
            if bytes.is_empty() {
                println!("no input: file `{}` is empty", path.display());
                return Ok(());
            }
            (format!("file `{}`", path.display()), bytes)
        },
        None => ("inline sample".to_string(), SAMPLE_MTEF.to_vec()),
    };

    println!("Source        : {source_label}");
    println!("Byte length   : {}", bytes.len());

    // Inspect the parser without running a full conversion first. The
    // arena-tied lifetime means the parser borrows from `bytes`, so we
    // create a dedicated `Formula` for this scope.
    {
        let formula = Formula::new();
        let parser = MtefParser::new(formula.arena(), &bytes);
        println!("Parser valid  : {}", parser.is_valid());
        if let Some((mtef_version, platform, product, version, sub)) = parser.version_info() {
            println!(
                "MTEF header   : version={mtef_version}, platform={platform}, \
                 product={product}, app_version={version}.{sub}"
            );
        } else {
            println!("MTEF header   : (unavailable - data did not pass validation)");
        }
    }

    // Now run the full helper to produce a LaTeX string. The minimal
    // sample contains no glyphs, so the output for it is essentially an
    // empty display-style block - that is intentional and demonstrates
    // that the pipeline runs end-to-end without errors.
    match mtef_to_latex(&bytes) {
        Ok(latex) => {
            if latex.is_empty() {
                println!("LaTeX         : <empty - sample contains no equation body>");
            } else {
                println!("LaTeX         : {latex}");
            }
        },
        Err(e) => {
            println!("LaTeX         : <conversion failed: {e}>");
        },
    }

    Ok(())
}
