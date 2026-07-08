//! Convert an EMF or WMF file to PNG.
//!
//! This example demonstrates the direct entry points that take raw metafile
//! bytes (rather than a wrapping BLIP record):
//! [`litchi_imgconv::emf::convert_emf`] and [`litchi_imgconv::wmf::convert_wmf`].
//! The format is dispatched by file-extension; PICT files would be routed to
//! [`litchi_imgconv::pict::convert_pict`] in the same way.
//!
//! # Run
//!
//! ```bash
//! # Use the bundled samples:
//! cargo run -p litchi-imgconv --example convert_emf -- \
//!     test-data/images/emf/wrench.emf wrench.png
//!
//! cargo run -p litchi-imgconv --example convert_emf -- \
//!     test-data/images/wmf/santa.wmf santa.png
//!
//! # Or with no arguments — defaults to test-data/images/emf/wrench.emf:
//! cargo run -p litchi-imgconv --example convert_emf
//! ```

use std::path::{Path, PathBuf};

use image::ImageFormat;
use litchi_imgconv::{emf, pict, wmf};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // First arg: input path (defaults to a bundled EMF sample).
    // Second arg: output PNG path (defaults to alongside the input with .png).
    let mut args = std::env::args().skip(1);
    let input: PathBuf = args
        .next()
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("test-data/images/emf/wrench.emf"));
    let output: PathBuf = args
        .next()
        .map(PathBuf::from)
        .unwrap_or_else(|| input.with_extension("png"));

    if !input.exists() {
        return Err(format!(
            "input file does not exist: {} (run from repo root, or pass an absolute path)",
            input.display()
        )
        .into());
    }

    let bytes = std::fs::read(&input)?;
    println!("loaded {} ({} bytes)", input.display(), bytes.len());

    // Pick the converter by extension. The width parameter (1024) demonstrates
    // aspect-ratio-preserving resize: the converter computes the matching
    // height when only one dimension is supplied.
    let png = match extension_lower(&input).as_deref() {
        Some("emf") => emf::convert_emf(&bytes, ImageFormat::Png, Some(1024), None)?,
        Some("wmf") => wmf::convert_wmf(&bytes, ImageFormat::Png, Some(1024), None)?,
        Some("pict" | "pct") => pict::convert_pict(&bytes, ImageFormat::Png, Some(1024), None)?,
        other => {
            return Err(format!(
                "unsupported extension {:?}; expected .emf, .wmf, or .pict",
                other
            )
            .into());
        },
    };

    std::fs::write(&output, &png)?;
    println!("wrote {} ({} bytes PNG)", output.display(), png.len());
    Ok(())
}

fn extension_lower(p: &Path) -> Option<String> {
    p.extension()
        .and_then(|e| e.to_str())
        .map(|s| s.to_ascii_lowercase())
}
