//! Convert an EMF or WMF file to PNG.
//!
//! This example demonstrates the bounded, typed raw-metafile API. It dispatches
//! EMF/WMF by extension and always requests PNG raster output.
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

use litchi_imgconv::{InputFormat, Options, OutputFormat, convert_metafile};

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

    let input_format = match extension_lower(&input).as_deref() {
        Some("emf") => InputFormat::Emf,
        Some("wmf") => InputFormat::Wmf,
        other => {
            return Err(format!("unsupported extension {:?}; expected .emf or .wmf", other).into());
        },
    };

    // The width demonstrates aspect-ratio-preserving resizing; the converter
    // computes the matching height and applies its resource limits first.
    let png = convert_metafile(
        &bytes,
        input_format,
        OutputFormat::Png,
        Options::default().width(1024),
    )?;

    std::fs::write(&output, &png.bytes)?;
    println!(
        "wrote {} ({} bytes {})",
        output.display(),
        png.bytes.len(),
        png.mime_type
    );
    Ok(())
}

fn extension_lower(p: &Path) -> Option<String> {
    p.extension()
        .and_then(|e| e.to_str())
        .map(|s| s.to_ascii_lowercase())
}
