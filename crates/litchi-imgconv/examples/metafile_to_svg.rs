//! Convert an EMF, WMF or PICT metafile to SVG.
//!
//! `litchi-imgconv` exposes three public, format-specific SVG entry points:
//!
//!   - [`litchi_imgconv::emf::convert_emf_to_svg`] (`emf/mod.rs`)
//!   - [`litchi_imgconv::wmf::convert_wmf_to_svg`] (`wmf/mod.rs`)
//!   - [`litchi_imgconv::pict::convert_pict_to_svg`] (`pict/mod.rs`)
//!
//! Each walks the parsed metafile records and emits a bounded SVG document.
//! Unsupported metafile operations are reported by the renderer rather than
//! being replaced with a placeholder raster image.
//!
//! # Run
//!
//! ```bash
//! # Use the bundled EMF sample:
//! cargo run -p litchi-imgconv --example metafile_to_svg
//!
//! # Or pass any EMF/WMF/PICT file:
//! cargo run -p litchi-imgconv --example metafile_to_svg -- \
//!     test-data/images/wmf/santa.wmf santa.svg
//! ```

use std::path::{Path, PathBuf};

use litchi_imgconv::{emf, pict, wmf};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let input: PathBuf = args
        .next()
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("test-data/images/emf/wrench.emf"));
    let output: PathBuf = args
        .next()
        .map(PathBuf::from)
        .unwrap_or_else(|| input.with_extension("svg"));

    if !input.exists() {
        return Err(format!(
            "input file does not exist: {} (run from repo root, or pass an absolute path)",
            input.display()
        )
        .into());
    }

    let bytes = std::fs::read(&input)?;
    println!("loaded {} ({} bytes)", input.display(), bytes.len());

    let svg: String = match extension_lower(&input).as_deref() {
        Some("emf") => emf::convert_emf_to_svg(&bytes)?,
        Some("wmf") => wmf::convert_wmf_to_svg(&bytes)?,
        Some("pict" | "pct") => pict::convert_pict_to_svg(&bytes)?,
        other => {
            return Err(format!(
                "unsupported extension {:?}; expected .emf, .wmf or .pict",
                other
            )
            .into());
        },
    };

    std::fs::write(&output, svg.as_bytes())?;
    println!(
        "wrote {} ({} bytes SVG, {} chars)",
        output.display(),
        svg.len(),
        svg.chars().count()
    );

    // Print the first line of the SVG so the user can see it really is a
    // valid document without having to open it.
    if let Some(first_line) = svg.lines().next() {
        println!("first line       : {}", first_line);
    }
    Ok(())
}

fn extension_lower(p: &Path) -> Option<String> {
    p.extension()
        .and_then(|e| e.to_str())
        .map(|s| s.to_ascii_lowercase())
}
