//! Convert an EMF or WMF metafile to SVG.
//!
//! `litchi-imgconv` exposes two public, format-specific SVG entry points:
//!
//!   - [`litchi_imgconv::emf::convert_emf_to_svg`] (`emf/mod.rs`)
//!   - [`litchi_imgconv::wmf::convert_wmf_to_svg`] (`wmf/mod.rs`)
//!
//! Both walk the parsed metafile records and emit a minimal SVG document
//! using the building blocks in the [`litchi_imgconv::svg`] module
//! (`SvgBuilder`, `SvgPath`, `SvgRect`, `SvgEllipse`, `SvgText`,
//! `SvgImage`). Embedded raster blits become base64 `data:image/png` URLs
//! so the resulting SVG is fully self-contained.
//!
//! The PICT decoder does not currently expose a `convert_pict_to_svg`
//! entry point - see `pict/mod.rs`, which only ships PNG/JPEG/WebP
//! converters - so this example only handles EMF and WMF.
//!
//! # Run
//!
//! ```bash
//! # Use the bundled EMF sample:
//! cargo run -p litchi-imgconv --example metafile_to_svg
//!
//! # Or pass any EMF/WMF file:
//! cargo run -p litchi-imgconv --example metafile_to_svg -- \
//!     test-data/images/wmf/santa.wmf santa.svg
//! ```

use std::path::{Path, PathBuf};

use litchi_imgconv::{emf, wmf};

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
        Some("pict" | "pct") => {
            // The pict module currently has no convert_pict_to_svg; raster-only.
            return Err(
                "PICT to SVG conversion is not exposed by litchi-imgconv; \
                 use convert_pict (raster) instead"
                    .into(),
            );
        },
        other => {
            return Err(format!(
                "unsupported extension {:?}; expected .emf or .wmf",
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
