//! Demonstrate glyph collection and font subsetting.
//!
//! Run with:
//!
//! ```bash
//! cargo run -p litchi-fonts --example subset_glyphs
//! ```
//!
//! This example:
//!   1. Loads a system font (with graceful fallbacks),
//!   2. Implements `CollectGlyphs` on a tiny in-memory document type to
//!      show how callers integrate with the trait,
//!   3. Builds a typed `Glyphs` set from the sample's Unicode scalars,
//!   4. Sends the code points through the canonical bounded preparation API,
//!      which maps them and runs the checked subsetter to produce a smaller
//!      font blob.
//!
//! Like `load_font.rs`, the example exits cleanly when no candidate font
//! is available on the host so it remains usable in CI.

use litchi_fonts::{
    CollectGlyphs, FontData, FontError, GlyphMap, Loader, Request,
    embedding::{Mode, PreparationLimits, prepare_with},
};

const CANDIDATE_FAMILIES: &[&str] = &[
    "Arial",
    "Helvetica",
    "DejaVu Sans",
    "Liberation Sans",
    "Times New Roman",
    "Sans",
];

/// A trivial document representation: each entry is a `(font name, text run)`
/// pair. Implementing `CollectGlyphs` shows the shape of the trait that
/// real document parsers (docx, pptx, ...) provide.
struct SimpleDocument<'a> {
    runs: Vec<(&'a str, &'a str)>,
}

impl CollectGlyphs for SimpleDocument<'_> {
    fn collect_glyphs(&self) -> GlyphMap {
        let mut out = GlyphMap::new();
        for (font_name, text) in &self.runs {
            let bitmap = out.entry(Request::regular(*font_name)).or_default();
            for ch in text.chars() {
                bitmap.insert(ch);
            }
        }
        out
    }
}

fn try_load(loader: &Loader, families: &[&str]) -> Option<FontData> {
    for family in families {
        match loader.load_system_font(family) {
            Ok(font) => {
                println!("Loaded font family: {family}");
                return Some(font);
            },
            Err(FontError::NotFound(_)) => continue,
            Err(err) => {
                println!("  - error loading '{family}': {err}");
            },
        }
    }
    None
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 1. Demonstrate `CollectGlyphs` on a fake document, regardless of
    //    whether any system font is installed. This part always runs.
    let doc = SimpleDocument {
        runs: vec![
            ("BodyFont", "The quick brown fox jumps over the lazy dog."),
            ("BodyFont", "Sphinx of black quartz, judge my vow."),
            ("HeadingFont", "Hello, world!"),
        ],
    };
    let glyph_map = doc.collect_glyphs();
    println!("CollectGlyphs result:");
    for (request, bitmap) in &glyph_map {
        println!(
            "  font '{font_name}' -> {} unique code points",
            bitmap.len(),
            font_name = request.family(),
        );
    }

    // 2. Try to load a real system font for the subsetting demo.
    let loader = Loader::new();
    let Some(font) = try_load(&loader, CANDIDATE_FAMILIES) else {
        println!();
        println!(
            "No candidate fonts could be loaded on this system; \
             skipping subsetting demo. Tried: {CANDIDATE_FAMILIES:?}"
        );
        return Ok(());
    };

    println!();
    println!("Original font data: {} bytes", font.data.len());

    // 3. Build a small set of code points to keep. The canonical preparation
    // path maps these to glyph IDs and retains `.notdef` automatically.
    let mut used_glyphs = GlyphMap::new();
    used_glyphs.insert(
        Request::regular(font.name.clone()),
        "The quick brown fox jumps over the lazy dog."
            .chars()
            .collect(),
    );

    // 4. Run the canonical preparation API. Its selected limits scope source
    // parsing, mapping, and subset output in one checked operation.
    match prepare_with(
        &loader,
        used_glyphs,
        Mode::Subset,
        PreparationLimits::default(),
    ) {
        Ok(prepared) => {
            for subset_font in prepared {
                println!(
                    "Subset font: {} bytes ({:.1}% of original)",
                    subset_font.data.len(),
                    100.0 * (subset_font.data.len() as f64) / (font.data.len() as f64),
                );
            }
        },
        Err(err) => {
            println!("Subsetting failed (this can happen with CFF or unusual fonts): {err}");
        },
    }

    Ok(())
}
