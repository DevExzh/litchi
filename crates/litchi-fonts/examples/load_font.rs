//! Load a system font and print its metadata.
//!
//! Run with:
//!
//! ```bash
//! cargo run -p litchi-fonts --example load_font
//! ```
//!
//! The example tries a list of common font families and stops at the first one
//! it can resolve on the host system. If none of the candidates exist, it
//! exits cleanly with an explanatory message rather than panicking — useful
//! for CI environments that may not have all fonts installed.

use litchi_fonts::{FontData, FontError, Loader};

/// Common font families that are usually available on at least one of the
/// major desktop / CI platforms (macOS, Windows, Linux distros with the
/// `fontconfig`/`liberation`/`dejavu` packages).
const CANDIDATE_FAMILIES: &[&str] = &[
    "Arial",
    "Helvetica",
    "DejaVu Sans",
    "Liberation Sans",
    "Times New Roman",
    "Sans",
];

fn try_load(loader: &Loader, families: &[&str]) -> Option<FontData> {
    for family in families {
        match loader.load_system_font(family) {
            Ok(font) => {
                println!("Loaded font family: {family}");
                return Some(font);
            },
            Err(FontError::NotFound(name)) => {
                println!("  - '{name}' not found, trying next candidate...");
            },
            Err(err) => {
                println!("  - error loading '{family}': {err}");
            },
        }
    }
    None
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let loader = Loader::new();

    let Some(font) = try_load(&loader, CANDIDATE_FAMILIES) else {
        println!(
            "No candidate fonts could be loaded on this system. \
             Tried: {CANDIDATE_FAMILIES:?}"
        );
        println!("Skipping report. (This is non-fatal — useful for CI without fonts.)");
        return Ok(());
    };

    println!();
    println!("Font report");
    println!("===========");
    println!("name       : {}", font.name);
    println!("data length: {} bytes", font.data.len());
    println!("face index : {}", font.index);

    match &font.properties {
        Some(props) => {
            println!("properties :");
            println!("  license: {:?}", props.license());
            println!("  panose : {:02X?}", props.panose().bytes());
            match props.charset() {
                Some(charset) => println!("  charset: {}", charset.code()),
                None => println!("  charset: <absent or ambiguous>"),
            }
            println!("  family : {:?}", props.family());
            println!("  pitch  : {:?}", props.pitch());
            println!(
                "  sig    : usb={:08X?} csb={:08X?}",
                props.signature().unicode(),
                props.signature().code_pages()
            );
        },
        None => {
            println!("properties : <none extracted — font lacks an OS/2 table or is too short>");
        },
    }

    Ok(())
}
