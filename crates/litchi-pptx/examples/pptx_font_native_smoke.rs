//! Re-author the checked-in Boldonse graph for native PowerPoint verification.
//!
//! Run with:
//!
//! ```sh
//! cargo run -p litchi-pptx --example pptx_font_native_smoke -- \
//!     target/office-verification
//! ```

use std::io;
use std::path::{Path, PathBuf};

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_pptx::font::{Face, Font, Fonts, Format, Style};
use litchi_pptx::{Package, font};
use sha2::{Digest, Sha256};

type Result<T> = std::result::Result<T, Box<dyn std::error::Error + Send + Sync>>;

const SOURCE: &str = "test-data/libreoffice-core/sd/qa/unit/data/BoldonseFontEmbedded.pptx";
const OUTPUT: &str = "pptx-font-crud-generated.pptx";
const FONT_NAME: &str = "Boldonse";
const FONT_BYTES: usize = 36_187;
const FONT_SHA256: [u8; 32] = [
    0x7b, 0x21, 0xa8, 0x2e, 0xf5, 0x34, 0xb4, 0xa1, 0x85, 0x28, 0x7d, 0xeb, 0x48, 0xc6, 0xb8, 0x7b,
    0xe0, 0xac, 0xd6, 0xd8, 0xd1, 0x2b, 0xb3, 0xac, 0x66, 0xdd, 0x31, 0x6a, 0xaf, 0xe9, 0x59, 0xe4,
];
const FONT_DATA: &str = "application/x-fontdata";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Payload {
    Exact,
    Office,
}

fn main() -> Result<()> {
    let args = std::env::args_os().skip(1).collect::<Vec<_>>();
    if args.first().and_then(|value| value.to_str()) == Some("--verify-office") {
        let path = args
            .get(1)
            .map(PathBuf::from)
            .ok_or_else(|| missing("PowerPoint-saved verification path"))?;
        verify(&path, b"Test Test", Payload::Office)?;
        println!("{}", path.canonicalize()?.display());
        return Ok(());
    }
    let directory = args
        .first()
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("target/office-verification"));
    std::fs::create_dir_all(&directory)?;

    let package = Package::open(SOURCE)?;
    let mut graph = package.opc()?.clone();
    let source = font::load(&graph)?.ok_or_else(|| missing("source embedded-font collection"))?;
    let source_font = source.get("boldonse")?;
    let source_face = source_font
        .get(Style::Regular)
        .ok_or_else(|| missing("source regular Boldonse face"))?;
    let data = source_face.data().clone();

    let mut font = Font::from_face(FONT_NAME, Face::new(Style::Regular, data.clone()))?;
    if let Some(value) = source_font.panose() {
        font = font.with_panose(value);
    }
    if let Some(value) = source_font.pitch_family() {
        font = font.with_pitch_family(value);
    }
    if let Some(value) = source_font.charset() {
        font = font.with_charset(value);
    }

    // Exercise semantic add/get/replace/remove and checked numeric reorder on a
    // detached value. Font programs stay Arc-shared throughout the CRUD path.
    let mut fonts = Fonts::new();
    fonts.add(font)?;
    let _ = fonts.get("BOLDONSE")?;
    fonts.add(Font::from_face(
        "Delete Probe",
        Face::new(Style::Regular, data.clone()),
    )?)?;
    let _ = fonts.replace(
        "delete probe",
        Font::from_face(
            "Delete Probe Updated",
            Face::new(Style::Regular, data.clone()),
        )?,
    )?;
    let _ = fonts.remove("DELETE PROBE UPDATED")?;
    fonts.reorder(&[0usize])?;

    let removed =
        font::remove(&mut graph)?.ok_or_else(|| missing("removed source font collection"))?;
    if removed.len() != 1 || !font::put(&mut graph, fonts)? {
        return Err(missing("fresh font graph publication").into());
    }
    let mut package = Package::from_opc_package(graph)?;
    let unchanged =
        font::load(package.opc()?)?.ok_or_else(|| missing("published font collection"))?;
    let mut no_op_graph = package.opc()?.clone();
    if font::put(&mut no_op_graph, unchanged)? {
        return Err(invalid("semantic no-op unexpectedly changed package").into());
    }

    let destination = directory.join(OUTPUT);
    package.save(&destination)?;
    verify(&destination, b">Test<", Payload::Exact)?;
    println!("{}", destination.canonicalize()?.display());
    Ok(())
}

fn verify(path: &Path, expected_text: &[u8], payload: Payload) -> Result<()> {
    let package = Package::open(path)?;
    let fonts = font::load(package.opc()?)?
        .ok_or_else(|| missing("round-tripped embedded-font collection"))?;
    if fonts.len() != 1 {
        return Err(invalid("expected exactly one embedded typeface").into());
    }
    let font = fonts.get("BOLDONSE")?;
    let face = font
        .get(Style::Regular)
        .ok_or_else(|| missing("round-tripped regular face"))?;
    if font.faces().len() != 1 || face.data().format() != Format::PowerPoint {
        return Err(invalid("unexpected embedded face profile").into());
    }
    let digest: [u8; 32] = Sha256::digest(face.data().bytes()).into();
    match payload {
        Payload::Exact if face.data().bytes().len() != FONT_BYTES || digest != FONT_SHA256 => {
            return Err(invalid("generated Boldonse payload changed").into());
        },
        Payload::Office => validate_office_payload(face.data().bytes())?,
        Payload::Exact => {},
    }

    let opc = package.opc()?;
    let mut font_parts = opc
        .iter_parts()
        .filter(|part| part.content_type() == FONT_DATA);
    let part = font_parts
        .next()
        .ok_or_else(|| missing("PowerPoint font-data part"))?;
    if font_parts.next().is_some() || part.rels().iter().next().is_some() {
        return Err(invalid("font part graph is not singular and inert").into());
    }
    let presentation = opc.main_document_part()?;
    if presentation
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == rt::FONT)
        .count()
        != 1
    {
        return Err(invalid("presentation does not own exactly one font relationship").into());
    }
    let presentation_xml = presentation.blob();
    let font_list =
        find(presentation_xml, b"<p:embeddedFontLst").ok_or_else(|| missing("embeddedFontLst"))?;
    let defaults = find(presentation_xml, b"<p:defaultTextStyle")
        .ok_or_else(|| missing("defaultTextStyle"))?;
    if !presentation_xml
        .windows(b"embedTrueTypeFonts=\"1\"".len())
        .any(|window| window == b"embedTrueTypeFonts=\"1\"")
        || font_list >= defaults
    {
        return Err(invalid("presentation embedding flag or schema order is invalid").into());
    }
    let slide_is_visible = opc
        .iter_parts()
        .filter(|part| part.content_type() == ct::PML_SLIDE)
        .any(|slide| {
            find(slide.blob(), expected_text).is_some()
                && find(slide.blob(), b"typeface=\"Boldonse\"").is_some()
        });
    if !slide_is_visible {
        return Err(missing("visible Boldonse test text").into());
    }
    Ok(())
}

fn validate_office_payload(bytes: &[u8]) -> Result<()> {
    let size = bytes
        .get(..4)
        .and_then(|value| <[u8; 4]>::try_from(value).ok())
        .map(u32::from_le_bytes)
        .ok_or_else(|| invalid("PowerPoint font payload has no EOT header"))?;
    if usize::try_from(size).ok() != Some(bytes.len()) {
        return Err(invalid("PowerPoint font payload has an invalid EOT size").into());
    }
    Ok(())
}

fn find(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}

fn missing(value: &'static str) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidData, format!("missing {value}"))
}

fn invalid(value: &'static str) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidData, value)
}
