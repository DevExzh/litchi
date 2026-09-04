//! Generate DOCX web-settings and glossary artifacts for native Office checks.
//!
//! Run with:
//!
//! ```sh
//! cargo run -p litchi-docx --example docx_owner_native_smoke -- \
//!     target/office-owner-smoke
//! ```

use std::io;
use std::path::{Path, PathBuf};

use litchi_docx::Package;
use litchi_docx::glossary::{
    Catalog, Category, Conformance as GlossaryConformance, Entry, Gallery, Id as GlossaryId,
    Insert, Kind, Name, Props,
};
use litchi_docx::web::{Div, Id as DivId, Screen};

type Result<T> = std::result::Result<T, Box<dyn std::error::Error + Send + Sync>>;

fn main() -> Result<()> {
    let directory = std::env::args_os()
        .nth(1)
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("target/office-owner-smoke"));
    std::fs::create_dir_all(&directory)?;

    create_docx(&directory)?;
    create_glossary_docx(&directory)?;
    let word_round_trip = directory.join("glossary-owner-word.dotx");
    if word_round_trip.exists() {
        verify_glossary_docx(&word_round_trip)?;
    }
    println!("{}", directory.canonicalize()?.display());
    Ok(())
}

fn create_glossary_docx(directory: &Path) -> Result<()> {
    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const BUILDING_BLOCK_TEXT: &str = "Litchi reusable native building block";

    let destination = directory.join("glossary-owner.dotx");
    let mut package = Package::new_template()?;
    package
        .document_mut()?
        .add_heading("Litchi DOCX glossary owner", 1)?;
    package
        .document_mut()?
        .add_paragraph_with_text("This document carries a typed AutoText building block.");

    let body = format!(
        r#"<w:docPartBody xmlns:w="{W}"><w:p><w:r><w:t>{BUILDING_BLOCK_TEXT}</w:t></w:r></w:p></w:docPartBody>"#
    );
    let props = Props {
        category: Some(Category::new("General", Gallery::new("autoTxt")?)?),
        kinds: Kind::NORMAL,
        inserts: Insert::CONTENT,
        description: Some("Litchi native Word verification building block".to_owned()),
        id: Some(GlossaryId::new("{B81EC3D0-09F5-4E43-9D1A-934D27B13F36}")?),
        ..Props::new(Name::new("Litchi AutoText")?)
    };
    let mut catalog = Catalog::new();
    catalog.add(Entry::new("Litchi AutoText", body.into_bytes())?.with_props(props)?)?;
    let _ = package.put_glossary(catalog, GlossaryConformance::Transitional)?;
    package.save(&destination)?;

    verify_glossary_docx(&destination)
}

fn verify_glossary_docx(path: &Path) -> Result<()> {
    const BUILDING_BLOCK_TEXT: &str = "Litchi reusable native building block";

    let reopened = Package::open(path)?;
    let (catalog, conformance) = reopened
        .glossary()?
        .ok_or_else(|| missing("DOCX glossary catalog"))?;
    let stored = catalog
        .get("litchi autotext")?
        .ok_or_else(|| missing("DOCX glossary entry"))?;
    if conformance != GlossaryConformance::Transitional
        || !stored.body().is_some_and(|body| {
            body.windows(BUILDING_BLOCK_TEXT.len())
                .any(|window| window == BUILDING_BLOCK_TEXT.as_bytes())
        })
    {
        return Err(missing("round-tripped DOCX glossary values").into());
    }
    Ok(())
}

fn create_docx(directory: &Path) -> Result<()> {
    let destination = directory.join("web-settings-owner.docx");
    let mut package = Package::new()?;
    package
        .document_mut()?
        .add_heading("Litchi DOCX web-settings owner", 1)?;
    package
        .document_mut()?
        .add_paragraph_with_text("Typed web settings survived package authoring.");

    let (mut settings, conformance) = package.web()?.unwrap_or_default();
    settings
        .set_encoding("utf-8")?
        .set_allow_png(true)
        .set_pixels_per_inch(96)?
        .set_target_screen_size(Screen::Pixels1024x768);
    let body_id = DivId::new(7)?;
    let mut body = Div::new(body_id);
    body.set_body_div(true);
    settings.add(body)?;
    let quote_id = DivId::new(8)?;
    let mut quote = Div::new(quote_id);
    quote.set_block_quote(true);
    settings.add(quote)?;
    let _ = package.put_web(settings, conformance)?;
    package.save(&destination)?;

    let reopened = Package::open(&destination)?;
    let (settings, _) = reopened
        .web()?
        .ok_or_else(|| missing("DOCX web settings"))?;
    if settings.encoding() != Some("utf-8")
        || settings.allow_png() != Some(true)
        || settings.get(body_id)?.is_none()
        || settings.get(quote_id)?.is_none()
    {
        return Err(missing("round-tripped DOCX web-settings values").into());
    }
    Ok(())
}

fn missing(value: &'static str) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidData, format!("missing {value}"))
}
