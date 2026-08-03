//! Generate DOCX web-settings and PPTX table-style artifacts for native Office checks.
//!
//! Run with:
//!
//! ```sh
//! cargo run -p litchi-ooxml --example owner_native_smoke --all-features -- \
//!     target/office-owner-smoke
//! ```

use std::io;
use std::path::{Path, PathBuf};

use litchi_docx::web::{Div, Id as DivId, Screen};
use litchi_ooxml::{docx, pptx};
use litchi_pptx::table::style::{Def, Id, Parts};

type Result<T> = std::result::Result<T, Box<dyn std::error::Error + Send + Sync>>;

const TABLE_STYLE: &str = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";

fn main() -> Result<()> {
    let directory = std::env::args_os()
        .nth(1)
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("target/office-owner-smoke"));
    std::fs::create_dir_all(&directory)?;

    create_docx(&directory)?;
    create_pptx(&directory)?;

    println!("{}", directory.canonicalize()?.display());
    Ok(())
}

fn create_docx(directory: &Path) -> Result<()> {
    let destination = directory.join("web-settings-owner.docx");
    let mut package = docx::Package::new()?;
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

    let reopened = docx::Package::open(&destination)?;
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

fn create_pptx(directory: &Path) -> Result<()> {
    let destination = directory.join("table-style-owner.pptx");
    let mut package = pptx::Package::new()?;
    let slide = package.presentation_mut()?.add_slide()?;
    slide.set_title("Litchi PPTX table-style owner");
    slide.add_table(
        vec![
            vec!["Capability".to_owned(), "Status".to_owned()],
            vec!["Typed style catalog".to_owned(), "Verified".to_owned()],
        ],
        914_400,
        4_000_000,
        7_315_200,
        1_828_800,
    );

    let id = Id::parse(TABLE_STYLE)?;
    let mut styles = package
        .styles()?
        .ok_or_else(|| missing("PPTX table-style catalog"))?;
    let mut definition = Def::new(id, "Litchi native smoke")?;
    let expected = Parts::BACKGROUND | Parts::WHOLE | Parts::FIRST_ROW;
    let _ = definition.reset_parts(expected);
    styles.add(definition)?;
    let _ = package.put_styles(styles)?;
    package.save(&destination)?;

    let reopened = pptx::Package::open(&destination)?;
    let styles = reopened
        .styles()?
        .ok_or_else(|| missing("round-tripped PPTX table-style catalog"))?;
    if !styles
        .get(id)
        .is_some_and(|style| style.parts() == expected)
    {
        return Err(missing("round-tripped PPTX table-style definition").into());
    }
    Ok(())
}

fn missing(value: &'static str) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidData, format!("missing {value}"))
}
