//! Author a Word document with one inert HTML alternative-format import.
//!
//! The output path is the first argument, or `target/docx-alt.docx` by
//! default. Litchi stores the HTML bytes and relationship; Microsoft Word is
//! responsible for importing them when it opens the document.

use std::path::PathBuf;

use litchi_docx::Package;
use litchi_docx::alt::{Data, Import};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = std::env::args_os()
        .nth(1)
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("target/docx-alt.docx"));
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent)?;
    }

    let mut package = Package::new()?;
    package
        .document_mut()?
        .add_heading("Litchi altChunk verification", 1)?;

    let html = br#"<!doctype html><html><head><meta charset="utf-8"></head><body><h2>Imported HTML</h2><p>Word imported this move-owned payload.</p></body></html>"#;
    package.add_alt(Import::data(Data::Html(html.to_vec())), Some(true))?;
    package
        .document_mut()?
        .add_paragraph()
        .add_run_with_text("Content after the imported block.");
    package.save(&path)?;

    let reopened = Package::open(&path)?;
    let count = reopened.document()?.alts()?.len();
    if count != 1 {
        return Err(std::io::Error::other(format!(
            "expected one alternative-format anchor, found {count}"
        ))
        .into());
    }

    println!("{}", path.canonicalize()?.display());
    Ok(())
}
