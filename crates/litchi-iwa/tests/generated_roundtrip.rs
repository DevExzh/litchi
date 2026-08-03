//! End-to-end coverage for packages produced without native iWork templates.
//!
//! These tests intentionally exercise both the format detector and the
//! application-specific readers. They provide a deterministic fixture source
//! while native-app verification remains an external, opt-in check.

use std::error::Error;
use std::fs;
use std::io::Cursor;
use std::path::Path;

use litchi_iwa::Document;
use litchi_iwa::detect::{self, Format};
use litchi_iwa::keynote::{KeynoteDocument, KeynoteDocumentBuilder, KeynoteEditor};
use litchi_iwa::numbers::{NumbersDocument, NumbersDocumentBuilder, NumbersEditor};
use litchi_iwa::pages::{PagesDocument, PagesEditor};
use litchi_iwa::registry::Application;
use tempfile::tempdir;

#[test]
fn builders_emit_packages_that_all_public_readers_can_open() -> Result<(), Box<dyn Error>> {
    let directory = tempdir()?;

    let pages_path = directory.path().join("generated.pages");
    PagesEditor::create_with_text("Generated Pages verification document")?.save(&pages_path)?;
    verify_package(&pages_path, Format::Pages)?;

    let numbers_path = directory.path().join("generated.numbers");
    NumbersDocumentBuilder::new()
        .table_name("Verification")
        .table_dimensions(2, 2)
        .build()?
        .save(&numbers_path)?;
    verify_package(&numbers_path, Format::Numbers)?;

    let keynote_path = directory.path().join("generated.key");
    KeynoteDocumentBuilder::new()
        .title("Generated Keynote verification")
        .subtitle("Created without a native template")
        .build()?
        .save(&keynote_path)?;
    verify_package(&keynote_path, Format::Keynote)?;

    Ok(())
}

fn verify_package(path: &Path, expected: Format) -> Result<(), Box<dyn Error>> {
    let bytes = fs::read(path)?;
    let application = match expected {
        Format::Pages => Application::Pages,
        Format::Numbers => Application::Numbers,
        Format::Keynote => Application::Keynote,
    };

    assert_eq!(detect::bytes(&bytes), Some(expected));

    let mut reader = Cursor::new(bytes.as_slice());
    reader.set_position(1);
    assert_eq!(detect::reader(&mut reader), Some(expected));
    assert_eq!(reader.position(), 1);
    assert_eq!(detect::path(path), Some(expected));

    let document = Document::open(path)?;
    assert_eq!(document.application(), application);
    assert_eq!(document.stats().application, application);
    assert!(document.stats().total_objects > 0);
    document.text()?;
    document.extract_structured_data()?;

    let document_from_bytes = Document::from_bytes(&bytes)?;
    assert_eq!(document_from_bytes.application(), application);
    assert!(document_from_bytes.stats().total_objects > 0);

    match expected {
        Format::Pages => {
            PagesEditor::open(path)?;
            PagesDocument::open(path)?;
        },
        Format::Numbers => {
            NumbersEditor::open(path)?;
            NumbersDocument::open(path)?;
        },
        Format::Keynote => {
            KeynoteEditor::open(path)?;
            KeynoteDocument::open(path)?;
        },
    }

    Ok(())
}
