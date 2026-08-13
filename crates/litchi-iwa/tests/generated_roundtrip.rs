//! End-to-end coverage for packages produced without native iWork templates.
//!
//! These tests intentionally exercise both the format detector and the
//! application-specific readers. They provide a deterministic fixture source
//! while native-app verification remains an external, opt-in check.

use std::error::Error;
use std::fs;
use std::io::Cursor;
use std::path::Path;
use std::sync::Arc;

use litchi_iwa::Document;
use litchi_iwa::application::Application;
use litchi_iwa::detect::{self, Format};
use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
use litchi_iwa::numbers::{NumbersDocument, NumbersDocumentBuilder, NumbersEditor};
use litchi_iwa::pages::{PagesDocument, PagesEditor};
use litchi_keynote::Package as KeynotePackage;
use tempfile::tempdir;

fn assert_send_sync<T: Send + Sync>() {}

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

    assert_eq!(detect::bytes(&bytes)?, Some(expected));

    let mut reader = Cursor::new(bytes.as_slice());
    reader.set_position(1);
    assert_eq!(detect::reader(&mut reader)?, Some(expected));
    assert_eq!(reader.position(), 1);
    assert_eq!(detect::path(path)?, Some(expected));

    let document = Document::open(path)?;
    assert_send_sync::<litchi_iwa::raw::bundle::Bundle>();
    assert_send_sync::<litchi_iwa::Document>();
    assert_send_sync::<PagesDocument>();
    assert_send_sync::<NumbersDocument>();
    assert_send_sync::<KeynotePackage>();
    assert_eq!(document.application(), application);
    let document_stats = document.stats()?;
    assert_eq!(document_stats.application, application);
    assert!(document_stats.total_objects > 0);
    let document_snapshot = document.snapshot();
    assert_eq!(document_snapshot.application(), document.application());
    let document_snapshot_stats = document_snapshot.stats()?;
    assert_eq!(
        document_snapshot_stats.total_objects,
        document_stats.total_objects
    );
    document.validate()?;
    document.text()?;

    let document_from_bytes = Document::from_bytes(&bytes)?;
    assert_eq!(document_from_bytes.application(), application);
    assert!(document_from_bytes.stats()?.total_objects > 0);

    match expected {
        Format::Pages => {
            PagesEditor::open(path)?;
            let specialized = PagesDocument::open(path)?;
            let snapshot = specialized.snapshot();
            let specialized_stats = specialized.stats()?;
            let snapshot_stats = snapshot.stats()?;
            assert_eq!(
                snapshot_stats.total_objects,
                specialized_stats.total_objects
            );
        },
        Format::Numbers => {
            NumbersEditor::open(path)?;
            let specialized = NumbersDocument::open(path)?;
            let snapshot = specialized.snapshot();
            let specialized_stats = specialized.stats()?;
            let snapshot_stats = snapshot.stats()?;
            assert_eq!(
                snapshot_stats.total_objects,
                specialized_stats.total_objects
            );
            let semantic_sheets = specialized.semantic_sheets();
            let snapshot_sheets = snapshot.semantic_sheets();
            assert!(Arc::ptr_eq(&semantic_sheets, &snapshot_sheets));
            assert_eq!(semantic_sheets.len(), specialized_stats.sheet_count);
            assert_eq!(
                semantic_sheets
                    .iter()
                    .map(litchi_numbers::Sheet::table_count)
                    .sum::<usize>(),
                specialized_stats.table_count
            );
        },
        Format::Keynote => {
            KeynoteEditor::open(path)?;
            let specialized = KeynotePackage::open(path)?;
            let snapshot = specialized.snapshot();
            let specialized_stats = specialized.stats()?;
            let snapshot_stats = snapshot.stats()?;
            assert_eq!(
                snapshot_stats.total_objects,
                specialized_stats.total_objects
            );
            specialized.validate()?;
            assert_eq!(
                specialized.show()?.slide_count(),
                specialized_stats.slide_count
            );
            assert_eq!(
                specialized.semantic_snapshot()?.slides().len(),
                specialized_stats.slide_count
            );
            assert!(!specialized.text()?.is_empty());
        },
    }

    Ok(())
}
