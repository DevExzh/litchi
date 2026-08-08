#![cfg(feature = "keynote")]

use std::io;
use std::path::PathBuf;

use litchi::keynote::{Package, Position, SlideSelector};

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/keynote/basic.key")
}

#[test]
fn slide_skip_state_is_available_through_the_root_facade() -> Result<(), Box<dyn std::error::Error>>
{
    let package = Package::open(fixture_path())?;
    let is_skipped = package
        .slides()?
        .first()
        .ok_or_else(|| io::Error::other("native Keynote file has no slide"))?
        .is_skipped();

    let mut edit = package.edit();
    edit.set_slide_skipped(SlideSelector::index(0), is_skipped)?;
    let commit = edit.commit()?;

    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.package().source_bytes(), package.source_bytes());
    Ok(())
}

#[test]
fn slide_order_transaction_is_available_through_the_root_facade()
-> Result<(), Box<dyn std::error::Error>> {
    let package = Package::open(fixture_path())?;
    let source_pointer = package.source_bytes().as_ptr();
    let mut edit = package.edit_slide_order();
    edit.move_slide(SlideSelector::index(0), Position::new(0))?;
    let commit = edit.commit()?;

    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.package().source_bytes().as_ptr(), source_pointer);

    let reapplied = package.apply_slide_order(commit.patch())?;
    assert!(reapplied.patch().is_noop());
    assert_eq!(reapplied.package().source_bytes().as_ptr(), source_pointer);
    Ok(())
}

#[test]
fn show_settings_transaction_is_available_through_the_root_facade()
-> Result<(), Box<dyn std::error::Error>> {
    let package = Package::open(fixture_path())?;
    assert_eq!(package.show_settings()?, *package.show()?.settings());
    let source_pointer = package.source_bytes().as_ptr();
    let mut edit = package.edit_show_settings()?;
    let settings = *edit.settings();
    edit.set_settings(settings)?;
    let commit = edit.commit()?;

    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.package().source_bytes().as_ptr(), source_pointer);

    let reapplied = package.apply_show_settings(commit.patch())?;
    assert!(reapplied.patch().is_noop());
    assert!(!reapplied.diagnostics().changed());
    assert_eq!(reapplied.package().source_bytes().as_ptr(), source_pointer);
    Ok(())
}
