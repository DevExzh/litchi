#![cfg(feature = "pages")]

use std::io;
use std::path::PathBuf;

use litchi::Document;
use litchi::pages::document_settings::{
    Commit, Diagnostics, Edit, Error, LimitKind, Patch, Settings,
};
use litchi::pages::{
    Package, PageLayoutCommit, PageLayoutDiagnostics, PageLayoutEdit, PageLayoutError,
    PageLayoutLimitKind, PageLayoutPatch, SectionSelector, SectionTextCommit,
    SectionTextDiagnostics, SectionTextEdit, SectionTextError, SectionTextLimitKind,
    SectionTextPatch, TextPosition, TextSpan,
    page_layout::{Layout, Orientation},
    section::{PageNumber, PageNumbering, Start},
};

fn assert_send_sync<T: Send + Sync>() {}

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/pages/basic.pages")
}

#[test]
fn body_storage_reaches_facade_paragraphs() -> Result<(), Box<dyn std::error::Error>> {
    let path = fixture_path();
    let native = Package::open(&path)?;
    let section = native
        .sections()
        .first()
        .ok_or_else(|| io::Error::other("native Pages file has no body section"))?;

    // This fixture reproduces the facade bug: Pages projects native body text
    // into a rich-text storage rather than the legacy paragraph collection.
    assert!(section.heading().is_none());
    assert!(section.paragraphs().is_empty());
    assert_eq!(section.text_storages().len(), 1);
    let expected = section
        .text_storages()
        .first()
        .ok_or_else(|| io::Error::other("native Pages body has no text storage"))?
        .text();
    assert!(!expected.is_empty());

    let document = Document::open(path)?;
    assert_eq!(document.text()?, expected);
    assert_eq!(document.paragraph_count()?, 1);

    let paragraphs = document.paragraphs()?;
    assert_eq!(paragraphs.len(), 1);
    assert_eq!(
        paragraphs
            .first()
            .ok_or_else(|| io::Error::other("facade returned no Pages paragraph"))?
            .text()?,
        expected
    );

    let elements = document.elements()?;
    assert_eq!(elements.len(), 1);
    assert_eq!(
        elements
            .first()
            .and_then(litchi::DocumentElement::as_paragraph)
            .ok_or_else(|| io::Error::other("facade returned no Pages paragraph element"))?
            .text()?,
        expected
    );

    Ok(())
}

#[test]
fn section_name_transaction_reaches_pages_facade() -> Result<(), Box<dyn std::error::Error>> {
    let package = Package::open(fixture_path())?;
    let source_pointer = package.source_bytes().as_ptr();

    let mut noop = package.edit_section_name(SectionSelector::index(0))?;
    noop.set_name(Some("Blank"))?;
    let noop = noop.commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().source_bytes().as_ptr(), source_pointer);

    let mut edit = package.edit_section_name(SectionSelector::name("Blank"))?;
    edit.set_name(Some("Facade Section"))?;
    let commit = edit.commit()?;
    assert_eq!(
        commit.package().sections()[0].name(),
        Some("Facade Section")
    );
    let restored = commit
        .package()
        .apply_section_name(&commit.patch().inverse())?;
    assert_eq!(restored.package().source_bytes(), package.source_bytes());
    Ok(())
}

#[test]
fn section_pagination_transaction_reaches_pages_facade() -> Result<(), Box<dyn std::error::Error>> {
    let package = Package::open(fixture_path())?;
    let source_pointer = package.source_bytes().as_ptr();

    let mut noop = package.edit_section_pagination(SectionSelector::index(0))?;
    noop.set_start(Some(Start::NextPage))?;
    noop.set_page_numbering(Some(PageNumbering::ContinueFromPrevious))?;
    noop.set_starting_page_number(Some(PageNumber::new(1)?));
    let noop = noop.commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().source_bytes().as_ptr(), source_pointer);

    let mut edit = package.edit_section_pagination(SectionSelector::name("Blank"))?;
    edit.set_start(Some(Start::LeftPage))?;
    edit.set_page_numbering(Some(PageNumbering::Restart))?;
    edit.set_starting_page_number(Some(PageNumber::new(17)?));
    let commit = edit.commit()?;
    let changed = commit
        .package()
        .section_pagination(SectionSelector::index(0))?;
    assert_eq!(changed.start(), Some(Start::LeftPage));
    assert_eq!(changed.page_numbering(), Some(PageNumbering::Restart));
    assert_eq!(
        changed.starting_page_number().map(PageNumber::get),
        Some(17)
    );

    let restored = commit
        .package()
        .apply_section_pagination(&commit.patch().inverse())?;
    assert_eq!(restored.package().source_bytes(), package.source_bytes());
    Ok(())
}

#[test]
fn section_text_transaction_reaches_pages_facade() -> Result<(), Box<dyn std::error::Error>> {
    assert_send_sync::<SectionTextCommit>();
    assert_send_sync::<SectionTextDiagnostics>();
    assert_send_sync::<SectionTextEdit<'static>>();
    assert_send_sync::<SectionTextError>();
    assert_send_sync::<SectionTextLimitKind>();
    assert_send_sync::<SectionTextPatch>();
    assert_send_sync::<TextPosition>();

    let package = Package::open(fixture_path())?;
    let selector = SectionSelector::index(0);
    let original = package.section_text(selector)?.to_owned();
    let source_pointer = package.source_bytes().as_ptr();

    let mut noop = package.edit_section_text(selector)?;
    noop.set(&original)?;
    let noop = noop.commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().source_bytes().as_ptr(), source_pointer);

    let prefix = "Facade section: ";
    let mut edit = package.edit_section_text(SectionSelector::name("Blank"))?;
    edit.replace(TextSpan::from_utf16_indexes(0, 0)?, prefix)?;
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    assert_eq!(
        commit.package().section_text(SectionSelector::index(0))?,
        format!("{prefix}{original}")
    );

    let inverse = commit.patch().inverse();
    let restored = commit.package().apply_section_text(&inverse)?;
    assert_eq!(restored.package().source_bytes(), package.source_bytes());
    Ok(())
}

#[test]
fn page_layout_transaction_reaches_pages_facade() -> Result<(), Box<dyn std::error::Error>> {
    assert_send_sync::<Layout>();
    assert_send_sync::<PageLayoutCommit>();
    assert_send_sync::<PageLayoutDiagnostics>();
    assert_send_sync::<PageLayoutEdit<'static>>();
    assert_send_sync::<PageLayoutError>();
    assert_send_sync::<PageLayoutLimitKind>();
    assert_send_sync::<PageLayoutPatch>();

    let package = Package::open(fixture_path())?;
    let source_bytes = package.source_bytes();
    let source_pointer = source_bytes.as_ptr();
    let before = package.page_layout()?;
    assert!(before.page_width().is_some_and(|width| width > 0.0));
    assert!(before.page_height().is_some_and(|height| height > 0.0));

    let replacement_orientation = if before.orientation() == Some(Orientation::Landscape) {
        Orientation::Portrait
    } else {
        Orientation::Landscape
    };
    let mut after = before;
    after.set_orientation(Some(replacement_orientation))?;
    assert_ne!(after, before);

    let mut edit = package.edit_page_layout()?;
    assert_eq!(edit.layout(), before);
    edit.set_layout(after)?;
    let edit_debug = format!("{edit:?}");
    assert!(edit_debug.contains("PageLayoutEdit"));
    assert!(!edit_debug.contains("Index/"));
    assert!(!edit_debug.contains(".iwa"));
    assert!(!edit_debug.contains("identifier"));

    let changed = edit.commit()?;
    assert_eq!(changed.patch().before(), before);
    assert_eq!(changed.patch().after(), after);
    assert_eq!(changed.package().page_layout()?, after);
    assert!(changed.diagnostics().changed());
    assert!(changed.diagnostics().touched_components() >= 1);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_eq!(package.source_bytes().as_ptr(), source_pointer);
    assert_eq!(package.source_bytes(), source_bytes);
    assert_ne!(changed.package().source_bytes(), source_bytes);

    let patch_debug = format!("{:?}", changed.patch());
    assert!(patch_debug.contains("PageLayoutPatch"));
    assert!(!patch_debug.contains("Index/"));
    assert!(!patch_debug.contains(".iwa"));
    assert!(!patch_debug.contains("identifier"));
    assert!(!patch_debug.contains("fingerprint"));

    let restored = changed
        .package()
        .apply_page_layout(&changed.patch().inverse())?;
    assert_eq!(restored.package().source_bytes(), source_bytes);
    assert_eq!(restored.package().page_layout()?, before);
    Ok(())
}

#[test]
fn document_settings_transaction_reaches_pages_facade() -> Result<(), Box<dyn std::error::Error>> {
    assert_send_sync::<Settings>();
    assert_send_sync::<Edit<'static>>();
    assert_send_sync::<Patch>();
    assert_send_sync::<Commit>();
    assert_send_sync::<Diagnostics>();
    assert_send_sync::<Error>();
    assert_send_sync::<LimitKind>();

    let package = Package::open(fixture_path())?;
    let source_bytes = package.source_bytes();
    let source_pointer = source_bytes.as_ptr();
    let before = package.document_settings()?;

    let noop = package.edit_document_settings()?.set(before).commit()?;
    assert_eq!(noop.patch().before(), before);
    assert_eq!(noop.patch().after(), before);
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert!(!noop.diagnostics().full_reparse_performed());
    assert_eq!(noop.package().source_bytes().as_ptr(), source_pointer);

    let mut options = before.options();
    options.set_automatic_hyphenation(Some(!options.uses_automatic_hyphenation()));
    let mut after = before;
    after.set_options(options);
    assert_ne!(after, before);

    let changed = package.edit_document_settings()?.set(after).commit()?;
    assert_eq!(changed.patch().before(), before);
    assert_eq!(changed.patch().after(), after);
    assert_eq!(changed.package().document_settings()?, after);
    assert!(changed.diagnostics().changed());
    assert!(changed.diagnostics().touched_components() >= 1);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_eq!(package.source_bytes().as_ptr(), source_pointer);
    assert_eq!(package.source_bytes(), source_bytes);
    assert_ne!(changed.package().source_bytes(), source_bytes);

    let restored = changed
        .package()
        .apply_document_settings(&changed.patch().inverse())?;
    assert_eq!(restored.package().source_bytes(), source_bytes);
    assert_eq!(restored.package().document_settings()?, before);
    Ok(())
}
