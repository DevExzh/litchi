#![cfg(feature = "pages")]

use std::io;
use std::path::PathBuf;

use litchi::Document;
use litchi::pages::{
    Package, SectionSelector,
    section::{PageNumber, PageNumbering, Start},
};

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
