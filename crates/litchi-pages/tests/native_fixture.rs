use std::path::PathBuf;

use litchi_pages::{Package, SectionSelector};

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/pages/basic.pages")
}

#[test]
fn native_pages_fixture_opens_from_path_and_bytes() -> Result<(), Box<dyn std::error::Error>> {
    let path = fixture_path();
    let package = Package::open(&path)?;
    package.validate()?;

    let text = package.text()?;
    assert!(text.contains("Litchi native Pages fixture"));
    assert!(text.contains("Buffa lazy-view migration verification"));
    assert!(text.contains("2026-08-07"));
    assert!(package.stats().total_objects() > 0);
    assert_eq!(package.stats().section_count(), 1);
    let document = package.semantic_document();
    let selected = document
        .select_section(SectionSelector::index(0))?
        .ok_or_else(|| std::io::Error::other("native Pages fixture has no first section"))?;
    assert_eq!(selected.index(), 0);

    let bytes = std::fs::read(path)?;
    let from_bytes = Package::from_bytes(&bytes)?;
    from_bytes.validate()?;
    assert_eq!(from_bytes.text()?, text);
    assert_eq!(from_bytes.stats(), package.stats());
    Ok(())
}
