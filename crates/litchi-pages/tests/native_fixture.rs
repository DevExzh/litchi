use std::path::PathBuf;

use litchi_pages::{Package, SectionSelector};

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/pages/basic.pages")
}

#[test]
fn native_unicode_body_sections_open_with_exact_source_preservation()
-> Result<(), Box<dyn std::error::Error>> {
    let source = std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/pages/body-sections-unicode.pages"),
    )?;
    let package = Package::from_bytes(&source)?;
    package.validate()?;
    let text = package.text()?;
    assert!(text.contains("Pages borrowed body 😀"));
    assert!(text.contains("First section marker."));
    assert!(text.contains("Second section marker — end."));
    assert_eq!(package.stats().section_count(), 2);
    assert!(package.select_section(SectionSelector::index(1))?.is_some());
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    assert_eq!(output, source);
    Ok(())
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
    assert_eq!(selected.name(), Some("Blank"));
    assert_eq!(
        package
            .select_section("Blank")?
            .map(litchi_pages::Section::index),
        Some(0)
    );
    assert_eq!(
        document
            .section_named("Blank")?
            .map(litchi_pages::Section::index),
        Some(0)
    );

    let bytes = std::fs::read(path)?;
    let mut streamed = Vec::new();
    package.write_to(&mut streamed)?;
    assert_eq!(streamed, bytes);
    let from_bytes = Package::from_bytes(&bytes)?;
    from_bytes.validate()?;
    assert_eq!(from_bytes.text()?, text);
    assert_eq!(from_bytes.stats(), package.stats());
    Ok(())
}

#[test]
fn native_visible_body_table_reads_empty_and_preserves_exact_noop()
-> Result<(), Box<dyn std::error::Error>> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/pages/body-table-visible.pages");
    let bytes = std::fs::read(&path)?;
    let package = Package::open(&path)?;
    assert!(package.text()?.contains("Pages hidden-axis native oracle"));
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    assert_eq!(output, bytes);
    assert_eq!(
        package.body_table_hidden_axes(0usize)?,
        litchi_pages::table::hidden_axes::HiddenAxes::empty()
    );
    let noop = package
        .edit_body_table_hidden_axes(0usize)?
        .clear()
        .commit()?;
    assert!(noop.patch().is_noop());
    let mut noop_bytes = Vec::new();
    noop.package().write_to(&mut noop_bytes)?;
    assert_eq!(noop_bytes, bytes);
    let replay = package.apply_body_table_hidden_axes(noop.patch())?;
    let inverse = replay
        .package()
        .apply_body_table_hidden_axes(&noop.patch().inverse())?;
    let mut replay_bytes = Vec::new();
    inverse.package().write_to(&mut replay_bytes)?;
    assert_eq!(replay_bytes, bytes);

    let changed = package
        .edit_body_table_hidden_axes(0usize)?
        .set(litchi_pages::table::hidden_axes::HiddenAxes::new([
            litchi_pages::table::hidden_axes::AxisIndex::row(0),
        ])?)
        .commit()?;
    assert_eq!(
        changed.package().body_table_hidden_axes(0usize)?,
        litchi_pages::table::hidden_axes::HiddenAxes::new([
            litchi_pages::table::hidden_axes::AxisIndex::row(0),
        ])?
    );
    let mut changed_bytes = Vec::new();
    changed.package().write_to(&mut changed_bytes)?;
    let reopened = Package::from_bytes(&changed_bytes)?;
    assert_eq!(
        reopened.body_table_hidden_axes(0usize)?,
        litchi_pages::table::hidden_axes::HiddenAxes::new([
            litchi_pages::table::hidden_axes::AxisIndex::row(0),
        ])?
    );
    let restored = changed
        .package()
        .apply_body_table_hidden_axes(&changed.patch().inverse())?;
    let mut restored_bytes = Vec::new();
    restored.package().write_to(&mut restored_bytes)?;
    assert_eq!(restored_bytes, bytes);
    assert_eq!(
        restored.package().body_table_hidden_axes(0usize)?,
        litchi_pages::table::hidden_axes::HiddenAxes::empty()
    );
    let mut source_after_edit = Vec::new();
    package.write_to(&mut source_after_edit)?;
    assert_eq!(source_after_edit, bytes);
    Ok(())
}
