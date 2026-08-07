use std::path::PathBuf;

use litchi_keynote::{Package, SlideSelector};

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/keynote/basic.key")
}

fn assert_expected_slide(package: &Package) -> Result<(), Box<dyn std::error::Error>> {
    package.validate()?;
    let show = package.show()?;
    assert_eq!(show.slide_count(), 1);
    let selected = show
        .select_slide(SlideSelector::index(0))?
        .ok_or_else(|| std::io::Error::other("native Keynote fixture has no first slide"))?;
    assert_eq!(selected.index(), 0);
    if let Some(name) = selected.name() {
        assert_eq!(show.select_slide(name)?, Some(selected));
    }

    let text = show.all_text().join("\n");
    assert!(text.contains("Litchi native Keynote fixture"));
    assert!(text.contains("Buffa lazy-view migration verification"));
    assert!(text.contains("2026-08-07"));

    let stats = package.stats()?;
    assert!(stats.total_objects > 0);
    assert_eq!(stats.slide_count, 1);
    Ok(())
}

#[test]
fn native_keynote_fixture_opens_from_path_and_bytes() -> Result<(), Box<dyn std::error::Error>> {
    let path = fixture_path();
    let package = Package::open(&path)?;
    assert_expected_slide(&package)?;

    let bytes = std::fs::read(path)?;
    let from_bytes = Package::from_bytes(&bytes)?;
    assert_expected_slide(&from_bytes)?;
    assert_eq!(from_bytes.show()?, package.show()?);
    Ok(())
}
