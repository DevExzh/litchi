use std::path::PathBuf;

use litchi_pages::{
    Package, PackageError, SelectiveTextOptions, SelectiveTextSelector, SourceMetrics,
};

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/pages/basic.pages")
}

#[test]
fn selected_section_text_retains_exact_source_and_only_selected_output()
-> Result<(), Box<dyn std::error::Error>> {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let expected = package
        .sections()
        .first()
        .and_then(|section| section.text_storages().first())
        .map(|storage| storage.text().to_owned())
        .ok_or_else(|| std::io::Error::other("Pages fixture has no first section text"))?;

    let selected = Package::select_section_text(&source, SelectiveTextSelector::section(0))?;
    assert_eq!(selected.text(), expected);
    assert_eq!(selected.selector(), SelectiveTextSelector::section(0));
    assert_eq!(selected.source_bytes(), source.as_slice());
    assert!(selected.source_is_exact());
    assert!(selected.source_metrics().is_none());
    Ok(())
}

#[test]
fn source_metrics_are_explicit_and_content_free() -> Result<(), Box<dyn std::error::Error>> {
    let source = std::fs::read(fixture_path())?;
    let options =
        SelectiveTextOptions::new(litchi_pages::Limits::default())?.with_source_metrics(true);
    let selected = Package::select_section_text_with_options(
        &source,
        SelectiveTextSelector::section(0),
        options,
    )?;
    let metrics: SourceMetrics = selected
        .source_metrics()
        .ok_or_else(|| std::io::Error::other("selective metrics were not enabled"))?;
    assert_eq!(metrics.source_bytes(), source.len());
    assert!(metrics.component_count() > 0);
    assert!(metrics.object_count() >= metrics.selected_message_count());
    assert!(metrics.message_count() >= metrics.selected_message_count());
    assert!(metrics.selected_payload_bytes() > 0);
    assert!(selected.text().len() <= options.max_text_bytes());
    Ok(())
}

#[test]
fn missing_section_is_a_typed_selective_lifecycle_failure() -> Result<(), Box<dyn std::error::Error>>
{
    let source = std::fs::read(fixture_path())?;
    let error = Package::select_section_text(&source, SelectiveTextSelector::section(1))
        .expect_err("fixture has one section");
    assert!(matches!(
        error,
        PackageError::SelectiveSectionNotFound { index: 1 }
    ));
    Ok(())
}
