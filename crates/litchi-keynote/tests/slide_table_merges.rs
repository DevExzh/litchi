//! Selector-first, source-preserving merged-cell reads for Keynote tables.

use std::error::Error as StdError;

use litchi_keynote::{
    Limits, Package, ReadOptions, SemanticLimits, SlideSelector, SlideTableMergesError,
    SlideTableMergesLimitKind, TableSelector, slide::table::merge::Region,
};

const NATIVE_SOURCE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/slide-table-merges-native.key");

fn exact_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn StdError>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

#[test]
fn native_merge_read_uses_selectors_and_preserves_exact_source() -> Result<(), Box<dyn StdError>> {
    let package = Package::from_bytes(NATIVE_SOURCE)?;
    let before = exact_bytes(&package)?;

    let regions = package.slide_table_merges(SlideSelector::index(0), TableSelector::index(0))?;

    assert_eq!(regions, [Region::new(3, 1, 2, 2)?]);
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn merge_read_reports_invalid_table_selector_without_exposing_native_ids()
-> Result<(), Box<dyn StdError>> {
    let package = Package::from_bytes(NATIVE_SOURCE)?;

    let error = package
        .slide_table_merges(SlideSelector::index(0), TableSelector::index(99))
        .expect_err("the native fixture contains only one table");

    assert!(matches!(
        error,
        SlideTableMergesError::TablePositionNotFound { .. }
    ));
    Ok(())
}

#[test]
fn selection_and_merge_decode_share_the_cumulative_reference_budget()
-> Result<(), Box<dyn StdError>> {
    let semantic = SemanticLimits::new(
        SemanticLimits::MAX_OBJECTS,
        SemanticLimits::MAX_SLIDES,
        1,
        SemanticLimits::MAX_TEXT_STORAGES,
        SemanticLimits::MAX_TEXT_FRAGMENTS,
        SemanticLimits::MAX_TEXT_BYTES,
    )?;
    let package = Package::from_bytes_with_options(
        NATIVE_SOURCE,
        ReadOptions::new(Limits::default(), semantic),
    )?;

    let error = package
        .slide_table_merges(SlideSelector::index(0), TableSelector::index(0))
        .expect_err("root selection must consume the single permitted reference");

    assert!(matches!(
        error,
        SlideTableMergesError::LimitExceeded {
            kind: SlideTableMergesLimitKind::References,
            ..
        }
    ));
    Ok(())
}
