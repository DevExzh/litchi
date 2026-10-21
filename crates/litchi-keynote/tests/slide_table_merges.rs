//! Selector-first, source-preserving merged-cell reads for Keynote tables.

use std::error::Error as StdError;

use litchi_keynote::{
    OpenOptions, Package, SemanticLimits, SlideSelector, TableSelector,
    slide::table::merge::Region, slide::table::merge::transaction::Error as SlideTableMergesError,
    slide::table::merge::transaction::LimitKind as SlideTableMergesLimitKind,
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
        SlideTableMergesError::Refused(
            litchi_keynote::slide::table::merge::transaction::Refusal::TablePositionNotFound { .. }
        )
    ));
    Ok(())
}

#[test]
fn selection_and_merge_decode_share_the_cumulative_reference_budget()
-> Result<(), Box<dyn StdError>> {
    let semantic = SemanticLimits::builder().max_references(1)?.build()?;
    let package = Package::from_bytes_with_options(
        NATIVE_SOURCE,
        OpenOptions::new().semantic_limits(semantic),
    )?;

    let error = package
        .slide_table_merges(SlideSelector::index(0), TableSelector::index(0))
        .expect_err("root selection must consume the single permitted reference");

    assert!(matches!(
        error,
        SlideTableMergesError::Limit {
            kind: SlideTableMergesLimitKind::References,
            ..
        }
    ));
    Ok(())
}

/// A merge owner the commit creates takes a UUID derived from the table
/// model (record 0300, `rewrite_table_merges_derived`): two identical merge
/// rewrites of identical sources are byte-identical.
#[test]
fn merge_rewrites_are_deterministic_including_a_new_owner() -> Result<(), Box<dyn StdError>> {
    let package = Package::from_bytes(NATIVE_SOURCE)?;
    let existing = Region::new(3, 1, 2, 2)?;
    let fresh = Region::new(0, 0, 2, 2)?;
    // Removing the only merge leaves a table without merge owner state.
    let mut edit = package.edit_slide_table_merges(0, 0)?;
    edit.unmerge(existing)?;
    let unmerged = Package::from_bytes(&exact_bytes(edit.commit()?.package())?)?;
    let merge = |source: &Package| -> Result<Vec<u8>, Box<dyn StdError>> {
        let mut edit = source.edit_slide_table_merges(0, 0)?;
        edit.merge(fresh)?;
        exact_bytes(edit.commit()?.package())
    };
    let first = merge(&unmerged)?;
    let second = merge(&Package::from_bytes(&exact_bytes(&unmerged)?)?)?;
    assert_eq!(first, second);
    let reopened = Package::from_bytes(&first)?;
    assert_eq!(
        reopened.slide_table_merges(SlideSelector::index(0), TableSelector::index(0))?,
        [fresh]
    );
    // And on the native source, which keeps its owner.
    let mut edit = package.edit_slide_table_merges(0, 0)?;
    edit.merge(fresh)?;
    let once = exact_bytes(edit.commit()?.package())?;
    let mut edit = package.edit_slide_table_merges(0, 0)?;
    edit.merge(fresh)?;
    assert_eq!(exact_bytes(edit.commit()?.package())?, once);
    Ok(())
}
