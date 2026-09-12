//! Selector-first merge transactions through the supported root facades.

#[cfg(any(feature = "numbers", feature = "pages", feature = "keynote"))]
type TestResult = Result<(), Box<dyn std::error::Error>>;

#[cfg(feature = "numbers")]
#[test]
fn numbers_merge_transactions_preserve_source_and_support_inverse() -> TestResult {
    use litchi::numbers::{Package, SheetSelector, TableSelector};
    let bytes = include_bytes!("../../../test-data/iwork/numbers/table-merges-native.numbers");
    let source = Package::from_bytes(bytes)?;
    let before = source.table_merges(SheetSelector::index(0), TableSelector::index(0))?;
    let region = litchi::numbers::table::merge::Region::new(1, 1, 1, 2)?;
    let outside = litchi::numbers::table::merge::Region::new(u32::MAX, 0, 1, 2)?;
    let mut edit = source.edit_table_merges(SheetSelector::index(0), TableSelector::index(0))?;
    assert!(edit.merge(before[0]).is_err());
    assert!(edit.merge(outside).is_err());
    edit.merge(region)?;
    let committed = edit.commit()?;
    assert!(!committed.patch().is_noop());
    let mut expected = before.clone();
    expected.push(region);
    assert_eq!(
        committed
            .package()
            .table_merges(SheetSelector::index(0), TableSelector::index(0))?,
        expected
    );
    assert_eq!(
        source.table_merges(SheetSelector::index(0), TableSelector::index(0))?,
        before
    );
    assert!(
        committed
            .package()
            .apply_table_merges(committed.patch())
            .is_err()
    );
    let restored = committed
        .package()
        .apply_table_merges(&committed.patch().inverse())?;
    let mut restored_bytes = Vec::new();
    restored.package().write_to(&mut restored_bytes)?;
    assert_eq!(restored_bytes, bytes);
    let mut unmerge = committed
        .package()
        .edit_table_merges(SheetSelector::index(0), TableSelector::index(0))?;
    assert!(unmerge.unmerge(region)?);
    let removed = unmerge.commit()?;
    assert_eq!(
        removed
            .package()
            .table_merges(SheetSelector::index(0), TableSelector::index(0))?,
        before
    );
    let mut cancelled =
        source.edit_table_merges(SheetSelector::index(0), TableSelector::index(0))?;
    assert!(cancelled.unmerge(before[0])?);
    cancelled.merge(before[0])?;
    assert!(cancelled.commit()?.patch().is_noop());
    let mut missing = source.edit_table_merges(SheetSelector::index(0), TableSelector::index(0))?;
    assert!(!missing.unmerge(region)?);
    assert!(!missing.unmerge(outside)?);
    let unchanged = missing.commit()?;
    assert!(unchanged.patch().is_noop());
    let mut unchanged_bytes = Vec::new();
    unchanged.package().write_to(&mut unchanged_bytes)?;
    assert_eq!(unchanged_bytes, bytes);
    Ok(())
}

#[cfg(feature = "pages")]
#[test]
fn pages_merge_transactions_preserve_source_and_support_inverse() -> TestResult {
    use litchi::pages::{BodyTableSelector, Package};
    let bytes = include_bytes!("../../../test-data/iwork/pages/body-table-merges-native.pages");
    let source = Package::from_bytes(bytes)?;
    let before = source.body_table_merges(BodyTableSelector::index(0))?;
    let region = litchi::pages::table::merge::Region::new(1, 1, 1, 2)?;
    let outside = litchi::pages::table::merge::Region::new(u32::MAX, 0, 1, 2)?;
    let mut edit = source.edit_body_table_merges(BodyTableSelector::index(0))?;
    assert!(edit.merge(before[0]).is_err());
    assert!(edit.merge(outside).is_err());
    edit.merge(region)?;
    let committed = edit.commit()?;
    assert!(!committed.patch().is_noop());
    let mut expected = before.clone();
    expected.push(region);
    assert_eq!(
        committed
            .package()
            .body_table_merges(BodyTableSelector::index(0))?,
        expected
    );
    assert_eq!(
        source.body_table_merges(BodyTableSelector::index(0))?,
        before
    );
    assert!(
        committed
            .package()
            .apply_body_table_merges(committed.patch())
            .is_err()
    );
    let restored = committed
        .package()
        .apply_body_table_merges(&committed.patch().inverse())?;
    let mut restored_bytes = Vec::new();
    restored.package().write_to(&mut restored_bytes)?;
    assert_eq!(restored_bytes, bytes);
    let mut unmerge = committed
        .package()
        .edit_body_table_merges(BodyTableSelector::index(0))?;
    assert!(unmerge.unmerge(region)?);
    let removed = unmerge.commit()?;
    assert_eq!(
        removed
            .package()
            .body_table_merges(BodyTableSelector::index(0))?,
        before
    );
    let mut cancelled = source.edit_body_table_merges(BodyTableSelector::index(0))?;
    assert!(cancelled.unmerge(before[0])?);
    cancelled.merge(before[0])?;
    assert!(cancelled.commit()?.patch().is_noop());
    let mut missing = source.edit_body_table_merges(BodyTableSelector::index(0))?;
    assert!(!missing.unmerge(region)?);
    assert!(!missing.unmerge(outside)?);
    let unchanged = missing.commit()?;
    assert!(unchanged.patch().is_noop());
    let mut unchanged_bytes = Vec::new();
    unchanged.package().write_to(&mut unchanged_bytes)?;
    assert_eq!(unchanged_bytes, bytes);
    Ok(())
}

#[cfg(feature = "keynote")]
#[test]
fn keynote_merge_transactions_preserve_source_and_support_inverse() -> TestResult {
    use litchi::keynote::{Package, SlideSelector, TableSelector};
    let bytes = include_bytes!("../../../test-data/iwork/keynote/slide-table-merges-native.key");
    let source = Package::from_bytes(bytes)?;
    let before = source.slide_table_merges(SlideSelector::index(0), TableSelector::index(0))?;
    let region = litchi::keynote::slide::table::merge::Region::new(1, 1, 1, 2)?;
    let outside = litchi::keynote::slide::table::merge::Region::new(u32::MAX, 0, 1, 2)?;
    let mut edit =
        source.edit_slide_table_merges(SlideSelector::index(0), TableSelector::index(0))?;
    assert!(edit.merge(before[0]).is_err());
    assert!(edit.merge(outside).is_err());
    edit.merge(region)?;
    let committed = edit.commit()?;
    assert!(!committed.patch().is_noop());
    let mut expected = before.clone();
    expected.push(region);
    assert_eq!(
        committed
            .package()
            .slide_table_merges(SlideSelector::index(0), TableSelector::index(0))?,
        expected
    );
    assert_eq!(
        source.slide_table_merges(SlideSelector::index(0), TableSelector::index(0))?,
        before
    );
    assert!(
        committed
            .package()
            .apply_slide_table_merges(committed.patch())
            .is_err()
    );
    let restored = committed
        .package()
        .apply_slide_table_merges(&committed.patch().inverse())?;
    let mut restored_bytes = Vec::new();
    restored.package().write_to(&mut restored_bytes)?;
    assert_eq!(restored_bytes, bytes);
    let mut unmerge = committed
        .package()
        .edit_slide_table_merges(SlideSelector::index(0), TableSelector::index(0))?;
    assert!(unmerge.unmerge(region)?);
    let removed = unmerge.commit()?;
    assert_eq!(
        removed
            .package()
            .slide_table_merges(SlideSelector::index(0), TableSelector::index(0))?,
        before
    );
    let mut cancelled =
        source.edit_slide_table_merges(SlideSelector::index(0), TableSelector::index(0))?;
    assert!(cancelled.unmerge(before[0])?);
    cancelled.merge(before[0])?;
    assert!(cancelled.commit()?.patch().is_noop());
    let mut missing =
        source.edit_slide_table_merges(SlideSelector::index(0), TableSelector::index(0))?;
    assert!(!missing.unmerge(region)?);
    assert!(!missing.unmerge(outside)?);
    let unchanged = missing.commit()?;
    assert!(unchanged.patch().is_noop());
    let mut unchanged_bytes = Vec::new();
    unchanged.package().write_to(&mut unchanged_bytes)?;
    assert_eq!(unchanged_bytes, bytes);
    Ok(())
}
