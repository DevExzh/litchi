//! Exact-source selector-first Numbers merge transactions.

use litchi_numbers::table::merge::Region;
use litchi_numbers::table::merge::transaction::Error as TableMergesError;
use litchi_numbers::{Package, SheetSelector, TableSelector};

type TestResult = Result<(), Box<dyn std::error::Error>>;

const NATIVE_SOURCE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/numbers/table-merges-native.numbers"
));

#[test]
fn merge_and_unmerge_preserve_exact_source_and_inverse() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let before = source.table_merges(SheetSelector::index(0), TableSelector::index(0))?;
    let region = Region::new(1, 1, 1, 2)?;

    let mut edit = source.edit_table_merges(SheetSelector::index(0), TableSelector::index(0))?;
    edit.merge(region)?;
    assert!(edit.regions()?.contains(&region));
    let committed = edit.commit()?;
    assert!(committed.diagnostics().changed());
    assert_eq!(committed.diagnostics().regions_added(), 1);
    let mut expected = before.clone();
    expected.push(region);
    assert_eq!(
        committed
            .package()
            .table_merges(SheetSelector::index(0), TableSelector::index(0))?,
        expected
    );

    let restored = committed
        .package()
        .apply_table_merges(&committed.patch().inverse())?;
    let mut restored_bytes = Vec::new();
    restored.package().write_to(&mut restored_bytes)?;
    assert_eq!(restored_bytes, NATIVE_SOURCE);

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
    Ok(())
}

#[test]
fn missing_unmerge_is_a_byte_exact_noop() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let region = Region::new(1, 1, 1, 2)?;
    let mut edit = source.edit_table_merges(SheetSelector::index(0), TableSelector::index(0))?;
    assert!(!edit.unmerge(region)?);
    let committed = edit.commit()?;
    assert!(committed.patch().is_noop());
    assert!(!committed.diagnostics().changed());
    let mut bytes = Vec::new();
    committed.package().write_to(&mut bytes)?;
    assert_eq!(bytes, NATIVE_SOURCE);
    Ok(())
}

#[test]
fn staging_rejects_overlap_and_out_of_bounds_without_mutation() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let mut edit = source.edit_table_merges(SheetSelector::index(0), TableSelector::index(0))?;
    let existing = Region::new(10, 1, 2, 2)?;
    assert!(matches!(
        edit.merge(existing),
        Err(TableMergesError::Refused(
            litchi_numbers::table::merge::transaction::Refusal::OverlappingRegion
        ))
    ));
    let outside = Region::new(11, 2, 2, 2)?;
    assert!(matches!(
        edit.merge(outside),
        Err(TableMergesError::Refused(
            litchi_numbers::table::merge::transaction::Refusal::InvalidRegion
        ))
    ));
    let committed = edit.commit()?;
    assert!(committed.patch().is_noop());
    Ok(())
}

#[test]
fn inverse_patch_rejects_the_wrong_source_artifact() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let region = Region::new(1, 1, 1, 2)?;
    let mut edit = source.edit_table_merges(SheetSelector::index(0), TableSelector::index(0))?;
    edit.merge(region)?;
    let committed = edit.commit()?;
    assert!(matches!(
        source.apply_table_merges(&committed.patch().inverse()),
        Err(TableMergesError::PatchConflict)
    ));
    Ok(())
}

#[test]
fn multiple_staged_merges_can_remove_one_and_then_all() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let first = Region::new(10, 1, 2, 2)?;
    let second = Region::new(1, 1, 1, 2)?;

    let mut add = source.edit_table_merges(SheetSelector::index(0), TableSelector::index(0))?;
    add.merge(second)?;
    let added = add.commit()?;
    assert_eq!(
        added
            .package()
            .table_merges(SheetSelector::index(0), TableSelector::index(0))?,
        [first, second]
    );

    let mut remove_one = added
        .package()
        .edit_table_merges(SheetSelector::index(0), TableSelector::index(0))?;
    assert!(remove_one.unmerge(first)?);
    let one = remove_one.commit()?;
    assert_eq!(
        one.package()
            .table_merges(SheetSelector::index(0), TableSelector::index(0))?,
        [second]
    );

    let mut remove_all = one
        .package()
        .edit_table_merges(SheetSelector::index(0), TableSelector::index(0))?;
    assert!(remove_all.unmerge(second)?);
    let empty = remove_all.commit()?;
    assert!(
        empty
            .package()
            .table_merges(SheetSelector::index(0), TableSelector::index(0))?
            .is_empty()
    );
    Ok(())
}

#[test]
fn removing_and_readding_an_existing_region_is_an_exact_noop() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let region = Region::new(10, 1, 2, 2)?;
    let mut edit = source.edit_table_merges(SheetSelector::index(0), TableSelector::index(0))?;
    assert!(edit.unmerge(region)?);
    edit.merge(region)?;
    let committed = edit.commit()?;
    assert!(committed.patch().is_noop());
    let mut bytes = Vec::new();
    committed.package().write_to(&mut bytes)?;
    assert_eq!(bytes, NATIVE_SOURCE);
    Ok(())
}

#[test]
fn changed_publication_honors_the_package_output_ceiling() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let limits = litchi_numbers::SourceLimits::builder()
        .max_input_bytes(NATIVE_SOURCE.len() as u64)
        .and_then(|b| b.max_members(litchi_numbers::SourceLimits::MAXIMUM.max_members()))
        .and_then(|b| b.max_entry_bytes(litchi_numbers::SourceLimits::MAXIMUM.max_entry_bytes()))
        .and_then(|b| {
            b.max_aggregate_bytes(litchi_numbers::SourceLimits::MAXIMUM.max_aggregate_bytes())
        })
        .and_then(|b| {
            b.max_component_bytes(litchi_numbers::SourceLimits::MAXIMUM.max_component_bytes())
        })
        .and_then(|b| b.build())?;
    let bounded = Package::from_bytes_with_options(
        NATIVE_SOURCE,
        litchi_numbers::OpenOptions::new()
            .source_limits(limits)
            .semantic_limits(litchi_numbers::SemanticLimits::DEFAULT),
    )?;
    let region = Region::new(1, 1, 1, 2)?;
    let mut edit = bounded.edit_table_merges(SheetSelector::index(0), TableSelector::index(0))?;
    edit.merge(region)?;
    assert!(matches!(edit.commit(), Err(TableMergesError::Limit { .. })));
    assert_eq!(
        source.table_merges(SheetSelector::index(0), TableSelector::index(0))?,
        [Region::new(10, 1, 2, 2)?,]
    );
    Ok(())
}

#[test]
fn identical_merges_are_byte_identical() -> TestResult {
    // Record 0300: creating a merge owner used to draw a random UUID, so two
    // identical merges on the same workbook differed byte for byte.
    let region = Region::new(1, 1, 2, 2)?;
    let merge = || -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        let source = Package::new();
        let mut edit =
            source.edit_table_merges(SheetSelector::index(0), TableSelector::index(0))?;
        edit.merge(region)?;
        let committed = edit.commit()?;
        assert_eq!(
            committed
                .package()
                .table_merges(SheetSelector::index(0), TableSelector::index(0))?,
            vec![region]
        );
        let mut bytes = Vec::new();
        committed.package().write_to(&mut bytes)?;
        Ok(bytes)
    };
    assert_eq!(merge()?, merge()?);
    Ok(())
}
