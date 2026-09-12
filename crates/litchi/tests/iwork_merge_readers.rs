//! Public facade access to the format-owned, selector-first merge readers.

#[cfg(feature = "numbers")]
#[test]
fn numbers_merge_reader_is_available_without_the_migration_host() {
    use litchi::numbers::{MergeReader, SheetSelector, TableMergesError, TableSelector};
    let source = include_bytes!("../../../test-data/iwork/numbers/table-merges-native.numbers");
    let reader = MergeReader::from_bytes(source).unwrap();
    let regions = reader
        .table_merges(SheetSelector::index(0), TableSelector::index(0))
        .unwrap();
    assert_eq!(
        regions,
        vec![litchi::numbers::table::merge::Region::new(10, 1, 2, 2).unwrap()]
    );
    assert!(matches!(
        reader.table_merges(SheetSelector::index(999), TableSelector::index(0)),
        Err(TableMergesError::SheetNotFound)
    ));
}

#[cfg(feature = "pages")]
#[test]
fn pages_merge_reader_is_available_without_the_migration_host() {
    use litchi::pages::{BodyTableSelector, MergeReader};
    let source = include_bytes!("../../../test-data/iwork/pages/body-table-merges-native.pages");
    let reader = MergeReader::from_bytes(source).unwrap();
    let regions = reader
        .body_table_merges(BodyTableSelector::index(0))
        .unwrap();
    assert_eq!(
        regions,
        vec![litchi::pages::table::merge::Region::new(3, 2, 1, 2).unwrap()]
    );
}

#[cfg(feature = "keynote")]
#[test]
fn keynote_merge_reader_is_available_without_the_migration_host() {
    use litchi::keynote::{MergeReader, SlideSelector, TableSelector};
    let source = include_bytes!("../../../test-data/iwork/keynote/slide-table-merges-native.key");
    let reader = MergeReader::from_bytes(source).unwrap();
    let regions = reader
        .slide_table_merges(SlideSelector::index(0), TableSelector::index(0))
        .unwrap();
    assert_eq!(
        regions,
        vec![litchi::keynote::slide::table::merge::Region::new(3, 1, 2, 2).unwrap()]
    );
}
