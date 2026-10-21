//! Source lifetime and concurrency contracts of metadata-only Numbers reads.
use litchi_numbers::table::merge::transaction::Error as TableMergesError;
use litchi_numbers::table::merge::transaction::LimitKind as TableMergesLimitKind;

use std::sync::Arc;

use litchi_numbers::{MergeReader, table::merge::Region};

const NATIVE: &[u8] =
    include_bytes!("../../../test-data/iwork/numbers/table-merges-native.numbers");

#[test]
fn shared_source_survives_clones_and_is_released_with_the_last_reader() {
    let source: Arc<[u8]> = Arc::from(NATIVE);
    let source_lifetime = Arc::downgrade(&source);
    let reader = MergeReader::from_shared_bytes(Arc::clone(&source)).unwrap();
    let retained_owners = Arc::strong_count(&source);
    let clone = reader.clone();
    assert_eq!(Arc::strong_count(&source), retained_owners);

    drop(source);
    drop(reader);
    assert!(source_lifetime.upgrade().is_some());
    assert_eq!(
        clone.table_merges("Sheet 1", "shared-model").unwrap(),
        vec![Region::new(10, 1, 2, 2).unwrap()]
    );
    assert_eq!(source_lifetime.upgrade().unwrap().as_ref(), NATIVE);

    drop(clone);
    assert!(source_lifetime.upgrade().is_none());
}

#[test]
fn cloned_readers_can_query_the_same_immutable_source_concurrently() {
    let reader = MergeReader::from_bytes(NATIVE).unwrap();
    let clones = [reader.clone(), reader.clone()];
    let results = std::thread::scope(|scope| {
        let [first, second] = clones;
        let by_name = scope.spawn(move || first.table_merges("Sheet 1", "shared-model"));
        let by_position = scope.spawn(move || {
            second.table_merges(
                litchi_numbers::SheetSelector::index(0),
                litchi_numbers::TableSelector::index(0),
            )
        });
        (by_name.join().unwrap(), by_position.join().unwrap())
    });
    let expected = vec![Region::new(10, 1, 2, 2).unwrap()];
    assert_eq!(results.0.unwrap(), expected);
    assert_eq!(results.1.unwrap(), expected);
    assert_eq!(
        reader.table_merges("Sheet 1", "shared-model").unwrap(),
        expected
    );
}

#[test]
fn bounded_path_ingress_retains_an_independent_immutable_source() {
    let file = tempfile::NamedTempFile::new().unwrap();
    std::fs::write(file.path(), NATIVE).unwrap();
    let maximum = u64::try_from(NATIVE.len() - 1).unwrap();
    let physical = litchi_numbers::SourceLimits::builder()
        .max_input_bytes(maximum)
        .and_then(|b| b.max_members(litchi_numbers::SourceLimits::MAXIMUM.max_members()))
        .and_then(|b| b.max_entry_bytes(litchi_numbers::SourceLimits::MAXIMUM.max_entry_bytes()))
        .and_then(|b| {
            b.max_aggregate_bytes(litchi_numbers::SourceLimits::MAXIMUM.max_aggregate_bytes())
        })
        .and_then(|b| {
            b.max_component_bytes(litchi_numbers::SourceLimits::MAXIMUM.max_component_bytes())
        })
        .and_then(|b| b.build())
        .unwrap();
    let options = litchi_numbers::OpenOptions::new()
        .source_limits(physical)
        .semantic_limits(litchi_numbers::SemanticLimits::DEFAULT);
    assert!(matches!(
        MergeReader::open_with_options(file.path(), options),
        Err(TableMergesError::Limit {
            kind: TableMergesLimitKind::InputBytes,
            observed,
            maximum: actual_maximum,
        }) if observed == maximum + 1 && actual_maximum == maximum
    ));

    let reader = MergeReader::open(file.path()).unwrap();
    drop(file);
    assert_eq!(
        reader.table_merges("Sheet 1", "shared-model").unwrap(),
        vec![Region::new(10, 1, 2, 2).unwrap()]
    );
}

#[cfg(feature = "internal-iwork-source")]
#[test]
fn shared_catalog_reader_pins_original_parsed_archives_until_last_clone_drops() {
    use litchi_iwa_archive::{ComponentCatalog, Limits};

    let records = ComponentCatalog::from_bytes(NATIVE, Default::default())
        .unwrap()
        .into_iter()
        .map(|component| {
            let (name, archive) = component.into_parts();
            (name, Arc::new(archive))
        })
        .collect::<Vec<_>>();
    let lifetimes = records
        .iter()
        .map(|(_, archive)| Arc::downgrade(archive))
        .collect::<Vec<_>>();
    let catalog = Arc::new(
        ComponentCatalog::__from_shared_archives(
            records
                .iter()
                .map(|(name, archive)| (name.as_str(), Arc::clone(archive))),
            Limits::default(),
        )
        .unwrap(),
    );
    let object_count: usize = records
        .iter()
        .map(|(_, archive)| archive.objects.len())
        .sum();
    let defaults = litchi_numbers::SemanticLimits::DEFAULT;
    let limited = litchi_numbers::SemanticLimits::builder()
        .max_objects(object_count - 1)
        .and_then(|b| b.max_sheets(defaults.max_sheets()))
        .and_then(|b| b.max_tables(defaults.max_tables()))
        .and_then(|b| b.max_references(defaults.max_references()))
        .and_then(|b| b.build())
        .unwrap();
    assert!(matches!(
        MergeReader::__from_shared_catalog(
            Arc::clone(&catalog),
            litchi_numbers::OpenOptions::new().semantic_limits(limited),
        ),
        Err(TableMergesError::Limit {
            kind: TableMergesLimitKind::PayloadObjects,
            observed,
            maximum,
        }) if observed == object_count as u64 && maximum == (object_count - 1) as u64
    ));
    let reader = MergeReader::__from_shared_catalog(
        Arc::clone(&catalog),
        litchi_numbers::OpenOptions::new(),
    )
    .unwrap();
    let clone = reader.clone();
    for (name, archive) in &records {
        assert!(std::ptr::eq(
            catalog.get(name).unwrap().archive(),
            Arc::as_ptr(archive),
        ));
        assert_eq!(Arc::strong_count(archive), 2);
    }
    drop(records);
    drop(catalog);
    drop(reader);
    assert!(lifetimes.iter().all(|record| record.upgrade().is_some()));
    assert_eq!(
        clone.table_merges("Sheet 1", "shared-model").unwrap(),
        vec![Region::new(10, 1, 2, 2).unwrap()]
    );
    drop(clone);
    assert!(lifetimes.iter().all(|record| record.upgrade().is_none()));
}
