//! Focused Pages merge-reader coverage.
//!
//! The reader must resolve semantic selectors and decode only the selected
//! merge metadata while sharing immutable parsed archives across clones and
//! the migration-only component-catalog ingress.

use std::error::Error as StdError;
use std::sync::Arc;

use litchi_iwa_archive::Limits;
#[cfg(feature = "internal-iwork-source")]
use litchi_iwa_archive::SourceCatalog;
use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::table::merge::Region;
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::tswp;
use litchi_pages::{BodyTableMergesError, BodyTableSelector, MergeReader, Package};
use prost::Message as _;

const NATIVE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/body-table-merges-native.pages"
));
const BASIC: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/basic.pages"
));
const NUMBERS: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/numbers/table-merges-native.numbers"
));
const KEYNOTE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/keynote/slide-table-merges-native.key"
));

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

fn expected() -> [Region; 1] {
    [Region::new(3, 2, 1, 2).expect("native Pages merge geometry is valid")]
}

fn append_irrelevant_body_run(source: &[u8], length: usize) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == "Index/Document.iwa")
        .ok_or("missing Pages document component")?;
    let stream = SnappyStream::decompress(entry.data())?;
    let mut archive = Archive::parse(stream.as_bytes())?;
    for object in &mut archive.objects {
        for message_index in 0..object.messages.len() {
            let message_type = object.messages[message_index].type_;
            if !matches!(message_type, 2_001 | 2_022) {
                continue;
            }
            let Ok(mut storage) =
                tswp::StorageArchive::decode(object.messages[message_index].data.as_slice())
            else {
                continue;
            };
            if !storage
                .table_attachment
                .as_ref()
                .is_some_and(|table| !table.entries.is_empty())
            {
                continue;
            }
            storage.text.push("x".repeat(length));
            object.replace_message_preserving_header(
                message_index,
                RawMessage {
                    type_: message_type,
                    data: storage.encode_to_vec(),
                },
            )?;
            let component = SnappyStream::compress(&archive.to_bytes()?)?;
            return Ok(catalog.reassemble_to_bytes(
                &[EntryEdit::new(entry.name(), component.as_ref())],
                Limits::default(),
            )?);
        }
    }
    Err("native Pages body table storage is missing".into())
}

#[test]
fn reader_matches_package_for_name_and_position() -> TestResult {
    let reader = MergeReader::from_bytes(NATIVE)?;
    let clone = reader.clone();
    let package = Package::from_bytes(NATIVE)?;

    assert_eq!(reader.body_table_merges("Table 1")?, expected());
    assert_eq!(
        clone.body_table_merges(BodyTableSelector::index(0))?,
        expected()
    );
    assert!(
        reader
            .body_table_merges(BodyTableSelector::index(1))?
            .is_empty()
    );
    assert_eq!(
        reader.body_table_merges(BodyTableSelector::name("missing")),
        Err(BodyTableMergesError::TableNotFound)
    );
    assert_eq!(package.body_table_merges(0usize)?, expected());
    Ok(())
}

#[test]
fn reader_keeps_a_bounded_snapshot_after_path_removal() -> TestResult {
    let directory = tempfile::tempdir()?;
    let path = directory.path().join("merge-reader.pages");
    std::fs::write(&path, NATIVE)?;

    let reader = MergeReader::open(&path)?;
    std::fs::remove_file(&path)?;

    assert_eq!(reader.body_table_merges(0usize)?, expected());
    Ok(())
}

#[test]
fn reader_accepts_an_inclusive_input_limit_and_rejects_one_below() -> TestResult {
    let input_bytes = u64::try_from(NATIVE.len())?;
    let exact = Limits::new(
        input_bytes,
        Limits::MAX_ENTRIES,
        Limits::MAX_ENTRY_BYTES,
        Limits::MAX_TOTAL_BYTES,
        Limits::MAX_IWA_STREAM_BYTES,
    )?;
    assert_eq!(
        MergeReader::from_bytes_with_limits(NATIVE, exact)?.body_table_merges(0usize)?,
        expected()
    );

    let below = Limits::new(
        input_bytes - 1,
        Limits::MAX_ENTRIES,
        Limits::MAX_ENTRY_BYTES,
        Limits::MAX_TOTAL_BYTES,
        Limits::MAX_IWA_STREAM_BYTES,
    )?;
    assert!(matches!(
        MergeReader::from_bytes_with_limits(NATIVE, below),
        Err(BodyTableMergesError::LimitExceeded {
            kind: litchi_pages::BodyTableMergesLimitKind::InputBytes,
            observed,
            maximum,
        }) if observed == input_bytes && maximum == input_bytes - 1
    ));
    Ok(())
}

#[test]
fn reader_preserves_empty_body_and_table_miss_semantics() -> TestResult {
    let reader = MergeReader::from_bytes(BASIC)?;

    assert_eq!(
        reader.body_table_merges(0usize),
        Err(BodyTableMergesError::TableNotFound)
    );
    assert_eq!(
        reader.body_table_merges(BodyTableSelector::name("missing")),
        Err(BodyTableMergesError::TableNotFound)
    );
    Ok(())
}

#[test]
fn reader_keeps_large_irrelevant_trailing_body_runs_bounded() -> TestResult {
    let source = append_irrelevant_body_run(NATIVE, 32 * 1024)?;
    let reader = MergeReader::from_bytes(&source)?;

    assert_eq!(reader.body_table_merges(0usize)?, expected());
    Ok(())
}

#[test]
fn reader_rejects_foreign_iwork_formats_at_ingress() {
    for source in [NUMBERS, KEYNOTE] {
        let error = MergeReader::from_bytes(source).expect_err("foreign format must fail");
        assert_eq!(error, BodyTableMergesError::InvalidSource);
    }
}

#[test]
fn shared_byte_reader_clones_survive_input_source_drop() -> TestResult {
    let source: Arc<[u8]> = Arc::from(NATIVE);
    let source_lifetime = Arc::downgrade(&source);
    let reader = MergeReader::from_shared_bytes(Arc::clone(&source))?;
    let clone = reader.clone();

    drop(source);
    assert!(source_lifetime.upgrade().is_none());
    drop(reader);
    assert_eq!(clone.body_table_merges("Table 1")?, expected());
    drop(clone);
    Ok(())
}

#[cfg(feature = "internal-iwork-source")]
#[test]
fn shared_catalog_ingress_retains_the_original_archive_records() -> TestResult {
    let source = SourceCatalog::from_bytes_with_limits(NATIVE, Limits::default())?;
    let components = Arc::new(source.into_components());
    let components_lifetime = Arc::downgrade(&components);
    let reader = MergeReader::__from_shared_catalog(Arc::clone(&components), Limits::default())?;
    let clone = reader.clone();

    assert_eq!(Arc::strong_count(&components), 2);
    assert_eq!(
        reader.body_table_merges(BodyTableSelector::index(0))?,
        expected()
    );

    drop(components);
    drop(reader);
    assert!(components_lifetime.upgrade().is_some());
    assert_eq!(
        clone.body_table_merges(BodyTableSelector::index(0))?,
        expected()
    );
    drop(clone);
    assert!(components_lifetime.upgrade().is_none());
    Ok(())
}

#[cfg(feature = "internal-iwork-source")]
fn component_payload_limits(
    components: &litchi_iwa_archive::ComponentCatalog,
    total: u64,
) -> Limits {
    let largest = components
        .iter()
        .map(|component| {
            component
                .archive()
                .objects
                .last()
                .map(|object| {
                    object
                        .data_offset
                        .checked_add(object.data_length)
                        .expect("native archive stream extent fits u64")
                })
                .unwrap_or(0)
        })
        .max()
        .expect("native Pages package has components");
    assert!(largest > 0);
    Limits::new(
        Limits::MAX_INPUT_BYTES,
        Limits::MAX_ENTRIES,
        Limits::MAX_ENTRY_BYTES,
        total,
        usize::try_from(largest).expect("native archive stream extent fits usize"),
    )
    .expect("native payload budget is within the physical hard ceilings")
}

#[cfg(feature = "internal-iwork-source")]
fn total_component_payload_bytes(components: &litchi_iwa_archive::ComponentCatalog) -> u64 {
    components
        .iter()
        .map(|component| {
            component
                .archive()
                .objects
                .last()
                .map(|object| {
                    object
                        .data_offset
                        .checked_add(object.data_length)
                        .expect("native archive stream extent fits u64")
                })
                .unwrap_or(0)
        })
        .sum()
}

#[cfg(feature = "internal-iwork-source")]
#[test]
fn shared_catalog_reader_accepts_the_exact_component_payload_budget() -> TestResult {
    let source = SourceCatalog::from_bytes_with_limits(NATIVE, Limits::default())?;
    let components = Arc::new(source.into_components());
    let total = total_component_payload_bytes(&components);
    let limits = component_payload_limits(&components, total);
    let reader = MergeReader::__from_shared_catalog(Arc::clone(&components), limits)?;

    assert_eq!(reader.body_table_merges(0usize)?, expected());
    Ok(())
}

#[cfg(feature = "internal-iwork-source")]
#[test]
fn shared_catalog_reader_rejects_one_below_component_payload_budget() -> TestResult {
    let source = SourceCatalog::from_bytes_with_limits(NATIVE, Limits::default())?;
    let components = Arc::new(source.into_components());
    let total = total_component_payload_bytes(&components);
    assert!(total > 0);
    let limits = component_payload_limits(&components, total - 1);
    let reader = MergeReader::__from_shared_catalog(Arc::clone(&components), limits)?;

    assert!(matches!(
        reader.body_table_merges(0usize),
        Err(BodyTableMergesError::LimitExceeded {
            kind: litchi_pages::BodyTableMergesLimitKind::TotalPayloadBytes,
            observed,
            maximum,
        }) if observed == total && maximum == total - 1
    ));
    Ok(())
}
