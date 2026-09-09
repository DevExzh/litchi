//! Focused, read-only catalog coverage for Pages body tables.
//!
//! The catalog exposes only rooted source order, validated names, and table
//! dimensions.  These tests keep the native object graph and archive IDs
//! behind the package boundary while checking exact source preservation,
//! selector parity, malformed-source rejection, and bounded ingress.

use std::error::Error as StdError;

use litchi_iwa_archive::{Limits as ArchiveLimits, package::EntryEdit};
use litchi_iwa_core::{
    Archive, ArchiveLimits as CoreArchiveLimits, FieldInfo, FieldPath, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{tp, tst};
use litchi_pages::{
    BodyTableCatalogError, BodyTableCatalogLimitKind, BodyTableSelector, Limits, Package,
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

const BODY_MARKER: &str = "Pages hidden-axis native oracle — 2026-09-05";
const NATIVE_VISIBLE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/body-table-name-visible-native.pages"
));
const FOCUSED_RUST: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/body-table-name-focused.pages"
));
const FOCUSED_NATIVE_SAVED: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/body-table-name-focused-native-saved.pages"
));
const SOURCE_BUILT_BEFORE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/source-built-table-stylesheet-before.pages"
));
const SOURCE_BUILT_NATIVE_SAVED: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/source-built-table-stylesheet-native-saved.pages"
));
const NATIVE_MULTI: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/body-table-catalog-native.pages"
));
const BASIC: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/basic.pages"
));

#[derive(Clone, Copy)]
struct ExpectedTable {
    name: &'static str,
    rows: u32,
    columns: u32,
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn assert_catalog(source: &[u8], expected: &[ExpectedTable], label: &str) -> TestResult {
    let package = Package::from_bytes(source)?;
    let catalog = package.body_tables()?;
    assert_eq!(catalog.len(), expected.len(), "{label}: table count");
    assert_eq!(
        catalog.is_empty(),
        expected.is_empty(),
        "{label}: empty state"
    );
    assert_eq!(
        catalog.as_slice().len(),
        expected.len(),
        "{label}: slice length"
    );

    for (index, expected_table) in expected.iter().enumerate() {
        let snapshot = catalog
            .get(index)
            .ok_or_else(|| format!("{label}: missing table {index}"))?;
        assert_eq!(snapshot.index(), index, "{label}: source index");
        assert_eq!(snapshot.name(), expected_table.name, "{label}: name");
        assert_eq!(snapshot.rows(), expected_table.rows, "{label}: rows");
        assert_eq!(
            snapshot.columns(),
            expected_table.columns,
            "{label}: columns"
        );
        let by_position = catalog
            .select(snapshot.selector())?
            .ok_or_else(|| format!("{label}: position selector {index}"))?;
        assert_eq!(by_position, snapshot, "{label}: position selector parity");
        let by_name = catalog
            .select(BodyTableSelector::name(expected_table.name))?
            .ok_or_else(|| format!("{label}: name selector {index}"))?;
        assert_eq!(by_name, snapshot, "{label}: name selector parity");
    }

    assert!(
        catalog
            .select(BodyTableSelector::index(expected.len()))?
            .is_none(),
        "{label}: index bound"
    );
    assert!(
        catalog
            .select(BodyTableSelector::name("missing table"))?
            .is_none()
    );
    assert_eq!(exact_bytes(&package)?, source, "{label}: exact source");

    let reopened = Package::from_bytes(&exact_bytes(&package)?)?;
    assert_eq!(reopened.body_tables()?, catalog, "{label}: reopen catalog");
    assert_eq!(exact_bytes(&reopened)?, source, "{label}: reopen bytes");
    Ok(())
}

fn remove_root_body_payload_keep_data_reference(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = litchi_iwa_archive::package::Catalog::from_bytes(source)?;
    let mut replacements = Vec::<(String, Vec<u8>)>::new();
    let mut roots = 0usize;

    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        let Some((object_index, message_index)) =
            archive
                .objects
                .iter()
                .enumerate()
                .find_map(|(object_index, object)| {
                    object
                        .messages
                        .iter()
                        .position(|message| message.type_ == 10_000)
                        .map(|index| (object_index, index))
                })
        else {
            continue;
        };
        roots = roots.saturating_add(1);
        let root = tp::DocumentArchive::decode(
            archive.objects[object_index].messages[message_index]
                .data
                .as_slice(),
        )?;
        let body_identifier = root
            .body_storage
            .as_ref()
            .map(|reference| reference.identifier)
            .filter(|identifier| *identifier != 0)
            .ok_or("root body reference is missing")?;
        let object = &mut archive.objects[object_index];
        let message_info = object
            .archive_info
            .message_infos
            .get_mut(message_index)
            .ok_or("root message metadata is missing")?;
        if let Some(field) = message_info
            .field_infos
            .iter_mut()
            .find(|field| field.path.as_slice() == [4])
        {
            field.object_references.clear();
            field.data_references.clear();
            field.data_references.push(body_identifier);
        } else {
            let mut field = FieldInfo::new(FieldPath::from(vec![4]));
            field.data_references.push(body_identifier);
            message_info.field_infos.push(field);
        }
        message_info
            .object_references
            .retain(|identifier| *identifier != body_identifier);
        let mut root = root;
        root.body_storage = None;
        let data = root.encode_to_vec();
        object.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: 10_000,
                data,
            },
        )?;
        replacements.push((
            entry.name().to_owned(),
            SnappyStream::compress(&archive.to_bytes()?)?.to_vec(),
        ));
    }
    assert_eq!(roots, 1, "expected one Pages document root");
    let edits = replacements
        .iter()
        .map(|(name, bytes)| EntryEdit::new(name, bytes.as_slice()))
        .collect::<Vec<_>>();
    Ok(catalog.reassemble_to_bytes(&edits, ArchiveLimits::default())?)
}

fn rewrite_table_models(
    source: &[u8],
    rename_second_to: Option<&str>,
    append_to_first: Option<&[u8]>,
) -> TestResult<Vec<u8>> {
    let catalog = litchi_iwa_archive::package::Catalog::from_bytes(source)?;
    let mut replacements = Vec::<(String, Vec<u8>)>::new();
    let mut model_count = 0usize;

    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        let mut changed = false;
        for object in &mut archive.objects {
            let mut message_index = 0usize;
            while message_index < object.messages.len() {
                if object.messages[message_index].type_ != 6_001 {
                    message_index += 1;
                    continue;
                }
                let model_index = model_count;
                model_count = model_count.saturating_add(1);
                let mut data = object.messages[message_index].data.clone();
                if model_index == 0 {
                    if let Some(raw) = append_to_first {
                        data.extend_from_slice(raw);
                    }
                }
                if model_index == 1 {
                    if let Some(name) = rename_second_to {
                        let mut model = tst::TableModelArchive::decode(data.as_slice())?;
                        model.table_name = name.to_owned();
                        data = model.encode_to_vec();
                    }
                }
                if data != object.messages[message_index].data {
                    object.replace_message_preserving_header(
                        message_index,
                        RawMessage { type_: 6_001, data },
                    )?;
                    changed = true;
                }
                message_index += 1;
            }
        }
        if changed {
            replacements.push((
                entry.name().to_owned(),
                SnappyStream::compress(&archive.to_bytes()?)?.to_vec(),
            ));
        }
    }
    assert_eq!(model_count, 2, "expected the two native catalog models");

    let edits = replacements
        .iter()
        .map(|(name, bytes)| EntryEdit::new(name, bytes.as_slice()))
        .collect::<Vec<_>>();
    Ok(catalog.reassemble_to_bytes(&edits, ArchiveLimits::default())?)
}

fn add_root_body_data_reference(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = litchi_iwa_archive::package::Catalog::from_bytes(source)?;
    let mut replacements = Vec::<(String, Vec<u8>)>::new();
    let mut roots = 0usize;

    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        let mut changed = false;
        for object in &mut archive.objects {
            let Some(message_index) = object
                .messages
                .iter()
                .position(|message| message.type_ == 10_000)
            else {
                continue;
            };
            roots = roots.saturating_add(1);
            let root = tp::DocumentArchive::decode(object.messages[message_index].data.as_slice())?;
            let body_identifier = root
                .body_storage
                .as_ref()
                .map(|reference| reference.identifier)
                .filter(|identifier| *identifier != 0)
                .ok_or("root body reference is missing")?;
            let message_info = object
                .archive_info
                .message_infos
                .get_mut(message_index)
                .ok_or("root message metadata is missing")?;
            let field = message_info
                .field_infos
                .iter_mut()
                .find(|field| field.path.as_slice() == [4]);
            if let Some(field) = field {
                field.data_references.push(body_identifier);
            } else {
                let mut field = FieldInfo::new(FieldPath::from(vec![4]));
                field.data_references.push(body_identifier);
                message_info.field_infos.push(field);
            }
            changed = true;
        }
        if changed {
            replacements.push((
                entry.name().to_owned(),
                SnappyStream::compress(&archive.to_bytes()?)?.to_vec(),
            ));
        }
    }
    assert_eq!(roots, 1, "expected one Pages document root");
    let edits = replacements
        .iter()
        .map(|(name, bytes)| EntryEdit::new(name, bytes.as_slice()))
        .collect::<Vec<_>>();
    Ok(catalog.reassemble_to_bytes(&edits, ArchiveLimits::default())?)
}

#[test]
fn native_visible_catalog_reads_name_and_dimensions_without_ids() -> TestResult {
    assert_catalog(
        NATIVE_VISIBLE,
        &[ExpectedTable {
            name: "Table 1",
            rows: 5,
            columns: 4,
        }],
        "native-visible",
    )
}

#[test]
fn focused_rust_and_native_saved_catalogs_reopen_exactly() -> TestResult {
    let expected = [ExpectedTable {
        name: "Native renamed table",
        rows: 5,
        columns: 4,
    }];
    assert_catalog(FOCUSED_RUST, &expected, "focused-rust")?;
    assert_catalog(FOCUSED_NATIVE_SAVED, &expected, "focused-native-saved")
}

#[test]
fn source_built_and_native_saved_catalogs_reopen_exactly() -> TestResult {
    let expected = [ExpectedTable {
        name: "Cities",
        rows: 5,
        columns: 4,
    }];
    assert_catalog(SOURCE_BUILT_BEFORE, &expected, "source-built-before")?;
    assert_catalog(
        SOURCE_BUILT_NATIVE_SAVED,
        &expected,
        "source-built-native-saved",
    )
}

#[test]
fn native_catalog_preserves_two_table_order_and_dimensions() -> TestResult {
    let package = Package::from_bytes(NATIVE_MULTI)?;
    assert!(package.text()?.contains(BODY_MARKER));
    assert_catalog(
        NATIVE_MULTI,
        &[
            ExpectedTable {
                name: "Table 1",
                rows: 5,
                columns: 4,
            },
            ExpectedTable {
                name: "Table 2",
                rows: 3,
                columns: 2,
            },
        ],
        "native-multi",
    )
}

#[test]
fn basic_pages_without_tables_returns_an_exact_empty_catalog() -> TestResult {
    assert_catalog(BASIC, &[], "basic-empty")
}

#[test]
fn duplicate_names_are_cataloged_but_name_selection_is_ambiguous() -> TestResult {
    let source = rewrite_table_models(NATIVE_MULTI, Some("Table 1"), None)?;
    let package = Package::from_bytes(&source)?;
    let catalog = package.body_tables()?;
    assert_eq!(catalog.len(), 2);
    assert_eq!(catalog.get(0).map(|table| table.name()), Some("Table 1"));
    assert_eq!(catalog.get(1).map(|table| table.name()), Some("Table 1"));
    for index in 0..catalog.len() {
        let snapshot = catalog.get(index).expect("catalog snapshot");
        assert_eq!(catalog.select(snapshot.selector())?, Some(snapshot));
    }
    assert!(matches!(
        catalog.select(BodyTableSelector::name("Table 1")),
        Err(BodyTableCatalogError::AmbiguousTableName)
    ));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn malformed_model_name_is_rejected_without_source_mutation() -> TestResult {
    let malformed = rewrite_table_models(NATIVE_MULTI, None, Some(&[0x42, 0x01, b'X']))?;
    let package = Package::from_bytes(&malformed)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.body_tables(),
        Err(BodyTableCatalogError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    assert_eq!(before, malformed);
    Ok(())
}

#[test]
fn root_body_field_data_reference_is_rejected_without_source_mutation() -> TestResult {
    let malformed = add_root_body_data_reference(NATIVE_MULTI)?;
    let package = Package::from_bytes(&malformed)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.body_tables(),
        Err(BodyTableCatalogError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    assert_eq!(before, malformed);
    Ok(())
}

#[test]
fn absent_root_body_payload_with_data_edge_is_rejected_without_source_mutation() -> TestResult {
    let malformed = remove_root_body_payload_keep_data_reference(NATIVE_MULTI)?;
    let package = Package::from_bytes(&malformed)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.body_tables(),
        Err(BodyTableCatalogError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    assert_eq!(before, malformed);
    Ok(())
}

#[test]
fn catalog_resource_limit_is_reported_after_successful_package_ingress() -> TestResult {
    let maximum = 512usize;
    let archive_limits = CoreArchiveLimits::default().with_objects(maximum)?;
    let limits = Limits::new(
        Limits::MAX_INPUT_BYTES,
        Limits::MAX_ENTRIES,
        Limits::MAX_ENTRY_BYTES,
        Limits::MAX_TOTAL_BYTES,
        Limits::MAX_IWA_STREAM_BYTES,
    )?
    .with_archive_limits(archive_limits)?;
    let package = Package::from_bytes_with_limits(NATIVE_MULTI, limits)?;
    let before = exact_bytes(&package)?;
    match package.body_tables() {
        Err(BodyTableCatalogError::LimitExceeded {
            kind: BodyTableCatalogLimitKind::PayloadObjects,
            observed,
            maximum: reported_maximum,
        }) => {
            assert_eq!(reported_maximum, u64::try_from(maximum)?);
            assert!(observed > reported_maximum);
        },
        result => panic!("expected catalog-time object limit, got {result:?}"),
    }
    assert_eq!(exact_bytes(&package)?, before);
    assert_eq!(before, NATIVE_MULTI);
    Ok(())
}

#[test]
fn catalog_input_limit_rejects_before_projection() -> TestResult {
    let limits = Limits::new(
        u64::try_from(NATIVE_MULTI.len().saturating_sub(1))?,
        Limits::MAX_ENTRIES,
        Limits::MAX_ENTRY_BYTES,
        Limits::MAX_TOTAL_BYTES,
        Limits::MAX_IWA_STREAM_BYTES,
    )?;
    assert!(Package::from_bytes_with_limits(NATIVE_MULTI, limits).is_err());
    Ok(())
}
