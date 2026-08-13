//! Public integration coverage for exact-source row and column size edits.

use std::{fmt::Debug, path::PathBuf};

use litchi_iwa_archive::{
    Limits as ArchiveLimits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::wire::{
    append_varint_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields,
};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, RawMessage, SnappyStream};
use litchi_iwa_protos::{tsp, tst};
use litchi_numbers::{
    Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
    cell::Value,
    table::{
        dimension::{
            Dimension, Points, Size,
            transaction::{Commit, Diagnostics, Edit, LimitKind, Patch, Path, TransactionError},
        },
        lock::State as LockState,
    },
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const FIXTURE_MARKER: &str = "Litchi native Numbers fixture";
const CANONICAL_PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

trait ExactBytes {
    fn exact_bytes(&self) -> Vec<u8>;
}

impl ExactBytes for Package {
    fn exact_bytes(&self) -> Vec<u8> {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)
            .expect("an in-memory Vec accepts package bytes");
        bytes
    }
}

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/numbers/basic.numbers")
}

fn points(value: f32) -> TestResult<Size> {
    Ok(Size::Points(Points::new(value)?))
}

fn assert_native_invariants(package: &Package) -> TestResult {
    let table = package
        .document()
        .sheet(0)?
        .and_then(|sheet| sheet.tables().next())
        .ok_or_else(|| std::io::Error::other("native table is missing"))?;
    assert_eq!((table.row_count(), table.column_count()), (22, 7));
    assert!(matches!(
        table.get_a1("B2")?,
        Some(Value::Text(value)) if value == FIXTURE_MARKER
    ));
    let headers = package.table_header_settings(0usize, 0usize)?;
    assert_eq!(headers.header_rows.map(|count| count.get()), Some(1));
    assert_eq!(headers.header_columns.map(|count| count.get()), Some(1));
    assert_eq!(headers.footer_rows, None);
    assert!(headers.header_rows_are_frozen());
    assert!(headers.header_columns_are_frozen());
    assert_eq!(headers.repeating_header_rows_enabled, Some(true));
    assert_eq!(headers.repeating_header_columns_enabled, Some(true));
    Ok(())
}

fn assert_changed_locality(source: &[u8], target: &[u8], expected_member: &str) -> TestResult {
    let source = Catalog::from_bytes(source)?;
    let target = Catalog::from_bytes(target)?;
    let mut changed = Vec::new();
    for before in source.iter() {
        let after = target
            .iter()
            .find(|candidate| candidate.name() == before.name());
        if CANONICAL_PREVIEWS.contains(&before.name()) {
            assert!(after.is_none(), "preview {} was retained", before.name());
            continue;
        }
        let after = after.ok_or_else(|| {
            std::io::Error::other(format!("member {} was unexpectedly deleted", before.name()))
        })?;
        if before.data() == after.data() {
            assert_eq!(
                before.raw_record().local_record(),
                after.raw_record().local_record(),
                "unchanged member {} lost its exact local record",
                before.name()
            );
        } else {
            changed.push(before.name());
        }
    }
    assert_eq!(changed, [expected_member]);
    assert_eq!(target.len() + CANONICAL_PREVIEWS.len(), source.len());
    Ok(())
}

fn rewrite_component(
    source: &[u8],
    member: &str,
    mutate: impl FnOnce(&mut Archive) -> TestResult,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| std::io::Error::other("selected component is missing"))?;
    let stream = SnappyStream::decompress(entry.data())?;
    let mut archive = Archive::parse(stream.as_bytes())?;
    mutate(&mut archive)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(member, &compressed)],
        ArchiveLimits::default(),
    )?)
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn rewrite_native_table_model(
    source: &[u8],
    mutate: impl FnOnce(&mut ArchiveObject, usize, &mut tst::TableModelArchive) -> TestResult,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut selected = None;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(stream.as_bytes())?;
        if archive.objects.iter().any(|object| {
            object.messages.iter().any(|message| {
                message.type_ == 6_001
                    && tst::TableModelArchive::decode(message.data.as_slice())
                        .is_ok_and(|model| model.table_name == "Table 1")
            })
        }) {
            selected = Some((entry.name().to_owned(), archive));
            break;
        }
    }
    let (member, mut archive) =
        selected.ok_or_else(|| std::io::Error::other("native Table 1 model is missing"))?;
    let (object_index, message_index) = archive
        .objects
        .iter()
        .enumerate()
        .find_map(|(object_index, object)| {
            object
                .messages
                .iter()
                .enumerate()
                .find_map(|(message_index, message)| {
                    (message.type_ == 6_001
                        && tst::TableModelArchive::decode(message.data.as_slice())
                            .is_ok_and(|model| model.table_name == "Table 1"))
                    .then_some((object_index, message_index))
                })
        })
        .ok_or_else(|| std::io::Error::other("native Table 1 payload is missing"))?;
    let object = &mut archive.objects[object_index];
    let mut model = tst::TableModelArchive::decode(object.messages[message_index].data.as_slice())?;
    mutate(object, message_index, &mut model)?;
    object.replace_message_preserving_header(
        message_index,
        RawMessage {
            type_: 6_001,
            data: model.encode_to_vec(),
        },
    )?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(&member, &compressed)],
        ArchiveLimits::default(),
    )?)
}

fn native_table_model(source: &[u8]) -> TestResult<tst::TableModelArchive> {
    let catalog = Catalog::from_bytes(source)?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(stream.as_bytes())?;
        for message in archive
            .objects
            .iter()
            .flat_map(|object| object.messages.iter())
            .filter(|message| message.type_ == 6_001)
        {
            if let Ok(model) = tst::TableModelArchive::decode(message.data.as_slice())
                && model.table_name == "Table 1"
            {
                return Ok(model);
            }
        }
    }
    Err(std::io::Error::other("native Table 1 model is missing").into())
}

fn expand_native_to_second_row_bucket(source: &[u8], rows: u32) -> TestResult<Vec<u8>> {
    rewrite_native_table_model(source, |object, message_index, model| {
        model.number_of_rows = rows;
        model
            .base_data_store
            .row_headers
            .buckets
            .push(reference(904_937));
        let info = &mut object.archive_info.message_infos[message_index];
        info.object_references
            .retain(|identifier| *identifier != 904_937);
        info.object_references.push(904_937);
        info.field_infos
            .retain(|field| field.path.as_slice() != [4, 1, 2]);
        let mut field = FieldInfo::new(vec![4, 1, 2]);
        field.object_references = vec![904_855, 904_937];
        info.field_infos.push(field);
        Ok(())
    })
}

fn selected_bucket_payload(source: &[u8], member: &str) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| std::io::Error::other("selected component is missing"))?;
    let stream = SnappyStream::decompress(entry.data())?;
    let archive = Archive::parse(stream.as_bytes())?;
    Ok(archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .find(|message| message.type_ == 6_006)
        .ok_or_else(|| std::io::Error::other("selected bucket payload is missing"))?
        .data
        .clone())
}

fn transaction_work_precharge(source: &[u8], target: &[u8]) -> TestResult<usize> {
    let catalog = Catalog::from_bytes(source)?;
    let mut observed = source.len().saturating_mul(2);
    if source != target {
        observed = observed.saturating_add(target.len().saturating_mul(2));
    }
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(stream.as_bytes())?;
        for object in &archive.objects {
            for message in &object.messages {
                observed = observed.saturating_add(message.data.len().saturating_mul(16));
            }
        }
    }
    Ok(observed)
}

fn package_with_reference_limit(source: &[u8], maximum: usize) -> TestResult<Package> {
    let semantic = PackageSemanticLimits::new(
        PackageSemanticLimits::MAX_OBJECTS,
        PackageSemanticLimits::MAX_SHEETS,
        PackageSemanticLimits::MAX_TABLES,
        maximum,
    )?;
    Ok(Package::from_bytes_with_options(
        source,
        PackageReadOptions::new(PackageLimits::default(), semantic),
    )?)
}

fn minimum_reference_limit(
    source: &[u8],
    succeeds: impl Fn(&Package) -> bool,
) -> TestResult<usize> {
    let mut lower = 1usize;
    let mut upper = PackageSemanticLimits::MAX_REFERENCES;
    while lower < upper {
        let middle = lower + (upper - lower) / 2;
        let success =
            package_with_reference_limit(source, middle).is_ok_and(|package| succeeds(&package));
        if success {
            upper = middle;
        } else {
            lower = middle + 1;
        }
    }
    Ok(lower)
}

#[test]
fn native_oracle_distinguishes_default_from_explicit_model_default() -> TestResult {
    let path = fixture_path();
    let bytes = std::fs::read(&path)?;
    for package in [Package::open(&path)?, Package::from_bytes(&bytes)?] {
        for row in 0..22 {
            assert_eq!(
                package.table_dimension_size(0usize, 0usize, Dimension::Row(row))?,
                Size::Default
            );
        }
        for column in 0..5 {
            assert_eq!(
                package.table_dimension_size(0usize, 0usize, Dimension::Column(column))?,
                Size::Default
            );
        }
        for column in 5..7 {
            assert_eq!(
                package.table_dimension_size(
                    SheetSelector::name("Sheet 1"),
                    TableSelector::name("Table 1"),
                    Dimension::Column(column),
                )?,
                points(98.0)?
            );
        }
        assert_native_invariants(&package)?;
    }
    Ok(())
}

#[test]
fn native_noops_share_exact_bytes_and_preserve_explicit_peer() -> TestResult {
    let package = Package::open(fixture_path())?;
    let source = package.exact_bytes();
    for (dimension, size) in [
        (Dimension::Row(0), Size::Default),
        (Dimension::Column(5), points(98.0)?),
    ] {
        let edit = package.edit_table_dimension_size("Sheet 1", "Table 1", dimension)?;
        assert_eq!(
            edit.path(),
            Path::Dimension {
                sheet: 0,
                table: 0,
                dimension,
            }
        );
        assert_eq!(edit.size(), size);
        let commit = edit.set(size).commit()?;
        assert!(commit.patch().is_noop());
        assert_eq!(commit.patch().before(), size);
        assert_eq!(commit.patch().after(), size);
        assert!(!commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().touched_components(), 0);
        assert_eq!(commit.diagnostics().deleted_previews(), 0);
        assert!(!commit.diagnostics().full_reparse_performed());
        assert_eq!(commit.package().exact_bytes(), source);
        assert_eq!(
            package
                .apply_table_dimension_size(commit.patch())?
                .package()
                .exact_bytes(),
            source
        );
    }
    assert_eq!(
        package.table_dimension_size(0usize, 0usize, Dimension::Column(6))?,
        points(98.0)?
    );
    Ok(())
}

#[test]
fn native_row_and_column_changes_are_local_exact_and_reversible() -> TestResult {
    for (dimension, after, member) in [
        (
            Dimension::Row(1),
            points(32.0)?,
            "Index/Tables/HeaderStorageBucket-904855-2.iwa",
        ),
        (
            Dimension::Column(2),
            points(124.0)?,
            "Index/Tables/HeaderStorageBucket-904899-2.iwa",
        ),
    ] {
        let package = Package::open(fixture_path())?;
        let source = package.exact_bytes();
        let before = package.table_dimension_size(0usize, 0usize, dimension)?;
        assert_eq!(before, Size::Default);
        let commit = package
            .edit_table_dimension_size(0usize, 0usize, dimension)?
            .set(after)
            .commit()?;
        assert_eq!(commit.patch().dimension(), dimension);
        assert_eq!(commit.patch().before(), before);
        assert_eq!(commit.patch().after(), after);
        assert!(!commit.patch().is_noop());
        assert_ne!(
            commit.patch().source_fingerprint(),
            commit.patch().target_fingerprint()
        );
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().touched_components(), 1);
        assert_eq!(commit.diagnostics().deleted_previews(), 3);
        assert!(commit.diagnostics().full_reparse_performed());
        assert_eq!(
            commit
                .package()
                .table_dimension_size("Sheet 1", "Table 1", dimension)?,
            after
        );
        assert_eq!(
            commit
                .package()
                .table_dimension_size(0usize, 0usize, Dimension::Column(6))?,
            points(98.0)?
        );
        assert_native_invariants(commit.package())?;

        let target = commit.package().exact_bytes();
        assert_changed_locality(&source, &target, member)?;
        assert_eq!(
            package
                .apply_table_dimension_size(commit.patch())?
                .package()
                .exact_bytes(),
            target
        );
        let reopened = Package::from_bytes(&target)?;
        assert!(matches!(
            reopened.apply_table_dimension_size(commit.patch()),
            Err(TransactionError::PatchConflict)
        ));
        let inverse = commit.patch().inverse();
        assert_eq!(inverse.inverse(), *commit.patch());
        let restored = reopened.apply_table_dimension_size(&inverse)?;
        assert_eq!(restored.package().exact_bytes(), source);
        assert_native_invariants(restored.package())?;
        assert_eq!(
            restored
                .package()
                .table_dimension_size(0usize, 0usize, Dimension::Column(6))?,
            points(98.0)?
        );
    }
    Ok(())
}

#[test]
fn changed_owned_dimension_package_composes_with_table_lock_noop_and_apply() -> TestResult {
    let source = Package::open(fixture_path())?;
    let changed = source
        .edit_table_dimension_size(0usize, 0usize, Dimension::Column(2))?
        .set(points(124.0)?)
        .commit()?
        .into_package();
    let changed_bytes = changed.exact_bytes();

    let lock = changed.edit_table_lock(0usize, 0usize)?.commit()?;
    assert!(lock.patch().is_noop());
    assert!(!lock.diagnostics().changed());
    assert_eq!(lock.package().exact_bytes(), changed_bytes);
    assert_eq!(
        lock.package()
            .table_dimension_size(0usize, 0usize, Dimension::Column(2))?,
        points(124.0)?
    );

    let applied = changed.apply_table_lock(lock.patch())?;
    assert!(applied.patch().is_noop());
    assert_eq!(applied.package().exact_bytes(), changed_bytes);
    assert_eq!(
        applied
            .package()
            .table_dimension_size(0usize, 0usize, Dimension::Column(2))?,
        points(124.0)?
    );
    Ok(())
}

#[test]
fn selector_bounds_locking_and_public_values_are_typed_and_redacted() -> TestResult {
    fn assert_send_sync_debug<T: Send + Sync + Debug>() {}
    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<Dimension>();
    assert_send_sync_debug::<Points>();
    assert_send_sync_debug::<Size>();
    assert_send_sync_debug::<Edit<'static>>();
    assert_send_sync_debug::<Patch>();
    assert_send_sync_debug::<Commit>();
    assert_send_sync_debug::<Diagnostics>();
    assert_send_sync_debug::<Path>();
    assert_send_sync_debug::<LimitKind>();
    assert_send_sync_debug::<TransactionError>();

    let package = Package::open(fixture_path())?;
    assert!(matches!(
        package.table_dimension_size("private missing sheet", 0usize, Dimension::Row(0)),
        Err(TransactionError::SheetNotFound)
    ));
    assert!(matches!(
        package.table_dimension_size(0usize, "private missing table", Dimension::Row(0)),
        Err(TransactionError::TableNotFound)
    ));
    for (dimension, length) in [(Dimension::Row(22), 22), (Dimension::Column(7), 7)] {
        assert!(matches!(
            package.table_dimension_size(0usize, 0usize, dimension),
            Err(TransactionError::OutOfBounds {
                path: Path::Dimension {
                    sheet: 0,
                    table: 0,
                    dimension: actual,
                },
                length: actual_length,
            }) if actual == dimension && actual_length == length
        ));
    }
    for invalid in [0.0, -1.0, f32::INFINITY, f32::NEG_INFINITY, f32::NAN] {
        assert!(Points::new(invalid).is_err());
        assert!(Size::points(invalid).is_err());
    }

    let mut lock = package.edit_table_lock(0usize, 0usize)?;
    lock.lock();
    let locked = lock.commit()?.into_package();
    assert_eq!(locked.table_lock(0usize, 0usize)?, LockState::Locked);
    let noop = locked
        .edit_table_dimension_size(0usize, 0usize, Dimension::Row(0))?
        .set(Size::Default)
        .commit()?;
    assert!(noop.patch().is_noop());
    let error = locked
        .edit_table_dimension_size(0usize, 0usize, Dimension::Row(0))?
        .set(points(32.0)?)
        .commit()
        .expect_err("changed locked dimension must refuse");
    assert!(matches!(
        error,
        TransactionError::TableLocked {
            path: Path::Dimension {
                sheet: 0,
                table: 0,
                dimension: Dimension::Row(0),
            }
        }
    ));
    for rendered in [format!("{error:?}"), error.to_string()] {
        assert!(!rendered.contains(FIXTURE_MARKER));
        assert!(!rendered.contains("private missing"));
        assert!(!rendered.contains("Index/"));
    }
    Ok(())
}

#[test]
fn clearing_canonical_and_nonminimal_headers_preserves_native_structure() -> TestResult {
    const ROWS: &str = "Index/Tables/HeaderStorageBucket-904855-2.iwa";
    const COLUMNS: &str = "Index/Tables/HeaderStorageBucket-904899-2.iwa";
    let package = Package::open(fixture_path())?;
    let source = package.exact_bytes();

    // Column 5 is a canonical size-only explicit 98pt entry. Clearing it
    // removes only that entry; column 6 remains explicitly 98pt.
    let cleared_column = package
        .edit_table_dimension_size(0usize, 0usize, Dimension::Column(5))?
        .set(Size::Default)
        .commit()?;
    assert_eq!(
        cleared_column
            .package()
            .table_dimension_size(0usize, 0usize, Dimension::Column(5))?,
        Size::Default
    );
    assert_eq!(
        cleared_column
            .package()
            .table_dimension_size(0usize, 0usize, Dimension::Column(6))?,
        points(98.0)?
    );
    assert_changed_locality(&source, &cleared_column.package().exact_bytes(), COLUMNS)?;
    assert_eq!(
        cleared_column
            .package()
            .apply_table_dimension_size(&cleared_column.patch().inverse())?
            .package()
            .exact_bytes(),
        source
    );

    // Row 1 already has a nonminimal retained header with a stored-cell
    // count. Add hiding/style facets and an unknown field to prove clearing a
    // staged size keeps the entire source-ordered record byte-exact.
    let faceted = rewrite_component(&source, ROWS, |archive| {
        let object = archive
            .object_mut(904_855)
            .ok_or_else(|| std::io::Error::other("row bucket is missing"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == 6_006)
            .ok_or_else(|| std::io::Error::other("row bucket message is missing"))?;
        let original = object.messages[message_index].data.clone();
        let mut raw = repeated_length_delimited_payloads(&original, 2)?
            .into_iter()
            .map(ToOwned::to_owned)
            .collect::<Vec<_>>();
        let selected = raw
            .iter_mut()
            .find(|record| {
                tst::header_storage_bucket::Header::decode(record.as_slice())
                    .is_ok_and(|header| header.index == 1)
            })
            .ok_or_else(|| std::io::Error::other("row 1 header is missing"))?;
        let mut header = tst::header_storage_bucket::Header::decode(selected.as_slice())?;
        header.hiding_state = 7;
        header.cell_style = Some(reference(123_456));
        *selected = header.encode_to_vec();
        append_varint_field(selected, 99, 9_999)?;
        let rewritten = rewrite_repeated_length_delimited_fields(&original, 2, &raw)?;
        object.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: 6_006,
                data: rewritten,
            },
        )?;
        Ok(())
    })?;
    let row_package = Package::from_bytes(&faceted)?;
    let changed_row = row_package
        .edit_table_dimension_size(0usize, 0usize, Dimension::Row(1))?
        .set(points(32.0)?)
        .commit()?;
    let cleared_row = changed_row
        .package()
        .edit_table_dimension_size(0usize, 0usize, Dimension::Row(1))?
        .set(Size::Default)
        .commit()?;
    assert_eq!(
        cleared_row
            .package()
            .table_dimension_size(0usize, 0usize, Dimension::Row(1))?,
        Size::Default
    );
    assert_eq!(
        selected_bucket_payload(&cleared_row.package().exact_bytes(), ROWS)?,
        selected_bucket_payload(&faceted, ROWS)?
    );
    assert_native_invariants(cleared_row.package())?;
    Ok(())
}

#[test]
fn malformed_duplicate_wrong_wire_and_wrong_type_buckets_refuse_atomically() -> TestResult {
    const COLUMNS: &str = "Index/Tables/HeaderStorageBucket-904899-2.iwa";
    let source = std::fs::read(fixture_path())?;
    let malformed_duplicate = rewrite_component(&source, COLUMNS, |archive| {
        let object = archive
            .object_mut(904_899)
            .ok_or_else(|| std::io::Error::other("column bucket is missing"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == 6_006)
            .ok_or_else(|| std::io::Error::other("column bucket message is missing"))?;
        let mut bucket =
            tst::HeaderStorageBucket::decode(object.messages[message_index].data.as_slice())?;
        bucket.headers.push(bucket.headers[0]);
        object.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: 6_006,
                data: bucket.encode_to_vec(),
            },
        )?;
        Ok(())
    })?;
    let out_of_range = rewrite_component(&source, COLUMNS, |archive| {
        let object = archive
            .object_mut(904_899)
            .ok_or_else(|| std::io::Error::other("column bucket is missing"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == 6_006)
            .ok_or_else(|| std::io::Error::other("column bucket message is missing"))?;
        let mut bucket =
            tst::HeaderStorageBucket::decode(object.messages[message_index].data.as_slice())?;
        let mut header = bucket.headers[0];
        header.index = 7;
        bucket.headers.push(header);
        object.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: 6_006,
                data: bucket.encode_to_vec(),
            },
        )?;
        Ok(())
    })?;
    let malformed_wire = rewrite_component(&source, COLUMNS, |archive| {
        let object = archive
            .object_mut(904_899)
            .ok_or_else(|| std::io::Error::other("column bucket is missing"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == 6_006)
            .ok_or_else(|| std::io::Error::other("column bucket message is missing"))?;
        let mut data = object.messages[message_index].data.clone();
        append_varint_field(&mut data, 2, 1)?;
        object
            .replace_message_preserving_header(message_index, RawMessage { type_: 6_006, data })?;
        Ok(())
    })?;
    let wrong_type = rewrite_component(&source, COLUMNS, |archive| {
        let object = archive
            .object_mut(904_899)
            .ok_or_else(|| std::io::Error::other("column bucket is missing"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == 6_006)
            .ok_or_else(|| std::io::Error::other("column bucket message is missing"))?;
        let data = object.messages[message_index].data.clone();
        object.replace_message_preserving_header(message_index, RawMessage { type_: 777, data })?;
        Ok(())
    })?;
    let duplicate_message = rewrite_component(&source, COLUMNS, |archive| {
        let object = archive
            .object_mut(904_899)
            .ok_or_else(|| std::io::Error::other("column bucket is missing"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == 6_006)
            .ok_or_else(|| std::io::Error::other("column bucket message is missing"))?;
        object.messages.push(object.messages[message_index].clone());
        object
            .archive_info
            .message_infos
            .push(object.archive_info.message_infos[message_index].clone());
        Ok(())
    })?;
    let wrong_message_metadata = rewrite_component(&source, COLUMNS, |archive| {
        let object = archive
            .object_mut(904_899)
            .ok_or_else(|| std::io::Error::other("column bucket is missing"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == 6_006)
            .ok_or_else(|| std::io::Error::other("column bucket message is missing"))?;
        object.archive_info.message_infos[message_index].base_message_index = Some(0);
        Ok(())
    })?;
    let wrong_object_identifier = rewrite_component(&source, COLUMNS, |archive| {
        let object = archive
            .object_mut(904_899)
            .ok_or_else(|| std::io::Error::other("column bucket is missing"))?;
        object.archive_info.identifier = Some(999_899);
        Ok(())
    })?;
    let invalid_sizes = [
        ("negative zero", -0.0_f32),
        ("negative", -1.0),
        ("NaN", f32::NAN),
    ]
    .into_iter()
    .map(|(case, size)| {
        rewrite_component(&source, COLUMNS, |archive| {
            let object = archive
                .object_mut(904_899)
                .ok_or_else(|| std::io::Error::other("column bucket is missing"))?;
            let message_index = object
                .messages
                .iter()
                .position(|message| message.type_ == 6_006)
                .ok_or_else(|| std::io::Error::other("column bucket message is missing"))?;
            let mut bucket =
                tst::HeaderStorageBucket::decode(object.messages[message_index].data.as_slice())?;
            bucket.headers[0].size = size;
            object.replace_message_preserving_header(
                message_index,
                RawMessage {
                    type_: 6_006,
                    data: bucket.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .map(|bytes| (case, bytes))
    })
    .collect::<TestResult<Vec<_>>>()?;

    for (case, bytes) in [
        ("duplicate nonselected header", malformed_duplicate),
        ("out-of-range nonselected header", out_of_range),
        ("wrong wire", malformed_wire),
        ("missing typed message", wrong_type),
        ("duplicate typed message", duplicate_message),
        ("wrong message metadata", wrong_message_metadata),
        ("wrong object identifier", wrong_object_identifier),
    ]
    .into_iter()
    .chain(invalid_sizes)
    {
        let package = Package::from_bytes(&bytes)?;
        let before = package.exact_bytes();
        let read_error = package
            .table_dimension_size(0usize, 0usize, Dimension::Column(2))
            .expect_err(case);
        assert!(matches!(read_error, TransactionError::InvalidSource { .. }));
        let edit_error = package
            .edit_table_dimension_size(0usize, 0usize, Dimension::Column(2))
            .and_then(|edit| edit.set(points(80.0).unwrap()).commit())
            .expect_err("malformed nonselected header must fail the focused edit");
        assert!(matches!(edit_error, TransactionError::InvalidSource { .. }));
        for rendered in [
            format!("{read_error:?}"),
            read_error.to_string(),
            format!("{edit_error:?}"),
            edit_error.to_string(),
        ] {
            assert!(!rendered.contains(FIXTURE_MARKER));
            assert!(!rendered.contains("Index/"));
            assert!(!rendered.contains("904899"));
        }
        assert_eq!(package.exact_bytes(), before);
    }
    Ok(())
}

#[test]
fn changed_commit_preserves_unknown_header_bytes_and_existing_source_order() -> TestResult {
    const COLUMNS: &str = "Index/Tables/HeaderStorageBucket-904899-2.iwa";
    let source = std::fs::read(fixture_path())?;
    let prepared = rewrite_component(&source, COLUMNS, |archive| {
        let object = archive
            .object_mut(904_899)
            .ok_or_else(|| std::io::Error::other("column bucket is missing"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == 6_006)
            .ok_or_else(|| std::io::Error::other("column bucket message is missing"))?;
        let original = object.messages[message_index].data.clone();
        let raw = repeated_length_delimited_payloads(&original, 2)?;
        let mut reordered = vec![raw[1].to_vec(), raw[0].to_vec(), raw[2].to_vec()];
        append_varint_field(&mut reordered[1], 99, 9_999)?;
        let rewritten = rewrite_repeated_length_delimited_fields(&original, 2, &reordered)?;
        object.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: 6_006,
                data: rewritten,
            },
        )?;
        Ok(())
    })?;
    let package = Package::from_bytes(&prepared)?;
    let before = selected_bucket_payload(&prepared, COLUMNS)?;
    let before_raw = repeated_length_delimited_payloads(&before, 2)?
        .into_iter()
        .map(ToOwned::to_owned)
        .collect::<Vec<_>>();
    assert_eq!(
        before_raw
            .iter()
            .map(|raw| tst::header_storage_bucket::Header::decode(raw.as_slice()))
            .collect::<Result<Vec<_>, _>>()?
            .iter()
            .map(|header| header.index)
            .collect::<Vec<_>>(),
        [5, 1, 6]
    );

    let commit = package
        .edit_table_dimension_size(0usize, 0usize, Dimension::Column(2))?
        .set(points(124.0)?)
        .commit()?;
    let target = commit.package().exact_bytes();
    let after = selected_bucket_payload(&target, COLUMNS)?;
    let after_raw = repeated_length_delimited_payloads(&after, 2)?
        .into_iter()
        .map(ToOwned::to_owned)
        .collect::<Vec<_>>();
    let retained = after_raw
        .iter()
        .filter(|raw| {
            tst::header_storage_bucket::Header::decode(raw.as_slice())
                .is_ok_and(|header| header.index != 2)
        })
        .cloned()
        .collect::<Vec<_>>();
    assert_eq!(retained, before_raw);
    assert_eq!(
        commit
            .package()
            .table_dimension_size(0usize, 0usize, Dimension::Column(2))?,
        points(124.0)?
    );
    assert_native_invariants(commit.package())?;
    Ok(())
}

#[test]
fn row_bucket_authority_and_declared_reference_contradictions_refuse() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let duplicate = rewrite_native_table_model(&source, |_object, _message, model| {
        model
            .base_data_store
            .row_headers
            .buckets
            .push(reference(904_855));
        Ok(())
    })?;
    let mismatched = rewrite_native_table_model(&source, |_object, _message, model| {
        model.base_data_store.row_headers.buckets[0] = reference(904_899);
        Ok(())
    })?;
    let missing_declaration =
        rewrite_native_table_model(&source, |object, message_index, _model| {
            object.archive_info.message_infos[message_index]
                .object_references
                .retain(|identifier| *identifier != 904_855);
            Ok(())
        })?;
    let duplicate_declaration =
        rewrite_native_table_model(&source, |object, message_index, _model| {
            object.archive_info.message_infos[message_index]
                .object_references
                .push(904_855);
            Ok(())
        })?;
    let wrong_field_authority =
        rewrite_native_table_model(&source, |object, message_index, _model| {
            let mut field = FieldInfo::new(vec![4, 2]);
            field.object_references.push(904_855);
            object.archive_info.message_infos[message_index]
                .field_infos
                .push(field);
            Ok(())
        })?;
    let extra_declared_bucket =
        rewrite_native_table_model(&source, |object, message_index, _model| {
            object.archive_info.message_infos[message_index]
                .object_references
                .push(904_937);
            let mut field = FieldInfo::new(vec![4, 1, 2]);
            field.object_references.push(904_937);
            object.archive_info.message_infos[message_index]
                .field_infos
                .push(field);
            Ok(())
        })?;
    let zero_reference = rewrite_native_table_model(&source, |_object, _message, model| {
        model.base_data_store.row_headers.buckets[0] = reference(0);
        Ok(())
    })?;
    let external_reference = rewrite_native_table_model(&source, |_object, _message, model| {
        model.base_data_store.row_headers.buckets[0].deprecated_is_external = Some(true);
        Ok(())
    })?;
    assert_eq!(
        native_table_model(&duplicate)?
            .base_data_store
            .row_headers
            .buckets
            .len(),
        2
    );
    assert_eq!(
        native_table_model(&mismatched)?
            .base_data_store
            .row_headers
            .buckets[0]
            .identifier,
        904_899
    );
    for (case, bytes) in [
        ("duplicate payload reference", duplicate),
        ("cross-axis bucket authority", mismatched),
        ("missing aggregate declaration", missing_declaration),
        ("duplicate aggregate declaration", duplicate_declaration),
        ("wrong field authority", wrong_field_authority),
        ("extra declared row bucket", extra_declared_bucket),
        ("zero row bucket", zero_reference),
        ("external row bucket", external_reference),
    ] {
        let package = Package::from_bytes(&bytes)?;
        let unchanged = package.exact_bytes();
        for dimension in [Dimension::Row(0), Dimension::Column(2)] {
            assert!(
                matches!(
                    package.table_dimension_size(0usize, 0usize, dimension),
                    Err(TransactionError::InvalidSource { .. })
                ),
                "{case} row-bucket authority was accepted while reading {dimension:?}"
            );
            assert!(matches!(
                package
                    .edit_table_dimension_size(0usize, 0usize, dimension)
                    .and_then(|edit| edit.set(points(40.0).unwrap()).commit()),
                Err(TransactionError::InvalidSource { .. })
            ));
        }
        assert_eq!(package.exact_bytes(), unchanged);
    }
    Ok(())
}

#[test]
fn second_row_bucket_routes_by_global_row_without_touching_the_first() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let expanded = expand_native_to_second_row_bucket(&source, 65_537)?;
    let package = Package::from_bytes(&expanded)?;
    let first_before =
        selected_bucket_payload(&expanded, "Index/Tables/HeaderStorageBucket-904855-2.iwa")?;
    assert_eq!(
        package.table_dimension_size(0usize, 0usize, Dimension::Row(65_536))?,
        Size::Default
    );
    let commit = package
        .edit_table_dimension_size(0usize, 0usize, Dimension::Row(65_536))?
        .set(points(40.0)?)
        .commit()?;
    assert_eq!(
        commit
            .package()
            .table_dimension_size(0usize, 0usize, Dimension::Row(65_536))?,
        points(40.0)?
    );
    assert_eq!(
        selected_bucket_payload(
            &commit.package().exact_bytes(),
            "Index/Tables/HeaderStorageBucket-904855-2.iwa",
        )?,
        first_before
    );
    assert_changed_locality(
        &expanded,
        &commit.package().exact_bytes(),
        "Index/Tables/HeaderStorageBucket-904937-2.iwa",
    )?;
    Ok(())
}

#[test]
fn row_bucket_local_index_intervals_are_authoritative() -> TestResult {
    const FIRST: &str = "Index/Tables/HeaderStorageBucket-904855-2.iwa";
    const SECOND: &str = "Index/Tables/HeaderStorageBucket-904937-2.iwa";
    let source = std::fs::read(fixture_path())?;
    let expanded = expand_native_to_second_row_bucket(&source, 70_001)?;
    let hostile_first = rewrite_component(&expanded, FIRST, |archive| {
        let object = archive
            .object_mut(904_855)
            .ok_or_else(|| std::io::Error::other("first row bucket is missing"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == 6_006)
            .ok_or_else(|| std::io::Error::other("first row bucket payload is missing"))?;
        let mut bucket =
            tst::HeaderStorageBucket::decode(object.messages[message_index].data.as_slice())?;
        bucket.headers.push(tst::header_storage_bucket::Header {
            index: 70_000,
            size: 32.0,
            ..Default::default()
        });
        object.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: 6_006,
                data: bucket.encode_to_vec(),
            },
        )?;
        Ok(())
    })?;
    let hostile_second = rewrite_component(&expanded, SECOND, |archive| {
        let object = archive
            .object_mut(904_937)
            .ok_or_else(|| std::io::Error::other("second row bucket is missing"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == 6_006)
            .ok_or_else(|| std::io::Error::other("second row bucket payload is missing"))?;
        let mut bucket =
            tst::HeaderStorageBucket::decode(object.messages[message_index].data.as_slice())?;
        bucket.headers.push(tst::header_storage_bucket::Header {
            index: 1,
            size: 32.0,
            ..Default::default()
        });
        object.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: 6_006,
                data: bucket.encode_to_vec(),
            },
        )?;
        Ok(())
    })?;

    for (bytes, selected) in [
        (hostile_first, Dimension::Row(0)),
        (hostile_second, Dimension::Row(65_536)),
    ] {
        let package = Package::from_bytes(&bytes)?;
        let before = package.exact_bytes();
        assert!(matches!(
            package.table_dimension_size(0usize, 0usize, selected),
            Err(TransactionError::InvalidSource { .. })
        ));
        assert!(matches!(
            package
                .edit_table_dimension_size(0usize, 0usize, selected)
                .and_then(|edit| edit.set(points(40.0).unwrap()).commit()),
            Err(TransactionError::InvalidSource { .. })
        ));
        assert_eq!(package.exact_bytes(), before);
    }
    Ok(())
}

#[test]
fn inverse_transaction_work_limit_is_inclusive_and_max_minus_one_is_atomic() -> TestResult {
    assert!(
        PackageLimits::new(
            PackageLimits::MAX_INPUT_BYTES,
            PackageLimits::MAX_ENTRIES,
            PackageLimits::MAX_ENTRY_BYTES,
            0,
            PackageLimits::MAX_IWA_STREAM_BYTES,
        )
        .is_err()
    );
    let package = Package::open(fixture_path())?;
    let source = package.exact_bytes();
    let commit = package
        .edit_table_dimension_size(0usize, 0usize, Dimension::Column(2))?
        .set(points(124.0)?)
        .commit()?;
    let candidate = commit.package().exact_bytes();
    let inverse = commit.patch().inverse();
    let observed = transaction_work_precharge(&candidate, &source)?;

    let package_with_limit = |maximum: usize| -> TestResult<Package> {
        let limits = PackageLimits::new(
            PackageLimits::MAX_INPUT_BYTES,
            PackageLimits::MAX_ENTRIES,
            PackageLimits::MAX_ENTRY_BYTES,
            u64::try_from(maximum)?,
            PackageLimits::MAX_IWA_STREAM_BYTES,
        )?;
        Ok(Package::from_bytes_with_options(
            &candidate,
            PackageReadOptions::new(limits, PackageSemanticLimits::default()),
        )?)
    };

    let exact = package_with_limit(observed)?;
    assert_eq!(
        exact
            .apply_table_dimension_size(&inverse)?
            .package()
            .exact_bytes(),
        source
    );

    let maximum = observed - 1;
    let restricted = package_with_limit(maximum)?;
    let unchanged = restricted.exact_bytes();
    let error = restricted
        .apply_table_dimension_size(&inverse)
        .expect_err("max-minus-one transaction work must refuse before reopen");
    assert!(matches!(
        error,
        TransactionError::LimitExceeded {
            kind: LimitKind::TransactionWork,
            observed: actual,
            maximum: actual_maximum,
            path: Path::Package,
        } if actual == u64::try_from(observed)? && actual_maximum == u64::try_from(maximum)?
    ));
    assert_eq!(restricted.exact_bytes(), unchanged);
    Ok(())
}

#[test]
fn cumulative_public_read_and_changed_commit_limits_are_inclusive() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    assert!(
        PackageSemanticLimits::new(
            PackageSemanticLimits::MAX_OBJECTS,
            PackageSemanticLimits::MAX_SHEETS,
            PackageSemanticLimits::MAX_TABLES,
            0,
        )
        .is_err()
    );

    let read_exact = minimum_reference_limit(&source, |package| {
        package
            .table_dimension_size(0usize, 0usize, Dimension::Column(2))
            .is_ok()
    })?;
    assert!(read_exact > 1);
    assert_eq!(
        package_with_reference_limit(&source, read_exact)?.table_dimension_size(
            0usize,
            0usize,
            Dimension::Column(2),
        )?,
        Size::Default
    );
    let read_restricted = package_with_reference_limit(&source, read_exact - 1)?;
    let read_before = read_restricted.exact_bytes();
    assert!(matches!(
        read_restricted.table_dimension_size(0usize, 0usize, Dimension::Column(2)),
        Err(TransactionError::LimitExceeded {
            kind: LimitKind::PayloadReferences,
            observed,
            maximum,
            path: Path::Dimension {
                sheet: 0,
                table: 0,
                dimension: Dimension::Column(2),
            },
        }) if observed == u64::try_from(read_exact)?
            && maximum == u64::try_from(read_exact - 1)?
    ));
    assert_eq!(read_restricted.exact_bytes(), read_before);

    let commit_exact = minimum_reference_limit(&source, |package| {
        package
            .edit_table_dimension_size(0usize, 0usize, Dimension::Column(2))
            .and_then(|edit| edit.set(Size::points(124.0).unwrap()).commit())
            .is_ok()
    })?;
    assert!(commit_exact >= read_exact);
    let exact_commit = package_with_reference_limit(&source, commit_exact)?
        .edit_table_dimension_size(0usize, 0usize, Dimension::Column(2))?
        .set(points(124.0)?)
        .commit()?;
    assert_eq!(
        exact_commit
            .package()
            .table_dimension_size(0usize, 0usize, Dimension::Column(2))?,
        points(124.0)?
    );
    let restricted = package_with_reference_limit(&source, commit_exact - 1)?;
    let before = restricted.exact_bytes();
    let error = restricted
        .edit_table_dimension_size(0usize, 0usize, Dimension::Column(2))
        .and_then(|edit| edit.set(points(124.0).unwrap()).commit())
        .expect_err("cumulative max-minus-one commit must refuse atomically");
    assert!(matches!(
        error,
        TransactionError::LimitExceeded {
            kind: LimitKind::PayloadReferences,
            observed,
            maximum,
            path: Path::Dimension {
                sheet: 0,
                table: 0,
                dimension: Dimension::Column(2),
            },
        } if observed == u64::try_from(commit_exact)?
            && maximum == u64::try_from(commit_exact - 1)?
    ));
    assert_eq!(restricted.exact_bytes(), before);
    Ok(())
}
