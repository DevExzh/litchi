//! Native Numbers integration coverage for document-scoped Custom Date & Time.
//!
//! The source workbook was authored and reopened by Numbers 14.4. B2 stores a
//! date value with the native `Native Calendar` Custom Date & Time format;
//! C2 is an adjacent marker used to prove that format edits preserve both the
//! semantic value and its native cached scalar.

use std::{io, path::PathBuf};

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::wire::{
    WireView, append_length_delimited_field, patch_length_delimited_field,
};
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::tst;
use litchi_numbers::cell::{
    Value,
    data_format::custom::{Custom, DateTime as CustomDateTime, DateTimePattern, Name},
};
use litchi_numbers::table::cells::Storage;
use litchi_numbers::{CellPosition, Package};
use litchi_numbers_wire::{BncCell, CachedScalar, StoredValue};
use prost::Message as _;

#[path = "support/table_cell_data_format_fixture.rs"]
mod fixture;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const SHEET_NAME: &str = "Sheet 1";
const TABLE_NAME: &str = "Table 1";
const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const DOCUMENT_ROOT_ID: u64 = 1;
const DOCUMENT_SUPER_FIELD: u32 = 8;
const DOCUMENT_REGISTRY_FIELD: u32 = 9;
const TSA_REGISTRY_FIELD: u32 = 12;
const FORMAT_MEMBER: &str = "Index/Tables/DataList-904498-2.iwa";
const TILE_MEMBER: &str = "Index/Tables/Tile.iwa";

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/custom-datetime-native.numbers")
}

fn resaved_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/custom-datetime-native-resaved.numbers")
}

fn position() -> CellPosition {
    CellPosition::new(1, 1)
}

fn marker_position() -> CellPosition {
    CellPosition::new(1, 2)
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn native_datetime(package: &Package) -> TestResult<CustomDateTime> {
    let format = package
        .table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())?
        .ok_or_else(|| io::Error::other("native Custom Date & Time format is missing"))?;
    match format {
        Custom::DateTime(date_time) => Ok(date_time),
        Custom::Number(_) | Custom::Text(_) => {
            Err(io::Error::other("native fixture returned the wrong Custom family").into())
        },
    }
}

fn date_value(package: &Package) -> TestResult<f64> {
    let cell = package.table_cell(SHEET_NAME, TABLE_NAME, position())?;
    match cell.storage() {
        Storage::Stored(Value::Date(value)) => Ok(value.get()),
        storage => Err(io::Error::other(format!(
            "B2 value mismatch: expected Stored(Date), got {storage:?}"
        ))
        .into()),
    }
}

fn date_cache(source: &[u8]) -> TestResult<CachedScalar> {
    let cell = BncCell::parse(&native_date_cell(source)?)?;
    if cell.stored_value() != StoredValue::Date {
        return Err(io::Error::other("native Date & Time B2 cell is not a date").into());
    }
    match cell.cached_scalar()? {
        Some(cache @ CachedScalar::Date(_)) => Ok(cache),
        Some(cache) => Err(io::Error::other(format!(
            "native Date & Time B2 cache has the wrong family: {cache:?}"
        ))
        .into()),
        None => Err(io::Error::other("native Date & Time B2 cache is missing").into()),
    }
}

fn native_date_cell(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == TILE_MEMBER)
        .ok_or_else(|| io::Error::other("native tile member is missing"))?;
    let stream = SnappyStream::decompress(entry.data())?;
    let archive = Archive::parse(stream.as_bytes())?;
    let tile = archive
        .objects
        .iter()
        .flat_map(|object| object.messages.iter())
        .filter(|message| message.type_ == 6_002)
        .find_map(|message| {
            tst::Tile::decode(message.data.as_slice())
                .ok()
                .filter(|tile| {
                    tile.row_infos
                        .iter()
                        .any(|row| row.tile_row_index == position().row())
                })
        })
        .ok_or_else(|| io::Error::other("native Date & Time tile message is missing"))?;
    let row = tile
        .row_infos
        .iter()
        .find(|row| row.tile_row_index == position().row())
        .ok_or_else(|| io::Error::other("native Date & Time tile row is missing"))?;
    let offsets = row
        .cell_offsets
        .as_deref()
        .ok_or_else(|| io::Error::other("native Date & Time tile offsets are missing"))?;
    let column = usize::try_from(position().column())?;
    let offset_start = column
        .checked_mul(2)
        .ok_or_else(|| io::Error::other("native Date & Time offset range overflow"))?;
    let offset_end = offset_start
        .checked_add(2)
        .ok_or_else(|| io::Error::other("native Date & Time offset range overflow"))?;
    let slot = offsets
        .get(offset_start..offset_end)
        .ok_or_else(|| io::Error::other("native Date & Time cell offset is missing"))?;
    let raw_start = u16::from_le_bytes([slot[0], slot[1]]);
    if raw_start == u16::MAX {
        return Err(io::Error::other("native Date & Time B2 cell is absent").into());
    }
    let width = if row.has_wide_offsets.unwrap_or(false) {
        4usize
    } else {
        1usize
    };
    let start = usize::from(raw_start)
        .checked_mul(width)
        .ok_or_else(|| io::Error::other("native Date & Time cell offset overflow"))?;
    let end = offsets
        .chunks_exact(2)
        .skip(column + 1)
        .find_map(|next| {
            let raw = u16::from_le_bytes([next[0], next[1]]);
            (raw != u16::MAX).then(|| usize::from(raw).saturating_mul(width))
        })
        .unwrap_or_else(|| row.cell_storage_buffer.as_ref().map_or(0, Vec::len));
    let storage = row
        .cell_storage_buffer
        .as_deref()
        .ok_or_else(|| io::Error::other("native Date & Time cell storage is missing"))?;
    storage
        .get(start..end)
        .map(ToOwned::to_owned)
        .ok_or_else(|| io::Error::other("native Date & Time cell range is invalid").into())
}

fn assert_cell_values(
    package: &Package,
    expected_date: f64,
    expected_cache: CachedScalar,
    marker: &str,
) -> TestResult {
    let actual_date = date_value(package)?;
    assert_eq!(actual_date.to_bits(), expected_date.to_bits());
    assert_eq!(date_cache(&exact_bytes(package)?)?, expected_cache);

    let marker_cell = package.table_cell(SHEET_NAME, TABLE_NAME, marker_position())?;
    match marker_cell.storage() {
        Storage::Stored(Value::Text(value)) => assert_eq!(value, marker),
        storage => panic!("C2 marker mismatch: expected Stored(Text({marker:?})), got {storage:?}"),
    }
    Ok(())
}

fn assert_datetime(package: &Package, expected: &CustomDateTime) -> TestResult {
    let actual = native_datetime(package)?;
    assert_eq!(actual.name().as_str(), expected.name().as_str());
    assert_eq!(actual.pattern().as_str(), expected.pattern().as_str());
    Ok(())
}

fn assert_exact_locality(source: &[u8], target: &[u8]) -> TestResult {
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    for member in [fixture::DOCUMENT_MEMBER, FORMAT_MEMBER, TILE_MEMBER] {
        if !before.iter().any(|entry| entry.name() == member) {
            return Err(io::Error::other(format!(
                "native Date & Time fixture member {member} is missing"
            ))
            .into());
        }
    }
    let mut changed = Vec::new();
    for entry in before.iter() {
        let candidate = after
            .iter()
            .find(|other| other.name() == entry.name())
            .ok_or_else(|| io::Error::other("native edit removed a source member"))?;
        if entry.data() != candidate.data() {
            changed.push(entry.name().to_owned());
        } else {
            assert_eq!(
                entry.raw_record().local_record(),
                candidate.raw_record().local_record(),
                "unchanged member {} lost its exact local ZIP record",
                entry.name()
            );
        }
    }
    changed.sort_unstable();
    let mut expected = vec![
        fixture::DOCUMENT_MEMBER.to_owned(),
        FORMAT_MEMBER.to_owned(),
        TILE_MEMBER.to_owned(),
    ];
    expected.sort_unstable();
    assert_eq!(changed, expected);
    assert_eq!(before.len(), after.len());
    Ok(())
}

#[derive(Debug, Clone, Copy)]
enum RouteCorruption {
    DuplicateTsa12,
    DuplicateTnSuper8,
    MixedRoot9AndTsa12,
    MalformedReference,
}

fn rewrite_document_root(source: &[u8], corruption: RouteCorruption) -> TestResult<Vec<u8>> {
    fixture::rewrite_member(source, fixture::DOCUMENT_MEMBER, |archive| {
        let document = archive
            .object_mut(DOCUMENT_ROOT_ID)
            .ok_or_else(|| io::Error::other("native document root is missing"))?;
        let message = document
            .messages
            .iter_mut()
            .find(|message| message.type_ == DOCUMENT_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("native document message is missing"))?;
        let root = message.data.clone();
        message.data = corrupt_document_route(&root, corruption)?;
        Ok(())
    })
}

fn corrupt_document_route(source: &[u8], corruption: RouteCorruption) -> TestResult<Vec<u8>> {
    let root = WireView::parse(source)?;
    let super_field = root
        .fields()
        .find(|field| field.number() == DOCUMENT_SUPER_FIELD)
        .ok_or_else(|| io::Error::other("native TSA super envelope is missing"))?;
    let tsa = WireView::parse(super_field.payload())?;
    let registry = tsa
        .fields()
        .find(|field| field.number() == TSA_REGISTRY_FIELD)
        .ok_or_else(|| io::Error::other("native TSA field-12 registry reference is missing"))?;

    match corruption {
        RouteCorruption::DuplicateTsa12 => {
            let mut tsa_payload = super_field.payload().to_vec();
            append_length_delimited_field(
                &mut tsa_payload,
                TSA_REGISTRY_FIELD,
                registry.payload(),
            )?;
            Ok(patch_length_delimited_field(
                source,
                DOCUMENT_SUPER_FIELD,
                true,
                Some(&tsa_payload),
            )?)
        },
        RouteCorruption::DuplicateTnSuper8 => {
            let mut root_payload = source.to_vec();
            append_length_delimited_field(
                &mut root_payload,
                DOCUMENT_SUPER_FIELD,
                super_field.payload(),
            )?;
            Ok(root_payload)
        },
        RouteCorruption::MixedRoot9AndTsa12 => {
            let mut root_payload = source.to_vec();
            append_length_delimited_field(
                &mut root_payload,
                DOCUMENT_REGISTRY_FIELD,
                registry.payload(),
            )?;
            Ok(root_payload)
        },
        RouteCorruption::MalformedReference => {
            let tsa_payload = patch_length_delimited_field(
                super_field.payload(),
                TSA_REGISTRY_FIELD,
                true,
                Some(&[0x80]),
            )?;
            Ok(patch_length_delimited_field(
                source,
                DOCUMENT_SUPER_FIELD,
                true,
                Some(&tsa_payload),
            )?)
        },
    }
}

#[test]
fn native_custom_datetime_source_read_is_typed_and_exact_noop() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::open(fixture_path())?;
    assert_eq!(exact_bytes(&package)?, source);

    let original = native_datetime(&package)?;
    assert_eq!(original.name().as_str(), "Native Calendar");
    assert_eq!(original.pattern().as_str(), "MMM d, y");
    let expected_date = date_value(&package)?;
    let expected_cache = date_cache(&source)?;
    assert_cell_values(
        &package,
        expected_date,
        expected_cache,
        "Native DateTime marker",
    )?;

    let no_op = package
        .edit_table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())?
        .set(Custom::DateTime(original.clone()))
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(exact_bytes(no_op.package())?, source);
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());
    assert_datetime(no_op.package(), &original)?;
    assert_cell_values(
        no_op.package(),
        expected_date,
        expected_cache,
        "Native DateTime marker",
    )?;
    Ok(())
}

#[test]
fn native_custom_datetime_replacement_reopens_preserves_values_cache_and_inverse() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::open(fixture_path())?;
    let original = native_datetime(&package)?;
    let original_custom = Custom::DateTime(original.clone());
    let expected_date = date_value(&package)?;
    let expected_cache = date_cache(&source)?;
    let replacement = CustomDateTime::new(
        Name::new("Rust Calendar")?,
        DateTimePattern::new("yyyy-MM-dd HH:mm:ss")?,
    );
    let replacement_custom = Custom::DateTime(replacement.clone());

    let changed = package
        .edit_table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())?
        .set(replacement_custom.clone())
        .commit()?;
    let target = exact_bytes(changed.package())?;
    assert!(!changed.patch().is_noop());
    assert_eq!(changed.patch().before(), Some(&original_custom));
    assert_eq!(changed.patch().after(), Some(&replacement_custom));
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().touched_components(), 3);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_datetime(changed.package(), &replacement)?;
    assert_cell_values(
        changed.package(),
        expected_date,
        expected_cache,
        "Native DateTime marker",
    )?;
    assert_exact_locality(&source, &target)?;

    let reopened = Package::from_bytes(&target)?;
    assert_datetime(&reopened, &replacement)?;
    assert_cell_values(
        &reopened,
        expected_date,
        expected_cache,
        "Native DateTime marker",
    )?;

    let applied = package.apply_table_cell_custom_format(changed.patch())?;
    assert_eq!(exact_bytes(applied.package())?, target);
    let inverse = changed.patch().inverse();
    assert_eq!(inverse.inverse(), *changed.patch());
    let restored = reopened.apply_table_cell_custom_format(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_datetime(restored.package(), &original)?;
    assert_cell_values(
        restored.package(),
        expected_date,
        expected_cache,
        "Native DateTime marker",
    )?;
    Ok(())
}

#[test]
fn native_custom_datetime_clear_preserves_value_cache_locality_and_exact_inverse() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::open(fixture_path())?;
    let original = native_datetime(&package)?;
    let original_custom = Custom::DateTime(original.clone());
    let expected_date = date_value(&package)?;
    let expected_cache = date_cache(&source)?;

    let cleared = package
        .edit_table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())?
        .clear()
        .commit()?;
    let target = exact_bytes(cleared.package())?;
    assert_eq!(cleared.patch().before(), Some(&original_custom));
    assert_eq!(cleared.patch().after(), None);
    assert!(cleared.diagnostics().changed());
    assert_eq!(cleared.diagnostics().touched_components(), 3);
    assert!(cleared.diagnostics().full_reparse_performed());
    assert_eq!(
        cleared
            .package()
            .table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())?,
        None
    );
    assert_cell_values(
        cleared.package(),
        expected_date,
        expected_cache,
        "Native DateTime marker",
    )?;
    assert_exact_locality(&source, &target)?;

    let reopened = Package::from_bytes(&target)?;
    assert_eq!(
        reopened.table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())?,
        None
    );
    assert_cell_values(
        &reopened,
        expected_date,
        expected_cache,
        "Native DateTime marker",
    )?;

    let inverse = cleared.patch().inverse();
    assert_eq!(inverse.inverse(), *cleared.patch());
    let restored = reopened.apply_table_cell_custom_format(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_datetime(restored.package(), &original)?;
    assert_cell_values(
        restored.package(),
        expected_date,
        expected_cache,
        "Native DateTime marker",
    )?;
    Ok(())
}

#[test]
fn native_custom_datetime_resaved_fixture_reopens_clears_and_inverts_exactly() -> TestResult {
    let source = std::fs::read(resaved_fixture_path())?;
    let package = Package::open(resaved_fixture_path())?;
    assert_eq!(exact_bytes(&package)?, source);

    let original = native_datetime(&package)?;
    assert_eq!(original.name().as_str(), "Rust Calendar");
    assert_eq!(original.pattern().as_str(), "yyyy-MM-dd HH:mm:ss");
    let original_custom = Custom::DateTime(original.clone());
    let expected_date = date_value(&package)?;
    let expected_cache = date_cache(&source)?;
    assert_cell_values(
        &package,
        expected_date,
        expected_cache,
        "Native DateTime marker saved",
    )?;

    let no_op = package
        .edit_table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())?
        .set(Custom::DateTime(original.clone()))
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(exact_bytes(no_op.package())?, source);
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());
    assert_datetime(no_op.package(), &original)?;
    assert_cell_values(
        no_op.package(),
        expected_date,
        expected_cache,
        "Native DateTime marker saved",
    )?;

    let cleared = package
        .edit_table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())?
        .clear()
        .commit()?;
    let cleared_bytes = exact_bytes(cleared.package())?;
    assert_eq!(cleared.patch().before(), Some(&original_custom));
    assert_eq!(cleared.patch().after(), None);
    assert!(cleared.diagnostics().changed());
    assert_eq!(cleared.diagnostics().touched_components(), 3);
    assert!(cleared.diagnostics().full_reparse_performed());
    assert_eq!(
        cleared
            .package()
            .table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())?,
        None
    );
    assert_cell_values(
        cleared.package(),
        expected_date,
        expected_cache,
        "Native DateTime marker saved",
    )?;
    assert_exact_locality(&source, &cleared_bytes)?;

    let reopened = Package::from_bytes(&cleared_bytes)?;
    assert_eq!(
        reopened.table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())?,
        None
    );
    assert_cell_values(
        &reopened,
        expected_date,
        expected_cache,
        "Native DateTime marker saved",
    )?;

    let inverse = cleared.patch().inverse();
    assert_eq!(inverse.inverse(), *cleared.patch());
    let restored = reopened.apply_table_cell_custom_format(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_datetime(restored.package(), &original)?;
    assert_cell_values(
        restored.package(),
        expected_date,
        expected_cache,
        "Native DateTime marker saved",
    )?;
    Ok(())
}

#[test]
fn native_custom_datetime_registry_routes_fail_closed_and_atomically() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    for corruption in [
        RouteCorruption::DuplicateTsa12,
        RouteCorruption::DuplicateTnSuper8,
        RouteCorruption::MixedRoot9AndTsa12,
        RouteCorruption::MalformedReference,
    ] {
        let hostile = rewrite_document_root(&source, corruption)?;
        assert_ne!(hostile, source);
        let Ok(package) = Package::from_bytes(&hostile) else {
            continue;
        };
        let before = exact_bytes(&package)?;
        assert!(
            package
                .table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())
                .is_err(),
            "hostile native Date & Time route {corruption:?} was accepted by the reader"
        );
        assert!(
            package
                .edit_table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())
                .is_err(),
            "hostile native Date & Time route {corruption:?} was accepted by the editor"
        );
        assert_eq!(exact_bytes(&package)?, before);
    }
    Ok(())
}
