//! Native Numbers integration coverage for an ordinary Duration format.
//!
//! The source workbook was authored and reopened by Numbers 14.4. B2 stores a
//! duration value entered as `1h 23m 45s` with an explicit abbreviated,
//! automatic-units Duration format; C2 is an adjacent marker used to prove
//! that format edits preserve both the semantic value and its native cache.

use std::{io, path::PathBuf};

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::tst;
use litchi_numbers::cell::{
    Value,
    data_format::duration::{Duration, Style, Unit, UnitRange, Units},
};
use litchi_numbers::table::cells::Storage;
use litchi_numbers::{CellPosition, Package};
use litchi_numbers_wire::{BncCell, CachedScalar, StoredValue};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const SHEET_NAME: &str = "Sheet 1";
const TABLE_NAME: &str = "Table 1";
const TILE_MEMBER: &str = "Index/Tables/Tile.iwa";
const FORMAT_MEMBER: &str = "Index/Tables/DataList-904498-2.iwa";

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/duration-native.numbers")
}

fn resaved_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/duration-native-resaved.numbers")
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

fn native_duration(package: &Package) -> TestResult<Duration> {
    package
        .table_cell_duration_format(SHEET_NAME, TABLE_NAME, position())?
        .ok_or_else(|| io::Error::other("native Duration format is missing").into())
}

fn duration_value(package: &Package) -> TestResult<f64> {
    let cell = package.table_cell(SHEET_NAME, TABLE_NAME, position())?;
    match cell.storage() {
        Storage::Stored(Value::Duration(value)) => Ok(value.get()),
        storage => Err(io::Error::other(format!(
            "B2 value mismatch: expected Stored(Duration), got {storage:?}"
        ))
        .into()),
    }
}

fn duration_cache(source: &[u8]) -> TestResult<CachedScalar> {
    let cell = BncCell::parse(&native_duration_cell(source)?)?;
    if cell.stored_value() != StoredValue::Duration {
        return Err(io::Error::other("native Duration B2 cell is not a duration").into());
    }
    match cell.cached_scalar()? {
        Some(cache @ CachedScalar::Duration(_)) => Ok(cache),
        Some(cache) => Err(io::Error::other(format!(
            "native Duration B2 cache has the wrong family: {cache:?}"
        ))
        .into()),
        None => Err(io::Error::other("native Duration B2 cache is missing").into()),
    }
}

fn native_duration_cell(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == TILE_MEMBER)
        .ok_or_else(|| io::Error::other("native Duration tile member is missing"))?;
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
        .ok_or_else(|| io::Error::other("native Duration tile message is missing"))?;
    let row = tile
        .row_infos
        .iter()
        .find(|row| row.tile_row_index == position().row())
        .ok_or_else(|| io::Error::other("native Duration tile row is missing"))?;
    let offsets = row
        .cell_offsets
        .as_deref()
        .ok_or_else(|| io::Error::other("native Duration tile offsets are missing"))?;
    let column = usize::try_from(position().column())?;
    let offset_start = column
        .checked_mul(2)
        .ok_or_else(|| io::Error::other("native Duration offset range overflow"))?;
    let offset_end = offset_start
        .checked_add(2)
        .ok_or_else(|| io::Error::other("native Duration offset range overflow"))?;
    let slot = offsets
        .get(offset_start..offset_end)
        .ok_or_else(|| io::Error::other("native Duration cell offset is missing"))?;
    let raw_start = u16::from_le_bytes([slot[0], slot[1]]);
    if raw_start == u16::MAX {
        return Err(io::Error::other("native Duration B2 cell is absent").into());
    }
    let width = if row.has_wide_offsets.unwrap_or(false) {
        4usize
    } else {
        1usize
    };
    let start = usize::from(raw_start)
        .checked_mul(width)
        .ok_or_else(|| io::Error::other("native Duration cell offset overflow"))?;
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
        .ok_or_else(|| io::Error::other("native Duration cell storage is missing"))?;
    storage
        .get(start..end)
        .map(ToOwned::to_owned)
        .ok_or_else(|| io::Error::other("native Duration cell range is invalid").into())
}

fn assert_cell_values(
    package: &Package,
    expected_duration: f64,
    expected_cache: CachedScalar,
    marker: &str,
) -> TestResult {
    let actual_duration = duration_value(package)?;
    assert_eq!(actual_duration.to_bits(), expected_duration.to_bits());
    assert_eq!(duration_cache(&exact_bytes(package)?)?, expected_cache);

    let marker_cell = package.table_cell(SHEET_NAME, TABLE_NAME, marker_position())?;
    match marker_cell.storage() {
        Storage::Stored(Value::Text(value)) => assert_eq!(value, marker),
        storage => panic!("C2 marker mismatch: expected Stored(Text({marker:?})), got {storage:?}"),
    }
    Ok(())
}

fn assert_duration(package: &Package, expected: Duration) -> TestResult {
    assert_eq!(native_duration(package)?, expected);
    Ok(())
}

fn assert_exact_locality(source: &[u8], target: &[u8]) -> TestResult {
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    for member in [FORMAT_MEMBER, TILE_MEMBER] {
        if !before.iter().any(|entry| entry.name() == member) {
            return Err(io::Error::other(format!(
                "native Duration fixture member {member} is missing"
            ))
            .into());
        }
    }
    let mut changed = Vec::new();
    for entry in before.iter() {
        let candidate = after
            .iter()
            .find(|other| other.name() == entry.name())
            .ok_or_else(|| io::Error::other("native Duration edit removed a source member"))?;
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
    let mut expected = vec![FORMAT_MEMBER.to_owned(), TILE_MEMBER.to_owned()];
    expected.sort_unstable();
    assert_eq!(changed, expected);
    assert_eq!(before.len(), after.len());
    Ok(())
}

fn native_baseline() -> TestResult<Duration> {
    Ok(Duration::new(
        Style::Abbreviated,
        Units::Automatic(UnitRange::new(Unit::Hours, Unit::Seconds)?),
    ))
}

fn replacement_duration() -> TestResult<Duration> {
    Ok(Duration::custom(
        Style::Colon,
        UnitRange::new(Unit::Hours, Unit::Seconds)?,
    ))
}

#[test]
fn native_duration_source_read_is_typed_and_exact_noop() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::open(fixture_path())?;
    assert_eq!(exact_bytes(&package)?, source);

    let original = native_duration(&package)?;
    assert_eq!(original, native_baseline()?);
    let expected_duration = duration_value(&package)?;
    let expected_cache = duration_cache(&source)?;
    assert_cell_values(
        &package,
        expected_duration,
        expected_cache,
        "Native Duration marker",
    )?;

    let no_op = package
        .edit_table_cell_duration_format(SHEET_NAME, TABLE_NAME, position())?
        .set(original)
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(exact_bytes(no_op.package())?, source);
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert_eq!(no_op.diagnostics().deleted_previews(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());
    assert_duration(no_op.package(), original)?;
    assert_cell_values(
        no_op.package(),
        expected_duration,
        expected_cache,
        "Native Duration marker",
    )?;
    Ok(())
}

#[test]
fn native_duration_replacement_reopens_preserves_value_cache_and_inverse() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::open(fixture_path())?;
    let original = native_duration(&package)?;
    assert_eq!(original, native_baseline()?);
    let expected_duration = duration_value(&package)?;
    let expected_cache = duration_cache(&source)?;
    let replacement = replacement_duration()?;

    let changed = package
        .edit_table_cell_duration_format(SHEET_NAME, TABLE_NAME, position())?
        .set(replacement)
        .commit()?;
    let target = exact_bytes(changed.package())?;
    assert!(!changed.patch().is_noop());
    assert_eq!(changed.patch().before(), Some(&original));
    assert_eq!(changed.patch().after(), Some(&replacement));
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().touched_components(), 2);
    assert_eq!(changed.diagnostics().deleted_previews(), 0);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_duration(changed.package(), replacement)?;
    assert_cell_values(
        changed.package(),
        expected_duration,
        expected_cache,
        "Native Duration marker",
    )?;
    assert_exact_locality(&source, &target)?;

    let reopened = Package::from_bytes(&target)?;
    assert_duration(&reopened, replacement)?;
    assert_cell_values(
        &reopened,
        expected_duration,
        expected_cache,
        "Native Duration marker",
    )?;

    let applied = package.apply_table_cell_duration_format(changed.patch())?;
    assert_eq!(exact_bytes(applied.package())?, target);
    let inverse = changed.patch().inverse();
    assert_eq!(inverse.inverse(), *changed.patch());
    let restored = reopened.apply_table_cell_duration_format(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_duration(restored.package(), original)?;
    assert_cell_values(
        restored.package(),
        expected_duration,
        expected_cache,
        "Native Duration marker",
    )?;
    Ok(())
}

#[test]
fn native_duration_clear_reopens_preserves_value_cache_locality_and_inverse() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::open(fixture_path())?;
    let original = native_duration(&package)?;
    assert_eq!(original, native_baseline()?);
    let expected_duration = duration_value(&package)?;
    let expected_cache = duration_cache(&source)?;

    let cleared = package
        .edit_table_cell_duration_format(SHEET_NAME, TABLE_NAME, position())?
        .clear()
        .commit()?;
    let target = exact_bytes(cleared.package())?;
    assert_eq!(cleared.patch().before(), Some(&original));
    assert_eq!(cleared.patch().after(), None);
    assert!(cleared.diagnostics().changed());
    assert_eq!(cleared.diagnostics().touched_components(), 2);
    assert_eq!(cleared.diagnostics().deleted_previews(), 0);
    assert!(cleared.diagnostics().full_reparse_performed());
    assert_eq!(
        cleared
            .package()
            .table_cell_duration_format(SHEET_NAME, TABLE_NAME, position())?,
        None
    );
    assert_cell_values(
        cleared.package(),
        expected_duration,
        expected_cache,
        "Native Duration marker",
    )?;
    assert_exact_locality(&source, &target)?;

    let reopened = Package::from_bytes(&target)?;
    assert_eq!(
        reopened.table_cell_duration_format(SHEET_NAME, TABLE_NAME, position())?,
        None
    );
    assert_cell_values(
        &reopened,
        expected_duration,
        expected_cache,
        "Native Duration marker",
    )?;

    let inverse = cleared.patch().inverse();
    assert_eq!(inverse.inverse(), *cleared.patch());
    let restored = reopened.apply_table_cell_duration_format(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_duration(restored.package(), original)?;
    assert_cell_values(
        restored.package(),
        expected_duration,
        expected_cache,
        "Native Duration marker",
    )?;
    Ok(())
}

#[test]
fn native_duration_resaved_fixture_reopens_clears_and_inverts_exactly() -> TestResult {
    let source = std::fs::read(resaved_fixture_path())?;
    let package = Package::open(resaved_fixture_path())?;
    assert_eq!(exact_bytes(&package)?, source);

    let original = native_duration(&package)?;
    assert_eq!(original, replacement_duration()?);
    let expected_duration = duration_value(&package)?;
    let expected_cache = duration_cache(&source)?;
    assert_cell_values(
        &package,
        expected_duration,
        expected_cache,
        "Native Duration marker saved",
    )?;

    let no_op = package
        .edit_table_cell_duration_format(SHEET_NAME, TABLE_NAME, position())?
        .set(original)
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(exact_bytes(no_op.package())?, source);
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert_eq!(no_op.diagnostics().deleted_previews(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());
    assert_duration(no_op.package(), original)?;
    assert_cell_values(
        no_op.package(),
        expected_duration,
        expected_cache,
        "Native Duration marker saved",
    )?;

    let cleared = package
        .edit_table_cell_duration_format(SHEET_NAME, TABLE_NAME, position())?
        .clear()
        .commit()?;
    let target = exact_bytes(cleared.package())?;
    assert_eq!(cleared.patch().before(), Some(&original));
    assert_eq!(cleared.patch().after(), None);
    assert!(cleared.diagnostics().changed());
    assert_eq!(cleared.diagnostics().touched_components(), 2);
    assert_eq!(cleared.diagnostics().deleted_previews(), 0);
    assert!(cleared.diagnostics().full_reparse_performed());
    assert_eq!(
        cleared
            .package()
            .table_cell_duration_format(SHEET_NAME, TABLE_NAME, position())?,
        None
    );
    assert_cell_values(
        cleared.package(),
        expected_duration,
        expected_cache,
        "Native Duration marker saved",
    )?;
    assert_exact_locality(&source, &target)?;

    let reopened = Package::from_bytes(&target)?;
    assert_eq!(
        reopened.table_cell_duration_format(SHEET_NAME, TABLE_NAME, position())?,
        None
    );
    assert_cell_values(
        &reopened,
        expected_duration,
        expected_cache,
        "Native Duration marker saved",
    )?;

    let inverse = cleared.patch().inverse();
    assert_eq!(inverse.inverse(), *cleared.patch());
    let restored = reopened.apply_table_cell_duration_format(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_duration(restored.package(), original)?;
    assert_cell_values(
        restored.package(),
        expected_duration,
        expected_cache,
        "Native Duration marker saved",
    )?;
    Ok(())
}
