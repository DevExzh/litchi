//! Public integration coverage for selector-first Numbers cell reads and batches.

use std::{fmt::Debug, path::PathBuf};

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::tst;
use litchi_numbers::{
    Package, SheetSelector, TableSelector,
    cell::Value,
    table::{
        CellPosition, CellRange, Dimensions,
        cells::{
            Change, Commit, DependencyKind, Diagnostics, Edit, Error, Input, LimitKind, Patch,
            Path, State, Storage,
        },
    },
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const FIXTURE_MARKER: &str = "Litchi native Numbers fixture";
const CANONICAL_PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/numbers/basic.numbers")
}

fn assert_send_sync_debug<T: Send + Sync + Debug>() {}

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

fn assert_exact_scalar_locality(
    source: &[u8],
    target: &[u8],
    expected_changed_members: &[&str],
) -> TestResult {
    let source = Catalog::from_bytes(source)?;
    let target = Catalog::from_bytes(target)?;
    let mut changed_members = Vec::new();

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
            changed_members.push(before.name());
        }
    }
    assert_eq!(changed_members, expected_changed_members);
    assert_eq!(target.len() + CANONICAL_PREVIEWS.len(), source.len());
    Ok(())
}

fn string_lists(package: &[u8]) -> TestResult<Vec<tst::TableDataList>> {
    let catalog = Catalog::from_bytes(package)?;
    let mut lists = Vec::new();
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
            .filter(|message| matches!(message.type_, 6_005 | 6_201))
        {
            let Ok(list) = tst::TableDataList::decode(message.data.as_slice()) else {
                continue;
            };
            if list.list_type == tst::table_data_list::ListType::String as i32 {
                lists.push(list);
            }
        }
    }
    Ok(lists)
}

fn string_list_with_text(package: &[u8], text: &str) -> TestResult<tst::TableDataList> {
    let mut lists = string_lists(package)?
        .into_iter()
        .filter(|list| {
            list.entries
                .iter()
                .any(|entry| entry.string.as_deref() == Some(text))
        })
        .collect::<Vec<_>>();
    if lists.len() != 1 {
        return Err(std::io::Error::other(format!(
            "expected one native string list containing selected text, found {}",
            lists.len()
        ))
        .into());
    }
    Ok(lists.remove(0))
}

#[test]
fn native_single_reads_use_checked_index_and_name_selectors() -> TestResult {
    let package = Package::open(fixture_path())?;
    let b2 = CellPosition::from_a1("B2")?;
    let b3 = CellPosition::from_a1("B3")?;
    let g22 = CellPosition::from_a1("G22")?;

    let text_by_index = package.table_cell(0usize, 0usize, b2)?;
    let text_by_name = package.table_cell(
        SheetSelector::name("Sheet 1"),
        TableSelector::name("Table 1"),
        b2,
    )?;
    assert_eq!(text_by_name, text_by_index);
    assert_eq!(text_by_index.position(), b2);
    assert!(matches!(
        text_by_index.storage(),
        Storage::Stored(Value::Text(text)) if text == FIXTURE_MARKER
    ));
    assert_eq!(
        text_by_index.storage().value().map(Value::cell_type),
        Some(litchi_numbers::cell::Type::Text)
    );

    let number = package.table_cell("Sheet 1", "Table 1", b3)?;
    assert_eq!(number.position(), b3);
    assert!(matches!(
        number.storage(),
        Storage::Stored(Value::Number(value)) if value.get() == 42.0
    ));

    let missing = package.table_cell("Sheet 1", "Table 1", g22)?;
    assert_eq!(missing.position(), g22);
    assert!(missing.storage().is_missing());
    assert_eq!(missing.storage().value(), None);
    assert!(matches!(missing.storage(), Storage::Missing));
    Ok(())
}

#[test]
fn native_dense_range_is_row_major_and_presence_preserving() -> TestResult {
    let package = Package::open(fixture_path())?;
    let states = package.table_cells(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellRange::from_a1("B2:B3")?,
    )?;

    assert_eq!(states.len(), 2);
    assert_eq!(states[0].position(), CellPosition::from_a1("B2")?);
    assert_eq!(states[1].position(), CellPosition::from_a1("B3")?);
    assert!(matches!(
        states[0].storage(),
        Storage::Stored(Value::Text(text)) if text == FIXTURE_MARKER
    ));
    assert!(matches!(
        states[1].storage(),
        Storage::Stored(Value::Number(value)) if value.get() == 42.0
    ));

    let missing = package.table_cells(0usize, 0usize, CellRange::from_a1("F21:G22")?)?;
    assert_eq!(missing.len(), 4);
    assert_eq!(
        missing.iter().map(State::position).collect::<Vec<_>>(),
        ["F21", "G21", "F22", "G22"]
            .map(CellPosition::from_a1)
            .into_iter()
            .collect::<Result<Vec<_>, _>>()?
    );
    assert!(missing.iter().all(|state| state.storage().is_missing()));
    Ok(())
}

#[test]
fn missing_selectors_and_bounds_return_typed_errors() -> TestResult {
    let package = Package::open(fixture_path())?;
    let b2 = CellPosition::from_a1("B2")?;
    assert!(matches!(
        package.table_cell("missing sheet", 0usize, b2),
        Err(Error::SheetNotFound)
    ));
    assert!(matches!(
        package.table_cell(0usize, "missing table", b2),
        Err(Error::TableNotFound)
    ));

    let dimensions = Dimensions::new(22, 7);
    let outside = CellPosition::new(22, 0);
    assert!(matches!(
        package.table_cell(0usize, 0usize, outside),
        Err(Error::OutOfBounds {
            position,
            dimensions: actual,
        }) if position == outside && actual == dimensions
    ));

    let end = CellPosition::new(22, 8);
    let outside_range = CellRange::new(CellPosition::new(21, 6), end)?;
    assert!(matches!(
        package.table_cells(0usize, 0usize, outside_range),
        Err(Error::OutOfBounds {
            position,
            dimensions: actual,
        }) if position == end && actual == dimensions
    ));
    Ok(())
}

#[test]
fn public_read_values_are_send_sync_and_debug_redacted() -> TestResult {
    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<Storage>();
    assert_send_sync_debug::<State>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<LimitKind>();
    assert_send_sync_debug::<Path>();

    let package = Package::open(fixture_path())?;
    let state = package.table_cell("Sheet 1", "Table 1", CellPosition::from_a1("B2")?)?;
    for rendered in [format!("{state:?}"), format!("{:?}", state.storage())] {
        assert!(rendered.contains("Text"));
        assert!(!rendered.contains(FIXTURE_MARKER));
    }

    let error = package
        .table_cell(
            "private missing sheet name",
            0usize,
            CellPosition::new(0, 0),
        )
        .expect_err("missing selector must fail");
    for rendered in [format!("{error:?}"), error.to_string()] {
        assert!(!rendered.contains("private missing sheet name"));
        assert!(!rendered.contains(FIXTURE_MARKER));
    }
    Ok(())
}

#[test]
fn native_number_change_is_exact_reversible_and_source_bound() -> TestResult {
    let package = Package::open(fixture_path())?;
    let source = package.exact_bytes();
    let b3 = CellPosition::from_a1("B3")?;

    let commit = package
        .edit_table_cells(0usize, 0usize)?
        .set(b3, Input::number(43.0)?)?
        .commit()?;
    let diagnostics = commit.diagnostics();
    assert!(diagnostics.changed());
    assert_eq!(diagnostics.requested_cells(), 1);
    assert_eq!(diagnostics.changed_cells(), 1);
    assert_eq!(diagnostics.touched_components(), 1);
    assert_eq!(diagnostics.refreshed_formula_caches(), 0);
    assert_eq!(diagnostics.deleted_previews(), 3);
    assert!(diagnostics.full_reparse_performed());

    let patch = commit.patch();
    assert_eq!(patch.path(), Path::Table { sheet: 0, table: 0 });
    assert_eq!(patch.len(), 1);
    assert!(!patch.is_noop());
    assert_ne!(patch.source_fingerprint(), patch.target_fingerprint());
    assert!(matches!(
        commit.package().table_cell(0usize, 0usize, b3)?.storage(),
        Storage::Stored(Value::Number(value)) if value.get() == 43.0
    ));

    let target = commit.package().exact_bytes();
    assert_ne!(target, source);
    assert_exact_scalar_locality(&source, &target, &["Index/Tables/Tile.iwa"])?;
    assert_eq!(
        package.apply_table_cells(patch)?.package().exact_bytes(),
        target
    );

    let reopened = Package::from_bytes(&target)?;
    assert!(matches!(
        reopened.apply_table_cells(patch),
        Err(Error::PatchConflict)
    ));
    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), *patch);
    let restored = reopened.apply_table_cells(&inverse)?;
    assert_eq!(restored.package().exact_bytes(), source);
    assert!(matches!(
        restored
            .package()
            .table_cell("Sheet 1", "Table 1", b3)?
            .storage(),
        Storage::Stored(Value::Number(value)) if value.get() == 42.0
    ));
    Ok(())
}

#[test]
fn changed_patch_preserves_exact_source_preview_subset() -> TestResult {
    let native = std::fs::read(fixture_path())?;
    let masked = Catalog::from_bytes(&native)?.reassemble_with_deletions_to_bytes(
        &[],
        &["preview.jpg", "preview-web.jpg"],
        Limits::default(),
    )?;
    let masked_catalog = Catalog::from_bytes(&masked)?;
    assert_eq!(
        CANONICAL_PREVIEWS
            .iter()
            .copied()
            .filter(|name| masked_catalog.iter().any(|entry| entry.name() == *name))
            .collect::<Vec<_>>(),
        ["preview-micro.jpg"]
    );

    let package = Package::from_bytes(&masked)?;
    let source = package.exact_bytes();
    let commit = package
        .edit_table_cells(0usize, 0usize)?
        .set_a1("B3", Input::number(43.0)?)?
        .commit()?;
    assert_eq!(commit.diagnostics().deleted_previews(), 1);
    let target = commit.package().exact_bytes();
    let target_catalog = Catalog::from_bytes(&target)?;
    assert!(
        CANONICAL_PREVIEWS
            .iter()
            .all(|preview| target_catalog.iter().all(|entry| entry.name() != *preview))
    );
    assert_eq!(
        package
            .apply_table_cells(commit.patch())?
            .package()
            .exact_bytes(),
        target
    );
    assert_eq!(
        commit
            .package()
            .apply_table_cells(&commit.patch().inverse())?
            .package()
            .exact_bytes(),
        source
    );
    Ok(())
}

#[test]
fn native_mixed_scalar_batch_roundtrips_from_one_final_state() -> TestResult {
    let package = Package::open(fixture_path())?;
    let source = package.exact_bytes();
    let commit = package
        .edit_table_cells("Sheet 1", "Table 1")?
        .set_a1("B3", Input::number(43.5)?)?
        .set_a1("C3", Input::boolean(true))?
        .set_a1("D3", Input::date(123_456.25)?)?
        .set_a1("E3", Input::duration(90.5)?)?
        .commit()?;

    assert_eq!(commit.patch().len(), 4);
    assert_eq!(commit.diagnostics().requested_cells(), 4);
    assert_eq!(commit.diagnostics().changed_cells(), 4);
    assert_eq!(commit.diagnostics().touched_components(), 2);
    let states = commit
        .package()
        .table_cells(0usize, 0usize, CellRange::from_a1("B3:E3")?)?;
    assert!(matches!(
        states[0].storage(),
        Storage::Stored(Value::Number(value)) if value.get() == 43.5
    ));
    assert!(matches!(
        states[1].storage(),
        Storage::Stored(Value::Boolean(true))
    ));
    assert!(matches!(
        states[2].storage(),
        Storage::Stored(Value::Date(value)) if value.get() == 123_456.25
    ));
    assert!(matches!(
        states[3].storage(),
        Storage::Stored(Value::Duration(value)) if value.get() == 90.5
    ));

    let restored = commit
        .package()
        .apply_table_cells(&commit.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), source);
    Ok(())
}

#[test]
fn native_same_tile_mixed_equal_and_changed_counts_final_delta() -> TestResult {
    let package = Package::open(fixture_path())?;
    let source = package.exact_bytes();
    let commit = package
        .edit_table_cells(0usize, 0usize)?
        .set_a1("B3", Input::number(42.0)?)?
        .set_a1("C3", Input::boolean(true))?
        .commit()?;

    assert_eq!(commit.patch().len(), 2);
    assert_eq!(commit.diagnostics().requested_cells(), 2);
    assert_eq!(commit.diagnostics().changed_cells(), 1);
    assert_eq!(commit.diagnostics().touched_components(), 2);
    assert!(matches!(
        commit
            .package()
            .table_cell(0usize, 0usize, CellPosition::from_a1("B3")?)?
            .storage(),
        Storage::Stored(Value::Number(value)) if value.get() == 42.0
    ));
    assert!(matches!(
        commit
            .package()
            .table_cell(0usize, 0usize, CellPosition::from_a1("C3")?)?
            .storage(),
        Storage::Stored(Value::Boolean(true))
    ));
    let target = commit.package().exact_bytes();
    assert_exact_scalar_locality(
        &source,
        &target,
        &[
            "Index/Tables/Tile.iwa",
            "Index/Tables/HeaderStorageBucket-904855-2.iwa",
        ],
    )?;
    assert_eq!(
        commit
            .package()
            .apply_table_cells(&commit.patch().inverse())?
            .package()
            .exact_bytes(),
        source
    );
    Ok(())
}

#[test]
fn native_clear_preserves_stored_empty_as_distinct_from_missing() -> TestResult {
    let package = Package::open(fixture_path())?;
    let source = package.exact_bytes();
    let b3 = CellPosition::from_a1("B3")?;
    let g22 = CellPosition::from_a1("G22")?;
    let commit = package
        .edit_table_cells(0usize, 0usize)?
        .clear(b3)?
        .commit()?;

    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().requested_cells(), 1);
    assert_eq!(commit.diagnostics().changed_cells(), 1);
    assert!(matches!(
        commit.package().table_cell(0usize, 0usize, b3)?.storage(),
        Storage::Stored(Value::Empty)
    ));
    assert!(matches!(
        commit.package().table_cell(0usize, 0usize, g22)?.storage(),
        Storage::Missing
    ));

    let restored = commit
        .package()
        .apply_table_cells(&commit.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), source);
    assert!(matches!(
        restored.package().table_cell(0usize, 0usize, b3)?.storage(),
        Storage::Stored(Value::Number(value)) if value.get() == 42.0
    ));

    let cleared = commit.into_package();
    let cleared_source = cleared.exact_bytes();
    let noop = cleared
        .edit_table_cells(0usize, 0usize)?
        .clear(b3)?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().requested_cells(), 1);
    assert_eq!(noop.diagnostics().changed_cells(), 0);
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert_eq!(noop.diagnostics().deleted_previews(), 0);
    assert_eq!(noop.package().exact_bytes(), cleared_source);
    Ok(())
}

#[test]
fn native_text_replacement_and_clear_update_exact_string_ownership() -> TestResult {
    let package = Package::open(fixture_path())?;
    let source = package.exact_bytes();
    let source_strings = string_list_with_text(&source, FIXTURE_MARKER)?;
    let original = source_strings
        .entries
        .iter()
        .find(|entry| entry.string.as_deref() == Some(FIXTURE_MARKER))
        .ok_or_else(|| std::io::Error::other("fixture marker string entry is missing"))?;
    let original_key = original.key;
    assert_eq!(original.refcount, 1);

    let replacement = "Focused native batch text";
    let changed = package
        .edit_table_cells(0usize, 0usize)?
        .set_a1("B2", Input::text(replacement)?)?
        .commit()?;
    assert!(matches!(
        changed
            .package()
            .table_cell(0usize, 0usize, CellPosition::from_a1("B2")?)?
            .storage(),
        Storage::Stored(Value::Text(value)) if value == replacement
    ));
    let changed_bytes = changed.package().exact_bytes();
    let changed_strings = string_list_with_text(&changed_bytes, replacement)?;
    assert!(
        changed_strings
            .entries
            .iter()
            .all(|entry| entry.key != original_key)
    );
    let inserted = changed_strings
        .entries
        .iter()
        .find(|entry| entry.string.as_deref() == Some(replacement))
        .ok_or_else(|| std::io::Error::other("replacement string entry is missing"))?;
    assert_eq!(inserted.refcount, 1);
    assert_ne!(inserted.key, original_key);
    assert_eq!(
        changed_strings.next_list_id,
        source_strings.next_list_id + 1
    );

    let cleared = changed
        .package()
        .edit_table_cells(0usize, 0usize)?
        .clear_a1("B2")?
        .commit()?;
    assert!(matches!(
        cleared
            .package()
            .table_cell(0usize, 0usize, CellPosition::from_a1("B2")?)?
            .storage(),
        Storage::Stored(Value::Empty)
    ));
    assert!(
        string_lists(&cleared.package().exact_bytes())?
            .iter()
            .flat_map(|list| list.entries.iter())
            .all(|entry| entry.string.as_deref() != Some(replacement))
    );

    let restored = changed
        .package()
        .apply_table_cells(&changed.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), source);
    Ok(())
}

#[test]
fn native_exact_noop_bypasses_changed_path_and_preserves_source() -> TestResult {
    let package = Package::open(fixture_path())?;
    let source = package.exact_bytes();
    let commit = package
        .edit_table_cells("Sheet 1", "Table 1")?
        .set_a1("B3", Input::number(42.0)?)?
        .clear_a1("G22")?
        .commit()?;

    assert!(commit.patch().is_noop());
    assert_eq!(commit.patch().len(), 2);
    assert_eq!(
        commit.patch().source_fingerprint(),
        commit.patch().target_fingerprint()
    );
    assert_eq!(commit.package().exact_bytes(), source);
    let diagnostics = commit.diagnostics();
    assert!(!diagnostics.changed());
    assert_eq!(diagnostics.requested_cells(), 2);
    assert_eq!(diagnostics.changed_cells(), 0);
    assert_eq!(diagnostics.touched_components(), 0);
    assert_eq!(diagnostics.refreshed_formula_caches(), 0);
    assert_eq!(diagnostics.deleted_previews(), 0);
    assert!(!diagnostics.full_reparse_performed());

    let applied = package.apply_table_cells(commit.patch())?;
    assert!(applied.patch().is_noop());
    assert_eq!(applied.package().exact_bytes(), source);
    assert_eq!(
        applied
            .package()
            .apply_table_cells(&applied.patch().inverse())?
            .package()
            .exact_bytes(),
        source
    );
    Ok(())
}

#[test]
fn duplicate_bounds_and_addresses_fail_before_publication() -> TestResult {
    let package = Package::open(fixture_path())?;
    let source = package.exact_bytes();
    let b3 = CellPosition::from_a1("B3")?;
    let duplicate = package
        .edit_table_cells(0usize, 0usize)?
        .set(b3, Input::number(43.0)?)?
        .clear(b3)?
        .commit()
        .expect_err("duplicate coordinates must fail");
    assert!(matches!(
        duplicate,
        Error::DuplicatePosition { position } if position == b3
    ));

    let outside = CellPosition::new(22, 0);
    let bounds = package
        .edit_table_cells(0usize, 0usize)?
        .set(outside, Input::number(43.0)?)
        .expect_err("out-of-bounds coordinate must fail while staging");
    assert!(matches!(
        bounds,
        Error::OutOfBounds {
            position,
            dimensions,
        } if position == outside && dimensions == Dimensions::new(22, 7)
    ));
    assert!(matches!(
        package
            .edit_table_cells(0usize, 0usize)?
            .set_a1("private invalid address", Input::boolean(true)),
        Err(Error::InvalidAddress)
    ));
    assert_eq!(package.exact_bytes(), source);
    Ok(())
}

#[test]
fn native_header_name_index_allows_noop_and_refuses_change() -> TestResult {
    let package = Package::open(fixture_path())?;
    let source = package.exact_bytes();
    let noop = package
        .edit_table_cells(0usize, 0usize)?
        .clear_a1("A1")?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().exact_bytes(), source);

    let error = package
        .edit_table_cells(0usize, 0usize)?
        .set_a1("A1", Input::text("Input batch")?)?
        .commit()
        .expect_err("header-name indexed edits must refuse atomically");
    assert!(matches!(
        error,
        Error::UnsupportedDependency {
            path: Path::Table { sheet: 0, table: 0 },
            kind: DependencyKind::HeaderNameIndex,
        }
    ));
    assert_eq!(package.exact_bytes(), source);
    Ok(())
}

#[test]
fn locked_table_allows_exact_noop_and_refuses_changed_cell() -> TestResult {
    let package = Package::open(fixture_path())?;
    let mut lock = package.edit_table_lock(0usize, 0usize)?;
    lock.lock();
    let locked = lock.commit()?.into_package();
    let source = locked.exact_bytes();

    let noop = locked
        .edit_table_cells(0usize, 0usize)?
        .set_a1("B3", Input::number(42.0)?)?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().exact_bytes(), source);

    let error = locked
        .edit_table_cells(0usize, 0usize)?
        .set_a1("B3", Input::number(43.0)?)?
        .commit()
        .expect_err("a locked table must refuse a changed cell batch");
    assert!(matches!(
        error,
        Error::TableLocked {
            path: Path::Table { sheet: 0, table: 0 }
        }
    ));
    assert_eq!(locked.exact_bytes(), source);
    Ok(())
}

#[test]
fn public_mutation_values_are_send_sync_and_debug_redacted() -> TestResult {
    assert_send_sync_debug::<Input>();
    assert_send_sync_debug::<Change>();
    assert_send_sync_debug::<Edit<'static>>();
    assert_send_sync_debug::<Patch>();
    assert_send_sync_debug::<Commit>();
    assert_send_sync_debug::<Diagnostics>();
    assert_send_sync_debug::<DependencyKind>();

    let secret = "private authored mutation text";
    let input = Input::text(secret)?;
    let change = Change::set_a1("B3", input.clone())?;
    for rendered in [format!("{input:?}"), format!("{change:?}")] {
        assert!(!rendered.contains(secret));
        assert!(!rendered.contains(FIXTURE_MARKER));
    }

    let package = Package::open(fixture_path())?;
    let edit = package
        .edit_table_cells("Sheet 1", "Table 1")?
        .change(change)?;
    let rendered = format!("{edit:?}");
    assert!(!rendered.contains(secret));
    assert!(!rendered.contains(FIXTURE_MARKER));

    let commit = package
        .edit_table_cells(0usize, 0usize)?
        .set_a1("B3", Input::number(42.0)?)?
        .commit()?;
    for rendered in [format!("{:?}", commit.patch()), format!("{commit:?}")] {
        assert!(!rendered.contains(secret));
        assert!(!rendered.contains(FIXTURE_MARKER));
        assert!(!rendered.contains("Index/Tables.iwa"));
        assert!(!rendered.contains("preview.jpg"));
    }
    Ok(())
}
