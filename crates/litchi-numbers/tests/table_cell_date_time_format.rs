//! Exact-source integration coverage for Numbers table-cell Date & Time
//! display formats.
//!
//! The tests exercise the selector-first owner over a small rooted package.
//! Native BNC metadata and format-list identifiers stay in the test fixture;
//! callers of the focused API see only [`DateTime`] and checked selectors.

use std::{fmt::Debug, io};

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::{
    varint::{decode_varint_from_bytes, encode_varint},
    wire::{RawWireFields, append_varint_field},
};
use litchi_iwa_protos::tsk;
use litchi_numbers::cell::data_format::{
    DateTime,
    date_time::transaction::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path},
};
use litchi_numbers::{CellPosition, Package, SheetSelector, TableSelector};
use litchi_numbers_wire::{BncCell, CellDataFormatKind, StoredValue};
use prost::Message as _;

#[path = "support/table_cell_data_format_fixture.rs"]
mod fixture;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

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

fn selected_position() -> CellPosition {
    CellPosition::new(fixture::FIRST_CELL.0 as u32, fixture::FIRST_CELL.1 as u32)
}

fn source() -> TestResult<Vec<u8>> {
    fixture::synthetic_package_for(
        fixture::FormatFamily::DateTime,
        fixture::FormatSharing::Shared,
    )
}

fn date_time(pattern: &str) -> TestResult<DateTime> {
    Ok(DateTime::new(pattern)?)
}

fn assert_exact_locality(source: &[u8], target: &[u8]) -> TestResult {
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    let mut changed = Vec::new();
    for entry in before.iter() {
        let candidate = after
            .iter()
            .find(|other| other.name() == entry.name())
            .ok_or_else(|| io::Error::other("candidate removed a source member"))?;
        if entry.data() != candidate.data() {
            changed.push(entry.name().to_owned());
        } else {
            assert_eq!(
                entry.raw_record().local_record(),
                candidate.raw_record().local_record()
            );
        }
        if entry.name().contains("Metadata")
            || fixture::PREVIEW_MEMBERS.contains(&entry.name())
            || entry.name() == fixture::UNRELATED_MEMBER
            || entry.name() == fixture::SENTINEL_MEMBER
        {
            assert_eq!(entry.data(), candidate.data(), "unrelated member changed");
        }
    }
    assert_eq!(changed, [fixture::TABLES_MEMBER.to_owned()]);
    assert_eq!(before.len(), after.len());
    Ok(())
}

fn assert_non_format_cell_state(source: &[u8], target: &[u8]) -> TestResult {
    let before = fixture::tile_cells(source)?;
    let after = fixture::tile_cells(target)?;
    assert_eq!(before.len(), after.len());
    for (before, after) in before.iter().zip(after.iter()) {
        let before_cell = BncCell::parse(before)?;
        let after_cell = BncCell::parse(after)?;
        assert_eq!(before_cell.stored_value(), after_cell.stored_value());
        assert_eq!(before_cell.cached_scalar()?, after_cell.cached_scalar()?);
        let mut before_metadata_free = before_cell;
        let mut after_metadata_free = after_cell;
        before_metadata_free.clear_explicit_format();
        after_metadata_free.clear_explicit_format();
        assert_eq!(before_metadata_free.encode(), after_metadata_free.encode());
    }
    Ok(())
}

fn first_cell(source: &[u8]) -> TestResult<BncCell> {
    let cells = fixture::tile_cells(source)?;
    let bytes = cells
        .first()
        .ok_or_else(|| io::Error::other("Date & Time fixture first cell is missing"))?;
    Ok(BncCell::parse(bytes)?)
}

fn assert_owner_rejects(source: &[u8]) -> TestResult {
    let package = match Package::from_bytes(source) {
        Ok(package) => package,
        Err(_) => return Ok(()),
    };
    let before = package.exact_bytes();
    assert!(
        package
            .table_cell_date_time_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert!(
        package
            .edit_table_cell_date_time_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

fn raw_field_record(source: &[u8], number: u32) -> TestResult<Vec<u8>> {
    let mut fields = RawWireFields::new(source);
    while let Some(field) = fields.next()? {
        if field.number() == number {
            return Ok(field.raw().to_vec());
        }
    }
    Err(io::Error::other(format!("wire field {number} is missing")).into())
}

fn raw_format_payload_by_key(source: &[u8], key: u32) -> TestResult<Vec<u8>> {
    let list_payload = fixture::format_list_payload(source)?;
    let mut list_fields = RawWireFields::new(&list_payload);
    while let Some(entry_field) = list_fields.next()? {
        if entry_field.number() != 3 {
            continue;
        }
        let mut entry_fields = RawWireFields::new(entry_field.payload());
        let mut matching_key = false;
        while let Some(field) = entry_fields.next()? {
            if field.number() == 1 {
                let (value, _) = decode_varint_from_bytes(field.payload())
                    .map_err(|error| io::Error::other(error.to_string()))?;
                matching_key = value == u64::from(key);
            } else if matching_key && field.number() == 6 {
                return Ok(field.payload().to_vec());
            }
        }
    }
    Err(io::Error::other(format!("format entry key {key} is missing")).into())
}

#[test]
fn date_time_transaction_types_are_strictly_typed_and_redacted() -> TestResult {
    fn assert_send_sync_debug<T: Send + Sync + Debug>() {}

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<DateTime>();
    assert_send_sync_debug::<Edit<'static>>();
    assert_send_sync_debug::<Commit>();
    assert_send_sync_debug::<Patch>();
    assert_send_sync_debug::<Diagnostics>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<LimitKind>();
    assert_send_sync_debug::<Path>();

    let package = Package::from_bytes(&source()?)?;
    let edit = package.edit_table_cell_date_time_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        selected_position(),
    )?;
    let rendered = format!("{edit:?}");
    assert!(!rendered.contains("Index/"));
    assert!(!rendered.contains("data-format-table-id"));
    let commit = edit.set(date_time("MM/dd/yyyy")?).commit()?;
    for rendered in [
        format!("{commit:?}"),
        format!("{:?}", commit.patch()),
        format!("{:?}", commit.diagnostics()),
    ] {
        assert!(!rendered.contains("Index/"));
        assert!(!rendered.contains("data-format-table-id"));
    }
    Ok(())
}

#[test]
fn date_time_read_uses_selectors_and_preserves_pattern_exactly() -> TestResult {
    let bytes = source()?;
    let package = Package::from_bytes(&bytes)?;
    let expected = date_time("yyyy-MM-dd H:mm:ss")?;
    assert_eq!(
        package.table_cell_date_time_format(0usize, 0usize, selected_position())?,
        Some(expected.clone())
    );
    assert_eq!(
        package.table_cell_date_time_format(
            SheetSelector::name("Data Format Sheet"),
            TableSelector::name("Data Formats"),
            selected_position(),
        )?,
        Some(expected)
    );
    let cell = first_cell(&bytes)?;
    assert_eq!(
        cell.explicit_format_flags(),
        litchi_numbers_wire::EXPLICIT_DATE_TIME_FORMAT
    );
    assert_eq!(
        cell.cell_format_kind(),
        Some(litchi_numbers_wire::DATE_TIME_CELL_FORMAT_KIND)
    );
    assert!(cell.format_identifier().is_some());
    assert_eq!(cell.stored_value(), StoredValue::Date);
    Ok(())
}

#[test]
fn date_time_changed_commit_is_local_reversible_and_value_preserving() -> TestResult {
    let bytes = source()?;
    let package = Package::from_bytes(&bytes)?;
    let original = package
        .table_cell_date_time_format(0usize, 0usize, selected_position())?
        .ok_or_else(|| io::Error::other("Date & Time fixture format is missing"))?;

    let no_op = package
        .edit_table_cell_date_time_format(0usize, 0usize, selected_position())?
        .set(original.clone())
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().exact_bytes(), bytes);
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());

    let replacement = date_time("MM/dd/yyyy HH:mm:ss")?;
    let changed = package
        .edit_table_cell_date_time_format(0usize, 0usize, selected_position())?
        .set(replacement.clone())
        .commit()?;
    let target = changed.package().exact_bytes();
    assert!(!changed.patch().is_noop());
    assert_eq!(changed.patch().before(), Some(&original));
    assert_eq!(changed.patch().after(), Some(&replacement));
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().touched_components(), 1);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_eq!(
        changed
            .package()
            .table_cell_date_time_format(0usize, 0usize, selected_position())?,
        Some(replacement.clone())
    );
    assert_exact_locality(&bytes, &target)?;
    assert_non_format_cell_state(&bytes, &target)?;

    let payload = fixture::format_payload_by_key(
        &target,
        first_cell(&target)?
            .format_identifier()
            .ok_or_else(|| io::Error::other("Date & Time format identifier disappeared"))?,
    )?;
    let native = tsk::FormatStructArchive::decode(payload.as_slice())?;
    assert_eq!(
        native.format_type,
        Some(fixture::NATIVE_DATE_TIME_FORMAT_TYPE)
    );
    assert_eq!(
        native.date_time_format.as_deref(),
        Some(replacement.pattern())
    );

    let applied = package.apply_table_cell_date_time_format(changed.patch())?;
    assert_eq!(applied.package().exact_bytes(), target);
    let inverse = changed.patch().inverse();
    assert_eq!(inverse.inverse(), *changed.patch());
    let restored = Package::from_bytes(&target)?.apply_table_cell_date_time_format(&inverse)?;
    assert_eq!(restored.package().exact_bytes(), bytes);
    assert_eq!(
        restored
            .package()
            .table_cell_date_time_format(0usize, 0usize, selected_position())?,
        Some(original)
    );
    Ok(())
}

#[test]
fn date_time_clear_and_reset_remove_only_explicit_metadata() -> TestResult {
    let bytes = source()?;
    let package = Package::from_bytes(&bytes)?;
    let cleared = package
        .edit_table_cell_date_time_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .table_cell_date_time_format(0usize, 0usize, selected_position())?,
        None
    );
    let cell = first_cell(&cleared.package().exact_bytes())?;
    assert_eq!(cell.explicit_format_flags(), 0);
    assert_eq!(cell.cell_format_kind(), None);
    assert_eq!(cell.format_identifier(), None);
    assert_eq!(cell.stored_value(), StoredValue::Date);
    assert_non_format_cell_state(&bytes, &cleared.package().exact_bytes())?;

    let reset = cleared
        .package()
        .edit_table_cell_date_time_format(0usize, 0usize, selected_position())?
        .reset()
        .commit()?;
    assert!(reset.patch().is_noop());
    Ok(())
}

#[test]
fn date_time_wrong_family_and_automatic_marker_fail_closed() -> TestResult {
    let number = Package::from_bytes(&fixture::synthetic_package_for(
        fixture::FormatFamily::Number,
        fixture::FormatSharing::Shared,
    )?)?;
    assert!(matches!(
        number.table_cell_date_time_format(0usize, 0usize, selected_position()),
        Err(Error::WrongFormatFamily { .. })
    ));

    let secondary = fixture::rewrite_tile_cells(&source()?, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("Date & Time fixture first cell is missing"))?;
        let mut encoded = BncCell::parse(first)?.encode();
        let flags = u32::from_le_bytes(
            encoded[8..12]
                .try_into()
                .map_err(|_| io::Error::other("Date & Time flags are truncated"))?,
        );
        encoded[8..12].copy_from_slice(&(flags | 0x0000_2000).to_le_bytes());
        // The generic decimal identifier follows the kind field and precedes
        // the Date & Time identifier in the fixed BNC field layout.
        encoded.splice(16..16, 7_u32.to_le_bytes());
        *first = encoded;
        Ok(())
    })?;
    assert_owner_rejects(&secondary)?;

    let marked_automatic = fixture::rewrite_tile_cells(&source()?, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("Date & Time fixture first cell is missing"))?;
        if first.len() < 8 {
            return Err(io::Error::other("Date & Time fixture cell is truncated").into());
        }
        first[6..8].copy_from_slice(&0_u16.to_le_bytes());
        Ok(())
    })?;
    let package = Package::from_bytes(&marked_automatic)?;
    let before = package.exact_bytes();
    assert!(
        package
            .table_cell_date_time_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert!(
        package
            .edit_table_cell_date_time_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn date_time_accepts_empty_and_evidenced_type_nine_but_rejects_plain_number() -> TestResult {
    let empty = fixture::rewrite_tile_cells(&source()?, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("Date & Time fixture first cell is missing"))?;
        *first = BncCell::minimal().encode();
        Ok(())
    })?;
    // The selected cell no longer references key 1, so keep the fixture's
    // list census valid for the sibling cell that still does.
    let empty = fixture::rewrite_format_list_payload_for_test(&empty, |list| {
        let entry = list
            .entries
            .iter_mut()
            .find(|entry| entry.key == fixture::FIRST_FORMAT_KEY)
            .ok_or_else(|| io::Error::other("Date & Time format entry is missing"))?;
        entry.refcount = 1;
        Ok(())
    })?;
    let package = Package::from_bytes(&empty)?;
    assert_eq!(
        package.table_cell_date_time_format(0usize, 0usize, selected_position())?,
        None
    );
    let attached = package
        .edit_table_cell_date_time_format(0usize, 0usize, selected_position())?
        .set(date_time("yyyy-MM-dd")?)
        .commit()?;
    assert_eq!(
        attached
            .package()
            .table_cell_date_time_format(0usize, 0usize, selected_position())?,
        Some(date_time("yyyy-MM-dd")?)
    );
    assert_eq!(
        first_cell(&attached.package().exact_bytes())?.stored_value(),
        StoredValue::Empty
    );

    let type_nine = fixture::rewrite_tile_cells(&source()?, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("Date & Time fixture first cell is missing"))?;
        let mut cell = BncCell::minimal();
        cell.set_number(42.5)?;
        cell.set_data_format_identifier(
            fixture::FIRST_FORMAT_KEY,
            CellDataFormatKind::DateTime,
            None,
        )?;
        *first = cell.encode();
        Ok(())
    })?;
    let type_nine_cell = first_cell(&type_nine)?;
    assert_eq!(type_nine_cell.encode()[1], 9);
    assert_eq!(
        type_nine_cell.explicit_format_flags(),
        litchi_numbers_wire::EXPLICIT_DATE_TIME_FORMAT
    );
    assert_eq!(
        type_nine_cell.cell_format_kind(),
        Some(litchi_numbers_wire::DATE_TIME_CELL_FORMAT_KIND)
    );
    assert_eq!(
        type_nine_cell.format_identifier(),
        Some(fixture::FIRST_FORMAT_KEY)
    );
    assert_eq!(
        Package::from_bytes(&type_nine)?.table_cell_date_time_format(
            0usize,
            0usize,
            selected_position(),
        )?,
        Some(date_time("yyyy-MM-dd H:mm:ss")?)
    );

    let plain_number = fixture::rewrite_tile_cells(&source()?, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("Date & Time fixture first cell is missing"))?;
        // Preserve the valid Date & Time list reference and every BNC field,
        // changing only the native cell-type discriminator from type 9 to
        // ordinary numeric type 2. This isolates value-shape admission from
        // refcount/list-census validation.
        let mut type_nine = BncCell::minimal();
        type_nine.set_number(42.5)?;
        type_nine.set_data_format_identifier(
            fixture::FIRST_FORMAT_KEY,
            CellDataFormatKind::DateTime,
            None,
        )?;
        let mut encoded = type_nine.encode();
        assert_eq!(encoded.get(1), Some(&9));
        encoded[1] = 2;
        *first = encoded;
        Ok(())
    })?;
    // The list graph is still valid: both cells retain the shared key and its
    // census remains two. The owner must reject only the unsupported type-2
    // value shape, not a malformed reference graph.
    assert_eq!(fixture::format_keys(&plain_number)?, vec![Some(1), Some(1)]);
    assert_eq!(fixture::format_entry_facts(&plain_number)?, vec![(1, 2)]);
    let plain_package = Package::from_bytes(&plain_number)?;
    let plain_before = plain_package.exact_bytes();
    assert!(matches!(
        plain_package.table_cell_date_time_format(0usize, 0usize, selected_position()),
        Err(Error::WrongFormatFamily { .. })
    ));
    assert_eq!(plain_package.exact_bytes(), plain_before);

    let unformatted_number = fixture::rewrite_tile_cells(&source()?, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("Date & Time fixture first cell is missing"))?;
        let mut cell = BncCell::minimal();
        cell.set_plain_number(42.5)?;
        *first = cell.encode();
        Ok(())
    })?;
    let unformatted_number =
        fixture::rewrite_format_list_payload_for_test(&unformatted_number, |list| {
            let entry = list
                .entries
                .iter_mut()
                .find(|entry| entry.key == fixture::FIRST_FORMAT_KEY)
                .ok_or_else(|| io::Error::other("Date & Time format entry is missing"))?;
            entry.refcount = 1;
            Ok(())
        })?;
    assert_eq!(
        fixture::format_keys(&unformatted_number)?,
        vec![None, Some(1)]
    );
    assert_eq!(
        fixture::format_entry_facts(&unformatted_number)?,
        vec![(1, 1)]
    );
    assert_owner_rejects(&unformatted_number)?;
    Ok(())
}

#[test]
fn date_time_pattern_validation_is_bounded_and_lossless() -> TestResult {
    assert!(DateTime::new("   ").is_err());
    assert!(DateTime::new("yyyy\0MM").is_err());
    assert!(
        DateTime::new(
            &"x".repeat(litchi_numbers::cell::data_format::date_time::MAX_PATTERN_BYTES + 1)
        )
        .is_err()
    );
    let pattern = date_time("EEEE, MMMM d, y")?;
    assert_eq!(pattern.pattern(), "EEEE, MMMM d, y");
    Ok(())
}

#[test]
fn date_time_shared_format_key_uses_cow_and_keeps_refcounts_consistent() -> TestResult {
    let bytes = source()?;
    assert_eq!(fixture::format_keys(&bytes)?, vec![Some(1), Some(1)]);
    assert_eq!(fixture::format_entry_facts(&bytes)?, vec![(1, 2)]);
    assert_eq!(fixture::format_next_list_id(&bytes)?, 32);

    let replacement = date_time("MM/dd/yyyy HH:mm:ss")?;
    let first = Package::from_bytes(&bytes)?
        .edit_table_cell_date_time_format(0usize, 0usize, selected_position())?
        .set(replacement.clone())
        .commit()?;
    let first_bytes = first.package().exact_bytes();
    let first_keys = fixture::format_keys(&first_bytes)?;
    assert_ne!(first_keys[0], first_keys[1]);
    assert_eq!(first_keys[1], Some(1));
    let new_key = first_keys[0].ok_or_else(|| io::Error::other("COW key is missing"))?;
    assert_eq!(
        fixture::format_entry_facts(&first_bytes)?,
        vec![(1, 1), (new_key, 1)]
    );
    assert_eq!(fixture::format_next_list_id(&first_bytes)?, 32);
    assert_non_format_cell_state(&bytes, &first_bytes)?;

    let both = first
        .package()
        .edit_table_cell_date_time_format(0usize, 0usize, CellPosition::new(0, 1))?
        .set(replacement)
        .commit()?;
    let both_bytes = both.package().exact_bytes();
    assert_eq!(
        fixture::format_keys(&both_bytes)?,
        vec![Some(new_key), Some(new_key)]
    );
    assert_eq!(
        fixture::format_entry_facts(&both_bytes)?,
        vec![(new_key, 2)]
    );

    let one_cleared = both
        .package()
        .edit_table_cell_date_time_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    let one_cleared_bytes = one_cleared.package().exact_bytes();
    assert_eq!(
        fixture::format_keys(&one_cleared_bytes)?,
        vec![None, Some(new_key)]
    );
    assert_eq!(
        fixture::format_entry_facts(&one_cleared_bytes)?,
        vec![(new_key, 1)]
    );

    let all_cleared = one_cleared
        .package()
        .edit_table_cell_date_time_format(0usize, 0usize, CellPosition::new(0, 1))?
        .clear()
        .commit()?;
    let all_cleared_bytes = all_cleared.package().exact_bytes();
    assert_eq!(fixture::format_keys(&all_cleared_bytes)?, vec![None, None]);
    assert!(fixture::format_entry_facts(&all_cleared_bytes)?.is_empty());
    assert_eq!(fixture::format_next_list_id(&all_cleared_bytes)?, 32);
    Ok(())
}

#[test]
fn date_time_rewrite_preserves_unknown_format_fields_and_nested_groups() -> TestResult {
    let bytes = source()?;
    let mut payload = raw_format_payload_by_key(&bytes, fixture::FIRST_FORMAT_KEY)?;
    let nested_before = raw_field_record(&payload, 94)?;
    append_varint_field(&mut payload, 46, 0x80_03)?;
    let group_body = {
        let mut body = Vec::new();
        append_varint_field(&mut body, 51, 9)?;
        body
    };
    payload.extend_from_slice(&encode_varint((50_u64 << 3) | 3));
    payload.extend_from_slice(&group_body);
    payload.extend_from_slice(&encode_varint((50_u64 << 3) | 4));
    let hostile =
        fixture::rewrite_format_payload_by_key(&bytes, fixture::FIRST_FORMAT_KEY, &payload)
            .map_err(|error| {
                io::Error::other(format!("Date & Time hostile fixture rewrite: {error}"))
            })?;
    let root_before = fixture::format_list_payload(&hostile)?;
    let root_90 = fixture::unknown_field_record(&root_before, 90)?;
    let root_94 = fixture::unknown_field_record(&root_before, 94)?;

    let hostile_package = Package::from_bytes(&hostile)
        .map_err(|error| io::Error::other(format!("Date & Time package ingress: {error}")))?;
    let target = hostile_package
        .edit_table_cell_date_time_format(0usize, 0usize, selected_position())?
        .set(date_time("MM/dd/yyyy")?)
        .commit()
        .map_err(|error| io::Error::other(format!("Date & Time rewrite: {error}")))?
        .package()
        .exact_bytes();
    let target_list = fixture::format_list_payload(&target)?;
    assert_eq!(fixture::unknown_field_record(&target_list, 90)?, root_90);
    assert_eq!(fixture::unknown_field_record(&target_list, 94)?, root_94);
    let target_key = fixture::format_keys(&target)?[0]
        .ok_or_else(|| io::Error::other("Date & Time target key is missing"))?;
    let target_payload = raw_format_payload_by_key(&target, target_key)?;
    assert_eq!(raw_field_record(&target_payload, 94)?, nested_before);
    assert_eq!(raw_field_record(&target_payload, 46)?, {
        let mut expected = Vec::new();
        append_varint_field(&mut expected, 46, 0x80_03)?;
        expected
    });
    let group = {
        let mut expected = encode_varint((50_u64 << 3) | 3);
        let body = {
            let mut body = Vec::new();
            append_varint_field(&mut body, 51, 9)?;
            body
        };
        expected.extend_from_slice(&body);
        expected.extend_from_slice(&encode_varint((50_u64 << 3) | 4));
        expected
    };
    assert_eq!(raw_field_record(&target_payload, 50)?, group);
    assert_exact_locality(&hostile, &target)?;
    assert_non_format_cell_state(&hostile, &target)?;
    Ok(())
}

#[test]
fn date_time_malformed_format_entries_fail_closed_atomically() -> TestResult {
    for corruption in [
        fixture::Corruption::DuplicateFormatKey,
        fixture::Corruption::MissingFormatEntry,
        fixture::Corruption::RefcountMismatch,
        fixture::Corruption::MalformedFormatPayload,
        fixture::Corruption::UnsupportedFormatType,
        fixture::Corruption::WrongCellFormatKey,
    ] {
        assert_owner_rejects(&fixture::corrupted_package_for(
            fixture::FormatFamily::DateTime,
            corruption,
        )?)?;
    }

    let source = source()?;
    let zero_type =
        fixture::rewrite_format_varint_by_key(&source, fixture::FIRST_FORMAT_KEY, 1, 0)?;
    assert_owner_rejects(&zero_type)?;

    // A zero list reference is a malformed owner graph even though the
    // surrounding package remains parseable.
    let zero_key = fixture::rewrite_tile_cells(&source, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("Date & Time fixture first cell is missing"))?;
        let mut cell = BncCell::parse(first)?;
        cell.set_data_format_identifier(0, CellDataFormatKind::DateTime, None)?;
        *first = cell.encode();
        Ok(())
    })?;
    assert_owner_rejects(&zero_key)?;
    Ok(())
}

#[test]
fn date_time_changed_edit_reads_locked_but_refuses_publication() -> TestResult {
    let bytes = fixture::locked_table_package(&source()?)?;
    let package = Package::from_bytes(&bytes)?;
    assert_eq!(
        package.table_cell_date_time_format(0usize, 0usize, selected_position())?,
        Some(date_time("yyyy-MM-dd H:mm:ss")?)
    );
    let before = package.exact_bytes();
    let error = package
        .edit_table_cell_date_time_format(0usize, 0usize, selected_position())?
        .set(date_time("MM/dd/yyyy")?)
        .commit()
        .expect_err("a changed Date & Time edit must refuse a locked table");
    assert!(
        matches!(error, Error::TableLocked { path: Path::Cell { sheet: 0, table: 0, position } } if position == selected_position())
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn date_time_patch_rejects_stale_and_foreign_sources_atomically() -> TestResult {
    let bytes = source()?;
    let package = Package::from_bytes(&bytes)?;
    let commit = package
        .edit_table_cell_date_time_format(0usize, 0usize, selected_position())?
        .set(date_time("MM/dd/yyyy")?)
        .commit()?;
    let target = commit.package().exact_bytes();

    let stale = Package::from_bytes(&target)?;
    let stale_before = stale.exact_bytes();
    assert!(matches!(
        stale.apply_table_cell_date_time_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(stale.exact_bytes(), stale_before);

    let foreign = Package::from_bytes(&fixture::synthetic_package_for(
        fixture::FormatFamily::DateTime,
        fixture::FormatSharing::Unshared,
    )?)?;
    let foreign_before = foreign.exact_bytes();
    assert!(matches!(
        foreign.apply_table_cell_date_time_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(foreign.exact_bytes(), foreign_before);
    Ok(())
}
