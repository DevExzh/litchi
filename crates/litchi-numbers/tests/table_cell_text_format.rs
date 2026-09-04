//! Exact-source integration coverage for Numbers table-cell Text formats.
//!
//! Text is a marker format: unlike the configurable numeric families, its
//! public value carries no native identifier or style fields.  The fixture
//! still exercises the complete native graph, including the converted-text
//! (`0x81`) BNC shape, so these tests prove that the small public API does not
//! weaken ownership, copy-on-write, wire-preservation, or resource-boundary
//! guarantees.

use std::{fmt::Debug, io, path::PathBuf, sync::Arc, thread};

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{WireView, append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::{tsk, tst};
use litchi_numbers::cell::data_format::{
    Text,
    currency::transaction as currency_transaction,
    fraction::transaction as fraction_transaction,
    number::transaction as number_transaction,
    percentage::transaction as percentage_transaction,
    scientific::transaction as scientific_transaction,
    text::transaction::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path},
};
use litchi_numbers::{
    CellPosition, Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
};
use litchi_numbers_wire::{BncCell, CellDataFormatKind, StoredValue};
use prost::Message as _;

#[path = "support/table_cell_data_format_fixture.rs"]
mod fixture;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const BNC_CELL_FORMAT_IDENTIFIER_FLAG: u32 = 0x0000_2000;
const BNC_RESERVED_KNOWN_FIELD_FLAG: u32 = 0x0010_0000;

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

fn sibling_position() -> CellPosition {
    CellPosition::new(fixture::SECOND_CELL.0 as u32, fixture::SECOND_CELL.1 as u32)
}

fn shared_source() -> TestResult<Vec<u8>> {
    fixture::synthetic_package_for(fixture::FormatFamily::Text, fixture::FormatSharing::Shared)
}

fn inherited_source() -> TestResult<Vec<u8>> {
    fixture::text_inherited_first_package()
}

fn converted_source() -> TestResult<Vec<u8>> {
    fixture::text_converted_package()
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
                candidate.raw_record().local_record(),
                "unchanged member {} lost its exact local ZIP record",
                entry.name()
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

fn normalized_cell(cell: &[u8]) -> TestResult<Vec<u8>> {
    let mut parsed = BncCell::parse(cell)?;
    parsed.clear_explicit_format();
    Ok(parsed.encode())
}

fn assert_non_format_bnc_bytes(source: &[u8], target: &[u8]) -> TestResult {
    let before = fixture::tile_cells(source)?;
    let after = fixture::tile_cells(target)?;
    assert_eq!(before.len(), after.len());
    for (before, after) in before.iter().zip(after.iter()) {
        assert_eq!(
            BncCell::parse(before)?.stored_value(),
            BncCell::parse(after)?.stored_value(),
            "Text format edit changed the stored scalar"
        );
        assert_eq!(
            normalized_cell(before)?,
            normalized_cell(after)?,
            "Text format edit changed non-format BNC bytes"
        );
    }
    Ok(())
}

/// A malformed envelope may be refused by package ingress.  When ingress
/// admits it, require both focused Text entry points to refuse it without
/// publishing bytes.
fn assert_rejected_or_owner(source: &[u8]) -> TestResult {
    let Ok(package) = Package::from_bytes(source) else {
        return Ok(());
    };
    let before = package.exact_bytes();
    assert!(
        package
            .table_cell_text_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert!(
        package
            .edit_table_cell_text_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

fn assert_text_cell(source: &[u8], expected_flags: u16, expected_primary: u32) -> TestResult {
    let cell_bytes = fixture::tile_cells(source)?
        .into_iter()
        .next()
        .ok_or_else(|| io::Error::other("Text fixture first cell is missing"))?;
    let cell = BncCell::parse(&cell_bytes)?;
    assert_eq!(cell.explicit_format_flags(), expected_flags);
    assert_eq!(
        cell.cell_format_kind(),
        Some(litchi_numbers_wire::TEXT_CELL_FORMAT_KIND)
    );
    assert_eq!(cell.format_identifier(), Some(expected_primary));
    assert!(matches!(cell.stored_value(), StoredValue::Text(_)));
    Ok(())
}

/// Recover the generic identifier in the converted Text shape without adding
/// that native ID to the low-level/public API.  The fixture is intentionally
/// the canonical fixed-width BNC v5 layout: string key, kind, generic Number
/// key, then Text key.
fn converted_generic_identifier(source: &[u8]) -> TestResult<u32> {
    let cell = fixture::tile_cells(source)?
        .into_iter()
        .next()
        .ok_or_else(|| io::Error::other("Text fixture first cell is missing"))?;
    let parsed = BncCell::parse(&cell)?;
    assert_eq!(
        parsed.explicit_format_flags(),
        litchi_numbers_wire::EXPLICIT_CONVERTED_TEXT_FORMAT
    );
    assert!(cell.len() >= 28);
    Ok(u32::from_le_bytes(cell[20..24].try_into()?))
}

fn rewrite_marker(source: &[u8], marker: u16) -> TestResult<Vec<u8>> {
    fixture::rewrite_tile_cells(source, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("Text fixture first cell is missing"))?;
        if first.len() < 8 {
            return Err(io::Error::other("Text fixture first cell is truncated").into());
        }
        first[6..8].copy_from_slice(&marker.to_le_bytes());
        Ok(())
    })
}

/// Read one known materialized cell from the checked-in native fixture.  The
/// helper is deliberately scoped to the fixture's one tile member rather
/// than exposing a general raw-package reader in the test API.
fn native_fixture_cell(source: &[u8], row_index: u32, column: usize) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == "Index/Tables/Tile.iwa")
        .ok_or_else(|| io::Error::other("native Tile member is missing"))?;
    let decompressed = SnappyStream::decompress(entry.data())?;
    let archive = Archive::parse(decompressed.as_bytes())?;
    let message = archive
        .objects
        .iter()
        .flat_map(|object| object.messages.iter())
        .find(|message| message.type_ == 6_002)
        .ok_or_else(|| io::Error::other("native Tile message is missing"))?;
    let tile = tst::Tile::decode(message.data.as_slice())?;
    let row = tile
        .row_infos
        .iter()
        .find(|row| row.tile_row_index == row_index)
        .ok_or_else(|| io::Error::other("native target row is missing"))?;
    let storage = row
        .cell_storage_buffer
        .as_deref()
        .ok_or_else(|| io::Error::other("native target row storage is missing"))?;
    let offsets = row
        .cell_offsets
        .as_deref()
        .ok_or_else(|| io::Error::other("native target row offsets are missing"))?;
    let offset = column
        .checked_mul(2)
        .ok_or_else(|| io::Error::other("native target column overflows"))?;
    let raw_start = u16::from_le_bytes(
        offsets
            .get(offset..offset + 2)
            .ok_or_else(|| io::Error::other("native target column is outside offsets"))?
            .try_into()?,
    );
    if raw_start == u16::MAX {
        return Err(io::Error::other("native target cell is not materialized").into());
    }
    let unit = if row.has_wide_offsets.unwrap_or(false) {
        4
    } else {
        1
    };
    let start = usize::from(raw_start)
        .checked_mul(unit)
        .ok_or_else(|| io::Error::other("native target cell offset overflows"))?;
    let end = offsets
        .chunks_exact(2)
        .skip(column + 1)
        .filter_map(|encoded| {
            let raw = u16::from_le_bytes([encoded[0], encoded[1]]);
            (raw != u16::MAX).then_some(usize::from(raw).saturating_mul(unit))
        })
        .next()
        .unwrap_or(storage.len());
    if start > end || end > storage.len() {
        return Err(io::Error::other("native target cell range is invalid").into());
    }
    Ok(storage[start..end].to_vec())
}

/// Find the native format-list entry for one key in the checked-in fixture.
/// This validates the type-260 provenance asserted by the focused Text test
/// without making native identifiers part of the public Numbers API.
fn native_fixture_format_type(source: &[u8], key: u32) -> TestResult<u32> {
    let catalog = Catalog::from_bytes(source)?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().starts_with("Index/Tables/DataList"))
    {
        let decompressed = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(decompressed.as_bytes())?;
        for message in archive
            .objects
            .iter()
            .flat_map(|object| object.messages.iter())
            .filter(|message| message.type_ == 6_005)
        {
            let Ok(list) = tst::TableDataList::decode(message.data.as_slice()) else {
                continue;
            };
            if list.list_type != tst::table_data_list::ListType::Format as i32 {
                continue;
            }
            if let Some(format) = list
                .entries
                .iter()
                .find(|entry| entry.key == key)
                .and_then(|entry| entry.format.as_ref())
                .and_then(|format| format.format_type)
            {
                return Ok(format);
            }
        }
    }
    Err(io::Error::other(format!("native format entry {key} is missing")).into())
}

#[test]
fn text_transaction_types_are_strictly_typed_send_sync_and_redacted() -> TestResult {
    fn assert_send_sync_debug<T: Send + Sync + Debug>() {}

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<Text>();
    assert_send_sync_debug::<Edit<'static>>();
    assert_send_sync_debug::<Commit>();
    assert_send_sync_debug::<Patch>();
    assert_send_sync_debug::<Diagnostics>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<LimitKind>();
    assert_send_sync_debug::<Path>();

    let package = Package::from_bytes(&shared_source()?)?;
    let edit = package.edit_table_cell_text_format(
        SheetSelector::name("Data Format Sheet"),
        TableSelector::name("Data Formats"),
        selected_position(),
    )?;
    let rendered = format!("{edit:?}");
    assert!(!rendered.contains("Index/"));
    assert!(!rendered.contains("Document.iwa"));
    assert!(!rendered.contains("data-format-table-id"));
    let commit = edit.set(Text).commit()?;
    for rendered in [
        format!("{commit:?}"),
        format!("{:?}", commit.patch()),
        format!("{:?}", commit.diagnostics()),
    ] {
        assert!(!rendered.contains("Index/"));
        assert!(!rendered.contains("Document.iwa"));
        assert!(!rendered.contains("data-format-table-id"));
    }
    Ok(())
}

#[test]
fn text_selectors_and_option_semantics_are_explicit() -> TestResult {
    let source = shared_source()?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.table_cell_text_format(0usize, 0usize, selected_position())?,
        Some(Text)
    );
    assert_eq!(
        package.table_cell_text_format(
            SheetSelector::name("Data Format Sheet"),
            TableSelector::name("Data Formats"),
            selected_position(),
        )?,
        Some(Text)
    );
    assert!(matches!(
        package.table_cell_text_format("missing sheet", 0usize, selected_position()),
        Err(Error::SheetNotFound)
    ));
    assert!(matches!(
        package.table_cell_text_format(0usize, "missing table", selected_position()),
        Err(Error::TableNotFound)
    ));

    let cleared = package
        .edit_table_cell_text_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .table_cell_text_format(0usize, 0usize, selected_position())?,
        None
    );
    let explicit = cleared
        .package()
        .edit_table_cell_text_format(0usize, 0usize, selected_position())?
        .set(Text)
        .commit()?;
    assert_eq!(
        explicit
            .package()
            .table_cell_text_format(0usize, 0usize, selected_position())?,
        Some(Text)
    );
    Ok(())
}

#[test]
fn text_plain_0x80_reads_and_same_value_is_an_exact_noop() -> TestResult {
    let source = shared_source()?;
    assert_text_cell(
        &source,
        litchi_numbers_wire::EXPLICIT_TEXT_FORMAT,
        fixture::FIRST_FORMAT_KEY,
    )?;
    let first = BncCell::parse(&fixture::tile_cells(&source)?[0])?;
    assert_eq!(first.secondary_format_identifier(), None);
    assert_eq!(first.stored_value(), StoredValue::Text(1));

    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_table_cell_text_format(0usize, 0usize, selected_position())?
        .set(Text)
        .commit()?;
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().exact_bytes(), source);
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert!(!commit.diagnostics().full_reparse_performed());
    Ok(())
}

#[test]
fn text_converted_0x81_requires_and_preserves_generic_number_reference() -> TestResult {
    let source = converted_source()?;
    assert_text_cell(
        &source,
        litchi_numbers_wire::EXPLICIT_CONVERTED_TEXT_FORMAT,
        fixture::FIRST_FORMAT_KEY,
    )?;
    assert_eq!(
        converted_generic_identifier(&source)?,
        fixture::SECOND_FORMAT_KEY
    );
    assert_eq!(
        fixture::format_payload_by_key(&source, fixture::SECOND_FORMAT_KEY)
            .ok()
            .and_then(|payload| tsk::FormatStructArchive::decode(payload.as_slice()).ok())
            .and_then(|format| format.format_type),
        Some(fixture::NATIVE_NUMBER_FORMAT_TYPE)
    );

    let package = Package::from_bytes(&source)?;
    let no_op = package
        .edit_table_cell_text_format(0usize, 0usize, selected_position())?
        .set(Text)
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().exact_bytes(), source);

    let cleared = package
        .edit_table_cell_text_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    let cleared_bytes = cleared.package().exact_bytes();
    let cleared_cell = BncCell::parse(&fixture::tile_cells(&cleared_bytes)?[0])?;
    assert_eq!(cleared_cell.explicit_format_flags(), 0);
    assert_eq!(cleared_cell.cell_format_kind(), None);
    assert_eq!(cleared_cell.format_identifier(), None);
    assert_eq!(fixture::format_keys(&cleared_bytes)?, vec![None, Some(1)]);
    assert_eq!(fixture::format_entry_facts(&cleared_bytes)?, vec![(1, 1)]);
    assert_non_format_bnc_bytes(&source, &cleared_bytes)?;

    // The inverse of clearing converted Text must restore the exact native
    // marker and both format-list edges, not merely the public `Some(Text)`
    // value.
    let inverse = cleared.patch().inverse();
    let restored_exact =
        Package::from_bytes(&cleared_bytes)?.apply_table_cell_text_format(&inverse)?;
    let restored_exact_bytes = restored_exact.package().exact_bytes();
    assert_eq!(restored_exact_bytes, source);
    assert_eq!(
        converted_generic_identifier(&restored_exact_bytes)?,
        fixture::SECOND_FORMAT_KEY
    );
    assert_text_cell(
        &restored_exact_bytes,
        litchi_numbers_wire::EXPLICIT_CONVERTED_TEXT_FORMAT,
        fixture::FIRST_FORMAT_KEY,
    )?;
    assert_eq!(
        fixture::format_keys(&restored_exact_bytes)?,
        vec![
            Some(fixture::FIRST_FORMAT_KEY),
            Some(fixture::FIRST_FORMAT_KEY)
        ]
    );
    assert_eq!(
        fixture::format_entry_facts(&restored_exact_bytes)?,
        vec![
            (fixture::FIRST_FORMAT_KEY, 2),
            (fixture::SECOND_FORMAT_KEY, 1)
        ]
    );

    let restored = cleared
        .package()
        .edit_table_cell_text_format(0usize, 0usize, selected_position())?
        .set(Text)
        .commit()?;
    let restored_bytes = restored.package().exact_bytes();
    assert_text_cell(
        &restored_bytes,
        litchi_numbers_wire::EXPLICIT_TEXT_FORMAT,
        fixture::FIRST_FORMAT_KEY,
    )?;
    assert_eq!(
        fixture::format_keys(&restored_bytes)?,
        vec![Some(1), Some(1)]
    );
    assert_eq!(fixture::format_entry_facts(&restored_bytes)?, vec![(1, 2)]);
    assert_non_format_bnc_bytes(&source, &restored_bytes)?;
    Ok(())
}

#[test]
fn text_set_clear_reset_inverse_and_apply_are_exact() -> TestResult {
    let source = inherited_source()?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.table_cell_text_format(0usize, 0usize, selected_position())?,
        None
    );
    assert_eq!(
        package.table_cell_text_format(0usize, 0usize, sibling_position())?,
        Some(Text)
    );

    let changed = package
        .edit_table_cell_text_format(0usize, 0usize, selected_position())?
        .set(Text)
        .commit()?;
    let target = changed.package().exact_bytes();
    assert!(!changed.patch().is_noop());
    assert_eq!(changed.patch().before(), None);
    assert_eq!(changed.patch().after(), Some(&Text));
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().touched_components(), 1);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_eq!(
        changed
            .package()
            .table_cell_text_format(0usize, 0usize, selected_position())?,
        Some(Text)
    );
    assert_eq!(
        changed
            .package()
            .table_cell_text_format(0usize, 0usize, sibling_position())?,
        Some(Text)
    );
    assert_exact_locality(&source, &target)?;
    assert_non_format_bnc_bytes(&source, &target)?;

    let applied = package.apply_table_cell_text_format(changed.patch())?;
    assert_eq!(applied.package().exact_bytes(), target);
    let inverse = changed.patch().inverse();
    assert_eq!(inverse.inverse(), *changed.patch());
    let restored = Package::from_bytes(&target)?.apply_table_cell_text_format(&inverse)?;
    assert_eq!(restored.package().exact_bytes(), source);
    assert_eq!(
        restored
            .package()
            .table_cell_text_format(0usize, 0usize, selected_position())?,
        None
    );

    let cleared = changed
        .package()
        .edit_table_cell_text_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .table_cell_text_format(0usize, 0usize, selected_position())?,
        None
    );
    let reset = cleared
        .package()
        .edit_table_cell_text_format(0usize, 0usize, selected_position())?
        .reset()
        .commit()?;
    assert!(reset.patch().is_noop());
    assert_eq!(
        reset.package().exact_bytes(),
        cleared.package().exact_bytes()
    );
    Ok(())
}

#[test]
fn text_patch_apply_rejects_stale_foreign_and_unrelated_graph_sources_atomically() -> TestResult {
    let source = inherited_source()?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_table_cell_text_format(0usize, 0usize, selected_position())?
        .set(Text)
        .commit()?;
    let target = commit.package().exact_bytes();

    let stale = Package::from_bytes(&target)?;
    let stale_before = stale.exact_bytes();
    assert!(matches!(
        stale.apply_table_cell_text_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(stale.exact_bytes(), stale_before);

    let foreign = Package::from_bytes(&shared_source()?)?;
    let foreign_before = foreign.exact_bytes();
    assert!(matches!(
        foreign.apply_table_cell_text_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(foreign.exact_bytes(), foreign_before);

    let malformed = Package::from_bytes(&fixture::corrupted_package_for(
        fixture::FormatFamily::Text,
        fixture::Corruption::UnexpectedFieldReference,
    )?)?;
    let malformed_before = malformed.exact_bytes();
    assert!(matches!(
        malformed.apply_table_cell_text_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(malformed.exact_bytes(), malformed_before);
    Ok(())
}

#[test]
fn text_shared_copy_on_write_reuses_keys_and_culls_zero_refcounts() -> TestResult {
    let source = inherited_source()?;
    assert_eq!(fixture::format_entry_facts(&source)?, vec![(1, 1)]);
    assert_eq!(fixture::format_next_list_id(&source)?, 32);
    assert_eq!(fixture::format_keys(&source)?, vec![None, Some(1)]);

    let attached = Package::from_bytes(&source)?
        .edit_table_cell_text_format(0usize, 0usize, selected_position())?
        .set(Text)
        .commit()?;
    let attached_bytes = attached.package().exact_bytes();
    assert_eq!(
        fixture::format_keys(&attached_bytes)?,
        vec![Some(1), Some(1)]
    );
    assert_eq!(fixture::format_entry_facts(&attached_bytes)?, vec![(1, 2)]);
    assert_eq!(fixture::format_next_list_id(&attached_bytes)?, 32);

    let one_cleared = attached
        .package()
        .edit_table_cell_text_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    let one_cleared_bytes = one_cleared.package().exact_bytes();
    assert_eq!(
        fixture::format_keys(&one_cleared_bytes)?,
        vec![None, Some(1)]
    );
    assert_eq!(
        fixture::format_entry_facts(&one_cleared_bytes)?,
        vec![(1, 1)]
    );

    let all_cleared = one_cleared
        .package()
        .edit_table_cell_text_format(0usize, 0usize, sibling_position())?
        .clear()
        .commit()?;
    let all_cleared_bytes = all_cleared.package().exact_bytes();
    assert_eq!(fixture::format_keys(&all_cleared_bytes)?, vec![None, None]);
    assert!(fixture::format_entry_facts(&all_cleared_bytes)?.is_empty());
    assert_eq!(fixture::format_next_list_id(&all_cleared_bytes)?, 32);
    Ok(())
}

#[test]
fn text_rewrite_preserves_unknown_format_wire_and_member_locality() -> TestResult {
    // Clearing the selected member leaves the shared entry live through the
    // sibling.  This makes the changed operation exercise list surgery while
    // still requiring the old entry's unknown bytes to survive unchanged.
    let source = shared_source()?;
    let mut payload = fixture::format_payload_by_key(&source, fixture::FIRST_FORMAT_KEY)?;
    let original_nested_extension = fixture::unknown_field_record(&payload, 94)?;
    append_varint_field(&mut payload, 46, 0x80_03)?;
    append_length_delimited_field(&mut payload, 47, b"text extension")?;
    let hostile =
        fixture::rewrite_format_payload_by_key(&source, fixture::FIRST_FORMAT_KEY, &payload)?;
    let root_before = fixture::format_list_payload(&hostile)?;
    let root_90 = fixture::unknown_field_record(&root_before, 90)?;
    let root_94 = fixture::unknown_field_record(&root_before, 94)?;

    let commit = Package::from_bytes(&hostile)?
        .edit_table_cell_text_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    let target = commit.package().exact_bytes();
    let target_payload = fixture::format_list_payload(&target)?;
    assert_eq!(fixture::unknown_field_record(&target_payload, 90)?, root_90);
    assert_eq!(fixture::unknown_field_record(&target_payload, 94)?, root_94);
    assert_eq!(fixture::format_keys(&target)?, vec![None, Some(1)]);
    let target_nested = fixture::format_payload_by_key(&target, fixture::FIRST_FORMAT_KEY)?;
    assert_eq!(
        fixture::unknown_field_record(&target_nested, 94)?,
        original_nested_extension
    );
    for field_number in [46, 47] {
        assert!(
            WireView::parse(&target_nested)?
                .fields()
                .any(|field| field.number() == field_number),
            "unknown field {field_number} was not retained"
        );
    }
    let decoded = tsk::FormatStructArchive::decode(target_nested.as_slice())?;
    assert_eq!(decoded.format_type, Some(fixture::NATIVE_TEXT_FORMAT_TYPE));
    assert_exact_locality(&hostile, &target)?;
    assert_non_format_bnc_bytes(&hostile, &target)?;
    Ok(())
}

#[test]
fn text_native_fixture_marker_zero_promotes_and_number_boundary_is_typed() -> TestResult {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/basic.numbers");
    let package = Package::open(path)?;
    let source = package.exact_bytes();
    let b2 = CellPosition::from_a1("B2")?;
    let b2_before = native_fixture_cell(&source, 1, 1)?;
    let b2_before = BncCell::parse(&b2_before)?;
    assert_eq!(b2_before.explicit_format_flags(), 0);
    assert_eq!(
        b2_before.cell_format_kind(),
        Some(litchi_numbers_wire::TEXT_CELL_FORMAT_KIND)
    );
    assert_eq!(
        b2_before.format_identifier(),
        Some(fixture::FIRST_FORMAT_KEY)
    );
    assert!(matches!(b2_before.stored_value(), StoredValue::Text(_)));
    assert_eq!(
        native_fixture_format_type(&source, fixture::FIRST_FORMAT_KEY)?,
        fixture::NATIVE_TEXT_FORMAT_TYPE
    );
    assert_eq!(
        package.table_cell_text_format(0usize, 0usize, b2)?,
        None,
        "native marker-zero Text remains automatic rather than explicit"
    );
    let b2_clear = package
        .edit_table_cell_text_format(0usize, 0usize, b2)?
        .clear()
        .commit()?;
    assert!(b2_clear.patch().is_noop());
    assert_eq!(b2_clear.package().exact_bytes(), source);

    let b2_set = package
        .edit_table_cell_text_format(0usize, 0usize, b2)?
        .set(Text)
        .commit()?;
    let b2_after = native_fixture_cell(&b2_set.package().exact_bytes(), 1, 1)?;
    let b2_after = BncCell::parse(&b2_after)?;
    assert_eq!(
        b2_after.explicit_format_flags(),
        litchi_numbers_wire::EXPLICIT_TEXT_FORMAT
    );
    assert_eq!(
        b2_after.cell_format_kind(),
        Some(litchi_numbers_wire::TEXT_CELL_FORMAT_KIND)
    );
    assert_eq!(
        b2_after.format_identifier(),
        Some(fixture::FIRST_FORMAT_KEY)
    );
    assert_eq!(b2_after.stored_value(), b2_before.stored_value());
    assert_eq!(
        b2_set
            .package()
            .table_cell_text_format(0usize, 0usize, b2)?,
        Some(Text)
    );

    let b3 = CellPosition::from_a1("B3")?;
    assert!(matches!(
        package.table_cell_text_format(0usize, 0usize, b3),
        Err(Error::WrongFormatFamily { .. })
    ));
    Ok(())
}

#[test]
fn text_rewrite_preserves_scalar_style_comment_and_opaque_cell_bytes() -> TestResult {
    let source = fixture::rewrite_tile_cells(&inherited_source()?, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("Text fixture first cell is missing"))?;
        let mut cell = BncCell::parse(first)?;
        cell.set_style_identifier(Some(23));
        cell.set_text_style_identifier(Some(29));
        cell.set_comment_identifier(Some(31));
        let mut encoded = cell.encode();
        encoded.extend_from_slice(b"text-format-tail");
        *first = encoded;
        Ok(())
    })?;
    let before = BncCell::parse(&fixture::tile_cells(&source)?[0])?;
    assert_eq!(before.stored_value(), StoredValue::Text(1));
    let target = Package::from_bytes(&source)?
        .edit_table_cell_text_format(0usize, 0usize, selected_position())?
        .set(Text)
        .commit()?
        .package()
        .exact_bytes();
    let after = BncCell::parse(&fixture::tile_cells(&target)?[0])?;
    assert_eq!(after.stored_value(), before.stored_value());
    assert_eq!(after.style_identifier(), before.style_identifier());
    assert_eq!(
        after.text_style_identifier(),
        before.text_style_identifier()
    );
    assert_eq!(after.comment_identifier(), before.comment_identifier());
    assert_non_format_bnc_bytes(&source, &target)?;
    Ok(())
}

#[test]
fn text_wrong_family_boundaries_are_typed_and_symmetric() -> TestResult {
    for family in [
        fixture::FormatFamily::Number,
        fixture::FormatFamily::Percentage,
        fixture::FormatFamily::Currency,
        fixture::FormatFamily::Scientific,
        fixture::FormatFamily::Fraction,
    ] {
        let source = fixture::synthetic_package_for(family, fixture::FormatSharing::Shared)?;
        let package = Package::from_bytes(&source)?;
        let before = package.exact_bytes();
        assert!(matches!(
            package.table_cell_text_format(0usize, 0usize, selected_position()),
            Err(Error::WrongFormatFamily { .. })
        ));
        assert!(matches!(
            package.edit_table_cell_text_format(0usize, 0usize, selected_position()),
            Err(Error::WrongFormatFamily { .. })
        ));
        assert_eq!(package.exact_bytes(), before);
    }

    let source = shared_source()?;
    let package = Package::from_bytes(&source)?;
    assert!(matches!(
        package.table_cell_number_format(0usize, 0usize, selected_position()),
        Err(number_transaction::Error::WrongFormatFamily { .. })
    ));
    assert!(matches!(
        package.table_cell_percentage_format(0usize, 0usize, selected_position()),
        Err(percentage_transaction::Error::WrongFormatFamily { .. })
    ));
    assert!(matches!(
        package.table_cell_currency_format(0usize, 0usize, selected_position()),
        Err(currency_transaction::Error::WrongFormatFamily { .. })
    ));
    assert!(matches!(
        package.table_cell_scientific_format(0usize, 0usize, selected_position()),
        Err(scientific_transaction::Error::WrongFormatFamily { .. })
    ));
    assert!(matches!(
        package.table_cell_fraction_format(0usize, 0usize, selected_position()),
        Err(fraction_transaction::Error::WrongFormatFamily { .. })
    ));
    Ok(())
}

#[test]
fn text_coordinate_boundaries_return_cell_not_found_without_mutation() -> TestResult {
    let package = Package::from_bytes(&shared_source()?)?;
    let before = package.exact_bytes();
    for position in [
        CellPosition::new(1, 0),
        CellPosition::new(0, 2),
        CellPosition::new(u32::MAX, u32::MAX),
    ] {
        assert!(matches!(
            package.table_cell_text_format(0usize, 0usize, position),
            Err(Error::CellNotFound)
        ));
        assert!(matches!(
            package.edit_table_cell_text_format(0usize, 0usize, position),
            Err(Error::CellNotFound)
        ));
    }
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn text_malformed_wire_and_graph_inputs_fail_closed_and_remain_atomic() -> TestResult {
    for corruption in [
        fixture::Corruption::DuplicateFormatKey,
        fixture::Corruption::MissingFormatEntry,
        fixture::Corruption::RefcountMismatch,
        fixture::Corruption::DuplicateFormatList,
        fixture::Corruption::AliasedFormatList,
        fixture::Corruption::MalformedFormatPayload,
        fixture::Corruption::UnterminatedUnknownGroup,
        fixture::Corruption::UnsupportedFormatType,
        fixture::Corruption::WrongCellFormatKey,
        fixture::Corruption::FractionReplacementMetadataTrue,
    ] {
        assert_rejected_or_owner(&fixture::corrupted_package_for(
            fixture::FormatFamily::Text,
            corruption,
        )?)?;
    }

    // Converted Text is not valid merely because its marker is changed: the
    // 0x81 shape must carry a nonzero generic Number identifier.
    assert_rejected_or_owner(&rewrite_marker(
        &shared_source()?,
        litchi_numbers_wire::EXPLICIT_CONVERTED_TEXT_FORMAT,
    )?)?;

    // Conversely, a generic secondary field under the plain 0x80 marker is
    // also a family-shape mismatch.  Insert it in the fixed BNC field order.
    let generic_under_plain = fixture::rewrite_tile_cells(&shared_source()?, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("Text fixture first cell is missing"))?;
        if first.len() < 24 {
            return Err(io::Error::other("Text fixture first cell is truncated").into());
        }
        let flags = u32::from_le_bytes(first[8..12].try_into()?);
        first[8..12].copy_from_slice(&(flags | BNC_CELL_FORMAT_IDENTIFIER_FLAG).to_le_bytes());
        first.splice(20..20, fixture::SECOND_FORMAT_KEY.to_le_bytes());
        Ok(())
    })?;
    assert_rejected_or_owner(&generic_under_plain)?;

    // A converted cell with a zero generic key has the right field shape but
    // still violates the native 0x81 contract.
    let zero_generic = fixture::rewrite_tile_cells(&converted_source()?, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("Text fixture first cell is missing"))?;
        if first.len() < 24 {
            return Err(io::Error::other("Text fixture first cell is truncated").into());
        }
        first[20..24].fill(0);
        Ok(())
    })?;
    assert_rejected_or_owner(&zero_generic)?;

    let source = shared_source()?;
    for payload in [
        vec![0x08],
        vec![0x08, 0x84, 0x02, 0x5a],
        vec![0x0a, 0x01, 0x08],
    ] {
        assert_rejected_or_owner(&fixture::rewrite_format_payload_by_key(
            &source,
            fixture::FIRST_FORMAT_KEY,
            &payload,
        )?)?;
    }

    let mut wrong_type = Vec::new();
    append_varint_field(
        &mut wrong_type,
        1,
        u64::from(fixture::NATIVE_NUMBER_FORMAT_TYPE),
    )?;
    assert_rejected_or_owner(&fixture::rewrite_format_payload_by_key(
        &source,
        fixture::FIRST_FORMAT_KEY,
        &wrong_type,
    )?)?;

    let mut duplicate_type = Vec::new();
    append_varint_field(
        &mut duplicate_type,
        1,
        u64::from(fixture::NATIVE_TEXT_FORMAT_TYPE),
    )?;
    append_varint_field(
        &mut duplicate_type,
        1,
        u64::from(fixture::NATIVE_TEXT_FORMAT_TYPE),
    )?;
    assert_rejected_or_owner(&fixture::rewrite_format_payload_by_key(
        &source,
        fixture::FIRST_FORMAT_KEY,
        &duplicate_type,
    )?)?;
    Ok(())
}

#[test]
fn text_input_and_operation_budgets_reject_before_publication() -> TestResult {
    let source = inherited_source()?;
    let exact_limits = PackageLimits::new(
        u64::try_from(source.len())?,
        PackageLimits::MAX_ENTRIES,
        PackageLimits::MAX_ENTRY_BYTES,
        PackageLimits::MAX_TOTAL_BYTES,
        PackageLimits::MAX_IWA_STREAM_BYTES,
    )?;
    let exact = Package::from_bytes_with_options(
        &source,
        PackageReadOptions::new(exact_limits, PackageSemanticLimits::default()),
    )?;
    assert_eq!(
        exact.table_cell_text_format(0usize, 0usize, selected_position())?,
        None
    );

    let tight = PackageLimits::new(
        u64::try_from(source.len().saturating_sub(1))?,
        PackageLimits::MAX_ENTRIES,
        PackageLimits::MAX_ENTRY_BYTES,
        PackageLimits::MAX_TOTAL_BYTES,
        PackageLimits::MAX_IWA_STREAM_BYTES,
    )?;
    assert!(
        Package::from_bytes_with_options(
            &source,
            PackageReadOptions::new(tight, PackageSemanticLimits::default()),
        )
        .is_err()
    );
    let semantic = PackageSemanticLimits::new(1, 1, 1, 1)?;
    assert!(
        Package::from_bytes_with_options(
            &source,
            PackageReadOptions::new(PackageLimits::default(), semantic),
        )
        .is_err()
    );

    let before = exact.exact_bytes();
    let result = exact
        .edit_table_cell_text_format(0usize, 0usize, selected_position())?
        .set(Text)
        .commit();
    assert!(
        matches!(result, Err(Error::LimitExceeded { .. })),
        "unexpected Text operation-budget result: {result:?}"
    );
    assert_eq!(exact.exact_bytes(), before);
    Ok(())
}

#[test]
fn text_concurrent_arc_reads_and_edits_are_independent_and_send_sync() -> TestResult {
    let package = Arc::new(Package::from_bytes(&inherited_source()?)?);
    let handles = (0..8)
        .map(|_| {
            let package = Arc::clone(&package);
            thread::spawn(move || -> Result<(Option<Text>, Option<Text>), Error> {
                let observed = package.table_cell_text_format(
                    SheetSelector::index(0),
                    TableSelector::index(0),
                    selected_position(),
                )?;
                let commit = package
                    .edit_table_cell_text_format(
                        SheetSelector::index(0),
                        TableSelector::index(0),
                        selected_position(),
                    )?
                    .set(Text)
                    .commit()?;
                let after = commit.package().table_cell_text_format(
                    SheetSelector::index(0),
                    TableSelector::index(0),
                    selected_position(),
                )?;
                Ok((observed, after))
            })
        })
        .collect::<Vec<_>>();
    for handle in handles {
        let (observed, after) = handle
            .join()
            .map_err(|_| io::Error::other("concurrent Text worker panicked"))??;
        assert_eq!(observed, None);
        assert_eq!(after, Some(Text));
    }
    assert_eq!(
        package.table_cell_text_format(0usize, 0usize, selected_position())?,
        None
    );
    Ok(())
}

#[test]
fn text_reopened_package_keeps_marker_family_and_selector_equivalence() -> TestResult {
    let source = inherited_source()?;
    let commit = Package::from_bytes(&source)?
        .edit_table_cell_text_format(0usize, 0usize, selected_position())?
        .set(Text)
        .commit()?;
    let bytes = commit.package().exact_bytes();
    let reopened = Package::from_bytes(&bytes)?;
    let by_index = reopened.table_cell_text_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        selected_position(),
    )?;
    let by_name = reopened.table_cell_text_format(
        SheetSelector::name("Data Format Sheet"),
        TableSelector::name("Data Formats"),
        selected_position(),
    )?;
    assert_eq!(by_index, Some(Text));
    assert_eq!(by_name, by_index);
    assert_text_cell(
        &bytes,
        litchi_numbers_wire::EXPLICIT_TEXT_FORMAT,
        fixture::FIRST_FORMAT_KEY,
    )?;
    Ok(())
}

#[test]
fn text_missing_metadata_is_rejected_without_publication() -> TestResult {
    let source = shared_source()?;
    let without_metadata = Catalog::from_bytes(&source)?.reassemble_with_deletions_to_bytes(
        &[],
        &[fixture::METADATA_MEMBER],
        Limits::default(),
    )?;
    assert_rejected_or_owner(&without_metadata)?;
    Ok(())
}

#[test]
fn text_bnc_shape_rejects_non_text_storage_without_mutation() -> TestResult {
    let source = shared_source()?;
    let numeric = fixture::rewrite_tile_cells(&source, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("Text fixture first cell is missing"))?;
        let mut cell = BncCell::parse(first)?;
        cell.set_number_or_percentage_format_identifier_preserving_value(Some(
            fixture::FIRST_FORMAT_KEY,
        ))?;
        *first = cell.encode();
        Ok(())
    })?;
    assert_rejected_or_owner(&numeric)?;

    let control = fixture::rewrite_tile_cells(&source, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("Text fixture first cell is missing"))?;
        let mut cell = BncCell::parse(first)?;
        cell.set_data_format_identifier(
            fixture::FIRST_FORMAT_KEY,
            CellDataFormatKind::PopUpMenu,
            Some(77),
        )?;
        *first = cell.encode();
        Ok(())
    })?;
    assert_rejected_or_owner(&control)?;
    Ok(())
}

#[test]
fn text_reserved_known_bnc_field_is_rejected_without_mutation() -> TestResult {
    let source = fixture::rewrite_tile_cells(&shared_source()?, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("Text fixture first cell is missing"))?;
        if first.len() < 24 {
            return Err(io::Error::other("Text fixture first cell is truncated").into());
        }
        let flags = u32::from_le_bytes(first[8..12].try_into()?);
        first[8..12].copy_from_slice(&(flags | BNC_RESERVED_KNOWN_FIELD_FLAG).to_le_bytes());
        // The reserved known field is a fixed-width four-byte slot at the end
        // of the BNC v5 metadata layout.  Keep it parseable so the focused
        // owner, rather than package ingress, makes the refusal decision.
        first.extend_from_slice(&[0, 0, 0, 0]);
        Ok(())
    })?;
    assert_rejected_or_owner(&source)?;
    Ok(())
}
