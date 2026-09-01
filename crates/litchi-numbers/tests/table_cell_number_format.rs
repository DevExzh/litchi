//! Exact-source integration coverage for Numbers table-cell number formats.
//!
//! The shared fixture models a rooted table, tile, format list, and metadata
//! registry.  The public tests intentionally select a sheet, table, and
//! [`CellPosition`] and never expose the native graph outside fixture helpers.

use std::{fmt::Debug, io, path::PathBuf, sync::Arc, thread};

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::wire::{WireView, append_length_delimited_field, patch_nested_varint_field};
use litchi_iwa_protos::tst;
use litchi_numbers::cell::data_format::number::transaction::{
    Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path,
};
use litchi_numbers::cell::data_format::number::{
    DecimalPlaces, NegativeStyle, Number, ThousandsSeparator,
};
use litchi_numbers::{
    CellPosition, Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
};
use litchi_numbers_wire::BncCell;
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

fn number(
    decimal_places: DecimalPlaces,
    negative_style: NegativeStyle,
    separator: ThousandsSeparator,
) -> Number {
    Number::new(decimal_places, negative_style, separator)
}

fn fixed_number(
    places: u8,
    negative_style: NegativeStyle,
    separator: ThousandsSeparator,
) -> Number {
    number(
        DecimalPlaces::fixed(places).expect("fixture precision is valid"),
        negative_style,
        separator,
    )
}

fn selected_position() -> CellPosition {
    CellPosition::new(fixture::FIRST_CELL.0 as u32, fixture::FIRST_CELL.1 as u32)
}

fn sibling_position() -> CellPosition {
    CellPosition::new(fixture::SECOND_CELL.0 as u32, fixture::SECOND_CELL.1 as u32)
}

fn read_format_list(source: &[u8]) -> TestResult<Vec<u8>> {
    let archive = fixture::member_archive(source, fixture::TABLES_MEMBER)?;
    let sidecars = archive
        .object(fixture::SIDECAR_ID)
        .ok_or_else(|| io::Error::other("format sidecar is missing"))?;
    sidecars
        .messages
        .iter()
        .find(|message| {
            message.type_ == fixture::TABLE_DATA_LIST_TYPE
                && tst::TableDataList::decode(message.data.as_slice())
                    .map(|list| list.list_type == tst::table_data_list::ListType::Format as i32)
                    .unwrap_or(false)
        })
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("format list message is missing").into())
}

fn format_entry_facts(source: &[u8]) -> TestResult<Vec<(u32, u32)>> {
    let list = tst::TableDataList::decode(read_format_list(source)?.as_slice())?;
    let mut facts = list
        .entries
        .into_iter()
        .map(|entry| (entry.key, entry.refcount))
        .collect::<Vec<_>>();
    facts.sort_unstable();
    Ok(facts)
}

fn format_next_list_id(source: &[u8]) -> TestResult<u32> {
    Ok(tst::TableDataList::decode(read_format_list(source)?.as_slice())?.next_list_id)
}

fn first_format_payload(source: &[u8]) -> TestResult<Vec<u8>> {
    let payload = read_format_list(source)?;
    let list = WireView::parse(&payload)?;
    let entry = list
        .fields()
        .find(|field| field.number() == 3)
        .ok_or_else(|| io::Error::other("format entry field is missing"))?;
    let entry = WireView::parse(entry.payload())?;
    entry
        .fields()
        .find(|field| field.number() == 6)
        .map(|field| field.payload().to_vec())
        .ok_or_else(|| io::Error::other("format payload field is missing").into())
}

fn unknown_field_record(source: &[u8], field_number: u32) -> TestResult<Vec<u8>> {
    WireView::parse(source)?
        .fields()
        .find(|field| field.number() == field_number)
        .map(|field| field.raw().to_vec())
        .ok_or_else(|| io::Error::other("unknown extension is missing").into())
}

fn rewrite_first_format_varint(
    source: &[u8],
    field_number: u32,
    value: u64,
) -> TestResult<Vec<u8>> {
    fixture::rewrite_tables(source, |archive| {
        let sidecars = archive
            .object_mut(fixture::SIDECAR_ID)
            .ok_or_else(|| io::Error::other("format sidecar is missing"))?;
        let message = sidecars
            .messages
            .iter_mut()
            .find(|message| {
                message.type_ == fixture::TABLE_DATA_LIST_TYPE
                    && tst::TableDataList::decode(message.data.as_slice())
                        .map(|list| list.list_type == tst::table_data_list::ListType::Format as i32)
                        .unwrap_or(false)
            })
            .ok_or_else(|| io::Error::other("format list message is missing"))?;
        message.data =
            rewrite_first_format_varint_payload(message.data.as_slice(), field_number, value)?;
        Ok(())
    })
}

fn rewrite_first_format_varint_payload(
    source: &[u8],
    field_number: u32,
    value: u64,
) -> TestResult<Vec<u8>> {
    let list = WireView::parse(source)?;
    let mut output = Vec::with_capacity(source.len());
    let mut replaced = false;
    for field in list.fields() {
        if field.number() != 3 || replaced {
            output.extend_from_slice(field.raw());
            continue;
        }
        let entry_view = WireView::parse(field.payload())?;
        let mut entry = Vec::with_capacity(field.payload().len());
        let mut format_replaced = false;
        for entry_field in entry_view.fields() {
            if entry_field.number() != 6 || format_replaced {
                entry.extend_from_slice(entry_field.raw());
                continue;
            }
            let format = patch_nested_varint_field(
                entry_field.payload(),
                &[field_number],
                true,
                Some(value),
            )?;
            append_length_delimited_field(&mut entry, 6, &format)?;
            format_replaced = true;
        }
        if !format_replaced {
            return Err(io::Error::other("format payload field is missing").into());
        }
        append_length_delimited_field(&mut output, 3, &entry)?;
        replaced = true;
    }
    if !replaced {
        return Err(io::Error::other("format entry field is missing").into());
    }
    Ok(output)
}

fn tile_cells(source: &[u8]) -> TestResult<Vec<Vec<u8>>> {
    let payload = fixture::object_message(
        source,
        fixture::TABLES_MEMBER,
        fixture::TILE_ID,
        fixture::TILE_TYPE,
    )?;
    let tile = tst::Tile::decode(payload.as_slice())?;
    let row = tile
        .row_infos
        .first()
        .ok_or_else(|| io::Error::other("format fixture row is missing"))?;
    let storage = row
        .cell_storage_buffer
        .as_deref()
        .ok_or_else(|| io::Error::other("format fixture storage is missing"))?;
    let count = usize::try_from(row.cell_count)
        .map_err(|_| io::Error::other("format fixture cell count overflows usize"))?;
    let offsets = row
        .cell_offsets
        .as_deref()
        .ok_or_else(|| io::Error::other("format fixture offsets are missing"))?;
    if offsets.len() != count.saturating_mul(2) {
        return Err(io::Error::other("format fixture offsets have an invalid length").into());
    }
    let mut cells = Vec::with_capacity(count);
    for index in 0..count {
        let begin = usize::from(u16::from_le_bytes([
            offsets[index * 2],
            offsets[index * 2 + 1],
        ]));
        let end = if index + 1 == count {
            storage.len()
        } else {
            usize::from(u16::from_le_bytes([
                offsets[(index + 1) * 2],
                offsets[(index + 1) * 2 + 1],
            ]))
        };
        if begin > end || end > storage.len() {
            return Err(io::Error::other("format fixture cell offset is out of range").into());
        }
        cells.push(storage[begin..end].to_vec());
    }
    Ok(cells)
}

fn format_keys(source: &[u8]) -> TestResult<Vec<Option<u32>>> {
    tile_cells(source)?
        .into_iter()
        .map(|cell| Ok(BncCell::parse(&cell)?.format_identifier()))
        .collect()
}

fn pack_cells(cells: &[Vec<u8>]) -> TestResult<(Vec<u8>, Vec<u8>)> {
    let mut storage = Vec::new();
    let mut offsets = Vec::with_capacity(cells.len().saturating_mul(2));
    for cell in cells {
        offsets.extend_from_slice(
            &u16::try_from(storage.len())
                .map_err(|_| io::Error::other("format fixture row exceeds narrow offsets"))?
                .to_le_bytes(),
        );
        storage.extend_from_slice(cell);
    }
    Ok((storage, offsets))
}

fn original_automatic_with_format_id(source: &[u8]) -> TestResult<Vec<u8>> {
    fixture::rewrite_tables(source, |archive| {
        let tile = archive
            .object_mut(fixture::TILE_ID)
            .ok_or_else(|| io::Error::other("format fixture tile is missing"))?;
        let message = tile
            .messages
            .first_mut()
            .ok_or_else(|| io::Error::other("format fixture tile payload is missing"))?;
        let mut decoded = tst::Tile::decode(message.data.as_slice())?;
        let row = decoded
            .row_infos
            .first_mut()
            .ok_or_else(|| io::Error::other("format fixture row is missing"))?;
        let storage = row
            .cell_storage_buffer
            .as_deref()
            .ok_or_else(|| io::Error::other("format fixture storage is missing"))?;
        let offsets = row
            .cell_offsets
            .as_deref()
            .ok_or_else(|| io::Error::other("format fixture offsets are missing"))?;
        let count = usize::try_from(row.cell_count)
            .map_err(|_| io::Error::other("format fixture cell count overflows usize"))?;
        if offsets.len() != count.saturating_mul(2) {
            return Err(io::Error::other("format fixture offsets have an invalid length").into());
        }
        let mut cells = Vec::with_capacity(count);
        for index in 0..count {
            let begin = usize::from(u16::from_le_bytes([
                offsets[index * 2],
                offsets[index * 2 + 1],
            ]));
            let end = if index + 1 == count {
                storage.len()
            } else {
                usize::from(u16::from_le_bytes([
                    offsets[(index + 1) * 2],
                    offsets[(index + 1) * 2 + 1],
                ]))
            };
            if begin > end || end > storage.len() {
                return Err(io::Error::other("format fixture cell offset is out of range").into());
            }
            cells.push(storage[begin..end].to_vec());
        }
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("format fixture first cell is missing"))?;
        let cell = BncCell::parse(first)?;
        let mut encoded = cell.encode();
        if encoded.len() < 8 {
            return Err(io::Error::other("format fixture cell prefix is truncated").into());
        }
        // Keep the native format kind and identifier while clearing only the
        // explicit-format flags.  This is the original automatic state that
        // Numbers emits for a format-ID-bearing cell.
        encoded[6..8].fill(0);
        *first = encoded;
        let (storage, offsets) = pack_cells(&cells)?;
        row.cell_storage_buffer = Some(storage.clone());
        row.cell_offsets = Some(offsets.clone());
        row.cell_storage_buffer_pre_bnc = storage;
        row.cell_offsets_pre_bnc = offsets;
        message.data = decoded.encode_to_vec();
        Ok(())
    })
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

fn assert_rejected_atomically(source: &[u8]) -> TestResult {
    let Ok(package) = Package::from_bytes(source) else {
        // Strict ingress is allowed to reject the malformed package before a
        // semantic owner is constructed.
        return Ok(());
    };
    assert_owner_rejects_atomically(&package)
}

fn assert_owner_rejects_atomically(package: &Package) -> TestResult {
    let before = package.exact_bytes();
    assert!(
        package
            .table_cell_number_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert!(
        package
            .edit_table_cell_number_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

fn assert_write_rejected_atomically(package: &Package) -> TestResult {
    let before = package.exact_bytes();
    assert!(
        package
            .edit_table_cell_number_format(0usize, 0usize, selected_position())?
            .set(fixed_number(
                4,
                NegativeStyle::MinusSign,
                ThousandsSeparator::Shown,
            ))
            .commit()
            .is_err()
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn transaction_types_are_strictly_typed_send_sync_and_redacted() -> TestResult {
    fn assert_send_sync_debug<T: Send + Sync + Debug>() {}

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<Number>();
    assert_send_sync_debug::<Edit<'static>>();
    assert_send_sync_debug::<Commit>();
    assert_send_sync_debug::<Patch>();
    assert_send_sync_debug::<Diagnostics>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<LimitKind>();
    assert_send_sync_debug::<Path>();

    let package =
        Package::from_bytes(&fixture::synthetic_package(fixture::FormatSharing::Shared)?)?;
    let edit = package.edit_table_cell_number_format(
        SheetSelector::name("Data Format Sheet"),
        TableSelector::name("Data Formats"),
        selected_position(),
    )?;
    let rendered = format!("{edit:?}");
    assert!(!rendered.contains("Index/"));
    assert!(!rendered.contains("Document.iwa"));
    assert!(!rendered.contains("data-format-table-id"));
    Ok(())
}

#[test]
fn selectors_and_option_number_semantics_are_explicit() -> TestResult {
    let source = fixture::synthetic_package(fixture::FormatSharing::Shared)?;
    let package = Package::from_bytes(&source)?;
    let expected = fixed_number(2, NegativeStyle::Parentheses, ThousandsSeparator::Shown);
    assert_eq!(
        package.table_cell_number_format(0usize, 0usize, selected_position())?,
        Some(expected)
    );
    assert_eq!(
        package.table_cell_number_format(
            SheetSelector::name("Data Format Sheet"),
            TableSelector::name("Data Formats"),
            selected_position(),
        )?,
        Some(expected)
    );
    assert!(
        package
            .table_cell_number_format("missing sheet", 0usize, selected_position())
            .is_err()
    );
    assert!(
        package
            .table_cell_number_format(0usize, "missing table", selected_position())
            .is_err()
    );

    let cleared = package
        .edit_table_cell_number_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .table_cell_number_format(0usize, 0usize, selected_position())?,
        None
    );
    let automatic = number(
        DecimalPlaces::Automatic,
        NegativeStyle::MinusSign,
        ThousandsSeparator::Hidden,
    );
    let explicit_automatic = cleared
        .package()
        .edit_table_cell_number_format(0usize, 0usize, selected_position())?
        .set(automatic)
        .commit()?;
    assert_eq!(
        explicit_automatic.package().table_cell_number_format(
            0usize,
            0usize,
            selected_position()
        )?,
        Some(automatic)
    );
    Ok(())
}

#[test]
fn every_decimal_negative_and_separator_combination_roundtrips() -> TestResult {
    let decimal_places = [
        DecimalPlaces::Automatic,
        DecimalPlaces::fixed(0)?,
        DecimalPlaces::fixed(3)?,
    ];
    let negative_styles = [
        NegativeStyle::MinusSign,
        NegativeStyle::Red,
        NegativeStyle::Parentheses,
        NegativeStyle::RedParentheses,
    ];
    let separators = [ThousandsSeparator::Hidden, ThousandsSeparator::Shown];
    for decimal_places in decimal_places {
        for negative_style in negative_styles {
            for separator in separators {
                let package = Package::from_bytes(&fixture::synthetic_package(
                    fixture::FormatSharing::Shared,
                )?)?;
                let expected = number(decimal_places, negative_style, separator);
                let commit = package
                    .edit_table_cell_number_format(0usize, 0usize, selected_position())?
                    .set(expected)
                    .commit()?;
                assert_eq!(
                    commit.package().table_cell_number_format(
                        SheetSelector::index(0),
                        TableSelector::index(0),
                        selected_position(),
                    )?,
                    Some(expected)
                );
                assert_eq!(
                    commit.package().table_cell_number_format(
                        SheetSelector::index(0),
                        TableSelector::index(0),
                        sibling_position(),
                    )?,
                    Some(fixed_number(
                        2,
                        NegativeStyle::Parentheses,
                        ThousandsSeparator::Shown,
                    ))
                );
            }
        }
    }
    Ok(())
}

#[test]
fn automatic_and_fixed_thirty_match_native_number_boundaries() -> TestResult {
    let source = fixture::synthetic_package(fixture::FormatSharing::Shared)?;
    let fixed_thirty = fixed_number(30, NegativeStyle::MinusSign, ThousandsSeparator::Hidden);
    let fixed = Package::from_bytes(&source)?
        .edit_table_cell_number_format(0usize, 0usize, selected_position())?
        .set(fixed_thirty)
        .commit()?;
    assert_eq!(
        fixed
            .package()
            .table_cell_number_format(0usize, 0usize, selected_position())?,
        Some(fixed_thirty)
    );

    let automatic_source = original_automatic_with_format_id(&source)?;
    let automatic_source = rewrite_first_format_varint(&automatic_source, 2, 253)?;
    let automatic_source = rewrite_first_format_varint(&automatic_source, 4, 0)?;
    let automatic_source = rewrite_first_format_varint(&automatic_source, 5, 0)?;
    let automatic_package = Package::from_bytes(&automatic_source)?;
    assert_eq!(
        format_keys(&automatic_source)?,
        vec![
            Some(fixture::FIRST_FORMAT_KEY),
            Some(fixture::FIRST_FORMAT_KEY)
        ]
    );
    let automatic_list =
        tst::TableDataList::decode(read_format_list(&automatic_source)?.as_slice())?;
    let automatic_format = automatic_list
        .entries
        .first()
        .and_then(|entry| entry.format.as_ref())
        .ok_or_else(|| io::Error::other("automatic format entry is missing"))?;
    assert_eq!(automatic_format.decimal_places, Some(253));
    assert_eq!(automatic_format.negative_style, Some(0));
    assert_eq!(automatic_format.show_thousands_separator, Some(false));
    assert_eq!(
        automatic_package.table_cell_number_format(0usize, 0usize, selected_position())?,
        None,
        "flags=0 with a format ID is native automatic/inherited state"
    );
    let explicit_automatic = number(
        DecimalPlaces::Automatic,
        NegativeStyle::MinusSign,
        ThousandsSeparator::Hidden,
    );
    let committed = automatic_package
        .edit_table_cell_number_format(0usize, 0usize, selected_position())?
        .set(explicit_automatic)
        .commit()?;
    assert_eq!(
        committed
            .package()
            .table_cell_number_format(0usize, 0usize, selected_position())?,
        Some(explicit_automatic)
    );
    Ok(())
}

#[test]
fn set_clear_reset_noop_inverse_and_locality_are_exact() -> TestResult {
    let source = fixture::synthetic_package(fixture::FormatSharing::Shared)?;
    let package = Package::from_bytes(&source)?;
    let original = package
        .table_cell_number_format(0usize, 0usize, selected_position())?
        .ok_or_else(|| io::Error::other("shared fixture number format is missing"))?;

    let no_op = package
        .edit_table_cell_number_format(0usize, 0usize, selected_position())?
        .set(original)
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().exact_bytes(), source);
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert_eq!(no_op.diagnostics().deleted_previews(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());

    let replacement = fixed_number(4, NegativeStyle::RedParentheses, ThousandsSeparator::Hidden);
    let changed = package
        .edit_table_cell_number_format(0usize, 0usize, selected_position())?
        .set(replacement)
        .commit()?;
    let target = changed.package().exact_bytes();
    assert!(!changed.patch().is_noop());
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().touched_components(), 1);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_eq!(
        changed
            .package()
            .table_cell_number_format(0usize, 0usize, selected_position())?,
        Some(replacement)
    );
    assert_exact_locality(&source, &target)?;

    let applied = package.apply_table_cell_number_format(changed.patch())?;
    assert_eq!(applied.package().exact_bytes(), target);
    let inverse = changed.patch().inverse();
    assert_eq!(inverse.inverse(), *changed.patch());
    let restored = Package::from_bytes(&target)?.apply_table_cell_number_format(&inverse)?;
    assert_eq!(restored.package().exact_bytes(), source);
    assert_eq!(
        restored
            .package()
            .table_cell_number_format(0usize, 0usize, selected_position())?,
        Some(original)
    );

    let cleared = changed
        .package()
        .edit_table_cell_number_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .table_cell_number_format(0usize, 0usize, selected_position())?,
        None
    );
    let reset = cleared
        .package()
        .edit_table_cell_number_format(0usize, 0usize, selected_position())?
        .reset()
        .commit()?;
    assert!(reset.patch().is_noop());
    Ok(())
}

#[test]
fn shared_format_copy_on_write_reuses_keys_and_removes_zero_refcounts() -> TestResult {
    let source = fixture::synthetic_package(fixture::FormatSharing::Shared)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        format_entry_facts(&source)?,
        vec![(fixture::FIRST_FORMAT_KEY, 2)]
    );
    assert_eq!(format_next_list_id(&source)?, 32);
    assert_eq!(
        format_keys(&source)?,
        vec![
            Some(fixture::FIRST_FORMAT_KEY),
            Some(fixture::FIRST_FORMAT_KEY)
        ]
    );

    let replacement = fixed_number(5, NegativeStyle::Red, ThousandsSeparator::Shown);
    let first = package
        .edit_table_cell_number_format(0usize, 0usize, selected_position())?
        .set(replacement)
        .commit()?;
    let first_bytes = first.package().exact_bytes();
    let first_keys = format_keys(&first_bytes)?;
    assert_eq!(first_keys.len(), 2);
    assert_ne!(first_keys[0], first_keys[1], "a shared entry must COW");
    assert_eq!(first_keys[1], Some(fixture::FIRST_FORMAT_KEY));
    let new_key = first_keys[0].ok_or_else(|| io::Error::other("COW format key is missing"))?;
    assert_eq!(format_entry_facts(&first_bytes)?.len(), 2);
    assert_eq!(
        format_entry_facts(&first_bytes)?,
        vec![(fixture::FIRST_FORMAT_KEY, 1), (new_key, 1)]
    );
    assert!(new_key < format_next_list_id(&first_bytes)?);
    assert_eq!(format_next_list_id(&first_bytes)?, 32);

    // A second cell asking for the exact same semantic payload must reuse the
    // existing COW entry instead of allocating another native key.
    let both = first
        .package()
        .edit_table_cell_number_format(0usize, 0usize, sibling_position())?
        .set(replacement)
        .commit()?;
    let both_bytes = both.package().exact_bytes();
    assert_eq!(
        format_keys(&both_bytes)?,
        vec![Some(new_key), Some(new_key)]
    );
    assert_eq!(format_entry_facts(&both_bytes)?, vec![(new_key, 2)]);
    assert_eq!(format_next_list_id(&both_bytes)?, 32);

    // Clearing decrements/removes the shared COW entry, and then removes the
    // original entry once its last native cell reference is cleared.
    let one_cleared = both
        .package()
        .edit_table_cell_number_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    let one_cleared_bytes = one_cleared.package().exact_bytes();
    assert_eq!(format_keys(&one_cleared_bytes)?, vec![None, Some(new_key)]);
    assert_eq!(format_entry_facts(&one_cleared_bytes)?, vec![(new_key, 1)]);

    let all_cleared = one_cleared
        .package()
        .edit_table_cell_number_format(0usize, 0usize, sibling_position())?
        .clear()
        .commit()?;
    assert_eq!(
        format_keys(&all_cleared.package().exact_bytes())?,
        vec![None, None]
    );
    assert!(format_entry_facts(&all_cleared.package().exact_bytes())?.is_empty());
    assert_eq!(
        format_next_list_id(&all_cleared.package().exact_bytes())?,
        32
    );
    Ok(())
}

#[test]
fn exact_source_apply_rejects_stale_and_foreign_patches_atomically() -> TestResult {
    let source = fixture::synthetic_package(fixture::FormatSharing::Shared)?;
    let package = Package::from_bytes(&source)?;
    let replacement = fixed_number(1, NegativeStyle::Red, ThousandsSeparator::Hidden);
    let commit = package
        .edit_table_cell_number_format(0usize, 0usize, selected_position())?
        .set(replacement)
        .commit()?;
    let target = commit.package().exact_bytes();

    let stale = Package::from_bytes(&target)?;
    let stale_before = stale.exact_bytes();
    assert!(
        stale
            .apply_table_cell_number_format(commit.patch())
            .is_err()
    );
    assert_eq!(stale.exact_bytes(), stale_before);

    let foreign = Package::from_bytes(&fixture::synthetic_package(
        fixture::FormatSharing::Unshared,
    )?)?;
    let foreign_before = foreign.exact_bytes();
    assert!(
        foreign
            .apply_table_cell_number_format(commit.patch())
            .is_err()
    );
    assert_eq!(foreign.exact_bytes(), foreign_before);
    Ok(())
}

#[test]
fn shared_and_wrong_data_format_cells_fail_with_typed_option_boundaries() -> TestResult {
    let package = Package::from_bytes(&fixture::synthetic_package(
        fixture::FormatSharing::Unshared,
    )?)?;
    assert!(
        package
            .table_cell_number_format(0usize, 0usize, sibling_position())
            .is_err()
    );
    assert!(
        package
            .edit_table_cell_number_format(0usize, 0usize, sibling_position())
            .is_err()
    );
    Ok(())
}

#[test]
fn malformed_wire_and_graph_inputs_fail_closed_and_remain_atomic() -> TestResult {
    for corruption in [
        fixture::Corruption::DuplicateFormatKey,
        fixture::Corruption::MissingFormatEntry,
        fixture::Corruption::RefcountMismatch,
        fixture::Corruption::DuplicateFormatList,
        fixture::Corruption::AliasedFormatList,
        fixture::Corruption::MalformedFormatPayload,
        fixture::Corruption::UnterminatedUnknownGroup,
    ] {
        assert_rejected_atomically(&fixture::corrupted_package(corruption)?)?;
    }

    for corruption in [
        fixture::Corruption::UnsupportedFormatType,
        fixture::Corruption::WrongCellFormatKey,
    ] {
        let source = fixture::corrupted_package(corruption)?;
        let package = Package::from_bytes(&source).map_err(|error| {
            io::Error::other(format!(
                "{corruption:?} should reach the Number-format owner: {error}"
            ))
        })?;
        assert_owner_rejects_atomically(&package)?;
    }

    let shared_tile = fixture::corrupted_package(fixture::Corruption::UnexpectedFieldReference)?;
    let package = Package::from_bytes(&shared_tile)?;
    assert!(
        package
            .table_cell_number_format(0usize, 0usize, selected_position())
            .is_ok()
    );
    assert_write_rejected_atomically(&package)?;

    // These malformed native values are valid package/IWA envelopes.  Require
    // the semantic owner to admit them so this test cannot pass by exercising
    // only the package parser's ingress rejection path.
    for (field, value) in [(2, 31), (2, 254), (4, 4), (5, 2)] {
        let source = rewrite_first_format_varint(
            &fixture::synthetic_package(fixture::FormatSharing::Shared)?,
            field,
            value,
        )?;
        let package = Package::from_bytes(&source).map_err(|error| {
            io::Error::other(format!(
                "semantic malformed format was rejected before owner admission: {field}={value}: {error}"
            ))
        })?;
        assert_owner_rejects_atomically(&package)?;
    }
    Ok(())
}

#[test]
fn unknown_table_data_list_wire_survives_an_admitted_rewrite() -> TestResult {
    let source = fixture::synthetic_package(fixture::FormatSharing::Shared)?;
    let package = Package::from_bytes(&source)?;
    let replacement = fixed_number(6, NegativeStyle::Parentheses, ThousandsSeparator::Shown);
    let target = package
        .edit_table_cell_number_format(0usize, 0usize, selected_position())?
        .set(replacement)
        .commit()?
        .package()
        .exact_bytes();
    let payload = read_format_list(&target)?;
    let view = WireView::parse(&payload)?;
    assert!(view.fields().any(|field| field.number() == 90));
    assert!(view.fields().any(|field| field.number() == 94));
    assert_exact_locality(&source, &target)?;
    Ok(())
}

#[test]
fn unknown_nested_format_extension_is_copied_byte_for_byte() -> TestResult {
    let source = fixture::synthetic_package(fixture::FormatSharing::Shared)?;
    let original_extension = unknown_field_record(&first_format_payload(&source)?, 94)?;
    let replacement = fixed_number(7, NegativeStyle::Parentheses, ThousandsSeparator::Hidden);
    let target = Package::from_bytes(&source)?
        .edit_table_cell_number_format(0usize, 0usize, selected_position())?
        .set(replacement)
        .commit()?
        .package()
        .exact_bytes();
    let retained_extension = unknown_field_record(&first_format_payload(&target)?, 94)?;
    assert_eq!(retained_extension, original_extension);
    assert_exact_locality(&source, &target)?;
    Ok(())
}

#[test]
fn concurrent_arc_reads_and_edits_are_independent_and_send_sync() -> TestResult {
    let package = Arc::new(Package::from_bytes(&fixture::synthetic_package(
        fixture::FormatSharing::Shared,
    )?)?);
    let expected = fixed_number(8, NegativeStyle::RedParentheses, ThousandsSeparator::Shown);
    let handles = (0..8)
        .map(|_| {
            let package = Arc::clone(&package);
            thread::spawn(move || -> Result<(Option<Number>, Option<Number>), Error> {
                let observed = package.table_cell_number_format(
                    SheetSelector::index(0),
                    TableSelector::index(0),
                    selected_position(),
                )?;
                let commit = package
                    .edit_table_cell_number_format(
                        SheetSelector::index(0),
                        TableSelector::index(0),
                        selected_position(),
                    )?
                    .set(expected)
                    .commit()?;
                let after = commit.package().table_cell_number_format(
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
            .map_err(|_| io::Error::other("concurrent Number-format worker panicked"))??;
        assert_eq!(
            observed,
            Some(fixed_number(
                2,
                NegativeStyle::Parentheses,
                ThousandsSeparator::Shown
            ))
        );
        assert_eq!(after, Some(expected));
    }
    // The source owner is immutable; concurrent edits publish independent
    // snapshots and therefore cannot mutate the shared input behind the Arc.
    assert_eq!(
        package.table_cell_number_format(0usize, 0usize, selected_position())?,
        Some(fixed_number(
            2,
            NegativeStyle::Parentheses,
            ThousandsSeparator::Shown
        ))
    );
    Ok(())
}

#[test]
fn public_input_and_semantic_budgets_reject_before_publication() -> TestResult {
    let source = fixture::synthetic_package(fixture::FormatSharing::Shared)?;
    let tight = PackageLimits::new(
        source.len().saturating_sub(1) as u64,
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
    Ok(())
}

#[test]
fn native_fixture_read_is_selector_equivalent() -> TestResult {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/basic.numbers");
    let package = Package::open(path)?;
    let position = CellPosition::new(2, 1);
    let by_index = package.table_cell_number_format(0usize, 0usize, position)?;
    let by_name = package.table_cell_number_format(
        SheetSelector::name("Sheet 1"),
        TableSelector::name("Table 1"),
        position,
    )?;
    assert_eq!(by_name, by_index);
    Ok(())
}
