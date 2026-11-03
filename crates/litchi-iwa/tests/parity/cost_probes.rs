//! Callgrind ABBA cost probes: each migrated route against its host route
//! (record 0363 §5, record 0370).
//!
//! Each probe runs the host route and the leaf route of one operation in
//! ABBA order on the same fixture bytes: open, edit, commit and serialize.
//! The probes are ignored tests; measure one route at a time from a release
//! build with debug information:
//!
//! ```sh
//! CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-iwa --test parity --no-run
//! taskset -c 3 valgrind --tool=callgrind '--toggle-collect=*cost_probes::host_topology*' \
//!   target/release/deps/parity-<hash> --ignored --exact \
//!   cost_probes::numbers_topology_insert_cost_abba --test-threads=1
//! # likewise with '--toggle-collect=*cost_probes::leaf_topology*'
//! ```

use std::hint::black_box;

use litchi_iwa::numbers;
use litchi_iwa_archive::iwa::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_archive::{
    Limits as ArchiveLimits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::wire::append_varint_field;
use litchi_iwa_common::{WireLimits, wire::WireView};
use litchi_iwa_protos::{tn, tsp, tss, tst, tswp};
use litchi_numbers::table::topology::{ColumnInsertion, RowInsertion};
use prost::Message as _;

use crate::harness::{TestResult, fixture};
use crate::leaf_bytes;

#[derive(Clone, Copy)]
enum Insert {
    Row(RowInsertion),
    Column(ColumnInsertion),
}

#[inline(never)]
fn host_topology(source: &[u8], op: Insert) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let table = litchi_numbers::TableSelector::index(0);
    match op {
        Insert::Row(at) => editor.insert_table_row(table, at)?,
        Insert::Column(at) => editor.insert_table_column(table, at)?,
    }
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_topology(source: &[u8], op: Insert) -> TestResult<Vec<u8>> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let edit = package.edit_table_topology(0usize, 0usize)?;
    let edit = match op {
        Insert::Row(at) => edit.insert_row(at)?,
        Insert::Column(at) => edit.insert_column(at)?,
    };
    Ok(leaf_bytes!(edit.commit()?.package()))
}

/// Row and column insertion on `basic.numbers` and a native fixture.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_topology_insert_cost_abba() -> TestResult {
    let inputs = [
        fixture("iwork/numbers/basic.numbers")?,
        fixture("iwork/numbers/custom-number-native.numbers")?,
    ];
    let ops = [
        Insert::Row(RowInsertion::body(1)),
        Insert::Column(ColumnInsertion::body(0)),
    ];
    for input in &inputs {
        for op in ops {
            black_box(host_topology(&input.bytes, op)?);
            black_box(leaf_topology(&input.bytes, op)?);
            black_box(leaf_topology(&input.bytes, op)?);
            black_box(host_topology(&input.bytes, op)?);
        }
    }
    Ok(())
}

#[inline(never)]
fn host_text_style(source: &[u8], cells: &[(u32, u32)]) -> TestResult<Vec<u8>> {
    use litchi_iwa::text::{TextPointSize, TextStyle};
    let mut editor = numbers::Draft::from_bytes(source)?;
    let id = editor
        .package()
        .iwa_entry_names()
        .find_map(|name| {
            editor
                .package()
                .archive(name)
                .ok()?
                .objects
                .into_iter()
                .find_map(|object| {
                    object
                        .messages
                        .iter()
                        .any(|message| message.type_ == 6_001)
                        .then_some(object.archive_info.identifier)
                        .flatten()
                })
        })
        .ok_or("no table model")?;
    let style = TextStyle::new(TextPointSize::from_points(18.0)?).with_bold(true);
    for (row, column) in cells {
        editor.set_table_cell_text_style(id, *row as usize, *column as usize, style)?;
    }
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_text_style(source: &[u8], cells: &[(u32, u32)]) -> TestResult<Vec<u8>> {
    use litchi_numbers::cell::text_style::CharacterProperties;
    let package = litchi_numbers::Package::from_bytes(source)?;
    let mut properties = CharacterProperties::new();
    properties.size = Some(18.0);
    properties.bold = Some(true);
    let mut edit = package.edit_table_cell_text_style(0usize, 0usize)?;
    for (row, column) in cells {
        edit = edit.set(
            litchi_numbers::table::At::new(*row, *column),
            properties.clone(),
        )?;
    }
    Ok(leaf_bytes!(edit.commit()?.package()))
}

/// Whole-cell text style on `basic.numbers`: one cell, then two.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_text_style_cost_abba() -> TestResult {
    let input = fixture("iwork/numbers/basic.numbers")?;
    for cells in [&[(1, 1)][..], &[(1, 1), (2, 1)][..]] {
        black_box(host_text_style(&input.bytes, cells)?);
        black_box(leaf_text_style(&input.bytes, cells)?);
        black_box(leaf_text_style(&input.bytes, cells)?);
        black_box(host_text_style(&input.bytes, cells)?);
    }
    Ok(())
}

fn whole_cell_style_properties(
    size: f32,
    bold: bool,
    italic: bool,
) -> litchi_numbers::cell::text_style::CharacterProperties {
    let mut properties = litchi_numbers::cell::text_style::CharacterProperties::new();
    properties.size = Some(size);
    properties.bold = Some(bold);
    properties.italic = Some(italic);
    properties
}

fn prepare_w1g_host_whole_cell_style_single(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let model = crate::numbers_cases::host_first_table_model_id(&editor)?;
    let style =
        litchi_iwa::text::TextStyle::new(litchi_iwa::text::TextPointSize::from_points(18.0)?)
            .with_bold(true);
    editor.set_table_cell_text_style(model, 1, 1, style)?;
    Ok(editor.to_bytes()?)
}

fn host_w1g_whole_cell_style_at(source: &[u8], row: u32, column: u32) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let model = crate::numbers_cases::host_first_table_model_id(&editor)?;
    let style =
        litchi_iwa::text::TextStyle::new(litchi_iwa::text::TextPointSize::from_points(18.0)?)
            .with_bold(true);
    editor.set_table_cell_text_style(model, row as usize, column as usize, style)?;
    Ok(editor.to_bytes()?)
}

fn leaf_w1g_whole_cell_style_at(
    source: &[u8],
    row: u32,
    column: u32,
) -> TestResult<(Vec<u8>, Vec<u8>)> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let properties = whole_cell_style_properties(18.0, true, false);
    let commit = package
        .edit_table_cell_text_style(0usize, 0usize)?
        .set(litchi_numbers::table::At::new(row, column), properties)?
        .commit()?;
    let output = leaf_bytes!(commit.package());
    let restored = commit
        .package()
        .apply_table_cell_text_style(&commit.patch().inverse())?;
    Ok((output, leaf_bytes!(restored.package())))
}

fn prepare_w1g_leaf_whole_cell_style_single(source: &[u8]) -> TestResult<Vec<u8>> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let properties = whole_cell_style_properties(18.0, true, false);
    let commit = package
        .edit_table_cell_text_style(0usize, 0usize)?
        .set(litchi_numbers::table::At::new(1, 1), properties)?
        .commit()?;
    Ok(leaf_bytes!(commit.package()))
}

fn prepare_w1g_host_whole_cell_style_exclusive_repeat(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let model = crate::numbers_cases::host_first_table_model_id(&editor)?;
    let style =
        litchi_iwa::text::TextStyle::new(litchi_iwa::text::TextPointSize::from_points(21.0)?)
            .with_italic(true);
    editor.set_table_cell_text_style(model, 1, 1, style)?;
    Ok(editor.to_bytes()?)
}

fn prepare_w1g_leaf_whole_cell_style_exclusive_repeat(source: &[u8]) -> TestResult<Vec<u8>> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let properties = whole_cell_style_properties(21.0, false, true);
    let commit = package
        .edit_table_cell_text_style(0usize, 0usize)?
        .set(litchi_numbers::table::At::new(1, 1), properties)?
        .commit()?;
    Ok(leaf_bytes!(commit.package()))
}

fn prepare_w1g_host_whole_cell_style_batch_1000(
    source: &[u8],
    cells: &[(u32, u32)],
) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let model = crate::numbers_cases::host_first_table_model_id(&editor)?;
    let style =
        litchi_iwa::text::TextStyle::new(litchi_iwa::text::TextPointSize::from_points(19.0)?)
            .with_bold(true);
    for &(row, column) in cells {
        editor.set_table_cell_text_style(model, row as usize, column as usize, style)?;
    }
    Ok(editor.to_bytes()?)
}

fn prepare_w1g_leaf_whole_cell_style_batch_1000(
    source: &[u8],
    cells: &[(u32, u32)],
) -> TestResult<Vec<u8>> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let properties = whole_cell_style_properties(19.0, true, false);
    let mut edit = package.edit_table_cell_text_style(0usize, 0usize)?;
    for &(row, column) in cells {
        edit = edit.set(
            litchi_numbers::table::At::new(row, column),
            properties.clone(),
        )?;
    }
    Ok(leaf_bytes!(edit.commit()?.package()))
}

fn validate_w1g_leaf_whole_cell_style_batch_1000(
    source: &[u8],
    cells: &[(u32, u32)],
) -> TestResult<Vec<u8>> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let properties = whole_cell_style_properties(19.0, true, false);
    let mut edit = package.edit_table_cell_text_style(0usize, 0usize)?;
    for &(row, column) in cells {
        edit = edit.set(
            litchi_numbers::table::At::new(row, column),
            properties.clone(),
        )?;
    }
    let commit = edit.commit()?;
    let output = leaf_bytes!(commit.package());
    let restored = commit
        .package()
        .apply_table_cell_text_style(&commit.patch().inverse())?;
    assert_eq!(leaf_bytes!(restored.package()), source);
    Ok(output)
}

#[inline(never)]
fn host_whole_cell_style_single(source: &[u8]) -> TestResult<Vec<u8>> {
    prepare_w1g_host_whole_cell_style_single(source)
}

#[inline(never)]
fn leaf_whole_cell_style_single(source: &[u8]) -> TestResult<Vec<u8>> {
    prepare_w1g_leaf_whole_cell_style_single(source)
}

#[inline(never)]
fn host_whole_cell_style_exclusive_repeat(source: &[u8]) -> TestResult<Vec<u8>> {
    prepare_w1g_host_whole_cell_style_exclusive_repeat(source)
}

#[inline(never)]
fn leaf_whole_cell_style_exclusive_repeat(source: &[u8]) -> TestResult<Vec<u8>> {
    prepare_w1g_leaf_whole_cell_style_exclusive_repeat(source)
}

#[inline(never)]
fn host_whole_cell_style_batch_1000(source: &[u8], cells: &[(u32, u32)]) -> TestResult<Vec<u8>> {
    prepare_w1g_host_whole_cell_style_batch_1000(source, cells)
}

#[inline(never)]
fn leaf_whole_cell_style_batch_1000(source: &[u8], cells: &[(u32, u32)]) -> TestResult<Vec<u8>> {
    prepare_w1g_leaf_whole_cell_style_batch_1000(source, cells)
}

fn assert_whole_cell_style_output(
    host_output: &[u8],
    leaf_output: &[u8],
    cells: &[(u32, u32)],
    size: f32,
    bold: bool,
    italic: bool,
) -> TestResult {
    let host = litchi_numbers::Package::from_bytes(host_output)?;
    let leaf = litchi_numbers::Package::from_bytes(leaf_output)?;
    let reopened = litchi_numbers::Package::from_bytes(leaf_output)?;
    for &(row, column) in cells {
        let host_properties = host.table_cell_text_style(
            0usize,
            0usize,
            litchi_numbers::table::At::new(row, column),
        )?;
        for candidate in [&leaf, &reopened] {
            let properties = candidate.table_cell_text_style(
                0usize,
                0usize,
                litchi_numbers::table::At::new(row, column),
            )?;
            assert_eq!(properties, host_properties);
            assert_eq!(properties.size, Some(size));
            if bold {
                assert_eq!(properties.bold, Some(true));
            }
            if italic {
                assert_eq!(properties.italic, Some(true));
            }
        }
    }
    Ok(())
}

fn assert_untouched_member_preserved(
    source: &[u8],
    output: &[u8],
    excluded: &[&str],
) -> TestResult {
    let source = Catalog::from_bytes(source, ArchiveLimits::default())?;
    let output = Catalog::from_bytes(output, ArchiveLimits::default())?;
    for entry in source.iter() {
        if excluded.contains(&entry.name()) {
            continue;
        }
        let other = output
            .iter()
            .find(|candidate| candidate.name() == entry.name())
            .ok_or("untouched member disappeared")?;
        assert_eq!(
            other.data(),
            entry.data(),
            "member changed: {}",
            entry.name()
        );
    }
    Ok(())
}

fn table_model(source: &[u8]) -> TestResult<(u64, tst::TableModelArchive)> {
    let editor = numbers::Draft::from_bytes(source)?;
    for name in editor.package().iwa_entry_names() {
        let archive = editor.package().archive(name)?;
        for object in archive.objects {
            let Some(message) = object
                .messages
                .iter()
                .find(|message| message.type_ == 6_001)
            else {
                continue;
            };
            let identifier = object
                .archive_info
                .identifier
                .ok_or("table model has no id")?;
            return Ok((
                identifier,
                tst::TableModelArchive::decode(message.data.as_slice())?,
            ));
        }
    }
    Err("table model missing".into())
}

fn object_message(source: &[u8], identifier: u64, message_type: u32) -> TestResult<Vec<u8>> {
    let editor = numbers::Draft::from_bytes(source)?;
    for name in editor.package().iwa_entry_names() {
        let archive = editor.package().archive(name)?;
        let Some(object) = archive
            .objects
            .into_iter()
            .find(|object| object.archive_info.identifier == Some(identifier))
        else {
            continue;
        };
        let message = object
            .messages
            .into_iter()
            .find(|message| message.type_ == message_type)
            .ok_or("requested object message missing")?;
        return Ok(message.data);
    }
    Err("requested object missing".into())
}

fn first_missing_row_existing_tile(source: &[u8]) -> TestResult<(u32, u32)> {
    let (_, model) = table_model(source)?;
    let rows = model.number_of_rows;
    let tile_size = model
        .base_data_store
        .tiles
        .tile_size
        .ok_or("tile size missing")?;
    for tile_ref in model.base_data_store.tiles.tiles {
        let tile_key = tile_ref.tileid;
        let tile =
            tst::Tile::decode(object_message(source, tile_ref.tile.identifier, 6_002)?.as_slice())?;
        for local in 0..tile_size {
            let Some(row) = tile_key
                .checked_mul(tile_size)
                .and_then(|base| base.checked_add(local))
            else {
                continue;
            };
            if row >= rows {
                break;
            }
            if tile
                .row_infos
                .iter()
                .all(|candidate| candidate.tile_row_index != local)
            {
                return Ok((row, 0));
            }
        }
    }
    Err("source has no missing row in an existing tile".into())
}

#[derive(Debug, PartialEq, Eq)]
struct HeaderAndTileProbe {
    row_count: Option<u32>,
    column_count: Option<u32>,
    tile_num_cells: u32,
    tile_numrows: u32,
    target_row_present: bool,
    target_cell_present: bool,
    row_unknown_fields: Vec<Vec<u8>>,
    column_unknown_fields: Vec<Vec<u8>>,
}

fn header_unknown_fields(payload: &[u8], index: u32) -> TestResult<Vec<Vec<u8>>> {
    let bucket = WireView::parse(payload, WireLimits::DEFAULT)?;
    let mut unknown = bucket
        .fields()
        .filter(|field| !matches!(field.number(), 1 | 2))
        .map(|field| field.raw().to_vec())
        .collect::<Vec<_>>();
    for field in bucket.fields().filter(|field| field.number() == 2) {
        let header = WireView::parse(field.payload(), WireLimits::DEFAULT)?;
        let header_index = header
            .fields()
            .find(|candidate| candidate.number() == 1)
            .map(|candidate| litchi_iwa_common::decode_varint_from_bytes(candidate.payload()))
            .transpose()?
            .map(|(value, _)| value);
        if header_index == Some(u64::from(index)) {
            unknown.extend(
                header
                    .fields()
                    .filter(|candidate| candidate.number() > 6)
                    .map(|candidate| candidate.raw().to_vec())
                    .collect::<Vec<_>>(),
            );
            break;
        }
    }
    Ok(unknown)
}

fn add_header_unknowns(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source, ArchiveLimits::default())?;
    let mut replacements = Vec::new();
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let bytes = SnappyStream::decompress(entry.data(), Default::default())?;
        let mut archive = Archive::parse(bytes.as_bytes(), Default::default())?;
        let mut changed = false;
        for object in &mut archive.objects {
            for message in &mut object.messages {
                if message.type_ != 6_006 {
                    continue;
                }
                append_varint_field(&mut message.data, 90, 0x51)?;
                changed = true;
            }
        }
        if changed {
            let compressed =
                SnappyStream::compress(&archive.to_bytes(Default::default())?, Default::default())?;
            replacements.push((entry.name().to_owned(), compressed));
        }
    }
    if replacements.is_empty() {
        return Err("source has no header buckets to decorate".into());
    }
    let edits = replacements
        .iter()
        .map(|(name, bytes)| EntryEdit::new(name, bytes))
        .collect::<Vec<_>>();
    Ok(catalog.reassemble_to_bytes(&edits, ArchiveLimits::default())?)
}

fn header_and_tile_probe(source: &[u8], row: u32, column: u32) -> TestResult<HeaderAndTileProbe> {
    let (_, model) = table_model(source)?;
    let tile_size = model
        .base_data_store
        .tiles
        .tile_size
        .ok_or("tile size missing")?;
    let tile_key = row / tile_size;
    let tile_ref = model
        .base_data_store
        .tiles
        .tiles
        .iter()
        .find(|tile| tile.tileid == tile_key)
        .ok_or("selected tile reference missing")?;
    let tile_payload = object_message(source, tile_ref.tile.identifier, 6_002)?;
    let tile = tst::Tile::decode(tile_payload.as_slice())?;
    let local = row % tile_size;
    let target_row = tile
        .row_infos
        .iter()
        .find(|candidate| candidate.tile_row_index == local);
    let target_cell_present = target_row.is_some_and(|candidate| {
        candidate
            .cell_offsets
            .as_deref()
            .and_then(|offsets| offsets.get(column as usize * 2..column as usize * 2 + 2))
            .is_some_and(|offset| offset != [0xff, 0xff])
    });
    let row_bucket_reference = model
        .base_data_store
        .row_headers
        .buckets
        .get(usize::try_from(
            u64::from(row) / litchi_numbers_wire::table_topology::storage::HEADER_BUCKET_ROWS,
        )?)
        .ok_or("selected row bucket reference missing")?;
    let column_bucket_reference = &model.base_data_store.column_headers;
    let row_bucket = object_message(source, row_bucket_reference.identifier, 6_006)?;
    let column_bucket = object_message(source, column_bucket_reference.identifier, 6_006)?;
    let row_header = tst::HeaderStorageBucket::decode(row_bucket.as_slice())?
        .headers
        .into_iter()
        .find(|header| header.index == row)
        .map(|header| header.number_of_cells);
    let column_header = tst::HeaderStorageBucket::decode(column_bucket.as_slice())?
        .headers
        .into_iter()
        .find(|header| header.index == column)
        .map(|header| header.number_of_cells);
    Ok(HeaderAndTileProbe {
        row_count: row_header,
        column_count: column_header,
        tile_num_cells: tile.num_cells,
        tile_numrows: tile.numrows,
        target_row_present: target_row.is_some(),
        target_cell_present,
        row_unknown_fields: header_unknown_fields(&row_bucket, row)?,
        column_unknown_fields: header_unknown_fields(&column_bucket, column)?,
    })
}

fn assert_w1g_materialization_case(source: &[u8], row: u32, column: u32) -> TestResult {
    let source_probe = header_and_tile_probe(source, row, column)?;
    assert!(!source_probe.target_cell_present);
    let host = host_w1g_whole_cell_style_at(source, row, column)?;
    let (leaf, restored) = leaf_w1g_whole_cell_style_at(source, row, column)?;
    assert_eq!(
        restored, source,
        "leaf inverse must restore the exact source"
    );
    let host_probe = header_and_tile_probe(&host, row, column)?;
    let leaf_probe = header_and_tile_probe(&leaf, row, column)?;
    assert_eq!(
        host_probe, leaf_probe,
        "host/leaf storage materialization differs"
    );
    assert!(host_probe.target_row_present);
    assert!(host_probe.target_cell_present);
    assert!(host_probe.row_count.is_some_and(|count| count > 0));
    assert!(host_probe.column_count.is_some_and(|count| count > 0));
    assert!(host_probe.tile_numrows >= source_probe.tile_numrows);
    for output in [&host, &leaf] {
        let editor = numbers::Draft::from_bytes(output)?;
        let model = crate::numbers_cases::host_first_table_model_id(&editor)?;
        let style = editor.table_cell_text_style(model, row as usize, column as usize)?;
        assert_eq!(style.point_size.points(), 18.0);
        assert!(style.bold);
        let package = litchi_numbers::Package::from_bytes(output)?;
        let properties = package.table_cell_text_style(
            0usize,
            0usize,
            litchi_numbers::table::At::new(row, column),
        )?;
        assert_eq!(properties.size, Some(18.0));
        assert_eq!(properties.bold, Some(true));
    }
    assert_eq!(
        host_probe.row_unknown_fields,
        source_probe.row_unknown_fields
    );
    assert_eq!(
        host_probe.column_unknown_fields,
        source_probe.column_unknown_fields
    );
    Ok(())
}

/// Source-relative host/leaf controls for style-only materialization.  G22
/// exercises an absent cell in an existing row; the second target is an
/// absent row inside an existing tile.  Both controls compare row/column
/// buckets, tile presence/counts, semantic host/leaf reads, unknown fields,
/// and the exact leaf inverse.
#[test]
fn numbers_whole_cell_style_materialization_header_controls() -> TestResult {
    let input = fixture("iwork/numbers/basic.numbers")?;
    let source = add_header_unknowns(&input.bytes)?;
    let source_probe = header_and_tile_probe(&source, 21, 6)?;
    assert!(!source_probe.row_unknown_fields.is_empty());
    assert!(!source_probe.column_unknown_fields.is_empty());
    assert_w1g_materialization_case(&source, 21, 6)?;
    let missing = first_missing_row_existing_tile(&source)?;
    assert_w1g_materialization_case(&source, missing.0, missing.1)
}

fn assert_batch_style_identity(output: &[u8]) -> TestResult {
    let catalog = Catalog::from_bytes(output, ArchiveLimits::default())?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == "Index/Tables/DataList-904496-2.iwa")
        .ok_or("style-list member missing")?;
    let bytes = SnappyStream::decompress(entry.data(), Default::default())?;
    let archive = Archive::parse(bytes.as_bytes(), Default::default())?;
    let message = archive
        .objects
        .iter()
        .flat_map(|object| object.messages.iter())
        .find(|message| message.type_ == 6_005)
        .ok_or("style-list payload missing")?;
    let list = tst::TableDataList::decode(message.data.as_slice())?;
    assert!(!list.entries.iter().any(|entry| entry.key == 1));
    let fresh = list
        .entries
        .iter()
        .filter(|entry| entry.key >= 2 && entry.refcount == 1)
        .count();
    assert_eq!(fresh, 1_000);
    Ok(())
}

fn leaf_whole_cell_style_inverse(
    source: &[u8],
    cells: &[(u32, u32)],
    properties: litchi_numbers::cell::text_style::CharacterProperties,
) -> TestResult<Vec<u8>> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let mut edit = package.edit_table_cell_text_style(0usize, 0usize)?;
    for &(row, column) in cells {
        edit = edit.set(
            litchi_numbers::table::At::new(row, column),
            properties.clone(),
        )?;
    }
    let commit = edit.commit()?;
    let output = leaf_bytes!(commit.package());
    let restored = commit
        .package()
        .apply_table_cell_text_style(&commit.patch().inverse())?;
    assert_eq!(leaf_bytes!(restored.package()), source);
    Ok(output)
}

/// Current-source single whole-cell CharacterProperties cost route.
#[test]
#[ignore = "callgrind cost probe; run explicitly after coordinator acceptance"]
fn numbers_whole_cell_style_single_cost_abba() -> TestResult {
    let input = fixture("iwork/numbers/basic.numbers")?;
    let cells = [(1, 1)];
    let host_setup = prepare_w1g_host_whole_cell_style_single(&input.bytes)?;
    let leaf_setup = prepare_w1g_leaf_whole_cell_style_single(&input.bytes)?;
    assert_whole_cell_style_output(&host_setup, &leaf_setup, &cells, 18.0, true, false)?;
    let _ = leaf_whole_cell_style_inverse(
        &input.bytes,
        &cells,
        whole_cell_style_properties(18.0, true, false),
    )?;
    black_box(host_whole_cell_style_single(&input.bytes)?);
    black_box(leaf_whole_cell_style_single(&input.bytes)?);
    black_box(leaf_whole_cell_style_single(&input.bytes)?);
    black_box(host_whole_cell_style_single(&input.bytes)?);
    Ok(())
}

/// Current-source repeated exclusive single-cell CharacterProperties route.
#[test]
#[ignore = "callgrind cost probe; run explicitly after coordinator acceptance"]
fn numbers_whole_cell_style_exclusive_repeat_cost_abba() -> TestResult {
    let input = fixture("iwork/numbers/basic.numbers")?;
    let cells = [(1, 1)];
    let seeded = prepare_w1g_host_whole_cell_style_single(&input.bytes)?;
    let host_setup = prepare_w1g_host_whole_cell_style_exclusive_repeat(&seeded)?;
    let leaf_setup = prepare_w1g_leaf_whole_cell_style_exclusive_repeat(&seeded)?;
    assert_whole_cell_style_output(&host_setup, &leaf_setup, &cells, 21.0, false, true)?;
    let _ = leaf_whole_cell_style_inverse(
        &seeded,
        &cells,
        whole_cell_style_properties(21.0, false, true),
    )?;
    black_box(host_whole_cell_style_exclusive_repeat(&seeded)?);
    black_box(leaf_whole_cell_style_exclusive_repeat(&seeded)?);
    black_box(leaf_whole_cell_style_exclusive_repeat(&seeded)?);
    black_box(host_whole_cell_style_exclusive_repeat(&seeded)?);
    Ok(())
}

/// Current-source 1000-cell shared-key whole-cell CharacterProperties route.
#[test]
#[ignore = "callgrind cost probe; run explicitly after coordinator acceptance"]
fn numbers_whole_cell_style_batch_1000_cost_abba() -> TestResult {
    let fixture = shared_1000_cell_style_fixture()?;
    assert_eq!(fixture.cells.len(), 1_000);
    let host_setup = prepare_w1g_host_whole_cell_style_batch_1000(&fixture.bytes, &fixture.cells)?;
    let leaf_setup = validate_w1g_leaf_whole_cell_style_batch_1000(&fixture.bytes, &fixture.cells)?;
    assert_whole_cell_style_output(&host_setup, &leaf_setup, &fixture.cells, 19.0, true, false)?;
    assert_untouched_member_preserved(
        &fixture.bytes,
        &host_setup,
        &[
            "Index/Tables/Tile.iwa",
            "Index/Tables/DataList-904496-2.iwa",
            "Index/CalculationEngine.iwa",
            "Index/DocumentStylesheet.iwa",
            "Index/Metadata.iwa",
        ],
    )?;
    assert_untouched_member_preserved(
        &fixture.bytes,
        &leaf_setup,
        &[
            "Index/Tables/Tile.iwa",
            "Index/Tables/DataList-904496-2.iwa",
            "Index/CalculationEngine.iwa",
            "Index/DocumentStylesheet.iwa",
            "Index/Metadata.iwa",
        ],
    )?;
    assert_batch_style_identity(&host_setup)?;
    assert_batch_style_identity(&leaf_setup)?;
    black_box(host_whole_cell_style_batch_1000(
        &fixture.bytes,
        &fixture.cells,
    )?);
    black_box(leaf_whole_cell_style_batch_1000(
        &fixture.bytes,
        &fixture.cells,
    )?);
    black_box(leaf_whole_cell_style_batch_1000(
        &fixture.bytes,
        &fixture.cells,
    )?);
    black_box(host_whole_cell_style_batch_1000(
        &fixture.bytes,
        &fixture.cells,
    )?);
    Ok(())
}

#[derive(Clone)]
struct WholeCellFixture {
    bytes: Vec<u8>,
    cells: Vec<(u32, u32)>,
}

#[allow(
    dead_code,
    reason = "kept as a bounded source-fixture construction reference"
)]
fn whole_cell_reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

#[allow(
    dead_code,
    reason = "kept as a bounded source-fixture construction reference"
)]
fn whole_cell_object(
    identifier: u64,
    message_type: u32,
    data: Vec<u8>,
) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data,
        }],
        Default::default(),
    )?)
}

#[allow(
    dead_code,
    reason = "kept as a bounded source-fixture construction reference"
)]
fn minimal_1000_cell_style_fixture() -> TestResult<WholeCellFixture> {
    const DOCUMENT: u64 = 9_000_000_001;
    const SHEET: u64 = 9_000_000_002;
    const TABLE_INFO: u64 = 9_000_000_003;
    const TABLE_MODEL: u64 = 9_000_000_004;
    const STYLE_LIST: u64 = 9_000_000_005;
    const TILE: u64 = 9_000_000_006;
    const STYLE: u64 = 9_000_000_007;
    const STRING_LIST: u64 = 9_000_000_008;
    const FORMULA_LIST: u64 = 9_000_000_009;
    const FORMAT_LIST: u64 = 9_000_000_010;
    const FORMULA_ERROR_LIST: u64 = 9_000_000_011;
    const RICH_TEXT_LIST: u64 = 9_000_000_012;
    const STYLESHEET: u64 = 9_000_000_013;
    const DOCUMENT_TYPE: u32 = 1;
    const SHEET_TYPE: u32 = 2;
    const TABLE_INFO_TYPE: u32 = 6_000;
    const TABLE_MODEL_TYPE: u32 = 6_001;
    const TILE_TYPE: u32 = 6_002;
    const LIST_TYPE: u32 = 6_005;
    const STYLE_TYPE: u32 = 2_022;
    const STYLESHEET_TYPE: u32 = 401;

    let shared_key = 1;
    let mut cell = litchi_numbers_wire::BncCell::minimal();
    cell.set_text_style_identifier(Some(shared_key));
    let cell = cell.encode(&litchi_numbers_wire::Limits::DEFAULT)?;
    let mut rows = Vec::new();
    for row_index in 0..50_u32 {
        let mut storage = Vec::new();
        let mut offsets = vec![u16::MAX.to_le_bytes(); 20]
            .into_iter()
            .flatten()
            .collect::<Vec<_>>();
        for column in 0..20_usize {
            let offset = u16::try_from(storage.len())?;
            offsets[column * 2..column * 2 + 2].copy_from_slice(&offset.to_le_bytes());
            storage.extend_from_slice(&cell);
        }
        rows.push(tst::TileRowInfo {
            tile_row_index: row_index,
            cell_count: 20,
            cell_storage_buffer_pre_bnc: Vec::new(),
            cell_offsets_pre_bnc: Vec::new(),
            storage_version: Some(5),
            cell_storage_buffer: Some(storage),
            cell_offsets: Some(offsets),
            has_wide_offsets: Some(false),
        });
    }
    let tile = tst::Tile {
        max_column: 19,
        max_row: 49,
        num_cells: 1_000,
        numrows: 50,
        row_infos: rows,
        storage_version: Some(5),
        last_saved_in_bnc: Some(true),
        should_use_wide_rows: Some(false),
    };

    let style_list = whole_cell_object(
        STYLE_LIST,
        LIST_TYPE,
        tst::TableDataList {
            list_type: tst::table_data_list::ListType::Style as i32,
            next_list_id: shared_key + 1,
            entries: vec![tst::table_data_list::ListEntry {
                key: shared_key,
                refcount: 1_000,
                reference: Some(whole_cell_reference(STYLE)),
                ..Default::default()
            }],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    let empty_list = |identifier, list_type| {
        whole_cell_object(
            identifier,
            LIST_TYPE,
            tst::TableDataList {
                list_type: list_type as i32,
                next_list_id: 1,
                ..Default::default()
            }
            .encode_to_vec(),
        )
    };
    let string_list = empty_list(STRING_LIST, tst::table_data_list::ListType::String)?;
    let formula_list = empty_list(FORMULA_LIST, tst::table_data_list::ListType::Formula)?;
    let format_list = empty_list(FORMAT_LIST, tst::table_data_list::ListType::Format)?;
    let formula_error_list = empty_list(
        FORMULA_ERROR_LIST,
        tst::table_data_list::ListType::FormulaError,
    )?;
    let rich_text_list = empty_list(
        RICH_TEXT_LIST,
        tst::table_data_list::ListType::RichTextPayload,
    )?;
    let style = whole_cell_object(
        STYLE,
        STYLE_TYPE,
        tswp::ParagraphStyleArchive {
            super_: tss::StyleArchive {
                stylesheet: Some(whole_cell_reference(STYLESHEET)),
                ..Default::default()
            },
            char_properties: Some(tswp::CharacterStylePropertiesArchive::default()),
            para_properties: Some(tswp::ParagraphStylePropertiesArchive::default()),
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    let stylesheet = whole_cell_object(
        STYLESHEET,
        STYLESHEET_TYPE,
        tss::StylesheetArchive {
            styles: vec![whole_cell_reference(STYLE)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    let model = tst::TableModelArchive {
        table_id: "whole-cell-style-cost".to_owned(),
        table_style: whole_cell_reference(STYLE),
        body_text_style: whole_cell_reference(STYLE),
        header_row_text_style: whole_cell_reference(STYLE),
        header_column_text_style: whole_cell_reference(STYLE),
        footer_row_text_style: whole_cell_reference(STYLE),
        body_cell_style: whole_cell_reference(STYLE),
        header_row_style: whole_cell_reference(STYLE),
        header_column_style: whole_cell_reference(STYLE),
        footer_row_style: whole_cell_reference(STYLE),
        base_data_store: tst::DataStore {
            row_headers: tst::HeaderStorage {
                bucket_hash_function: 1,
                ..Default::default()
            },
            column_headers: whole_cell_reference(STRING_LIST),
            tiles: tst::TileStorage {
                tiles: vec![tst::tile_storage::Tile {
                    tileid: 0,
                    tile: whole_cell_reference(TILE),
                }],
                tile_size: Some(256),
                should_use_wide_rows: Some(false),
            },
            string_table: whole_cell_reference(STRING_LIST),
            style_table: whole_cell_reference(STYLE_LIST),
            formula_table: whole_cell_reference(FORMULA_LIST),
            formula_error_table: Some(whole_cell_reference(FORMULA_ERROR_LIST)),
            format_table_pre_bnc: whole_cell_reference(FORMAT_LIST),
            format_table: Some(whole_cell_reference(FORMAT_LIST)),
            rich_text_table: Some(whole_cell_reference(RICH_TEXT_LIST)),
            next_row_strip_id: 1,
            next_column_strip_id: 1,
            row_tile_tree: tst::TableRbTree::default(),
            column_tile_tree: tst::TableRbTree::default(),
            ..Default::default()
        },
        number_of_rows: 50,
        number_of_columns: 20,
        table_name: "Whole-cell style cost".to_owned(),
        default_row_height: 17.0,
        default_column_width: 60.0,
        ..Default::default()
    };
    let mut document = whole_cell_object(
        DOCUMENT,
        DOCUMENT_TYPE,
        tn::DocumentArchive {
            sheets: vec![whole_cell_reference(SHEET)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    let mut sheet = whole_cell_object(
        SHEET,
        SHEET_TYPE,
        tn::SheetArchive {
            name: "Whole-cell style sheet".to_owned(),
            drawable_infos: vec![whole_cell_reference(TABLE_INFO)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    let table_info = whole_cell_object(
        TABLE_INFO,
        TABLE_INFO_TYPE,
        tst::TableInfoArchive {
            table_model: whole_cell_reference(TABLE_MODEL),
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    let model = whole_cell_object(TABLE_MODEL, TABLE_MODEL_TYPE, model.encode_to_vec())?;
    let tile = whole_cell_object(TILE, TILE_TYPE, tile.encode_to_vec())?;
    document.archive_info.message_infos[0].object_references = vec![SHEET];
    sheet.archive_info.message_infos[0].object_references = vec![TABLE_INFO];
    let archive = Archive {
        objects: vec![
            document,
            sheet,
            table_info,
            model,
            style_list,
            tile,
            style,
            string_list,
            formula_list,
            format_list,
            formula_error_list,
            rich_text_list,
            stylesheet,
        ],
    };
    let compressed =
        SnappyStream::compress(&archive.to_bytes(Default::default())?, Default::default())?;
    // Preserve the native registry members needed by Numbers' publication
    // transition.  The selected table graph above remains source-built and
    // bounded; these small registry members are not measured table payload.
    let input = fixture("iwork/numbers/basic.numbers")?;
    let catalog = Catalog::from_bytes(&input.bytes, ArchiveLimits::default())?;
    let mut entries = vec![("Index/Document.iwa", compressed.as_slice())];
    for entry in catalog.iter().filter(|entry| {
        matches!(
            entry.name(),
            "Index/ViewState.iwa"
                | "Index/CalculationEngine.iwa"
                | "Index/AnnotationAuthorStorage.iwa"
                | "Index/DocumentStylesheet.iwa"
                | "Index/DocumentMetadata.iwa"
                | "Index/Metadata.iwa"
        )
    }) {
        entries.push((entry.name(), entry.data()));
    }
    let bytes = litchi_iwa_archive::package::to_bytes(entries, ArchiveLimits::default())?;
    Ok(WholeCellFixture {
        bytes,
        cells: (0..50)
            .flat_map(|row| (0..20).map(move |column| (row, column)))
            .collect(),
    })
}

fn shared_1000_cell_style_fixture() -> TestResult<WholeCellFixture> {
    const TABLE_MODEL: u32 = 6_001;
    const TILE: u32 = 6_002;
    const LIST: u32 = 6_005;
    let input = fixture("iwork/numbers/basic.numbers")?;
    let catalog = Catalog::from_bytes(&input.bytes, ArchiveLimits::default())?;
    let mut model_id = None;
    let mut model_member = None;
    let mut style_list_id = None;
    let mut style_id = None;
    let mut tile_id = None;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let bytes = SnappyStream::decompress(entry.data(), Default::default())?;
        let archive = Archive::parse(bytes.as_bytes(), Default::default())?;
        for object in &archive.objects {
            for message in &object.messages {
                if message.type_ != TABLE_MODEL {
                    continue;
                }
                let model = tst::TableModelArchive::decode(message.data.as_slice())?;
                model_id = object.archive_info.identifier;
                model_member = Some(entry.name().to_owned());
                style_list_id = Some(model.base_data_store.style_table.identifier);
                style_id = Some(model.body_text_style.identifier);
                tile_id = model
                    .base_data_store
                    .tiles
                    .tiles
                    .first()
                    .map(|tile| tile.tile.identifier);
            }
        }
    }
    let model_id = model_id.ok_or("missing table model")?;
    let model_member = model_member.ok_or("missing table model member")?;
    let style_list_id = style_list_id.ok_or("missing style list")?;
    let style_id = style_id.ok_or("missing body style")?;
    let tile_id = tile_id.ok_or("missing tile")?;
    let style_entry = catalog
        .iter()
        .find(|entry| {
            let Ok(bytes) = SnappyStream::decompress(entry.data(), Default::default()) else {
                return false;
            };
            let Ok(archive) = Archive::parse(bytes.as_bytes(), Default::default()) else {
                return false;
            };
            archive
                .objects
                .iter()
                .any(|object| object.archive_info.identifier == Some(style_list_id))
        })
        .ok_or("missing style list member")?;
    let style_member = style_entry.name().to_owned();
    let style_bytes = SnappyStream::decompress(style_entry.data(), Default::default())?;
    let mut style_archive = Archive::parse(style_bytes.as_bytes(), Default::default())?;
    let style_object = style_archive
        .objects
        .iter_mut()
        .find(|object| object.archive_info.identifier == Some(style_list_id))
        .ok_or("missing style list object")?;
    let style_message = style_object
        .messages
        .iter_mut()
        .find(|message| message.type_ == LIST)
        .ok_or("missing style list payload")?;
    let mut style_list = tst::TableDataList::decode(style_message.data.as_slice())?;
    style_list.list_type = tst::table_data_list::ListType::Style as i32;
    // The generated table below contains no cells that retain the template's
    // other style keys.  Keep only the one selected entry so each measured
    // request exercises the shared-key path rather than repeatedly scanning
    // an unrelated native style corpus.
    let shared_key = 1;
    style_list.entries.clear();
    style_list.segments.clear();
    style_list.next_list_id = shared_key + 1;
    style_list.entries.push(tst::table_data_list::ListEntry {
        key: shared_key,
        refcount: 1_000,
        reference: Some(litchi_iwa_protos::tsp::Reference {
            identifier: style_id,
            ..Default::default()
        }),
        ..Default::default()
    });
    style_message.data = style_list.encode_to_vec();
    if let Some(info) = style_object.archive_info.message_infos.first_mut()
        && !info.object_references.contains(&style_id)
    {
        info.object_references.push(style_id);
    }
    let style_compressed = SnappyStream::compress(
        &style_archive.to_bytes(Default::default())?,
        Default::default(),
    )?;

    let style_payload_entry = catalog
        .iter()
        .find(|entry| {
            let Ok(bytes) = SnappyStream::decompress(entry.data(), Default::default()) else {
                return false;
            };
            let Ok(archive) = Archive::parse(bytes.as_bytes(), Default::default()) else {
                return false;
            };
            archive.objects.iter().any(|object| {
                object.archive_info.identifier == Some(style_id)
                    && object.messages.iter().any(|message| message.type_ == 2_022)
            })
        })
        .ok_or("missing paragraph-style member")?;
    let style_payload_member = style_payload_entry.name().to_owned();
    let style_payload_bytes =
        SnappyStream::decompress(style_payload_entry.data(), Default::default())?;
    let mut style_payload_archive =
        Archive::parse(style_payload_bytes.as_bytes(), Default::default())?;
    let style_payload_object = style_payload_archive
        .objects
        .iter_mut()
        .find(|object| object.archive_info.identifier == Some(style_id))
        .ok_or("missing paragraph-style object")?;
    let style_payload_message = style_payload_object
        .messages
        .iter_mut()
        .find(|message| message.type_ == 2_022)
        .ok_or("missing paragraph-style payload")?;
    let existing_style =
        tswp::ParagraphStyleArchive::decode(style_payload_message.data.as_slice())?;
    let _old_stylesheet = existing_style
        .super_
        .stylesheet
        .ok_or("paragraph style has no stylesheet")?;
    let stylesheet_id = 9_800_000_001;
    let stylesheet = litchi_iwa_protos::tsp::Reference {
        identifier: stylesheet_id,
        ..Default::default()
    };
    style_payload_message.data = tswp::ParagraphStyleArchive {
        super_: tss::StyleArchive {
            stylesheet: Some(stylesheet),
            ..Default::default()
        },
        char_properties: Some(tswp::CharacterStylePropertiesArchive::default()),
        para_properties: Some(tswp::ParagraphStylePropertiesArchive::default()),
        ..Default::default()
    }
    .encode_to_vec();
    if let Some(info) = style_payload_object.archive_info.message_infos.first_mut() {
        info.object_references
            .retain(|reference| *reference != _old_stylesheet.identifier);
        if !info.object_references.contains(&stylesheet_id) {
            info.object_references.push(stylesheet_id);
        }
    }
    let _ = style_payload_object;
    style_payload_archive.objects.push(ArchiveObject::new(
        stylesheet_id,
        vec![RawMessage {
            type_: 401,
            data: tss::StylesheetArchive {
                styles: vec![litchi_iwa_protos::tsp::Reference {
                    identifier: style_id,
                    ..Default::default()
                }],
                ..Default::default()
            }
            .encode_to_vec(),
        }],
        Default::default(),
    )?);
    let style_payload_compressed = SnappyStream::compress(
        &style_payload_archive.to_bytes(Default::default())?,
        Default::default(),
    )?;
    let stylesheet_member = style_payload_member.clone();
    let stylesheet_compressed = style_payload_compressed.clone();

    let tile_entry = catalog
        .iter()
        .find(|entry| {
            let Ok(bytes) = SnappyStream::decompress(entry.data(), Default::default()) else {
                return false;
            };
            let Ok(archive) = Archive::parse(bytes.as_bytes(), Default::default()) else {
                return false;
            };
            archive
                .objects
                .iter()
                .any(|object| object.archive_info.identifier == Some(tile_id))
        })
        .ok_or("missing tile member")?;
    let tile_member = tile_entry.name().to_owned();
    let tile_bytes = SnappyStream::decompress(tile_entry.data(), Default::default())?;
    let mut tile_archive = Archive::parse(tile_bytes.as_bytes(), Default::default())?;
    let tile_object = tile_archive
        .objects
        .iter_mut()
        .find(|object| object.archive_info.identifier == Some(tile_id))
        .ok_or("missing tile object")?;
    let tile_message = tile_object
        .messages
        .iter_mut()
        .find(|message| message.type_ == TILE)
        .ok_or("missing tile payload")?;
    let tile = tst::Tile::decode(tile_message.data.as_slice())?;
    let mut sample_cell = litchi_numbers_wire::BncCell::minimal();
    sample_cell.set_text_style_identifier(Some(shared_key));
    let sample = sample_cell.encode(&litchi_numbers_wire::Limits::DEFAULT)?;
    let mut rows = Vec::new();
    for row_index in 0..50_u32 {
        let mut storage = Vec::new();
        let mut offsets = vec![u16::MAX.to_le_bytes(); 20]
            .into_iter()
            .flatten()
            .collect::<Vec<_>>();
        for column in 0..20_usize {
            let offset = u16::try_from(storage.len())?;
            offsets[column * 2..column * 2 + 2].copy_from_slice(&offset.to_le_bytes());
            storage.extend_from_slice(&sample);
        }
        rows.push(tst::TileRowInfo {
            tile_row_index: row_index,
            cell_count: 20,
            cell_storage_buffer_pre_bnc: Vec::new(),
            cell_offsets_pre_bnc: Vec::new(),
            storage_version: Some(5),
            cell_storage_buffer: Some(storage),
            cell_offsets: Some(offsets),
            has_wide_offsets: Some(false),
        });
    }
    let rewritten_tile = tst::Tile {
        max_column: 19,
        max_row: 49,
        num_cells: 1_000,
        numrows: 50,
        row_infos: rows,
        storage_version: tile.storage_version,
        last_saved_in_bnc: tile.last_saved_in_bnc,
        should_use_wide_rows: tile.should_use_wide_rows,
    };
    tile_message.data = rewritten_tile.encode_to_vec();
    let tile_compressed = SnappyStream::compress(
        &tile_archive.to_bytes(Default::default())?,
        Default::default(),
    )?;

    let model_entry = catalog
        .iter()
        .find(|entry| entry.name() == model_member)
        .ok_or("missing model member")?;
    let model_bytes = SnappyStream::decompress(model_entry.data(), Default::default())?;
    let mut model_archive = Archive::parse(model_bytes.as_bytes(), Default::default())?;
    let model_object = model_archive
        .objects
        .iter_mut()
        .find(|object| object.archive_info.identifier == Some(model_id))
        .ok_or("missing model object")?;
    let model_message = model_object
        .messages
        .iter_mut()
        .find(|message| message.type_ == TABLE_MODEL)
        .ok_or("missing model payload")?;
    let original_model = tst::TableModelArchive::decode(model_message.data.as_slice())?;
    let original_store = original_model.base_data_store;
    let style_reference = litchi_iwa_protos::tsp::Reference {
        identifier: style_id,
        ..Default::default()
    };
    let model = tst::TableModelArchive {
        table_id: "whole-cell-style-cost".to_owned(),
        table_style: style_reference,
        body_text_style: style_reference,
        header_row_text_style: style_reference,
        header_column_text_style: style_reference,
        footer_row_text_style: style_reference,
        body_cell_style: style_reference,
        header_row_style: style_reference,
        header_column_style: style_reference,
        footer_row_style: style_reference,
        base_data_store: tst::DataStore {
            row_headers: original_store.row_headers,
            column_headers: original_store.column_headers,
            tiles: tst::TileStorage {
                tiles: vec![tst::tile_storage::Tile {
                    tileid: 0,
                    tile: litchi_iwa_protos::tsp::Reference {
                        identifier: tile_id,
                        ..Default::default()
                    },
                }],
                tile_size: Some(256),
                should_use_wide_rows: Some(false),
            },
            string_table: original_store.string_table,
            style_table: original_store.style_table,
            formula_table: original_store.formula_table,
            formula_error_table: original_store.formula_error_table,
            format_table_pre_bnc: original_store.format_table_pre_bnc,
            multiple_choice_list_format_table: original_store.multiple_choice_list_format_table,
            merge_region_map: original_store.merge_region_map,
            deprecated_custom_format_table: original_store.deprecated_custom_format_table,
            next_row_strip_id: 1,
            next_column_strip_id: 1,
            row_tile_tree: tst::TableRbTree::default(),
            column_tile_tree: tst::TableRbTree::default(),
            storage_version_pre_bnc: original_store.storage_version_pre_bnc,
            rich_text_table: original_store.rich_text_table,
            conditionalstyletable: original_store.conditionalstyletable,
            comment_storage_table: original_store.comment_storage_table,
            import_warning_set_table: original_store.import_warning_set_table,
            control_cell_spec_table: original_store.control_cell_spec_table,
            format_table: original_store.format_table,
        },
        number_of_rows: 50,
        number_of_columns: 20,
        table_name: "Whole-cell style cost".to_owned(),
        default_row_height: 17.0,
        default_column_width: 60.0,
        ..Default::default()
    };
    model_message.data = model.encode_to_vec();
    let model_compressed = SnappyStream::compress(
        &model_archive.to_bytes(Default::default())?,
        Default::default(),
    )?;
    // The source-built tile contains all 1,000 cells, so keep its row-header
    // bucket authoritative as well.  Otherwise the host route legitimately
    // repairs stale header counts (and the leaf tile-only route does not),
    // turning a fixture-integrity difference into a false locality failure.
    let header_member = "Index/Tables/HeaderStorageBucket-904855-2.iwa";
    let header_entry = catalog
        .iter()
        .find(|entry| entry.name() == header_member)
        .ok_or("missing row-header bucket member")?;
    let header_bytes = SnappyStream::decompress(header_entry.data(), Default::default())?;
    let mut header_archive = Archive::parse(header_bytes.as_bytes(), Default::default())?;
    let mut header_found = false;
    for object in &mut header_archive.objects {
        for message in &mut object.messages {
            if message.type_ != 6_006 {
                continue;
            }
            message.data = tst::HeaderStorageBucket {
                bucket_hash_function: 1,
                headers: (0..50)
                    .map(|index| tst::header_storage_bucket::Header {
                        index,
                        size: 0.0,
                        hiding_state: 0,
                        number_of_cells: 20,
                        cell_style: None,
                        text_style: None,
                    })
                    .collect(),
            }
            .encode_to_vec();
            header_found = true;
        }
    }
    if !header_found {
        return Err("row-header bucket payload missing".into());
    }
    let header_compressed = SnappyStream::compress(
        &header_archive.to_bytes(Default::default())?,
        Default::default(),
    )?;
    let mut edits = vec![EntryEdit::new(&style_member, &style_compressed)];
    if style_payload_member != stylesheet_member {
        edits.push(EntryEdit::new(
            &style_payload_member,
            &style_payload_compressed,
        ));
    }
    edits.push(EntryEdit::new(&stylesheet_member, &stylesheet_compressed));
    edits.push(EntryEdit::new(&tile_member, &tile_compressed));
    edits.push(EntryEdit::new(&model_member, &model_compressed));
    edits.push(EntryEdit::new(header_member, &header_compressed));
    let bytes = catalog.reassemble_to_bytes(&edits, ArchiveLimits::default())?;
    Ok(WholeCellFixture {
        bytes,
        cells: (0..50)
            .flat_map(|row| (0..20).map(move |column| (row, column)))
            .collect(),
    })
}

#[inline(never)]
fn host_cell_text(source: &[u8], reset: bool) -> TestResult<Vec<u8>> {
    use litchi_iwa::text::Alignment;
    let mut editor = numbers::Draft::from_bytes(source)?;
    let id = editor
        .package()
        .iwa_entry_names()
        .find_map(|name| {
            editor
                .package()
                .archive(name)
                .ok()?
                .objects
                .into_iter()
                .find_map(|object| {
                    object
                        .messages
                        .iter()
                        .any(|message| message.type_ == 6_001)
                        .then_some(object.archive_info.identifier)
                        .flatten()
                })
        })
        .ok_or("no table model")?;
    if reset {
        editor.reset_table_cell_text_alignment(id, 1, 1)?;
    } else {
        editor.set_table_cell_text_alignment(id, 1, 1, Alignment::Justified)?;
    }
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_cell_text(source: &[u8], reset: bool) -> TestResult<Vec<u8>> {
    use litchi_iwa::text::Alignment;
    use litchi_numbers::cell::text::{Property, PropertyKind};
    let package = litchi_numbers::Package::from_bytes(source)?;
    let at = litchi_numbers::table::At::new(1, 1);
    let edit = package.edit_table_cell_text(0usize, 0usize)?;
    let edit = if reset {
        edit.reset(at, PropertyKind::Alignment)?
    } else {
        edit.set(at, Property::Alignment(Alignment::Justified))?
    };
    Ok(leaf_bytes!(edit.commit()?.package()))
}

/// Cell paragraph alignment on `basic.numbers` B2: set, then reset on the
/// host's output.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_cell_text_cost_abba() -> TestResult {
    let input = fixture("iwork/numbers/basic.numbers")?;
    let set = host_cell_text(&input.bytes, false)?;
    for (bytes, reset) in [(&input.bytes, false), (&set, true)] {
        black_box(host_cell_text(bytes, reset)?);
        black_box(leaf_cell_text(bytes, reset)?);
        black_box(leaf_cell_text(bytes, reset)?);
        black_box(host_cell_text(bytes, reset)?);
    }
    Ok(())
}

#[inline(never)]
fn host_cell_list(source: &[u8], numbered: bool) -> TestResult<Vec<u8>> {
    use litchi_iwa_text::paragraph::list::ParagraphList;
    let mut editor = numbers::Draft::from_bytes(source)?;
    let id = crate::numbers_cases::host_first_table_model_id(&editor)?;
    let list = if numbered {
        ParagraphList::Numbered
    } else {
        ParagraphList::Bullet
    };
    editor.set_table_cell_paragraph_list(id, 1, 1, list)?;
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_cell_list(source: &[u8], numbered: bool) -> TestResult<Vec<u8>> {
    use litchi_numbers::cell::list::ParagraphList;
    let package = litchi_numbers::Package::from_bytes(source)?;
    let list = if numbered {
        ParagraphList::Numbered
    } else {
        ParagraphList::Bullet
    };
    let commit = package
        .edit_table_cell_lists(0usize, 0usize)?
        .set_list(litchi_numbers::table::At::new(1, 1), list)?
        .commit()?;
    Ok(leaf_bytes!(commit.package()))
}

/// Cell paragraph list on `basic.numbers` B2: a bullet list on the plain
/// cell (promotion), then a numbered list on the host's output.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_cell_list_cost_abba() -> TestResult {
    let input = fixture("iwork/numbers/basic.numbers")?;
    let listed = host_cell_list(&input.bytes, false)?;
    for (bytes, numbered) in [(&input.bytes, false), (&listed, true)] {
        black_box(host_cell_list(bytes, numbered)?);
        black_box(leaf_cell_list(bytes, numbered)?);
        black_box(leaf_cell_list(bytes, numbered)?);
        black_box(host_cell_list(bytes, numbered)?);
    }
    Ok(())
}

#[derive(Clone, Copy)]
enum Remove {
    Row(litchi_numbers::table::topology::RowDeletion),
    Column(litchi_numbers::table::topology::ColumnDeletion),
}

#[inline(never)]
fn host_removal(source: &[u8], op: Remove) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let table = litchi_numbers::TableSelector::index(0);
    match op {
        Remove::Row(at) => editor.remove_table_row(table, at)?,
        Remove::Column(at) => editor.remove_table_column(table, at)?,
    }
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_removal(source: &[u8], op: Remove) -> TestResult<Vec<u8>> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let edit = package.edit_table_topology(0usize, 0usize)?;
    let edit = match op {
        Remove::Row(at) => edit.remove_row(at)?,
        Remove::Column(at) => edit.remove_column(at)?,
    };
    Ok(leaf_bytes!(edit.commit()?.package()))
}

/// Row and column removal on `basic.numbers` and a native fixture.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_topology_remove_cost_abba() -> TestResult {
    use litchi_numbers::table::topology::{ColumnDeletion, RowDeletion};
    let inputs = [
        fixture("iwork/numbers/basic.numbers")?,
        fixture("iwork/numbers/custom-number-native.numbers")?,
    ];
    let ops = [
        Remove::Row(RowDeletion::body(0)),
        Remove::Column(ColumnDeletion::body(0)),
    ];
    for input in &inputs {
        for op in ops {
            black_box(host_removal(&input.bytes, op)?);
            black_box(leaf_removal(&input.bytes, op)?);
            black_box(leaf_removal(&input.bytes, op)?);
            black_box(host_removal(&input.bytes, op)?);
        }
    }
    Ok(())
}

fn highlight_rules() -> TestResult<Vec<litchi_numbers::cell::conditional_highlight::Rule>> {
    use litchi_iwa_common::color::{RgbColorSpace, Rgba};
    use litchi_numbers::cell::conditional_highlight::{Condition, Number, Rule, Style, Text};
    let fill = Rgba::new(1.0, 0.8, 0.2, 1.0, RgbColorSpace::Srgb)?;
    Ok(vec![
        Rule::new(
            Condition::GreaterThan(Number::new(2.5)?),
            Style::with_fill(fill),
        ),
        Rule::new(
            Condition::TextContains(Text::new("ot")?),
            Style::with_fill(fill),
        ),
    ])
}

#[inline(never)]
fn host_highlight(source: &[u8], clear: bool) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let id = crate::numbers_cases::host_first_table_model_id(&editor)?;
    if clear {
        editor.clear_cell_conditional_highlighting(id, 1, 1)?;
    } else {
        editor.set_cell_conditional_highlighting(id, 1, 1, &highlight_rules()?)?;
    }
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_highlight(source: &[u8], clear: bool) -> TestResult<Vec<u8>> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let at = litchi_numbers::table::At::new(1, 1);
    let edit = package.edit_table_cell_conditional_highlights(0usize, 0usize)?;
    let edit = if clear {
        edit.clear(at)?
    } else {
        edit.set(at, highlight_rules()?)?
    };
    Ok(leaf_bytes!(edit.commit()?.package()))
}

/// Conditional highlighting on `basic.numbers` B2: two rules, then a clear
/// on the host's output.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_conditional_highlight_cost_abba() -> TestResult {
    let input = fixture("iwork/numbers/basic.numbers")?;
    let set = host_highlight(&input.bytes, false)?;
    for (bytes, clear) in [(&input.bytes, false), (&set, true)] {
        black_box(host_highlight(bytes, clear)?);
        black_box(leaf_highlight(bytes, clear)?);
        black_box(leaf_highlight(bytes, clear)?);
        black_box(host_highlight(bytes, clear)?);
    }
    Ok(())
}

#[inline(never)]
fn host_resize(source: &[u8], rows: usize, columns: usize) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    editor.resize_table(litchi_numbers::TableSelector::index(0), rows, columns)?;
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_resize(source: &[u8], rows: u32, columns: u32) -> TestResult<Vec<u8>> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let commit = package
        .edit_table_resize(0usize, 0usize)?
        .resize(rows, columns)?
        .commit()?;
    Ok(leaf_bytes!(commit.package()))
}

/// Growing the first table of `basic.numbers` by five rows and four columns.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_table_resize_cost_abba() -> TestResult {
    let input = fixture("iwork/numbers/basic.numbers")?;
    black_box(host_resize(&input.bytes, 26, 11)?);
    black_box(leaf_resize(&input.bytes, 26, 11)?);
    black_box(leaf_resize(&input.bytes, 26, 11)?);
    black_box(host_resize(&input.bytes, 26, 11)?);
    Ok(())
}

#[inline(never)]
fn host_table_add(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    editor.add_empty_table(litchi_numbers::SheetSelector::index(0), "Added Table", 4, 3)?;
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_table_add(source: &[u8]) -> TestResult<Vec<u8>> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let commit = package
        .edit_table_lifecycle(0)?
        .add("Added Table", 4, 3)?
        .commit()?;
    Ok(leaf_bytes!(commit.package()))
}

/// Adding a 4×3 table to the first sheet of `basic.numbers`.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_table_add_cost_abba() -> TestResult {
    let input = fixture("iwork/numbers/basic.numbers")?;
    black_box(host_table_add(&input.bytes)?);
    black_box(leaf_table_add(&input.bytes)?);
    black_box(leaf_table_add(&input.bytes)?);
    black_box(host_table_add(&input.bytes)?);
    Ok(())
}

/// Adding a 4×3 table to a workbook without a table (the blank workbook
/// with its table removed): the host's bootstrap against the leaf's shared
/// from-scratch graph (record 0420).
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_table_add_without_template_cost_abba() -> TestResult {
    let tableless = litchi_numbers::Package::new()
        .edit_table_lifecycle(0)?
        .remove(0)?
        .commit()?;
    let input = leaf_bytes!(tableless.package());
    black_box(host_table_add(&input)?);
    black_box(leaf_table_add(&input)?);
    black_box(leaf_table_add(&input)?);
    black_box(host_table_add(&input)?);
    Ok(())
}

#[inline(never)]
fn host_lifecycle(source: &[u8], operation: u8) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    match operation {
        0 => {
            editor.add_empty_sheet("Added Sheet")?;
        },
        1 => {
            editor.remove_table(litchi_numbers::TableSelector::index(0))?;
        },
        3 => {
            editor.duplicate_table(litchi_numbers::TableSelector::index(0))?;
        },
        4 => {
            editor.duplicate_sheet(litchi_numbers::SheetSelector::index(0))?;
        },
        _ => {
            editor.remove_sheet(litchi_numbers::SheetSelector::index(1))?;
        },
    }
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_lifecycle(source: &[u8], operation: u8) -> TestResult<Vec<u8>> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let bytes = match operation {
        0 => {
            let commit = package
                .edit_sheet_lifecycle()
                .add("Added Sheet")?
                .commit()?;
            leaf_bytes!(commit.package())
        },
        1 => {
            let commit = package.edit_table_lifecycle(0)?.remove(0)?.commit()?;
            leaf_bytes!(commit.package())
        },
        3 => {
            let commit = package.edit_table_lifecycle(0)?.duplicate(0)?.commit()?;
            leaf_bytes!(commit.package())
        },
        4 => {
            let commit = package.edit_sheet_lifecycle().duplicate(0)?.commit()?;
            leaf_bytes!(commit.package())
        },
        _ => {
            let commit = package.edit_sheet_lifecycle().remove(1)?.commit()?;
            leaf_bytes!(commit.package())
        },
    };
    Ok(bytes)
}

/// Adding a sheet to `basic.numbers`, removing its first table, and
/// removing a second sheet holding a 4×3 table.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_lifecycle_cost_abba() -> TestResult {
    let input = fixture("iwork/numbers/basic.numbers")?;
    let two = litchi_numbers::Package::from_bytes(&input.bytes)?
        .edit_sheet_lifecycle()
        .add("Second")?
        .commit()?
        .into_package()
        .edit_table_lifecycle(1)?
        .add("Second Table", 4, 3)?
        .commit()?
        .into_package();
    let two = leaf_bytes!(&two);
    for (source, operation) in [(&input.bytes, 0), (&input.bytes, 1), (&two, 2)] {
        black_box(host_lifecycle(source, operation)?);
        black_box(leaf_lifecycle(source, operation)?);
        black_box(leaf_lifecycle(source, operation)?);
        black_box(host_lifecycle(source, operation)?);
    }
    Ok(())
}

/// Duplicating the first table of `basic.numbers`.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_table_duplicate_cost_abba() -> TestResult {
    let input = fixture("iwork/numbers/basic.numbers")?;
    black_box(host_lifecycle(&input.bytes, 3)?);
    black_box(leaf_lifecycle(&input.bytes, 3)?);
    black_box(leaf_lifecycle(&input.bytes, 3)?);
    black_box(host_lifecycle(&input.bytes, 3)?);
    Ok(())
}

#[inline(never)]
fn host_builder() -> TestResult<Vec<u8>> {
    Ok(numbers::Draft::builder()
        .table_dimensions(20, 8)
        .language("fr")
        .locale("fr_FR")
        .build()?
        .to_bytes()?)
}

#[inline(never)]
fn leaf_builder() -> TestResult<Vec<u8>> {
    let package = litchi_numbers::Builder::new()
        .identity(litchi_numbers::Identity::BLANK)
        .table_dimensions(20, 8)
        .language("fr")
        .locale("fr_FR")
        .build()?;
    Ok(leaf_bytes!(&package))
}

/// A 20×8 `fr`/`fr_FR` workbook from each builder.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_builder_cost_abba() -> TestResult {
    black_box(host_builder()?);
    black_box(leaf_builder()?);
    black_box(leaf_builder()?);
    black_box(host_builder()?);
    Ok(())
}

/// Duplicating the first sheet of `basic.numbers`.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_sheet_duplicate_cost_abba() -> TestResult {
    let input = fixture("iwork/numbers/basic.numbers")?;
    black_box(host_lifecycle(&input.bytes, 4)?);
    black_box(leaf_lifecycle(&input.bytes, 4)?);
    black_box(leaf_lifecycle(&input.bytes, 4)?);
    black_box(host_lifecycle(&input.bytes, 4)?);
    Ok(())
}

#[inline(never)]
fn host_data_format(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let model = crate::numbers_cases::host_first_table_model_id(&editor)?;
    let numeral = crate::numbers_cell_data_format_cases::numeral_samples()?[1];
    editor.set_table_cell_numeral_system_format(model, 1, 1, numeral)?;
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_data_format(source: &[u8]) -> TestResult<Vec<u8>> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let numeral = crate::numbers_cell_data_format_cases::numeral_samples()?[1];
    let commit = package
        .edit_table_cell_data_format(0usize, 0usize, litchi_numbers::At::new(1, 1))?
        .set(numeral)
        .commit()?;
    Ok(leaf_bytes!(commit.package()))
}

/// A base-16 numeral system on cell B2 of `custom-retirement-cleared.numbers`.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_cell_data_format_cost_abba() -> TestResult {
    let input = fixture("iwork/numbers/custom-retirement-cleared.numbers")?;
    black_box(host_data_format(&input.bytes)?);
    black_box(leaf_data_format(&input.bytes)?);
    black_box(leaf_data_format(&input.bytes)?);
    black_box(host_data_format(&input.bytes)?);
    Ok(())
}

#[inline(never)]
fn host_slide_table_add(source: &[u8]) -> TestResult<Vec<u8>> {
    use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
    let mut editor = litchi_iwa::keynote::Draft::from_bytes(source)?;
    editor.add_slide_table(
        0,
        "Litchi table",
        3,
        4,
        DrawablePoint { x: 120.0, y: 140.0 },
        DrawableSize {
            width: 360.0,
            height: 180.0,
        },
    )?;
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_slide_table_add(source: &[u8]) -> TestResult<Vec<u8>> {
    let package = litchi_keynote::Package::from_bytes(source)?;
    let commit = package
        .edit_slide_table_lifecycle()
        .add(
            litchi_keynote::SlideSelector::index(0),
            "Litchi table",
            3,
            4,
            litchi_keynote::slide::drawable::Geometry::new(120.0, 140.0, 360.0, 180.0),
        )?
        .commit()?;
    Ok(leaf_bytes!(commit.package()))
}

/// Adding a 3×4 slide table to a builder deck from its detached creation
/// table (record 0423).
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn keynote_slide_table_add_cost_abba() -> TestResult {
    let input = litchi_iwa::keynote::Builder::new()
        .title("Tables")
        .build()?
        .to_bytes()?;
    black_box(host_slide_table_add(&input)?);
    black_box(leaf_slide_table_add(&input)?);
    black_box(leaf_slide_table_add(&input)?);
    black_box(host_slide_table_add(&input)?);
    Ok(())
}

#[inline(never)]
fn host_shape_image_fill(source: &[u8], png: &[u8]) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let (sheet, identifier) = crate::numbers_drawable_cases::host_drawable(
        &editor,
        litchi_numbers::sheet::drawable::Kind::Shape,
        0,
    )?;
    editor.set_sheet_shape_image_fill(
        sheet,
        identifier,
        "poster.png",
        png,
        litchi_iwa::shapes::ShapeImageFillTechnique::Tile,
        None,
    )?;
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_shape_image_fill(source: &[u8], png: &[u8]) -> TestResult<Vec<u8>> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let mut edit = package.edit_sheet_drawable_style(
        0,
        litchi_numbers::sheet::drawable::DrawableSelector::of(
            litchi_numbers::sheet::drawable::Kind::Shape,
            0,
        ),
    )?;
    edit.set_image_fill(
        png,
        "poster.png",
        litchi_iwa_common::shape::fill::ImageTechnique::Tile,
        None,
    )?;
    Ok(leaf_bytes!(edit.commit()?.package()))
}

/// A tiled image fill on the first shape of the authored Numbers fixture
/// (record 0423).
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_shape_image_fill_cost_abba() -> TestResult {
    let input = crate::numbers_drawable_cases::authored_fixture()?.bytes;
    let png = fixture("iwork/keynote/media-replacement-assets/poster.png")?.bytes;
    black_box(host_shape_image_fill(&input, &png)?);
    black_box(leaf_shape_image_fill(&input, &png)?);
    black_box(leaf_shape_image_fill(&input, &png)?);
    black_box(host_shape_image_fill(&input, &png)?);
    Ok(())
}
