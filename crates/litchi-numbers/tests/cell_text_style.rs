//! Cell character properties (record 0309): bold, italic, size, font and
//! colour of cell text, on the created workbook and across the corpus.

use std::path::PathBuf;

use litchi_iwa_archive::{
    Limits as ArchiveLimits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::{WireLimits, wire::WireView};
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::tst;
use litchi_numbers::cell::text_style::{CharacterProperties, Error};
use litchi_numbers::table::cells::Input;
use litchi_numbers::{At, Package};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

fn bold() -> CharacterProperties {
    let mut properties = CharacterProperties::new();
    properties.bold = Some(true);
    properties
}

struct SharedTextStyleSource {
    bytes: Vec<u8>,
    key: u32,
    positions: [(u32, u32); 2],
    style_id: u64,
    style_payload: Vec<u8>,
}

fn append_wire_field(data: &mut Vec<u8>, field: u32, payload: &[u8]) {
    litchi_iwa_common::varint::encode_varint_into(data, (u64::from(field) << 3) | 2);
    litchi_iwa_common::varint::encode_varint_into(data, payload.len() as u64);
    data.extend_from_slice(payload);
}

fn shared_text_style_source() -> TestResult<SharedTextStyleSource> {
    const TABLE_MODEL: u32 = 6_001;
    const TILE: u32 = 6_002;
    const LIST: u32 = 6_005;
    let source = std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/numbers/basic.numbers"),
    )?;
    let catalog = Catalog::from_bytes(&source, ArchiveLimits::default())?;
    let mut style_list_id = None;
    let mut style_id = None;
    let mut tile_profiles = Vec::new();
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let bytes = SnappyStream::decompress(entry.data(), Default::default())?;
        let archive = Archive::parse(bytes.as_bytes(), Default::default())?;
        for object in &archive.objects {
            for message in &object.messages {
                if message.type_ == TABLE_MODEL {
                    let model = tst::TableModelArchive::decode(message.data.as_slice())?;
                    style_list_id = Some(model.base_data_store.style_table.identifier);
                    style_id = Some(model.body_text_style.identifier);
                    let tile_size = model
                        .base_data_store
                        .tiles
                        .tile_size
                        .ok_or("missing tile size")?;
                    for tile in &model.base_data_store.tiles.tiles {
                        tile_profiles.push((tile.tile.identifier, tile.tileid, tile_size));
                    }
                }
            }
        }
    }
    let style_list_id = style_list_id.ok_or("missing style-list reference")?;
    let style_id = style_id.ok_or("missing body text-style reference")?;
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
        .ok_or("missing style-list member")?;
    let style_member = style_entry.name().to_owned();
    let style_bytes = SnappyStream::decompress(style_entry.data(), Default::default())?;
    let mut style_archive = Archive::parse(style_bytes.as_bytes(), Default::default())?;
    let style_object = style_archive
        .objects
        .iter_mut()
        .find(|object| object.archive_info.identifier == Some(style_list_id))
        .ok_or("missing style-list object")?;
    let style_message = style_object
        .messages
        .iter_mut()
        .find(|message| message.type_ == LIST)
        .ok_or("missing style-list payload")?;
    let mut style_list = tst::TableDataList::decode(style_message.data.as_slice())?;
    style_list.list_type = tst::table_data_list::ListType::Style as i32;
    let shared_key = style_list
        .entries
        .iter()
        .map(|entry| entry.key)
        .max()
        .unwrap_or(0)
        .saturating_add(1);
    style_list.next_list_id = shared_key.saturating_add(1);
    style_list.entries.push(tst::table_data_list::ListEntry {
        key: shared_key,
        refcount: 2,
        reference: Some(litchi_iwa_protos::tsp::Reference {
            identifier: style_id,
            ..Default::default()
        }),
        ..Default::default()
    });
    style_message.data = style_list.encode_to_vec();
    if let Some(info) = style_object.archive_info.message_infos.first_mut() {
        if !info.object_references.contains(&style_id) {
            info.object_references.push(style_id);
        }
    }
    let style_compressed = SnappyStream::compress(
        &style_archive.to_bytes(Default::default())?,
        Default::default(),
    )?;
    let style_object_entry = catalog
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
    let style_object_member = style_object_entry.name().to_owned();
    if style_object_member == style_member {
        return Err("style-list and paragraph-style members unexpectedly overlap".into());
    }
    let style_object_bytes =
        SnappyStream::decompress(style_object_entry.data(), Default::default())?;
    let mut style_object_archive =
        Archive::parse(style_object_bytes.as_bytes(), Default::default())?;
    let style_object = style_object_archive
        .objects
        .iter_mut()
        .find(|object| object.archive_info.identifier == Some(style_id))
        .ok_or("missing paragraph-style object")?;
    let style_payload = style_object
        .messages
        .iter_mut()
        .find(|message| message.type_ == 2_022)
        .ok_or("missing paragraph-style payload")?;
    let mut preserved_payload = style_payload.data.clone();
    let existing_fields = WireView::parse(&preserved_payload, WireLimits::DEFAULT)?;
    if !existing_fields.fields().any(|field| field.number() == 12) {
        append_wire_field(&mut preserved_payload, 12, &[8, 1]);
    }
    append_wire_field(&mut preserved_payload, 99, b"unknown-style-tail");
    style_payload.data = preserved_payload.clone();
    let style_object_compressed = SnappyStream::compress(
        &style_object_archive.to_bytes(Default::default())?,
        Default::default(),
    )?;

    let tile_entry = catalog
        .iter()
        .find(|entry| entry.name() == "Index/Tables/Tile.iwa")
        .ok_or("missing tile member")?;
    let tile_bytes = SnappyStream::decompress(tile_entry.data(), Default::default())?;
    let mut tile_archive = Archive::parse(tile_bytes.as_bytes(), Default::default())?;
    let tile_object_id = tile_archive
        .objects
        .iter()
        .find(|object| object.messages.iter().any(|message| message.type_ == TILE))
        .and_then(|object| object.archive_info.identifier)
        .ok_or("missing tile object identifier")?;
    let (tile_index, tile_size) = tile_profiles
        .iter()
        .find(|(identifier, _, _)| *identifier == tile_object_id)
        .map(|(_, index, size)| (*index, *size))
        .ok_or("tile is not owned by the selected table")?;
    let tile_message = tile_archive
        .objects
        .iter_mut()
        .flat_map(|object| object.messages.iter_mut())
        .find(|message| message.type_ == TILE)
        .ok_or("missing tile payload")?;
    let mut tile = tst::Tile::decode(tile_message.data.as_slice())?;
    let mut changed = 0_usize;
    let mut positions = Vec::new();
    for row in &mut tile.row_infos {
        if changed == 2 || row.cell_count == 0 {
            continue;
        }
        let Some(storage) = row.cell_storage_buffer.clone() else {
            continue;
        };
        let Some(offsets) = row.cell_offsets.clone() else {
            continue;
        };
        let unit = if row.has_wide_offsets == Some(true) {
            4
        } else {
            1
        };
        let slots = offsets.len() / 2;
        let present = (0..slots)
            .filter(|index| {
                u16::from_le_bytes([offsets[index * 2], offsets[index * 2 + 1]]) != u16::MAX
            })
            .collect::<Vec<_>>();
        if present.is_empty() {
            continue;
        }
        let table_row = tile_index
            .checked_mul(tile_size)
            .and_then(|base| base.checked_add(row.tile_row_index))
            .ok_or("tile row coordinate overflow")?;
        positions.push((table_row, u32::try_from(present[0])?));
        let mut cells = vec![None; slots];
        for (position, index) in present.iter().enumerate() {
            let begin = usize::from(u16::from_le_bytes([
                offsets[index * 2],
                offsets[index * 2 + 1],
            ])) * unit;
            let end = present
                .get(position + 1)
                .map(|next| {
                    usize::from(u16::from_le_bytes([
                        offsets[next * 2],
                        offsets[next * 2 + 1],
                    ])) * unit
                })
                .unwrap_or(storage.len());
            cells[*index] = Some(storage.get(begin..end).ok_or("invalid tile span")?.to_vec());
        }
        let first = present[0];
        let mut cell = litchi_numbers_wire::BncCell::parse(
            cells[first].as_deref().ok_or("missing materialized cell")?,
            &litchi_numbers_wire::Limits::DEFAULT,
        )?;
        cell.set_text_style_identifier(Some(shared_key));
        cells[first] = Some(cell.encode(&litchi_numbers_wire::Limits::DEFAULT)?);
        let mut rewritten_storage = Vec::new();
        let mut rewritten_offsets = Vec::new();
        for cell in cells {
            if let Some(cell) = cell {
                let offset = u16::try_from(rewritten_storage.len() / unit)?;
                rewritten_offsets.extend_from_slice(&offset.to_le_bytes());
                rewritten_storage.extend_from_slice(&cell);
                while rewritten_storage.len() % unit != 0 {
                    rewritten_storage.push(0);
                }
            } else {
                rewritten_offsets.extend_from_slice(&u16::MAX.to_le_bytes());
            }
        }
        row.cell_storage_buffer = Some(rewritten_storage.clone());
        row.cell_offsets = Some(rewritten_offsets.clone());
        row.cell_storage_buffer_pre_bnc = rewritten_storage;
        row.cell_offsets_pre_bnc = rewritten_offsets;
        changed += 1;
    }
    assert_eq!(changed, 2, "source helper needs two materialized rows");
    let positions: [(u32, u32); 2] = positions
        .try_into()
        .map_err(|_| "source helper did not retain two positions")?;
    tile_message.data = tile.encode_to_vec();
    let tile_compressed = SnappyStream::compress(
        &tile_archive.to_bytes(Default::default())?,
        Default::default(),
    )?;
    let bytes = catalog.reassemble_to_bytes(
        &[
            EntryEdit::new(&style_member, &style_compressed),
            EntryEdit::new(&style_object_member, &style_object_compressed),
            EntryEdit::new("Index/Tables/Tile.iwa", &tile_compressed),
        ],
        ArchiveLimits::default(),
    )?;
    Ok(SharedTextStyleSource {
        bytes,
        key: shared_key,
        positions,
        style_id,
        style_payload: preserved_payload,
    })
}

fn style_entries(bytes: &[u8]) -> TestResult<Vec<(u32, u32)>> {
    const TABLE_MODEL: u32 = 6_001;
    const LIST: u32 = 6_005;
    let catalog = Catalog::from_bytes(bytes, ArchiveLimits::default())?;
    let mut style_list_id = None;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let bytes = SnappyStream::decompress(entry.data(), Default::default())?;
        let archive = Archive::parse(bytes.as_bytes(), Default::default())?;
        for object in &archive.objects {
            for message in &object.messages {
                if message.type_ == TABLE_MODEL {
                    style_list_id = Some(
                        tst::TableModelArchive::decode(message.data.as_slice())?
                            .base_data_store
                            .style_table
                            .identifier,
                    );
                }
            }
        }
    }
    let style_list_id = style_list_id.ok_or("missing style-list reference")?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let bytes = SnappyStream::decompress(entry.data(), Default::default())?;
        let archive = Archive::parse(bytes.as_bytes(), Default::default())?;
        for object in &archive.objects {
            if object.archive_info.identifier != Some(style_list_id) {
                continue;
            }
            for message in &object.messages {
                if message.type_ == LIST {
                    return Ok(tst::TableDataList::decode(message.data.as_slice())?
                        .entries
                        .into_iter()
                        .map(|entry| (entry.key, entry.refcount))
                        .collect());
                }
            }
        }
    }
    Err("missing style-list payload".into())
}

fn paragraph_style_payload(bytes: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(bytes, ArchiveLimits::default())?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let bytes = SnappyStream::decompress(entry.data(), Default::default())?;
        let archive = Archive::parse(bytes.as_bytes(), Default::default())?;
        for object in &archive.objects {
            if object.archive_info.identifier != Some(identifier) {
                continue;
            }
            if let Some(message) = object
                .messages
                .iter()
                .find(|message| message.type_ == 2_022)
            {
                return Ok(message.data.clone());
            }
        }
    }
    Err("missing paragraph-style payload".into())
}

fn mutate_paragraph_style_with_unknowns(
    bytes: &[u8],
    identifier: u64,
) -> TestResult<(Vec<u8>, Vec<u8>)> {
    let catalog = Catalog::from_bytes(bytes, ArchiveLimits::default())?;
    let entry = catalog
        .iter()
        .find(|entry| {
            let Ok(bytes) = SnappyStream::decompress(entry.data(), Default::default()) else {
                return false;
            };
            let Ok(archive) = Archive::parse(bytes.as_bytes(), Default::default()) else {
                return false;
            };
            archive.objects.iter().any(|object| {
                object.archive_info.identifier == Some(identifier)
                    && object.messages.iter().any(|message| message.type_ == 2_022)
            })
        })
        .ok_or("missing paragraph-style member")?;
    let member = entry.name().to_owned();
    let bytes = SnappyStream::decompress(entry.data(), Default::default())?;
    let mut archive = Archive::parse(bytes.as_bytes(), Default::default())?;
    let object = archive
        .objects
        .iter_mut()
        .find(|object| object.archive_info.identifier == Some(identifier))
        .ok_or("missing paragraph-style object")?;
    let message = object
        .messages
        .iter_mut()
        .find(|message| message.type_ == 2_022)
        .ok_or("missing paragraph-style payload")?;
    let payload = message.data.clone();
    let fields = WireView::parse(&payload, WireLimits::DEFAULT)?;
    let mut rewritten = Vec::new();
    let mut paragraph_seen = false;
    for field in fields.fields() {
        if field.number() == 12 {
            paragraph_seen = true;
            rewritten.extend_from_slice(field.key());
            litchi_iwa_common::varint::encode_varint_into(&mut rewritten, 2);
            rewritten.extend_from_slice(&[8, 1]);
        } else {
            rewritten.extend_from_slice(field.raw());
        }
    }
    if !paragraph_seen {
        append_wire_field(&mut rewritten, 12, &[8, 1]);
    }
    let mut payload = rewritten;
    append_wire_field(&mut payload, 99, b"unknown-exclusive-tail");
    message.data = payload.clone();
    let compressed =
        SnappyStream::compress(&archive.to_bytes(Default::default())?, Default::default())?;
    let mutated = catalog.reassemble_to_bytes(
        &[EntryEdit::new(&member, &compressed)],
        ArchiveLimits::default(),
    )?;
    Ok((mutated, payload))
}

fn cell_style_identity(bytes: &[u8], position: At) -> TestResult<(u32, u64, Option<u64>)> {
    const TABLE_MODEL: u32 = 6_001;
    const TILE: u32 = 6_002;
    const LIST: u32 = 6_005;
    let catalog = Catalog::from_bytes(bytes, ArchiveLimits::default())?;
    let mut model_id = None;
    let mut style_list_id = None;
    let mut tile_size = None;
    let mut tile_refs = Vec::new();
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
                style_list_id = Some(model.base_data_store.style_table.identifier);
                tile_size = model.base_data_store.tiles.tile_size;
                tile_refs = model
                    .base_data_store
                    .tiles
                    .tiles
                    .iter()
                    .map(|tile| (tile.tileid, tile.tile.identifier))
                    .collect();
            }
        }
    }
    let _model_id = model_id.ok_or("missing table model")?;
    let style_list_id = style_list_id.ok_or("missing style list")?;
    let tile_size = tile_size.ok_or("missing tile size")?;
    let tile_id = tile_refs
        .iter()
        .find(|(index, _)| *index == position.row() / tile_size)
        .map(|(_, identifier)| *identifier)
        .ok_or("missing selected tile")?;
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
        .ok_or("missing selected tile member")?;
    let tile_bytes = SnappyStream::decompress(tile_entry.data(), Default::default())?;
    let archive = Archive::parse(tile_bytes.as_bytes(), Default::default())?;
    let tile_message = archive
        .objects
        .iter()
        .flat_map(|object| object.messages.iter())
        .find(|message| message.type_ == TILE)
        .ok_or("missing selected tile payload")?;
    let tile = tst::Tile::decode(tile_message.data.as_slice())?;
    let row_index = position.row() % tile_size;
    let row = tile
        .row_infos
        .iter()
        .find(|row| row.tile_row_index == row_index)
        .ok_or("missing selected tile row")?;
    let storage = row
        .cell_storage_buffer
        .as_ref()
        .ok_or("missing selected row storage")?;
    let offsets = row
        .cell_offsets
        .as_ref()
        .ok_or("missing selected row offsets")?;
    let unit = if row.has_wide_offsets == Some(true) {
        4
    } else {
        1
    };
    let column = usize::try_from(position.column())?;
    let offset = column.checked_mul(2).ok_or("cell offset overflow")?;
    let begin = usize::from(u16::from_le_bytes(
        offsets
            .get(offset..offset + 2)
            .ok_or("cell offset missing")?
            .try_into()?,
    )) * unit;
    if begin == usize::from(u16::MAX) * unit {
        return Err("selected cell is not materialized".into());
    }
    let next = offsets
        .chunks_exact(2)
        .skip(column + 1)
        .find_map(|pair| {
            let value = u16::from_le_bytes([pair[0], pair[1]]);
            (value != u16::MAX).then_some(usize::from(value) * unit)
        })
        .unwrap_or(storage.len());
    let cell = storage
        .get(begin..next)
        .ok_or("selected cell span invalid")?;
    let cell = litchi_numbers_wire::BncCell::parse(cell, &litchi_numbers_wire::Limits::DEFAULT)?;
    let key = cell
        .text_style_identifier()
        .ok_or("selected cell has no text style key")?;

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
    let list_bytes = SnappyStream::decompress(style_entry.data(), Default::default())?;
    let list_archive = Archive::parse(list_bytes.as_bytes(), Default::default())?;
    let list_message = list_archive
        .objects
        .iter()
        .flat_map(|object| object.messages.iter())
        .find(|message| message.type_ == LIST)
        .ok_or("missing style list payload")?;
    let list = tst::TableDataList::decode(list_message.data.as_slice())?;
    let mut style_id = list
        .entries
        .iter()
        .find(|entry| entry.key == key)
        .and_then(|entry| entry.reference.as_ref())
        .map(|reference| reference.identifier);
    if style_id.is_none() {
        for segment_ref in &list.segments {
            let segment_id = segment_ref.identifier;
            let Some(segment_entry) = catalog.iter().find(|entry| {
                let Ok(bytes) = SnappyStream::decompress(entry.data(), Default::default()) else {
                    return false;
                };
                let Ok(archive) = Archive::parse(bytes.as_bytes(), Default::default()) else {
                    return false;
                };
                archive
                    .objects
                    .iter()
                    .any(|object| object.archive_info.identifier == Some(segment_id))
            }) else {
                continue;
            };
            let bytes = SnappyStream::decompress(segment_entry.data(), Default::default())?;
            let archive = Archive::parse(bytes.as_bytes(), Default::default())?;
            let Some(message) = archive
                .objects
                .iter()
                .flat_map(|object| object.messages.iter())
                .find(|message| message.type_ == 6_201)
            else {
                continue;
            };
            let segment = tst::TableDataListSegment::decode(message.data.as_slice())?;
            style_id = segment
                .entries
                .iter()
                .find(|entry| entry.key == key)
                .and_then(|entry| entry.reference.as_ref())
                .map(|reference| reference.identifier);
            if style_id.is_some() {
                break;
            }
        }
    }
    let style_id = style_id.ok_or("missing selected style object")?;
    let style_payload = paragraph_style_payload(bytes, style_id)?;
    let parent = litchi_iwa_text_wire::character_style_parent(&style_payload, WireLimits::DEFAULT)?;
    Ok((key, style_id, parent))
}

#[test]
fn cells_take_text_styles_on_a_created_workbook() -> TestResult {
    let package = Package::new()
        .edit_table_cells(0, 0)?
        .set(At::new(1, 1), Input::Text("b".into()))?
        .set(At::new(2, 1), Input::Text("c".into()))?
        .set(At::new(2, 2), Input::number(4.0)?)?
        .commit()?
        .into_package();
    let source = bytes(&package)?;
    let base = package.table_cell_text_style(0, 0, At::new(1, 1))?;
    let mut red = CharacterProperties::new();
    red.color = Some([200, 0, 0]);
    red.size = Some(18.0);
    red.font_name = Some("Helvetica".into());
    let commit = package
        .edit_table_cell_text_style(0, 0)?
        .set(At::new(1, 1), bold())?
        .set(At::new(2, 1), bold())?
        .set(At::new(2, 2), red.clone())?
        .set(At::new(2, 2), {
            let mut italic = CharacterProperties::new();
            italic.italic = Some(true);
            italic
        })?
        .commit()?;
    let styled = commit.package();
    let reopened = Package::from_bytes(&bytes(styled)?)?;
    for candidate in [styled, &reopened] {
        let b = candidate.table_cell_text_style(0, 0, At::new(1, 1))?;
        assert_eq!(b.bold, Some(true));
        // Unset properties stay inherited from the base style.
        assert_eq!(b.size, base.size);
        assert_eq!(b.font_name, base.font_name);
        let c = candidate.table_cell_text_style(0, 0, At::new(2, 2))?;
        assert_eq!(c.color, Some([200, 0, 0]));
        assert_eq!(c.size, Some(18.0));
        assert_eq!(c.italic, Some(true));
        assert_eq!(c.font_name.as_deref(), Some("Helvetica"));
        // Values are untouched.
        let table = candidate.sheets()[0].tables().next().expect("table");
        assert_eq!(table.iter_cells().count(), 3);
    }
    // Restyling a styled cell builds on its style and releases the old key.
    let mut unbold = CharacterProperties::new();
    unbold.bold = Some(false);
    let again_commit = styled
        .edit_table_cell_text_style(0, 0)?
        .set(At::new(1, 1), unbold)?
        .commit()?;
    let again_bytes = bytes(again_commit.package())?;
    let again_reopened = Package::from_bytes(&again_bytes)?;
    assert_eq!(
        again_reopened
            .table_cell_text_style(0, 0, At::new(1, 1))?
            .bold,
        Some(false)
    );
    assert_eq!(
        again_reopened
            .table_cell_text_style(0, 0, At::new(2, 1))?
            .bold,
        Some(true)
    );
    let restored_again = again_commit
        .package()
        .apply_table_cell_text_style(&again_commit.patch().inverse())?;
    assert_eq!(bytes(restored_again.package())?, bytes(styled)?);
    let again = again_commit.into_package();
    let restored = styled.apply_table_cell_text_style(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    // A value write keeps the style.
    let rewritten = again
        .edit_table_cells(0, 0)?
        .set(At::new(2, 1), Input::Text("changed".into()))?
        .commit()?
        .into_package();
    assert_eq!(
        rewritten.table_cell_text_style(0, 0, At::new(2, 1))?.bold,
        Some(true)
    );
    Ok(())
}

#[test]
fn shared_style_key_publishes_reopens_reads_and_reverses_exactly() -> TestResult {
    let generated = shared_text_style_source()?;
    let source = generated.bytes;
    let key = generated.key;
    let first = At::new(generated.positions[0].0, generated.positions[0].1);
    let second = At::new(generated.positions[1].0, generated.positions[1].1);
    assert!(
        style_entries(&source)?.contains(&(key, 2)),
        "source helper must publish one shared key"
    );
    let package = Package::from_bytes(&source)?;
    let sibling = (0..10)
        .flat_map(|row| (0..10).map(move |column| At::new(row, column)))
        .find(|candidate| *candidate != first && *candidate != second)
        .ok_or("missing locality candidate")?;
    let before_sibling = package.table_cell_text_style(0, 0, sibling)?;
    assert_eq!(
        package.table_cell_text_style(0, 0, first)?,
        package.table_cell_text_style(0, 0, second)?,
        "shared-key cells have the same direct read"
    );
    let mut first_change = CharacterProperties::new();
    first_change.size = Some(73.0);
    let first_commit = package
        .edit_table_cell_text_style(0, 0)?
        .set(first, first_change)?
        .commit()?;
    let first_bytes = bytes(first_commit.package())?;
    let first_identity = cell_style_identity(&first_bytes, first)?;
    assert_ne!(
        first_identity.2,
        Some(first_identity.1),
        "fresh style must not self-parent"
    );
    let first_reopened = Package::from_bytes(&first_bytes)?;
    let mut second_change = CharacterProperties::new();
    second_change.italic = Some(true);
    let second_commit = first_reopened
        .edit_table_cell_text_style(0, 0)?
        .set(first, second_change)?
        .commit()?;
    let second_bytes = bytes(second_commit.package())?;
    let second_identity = cell_style_identity(&second_bytes, first)?;
    assert_eq!(
        second_identity, first_identity,
        "exclusive edit reuses key/object/parent"
    );
    let second_reopened = Package::from_bytes(&second_bytes)?;
    let effective = second_reopened.table_cell_text_style(0, 0, first)?;
    assert_eq!(effective.size, Some(73.0));
    assert_eq!(effective.italic, Some(true));
    let restored_first = second_commit
        .package()
        .apply_table_cell_text_style(&second_commit.patch().inverse())?;
    assert_eq!(bytes(restored_first.package())?, first_bytes);
    let (unknown_bytes, unknown_payload) =
        mutate_paragraph_style_with_unknowns(&second_bytes, second_identity.1)?;
    let unknown_fields = WireView::parse(&unknown_payload, WireLimits::DEFAULT)?;
    assert_eq!(
        unknown_fields
            .fields()
            .find(|field| field.number() == 12)
            .map(|field| field.payload()),
        Some(&[8, 1][..]),
        "unknown fixture has nonempty paragraph data"
    );
    assert!(
        unknown_fields.fields().any(|field| field.number() == 99),
        "unknown fixture has a root unknown field"
    );
    let unknown_package = Package::from_bytes(&unknown_bytes)?;
    let mut unknown_change = CharacterProperties::new();
    unknown_change.color = Some([12, 34, 56]);
    let unknown_commit = unknown_package
        .edit_table_cell_text_style(0, 0)?
        .set(first, unknown_change)?
        .commit()?;
    let unknown_candidate_bytes = bytes(unknown_commit.package())?;
    assert_eq!(
        paragraph_style_payload(&unknown_candidate_bytes, second_identity.1)?,
        unknown_payload,
        "unknown/unrelated paragraph fields force fresh derivation and survive"
    );
    let unknown_identity = cell_style_identity(&unknown_candidate_bytes, first)?;
    assert_ne!(unknown_identity.1, second_identity.1);
    assert_eq!(
        Package::from_bytes(&unknown_candidate_bytes)?
            .table_cell_text_style(0, 0, first)?
            .color,
        Some([12, 34, 56])
    );
    let restored_unknown = unknown_commit
        .package()
        .apply_table_cell_text_style(&unknown_commit.patch().inverse())?;
    assert_eq!(bytes(restored_unknown.package())?, unknown_bytes);
    let mut requested = CharacterProperties::new();
    requested.size = Some(73.0);
    let commit = package
        .edit_table_cell_text_style(0, 0)?
        .set(first, requested.clone())?
        .set(second, requested.clone())?
        .commit()?;
    let candidate_bytes = bytes(commit.package())?;
    assert_eq!(
        paragraph_style_payload(&candidate_bytes, generated.style_id)?,
        generated.style_payload,
        "unknown and unrelated paragraph fields on the shared owner survive safe derivation"
    );
    assert!(
        !style_entries(&candidate_bytes)?
            .iter()
            .any(|(candidate, _)| *candidate == key),
        "the shared source key is released after both cells move"
    );
    let reopened = Package::from_bytes(&candidate_bytes)?;
    for position in [first, second] {
        assert_eq!(
            reopened.table_cell_text_style(0, 0, position)?.size,
            Some(73.0)
        );
    }
    assert_eq!(
        reopened.table_cell_text_style(0, 0, sibling)?,
        before_sibling,
        "an unrelated cell remains unchanged"
    );
    let restored = commit
        .package()
        .apply_table_cell_text_style(&commit.patch().inverse())?;
    assert_eq!(
        bytes(restored.package())?,
        source,
        "inverse is exact-source"
    );
    let current = reopened.table_cell_text_style(0, 0, first)?;
    let noop = reopened
        .edit_table_cell_text_style(0, 0)?
        .set(first, current)?
        .commit()?;
    assert_eq!(
        bytes(noop.package())?,
        candidate_bytes,
        "setting the effective style is an exact no-op"
    );
    Ok(())
}

#[test]
fn exclusive_existing_style_batch_rewrites_each_cell() -> TestResult {
    let package = Package::new()
        .edit_table_cells(0, 0)?
        .set(At::new(1, 1), Input::Text("b".into()))?
        .commit()?
        .into_package();
    let package = package
        .edit_table_cells(0, 0)?
        .set(At::new(2, 1), Input::Text("c".into()))?
        .commit()?
        .into_package();
    let mut bold = CharacterProperties::new();
    bold.bold = Some(true);
    let styled = package
        .edit_table_cell_text_style(0, 0)?
        .set(At::new(1, 1), bold)?
        .commit()?
        .into_package();
    let mut italic = CharacterProperties::new();
    italic.italic = Some(true);
    let styled = styled
        .edit_table_cell_text_style(0, 0)?
        .set(At::new(2, 1), italic)?
        .commit()?
        .into_package();
    let mut second = CharacterProperties::new();
    second.italic = Some(true);
    let mut first = CharacterProperties::new();
    first.bold = Some(true);
    let commit = styled
        .edit_table_cell_text_style(0, 0)?
        .set(At::new(1, 1), second)?
        .set(At::new(2, 1), first)?
        .commit()?;
    assert_eq!(
        commit
            .package()
            .table_cell_text_style(0, 0, At::new(1, 1))?
            .italic,
        Some(true)
    );
    assert_eq!(
        commit
            .package()
            .table_cell_text_style(0, 0, At::new(2, 1))?
            .bold,
        Some(true)
    );
    let restored = commit
        .package()
        .apply_table_cell_text_style(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&styled)?);
    Ok(())
}

#[test]
fn refusals_are_typed() -> TestResult {
    let package = Package::new();
    assert!(matches!(
        package
            .edit_table_cell_text_style(0, 0)?
            .set(At::new(0, 0), CharacterProperties::new()),
        Err(Error::Refused(
            litchi_numbers::cell::text_style::Refusal::EmptyStyle
        ))
    ));
    assert!(matches!(
        package
            .edit_table_cell_text_style(0, 0)?
            .set(At::new(99, 0), bold()),
        Err(Error::Refused(
            litchi_numbers::cell::text_style::Refusal::CellNotFound
        ))
    ));
    // An empty cell takes a style-only cell (record 0370).
    let commit = package
        .edit_table_cell_text_style(0, 0)?
        .set(At::new(3, 3), bold())?
        .commit()?;
    assert_eq!(
        commit
            .package()
            .table_cell_text_style(0, 0, At::new(3, 3))?
            .bold,
        Some(true)
    );
    assert!(matches!(
        package.edit_table_cell_text_style(0, 0)?.commit(),
        Err(Error::Refused(
            litchi_numbers::cell::text_style::Refusal::NoStagedOperation
        ))
    ));
    Ok(())
}

#[test]
fn setting_effective_style_as_a_noop_preserves_exact_source() -> TestResult {
    let package = Package::new()
        .edit_table_cells(0, 0)?
        .set(At::new(1, 1), Input::Text("source".into()))?
        .commit()?
        .into_package();
    let source = bytes(&package)?;
    let mut current = CharacterProperties::new();
    current.bold = Some(false);
    let commit = package
        .edit_table_cell_text_style(0, 0)?
        .set(At::new(1, 1), current)?
        .commit()?;
    assert_eq!(bytes(commit.package())?, source);
    let restored = commit
        .package()
        .apply_table_cell_text_style(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn all_noop_text_style_batch_keeps_exact_source_and_watermark() -> TestResult {
    let package = Package::new()
        .edit_table_cells(0, 0)?
        .set(At::new(1, 1), Input::Text("one".into()))?
        .set(At::new(2, 1), Input::Text("two".into()))?
        .commit()?
        .into_package();
    let mut first = CharacterProperties::new();
    first.bold = Some(false);
    let second = first.clone();
    let source = bytes(&package)?;
    let commit = package
        .edit_table_cell_text_style(0, 0)?
        .set(At::new(1, 1), first)?
        .set(At::new(2, 1), second)?
        .commit()?;
    assert_eq!(bytes(commit.package())?, source);
    Ok(())
}

#[test]
fn mixed_noop_batch_allocates_only_changed_style() -> TestResult {
    let generated = shared_text_style_source()?;
    let source = generated.bytes;
    let first = At::new(generated.positions[0].0, generated.positions[0].1);
    let second = At::new(generated.positions[1].0, generated.positions[1].1);
    let package = Package::from_bytes(&source)?;
    let first_before = cell_style_identity(&source, first)?;
    let second_before = cell_style_identity(&source, second)?;
    let noop = package.table_cell_text_style(0, 0, first)?;
    let mut changed = CharacterProperties::new();
    changed.size = Some(73.0);
    let commit = package
        .edit_table_cell_text_style(0, 0)?
        .set(first, noop)?
        .set(second, changed)?
        .commit()?;
    let output = bytes(commit.package())?;
    assert_eq!(cell_style_identity(&output, first)?, first_before);
    let second_after = cell_style_identity(&output, second)?;
    assert_ne!(second_after.0, second_before.0);
    assert_ne!(second_after.1, second_before.1);
    assert_eq!(
        Package::from_bytes(&output)?
            .table_cell_text_style(0, 0, second)?
            .size,
        Some(73.0)
    );
    let restored = commit
        .package()
        .apply_table_cell_text_style(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    Ok(())
}

/// Every stored cell of the corpus takes bold, or refuses typed.
#[test]
fn corpus_cells_take_bold_or_refuse_typed() -> TestResult {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/numbers");
    let mut styled = 0_usize;
    let mut refused = Vec::new();
    let mut paths: Vec<PathBuf> = std::fs::read_dir(root)?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<Result<_, _>>()?;
    paths.retain(|path| {
        path.extension()
            .is_some_and(|extension| extension == "numbers")
    });
    paths.sort();
    for path in paths {
        let package = Package::open(&path)?;
        let source = bytes(&package)?;
        let Some(cell) = package.sheets()[0]
            .tables()
            .next()
            .and_then(|table| table.iter_cells().next())
            .map(|cell| cell.position())
        else {
            continue;
        };
        match package
            .edit_table_cell_text_style(0, 0)?
            .set(cell, bold())?
            .commit()
        {
            Ok(commit) => {
                assert_eq!(
                    commit.package().table_cell_text_style(0, 0, cell)?.bold,
                    Some(true),
                    "{}",
                    path.display()
                );
                let restored = commit
                    .package()
                    .apply_table_cell_text_style(&commit.patch().inverse())?;
                assert_eq!(bytes(restored.package())?, source);
                styled += 1;
            },
            Err(
                error @ Error::Refused(
                    litchi_numbers::cell::text_style::Refusal::UnsupportedDependency,
                ),
            ) => {
                refused.push((path, error));
            },
            Err(error) => panic!("{}: {error:?}", path.display()),
        }
    }
    eprintln!("styled {styled}; refused {refused:?}");
    assert!(styled >= 25, "only {styled} styled");
    Ok(())
}

/// A created workbook authors formulas (request 18, record 0309): its
/// table owner carries the native table extents.
#[test]
fn formulas_work_on_a_created_workbook() -> TestResult {
    use litchi_numbers::cell::Value as CellValue;
    use litchi_numbers::formula::Expression;
    let package = Package::new()
        .edit_table_cells(0, 0)?
        .set(At::new(1, 1), Input::number(2.0)?)?
        .commit()?
        .into_package();
    let commit = package
        .edit_table_cells(0, 0)?
        .set_formula(At::new(2, 1), Expression::parse("B2*3")?)?
        .commit()?;
    let reopened = Package::from_bytes(&bytes(commit.package())?)?;
    let table = reopened.sheets()[0].tables().next().expect("table");
    assert!(matches!(
        table.get(At::new(2, 1)),
        Some(CellValue::Formula(_))
    ));
    Ok(())
}
