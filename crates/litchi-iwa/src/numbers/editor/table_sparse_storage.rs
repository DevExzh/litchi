//! Lazy allocation for sparse Numbers table storage.
//!
//! Numbers creates tile and row-header storage only when a cell in the
//! corresponding range receives content. Keeping that behavior avoids a large
//! graph of empty IWA objects for wide, source-created tables.

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};

use prost::Message;

use super::*;
use crate::wire::{
    patch_varint_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields, transform_length_delimited_field,
};

const DEFAULT_TILE_SIZE_ROWS: u32 = 256;
const TABLE_MODEL_DATA_STORE_FIELD: u32 = 4;
const DATA_STORE_ROW_HEADERS_FIELD: u32 = 1;
const DATA_STORE_TILES_FIELD: u32 = 3;
const DATA_STORE_NEXT_ROW_STRIP_ID_FIELD: u32 = 7;
const DATA_STORE_ROW_TILE_TREE_FIELD: u32 = 9;
const HEADER_STORAGE_BUCKETS_FIELD: u32 = 2;
const TILE_STORAGE_REFERENCES_FIELD: u32 = 1;
const TABLE_RB_TREE_NODES_FIELD: u32 = 1;

struct TileAllocation {
    key: u32,
    object_id: u64,
    archive_name: String,
}

struct HeaderBucketAllocation {
    object_id: u64,
    archive_name: String,
}

/// Allocate every sparse backing object needed by ordinary cell writes.
///
/// The operation is deliberately a no-op when all target tiles and row-header
/// buckets already exist, so an in-tile write preserves the existing object
/// graph byte-for-byte except for the cell itself.
pub(super) fn ensure_cell_storage(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<()> {
    ensure_cells_storage(package, table_id, &[(row, column)])
}

pub(super) fn ensure_attached_cell_storage(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<()> {
    ensure_attached_cells_storage(package, table_id, &[(row, column)])
}

pub(super) fn ensure_cells_storage(
    package: &mut IWorkPackage,
    table_id: u64,
    coordinates: &[(usize, usize)],
) -> Result<()> {
    let descriptor = table_models(package)?
        .into_iter()
        .find(|table| table.object_id == table_id)
        .ok_or_else(|| Error::ParseError(format!("Numbers table object {table_id} not found")))?;
    ensure_descriptor_storage(package, &descriptor, coordinates)
}

pub(super) fn ensure_attached_cells_storage(
    package: &mut IWorkPackage,
    table_id: u64,
    coordinates: &[(usize, usize)],
) -> Result<()> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    ensure_descriptor_storage(package, &descriptor, coordinates)
}

fn ensure_descriptor_storage(
    package: &mut IWorkPackage,
    descriptor: &model::TableDescriptor,
    coordinates: &[(usize, usize)],
) -> Result<()> {
    if coordinates.is_empty() {
        return Ok(());
    }

    let tile_size = descriptor
        .model
        .base_data_store
        .tiles
        .tile_size
        .unwrap_or(DEFAULT_TILE_SIZE_ROWS);
    if tile_size == 0 {
        return Err(Error::InvalidFormat(
            "Numbers table declares a zero tile size".to_owned(),
        ));
    }

    let mut requested_tiles = BTreeSet::new();
    let mut largest_header_bucket: Option<usize> = None;
    for &(row, column) in coordinates {
        validate_coordinate(&descriptor.model, row, column)?;
        let row_u32 = u32::try_from(row)
            .map_err(|_| Error::ParseError("Numbers row exceeds u32".to_owned()))?;
        requested_tiles.insert(row_u32 / tile_size);
        largest_header_bucket = Some(
            largest_header_bucket
                .unwrap_or_default()
                .max(row / HEADER_BUCKET_ROWS),
        );
    }

    let existing_tiles = tile_references_by_key(&descriptor.model)?;
    let missing_tiles = requested_tiles
        .into_iter()
        .filter(|key| !existing_tiles.contains_key(key))
        .collect::<Vec<_>>();
    let existing_bucket_count = descriptor.model.base_data_store.row_headers.buckets.len();
    let missing_bucket_count = largest_header_bucket
        .and_then(|index| index.checked_add(1))
        .and_then(|required| required.checked_sub(existing_bucket_count))
        .unwrap_or_default();
    if missing_tiles.is_empty() && missing_bucket_count == 0 {
        return Ok(());
    }

    let locations = object_locations(package)?;
    let tile_template = descriptor
        .model
        .base_data_store
        .tiles
        .tiles
        .first()
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table {:?} has no tile template",
                descriptor.model.table_name
            ))
        })?;
    let tile_template_archive = locations
        .get(&tile_template.tile.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers tile object {} is missing",
                tile_template.tile.identifier
            ))
        })?
        .to_owned();
    let tile_template_contents = package.archive(&tile_template_archive)?;
    let tile_template_object = tile_template_contents
        .object(tile_template.tile.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers tile object {} is missing",
                tile_template.tile.identifier
            ))
        })?;

    let header_template = descriptor
        .model
        .base_data_store
        .row_headers
        .buckets
        .first()
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table {:?} has no row-header template",
                descriptor.model.table_name
            ))
        })?;
    let header_template_archive = locations
        .get(&header_template.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers row-header object {} is missing",
                header_template.identifier
            ))
        })?
        .to_owned();
    let header_template_contents = package.archive(&header_template_archive)?;
    let header_template_object = header_template_contents
        .object(header_template.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers row-header object {} is missing",
                header_template.identifier
            ))
        })?;

    let mut next_identifier = next_object_identifier(package)?;
    let mut tiles = Vec::with_capacity(missing_tiles.len());
    for key in missing_tiles {
        let object_id = model::take_identifier(&mut next_identifier)?;
        let object = model::clone_empty_table_storage(
            tile_template_object,
            object_id,
            model::TableOwnedKind::Tile,
            descriptor.model.number_of_rows,
            descriptor.model.number_of_columns,
            descriptor.object_id,
        )?;
        package.update_archive(&tile_template_archive, |archive| {
            archive.insert_object(object)
        })?;
        tiles.push(TileAllocation {
            key,
            object_id,
            archive_name: tile_template_archive.clone(),
        });
    }

    let mut headers = Vec::with_capacity(missing_bucket_count);
    for _ in 0..missing_bucket_count {
        let object_id = model::take_identifier(&mut next_identifier)?;
        let object = model::clone_empty_table_storage(
            header_template_object,
            object_id,
            model::TableOwnedKind::Header,
            descriptor.model.number_of_rows,
            descriptor.model.number_of_columns,
            descriptor.object_id,
        )?;
        package.update_archive(&header_template_archive, |archive| {
            archive.insert_object(object)
        })?;
        headers.push(HeaderBucketAllocation {
            object_id,
            archive_name: header_template_archive.clone(),
        });
    }

    update_model_storage_references(package, descriptor, &locations, &tiles, &headers, tile_size)?;
    register_storage_objects(package, descriptor.object_id, &locations, &tiles, &headers)?;

    let last_identifier = next_identifier.checked_sub(1).ok_or_else(|| {
        Error::InvalidFormat("Numbers sparse storage allocated no identifiers".to_owned())
    })?;
    set_package_last_object_identifier(package, last_identifier)
}

fn validate_coordinate(model: &TableModelArchive, row: usize, column: usize) -> Result<()> {
    if row >= model.number_of_rows as usize || column >= model.number_of_columns as usize {
        return Err(Error::ParseError(format!(
            "Cell ({row}, {column}) is outside Numbers table {:?} dimensions {}x{}",
            model.table_name, model.number_of_rows, model.number_of_columns
        )));
    }
    Ok(())
}

fn tile_references_by_key(model: &TableModelArchive) -> Result<HashMap<u32, u64>> {
    let mut tiles = HashMap::with_capacity(model.base_data_store.tiles.tiles.len());
    for reference in &model.base_data_store.tiles.tiles {
        if reference.tile.identifier == 0 {
            return Err(Error::InvalidFormat(
                "Numbers table contains a zero tile object reference".to_owned(),
            ));
        }
        if let Some(previous) = tiles.insert(reference.tileid, reference.tile.identifier) {
            return Err(Error::InvalidFormat(format!(
                "Numbers table tile key {} maps to both {previous} and {}",
                reference.tileid, reference.tile.identifier
            )));
        }
    }
    Ok(tiles)
}

fn update_model_storage_references(
    package: &mut IWorkPackage,
    descriptor: &model::TableDescriptor,
    locations: &HashMap<u64, String>,
    tiles: &[TileAllocation],
    headers: &[HeaderBucketAllocation],
    tile_size: u32,
) -> Result<()> {
    let model_archive = locations
        .get(&descriptor.object_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table model object {} is missing",
                descriptor.object_id
            ))
        })?
        .to_owned();
    package.update_archive(&model_archive, |archive| {
        let object = archive.object_mut(descriptor.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table model object {} is missing",
                descriptor.object_id
            ))
        })?;
        let message_index = model::find_table_model_message(object)?;
        let message_type = object.messages[message_index].type_;
        let original = object.messages[message_index].data.clone();
        let previous = TableModelArchive::decode(original.as_slice())?;
        if previous != descriptor.model {
            return Err(Error::InvalidFormat(format!(
                "Numbers table model {} changed during sparse storage allocation",
                descriptor.object_id
            )));
        }
        let mut current = previous.clone();
        current
            .base_data_store
            .tiles
            .tiles
            .extend(tiles.iter().map(|tile| tst::tile_storage::Tile {
                tileid: tile.key,
                tile: tsp::Reference {
                    identifier: tile.object_id,
                    ..Default::default()
                },
            }));
        current
            .base_data_store
            .tiles
            .tiles
            .sort_by_key(|reference| reference.tileid);
        let tile_by_key = tile_references_by_key(&current)?;
        if tile_by_key.len() != current.base_data_store.tiles.tiles.len() {
            return Err(Error::InvalidFormat(
                "Numbers sparse tile allocation duplicated a tile key".to_owned(),
            ));
        }
        current
            .base_data_store
            .row_headers
            .buckets
            .extend(headers.iter().map(|header| tsp::Reference {
                identifier: header.object_id,
                ..Default::default()
            }));
        current.base_data_store.row_tile_tree.nodes = canonical_row_tile_tree(&current, tile_size)?;
        let next_row_strip_id = current
            .base_data_store
            .tiles
            .tiles
            .iter()
            .map(|reference| {
                reference.tileid.checked_add(1).ok_or_else(|| {
                    Error::ParseError("Numbers tile key exceeds row-strip capacity".to_owned())
                })
            })
            .collect::<Result<Vec<_>>>()?
            .into_iter()
            .max()
            .unwrap_or_default();
        current.base_data_store.next_row_strip_id = current
            .base_data_store
            .next_row_strip_id
            .max(next_row_strip_id);

        let data = rewrite_table_storage_wire(&original, &previous, &current)?;
        if TableModelArchive::decode(data.as_slice())? != current {
            return Err(Error::InvalidFormat(
                "Numbers sparse storage model wire mutation failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        let info = &mut object.archive_info.message_infos[message_index];
        extend_references(
            &mut info.object_references,
            tiles.iter().map(|tile| tile.object_id),
        );
        extend_references(
            &mut info.object_references,
            headers.iter().map(|header| header.object_id),
        );
        let existing_tile_ids = previous
            .base_data_store
            .tiles
            .tiles
            .iter()
            .map(|reference| reference.tile.identifier)
            .collect::<HashSet<_>>();
        let existing_header_ids = previous
            .base_data_store
            .row_headers
            .buckets
            .iter()
            .map(|reference| reference.identifier)
            .collect::<HashSet<_>>();
        for field in &mut info.field_infos {
            match field.path.path.as_slice() {
                [
                    TABLE_MODEL_DATA_STORE_FIELD,
                    DATA_STORE_TILES_FIELD,
                    TILE_STORAGE_REFERENCES_FIELD,
                    2,
                ] if field
                    .object_references
                    .iter()
                    .any(|identifier| existing_tile_ids.contains(identifier)) =>
                {
                    extend_references(
                        &mut field.object_references,
                        tiles.iter().map(|tile| tile.object_id),
                    );
                },
                [
                    TABLE_MODEL_DATA_STORE_FIELD,
                    DATA_STORE_ROW_HEADERS_FIELD,
                    HEADER_STORAGE_BUCKETS_FIELD,
                ] if field
                    .object_references
                    .iter()
                    .any(|identifier| existing_header_ids.contains(identifier)) =>
                {
                    extend_references(
                        &mut field.object_references,
                        headers.iter().map(|header| header.object_id),
                    );
                },
                _ => {},
            }
        }
        Ok(())
    })
}

fn canonical_row_tile_tree(
    model: &TableModelArchive,
    tile_size: u32,
) -> Result<Vec<tst::table_rb_tree::Node>> {
    model
        .base_data_store
        .tiles
        .tiles
        .iter()
        .enumerate()
        .map(|(index, reference)| {
            let key = reference.tileid.checked_mul(tile_size).ok_or_else(|| {
                Error::ParseError("Numbers tile row origin overflows u32".to_owned())
            })?;
            let value = u32::try_from(index)
                .map_err(|_| Error::ParseError("Numbers tile index exceeds u32".to_owned()))?;
            Ok(tst::table_rb_tree::Node { key, value })
        })
        .collect()
}

fn rewrite_table_storage_wire(
    original: &[u8],
    previous: &TableModelArchive,
    current: &TableModelArchive,
) -> Result<Vec<u8>> {
    transform_length_delimited_field(original, TABLE_MODEL_DATA_STORE_FIELD, |store| {
        let mut data =
            transform_length_delimited_field(store, DATA_STORE_ROW_HEADERS_FIELD, |headers| {
                rewrite_references_wire(
                    headers,
                    HEADER_STORAGE_BUCKETS_FIELD,
                    &previous.base_data_store.row_headers.buckets,
                    &current.base_data_store.row_headers.buckets,
                )
            })?;
        data = transform_length_delimited_field(&data, DATA_STORE_TILES_FIELD, |tiles| {
            rewrite_tile_references_wire(
                tiles,
                &previous.base_data_store.tiles.tiles,
                &current.base_data_store.tiles.tiles,
            )
        })?;
        data = patch_varint_field(
            &data,
            DATA_STORE_NEXT_ROW_STRIP_ID_FIELD,
            true,
            Some(u64::from(current.base_data_store.next_row_strip_id)),
        )?;
        transform_length_delimited_field(&data, DATA_STORE_ROW_TILE_TREE_FIELD, |tree| {
            rewrite_tree_nodes_wire(
                tree,
                &previous.base_data_store.row_tile_tree.nodes,
                &current.base_data_store.row_tile_tree.nodes,
            )
        })
    })
}

fn rewrite_references_wire(
    original: &[u8],
    field: u32,
    previous: &[tsp::Reference],
    current: &[tsp::Reference],
) -> Result<Vec<u8>> {
    let raw = repeated_length_delimited_payloads(original, field)?;
    if raw.len() != previous.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers reference field {field} has {} raw values but {} decoded values",
            raw.len(),
            previous.len()
        )));
    }
    let mut existing = HashMap::with_capacity(previous.len());
    for (reference, payload) in previous.iter().zip(raw) {
        if tsp::Reference::decode(payload)? != *reference {
            return Err(Error::InvalidFormat(format!(
                "Numbers reference field {field} changed during sparse allocation"
            )));
        }
        if existing.insert(reference.identifier, payload).is_some() {
            return Err(Error::InvalidFormat(format!(
                "Numbers reference field {field} repeats object {}",
                reference.identifier
            )));
        }
    }
    let mut seen = HashSet::with_capacity(current.len());
    let replacements = current
        .iter()
        .map(|reference| {
            if !seen.insert(reference.identifier) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers reference field {field} would repeat object {}",
                    reference.identifier
                )));
            }
            Ok(existing
                .get(&reference.identifier)
                .map_or_else(|| reference.encode_to_vec(), |payload| payload.to_vec()))
        })
        .collect::<Result<Vec<_>>>()?;
    rewrite_repeated_length_delimited_fields(original, field, &replacements)
}

fn rewrite_tile_references_wire(
    original: &[u8],
    previous: &[tst::tile_storage::Tile],
    current: &[tst::tile_storage::Tile],
) -> Result<Vec<u8>> {
    let raw = repeated_length_delimited_payloads(original, TILE_STORAGE_REFERENCES_FIELD)?;
    if raw.len() != previous.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers tile storage has {} raw references but {} decoded references",
            raw.len(),
            previous.len()
        )));
    }
    let mut existing = HashMap::with_capacity(previous.len());
    for (reference, payload) in previous.iter().zip(raw) {
        if tst::tile_storage::Tile::decode(payload)? != *reference {
            return Err(Error::InvalidFormat(
                "Numbers tile reference changed during sparse allocation".to_owned(),
            ));
        }
        let key = (reference.tileid, reference.tile.identifier);
        if existing.insert(key, payload).is_some() {
            return Err(Error::InvalidFormat(format!(
                "Numbers tile storage repeats tile key {}",
                reference.tileid
            )));
        }
    }
    let mut seen = HashSet::with_capacity(current.len());
    let replacements = current
        .iter()
        .map(|reference| {
            let key = (reference.tileid, reference.tile.identifier);
            if !seen.insert(key) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers tile storage would repeat tile key {}",
                    reference.tileid
                )));
            }
            Ok(existing
                .get(&key)
                .map_or_else(|| reference.encode_to_vec(), |payload| payload.to_vec()))
        })
        .collect::<Result<Vec<_>>>()?;
    rewrite_repeated_length_delimited_fields(original, TILE_STORAGE_REFERENCES_FIELD, &replacements)
}

fn rewrite_tree_nodes_wire(
    original: &[u8],
    previous: &[tst::table_rb_tree::Node],
    current: &[tst::table_rb_tree::Node],
) -> Result<Vec<u8>> {
    let raw = repeated_length_delimited_payloads(original, TABLE_RB_TREE_NODES_FIELD)?;
    if raw.len() != previous.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers row-tile tree has {} raw nodes but {} decoded nodes",
            raw.len(),
            previous.len()
        )));
    }
    let mut existing = HashMap::with_capacity(previous.len());
    for (node, payload) in previous.iter().zip(raw) {
        if tst::table_rb_tree::Node::decode(payload)? != *node {
            return Err(Error::InvalidFormat(
                "Numbers row-tile tree changed during sparse allocation".to_owned(),
            ));
        }
        if existing.insert((node.key, node.value), payload).is_some() {
            return Err(Error::InvalidFormat(format!(
                "Numbers row-tile tree repeats node key {}",
                node.key
            )));
        }
    }
    let mut seen = HashSet::with_capacity(current.len());
    let replacements = current
        .iter()
        .map(|node| {
            let key = (node.key, node.value);
            if !seen.insert(key) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers row-tile tree would repeat node key {}",
                    node.key
                )));
            }
            Ok(existing
                .get(&key)
                .map_or_else(|| node.encode_to_vec(), |payload| payload.to_vec()))
        })
        .collect::<Result<Vec<_>>>()?;
    rewrite_repeated_length_delimited_fields(original, TABLE_RB_TREE_NODES_FIELD, &replacements)
}

fn register_storage_objects(
    package: &mut IWorkPackage,
    model_id: u64,
    locations: &HashMap<u64, String>,
    tiles: &[TileAllocation],
    headers: &[HeaderBucketAllocation],
) -> Result<()> {
    let model_archive = locations.get(&model_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers table model object {model_id} is missing"))
    })?;
    let mut component_objects = BTreeMap::<u64, Vec<u64>>::new();
    for (object_id, archive_name) in tiles
        .iter()
        .map(|tile| (tile.object_id, tile.archive_name.as_str()))
        .chain(
            headers
                .iter()
                .map(|header| (header.object_id, header.archive_name.as_str())),
        )
    {
        if let Some(component) = component_identifier_for_entry(package, archive_name)? {
            component_objects
                .entry(component)
                .or_default()
                .push(object_id);
        }
        table_duplicate::register_numbers_component_reference(
            package,
            model_archive,
            archive_name,
            object_id,
        )?;
    }
    for (component, object_ids) in component_objects {
        add_component_object_uuids(package, component, &object_ids)?;
    }
    Ok(())
}

fn extend_references(references: &mut Vec<u64>, additions: impl IntoIterator<Item = u64>) {
    for identifier in additions {
        if !references.contains(&identifier) {
            references.push(identifier);
        }
    }
}
