use std::collections::HashMap;
use std::env;

use litchi_iwa::IWorkPackage;
use litchi_iwa::protobuf::{tsd, tst};
use prost::Message;

#[allow(deprecated)]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args()
        .nth(1)
        .ok_or("usage: inspect_numbers_storage <file>")?;
    let package = IWorkPackage::open(path)?;
    let mut objects = HashMap::new();
    for name in package.entry_names().filter(|name| name.ends_with(".iwa")) {
        let archive = package.archive(name)?;
        for object in archive.objects {
            let Some(identifier) = object.archive_info.identifier else {
                continue;
            };
            objects.insert(identifier, (name.to_owned(), object));
        }
    }

    for (info_id, (_, object)) in &objects {
        for message in &object.messages {
            let Ok(info) = tst::TableInfoArchive::decode(message.data.as_slice()) else {
                continue;
            };
            if !objects.contains_key(&info.table_model.identifier) {
                continue;
            }
            println!(
                "table_info={info_id} model={} parent={:?} comment={:?} pencils={:?} title={:?} caption={:?} editing={:?} summary={:?} category_order={:?} view_uids={:?} pivot_model={:?} pivot_order={:?}",
                info.table_model.identifier,
                info.super_.parent.as_ref().map(|value| value.identifier),
                info.super_.comment.as_ref().map(|value| value.identifier),
                info.super_
                    .pencil_annotations
                    .iter()
                    .map(|value| value.identifier)
                    .collect::<Vec<_>>(),
                info.super_.title.as_ref().map(|value| value.identifier),
                info.super_.caption.as_ref().map(|value| value.identifier),
                info.editing_state.as_ref().map(|value| value.identifier),
                info.summary_model.as_ref().map(|value| value.identifier),
                info.category_order.as_ref().map(|value| value.identifier),
                info.view_column_row_uids
                    .as_ref()
                    .map(|value| value.identifier),
                info.pivot_data_model.as_ref().map(|value| value.identifier),
                info.pivot_order.as_ref().map(|value| value.identifier),
            );
        }
    }

    for (model_id, (_, object)) in &objects {
        for message in &object.messages {
            if message.type_ != 6000 && message.type_ != 6001 {
                continue;
            }
            let Ok(model) = tst::TableModelArchive::decode(message.data.as_slice()) else {
                continue;
            };
            println!(
                "model={model_id} table_id={:?} name={:?} title=(visible={:?}, height={:?}, outlined={:?}) rows={} cols={} tile_size={:?} string_table={} formula_table={} formula_error_table={:?} rich_text_table={:?} row_buckets={:?} column_headers={} next_strips=({}, {}) uid_map={:?} stroke_sidecar={:?}",
                model.table_id,
                model.table_name,
                model.table_name_enabled,
                model.table_name_height,
                model.table_name_border_enabled,
                model.number_of_rows,
                model.number_of_columns,
                model.base_data_store.tiles.tile_size,
                model.base_data_store.string_table.identifier,
                model.base_data_store.formula_table.identifier,
                model
                    .base_data_store
                    .formula_error_table
                    .as_ref()
                    .map(|reference| reference.identifier),
                model
                    .base_data_store
                    .rich_text_table
                    .as_ref()
                    .map(|reference| reference.identifier),
                model
                    .base_data_store
                    .row_headers
                    .buckets
                    .iter()
                    .map(|reference| reference.identifier)
                    .collect::<Vec<_>>(),
                model.base_data_store.column_headers.identifier,
                model.base_data_store.next_row_strip_id,
                model.base_data_store.next_column_strip_id,
                model
                    .base_column_row_uids
                    .as_ref()
                    .map(|reference| reference.identifier),
                model
                    .stroke_sidecar
                    .as_ref()
                    .map(|reference| reference.identifier)
            );
            for header_id in model
                .base_data_store
                .row_headers
                .buckets
                .iter()
                .map(|reference| reference.identifier)
                .chain(std::iter::once(
                    model.base_data_store.column_headers.identifier,
                ))
            {
                let Some((archive_name, header_object)) = objects.get(&header_id) else {
                    continue;
                };
                for header_message in &header_object.messages {
                    if let Ok(bucket) =
                        tst::HeaderStorageBucket::decode(header_message.data.as_slice())
                    {
                        println!(
                            " header={} archive={} hash={} entries={:?}",
                            header_id,
                            archive_name,
                            bucket.bucket_hash_function,
                            bucket
                                .headers
                                .iter()
                                .map(|header| (
                                    header.index,
                                    header.size,
                                    header.hiding_state,
                                    header.number_of_cells
                                ))
                                .collect::<Vec<_>>()
                        );
                    }
                }
            }
            if let Some(uid_id) = model
                .base_column_row_uids
                .as_ref()
                .map(|reference| reference.identifier)
                && let Some((archive_name, uid_object)) = objects.get(&uid_id)
            {
                for uid_message in &uid_object.messages {
                    if let Ok(map) =
                        tst::ColumnRowUidMapArchive::decode(uid_message.data.as_slice())
                    {
                        println!(
                            " uid_map={} archive={} columns=({},{},{}) rows=({},{},{})",
                            uid_id,
                            archive_name,
                            map.sorted_column_uids.len(),
                            map.column_index_for_uid.len(),
                            map.column_uid_for_index.len(),
                            map.sorted_row_uids.len(),
                            map.row_index_for_uid.len(),
                            map.row_uid_for_index.len()
                        );
                        println!(
                            "  column_index_for_uid={:?} column_uid_for_index={:?}",
                            map.column_index_for_uid, map.column_uid_for_index
                        );
                        println!(
                            "  row_index_for_uid={:?} row_uid_for_index={:?}",
                            map.row_index_for_uid, map.row_uid_for_index
                        );
                    }
                }
            }
            if let Some(stroke_id) = model
                .stroke_sidecar
                .as_ref()
                .map(|reference| reference.identifier)
                && let Some((archive_name, stroke_object)) = objects.get(&stroke_id)
            {
                for stroke_message in &stroke_object.messages {
                    if let Ok(stroke) =
                        tst::StrokeSidecarArchive::decode(stroke_message.data.as_slice())
                    {
                        println!(
                            " stroke={} archive={} rows={:?} columns={:?} layers=({},{},{},{})",
                            stroke_id,
                            archive_name,
                            stroke.row_count,
                            stroke.column_count,
                            stroke.left_column_stroke_layers.len(),
                            stroke.right_column_stroke_layers.len(),
                            stroke.top_row_stroke_layers.len(),
                            stroke.bottom_row_stroke_layers.len()
                        );
                    }
                }
            }
            if let Some((archive_name, string_object)) =
                objects.get(&model.base_data_store.string_table.identifier)
            {
                println!(" string table archive={archive_name}");
                for string_message in &string_object.messages {
                    if let Ok(list) = tst::TableDataList::decode(string_message.data.as_slice()) {
                        println!(
                            "  list_type={} next={} segmented={} entries={}",
                            list.list_type,
                            list.next_list_id,
                            list.segments.len(),
                            list.entries.len()
                        );
                        for entry in list.entries {
                            println!(
                                "   key={} refs={} string={:?}",
                                entry.key, entry.refcount, entry.string
                            );
                        }
                    }
                }
            }
            if let Some(rich_text_id) = model
                .base_data_store
                .rich_text_table
                .as_ref()
                .map(|reference| reference.identifier)
                && let Some((archive_name, rich_text_object)) = objects.get(&rich_text_id)
            {
                println!(" rich text table archive={archive_name}");
                for rich_text_message in &rich_text_object.messages {
                    if let Ok(list) = tst::TableDataList::decode(rich_text_message.data.as_slice())
                    {
                        for entry in list.entries {
                            let payload = entry
                                .rich_text_payload
                                .as_ref()
                                .map(|reference| reference.identifier);
                            let storage = payload
                                .and_then(|identifier| objects.get(&identifier))
                                .and_then(|(_, object)| {
                                    object.messages.iter().find_map(|message| {
                                        tst::RichTextPayloadArchive::decode(message.data.as_slice())
                                            .ok()
                                    })
                                })
                                .map(|payload| payload.storage.identifier);
                            println!(
                                "  rich key={} refs={} payload={payload:?} storage={storage:?}",
                                entry.key, entry.refcount
                            );
                        }
                    }
                }
            }
            if let Some(error_table_id) = model
                .base_data_store
                .formula_error_table
                .as_ref()
                .map(|reference| reference.identifier)
                && let Some((archive_name, error_object)) = objects.get(&error_table_id)
            {
                println!(" formula error table archive={archive_name}");
                for error_message in &error_object.messages {
                    if let Ok(list) = tst::TableDataList::decode(error_message.data.as_slice()) {
                        println!(
                            "  list_type={} next={} segmented={} entries={}",
                            list.list_type,
                            list.next_list_id,
                            list.segments.len(),
                            list.entries.len()
                        );
                        for entry in list.entries {
                            println!(
                                "   error key={} refs={} reference={:?} string={:?}",
                                entry.key,
                                entry.refcount,
                                entry.reference.map(|reference| reference.identifier),
                                entry.string
                            );
                        }
                    }
                }
            }
            if let Some(comment_table_id) = model
                .base_data_store
                .comment_storage_table
                .as_ref()
                .map(|reference| reference.identifier)
                && let Some((archive_name, comment_object)) = objects.get(&comment_table_id)
            {
                println!(" comment table archive={archive_name}");
                for comment_message in &comment_object.messages {
                    if let Ok(list) = tst::TableDataList::decode(comment_message.data.as_slice()) {
                        println!(
                            "  list_type={} next={} segmented={} entries={}",
                            list.list_type,
                            list.next_list_id,
                            list.segments.len(),
                            list.entries.len()
                        );
                        for entry in list.entries {
                            let reference = entry
                                .comment_storage
                                .as_ref()
                                .map(|reference| reference.identifier);
                            let payloads = reference
                                .and_then(|identifier| objects.get(&identifier))
                                .map(|(archive, object)| {
                                    (
                                        archive,
                                        object
                                            .messages
                                            .iter()
                                            .map(|message| {
                                                let comment = tsd::CommentStorageArchive::decode(
                                                    message.data.as_slice(),
                                                )
                                                .ok()
                                                .map(|comment| {
                                                    (
                                                        comment.text,
                                                        comment.creation_date,
                                                        comment
                                                            .author
                                                            .map(|author| author.identifier),
                                                        comment
                                                            .replies
                                                            .iter()
                                                            .map(|reply| reply.identifier)
                                                            .collect::<Vec<_>>(),
                                                        comment.storage_uuid,
                                                    )
                                                });
                                                (message.type_, comment)
                                            })
                                            .collect::<Vec<_>>(),
                                    )
                                });
                            println!(
                                "   comment key={} refs={} reference={reference:?} payloads={payloads:?}",
                                entry.key, entry.refcount
                            );
                        }
                    }
                }
            }
            let tile_size = model.base_data_store.tiles.tile_size.unwrap_or(256);
            for tile_ref in model.base_data_store.tiles.tiles {
                let tile_id = tile_ref.tile.identifier;
                let (archive_name, tile_object) = objects
                    .get(&tile_id)
                    .ok_or("Numbers tile object is missing")?;
                println!(" tile={tile_id} archive={archive_name}");
                for tile_message in &tile_object.messages {
                    let Ok(tile) = tst::Tile::decode(tile_message.data.as_slice()) else {
                        continue;
                    };
                    println!(
                        "  max_column={} max_row={} cells={} rows={} storage={:?} bnc={:?} wide={:?}",
                        tile.max_column,
                        tile.max_row,
                        tile.num_cells,
                        tile.numrows,
                        tile.storage_version,
                        tile.last_saved_in_bnc,
                        tile.should_use_wide_rows
                    );
                    for row in tile.row_infos {
                        let storage = row
                            .cell_storage_buffer
                            .as_deref()
                            .unwrap_or(&row.cell_storage_buffer_pre_bnc);
                        let offsets = row
                            .cell_offsets
                            .as_deref()
                            .unwrap_or(&row.cell_offsets_pre_bnc);
                        for (column, cell) in
                            split_cells(storage, offsets, row.has_wide_offsets.unwrap_or(false))?
                        {
                            if let Some(comment_id) = bnc_u32_field(cell, 0x080000) {
                                println!(
                                    "   comment cell=({}, {}) id={comment_id}",
                                    tile_ref.tileid * tile_size + row.tile_row_index,
                                    column
                                );
                            }
                        }
                        println!(
                            "   row={} count={} storage_version={:?} wide={:?} offsets={} storage={}",
                            row.tile_row_index,
                            row.cell_count,
                            row.storage_version,
                            row.has_wide_offsets,
                            hex(offsets),
                            hex(storage)
                        );
                    }
                }
            }
        }
    }
    for (object_id, (archive_name, object)) in &objects {
        for message in &object.messages {
            match message.type_ {
                6370 => {
                    if let Ok(pivot) = tst::PivotOwnerArchive::decode(message.data.as_slice()) {
                        println!(
                            "pivot_owner={object_id} archive={archive_name} uid={:?} source_uid={:?} source_name={:?} row_groups={:?} column_groups={:?} aggregates={:?}",
                            pivot.pivot_owner_uid,
                            pivot.source_table_uid,
                            pivot.source_table_name,
                            pivot.grouping_columns_for_rows.as_ref().map(|list| list
                                .group_column
                                .iter()
                                .map(|column| (
                                    &column.column_uid,
                                    column.grouping_type,
                                    column.grouping_column_uid.as_ref()
                                ))
                                .collect::<Vec<_>>()),
                            pivot.grouping_columns_for_columns.as_ref().map(|list| list
                                .group_column
                                .iter()
                                .map(|column| (
                                    &column.column_uid,
                                    column.grouping_type,
                                    column.grouping_column_uid.as_ref()
                                ))
                                .collect::<Vec<_>>()),
                            pivot.aggregate_columns.as_ref().map(|list| list
                                .aggregates
                                .iter()
                                .map(|column| (
                                    &column.column_uid,
                                    column.level,
                                    column.agg_type,
                                    column.show_as_type,
                                    column.column_aggregate_uid.as_ref()
                                ))
                                .collect::<Vec<_>>()),
                        );
                    }
                },
                6373 => {
                    if let Ok(group) = tst::GroupByArchive::decode(message.data.as_slice()) {
                        println!(
                            "group_by={object_id} archive={archive_name} uid={:?} enabled={} owner_index={:?} columns={:?} aggregate_types={:?} formulas={:?} root_uid={:?} root_ref={:?} aggregators={:?} aggregator_refs={:?}",
                            group.group_by_uid,
                            group.is_enabled,
                            group.owner_index,
                            group
                                .group_column
                                .iter()
                                .map(|column| (
                                    &column.column_uid,
                                    column.grouping_type,
                                    column.grouping_column_uid.as_ref()
                                ))
                                .collect::<Vec<_>>(),
                            group.column_agg_type,
                            [
                                group.indirect_agg_type_change_formula.as_ref(),
                                group.grouping_columns_formula.as_ref(),
                                group.aggs_in_group_root_formula.as_ref(),
                                group.grouping_column_headers_formula.as_ref(),
                                group.column_order_changed_formula.as_ref(),
                                group.row_order_changed_formula.as_ref(),
                                group.row_order_changed_ignoring_recalc_formula.as_ref(),
                                group.hidden_states_changed_formula.as_ref(),
                            ],
                            group.group_node_root.as_ref().map(|root| &root.group_uid),
                            group
                                .group_node_root_ref
                                .as_ref()
                                .map(|reference| reference.identifier),
                            group
                                .aggregator
                                .iter()
                                .map(|aggregator| (
                                    &aggregator.column_uid,
                                    aggregator.agg_node.as_ref().map(|node| &node.formula_coord)
                                ))
                                .collect::<Vec<_>>(),
                            group
                                .aggregator_ref
                                .iter()
                                .map(|reference| reference.identifier)
                                .collect::<Vec<_>>(),
                        );
                    }
                },
                6382 => {
                    if let Ok(aggregator) =
                        tst::group_by_archive::AggregatorArchive::decode(message.data.as_slice())
                    {
                        println!(
                            "group_aggregator={object_id} archive={archive_name} column_uid={:?} formula={:?} children={}",
                            aggregator.column_uid,
                            aggregator.agg_node.as_ref().map(|node| &node.formula_coord),
                            aggregator
                                .agg_node
                                .as_ref()
                                .map_or(0, |node| node.child.len()),
                        );
                        if let Some(node) = &aggregator.agg_node {
                            print_aggregate_node(node, &mut Vec::new());
                        }
                    }
                },
                6383 => {
                    if let Ok(node) =
                        tst::group_by_archive::GroupNodeArchive::decode(message.data.as_slice())
                    {
                        println!(
                            "group_node={object_id} archive={archive_name} uid={:?} value={:?} rows={:?} aggregate_coords={:?} children={} child_refs={:?}",
                            node.group_uid,
                            node.group_cell_value.as_ref().map(compact_cell_value),
                            node.row_uid,
                            node.agg_formula_coords,
                            node.child.len(),
                            node.child_ref,
                        );
                    }
                },
                _ => {},
            }
        }
    }
    Ok(())
}

fn print_aggregate_node(node: &tst::group_by_archive::AggNodeArchive, path: &mut Vec<usize>) {
    println!(
        "  agg path={path:?} coord=({:?}, {:?}, packed={:?}) counts={:?} children={}",
        node.formula_coord.column,
        node.formula_coord.row,
        node.formula_coord.packed_data,
        node.accum.as_ref().map(|accumulator| (
            accumulator.bool_count,
            accumulator.number_count,
            accumulator.date_count,
            accumulator.duration_count,
            accumulator.string_count,
            accumulator.no_content_count,
        )),
        node.child.len(),
    );
    for (index, child) in node.child.iter().enumerate() {
        path.push(index);
        print_aggregate_node(child, path);
        path.pop();
    }
}

fn compact_cell_value(value: &litchi_iwa::protobuf::tsce::CellValueArchive) -> String {
    if let Some(string) = &value.string_value {
        return format!("string:{:?}", string.value);
    }
    if let Some(number) = &value.number_value {
        return format!("number:{:?}", number.value);
    }
    if let Some(boolean) = &value.boolean_value {
        return format!("bool:{}", boolean.value);
    }
    format!("{:?}", value.cell_value_type())
}

type IndexedCell<'a> = (usize, &'a [u8]);

fn split_cells<'a>(
    storage: &'a [u8],
    offsets: &[u8],
    wide: bool,
) -> std::io::Result<Vec<IndexedCell<'a>>> {
    let width = if wide { 4 } else { 1 };
    let starts = offsets
        .chunks_exact(2)
        .enumerate()
        .filter_map(|(column, bytes)| {
            let offset = u16::from_le_bytes([bytes[0], bytes[1]]);
            (offset != u16::MAX).then_some((column, usize::from(offset) * width))
        })
        .collect::<Vec<_>>();
    starts
        .iter()
        .enumerate()
        .map(|(index, &(column, start))| {
            let end = starts
                .get(index + 1)
                .map_or(storage.len(), |(_, offset)| *offset);
            let cell = storage.get(start..end).ok_or_else(|| {
                std::io::Error::new(std::io::ErrorKind::InvalidData, "invalid cell range")
            })?;
            Ok((column, cell))
        })
        .collect()
}

fn bnc_u32_field(cell: &[u8], target: u32) -> Option<u32> {
    if cell.first() != Some(&5) || cell.len() < 12 {
        return None;
    }
    let flags = u32::from_le_bytes(cell[8..12].try_into().ok()?);
    if flags & target == 0 {
        return None;
    }
    let mut cursor = 12;
    for (flag, size) in [
        (0x000001, 16),
        (0x000002, 8),
        (0x000004, 8),
        (0x000008, 4),
        (0x000010, 4),
        (0x000020, 4),
        (0x000040, 4),
        (0x000080, 4),
        (0x000100, 4),
        (0x000200, 4),
        (0x000400, 4),
        (0x000800, 4),
        (0x001000, 4),
        (0x002000, 4),
        (0x004000, 4),
        (0x008000, 4),
        (0x010000, 4),
        (0x020000, 4),
        (0x040000, 4),
        (0x080000, 4),
        (0x100000, 4),
    ] {
        if flags & flag == 0 {
            continue;
        }
        let end = cursor + size;
        let field = cell.get(cursor..end)?;
        if flag == target {
            return Some(u32::from_le_bytes(field.try_into().ok()?));
        }
        cursor = end;
    }
    None
}

fn hex(data: &[u8]) -> String {
    data.iter().map(|byte| format!("{byte:02x}")).collect()
}
