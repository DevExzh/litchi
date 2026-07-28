//! CalculationEngine registration for volatile conditional-style formulas.

use prost::Message;

use super::*;

const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
const CELL_RECORD_TILE_MESSAGE_TYPE: u32 = 4_009;
const CONDITIONAL_STYLE_OWNER_KIND: u32 = 3;
const STYLE_STORAGE_OWNER_KIND: u32 = 35;
const REGISTERED_FORMULA_OWNER_COUNT: u64 = 2;
const STYLE_STORAGE_OWNER_OFFSET: u64 = 35;
const CONDITIONAL_BASE_OWNER_OFFSET: u64 = 3;
const STYLE_STORAGE_ROW: u32 = 0;
const STYLE_STORAGE_COLUMN: u32 = 2;
const WHOLE_ROW_COLUMN_SENTINEL: u32 = i16::MAX as u32;
const WHOLE_COLUMN_ROW_SENTINEL: u32 = i32::MAX as u32;
const NATIVE_MESSAGE_VERSION: &[u32] = &[3, 2, 10];

pub(super) fn ensure_volatile_owner(
    package: &mut IWorkPackage,
    table_uuid: &tsp::Uuid,
    conditional_owner_uid: tsp::Uuid,
    row: usize,
    column: usize,
) -> Result<()> {
    let Some(entry_name) = package.calculation_engine_entry_name()?.map(str::to_owned) else {
        return Ok(());
    };
    let archive = package.archive(&entry_name)?;
    let table_owner_uid = formula_owner_uuid_for_table(table_uuid);
    let table_internal_id = owner_internal_id(&archive, &table_owner_uid)?.ok_or_else(|| {
        Error::InvalidFormat("Numbers table has no CalculationEngine owner".to_owned())
    })?;
    let style_owner_uid = tsp::Uuid {
        lower: table_owner_uid
            .lower
            .checked_add(STYLE_STORAGE_OWNER_OFFSET)
            .ok_or_else(|| Error::ParseError("Numbers formula-owner UUID overflow".to_owned()))?,
        upper: table_owner_uid.upper,
    };
    if owner_internal_id(&archive, &conditional_owner_uid)?.is_some() {
        let style_internal_id =
            owner_internal_id(&archive, &style_owner_uid)?.ok_or_else(|| {
                Error::InvalidFormat(
                    "Numbers conditional-style storage owner is missing".to_owned(),
                )
            })?;
        let host = HostDependency::new(
            row,
            column,
            style_internal_id,
            table_internal_id,
            table_owner_uid,
        )?;
        return append_host(package, &entry_name, conditional_owner_uid, &host);
    }
    let style_internal_id = super::super::formula_clone::next_internal_owner_id(&archive)?;
    let conditional_internal_id = style_internal_id
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("Numbers formula-owner ID overflow".to_owned()))?;
    let host = HostDependency::new(
        row,
        column,
        style_internal_id,
        table_internal_id,
        table_owner_uid,
    )?;

    let first_id = next_object_identifier(package)?;
    let style_owner_id = first_id;
    let style_tile_id = checked_next_identifier(style_owner_id)?;
    let conditional_owner_id = checked_next_identifier(style_tile_id)?;
    let conditional_tile_id = checked_next_identifier(conditional_owner_id)?;
    let owner_ids = [style_owner_id, conditional_owner_id];
    let component_ids = [
        style_owner_id,
        style_tile_id,
        conditional_owner_id,
        conditional_tile_id,
    ];
    let style_owner = auxiliary_owner(
        style_owner_uid,
        style_internal_id,
        STYLE_STORAGE_OWNER_KIND,
        table_owner_uid,
        style_tile_id,
    );
    let conditional_owner = conditional_owner(
        conditional_owner_uid,
        conditional_internal_id,
        conditional_tile_id,
        &host,
    )?;
    let map_entries = [
        owner_map_entry(style_internal_id, style_owner_uid),
        owner_map_entry(conditional_internal_id, conditional_owner_uid),
    ];

    package.update_archive(&entry_name, |archive| {
        let (engine_id, message_index) =
            super::super::formula_clone::calculation_engine_location(archive)?;
        let engine = archive.object_mut(engine_id).ok_or_else(|| {
            Error::InvalidFormat("Numbers CalculationEngine root is missing".to_owned())
        })?;
        let original = engine.messages[message_index].clone();
        let data = super::super::formula_clone::append_formula_owners_to_engine(
            &original.data,
            &owner_ids,
            &map_entries,
            REGISTERED_FORMULA_OWNER_COUNT,
        )?;
        engine.replace_message(
            message_index,
            RawMessage {
                type_: CALCULATION_ENGINE_MESSAGE_TYPE,
                data,
            },
        )?;
        let references = &mut engine.archive_info.message_infos[message_index].object_references;
        for identifier in owner_ids {
            if !references.contains(&identifier) {
                references.push(identifier);
            }
        }

        insert_owner(archive, style_owner_id, style_owner, style_tile_id)?;
        insert_tile(
            archive,
            style_tile_id,
            cell_record_tile(style_internal_id, Vec::new()),
        )?;
        insert_owner(
            archive,
            conditional_owner_id,
            conditional_owner,
            conditional_tile_id,
        )?;
        insert_tile(
            archive,
            conditional_tile_id,
            cell_record_tile(conditional_internal_id, vec![host.record.clone()]),
        )
    })?;
    set_package_last_object_identifier(package, conditional_tile_id)?;
    if let Some(component) = component_identifier_for_entry(package, &entry_name)? {
        add_component_object_uuids(package, component, &component_ids)?;
    }
    Ok(())
}

fn append_host(
    package: &mut IWorkPackage,
    entry_name: &str,
    conditional_owner_uid: tsp::Uuid,
    host: &HostDependency,
) -> Result<()> {
    package.update_archive(entry_name, |archive| {
        let (owner_id, owner_message_index, mut owner) =
            owner_location(archive, &conditional_owner_uid)?.ok_or_else(|| {
                Error::InvalidFormat(
                    "Numbers conditional-style formula owner is missing".to_owned(),
                )
            })?;
        let owner_data = archive
            .object(owner_id)
            .and_then(|object| object.messages.get(owner_message_index))
            .map(|message| message.data.clone())
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "Numbers conditional-style formula owner disappeared".to_owned(),
                )
            })?;
        if owner.encode_to_vec() != owner_data {
            return Err(Error::InvalidFormat(
                "cannot safely extend a conditional-style owner with unknown wire fields"
                    .to_owned(),
            ));
        }
        if owner
            .cell_dependencies
            .as_ref()
            .is_some_and(|dependencies| {
                dependencies
                    .cell_record
                    .iter()
                    .any(|record| same_host(record, &host.record))
            })
        {
            return Ok(());
        }

        owner
            .cell_dependencies
            .get_or_insert_default()
            .cell_record
            .push(host.record.clone());
        merge_coordinate_set(
            owner
                .volatile_dependencies
                .get_or_insert_with(empty_volatile_dependencies)
                .volatile_time_cells
                .get_or_insert_default(),
            &host.coordinates,
        );
        let references = owner.uuid_references.get_or_insert_default();
        let table_ref = if let Some(index) = references
            .table_refs
            .iter()
            .position(|table_ref| table_ref.owner_uuid == host.table_owner_uid)
        {
            &mut references.table_refs[index]
        } else {
            references
                .table_refs
                .push(tsce::uuid_references_archive::TableRef {
                    owner_uuid: host.table_owner_uid,
                    coord_set: Some(tsce::CellCoordSetArchive::default()),
                });
            references
                .table_refs
                .last_mut()
                .expect("a table reference was just inserted")
        };
        merge_coordinate_set(
            table_ref.coord_set.get_or_insert_default(),
            &host.coordinates,
        );
        let tile_id = owner
            .tiled_cell_dependencies
            .as_ref()
            .and_then(|dependencies| dependencies.cell_record_tiles.first())
            .map(|reference| reference.identifier)
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "Numbers conditional-style dependency tile is missing".to_owned(),
                )
            })?;

        let owner_object = archive.object_mut(owner_id).ok_or_else(|| {
            Error::InvalidFormat("Numbers conditional-style formula owner disappeared".to_owned())
        })?;
        let owner_message_type = owner_object.messages[owner_message_index].type_;
        owner_object.replace_message(
            owner_message_index,
            RawMessage {
                type_: owner_message_type,
                data: owner.encode_to_vec(),
            },
        )?;

        let tile_object = archive.object_mut(tile_id).ok_or_else(|| {
            Error::InvalidFormat("Numbers conditional-style dependency tile is missing".to_owned())
        })?;
        let tile_message_index = tile_object
            .messages
            .iter()
            .position(|message| message.type_ == CELL_RECORD_TILE_MESSAGE_TYPE)
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "Numbers conditional-style dependency tile payload is missing".to_owned(),
                )
            })?;
        let mut tile = tsce::CellRecordTileArchive::decode(
            tile_object.messages[tile_message_index].data.as_slice(),
        )?;
        if tile.encode_to_vec() != tile_object.messages[tile_message_index].data {
            return Err(Error::InvalidFormat(
                "cannot safely extend a conditional-style tile with unknown wire fields".to_owned(),
            ));
        }
        tile.cell_records.push(host.record.clone());
        tile_object.replace_message(
            tile_message_index,
            RawMessage {
                type_: CELL_RECORD_TILE_MESSAGE_TYPE,
                data: tile.encode_to_vec(),
            },
        )?;
        Ok(())
    })
}

struct HostDependency {
    record: tsce::CellRecordExpandedArchive,
    coordinates: tsce::CellCoordSetArchive,
    table_owner_uid: tsp::Uuid,
}

impl HostDependency {
    fn new(
        row: usize,
        column: usize,
        style_internal_id: u32,
        table_internal_id: u32,
        table_owner_uid: tsp::Uuid,
    ) -> Result<Self> {
        let row = u32::try_from(row)
            .map_err(|_| Error::ParseError("conditional-highlight row exceeds u32".to_owned()))?;
        let row_index = i32::try_from(row).map_err(|_| {
            Error::ParseError(
                "conditional-highlight row exceeds the native signed index".to_owned(),
            )
        })?;
        let column = u32::try_from(column).map_err(|_| {
            Error::ParseError("conditional-highlight column exceeds u32".to_owned())
        })?;
        Ok(Self {
            record: tsce::CellRecordExpandedArchive {
                column,
                row,
                expanded_edges: Some(tsce::ExpandedEdgesArchive {
                    edge_with_owner_rows: vec![STYLE_STORAGE_ROW, row],
                    edge_with_owner_columns: vec![STYLE_STORAGE_COLUMN, column],
                    internal_owner_id_for_edge: vec![style_internal_id, table_internal_id],
                    ..Default::default()
                }),
                ..Default::default()
            },
            coordinates: coordinate_set(row_index, column),
            table_owner_uid,
        })
    }
}

fn owner_internal_id(archive: &Archive, uid: &tsp::Uuid) -> Result<Option<u32>> {
    archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .filter(|message| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
        .map(|message| tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice()))
        .find_map(|owner| match owner {
            Ok(owner) if owner.formula_owner_uid == *uid => {
                Some(Ok(owner.internal_formula_owner_id))
            },
            Ok(_) => None,
            Err(error) => Some(Err(error.into())),
        })
        .transpose()
}

fn owner_location(
    archive: &Archive,
    uid: &tsp::Uuid,
) -> Result<Option<(u64, usize, tsce::FormulaOwnerDependenciesArchive)>> {
    for object in &archive.objects {
        for (message_index, message) in object.messages.iter().enumerate() {
            if message.type_ != FORMULA_OWNER_MESSAGE_TYPE {
                continue;
            }
            let owner = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
            if owner.formula_owner_uid == *uid {
                let identifier = object.archive_info.identifier.ok_or_else(|| {
                    Error::InvalidFormat("Numbers formula owner has no identifier".to_owned())
                })?;
                return Ok(Some((identifier, message_index, owner)));
            }
        }
    }
    Ok(None)
}

fn same_host(
    left: &tsce::CellRecordExpandedArchive,
    right: &tsce::CellRecordExpandedArchive,
) -> bool {
    left.row == right.row && left.column == right.column
}

fn merge_coordinate_set(
    target: &mut tsce::CellCoordSetArchive,
    addition: &tsce::CellCoordSetArchive,
) {
    for source_column in &addition.column_entries {
        let target_column = if let Some(index) = target
            .column_entries
            .iter()
            .position(|entry| entry.column == source_column.column)
        {
            &mut target.column_entries[index]
        } else {
            target
                .column_entries
                .push(tsce::cell_coord_set_archive::ColumnEntry {
                    column: source_column.column,
                    row_set: tsce::IndexSetArchive::default(),
                });
            target
                .column_entries
                .last_mut()
                .expect("a coordinate column was just inserted")
        };
        for source_entry in &source_column.row_set.entries {
            let row = source_entry.range_begin;
            if !target_column.row_set.entries.iter().any(|entry| {
                let end = entry.range_end.unwrap_or(entry.range_begin);
                (entry.range_begin..=end).contains(&row)
            }) {
                target_column.row_set.entries.push(*source_entry);
            }
        }
        target_column
            .row_set
            .entries
            .sort_unstable_by_key(|entry| entry.range_begin);
    }
    target
        .column_entries
        .sort_unstable_by_key(|entry| entry.column);
}

fn auxiliary_owner(
    uid: tsp::Uuid,
    internal_id: u32,
    kind: u32,
    base_uid: tsp::Uuid,
    tile_id: u64,
) -> tsce::FormulaOwnerDependenciesArchive {
    tsce::FormulaOwnerDependenciesArchive {
        formula_owner_uid: uid,
        internal_formula_owner_id: internal_id,
        owner_kind: Some(kind),
        cell_dependencies: Some(tsce::CellDependenciesExpandedArchive::default()),
        range_dependencies: Some(tsce::RangeDependenciesArchive::default()),
        volatile_dependencies: Some(empty_volatile_dependencies()),
        spanning_column_dependencies: Some(empty_spanning_dependencies()),
        spanning_row_dependencies: Some(empty_spanning_dependencies()),
        whole_owner_dependencies: Some(tsce::WholeOwnerDependenciesExpandedArchive {
            dependent_cells: Some(tsce::InternalCellRefSetArchive::default()),
        }),
        cell_errors: Some(tsce::CellErrorsArchive::default()),
        base_owner_uid: Some(base_uid),
        tiled_cell_dependencies: Some(tsce::CellDependenciesTiledArchive {
            cell_record_tiles: vec![tsp::Reference {
                identifier: tile_id,
                ..Default::default()
            }],
        }),
        uuid_references: Some(tsce::UuidReferencesArchive::default()),
        tiled_range_dependencies: Some(tsce::RangeDependenciesTiledArchive::default()),
        spill_range_sizes: Some(tsce::CellSpillSizesArchive::default()),
        ..Default::default()
    }
}

fn conditional_owner(
    uid: tsp::Uuid,
    internal_id: u32,
    tile_id: u64,
    host: &HostDependency,
) -> Result<tsce::FormulaOwnerDependenciesArchive> {
    let mut owner = auxiliary_owner(
        uid,
        internal_id,
        CONDITIONAL_STYLE_OWNER_KIND,
        tsp::Uuid {
            lower: uid
                .lower
                .checked_sub(CONDITIONAL_BASE_OWNER_OFFSET)
                .ok_or_else(|| {
                    Error::ParseError("conditional-style base UUID underflow".to_owned())
                })?,
            upper: uid.upper,
        },
        tile_id,
    );
    owner.cell_dependencies = Some(tsce::CellDependenciesExpandedArchive {
        cell_record: vec![host.record.clone()],
    });
    owner.volatile_dependencies = Some(tsce::VolatileDependenciesExpandedArchive {
        volatile_time_cells: Some(host.coordinates.clone()),
        ..empty_volatile_dependencies()
    });
    owner.uuid_references = Some(tsce::UuidReferencesArchive {
        table_refs: vec![tsce::uuid_references_archive::TableRef {
            owner_uuid: host.table_owner_uid,
            coord_set: Some(host.coordinates.clone()),
        }],
        ..Default::default()
    });
    Ok(owner)
}

fn cell_record_tile(
    internal_owner_id: u32,
    cell_records: Vec<tsce::CellRecordExpandedArchive>,
) -> tsce::CellRecordTileArchive {
    tsce::CellRecordTileArchive {
        internal_owner_id,
        tile_column_begin: 0,
        tile_row_begin: 0,
        cell_records,
    }
}

fn coordinate_set(row: i32, column: u32) -> tsce::CellCoordSetArchive {
    tsce::CellCoordSetArchive {
        column_entries: vec![tsce::cell_coord_set_archive::ColumnEntry {
            column,
            row_set: tsce::IndexSetArchive {
                entries: vec![tsce::index_set_archive::IndexSetEntry {
                    range_begin: row,
                    ..Default::default()
                }],
            },
        }],
    }
}

fn empty_volatile_dependencies() -> tsce::VolatileDependenciesExpandedArchive {
    tsce::VolatileDependenciesExpandedArchive {
        volatile_time_cells: Some(tsce::CellCoordSetArchive::default()),
        volatile_random_cells: Some(tsce::CellCoordSetArchive::default()),
        volatile_locale_cells: Some(tsce::CellCoordSetArchive::default()),
        volatile_sheet_table_name_cells: Some(tsce::CellCoordSetArchive::default()),
        volatile_remote_data_cells: Some(tsce::CellCoordSetArchive::default()),
        volatile_geometry_cell_refs: Some(tsce::InternalCellRefSetArchive::default()),
    }
}

fn empty_spanning_dependencies() -> tsce::SpanningDependenciesExpandedArchive {
    let sentinel = tsce::RangeCoordinateArchive {
        top_left_column: WHOLE_ROW_COLUMN_SENTINEL,
        top_left_row: WHOLE_COLUMN_ROW_SENTINEL,
        bottom_right_column: WHOLE_ROW_COLUMN_SENTINEL,
        bottom_right_row: WHOLE_COLUMN_ROW_SENTINEL,
    };
    tsce::SpanningDependenciesExpandedArchive {
        total_range_for_table: Some(sentinel),
        body_range_for_table: Some(sentinel),
        ..Default::default()
    }
}

fn owner_map_entry(
    internal_owner_id: u32,
    owner_uid: tsp::Uuid,
) -> tsce::owner_id_map_archive::OwnerIdMapArchiveEntry {
    tsce::owner_id_map_archive::OwnerIdMapArchiveEntry {
        internal_owner_id,
        owner_id: uuid_as_cfuuid(&owner_uid),
    }
}

fn insert_owner(
    archive: &mut Archive,
    identifier: u64,
    owner: tsce::FormulaOwnerDependenciesArchive,
    tile_id: u64,
) -> Result<()> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: FORMULA_OWNER_MESSAGE_TYPE,
            data: owner.encode_to_vec(),
        }],
    )?;
    object.archive_info.message_infos[0].versions = NATIVE_MESSAGE_VERSION.to_vec();
    object.archive_info.message_infos[0]
        .object_references
        .push(tile_id);
    archive.insert_object(object)
}

fn insert_tile(
    archive: &mut Archive,
    identifier: u64,
    tile: tsce::CellRecordTileArchive,
) -> Result<()> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: CELL_RECORD_TILE_MESSAGE_TYPE,
            data: tile.encode_to_vec(),
        }],
    )?;
    object.archive_info.message_infos[0].versions = NATIVE_MESSAGE_VERSION.to_vec();
    archive.insert_object(object)
}

fn checked_next_identifier(identifier: u64) -> Result<u64> {
    identifier
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))
}
