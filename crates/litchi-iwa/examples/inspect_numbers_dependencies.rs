use std::env;

use litchi_iwa::raw::package::IWorkPackage;
use litchi_iwa_protos::tsce;
use prost::Message;

fn print_record(record: &tsce::CellRecordExpandedArchive, indent: &str) {
    let Some(edges) = &record.expanded_edges else {
        println!(
            "{indent}cell=({}, {}) edges=<absent>",
            record.column, record.row
        );
        return;
    };
    let local = edges
        .edge_without_owner_columns
        .iter()
        .copied()
        .zip(edges.edge_without_owner_rows.iter().copied())
        .collect::<Vec<_>>();
    let external = edges
        .edge_with_owner_columns
        .iter()
        .copied()
        .zip(edges.edge_with_owner_rows.iter().copied())
        .zip(edges.internal_owner_id_for_edge.iter().copied())
        .map(|((column, row), owner)| (owner, column, row))
        .collect::<Vec<_>>();
    println!(
        "{indent}cell=({}, {}) dirty={:?} calculated={:?} local={local:?} external={external:?}",
        record.column,
        record.row,
        record.dirty_self_plus_precedents_count,
        record.has_calculated_precedents,
    );
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let arguments = env::args().skip(1).collect::<Vec<_>>();
    let path = arguments
        .first()
        .ok_or("usage: inspect_numbers_dependencies <file.numbers>")?;
    let debug = arguments
        .get(1)
        .is_some_and(|argument| argument == "--debug");
    let package = IWorkPackage::open(path)?;
    let component = package
        .calculation_engine_entry_name()?
        .ok_or("package has no CalculationEngine component")?;
    println!("component={component}");
    let archive = package.archive(component)?;
    for object in archive.objects {
        let id = object.archive_info.identifier.unwrap_or(0);
        for message in object.messages {
            match message.type_ {
                4000 => {
                    let engine = tsce::CalculationEngineArchive::decode(message.data.as_slice())?;
                    println!(
                        "engine={id} formulas={:?} owners={} owner_map={}",
                        engine.dependency_tracker.number_of_formulas,
                        engine.dependency_tracker.formula_owner_dependencies.len(),
                        engine
                            .dependency_tracker
                            .owner_id_map
                            .as_ref()
                            .map_or(0, |map| map.map_entry.len())
                    );
                },
                4008 => {
                    let owner =
                        tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
                    println!(
                        "owner={id} uid={:016x}{:016x} internal={} kind={:?} base={:?} formula_owner={:?} cells={:?} tiled={}",
                        owner.formula_owner_uid.upper,
                        owner.formula_owner_uid.lower,
                        owner.internal_formula_owner_id,
                        owner.owner_kind,
                        owner
                            .base_owner_uid
                            .as_ref()
                            .map(|uid| (uid.upper, uid.lower)),
                        owner
                            .formula_owner
                            .as_ref()
                            .map(|reference| reference.identifier),
                        owner
                            .cell_dependencies
                            .as_ref()
                            .map(|dependencies| dependencies.cell_record.len()),
                        owner
                            .tiled_cell_dependencies
                            .as_ref()
                            .map_or(0, |dependencies| dependencies.cell_record_tiles.len()),
                    );
                    if let Some(dependencies) = &owner.cell_dependencies {
                        for record in &dependencies.cell_record {
                            print_record(record, "  inline ");
                        }
                    }
                    if debug {
                        println!("  {owner:#?}");
                    }
                },
                4009 => {
                    let tile = tsce::CellRecordTileArchive::decode(message.data.as_slice())?;
                    println!(
                        "cell_tile={id} owner={} origin=({}, {}) records={}",
                        tile.internal_owner_id,
                        tile.tile_column_begin,
                        tile.tile_row_begin,
                        tile.cell_records.len()
                    );
                    for record in tile.cell_records {
                        print_record(&record, "  ");
                    }
                },
                4010 => {
                    let tile = tsce::RangePrecedentsTileArchive::decode(message.data.as_slice())?;
                    println!(
                        "range_tile={id} owner={} records={}",
                        tile.to_owner_id,
                        tile.from_to_range.len()
                    );
                    if debug {
                        println!("  {tile:#?}");
                    }
                },
                _ => {},
            }
        }
    }
    Ok(())
}
