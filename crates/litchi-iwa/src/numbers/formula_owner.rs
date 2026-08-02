//! Typed construction and identity helpers for Numbers table formula owners.

use crate::protobuf::{tsce, tsp};

pub(crate) const TABLE_FORMULA_OWNER_KIND: u32 = 1;

pub(crate) fn empty_table_formula_owner(
    table_uuid: &tsp::Uuid,
    table_info_id: u64,
    internal_owner_id: u32,
) -> tsce::FormulaOwnerDependenciesArchive {
    tsce::FormulaOwnerDependenciesArchive {
        formula_owner_uid: formula_owner_uuid_for_table(table_uuid),
        internal_formula_owner_id: internal_owner_id,
        owner_kind: Some(TABLE_FORMULA_OWNER_KIND),
        cell_dependencies: Some(tsce::CellDependenciesExpandedArchive::default()),
        range_dependencies: Some(tsce::RangeDependenciesArchive::default()),
        volatile_dependencies: Some(tsce::VolatileDependenciesExpandedArchive {
            volatile_time_cells: Some(tsce::CellCoordSetArchive::default()),
            volatile_random_cells: Some(tsce::CellCoordSetArchive::default()),
            volatile_locale_cells: Some(tsce::CellCoordSetArchive::default()),
            volatile_sheet_table_name_cells: Some(tsce::CellCoordSetArchive::default()),
            volatile_remote_data_cells: Some(tsce::CellCoordSetArchive::default()),
            volatile_geometry_cell_refs: Some(tsce::InternalCellRefSetArchive::default()),
        }),
        spanning_column_dependencies: Some(tsce::SpanningDependenciesExpandedArchive::default()),
        spanning_row_dependencies: Some(tsce::SpanningDependenciesExpandedArchive::default()),
        whole_owner_dependencies: Some(tsce::WholeOwnerDependenciesExpandedArchive {
            dependent_cells: Some(tsce::InternalCellRefSetArchive::default()),
        }),
        cell_errors: Some(tsce::CellErrorsArchive::default()),
        formula_owner: Some(tsp::Reference {
            identifier: table_info_id,
            ..Default::default()
        }),
        tiled_cell_dependencies: Some(tsce::CellDependenciesTiledArchive::default()),
        uuid_references: Some(tsce::UuidReferencesArchive::default()),
        tiled_range_dependencies: Some(tsce::RangeDependenciesTiledArchive::default()),
        spill_range_sizes: Some(tsce::CellSpillSizesArchive::default()),
        ..Default::default()
    }
}

pub(crate) fn formula_owner_uuid_for_table(table: &tsp::Uuid) -> tsp::Uuid {
    tsp::Uuid {
        lower: table.upper.swap_bytes(),
        upper: table.lower.swap_bytes(),
    }
}

pub(crate) fn uuid_as_cfuuid(uuid: &tsp::Uuid) -> tsp::CfuuidArchive {
    tsp::CfuuidArchive {
        uuid_bytes: None,
        uuid_w0: Some(uuid.lower as u32),
        uuid_w1: Some((uuid.lower >> 32) as u32),
        uuid_w2: Some(uuid.upper as u32),
        uuid_w3: Some((uuid.upper >> 32) as u32),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn native_formula_owner_uuid_reverses_table_uuid_bytes() {
        let table = tsp::Uuid {
            lower: 0xa1f5_067b_8152_4a6a,
            upper: 0x68b3_84c7_9e7f_4b90,
        };
        assert_eq!(
            formula_owner_uuid_for_table(&table),
            tsp::Uuid {
                lower: 0x904b_7f9e_c784_b368,
                upper: 0x6a4a_5281_7b06_f5a1,
            }
        );
    }
}
