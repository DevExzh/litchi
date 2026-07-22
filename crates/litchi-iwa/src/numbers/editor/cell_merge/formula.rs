//! Canonical native merge-formula construction and strict semantic parsing.

use super::*;
use crate::numbers::formula_owner::{formula_owner_uuid_for_table, uuid_as_cfuuid};

const NATIVE_MERGE_FUNCTION_INDEX: u32 = 168;
const NATIVE_MERGE_FUNCTION_ARGUMENTS: u32 = 1;

pub(super) fn parse_regions(model: &TableModelArchive) -> Result<Vec<IWorkTableCellRegion>> {
    let Some(owner) = &model.merge_owner else {
        return Ok(Vec::new());
    };
    let Some(store) = &owner.formula_store else {
        return Ok(Vec::new());
    };
    let expected_table = uuid_as_cfuuid(&formula_owner_uuid_for_table(&parse_table_uuid(
        &model.table_id,
    )?));
    let mut regions = Vec::with_capacity(store.formulas.len());
    let mut formula_indexes = HashSet::with_capacity(store.formulas.len());
    for pair in &store.formulas {
        if pair.formula_index >= store.next_formula_index {
            return Err(Error::InvalidFormat(format!(
                "iWork merge formula index {} reaches next index {}",
                pair.formula_index, store.next_formula_index
            )));
        }
        if !formula_indexes.insert(pair.formula_index) {
            return Err(Error::InvalidFormat(format!(
                "iWork merge formula index {} is duplicated",
                pair.formula_index
            )));
        }
        let region = parse_merge_formula(&pair.formula, &expected_table)?;
        validate_region_bounds(model, region)?;
        if let Some(overlap) = regions
            .iter()
            .find(|candidate: &&IWorkTableCellRegion| candidate.overlaps(region))
        {
            return Err(Error::InvalidFormat(format!(
                "iWork merge regions {overlap:?} and {region:?} overlap"
            )));
        }
        regions.push(region);
    }
    Ok(regions)
}

pub(super) fn parse_merge_formula(
    formula: &tsce::FormulaArchive,
    expected_table: &tsp::CfuuidArchive,
) -> Result<IWorkTableCellRegion> {
    let nodes = &formula.ast_node_array.ast_node;
    let [range_node, function_node] = nodes.as_slice() else {
        return Err(unsupported_merge_formula());
    };
    use tsce::ast_node_array_archive::AstNodeType;
    if range_node.ast_node_type != AstNodeType::ColonTractNode as i32
        || function_node.ast_node_type != AstNodeType::FunctionNode as i32
        || function_node.ast_function_node_index != Some(NATIVE_MERGE_FUNCTION_INDEX)
        || function_node.ast_function_node_num_args != Some(NATIVE_MERGE_FUNCTION_ARGUMENTS)
        || range_node
            .ast_cross_table_reference_extra_info
            .as_ref()
            .map(|info| &info.table_id)
            != Some(expected_table)
    {
        return Err(unsupported_merge_formula());
    }
    let sticky = range_node
        .ast_sticky_bits
        .as_ref()
        .ok_or_else(unsupported_merge_formula)?;
    if !(sticky.begin_row_is_absolute
        && sticky.begin_column_is_absolute
        && sticky.end_row_is_absolute
        && sticky.end_column_is_absolute)
    {
        return Err(unsupported_merge_formula());
    }
    let tract = range_node
        .ast_colon_tract
        .as_ref()
        .ok_or_else(unsupported_merge_formula)?;
    if !tract.relative_column.is_empty()
        || !tract.relative_row.is_empty()
        || tract.absolute_column.len() != 1
        || tract.absolute_row.len() != 1
        || tract.preserve_rectangular == Some(false)
    {
        return Err(unsupported_merge_formula());
    }
    let columns = &tract.absolute_column[0];
    let rows = &tract.absolute_row[0];
    let end_column = columns.range_end.unwrap_or(columns.range_begin);
    let end_row = rows.range_end.unwrap_or(rows.range_begin);
    let column_count = end_column
        .checked_sub(columns.range_begin)
        .and_then(|difference| difference.checked_add(1))
        .ok_or_else(unsupported_merge_formula)?;
    let row_count = end_row
        .checked_sub(rows.range_begin)
        .and_then(|difference| difference.checked_add(1))
        .ok_or_else(unsupported_merge_formula)?;
    IWorkTableCellRegion::new(
        rows.range_begin as usize,
        columns.range_begin as usize,
        row_count as usize,
        column_count as usize,
    )
    .map_err(|_| unsupported_merge_formula())
}

pub(super) fn merge_formula(
    region: IWorkTableCellRegion,
    table_id: tsp::CfuuidArchive,
) -> Result<tsce::FormulaArchive> {
    use tsce::ast_node_array_archive::AstNodeType;
    use tsce::ast_node_array_archive::ast_colon_tract_archive::AstColonTractAbsoluteRangeArchive;
    let begin_row = u32::try_from(region.row())
        .map_err(|_| Error::ParseError("Table-cell merge row exceeds u32".to_owned()))?;
    let end_row = u32::try_from(region.end_row())
        .map_err(|_| Error::ParseError("Table-cell merge row exceeds u32".to_owned()))?;
    let begin_column = u32::try_from(region.column())
        .map_err(|_| Error::ParseError("Table-cell merge column exceeds u32".to_owned()))?;
    let end_column = u32::try_from(region.end_column())
        .map_err(|_| Error::ParseError("Table-cell merge column exceeds u32".to_owned()))?;
    let range_node = tsce::ast_node_array_archive::AstNodeArchive {
        ast_node_type: AstNodeType::ColonTractNode as i32,
        ast_cross_table_reference_extra_info: Some(
            tsce::ast_node_array_archive::AstCrossTableReferenceExtraInfoArchive {
                table_id,
                ..Default::default()
            },
        ),
        ast_sticky_bits: Some(tsce::ast_node_array_archive::AstStickyBits {
            begin_row_is_absolute: true,
            begin_column_is_absolute: true,
            end_row_is_absolute: true,
            end_column_is_absolute: true,
        }),
        ast_colon_tract: Some(tsce::ast_node_array_archive::AstColonTractArchive {
            absolute_column: vec![AstColonTractAbsoluteRangeArchive {
                range_begin: begin_column,
                range_end: Some(end_column),
            }],
            absolute_row: vec![AstColonTractAbsoluteRangeArchive {
                range_begin: begin_row,
                range_end: Some(end_row),
            }],
            preserve_rectangular: Some(true),
            ..Default::default()
        }),
        ..Default::default()
    };
    let function_node = tsce::ast_node_array_archive::AstNodeArchive {
        ast_node_type: AstNodeType::FunctionNode as i32,
        ast_function_node_index: Some(NATIVE_MERGE_FUNCTION_INDEX),
        ast_function_node_num_args: Some(NATIVE_MERGE_FUNCTION_ARGUMENTS),
        ..Default::default()
    };
    Ok(tsce::FormulaArchive {
        ast_node_array: tsce::AstNodeArrayArchive {
            ast_node: vec![range_node, function_node],
        },
        ..Default::default()
    })
}

/// Retarget an already-validated merge formula without changing its unrelated
/// semantic fields or optional-field presence.
///
/// The matching wire rewrite updates only these changed coordinates, retaining
/// unknown protobuf fields in native formulas byte-for-byte.
pub(super) fn rewrite_formula_region(
    formula: &tsce::FormulaArchive,
    region: IWorkTableCellRegion,
) -> Result<tsce::FormulaArchive> {
    let begin_row = merge_coordinate(region.row(), "row")?;
    let end_row = merge_coordinate(region.end_row(), "row")?;
    let begin_column = merge_coordinate(region.column(), "column")?;
    let end_column = merge_coordinate(region.end_column(), "column")?;
    let mut rewritten = formula.clone();
    let range_node = rewritten
        .ast_node_array
        .ast_node
        .first_mut()
        .ok_or_else(unsupported_merge_formula)?;
    let tract = range_node
        .ast_colon_tract
        .as_mut()
        .ok_or_else(unsupported_merge_formula)?;
    rewrite_absolute_range(&mut tract.absolute_row, begin_row, end_row, "row")?;
    rewrite_absolute_range(
        &mut tract.absolute_column,
        begin_column,
        end_column,
        "column",
    )?;
    Ok(rewritten)
}

fn merge_coordinate(value: usize, axis: &str) -> Result<u32> {
    u32::try_from(value)
        .map_err(|_| Error::ParseError(format!("Table-cell merge {axis} exceeds u32")))
}

fn rewrite_absolute_range(
    ranges: &mut [
        tsce::ast_node_array_archive::ast_colon_tract_archive::AstColonTractAbsoluteRangeArchive
    ],
    begin: u32,
    end: u32,
    axis: &str,
) -> Result<()> {
    let [range] = ranges else {
        return Err(Error::InvalidFormat(format!(
            "iWork merge formula has unsupported absolute {axis} ranges"
        )));
    };
    let had_explicit_end = range.range_end.is_some();
    range.range_begin = begin;
    range.range_end = (had_explicit_end || begin != end).then_some(end);
    Ok(())
}

pub(super) fn validate_region_bounds(
    model: &TableModelArchive,
    region: IWorkTableCellRegion,
) -> Result<()> {
    if region.end_row() >= model.number_of_rows as usize
        || region.end_column() >= model.number_of_columns as usize
    {
        return Err(Error::ParseError(format!(
            "Table-cell region {region:?} exceeds table dimensions {}x{}",
            model.number_of_rows, model.number_of_columns
        )));
    }
    Ok(())
}

pub(super) fn parse_table_uuid(value: &str) -> Result<tsp::Uuid> {
    let compact = value
        .chars()
        .filter(|character| *character != '-')
        .collect::<String>();
    if compact.len() != 32 || !compact.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(Error::InvalidFormat(format!(
            "iWork table UUID {value:?} is malformed"
        )));
    }
    let raw = u128::from_str_radix(&compact, 16)
        .map_err(|_| Error::InvalidFormat(format!("iWork table UUID {value:?} is malformed")))?;
    Ok(tsp::Uuid {
        lower: raw as u64,
        upper: (raw >> 64) as u64,
    })
}

fn unsupported_merge_formula() -> Error {
    Error::InvalidFormat("iWork table contains an unsupported merge formula".to_owned())
}
