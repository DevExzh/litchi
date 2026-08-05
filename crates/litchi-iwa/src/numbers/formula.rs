//! Typed construction of native iWork table formula ASTs and dependency edges.
//!
//! The dependency-free formula vocabulary is owned by
//! `litchi_numbers::formula`; this module retains only archive-boundary
//! compilation and calculation-engine metadata.

use std::collections::{BTreeSet, HashMap};

use super::function_map::function_identifier;
use crate::protobuf::tsce;
use crate::{Error, Result};

use tsce::ast_node_array_archive::AstNodeArchive;
use tsce::ast_node_array_archive::AstNodeType;

pub use litchi_numbers::formula::{
    FormulaAxisReference, FormulaBinaryOperator, FormulaCachedValue, FormulaCellReference,
    FormulaExpression, FormulaPivotCategoryReference, FormulaUuid,
};

const MAX_FORMULA_PRECEDENTS: usize = 1_100_000;
const MAX_FORMULA_NODES: usize = 65_536;
const MAX_FORMULA_DEPTH: usize = 512;
const MAX_FUNCTION_ARGUMENTS: usize = 256;
const MAX_TOTAL_FUNCTION_ARGUMENTS: usize = 16_384;
const WHOLE_ROW_COLUMN_SENTINEL: u32 = i16::MAX as u32;
const WHOLE_COLUMN_ROW_SENTINEL: u32 = i32::MAX as u32;

/// Compile a leaf formula expression using the owning workbook's archive
/// context and calculation-engine tables.
pub(crate) fn compile_formula(
    expression: &FormulaExpression,
    host_row: usize,
    host_column: usize,
    table_rows: usize,
    table_columns: usize,
    external_tables: &HashMap<u64, ExternalFormulaTable>,
    pivot_categories: &HashMap<PivotFormulaKey, ExternalPivotCategory>,
) -> Result<CompiledFormula> {
    let host_row = u32::try_from(host_row)
        .map_err(|_| Error::ParseError("Numbers formula host row exceeds u32".to_owned()))?;
    let host_column = u32::try_from(host_column)
        .map_err(|_| Error::ParseError("Numbers formula host column exceeds u32".to_owned()))?;
    let table_rows = u32::try_from(table_rows)
        .map_err(|_| Error::ParseError("Numbers table row count exceeds u32".to_owned()))?;
    let table_columns = u32::try_from(table_columns)
        .map_err(|_| Error::ParseError("Numbers table column count exceeds u32".to_owned()))?;
    if host_row >= table_rows || host_column >= table_columns {
        return Err(Error::ParseError(format!(
            "Numbers formula host cell ({host_row}, {host_column}) is outside the {table_rows}x{table_columns} table"
        )));
    }

    let node_count = validate_expression(expression)?;
    let mut ast_node = Vec::with_capacity(node_count);
    let mut context = FormulaCompileContext {
        host_row,
        host_column,
        table_rows,
        table_columns,
        external_tables,
        pivot_categories,
        local_precedents: BTreeSet::new(),
        external_precedents: BTreeSet::new(),
        contains_pivot_category: false,
    };
    append_nodes(expression, &mut ast_node, &mut context)?;
    if ast_node.is_empty() {
        return Err(Error::ParseError(
            "Numbers formula AST cannot be empty".to_owned(),
        ));
    }
    Ok(CompiledFormula {
        archive: tsce::FormulaArchive {
            ast_node_array: tsce::AstNodeArrayArchive { ast_node },
            translation_flags: context.contains_pivot_category.then_some(
                tsce::FormulaTranslationFlagsArchive {
                    excel_import_translation: Some(false),
                    number_to_date_coercion_removal_translation: Some(false),
                    contains_uid_form_references: Some(false),
                    contains_frozen_references: Some(false),
                    returns_percent_formatted: Some(true),
                },
            ),
            ..Default::default()
        },
        local_precedents: context.local_precedents.into_iter().collect(),
        external_precedents: context.external_precedents.into_iter().collect(),
    })
}

/// Validate the expression graph before the recursive archive walk.
///
/// Formula values are built by callers, so the compiler must bound both the
/// work it performs and the temporary memory it reserves. The iterative
/// preflight also makes the recursion depth used by `append_nodes` explicit.
fn validate_expression(expression: &FormulaExpression) -> Result<usize> {
    let mut pending = vec![(expression, 1_usize)];
    let mut node_count = 0_usize;
    let mut function_argument_count = 0_usize;

    while let Some((expression, depth)) = pending.pop() {
        if depth > MAX_FORMULA_DEPTH {
            return Err(Error::ParseError(format!(
                "Numbers formula nesting exceeds the safety limit {MAX_FORMULA_DEPTH}"
            )));
        }
        node_count = node_count
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("Numbers formula node count overflow".to_owned()))?;
        if node_count > MAX_FORMULA_NODES {
            return Err(Error::ParseError(format!(
                "Numbers formula has more than {MAX_FORMULA_NODES} expression nodes"
            )));
        }

        match expression {
            FormulaExpression::Function { arguments, .. } => {
                if arguments.len() > MAX_FUNCTION_ARGUMENTS {
                    return Err(Error::ParseError(format!(
                        "Numbers formula function has more than {MAX_FUNCTION_ARGUMENTS} arguments"
                    )));
                }
                function_argument_count = function_argument_count
                    .checked_add(arguments.len())
                    .ok_or_else(|| {
                        Error::ParseError("Numbers formula argument count overflow".to_owned())
                    })?;
                if function_argument_count > MAX_TOTAL_FUNCTION_ARGUMENTS {
                    return Err(Error::ParseError(format!(
                        "Numbers formula has more than {MAX_TOTAL_FUNCTION_ARGUMENTS} function arguments"
                    )));
                }
                pending.extend(arguments.iter().rev().map(|argument| (argument, depth + 1)));
            },
            FormulaExpression::Binary { left, right, .. } => {
                pending.push((right, depth + 1));
                pending.push((left, depth + 1));
            },
            FormulaExpression::Negate(value) | FormulaExpression::Percent(value) => {
                pending.push((value, depth + 1));
            },
            FormulaExpression::Number(_)
            | FormulaExpression::Text(_)
            | FormulaExpression::Boolean(_)
            | FormulaExpression::PivotCategory(_)
            | FormulaExpression::Cell(_)
            | FormulaExpression::TableCell { .. }
            | FormulaExpression::TableRange { .. }
            | FormulaExpression::Rows { .. }
            | FormulaExpression::Columns { .. }
            | FormulaExpression::TableRows { .. }
            | FormulaExpression::TableColumns { .. }
            | FormulaExpression::Range { .. } => {},
        }
    }

    Ok(node_count)
}

fn append_nodes(
    expression: &FormulaExpression,
    output: &mut Vec<AstNodeArchive>,
    context: &mut FormulaCompileContext,
) -> Result<()> {
    match expression {
        FormulaExpression::Number(value) => append_number(*value, output),
        FormulaExpression::Text(value) => {
            if value.contains('\0') {
                return Err(Error::ParseError(
                    "Numbers formula strings cannot contain NUL".to_owned(),
                ));
            }
            output.push(AstNodeArchive {
                ast_node_type: AstNodeType::StringNode as i32,
                ast_string_node_string: Some(value.clone()),
                ..Default::default()
            });
            Ok(())
        },
        FormulaExpression::Boolean(value) => {
            output.push(AstNodeArchive {
                ast_node_type: AstNodeType::BooleanNode as i32,
                ast_boolean_node_boolean: Some(*value),
                ..Default::default()
            });
            Ok(())
        },
        FormulaExpression::PivotCategory(reference) => {
            append_pivot_category_reference(*reference, output, context)
        },
        FormulaExpression::Cell(reference) => append_cell_reference(*reference, output, context),
        FormulaExpression::TableCell {
            table_id,
            reference,
        } => append_table_cell_reference(*table_id, *reference, output, context),
        FormulaExpression::TableRange {
            table_id,
            start,
            end,
        } => append_table_range_reference(*table_id, *start, *end, output, context),
        FormulaExpression::Rows { start, end } => {
            append_axis_range_reference(Axis::Row, *start, *end, None, output, context)
        },
        FormulaExpression::Columns { start, end } => {
            append_axis_range_reference(Axis::Column, *start, *end, None, output, context)
        },
        FormulaExpression::TableRows {
            table_id,
            start,
            end,
        } => append_axis_range_reference(Axis::Row, *start, *end, Some(*table_id), output, context),
        FormulaExpression::TableColumns {
            table_id,
            start,
            end,
        } => append_axis_range_reference(
            Axis::Column,
            *start,
            *end,
            Some(*table_id),
            output,
            context,
        ),
        FormulaExpression::Range { start, end } => {
            append_cell_reference(*start, output, context)?;
            append_cell_reference(*end, output, context)?;
            output.push(AstNodeArchive {
                ast_node_type: AstNodeType::ColonNode as i32,
                ..Default::default()
            });
            let start_row = u32::try_from(start.row).map_err(|_| {
                Error::ParseError("Numbers formula range row exceeds u32".to_owned())
            })?;
            let end_row = u32::try_from(end.row).map_err(|_| {
                Error::ParseError("Numbers formula range row exceeds u32".to_owned())
            })?;
            let start_column = u32::try_from(start.column).map_err(|_| {
                Error::ParseError("Numbers formula range column exceeds u32".to_owned())
            })?;
            let end_column = u32::try_from(end.column).map_err(|_| {
                Error::ParseError("Numbers formula range column exceeds u32".to_owned())
            })?;
            let top = start_row.min(end_row);
            let bottom = start_row.max(end_row);
            let left = start_column.min(end_column);
            let right = start_column.max(end_column);
            let rows = u64::from(bottom - top) + 1;
            let columns = u64::from(right - left) + 1;
            let cells = rows.checked_mul(columns).ok_or_else(|| {
                Error::ParseError("Numbers formula range size overflow".to_owned())
            })?;
            if cells > MAX_FORMULA_PRECEDENTS as u64 {
                return Err(Error::ParseError(format!(
                    "Numbers formula range expands to {cells} precedents, exceeding the safety limit {MAX_FORMULA_PRECEDENTS}"
                )));
            }
            for row in top..=bottom {
                for column in left..=right {
                    context.insert_local_precedent((row, column))?;
                }
            }
            Ok(())
        },
        FormulaExpression::Function { name, arguments } => {
            let identifier = function_identifier(name).ok_or_else(|| {
                Error::ParseError(format!("Unknown Numbers formula function {name:?}"))
            })?;
            validate_function_arity(name, identifier, arguments.len())?;
            if is_lazy_function(identifier) {
                return Err(Error::ParseError(format!(
                    "Numbers function {name} uses thunk/lambda AST nodes that are not yet writable"
                )));
            }
            if requires_dependency_records(identifier) {
                return Err(Error::ParseError(format!(
                    "Numbers function {name} requires volatile, reference, remote-data, or spill dependency records that are not yet writable"
                )));
            }
            for argument in arguments {
                append_nodes(argument, output, context)?;
            }
            let argument_count = u32::try_from(arguments.len()).map_err(|_| {
                Error::ParseError("Numbers formula has too many function arguments".to_owned())
            })?;
            output.push(AstNodeArchive {
                ast_node_type: AstNodeType::FunctionNode as i32,
                ast_function_node_index: Some(identifier),
                ast_function_node_num_args: Some(argument_count),
                ..Default::default()
            });
            Ok(())
        },
        FormulaExpression::Binary {
            operator,
            left,
            right,
        } => {
            append_nodes(left, output, context)?;
            append_nodes(right, output, context)?;
            output.push(AstNodeArchive {
                ast_node_type: binary_node_type(*operator) as i32,
                ..Default::default()
            });
            Ok(())
        },
        FormulaExpression::Negate(value) => {
            append_nodes(value, output, context)?;
            output.push(AstNodeArchive {
                ast_node_type: AstNodeType::NegationNode as i32,
                ..Default::default()
            });
            Ok(())
        },
        FormulaExpression::Percent(value) => {
            append_nodes(value, output, context)?;
            output.push(AstNodeArchive {
                ast_node_type: AstNodeType::PercentNode as i32,
                ..Default::default()
            });
            Ok(())
        },
    }
}

pub(crate) struct CompiledFormula {
    pub(crate) archive: tsce::FormulaArchive,
    /// Sorted `(row, column)` coordinates in the formula's own table.
    pub(crate) local_precedents: Vec<(u32, u32)>,
    /// Sorted `(internal owner, row, column)` coordinates in other tables.
    pub(crate) external_precedents: Vec<(u32, u32, u32)>,
}

#[derive(Debug, Clone)]
pub(crate) struct ExternalFormulaTable {
    pub(crate) rows: u32,
    pub(crate) columns: u32,
    pub(crate) owner_uid: crate::protobuf::tsp::Uuid,
    pub(crate) internal_owner_id: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) struct PivotFormulaKey {
    pub(crate) group_by_uid: FormulaUuid,
    pub(crate) column_uid: FormulaUuid,
    pub(crate) group_uid: FormulaUuid,
}

impl PivotFormulaKey {
    pub(crate) fn new(
        group_by_uid: FormulaUuid,
        column_uid: FormulaUuid,
        group_uid: FormulaUuid,
    ) -> Self {
        Self {
            group_by_uid,
            column_uid,
            group_uid,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ExternalPivotCategory {
    pub(crate) internal_owner_id: u32,
    pub(crate) grouping_columns: (u32, u32),
    pub(crate) aggregate: (u32, u32),
    pub(crate) aggregate_type: u32,
    pub(crate) group_level: i32,
    pub(crate) label: Option<String>,
}

struct FormulaCompileContext<'a> {
    host_row: u32,
    host_column: u32,
    table_rows: u32,
    table_columns: u32,
    external_tables: &'a HashMap<u64, ExternalFormulaTable>,
    pivot_categories: &'a HashMap<PivotFormulaKey, ExternalPivotCategory>,
    local_precedents: BTreeSet<(u32, u32)>,
    external_precedents: BTreeSet<(u32, u32, u32)>,
    contains_pivot_category: bool,
}

impl FormulaCompileContext<'_> {
    fn insert_local_precedent(&mut self, precedent: (u32, u32)) -> Result<()> {
        if !self.local_precedents.contains(&precedent)
            && self.local_precedents.len() + self.external_precedents.len()
                >= MAX_FORMULA_PRECEDENTS
        {
            return Err(Error::ParseError(format!(
                "Numbers formula has more than {MAX_FORMULA_PRECEDENTS} aggregate precedents"
            )));
        }
        self.local_precedents.insert(precedent);
        Ok(())
    }

    fn insert_external_precedent(&mut self, precedent: (u32, u32, u32)) -> Result<()> {
        if !self.external_precedents.contains(&precedent)
            && self.local_precedents.len() + self.external_precedents.len()
                >= MAX_FORMULA_PRECEDENTS
        {
            return Err(Error::ParseError(format!(
                "Numbers formula has more than {MAX_FORMULA_PRECEDENTS} aggregate precedents"
            )));
        }
        self.external_precedents.insert(precedent);
        Ok(())
    }
}

fn append_pivot_category_reference(
    reference: FormulaPivotCategoryReference,
    output: &mut Vec<AstNodeArchive>,
    context: &mut FormulaCompileContext,
) -> Result<()> {
    let key = PivotFormulaKey::new(
        reference.group_by_uid,
        reference.column_uid,
        reference.group_uid,
    );
    let category = context.pivot_categories.get(&key).ok_or_else(|| {
        Error::ParseError(format!(
            "Numbers pivot category {:?} is not present in the workbook's group-by model",
            reference.group_uid
        ))
    })?;
    context.contains_pivot_category = true;
    if reference.aggregate_type != category.aggregate_type {
        return Err(Error::ParseError(format!(
            "Numbers pivot category aggregate type {} does not match the model's {}",
            reference.aggregate_type, category.aggregate_type
        )));
    }
    if reference.group_level != category.group_level {
        return Err(Error::ParseError(format!(
            "Numbers pivot category group level {} does not match the model's {}",
            reference.group_level, category.group_level
        )));
    }

    output.push(AstNodeArchive {
        ast_node_type: AstNodeType::CategoryRefNode as i32,
        ast_category_ref: Some(tsce::ast_node_array_archive::AstCategoryReferenceArchive {
            category_ref: tsce::CategoryReferenceArchive {
                group_by_uid: uuid_archive(reference.group_by_uid),
                column_uid: uuid_archive(reference.column_uid),
                aggregate_type: reference.aggregate_type,
                group_level: reference.group_level,
                preserve_flags: Some(tsce::PreserveColumnRowFlagsArchive {
                    begin_row_is_absolute: false,
                    begin_column_is_absolute: true,
                    end_row_is_absolute: Some(false),
                    end_column_is_absolute: Some(false),
                }),
                absolute_group_uid: Some(uuid_archive(reference.group_uid)),
                agg_index_level: Some(u16::MAX as u32),
                ..Default::default()
            },
        }),
        ..Default::default()
    });
    for &(row, column) in &[category.grouping_columns, category.aggregate] {
        context.insert_external_precedent((category.internal_owner_id, row, column))?;
    }
    Ok(())
}

fn uuid_archive(uuid: FormulaUuid) -> crate::protobuf::tsp::Uuid {
    crate::protobuf::tsp::Uuid {
        lower: uuid.lower,
        upper: uuid.upper,
    }
}

fn append_cell_reference(
    reference: FormulaCellReference,
    output: &mut Vec<AstNodeArchive>,
    context: &mut FormulaCompileContext,
) -> Result<()> {
    let row = u32::try_from(reference.row)
        .map_err(|_| Error::ParseError("Numbers formula reference row exceeds u32".to_owned()))?;
    let column = u32::try_from(reference.column).map_err(|_| {
        Error::ParseError("Numbers formula reference column exceeds u32".to_owned())
    })?;
    if row >= context.table_rows || column >= context.table_columns {
        return Err(Error::ParseError(format!(
            "Numbers formula reference ({row}, {column}) is outside the {}x{} table",
            context.table_rows, context.table_columns
        )));
    }
    let ast_row = coordinate_value(row, context.host_row, reference.absolute_row, "row")?;
    let ast_column = coordinate_value(
        column,
        context.host_column,
        reference.absolute_column,
        "column",
    )?;
    output.push(AstNodeArchive {
        ast_node_type: AstNodeType::CellReferenceNode as i32,
        ast_column: Some(tsce::ast_node_array_archive::AstColumnCoordinateArchive {
            column: ast_column,
            absolute: Some(reference.absolute_column),
        }),
        ast_row: Some(tsce::ast_node_array_archive::AstRowCoordinateArchive {
            row: ast_row,
            absolute: Some(reference.absolute_row),
        }),
        ..Default::default()
    });
    context.insert_local_precedent((row, column))
}

fn append_table_cell_reference(
    table_id: u64,
    reference: FormulaCellReference,
    output: &mut Vec<AstNodeArchive>,
    context: &mut FormulaCompileContext,
) -> Result<()> {
    let table = context.external_tables.get(&table_id).ok_or_else(|| {
        Error::ParseError(format!(
            "Numbers cross-table formula target {table_id} has no CalculationEngine owner"
        ))
    })?;
    let row = u32::try_from(reference.row)
        .map_err(|_| Error::ParseError("Numbers formula reference row exceeds u32".to_owned()))?;
    let column = u32::try_from(reference.column).map_err(|_| {
        Error::ParseError("Numbers formula reference column exceeds u32".to_owned())
    })?;
    if row >= table.rows || column >= table.columns {
        return Err(Error::ParseError(format!(
            "Numbers cross-table formula reference ({row}, {column}) is outside target table {table_id} dimensions {}x{}",
            table.rows, table.columns
        )));
    }
    let ast_row = coordinate_value(row, context.host_row, reference.absolute_row, "row")?;
    let ast_column = coordinate_value(
        column,
        context.host_column,
        reference.absolute_column,
        "column",
    )?;
    output.push(AstNodeArchive {
        ast_node_type: AstNodeType::CellReferenceNode as i32,
        ast_column: Some(tsce::ast_node_array_archive::AstColumnCoordinateArchive {
            column: ast_column,
            absolute: Some(reference.absolute_column),
        }),
        ast_row: Some(tsce::ast_node_array_archive::AstRowCoordinateArchive {
            row: ast_row,
            absolute: Some(reference.absolute_row),
        }),
        ast_cross_table_reference_extra_info: Some(
            tsce::ast_node_array_archive::AstCrossTableReferenceExtraInfoArchive {
                table_id: owner_uid_as_cfuuid(&table.owner_uid),
                ..Default::default()
            },
        ),
        ..Default::default()
    });
    context.insert_external_precedent((table.internal_owner_id, row, column))
}

fn append_table_range_reference(
    table_id: u64,
    start: FormulaCellReference,
    end: FormulaCellReference,
    output: &mut Vec<AstNodeArchive>,
    context: &mut FormulaCompileContext,
) -> Result<()> {
    let table = context.external_tables.get(&table_id).ok_or_else(|| {
        Error::ParseError(format!(
            "Numbers cross-table formula target {table_id} has no CalculationEngine owner"
        ))
    })?;
    let (table_rows, table_columns, owner_uid, internal_owner_id) = (
        table.rows,
        table.columns,
        table.owner_uid,
        table.internal_owner_id,
    );
    let start_row = u32::try_from(start.row)
        .map_err(|_| Error::ParseError("Numbers formula range row exceeds u32".to_owned()))?;
    let end_row = u32::try_from(end.row)
        .map_err(|_| Error::ParseError("Numbers formula range row exceeds u32".to_owned()))?;
    let start_column = u32::try_from(start.column)
        .map_err(|_| Error::ParseError("Numbers formula range column exceeds u32".to_owned()))?;
    let end_column = u32::try_from(end.column)
        .map_err(|_| Error::ParseError("Numbers formula range column exceeds u32".to_owned()))?;
    if start_row >= table_rows
        || end_row >= table_rows
        || start_column >= table_columns
        || end_column >= table_columns
    {
        return Err(Error::ParseError(format!(
            "Numbers cross-table formula range ({start_row}, {start_column}):({end_row}, {end_column}) is outside target table {table_id} dimensions {table_rows}x{table_columns}"
        )));
    }

    let (relative_column, absolute_column) = colon_axis_ranges(
        start_column,
        end_column,
        start.absolute_column,
        end.absolute_column,
        context.host_column,
        "column",
    )?;
    let (relative_row, absolute_row) = colon_axis_ranges(
        start_row,
        end_row,
        start.absolute_row,
        end.absolute_row,
        context.host_row,
        "row",
    )?;
    output.push(AstNodeArchive {
        ast_node_type: AstNodeType::ColonTractNode as i32,
        ast_cross_table_reference_extra_info: Some(
            tsce::ast_node_array_archive::AstCrossTableReferenceExtraInfoArchive {
                table_id: owner_uid_as_cfuuid(&owner_uid),
                ..Default::default()
            },
        ),
        ast_sticky_bits: Some(tsce::ast_node_array_archive::AstStickyBits {
            begin_row_is_absolute: start.absolute_row,
            begin_column_is_absolute: start.absolute_column,
            end_row_is_absolute: end.absolute_row,
            end_column_is_absolute: end.absolute_column,
        }),
        ast_colon_tract: Some(tsce::ast_node_array_archive::AstColonTractArchive {
            relative_column,
            relative_row,
            absolute_column,
            absolute_row,
            preserve_rectangular: Some(true),
        }),
        ..Default::default()
    });

    let top = start_row.min(end_row);
    let bottom = start_row.max(end_row);
    let left = start_column.min(end_column);
    let right = start_column.max(end_column);
    let rows = u64::from(bottom - top) + 1;
    let columns = u64::from(right - left) + 1;
    let cells = rows
        .checked_mul(columns)
        .ok_or_else(|| Error::ParseError("Numbers formula range size overflow".to_owned()))?;
    if cells > MAX_FORMULA_PRECEDENTS as u64 {
        return Err(Error::ParseError(format!(
            "Numbers formula range expands to {cells} precedents, exceeding the safety limit {MAX_FORMULA_PRECEDENTS}"
        )));
    }
    for row in top..=bottom {
        for column in left..=right {
            context.insert_external_precedent((internal_owner_id, row, column))?;
        }
    }
    Ok(())
}

#[derive(Debug, Clone, Copy)]
enum Axis {
    Row,
    Column,
}

fn append_axis_range_reference(
    axis: Axis,
    start: FormulaAxisReference,
    end: FormulaAxisReference,
    table_id: Option<u64>,
    output: &mut Vec<AstNodeArchive>,
    context: &mut FormulaCompileContext,
) -> Result<()> {
    let (rows, columns, owner) = if let Some(table_id) = table_id {
        let table = context.external_tables.get(&table_id).ok_or_else(|| {
            Error::ParseError(format!(
                "Numbers cross-table formula target {table_id} has no CalculationEngine owner"
            ))
        })?;
        (
            table.rows,
            table.columns,
            Some((table.internal_owner_id, table.owner_uid)),
        )
    } else {
        (context.table_rows, context.table_columns, None)
    };
    let start_index = u32::try_from(start.index)
        .map_err(|_| Error::ParseError("Numbers formula axis index exceeds u32".to_owned()))?;
    let end_index = u32::try_from(end.index)
        .map_err(|_| Error::ParseError("Numbers formula axis index exceeds u32".to_owned()))?;
    let (axis_length, host, axis_name) = match axis {
        Axis::Row => (rows, context.host_row, "row"),
        Axis::Column => (columns, context.host_column, "column"),
    };
    if start_index >= axis_length || end_index >= axis_length {
        let target = table_id.map_or_else(
            || "host table".to_owned(),
            |identifier| format!("target table {identifier}"),
        );
        return Err(Error::ParseError(format!(
            "Numbers formula whole-{axis_name} range {start_index}:{end_index} is outside {target} {axis_name} count {axis_length}"
        )));
    }

    let (relative_axis, absolute_axis) = colon_axis_ranges(
        start_index,
        end_index,
        start.absolute,
        end.absolute,
        host,
        axis_name,
    )?;
    let (relative_column, relative_row, absolute_column, absolute_row, sticky) = match axis {
        Axis::Row => (
            Vec::new(),
            relative_axis,
            vec![AbsoluteRange {
                range_begin: WHOLE_ROW_COLUMN_SENTINEL,
                range_end: None,
            }],
            absolute_axis,
            tsce::ast_node_array_archive::AstStickyBits {
                begin_row_is_absolute: start.absolute,
                begin_column_is_absolute: false,
                end_row_is_absolute: end.absolute,
                end_column_is_absolute: false,
            },
        ),
        Axis::Column => (
            relative_axis,
            Vec::new(),
            absolute_axis,
            vec![AbsoluteRange {
                range_begin: WHOLE_COLUMN_ROW_SENTINEL,
                range_end: None,
            }],
            tsce::ast_node_array_archive::AstStickyBits {
                begin_row_is_absolute: false,
                begin_column_is_absolute: start.absolute,
                end_row_is_absolute: false,
                end_column_is_absolute: end.absolute,
            },
        ),
    };
    output.push(AstNodeArchive {
        ast_node_type: AstNodeType::ColonTractNode as i32,
        ast_cross_table_reference_extra_info: owner.as_ref().map(|(_, uid)| {
            tsce::ast_node_array_archive::AstCrossTableReferenceExtraInfoArchive {
                table_id: owner_uid_as_cfuuid(uid),
                ..Default::default()
            }
        }),
        ast_sticky_bits: Some(sticky),
        ast_colon_tract: Some(tsce::ast_node_array_archive::AstColonTractArchive {
            relative_column,
            relative_row,
            absolute_column,
            absolute_row,
            preserve_rectangular: Some(true),
        }),
        ..Default::default()
    });

    let begin = start_index.min(end_index);
    let finish = start_index.max(end_index);
    let selected = u64::from(finish - begin) + 1;
    let perpendicular = u64::from(match axis {
        Axis::Row => columns,
        Axis::Column => rows,
    });
    let cells = selected.checked_mul(perpendicular).ok_or_else(|| {
        Error::ParseError("Numbers formula whole-axis range size overflow".to_owned())
    })?;
    if cells > MAX_FORMULA_PRECEDENTS as u64 {
        return Err(Error::ParseError(format!(
            "Numbers formula whole-{axis_name} range expands to {cells} precedents, exceeding the safety limit {MAX_FORMULA_PRECEDENTS}"
        )));
    }
    match (axis, owner) {
        (Axis::Row, Some((owner_id, _))) => {
            for row in begin..=finish {
                for column in 0..columns {
                    context.insert_external_precedent((owner_id, row, column))?;
                }
            }
        },
        (Axis::Column, Some((owner_id, _))) => {
            for row in 0..rows {
                for column in begin..=finish {
                    context.insert_external_precedent((owner_id, row, column))?;
                }
            }
        },
        (Axis::Row, None) => {
            for row in begin..=finish {
                for column in 0..columns {
                    context.insert_local_precedent((row, column))?;
                }
            }
        },
        (Axis::Column, None) => {
            for row in 0..rows {
                for column in begin..=finish {
                    context.insert_local_precedent((row, column))?;
                }
            }
        },
    }
    Ok(())
}

type RelativeRange =
    tsce::ast_node_array_archive::ast_colon_tract_archive::AstColonTractRelativeRangeArchive;
type AbsoluteRange =
    tsce::ast_node_array_archive::ast_colon_tract_archive::AstColonTractAbsoluteRangeArchive;

fn colon_axis_ranges(
    start: u32,
    end: u32,
    start_absolute: bool,
    end_absolute: bool,
    host: u32,
    axis: &str,
) -> Result<(Vec<RelativeRange>, Vec<AbsoluteRange>)> {
    let mut relative = Vec::new();
    let mut absolute = Vec::new();
    match (start_absolute, end_absolute) {
        (false, false) => relative.push(RelativeRange {
            range_begin: coordinate_value(start, host, false, axis)?,
            range_end: (end != start)
                .then(|| coordinate_value(end, host, false, axis))
                .transpose()?,
        }),
        (true, true) => absolute.push(AbsoluteRange {
            range_begin: start,
            range_end: (end != start).then_some(end),
        }),
        (false, true) => {
            relative.push(RelativeRange {
                range_begin: coordinate_value(start, host, false, axis)?,
                range_end: None,
            });
            absolute.push(AbsoluteRange {
                range_begin: end,
                range_end: None,
            });
        },
        (true, false) => {
            relative.push(RelativeRange {
                range_begin: coordinate_value(end, host, false, axis)?,
                range_end: None,
            });
            absolute.push(AbsoluteRange {
                range_begin: start,
                range_end: None,
            });
        },
    }
    Ok((relative, absolute))
}

fn owner_uid_as_cfuuid(uid: &crate::protobuf::tsp::Uuid) -> crate::protobuf::tsp::CfuuidArchive {
    crate::protobuf::tsp::CfuuidArchive {
        uuid_bytes: None,
        uuid_w0: Some(uid.lower as u32),
        uuid_w1: Some((uid.lower >> 32) as u32),
        uuid_w2: Some(uid.upper as u32),
        uuid_w3: Some((uid.upper >> 32) as u32),
    }
}

fn coordinate_value(target: u32, host: u32, absolute: bool, axis: &str) -> Result<i32> {
    let value = if absolute {
        i64::from(target)
    } else {
        i64::from(target) - i64::from(host)
    };
    i32::try_from(value).map_err(|_| {
        Error::ParseError(format!(
            "Numbers formula {axis} coordinate does not fit the AST format"
        ))
    })
}

fn append_number(value: f64, output: &mut Vec<AstNodeArchive>) -> Result<()> {
    if !value.is_finite() {
        return Err(Error::ParseError(
            "Numbers formula literals must be finite".to_owned(),
        ));
    }
    if value.is_sign_negative() && value != 0.0 {
        append_number(-value, output)?;
        output.push(AstNodeArchive {
            ast_node_type: AstNodeType::NegationNode as i32,
            ..Default::default()
        });
        return Ok(());
    }
    let (decimal_low, decimal_high) = decimal128_parts(value)?;
    output.push(AstNodeArchive {
        ast_node_type: AstNodeType::NumberNode as i32,
        ast_number_node_number: Some(value),
        ast_number_node_decimal_low: Some(decimal_low),
        ast_number_node_decimal_high: Some(decimal_high),
        ..Default::default()
    });
    Ok(())
}

/// Encode the canonical decimal spelling used by app-generated formula ASTs.
/// `f64::to_string` supplies the shortest round-tripping decimal, whose
/// coefficient is at most seventeen digits and therefore fits comfortably in
/// decimal128's 113-bit coefficient field.
fn decimal128_parts(value: f64) -> Result<(u64, u64)> {
    let encoded = u128::from_le_bytes(super::bnc::decimal128_le(value)?);
    Ok((encoded as u64, (encoded >> 64) as u64))
}

fn binary_node_type(operator: FormulaBinaryOperator) -> AstNodeType {
    match operator {
        FormulaBinaryOperator::Add => AstNodeType::AdditionNode,
        FormulaBinaryOperator::Subtract => AstNodeType::SubtractionNode,
        FormulaBinaryOperator::Multiply => AstNodeType::MultiplicationNode,
        FormulaBinaryOperator::Divide => AstNodeType::DivisionNode,
        FormulaBinaryOperator::Power => AstNodeType::PowerNode,
        FormulaBinaryOperator::Concatenate => AstNodeType::ConcatenationNode,
        FormulaBinaryOperator::GreaterThan => AstNodeType::GreaterThanNode,
        FormulaBinaryOperator::GreaterThanOrEqual => AstNodeType::GreaterThanOrEqualToNode,
        FormulaBinaryOperator::LessThan => AstNodeType::LessThanNode,
        FormulaBinaryOperator::LessThanOrEqual => AstNodeType::LessThanOrEqualToNode,
        FormulaBinaryOperator::Equal => AstNodeType::EqualToNode,
        FormulaBinaryOperator::NotEqual => AstNodeType::NotEqualToNode,
    }
}

fn is_lazy_function(identifier: u32) -> bool {
    matches!(
        identifier,
        62 | 235 | 313 | 336 | 363..=372 // IF/IFERROR/IFS/SWITCH and lambda family
    )
}

fn validate_function_arity(name: &str, identifier: u32, actual: usize) -> Result<()> {
    let (minimum, maximum) = match identifier {
        1
        | 4
        | 5
        | 9
        | 10
        | 11
        | 13
        | 18
        | 20
        | 21
        | 28
        | 29
        | 32
        | 41
        | 44
        | 48
        | 50
        | 51
        | 60
        | 65
        | 69..=72
        | 77
        | 78
        | 80
        | 82
        | 90
        | 94
        | 96
        | 104
        | 117
        | 124
        | 125
        | 129
        | 132
        | 133
        | 134
        | 135
        | 139
        | 149
        | 150..=151
        | 155
        | 157..=159
        | 167 => (1, 1),
        7 | 15..=16 | 25 | 30..=31 | 84..=89 | 102 | 113 | 138 | 140..=143 | 160..=163 | 168 => {
            (1, MAX_FUNCTION_ARGUMENTS)
        },
        12
        | 19
        | 24
        | 27
        | 49
        | 53
        | 66
        | 74
        | 75
        | 81
        | 83
        | 92
        | 95
        | 103
        | 107
        | 120
        | 126..=128
        | 137
        | 145
        | 146
        | 148
        | 152
        | 164
        | 216..=218
        | 221
        | 223
        | 225
        | 226
        | 231..=234
        | 314 => (2, 2),
        39 => (3, 3),
        52 | 156 | 97 => (0, 0),
        62 => (3, 3),
        212 => (2, 2),
        _ => {
            return Err(Error::ParseError(format!(
                "Numbers function {name} has no validated arity metadata"
            )));
        },
    };
    if actual < minimum || actual > maximum {
        return Err(Error::ParseError(format!(
            "Numbers function {name} expects {minimum}..={maximum} arguments, got {actual}"
        )));
    }
    Ok(())
}

fn requires_dependency_records(identifier: u32) -> bool {
    matches!(
        identifier,
        64 | 97 | 101 | 118 | 119 | 154 | 298..=303 | 322 | 323 | 325 | 338 | 342..=362
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    fn compile(
        expression: &FormulaExpression,
        host_row: usize,
        host_column: usize,
        table_rows: usize,
        table_columns: usize,
        external_tables: &HashMap<u64, ExternalFormulaTable>,
        pivot_categories: &HashMap<PivotFormulaKey, ExternalPivotCategory>,
    ) -> Result<CompiledFormula> {
        compile_formula(
            expression,
            host_row,
            host_column,
            table_rows,
            table_columns,
            external_tables,
            pivot_categories,
        )
    }

    #[test]
    fn compiles_observed_sum_ast() {
        let formula = FormulaExpression::function(
            "SUM",
            [
                FormulaExpression::Number(1.0),
                FormulaExpression::Number(2.0),
            ],
        );
        let compiled = compile(&formula, 5, 2, 20, 10, &HashMap::new(), &HashMap::new()).unwrap();
        let nodes = &compiled.archive.ast_node_array.ast_node;
        assert_eq!(nodes.len(), 3);
        assert_eq!(nodes[0].ast_node_type(), AstNodeType::NumberNode);
        assert_eq!(nodes[0].ast_number_node_decimal_low, Some(1));
        assert_eq!(nodes[1].ast_number_node_decimal_low, Some(2));
        assert_eq!(nodes[2].ast_node_type(), AstNodeType::FunctionNode);
        assert_eq!(nodes[2].ast_function_node_index, Some(168));
        assert_eq!(nodes[2].ast_function_node_num_args, Some(2));
    }

    #[test]
    fn pivot_categories_match_app_ast_and_group_aggregate_dependencies() {
        let group_by = FormulaUuid::new(10, 20);
        let column = FormulaUuid::new(30, 40);
        let north = FormulaUuid::new(50, 60);
        let total = FormulaUuid::new(1, 0);
        let categories = HashMap::from([
            (
                PivotFormulaKey::new(group_by, column, north),
                ExternalPivotCategory {
                    internal_owner_id: 38,
                    grouping_columns: (0, 1),
                    aggregate: (0, 8),
                    aggregate_type: 2,
                    group_level: 2,
                    label: Some("North".to_owned()),
                },
            ),
            (
                PivotFormulaKey::new(group_by, column, total),
                ExternalPivotCategory {
                    internal_owner_id: 38,
                    grouping_columns: (0, 1),
                    aggregate: (0, 18),
                    aggregate_type: 2,
                    group_level: 0,
                    label: Some("Grand Total".to_owned()),
                },
            ),
        ]);
        let north_reference = FormulaPivotCategoryReference::new(group_by, column, north, 2, 2);
        let total_reference = FormulaPivotCategoryReference::new(group_by, column, total, 2, 0);
        let expression = FormulaExpression::binary(
            FormulaBinaryOperator::Divide,
            FormulaExpression::pivot_category(north_reference),
            FormulaExpression::pivot_category(total_reference),
        );
        let compiled = compile(&expression, 0, 0, 1, 1, &HashMap::new(), &categories).unwrap();
        assert_eq!(
            compiled.external_precedents,
            [(38, 0, 1), (38, 0, 8), (38, 0, 18)]
        );
        let nodes = &compiled.archive.ast_node_array.ast_node;
        assert_eq!(nodes[0].ast_node_type(), AstNodeType::CategoryRefNode);
        let category = &nodes[0].ast_category_ref.as_ref().unwrap().category_ref;
        assert_eq!(category.group_by_uid, uuid_archive(group_by));
        assert_eq!(category.column_uid, uuid_archive(column));
        assert_eq!(category.absolute_group_uid, Some(uuid_archive(north)));
        assert_eq!(category.aggregate_type, 2);
        assert_eq!(category.group_level, 2);
        assert_eq!(category.agg_index_level, Some(u16::MAX as u32));
        assert_eq!(nodes[2].ast_node_type(), AstNodeType::DivisionNode);
        assert_eq!(
            compiled
                .archive
                .translation_flags
                .as_ref()
                .and_then(|flags| flags.returns_percent_formatted),
            Some(true)
        );

        let mismatched = FormulaExpression::pivot_category(FormulaPivotCategoryReference::new(
            group_by, column, north, 2, 1,
        ));
        assert!(compile(&mismatched, 0, 0, 1, 1, &HashMap::new(), &categories,).is_err());
    }

    #[test]
    fn decimal_literals_match_app_generated_formula_nodes() {
        let mut nodes = Vec::new();
        append_number(0.75, &mut nodes).unwrap();
        assert_eq!(nodes[0].ast_number_node_decimal_low, Some(75));
        assert_eq!(
            nodes[0].ast_number_node_decimal_high,
            Some(0x303c_0000_0000_0000)
        );

        nodes.clear();
        append_number(1e-5, &mut nodes).unwrap();
        assert_eq!(nodes[0].ast_number_node_decimal_low, Some(1));
        assert_eq!(
            nodes[0].ast_number_node_decimal_high,
            Some(0x3036_0000_0000_0000)
        );
    }

    #[test]
    fn rejects_unknown_and_lazy_functions() {
        assert!(
            compile(
                &FormulaExpression::function("MISSING", []),
                0,
                0,
                1,
                1,
                &HashMap::new(),
                &HashMap::new(),
            )
            .is_err()
        );
        assert!(
            compile(
                &FormulaExpression::function("IF", [FormulaExpression::Boolean(true)]),
                0,
                0,
                1,
                1,
                &HashMap::new(),
                &HashMap::new(),
            )
            .is_err()
        );
        assert!(
            compile(
                &FormulaExpression::function("NOW", []),
                0,
                0,
                1,
                1,
                &HashMap::new(),
                &HashMap::new(),
            )
            .is_err()
        );
        assert!(
            compile(
                &FormulaExpression::function("COUNTBLANK", []),
                0,
                0,
                1,
                1,
                &HashMap::new(),
                &HashMap::new(),
            )
            .is_err()
        );
        assert!(
            compile(
                &FormulaExpression::function("SEQUENCE", [FormulaExpression::Number(2.0)]),
                0,
                0,
                1,
                1,
                &HashMap::new(),
                &HashMap::new(),
            )
            .is_err()
        );
    }

    #[test]
    fn rejects_invalid_arity_and_unbounded_expression_work() {
        assert!(
            compile(
                &FormulaExpression::function(
                    "SIN",
                    [
                        FormulaExpression::Number(1.0),
                        FormulaExpression::Number(2.0)
                    ],
                ),
                0,
                0,
                1,
                1,
                &HashMap::new(),
                &HashMap::new(),
            )
            .is_err()
        );

        let too_many_arguments = FormulaExpression::function(
            "SUM",
            (0..=MAX_FUNCTION_ARGUMENTS).map(|value| FormulaExpression::Number(value as f64)),
        );
        assert!(
            compile(
                &too_many_arguments,
                0,
                0,
                1,
                1,
                &HashMap::new(),
                &HashMap::new(),
            )
            .is_err()
        );

        let deeply_nested = (0..MAX_FORMULA_DEPTH)
            .fold(FormulaExpression::Number(1.0), |expression, _| {
                FormulaExpression::negate(expression)
            });
        assert!(compile(&deeply_nested, 0, 0, 1, 1, &HashMap::new(), &HashMap::new(),).is_err());
    }

    #[test]
    fn cell_coordinates_match_app_generated_relative_and_mixed_references() {
        let relative = compile(
            &FormulaExpression::relative_cell(4, 1),
            5,
            2,
            20,
            10,
            &HashMap::new(),
            &HashMap::new(),
        )
        .unwrap();
        let node = &relative.archive.ast_node_array.ast_node[0];
        assert_eq!(node.ast_node_type(), AstNodeType::CellReferenceNode);
        assert_eq!(node.ast_column.as_ref().unwrap().column, -1);
        assert_eq!(node.ast_column.as_ref().unwrap().absolute, Some(false));
        assert_eq!(node.ast_row.as_ref().unwrap().row, -1);
        assert_eq!(node.ast_row.as_ref().unwrap().absolute, Some(false));
        assert_eq!(relative.local_precedents, [(4, 1)]);

        let mixed = compile(
            &FormulaExpression::cell(FormulaCellReference::mixed(1, 0, true, false)),
            5,
            2,
            20,
            10,
            &HashMap::new(),
            &HashMap::new(),
        )
        .unwrap();
        let node = &mixed.archive.ast_node_array.ast_node[0];
        assert_eq!(node.ast_column.as_ref().unwrap().column, -2);
        assert_eq!(node.ast_row.as_ref().unwrap().row, 1);
        assert_eq!(node.ast_row.as_ref().unwrap().absolute, Some(true));
    }

    #[test]
    fn ranges_expand_sorted_unique_dependency_edges() {
        let expression = FormulaExpression::function(
            "SUM",
            [FormulaExpression::range(
                FormulaCellReference::relative(2, 1),
                FormulaCellReference::absolute(1, 0),
            )],
        );
        let compiled =
            compile(&expression, 5, 2, 20, 10, &HashMap::new(), &HashMap::new()).unwrap();
        assert_eq!(compiled.local_precedents, [(1, 0), (1, 1), (2, 0), (2, 1)]);
        assert_eq!(
            compiled
                .archive
                .ast_node_array
                .ast_node
                .iter()
                .map(|node| node.ast_node_type())
                .collect::<Vec<_>>(),
            [
                AstNodeType::CellReferenceNode,
                AstNodeType::CellReferenceNode,
                AstNodeType::ColonNode,
                AstNodeType::FunctionNode,
            ]
        );
    }

    #[test]
    fn cross_table_cells_use_target_owner_uid_and_external_edges() {
        let tables = HashMap::from([(
            42,
            ExternalFormulaTable {
                rows: 5,
                columns: 9,
                owner_uid: crate::protobuf::tsp::Uuid {
                    upper: 0xbd57_cadc_10f6_658a,
                    lower: 0xe24e_294b_3170_bbd8,
                },
                internal_owner_id: 7,
            },
        )]);
        let compiled = compile(
            &FormulaExpression::table_cell(42, FormulaCellReference::relative(0, 1)),
            4,
            1,
            5,
            4,
            &tables,
            &HashMap::new(),
        )
        .unwrap();
        let node = &compiled.archive.ast_node_array.ast_node[0];
        assert_eq!(node.ast_node_type(), AstNodeType::CellReferenceNode);
        assert_eq!(node.ast_column.as_ref().unwrap().column, 0);
        assert_eq!(node.ast_row.as_ref().unwrap().row, -4);
        let table_id = &node
            .ast_cross_table_reference_extra_info
            .as_ref()
            .unwrap()
            .table_id;
        assert_eq!(table_id.uuid_w0, Some(0x3170_bbd8));
        assert_eq!(table_id.uuid_w1, Some(0xe24e_294b));
        assert_eq!(table_id.uuid_w2, Some(0x10f6_658a));
        assert_eq!(table_id.uuid_w3, Some(0xbd57_cadc));
        assert_eq!(compiled.local_precedents, []);
        assert_eq!(compiled.external_precedents, [(7, 0, 1)]);
    }

    #[test]
    fn cross_table_ranges_match_app_generated_colon_tracts() {
        let tables = HashMap::from([(
            42,
            ExternalFormulaTable {
                rows: 5,
                columns: 9,
                owner_uid: crate::protobuf::tsp::Uuid {
                    upper: 0xbd57_cadc_10f6_658a,
                    lower: 0xe24e_294b_3170_bbd8,
                },
                internal_owner_id: 7,
            },
        )]);
        let compiled = compile(
            &FormulaExpression::table_range(
                42,
                FormulaCellReference::relative(0, 2),
                FormulaCellReference::relative(1, 2),
            ),
            4,
            2,
            5,
            4,
            &tables,
            &HashMap::new(),
        )
        .unwrap();
        let node = &compiled.archive.ast_node_array.ast_node[0];
        assert_eq!(node.ast_node_type(), AstNodeType::ColonTractNode);
        assert_eq!(
            node.ast_sticky_bits,
            Some(tsce::ast_node_array_archive::AstStickyBits {
                begin_row_is_absolute: false,
                begin_column_is_absolute: false,
                end_row_is_absolute: false,
                end_column_is_absolute: false,
            })
        );
        let tract = node.ast_colon_tract.as_ref().unwrap();
        assert_eq!(
            tract.relative_column,
            [RelativeRange {
                range_begin: 0,
                range_end: None,
            }]
        );
        assert_eq!(
            tract.relative_row,
            [RelativeRange {
                range_begin: -4,
                range_end: Some(-3),
            }]
        );
        assert!(tract.absolute_column.is_empty());
        assert!(tract.absolute_row.is_empty());
        assert_eq!(compiled.local_precedents, []);
        assert_eq!(compiled.external_precedents, [(7, 0, 2), (7, 1, 2)]);

        let mixed = compile(
            &FormulaExpression::table_range(
                42,
                FormulaCellReference::absolute(1, 0),
                FormulaCellReference::relative(3, 1),
            ),
            4,
            2,
            5,
            4,
            &tables,
            &HashMap::new(),
        )
        .unwrap();
        let mixed = mixed.archive.ast_node_array.ast_node[0]
            .ast_colon_tract
            .as_ref()
            .unwrap();
        assert_eq!(mixed.absolute_column[0].range_begin, 0);
        assert_eq!(mixed.relative_column[0].range_begin, -1);
        assert_eq!(mixed.absolute_row[0].range_begin, 1);
        assert_eq!(mixed.relative_row[0].range_begin, -1);
    }

    #[test]
    fn whole_axis_ranges_match_app_sentinels_and_expand_dependencies() {
        let tables = HashMap::from([(
            42,
            ExternalFormulaTable {
                rows: 4,
                columns: 3,
                owner_uid: crate::protobuf::tsp::Uuid {
                    upper: 0xbd57_cadc_10f6_658a,
                    lower: 0xe24e_294b_3170_bbd8,
                },
                internal_owner_id: 7,
            },
        )]);
        let rows = compile(
            &FormulaExpression::table_rows(
                42,
                FormulaAxisReference::relative(0),
                FormulaAxisReference::relative(1),
            ),
            1,
            0,
            5,
            4,
            &tables,
            &HashMap::new(),
        )
        .unwrap();
        let node = &rows.archive.ast_node_array.ast_node[0];
        assert_eq!(node.ast_node_type(), AstNodeType::ColonTractNode);
        assert_eq!(
            node.ast_sticky_bits,
            Some(tsce::ast_node_array_archive::AstStickyBits {
                begin_row_is_absolute: false,
                begin_column_is_absolute: false,
                end_row_is_absolute: false,
                end_column_is_absolute: false,
            })
        );
        let tract = node.ast_colon_tract.as_ref().unwrap();
        assert!(tract.relative_column.is_empty());
        assert_eq!(
            tract.relative_row,
            [RelativeRange {
                range_begin: -1,
                range_end: Some(0),
            }]
        );
        assert_eq!(
            tract.absolute_column,
            [AbsoluteRange {
                range_begin: i16::MAX as u32,
                range_end: None,
            }]
        );
        assert!(tract.absolute_row.is_empty());
        assert_eq!(
            rows.external_precedents,
            [
                (7, 0, 0),
                (7, 0, 1),
                (7, 0, 2),
                (7, 1, 0),
                (7, 1, 1),
                (7, 1, 2),
            ]
        );

        let columns = compile(
            &FormulaExpression::table_columns(
                42,
                FormulaAxisReference::relative(1),
                FormulaAxisReference::absolute(2),
            ),
            2,
            0,
            5,
            4,
            &tables,
            &HashMap::new(),
        )
        .unwrap();
        let node = &columns.archive.ast_node_array.ast_node[0];
        let tract = node.ast_colon_tract.as_ref().unwrap();
        assert_eq!(tract.relative_column[0].range_begin, 1);
        assert_eq!(tract.absolute_column[0].range_begin, 2);
        assert!(tract.relative_row.is_empty());
        assert_eq!(tract.absolute_row[0].range_begin, i32::MAX as u32);
        assert!(
            !node
                .ast_sticky_bits
                .as_ref()
                .unwrap()
                .begin_column_is_absolute
        );
        assert!(
            node.ast_sticky_bits
                .as_ref()
                .unwrap()
                .end_column_is_absolute
        );
        assert_eq!(columns.external_precedents.len(), 8);

        let local = compile(
            &FormulaExpression::rows(
                FormulaAxisReference::absolute(2),
                FormulaAxisReference::absolute(2),
            ),
            3,
            1,
            5,
            4,
            &HashMap::new(),
            &HashMap::new(),
        )
        .unwrap();
        assert_eq!(local.local_precedents, [(2, 0), (2, 1), (2, 2), (2, 3)]);
        assert!(local.external_precedents.is_empty());
    }

    #[test]
    fn rejects_out_of_bounds_references_transactionally() {
        assert!(
            compile(
                &FormulaExpression::relative_cell(10, 0),
                0,
                0,
                10,
                1,
                &HashMap::new(),
                &HashMap::new(),
            )
            .is_err()
        );
        assert!(
            compile(
                &FormulaExpression::rows(
                    FormulaAxisReference::relative(0),
                    FormulaAxisReference::relative(10),
                ),
                0,
                0,
                10,
                1,
                &HashMap::new(),
                &HashMap::new(),
            )
            .is_err()
        );
        assert!(
            compile(
                &FormulaExpression::columns(
                    FormulaAxisReference::relative(0),
                    FormulaAxisReference::relative(0),
                ),
                0,
                0,
                MAX_FORMULA_PRECEDENTS + 1,
                1,
                &HashMap::new(),
                &HashMap::new(),
            )
            .is_err()
        );
    }
}
