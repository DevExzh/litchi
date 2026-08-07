//! Shared typed formula mutation for native tables attached in any iWork app.

use super::*;

pub(super) fn set_attached_table_formula(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    expression: FormulaExpression,
    cached_value: Option<FormulaCachedValue>,
) -> Result<()> {
    table_sparse_storage::ensure_attached_cell_storage(package, table_id, row, column)?;
    let descriptors = attached_table_descriptors(package)?;
    let descriptor = descriptors
        .iter()
        .find(|table| table.object_id == table_id)
        .cloned()
        .ok_or_else(|| Error::ParseError(format!("iWork table object {table_id} not found")))?;
    let external_tables = formula_external_tables(package, &descriptors)?;
    let pivot_categories = formula_pivot_categories(package)?;
    let compiled = compile_formula(
        &expression,
        row,
        column,
        descriptor.model.number_of_rows as usize,
        descriptor.model.number_of_columns as usize,
        &external_tables,
        &pivot_categories,
    )?;
    let formula = compiled.archive;

    let locations = object_locations(package)?;
    let (old_formula, old_formula_error) = {
        let location = locate_attached_cell(package, table_id, row, column)?;
        let cell = read_tile_cell(
            package,
            &location.tile_archive,
            location.tile_id,
            location.tile_row,
            column,
        )?
        .as_deref()
        .map(BncCell::parse)
        .transpose()?;
        let formula = cell.as_ref().and_then(|cell| match cell.stored_value() {
            StoredValue::Formula(identifier) => Some(identifier),
            _ => None,
        });
        let error = cell.as_ref().and_then(BncCell::formula_error_identifier);
        (formula, error)
    };
    if let Some(cached_value) = cached_value {
        set_attached_cell_in_package(package, table_id, row, column, cached_value.into_value())?;
    } else {
        if old_formula.is_some()
            && let Some(identifier) = old_formula_error
        {
            decrement_formula_error_table(package, &locations, &descriptor.model, identifier)?;
        }
        if let Some(identifier) = old_formula {
            decrement_formula_table(
                package,
                &locations,
                descriptor.model.base_data_store.formula_table.identifier,
                identifier,
            )?;
            update_formula_dependencies(
                package,
                descriptor.table_info_id,
                row,
                column,
                false,
                &[],
                &[],
            )?;
        } else {
            let empty_value = CellValue::number(0.0).map_err(|error| {
                Error::InvalidFormat(format!(
                    "Numbers formula placeholder value is not finite: {error}"
                ))
            })?;
            set_attached_cell_in_package(package, table_id, row, column, empty_value)?;
        }
    }

    let formula_id = insert_formula_table(
        package,
        &locations,
        descriptor.model.base_data_store.formula_table.identifier,
        formula.clone(),
    )?;
    set_encoded_cell_value(
        package,
        table_id,
        row,
        column,
        EncodedValue::Formula(formula_id),
    )?;
    update_formula_dependencies(
        package,
        descriptor.table_info_id,
        row,
        column,
        true,
        &compiled.local_precedents,
        &compiled.external_precedents,
    )?;

    let verified = IWorkPackage::from_bytes(&package.to_bytes()?)?;
    verify_formula_link(&verified, table_id, row, column, formula_id, &formula)?;
    verify_formula_dependency(
        &verified,
        descriptor.table_info_id,
        row,
        column,
        &compiled.local_precedents,
        &compiled.external_precedents,
    )
}
