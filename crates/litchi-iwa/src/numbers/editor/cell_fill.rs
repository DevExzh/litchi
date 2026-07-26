//! Table-cell fill CRUD backed by the shared cell-style graph.

use super::*;
use crate::shapes::{ShapeFill, fill_from_native, fill_to_native, validate_image_asset};

pub(super) fn cell_fill(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<ShapeFill> {
    cell_style::effective_property(
        package,
        table_id,
        row,
        column,
        cell_style::CellStylePropertyKind::Fill,
    )?
    .map_or(Ok(ShapeFill::None), |property| {
        let cell_style::CellStyleProperty::Fill(fill) = property else {
            return Err(Error::InvalidFormat(
                "iWork cell-style fill resolved to another property".to_owned(),
            ));
        };
        fill_from_native(&fill)
    })
}

pub(super) fn set_cell_fill(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    fill: &ShapeFill,
) -> Result<()> {
    validate_image_asset(package, fill)?;
    cell_style::set_property(
        package,
        table_id,
        row,
        column,
        cell_style::CellStyleProperty::Fill(Box::new(fill_to_native(fill))),
    )
}

pub(super) fn reset_cell_fill(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    cell_style::reset_property(
        package,
        table_id,
        row,
        column,
        cell_style::CellStylePropertyKind::Fill,
    )
}
