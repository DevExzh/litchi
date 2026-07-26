//! Table-cell text layout backed by the shared cell-style graph.

use super::*;
use crate::table_cell_layout::{
    TableCellInset, TableCellInsets, TableCellLayout, TableCellTextWrap, TableCellVerticalAlignment,
};

const NATIVE_ALIGN_TOP: i32 =
    tswp::shape_style_properties_archive::VerticalAlignmentType::KFrameAlignTop as i32;
const NATIVE_ALIGN_MIDDLE: i32 =
    tswp::shape_style_properties_archive::VerticalAlignmentType::KFrameAlignMiddle as i32;
const NATIVE_ALIGN_BOTTOM: i32 =
    tswp::shape_style_properties_archive::VerticalAlignmentType::KFrameAlignBottom as i32;

pub(super) fn cell_layout(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TableCellLayout> {
    let text_wrap = cell_style::effective_property(
        package,
        table_id,
        row,
        column,
        cell_style::CellStylePropertyKind::TextWrap,
    )?
    .map_or(Ok(TableCellTextWrap::Unwrapped), |property| {
        let cell_style::CellStyleProperty::TextWrap(value) = property else {
            return property_mismatch("text wrapping");
        };
        Ok(if value {
            TableCellTextWrap::Wrapped
        } else {
            TableCellTextWrap::Unwrapped
        })
    })?;
    let vertical_alignment = cell_style::effective_property(
        package,
        table_id,
        row,
        column,
        cell_style::CellStylePropertyKind::VerticalAlignment,
    )?
    .map_or(Ok(TableCellVerticalAlignment::Top), |property| {
        let cell_style::CellStyleProperty::VerticalAlignment(value) = property else {
            return property_mismatch("vertical alignment");
        };
        alignment_from_native(value)
    })?;
    let insets = cell_style::effective_property(
        package,
        table_id,
        row,
        column,
        cell_style::CellStylePropertyKind::Padding,
    )?
    .map_or(Ok(TableCellInsets::ZERO), |property| {
        let cell_style::CellStyleProperty::Padding(value) = property else {
            return property_mismatch("text insets");
        };
        insets_from_native(&value)
    })?;
    Ok(TableCellLayout::new(text_wrap, vertical_alignment, insets))
}

pub(super) fn set_cell_layout(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    layout: TableCellLayout,
) -> Result<()> {
    let properties = [
        cell_style::CellStyleProperty::TextWrap(matches!(
            layout.text_wrap(),
            TableCellTextWrap::Wrapped
        )),
        cell_style::CellStyleProperty::VerticalAlignment(alignment_to_native(
            layout.vertical_alignment(),
        )),
        cell_style::CellStyleProperty::Padding(insets_to_native(layout.insets())),
    ];
    cell_style::set_properties(package, table_id, row, column, &properties)
}

pub(super) fn reset_cell_layout(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    let mut staged = package.clone();
    let mut changed = false;
    for kind in [
        cell_style::CellStylePropertyKind::TextWrap,
        cell_style::CellStylePropertyKind::VerticalAlignment,
        cell_style::CellStylePropertyKind::Padding,
    ] {
        changed |= cell_style::reset_property(&mut staged, table_id, row, column, kind)?;
    }
    if changed {
        *package = staged;
    }
    Ok(changed)
}

fn alignment_from_native(value: i32) -> Result<TableCellVerticalAlignment> {
    match value {
        NATIVE_ALIGN_TOP => Ok(TableCellVerticalAlignment::Top),
        NATIVE_ALIGN_MIDDLE => Ok(TableCellVerticalAlignment::Middle),
        NATIVE_ALIGN_BOTTOM => Ok(TableCellVerticalAlignment::Bottom),
        _ => Err(Error::InvalidFormat(format!(
            "iWork table cell uses unknown vertical alignment {value}"
        ))),
    }
}

const fn alignment_to_native(value: TableCellVerticalAlignment) -> i32 {
    match value {
        TableCellVerticalAlignment::Top => NATIVE_ALIGN_TOP,
        TableCellVerticalAlignment::Middle => NATIVE_ALIGN_MIDDLE,
        TableCellVerticalAlignment::Bottom => NATIVE_ALIGN_BOTTOM,
    }
}

fn insets_from_native(native: &tswp::PaddingArchive) -> Result<TableCellInsets> {
    Ok(TableCellInsets::new(
        TableCellInset::from_points(native.left.unwrap_or(0.0))?,
        TableCellInset::from_points(native.top.unwrap_or(0.0))?,
        TableCellInset::from_points(native.right.unwrap_or(0.0))?,
        TableCellInset::from_points(native.bottom.unwrap_or(0.0))?,
    ))
}

fn insets_to_native(insets: TableCellInsets) -> tswp::PaddingArchive {
    tswp::PaddingArchive {
        left: Some(insets.left().points()),
        top: Some(insets.top().points()),
        right: Some(insets.right().points()),
        bottom: Some(insets.bottom().points()),
    }
}

fn property_mismatch<T>(name: &str) -> Result<T> {
    Err(Error::InvalidFormat(format!(
        "iWork cell-style {name} resolved to another property"
    )))
}
