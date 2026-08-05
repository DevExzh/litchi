//! Table-cell text layout backed by the shared cell-style graph.

use super::*;
use litchi_iwa_common::table::cell::layout::{Inset, Insets, Layout, TextWrap, VerticalAlignment};

impl From<litchi_iwa_common::table::cell::layout::Error> for Error {
    fn from(error: litchi_iwa_common::table::cell::layout::Error) -> Self {
        Self::ParseError(error.to_string())
    }
}

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
) -> Result<Layout> {
    let text_wrap = cell_style::effective_property(
        package,
        table_id,
        row,
        column,
        cell_style::CellStylePropertyKind::TextWrap,
    )?
    .map_or(Ok(TextWrap::Unwrapped), |property| {
        let cell_style::CellStyleProperty::TextWrap(value) = property else {
            return property_mismatch("text wrapping");
        };
        Ok(if value {
            TextWrap::Wrapped
        } else {
            TextWrap::Unwrapped
        })
    })?;
    let vertical_alignment = cell_style::effective_property(
        package,
        table_id,
        row,
        column,
        cell_style::CellStylePropertyKind::VerticalAlignment,
    )?
    .map_or(Ok(VerticalAlignment::Top), |property| {
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
    .map_or(Ok(Insets::ZERO), |property| {
        let cell_style::CellStyleProperty::Padding(value) = property else {
            return property_mismatch("text insets");
        };
        insets_from_native(&value)
    })?;
    Ok(Layout::new(text_wrap, vertical_alignment, insets))
}

pub(super) fn set_cell_layout(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    layout: Layout,
) -> Result<()> {
    let properties = [
        cell_style::CellStyleProperty::TextWrap(matches!(layout.text_wrap(), TextWrap::Wrapped)),
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

fn alignment_from_native(value: i32) -> Result<VerticalAlignment> {
    match value {
        NATIVE_ALIGN_TOP => Ok(VerticalAlignment::Top),
        NATIVE_ALIGN_MIDDLE => Ok(VerticalAlignment::Middle),
        NATIVE_ALIGN_BOTTOM => Ok(VerticalAlignment::Bottom),
        _ => Err(Error::InvalidFormat(format!(
            "iWork table cell uses unknown vertical alignment {value}"
        ))),
    }
}

const fn alignment_to_native(value: VerticalAlignment) -> i32 {
    match value {
        VerticalAlignment::Top => NATIVE_ALIGN_TOP,
        VerticalAlignment::Middle => NATIVE_ALIGN_MIDDLE,
        VerticalAlignment::Bottom => NATIVE_ALIGN_BOTTOM,
    }
}

fn insets_from_native(native: &tswp::PaddingArchive) -> Result<Insets> {
    Ok(Insets::new(
        Inset::from_points(native.left.unwrap_or(0.0))?,
        Inset::from_points(native.top.unwrap_or(0.0))?,
        Inset::from_points(native.right.unwrap_or(0.0))?,
        Inset::from_points(native.bottom.unwrap_or(0.0))?,
    ))
}

fn insets_to_native(insets: Insets) -> tswp::PaddingArchive {
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn unknown_native_alignment_is_rejected() {
        assert!(alignment_from_native(i32::MAX).is_err());
    }

    #[test]
    fn malformed_native_padding_is_rejected_before_layout_construction() {
        let padding = tswp::PaddingArchive {
            left: Some(f32::NAN),
            ..Default::default()
        };
        assert!(insets_from_native(&padding).is_err());
    }
}
