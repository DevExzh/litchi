//! Native value-axis scale-bound CRUD for Numbers sheet charts.

use super::*;
use crate::charts::ChartValueAxisBounds;
use crate::charts::axis_bounds::{
    chart_value_axis_bounds as read_native_chart_value_axis_bounds,
    set_chart_value_axis_bounds as set_native_chart_value_axis_bounds,
};

impl NumbersEditor {
    /// Read the manual minimum and maximum of a native sheet chart's value axis.
    ///
    /// Missing endpoints use Numbers' automatic scale calculation.
    pub fn sheet_chart_value_axis_bounds(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ChartValueAxisBounds> {
        sheet_chart_value_axis_bounds(self, sheet_id, drawable_object_id)
    }

    /// Set the manual minimum and maximum of a native sheet chart's value axis.
    ///
    /// Use [`ChartValueAxisBounds::automatic`] to restore Numbers' automatic
    /// bounds for both endpoints.
    pub fn set_sheet_chart_value_axis_bounds(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        bounds: ChartValueAxisBounds,
    ) -> Result<()> {
        set_sheet_chart_value_axis_bounds(self, sheet_id, drawable_object_id, bounds)
    }
}

fn sheet_chart_value_axis_bounds(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<ChartValueAxisBounds> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_chart_value_axis_bounds(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
    )
}

fn set_sheet_chart_value_axis_bounds(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    bounds: ChartValueAxisBounds,
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_chart_value_axis_bounds(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        bounds,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_value_axis_bounds(sheet_id, drawable_object_id)? != bounds {
        return Err(Error::InvalidFormat(
            "Numbers chart value-axis bounds update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
