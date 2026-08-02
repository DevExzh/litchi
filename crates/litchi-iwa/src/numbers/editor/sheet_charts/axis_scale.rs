//! Native value-axis scale CRUD for Numbers sheet charts.

use super::*;
use crate::charts::ChartValueAxisScale;
use crate::charts::axis_scale::{
    chart_value_axis_scale as read_native_chart_value_axis_scale,
    set_chart_value_axis_scale as set_native_chart_value_axis_scale,
};

impl NumbersEditor {
    /// Read the primary value-axis scale of one native sheet chart.
    pub fn sheet_chart_value_axis_scale(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ChartValueAxisScale> {
        sheet_chart_value_axis_scale(self, sheet_id, drawable_object_id)
    }

    /// Set the primary value-axis scale of one native sheet chart.
    pub fn set_sheet_chart_value_axis_scale(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        scale: ChartValueAxisScale,
    ) -> Result<()> {
        set_sheet_chart_value_axis_scale(self, sheet_id, drawable_object_id, scale)
    }
}

fn sheet_chart_value_axis_scale(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<ChartValueAxisScale> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_chart_value_axis_scale(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
    )
}

fn set_sheet_chart_value_axis_scale(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    scale: ChartValueAxisScale,
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_chart_value_axis_scale(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        scale,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_value_axis_scale(sheet_id, drawable_object_id)? != scale {
        return Err(Error::InvalidFormat(
            "Numbers chart value-axis scale update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
