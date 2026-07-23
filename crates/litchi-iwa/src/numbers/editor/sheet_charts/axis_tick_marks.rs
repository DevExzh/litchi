//! Native minor-tick-mark visibility CRUD for Numbers sheet charts.

use super::*;
use crate::charts::ChartAxis;
use crate::charts::axis_style::{
    chart_axis_minor_tick_marks_visible as read_native_chart_axis_minor_tick_marks_visible,
    set_chart_axis_minor_tick_marks_visible as set_native_chart_axis_minor_tick_marks_visible,
};

impl NumbersEditor {
    /// Read whether Numbers shows minor tick marks for one native sheet-chart axis.
    pub fn sheet_chart_axis_minor_tick_marks_visible(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: ChartAxis,
    ) -> Result<bool> {
        sheet_chart_axis_minor_tick_marks_visible(self, sheet_id, drawable_object_id, axis)
    }

    /// Set whether Numbers shows minor tick marks for one native sheet-chart axis.
    pub fn set_sheet_chart_axis_minor_tick_marks_visible(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: ChartAxis,
        visible: bool,
    ) -> Result<()> {
        set_sheet_chart_axis_minor_tick_marks_visible(
            self,
            sheet_id,
            drawable_object_id,
            axis,
            visible,
        )
    }
}

fn sheet_chart_axis_minor_tick_marks_visible(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    axis: ChartAxis,
) -> Result<bool> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_chart_axis_minor_tick_marks_visible(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        axis,
    )
}

fn set_sheet_chart_axis_minor_tick_marks_visible(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    axis: ChartAxis,
    visible: bool,
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_chart_axis_minor_tick_marks_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        axis,
        visible,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_axis_minor_tick_marks_visible(sheet_id, drawable_object_id, axis)?
        != visible
    {
        return Err(Error::InvalidFormat(
            "Numbers chart axis minor-tick-mark update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
