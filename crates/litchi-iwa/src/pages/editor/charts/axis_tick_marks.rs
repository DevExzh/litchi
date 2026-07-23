//! Native minor-tick-mark visibility CRUD for Pages body charts.

use super::*;
use crate::charts::ChartAxis;
use crate::charts::axis_style::{
    chart_axis_minor_tick_marks_visible as read_native_chart_axis_minor_tick_marks_visible,
    set_chart_axis_minor_tick_marks_visible as set_native_chart_axis_minor_tick_marks_visible,
};

impl PagesEditor {
    /// Read whether Pages shows minor tick marks for one native body-chart axis.
    pub fn body_chart_axis_minor_tick_marks_visible(
        &self,
        drawable_object_id: u64,
        axis: ChartAxis,
    ) -> Result<bool> {
        body_chart_axis_minor_tick_marks_visible(self, drawable_object_id, axis)
    }

    /// Set whether Pages shows minor tick marks for one native body-chart axis.
    pub fn set_body_chart_axis_minor_tick_marks_visible(
        &mut self,
        drawable_object_id: u64,
        axis: ChartAxis,
        visible: bool,
    ) -> Result<()> {
        set_body_chart_axis_minor_tick_marks_visible(self, drawable_object_id, axis, visible)
    }
}

fn body_chart_axis_minor_tick_marks_visible(
    editor: &PagesEditor,
    drawable_object_id: u64,
    axis: ChartAxis,
) -> Result<bool> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_chart_axis_minor_tick_marks_visible(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        axis,
    )
}

fn set_body_chart_axis_minor_tick_marks_visible(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    axis: ChartAxis,
    visible: bool,
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_axis_minor_tick_marks_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        axis,
        visible,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_axis_minor_tick_marks_visible(drawable_object_id, axis)? != visible {
        return Err(Error::InvalidFormat(
            "Pages chart axis minor-tick-mark update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
