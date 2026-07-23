//! Native minor-tick-mark visibility CRUD for Keynote slide charts.

use super::*;
use crate::charts::ChartAxis;
use crate::charts::axis_style::{
    chart_axis_minor_tick_marks_visible as read_native_chart_axis_minor_tick_marks_visible,
    set_chart_axis_minor_tick_marks_visible as set_native_chart_axis_minor_tick_marks_visible,
};

impl KeynoteEditor {
    /// Read whether Keynote shows minor tick marks for one native slide-chart axis.
    pub fn slide_chart_axis_minor_tick_marks_visible(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: ChartAxis,
    ) -> Result<bool> {
        slide_chart_axis_minor_tick_marks_visible(self, slide_index, drawable_object_id, axis)
    }

    /// Set whether Keynote shows minor tick marks for one native slide-chart axis.
    pub fn set_slide_chart_axis_minor_tick_marks_visible(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: ChartAxis,
        visible: bool,
    ) -> Result<()> {
        set_slide_chart_axis_minor_tick_marks_visible(
            self,
            slide_index,
            drawable_object_id,
            axis,
            visible,
        )
    }
}

fn slide_chart_axis_minor_tick_marks_visible(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: ChartAxis,
) -> Result<bool> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_chart_axis_minor_tick_marks_visible(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
    )
}

fn set_slide_chart_axis_minor_tick_marks_visible(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: ChartAxis,
    visible: bool,
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_axis_minor_tick_marks_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
        visible,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_axis_minor_tick_marks_visible(slide_index, drawable_object_id, axis)?
        != visible
    {
        return Err(Error::InvalidFormat(
            "Keynote chart axis minor-tick-mark update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
