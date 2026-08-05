//! Native tick-mark CRUD for Keynote slide charts.

use super::*;
use crate::charts::axis_style::{
    chart_axis_minor_tick_marks_visible as read_native_chart_axis_minor_tick_marks_visible,
    chart_axis_tick_mark_location as read_native_chart_axis_tick_mark_location,
    set_chart_axis_minor_tick_marks_visible as set_native_chart_axis_minor_tick_marks_visible,
    set_chart_axis_tick_mark_location as set_native_chart_axis_tick_mark_location,
};
use crate::charts::{Axis, TickMarkLocation};

impl KeynoteEditor {
    /// Read whether Keynote shows minor tick marks for one native slide-chart axis.
    pub fn slide_chart_axis_minor_tick_marks_visible(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<bool> {
        slide_chart_axis_minor_tick_marks_visible(self, slide_index, drawable_object_id, axis)
    }

    /// Set whether Keynote shows minor tick marks for one native slide-chart axis.
    pub fn set_slide_chart_axis_minor_tick_marks_visible(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: Axis,
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

    /// Read where Keynote draws major tick marks for one native slide-chart axis.
    pub fn slide_chart_axis_tick_mark_location(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<TickMarkLocation> {
        slide_chart_axis_tick_mark_location(self, slide_index, drawable_object_id, axis)
    }

    /// Set where Keynote draws major tick marks for one native slide-chart axis.
    pub fn set_slide_chart_axis_tick_mark_location(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: Axis,
        location: TickMarkLocation,
    ) -> Result<()> {
        set_slide_chart_axis_tick_mark_location(
            self,
            slide_index,
            drawable_object_id,
            axis,
            location,
        )
    }
}

fn slide_chart_axis_minor_tick_marks_visible(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: Axis,
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
    axis: Axis,
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

fn slide_chart_axis_tick_mark_location(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: Axis,
) -> Result<TickMarkLocation> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_chart_axis_tick_mark_location(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
    )
}

fn set_slide_chart_axis_tick_mark_location(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: Axis,
    location: TickMarkLocation,
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_axis_tick_mark_location(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
        location,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_axis_tick_mark_location(slide_index, drawable_object_id, axis)?
        != location
    {
        return Err(Error::InvalidFormat(
            "Keynote chart axis tick-mark-location update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
