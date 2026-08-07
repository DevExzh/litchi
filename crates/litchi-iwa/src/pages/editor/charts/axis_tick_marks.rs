//! Native tick-mark CRUD for Pages body charts.

use super::*;
use crate::charts::axis_style::{
    chart_axis_minor_tick_marks_visible as read_native_chart_axis_minor_tick_marks_visible,
    chart_axis_tick_mark_location as read_native_chart_axis_tick_mark_location,
    set_chart_axis_minor_tick_marks_visible as set_native_chart_axis_minor_tick_marks_visible,
    set_chart_axis_tick_mark_location as set_native_chart_axis_tick_mark_location,
};
use litchi_iwa_common::chart::axis::style::Visibility;
use litchi_iwa_common::chart::axis::{Axis, TickMarkLocation};

impl PagesEditor {
    /// Read whether Pages shows minor tick marks for one native body-chart axis.
    pub fn body_chart_axis_minor_tick_marks_visible(
        &self,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<Visibility> {
        body_chart_axis_minor_tick_marks_visible(self, drawable_object_id, axis)
    }

    /// Set whether Pages shows minor tick marks for one native body-chart axis.
    pub fn set_body_chart_axis_minor_tick_marks_visible(
        &mut self,
        drawable_object_id: u64,
        axis: Axis,
        visible: Visibility,
    ) -> Result<()> {
        set_body_chart_axis_minor_tick_marks_visible(self, drawable_object_id, axis, visible)
    }

    /// Read where Pages draws major tick marks for one native body-chart axis.
    pub fn body_chart_axis_tick_mark_location(
        &self,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<TickMarkLocation> {
        body_chart_axis_tick_mark_location(self, drawable_object_id, axis)
    }

    /// Set where Pages draws major tick marks for one native body-chart axis.
    pub fn set_body_chart_axis_tick_mark_location(
        &mut self,
        drawable_object_id: u64,
        axis: Axis,
        location: TickMarkLocation,
    ) -> Result<()> {
        set_body_chart_axis_tick_mark_location(self, drawable_object_id, axis, location)
    }
}

fn body_chart_axis_minor_tick_marks_visible(
    editor: &PagesEditor,
    drawable_object_id: u64,
    axis: Axis,
) -> Result<Visibility> {
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
    axis: Axis,
    visible: Visibility,
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

fn body_chart_axis_tick_mark_location(
    editor: &PagesEditor,
    drawable_object_id: u64,
    axis: Axis,
) -> Result<TickMarkLocation> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_chart_axis_tick_mark_location(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        axis,
    )
}

fn set_body_chart_axis_tick_mark_location(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    axis: Axis,
    location: TickMarkLocation,
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_axis_tick_mark_location(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        axis,
        location,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_axis_tick_mark_location(drawable_object_id, axis)? != location {
        return Err(Error::InvalidFormat(
            "Pages chart axis tick-mark-location update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
