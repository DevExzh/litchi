//! Native tick-mark CRUD for Numbers sheet charts.

use super::*;
use crate::charts::axis_style::{
    chart_axis_minor_tick_marks_visible as read_native_chart_axis_minor_tick_marks_visible,
    chart_axis_tick_mark_location as read_native_chart_axis_tick_mark_location,
    set_chart_axis_minor_tick_marks_visible as set_native_chart_axis_minor_tick_marks_visible,
    set_chart_axis_tick_mark_location as set_native_chart_axis_tick_mark_location,
};
use crate::charts::{Axis, TickMarkLocation};

impl NumbersEditor {
    /// Read whether Numbers shows minor tick marks for one native sheet-chart axis.
    pub fn sheet_chart_axis_minor_tick_marks_visible(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<bool> {
        sheet_chart_axis_minor_tick_marks_visible(self, sheet_id, drawable_object_id, axis)
    }

    /// Set whether Numbers shows minor tick marks for one native sheet-chart axis.
    pub fn set_sheet_chart_axis_minor_tick_marks_visible(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: Axis,
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

    /// Read where Numbers draws major tick marks for one native sheet-chart axis.
    pub fn sheet_chart_axis_tick_mark_location(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<TickMarkLocation> {
        sheet_chart_axis_tick_mark_location(self, sheet_id, drawable_object_id, axis)
    }

    /// Set where Numbers draws major tick marks for one native sheet-chart axis.
    pub fn set_sheet_chart_axis_tick_mark_location(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: Axis,
        location: TickMarkLocation,
    ) -> Result<()> {
        set_sheet_chart_axis_tick_mark_location(self, sheet_id, drawable_object_id, axis, location)
    }
}

fn sheet_chart_axis_minor_tick_marks_visible(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    axis: Axis,
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
    axis: Axis,
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

fn sheet_chart_axis_tick_mark_location(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    axis: Axis,
) -> Result<TickMarkLocation> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_chart_axis_tick_mark_location(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        axis,
    )
}

fn set_sheet_chart_axis_tick_mark_location(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    axis: Axis,
    location: TickMarkLocation,
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_chart_axis_tick_mark_location(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        axis,
        location,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_axis_tick_mark_location(sheet_id, drawable_object_id, axis)? != location
    {
        return Err(Error::InvalidFormat(
            "Numbers chart axis tick-mark-location update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
