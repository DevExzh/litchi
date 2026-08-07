//! Native axis-line visibility CRUD for Keynote slide charts.

use super::*;
use crate::charts::axis_style::{
    chart_axis_line_visible as read_native_chart_axis_line_visible,
    set_chart_axis_line_visible as set_native_chart_axis_line_visible,
};
use litchi_iwa_common::chart::axis::Axis;
use litchi_iwa_common::chart::axis::style::Visibility;

impl KeynoteEditor {
    /// Read whether Keynote shows the line for one native slide-chart axis.
    pub fn slide_chart_axis_line_visible(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<Visibility> {
        slide_chart_axis_line_visible(self, slide_index, drawable_object_id, axis)
    }

    /// Set whether Keynote shows the line for one native slide-chart axis.
    pub fn set_slide_chart_axis_line_visible(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: Axis,
        visible: Visibility,
    ) -> Result<()> {
        set_slide_chart_axis_line_visible(self, slide_index, drawable_object_id, axis, visible)
    }
}

fn slide_chart_axis_line_visible(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: Axis,
) -> Result<Visibility> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_chart_axis_line_visible(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
    )
}

fn set_slide_chart_axis_line_visible(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: Axis,
    visible: Visibility,
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_axis_line_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
        visible,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_axis_line_visible(slide_index, drawable_object_id, axis)? != visible {
        return Err(Error::InvalidFormat(
            "Keynote chart axis-line update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
