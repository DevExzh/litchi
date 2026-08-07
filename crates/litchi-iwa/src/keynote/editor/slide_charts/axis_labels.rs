//! Native axis-label visibility CRUD for Keynote slide charts.

use super::*;
use crate::charts::Axis;
use crate::charts::axis::{
    chart_axis_labels_visible as read_native_chart_axis_labels_visible,
    set_chart_axis_labels_visible as set_native_chart_axis_labels_visible,
};

impl KeynoteEditor {
    /// Read whether Keynote shows labels for one native slide-chart axis.
    pub fn slide_chart_axis_labels_visible(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<bool> {
        slide_chart_axis_labels_visible(self, slide_index, drawable_object_id, axis)
    }

    /// Set whether Keynote shows labels for one native slide-chart axis.
    pub fn set_slide_chart_axis_labels_visible(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: Axis,
        visible: bool,
    ) -> Result<()> {
        set_slide_chart_axis_labels_visible(self, slide_index, drawable_object_id, axis, visible)
    }
}

fn slide_chart_axis_labels_visible(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: Axis,
) -> Result<bool> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_chart_axis_labels_visible(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
    )
}

fn set_slide_chart_axis_labels_visible(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: Axis,
    visible: bool,
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_axis_labels_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
        visible,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_axis_labels_visible(slide_index, drawable_object_id, axis)? != visible {
        return Err(Error::InvalidFormat(
            "Keynote chart axis-label update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
