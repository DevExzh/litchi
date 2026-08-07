//! Native value-axis scale CRUD for Keynote slide charts.

use super::*;
use crate::charts::Scale;
use crate::charts::axis_scale::{
    chart_value_axis_scale as read_native_chart_value_axis_scale,
    set_chart_value_axis_scale as set_native_chart_value_axis_scale,
};

impl KeynoteEditor {
    /// Read the primary value-axis scale of one native slide chart.
    pub fn slide_chart_value_axis_scale(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Scale> {
        slide_chart_value_axis_scale(self, slide_index, drawable_object_id)
    }

    /// Set the primary value-axis scale of one native slide chart.
    pub fn set_slide_chart_value_axis_scale(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        scale: Scale,
    ) -> Result<()> {
        set_slide_chart_value_axis_scale(self, slide_index, drawable_object_id, scale)
    }
}

fn slide_chart_value_axis_scale(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<Scale> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_chart_value_axis_scale(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
    )
}

fn set_slide_chart_value_axis_scale(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    scale: Scale,
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_value_axis_scale(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        scale,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_value_axis_scale(slide_index, drawable_object_id)? != scale {
        return Err(Error::InvalidFormat(
            "Keynote chart value-axis scale update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
