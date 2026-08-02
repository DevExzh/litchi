//! Native axis-title CRUD for Keynote slide charts.

use super::*;
use crate::charts::ChartAxis;
use crate::charts::axis::{
    chart_axis_title as read_native_chart_axis_title,
    remove_chart_axis_title as remove_native_chart_axis_title,
    set_chart_axis_title as set_native_chart_axis_title,
};

impl KeynoteEditor {
    /// Read the title shown by Keynote for one native slide-chart axis.
    pub fn slide_chart_axis_title(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: ChartAxis,
    ) -> Result<Option<String>> {
        slide_chart_axis_title(self, slide_index, drawable_object_id, axis)
    }

    /// Create or replace the title shown by Keynote for one native slide-chart axis.
    pub fn set_slide_chart_axis_title(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: ChartAxis,
        title: &str,
    ) -> Result<()> {
        set_slide_chart_axis_title(self, slide_index, drawable_object_id, axis, title)
    }

    /// Remove the title shown by Keynote for one native slide-chart axis.
    ///
    /// Returns whether a visible title was present.
    pub fn remove_slide_chart_axis_title(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: ChartAxis,
    ) -> Result<bool> {
        remove_slide_chart_axis_title(self, slide_index, drawable_object_id, axis)
    }
}

fn slide_chart_axis_title(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: ChartAxis,
) -> Result<Option<String>> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_chart_axis_title(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
    )
}

fn set_slide_chart_axis_title(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: ChartAxis,
    title: &str,
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_axis_title(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
        title,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified
        .slide_chart_axis_title(slide_index, drawable_object_id, axis)?
        .as_deref()
        != Some(title)
    {
        return Err(Error::InvalidFormat(
            "Keynote chart axis title update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn remove_slide_chart_axis_title(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: ChartAxis,
) -> Result<bool> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    let removed = remove_native_chart_axis_title(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
    )?;
    if !removed {
        return Ok(false);
    }
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified
        .slide_chart_axis_title(slide_index, drawable_object_id, axis)?
        .is_some()
    {
        return Err(Error::InvalidFormat(
            "Keynote chart axis title removal failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(true)
}
