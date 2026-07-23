//! Native axis-title CRUD for Pages body charts.

use super::*;
use crate::charts::ChartAxis;
use crate::charts::axis::{
    chart_axis_title as read_native_chart_axis_title,
    remove_chart_axis_title as remove_native_chart_axis_title,
    set_chart_axis_title as set_native_chart_axis_title,
};

impl PagesEditor {
    /// Read the title shown by Pages for one native body-chart axis.
    pub fn body_chart_axis_title(
        &self,
        drawable_object_id: u64,
        axis: ChartAxis,
    ) -> Result<Option<String>> {
        body_chart_axis_title(self, drawable_object_id, axis)
    }

    /// Create or replace the title shown by Pages for one native body-chart axis.
    pub fn set_body_chart_axis_title(
        &mut self,
        drawable_object_id: u64,
        axis: ChartAxis,
        title: &str,
    ) -> Result<()> {
        set_body_chart_axis_title(self, drawable_object_id, axis, title)
    }

    /// Remove the title shown by Pages for one native body-chart axis.
    ///
    /// Returns whether a visible title was present.
    pub fn remove_body_chart_axis_title(
        &mut self,
        drawable_object_id: u64,
        axis: ChartAxis,
    ) -> Result<bool> {
        remove_body_chart_axis_title(self, drawable_object_id, axis)
    }
}

fn body_chart_axis_title(
    editor: &PagesEditor,
    drawable_object_id: u64,
    axis: ChartAxis,
) -> Result<Option<String>> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_chart_axis_title(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        axis,
    )
}

fn set_body_chart_axis_title(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    axis: ChartAxis,
    title: &str,
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_axis_title(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        axis,
        title,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified
        .body_chart_axis_title(drawable_object_id, axis)?
        .as_deref()
        != Some(title)
    {
        return Err(Error::InvalidFormat(
            "Pages chart axis title update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn remove_body_chart_axis_title(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    axis: ChartAxis,
) -> Result<bool> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    let removed = remove_native_chart_axis_title(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        axis,
    )?;
    if !removed {
        return Ok(false);
    }
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified
        .body_chart_axis_title(drawable_object_id, axis)?
        .is_some()
    {
        return Err(Error::InvalidFormat(
            "Pages chart axis title removal failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(true)
}
