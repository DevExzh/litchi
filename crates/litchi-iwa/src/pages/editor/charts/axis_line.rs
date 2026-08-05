//! Native axis-line visibility CRUD for Pages body charts.

use super::*;
use crate::charts::axis_style::{
    chart_axis_line_visible as read_native_chart_axis_line_visible,
    set_chart_axis_line_visible as set_native_chart_axis_line_visible,
};
use litchi_iwa_common::chart::axis::Axis;
use litchi_iwa_common::chart::axis::style::Visibility;

impl PagesEditor {
    /// Read whether Pages shows the line for one native body-chart axis.
    pub fn body_chart_axis_line_visible(
        &self,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<Visibility> {
        body_chart_axis_line_visible(self, drawable_object_id, axis)
    }

    /// Set whether Pages shows the line for one native body-chart axis.
    pub fn set_body_chart_axis_line_visible(
        &mut self,
        drawable_object_id: u64,
        axis: Axis,
        visible: Visibility,
    ) -> Result<()> {
        set_body_chart_axis_line_visible(self, drawable_object_id, axis, visible)
    }
}

fn body_chart_axis_line_visible(
    editor: &PagesEditor,
    drawable_object_id: u64,
    axis: Axis,
) -> Result<Visibility> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_chart_axis_line_visible(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        axis,
    )
}

fn set_body_chart_axis_line_visible(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    axis: Axis,
    visible: Visibility,
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_axis_line_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        axis,
        visible,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_axis_line_visible(drawable_object_id, axis)? != visible {
        return Err(Error::InvalidFormat(
            "Pages chart axis-line update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
