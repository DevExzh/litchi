//! Native value-axis scale CRUD for Pages body charts.

use super::*;
use crate::charts::ChartValueAxisScale;
use crate::charts::axis_scale::{
    chart_value_axis_scale as read_native_chart_value_axis_scale,
    set_chart_value_axis_scale as set_native_chart_value_axis_scale,
};

impl PagesEditor {
    /// Read the primary value-axis scale of one native body chart.
    pub fn body_chart_value_axis_scale(
        &self,
        drawable_object_id: u64,
    ) -> Result<ChartValueAxisScale> {
        body_chart_value_axis_scale(self, drawable_object_id)
    }

    /// Set the primary value-axis scale of one native body chart.
    pub fn set_body_chart_value_axis_scale(
        &mut self,
        drawable_object_id: u64,
        scale: ChartValueAxisScale,
    ) -> Result<()> {
        set_body_chart_value_axis_scale(self, drawable_object_id, scale)
    }
}

fn body_chart_value_axis_scale(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<ChartValueAxisScale> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_chart_value_axis_scale(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
    )
}

fn set_body_chart_value_axis_scale(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    scale: ChartValueAxisScale,
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_value_axis_scale(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        scale,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_value_axis_scale(drawable_object_id)? != scale {
        return Err(Error::InvalidFormat(
            "Pages chart value-axis scale update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
