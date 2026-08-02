//! Native value-axis scale-bound CRUD for Pages body charts.

use super::*;
use crate::charts::ChartValueAxisBounds;
use crate::charts::axis_bounds::{
    chart_value_axis_bounds as read_native_chart_value_axis_bounds,
    set_chart_value_axis_bounds as set_native_chart_value_axis_bounds,
};

impl PagesEditor {
    /// Read the manual minimum and maximum of a native body chart's value axis.
    ///
    /// Missing endpoints use Pages' automatic scale calculation.
    pub fn body_chart_value_axis_bounds(
        &self,
        drawable_object_id: u64,
    ) -> Result<ChartValueAxisBounds> {
        body_chart_value_axis_bounds(self, drawable_object_id)
    }

    /// Set the manual minimum and maximum of a native body chart's value axis.
    ///
    /// Use [`ChartValueAxisBounds::automatic`] to restore Pages' automatic
    /// bounds for both endpoints.
    pub fn set_body_chart_value_axis_bounds(
        &mut self,
        drawable_object_id: u64,
        bounds: ChartValueAxisBounds,
    ) -> Result<()> {
        set_body_chart_value_axis_bounds(self, drawable_object_id, bounds)
    }
}

fn body_chart_value_axis_bounds(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<ChartValueAxisBounds> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_chart_value_axis_bounds(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
    )
}

fn set_body_chart_value_axis_bounds(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    bounds: ChartValueAxisBounds,
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_value_axis_bounds(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        bounds,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_value_axis_bounds(drawable_object_id)? != bounds {
        return Err(Error::InvalidFormat(
            "Pages chart value-axis bounds update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
