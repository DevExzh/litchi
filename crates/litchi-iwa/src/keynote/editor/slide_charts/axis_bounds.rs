//! Native value-axis scale-bound CRUD for Keynote slide charts.

use super::*;
use crate::charts::Bounds;
use crate::charts::axis_bounds::{
    chart_value_axis_bounds as read_native_chart_value_axis_bounds,
    set_chart_value_axis_bounds as set_native_chart_value_axis_bounds,
};

impl KeynoteEditor {
    /// Read the manual minimum and maximum of a native slide chart's value axis.
    ///
    /// Missing endpoints use Keynote's automatic scale calculation.
    pub fn slide_chart_value_axis_bounds(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Bounds> {
        slide_chart_value_axis_bounds(self, slide_index, drawable_object_id)
    }

    /// Set the manual minimum and maximum of a native slide chart's value axis.
    ///
    /// Use [`Bounds::automatic`] to restore Keynote's automatic
    /// bounds for both endpoints.
    pub fn set_slide_chart_value_axis_bounds(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        bounds: Bounds,
    ) -> Result<()> {
        set_slide_chart_value_axis_bounds(self, slide_index, drawable_object_id, bounds)
    }
}

fn slide_chart_value_axis_bounds(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<Bounds> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_chart_value_axis_bounds(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
    )
}

fn set_slide_chart_value_axis_bounds(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    bounds: Bounds,
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_value_axis_bounds(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        bounds,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_value_axis_bounds(slide_index, drawable_object_id)? != bounds {
        return Err(Error::InvalidFormat(
            "Keynote chart value-axis bounds update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
