//! Native value-axis scale-step CRUD for Keynote slide charts.

use super::*;
use crate::charts::ChartValueAxisSteps;
use crate::charts::axis_steps::{
    chart_value_axis_steps as read_native_chart_value_axis_steps,
    set_chart_value_axis_steps as set_native_chart_value_axis_steps,
};

impl KeynoteEditor {
    /// Read the major and minor scale steps of a native slide chart's value axis.
    ///
    /// Missing step counts use Keynote's automatic scale calculation.
    pub fn slide_chart_value_axis_steps(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ChartValueAxisSteps> {
        slide_chart_value_axis_steps(self, slide_index, drawable_object_id)
    }

    /// Set the major and minor scale steps of a native slide chart's value axis.
    ///
    /// Use [`ChartValueAxisSteps::automatic`] to restore Keynote's automatic
    /// calculation for both step counts.
    pub fn set_slide_chart_value_axis_steps(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        steps: ChartValueAxisSteps,
    ) -> Result<()> {
        set_slide_chart_value_axis_steps(self, slide_index, drawable_object_id, steps)
    }
}

fn slide_chart_value_axis_steps(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<ChartValueAxisSteps> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_chart_value_axis_steps(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
    )
}

fn set_slide_chart_value_axis_steps(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    steps: ChartValueAxisSteps,
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_value_axis_steps(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        steps,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_value_axis_steps(slide_index, drawable_object_id)? != steps {
        return Err(Error::InvalidFormat(
            "Keynote chart value-axis steps update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
