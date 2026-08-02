//! Native value-axis scale-step CRUD for Numbers sheet charts.

use super::*;
use crate::charts::ChartValueAxisSteps;
use crate::charts::axis_steps::{
    chart_value_axis_steps as read_native_chart_value_axis_steps,
    set_chart_value_axis_steps as set_native_chart_value_axis_steps,
};

impl NumbersEditor {
    /// Read the major and minor scale steps of a native sheet chart's value axis.
    ///
    /// Missing step counts use Numbers' automatic scale calculation.
    pub fn sheet_chart_value_axis_steps(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ChartValueAxisSteps> {
        sheet_chart_value_axis_steps(self, sheet_id, drawable_object_id)
    }

    /// Set the major and minor scale steps of a native sheet chart's value axis.
    ///
    /// Use [`ChartValueAxisSteps::automatic`] to restore Numbers' automatic
    /// calculation for both step counts.
    pub fn set_sheet_chart_value_axis_steps(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        steps: ChartValueAxisSteps,
    ) -> Result<()> {
        set_sheet_chart_value_axis_steps(self, sheet_id, drawable_object_id, steps)
    }
}

fn sheet_chart_value_axis_steps(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<ChartValueAxisSteps> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_chart_value_axis_steps(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
    )
}

fn set_sheet_chart_value_axis_steps(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    steps: ChartValueAxisSteps,
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_chart_value_axis_steps(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        steps,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_value_axis_steps(sheet_id, drawable_object_id)? != steps {
        return Err(Error::InvalidFormat(
            "Numbers chart value-axis steps update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
