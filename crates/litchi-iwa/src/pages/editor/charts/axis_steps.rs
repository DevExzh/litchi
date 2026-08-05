//! Native value-axis scale-step CRUD for Pages body charts.

use super::*;
use crate::charts::Steps;
use crate::charts::axis_steps::{
    chart_value_axis_steps as read_native_chart_value_axis_steps,
    set_chart_value_axis_steps as set_native_chart_value_axis_steps,
};

impl PagesEditor {
    /// Read the major and minor scale steps of a native body chart's value axis.
    ///
    /// Missing step counts use Pages' automatic scale calculation.
    pub fn body_chart_value_axis_steps(&self, drawable_object_id: u64) -> Result<Steps> {
        body_chart_value_axis_steps(self, drawable_object_id)
    }

    /// Set the major and minor scale steps of a native body chart's value axis.
    ///
    /// Use [`Steps::automatic`] to restore Pages' automatic
    /// calculation for both step counts.
    pub fn set_body_chart_value_axis_steps(
        &mut self,
        drawable_object_id: u64,
        steps: Steps,
    ) -> Result<()> {
        set_body_chart_value_axis_steps(self, drawable_object_id, steps)
    }
}

fn body_chart_value_axis_steps(editor: &PagesEditor, drawable_object_id: u64) -> Result<Steps> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_chart_value_axis_steps(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
    )
}

fn set_body_chart_value_axis_steps(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    steps: Steps,
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_value_axis_steps(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        steps,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_value_axis_steps(drawable_object_id)? != steps {
        return Err(Error::InvalidFormat(
            "Pages chart value-axis steps update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
