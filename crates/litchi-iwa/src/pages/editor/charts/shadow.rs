//! Native shadow CRUD for Pages body charts.

use super::*;
use crate::charts::ChartShadow;
use crate::charts::shadow::{
    chart_shadow as read_native_chart_shadow, set_chart_shadow as set_native_chart_shadow,
};

impl PagesEditor {
    /// Read the native shadow scope and drop-shadow appearance of one body chart.
    pub fn body_chart_shadow(&self, drawable_object_id: u64) -> Result<ChartShadow> {
        body_chart_shadow(self, drawable_object_id)
    }

    /// Set the native shadow scope and drop-shadow appearance of one body chart.
    pub fn set_body_chart_shadow(
        &mut self,
        drawable_object_id: u64,
        shadow: ChartShadow,
    ) -> Result<()> {
        set_body_chart_shadow(self, drawable_object_id, shadow)
    }
}

fn body_chart_shadow(editor: &PagesEditor, drawable_object_id: u64) -> Result<ChartShadow> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_chart_shadow(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
    )
}

fn set_body_chart_shadow(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    shadow: ChartShadow,
) -> Result<()> {
    if body_chart_shadow(editor, drawable_object_id)? == shadow {
        return Ok(());
    }
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_shadow(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        shadow,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_shadow(drawable_object_id)? != shadow {
        return Err(Error::InvalidFormat(
            "Pages chart shadow update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
