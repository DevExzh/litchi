//! Native shadow CRUD for Keynote slide charts.

use super::*;
use crate::charts::ChartShadow;
use crate::charts::shadow::{
    chart_shadow as read_native_chart_shadow, set_chart_shadow as set_native_chart_shadow,
};

impl KeynoteEditor {
    /// Read the native shadow scope and drop-shadow appearance of one chart.
    pub fn slide_chart_shadow(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ChartShadow> {
        slide_chart_shadow(self, slide_index, drawable_object_id)
    }

    /// Set the native shadow scope and drop-shadow appearance of one chart.
    pub fn set_slide_chart_shadow(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        shadow: ChartShadow,
    ) -> Result<()> {
        set_slide_chart_shadow(self, slide_index, drawable_object_id, shadow)
    }
}

fn slide_chart_shadow(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<ChartShadow> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_chart_shadow(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
    )
}

fn set_slide_chart_shadow(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    shadow: ChartShadow,
) -> Result<()> {
    if slide_chart_shadow(editor, slide_index, drawable_object_id)? == shadow {
        return Ok(());
    }
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_shadow(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        shadow,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_shadow(slide_index, drawable_object_id)? != shadow {
        return Err(Error::InvalidFormat(
            "Keynote chart shadow update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
