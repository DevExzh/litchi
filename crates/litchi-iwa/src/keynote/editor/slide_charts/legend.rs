//! Native legend visibility CRUD for Keynote slide charts.

use super::*;
use crate::charts::options::{
    chart_legend_visible as read_native_chart_legend_visible,
    set_chart_legend_visible as set_native_chart_legend_visible,
};

impl KeynoteEditor {
    /// Read whether Keynote shows the native legend for one slide chart.
    pub fn slide_chart_legend_visible(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        slide_chart_legend_visible(self, slide_index, drawable_object_id)
    }

    /// Set whether Keynote shows the native legend for one slide chart.
    pub fn set_slide_chart_legend_visible(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        visible: bool,
    ) -> Result<()> {
        set_slide_chart_legend_visible(self, slide_index, drawable_object_id, visible)
    }
}

fn slide_chart_legend_visible(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<bool> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_chart_legend_visible(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
    )
}

fn set_slide_chart_legend_visible(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    visible: bool,
) -> Result<()> {
    if slide_chart_legend_visible(editor, slide_index, drawable_object_id)? == visible {
        return Ok(());
    }
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_legend_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        visible,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_legend_visible(slide_index, drawable_object_id)? != visible {
        return Err(Error::InvalidFormat(
            "Keynote chart legend update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
