//! Native gap-spacing CRUD for Keynote slide charts.

use super::*;
use litchi_iwa_common::chart::gaps::Spacing;

use crate::charts::gaps::{
    chart_gap_spacing as read_native_chart_gap_spacing,
    set_chart_gap_spacing as set_native_chart_gap_spacing,
};

impl KeynoteEditor {
    /// Read the spacing between items and sets for one native slide chart.
    pub fn slide_chart_gap_spacing(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Spacing> {
        slide_chart_gap_spacing(self, slide_index, drawable_object_id)
    }

    /// Set the spacing between items and sets for one native slide chart.
    pub fn set_slide_chart_gap_spacing(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        spacing: Spacing,
    ) -> Result<()> {
        set_slide_chart_gap_spacing(self, slide_index, drawable_object_id, spacing)
    }
}

fn slide_chart_gap_spacing(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<Spacing> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_chart_gap_spacing(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
    )
}

fn set_slide_chart_gap_spacing(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    spacing: Spacing,
) -> Result<()> {
    if slide_chart_gap_spacing(editor, slide_index, drawable_object_id)? == spacing {
        return Ok(());
    }
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_gap_spacing(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        spacing,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_gap_spacing(slide_index, drawable_object_id)? != spacing {
        return Err(Error::InvalidFormat(
            "Keynote chart gap update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
