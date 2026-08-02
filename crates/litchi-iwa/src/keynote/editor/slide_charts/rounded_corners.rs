//! Native rounded-corner CRUD for Keynote slide charts.

use super::*;
use crate::charts::ChartRoundedCorners;
use crate::charts::rounded_corners::{
    chart_rounded_corners as read_native_chart_rounded_corners,
    set_chart_rounded_corners as set_native_chart_rounded_corners,
};

impl KeynoteEditor {
    /// Read the rounded-corner settings for one native Keynote slide chart.
    pub fn slide_chart_rounded_corners(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ChartRoundedCorners> {
        slide_chart_rounded_corners(self, slide_index, drawable_object_id)
    }

    /// Set the rounded-corner settings for one native Keynote slide chart.
    pub fn set_slide_chart_rounded_corners(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        rounded_corners: ChartRoundedCorners,
    ) -> Result<()> {
        set_slide_chart_rounded_corners(self, slide_index, drawable_object_id, rounded_corners)
    }
}

fn slide_chart_rounded_corners(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<ChartRoundedCorners> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_chart_rounded_corners(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
    )
}

fn set_slide_chart_rounded_corners(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    rounded_corners: ChartRoundedCorners,
) -> Result<()> {
    if slide_chart_rounded_corners(editor, slide_index, drawable_object_id)? == rounded_corners {
        return Ok(());
    }
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_rounded_corners(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        rounded_corners,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_rounded_corners(slide_index, drawable_object_id)? != rounded_corners {
        return Err(Error::InvalidFormat(
            "Keynote chart rounded-corner update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
