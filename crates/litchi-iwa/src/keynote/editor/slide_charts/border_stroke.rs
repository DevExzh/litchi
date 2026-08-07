//! Native chart-area border stroke CRUD for Keynote slide charts.

use super::*;
use crate::charts::border_stroke::{
    chart_border_stroke as read_native_chart_border_stroke,
    set_chart_border_stroke as set_native_chart_border_stroke,
};
use crate::shapes::Stroke;

impl KeynoteEditor {
    /// Read the chart-area border stroke independently of border visibility.
    ///
    /// `None` represents the native inspector's empty-stroke option.
    pub fn slide_chart_border_stroke(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Option<Stroke>> {
        slide_chart_border_stroke(self, slide_index, drawable_object_id)
    }

    /// Set the chart-area border stroke independently of border visibility.
    pub fn set_slide_chart_border_stroke(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        stroke: Option<Stroke>,
    ) -> Result<()> {
        set_slide_chart_border_stroke(self, slide_index, drawable_object_id, stroke)
    }
}

fn slide_chart_border_stroke(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<Option<Stroke>> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_chart_border_stroke(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
    )
}

fn set_slide_chart_border_stroke(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    stroke: Option<Stroke>,
) -> Result<()> {
    if slide_chart_border_stroke(editor, slide_index, drawable_object_id)? == stroke {
        return Ok(());
    }
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_border_stroke(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        stroke,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_border_stroke(slide_index, drawable_object_id)? != stroke {
        return Err(Error::InvalidFormat(
            "Keynote chart border stroke update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
