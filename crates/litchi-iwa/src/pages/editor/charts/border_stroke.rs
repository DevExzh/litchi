//! Native chart-area border stroke CRUD for Pages body charts.

use super::*;
use crate::charts::border_stroke::{
    chart_border_stroke as read_native_chart_border_stroke,
    set_chart_border_stroke as set_native_chart_border_stroke,
};
use crate::shapes::Stroke;

impl PagesEditor {
    /// Read the chart-area border stroke independently of border visibility.
    ///
    /// `None` represents the native inspector's empty-stroke option.
    pub fn body_chart_border_stroke(&self, drawable_object_id: u64) -> Result<Option<Stroke>> {
        body_chart_border_stroke(self, drawable_object_id)
    }

    /// Set the chart-area border stroke independently of border visibility.
    pub fn set_body_chart_border_stroke(
        &mut self,
        drawable_object_id: u64,
        stroke: Option<Stroke>,
    ) -> Result<()> {
        set_body_chart_border_stroke(self, drawable_object_id, stroke)
    }
}

fn body_chart_border_stroke(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<Option<Stroke>> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_chart_border_stroke(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
    )
}

fn set_body_chart_border_stroke(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    stroke: Option<Stroke>,
) -> Result<()> {
    if body_chart_border_stroke(editor, drawable_object_id)? == stroke {
        return Ok(());
    }
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_border_stroke(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        stroke,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_border_stroke(drawable_object_id)? != stroke {
        return Err(Error::InvalidFormat(
            "Pages chart border stroke update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
