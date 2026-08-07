//! Native chart-area border stroke CRUD for Numbers sheet charts.

use super::*;
use crate::charts::border_stroke::{
    chart_border_stroke as read_native_chart_border_stroke,
    set_chart_border_stroke as set_native_chart_border_stroke,
};
use crate::shapes::Stroke;

impl NumbersEditor {
    /// Read the chart-area border stroke independently of border visibility.
    ///
    /// `None` represents the native inspector's empty-stroke option.
    pub fn sheet_chart_border_stroke(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Option<Stroke>> {
        sheet_chart_border_stroke(self, sheet_id, drawable_object_id)
    }

    /// Set the chart-area border stroke independently of border visibility.
    pub fn set_sheet_chart_border_stroke(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        stroke: Option<Stroke>,
    ) -> Result<()> {
        set_sheet_chart_border_stroke(self, sheet_id, drawable_object_id, stroke)
    }
}

fn sheet_chart_border_stroke(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<Option<Stroke>> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_chart_border_stroke(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
    )
}

fn set_sheet_chart_border_stroke(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    stroke: Option<Stroke>,
) -> Result<()> {
    if sheet_chart_border_stroke(editor, sheet_id, drawable_object_id)? == stroke {
        return Ok(());
    }
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_chart_border_stroke(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        stroke,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_border_stroke(sheet_id, drawable_object_id)? != stroke {
        return Err(Error::InvalidFormat(
            "Numbers chart border stroke update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
