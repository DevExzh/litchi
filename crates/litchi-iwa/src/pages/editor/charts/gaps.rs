//! Native gap-spacing CRUD for Pages body charts.

use super::*;
use crate::charts::ChartGapSpacing;
use crate::charts::gaps::{
    chart_gap_spacing as read_native_chart_gap_spacing,
    set_chart_gap_spacing as set_native_chart_gap_spacing,
};

impl PagesEditor {
    /// Read the spacing between items and sets for one native body chart.
    pub fn body_chart_gap_spacing(&self, drawable_object_id: u64) -> Result<ChartGapSpacing> {
        body_chart_gap_spacing(self, drawable_object_id)
    }

    /// Set the spacing between items and sets for one native body chart.
    pub fn set_body_chart_gap_spacing(
        &mut self,
        drawable_object_id: u64,
        spacing: ChartGapSpacing,
    ) -> Result<()> {
        set_body_chart_gap_spacing(self, drawable_object_id, spacing)
    }
}

fn body_chart_gap_spacing(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<ChartGapSpacing> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_chart_gap_spacing(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
    )
}

fn set_body_chart_gap_spacing(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    spacing: ChartGapSpacing,
) -> Result<()> {
    if body_chart_gap_spacing(editor, drawable_object_id)? == spacing {
        return Ok(());
    }
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_gap_spacing(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        spacing,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_gap_spacing(drawable_object_id)? != spacing {
        return Err(Error::InvalidFormat(
            "Pages chart gap update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
