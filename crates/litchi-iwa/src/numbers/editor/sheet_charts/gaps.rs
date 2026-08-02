//! Native gap-spacing CRUD for Numbers sheet charts.

use super::*;
use crate::charts::ChartGapSpacing;
use crate::charts::gaps::{
    chart_gap_spacing as read_native_chart_gap_spacing,
    set_chart_gap_spacing as set_native_chart_gap_spacing,
};

impl NumbersEditor {
    /// Read the spacing between items and sets for one native sheet chart.
    pub fn sheet_chart_gap_spacing(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ChartGapSpacing> {
        sheet_chart_gap_spacing(self, sheet_id, drawable_object_id)
    }

    /// Set the spacing between items and sets for one native sheet chart.
    pub fn set_sheet_chart_gap_spacing(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        spacing: ChartGapSpacing,
    ) -> Result<()> {
        set_sheet_chart_gap_spacing(self, sheet_id, drawable_object_id, spacing)
    }
}

fn sheet_chart_gap_spacing(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<ChartGapSpacing> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_chart_gap_spacing(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
    )
}

fn set_sheet_chart_gap_spacing(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    spacing: ChartGapSpacing,
) -> Result<()> {
    if sheet_chart_gap_spacing(editor, sheet_id, drawable_object_id)? == spacing {
        return Ok(());
    }
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_chart_gap_spacing(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        spacing,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_gap_spacing(sheet_id, drawable_object_id)? != spacing {
        return Err(Error::InvalidFormat(
            "Numbers chart gap update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
