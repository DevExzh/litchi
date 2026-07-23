//! Native chart-border CRUD for Numbers sheet charts.

use super::*;
use crate::charts::border::{
    chart_border_visible as read_native_chart_border_visible,
    set_chart_border_visible as set_native_chart_border_visible,
};

impl NumbersEditor {
    /// Read whether Numbers shows the chart-area border for one native sheet chart.
    pub fn sheet_chart_border_visible(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        sheet_chart_border_visible(self, sheet_id, drawable_object_id)
    }

    /// Set whether Numbers shows the chart-area border for one native sheet chart.
    pub fn set_sheet_chart_border_visible(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        visible: bool,
    ) -> Result<()> {
        set_sheet_chart_border_visible(self, sheet_id, drawable_object_id, visible)
    }
}

fn sheet_chart_border_visible(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<bool> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_chart_border_visible(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
    )
}

fn set_sheet_chart_border_visible(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    visible: bool,
) -> Result<()> {
    if sheet_chart_border_visible(editor, sheet_id, drawable_object_id)? == visible {
        return Ok(());
    }
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_chart_border_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        visible,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_border_visible(sheet_id, drawable_object_id)? != visible {
        return Err(Error::InvalidFormat(
            "Numbers chart border update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
