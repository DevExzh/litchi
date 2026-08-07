//! Native axis-title CRUD for Numbers sheet charts.

use super::*;
use crate::charts::Axis;
use crate::charts::axis::{
    chart_axis_title as read_native_chart_axis_title,
    remove_chart_axis_title as remove_native_chart_axis_title,
    set_chart_axis_title as set_native_chart_axis_title,
};

impl NumbersEditor {
    /// Read the title shown by Numbers for one native sheet-chart axis.
    pub fn sheet_chart_axis_title(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<Option<String>> {
        sheet_chart_axis_title(self, sheet_id, drawable_object_id, axis)
    }

    /// Create or replace the title shown by Numbers for one native sheet-chart axis.
    pub fn set_sheet_chart_axis_title(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: Axis,
        title: &str,
    ) -> Result<()> {
        set_sheet_chart_axis_title(self, sheet_id, drawable_object_id, axis, title)
    }

    /// Remove the title shown by Numbers for one native sheet-chart axis.
    ///
    /// Returns whether a visible title was present.
    pub fn remove_sheet_chart_axis_title(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<bool> {
        remove_sheet_chart_axis_title(self, sheet_id, drawable_object_id, axis)
    }
}

fn sheet_chart_axis_title(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    axis: Axis,
) -> Result<Option<String>> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_chart_axis_title(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        axis,
    )
}

fn set_sheet_chart_axis_title(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    axis: Axis,
    title: &str,
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_chart_axis_title(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        axis,
        title,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified
        .sheet_chart_axis_title(sheet_id, drawable_object_id, axis)?
        .as_deref()
        != Some(title)
    {
        return Err(Error::InvalidFormat(
            "Numbers chart axis title update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn remove_sheet_chart_axis_title(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    axis: Axis,
) -> Result<bool> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    let removed = remove_native_chart_axis_title(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        axis,
    )?;
    if !removed {
        return Ok(false);
    }
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified
        .sheet_chart_axis_title(sheet_id, drawable_object_id, axis)?
        .is_some()
    {
        return Err(Error::InvalidFormat(
            "Numbers chart axis title removal failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(true)
}
