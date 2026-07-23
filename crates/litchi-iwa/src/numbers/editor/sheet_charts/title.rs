//! Native title CRUD for Numbers sheet charts.

use super::*;
use crate::charts::title::{
    chart_title as read_native_chart_title, remove_chart_title as remove_native_chart_title,
    set_chart_title as set_native_chart_title,
};

impl NumbersEditor {
    /// Read the chart title shown by Numbers for one sheet chart.
    pub fn sheet_chart_title(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Option<String>> {
        sheet_chart_title(self, sheet_id, drawable_object_id)
    }

    /// Create or replace the native title shown by Numbers for one sheet chart.
    pub fn set_sheet_chart_title(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        title: &str,
    ) -> Result<()> {
        set_sheet_chart_title(self, sheet_id, drawable_object_id, title)
    }

    /// Remove the native title shown by Numbers for one sheet chart.
    ///
    /// Returns whether a visible title was present.
    pub fn remove_sheet_chart_title(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        remove_sheet_chart_title(self, sheet_id, drawable_object_id)
    }
}

fn sheet_chart_title(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<Option<String>> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_chart_title(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
    )
}

fn set_sheet_chart_title(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    title: &str,
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_chart_title(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        title,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified
        .sheet_chart_title(sheet_id, drawable_object_id)?
        .as_deref()
        != Some(title)
    {
        return Err(Error::InvalidFormat(
            "Numbers chart title update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn remove_sheet_chart_title(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<bool> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    let removed = remove_native_chart_title(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
    )?;
    if !removed {
        return Ok(false);
    }
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified
        .sheet_chart_title(sheet_id, drawable_object_id)?
        .is_some()
    {
        return Err(Error::InvalidFormat(
            "Numbers chart title removal failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(true)
}
