//! Native hidden-row and hidden-column participation CRUD for Numbers charts.

use super::*;
use crate::charts::hidden_data::{
    chart_includes_hidden_data as read_native_hidden_data,
    set_chart_includes_hidden_data as set_native_hidden_data,
};

impl NumbersEditor {
    /// Read whether a sheet chart includes data from hidden rows and columns.
    pub fn sheet_chart_includes_hidden_data(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        sheet_chart_includes_hidden_data(self, sheet_id, drawable_object_id)
    }

    /// Set whether a sheet chart includes data from hidden rows and columns.
    pub fn set_sheet_chart_includes_hidden_data(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        includes_hidden_data: bool,
    ) -> Result<()> {
        set_sheet_chart_includes_hidden_data(
            self,
            sheet_id,
            drawable_object_id,
            includes_hidden_data,
        )
    }
}

fn sheet_chart_includes_hidden_data(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<bool> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_hidden_data(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
    )
}

fn set_sheet_chart_includes_hidden_data(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    includes_hidden_data: bool,
) -> Result<()> {
    if sheet_chart_includes_hidden_data(editor, sheet_id, drawable_object_id)?
        == includes_hidden_data
    {
        return Ok(());
    }
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_hidden_data(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        includes_hidden_data,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_includes_hidden_data(sheet_id, drawable_object_id)?
        != includes_hidden_data
    {
        return Err(Error::InvalidFormat(
            "Numbers chart hidden-data update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
