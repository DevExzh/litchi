//! Native category-axis series-names CRUD for Numbers sheet charts.

use super::*;
use crate::charts::axis::{
    chart_category_axis_series_names_visible as read_native_chart_category_axis_series_names_visible,
    set_chart_category_axis_series_names_visible as set_native_chart_category_axis_series_names_visible,
};

impl NumbersEditor {
    /// Read whether Numbers shows series names on a native sheet chart category axis.
    pub fn sheet_chart_category_axis_series_names_visible(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        sheet_chart_category_axis_series_names_visible(self, sheet_id, drawable_object_id)
    }

    /// Set whether Numbers shows series names on a native sheet chart category axis.
    pub fn set_sheet_chart_category_axis_series_names_visible(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        visible: bool,
    ) -> Result<()> {
        set_sheet_chart_category_axis_series_names_visible(
            self,
            sheet_id,
            drawable_object_id,
            visible,
        )
    }
}

fn sheet_chart_category_axis_series_names_visible(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<bool> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_chart_category_axis_series_names_visible(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
    )
}

fn set_sheet_chart_category_axis_series_names_visible(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    visible: bool,
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_chart_category_axis_series_names_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        visible,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_category_axis_series_names_visible(sheet_id, drawable_object_id)?
        != visible
    {
        return Err(Error::InvalidFormat(
            "Numbers chart category-axis series-names update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
