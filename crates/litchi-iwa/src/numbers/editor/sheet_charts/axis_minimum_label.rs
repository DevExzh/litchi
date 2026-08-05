//! Native value-axis minimum-label CRUD for Numbers sheet charts.

use super::*;
use crate::charts::axis_style::{
    chart_value_axis_minimum_label_visible as read_native_chart_value_axis_minimum_label_visible,
    set_chart_value_axis_minimum_label_visible as set_native_chart_value_axis_minimum_label_visible,
};
use litchi_iwa_common::chart::axis::style::Visibility;

impl NumbersEditor {
    /// Read whether Numbers shows the minimum value label on a native sheet chart.
    pub fn sheet_chart_value_axis_minimum_label_visible(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Visibility> {
        sheet_chart_value_axis_minimum_label_visible(self, sheet_id, drawable_object_id)
    }

    /// Set whether Numbers shows the minimum value label on a native sheet chart.
    pub fn set_sheet_chart_value_axis_minimum_label_visible(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        visible: Visibility,
    ) -> Result<()> {
        set_sheet_chart_value_axis_minimum_label_visible(
            self,
            sheet_id,
            drawable_object_id,
            visible,
        )
    }
}

fn sheet_chart_value_axis_minimum_label_visible(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<Visibility> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_chart_value_axis_minimum_label_visible(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
    )
}

fn set_sheet_chart_value_axis_minimum_label_visible(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    visible: Visibility,
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_chart_value_axis_minimum_label_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        visible,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_value_axis_minimum_label_visible(sheet_id, drawable_object_id)?
        != visible
    {
        return Err(Error::InvalidFormat(
            "Numbers chart value-axis minimum-label update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
