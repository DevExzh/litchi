//! Native value-axis minimum-label CRUD for Pages body charts.

use super::*;
use crate::charts::axis_style::{
    chart_value_axis_minimum_label_visible as read_native_chart_value_axis_minimum_label_visible,
    set_chart_value_axis_minimum_label_visible as set_native_chart_value_axis_minimum_label_visible,
};
use litchi_iwa_common::chart::axis::style::Visibility;

impl PagesEditor {
    /// Read whether Pages shows the minimum value label on a native body chart.
    pub fn body_chart_value_axis_minimum_label_visible(
        &self,
        drawable_object_id: u64,
    ) -> Result<Visibility> {
        body_chart_value_axis_minimum_label_visible(self, drawable_object_id)
    }

    /// Set whether Pages shows the minimum value label on a native body chart.
    pub fn set_body_chart_value_axis_minimum_label_visible(
        &mut self,
        drawable_object_id: u64,
        visible: Visibility,
    ) -> Result<()> {
        set_body_chart_value_axis_minimum_label_visible(self, drawable_object_id, visible)
    }
}

fn body_chart_value_axis_minimum_label_visible(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<Visibility> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_chart_value_axis_minimum_label_visible(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
    )
}

fn set_body_chart_value_axis_minimum_label_visible(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    visible: Visibility,
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_value_axis_minimum_label_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        visible,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_value_axis_minimum_label_visible(drawable_object_id)? != visible {
        return Err(Error::InvalidFormat(
            "Pages chart value-axis minimum-label update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
