//! Native category-axis series-names CRUD for Pages body charts.

use super::*;
use crate::charts::axis::{
    chart_category_axis_series_names_visible as read_native_chart_category_axis_series_names_visible,
    set_chart_category_axis_series_names_visible as set_native_chart_category_axis_series_names_visible,
};

impl PagesEditor {
    /// Read whether Pages shows series names on a native body chart category axis.
    pub fn body_chart_category_axis_series_names_visible(
        &self,
        drawable_object_id: u64,
    ) -> Result<bool> {
        body_chart_category_axis_series_names_visible(self, drawable_object_id)
    }

    /// Set whether Pages shows series names on a native body chart category axis.
    pub fn set_body_chart_category_axis_series_names_visible(
        &mut self,
        drawable_object_id: u64,
        visible: bool,
    ) -> Result<()> {
        set_body_chart_category_axis_series_names_visible(self, drawable_object_id, visible)
    }
}

fn body_chart_category_axis_series_names_visible(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<bool> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_chart_category_axis_series_names_visible(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
    )
}

fn set_body_chart_category_axis_series_names_visible(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    visible: bool,
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_category_axis_series_names_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        visible,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_category_axis_series_names_visible(drawable_object_id)? != visible {
        return Err(Error::InvalidFormat(
            "Pages chart category-axis series-names update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
