//! Native title CRUD for Pages body charts.

use super::*;
use crate::charts::options::{
    chart_title as read_native_chart_title, remove_chart_title as remove_native_chart_title,
    set_chart_title as set_native_chart_title,
};

impl PagesEditor {
    /// Read the chart title shown by Pages for one body chart.
    pub fn body_chart_title(&self, drawable_object_id: u64) -> Result<Option<String>> {
        body_chart_title(self, drawable_object_id)
    }

    /// Create or replace the native title shown by Pages for one body chart.
    pub fn set_body_chart_title(&mut self, drawable_object_id: u64, title: &str) -> Result<()> {
        set_body_chart_title(self, drawable_object_id, title)
    }

    /// Remove the native title shown by Pages for one body chart.
    ///
    /// Returns whether a visible title was present.
    pub fn remove_body_chart_title(&mut self, drawable_object_id: u64) -> Result<bool> {
        remove_body_chart_title(self, drawable_object_id)
    }
}

fn body_chart_title(editor: &PagesEditor, drawable_object_id: u64) -> Result<Option<String>> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_chart_title(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
    )
}

fn set_body_chart_title(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    title: &str,
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_title(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        title,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_title(drawable_object_id)?.as_deref() != Some(title) {
        return Err(Error::InvalidFormat(
            "Pages chart title update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn remove_body_chart_title(editor: &mut PagesEditor, drawable_object_id: u64) -> Result<bool> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    let removed = remove_native_chart_title(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
    )?;
    if !removed {
        return Ok(false);
    }
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_title(drawable_object_id)?.is_some() {
        return Err(Error::InvalidFormat(
            "Pages chart title removal failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(true)
}
