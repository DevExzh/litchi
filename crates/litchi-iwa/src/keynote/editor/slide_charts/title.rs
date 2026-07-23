//! Native title CRUD for Keynote slide charts.

use super::*;
use crate::charts::options::{
    chart_title as read_native_chart_title, remove_chart_title as remove_native_chart_title,
    set_chart_title as set_native_chart_title,
};

impl KeynoteEditor {
    /// Read the chart title shown by Keynote for one slide chart.
    pub fn slide_chart_title(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Option<String>> {
        slide_chart_title(self, slide_index, drawable_object_id)
    }

    /// Create or replace the native title shown by Keynote for one slide chart.
    pub fn set_slide_chart_title(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        title: &str,
    ) -> Result<()> {
        set_slide_chart_title(self, slide_index, drawable_object_id, title)
    }

    /// Remove the native title shown by Keynote for one slide chart.
    ///
    /// Returns whether a visible title was present.
    pub fn remove_slide_chart_title(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        remove_slide_chart_title(self, slide_index, drawable_object_id)
    }
}

fn slide_chart_title(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<Option<String>> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_chart_title(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
    )
}

fn set_slide_chart_title(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    title: &str,
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_title(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        title,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified
        .slide_chart_title(slide_index, drawable_object_id)?
        .as_deref()
        != Some(title)
    {
        return Err(Error::InvalidFormat(
            "Keynote chart title update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn remove_slide_chart_title(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<bool> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    let removed = remove_native_chart_title(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
    )?;
    if !removed {
        return Ok(false);
    }
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified
        .slide_chart_title(slide_index, drawable_object_id)?
        .is_some()
    {
        return Err(Error::InvalidFormat(
            "Keynote chart title removal failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(true)
}
