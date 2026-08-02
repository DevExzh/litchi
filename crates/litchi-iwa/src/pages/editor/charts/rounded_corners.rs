//! Native rounded-corner CRUD for Pages body charts.

use super::*;
use crate::charts::ChartRoundedCorners;
use crate::charts::rounded_corners::{
    chart_rounded_corners as read_native_chart_rounded_corners,
    set_chart_rounded_corners as set_native_chart_rounded_corners,
};

impl PagesEditor {
    /// Read the rounded-corner settings for one native Pages body chart.
    pub fn body_chart_rounded_corners(
        &self,
        drawable_object_id: u64,
    ) -> Result<ChartRoundedCorners> {
        body_chart_rounded_corners(self, drawable_object_id)
    }

    /// Set the rounded-corner settings for one native Pages body chart.
    pub fn set_body_chart_rounded_corners(
        &mut self,
        drawable_object_id: u64,
        rounded_corners: ChartRoundedCorners,
    ) -> Result<()> {
        set_body_chart_rounded_corners(self, drawable_object_id, rounded_corners)
    }
}

fn body_chart_rounded_corners(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<ChartRoundedCorners> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_chart_rounded_corners(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
    )
}

fn set_body_chart_rounded_corners(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    rounded_corners: ChartRoundedCorners,
) -> Result<()> {
    if body_chart_rounded_corners(editor, drawable_object_id)? == rounded_corners {
        return Ok(());
    }
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_rounded_corners(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        rounded_corners,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_rounded_corners(drawable_object_id)? != rounded_corners {
        return Err(Error::InvalidFormat(
            "Pages chart rounded-corner update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
