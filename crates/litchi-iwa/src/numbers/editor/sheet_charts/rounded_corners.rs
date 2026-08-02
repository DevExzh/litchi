//! Native rounded-corner CRUD for Numbers sheet charts.

use super::*;
use crate::charts::ChartRoundedCorners;
use crate::charts::rounded_corners::{
    chart_rounded_corners as read_native_chart_rounded_corners,
    set_chart_rounded_corners as set_native_chart_rounded_corners,
};

impl NumbersEditor {
    /// Read the rounded-corner settings for one native Numbers sheet chart.
    pub fn sheet_chart_rounded_corners(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ChartRoundedCorners> {
        sheet_chart_rounded_corners(self, sheet_id, drawable_object_id)
    }

    /// Set the rounded-corner settings for one native Numbers sheet chart.
    pub fn set_sheet_chart_rounded_corners(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        rounded_corners: ChartRoundedCorners,
    ) -> Result<()> {
        set_sheet_chart_rounded_corners(self, sheet_id, drawable_object_id, rounded_corners)
    }
}

fn sheet_chart_rounded_corners(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<ChartRoundedCorners> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_chart_rounded_corners(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
    )
}

fn set_sheet_chart_rounded_corners(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    rounded_corners: ChartRoundedCorners,
) -> Result<()> {
    if sheet_chart_rounded_corners(editor, sheet_id, drawable_object_id)? == rounded_corners {
        return Ok(());
    }
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_chart_rounded_corners(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        rounded_corners,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_rounded_corners(sheet_id, drawable_object_id)? != rounded_corners {
        return Err(Error::InvalidFormat(
            "Numbers chart rounded-corner update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
