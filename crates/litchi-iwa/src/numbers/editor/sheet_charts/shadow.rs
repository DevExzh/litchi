//! Native shadow CRUD for Numbers sheet charts.

use super::*;
use crate::charts::ChartShadow;
use crate::charts::shadow::{
    chart_shadow as read_native_chart_shadow, set_chart_shadow as set_native_chart_shadow,
};

impl NumbersEditor {
    /// Read the native shadow scope and drop-shadow appearance of one chart.
    pub fn sheet_chart_shadow(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ChartShadow> {
        sheet_chart_shadow(self, sheet_id, drawable_object_id)
    }

    /// Set the native shadow scope and drop-shadow appearance of one chart.
    pub fn set_sheet_chart_shadow(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        shadow: ChartShadow,
    ) -> Result<()> {
        set_sheet_chart_shadow(self, sheet_id, drawable_object_id, shadow)
    }
}

fn sheet_chart_shadow(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<ChartShadow> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_chart_shadow(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
    )
}

fn set_sheet_chart_shadow(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    shadow: ChartShadow,
) -> Result<()> {
    if sheet_chart_shadow(editor, sheet_id, drawable_object_id)? == shadow {
        return Ok(());
    }
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_shadow(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        shadow,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_shadow(sheet_id, drawable_object_id)? != shadow {
        return Err(Error::InvalidFormat(
            "Numbers chart shadow update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
