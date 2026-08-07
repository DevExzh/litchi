//! Native pie and donut start-angle CRUD for Numbers sheet charts.

use super::*;
use crate::charts::ChartPieStartAngle;
use crate::charts::pie_start_angle::{
    chart_pie_start_angle as read_native_chart_pie_start_angle,
    set_chart_pie_start_angle as set_native_chart_pie_start_angle,
};

impl NumbersEditor {
    /// Read the Wedges rotation angle of one pie or donut chart.
    pub fn sheet_chart_pie_start_angle(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ChartPieStartAngle> {
        sheet_chart_pie_start_angle(self, sheet_id, drawable_object_id)
    }

    /// Set the Wedges rotation angle of one pie or donut chart.
    pub fn set_sheet_chart_pie_start_angle(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        angle: ChartPieStartAngle,
    ) -> Result<()> {
        set_sheet_chart_pie_start_angle(self, sheet_id, drawable_object_id, angle)
    }
}

fn sheet_chart_pie_start_angle(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<ChartPieStartAngle> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    require_pie_start_angle(graph.info.kind, drawable_object_id)?;
    read_native_chart_pie_start_angle(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
    )
}

fn set_sheet_chart_pie_start_angle(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    angle: ChartPieStartAngle,
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    require_pie_start_angle(graph.info.kind, drawable_object_id)?;
    if read_native_chart_pie_start_angle(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
    )? == angle
    {
        return Ok(());
    }
    let mut staged = editor.package().clone();
    set_native_chart_pie_start_angle(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        angle,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_pie_start_angle(sheet_id, drawable_object_id)? != angle {
        return Err(Error::InvalidFormat(
            "Numbers chart pie start-angle update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn require_pie_start_angle(kind: Kind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_pie_start_angle() {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} kind {kind:?} has no Wedges rotation angle"
        )));
    }
    Ok(())
}
