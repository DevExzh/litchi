//! Native pie and donut start-angle CRUD for Pages body charts.

use super::*;
use crate::charts::ChartPieStartAngle;
use crate::charts::pie_start_angle::{
    chart_pie_start_angle as read_native_chart_pie_start_angle,
    set_chart_pie_start_angle as set_native_chart_pie_start_angle,
};

impl PagesEditor {
    /// Read the Wedges rotation angle of one pie or donut body chart.
    pub fn body_chart_pie_start_angle(
        &self,
        drawable_object_id: u64,
    ) -> Result<ChartPieStartAngle> {
        body_chart_pie_start_angle(self, drawable_object_id)
    }

    /// Set the Wedges rotation angle of one pie or donut body chart.
    pub fn set_body_chart_pie_start_angle(
        &mut self,
        drawable_object_id: u64,
        angle: ChartPieStartAngle,
    ) -> Result<()> {
        set_body_chart_pie_start_angle(self, drawable_object_id, angle)
    }
}

fn body_chart_pie_start_angle(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<ChartPieStartAngle> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    require_pie_start_angle(graph.info.kind, drawable_object_id)?;
    read_native_chart_pie_start_angle(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
    )
}

fn set_body_chart_pie_start_angle(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    angle: ChartPieStartAngle,
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    require_pie_start_angle(graph.info.kind, drawable_object_id)?;
    if read_native_chart_pie_start_angle(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
    )? == angle
    {
        return Ok(());
    }
    let mut staged = editor.package().clone();
    set_native_chart_pie_start_angle(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        angle,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_pie_start_angle(drawable_object_id)? != angle {
        return Err(Error::InvalidFormat(
            "Pages chart pie start-angle update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn require_pie_start_angle(kind: ChartKind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_pie_start_angle() {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} kind {kind:?} has no Wedges rotation angle"
        )));
    }
    Ok(())
}
