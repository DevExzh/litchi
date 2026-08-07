//! Native pie and donut start-angle CRUD for Keynote slide charts.

use super::*;
use crate::charts::ChartPieStartAngle;
use crate::charts::pie_start_angle::{
    chart_pie_start_angle as read_native_chart_pie_start_angle,
    set_chart_pie_start_angle as set_native_chart_pie_start_angle,
};

impl KeynoteEditor {
    /// Read the Wedges rotation angle of one pie or donut chart.
    pub fn slide_chart_pie_start_angle(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ChartPieStartAngle> {
        slide_chart_pie_start_angle(self, slide_index, drawable_object_id)
    }

    /// Set the Wedges rotation angle of one pie or donut chart.
    pub fn set_slide_chart_pie_start_angle(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        angle: ChartPieStartAngle,
    ) -> Result<()> {
        set_slide_chart_pie_start_angle(self, slide_index, drawable_object_id, angle)
    }
}

fn slide_chart_pie_start_angle(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<ChartPieStartAngle> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    require_pie_start_angle(graph.info.kind, drawable_object_id)?;
    read_native_chart_pie_start_angle(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
    )
}

fn set_slide_chart_pie_start_angle(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    angle: ChartPieStartAngle,
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    require_pie_start_angle(graph.info.kind, drawable_object_id)?;
    if read_native_chart_pie_start_angle(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
    )? == angle
    {
        return Ok(());
    }
    let mut staged = editor.package().clone();
    set_native_chart_pie_start_angle(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        angle,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_pie_start_angle(slide_index, drawable_object_id)? != angle {
        return Err(Error::InvalidFormat(
            "Keynote chart pie start-angle update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn require_pie_start_angle(kind: Kind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_pie_start_angle() {
        return Err(Error::InvalidFormat(format!(
            "Keynote chart {drawable_object_id} kind {kind:?} has no Wedges rotation angle"
        )));
    }
    Ok(())
}
