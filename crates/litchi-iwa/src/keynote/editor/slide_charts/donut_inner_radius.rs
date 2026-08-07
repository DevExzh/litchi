//! Native donut inner-radius CRUD for Keynote slide charts.

use super::*;
use crate::charts::ChartDonutInnerRadius;
use crate::charts::donut_inner_radius::{
    chart_donut_inner_radius as read_native_chart_donut_inner_radius,
    set_chart_donut_inner_radius as set_native_chart_donut_inner_radius,
};

impl KeynoteEditor {
    /// Read the Segments inner radius of one donut chart.
    pub fn slide_chart_donut_inner_radius(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ChartDonutInnerRadius> {
        slide_chart_donut_inner_radius(self, slide_index, drawable_object_id)
    }

    /// Set the Segments inner radius of one donut chart.
    pub fn set_slide_chart_donut_inner_radius(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        radius: ChartDonutInnerRadius,
    ) -> Result<()> {
        set_slide_chart_donut_inner_radius(self, slide_index, drawable_object_id, radius)
    }
}

fn slide_chart_donut_inner_radius(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<ChartDonutInnerRadius> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    require_donut_inner_radius(graph.info.kind, drawable_object_id)?;
    read_native_chart_donut_inner_radius(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
    )
}

fn set_slide_chart_donut_inner_radius(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    radius: ChartDonutInnerRadius,
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    require_donut_inner_radius(graph.info.kind, drawable_object_id)?;
    if read_native_chart_donut_inner_radius(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
    )? == radius
    {
        return Ok(());
    }
    let mut staged = editor.package().clone();
    set_native_chart_donut_inner_radius(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        radius,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_donut_inner_radius(slide_index, drawable_object_id)? != radius {
        return Err(Error::InvalidFormat(
            "Keynote chart donut inner-radius update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn require_donut_inner_radius(kind: Kind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_donut_inner_radius() {
        return Err(Error::InvalidFormat(format!(
            "Keynote chart {drawable_object_id} kind {kind:?} has no Segments inner radius"
        )));
    }
    Ok(())
}
