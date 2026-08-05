//! Native donut inner-radius CRUD for Pages body charts.

use super::*;
use crate::charts::ChartDonutInnerRadius;
use crate::charts::donut_inner_radius::{
    chart_donut_inner_radius as read_native_chart_donut_inner_radius,
    set_chart_donut_inner_radius as set_native_chart_donut_inner_radius,
};

impl PagesEditor {
    /// Read the Segments inner radius of one donut body chart.
    pub fn body_chart_donut_inner_radius(
        &self,
        drawable_object_id: u64,
    ) -> Result<ChartDonutInnerRadius> {
        body_chart_donut_inner_radius(self, drawable_object_id)
    }

    /// Set the Segments inner radius of one donut body chart.
    pub fn set_body_chart_donut_inner_radius(
        &mut self,
        drawable_object_id: u64,
        radius: ChartDonutInnerRadius,
    ) -> Result<()> {
        set_body_chart_donut_inner_radius(self, drawable_object_id, radius)
    }
}

fn body_chart_donut_inner_radius(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<ChartDonutInnerRadius> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    require_donut_inner_radius(graph.info.kind, drawable_object_id)?;
    read_native_chart_donut_inner_radius(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
    )
}

fn set_body_chart_donut_inner_radius(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    radius: ChartDonutInnerRadius,
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    require_donut_inner_radius(graph.info.kind, drawable_object_id)?;
    if read_native_chart_donut_inner_radius(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
    )? == radius
    {
        return Ok(());
    }
    let mut staged = editor.package().clone();
    set_native_chart_donut_inner_radius(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        radius,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_donut_inner_radius(drawable_object_id)? != radius {
        return Err(Error::InvalidFormat(
            "Pages chart donut inner-radius update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn require_donut_inner_radius(kind: Kind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_donut_inner_radius() {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} kind {kind:?} has no Segments inner radius"
        )));
    }
    Ok(())
}
