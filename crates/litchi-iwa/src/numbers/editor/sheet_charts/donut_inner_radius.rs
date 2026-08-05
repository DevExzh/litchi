//! Native donut inner-radius CRUD for Numbers sheet charts.

use super::*;
use crate::charts::ChartDonutInnerRadius;
use crate::charts::donut_inner_radius::{
    chart_donut_inner_radius as read_native_chart_donut_inner_radius,
    set_chart_donut_inner_radius as set_native_chart_donut_inner_radius,
};

impl NumbersEditor {
    /// Read the Segments inner radius of one donut chart.
    pub fn sheet_chart_donut_inner_radius(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ChartDonutInnerRadius> {
        sheet_chart_donut_inner_radius(self, sheet_id, drawable_object_id)
    }

    /// Set the Segments inner radius of one donut chart.
    pub fn set_sheet_chart_donut_inner_radius(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        radius: ChartDonutInnerRadius,
    ) -> Result<()> {
        set_sheet_chart_donut_inner_radius(self, sheet_id, drawable_object_id, radius)
    }
}

fn sheet_chart_donut_inner_radius(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<ChartDonutInnerRadius> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    require_donut_inner_radius(graph.info.kind, drawable_object_id)?;
    read_native_chart_donut_inner_radius(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
    )
}

fn set_sheet_chart_donut_inner_radius(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    radius: ChartDonutInnerRadius,
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    require_donut_inner_radius(graph.info.kind, drawable_object_id)?;
    if read_native_chart_donut_inner_radius(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
    )? == radius
    {
        return Ok(());
    }
    let mut staged = editor.package().clone();
    set_native_chart_donut_inner_radius(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        radius,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_donut_inner_radius(sheet_id, drawable_object_id)? != radius {
        return Err(Error::InvalidFormat(
            "Numbers chart donut inner-radius update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn require_donut_inner_radius(kind: Kind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_donut_inner_radius() {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} kind {kind:?} has no Segments inner radius"
        )));
    }
    Ok(())
}
