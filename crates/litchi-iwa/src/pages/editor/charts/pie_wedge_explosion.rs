//! Native per-wedge position CRUD for Pages pie and donut body charts.

use super::*;
use crate::charts::pie_wedge_explosion::{
    chart_pie_wedge_explosions as read_native_wedge_explosions,
    set_chart_pie_wedge_explosions as set_native_wedge_explosions,
};
use crate::charts::{ChartPieWedgeExplosion, ChartPieWedgeIndex};

impl PagesEditor {
    /// Read every pie or donut wedge position in chart series order.
    pub fn body_chart_pie_wedge_explosions(
        &self,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartPieWedgeExplosion>> {
        body_chart_pie_wedge_explosions(self, drawable_object_id)
    }

    /// Read one pie or donut wedge position by zero-based series index.
    pub fn body_chart_pie_wedge_explosion(
        &self,
        drawable_object_id: u64,
        wedge: ChartPieWedgeIndex,
    ) -> Result<ChartPieWedgeExplosion> {
        let explosions = body_chart_pie_wedge_explosions(self, drawable_object_id)?;
        explosions
            .get(wedge.zero_based())
            .copied()
            .ok_or_else(|| wedge_index_error("Pages", drawable_object_id, wedge, explosions.len()))
    }

    /// Set every pie or donut wedge position in chart series order.
    pub fn set_body_chart_pie_wedge_explosions(
        &mut self,
        drawable_object_id: u64,
        explosions: &[ChartPieWedgeExplosion],
    ) -> Result<()> {
        set_body_chart_pie_wedge_explosions(self, drawable_object_id, explosions)
    }

    /// Set one pie or donut wedge position by zero-based series index.
    pub fn set_body_chart_pie_wedge_explosion(
        &mut self,
        drawable_object_id: u64,
        wedge: ChartPieWedgeIndex,
        explosion: ChartPieWedgeExplosion,
    ) -> Result<()> {
        let mut explosions = body_chart_pie_wedge_explosions(self, drawable_object_id)?;
        let explosion_count = explosions.len();
        let target = explosions.get_mut(wedge.zero_based()).ok_or_else(|| {
            wedge_index_error("Pages", drawable_object_id, wedge, explosion_count)
        })?;
        if *target == explosion {
            return Ok(());
        }
        *target = explosion;
        set_body_chart_pie_wedge_explosions(self, drawable_object_id, &explosions)
    }
}

fn body_chart_pie_wedge_explosions(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<Vec<ChartPieWedgeExplosion>> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    require_pie_wedges(graph.info.kind, drawable_object_id)?;
    let series_count = chart_series_count(
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    read_native_wedge_explosions(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        series_count,
    )
}

fn set_body_chart_pie_wedge_explosions(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    explosions: &[ChartPieWedgeExplosion],
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    require_pie_wedges(graph.info.kind, drawable_object_id)?;
    let series_count = chart_series_count(
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    if explosions.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} requires {series_count} wedge positions, got {}",
            explosions.len()
        )));
    }
    if read_native_wedge_explosions(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        series_count,
    )? == explosions
    {
        return Ok(());
    }
    let expected = explosions.to_vec();
    let mut staged = editor.package().clone();
    set_native_wedge_explosions(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        &expected,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_pie_wedge_explosions(drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Pages chart pie wedge-explosion update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn require_pie_wedges(kind: ChartKind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_pie_start_angle() {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} kind {kind:?} has no pie wedges"
        )));
    }
    Ok(())
}

fn chart_series_count(
    direction: Direction,
    data: &ChartData,
    drawable_label: &str,
    drawable_object_id: u64,
) -> Result<usize> {
    match direction.kind() {
        Some(DirectionKind::Rows) => Ok(data.row_names().len()),
        Some(DirectionKind::Columns) => Ok(data.column_names().len()),
        None => Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has unsupported series direction {}",
            direction.native_value()
        ))),
    }
}

fn wedge_index_error(
    drawable_label: &str,
    drawable_object_id: u64,
    wedge: ChartPieWedgeIndex,
    series_count: usize,
) -> Error {
    Error::InvalidFormat(format!(
        "{drawable_label} chart {drawable_object_id} wedge index {} exceeds series count {series_count}",
        wedge.zero_based()
    ))
}
