//! Native per-wedge leader-line visibility CRUD for Pages pie and donut charts.

use super::*;
use crate::charts::pie_leader_lines::{
    chart_pie_leader_line_visibilities as read_native_leader_line_visibilities,
    set_chart_pie_leader_line_visibilities as set_native_leader_line_visibilities,
};
use crate::charts::{ChartPieWedgeIndex, LeaderLineVisibility};

impl PagesEditor {
    /// Read every pie or donut wedge's leader-line visibility.
    pub fn body_chart_pie_leader_line_visibilities(
        &self,
        drawable_object_id: u64,
    ) -> Result<Vec<LeaderLineVisibility>> {
        body_chart_pie_leader_line_visibilities(self, drawable_object_id)
    }

    /// Read one pie or donut wedge's leader-line visibility.
    pub fn body_chart_pie_leader_line_visibility(
        &self,
        drawable_object_id: u64,
        wedge: ChartPieWedgeIndex,
    ) -> Result<LeaderLineVisibility> {
        let visibilities = body_chart_pie_leader_line_visibilities(self, drawable_object_id)?;
        visibilities
            .get(wedge.zero_based())
            .copied()
            .ok_or_else(|| {
                leader_line_index_error("Pages", drawable_object_id, wedge, visibilities.len())
            })
    }

    /// Set every pie or donut wedge's leader-line visibility.
    pub fn set_body_chart_pie_leader_line_visibilities(
        &mut self,
        drawable_object_id: u64,
        visibilities: &[LeaderLineVisibility],
    ) -> Result<()> {
        set_body_chart_pie_leader_line_visibilities(self, drawable_object_id, visibilities)
    }

    /// Set one pie or donut wedge's leader-line visibility.
    pub fn set_body_chart_pie_leader_line_visibility(
        &mut self,
        drawable_object_id: u64,
        wedge: ChartPieWedgeIndex,
        visibility: LeaderLineVisibility,
    ) -> Result<()> {
        let mut visibilities = body_chart_pie_leader_line_visibilities(self, drawable_object_id)?;
        let count = visibilities.len();
        let target = visibilities
            .get_mut(wedge.zero_based())
            .ok_or_else(|| leader_line_index_error("Pages", drawable_object_id, wedge, count))?;
        if *target == visibility {
            return Ok(());
        }
        *target = visibility;
        set_body_chart_pie_leader_line_visibilities(self, drawable_object_id, &visibilities)
    }
}

fn body_chart_pie_leader_line_visibilities(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<Vec<LeaderLineVisibility>> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    require_pie_leader_lines(graph.info.kind, drawable_object_id)?;
    let series_count = leader_line_series_count(
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    read_native_leader_line_visibilities(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        series_count,
    )
}

fn set_body_chart_pie_leader_line_visibilities(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    visibilities: &[LeaderLineVisibility],
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    require_pie_leader_lines(graph.info.kind, drawable_object_id)?;
    let series_count = leader_line_series_count(
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    if visibilities.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} requires {series_count} pie leader-line visibilities, got {}",
            visibilities.len()
        )));
    }
    if read_native_leader_line_visibilities(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        series_count,
    )? == visibilities
    {
        return Ok(());
    }
    let expected = visibilities.to_vec();
    let mut staged = editor.package().clone();
    set_native_leader_line_visibilities(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        &expected,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_pie_leader_line_visibilities(drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Pages chart pie leader-line visibility update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn require_pie_leader_lines(kind: Kind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_pie_start_angle() {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} kind {kind:?} has no pie leader lines"
        )));
    }
    Ok(())
}

fn leader_line_series_count(
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

fn leader_line_index_error(
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
