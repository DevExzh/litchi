//! Native per-wedge label-distance CRUD for Numbers pie and donut charts.

use super::*;
use crate::charts::pie_label_distance::{
    chart_pie_label_distances as read_native_label_distances,
    set_chart_pie_label_distances as set_native_label_distances,
};
use crate::charts::{ChartPieLabelDistance, ChartPieWedgeIndex};

impl NumbersEditor {
    /// Read every pie or donut wedge's label distance in chart-series order.
    pub fn sheet_chart_pie_label_distances(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartPieLabelDistance>> {
        sheet_chart_pie_label_distances(self, sheet_id, drawable_object_id)
    }

    /// Read one pie or donut wedge's label distance.
    pub fn sheet_chart_pie_label_distance(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        wedge: ChartPieWedgeIndex,
    ) -> Result<ChartPieLabelDistance> {
        let distances = sheet_chart_pie_label_distances(self, sheet_id, drawable_object_id)?;
        distances.get(wedge.zero_based()).copied().ok_or_else(|| {
            label_distance_index_error("Numbers", drawable_object_id, wedge, distances.len())
        })
    }

    /// Set every pie or donut wedge's label distance in chart-series order.
    pub fn set_sheet_chart_pie_label_distances(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        distances: &[ChartPieLabelDistance],
    ) -> Result<()> {
        set_sheet_chart_pie_label_distances(self, sheet_id, drawable_object_id, distances)
    }

    /// Set one pie or donut wedge's label distance.
    pub fn set_sheet_chart_pie_label_distance(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        wedge: ChartPieWedgeIndex,
        distance: ChartPieLabelDistance,
    ) -> Result<()> {
        let mut distances = sheet_chart_pie_label_distances(self, sheet_id, drawable_object_id)?;
        let count = distances.len();
        let target = distances.get_mut(wedge.zero_based()).ok_or_else(|| {
            label_distance_index_error("Numbers", drawable_object_id, wedge, count)
        })?;
        if *target == distance {
            return Ok(());
        }
        *target = distance;
        set_sheet_chart_pie_label_distances(self, sheet_id, drawable_object_id, &distances)
    }
}

fn sheet_chart_pie_label_distances(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<Vec<ChartPieLabelDistance>> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    require_pie_label_distance(graph.info.kind, drawable_object_id)?;
    let series_count = label_distance_series_count(
        graph.info.direction,
        &graph.info.data,
        "Numbers",
        drawable_object_id,
    )?;
    read_native_label_distances(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        series_count,
    )
}

fn set_sheet_chart_pie_label_distances(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    distances: &[ChartPieLabelDistance],
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    require_pie_label_distance(graph.info.kind, drawable_object_id)?;
    let series_count = label_distance_series_count(
        graph.info.direction,
        &graph.info.data,
        "Numbers",
        drawable_object_id,
    )?;
    if distances.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} requires {series_count} pie label distances, got {}",
            distances.len()
        )));
    }
    if read_native_label_distances(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        series_count,
    )? == distances
    {
        return Ok(());
    }
    let expected = distances.to_vec();
    let mut staged = editor.package().clone();
    set_native_label_distances(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        &expected,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_pie_label_distances(sheet_id, drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Numbers chart pie label-distance update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn require_pie_label_distance(kind: Kind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_pie_start_angle() {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} kind {kind:?} has no pie labels"
        )));
    }
    Ok(())
}

fn label_distance_series_count(
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

fn label_distance_index_error(
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
