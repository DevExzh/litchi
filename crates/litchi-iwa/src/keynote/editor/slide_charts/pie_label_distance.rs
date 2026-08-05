//! Native per-wedge label-distance CRUD for Keynote pie and donut charts.

use super::*;
use crate::charts::pie_label_distance::{
    chart_pie_label_distances as read_native_label_distances,
    set_chart_pie_label_distances as set_native_label_distances,
};
use crate::charts::{ChartPieLabelDistance, ChartPieWedgeIndex};

impl KeynoteEditor {
    /// Read every pie or donut wedge's label distance in chart-series order.
    pub fn slide_chart_pie_label_distances(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartPieLabelDistance>> {
        slide_chart_pie_label_distances(self, slide_index, drawable_object_id)
    }

    /// Read one pie or donut wedge's label distance.
    pub fn slide_chart_pie_label_distance(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        wedge: ChartPieWedgeIndex,
    ) -> Result<ChartPieLabelDistance> {
        let distances = slide_chart_pie_label_distances(self, slide_index, drawable_object_id)?;
        distances.get(wedge.zero_based()).copied().ok_or_else(|| {
            label_distance_index_error("Keynote", drawable_object_id, wedge, distances.len())
        })
    }

    /// Set every pie or donut wedge's label distance in chart-series order.
    pub fn set_slide_chart_pie_label_distances(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        distances: &[ChartPieLabelDistance],
    ) -> Result<()> {
        set_slide_chart_pie_label_distances(self, slide_index, drawable_object_id, distances)
    }

    /// Set one pie or donut wedge's label distance.
    pub fn set_slide_chart_pie_label_distance(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        wedge: ChartPieWedgeIndex,
        distance: ChartPieLabelDistance,
    ) -> Result<()> {
        let mut distances = slide_chart_pie_label_distances(self, slide_index, drawable_object_id)?;
        let count = distances.len();
        let target = distances.get_mut(wedge.zero_based()).ok_or_else(|| {
            label_distance_index_error("Keynote", drawable_object_id, wedge, count)
        })?;
        if *target == distance {
            return Ok(());
        }
        *target = distance;
        set_slide_chart_pie_label_distances(self, slide_index, drawable_object_id, &distances)
    }
}

fn slide_chart_pie_label_distances(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<Vec<ChartPieLabelDistance>> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    require_pie_label_distance(graph.info.kind, drawable_object_id)?;
    let series_count = label_distance_series_count(
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )?;
    read_native_label_distances(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        series_count,
    )
}

fn set_slide_chart_pie_label_distances(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    distances: &[ChartPieLabelDistance],
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    require_pie_label_distance(graph.info.kind, drawable_object_id)?;
    let series_count = label_distance_series_count(
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )?;
    if distances.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Keynote chart {drawable_object_id} requires {series_count} pie label distances, got {}",
            distances.len()
        )));
    }
    if read_native_label_distances(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
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
        "Keynote",
        &expected,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_pie_label_distances(slide_index, drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Keynote chart pie label-distance update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn require_pie_label_distance(kind: ChartKind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_pie_start_angle() {
        return Err(Error::InvalidFormat(format!(
            "Keynote chart {drawable_object_id} kind {kind:?} has no pie labels"
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
