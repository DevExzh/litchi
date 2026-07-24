//! Native per-wedge data-label visibility CRUD for Keynote pie and donut charts.

use super::*;
use crate::charts::pie_labels::{
    chart_pie_label_visibilities as read_native_label_visibilities,
    set_chart_pie_label_visibilities as set_native_label_visibilities,
};
use crate::charts::{ChartPieLabelVisibility, ChartPieWedgeIndex};

impl KeynoteEditor {
    /// Read label visibility for every pie or donut wedge in chart-series order.
    pub fn slide_chart_pie_label_visibilities(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartPieLabelVisibility>> {
        slide_chart_pie_label_visibilities(self, slide_index, drawable_object_id)
    }

    /// Read label visibility for one pie or donut wedge.
    pub fn slide_chart_pie_label_visibility(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        wedge: ChartPieWedgeIndex,
    ) -> Result<ChartPieLabelVisibility> {
        let visibilities =
            slide_chart_pie_label_visibilities(self, slide_index, drawable_object_id)?;
        visibilities
            .get(wedge.zero_based())
            .copied()
            .ok_or_else(|| {
                label_index_error("Keynote", drawable_object_id, wedge, visibilities.len())
            })
    }

    /// Set label visibility for every pie or donut wedge in chart-series order.
    pub fn set_slide_chart_pie_label_visibilities(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        visibilities: &[ChartPieLabelVisibility],
    ) -> Result<()> {
        set_slide_chart_pie_label_visibilities(self, slide_index, drawable_object_id, visibilities)
    }

    /// Set label visibility for one pie or donut wedge.
    pub fn set_slide_chart_pie_label_visibility(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        wedge: ChartPieWedgeIndex,
        visibility: ChartPieLabelVisibility,
    ) -> Result<()> {
        let mut visibilities =
            slide_chart_pie_label_visibilities(self, slide_index, drawable_object_id)?;
        let count = visibilities.len();
        let target = visibilities
            .get_mut(wedge.zero_based())
            .ok_or_else(|| label_index_error("Keynote", drawable_object_id, wedge, count))?;
        if *target == visibility {
            return Ok(());
        }
        *target = visibility;
        set_slide_chart_pie_label_visibilities(self, slide_index, drawable_object_id, &visibilities)
    }
}

fn slide_chart_pie_label_visibilities(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<Vec<ChartPieLabelVisibility>> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    require_pie_labels(graph.info.kind, drawable_object_id)?;
    let series_count = label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )?;
    read_native_label_visibilities(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        series_count,
    )
}

fn set_slide_chart_pie_label_visibilities(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    visibilities: &[ChartPieLabelVisibility],
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    require_pie_labels(graph.info.kind, drawable_object_id)?;
    let series_count = label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )?;
    if visibilities.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Keynote chart {drawable_object_id} requires {series_count} pie label settings, got {}",
            visibilities.len()
        )));
    }
    if read_native_label_visibilities(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        series_count,
    )? == visibilities
    {
        return Ok(());
    }
    let expected = visibilities.to_vec();
    let mut staged = editor.package().clone();
    set_native_label_visibilities(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        &expected,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_pie_label_visibilities(slide_index, drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Keynote chart pie label-visibility update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn require_pie_labels(kind: ChartKind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_pie_start_angle() {
        return Err(Error::InvalidFormat(format!(
            "Keynote chart {drawable_object_id} kind {kind:?} has no pie labels"
        )));
    }
    Ok(())
}

fn label_series_count(
    direction: ChartSeriesDirection,
    data: &ChartData,
    drawable_label: &str,
    drawable_object_id: u64,
) -> Result<usize> {
    match direction {
        ChartSeriesDirection::Rows => Ok(data.row_names().len()),
        ChartSeriesDirection::Columns => Ok(data.column_names().len()),
        ChartSeriesDirection::Unsupported(value) => Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has unsupported series direction {value}"
        ))),
    }
}

fn label_index_error(
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
