//! Native per-wedge data-label visibility CRUD for Pages pie and donut body charts.

use super::*;
use crate::charts::pie_labels::{
    chart_pie_label_visibilities as read_native_label_visibilities,
    set_chart_pie_label_visibilities as set_native_label_visibilities,
};
use crate::charts::{ChartPieWedgeIndex, LabelVisibility};

impl PagesEditor {
    /// Read label visibility for every pie or donut wedge in chart-series order.
    pub fn body_chart_pie_label_visibilities(
        &self,
        drawable_object_id: u64,
    ) -> Result<Vec<LabelVisibility>> {
        body_chart_pie_label_visibilities(self, drawable_object_id)
    }

    /// Read label visibility for one pie or donut wedge.
    pub fn body_chart_pie_label_visibility(
        &self,
        drawable_object_id: u64,
        wedge: ChartPieWedgeIndex,
    ) -> Result<LabelVisibility> {
        let values = body_chart_pie_label_visibilities(self, drawable_object_id)?;
        values
            .get(wedge.zero_based())
            .copied()
            .ok_or_else(|| label_index_error("Pages", drawable_object_id, wedge, values.len()))
    }

    /// Set label visibility for every pie or donut wedge in chart-series order.
    pub fn set_body_chart_pie_label_visibilities(
        &mut self,
        drawable_object_id: u64,
        visibilities: &[LabelVisibility],
    ) -> Result<()> {
        set_body_chart_pie_label_visibilities(self, drawable_object_id, visibilities)
    }

    /// Set label visibility for one pie or donut wedge.
    pub fn set_body_chart_pie_label_visibility(
        &mut self,
        drawable_object_id: u64,
        wedge: ChartPieWedgeIndex,
        visibility: LabelVisibility,
    ) -> Result<()> {
        let mut values = body_chart_pie_label_visibilities(self, drawable_object_id)?;
        let count = values.len();
        let target = values
            .get_mut(wedge.zero_based())
            .ok_or_else(|| label_index_error("Pages", drawable_object_id, wedge, count))?;
        if *target == visibility {
            return Ok(());
        }
        *target = visibility;
        set_body_chart_pie_label_visibilities(self, drawable_object_id, &values)
    }
}

fn body_chart_pie_label_visibilities(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<Vec<LabelVisibility>> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    require_pie_labels(graph.info.kind, drawable_object_id)?;
    let series_count = label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    read_native_label_visibilities(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        series_count,
    )
}

fn set_body_chart_pie_label_visibilities(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    visibilities: &[LabelVisibility],
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    require_pie_labels(graph.info.kind, drawable_object_id)?;
    let series_count = label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    if visibilities.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} requires {series_count} pie label settings, got {}",
            visibilities.len()
        )));
    }
    if read_native_label_visibilities(
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
    set_native_label_visibilities(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        &expected,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_pie_label_visibilities(drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Pages chart pie label-visibility update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn require_pie_labels(kind: Kind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_pie_start_angle() {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} kind {kind:?} has no pie labels"
        )));
    }
    Ok(())
}

fn label_series_count(
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
