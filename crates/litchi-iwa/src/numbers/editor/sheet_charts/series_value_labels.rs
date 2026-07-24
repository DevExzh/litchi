//! Native per-series value-label visibility CRUD for Numbers charts.

use super::*;
use crate::charts::series_value_labels::{
    chart_series_value_label_visibilities as read_native_value_labels,
    set_chart_series_value_label_visibilities as set_native_value_labels,
};
use crate::charts::{ChartSeriesIndex, ChartSeriesValueLabelVisibility};

impl NumbersEditor {
    /// Read every series' value-label visibility in native series order.
    pub fn sheet_chart_series_value_label_visibilities(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartSeriesValueLabelVisibility>> {
        sheet_chart_series_value_label_visibilities(self, sheet_id, drawable_object_id)
    }

    /// Read one series' value-label visibility.
    pub fn sheet_chart_series_value_label_visibility(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<ChartSeriesValueLabelVisibility> {
        let visibilities =
            sheet_chart_series_value_label_visibilities(self, sheet_id, drawable_object_id)?;
        visibilities
            .get(series.zero_based())
            .copied()
            .ok_or_else(|| {
                value_label_index_error("Numbers", drawable_object_id, series, visibilities.len())
            })
    }

    /// Set every series' value-label visibility in native series order.
    pub fn set_sheet_chart_series_value_label_visibilities(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        visibilities: &[ChartSeriesValueLabelVisibility],
    ) -> Result<()> {
        set_sheet_chart_series_value_label_visibilities(
            self,
            sheet_id,
            drawable_object_id,
            visibilities,
        )
    }

    /// Set one series' value-label visibility.
    pub fn set_sheet_chart_series_value_label_visibility(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        visibility: ChartSeriesValueLabelVisibility,
    ) -> Result<()> {
        let mut visibilities =
            sheet_chart_series_value_label_visibilities(self, sheet_id, drawable_object_id)?;
        let count = visibilities.len();
        let target = visibilities
            .get_mut(series.zero_based())
            .ok_or_else(|| value_label_index_error("Numbers", drawable_object_id, series, count))?;
        if *target == visibility {
            return Ok(());
        }
        *target = visibility;
        set_sheet_chart_series_value_label_visibilities(
            self,
            sheet_id,
            drawable_object_id,
            &visibilities,
        )
    }
}

fn sheet_chart_series_value_label_visibilities(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<Vec<ChartSeriesValueLabelVisibility>> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Numbers",
        drawable_object_id,
    )?;
    read_native_value_labels(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        graph.info.kind,
        series_count,
    )
}

fn set_sheet_chart_series_value_label_visibilities(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    visibilities: &[ChartSeriesValueLabelVisibility],
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Numbers",
        drawable_object_id,
    )?;
    if visibilities.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} requires {series_count} series value-label visibilities, got {}",
            visibilities.len()
        )));
    }
    if read_native_value_labels(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        graph.info.kind,
        series_count,
    )? == visibilities
    {
        return Ok(());
    }
    let expected = visibilities.to_vec();
    let mut staged = editor.package().clone();
    set_native_value_labels(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        graph.info.kind,
        &expected,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_series_value_label_visibilities(sheet_id, drawable_object_id)?
        != expected
    {
        return Err(Error::InvalidFormat(
            "Numbers chart series value-label update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn value_label_series_count(
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

fn value_label_index_error(
    drawable_label: &str,
    drawable_object_id: u64,
    series: ChartSeriesIndex,
    series_count: usize,
) -> Error {
    Error::InvalidFormat(format!(
        "{drawable_label} chart {drawable_object_id} series index {} exceeds series count {series_count}",
        series.zero_based()
    ))
}
