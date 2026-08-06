//! Native per-series error-bar CRUD for Pages charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::Index;
use crate::charts::series_error_bars::{
    chart_series_error_bars as read_native_error_bars,
    set_chart_series_error_bars as set_native_error_bars,
};
use litchi_iwa_common::chart::error_bar::Series;

impl PagesEditor {
    /// Read every body chart series' error bars in native series order.
    pub fn body_chart_series_error_bars(&self, drawable_object_id: u64) -> Result<Vec<Series>> {
        body_chart_series_error_bars(self, drawable_object_id)
    }

    /// Read one body chart series' error bars.
    pub fn body_chart_series_error_bar(
        &self,
        drawable_object_id: u64,
        series: Index,
    ) -> Result<Series> {
        let values = body_chart_series_error_bars(self, drawable_object_id)?;
        values
            .get(series.zero_based())
            .cloned()
            .ok_or_else(|| error_bar_index_error("Pages", drawable_object_id, series, values.len()))
    }

    /// Set every body chart series' error bars in native series order.
    pub fn set_body_chart_series_error_bars(
        &mut self,
        drawable_object_id: u64,
        values: &[Series],
    ) -> Result<()> {
        set_body_chart_series_error_bars(self, drawable_object_id, values)
    }

    /// Set one body chart series' error bars.
    pub fn set_body_chart_series_error_bar(
        &mut self,
        drawable_object_id: u64,
        series: Index,
        value: Series,
    ) -> Result<()> {
        let mut values = body_chart_series_error_bars(self, drawable_object_id)?;
        let count = values.len();
        let target = values
            .get_mut(series.zero_based())
            .ok_or_else(|| error_bar_index_error("Pages", drawable_object_id, series, count))?;
        if *target == value {
            return Ok(());
        }
        *target = value;
        set_body_chart_series_error_bars(self, drawable_object_id, &values)
    }
}

fn body_chart_series_error_bars(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<Vec<Series>> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    read_native_error_bars(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        series_count,
    )
}

fn set_body_chart_series_error_bars(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    values: &[Series],
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    if values.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} requires {series_count} series error-bar settings, got {}",
            values.len()
        )));
    }
    if read_native_error_bars(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        series_count,
    )? == values
    {
        return Ok(());
    }
    let expected = values.to_vec();
    let mut staged = editor.package().clone();
    set_native_error_bars(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        &expected,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_series_error_bars(drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Pages chart series error-bar update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn error_bar_index_error(
    drawable_label: &str,
    drawable_object_id: u64,
    series: Index,
    series_count: usize,
) -> Error {
    Error::InvalidFormat(format!(
        "{drawable_label} chart {drawable_object_id} series index {} exceeds series count {series_count}",
        series.zero_based()
    ))
}
