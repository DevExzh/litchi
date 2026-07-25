//! Native per-series error-bar CRUD for Numbers charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_error_bars::{
    chart_series_error_bars as read_native_error_bars,
    set_chart_series_error_bars as set_native_error_bars,
};
use crate::charts::{ChartSeriesErrorBars, ChartSeriesIndex};

impl NumbersEditor {
    /// Read every series' error bars in native series order.
    pub fn sheet_chart_series_error_bars(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartSeriesErrorBars>> {
        sheet_chart_series_error_bars(self, sheet_id, drawable_object_id)
    }

    /// Read one series' error bars.
    pub fn sheet_chart_series_error_bar(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<ChartSeriesErrorBars> {
        let values = sheet_chart_series_error_bars(self, sheet_id, drawable_object_id)?;
        values.get(series.zero_based()).cloned().ok_or_else(|| {
            error_bar_index_error("Numbers", drawable_object_id, series, values.len())
        })
    }

    /// Set every series' error bars in native series order.
    pub fn set_sheet_chart_series_error_bars(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        values: &[ChartSeriesErrorBars],
    ) -> Result<()> {
        set_sheet_chart_series_error_bars(self, sheet_id, drawable_object_id, values)
    }

    /// Set one series' error bars.
    pub fn set_sheet_chart_series_error_bar(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        value: ChartSeriesErrorBars,
    ) -> Result<()> {
        let mut values = sheet_chart_series_error_bars(self, sheet_id, drawable_object_id)?;
        let count = values.len();
        let target = values
            .get_mut(series.zero_based())
            .ok_or_else(|| error_bar_index_error("Numbers", drawable_object_id, series, count))?;
        if *target == value {
            return Ok(());
        }
        *target = value;
        set_sheet_chart_series_error_bars(self, sheet_id, drawable_object_id, &values)
    }
}

fn sheet_chart_series_error_bars(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<Vec<ChartSeriesErrorBars>> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Numbers",
        drawable_object_id,
    )?;
    read_native_error_bars(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        series_count,
    )
}

fn set_sheet_chart_series_error_bars(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    values: &[ChartSeriesErrorBars],
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Numbers",
        drawable_object_id,
    )?;
    if values.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} requires {series_count} series error-bar settings, got {}",
            values.len()
        )));
    }
    if read_native_error_bars(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
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
        "Numbers",
        &expected,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_series_error_bars(sheet_id, drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Numbers chart series error-bar update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn error_bar_index_error(
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
