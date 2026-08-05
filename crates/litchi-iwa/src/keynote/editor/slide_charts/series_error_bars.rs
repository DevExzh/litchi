//! Native per-series error-bar CRUD for Keynote charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_error_bars::{
    chart_series_error_bars as read_native_error_bars,
    set_chart_series_error_bars as set_native_error_bars,
};
use crate::charts::{ChartSeriesErrorBars, Index};

impl KeynoteEditor {
    /// Read every series' error bars in native series order.
    pub fn slide_chart_series_error_bars(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartSeriesErrorBars>> {
        slide_chart_series_error_bars(self, slide_index, drawable_object_id)
    }

    /// Read one series' error bars.
    pub fn slide_chart_series_error_bar(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        series: Index,
    ) -> Result<ChartSeriesErrorBars> {
        let values = slide_chart_series_error_bars(self, slide_index, drawable_object_id)?;
        values.get(series.zero_based()).cloned().ok_or_else(|| {
            error_bar_index_error("Keynote", drawable_object_id, series, values.len())
        })
    }

    /// Set every series' error bars in native series order.
    pub fn set_slide_chart_series_error_bars(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        values: &[ChartSeriesErrorBars],
    ) -> Result<()> {
        set_slide_chart_series_error_bars(self, slide_index, drawable_object_id, values)
    }

    /// Set one series' error bars.
    pub fn set_slide_chart_series_error_bar(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        series: Index,
        value: ChartSeriesErrorBars,
    ) -> Result<()> {
        let mut values = slide_chart_series_error_bars(self, slide_index, drawable_object_id)?;
        let count = values.len();
        let target = values
            .get_mut(series.zero_based())
            .ok_or_else(|| error_bar_index_error("Keynote", drawable_object_id, series, count))?;
        if *target == value {
            return Ok(());
        }
        *target = value;
        set_slide_chart_series_error_bars(self, slide_index, drawable_object_id, &values)
    }
}

fn slide_chart_series_error_bars(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<Vec<ChartSeriesErrorBars>> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )?;
    read_native_error_bars(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        series_count,
    )
}

fn set_slide_chart_series_error_bars(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    values: &[ChartSeriesErrorBars],
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )?;
    if values.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Keynote chart {drawable_object_id} requires {series_count} series error-bar settings, got {}",
            values.len()
        )));
    }
    if read_native_error_bars(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
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
        "Keynote",
        &expected,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_series_error_bars(slide_index, drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Keynote chart series error-bar update failed validation".to_owned(),
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
