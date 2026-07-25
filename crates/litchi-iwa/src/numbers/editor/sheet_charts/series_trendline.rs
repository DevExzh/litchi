//! Native per-series trendline CRUD for Numbers charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_trendline::{
    chart_series_trendlines as read_native_trendlines,
    set_chart_series_trendlines as set_native_trendlines,
};
use crate::charts::{ChartSeriesIndex, ChartSeriesTrendline};

impl NumbersEditor {
    /// Read every series' trendline in native series order.
    pub fn sheet_chart_series_trendlines(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartSeriesTrendline>> {
        sheet_chart_series_trendlines(self, sheet_id, drawable_object_id)
    }

    /// Read one series' trendline.
    pub fn sheet_chart_series_trendline(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<ChartSeriesTrendline> {
        let trendlines = sheet_chart_series_trendlines(self, sheet_id, drawable_object_id)?;
        trendlines.get(series.zero_based()).copied().ok_or_else(|| {
            trendline_index_error("Numbers", drawable_object_id, series, trendlines.len())
        })
    }

    /// Set every series' trendline in native series order.
    pub fn set_sheet_chart_series_trendlines(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        trendlines: &[ChartSeriesTrendline],
    ) -> Result<()> {
        set_sheet_chart_series_trendlines(self, sheet_id, drawable_object_id, trendlines)
    }

    /// Set one series' trendline.
    pub fn set_sheet_chart_series_trendline(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        trendline: ChartSeriesTrendline,
    ) -> Result<()> {
        let mut trendlines = sheet_chart_series_trendlines(self, sheet_id, drawable_object_id)?;
        let count = trendlines.len();
        let target = trendlines
            .get_mut(series.zero_based())
            .ok_or_else(|| trendline_index_error("Numbers", drawable_object_id, series, count))?;
        if *target == trendline {
            return Ok(());
        }
        *target = trendline;
        set_sheet_chart_series_trendlines(self, sheet_id, drawable_object_id, &trendlines)
    }
}

fn sheet_chart_series_trendlines(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<Vec<ChartSeriesTrendline>> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Numbers",
        drawable_object_id,
    )?;
    read_native_trendlines(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        series_count,
    )
}

fn set_sheet_chart_series_trendlines(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    trendlines: &[ChartSeriesTrendline],
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Numbers",
        drawable_object_id,
    )?;
    if trendlines.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} requires {series_count} series trendlines, got {}",
            trendlines.len()
        )));
    }
    if read_native_trendlines(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        series_count,
    )? == trendlines
    {
        return Ok(());
    }
    let expected = trendlines.to_vec();
    let mut staged = editor.package().clone();
    set_native_trendlines(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        &expected,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_series_trendlines(sheet_id, drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Numbers chart series trendline update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn trendline_index_error(
    suite: &str,
    drawable_object_id: u64,
    series: ChartSeriesIndex,
    count: usize,
) -> Error {
    Error::InvalidFormat(format!(
        "{suite} chart {drawable_object_id} series index {} exceeds series count {count}",
        series.zero_based()
    ))
}
