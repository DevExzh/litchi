//! Native per-series trendline CRUD for Pages charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_trendline::{
    chart_series_trendlines as read_native_trendlines,
    set_chart_series_trendlines as set_native_trendlines,
};
use crate::charts::{ChartSeriesIndex, ChartSeriesTrendline};

impl PagesEditor {
    /// Read every body chart series' trendline in native series order.
    pub fn body_chart_series_trendlines(
        &self,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartSeriesTrendline>> {
        body_chart_series_trendlines(self, drawable_object_id)
    }

    /// Read one body chart series' trendline.
    pub fn body_chart_series_trendline(
        &self,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<ChartSeriesTrendline> {
        let trendlines = body_chart_series_trendlines(self, drawable_object_id)?;
        trendlines.get(series.zero_based()).copied().ok_or_else(|| {
            trendline_index_error("Pages", drawable_object_id, series, trendlines.len())
        })
    }

    /// Set every body chart series' trendline in native series order.
    pub fn set_body_chart_series_trendlines(
        &mut self,
        drawable_object_id: u64,
        trendlines: &[ChartSeriesTrendline],
    ) -> Result<()> {
        set_body_chart_series_trendlines(self, drawable_object_id, trendlines)
    }

    /// Set one body chart series' trendline.
    pub fn set_body_chart_series_trendline(
        &mut self,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        trendline: ChartSeriesTrendline,
    ) -> Result<()> {
        let mut trendlines = body_chart_series_trendlines(self, drawable_object_id)?;
        let count = trendlines.len();
        let target = trendlines
            .get_mut(series.zero_based())
            .ok_or_else(|| trendline_index_error("Pages", drawable_object_id, series, count))?;
        if *target == trendline {
            return Ok(());
        }
        *target = trendline;
        set_body_chart_series_trendlines(self, drawable_object_id, &trendlines)
    }
}

fn body_chart_series_trendlines(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<Vec<ChartSeriesTrendline>> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    read_native_trendlines(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        series_count,
    )
}

fn set_body_chart_series_trendlines(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    trendlines: &[ChartSeriesTrendline],
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    if trendlines.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} requires {series_count} series trendlines, got {}",
            trendlines.len()
        )));
    }
    if read_native_trendlines(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
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
        "Pages",
        &expected,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_series_trendlines(drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Pages chart series trendline update failed validation".to_owned(),
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
