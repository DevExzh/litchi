//! Native per-series trendline CRUD for Keynote charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_trendline::{
    chart_series_trendlines as read_native_trendlines,
    set_chart_series_trendlines as set_native_trendlines,
};
use crate::charts::{ChartSeriesIndex, ChartSeriesTrendline};

impl KeynoteEditor {
    /// Read every series' trendline in native series order.
    pub fn slide_chart_series_trendlines(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartSeriesTrendline>> {
        slide_chart_series_trendlines(self, slide_index, drawable_object_id)
    }

    /// Read one series' trendline.
    pub fn slide_chart_series_trendline(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<ChartSeriesTrendline> {
        let trendlines = slide_chart_series_trendlines(self, slide_index, drawable_object_id)?;
        trendlines.get(series.zero_based()).copied().ok_or_else(|| {
            trendline_index_error("Keynote", drawable_object_id, series, trendlines.len())
        })
    }

    /// Set every series' trendline in native series order.
    pub fn set_slide_chart_series_trendlines(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        trendlines: &[ChartSeriesTrendline],
    ) -> Result<()> {
        set_slide_chart_series_trendlines(self, slide_index, drawable_object_id, trendlines)
    }

    /// Set one series' trendline.
    pub fn set_slide_chart_series_trendline(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        trendline: ChartSeriesTrendline,
    ) -> Result<()> {
        let mut trendlines = slide_chart_series_trendlines(self, slide_index, drawable_object_id)?;
        let count = trendlines.len();
        let target = trendlines
            .get_mut(series.zero_based())
            .ok_or_else(|| trendline_index_error("Keynote", drawable_object_id, series, count))?;
        if *target == trendline {
            return Ok(());
        }
        *target = trendline;
        set_slide_chart_series_trendlines(self, slide_index, drawable_object_id, &trendlines)
    }
}

fn slide_chart_series_trendlines(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<Vec<ChartSeriesTrendline>> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )?;
    read_native_trendlines(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        series_count,
    )
}

fn set_slide_chart_series_trendlines(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    trendlines: &[ChartSeriesTrendline],
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )?;
    if trendlines.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Keynote chart {drawable_object_id} requires {series_count} series trendlines, got {}",
            trendlines.len()
        )));
    }
    if read_native_trendlines(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
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
        "Keynote",
        &expected,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_series_trendlines(slide_index, drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Keynote chart series trendline update failed validation".to_owned(),
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
