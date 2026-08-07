//! Native per-series error-bar Auto-Fit CRUD for Keynote charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_error_bar_auto_fit::{
    chart_series_error_bar_auto_fits as read_native_auto_fits,
    set_chart_series_error_bar_auto_fits as set_native_auto_fits,
};
use crate::charts::{ChartSeriesErrorBarAutoFit, Index};

impl KeynoteEditor {
    /// Read every series' error-bar Auto-Fit setting.
    pub fn slide_chart_series_error_bar_auto_fits(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartSeriesErrorBarAutoFit>> {
        slide_chart_series_error_bar_auto_fits(self, slide_index, drawable_object_id)
    }

    /// Read one series' error-bar Auto-Fit setting.
    pub fn slide_chart_series_error_bar_auto_fit(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        series: Index,
    ) -> Result<ChartSeriesErrorBarAutoFit> {
        let values = slide_chart_series_error_bar_auto_fits(self, slide_index, drawable_object_id)?;
        values.get(series.zero_based()).copied().ok_or_else(|| {
            auto_fit_index_error("Keynote", drawable_object_id, series, values.len())
        })
    }

    /// Set every series' error-bar Auto-Fit setting.
    pub fn set_slide_chart_series_error_bar_auto_fits(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        values: &[ChartSeriesErrorBarAutoFit],
    ) -> Result<()> {
        set_slide_chart_series_error_bar_auto_fits(self, slide_index, drawable_object_id, values)
    }

    /// Set one series' error-bar Auto-Fit setting.
    pub fn set_slide_chart_series_error_bar_auto_fit(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        series: Index,
        value: ChartSeriesErrorBarAutoFit,
    ) -> Result<()> {
        let mut values =
            slide_chart_series_error_bar_auto_fits(self, slide_index, drawable_object_id)?;
        let count = values.len();
        let target = values
            .get_mut(series.zero_based())
            .ok_or_else(|| auto_fit_index_error("Keynote", drawable_object_id, series, count))?;
        if *target == value {
            return Ok(());
        }
        *target = value;
        set_slide_chart_series_error_bar_auto_fits(self, slide_index, drawable_object_id, &values)
    }
}

fn slide_chart_series_error_bar_auto_fits(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<Vec<ChartSeriesErrorBarAutoFit>> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )?;
    read_native_auto_fits(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        series_count,
    )
}

fn set_slide_chart_series_error_bar_auto_fits(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    values: &[ChartSeriesErrorBarAutoFit],
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
            "Keynote chart {drawable_object_id} requires {series_count} series error-bar Auto-Fit settings, got {}",
            values.len()
        )));
    }
    if read_native_auto_fits(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        series_count,
    )? == values
    {
        return Ok(());
    }
    let mut staged = editor.package().clone();
    set_native_auto_fits(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        values,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_series_error_bar_auto_fits(slide_index, drawable_object_id)? != values {
        return Err(Error::InvalidFormat(
            "Keynote chart series error-bar Auto-Fit update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn auto_fit_index_error(
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
