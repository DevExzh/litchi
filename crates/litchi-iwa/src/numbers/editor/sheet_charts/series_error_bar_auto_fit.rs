//! Native per-series error-bar Auto-Fit CRUD for Numbers charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_error_bar_auto_fit::{
    chart_series_error_bar_auto_fits as read_native_auto_fits,
    set_chart_series_error_bar_auto_fits as set_native_auto_fits,
};
use crate::charts::{ChartSeriesErrorBarAutoFit, Index};

impl NumbersEditor {
    /// Read every series' error-bar Auto-Fit setting.
    pub fn sheet_chart_series_error_bar_auto_fits(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartSeriesErrorBarAutoFit>> {
        sheet_chart_series_error_bar_auto_fits(self, sheet_id, drawable_object_id)
    }

    /// Read one series' error-bar Auto-Fit setting.
    pub fn sheet_chart_series_error_bar_auto_fit(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: Index,
    ) -> Result<ChartSeriesErrorBarAutoFit> {
        let values = sheet_chart_series_error_bar_auto_fits(self, sheet_id, drawable_object_id)?;
        values.get(series.zero_based()).copied().ok_or_else(|| {
            auto_fit_index_error("Numbers", drawable_object_id, series, values.len())
        })
    }

    /// Set every series' error-bar Auto-Fit setting.
    pub fn set_sheet_chart_series_error_bar_auto_fits(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        values: &[ChartSeriesErrorBarAutoFit],
    ) -> Result<()> {
        set_sheet_chart_series_error_bar_auto_fits(self, sheet_id, drawable_object_id, values)
    }

    /// Set one series' error-bar Auto-Fit setting.
    pub fn set_sheet_chart_series_error_bar_auto_fit(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: Index,
        value: ChartSeriesErrorBarAutoFit,
    ) -> Result<()> {
        let mut values =
            sheet_chart_series_error_bar_auto_fits(self, sheet_id, drawable_object_id)?;
        let count = values.len();
        let target = values
            .get_mut(series.zero_based())
            .ok_or_else(|| auto_fit_index_error("Numbers", drawable_object_id, series, count))?;
        if *target == value {
            return Ok(());
        }
        *target = value;
        set_sheet_chart_series_error_bar_auto_fits(self, sheet_id, drawable_object_id, &values)
    }
}

fn sheet_chart_series_error_bar_auto_fits(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<Vec<ChartSeriesErrorBarAutoFit>> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Numbers",
        drawable_object_id,
    )?;
    read_native_auto_fits(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        series_count,
    )
}

fn set_sheet_chart_series_error_bar_auto_fits(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    values: &[ChartSeriesErrorBarAutoFit],
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Numbers",
        drawable_object_id,
    )?;
    if values.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} requires {series_count} series error-bar Auto-Fit settings, got {}",
            values.len()
        )));
    }
    if read_native_auto_fits(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
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
        "Numbers",
        values,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_series_error_bar_auto_fits(sheet_id, drawable_object_id)? != values {
        return Err(Error::InvalidFormat(
            "Numbers chart series error-bar Auto-Fit update failed validation".to_owned(),
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
