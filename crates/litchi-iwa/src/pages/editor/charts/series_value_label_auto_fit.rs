//! Native per-series value-label Auto-Fit CRUD for Pages charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_value_label_auto_fit::{
    chart_series_value_label_auto_fits as read_native_auto_fits,
    set_chart_series_value_label_auto_fits as set_native_auto_fits,
};
use crate::charts::{ChartSeriesValueLabelAutoFit, Index};

impl PagesEditor {
    /// Read every series' value-label Auto-Fit setting in native series order.
    pub fn body_chart_series_value_label_auto_fits(
        &self,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartSeriesValueLabelAutoFit>> {
        body_chart_series_value_label_auto_fits(self, drawable_object_id)
    }

    /// Read one series' value-label Auto-Fit setting.
    pub fn body_chart_series_value_label_auto_fit(
        &self,
        drawable_object_id: u64,
        series: Index,
    ) -> Result<ChartSeriesValueLabelAutoFit> {
        let values = body_chart_series_value_label_auto_fits(self, drawable_object_id)?;
        values
            .get(series.zero_based())
            .copied()
            .ok_or_else(|| auto_fit_index_error("Pages", drawable_object_id, series, values.len()))
    }

    /// Set every series' value-label Auto-Fit setting in native series order.
    pub fn set_body_chart_series_value_label_auto_fits(
        &mut self,
        drawable_object_id: u64,
        values: &[ChartSeriesValueLabelAutoFit],
    ) -> Result<()> {
        set_body_chart_series_value_label_auto_fits(self, drawable_object_id, values)
    }

    /// Set one series' value-label Auto-Fit setting.
    pub fn set_body_chart_series_value_label_auto_fit(
        &mut self,
        drawable_object_id: u64,
        series: Index,
        value: ChartSeriesValueLabelAutoFit,
    ) -> Result<()> {
        let mut values = body_chart_series_value_label_auto_fits(self, drawable_object_id)?;
        let count = values.len();
        let target = values
            .get_mut(series.zero_based())
            .ok_or_else(|| auto_fit_index_error("Pages", drawable_object_id, series, count))?;
        if *target == value {
            return Ok(());
        }
        *target = value;
        set_body_chart_series_value_label_auto_fits(self, drawable_object_id, &values)
    }
}

fn body_chart_series_value_label_auto_fits(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<Vec<ChartSeriesValueLabelAutoFit>> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    read_native_auto_fits(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        series_count,
    )
}

fn set_body_chart_series_value_label_auto_fits(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    values: &[ChartSeriesValueLabelAutoFit],
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
            "Pages chart {drawable_object_id} requires {series_count} series value-label Auto-Fit settings, got {}",
            values.len()
        )));
    }
    if read_native_auto_fits(
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
    set_native_auto_fits(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        &expected,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_series_value_label_auto_fits(drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Pages chart series value-label Auto-Fit update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn auto_fit_index_error(
    suite: &str,
    drawable_object_id: u64,
    series: Index,
    count: usize,
) -> Error {
    Error::InvalidFormat(format!(
        "{suite} chart {drawable_object_id} series index {} exceeds series count {count}",
        series.zero_based()
    ))
}
