//! Native per-series value-label number-format CRUD for Numbers charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_value_label_number_format::{
    chart_series_value_label_number_formats as read_native_formats,
    set_chart_series_value_label_number_formats as set_native_formats,
};
use crate::charts::{ChartSeriesIndex, ChartSeriesValueLabelNumberFormat};

impl NumbersEditor {
    /// Read every series' value-label number format in native series order.
    pub fn sheet_chart_series_value_label_number_formats(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartSeriesValueLabelNumberFormat>> {
        sheet_chart_series_value_label_number_formats(self, sheet_id, drawable_object_id)
    }

    /// Read one series' value-label number format.
    pub fn sheet_chart_series_value_label_number_format(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<ChartSeriesValueLabelNumberFormat> {
        let formats =
            sheet_chart_series_value_label_number_formats(self, sheet_id, drawable_object_id)?;
        formats
            .get(series.zero_based())
            .copied()
            .ok_or_else(|| format_index_error("Numbers", drawable_object_id, series, formats.len()))
    }

    /// Set every series' value-label number format in native series order.
    pub fn set_sheet_chart_series_value_label_number_formats(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        formats: &[ChartSeriesValueLabelNumberFormat],
    ) -> Result<()> {
        set_sheet_chart_series_value_label_number_formats(
            self,
            sheet_id,
            drawable_object_id,
            formats,
        )
    }

    /// Set one series' value-label number format.
    pub fn set_sheet_chart_series_value_label_number_format(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        format: ChartSeriesValueLabelNumberFormat,
    ) -> Result<()> {
        let mut formats =
            sheet_chart_series_value_label_number_formats(self, sheet_id, drawable_object_id)?;
        let count = formats.len();
        let target = formats
            .get_mut(series.zero_based())
            .ok_or_else(|| format_index_error("Numbers", drawable_object_id, series, count))?;
        if *target == format {
            return Ok(());
        }
        *target = format;
        set_sheet_chart_series_value_label_number_formats(
            self,
            sheet_id,
            drawable_object_id,
            &formats,
        )
    }
}

fn sheet_chart_series_value_label_number_formats(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<Vec<ChartSeriesValueLabelNumberFormat>> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Numbers",
        drawable_object_id,
    )?;
    read_native_formats(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        graph.info.kind,
        series_count,
    )
}

fn set_sheet_chart_series_value_label_number_formats(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    formats: &[ChartSeriesValueLabelNumberFormat],
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Numbers",
        drawable_object_id,
    )?;
    if formats.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} requires {series_count} series value-label number formats, got {}",
            formats.len()
        )));
    }
    if read_native_formats(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        graph.info.kind,
        series_count,
    )? == formats
    {
        return Ok(());
    }
    let expected = formats.to_vec();
    let mut staged = editor.package().clone();
    set_native_formats(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        graph.info.kind,
        &expected,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_series_value_label_number_formats(sheet_id, drawable_object_id)?
        != expected
    {
        return Err(Error::InvalidFormat(
            "Numbers chart series value-label number-format update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn format_index_error(
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
