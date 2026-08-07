//! Native per-series value-label number-format CRUD for Pages charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_value_label_number_format::{
    chart_series_value_label_number_formats as read_native_formats,
    set_chart_series_value_label_number_formats as set_native_formats,
};
use crate::charts::{Index, NumberFormat};

impl PagesEditor {
    /// Read every series' value-label number format in native series order.
    pub fn body_chart_series_value_label_number_formats(
        &self,
        drawable_object_id: u64,
    ) -> Result<Vec<NumberFormat>> {
        body_chart_series_value_label_number_formats(self, drawable_object_id)
    }

    /// Read one series' value-label number format.
    pub fn body_chart_series_value_label_number_format(
        &self,
        drawable_object_id: u64,
        series: Index,
    ) -> Result<NumberFormat> {
        let formats = body_chart_series_value_label_number_formats(self, drawable_object_id)?;
        formats
            .get(series.zero_based())
            .copied()
            .ok_or_else(|| format_index_error("Pages", drawable_object_id, series, formats.len()))
    }

    /// Set every series' value-label number format in native series order.
    pub fn set_body_chart_series_value_label_number_formats(
        &mut self,
        drawable_object_id: u64,
        formats: &[NumberFormat],
    ) -> Result<()> {
        set_body_chart_series_value_label_number_formats(self, drawable_object_id, formats)
    }

    /// Set one series' value-label number format.
    pub fn set_body_chart_series_value_label_number_format(
        &mut self,
        drawable_object_id: u64,
        series: Index,
        format: NumberFormat,
    ) -> Result<()> {
        let mut formats = body_chart_series_value_label_number_formats(self, drawable_object_id)?;
        let count = formats.len();
        let target = formats
            .get_mut(series.zero_based())
            .ok_or_else(|| format_index_error("Pages", drawable_object_id, series, count))?;
        if *target == format {
            return Ok(());
        }
        *target = format;
        set_body_chart_series_value_label_number_formats(self, drawable_object_id, &formats)
    }
}

fn body_chart_series_value_label_number_formats(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<Vec<NumberFormat>> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    read_native_formats(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        graph.info.kind,
        series_count,
    )
}

fn set_body_chart_series_value_label_number_formats(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    formats: &[NumberFormat],
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    if formats.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} requires {series_count} series value-label number formats, got {}",
            formats.len()
        )));
    }
    if read_native_formats(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
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
        "Pages",
        graph.info.kind,
        &expected,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_series_value_label_number_formats(drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Pages chart series value-label number-format update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn format_index_error(suite: &str, drawable_object_id: u64, series: Index, count: usize) -> Error {
    Error::InvalidFormat(format!(
        "{suite} chart {drawable_object_id} series index {} exceeds series count {count}",
        series.zero_based()
    ))
}
