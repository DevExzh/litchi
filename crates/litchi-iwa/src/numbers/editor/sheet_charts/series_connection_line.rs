//! Per-series connection-line geometry CRUD for Numbers charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_connection_line::{
    chart_series_connection_lines as read_native, set_chart_series_connection_lines as set_native,
};
use crate::charts::{ChartSeriesConnectionLine, ChartSeriesIndex};

impl NumbersEditor {
    /// Read connection geometry in native series order.
    pub fn sheet_chart_series_connection_lines(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartSeriesConnectionLine>> {
        read(self, sheet_id, drawable_object_id)
    }

    /// Read one series' connection geometry.
    pub fn sheet_chart_series_connection_line(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<ChartSeriesConnectionLine> {
        let values = read(self, sheet_id, drawable_object_id)?;
        values
            .get(series.zero_based())
            .copied()
            .ok_or_else(|| index_error(drawable_object_id, series, values.len()))
    }

    /// Replace every series' connection geometry transactionally.
    pub fn set_sheet_chart_series_connection_lines(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        values: &[ChartSeriesConnectionLine],
    ) -> Result<()> {
        set(self, sheet_id, drawable_object_id, values)
    }

    /// Replace one series' connection geometry transactionally.
    pub fn set_sheet_chart_series_connection_line(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        value: ChartSeriesConnectionLine,
    ) -> Result<()> {
        let mut values = read(self, sheet_id, drawable_object_id)?;
        let count = values.len();
        let target = values
            .get_mut(series.zero_based())
            .ok_or_else(|| index_error(drawable_object_id, series, count))?;
        if *target == value {
            return Ok(());
        }
        *target = value;
        set(self, sheet_id, drawable_object_id, &values)
    }
}

fn read(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<Vec<ChartSeriesConnectionLine>> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Numbers",
        drawable_object_id,
    )?;
    read_native(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        graph.info.kind,
        count,
    )
}

fn set(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    values: &[ChartSeriesConnectionLine],
) -> Result<()> {
    if read(editor, sheet_id, drawable_object_id)? == values {
        return Ok(());
    }
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Numbers",
        drawable_object_id,
    )?;
    let mut staged = editor.package().clone();
    set_native(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        graph.info.kind,
        count,
        values,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_series_connection_lines(sheet_id, drawable_object_id)? != values {
        return Err(Error::InvalidFormat(
            "Numbers chart series connection-line update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn index_error(drawable_object_id: u64, series: ChartSeriesIndex, count: usize) -> Error {
    Error::InvalidFormat(format!(
        "Numbers chart {drawable_object_id} has {count} series, not series {}",
        series.zero_based() + 1
    ))
}
