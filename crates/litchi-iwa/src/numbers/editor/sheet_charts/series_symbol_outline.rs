//! Inherited data-symbol outline CRUD for Numbers charts.

use super::graph::SheetChartGraph;
use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_symbol_outline::{
    chart_series_symbol_outlines as read_native, reset_chart_series_symbol_outline as reset_native,
    set_chart_series_symbol_outlines as set_native,
};
use crate::charts::{ChartSeriesIndex, ChartSeriesStroke};

impl NumbersEditor {
    pub fn sheet_chart_series_symbol_outlines(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<Option<ChartSeriesStroke>>> {
        read(self, sheet_id, drawable_object_id)
    }

    pub fn sheet_chart_series_symbol_outline(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<Option<ChartSeriesStroke>> {
        let values = read(self, sheet_id, drawable_object_id)?;
        values
            .get(series.zero_based())
            .copied()
            .ok_or_else(|| index_error("Numbers", drawable_object_id, series, values.len()))
    }

    pub fn set_sheet_chart_series_symbol_outlines(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        outlines: &[Option<ChartSeriesStroke>],
    ) -> Result<()> {
        set(self, sheet_id, drawable_object_id, outlines)
    }

    pub fn set_sheet_chart_series_symbol_outline(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        outline: Option<ChartSeriesStroke>,
    ) -> Result<()> {
        let mut values = read(self, sheet_id, drawable_object_id)?;
        let count = values.len();
        let target = values
            .get_mut(series.zero_based())
            .ok_or_else(|| index_error("Numbers", drawable_object_id, series, count))?;
        if *target == outline {
            return Ok(());
        }
        *target = outline;
        set(self, sheet_id, drawable_object_id, &values)
    }

    pub fn reset_sheet_chart_series_symbol_outline(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<Option<ChartSeriesStroke>> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        let count = series_count(&graph, drawable_object_id)?;
        let mut staged = self.package().clone();
        let inherited = reset_native(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            graph.info.kind,
            count,
            series.zero_based(),
        )?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.sheet_chart_series_symbol_outline(sheet_id, drawable_object_id, series)?
            != inherited
        {
            return Err(Error::InvalidFormat(
                "Numbers chart data-symbol outline reset failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(inherited)
    }
}

fn read(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<Vec<Option<ChartSeriesStroke>>> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        graph.info.kind,
        series_count(&graph, drawable_object_id)?,
    )
}

fn set(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    outlines: &[Option<ChartSeriesStroke>],
) -> Result<()> {
    if read(editor, sheet_id, drawable_object_id)? == outlines {
        return Ok(());
    }
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        graph.info.kind,
        series_count(&graph, drawable_object_id)?,
        outlines,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_series_symbol_outlines(sheet_id, drawable_object_id)? != outlines {
        return Err(Error::InvalidFormat(
            "Numbers chart data-symbol outline update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn series_count(graph: &SheetChartGraph, drawable_object_id: u64) -> Result<usize> {
    value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Numbers",
        drawable_object_id,
    )
}

fn index_error(
    application: &str,
    drawable_object_id: u64,
    series: ChartSeriesIndex,
    count: usize,
) -> Error {
    Error::InvalidFormat(format!(
        "{application} chart {drawable_object_id} has {count} series, not series {}",
        series.zero_based() + 1
    ))
}
