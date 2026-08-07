//! Inherited data-symbol outline CRUD for Pages charts.

use super::graph::BodyChartGraph;
use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_symbol_outline::{
    chart_series_symbol_outlines as read_native, reset_chart_series_symbol_outline as reset_native,
    set_chart_series_symbol_outlines as set_native,
};
use crate::charts::{ChartSeriesStroke, Index};

impl PagesEditor {
    pub fn body_chart_series_symbol_outlines(
        &self,
        drawable_object_id: u64,
    ) -> Result<Vec<Option<ChartSeriesStroke>>> {
        read(self, drawable_object_id)
    }

    pub fn body_chart_series_symbol_outline(
        &self,
        drawable_object_id: u64,
        series: Index,
    ) -> Result<Option<ChartSeriesStroke>> {
        let values = read(self, drawable_object_id)?;
        values
            .get(series.zero_based())
            .copied()
            .ok_or_else(|| index_error(drawable_object_id, series, values.len()))
    }

    pub fn set_body_chart_series_symbol_outlines(
        &mut self,
        drawable_object_id: u64,
        outlines: &[Option<ChartSeriesStroke>],
    ) -> Result<()> {
        set(self, drawable_object_id, outlines)
    }

    pub fn set_body_chart_series_symbol_outline(
        &mut self,
        drawable_object_id: u64,
        series: Index,
        outline: Option<ChartSeriesStroke>,
    ) -> Result<()> {
        let mut values = read(self, drawable_object_id)?;
        let count = values.len();
        let target = values
            .get_mut(series.zero_based())
            .ok_or_else(|| index_error(drawable_object_id, series, count))?;
        if *target == outline {
            return Ok(());
        }
        *target = outline;
        set(self, drawable_object_id, &values)
    }

    pub fn reset_body_chart_series_symbol_outline(
        &mut self,
        drawable_object_id: u64,
        series: Index,
    ) -> Result<Option<ChartSeriesStroke>> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        let count = series_count(&graph, drawable_object_id)?;
        let mut staged = self.package().clone();
        let inherited = reset_native(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            graph.info.kind,
            count,
            series.zero_based(),
        )?;
        let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.body_chart_series_symbol_outline(drawable_object_id, series)? != inherited {
            return Err(Error::InvalidFormat(
                "Pages chart data-symbol outline reset failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(inherited)
    }
}

fn read(editor: &PagesEditor, drawable_object_id: u64) -> Result<Vec<Option<ChartSeriesStroke>>> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        graph.info.kind,
        series_count(&graph, drawable_object_id)?,
    )
}

fn set(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    outlines: &[Option<ChartSeriesStroke>],
) -> Result<()> {
    if read(editor, drawable_object_id)? == outlines {
        return Ok(());
    }
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        graph.info.kind,
        series_count(&graph, drawable_object_id)?,
        outlines,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_series_symbol_outlines(drawable_object_id)? != outlines {
        return Err(Error::InvalidFormat(
            "Pages chart data-symbol outline update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn series_count(graph: &BodyChartGraph, drawable_object_id: u64) -> Result<usize> {
    value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )
}

fn index_error(drawable_object_id: u64, series: Index, count: usize) -> Error {
    Error::InvalidFormat(format!(
        "Pages chart {drawable_object_id} has {count} series, not series {}",
        series.zero_based() + 1
    ))
}
