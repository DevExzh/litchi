//! Inherited data-symbol outline CRUD for Keynote charts.

use super::graph::SlideChartGraph;
use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_symbol_outline::{
    chart_series_symbol_outlines as read_native, reset_chart_series_symbol_outline as reset_native,
    set_chart_series_symbol_outlines as set_native,
};
use crate::charts::{ChartSeriesStroke, Index};

impl KeynoteEditor {
    pub fn slide_chart_series_symbol_outlines(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<Option<ChartSeriesStroke>>> {
        read(self, slide_index, drawable_object_id)
    }

    pub fn slide_chart_series_symbol_outline(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        series: Index,
    ) -> Result<Option<ChartSeriesStroke>> {
        let values = read(self, slide_index, drawable_object_id)?;
        values
            .get(series.zero_based())
            .copied()
            .ok_or_else(|| index_error(drawable_object_id, series, values.len()))
    }

    pub fn set_slide_chart_series_symbol_outlines(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        outlines: &[Option<ChartSeriesStroke>],
    ) -> Result<()> {
        set(self, slide_index, drawable_object_id, outlines)
    }

    pub fn set_slide_chart_series_symbol_outline(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        series: Index,
        outline: Option<ChartSeriesStroke>,
    ) -> Result<()> {
        let mut values = read(self, slide_index, drawable_object_id)?;
        let count = values.len();
        let target = values
            .get_mut(series.zero_based())
            .ok_or_else(|| index_error(drawable_object_id, series, count))?;
        if *target == outline {
            return Ok(());
        }
        *target = outline;
        set(self, slide_index, drawable_object_id, &values)
    }

    pub fn reset_slide_chart_series_symbol_outline(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        series: Index,
    ) -> Result<Option<ChartSeriesStroke>> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        let count = series_count(&graph, drawable_object_id)?;
        let mut staged = self.package().clone();
        let inherited = reset_native(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            graph.info.kind,
            count,
            series.zero_based(),
        )?;
        let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.slide_chart_series_symbol_outline(slide_index, drawable_object_id, series)?
            != inherited
        {
            return Err(Error::InvalidFormat(
                "Keynote chart data-symbol outline reset failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(inherited)
    }
}

fn read(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<Vec<Option<ChartSeriesStroke>>> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        graph.info.kind,
        series_count(&graph, drawable_object_id)?,
    )
}

fn set(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    outlines: &[Option<ChartSeriesStroke>],
) -> Result<()> {
    if read(editor, slide_index, drawable_object_id)? == outlines {
        return Ok(());
    }
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        graph.info.kind,
        series_count(&graph, drawable_object_id)?,
        outlines,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_series_symbol_outlines(slide_index, drawable_object_id)? != outlines {
        return Err(Error::InvalidFormat(
            "Keynote chart data-symbol outline update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn series_count(graph: &SlideChartGraph, drawable_object_id: u64) -> Result<usize> {
    value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )
}

fn index_error(drawable_object_id: u64, series: Index, count: usize) -> Error {
    Error::InvalidFormat(format!(
        "Keynote chart {drawable_object_id} has {count} series, not series {}",
        series.zero_based() + 1
    ))
}
