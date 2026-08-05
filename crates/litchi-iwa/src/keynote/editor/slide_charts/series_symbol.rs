//! Per-series data-symbol CRUD for Keynote charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_symbol::{
    chart_series_symbols as read_native_symbols, set_chart_series_symbols as set_native_symbols,
};
use crate::charts::{ChartSeriesSymbol, Index};

impl KeynoteEditor {
    pub fn slide_chart_series_symbols(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<Option<ChartSeriesSymbol>>> {
        slide_chart_series_symbols(self, slide_index, drawable_object_id)
    }

    pub fn slide_chart_series_symbol(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        series: Index,
    ) -> Result<Option<ChartSeriesSymbol>> {
        let symbols = self.slide_chart_series_symbols(slide_index, drawable_object_id)?;
        symbols
            .get(series.zero_based())
            .copied()
            .ok_or_else(|| symbol_index_error("Keynote", drawable_object_id, series, symbols.len()))
    }

    pub fn set_slide_chart_series_symbols(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        symbols: &[Option<ChartSeriesSymbol>],
    ) -> Result<()> {
        set_slide_chart_series_symbols(self, slide_index, drawable_object_id, symbols)
    }

    pub fn set_slide_chart_series_symbol(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        series: Index,
        symbol: Option<ChartSeriesSymbol>,
    ) -> Result<()> {
        let mut symbols = self.slide_chart_series_symbols(slide_index, drawable_object_id)?;
        let count = symbols.len();
        let target = symbols
            .get_mut(series.zero_based())
            .ok_or_else(|| symbol_index_error("Keynote", drawable_object_id, series, count))?;
        *target = symbol;
        self.set_slide_chart_series_symbols(slide_index, drawable_object_id, &symbols)
    }
}

fn slide_chart_series_symbols(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<Vec<Option<ChartSeriesSymbol>>> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )?;
    read_native_symbols(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        graph.info.kind,
        series_count,
    )
}

fn set_slide_chart_series_symbols(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    symbols: &[Option<ChartSeriesSymbol>],
) -> Result<()> {
    if editor.slide_chart_series_symbols(slide_index, drawable_object_id)? == symbols {
        return Ok(());
    }
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )?;
    let mut staged = editor.package().clone();
    set_native_symbols(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        graph.info.kind,
        series_count,
        symbols,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_series_symbols(slide_index, drawable_object_id)? != symbols {
        return Err(Error::InvalidFormat(
            "Keynote chart data-symbol update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn symbol_index_error(
    application: &str,
    drawable_object_id: u64,
    series: Index,
    count: usize,
) -> Error {
    Error::InvalidFormat(format!(
        "{application} chart {drawable_object_id} has {count} series, not series {}",
        series.zero_based() + 1
    ))
}
