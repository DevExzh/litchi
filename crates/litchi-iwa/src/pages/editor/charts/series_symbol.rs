//! Per-series data-symbol CRUD for Pages charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_symbol::{
    chart_series_symbols as read_native_symbols, set_chart_series_symbols as set_native_symbols,
};
use crate::charts::{ChartSeriesIndex, ChartSeriesSymbol};

impl PagesEditor {
    pub fn body_chart_series_symbols(
        &self,
        drawable_object_id: u64,
    ) -> Result<Vec<Option<ChartSeriesSymbol>>> {
        body_chart_series_symbols(self, drawable_object_id)
    }

    pub fn body_chart_series_symbol(
        &self,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<Option<ChartSeriesSymbol>> {
        let symbols = self.body_chart_series_symbols(drawable_object_id)?;
        symbols
            .get(series.zero_based())
            .copied()
            .ok_or_else(|| symbol_index_error("Pages", drawable_object_id, series, symbols.len()))
    }

    pub fn set_body_chart_series_symbols(
        &mut self,
        drawable_object_id: u64,
        symbols: &[Option<ChartSeriesSymbol>],
    ) -> Result<()> {
        set_body_chart_series_symbols(self, drawable_object_id, symbols)
    }

    pub fn set_body_chart_series_symbol(
        &mut self,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        symbol: Option<ChartSeriesSymbol>,
    ) -> Result<()> {
        let mut symbols = self.body_chart_series_symbols(drawable_object_id)?;
        let count = symbols.len();
        let target = symbols
            .get_mut(series.zero_based())
            .ok_or_else(|| symbol_index_error("Pages", drawable_object_id, series, count))?;
        *target = symbol;
        self.set_body_chart_series_symbols(drawable_object_id, &symbols)
    }
}

fn body_chart_series_symbols(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<Vec<Option<ChartSeriesSymbol>>> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    read_native_symbols(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        graph.info.kind,
        series_count,
    )
}

fn set_body_chart_series_symbols(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    symbols: &[Option<ChartSeriesSymbol>],
) -> Result<()> {
    if editor.body_chart_series_symbols(drawable_object_id)? == symbols {
        return Ok(());
    }
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    let mut staged = editor.package().clone();
    set_native_symbols(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        graph.info.kind,
        series_count,
        symbols,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_series_symbols(drawable_object_id)? != symbols {
        return Err(Error::InvalidFormat(
            "Pages chart data-symbol update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn symbol_index_error(
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
