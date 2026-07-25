//! Per-series data-symbol CRUD for Numbers charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_symbol::{
    chart_series_symbols as read_native_symbols, set_chart_series_symbols as set_native_symbols,
};
use crate::charts::{ChartSeriesIndex, ChartSeriesSymbol};

impl NumbersEditor {
    pub fn sheet_chart_series_symbols(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<Option<ChartSeriesSymbol>>> {
        sheet_chart_series_symbols(self, sheet_id, drawable_object_id)
    }

    pub fn sheet_chart_series_symbol(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<Option<ChartSeriesSymbol>> {
        let symbols = self.sheet_chart_series_symbols(sheet_id, drawable_object_id)?;
        symbols
            .get(series.zero_based())
            .copied()
            .ok_or_else(|| symbol_index_error("Numbers", drawable_object_id, series, symbols.len()))
    }

    pub fn set_sheet_chart_series_symbols(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        symbols: &[Option<ChartSeriesSymbol>],
    ) -> Result<()> {
        set_sheet_chart_series_symbols(self, sheet_id, drawable_object_id, symbols)
    }

    pub fn set_sheet_chart_series_symbol(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        symbol: Option<ChartSeriesSymbol>,
    ) -> Result<()> {
        let mut symbols = self.sheet_chart_series_symbols(sheet_id, drawable_object_id)?;
        let count = symbols.len();
        let target = symbols
            .get_mut(series.zero_based())
            .ok_or_else(|| symbol_index_error("Numbers", drawable_object_id, series, count))?;
        *target = symbol;
        self.set_sheet_chart_series_symbols(sheet_id, drawable_object_id, &symbols)
    }
}

fn sheet_chart_series_symbols(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<Vec<Option<ChartSeriesSymbol>>> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Numbers",
        drawable_object_id,
    )?;
    read_native_symbols(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        graph.info.kind,
        series_count,
    )
}

fn set_sheet_chart_series_symbols(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    symbols: &[Option<ChartSeriesSymbol>],
) -> Result<()> {
    if editor.sheet_chart_series_symbols(sheet_id, drawable_object_id)? == symbols {
        return Ok(());
    }
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Numbers",
        drawable_object_id,
    )?;
    let mut staged = editor.package().clone();
    set_native_symbols(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        graph.info.kind,
        series_count,
        symbols,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_series_symbols(sheet_id, drawable_object_id)? != symbols {
        return Err(Error::InvalidFormat(
            "Numbers chart data-symbol update failed validation".to_owned(),
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
