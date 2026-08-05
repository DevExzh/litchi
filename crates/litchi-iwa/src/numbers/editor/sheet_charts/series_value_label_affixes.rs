//! Native per-series value-label prefix and suffix CRUD for Numbers charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_value_label_affixes::{
    chart_series_value_label_affixes as read_native_affixes,
    set_chart_series_value_label_affixes as set_native_affixes,
};
use crate::charts::{ChartSeriesIndex, LabelAffixes};

impl NumbersEditor {
    /// Read every series' value-label affixes in native series order.
    pub fn sheet_chart_series_value_label_affixes(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<LabelAffixes>> {
        sheet_chart_series_value_label_affixes(self, sheet_id, drawable_object_id)
    }

    /// Read one series' value-label affixes.
    pub fn sheet_chart_series_value_label_affix(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<LabelAffixes> {
        let affixes = sheet_chart_series_value_label_affixes(self, sheet_id, drawable_object_id)?;
        affixes
            .get(series.zero_based())
            .cloned()
            .ok_or_else(|| affix_index_error("Numbers", drawable_object_id, series, affixes.len()))
    }

    /// Set every series' value-label affixes in native series order.
    pub fn set_sheet_chart_series_value_label_affixes(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        affixes: &[LabelAffixes],
    ) -> Result<()> {
        set_sheet_chart_series_value_label_affixes(self, sheet_id, drawable_object_id, affixes)
    }

    /// Set one series' value-label affixes.
    pub fn set_sheet_chart_series_value_label_affix(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        affixes: LabelAffixes,
    ) -> Result<()> {
        let mut values =
            sheet_chart_series_value_label_affixes(self, sheet_id, drawable_object_id)?;
        let count = values.len();
        let target = values
            .get_mut(series.zero_based())
            .ok_or_else(|| affix_index_error("Numbers", drawable_object_id, series, count))?;
        if *target == affixes {
            return Ok(());
        }
        *target = affixes;
        set_sheet_chart_series_value_label_affixes(self, sheet_id, drawable_object_id, &values)
    }
}

fn sheet_chart_series_value_label_affixes(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<Vec<LabelAffixes>> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Numbers",
        drawable_object_id,
    )?;
    read_native_affixes(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        graph.info.kind,
        series_count,
    )
}

fn set_sheet_chart_series_value_label_affixes(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    affixes: &[LabelAffixes],
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Numbers",
        drawable_object_id,
    )?;
    if affixes.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} requires {series_count} series value-label affix pairs, got {}",
            affixes.len()
        )));
    }
    if read_native_affixes(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        graph.info.kind,
        series_count,
    )? == affixes
    {
        return Ok(());
    }
    let expected = affixes.to_vec();
    let mut staged = editor.package().clone();
    set_native_affixes(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        graph.info.kind,
        &expected,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_series_value_label_affixes(sheet_id, drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Numbers chart series value-label affix update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn affix_index_error(
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
