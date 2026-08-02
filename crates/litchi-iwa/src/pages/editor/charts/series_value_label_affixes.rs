//! Native per-series value-label prefix and suffix CRUD for Pages charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_value_label_affixes::{
    chart_series_value_label_affixes as read_native_affixes,
    set_chart_series_value_label_affixes as set_native_affixes,
};
use crate::charts::{ChartSeriesIndex, ChartSeriesValueLabelAffixes};

impl PagesEditor {
    /// Read every series' value-label affixes in native series order.
    pub fn body_chart_series_value_label_affixes(
        &self,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartSeriesValueLabelAffixes>> {
        body_chart_series_value_label_affixes(self, drawable_object_id)
    }

    /// Read one series' value-label affixes.
    pub fn body_chart_series_value_label_affix(
        &self,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<ChartSeriesValueLabelAffixes> {
        let affixes = body_chart_series_value_label_affixes(self, drawable_object_id)?;
        affixes
            .get(series.zero_based())
            .cloned()
            .ok_or_else(|| affix_index_error("Pages", drawable_object_id, series, affixes.len()))
    }

    /// Set every series' value-label affixes in native series order.
    pub fn set_body_chart_series_value_label_affixes(
        &mut self,
        drawable_object_id: u64,
        affixes: &[ChartSeriesValueLabelAffixes],
    ) -> Result<()> {
        set_body_chart_series_value_label_affixes(self, drawable_object_id, affixes)
    }

    /// Set one series' value-label affixes.
    pub fn set_body_chart_series_value_label_affix(
        &mut self,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        affixes: ChartSeriesValueLabelAffixes,
    ) -> Result<()> {
        let mut values = body_chart_series_value_label_affixes(self, drawable_object_id)?;
        let count = values.len();
        let target = values
            .get_mut(series.zero_based())
            .ok_or_else(|| affix_index_error("Pages", drawable_object_id, series, count))?;
        if *target == affixes {
            return Ok(());
        }
        *target = affixes;
        set_body_chart_series_value_label_affixes(self, drawable_object_id, &values)
    }
}

fn body_chart_series_value_label_affixes(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<Vec<ChartSeriesValueLabelAffixes>> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    read_native_affixes(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        graph.info.kind,
        series_count,
    )
}

fn set_body_chart_series_value_label_affixes(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    affixes: &[ChartSeriesValueLabelAffixes],
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    if affixes.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} requires {series_count} series value-label affix pairs, got {}",
            affixes.len()
        )));
    }
    if read_native_affixes(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
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
        "Pages",
        graph.info.kind,
        &expected,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_series_value_label_affixes(drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Pages chart series value-label affix update failed validation".to_owned(),
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
