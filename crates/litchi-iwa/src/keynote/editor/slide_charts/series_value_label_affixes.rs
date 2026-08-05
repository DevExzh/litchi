//! Native per-series value-label prefix and suffix CRUD for Keynote charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_value_label_affixes::{
    chart_series_value_label_affixes as read_native_affixes,
    set_chart_series_value_label_affixes as set_native_affixes,
};
use crate::charts::{ChartSeriesIndex, LabelAffixes};

impl KeynoteEditor {
    /// Read every series' value-label affixes in native series order.
    pub fn slide_chart_series_value_label_affixes(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<LabelAffixes>> {
        slide_chart_series_value_label_affixes(self, slide_index, drawable_object_id)
    }

    /// Read one series' value-label affixes.
    pub fn slide_chart_series_value_label_affix(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<LabelAffixes> {
        let affixes =
            slide_chart_series_value_label_affixes(self, slide_index, drawable_object_id)?;
        affixes
            .get(series.zero_based())
            .cloned()
            .ok_or_else(|| affix_index_error("Keynote", drawable_object_id, series, affixes.len()))
    }

    /// Set every series' value-label affixes in native series order.
    pub fn set_slide_chart_series_value_label_affixes(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        affixes: &[LabelAffixes],
    ) -> Result<()> {
        set_slide_chart_series_value_label_affixes(self, slide_index, drawable_object_id, affixes)
    }

    /// Set one series' value-label affixes.
    pub fn set_slide_chart_series_value_label_affix(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        affixes: LabelAffixes,
    ) -> Result<()> {
        let mut values =
            slide_chart_series_value_label_affixes(self, slide_index, drawable_object_id)?;
        let count = values.len();
        let target = values
            .get_mut(series.zero_based())
            .ok_or_else(|| affix_index_error("Keynote", drawable_object_id, series, count))?;
        if *target == affixes {
            return Ok(());
        }
        *target = affixes;
        set_slide_chart_series_value_label_affixes(self, slide_index, drawable_object_id, &values)
    }
}

fn slide_chart_series_value_label_affixes(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<Vec<LabelAffixes>> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )?;
    read_native_affixes(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        graph.info.kind,
        series_count,
    )
}

fn set_slide_chart_series_value_label_affixes(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    affixes: &[LabelAffixes],
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )?;
    if affixes.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Keynote chart {drawable_object_id} requires {series_count} series value-label affix pairs, got {}",
            affixes.len()
        )));
    }
    if read_native_affixes(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
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
        "Keynote",
        graph.info.kind,
        &expected,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_series_value_label_affixes(slide_index, drawable_object_id)? != expected
    {
        return Err(Error::InvalidFormat(
            "Keynote chart series value-label affix update failed validation".to_owned(),
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
