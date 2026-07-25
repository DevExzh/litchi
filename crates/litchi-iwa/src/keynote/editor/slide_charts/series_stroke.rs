//! Inherited per-series stroke CRUD for Keynote charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_stroke::{
    chart_series_strokes as read_native_strokes, reset_chart_series_stroke as reset_native_stroke,
    set_chart_series_strokes as set_native_strokes,
};
use crate::charts::{ChartSeriesIndex, ChartSeriesStroke};

impl KeynoteEditor {
    /// Read effective strokes in native series order.
    pub fn slide_chart_series_strokes(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<Option<ChartSeriesStroke>>> {
        slide_chart_series_strokes(self, slide_index, drawable_object_id)
    }

    /// Read one effective series stroke. `None` means the stroke is hidden.
    pub fn slide_chart_series_stroke(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<Option<ChartSeriesStroke>> {
        let strokes = slide_chart_series_strokes(self, slide_index, drawable_object_id)?;
        strokes.get(series.zero_based()).copied().ok_or_else(|| {
            series_stroke_index_error("Keynote", drawable_object_id, series, strokes.len())
        })
    }

    /// Replace every series stroke transactionally.
    pub fn set_slide_chart_series_strokes(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        strokes: &[Option<ChartSeriesStroke>],
    ) -> Result<()> {
        set_slide_chart_series_strokes(self, slide_index, drawable_object_id, strokes)
    }

    /// Replace one series stroke transactionally. `None` hides the stroke.
    pub fn set_slide_chart_series_stroke(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        stroke: Option<ChartSeriesStroke>,
    ) -> Result<()> {
        let mut strokes = slide_chart_series_strokes(self, slide_index, drawable_object_id)?;
        let count = strokes.len();
        let target = strokes.get_mut(series.zero_based()).ok_or_else(|| {
            series_stroke_index_error("Keynote", drawable_object_id, series, count)
        })?;
        if *target == stroke {
            return Ok(());
        }
        *target = stroke;
        set_slide_chart_series_strokes(self, slide_index, drawable_object_id, &strokes)
    }

    /// Remove one local override and reveal its inherited series stroke.
    pub fn reset_slide_chart_series_stroke(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<Option<ChartSeriesStroke>> {
        reset_slide_chart_series_stroke(self, slide_index, drawable_object_id, series)
    }
}

fn slide_chart_series_strokes(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<Vec<Option<ChartSeriesStroke>>> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )?;
    read_native_strokes(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        graph.info.kind,
        series_count,
    )
}

fn set_slide_chart_series_strokes(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    strokes: &[Option<ChartSeriesStroke>],
) -> Result<()> {
    if slide_chart_series_strokes(editor, slide_index, drawable_object_id)? == strokes {
        return Ok(());
    }
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )?;
    let mut staged = editor.package().clone();
    set_native_strokes(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        graph.info.kind,
        series_count,
        strokes,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_series_strokes(slide_index, drawable_object_id)? != strokes {
        return Err(Error::InvalidFormat(
            "Keynote chart series stroke update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn reset_slide_chart_series_stroke(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    series: ChartSeriesIndex,
) -> Result<Option<ChartSeriesStroke>> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )?;
    let mut staged = editor.package().clone();
    let inherited = reset_native_stroke(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        graph.info.kind,
        series_count,
        series.zero_based(),
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_series_stroke(slide_index, drawable_object_id, series)? != inherited {
        return Err(Error::InvalidFormat(
            "Keynote chart series stroke reset failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(inherited)
}

fn series_stroke_index_error(
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
