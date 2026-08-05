//! Native per-series value-label CRUD for Keynote charts.

use super::*;
use crate::charts::series_topology::chart_series_count;
use crate::charts::series_value_label_location::{
    chart_series_value_label_locations as read_native_locations,
    set_chart_series_value_label_locations as set_native_locations,
};
use crate::charts::series_value_labels::{
    chart_series_value_label_visibilities as read_native_value_labels,
    set_chart_series_value_label_visibilities as set_native_value_labels,
};
use crate::charts::{
    ChartSeriesIndex, ChartSeriesValueLabelLocation, ChartSeriesValueLabelVisibility,
};

impl KeynoteEditor {
    /// Read every series' value-label visibility in native series order.
    pub fn slide_chart_series_value_label_visibilities(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartSeriesValueLabelVisibility>> {
        slide_chart_series_value_label_visibilities(self, slide_index, drawable_object_id)
    }

    /// Read one series' value-label visibility.
    pub fn slide_chart_series_value_label_visibility(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<ChartSeriesValueLabelVisibility> {
        let visibilities =
            slide_chart_series_value_label_visibilities(self, slide_index, drawable_object_id)?;
        visibilities
            .get(series.zero_based())
            .copied()
            .ok_or_else(|| {
                value_label_index_error("Keynote", drawable_object_id, series, visibilities.len())
            })
    }

    /// Set every series' value-label visibility in native series order.
    pub fn set_slide_chart_series_value_label_visibilities(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        visibilities: &[ChartSeriesValueLabelVisibility],
    ) -> Result<()> {
        set_slide_chart_series_value_label_visibilities(
            self,
            slide_index,
            drawable_object_id,
            visibilities,
        )
    }

    /// Set one series' value-label visibility.
    pub fn set_slide_chart_series_value_label_visibility(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        visibility: ChartSeriesValueLabelVisibility,
    ) -> Result<()> {
        let mut visibilities =
            slide_chart_series_value_label_visibilities(self, slide_index, drawable_object_id)?;
        let count = visibilities.len();
        let target = visibilities
            .get_mut(series.zero_based())
            .ok_or_else(|| value_label_index_error("Keynote", drawable_object_id, series, count))?;
        if *target == visibility {
            return Ok(());
        }
        *target = visibility;
        set_slide_chart_series_value_label_visibilities(
            self,
            slide_index,
            drawable_object_id,
            &visibilities,
        )
    }

    /// Read every series' value-label Location setting in native series order.
    pub fn slide_chart_series_value_label_locations(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartSeriesValueLabelLocation>> {
        slide_chart_series_value_label_locations(self, slide_index, drawable_object_id)
    }

    /// Read one series' value-label Location setting.
    pub fn slide_chart_series_value_label_location(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<ChartSeriesValueLabelLocation> {
        let locations =
            slide_chart_series_value_label_locations(self, slide_index, drawable_object_id)?;
        locations.get(series.zero_based()).copied().ok_or_else(|| {
            value_label_index_error("Keynote", drawable_object_id, series, locations.len())
        })
    }

    /// Set every series' value-label Location setting in native series order.
    pub fn set_slide_chart_series_value_label_locations(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        locations: &[ChartSeriesValueLabelLocation],
    ) -> Result<()> {
        set_slide_chart_series_value_label_locations(
            self,
            slide_index,
            drawable_object_id,
            locations,
        )
    }

    /// Set one series' value-label Location setting.
    pub fn set_slide_chart_series_value_label_location(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        location: ChartSeriesValueLabelLocation,
    ) -> Result<()> {
        let mut locations =
            slide_chart_series_value_label_locations(self, slide_index, drawable_object_id)?;
        let count = locations.len();
        let target = locations
            .get_mut(series.zero_based())
            .ok_or_else(|| value_label_index_error("Keynote", drawable_object_id, series, count))?;
        if *target == location {
            return Ok(());
        }
        *target = location;
        set_slide_chart_series_value_label_locations(
            self,
            slide_index,
            drawable_object_id,
            &locations,
        )
    }
}

fn slide_chart_series_value_label_visibilities(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<Vec<ChartSeriesValueLabelVisibility>> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )?;
    read_native_value_labels(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        graph.info.kind,
        series_count,
    )
}

fn set_slide_chart_series_value_label_visibilities(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    visibilities: &[ChartSeriesValueLabelVisibility],
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )?;
    if visibilities.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Keynote chart {drawable_object_id} requires {series_count} series value-label visibilities, got {}",
            visibilities.len()
        )));
    }
    if read_native_value_labels(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        graph.info.kind,
        series_count,
    )? == visibilities
    {
        return Ok(());
    }
    let expected = visibilities.to_vec();
    let mut staged = editor.package().clone();
    set_native_value_labels(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        graph.info.kind,
        &expected,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_series_value_label_visibilities(slide_index, drawable_object_id)?
        != expected
    {
        return Err(Error::InvalidFormat(
            "Keynote chart series value-label update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn slide_chart_series_value_label_locations(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<Vec<ChartSeriesValueLabelLocation>> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )?;
    read_native_locations(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        graph.info.kind,
        series_count,
    )
}

fn set_slide_chart_series_value_label_locations(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    locations: &[ChartSeriesValueLabelLocation],
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Keynote",
        drawable_object_id,
    )?;
    if locations.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Keynote chart {drawable_object_id} requires {series_count} series value-label Location settings, got {}",
            locations.len()
        )));
    }
    if read_native_locations(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        graph.info.kind,
        series_count,
    )? == locations
    {
        return Ok(());
    }
    let expected = locations.to_vec();
    let mut staged = editor.package().clone();
    set_native_locations(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        graph.info.kind,
        series_count,
        &expected,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_series_value_label_locations(slide_index, drawable_object_id)?
        != expected
    {
        return Err(Error::InvalidFormat(
            "Keynote chart series value-label Location update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

pub(super) fn value_label_series_count(
    kind: ChartKind,
    direction: Direction,
    data: &ChartData,
    drawable_label: &str,
    drawable_object_id: u64,
) -> Result<usize> {
    chart_series_count(kind, direction, data, drawable_label, drawable_object_id)
}

fn value_label_index_error(
    drawable_label: &str,
    drawable_object_id: u64,
    series: ChartSeriesIndex,
    series_count: usize,
) -> Error {
    Error::InvalidFormat(format!(
        "{drawable_label} chart {drawable_object_id} series index {} exceeds series count {series_count}",
        series.zero_based()
    ))
}
