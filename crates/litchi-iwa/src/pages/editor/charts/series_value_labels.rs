//! Native per-series value-label CRUD for Pages charts.

use super::*;
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

impl PagesEditor {
    /// Read every series' value-label visibility in native series order.
    pub fn body_chart_series_value_label_visibilities(
        &self,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartSeriesValueLabelVisibility>> {
        body_chart_series_value_label_visibilities(self, drawable_object_id)
    }

    /// Read one series' value-label visibility.
    pub fn body_chart_series_value_label_visibility(
        &self,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<ChartSeriesValueLabelVisibility> {
        let visibilities = body_chart_series_value_label_visibilities(self, drawable_object_id)?;
        visibilities
            .get(series.zero_based())
            .copied()
            .ok_or_else(|| {
                value_label_index_error("Pages", drawable_object_id, series, visibilities.len())
            })
    }

    /// Set every series' value-label visibility in native series order.
    pub fn set_body_chart_series_value_label_visibilities(
        &mut self,
        drawable_object_id: u64,
        visibilities: &[ChartSeriesValueLabelVisibility],
    ) -> Result<()> {
        set_body_chart_series_value_label_visibilities(self, drawable_object_id, visibilities)
    }

    /// Set one series' value-label visibility.
    pub fn set_body_chart_series_value_label_visibility(
        &mut self,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        visibility: ChartSeriesValueLabelVisibility,
    ) -> Result<()> {
        let mut visibilities =
            body_chart_series_value_label_visibilities(self, drawable_object_id)?;
        let count = visibilities.len();
        let target = visibilities
            .get_mut(series.zero_based())
            .ok_or_else(|| value_label_index_error("Pages", drawable_object_id, series, count))?;
        if *target == visibility {
            return Ok(());
        }
        *target = visibility;
        set_body_chart_series_value_label_visibilities(self, drawable_object_id, &visibilities)
    }

    /// Read every series' value-label Location setting in native series order.
    pub fn body_chart_series_value_label_locations(
        &self,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartSeriesValueLabelLocation>> {
        body_chart_series_value_label_locations(self, drawable_object_id)
    }

    /// Read one series' value-label Location setting.
    pub fn body_chart_series_value_label_location(
        &self,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<ChartSeriesValueLabelLocation> {
        let locations = body_chart_series_value_label_locations(self, drawable_object_id)?;
        locations.get(series.zero_based()).copied().ok_or_else(|| {
            value_label_index_error("Pages", drawable_object_id, series, locations.len())
        })
    }

    /// Set every series' value-label Location setting in native series order.
    pub fn set_body_chart_series_value_label_locations(
        &mut self,
        drawable_object_id: u64,
        locations: &[ChartSeriesValueLabelLocation],
    ) -> Result<()> {
        set_body_chart_series_value_label_locations(self, drawable_object_id, locations)
    }

    /// Set one series' value-label Location setting.
    pub fn set_body_chart_series_value_label_location(
        &mut self,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        location: ChartSeriesValueLabelLocation,
    ) -> Result<()> {
        let mut locations = body_chart_series_value_label_locations(self, drawable_object_id)?;
        let count = locations.len();
        let target = locations
            .get_mut(series.zero_based())
            .ok_or_else(|| value_label_index_error("Pages", drawable_object_id, series, count))?;
        if *target == location {
            return Ok(());
        }
        *target = location;
        set_body_chart_series_value_label_locations(self, drawable_object_id, &locations)
    }
}

fn body_chart_series_value_label_visibilities(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<Vec<ChartSeriesValueLabelVisibility>> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    read_native_value_labels(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        graph.info.kind,
        series_count,
    )
}

fn set_body_chart_series_value_label_visibilities(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    visibilities: &[ChartSeriesValueLabelVisibility],
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    if visibilities.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} requires {series_count} series value-label visibilities, got {}",
            visibilities.len()
        )));
    }
    if read_native_value_labels(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
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
        "Pages",
        graph.info.kind,
        &expected,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_series_value_label_visibilities(drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Pages chart series value-label update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn body_chart_series_value_label_locations(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<Vec<ChartSeriesValueLabelLocation>> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    read_native_locations(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        graph.info.kind,
        series_count,
    )
}

fn set_body_chart_series_value_label_locations(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    locations: &[ChartSeriesValueLabelLocation],
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    if locations.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} requires {series_count} series value-label Location settings, got {}",
            locations.len()
        )));
    }
    if read_native_locations(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
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
        "Pages",
        graph.info.kind,
        series_count,
        &expected,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_series_value_label_locations(drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Pages chart series value-label Location update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

pub(super) fn value_label_series_count(
    direction: ChartSeriesDirection,
    data: &ChartData,
    drawable_label: &str,
    drawable_object_id: u64,
) -> Result<usize> {
    match direction {
        ChartSeriesDirection::Rows => Ok(data.row_names().len()),
        ChartSeriesDirection::Columns => Ok(data.column_names().len()),
        ChartSeriesDirection::Unsupported(value) => Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has unsupported series direction {value}"
        ))),
    }
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
