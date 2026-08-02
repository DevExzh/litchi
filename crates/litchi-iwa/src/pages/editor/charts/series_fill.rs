//! Inherited per-series fill CRUD for Pages charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::ChartSeriesIndex;
use crate::charts::series_fill::{
    chart_series_fills as read_native_fills, reset_chart_series_fill as reset_native_fill,
    set_chart_series_fills as set_native_fills,
    set_chart_series_image_fill_data as set_native_image_fill_data,
};
use crate::shapes::{RgbaColor, ShapeFill, ShapeImageFill, ShapeImageFillTechnique};

impl PagesEditor {
    /// Read effective fills in native series order.
    pub fn body_chart_series_fills(&self, drawable_object_id: u64) -> Result<Vec<ShapeFill>> {
        body_chart_series_fills(self, drawable_object_id)
    }

    /// Read one effective series fill.
    pub fn body_chart_series_fill(
        &self,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<ShapeFill> {
        let fills = body_chart_series_fills(self, drawable_object_id)?;
        fills.get(series.zero_based()).cloned().ok_or_else(|| {
            series_fill_index_error("Pages", drawable_object_id, series, fills.len())
        })
    }

    /// Replace every series fill transactionally.
    pub fn set_body_chart_series_fills(
        &mut self,
        drawable_object_id: u64,
        fills: &[ShapeFill],
    ) -> Result<()> {
        set_body_chart_series_fills(self, drawable_object_id, fills)
    }

    /// Replace one series fill transactionally.
    pub fn set_body_chart_series_fill(
        &mut self,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        fill: &ShapeFill,
    ) -> Result<()> {
        let mut fills = body_chart_series_fills(self, drawable_object_id)?;
        let count = fills.len();
        let target = fills
            .get_mut(series.zero_based())
            .ok_or_else(|| series_fill_index_error("Pages", drawable_object_id, series, count))?;
        if target == fill {
            return Ok(());
        }
        *target = fill.clone();
        set_body_chart_series_fills(self, drawable_object_id, &fills)
    }

    /// Remove one local override and reveal its inherited series fill.
    pub fn reset_body_chart_series_fill(
        &mut self,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<ShapeFill> {
        reset_body_chart_series_fill(self, drawable_object_id, series)
    }

    /// Embed image bytes and assign them to one series.
    #[allow(clippy::too_many_arguments)]
    pub fn set_body_chart_series_image_fill(
        &mut self,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        preferred_filename: &str,
        data: &[u8],
        technique: ShapeImageFillTechnique,
        tint: Option<RgbaColor>,
    ) -> Result<ShapeImageFill> {
        set_body_chart_series_image_fill(
            self,
            drawable_object_id,
            series,
            preferred_filename,
            data,
            technique,
            tint,
        )
    }
}

fn body_chart_series_fills(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<Vec<ShapeFill>> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    read_native_fills(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        graph.info.kind,
        series_count,
    )
}

fn set_body_chart_series_fills(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    fills: &[ShapeFill],
) -> Result<()> {
    if body_chart_series_fills(editor, drawable_object_id)? == fills {
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
    set_native_fills(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        graph.info.kind,
        series_count,
        fills,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_series_fills(drawable_object_id)? != fills {
        return Err(Error::InvalidFormat(
            "Pages chart series fill update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn reset_body_chart_series_fill(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    series: ChartSeriesIndex,
) -> Result<ShapeFill> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    let mut staged = editor.package().clone();
    let inherited = reset_native_fill(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        graph.info.kind,
        series_count,
        series.zero_based(),
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_series_fill(drawable_object_id, series)? != inherited {
        return Err(Error::InvalidFormat(
            "Pages chart series fill reset failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(inherited)
}

#[allow(clippy::too_many_arguments)]
fn set_body_chart_series_image_fill(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    series: ChartSeriesIndex,
    preferred_filename: &str,
    data: &[u8],
    technique: ShapeImageFillTechnique,
    tint: Option<RgbaColor>,
) -> Result<ShapeImageFill> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let series_count = value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Pages",
        drawable_object_id,
    )?;
    let fill_size = graph.info.geometry.size.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} has no series image-fill dimensions"
        ))
    })?;
    let mut staged = editor.package().clone();
    let image = set_native_image_fill_data(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        graph.info.kind,
        series_count,
        series.zero_based(),
        preferred_filename,
        data,
        technique,
        fill_size,
        tint,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_series_fill(drawable_object_id, series)?
        != ShapeFill::Image(image.clone())
    {
        return Err(Error::InvalidFormat(
            "Pages chart series image-fill update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(image)
}

fn series_fill_index_error(
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
