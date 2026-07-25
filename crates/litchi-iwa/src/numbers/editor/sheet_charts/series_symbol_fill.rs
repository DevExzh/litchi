//! Inherited data-symbol fill CRUD for Numbers charts.

use super::graph::SheetChartGraph;
use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::series_symbol_fill::{
    chart_series_symbol_fills as read_native, reset_chart_series_symbol_fill as reset_native,
    set_chart_series_symbol_fills as set_native,
    set_chart_series_symbol_image_fill_data as set_native_image,
};
use crate::charts::{ChartSeriesIndex, ChartSeriesSymbolFill};
use crate::shapes::{RgbaColor, ShapeImageFill, ShapeImageFillTechnique};

impl NumbersEditor {
    pub fn sheet_chart_series_symbol_fills(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartSeriesSymbolFill>> {
        read(self, sheet_id, drawable_object_id)
    }

    pub fn sheet_chart_series_symbol_fill(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<ChartSeriesSymbolFill> {
        let values = read(self, sheet_id, drawable_object_id)?;
        values
            .get(series.zero_based())
            .cloned()
            .ok_or_else(|| index_error("Numbers", drawable_object_id, series, values.len()))
    }

    pub fn set_sheet_chart_series_symbol_fills(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        fills: &[ChartSeriesSymbolFill],
    ) -> Result<()> {
        set(self, sheet_id, drawable_object_id, fills)
    }

    pub fn set_sheet_chart_series_symbol_fill(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        fill: ChartSeriesSymbolFill,
    ) -> Result<()> {
        let mut values = read(self, sheet_id, drawable_object_id)?;
        let count = values.len();
        let target = values
            .get_mut(series.zero_based())
            .ok_or_else(|| index_error("Numbers", drawable_object_id, series, count))?;
        if *target == fill {
            return Ok(());
        }
        *target = fill;
        set(self, sheet_id, drawable_object_id, &values)
    }

    pub fn reset_sheet_chart_series_symbol_fill(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
    ) -> Result<ChartSeriesSymbolFill> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        let count = series_count(&graph, drawable_object_id)?;
        let mut staged = self.package().clone();
        let inherited = reset_native(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            graph.info.kind,
            count,
            series.zero_based(),
        )?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.sheet_chart_series_symbol_fill(sheet_id, drawable_object_id, series)?
            != inherited
        {
            return Err(Error::InvalidFormat(
                "Numbers chart data-symbol fill reset failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(inherited)
    }

    #[allow(clippy::too_many_arguments)]
    pub fn set_sheet_chart_series_symbol_image_fill(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        series: ChartSeriesIndex,
        preferred_filename: &str,
        data: &[u8],
        technique: ShapeImageFillTechnique,
        tint: Option<RgbaColor>,
    ) -> Result<ShapeImageFill> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        let count = series_count(&graph, drawable_object_id)?;
        let fill_size = graph.info.geometry.size.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers chart {drawable_object_id} has no data-symbol image-fill dimensions"
            ))
        })?;
        let mut staged = self.package().clone();
        let image = set_native_image(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            graph.info.kind,
            count,
            series.zero_based(),
            preferred_filename,
            data,
            technique,
            fill_size,
            tint,
        )?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.sheet_chart_series_symbol_fill(sheet_id, drawable_object_id, series)?
            != ChartSeriesSymbolFill::Custom(crate::shapes::ShapeFill::Image(image.clone()))
        {
            return Err(Error::InvalidFormat(
                "Numbers chart data-symbol image-fill update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(image)
    }
}

fn read(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<Vec<ChartSeriesSymbolFill>> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        graph.info.kind,
        series_count(&graph, drawable_object_id)?,
    )
}

fn set(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    fills: &[ChartSeriesSymbolFill],
) -> Result<()> {
    if read(editor, sheet_id, drawable_object_id)? == fills {
        return Ok(());
    }
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        graph.info.kind,
        series_count(&graph, drawable_object_id)?,
        fills,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_series_symbol_fills(sheet_id, drawable_object_id)? != fills {
        return Err(Error::InvalidFormat(
            "Numbers chart data-symbol fill update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn series_count(graph: &SheetChartGraph, drawable_object_id: u64) -> Result<usize> {
    value_label_series_count(
        graph.info.kind,
        graph.info.direction,
        &graph.info.data,
        "Numbers",
        drawable_object_id,
    )
}

fn index_error(
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
