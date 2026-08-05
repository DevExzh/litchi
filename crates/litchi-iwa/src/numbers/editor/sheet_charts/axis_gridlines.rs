//! Native chart-gridline visibility CRUD for Numbers sheet charts.

use super::*;
use crate::charts::axis_gridline_stroke::{
    chart_axis_gridline_stroke as read_native_chart_axis_gridline_stroke,
    set_chart_axis_gridline_stroke as set_native_chart_axis_gridline_stroke,
};
use crate::charts::axis_style::{
    chart_axis_major_gridlines_visible as read_native_chart_axis_major_gridlines_visible,
    chart_axis_minor_gridlines_visible as read_native_chart_axis_minor_gridlines_visible,
    set_chart_axis_major_gridlines_visible as set_native_chart_axis_major_gridlines_visible,
    set_chart_axis_minor_gridlines_visible as set_native_chart_axis_minor_gridlines_visible,
};
use crate::charts::{Axis, ChartAxisGridline, ChartAxisGridlineStroke};

impl NumbersEditor {
    /// Read whether Numbers shows major gridlines for one native sheet-chart axis.
    pub fn sheet_chart_axis_major_gridlines_visible(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<bool> {
        sheet_chart_axis_major_gridlines_visible(self, sheet_id, drawable_object_id, axis)
    }

    /// Set whether Numbers shows major gridlines for one native sheet-chart axis.
    pub fn set_sheet_chart_axis_major_gridlines_visible(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: Axis,
        visible: bool,
    ) -> Result<()> {
        set_sheet_chart_axis_major_gridlines_visible(
            self,
            sheet_id,
            drawable_object_id,
            axis,
            visible,
        )
    }

    /// Read whether Numbers shows minor gridlines for one native sheet-chart axis.
    pub fn sheet_chart_axis_minor_gridlines_visible(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<bool> {
        sheet_chart_axis_minor_gridlines_visible(self, sheet_id, drawable_object_id, axis)
    }

    /// Set whether Numbers shows minor gridlines for one native sheet-chart axis.
    pub fn set_sheet_chart_axis_minor_gridlines_visible(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: Axis,
        visible: bool,
    ) -> Result<()> {
        set_sheet_chart_axis_minor_gridlines_visible(
            self,
            sheet_id,
            drawable_object_id,
            axis,
            visible,
        )
    }

    /// Read the exact inherited, empty, or explicit stroke for one gridline family.
    pub fn sheet_chart_axis_gridline_stroke(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: Axis,
        gridline: ChartAxisGridline,
    ) -> Result<ChartAxisGridlineStroke> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        read_native_chart_axis_gridline_stroke(
            &self.package,
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            axis,
            gridline,
        )
    }

    /// Set one gridline family's stroke without changing its visibility.
    pub fn set_sheet_chart_axis_gridline_stroke(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: Axis,
        gridline: ChartAxisGridline,
        stroke: ChartAxisGridlineStroke,
    ) -> Result<()> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_native_chart_axis_gridline_stroke(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            axis,
            gridline,
            stroke,
        )?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.sheet_chart_axis_gridline_stroke(
            sheet_id,
            drawable_object_id,
            axis,
            gridline,
        )? != stroke
        {
            return Err(Error::InvalidFormat(
                "Numbers chart axis gridline-stroke update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

fn sheet_chart_axis_major_gridlines_visible(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    axis: Axis,
) -> Result<bool> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_chart_axis_major_gridlines_visible(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        axis,
    )
}

fn set_sheet_chart_axis_major_gridlines_visible(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    axis: Axis,
    visible: bool,
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_chart_axis_major_gridlines_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        axis,
        visible,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_axis_major_gridlines_visible(sheet_id, drawable_object_id, axis)?
        != visible
    {
        return Err(Error::InvalidFormat(
            "Numbers chart axis major-gridline update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn sheet_chart_axis_minor_gridlines_visible(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    axis: Axis,
) -> Result<bool> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_chart_axis_minor_gridlines_visible(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        axis,
    )
}

fn set_sheet_chart_axis_minor_gridlines_visible(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    axis: Axis,
    visible: bool,
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_chart_axis_minor_gridlines_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        axis,
        visible,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_axis_minor_gridlines_visible(sheet_id, drawable_object_id, axis)?
        != visible
    {
        return Err(Error::InvalidFormat(
            "Numbers chart axis minor-gridline update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{ChartData, ChartKind};
    use crate::numbers::NumbersDocumentBuilder;
    use crate::shapes::{
        DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeStroke, StrokePattern,
        StrokeWidth,
    };

    #[test]
    fn scratch_spreadsheet_supports_axis_gridline_stroke_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let chart = editor
            .add_sheet_chart(
                sheet_id,
                ChartKind::Column2d,
                data(),
                DrawablePoint { x: 20.0, y: 20.0 },
                DrawableSize {
                    width: 400.0,
                    height: 300.0,
                },
            )
            .unwrap();
        let baseline = editor.to_bytes().unwrap();
        let stroke = ChartAxisGridlineStroke::Stroke(test_stroke());

        editor
            .set_sheet_chart_axis_gridline_stroke(
                sheet_id,
                chart.drawable_object_id,
                Axis::Value,
                ChartAxisGridline::Major,
                stroke,
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_axis_gridline_stroke(
                    sheet_id,
                    chart.drawable_object_id,
                    Axis::Value,
                    ChartAxisGridline::Major,
                )
                .unwrap(),
            stroke
        );
        assert!(
            editor
                .sheet_chart_axis_major_gridlines_visible(
                    sheet_id,
                    chart.drawable_object_id,
                    Axis::Value,
                )
                .unwrap()
        );
        editor
            .set_sheet_chart_axis_gridline_stroke(
                sheet_id,
                chart.drawable_object_id,
                Axis::Value,
                ChartAxisGridline::Major,
                ChartAxisGridlineStroke::Inherited,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    fn data() -> ChartData {
        ChartData::new(
            vec!["Series".to_owned()],
            vec!["A".to_owned(), "B".to_owned()],
            vec![vec![Some(8.0), Some(20.0)]],
        )
        .unwrap()
    }

    fn test_stroke() -> ShapeStroke {
        ShapeStroke::new(
            RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
            StrokeWidth::new(2.5).unwrap(),
            StrokePattern::LongDash,
        )
    }
}
