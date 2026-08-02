//! Native chart-gridline visibility CRUD for Pages body charts.

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
use crate::charts::{ChartAxis, ChartAxisGridline, ChartAxisGridlineStroke};

impl PagesEditor {
    /// Read whether Pages shows major gridlines for one native body-chart axis.
    pub fn body_chart_axis_major_gridlines_visible(
        &self,
        drawable_object_id: u64,
        axis: ChartAxis,
    ) -> Result<bool> {
        body_chart_axis_major_gridlines_visible(self, drawable_object_id, axis)
    }

    /// Set whether Pages shows major gridlines for one native body-chart axis.
    pub fn set_body_chart_axis_major_gridlines_visible(
        &mut self,
        drawable_object_id: u64,
        axis: ChartAxis,
        visible: bool,
    ) -> Result<()> {
        set_body_chart_axis_major_gridlines_visible(self, drawable_object_id, axis, visible)
    }

    /// Read whether Pages shows minor gridlines for one native body-chart axis.
    pub fn body_chart_axis_minor_gridlines_visible(
        &self,
        drawable_object_id: u64,
        axis: ChartAxis,
    ) -> Result<bool> {
        body_chart_axis_minor_gridlines_visible(self, drawable_object_id, axis)
    }

    /// Set whether Pages shows minor gridlines for one native body-chart axis.
    pub fn set_body_chart_axis_minor_gridlines_visible(
        &mut self,
        drawable_object_id: u64,
        axis: ChartAxis,
        visible: bool,
    ) -> Result<()> {
        set_body_chart_axis_minor_gridlines_visible(self, drawable_object_id, axis, visible)
    }

    /// Read the exact inherited, empty, or explicit stroke for one gridline family.
    pub fn body_chart_axis_gridline_stroke(
        &self,
        drawable_object_id: u64,
        axis: ChartAxis,
        gridline: ChartAxisGridline,
    ) -> Result<ChartAxisGridlineStroke> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        read_native_chart_axis_gridline_stroke(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            axis,
            gridline,
        )
    }

    /// Set one gridline family's stroke without changing its visibility.
    pub fn set_body_chart_axis_gridline_stroke(
        &mut self,
        drawable_object_id: u64,
        axis: ChartAxis,
        gridline: ChartAxisGridline,
        stroke: ChartAxisGridlineStroke,
    ) -> Result<()> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_native_chart_axis_gridline_stroke(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            axis,
            gridline,
            stroke,
        )?;
        let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.body_chart_axis_gridline_stroke(drawable_object_id, axis, gridline)? != stroke {
            return Err(Error::InvalidFormat(
                "Pages chart axis gridline-stroke update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

fn body_chart_axis_major_gridlines_visible(
    editor: &PagesEditor,
    drawable_object_id: u64,
    axis: ChartAxis,
) -> Result<bool> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_chart_axis_major_gridlines_visible(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        axis,
    )
}

fn set_body_chart_axis_major_gridlines_visible(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    axis: ChartAxis,
    visible: bool,
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_axis_major_gridlines_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        axis,
        visible,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_axis_major_gridlines_visible(drawable_object_id, axis)? != visible {
        return Err(Error::InvalidFormat(
            "Pages chart axis major-gridline update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn body_chart_axis_minor_gridlines_visible(
    editor: &PagesEditor,
    drawable_object_id: u64,
    axis: ChartAxis,
) -> Result<bool> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_chart_axis_minor_gridlines_visible(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        axis,
    )
}

fn set_body_chart_axis_minor_gridlines_visible(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    axis: ChartAxis,
    visible: bool,
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_axis_minor_gridlines_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        axis,
        visible,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_axis_minor_gridlines_visible(drawable_object_id, axis)? != visible {
        return Err(Error::InvalidFormat(
            "Pages chart axis minor-gridline update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{ChartData, ChartKind};
    use crate::pages::PagesDocumentBuilder;
    use crate::shapes::{
        DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeStroke, StrokePattern,
        StrokeWidth,
    };

    #[test]
    fn scratch_document_supports_axis_gridline_stroke_crud() {
        let body = "Gridline stroke";
        let mut editor = PagesDocumentBuilder::new().body_text(body).build().unwrap();
        let chart = editor
            .add_body_chart(
                body.encode_utf16().count(),
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
            .set_body_chart_axis_gridline_stroke(
                chart.drawable_object_id,
                ChartAxis::Value,
                ChartAxisGridline::Minor,
                stroke,
            )
            .unwrap();
        assert_eq!(
            editor
                .body_chart_axis_gridline_stroke(
                    chart.drawable_object_id,
                    ChartAxis::Value,
                    ChartAxisGridline::Minor,
                )
                .unwrap(),
            stroke
        );
        assert!(
            !editor
                .body_chart_axis_minor_gridlines_visible(
                    chart.drawable_object_id,
                    ChartAxis::Value,
                )
                .unwrap()
        );
        editor
            .set_body_chart_axis_gridline_stroke(
                chart.drawable_object_id,
                ChartAxis::Value,
                ChartAxisGridline::Minor,
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
            StrokeWidth::new(2.0).unwrap(),
            StrokePattern::RoundedDash,
        )
    }
}
