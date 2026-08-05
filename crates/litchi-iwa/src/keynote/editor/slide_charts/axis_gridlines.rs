//! Native chart-gridline visibility CRUD for Keynote slide charts.

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

impl KeynoteEditor {
    /// Read whether Keynote shows major gridlines for one native slide-chart axis.
    pub fn slide_chart_axis_major_gridlines_visible(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<bool> {
        slide_chart_axis_major_gridlines_visible(self, slide_index, drawable_object_id, axis)
    }

    /// Set whether Keynote shows major gridlines for one native slide-chart axis.
    pub fn set_slide_chart_axis_major_gridlines_visible(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: Axis,
        visible: bool,
    ) -> Result<()> {
        set_slide_chart_axis_major_gridlines_visible(
            self,
            slide_index,
            drawable_object_id,
            axis,
            visible,
        )
    }

    /// Read whether Keynote shows minor gridlines for one native slide-chart axis.
    pub fn slide_chart_axis_minor_gridlines_visible(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<bool> {
        slide_chart_axis_minor_gridlines_visible(self, slide_index, drawable_object_id, axis)
    }

    /// Set whether Keynote shows minor gridlines for one native slide-chart axis.
    pub fn set_slide_chart_axis_minor_gridlines_visible(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: Axis,
        visible: bool,
    ) -> Result<()> {
        set_slide_chart_axis_minor_gridlines_visible(
            self,
            slide_index,
            drawable_object_id,
            axis,
            visible,
        )
    }

    /// Read the exact inherited, empty, or explicit stroke for one gridline family.
    pub fn slide_chart_axis_gridline_stroke(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: Axis,
        gridline: ChartAxisGridline,
    ) -> Result<ChartAxisGridlineStroke> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        read_native_chart_axis_gridline_stroke(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            axis,
            gridline,
        )
    }

    /// Set one gridline family's stroke without changing its visibility.
    pub fn set_slide_chart_axis_gridline_stroke(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: Axis,
        gridline: ChartAxisGridline,
        stroke: ChartAxisGridlineStroke,
    ) -> Result<()> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_native_chart_axis_gridline_stroke(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            axis,
            gridline,
            stroke,
        )?;
        let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.slide_chart_axis_gridline_stroke(
            slide_index,
            drawable_object_id,
            axis,
            gridline,
        )? != stroke
        {
            return Err(Error::InvalidFormat(
                "Keynote chart axis gridline-stroke update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

fn slide_chart_axis_major_gridlines_visible(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: Axis,
) -> Result<bool> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_chart_axis_major_gridlines_visible(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
    )
}

fn set_slide_chart_axis_major_gridlines_visible(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: Axis,
    visible: bool,
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_axis_major_gridlines_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
        visible,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_axis_major_gridlines_visible(slide_index, drawable_object_id, axis)?
        != visible
    {
        return Err(Error::InvalidFormat(
            "Keynote chart axis major-gridline update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn slide_chart_axis_minor_gridlines_visible(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: Axis,
) -> Result<bool> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_chart_axis_minor_gridlines_visible(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
    )
}

fn set_slide_chart_axis_minor_gridlines_visible(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: Axis,
    visible: bool,
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_axis_minor_gridlines_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
        visible,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_axis_minor_gridlines_visible(slide_index, drawable_object_id, axis)?
        != visible
    {
        return Err(Error::InvalidFormat(
            "Keynote chart axis minor-gridline update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{ChartData, ChartKind};
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::shapes::{
        DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeStroke, StrokePattern,
        StrokeWidth,
    };

    #[test]
    fn scratch_presentation_supports_axis_gridline_stroke_crud() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_slide_chart(
                0,
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

        assert_eq!(
            editor
                .slide_chart_axis_gridline_stroke(
                    0,
                    chart.drawable_object_id,
                    Axis::Category,
                    ChartAxisGridline::Major,
                )
                .unwrap(),
            ChartAxisGridlineStroke::Inherited
        );
        editor
            .set_slide_chart_axis_gridline_stroke(
                0,
                chart.drawable_object_id,
                Axis::Category,
                ChartAxisGridline::Major,
                stroke,
            )
            .unwrap();
        assert_eq!(
            editor
                .slide_chart_axis_gridline_stroke(
                    0,
                    chart.drawable_object_id,
                    Axis::Category,
                    ChartAxisGridline::Major,
                )
                .unwrap(),
            stroke
        );
        assert!(
            !editor
                .slide_chart_axis_major_gridlines_visible(
                    0,
                    chart.drawable_object_id,
                    Axis::Category,
                )
                .unwrap()
        );
        editor
            .set_slide_chart_axis_gridline_stroke(
                0,
                chart.drawable_object_id,
                Axis::Category,
                ChartAxisGridline::Major,
                ChartAxisGridlineStroke::NoStroke,
            )
            .unwrap();
        editor
            .set_slide_chart_axis_gridline_stroke(
                0,
                chart.drawable_object_id,
                Axis::Category,
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
            StrokeWidth::new(3.0).unwrap(),
            StrokePattern::MediumDash,
        )
    }
}
