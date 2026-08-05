//! Native axis-label angle CRUD for Keynote slide charts.

use super::*;
use crate::charts::axis_label_angle::{
    chart_axis_label_angle as read_native_axis_label_angle,
    set_chart_axis_label_angle as set_native_axis_label_angle,
};
use crate::charts::{Axis, ChartAxisLabelAngle};

impl KeynoteEditor {
    /// Read one slide-chart axis' normalized label angle.
    pub fn slide_chart_axis_label_angle(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<ChartAxisLabelAngle> {
        slide_chart_axis_label_angle(self, slide_index, drawable_object_id, axis)
    }

    /// Set or reset one slide-chart axis' label angle.
    pub fn set_slide_chart_axis_label_angle(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: Axis,
        angle: ChartAxisLabelAngle,
    ) -> Result<()> {
        set_slide_chart_axis_label_angle(self, slide_index, drawable_object_id, axis, angle)
    }
}

fn slide_chart_axis_label_angle(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: Axis,
) -> Result<ChartAxisLabelAngle> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_axis_label_angle(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
    )
}

fn set_slide_chart_axis_label_angle(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: Axis,
    angle: ChartAxisLabelAngle,
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_axis_label_angle(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
        angle,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_axis_label_angle(slide_index, drawable_object_id, axis)? != angle {
        return Err(Error::InvalidFormat(
            "Keynote chart axis label-angle update failed validation".to_owned(),
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
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_presentation_supports_axis_label_angle_crud() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_slide_chart(
                0,
                ChartKind::Line2d,
                data(),
                DrawablePoint { x: 20.0, y: 20.0 },
                DrawableSize {
                    width: 400.0,
                    height: 300.0,
                },
            )
            .unwrap();
        let baseline = editor.to_bytes().unwrap();
        assert_eq!(
            editor
                .slide_chart_axis_label_angle(0, chart.drawable_object_id, Axis::Value)
                .unwrap(),
            ChartAxisLabelAngle::HORIZONTAL
        );
        let expected = ChartAxisLabelAngle::new(12.5).unwrap();
        editor
            .set_slide_chart_axis_label_angle(0, chart.drawable_object_id, Axis::Value, expected)
            .unwrap();
        let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_chart_axis_label_angle(0, chart.drawable_object_id, Axis::Value)
                .unwrap(),
            expected
        );
        reopened
            .set_slide_chart_axis_label_angle(
                0,
                chart.drawable_object_id,
                Axis::Value,
                ChartAxisLabelAngle::HORIZONTAL,
            )
            .unwrap();
        assert_eq!(reopened.to_bytes().unwrap(), baseline);
    }

    fn data() -> ChartData {
        ChartData::new(
            vec!["Series".to_owned()],
            vec!["A".to_owned(), "B".to_owned()],
            vec![vec![Some(1.0), Some(2.0)]],
        )
        .unwrap()
    }
}
