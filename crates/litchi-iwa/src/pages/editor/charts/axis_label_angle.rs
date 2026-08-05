//! Native axis-label angle CRUD for Pages body charts.

use super::*;
use crate::charts::axis_label_angle::{
    chart_axis_label_angle as read_native_axis_label_angle,
    set_chart_axis_label_angle as set_native_axis_label_angle,
};
use crate::charts::{Axis, LabelAngle};

impl PagesEditor {
    /// Read one body-chart axis' normalized label angle.
    pub fn body_chart_axis_label_angle(
        &self,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<LabelAngle> {
        body_chart_axis_label_angle(self, drawable_object_id, axis)
    }

    /// Set or reset one body-chart axis' label angle.
    pub fn set_body_chart_axis_label_angle(
        &mut self,
        drawable_object_id: u64,
        axis: Axis,
        angle: LabelAngle,
    ) -> Result<()> {
        set_body_chart_axis_label_angle(self, drawable_object_id, axis, angle)
    }
}

fn body_chart_axis_label_angle(
    editor: &PagesEditor,
    drawable_object_id: u64,
    axis: Axis,
) -> Result<LabelAngle> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_axis_label_angle(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        axis,
    )
}

fn set_body_chart_axis_label_angle(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    axis: Axis,
    angle: LabelAngle,
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_axis_label_angle(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        axis,
        angle,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_axis_label_angle(drawable_object_id, axis)? != angle {
        return Err(Error::InvalidFormat(
            "Pages chart axis label-angle update failed validation".to_owned(),
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
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_document_supports_axis_label_angle_crud() {
        let body = "Axis label angle";
        let mut editor = PagesDocumentBuilder::new().body_text(body).build().unwrap();
        let chart = editor
            .add_body_chart(
                body.encode_utf16().count(),
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
                .body_chart_axis_label_angle(chart.drawable_object_id, Axis::Value)
                .unwrap(),
            LabelAngle::HORIZONTAL
        );
        let expected = LabelAngle::new(12.5).unwrap();
        editor
            .set_body_chart_axis_label_angle(chart.drawable_object_id, Axis::Value, expected)
            .unwrap();
        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body_chart_axis_label_angle(chart.drawable_object_id, Axis::Value)
                .unwrap(),
            expected
        );
        reopened
            .set_body_chart_axis_label_angle(
                chart.drawable_object_id,
                Axis::Value,
                LabelAngle::HORIZONTAL,
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
