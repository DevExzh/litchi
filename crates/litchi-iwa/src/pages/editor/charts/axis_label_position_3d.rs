//! Native 3D value-axis label-position CRUD for Pages body charts.

use super::*;
use crate::charts::Chart3dAxisLabelPosition;
use crate::charts::axis_label_position_3d::{
    chart_3d_value_axis_label_position as read_native_3d_value_axis_label_position,
    set_chart_3d_value_axis_label_position as set_native_3d_value_axis_label_position,
};

impl PagesEditor {
    /// Read one body chart's 3D primary value-axis label position.
    pub fn body_chart_3d_value_axis_label_position(
        &self,
        drawable_object_id: u64,
    ) -> Result<Chart3dAxisLabelPosition> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        read_native_3d_value_axis_label_position(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            graph.info.kind,
        )
    }

    /// Set one body chart's 3D primary value-axis label position.
    pub fn set_body_chart_3d_value_axis_label_position(
        &mut self,
        drawable_object_id: u64,
        position: Chart3dAxisLabelPosition,
    ) -> Result<()> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        if read_native_3d_value_axis_label_position(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            graph.info.kind,
        )? == position
        {
            return Ok(());
        }
        let mut staged = self.package().clone();
        set_native_3d_value_axis_label_position(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            graph.info.kind,
            position,
        )?;
        let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.body_chart_3d_value_axis_label_position(drawable_object_id)? != position {
            return Err(Error::InvalidFormat(
                "Pages chart 3D value-axis label-position update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pages::PagesDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_document_supports_3d_axis_label_position_crud() {
        let mut editor = PagesDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_body_chart(
                0,
                ChartKind::Column3d,
                data(),
                DrawablePoint { x: 20.0, y: 20.0 },
                DrawableSize {
                    width: 400.0,
                    height: 300.0,
                },
            )
            .unwrap();
        assert_eq!(
            editor
                .body_chart_3d_value_axis_label_position(chart.drawable_object_id)
                .unwrap(),
            Chart3dAxisLabelPosition::Automatic
        );
        editor
            .set_body_chart_3d_value_axis_label_position(
                chart.drawable_object_id,
                Chart3dAxisLabelPosition::Trailing,
            )
            .unwrap();
        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body_chart_3d_value_axis_label_position(chart.drawable_object_id)
                .unwrap(),
            Chart3dAxisLabelPosition::Trailing
        );
    }

    #[test]
    fn charts_without_3d_value_axes_reject_label_position_access() {
        let mut editor = PagesDocumentBuilder::new().build().unwrap();
        for kind in [ChartKind::Column2d, ChartKind::Pie3d] {
            let chart = editor
                .add_body_chart(
                    0,
                    kind,
                    data(),
                    DrawablePoint { x: 20.0, y: 20.0 },
                    DrawableSize {
                        width: 400.0,
                        height: 300.0,
                    },
                )
                .unwrap();
            assert!(
                editor
                    .body_chart_3d_value_axis_label_position(chart.drawable_object_id)
                    .is_err()
            );
        }
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
