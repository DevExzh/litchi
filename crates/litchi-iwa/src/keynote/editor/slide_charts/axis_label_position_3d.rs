//! Native 3D value-axis label-position CRUD for Keynote slide charts.

use super::*;
use crate::charts::LabelPosition3d;
use crate::charts::axis_label_position_3d::{
    chart_3d_value_axis_label_position as read_native_3d_value_axis_label_position,
    set_chart_3d_value_axis_label_position as set_native_3d_value_axis_label_position,
};

impl KeynoteEditor {
    /// Read one slide chart's 3D primary value-axis label position.
    pub fn slide_chart_3d_value_axis_label_position(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<LabelPosition3d> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        read_native_3d_value_axis_label_position(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            graph.info.kind,
        )
    }

    /// Set one slide chart's 3D primary value-axis label position.
    pub fn set_slide_chart_3d_value_axis_label_position(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        position: LabelPosition3d,
    ) -> Result<()> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        if read_native_3d_value_axis_label_position(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
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
            "Keynote",
            graph.info.kind,
            position,
        )?;
        let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.slide_chart_3d_value_axis_label_position(slide_index, drawable_object_id)?
            != position
        {
            return Err(Error::InvalidFormat(
                "Keynote chart 3D value-axis label-position update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_presentation_supports_3d_axis_label_position_crud() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_slide_chart(
                0,
                ChartKind::Area3d,
                data(),
                DrawablePoint { x: 20.0, y: 20.0 },
                DrawableSize {
                    width: 400.0,
                    height: 300.0,
                },
            )
            .unwrap();
        editor
            .set_slide_chart_3d_value_axis_label_position(
                0,
                chart.drawable_object_id,
                LabelPosition3d::Leading,
            )
            .unwrap();
        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_chart_3d_value_axis_label_position(0, chart.drawable_object_id)
                .unwrap(),
            LabelPosition3d::Leading
        );
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
