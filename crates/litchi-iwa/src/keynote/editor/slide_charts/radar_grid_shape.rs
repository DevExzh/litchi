//! Native radar grid-shape CRUD for Keynote slide charts.

use super::*;
use crate::charts::ChartRadarGridShape;
use crate::charts::radar_grid_shape::{
    chart_radar_grid_shape as read_native_chart_radar_grid_shape,
    set_chart_radar_grid_shape as set_native_chart_radar_grid_shape,
};

impl KeynoteEditor {
    /// Read one radar chart's `Radar Chart Grid Shape`.
    pub fn slide_chart_radar_grid_shape(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ChartRadarGridShape> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        read_native_chart_radar_grid_shape(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            graph.info.kind,
        )
    }

    /// Set one radar chart's `Radar Chart Grid Shape`.
    pub fn set_slide_chart_radar_grid_shape(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        shape: ChartRadarGridShape,
    ) -> Result<()> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        if read_native_chart_radar_grid_shape(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            graph.info.kind,
        )? == shape
        {
            return Ok(());
        }
        let mut staged = self.package().clone();
        set_native_chart_radar_grid_shape(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            graph.info.kind,
            shape,
        )?;
        let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.slide_chart_radar_grid_shape(slide_index, drawable_object_id)? != shape {
            return Err(Error::InvalidFormat(
                "Keynote radar chart grid-shape update failed validation".to_owned(),
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
    fn scratch_presentation_supports_radar_grid_shape_crud() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_slide_chart(
                0,
                ChartKind::Radar2d,
                data(),
                DrawablePoint { x: 20.0, y: 20.0 },
                DrawableSize {
                    width: 400.0,
                    height: 300.0,
                },
            )
            .unwrap();
        editor
            .set_slide_chart_radar_grid_shape(
                0,
                chart.drawable_object_id,
                ChartRadarGridShape::Curved,
            )
            .unwrap();
        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_chart_radar_grid_shape(0, chart.drawable_object_id)
                .unwrap(),
            ChartRadarGridShape::Curved
        );
    }

    fn data() -> ChartData {
        ChartData::new(
            vec!["North".to_owned(), "South".to_owned()],
            vec!["Q1".to_owned(), "Q2".to_owned()],
            vec![vec![Some(12.0), Some(18.0)], vec![Some(9.0), Some(21.0)]],
        )
        .unwrap()
    }
}
