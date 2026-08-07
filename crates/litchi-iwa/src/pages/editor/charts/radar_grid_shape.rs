//! Native radar grid-shape CRUD for Pages body charts.

use super::*;
use crate::charts::ChartRadarGridShape;
use crate::charts::radar_grid_shape::{
    chart_radar_grid_shape as read_native_chart_radar_grid_shape,
    set_chart_radar_grid_shape as set_native_chart_radar_grid_shape,
};

impl PagesEditor {
    /// Read one radar chart's `Radar Chart Grid Shape`.
    pub fn body_chart_radar_grid_shape(
        &self,
        drawable_object_id: u64,
    ) -> Result<ChartRadarGridShape> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        read_native_chart_radar_grid_shape(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            graph.info.kind,
        )
    }

    /// Set one radar chart's `Radar Chart Grid Shape`.
    pub fn set_body_chart_radar_grid_shape(
        &mut self,
        drawable_object_id: u64,
        shape: ChartRadarGridShape,
    ) -> Result<()> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        if read_native_chart_radar_grid_shape(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
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
            "Pages",
            graph.info.kind,
            shape,
        )?;
        let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.body_chart_radar_grid_shape(drawable_object_id)? != shape {
            return Err(Error::InvalidFormat(
                "Pages radar chart grid-shape update failed validation".to_owned(),
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
    fn scratch_document_supports_radar_grid_shape_crud() {
        let mut editor = PagesDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_body_chart(
                0,
                Kind::Radar2d,
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
                .body_chart_radar_grid_shape(chart.drawable_object_id)
                .unwrap(),
            ChartRadarGridShape::Straight
        );
        editor
            .set_body_chart_radar_grid_shape(chart.drawable_object_id, ChartRadarGridShape::Curved)
            .unwrap();
        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body_chart_radar_grid_shape(chart.drawable_object_id)
                .unwrap(),
            ChartRadarGridShape::Curved
        );
    }

    #[test]
    fn non_radar_chart_rejects_grid_shape_access() {
        let mut editor = PagesDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_body_chart(
                0,
                Kind::Line2d,
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
                .body_chart_radar_grid_shape(chart.drawable_object_id)
                .is_err()
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
