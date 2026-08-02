//! Native radar start-angle CRUD for Pages body charts.

use super::*;
use crate::charts::ChartRadarStartAngle;
use crate::charts::radar_start_angle::{
    chart_radar_start_angle as read_native_chart_radar_start_angle,
    set_chart_radar_start_angle as set_native_chart_radar_start_angle,
};

impl PagesEditor {
    /// Read one radar chart's `Rotation Angle`.
    pub fn body_chart_radar_start_angle(
        &self,
        drawable_object_id: u64,
    ) -> Result<ChartRadarStartAngle> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        read_native_chart_radar_start_angle(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            graph.info.kind,
        )
    }

    /// Set one radar chart's `Rotation Angle`.
    pub fn set_body_chart_radar_start_angle(
        &mut self,
        drawable_object_id: u64,
        angle: ChartRadarStartAngle,
    ) -> Result<()> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        if read_native_chart_radar_start_angle(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            graph.info.kind,
        )? == angle
        {
            return Ok(());
        }
        let mut staged = self.package().clone();
        set_native_chart_radar_start_angle(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            graph.info.kind,
            angle,
        )?;
        let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.body_chart_radar_start_angle(drawable_object_id)? != angle {
            return Err(Error::InvalidFormat(
                "Pages radar chart start-angle update failed validation".to_owned(),
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
    fn scratch_document_supports_radar_start_angle_crud() {
        let mut editor = PagesDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_body_chart(
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
        assert_eq!(
            editor
                .body_chart_radar_start_angle(chart.drawable_object_id)
                .unwrap(),
            ChartRadarStartAngle::ZERO
        );
        let angle = ChartRadarStartAngle::from_degrees(45.0).unwrap();
        editor
            .set_body_chart_radar_start_angle(chart.drawable_object_id, angle)
            .unwrap();
        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body_chart_radar_start_angle(chart.drawable_object_id)
                .unwrap(),
            angle
        );
    }

    #[test]
    fn non_radar_chart_rejects_start_angle_access() {
        let mut editor = PagesDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_body_chart(
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
        assert!(
            editor
                .body_chart_radar_start_angle(chart.drawable_object_id)
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
