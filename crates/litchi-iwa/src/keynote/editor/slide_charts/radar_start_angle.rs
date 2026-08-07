//! Native radar start-angle CRUD for Keynote slide charts.

use super::*;
use crate::charts::ChartRadarStartAngle;
use crate::charts::radar_start_angle::{
    chart_radar_start_angle as read_native_chart_radar_start_angle,
    set_chart_radar_start_angle as set_native_chart_radar_start_angle,
};

impl KeynoteEditor {
    /// Read one radar chart's `Rotation Angle`.
    pub fn slide_chart_radar_start_angle(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ChartRadarStartAngle> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        read_native_chart_radar_start_angle(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            graph.info.kind,
        )
    }

    /// Set one radar chart's `Rotation Angle`.
    pub fn set_slide_chart_radar_start_angle(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        angle: ChartRadarStartAngle,
    ) -> Result<()> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        if read_native_chart_radar_start_angle(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
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
            "Keynote",
            graph.info.kind,
            angle,
        )?;
        let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.slide_chart_radar_start_angle(slide_index, drawable_object_id)? != angle {
            return Err(Error::InvalidFormat(
                "Keynote radar chart start-angle update failed validation".to_owned(),
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
    fn scratch_presentation_supports_radar_start_angle_crud() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_slide_chart(
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
        let angle = ChartRadarStartAngle::from_degrees(315.5).unwrap();
        editor
            .set_slide_chart_radar_start_angle(0, chart.drawable_object_id, angle)
            .unwrap();
        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_chart_radar_start_angle(0, chart.drawable_object_id)
                .unwrap(),
            angle
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
