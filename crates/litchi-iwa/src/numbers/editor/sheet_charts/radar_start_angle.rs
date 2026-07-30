//! Native radar start-angle CRUD for Numbers sheet charts.

use super::*;
use crate::charts::ChartRadarStartAngle;
use crate::charts::radar_start_angle::{
    chart_radar_start_angle as read_native_chart_radar_start_angle,
    set_chart_radar_start_angle as set_native_chart_radar_start_angle,
};

impl NumbersEditor {
    /// Read one radar chart's `Rotation Angle`.
    pub fn sheet_chart_radar_start_angle(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ChartRadarStartAngle> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        read_native_chart_radar_start_angle(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            graph.info.kind,
        )
    }

    /// Set one radar chart's `Rotation Angle`.
    pub fn set_sheet_chart_radar_start_angle(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        angle: ChartRadarStartAngle,
    ) -> Result<()> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        if read_native_chart_radar_start_angle(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
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
            "Numbers",
            graph.info.kind,
            angle,
        )?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.sheet_chart_radar_start_angle(sheet_id, drawable_object_id)? != angle {
            return Err(Error::InvalidFormat(
                "Numbers radar chart start-angle update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::NumbersDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn duplicated_radar_charts_have_copy_on_write_start_angles() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let chart = editor
            .add_sheet_chart(
                sheet_id,
                ChartKind::Radar2d,
                data(),
                DrawablePoint { x: 20.0, y: 20.0 },
                DrawableSize {
                    width: 400.0,
                    height: 300.0,
                },
            )
            .unwrap();
        let source_angle = ChartRadarStartAngle::from_degrees(45.0).unwrap();
        editor
            .set_sheet_chart_radar_start_angle(sheet_id, chart.drawable_object_id, source_angle)
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, chart.drawable_object_id)
            .unwrap();
        let duplicate_angle = ChartRadarStartAngle::from_degrees(225.5).unwrap();
        editor
            .set_sheet_chart_radar_start_angle(
                sheet_id,
                duplicate.drawable_object_id,
                duplicate_angle,
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_radar_start_angle(sheet_id, chart.drawable_object_id)
                .unwrap(),
            source_angle
        );
        assert_eq!(
            editor
                .sheet_chart_radar_start_angle(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            duplicate_angle
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
