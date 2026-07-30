//! Native radar grid-shape CRUD for Numbers sheet charts.

use super::*;
use crate::charts::ChartRadarGridShape;
use crate::charts::radar_grid_shape::{
    chart_radar_grid_shape as read_native_chart_radar_grid_shape,
    set_chart_radar_grid_shape as set_native_chart_radar_grid_shape,
};

impl NumbersEditor {
    /// Read one radar chart's `Radar Chart Grid Shape`.
    pub fn sheet_chart_radar_grid_shape(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ChartRadarGridShape> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        read_native_chart_radar_grid_shape(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            graph.info.kind,
        )
    }

    /// Set one radar chart's `Radar Chart Grid Shape`.
    pub fn set_sheet_chart_radar_grid_shape(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        shape: ChartRadarGridShape,
    ) -> Result<()> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        if read_native_chart_radar_grid_shape(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
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
            "Numbers",
            graph.info.kind,
            shape,
        )?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.sheet_chart_radar_grid_shape(sheet_id, drawable_object_id)? != shape {
            return Err(Error::InvalidFormat(
                "Numbers radar chart grid-shape update failed validation".to_owned(),
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
    fn duplicated_radar_charts_have_copy_on_write_grid_shapes() {
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
        editor
            .set_sheet_chart_radar_grid_shape(
                sheet_id,
                chart.drawable_object_id,
                ChartRadarGridShape::Curved,
            )
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, chart.drawable_object_id)
            .unwrap();
        editor
            .set_sheet_chart_radar_grid_shape(
                sheet_id,
                duplicate.drawable_object_id,
                ChartRadarGridShape::Straight,
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_radar_grid_shape(sheet_id, chart.drawable_object_id)
                .unwrap(),
            ChartRadarGridShape::Curved
        );
        assert_eq!(
            editor
                .sheet_chart_radar_grid_shape(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            ChartRadarGridShape::Straight
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
