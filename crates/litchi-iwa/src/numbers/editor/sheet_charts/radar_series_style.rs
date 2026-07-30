//! Native chart-wide Radar series-style CRUD for Numbers sheet charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::ChartRadarSeriesStyle;
use crate::charts::radar_series_style::{
    chart_radar_series_style as read_native_radar_style,
    set_chart_radar_series_style as set_native_radar_style,
};

impl NumbersEditor {
    /// Read the chart-wide Radar `Style` selection.
    pub fn sheet_chart_radar_series_style(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ChartRadarSeriesStyle> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        let series_count = value_label_series_count(
            graph.info.kind,
            graph.info.direction,
            &graph.info.data,
            "Numbers",
            drawable_object_id,
        )?;
        read_native_radar_style(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            graph.info.kind,
            series_count,
        )
    }

    /// Apply one Radar `Style` selection to every series transactionally.
    pub fn set_sheet_chart_radar_series_style(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        style: ChartRadarSeriesStyle,
    ) -> Result<()> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        let series_count = value_label_series_count(
            graph.info.kind,
            graph.info.direction,
            &graph.info.data,
            "Numbers",
            drawable_object_id,
        )?;
        let mut staged = self.package().clone();
        set_native_radar_style(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            graph.info.kind,
            series_count,
            style,
        )?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.sheet_chart_radar_series_style(sheet_id, drawable_object_id)? != style {
            return Err(Error::InvalidFormat(
                "Numbers Radar series-style update failed validation".to_owned(),
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
    fn duplicated_radar_charts_have_copy_on_write_series_styles() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
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
            .set_sheet_chart_radar_series_style(
                sheet_id,
                source.drawable_object_id,
                ChartRadarSeriesStyle::Fill,
            )
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, source.drawable_object_id)
            .unwrap();
        editor
            .set_sheet_chart_radar_series_style(
                sheet_id,
                duplicate.drawable_object_id,
                ChartRadarSeriesStyle::Stroke,
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_radar_series_style(sheet_id, source.drawable_object_id)
                .unwrap(),
            ChartRadarSeriesStyle::Fill
        );
        assert_eq!(
            editor
                .sheet_chart_radar_series_style(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            ChartRadarSeriesStyle::Stroke
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
