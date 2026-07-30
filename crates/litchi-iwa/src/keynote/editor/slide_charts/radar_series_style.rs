//! Native chart-wide Radar series-style CRUD for Keynote slide charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::ChartRadarSeriesStyle;
use crate::charts::radar_series_style::{
    chart_radar_series_style as read_native_radar_style,
    set_chart_radar_series_style as set_native_radar_style,
};

impl KeynoteEditor {
    /// Read the chart-wide Radar `Style` selection.
    pub fn slide_chart_radar_series_style(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ChartRadarSeriesStyle> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        let series_count = value_label_series_count(
            graph.info.kind,
            graph.info.direction,
            &graph.info.data,
            "Keynote",
            drawable_object_id,
        )?;
        read_native_radar_style(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            graph.info.kind,
            series_count,
        )
    }

    /// Apply one Radar `Style` selection to every series transactionally.
    pub fn set_slide_chart_radar_series_style(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        style: ChartRadarSeriesStyle,
    ) -> Result<()> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        let series_count = value_label_series_count(
            graph.info.kind,
            graph.info.direction,
            &graph.info.data,
            "Keynote",
            drawable_object_id,
        )?;
        let mut staged = self.package().clone();
        set_native_radar_style(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            graph.info.kind,
            series_count,
            style,
        )?;
        let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.slide_chart_radar_series_style(slide_index, drawable_object_id)? != style {
            return Err(Error::InvalidFormat(
                "Keynote Radar series-style update failed validation".to_owned(),
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
    fn scratch_presentation_supports_radar_series_style_crud() {
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
            .set_slide_chart_radar_series_style(
                0,
                chart.drawable_object_id,
                ChartRadarSeriesStyle::Stroke,
            )
            .unwrap();
        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_chart_radar_series_style(0, chart.drawable_object_id)
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
