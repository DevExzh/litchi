//! Native chart-wide Radar series-style CRUD for Pages body charts.

use super::series_value_labels::value_label_series_count;
use super::*;
use crate::charts::ChartRadarSeriesStyle;
use crate::charts::radar_series_style::{
    chart_radar_series_style as read_native_radar_style,
    set_chart_radar_series_style as set_native_radar_style,
};

impl PagesEditor {
    /// Read the chart-wide Radar `Style` selection.
    pub fn body_chart_radar_series_style(
        &self,
        drawable_object_id: u64,
    ) -> Result<ChartRadarSeriesStyle> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        let series_count = value_label_series_count(
            graph.info.kind,
            graph.info.direction,
            &graph.info.data,
            "Pages",
            drawable_object_id,
        )?;
        read_native_radar_style(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            graph.info.kind,
            series_count,
        )
    }

    /// Apply one Radar `Style` selection to every series transactionally.
    pub fn set_body_chart_radar_series_style(
        &mut self,
        drawable_object_id: u64,
        style: ChartRadarSeriesStyle,
    ) -> Result<()> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        let series_count = value_label_series_count(
            graph.info.kind,
            graph.info.direction,
            &graph.info.data,
            "Pages",
            drawable_object_id,
        )?;
        let mut staged = self.package().clone();
        set_native_radar_style(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            graph.info.kind,
            series_count,
            style,
        )?;
        let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.body_chart_radar_series_style(drawable_object_id)? != style {
            return Err(Error::InvalidFormat(
                "Pages Radar series-style update failed validation".to_owned(),
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
    fn scratch_document_supports_all_radar_series_styles() {
        let body = "Radar Series Style CRUD";
        let mut editor = PagesDocumentBuilder::new().body_text(body).build().unwrap();
        let chart = editor
            .add_body_chart(
                body.encode_utf16().count(),
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
                .body_chart_radar_series_style(chart.drawable_object_id)
                .unwrap(),
            ChartRadarSeriesStyle::FillAndStroke
        );
        for style in [
            ChartRadarSeriesStyle::Fill,
            ChartRadarSeriesStyle::Stroke,
            ChartRadarSeriesStyle::FillAndStroke,
        ] {
            editor
                .set_body_chart_radar_series_style(chart.drawable_object_id, style)
                .unwrap();
            let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
            assert_eq!(
                reopened
                    .body_chart_radar_series_style(chart.drawable_object_id)
                    .unwrap(),
                style
            );
        }
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
