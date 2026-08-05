//! Native 3D line/area series-gap CRUD for Pages body charts.

use super::*;
use crate::charts::Chart3dSeriesGap;
use crate::charts::series_gap_3d::{
    chart_3d_series_gap as read_native_chart_3d_series_gap,
    set_chart_3d_series_gap as set_native_chart_3d_series_gap,
};

impl PagesEditor {
    /// Read one 3D line/area chart's `Between Series` gap.
    pub fn body_chart_3d_series_gap(&self, drawable_object_id: u64) -> Result<Chart3dSeriesGap> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        read_native_chart_3d_series_gap(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            graph.info.kind,
        )
    }

    /// Set one 3D line/area chart's `Between Series` gap.
    pub fn set_body_chart_3d_series_gap(
        &mut self,
        drawable_object_id: u64,
        gap: Chart3dSeriesGap,
    ) -> Result<()> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        if read_native_chart_3d_series_gap(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            graph.info.kind,
        )? == gap
        {
            return Ok(());
        }
        let mut staged = self.package().clone();
        set_native_chart_3d_series_gap(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            graph.info.kind,
            gap,
        )?;
        let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.body_chart_3d_series_gap(drawable_object_id)? != gap {
            return Err(Error::InvalidFormat(
                "Pages chart 3D series-gap update failed validation".to_owned(),
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
    fn scratch_document_supports_3d_series_gap_crud() {
        let mut editor = PagesDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_body_chart(
                0,
                Kind::Area3d,
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
                .body_chart_3d_series_gap(chart.drawable_object_id)
                .unwrap(),
            Chart3dSeriesGap::NATIVE_DEFAULT
        );
        let gap = Chart3dSeriesGap::new(25).unwrap();
        editor
            .set_body_chart_3d_series_gap(chart.drawable_object_id, gap)
            .unwrap();
        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body_chart_3d_series_gap(chart.drawable_object_id)
                .unwrap(),
            gap
        );
    }

    #[test]
    fn charts_without_between_series_control_reject_access() {
        let mut editor = PagesDocumentBuilder::new().build().unwrap();
        for kind in [Kind::Area2d, Kind::StackedArea3d, Kind::Column3d] {
            let chart = editor
                .add_body_chart(
                    0,
                    kind,
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
                    .body_chart_3d_series_gap(chart.drawable_object_id)
                    .is_err()
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
