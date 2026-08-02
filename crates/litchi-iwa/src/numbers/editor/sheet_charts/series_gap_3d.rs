//! Native 3D line/area series-gap CRUD for Numbers sheet charts.

use super::*;
use crate::charts::Chart3dSeriesGap;
use crate::charts::series_gap_3d::{
    chart_3d_series_gap as read_native_chart_3d_series_gap,
    set_chart_3d_series_gap as set_native_chart_3d_series_gap,
};

impl NumbersEditor {
    /// Read one 3D line/area chart's `Between Series` gap.
    pub fn sheet_chart_3d_series_gap(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Chart3dSeriesGap> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        read_native_chart_3d_series_gap(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            graph.info.kind,
        )
    }

    /// Set one 3D line/area chart's `Between Series` gap.
    pub fn set_sheet_chart_3d_series_gap(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        gap: Chart3dSeriesGap,
    ) -> Result<()> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        if read_native_chart_3d_series_gap(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
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
            "Numbers",
            graph.info.kind,
            gap,
        )?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.sheet_chart_3d_series_gap(sheet_id, drawable_object_id)? != gap {
            return Err(Error::InvalidFormat(
                "Numbers chart 3D series-gap update failed validation".to_owned(),
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
    fn duplicated_charts_have_copy_on_write_3d_series_gaps() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let chart = editor
            .add_sheet_chart(
                sheet_id,
                ChartKind::Line3d,
                data(),
                DrawablePoint { x: 20.0, y: 20.0 },
                DrawableSize {
                    width: 400.0,
                    height: 300.0,
                },
            )
            .unwrap();
        editor
            .set_sheet_chart_3d_series_gap(
                sheet_id,
                chart.drawable_object_id,
                Chart3dSeriesGap::new(25).unwrap(),
            )
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, chart.drawable_object_id)
            .unwrap();
        editor
            .set_sheet_chart_3d_series_gap(
                sheet_id,
                duplicate.drawable_object_id,
                Chart3dSeriesGap::new(175).unwrap(),
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_3d_series_gap(sheet_id, chart.drawable_object_id)
                .unwrap()
                .percent(),
            25
        );
        assert_eq!(
            editor
                .sheet_chart_3d_series_gap(sheet_id, duplicate.drawable_object_id)
                .unwrap()
                .percent(),
            175
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
