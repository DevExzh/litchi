//! Native chart Arrange-panel CRUD for Numbers sheets.

use super::*;
use crate::charts::ChartArrangement;
use crate::charts::arrangement::{
    chart_arrangement as read_native_arrangement, set_chart_arrangement as set_native_arrangement,
};

impl NumbersEditor {
    /// Read one sheet chart's lock and aspect-ratio constraint.
    pub fn sheet_chart_arrangement(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ChartArrangement> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        read_native_arrangement(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
        )
    }

    /// Set one sheet chart's lock and aspect-ratio constraint.
    pub fn set_sheet_chart_arrangement(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        arrangement: ChartArrangement,
    ) -> Result<()> {
        if self.sheet_chart_arrangement(sheet_id, drawable_object_id)? == arrangement {
            return Ok(());
        }
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_native_arrangement(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            arrangement,
        )?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.sheet_chart_arrangement(sheet_id, drawable_object_id)? != arrangement {
            return Err(Error::InvalidFormat(
                "Numbers chart arrangement update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{ChartData, Kind};
    use crate::numbers::NumbersDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_spreadsheet_supports_chart_arrangement_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let chart = editor
            .add_sheet_chart(
                sheet_id,
                Kind::Line2d,
                data(),
                DrawablePoint { x: 20.0, y: 20.0 },
                DrawableSize {
                    width: 400.0,
                    height: 300.0,
                },
            )
            .unwrap();
        let baseline = editor.to_bytes().unwrap();
        assert_eq!(
            editor
                .sheet_chart_arrangement(sheet_id, chart.drawable_object_id)
                .unwrap(),
            ChartArrangement::default()
        );
        assert_eq!(
            editor.sheet_charts(sheet_id).unwrap()[0].arrangement,
            ChartArrangement::default()
        );

        let constrained = ChartArrangement::default().with_constrain_proportions(true);
        editor
            .set_sheet_chart_arrangement(sheet_id, chart.drawable_object_id, constrained)
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, chart.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_arrangement(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            constrained
        );

        let locked = ChartArrangement::default().with_locked(true);
        editor
            .set_sheet_chart_arrangement(sheet_id, duplicate.drawable_object_id, locked)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_arrangement(sheet_id, chart.drawable_object_id)
                .unwrap(),
            constrained
        );
        assert_eq!(
            editor
                .sheet_chart_arrangement(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            locked
        );

        editor
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        editor
            .set_sheet_chart_arrangement(
                sheet_id,
                chart.drawable_object_id,
                ChartArrangement::default(),
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    fn data() -> ChartData {
        ChartData::new(
            vec!["Series".to_owned()],
            vec!["A".to_owned(), "B".to_owned(), "C".to_owned()],
            vec![vec![Some(8.0), Some(20.0), Some(42.0)]],
        )
        .unwrap()
    }
}
