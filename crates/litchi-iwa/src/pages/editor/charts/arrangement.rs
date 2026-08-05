//! Native chart Arrange-panel CRUD for Pages body charts.

use super::*;
use crate::charts::ChartArrangement;
use crate::charts::arrangement::{
    chart_arrangement as read_native_arrangement, set_chart_arrangement as set_native_arrangement,
};

impl PagesEditor {
    /// Read one body chart's lock and aspect-ratio constraint.
    pub fn body_chart_arrangement(&self, drawable_object_id: u64) -> Result<ChartArrangement> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        read_native_arrangement(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
        )
    }

    /// Set one body chart's lock and aspect-ratio constraint.
    pub fn set_body_chart_arrangement(
        &mut self,
        drawable_object_id: u64,
        arrangement: ChartArrangement,
    ) -> Result<()> {
        if self.body_chart_arrangement(drawable_object_id)? == arrangement {
            return Ok(());
        }
        let graph = body_chart_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_native_arrangement(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            arrangement,
        )?;
        let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.body_chart_arrangement(drawable_object_id)? != arrangement {
            return Err(Error::InvalidFormat(
                "Pages chart arrangement update failed validation".to_owned(),
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
    use crate::pages::PagesDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_document_supports_chart_arrangement_crud() {
        let body = "Chart arrangement";
        let mut editor = PagesDocumentBuilder::new().body_text(body).build().unwrap();
        let chart = editor
            .add_body_chart(
                body.encode_utf16().count(),
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
                .body_chart_arrangement(chart.drawable_object_id)
                .unwrap(),
            ChartArrangement::default()
        );
        assert_eq!(
            editor.body_charts().unwrap()[0].arrangement,
            ChartArrangement::default()
        );

        let constrained = ChartArrangement::default().with_constrain_proportions(true);
        editor
            .set_body_chart_arrangement(chart.drawable_object_id, constrained)
            .unwrap();
        let duplicate = editor
            .duplicate_body_chart(
                chart.drawable_object_id,
                editor.body_text().unwrap().encode_utf16().count(),
            )
            .unwrap();
        assert_eq!(
            editor
                .body_chart_arrangement(duplicate.drawable_object_id)
                .unwrap(),
            constrained
        );

        let locked = ChartArrangement::default().with_locked(true);
        editor
            .set_body_chart_arrangement(duplicate.drawable_object_id, locked)
            .unwrap();
        assert_eq!(
            editor
                .body_chart_arrangement(chart.drawable_object_id)
                .unwrap(),
            constrained
        );
        assert_eq!(
            editor
                .body_chart_arrangement(duplicate.drawable_object_id)
                .unwrap(),
            locked
        );

        editor
            .remove_body_chart(duplicate.drawable_object_id)
            .unwrap();
        editor
            .set_body_chart_arrangement(chart.drawable_object_id, ChartArrangement::default())
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
