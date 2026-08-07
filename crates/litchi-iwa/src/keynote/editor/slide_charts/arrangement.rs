//! Native chart Arrange-panel CRUD for Keynote slide charts.

use super::*;
use crate::charts::ChartArrangement;
use crate::charts::arrangement::{
    chart_arrangement as read_native_arrangement, set_chart_arrangement as set_native_arrangement,
};

impl KeynoteEditor {
    /// Read one slide chart's lock and aspect-ratio constraint.
    pub fn slide_chart_arrangement(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ChartArrangement> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        read_native_arrangement(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
        )
    }

    /// Set one slide chart's lock and aspect-ratio constraint.
    pub fn set_slide_chart_arrangement(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        arrangement: ChartArrangement,
    ) -> Result<()> {
        if self.slide_chart_arrangement(slide_index, drawable_object_id)? == arrangement {
            return Ok(());
        }
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_native_arrangement(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            arrangement,
        )?;
        let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.slide_chart_arrangement(slide_index, drawable_object_id)? != arrangement {
            return Err(Error::InvalidFormat(
                "Keynote chart arrangement update failed validation".to_owned(),
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
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_presentation_supports_chart_arrangement_crud() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_slide_chart(
                0,
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
                .slide_chart_arrangement(0, chart.drawable_object_id)
                .unwrap(),
            ChartArrangement::default()
        );
        assert_eq!(
            editor.slide_charts(0).unwrap()[0].arrangement,
            ChartArrangement::default()
        );

        let constrained = ChartArrangement::default().with_constrain_proportions(true);
        editor
            .set_slide_chart_arrangement(0, chart.drawable_object_id, constrained)
            .unwrap();
        let duplicate = editor
            .duplicate_slide_chart(0, chart_selector(&editor, &chart))
            .unwrap();
        assert_eq!(
            editor
                .slide_chart_arrangement(0, duplicate.drawable_object_id)
                .unwrap(),
            constrained
        );

        let locked = ChartArrangement::default().with_locked(true);
        editor
            .set_slide_chart_arrangement(0, duplicate.drawable_object_id, locked)
            .unwrap();
        assert_eq!(
            editor
                .slide_chart_arrangement(0, chart.drawable_object_id)
                .unwrap(),
            constrained
        );
        assert_eq!(
            editor
                .slide_chart_arrangement(0, duplicate.drawable_object_id)
                .unwrap(),
            locked
        );

        editor
            .remove_slide_chart(0, chart_selector(&editor, &duplicate))
            .unwrap();
        editor
            .set_slide_chart_arrangement(0, chart.drawable_object_id, ChartArrangement::default())
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
