//! Native axis-label prefix and suffix CRUD for Keynote slide charts.

use super::*;
use crate::charts::axis_label_affixes::{
    chart_axis_label_affixes as read_native_axis_label_affixes,
    set_chart_axis_label_affixes as set_native_axis_label_affixes,
};
use crate::charts::{Axis, LabelAffixes};

impl KeynoteEditor {
    /// Read the prefix and suffix applied to one slide-chart axis' labels.
    pub fn slide_chart_axis_label_affixes(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<LabelAffixes> {
        slide_chart_axis_label_affixes(self, slide_index, drawable_object_id, axis)
    }

    /// Set or clear one slide-chart axis' label prefix and suffix.
    pub fn set_slide_chart_axis_label_affixes(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: Axis,
        affixes: LabelAffixes,
    ) -> Result<()> {
        set_slide_chart_axis_label_affixes(self, slide_index, drawable_object_id, axis, affixes)
    }
}

fn slide_chart_axis_label_affixes(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: Axis,
) -> Result<LabelAffixes> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_axis_label_affixes(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
    )
}

fn set_slide_chart_axis_label_affixes(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: Axis,
    affixes: LabelAffixes,
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_axis_label_affixes(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
        &affixes,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_axis_label_affixes(slide_index, drawable_object_id, axis)? != affixes {
        return Err(Error::InvalidFormat(
            "Keynote chart axis label-affix update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{ChartData, Kind};
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_presentation_supports_axis_label_affix_crud() {
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
        let expected = LabelAffixes::new("USD ", " net").unwrap();
        assert_eq!(
            editor
                .slide_chart_axis_label_affixes(0, chart.drawable_object_id, Axis::Value,)
                .unwrap(),
            LabelAffixes::default()
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_slide_chart_axis_label_affixes(
                0,
                chart.drawable_object_id,
                Axis::Value,
                expected.clone(),
            )
            .unwrap();
        let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_chart_axis_label_affixes(0, chart.drawable_object_id, Axis::Value,)
                .unwrap(),
            expected
        );
        reopened
            .set_slide_chart_axis_label_affixes(
                0,
                chart.drawable_object_id,
                Axis::Value,
                LabelAffixes::default(),
            )
            .unwrap();
        assert_eq!(
            reopened
                .slide_chart_axis_label_affixes(0, chart.drawable_object_id, Axis::Value,)
                .unwrap(),
            LabelAffixes::default()
        );
        assert_eq!(reopened.to_bytes().unwrap(), baseline);
    }

    fn data() -> ChartData {
        ChartData::new(
            vec!["Series".to_owned()],
            vec!["A".to_owned(), "B".to_owned()],
            vec![vec![Some(-1_000.5), Some(2_000.25)]],
        )
        .unwrap()
    }
}
