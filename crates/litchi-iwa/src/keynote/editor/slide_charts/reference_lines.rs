//! Native value-axis reference-line CRUD for Keynote slide charts.

use super::*;
use crate::charts::ChartReferenceLine;
use crate::charts::reference_lines::{
    chart_reference_lines as read_native_reference_lines,
    set_chart_reference_lines as set_native_reference_lines,
};

impl KeynoteEditor {
    /// Read ordered reference lines on one slide chart's primary value axis.
    pub fn slide_chart_reference_lines(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<ChartReferenceLine>> {
        slide_chart_reference_lines(self, slide_index, drawable_object_id)
    }

    /// Replace ordered reference lines on one slide chart's primary value axis.
    pub fn set_slide_chart_reference_lines(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        reference_lines: &[ChartReferenceLine],
    ) -> Result<()> {
        set_slide_chart_reference_lines(self, slide_index, drawable_object_id, reference_lines)
    }
}

fn slide_chart_reference_lines(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<Vec<ChartReferenceLine>> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_reference_lines(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
    )
}

fn set_slide_chart_reference_lines(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    reference_lines: &[ChartReferenceLine],
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_reference_lines(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        reference_lines,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_reference_lines(slide_index, drawable_object_id)? != reference_lines {
        return Err(Error::InvalidFormat(
            "Keynote chart reference-line update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{ChartData, ChartKind, ChartReferenceLineValue};
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_presentation_supports_reference_line_crud() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_slide_chart(
                0,
                ChartKind::Line2d,
                data(),
                DrawablePoint { x: 20.0, y: 20.0 },
                DrawableSize {
                    width: 400.0,
                    height: 300.0,
                },
            )
            .unwrap();
        let baseline = editor.to_bytes().unwrap();
        assert!(
            editor
                .slide_chart_reference_lines(0, chart.drawable_object_id)
                .unwrap()
                .is_empty()
        );
        let initial = vec![
            ChartReferenceLine::average(),
            ChartReferenceLine::custom(ChartReferenceLineValue::new(17.5).unwrap())
                .with_name("Threshold"),
        ];
        editor
            .set_slide_chart_reference_lines(0, chart.drawable_object_id, &initial)
            .unwrap();
        let duplicate = editor
            .duplicate_slide_chart(0, chart.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .slide_chart_reference_lines(0, duplicate.drawable_object_id)
                .unwrap(),
            initial
        );
        let updated = vec![
            ChartReferenceLine::median()
                .with_name("Middle")
                .with_value_visibility(true),
            ChartReferenceLine::minimum().with_name_visibility(false),
            ChartReferenceLine::maximum(),
        ];
        editor
            .set_slide_chart_reference_lines(0, chart.drawable_object_id, &updated)
            .unwrap();
        assert_eq!(
            editor
                .slide_chart_reference_lines(0, duplicate.drawable_object_id)
                .unwrap(),
            initial
        );
        editor
            .remove_slide_chart(0, duplicate.drawable_object_id)
            .unwrap();
        let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_chart_reference_lines(0, chart.drawable_object_id)
                .unwrap(),
            updated
        );
        reopened
            .set_slide_chart_reference_lines(0, chart.drawable_object_id, &[])
            .unwrap();
        assert_eq!(reopened.to_bytes().unwrap(), baseline);
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
