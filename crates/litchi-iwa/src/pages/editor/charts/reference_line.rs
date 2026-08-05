//! Native value-axis reference-line CRUD for Pages body charts.

use super::*;
use crate::charts::reference_line::Line;
use crate::charts::reference_line::{
    chart_reference_lines as read_native_reference_lines,
    set_chart_reference_lines as set_native_reference_lines,
};

impl PagesEditor {
    /// Read ordered reference lines on one body chart's primary value axis.
    pub fn body_chart_reference_lines(&self, drawable_object_id: u64) -> Result<Vec<Line>> {
        body_chart_reference_lines(self, drawable_object_id)
    }

    /// Replace ordered reference lines on one body chart's primary value axis.
    pub fn set_body_chart_reference_lines(
        &mut self,
        drawable_object_id: u64,
        reference_line: &[Line],
    ) -> Result<()> {
        set_body_chart_reference_lines(self, drawable_object_id, reference_line)
    }
}

fn body_chart_reference_lines(editor: &PagesEditor, drawable_object_id: u64) -> Result<Vec<Line>> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_reference_lines(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
    )
}

fn set_body_chart_reference_lines(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    reference_line: &[Line],
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_reference_lines(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        reference_line,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_reference_lines(drawable_object_id)? != reference_line {
        return Err(Error::InvalidFormat(
            "Pages chart reference-line update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::reference_line::Value;
    use crate::charts::{ChartData, ChartKind};
    use crate::pages::PagesDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_document_supports_reference_line_crud() {
        let body = "Reference lines";
        let mut editor = PagesDocumentBuilder::new().body_text(body).build().unwrap();
        let chart = editor
            .add_body_chart(
                body.encode_utf16().count(),
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
                .body_chart_reference_lines(chart.drawable_object_id)
                .unwrap()
                .is_empty()
        );
        let initial = vec![
            Line::average(),
            Line::custom(Value::new(17.5).unwrap())
                .try_with_name("Threshold")
                .unwrap(),
        ];
        editor
            .set_body_chart_reference_lines(chart.drawable_object_id, &initial)
            .unwrap();
        let duplicate = editor
            .duplicate_body_chart(
                chart.drawable_object_id,
                editor.body_text().unwrap().encode_utf16().count(),
            )
            .unwrap();
        assert_eq!(
            editor
                .body_chart_reference_lines(duplicate.drawable_object_id)
                .unwrap(),
            initial
        );
        let updated = vec![
            Line::median()
                .try_with_name("Middle")
                .unwrap()
                .with_value_visibility(true),
            Line::minimum().with_name_visibility(false),
            Line::maximum(),
        ];
        editor
            .set_body_chart_reference_lines(chart.drawable_object_id, &updated)
            .unwrap();
        assert_eq!(
            editor
                .body_chart_reference_lines(duplicate.drawable_object_id)
                .unwrap(),
            initial
        );
        editor
            .remove_body_chart(duplicate.drawable_object_id)
            .unwrap();
        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body_chart_reference_lines(chart.drawable_object_id)
                .unwrap(),
            updated
        );
        reopened
            .set_body_chart_reference_lines(chart.drawable_object_id, &[])
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
