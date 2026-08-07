//! Native value-axis reference-line CRUD for Numbers sheet charts.

use super::*;
use crate::charts::reference_line::Line;
use crate::charts::reference_line::{
    chart_reference_lines as read_native_reference_lines,
    set_chart_reference_lines as set_native_reference_lines,
};

impl NumbersEditor {
    /// Read ordered reference lines on one sheet chart's primary value axis.
    pub fn sheet_chart_reference_lines(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<Line>> {
        sheet_chart_reference_lines(self, sheet_id, drawable_object_id)
    }

    /// Replace ordered reference lines on one sheet chart's primary value axis.
    pub fn set_sheet_chart_reference_lines(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        reference_line: &[Line],
    ) -> Result<()> {
        set_sheet_chart_reference_lines(self, sheet_id, drawable_object_id, reference_line)
    }
}

fn sheet_chart_reference_lines(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<Vec<Line>> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_reference_lines(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
    )
}

fn set_sheet_chart_reference_lines(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    reference_line: &[Line],
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_reference_lines(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        reference_line,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_reference_lines(sheet_id, drawable_object_id)? != reference_line {
        return Err(Error::InvalidFormat(
            "Numbers chart reference-line update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::reference_line::Value;
    use crate::charts::{ChartData, Kind};
    use crate::numbers::NumbersDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_spreadsheet_supports_reference_line_crud() {
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
        assert!(
            editor
                .sheet_chart_reference_lines(sheet_id, chart.drawable_object_id)
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
            .set_sheet_chart_reference_lines(sheet_id, chart.drawable_object_id, &initial)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_reference_lines(sheet_id, chart.drawable_object_id)
                .unwrap(),
            initial
        );
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, chart.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_reference_lines(sheet_id, duplicate.drawable_object_id)
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
            .set_sheet_chart_reference_lines(sheet_id, chart.drawable_object_id, &updated)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_reference_lines(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            initial
        );
        editor
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_reference_lines(sheet_id, chart.drawable_object_id)
                .unwrap(),
            updated
        );
        reopened
            .set_sheet_chart_reference_lines(sheet_id, chart.drawable_object_id, &[])
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
