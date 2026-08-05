//! Native axis-label prefix and suffix CRUD for Numbers sheet charts.

use super::*;
use crate::charts::axis_label_affixes::{
    chart_axis_label_affixes as read_native_axis_label_affixes,
    set_chart_axis_label_affixes as set_native_axis_label_affixes,
};
use crate::charts::{Axis, LabelAffixes};

impl NumbersEditor {
    /// Read the prefix and suffix applied to one sheet-chart axis' labels.
    pub fn sheet_chart_axis_label_affixes(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<LabelAffixes> {
        sheet_chart_axis_label_affixes(self, sheet_id, drawable_object_id, axis)
    }

    /// Set or clear one sheet-chart axis' label prefix and suffix.
    pub fn set_sheet_chart_axis_label_affixes(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: Axis,
        affixes: LabelAffixes,
    ) -> Result<()> {
        set_sheet_chart_axis_label_affixes(self, sheet_id, drawable_object_id, axis, affixes)
    }
}

fn sheet_chart_axis_label_affixes(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    axis: Axis,
) -> Result<LabelAffixes> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_axis_label_affixes(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        axis,
    )
}

fn set_sheet_chart_axis_label_affixes(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    axis: Axis,
    affixes: LabelAffixes,
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_axis_label_affixes(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        axis,
        &affixes,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_axis_label_affixes(sheet_id, drawable_object_id, axis)? != affixes {
        return Err(Error::InvalidFormat(
            "Numbers chart axis label-affix update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{ChartData, ChartKind};
    use crate::numbers::NumbersDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_spreadsheet_supports_axis_label_affix_crud() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let chart = editor
            .add_sheet_chart(
                sheet_id,
                ChartKind::Line2d,
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
                .sheet_chart_axis_label_affixes(sheet_id, chart.drawable_object_id, Axis::Value,)
                .unwrap(),
            LabelAffixes::default()
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_axis_label_affixes(
                sheet_id,
                chart.drawable_object_id,
                Axis::Value,
                expected.clone(),
            )
            .unwrap();
        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_axis_label_affixes(sheet_id, chart.drawable_object_id, Axis::Value,)
                .unwrap(),
            expected
        );
        reopened
            .set_sheet_chart_axis_label_affixes(
                sheet_id,
                chart.drawable_object_id,
                Axis::Value,
                LabelAffixes::default(),
            )
            .unwrap();
        assert_eq!(
            reopened
                .sheet_chart_axis_label_affixes(sheet_id, chart.drawable_object_id, Axis::Value,)
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
