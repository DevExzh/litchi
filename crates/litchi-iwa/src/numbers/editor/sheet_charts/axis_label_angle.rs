//! Native axis-label angle CRUD for Numbers sheet charts.

use super::*;
use crate::charts::axis_label_angle::{
    chart_axis_label_angle as read_native_axis_label_angle,
    set_chart_axis_label_angle as set_native_axis_label_angle,
};
use crate::charts::{Axis, LabelAngle};

impl NumbersEditor {
    /// Read one sheet-chart axis' normalized label angle.
    pub fn sheet_chart_axis_label_angle(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<LabelAngle> {
        sheet_chart_axis_label_angle(self, sheet_id, drawable_object_id, axis)
    }

    /// Set or reset one sheet-chart axis' label angle.
    pub fn set_sheet_chart_axis_label_angle(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: Axis,
        angle: LabelAngle,
    ) -> Result<()> {
        set_sheet_chart_axis_label_angle(self, sheet_id, drawable_object_id, axis, angle)
    }
}

fn sheet_chart_axis_label_angle(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    axis: Axis,
) -> Result<LabelAngle> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_axis_label_angle(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        axis,
    )
}

fn set_sheet_chart_axis_label_angle(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    axis: Axis,
    angle: LabelAngle,
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_axis_label_angle(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        axis,
        angle,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_axis_label_angle(sheet_id, drawable_object_id, axis)? != angle {
        return Err(Error::InvalidFormat(
            "Numbers chart axis label-angle update failed validation".to_owned(),
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
    fn scratch_spreadsheet_supports_axis_label_angle_crud() {
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
        let baseline = editor.to_bytes().unwrap();
        assert_eq!(
            editor
                .sheet_chart_axis_label_angle(sheet_id, chart.drawable_object_id, Axis::Value,)
                .unwrap(),
            LabelAngle::HORIZONTAL
        );
        let expected = LabelAngle::new(12.5).unwrap();
        editor
            .set_sheet_chart_axis_label_angle(
                sheet_id,
                chart.drawable_object_id,
                Axis::Value,
                expected,
            )
            .unwrap();
        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_axis_label_angle(sheet_id, chart.drawable_object_id, Axis::Value,)
                .unwrap(),
            expected
        );
        reopened
            .set_sheet_chart_axis_label_angle(
                sheet_id,
                chart.drawable_object_id,
                Axis::Value,
                LabelAngle::HORIZONTAL,
            )
            .unwrap();
        assert_eq!(reopened.to_bytes().unwrap(), baseline);
    }

    fn data() -> ChartData {
        ChartData::new(
            vec!["Series".to_owned()],
            vec!["A".to_owned(), "B".to_owned()],
            vec![vec![Some(1.0), Some(2.0)]],
        )
        .unwrap()
    }
}
