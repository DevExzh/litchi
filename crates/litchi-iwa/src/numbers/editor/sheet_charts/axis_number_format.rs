//! Native axis-label number-format CRUD for Numbers sheet charts.

use super::*;
use crate::charts::axis_number_format::{
    chart_axis_number_format as read_native_chart_axis_number_format,
    set_chart_axis_number_format as set_native_chart_axis_number_format,
};
use crate::charts::{Axis, NumberFormat};

impl NumbersEditor {
    /// Read the decimal-number format of one native sheet-chart axis.
    pub fn sheet_chart_axis_number_format(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<NumberFormat> {
        sheet_chart_axis_number_format(self, sheet_id, drawable_object_id, axis)
    }

    /// Set or reset the decimal-number format of one native sheet-chart axis.
    pub fn set_sheet_chart_axis_number_format(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: Axis,
        format: NumberFormat,
    ) -> Result<()> {
        set_sheet_chart_axis_number_format(self, sheet_id, drawable_object_id, axis, format)
    }
}

fn sheet_chart_axis_number_format(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    axis: Axis,
) -> Result<NumberFormat> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_chart_axis_number_format(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        axis,
    )
}

fn set_sheet_chart_axis_number_format(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    axis: Axis,
    format: NumberFormat,
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_chart_axis_number_format(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        axis,
        format,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_axis_number_format(sheet_id, drawable_object_id, axis)? != format {
        return Err(Error::InvalidFormat(
            "Numbers chart axis number-format update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{ChartData, Kind, DecimalPlaces, NegativeStyle};
    use crate::numbers::NumbersDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_spreadsheet_supports_axis_number_format_crud() {
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
        let expected = NumberFormat::new(
            DecimalPlaces::fixed(2).unwrap(),
            NegativeStyle::Parentheses,
            true,
        );
        assert_eq!(
            editor
                .sheet_chart_axis_number_format(sheet_id, chart.drawable_object_id, Axis::Value,)
                .unwrap(),
            NumberFormat::AXIS_NATIVE_DEFAULT
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_sheet_chart_axis_number_format(
                sheet_id,
                chart.drawable_object_id,
                Axis::Value,
                expected,
            )
            .unwrap();
        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_axis_number_format(sheet_id, chart.drawable_object_id, Axis::Value,)
                .unwrap(),
            expected
        );
        reopened
            .set_sheet_chart_axis_number_format(
                sheet_id,
                chart.drawable_object_id,
                Axis::Value,
                NumberFormat::AXIS_NATIVE_DEFAULT,
            )
            .unwrap();
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
