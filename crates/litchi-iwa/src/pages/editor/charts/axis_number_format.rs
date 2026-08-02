//! Native axis-label number-format CRUD for Pages body charts.

use super::*;
use crate::charts::axis_number_format::{
    chart_axis_number_format as read_native_chart_axis_number_format,
    set_chart_axis_number_format as set_native_chart_axis_number_format,
};
use crate::charts::{ChartAxis, ChartNumberFormat};

impl PagesEditor {
    /// Read the decimal-number format of one native body-chart axis.
    pub fn body_chart_axis_number_format(
        &self,
        drawable_object_id: u64,
        axis: ChartAxis,
    ) -> Result<ChartNumberFormat> {
        body_chart_axis_number_format(self, drawable_object_id, axis)
    }

    /// Set or reset the decimal-number format of one native body-chart axis.
    pub fn set_body_chart_axis_number_format(
        &mut self,
        drawable_object_id: u64,
        axis: ChartAxis,
        format: ChartNumberFormat,
    ) -> Result<()> {
        set_body_chart_axis_number_format(self, drawable_object_id, axis, format)
    }
}

fn body_chart_axis_number_format(
    editor: &PagesEditor,
    drawable_object_id: u64,
    axis: ChartAxis,
) -> Result<ChartNumberFormat> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_chart_axis_number_format(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        axis,
    )
}

fn set_body_chart_axis_number_format(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    axis: ChartAxis,
    format: ChartNumberFormat,
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_axis_number_format(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        axis,
        format,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_axis_number_format(drawable_object_id, axis)? != format {
        return Err(Error::InvalidFormat(
            "Pages chart axis number-format update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{ChartData, ChartDecimalPlaces, ChartKind, ChartNegativeStyle};
    use crate::pages::PagesDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_document_supports_axis_number_format_crud() {
        let body = "Axis format";
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
        let expected = ChartNumberFormat::new(
            ChartDecimalPlaces::fixed(2).unwrap(),
            ChartNegativeStyle::Parentheses,
            true,
        );
        assert_eq!(
            editor
                .body_chart_axis_number_format(chart.drawable_object_id, ChartAxis::Value)
                .unwrap(),
            ChartNumberFormat::AXIS_NATIVE_DEFAULT
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_body_chart_axis_number_format(chart.drawable_object_id, ChartAxis::Value, expected)
            .unwrap();
        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body_chart_axis_number_format(chart.drawable_object_id, ChartAxis::Value)
                .unwrap(),
            expected
        );
        reopened
            .set_body_chart_axis_number_format(
                chart.drawable_object_id,
                ChartAxis::Value,
                ChartNumberFormat::AXIS_NATIVE_DEFAULT,
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
