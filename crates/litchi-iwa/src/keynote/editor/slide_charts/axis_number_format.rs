//! Native axis-label number-format CRUD for Keynote slide charts.

use super::*;
use crate::charts::axis_number_format::{
    chart_axis_number_format as read_native_chart_axis_number_format,
    set_chart_axis_number_format as set_native_chart_axis_number_format,
};
use crate::charts::{ChartAxis, ChartNumberFormat};

impl KeynoteEditor {
    /// Read the decimal-number format of one native slide-chart axis.
    pub fn slide_chart_axis_number_format(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: ChartAxis,
    ) -> Result<ChartNumberFormat> {
        slide_chart_axis_number_format(self, slide_index, drawable_object_id, axis)
    }

    /// Set or reset the decimal-number format of one native slide-chart axis.
    pub fn set_slide_chart_axis_number_format(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: ChartAxis,
        format: ChartNumberFormat,
    ) -> Result<()> {
        set_slide_chart_axis_number_format(self, slide_index, drawable_object_id, axis, format)
    }
}

fn slide_chart_axis_number_format(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: ChartAxis,
) -> Result<ChartNumberFormat> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_chart_axis_number_format(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
    )
}

fn set_slide_chart_axis_number_format(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    axis: ChartAxis,
    format: ChartNumberFormat,
) -> Result<()> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_axis_number_format(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        axis,
        format,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_axis_number_format(slide_index, drawable_object_id, axis)? != format {
        return Err(Error::InvalidFormat(
            "Keynote chart axis number-format update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{ChartData, ChartDecimalPlaces, ChartKind, ChartNegativeStyle};
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_presentation_supports_axis_number_format_crud() {
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
        exercise(&mut editor, chart.drawable_object_id);
    }

    fn exercise(editor: &mut KeynoteEditor, chart_id: u64) {
        let expected = ChartNumberFormat::new(
            ChartDecimalPlaces::fixed(2).unwrap(),
            ChartNegativeStyle::Parentheses,
            true,
        );
        assert_eq!(
            editor
                .slide_chart_axis_number_format(0, chart_id, ChartAxis::Value)
                .unwrap(),
            ChartNumberFormat::AXIS_NATIVE_DEFAULT
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_slide_chart_axis_number_format(0, chart_id, ChartAxis::Value, expected)
            .unwrap();
        let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_chart_axis_number_format(0, chart_id, ChartAxis::Value)
                .unwrap(),
            expected
        );
        reopened
            .set_slide_chart_axis_number_format(
                0,
                chart_id,
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
