//! Native axis-label prefix and suffix CRUD for Pages body charts.

use super::*;
use crate::charts::axis_label_affixes::{
    chart_axis_label_affixes as read_native_axis_label_affixes,
    set_chart_axis_label_affixes as set_native_axis_label_affixes,
};
use crate::charts::{Axis, ChartLabelAffixes};

impl PagesEditor {
    /// Read the prefix and suffix applied to one body-chart axis' labels.
    pub fn body_chart_axis_label_affixes(
        &self,
        drawable_object_id: u64,
        axis: Axis,
    ) -> Result<ChartLabelAffixes> {
        body_chart_axis_label_affixes(self, drawable_object_id, axis)
    }

    /// Set or clear one body-chart axis' label prefix and suffix.
    pub fn set_body_chart_axis_label_affixes(
        &mut self,
        drawable_object_id: u64,
        axis: Axis,
        affixes: ChartLabelAffixes,
    ) -> Result<()> {
        set_body_chart_axis_label_affixes(self, drawable_object_id, axis, affixes)
    }
}

fn body_chart_axis_label_affixes(
    editor: &PagesEditor,
    drawable_object_id: u64,
    axis: Axis,
) -> Result<ChartLabelAffixes> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_axis_label_affixes(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        axis,
    )
}

fn set_body_chart_axis_label_affixes(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    axis: Axis,
    affixes: ChartLabelAffixes,
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_axis_label_affixes(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        axis,
        &affixes,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_axis_label_affixes(drawable_object_id, axis)? != affixes {
        return Err(Error::InvalidFormat(
            "Pages chart axis label-affix update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{ChartData, ChartKind};
    use crate::pages::PagesDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_document_supports_axis_label_affix_crud() {
        let body = "Axis affixes";
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
        let expected = ChartLabelAffixes::new("USD ", " net");
        assert_eq!(
            editor
                .body_chart_axis_label_affixes(chart.drawable_object_id, Axis::Value)
                .unwrap(),
            ChartLabelAffixes::default()
        );
        let baseline = editor.to_bytes().unwrap();
        editor
            .set_body_chart_axis_label_affixes(
                chart.drawable_object_id,
                Axis::Value,
                expected.clone(),
            )
            .unwrap();
        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body_chart_axis_label_affixes(chart.drawable_object_id, Axis::Value)
                .unwrap(),
            expected
        );
        reopened
            .set_body_chart_axis_label_affixes(
                chart.drawable_object_id,
                Axis::Value,
                ChartLabelAffixes::default(),
            )
            .unwrap();
        assert_eq!(
            reopened
                .body_chart_axis_label_affixes(chart.drawable_object_id, Axis::Value)
                .unwrap(),
            ChartLabelAffixes::default()
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
