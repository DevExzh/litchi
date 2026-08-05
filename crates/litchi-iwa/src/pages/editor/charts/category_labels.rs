//! Native category-label layout CRUD for Pages body charts.

use super::*;
use crate::charts::ChartCategoryLabelLayout;
use crate::charts::category_labels::{
    chart_category_label_layout as read_native_category_label_layout,
    set_chart_category_label_layout as set_native_category_label_layout,
};

impl PagesEditor {
    /// Read the complete category-label menu state for one body chart.
    pub fn body_chart_category_label_layout(
        &self,
        drawable_object_id: u64,
    ) -> Result<ChartCategoryLabelLayout> {
        body_chart_category_label_layout(self, drawable_object_id)
    }

    /// Set the complete category-label menu state for one body chart.
    pub fn set_body_chart_category_label_layout(
        &mut self,
        drawable_object_id: u64,
        layout: ChartCategoryLabelLayout,
    ) -> Result<()> {
        set_body_chart_category_label_layout(self, drawable_object_id, layout)
    }
}

fn body_chart_category_label_layout(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<ChartCategoryLabelLayout> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_category_label_layout(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
    )
}

fn set_body_chart_category_label_layout(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    layout: ChartCategoryLabelLayout,
) -> Result<()> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_category_label_layout(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        layout,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_category_label_layout(drawable_object_id)? != layout {
        return Err(Error::InvalidFormat(
            "Pages chart category-label layout update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{ChartCategoryLabelFrequency, ChartCategoryLabelInterval, ChartData, Kind};
    use crate::pages::PagesDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_document_supports_category_label_layout_crud() {
        let body = "Category label layout";
        let mut editor = PagesDocumentBuilder::new().body_text(body).build().unwrap();
        let chart = editor
            .add_body_chart(
                body.encode_utf16().count(),
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
        assert_eq!(
            editor
                .body_chart_category_label_layout(chart.drawable_object_id)
                .unwrap(),
            ChartCategoryLabelLayout::default()
        );
        for layout in [
            ChartCategoryLabelLayout::new(ChartCategoryLabelFrequency::None, true),
            ChartCategoryLabelLayout::new(ChartCategoryLabelFrequency::All, true),
        ] {
            editor
                .set_body_chart_category_label_layout(chart.drawable_object_id, layout)
                .unwrap();
            assert_eq!(
                editor
                    .body_chart_category_label_layout(chart.drawable_object_id)
                    .unwrap(),
                layout
            );
        }
        let customized = ChartCategoryLabelLayout::new(
            ChartCategoryLabelFrequency::Every(ChartCategoryLabelInterval::new(3).unwrap()),
            false,
        );
        editor
            .set_body_chart_category_label_layout(chart.drawable_object_id, customized)
            .unwrap();
        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body_chart_category_label_layout(chart.drawable_object_id)
                .unwrap(),
            customized
        );
        reopened
            .set_body_chart_category_label_layout(
                chart.drawable_object_id,
                ChartCategoryLabelLayout::default(),
            )
            .unwrap();
        assert_eq!(reopened.to_bytes().unwrap(), baseline);
    }

    fn data() -> ChartData {
        ChartData::new(
            vec!["Series".to_owned()],
            vec![
                "A".to_owned(),
                "B".to_owned(),
                "C".to_owned(),
                "D".to_owned(),
            ],
            vec![vec![Some(1.0), Some(2.0), Some(3.0), Some(4.0)]],
        )
        .unwrap()
    }
}
