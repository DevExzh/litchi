//! Native category-label layout CRUD for Numbers sheet charts.

use super::*;
use crate::charts::ChartCategoryLabelLayout;
use crate::charts::category_labels::{
    chart_category_label_layout as read_native_category_label_layout,
    set_chart_category_label_layout as set_native_category_label_layout,
};

impl NumbersEditor {
    /// Read the complete category-label menu state for one sheet chart.
    pub fn sheet_chart_category_label_layout(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ChartCategoryLabelLayout> {
        sheet_chart_category_label_layout(self, sheet_id, drawable_object_id)
    }

    /// Set the complete category-label menu state for one sheet chart.
    pub fn set_sheet_chart_category_label_layout(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        layout: ChartCategoryLabelLayout,
    ) -> Result<()> {
        set_sheet_chart_category_label_layout(self, sheet_id, drawable_object_id, layout)
    }
}

fn sheet_chart_category_label_layout(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<ChartCategoryLabelLayout> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_category_label_layout(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
    )
}

fn set_sheet_chart_category_label_layout(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    layout: ChartCategoryLabelLayout,
) -> Result<()> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_category_label_layout(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        layout,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_category_label_layout(sheet_id, drawable_object_id)? != layout {
        return Err(Error::InvalidFormat(
            "Numbers chart category-label layout update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{ChartCategoryLabelFrequency, ChartCategoryLabelInterval, ChartData, Kind};
    use crate::numbers::NumbersDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_spreadsheet_supports_category_label_layout_crud() {
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
        assert_eq!(
            editor
                .sheet_chart_category_label_layout(sheet_id, chart.drawable_object_id)
                .unwrap(),
            ChartCategoryLabelLayout::default()
        );
        for layout in [
            ChartCategoryLabelLayout::new(ChartCategoryLabelFrequency::None, true),
            ChartCategoryLabelLayout::new(ChartCategoryLabelFrequency::All, true),
        ] {
            editor
                .set_sheet_chart_category_label_layout(sheet_id, chart.drawable_object_id, layout)
                .unwrap();
            assert_eq!(
                editor
                    .sheet_chart_category_label_layout(sheet_id, chart.drawable_object_id)
                    .unwrap(),
                layout
            );
        }
        let customized = ChartCategoryLabelLayout::new(
            ChartCategoryLabelFrequency::Every(ChartCategoryLabelInterval::new(3).unwrap()),
            false,
        );
        editor
            .set_sheet_chart_category_label_layout(sheet_id, chart.drawable_object_id, customized)
            .unwrap();
        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_chart_category_label_layout(sheet_id, chart.drawable_object_id)
                .unwrap(),
            customized
        );
        reopened
            .set_sheet_chart_category_label_layout(
                sheet_id,
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
