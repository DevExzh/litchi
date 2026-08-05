//! Native 3D value-axis label-position CRUD for Numbers sheet charts.

use super::*;
use crate::charts::LabelPosition3d;
use crate::charts::axis_label_position_3d::{
    chart_3d_value_axis_label_position as read_native_3d_value_axis_label_position,
    set_chart_3d_value_axis_label_position as set_native_3d_value_axis_label_position,
};

impl NumbersEditor {
    /// Read one sheet chart's 3D primary value-axis label position.
    pub fn sheet_chart_3d_value_axis_label_position(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<LabelPosition3d> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        read_native_3d_value_axis_label_position(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            graph.info.kind,
        )
    }

    /// Set one sheet chart's 3D primary value-axis label position.
    pub fn set_sheet_chart_3d_value_axis_label_position(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        position: LabelPosition3d,
    ) -> Result<()> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        if read_native_3d_value_axis_label_position(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            graph.info.kind,
        )? == position
        {
            return Ok(());
        }
        let mut staged = self.package().clone();
        set_native_3d_value_axis_label_position(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            graph.info.kind,
            position,
        )?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.sheet_chart_3d_value_axis_label_position(sheet_id, drawable_object_id)?
            != position
        {
            return Err(Error::InvalidFormat(
                "Numbers chart 3D value-axis label-position update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::NumbersDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn duplicated_sheet_charts_have_copy_on_write_axis_label_positions() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let chart = editor
            .add_sheet_chart(
                sheet_id,
                Kind::Bar3d,
                data(),
                DrawablePoint { x: 20.0, y: 20.0 },
                DrawableSize {
                    width: 400.0,
                    height: 300.0,
                },
            )
            .unwrap();
        editor
            .set_sheet_chart_3d_value_axis_label_position(
                sheet_id,
                chart.drawable_object_id,
                LabelPosition3d::Leading,
            )
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, chart.drawable_object_id)
            .unwrap();
        editor
            .set_sheet_chart_3d_value_axis_label_position(
                sheet_id,
                duplicate.drawable_object_id,
                LabelPosition3d::Trailing,
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_3d_value_axis_label_position(sheet_id, chart.drawable_object_id)
                .unwrap(),
            LabelPosition3d::Leading
        );
        assert_eq!(
            editor
                .sheet_chart_3d_value_axis_label_position(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            LabelPosition3d::Trailing
        );
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
