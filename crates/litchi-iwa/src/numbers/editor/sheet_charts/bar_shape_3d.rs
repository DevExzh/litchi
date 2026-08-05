//! Native 3D bar-shape CRUD for Numbers sheet charts.

use super::*;
use crate::charts::Chart3dBarShape;
use crate::charts::bar_shape_3d::{
    chart_3d_bar_shape as read_native_chart_3d_bar_shape,
    set_chart_3d_bar_shape as set_native_chart_3d_bar_shape,
};

impl NumbersEditor {
    /// Read the rectangular/cylindrical geometry of one 3D bar or column chart.
    pub fn sheet_chart_3d_bar_shape(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Chart3dBarShape> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        require_3d_bar_shape(graph.info.kind, drawable_object_id)?;
        read_native_chart_3d_bar_shape(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
        )
    }

    /// Set the rectangular/cylindrical geometry of one 3D bar or column chart.
    pub fn set_sheet_chart_3d_bar_shape(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        shape: Chart3dBarShape,
    ) -> Result<()> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        require_3d_bar_shape(graph.info.kind, drawable_object_id)?;
        if read_native_chart_3d_bar_shape(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
        )? == shape
        {
            return Ok(());
        }
        let mut staged = self.package().clone();
        set_native_chart_3d_bar_shape(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            shape,
        )?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.sheet_chart_3d_bar_shape(sheet_id, drawable_object_id)? != shape {
            return Err(Error::InvalidFormat(
                "Numbers chart 3D bar-shape update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

fn require_3d_bar_shape(kind: Kind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_3d_bar_shape() {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} kind {kind:?} has no 3D bar shape"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::NumbersDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn duplicated_sheet_charts_have_copy_on_write_3d_bar_shapes() {
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
            .set_sheet_chart_3d_bar_shape(
                sheet_id,
                chart.drawable_object_id,
                Chart3dBarShape::Cylinder,
            )
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, chart.drawable_object_id)
            .unwrap();
        editor
            .set_sheet_chart_3d_bar_shape(
                sheet_id,
                duplicate.drawable_object_id,
                Chart3dBarShape::Rectangle,
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_3d_bar_shape(sheet_id, chart.drawable_object_id)
                .unwrap(),
            Chart3dBarShape::Cylinder
        );
        assert_eq!(
            editor
                .sheet_chart_3d_bar_shape(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            Chart3dBarShape::Rectangle
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
