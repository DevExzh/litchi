//! Native 3D bar-shape CRUD for Pages body charts.

use super::*;
use crate::charts::Chart3dBarShape;
use crate::charts::bar_shape_3d::{
    chart_3d_bar_shape as read_native_chart_3d_bar_shape,
    set_chart_3d_bar_shape as set_native_chart_3d_bar_shape,
};

impl PagesEditor {
    /// Read the rectangular/cylindrical geometry of one 3D bar or column chart.
    pub fn body_chart_3d_bar_shape(&self, drawable_object_id: u64) -> Result<Chart3dBarShape> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        require_3d_bar_shape(graph.info.kind, drawable_object_id)?;
        read_native_chart_3d_bar_shape(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
        )
    }

    /// Set the rectangular/cylindrical geometry of one 3D bar or column chart.
    pub fn set_body_chart_3d_bar_shape(
        &mut self,
        drawable_object_id: u64,
        shape: Chart3dBarShape,
    ) -> Result<()> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        require_3d_bar_shape(graph.info.kind, drawable_object_id)?;
        if read_native_chart_3d_bar_shape(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
        )? == shape
        {
            return Ok(());
        }
        let mut staged = self.package().clone();
        set_native_chart_3d_bar_shape(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            shape,
        )?;
        let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.body_chart_3d_bar_shape(drawable_object_id)? != shape {
            return Err(Error::InvalidFormat(
                "Pages chart 3D bar-shape update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

fn require_3d_bar_shape(kind: Kind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_3d_bar_shape() {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} kind {kind:?} has no 3D bar shape"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pages::PagesDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_document_supports_3d_bar_shape_crud() {
        let mut editor = PagesDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_body_chart(
                0,
                Kind::Column3d,
                data(),
                DrawablePoint { x: 20.0, y: 20.0 },
                DrawableSize {
                    width: 400.0,
                    height: 300.0,
                },
            )
            .unwrap();
        assert_eq!(
            editor
                .body_chart_3d_bar_shape(chart.drawable_object_id)
                .unwrap(),
            Chart3dBarShape::Rectangle
        );
        editor
            .set_body_chart_3d_bar_shape(chart.drawable_object_id, Chart3dBarShape::Cylinder)
            .unwrap();
        let duplicate = editor
            .duplicate_body_chart(chart.drawable_object_id, 1)
            .unwrap();
        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body_chart_3d_bar_shape(chart.drawable_object_id)
                .unwrap(),
            Chart3dBarShape::Cylinder
        );
        assert_eq!(
            reopened
                .body_chart_3d_bar_shape(duplicate.drawable_object_id)
                .unwrap(),
            Chart3dBarShape::Cylinder
        );
    }

    #[test]
    fn charts_without_3d_bars_reject_bar_shape_access() {
        let mut editor = PagesDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_body_chart(
                0,
                Kind::Area3d,
                data(),
                DrawablePoint { x: 20.0, y: 20.0 },
                DrawableSize {
                    width: 400.0,
                    height: 300.0,
                },
            )
            .unwrap();
        assert!(
            editor
                .body_chart_3d_bar_shape(chart.drawable_object_id)
                .is_err()
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
