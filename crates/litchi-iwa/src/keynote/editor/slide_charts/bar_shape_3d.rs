//! Native 3D bar-shape CRUD for Keynote slide charts.

use super::*;
use crate::charts::Chart3dBarShape;
use crate::charts::bar_shape_3d::{
    chart_3d_bar_shape as read_native_chart_3d_bar_shape,
    set_chart_3d_bar_shape as set_native_chart_3d_bar_shape,
};

impl KeynoteEditor {
    /// Read the rectangular/cylindrical geometry of one 3D bar or column chart.
    pub fn slide_chart_3d_bar_shape(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Chart3dBarShape> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        require_3d_bar_shape(graph.info.kind, drawable_object_id)?;
        read_native_chart_3d_bar_shape(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
        )
    }

    /// Set the rectangular/cylindrical geometry of one 3D bar or column chart.
    pub fn set_slide_chart_3d_bar_shape(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        shape: Chart3dBarShape,
    ) -> Result<()> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        require_3d_bar_shape(graph.info.kind, drawable_object_id)?;
        if read_native_chart_3d_bar_shape(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
        )? == shape
        {
            return Ok(());
        }
        let mut staged = self.package().clone();
        set_native_chart_3d_bar_shape(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            shape,
        )?;
        let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.slide_chart_3d_bar_shape(slide_index, drawable_object_id)? != shape {
            return Err(Error::InvalidFormat(
                "Keynote chart 3D bar-shape update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

fn require_3d_bar_shape(kind: Kind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_3d_bar_shape() {
        return Err(Error::InvalidFormat(format!(
            "Keynote chart {drawable_object_id} kind {kind:?} has no 3D bar shape"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_presentation_supports_3d_bar_shape_crud() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_slide_chart(
                0,
                Kind::StackedColumn3d,
                data(),
                DrawablePoint { x: 20.0, y: 20.0 },
                DrawableSize {
                    width: 400.0,
                    height: 300.0,
                },
            )
            .unwrap();
        editor
            .set_slide_chart_3d_bar_shape(0, chart.drawable_object_id, Chart3dBarShape::Cylinder)
            .unwrap();
        let duplicate = editor
            .duplicate_slide_chart(0, chart_selector(&editor, &chart))
            .unwrap();
        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_chart_3d_bar_shape(0, chart.drawable_object_id)
                .unwrap(),
            Chart3dBarShape::Cylinder
        );
        assert_eq!(
            reopened
                .slide_chart_3d_bar_shape(0, duplicate.drawable_object_id)
                .unwrap(),
            Chart3dBarShape::Cylinder
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
