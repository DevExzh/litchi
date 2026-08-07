//! Native 3D scene rotation CRUD for Keynote slide charts.

use super::*;
use crate::charts::Chart3dRotation;
use crate::charts::scene_3d::{
    chart_3d_rotation as read_native_chart_3d_rotation,
    set_chart_3d_rotation as set_native_chart_3d_rotation,
};

impl KeynoteEditor {
    /// Read the X/Y orientation of one slide chart's native 3D scene.
    pub fn slide_chart_3d_rotation(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Chart3dRotation> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        require_3d_scene(graph.info.kind, drawable_object_id)?;
        read_native_chart_3d_rotation(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
        )
    }

    /// Set the X/Y orientation of one slide chart's native 3D scene.
    pub fn set_slide_chart_3d_rotation(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        rotation: Chart3dRotation,
    ) -> Result<()> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        require_3d_scene(graph.info.kind, drawable_object_id)?;
        if read_native_chart_3d_rotation(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
        )? == rotation
        {
            return Ok(());
        }
        let mut staged = self.package().clone();
        set_native_chart_3d_rotation(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            rotation,
        )?;
        let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.slide_chart_3d_rotation(slide_index, drawable_object_id)? != rotation {
            return Err(Error::InvalidFormat(
                "Keynote chart 3D rotation update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

fn require_3d_scene(kind: Kind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_3d_scene() {
        return Err(Error::InvalidFormat(format!(
            "Keynote chart {drawable_object_id} kind {kind:?} has no 3D scene"
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
    fn scratch_presentation_supports_3d_chart_rotation_crud() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_slide_chart(
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
        let rotation = Chart3dRotation::from_degrees(37.5, 22.5).unwrap();
        editor
            .set_slide_chart_3d_rotation(0, chart.drawable_object_id, rotation)
            .unwrap();
        let duplicate = editor
            .duplicate_slide_chart(0, chart_selector(&editor, &chart))
            .unwrap();
        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_chart_3d_rotation(0, chart.drawable_object_id)
                .unwrap(),
            rotation
        );
        assert_eq!(
            reopened
                .slide_chart_3d_rotation(0, duplicate.drawable_object_id)
                .unwrap(),
            rotation
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
