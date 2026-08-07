//! Native 3D scene rotation CRUD for Pages body charts.

use super::*;
use crate::charts::Chart3dRotation;
use crate::charts::scene_3d::{
    chart_3d_rotation as read_native_chart_3d_rotation,
    set_chart_3d_rotation as set_native_chart_3d_rotation,
};

impl PagesEditor {
    /// Read the X/Y orientation of one body chart's native 3D scene.
    pub fn body_chart_3d_rotation(&self, drawable_object_id: u64) -> Result<Chart3dRotation> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        require_3d_scene(graph.info.kind, drawable_object_id)?;
        read_native_chart_3d_rotation(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
        )
    }

    /// Set the X/Y orientation of one body chart's native 3D scene.
    pub fn set_body_chart_3d_rotation(
        &mut self,
        drawable_object_id: u64,
        rotation: Chart3dRotation,
    ) -> Result<()> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        require_3d_scene(graph.info.kind, drawable_object_id)?;
        if read_native_chart_3d_rotation(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
        )? == rotation
        {
            return Ok(());
        }
        let mut staged = self.package().clone();
        set_native_chart_3d_rotation(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            rotation,
        )?;
        let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.body_chart_3d_rotation(drawable_object_id)? != rotation {
            return Err(Error::InvalidFormat(
                "Pages chart 3D rotation update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

fn require_3d_scene(kind: Kind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_3d_scene() {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} kind {kind:?} has no 3D scene"
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
    fn scratch_document_supports_3d_chart_rotation_crud() {
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
                .body_chart_3d_rotation(chart.drawable_object_id)
                .unwrap(),
            Chart3dRotation::DEFAULT
        );
        let rotation = Chart3dRotation::from_degrees(30.0, -40.0).unwrap();
        editor
            .set_body_chart_3d_rotation(chart.drawable_object_id, rotation)
            .unwrap();
        let duplicate = editor
            .duplicate_body_chart(chart.drawable_object_id, 1)
            .unwrap();
        assert_eq!(
            editor
                .body_chart_3d_rotation(duplicate.drawable_object_id)
                .unwrap(),
            rotation
        );
        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body_chart_3d_rotation(chart.drawable_object_id)
                .unwrap(),
            rotation
        );
    }

    #[test]
    fn two_dimensional_charts_reject_3d_scene_access() {
        let mut editor = PagesDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_body_chart(
                0,
                Kind::Column2d,
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
                .body_chart_3d_rotation(chart.drawable_object_id)
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
