//! Native 3D scene rotation CRUD for Numbers sheet charts.

use super::*;
use crate::charts::Chart3dRotation;
use crate::charts::scene_3d::{
    chart_3d_rotation as read_native_chart_3d_rotation,
    set_chart_3d_rotation as set_native_chart_3d_rotation,
};

impl NumbersEditor {
    /// Read the X/Y orientation of one sheet chart's native 3D scene.
    pub fn sheet_chart_3d_rotation(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Chart3dRotation> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        require_3d_scene(graph.info.kind, drawable_object_id)?;
        read_native_chart_3d_rotation(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
        )
    }

    /// Set the X/Y orientation of one sheet chart's native 3D scene.
    pub fn set_sheet_chart_3d_rotation(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        rotation: Chart3dRotation,
    ) -> Result<()> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        require_3d_scene(graph.info.kind, drawable_object_id)?;
        if read_native_chart_3d_rotation(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
        )? == rotation
        {
            return Ok(());
        }
        let mut staged = self.package().clone();
        set_native_chart_3d_rotation(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            rotation,
        )?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.sheet_chart_3d_rotation(sheet_id, drawable_object_id)? != rotation {
            return Err(Error::InvalidFormat(
                "Numbers chart 3D rotation update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

fn require_3d_scene(kind: Kind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_3d_scene() {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} kind {kind:?} has no 3D scene"
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
    fn scratch_spreadsheet_supports_3d_chart_rotation_crud() {
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
        let rotation = Chart3dRotation::from_degrees(-12.5, 35.0).unwrap();
        editor
            .set_sheet_chart_3d_rotation(sheet_id, chart.drawable_object_id, rotation)
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, chart.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_3d_rotation(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            rotation
        );
        editor
            .set_sheet_chart_3d_rotation(
                sheet_id,
                duplicate.drawable_object_id,
                Chart3dRotation::DEFAULT,
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_3d_rotation(sheet_id, chart.drawable_object_id)
                .unwrap(),
            rotation
        );
        assert_eq!(
            editor
                .sheet_chart_3d_rotation(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            Chart3dRotation::DEFAULT
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
