//! Native 3D chart-depth CRUD for Keynote slide charts.

use super::*;
use crate::charts::Chart3dDepth;
use crate::charts::depth_3d::{
    chart_3d_depth as read_native_chart_3d_depth, set_chart_3d_depth as set_native_chart_3d_depth,
};

impl KeynoteEditor {
    /// Read one slide chart's depth in Chart-inspector percent.
    pub fn slide_chart_3d_depth(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Chart3dDepth> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        require_3d_depth(graph.info.kind, drawable_object_id)?;
        read_native_chart_3d_depth(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            graph.info.kind,
        )
    }

    /// Set one slide chart's depth in Chart-inspector percent.
    pub fn set_slide_chart_3d_depth(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        depth: Chart3dDepth,
    ) -> Result<()> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        require_3d_depth(graph.info.kind, drawable_object_id)?;
        if read_native_chart_3d_depth(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            graph.info.kind,
        )? == depth
        {
            return Ok(());
        }
        let mut staged = self.package().clone();
        set_native_chart_3d_depth(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            graph.info.kind,
            depth,
        )?;
        let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.slide_chart_3d_depth(slide_index, drawable_object_id)? != depth {
            return Err(Error::InvalidFormat(
                "Keynote chart 3D depth update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

fn require_3d_depth(kind: Kind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_3d_depth() {
        return Err(Error::InvalidFormat(format!(
            "Keynote chart {drawable_object_id} kind {kind:?} has no 3D depth"
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
    fn scratch_presentation_supports_3d_chart_depth_crud() {
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
        let depth = Chart3dDepth::from_percent(100.0).unwrap();
        editor
            .set_slide_chart_3d_depth(0, chart.drawable_object_id, depth)
            .unwrap();
        let duplicate = editor
            .duplicate_slide_chart(0, chart.drawable_object_id)
            .unwrap();
        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_chart_3d_depth(0, chart.drawable_object_id)
                .unwrap(),
            depth
        );
        assert_eq!(
            reopened
                .slide_chart_3d_depth(0, duplicate.drawable_object_id)
                .unwrap(),
            depth
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
