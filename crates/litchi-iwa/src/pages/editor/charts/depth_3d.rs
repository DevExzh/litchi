//! Native 3D chart-depth CRUD for Pages body charts.

use super::*;
use crate::charts::Chart3dDepth;
use crate::charts::depth_3d::{
    chart_3d_depth as read_native_chart_3d_depth, set_chart_3d_depth as set_native_chart_3d_depth,
};

impl PagesEditor {
    /// Read one body chart's depth in Chart-inspector percent.
    pub fn body_chart_3d_depth(&self, drawable_object_id: u64) -> Result<Chart3dDepth> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        require_3d_depth(graph.info.kind, drawable_object_id)?;
        read_native_chart_3d_depth(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            graph.info.kind,
        )
    }

    /// Set one body chart's depth in Chart-inspector percent.
    pub fn set_body_chart_3d_depth(
        &mut self,
        drawable_object_id: u64,
        depth: Chart3dDepth,
    ) -> Result<()> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        require_3d_depth(graph.info.kind, drawable_object_id)?;
        if read_native_chart_3d_depth(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
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
            "Pages",
            graph.info.kind,
            depth,
        )?;
        let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.body_chart_3d_depth(drawable_object_id)? != depth {
            return Err(Error::InvalidFormat(
                "Pages chart 3D depth update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

fn require_3d_depth(kind: Kind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_3d_depth() {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} kind {kind:?} has no 3D depth"
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
    fn scratch_document_supports_3d_chart_depth_crud() {
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
                .body_chart_3d_depth(chart.drawable_object_id)
                .unwrap(),
            Chart3dDepth::DEFAULT
        );
        let depth = Chart3dDepth::from_percent(50.0).unwrap();
        editor
            .set_body_chart_3d_depth(chart.drawable_object_id, depth)
            .unwrap();
        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body_chart_3d_depth(chart.drawable_object_id)
                .unwrap(),
            depth
        );
    }

    #[test]
    fn two_dimensional_charts_reject_3d_depth_access() {
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
                .body_chart_3d_depth(chart.drawable_object_id)
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
