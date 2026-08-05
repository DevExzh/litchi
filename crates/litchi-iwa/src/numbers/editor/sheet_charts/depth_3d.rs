//! Native 3D chart-depth CRUD for Numbers sheet charts.

use super::*;
use crate::charts::Chart3dDepth;
use crate::charts::depth_3d::{
    chart_3d_depth as read_native_chart_3d_depth, set_chart_3d_depth as set_native_chart_3d_depth,
};

impl NumbersEditor {
    /// Read one sheet chart's depth in Chart-inspector percent.
    pub fn sheet_chart_3d_depth(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Chart3dDepth> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        require_3d_depth(graph.info.kind, drawable_object_id)?;
        read_native_chart_3d_depth(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            graph.info.kind,
        )
    }

    /// Set one sheet chart's depth in Chart-inspector percent.
    pub fn set_sheet_chart_3d_depth(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        depth: Chart3dDepth,
    ) -> Result<()> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        require_3d_depth(graph.info.kind, drawable_object_id)?;
        if read_native_chart_3d_depth(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
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
            "Numbers",
            graph.info.kind,
            depth,
        )?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.sheet_chart_3d_depth(sheet_id, drawable_object_id)? != depth {
            return Err(Error::InvalidFormat(
                "Numbers chart 3D depth update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

fn require_3d_depth(kind: Kind, drawable_object_id: u64) -> Result<()> {
    if !kind.supports_3d_depth() {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} kind {kind:?} has no 3D depth"
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
    fn duplicated_sheet_charts_have_copy_on_write_3d_depths() {
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
        let original = Chart3dDepth::from_percent(25.0).unwrap();
        editor
            .set_sheet_chart_3d_depth(sheet_id, chart.drawable_object_id, original)
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, chart.drawable_object_id)
            .unwrap();
        let changed = Chart3dDepth::from_percent(75.0).unwrap();
        editor
            .set_sheet_chart_3d_depth(sheet_id, duplicate.drawable_object_id, changed)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_3d_depth(sheet_id, chart.drawable_object_id)
                .unwrap(),
            original
        );
        assert_eq!(
            editor
                .sheet_chart_3d_depth(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            changed
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
