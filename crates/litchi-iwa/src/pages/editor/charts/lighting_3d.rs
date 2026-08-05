//! Native 3D lighting-style CRUD for Pages body charts.

use super::*;
use crate::charts::Chart3dLightingStyle;
use crate::charts::lighting_3d::{
    chart_3d_lighting_style as read_native_chart_3d_lighting_style,
    set_chart_3d_lighting_style as set_native_chart_3d_lighting_style,
};

impl PagesEditor {
    /// Read one body chart's native Lighting Style choice.
    pub fn body_chart_3d_lighting_style(
        &self,
        drawable_object_id: u64,
    ) -> Result<Chart3dLightingStyle> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        read_native_chart_3d_lighting_style(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            graph.info.kind,
        )
    }

    /// Set one body chart's native Lighting Style choice.
    pub fn set_body_chart_3d_lighting_style(
        &mut self,
        drawable_object_id: u64,
        style: Chart3dLightingStyle,
    ) -> Result<()> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        if read_native_chart_3d_lighting_style(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            graph.info.kind,
        )? == style
        {
            return Ok(());
        }
        let mut staged = self.package().clone();
        set_native_chart_3d_lighting_style(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            graph.info.kind,
            style,
        )?;
        let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.body_chart_3d_lighting_style(drawable_object_id)? != style {
            return Err(Error::InvalidFormat(
                "Pages chart 3D lighting-style update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pages::PagesDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_document_supports_3d_lighting_style_crud() {
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
                .body_chart_3d_lighting_style(chart.drawable_object_id)
                .unwrap(),
            Chart3dLightingStyle::Default
        );
        editor
            .set_body_chart_3d_lighting_style(
                chart.drawable_object_id,
                Chart3dLightingStyle::MediumCenter,
            )
            .unwrap();
        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body_chart_3d_lighting_style(chart.drawable_object_id)
                .unwrap(),
            Chart3dLightingStyle::MediumCenter
        );
    }

    #[test]
    fn two_dimensional_charts_reject_3d_lighting_access() {
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
                .body_chart_3d_lighting_style(chart.drawable_object_id)
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
