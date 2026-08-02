//! Native 3D lighting-style CRUD for Keynote slide charts.

use super::*;
use crate::charts::Chart3dLightingStyle;
use crate::charts::lighting_3d::{
    chart_3d_lighting_style as read_native_chart_3d_lighting_style,
    set_chart_3d_lighting_style as set_native_chart_3d_lighting_style,
};

impl KeynoteEditor {
    /// Read one slide chart's native Lighting Style choice.
    pub fn slide_chart_3d_lighting_style(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Chart3dLightingStyle> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        read_native_chart_3d_lighting_style(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            graph.info.kind,
        )
    }

    /// Set one slide chart's native Lighting Style choice.
    pub fn set_slide_chart_3d_lighting_style(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        style: Chart3dLightingStyle,
    ) -> Result<()> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        if read_native_chart_3d_lighting_style(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
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
            "Keynote",
            graph.info.kind,
            style,
        )?;
        let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.slide_chart_3d_lighting_style(slide_index, drawable_object_id)? != style {
            return Err(Error::InvalidFormat(
                "Keynote chart 3D lighting-style update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_presentation_supports_3d_lighting_style_crud() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_slide_chart(
                0,
                ChartKind::Area3d,
                data(),
                DrawablePoint { x: 20.0, y: 20.0 },
                DrawableSize {
                    width: 400.0,
                    height: 300.0,
                },
            )
            .unwrap();
        editor
            .set_slide_chart_3d_lighting_style(
                0,
                chart.drawable_object_id,
                Chart3dLightingStyle::MediumLeft,
            )
            .unwrap();
        let duplicate = editor
            .duplicate_slide_chart(0, chart.drawable_object_id)
            .unwrap();
        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_chart_3d_lighting_style(0, chart.drawable_object_id)
                .unwrap(),
            Chart3dLightingStyle::MediumLeft
        );
        assert_eq!(
            reopened
                .slide_chart_3d_lighting_style(0, duplicate.drawable_object_id)
                .unwrap(),
            Chart3dLightingStyle::MediumLeft
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
