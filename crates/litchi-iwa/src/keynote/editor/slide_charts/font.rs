//! Chart-wide font CRUD for Keynote slide charts.

use super::*;
use crate::charts::ChartFont;
use crate::charts::font::{
    chart_font as read_native_font, reset_chart_font as reset_native_font,
    set_chart_font as set_native_font,
};

impl KeynoteEditor {
    /// Read the uniform effective font used by one slide chart.
    pub fn slide_chart_font(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ChartFont> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        read_native_font(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
        )
    }

    /// Set the uniform font used by every semantic text slot in one slide chart.
    pub fn set_slide_chart_font(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        font: ChartFont,
    ) -> Result<()> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_native_font(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            &font,
        )?;
        let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.slide_chart_font(slide_index, drawable_object_id)? != font {
            return Err(Error::InvalidFormat(
                "Keynote chart font update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Reset crate-owned chart font overrides to their inherited theme values.
    pub fn reset_slide_chart_font(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        if !reset_native_font(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
        )? {
            return Ok(false);
        }
        *self = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        Ok(true)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{ChartData, ChartKind};
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_presentation_supports_chart_font_crud_and_copy_on_write() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_slide_chart(
                0,
                ChartKind::Line2d,
                data(),
                DrawablePoint { x: 20.0, y: 20.0 },
                DrawableSize {
                    width: 400.0,
                    height: 300.0,
                },
            )
            .unwrap();
        let baseline = editor.to_bytes().unwrap();
        let inherited = editor
            .slide_chart_font(0, chart.drawable_object_id)
            .unwrap();
        let demi = ChartFont::named("AvenirNext-DemiBold")
            .unwrap()
            .with_bold(true);
        editor
            .set_slide_chart_font(0, chart.drawable_object_id, demi.clone())
            .unwrap();
        let duplicate = editor
            .duplicate_slide_chart(0, chart.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .slide_chart_font(0, duplicate.drawable_object_id)
                .unwrap(),
            demi
        );
        let italic = ChartFont::named("AvenirNext-Italic")
            .unwrap()
            .with_italic(true);
        editor
            .set_slide_chart_font(0, duplicate.drawable_object_id, italic.clone())
            .unwrap();
        assert_eq!(
            editor
                .slide_chart_font(0, chart.drawable_object_id)
                .unwrap(),
            demi
        );
        assert_eq!(
            editor
                .slide_chart_font(0, duplicate.drawable_object_id)
                .unwrap(),
            italic
        );
        assert!(
            editor
                .reset_slide_chart_font(0, duplicate.drawable_object_id)
                .unwrap()
        );
        editor
            .remove_slide_chart(0, duplicate.drawable_object_id)
            .unwrap();
        assert!(
            editor
                .reset_slide_chart_font(0, chart.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            editor
                .slide_chart_font(0, chart.drawable_object_id)
                .unwrap(),
            inherited
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    fn data() -> ChartData {
        ChartData::new(
            vec!["Series".to_owned()],
            vec!["A".to_owned(), "B".to_owned(), "C".to_owned()],
            vec![vec![Some(8.0), Some(20.0), Some(42.0)]],
        )
        .unwrap()
    }
}
