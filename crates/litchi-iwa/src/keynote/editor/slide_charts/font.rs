//! Chart-wide font CRUD for Keynote slide charts.

use super::*;
use crate::charts::font::{
    chart_font as read_native_font, chart_font_size as read_native_font_size,
    reset_chart_font as reset_native_font, reset_chart_font_size as reset_native_font_size,
    set_chart_font as set_native_font, set_chart_font_size as set_native_font_size,
};
use crate::charts::{ChartFont, ChartFontSize};

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

    /// Read the uniform effective point size used by one slide chart.
    pub fn slide_chart_font_size(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ChartFontSize> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        read_native_font_size(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
        )
    }

    /// Set the uniform point size used by every semantic text slot in one chart.
    pub fn set_slide_chart_font_size(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        size: ChartFontSize,
    ) -> Result<()> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_native_font_size(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            size,
        )?;
        let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.slide_chart_font_size(slide_index, drawable_object_id)? != size {
            return Err(Error::InvalidFormat(
                "Keynote chart font-size update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Reset crate-owned point-size overrides to the inherited theme size.
    pub fn reset_slide_chart_font_size(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        if !reset_native_font_size(
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
    use crate::charts::{ChartData, Kind};
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_presentation_supports_chart_font_crud_and_copy_on_write() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let chart = editor
            .add_slide_chart(
                0,
                Kind::Line2d,
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
        let inherited_size = editor
            .slide_chart_font_size(0, chart.drawable_object_id)
            .unwrap();
        let demi = ChartFont::named("AvenirNext-DemiBold")
            .unwrap()
            .with_bold(true);
        let large = ChartFontSize::from_points(18.0).unwrap();
        editor
            .set_slide_chart_font(0, chart.drawable_object_id, demi.clone())
            .unwrap();
        editor
            .set_slide_chart_font_size(0, chart.drawable_object_id, large)
            .unwrap();
        let duplicate = editor
            .duplicate_slide_chart(0, chart_selector(&editor, &chart))
            .unwrap();
        assert_eq!(
            editor
                .slide_chart_font(0, duplicate.drawable_object_id)
                .unwrap(),
            demi
        );
        assert_eq!(
            editor
                .slide_chart_font_size(0, duplicate.drawable_object_id)
                .unwrap(),
            large
        );
        let italic = ChartFont::named("AvenirNext-Italic")
            .unwrap()
            .with_italic(true);
        let larger = ChartFontSize::from_points(24.0).unwrap();
        editor
            .set_slide_chart_font(0, duplicate.drawable_object_id, italic.clone())
            .unwrap();
        editor
            .set_slide_chart_font_size(0, duplicate.drawable_object_id, larger)
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
        assert_eq!(
            editor
                .slide_chart_font_size(0, chart.drawable_object_id)
                .unwrap(),
            large
        );
        assert!(
            editor
                .reset_slide_chart_font(0, duplicate.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            editor
                .slide_chart_font_size(0, duplicate.drawable_object_id)
                .unwrap(),
            larger
        );
        assert!(
            editor
                .reset_slide_chart_font_size(0, duplicate.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            editor
                .slide_chart_font_size(0, duplicate.drawable_object_id)
                .unwrap(),
            inherited_size
        );
        editor
            .remove_slide_chart(0, chart_selector(&editor, &duplicate))
            .unwrap();
        assert!(
            editor
                .reset_slide_chart_font(0, chart.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            editor
                .slide_chart_font_size(0, chart.drawable_object_id)
                .unwrap(),
            large
        );
        assert!(
            editor
                .reset_slide_chart_font_size(0, chart.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            editor
                .slide_chart_font(0, chart.drawable_object_id)
                .unwrap(),
            inherited
        );
        assert_eq!(
            editor
                .slide_chart_font_size(0, chart.drawable_object_id)
                .unwrap(),
            inherited_size
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
