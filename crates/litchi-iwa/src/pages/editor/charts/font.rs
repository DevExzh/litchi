//! Chart-wide font CRUD for Pages body charts.

use super::*;
use crate::charts::font::{
    chart_font as read_native_font, chart_font_size as read_native_font_size,
    reset_chart_font as reset_native_font, reset_chart_font_size as reset_native_font_size,
    set_chart_font as set_native_font, set_chart_font_size as set_native_font_size,
};
use crate::charts::{ChartFont, ChartFontSize};

impl PagesEditor {
    /// Read the uniform effective font used by one body chart.
    pub fn body_chart_font(&self, drawable_object_id: u64) -> Result<ChartFont> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        read_native_font(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
        )
    }

    /// Set the uniform font used by every semantic text slot in one body chart.
    pub fn set_body_chart_font(&mut self, drawable_object_id: u64, font: ChartFont) -> Result<()> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_native_font(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            &font,
        )?;
        let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.body_chart_font(drawable_object_id)? != font {
            return Err(Error::InvalidFormat(
                "Pages chart font update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Reset crate-owned chart font overrides to their inherited theme values.
    pub fn reset_body_chart_font(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        if !reset_native_font(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Pages",
        )? {
            return Ok(false);
        }
        *self = PagesEditor::from_bytes(&staged.to_bytes()?)?;
        Ok(true)
    }

    /// Read the uniform effective point size used by one body chart.
    pub fn body_chart_font_size(&self, drawable_object_id: u64) -> Result<ChartFontSize> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        read_native_font_size(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
        )
    }

    /// Set the uniform point size used by every semantic text slot in one chart.
    pub fn set_body_chart_font_size(
        &mut self,
        drawable_object_id: u64,
        size: ChartFontSize,
    ) -> Result<()> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_native_font_size(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            size,
        )?;
        let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.body_chart_font_size(drawable_object_id)? != size {
            return Err(Error::InvalidFormat(
                "Pages chart font-size update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Reset crate-owned point-size overrides to the inherited theme size.
    pub fn reset_body_chart_font_size(&mut self, drawable_object_id: u64) -> Result<bool> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        if !reset_native_font_size(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Pages",
        )? {
            return Ok(false);
        }
        *self = PagesEditor::from_bytes(&staged.to_bytes()?)?;
        Ok(true)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{ChartData, ChartKind};
    use crate::pages::PagesDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_document_supports_chart_font_crud_and_copy_on_write() {
        let body = "Chart fonts";
        let mut editor = PagesDocumentBuilder::new().body_text(body).build().unwrap();
        let chart = editor
            .add_body_chart(
                body.encode_utf16().count(),
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
        let inherited = editor.body_chart_font(chart.drawable_object_id).unwrap();
        let inherited_size = editor
            .body_chart_font_size(chart.drawable_object_id)
            .unwrap();
        let demi = ChartFont::named("AvenirNext-DemiBold")
            .unwrap()
            .with_bold(true);
        let large = ChartFontSize::from_points(18.0).unwrap();
        editor
            .set_body_chart_font(chart.drawable_object_id, demi.clone())
            .unwrap();
        editor
            .set_body_chart_font_size(chart.drawable_object_id, large)
            .unwrap();
        let duplicate = editor
            .duplicate_body_chart(
                chart.drawable_object_id,
                editor.body_text().unwrap().encode_utf16().count(),
            )
            .unwrap();
        assert_eq!(
            editor
                .body_chart_font(duplicate.drawable_object_id)
                .unwrap(),
            demi
        );
        assert_eq!(
            editor
                .body_chart_font_size(duplicate.drawable_object_id)
                .unwrap(),
            large
        );
        let italic = ChartFont::named("AvenirNext-Italic")
            .unwrap()
            .with_italic(true);
        let larger = ChartFontSize::from_points(24.0).unwrap();
        editor
            .set_body_chart_font(duplicate.drawable_object_id, italic.clone())
            .unwrap();
        editor
            .set_body_chart_font_size(duplicate.drawable_object_id, larger)
            .unwrap();
        assert_eq!(
            editor.body_chart_font(chart.drawable_object_id).unwrap(),
            demi
        );
        assert_eq!(
            editor
                .body_chart_font(duplicate.drawable_object_id)
                .unwrap(),
            italic
        );
        assert_eq!(
            editor
                .body_chart_font_size(chart.drawable_object_id)
                .unwrap(),
            large
        );
        assert!(
            editor
                .reset_body_chart_font(duplicate.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            editor
                .body_chart_font_size(duplicate.drawable_object_id)
                .unwrap(),
            larger
        );
        assert!(
            editor
                .reset_body_chart_font_size(duplicate.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            editor
                .body_chart_font_size(duplicate.drawable_object_id)
                .unwrap(),
            inherited_size
        );
        editor
            .remove_body_chart(duplicate.drawable_object_id)
            .unwrap();
        assert!(
            editor
                .reset_body_chart_font(chart.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            editor
                .body_chart_font_size(chart.drawable_object_id)
                .unwrap(),
            large
        );
        assert!(
            editor
                .reset_body_chart_font_size(chart.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            editor.body_chart_font(chart.drawable_object_id).unwrap(),
            inherited
        );
        assert_eq!(
            editor
                .body_chart_font_size(chart.drawable_object_id)
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
