//! Chart-wide font CRUD for Pages body charts.

use super::*;
use crate::charts::ChartFont;
use crate::charts::font::{
    chart_font as read_native_font, reset_chart_font as reset_native_font,
    set_chart_font as set_native_font,
};

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
        let demi = ChartFont::named("AvenirNext-DemiBold")
            .unwrap()
            .with_bold(true);
        editor
            .set_body_chart_font(chart.drawable_object_id, demi.clone())
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
        let italic = ChartFont::named("AvenirNext-Italic")
            .unwrap()
            .with_italic(true);
        editor
            .set_body_chart_font(duplicate.drawable_object_id, italic.clone())
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
        assert!(
            editor
                .reset_body_chart_font(duplicate.drawable_object_id)
                .unwrap()
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
            editor.body_chart_font(chart.drawable_object_id).unwrap(),
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
