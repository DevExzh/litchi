//! Chart-wide font CRUD for Numbers sheet charts.

use super::*;
use crate::charts::ChartFont;
use crate::charts::font::{
    chart_font as read_native_font, reset_chart_font as reset_native_font,
    set_chart_font as set_native_font,
};

impl NumbersEditor {
    /// Read the uniform effective font used by one sheet chart.
    pub fn sheet_chart_font(&self, sheet_id: u64, drawable_object_id: u64) -> Result<ChartFont> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        read_native_font(
            &self.package,
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
        )
    }

    /// Set the uniform font used by every semantic text slot in one sheet chart.
    pub fn set_sheet_chart_font(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        font: ChartFont,
    ) -> Result<()> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_native_font(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            &font,
        )?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.sheet_chart_font(sheet_id, drawable_object_id)? != font {
            return Err(Error::InvalidFormat(
                "Numbers chart font update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Reset crate-owned chart font overrides to their inherited theme values.
    pub fn reset_sheet_chart_font(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        if !reset_native_font(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
        )? {
            return Ok(false);
        }
        *self = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        Ok(true)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{ChartData, ChartKind};
    use crate::numbers::NumbersDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_spreadsheet_supports_chart_font_crud_and_copy_on_write() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let chart = editor
            .add_sheet_chart(
                sheet_id,
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
            .sheet_chart_font(sheet_id, chart.drawable_object_id)
            .unwrap();
        assert!(!inherited.bold());
        assert!(!inherited.italic());

        let demi = ChartFont::named("AvenirNext-DemiBold")
            .unwrap()
            .with_bold(true);
        editor
            .set_sheet_chart_font(sheet_id, chart.drawable_object_id, demi.clone())
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, chart.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_font(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            demi
        );

        let italic = ChartFont::named("AvenirNext-Italic")
            .unwrap()
            .with_italic(true);
        editor
            .set_sheet_chart_font(sheet_id, duplicate.drawable_object_id, italic.clone())
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_font(sheet_id, chart.drawable_object_id)
                .unwrap(),
            demi
        );
        assert_eq!(
            editor
                .sheet_chart_font(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            italic
        );

        assert!(
            editor
                .reset_sheet_chart_font(sheet_id, duplicate.drawable_object_id)
                .unwrap()
        );
        editor
            .remove_sheet_chart(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert!(
            editor
                .reset_sheet_chart_font(sheet_id, chart.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            editor
                .sheet_chart_font(sheet_id, chart.drawable_object_id)
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
