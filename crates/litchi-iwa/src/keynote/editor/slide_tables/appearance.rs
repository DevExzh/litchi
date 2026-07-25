//! Copy-on-write appearance CRUD for Keynote slide tables.

use super::*;
use crate::table_appearance::{
    TableAppearance, set_table_appearance as set_native_table_appearance,
    table_appearance as read_native_table_appearance,
};

impl KeynoteEditor {
    /// Read the effective alternating-row and automatic-sizing settings.
    pub fn slide_table_appearance(
        &self,
        slide_index: usize,
        model_object_id: u64,
    ) -> Result<TableAppearance> {
        require_table_model(self, slide_index, model_object_id)?;
        read_native_table_appearance(self.package(), model_object_id)
    }

    /// Replace appearance settings without mutating styles shared by other tables.
    pub fn set_slide_table_appearance(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        appearance: TableAppearance,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
        if self.slide_table_appearance(slide_index, model_object_id)? == appearance {
            return Ok(());
        }
        let mut staged = self.package().clone();
        set_native_table_appearance(&mut staged, model_object_id, appearance)?;
        let verified = Self::from_package(staged)?;
        if verified.slide_table_appearance(slide_index, model_object_id)? != appearance {
            return Err(Error::InvalidFormat(
                "Keynote table appearance failed round-trip validation".to_owned(),
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
    use crate::table_appearance::{TableRowBanding, TableRowSizing};

    #[test]
    fn scratch_table_appearance_is_copy_on_write() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let source = editor
            .add_slide_table(
                0,
                "Source",
                3,
                2,
                DrawablePoint { x: 40.0, y: 40.0 },
                DrawableSize {
                    width: 320.0,
                    height: 180.0,
                },
            )
            .unwrap();
        let duplicate = editor
            .duplicate_slide_table(0, source.drawable_object_id)
            .unwrap();
        let appearance = TableAppearance {
            row_banding: TableRowBanding::Enabled,
            row_sizing: TableRowSizing::FitCellContents,
        };

        editor
            .set_slide_table_appearance(0, duplicate.model_object_id, appearance)
            .unwrap();

        assert_eq!(
            editor
                .slide_table_appearance(0, source.model_object_id)
                .unwrap(),
            TableAppearance::default()
        );
        assert_eq!(
            editor
                .slide_table_appearance(0, duplicate.model_object_id)
                .unwrap(),
            appearance
        );
        assert_eq!(
            editor
                .slide_tables(0)
                .unwrap()
                .into_iter()
                .find(|table| table.model_object_id == duplicate.model_object_id)
                .unwrap()
                .appearance,
            appearance
        );
    }
}
