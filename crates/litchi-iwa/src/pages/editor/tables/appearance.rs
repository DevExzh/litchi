//! Copy-on-write appearance CRUD for Pages body tables.

use super::*;
use crate::table_appearance::{
    TableAppearance, set_table_appearance as set_native_table_appearance,
    table_appearance as read_native_table_appearance,
};

impl PagesEditor {
    /// Read the effective alternating-row and automatic-sizing settings.
    pub fn body_table_appearance(&self, model_object_id: u64) -> Result<TableAppearance> {
        self.require_body_table(model_object_id)?;
        read_native_table_appearance(self.package(), model_object_id)
    }

    /// Replace appearance settings without mutating styles shared by other tables.
    pub fn set_body_table_appearance(
        &mut self,
        model_object_id: u64,
        appearance: TableAppearance,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        if self.body_table_appearance(model_object_id)? == appearance {
            return Ok(());
        }
        let mut staged = self.package().clone();
        set_native_table_appearance(&mut staged, model_object_id, appearance)?;
        let verified = Self::from_package(staged)?;
        if verified.body_table_appearance(model_object_id)? != appearance {
            return Err(Error::InvalidFormat(
                "Pages table appearance failed round-trip validation".to_owned(),
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
    use crate::table_appearance::{
        TableGridlineVisibility, TableGridlines, TableRowBanding, TableRowSizing,
    };

    #[test]
    fn scratch_table_appearance_is_copy_on_write() {
        let body = "Appearance";
        let mut editor = PagesDocumentBuilder::new().body_text(body).build().unwrap();
        let source = editor
            .add_table(body.encode_utf16().count(), "Source", 3, 2)
            .unwrap();
        let anchor = editor.body_text().unwrap().encode_utf16().count();
        let duplicate = editor
            .duplicate_table(source.model_object_id, anchor)
            .unwrap();
        let appearance = TableAppearance {
            row_banding: TableRowBanding::Enabled,
            row_sizing: TableRowSizing::FitCellContents,
            gridlines: TableGridlines {
                body_horizontal: TableGridlineVisibility::Hidden,
                body_vertical: TableGridlineVisibility::Visible,
            },
        };

        editor
            .set_body_table_appearance(duplicate.model_object_id, appearance)
            .unwrap();

        assert_eq!(
            editor
                .body_table_appearance(source.model_object_id)
                .unwrap(),
            TableAppearance::default()
        );
        assert_eq!(
            editor
                .body_table_appearance(duplicate.model_object_id)
                .unwrap(),
            appearance
        );
        assert_eq!(
            editor
                .tables()
                .unwrap()
                .into_iter()
                .find(|table| table.model_object_id == duplicate.model_object_id)
                .unwrap()
                .appearance,
            appearance
        );
    }
}
