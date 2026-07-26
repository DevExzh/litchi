//! Typed hidden-row and hidden-column CRUD for Pages body tables.

use super::*;
use crate::table_hidden_axes::{
    TableHiddenAxes, set_table_hidden_axes as set_native_table_hidden_axes,
    table_hidden_axes as read_native_table_hidden_axes,
};

impl PagesEditor {
    /// Read the canonical user-hidden rows and columns of one body table.
    pub fn body_table_hidden_axes(&self, model_object_id: u64) -> Result<TableHiddenAxes> {
        self.require_body_table(model_object_id)?;
        read_native_table_hidden_axes(self.package(), model_object_id)
    }

    /// Replace all user-hidden rows and columns transactionally.
    pub fn set_body_table_hidden_axes(
        &mut self,
        model_object_id: u64,
        hidden: &TableHiddenAxes,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        if self.body_table_hidden_axes(model_object_id)? == *hidden {
            return Ok(());
        }
        let mut staged = self.package().clone();
        set_native_table_hidden_axes(&mut staged, model_object_id, hidden)?;
        let verified = Self::from_package(staged)?;
        if verified.body_table_hidden_axes(model_object_id)? != *hidden {
            return Err(Error::InvalidFormat(
                "Pages table hidden axes failed round-trip validation".to_owned(),
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
    use crate::table_hidden_axes::TableAxisIndex;

    #[test]
    fn scratch_table_hidden_axes_round_trip_transactionally() {
        let body = "Hidden axes";
        let mut editor = PagesDocumentBuilder::new().body_text(body).build().unwrap();
        let table = editor
            .add_table(body.encode_utf16().count(), "Table", 4, 3)
            .unwrap();
        let hidden =
            TableHiddenAxes::new([TableAxisIndex::row(2), TableAxisIndex::column(1)]).unwrap();

        editor
            .set_body_table_hidden_axes(table.model_object_id, &hidden)
            .unwrap();
        assert_eq!(
            editor
                .body_table_hidden_axes(table.model_object_id)
                .unwrap(),
            hidden
        );

        let before = editor.package().to_bytes().unwrap();
        let invalid = TableHiddenAxes::new([TableAxisIndex::column(3)]).unwrap();
        assert!(
            editor
                .set_body_table_hidden_axes(table.model_object_id, &invalid)
                .is_err()
        );
        assert_eq!(editor.package().to_bytes().unwrap(), before);
    }
}
