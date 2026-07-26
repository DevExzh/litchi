//! Typed hidden-row and hidden-column CRUD for Keynote slide tables.

use super::*;
use crate::table_hidden_axes::{
    TableHiddenAxes, set_table_hidden_axes as set_native_table_hidden_axes,
    table_hidden_axes as read_native_table_hidden_axes,
};

impl KeynoteEditor {
    /// Read the canonical user-hidden rows and columns of one slide table.
    pub fn slide_table_hidden_axes(
        &self,
        slide_index: usize,
        model_object_id: u64,
    ) -> Result<TableHiddenAxes> {
        require_table_model(self, slide_index, model_object_id)?;
        read_native_table_hidden_axes(self.package(), model_object_id)
    }

    /// Replace all user-hidden rows and columns transactionally.
    pub fn set_slide_table_hidden_axes(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        hidden: &TableHiddenAxes,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
        if self.slide_table_hidden_axes(slide_index, model_object_id)? == *hidden {
            return Ok(());
        }
        let mut staged = self.package().clone();
        set_native_table_hidden_axes(&mut staged, model_object_id, hidden)?;
        let verified = Self::from_package(staged)?;
        if verified.slide_table_hidden_axes(slide_index, model_object_id)? != *hidden {
            return Err(Error::InvalidFormat(
                "Keynote table hidden axes failed round-trip validation".to_owned(),
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
    use crate::table_hidden_axes::TableAxisIndex;

    #[test]
    fn scratch_table_hidden_axes_round_trip_transactionally() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let table = editor
            .add_slide_table(
                0,
                "Table",
                4,
                3,
                DrawablePoint { x: 40.0, y: 40.0 },
                DrawableSize {
                    width: 320.0,
                    height: 180.0,
                },
            )
            .unwrap();
        let hidden =
            TableHiddenAxes::new([TableAxisIndex::row(2), TableAxisIndex::column(1)]).unwrap();

        editor
            .set_slide_table_hidden_axes(0, table.model_object_id, &hidden)
            .unwrap();
        assert_eq!(
            editor
                .slide_table_hidden_axes(0, table.model_object_id)
                .unwrap(),
            hidden
        );

        let before = editor.package().to_bytes().unwrap();
        let invalid = TableHiddenAxes::new([TableAxisIndex::row(4)]).unwrap();
        assert!(
            editor
                .set_slide_table_hidden_axes(0, table.model_object_id, &invalid)
                .is_err()
        );
        assert_eq!(editor.package().to_bytes().unwrap(), before);
    }
}
