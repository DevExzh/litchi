//! Native table lock CRUD for Pages body tables.

use super::*;
use crate::table_lock::{
    set_table_lock_state as set_native_table_lock_state,
    table_lock_state as read_native_table_lock_state,
};
use litchi_iwa_common::table::lock::State as TableLockState;

impl PagesEditor {
    /// Read one body table's interactive lock state.
    pub fn body_table_lock_state(&self, model_object_id: u64) -> Result<TableLockState> {
        let graph = self.require_body_table(model_object_id)?;
        read_native_table_lock_state(
            self.package(),
            &graph.drawable_archive,
            graph.info.drawable_object_id,
            "Pages",
        )
    }

    /// Set one body table's interactive lock state transactionally.
    pub fn set_body_table_lock_state(
        &mut self,
        model_object_id: u64,
        state: TableLockState,
    ) -> Result<()> {
        if self.body_table_lock_state(model_object_id)? == state {
            return Ok(());
        }
        let graph = self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        set_native_table_lock_state(
            &mut staged,
            &graph.drawable_archive,
            graph.info.drawable_object_id,
            "Pages",
            state,
        )?;
        let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.body_table_lock_state(model_object_id)? != state {
            return Err(Error::InvalidFormat(
                "Pages table lock update failed validation".to_owned(),
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

    #[test]
    fn scratch_document_supports_table_lock_crud() {
        let body = "Table lock";
        let mut editor = PagesDocumentBuilder::new().body_text(body).build().unwrap();
        let table = editor
            .add_table(body.encode_utf16().count(), "Locked Table", 3, 2)
            .unwrap();
        let baseline = editor.to_bytes().unwrap();
        assert_eq!(
            editor.body_table_lock_state(table.model_object_id).unwrap(),
            TableLockState::Unlocked
        );
        assert_eq!(
            editor.tables().unwrap()[0].lock_state,
            TableLockState::Unlocked
        );

        editor
            .set_body_table_lock_state(table.model_object_id, TableLockState::Locked)
            .unwrap();
        let anchor = editor.body_text().unwrap().encode_utf16().count();
        let duplicate = editor
            .duplicate_table(table.model_object_id, anchor)
            .unwrap();
        assert_eq!(
            editor
                .body_table_lock_state(duplicate.model_object_id)
                .unwrap(),
            TableLockState::Locked
        );

        editor
            .set_body_table_lock_state(duplicate.model_object_id, TableLockState::Unlocked)
            .unwrap();
        assert_eq!(
            editor.body_table_lock_state(table.model_object_id).unwrap(),
            TableLockState::Locked
        );
        editor.remove_table(duplicate.model_object_id).unwrap();
        editor
            .set_body_table_lock_state(table.model_object_id, TableLockState::Unlocked)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }
}
