//! Native table lock CRUD for Keynote slides.

use super::*;
use crate::table_lock::{
    TableLockState, set_table_lock_state as set_native_table_lock_state,
    table_lock_state as read_native_table_lock_state,
};

impl KeynoteEditor {
    /// Read one slide table's interactive lock state.
    pub fn slide_table_lock_state(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<TableLockState> {
        let graph = slide_table_graph(self, slide_index, drawable_object_id)?;
        read_native_table_lock_state(
            self.package(),
            &graph.table_archive,
            drawable_object_id,
            "Keynote",
        )
    }

    /// Set one slide table's interactive lock state transactionally.
    pub fn set_slide_table_lock_state(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        state: TableLockState,
    ) -> Result<()> {
        if self.slide_table_lock_state(slide_index, drawable_object_id)? == state {
            return Ok(());
        }
        let graph = slide_table_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_native_table_lock_state(
            &mut staged,
            &graph.table_archive,
            drawable_object_id,
            "Keynote",
            state,
        )?;
        let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.slide_table_lock_state(slide_index, drawable_object_id)? != state {
            return Err(Error::InvalidFormat(
                "Keynote table lock update failed validation".to_owned(),
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

    #[test]
    fn scratch_presentation_supports_table_lock_crud() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let table = editor
            .add_slide_table(
                0,
                "Locked Table",
                3,
                2,
                DrawablePoint { x: 40.0, y: 40.0 },
                DrawableSize {
                    width: 320.0,
                    height: 180.0,
                },
            )
            .unwrap();
        let baseline = editor.to_bytes().unwrap();
        assert_eq!(
            editor
                .slide_table_lock_state(0, table.drawable_object_id)
                .unwrap(),
            TableLockState::Unlocked
        );
        assert_eq!(
            editor.slide_tables(0).unwrap()[0].lock_state,
            TableLockState::Unlocked
        );

        editor
            .set_slide_table_lock_state(0, table.drawable_object_id, TableLockState::Locked)
            .unwrap();
        editor
            .set_slide_table_lock_state(0, table.drawable_object_id, TableLockState::Unlocked)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        editor
            .set_slide_table_lock_state(0, table.drawable_object_id, TableLockState::Locked)
            .unwrap();
        let duplicate = editor
            .duplicate_slide_table(0, table.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .slide_table_lock_state(0, duplicate.drawable_object_id)
                .unwrap(),
            TableLockState::Locked
        );

        editor
            .set_slide_table_lock_state(0, duplicate.drawable_object_id, TableLockState::Unlocked)
            .unwrap();
        assert_eq!(
            editor
                .slide_table_lock_state(0, table.drawable_object_id)
                .unwrap(),
            TableLockState::Locked
        );
        editor
            .remove_slide_table(0, duplicate.drawable_object_id)
            .unwrap();
        editor
            .set_slide_table_lock_state(0, table.drawable_object_id, TableLockState::Unlocked)
            .unwrap();
        assert_eq!(
            editor
                .slide_table_lock_state(0, table.drawable_object_id)
                .unwrap(),
            TableLockState::Unlocked
        );
    }
}
