//! Native table lock CRUD for Numbers tables.

use super::*;
use crate::table_lock::{
    TableLockState, set_table_lock_state_for_model as set_native_table_lock_state,
    table_lock_state_for_model as read_native_table_lock_state,
};

impl NumbersEditor {
    /// Read one table's interactive lock state.
    pub fn table_lock_state(&self, table_id: u64) -> Result<TableLockState> {
        let (drawable_id, archive_name) = table_lock_context(&self.package, table_id)?;
        read_native_table_lock_state(&self.package, &archive_name, drawable_id, table_id)
    }

    /// Set one table's interactive lock state transactionally.
    pub fn set_table_lock_state(&mut self, table_id: u64, state: TableLockState) -> Result<()> {
        if self.table_lock_state(table_id)? == state {
            return Ok(());
        }
        let (drawable_id, archive_name) = table_lock_context(&self.package, table_id)?;
        let mut staged = self.package.clone();
        set_native_table_lock_state(&mut staged, &archive_name, drawable_id, table_id, state)?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.table_lock_state(table_id)? != state {
            return Err(Error::InvalidFormat(
                "Numbers table lock update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

fn table_lock_context(package: &IWorkPackage, table_id: u64) -> Result<(u64, String)> {
    let descriptor = table_models(package)?
        .into_iter()
        .find(|descriptor| descriptor.object_id == table_id)
        .ok_or_else(|| Error::ParseError(format!("Numbers table {table_id} not found")))?;
    let archive_name = object_locations(package)?
        .remove(&descriptor.table_info_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table drawable {} is missing",
                descriptor.table_info_id
            ))
        })?;
    Ok((descriptor.table_info_id, archive_name))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::NumbersDocumentBuilder;

    #[test]
    fn scratch_spreadsheet_supports_table_lock_crud() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_name("Locked Table")
            .table_dimensions(3, 2)
            .build()
            .unwrap();
        let table = editor.tables().unwrap().remove(0);
        let baseline = editor.to_bytes().unwrap();
        assert_eq!(
            editor.table_lock_state(table.object_id).unwrap(),
            TableLockState::Unlocked
        );
        assert_eq!(
            editor.tables().unwrap()[0].lock_state,
            TableLockState::Unlocked
        );

        editor
            .set_table_lock_state(table.object_id, TableLockState::Locked)
            .unwrap();
        let duplicate = editor.duplicate_table(table.object_id).unwrap();
        assert_eq!(
            editor.table_lock_state(duplicate.object_id).unwrap(),
            TableLockState::Locked
        );

        editor
            .set_table_lock_state(duplicate.object_id, TableLockState::Unlocked)
            .unwrap();
        assert_eq!(
            editor.table_lock_state(table.object_id).unwrap(),
            TableLockState::Locked
        );
        editor.remove_table(duplicate.object_id).unwrap();
        editor
            .set_table_lock_state(table.object_id, TableLockState::Unlocked)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }
}
