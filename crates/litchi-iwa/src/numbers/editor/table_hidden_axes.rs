//! Typed hidden-row and hidden-column CRUD for Numbers tables.

use super::*;
use crate::table_hidden_axes::{
    set_table_hidden_axes as set_native_table_hidden_axes,
    table_hidden_axes as read_native_table_hidden_axes,
};
use litchi_iwa_common::table::axis::HiddenAxes;

impl NumbersEditor {
    /// Read the canonical user-hidden rows and columns of one attached table.
    pub fn table_hidden_axes(&self, table_id: u64) -> Result<HiddenAxes> {
        read_native_table_hidden_axes(&self.package, table_id)
    }

    /// Replace all user-hidden rows and columns transactionally.
    pub fn set_table_hidden_axes(&mut self, table_id: u64, hidden: &HiddenAxes) -> Result<()> {
        if self.table_hidden_axes(table_id)? == *hidden {
            return Ok(());
        }
        let mut staged = self.package.clone();
        set_native_table_hidden_axes(&mut staged, table_id, hidden)?;
        let verified = Self::from_package(staged)?;
        if verified.table_hidden_axes(table_id)? != *hidden {
            return Err(Error::InvalidFormat(
                "Numbers table hidden axes failed round-trip validation".to_owned(),
            ));
        }
        self.package = verified.package;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::NumbersDocumentBuilder;
    use crate::table_hidden_axes::{
        FILTER_SET_MESSAGE_TYPE, HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE,
    };
    use litchi_iwa_common::table::axis::{AxisIndex, HiddenAxes};

    #[test]
    fn scratch_table_hidden_axes_round_trip_transactionally() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(4, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let hidden = HiddenAxes::new([AxisIndex::row(2), AxisIndex::column(1)]).unwrap();

        editor.set_table_hidden_axes(table_id, &hidden).unwrap();
        assert_eq!(editor.table_hidden_axes(table_id).unwrap(), hidden);
        let mut filter_sets = 0;
        let mut formula_owners = 0;
        for archive_name in editor.package.iwa_entry_names() {
            for message in editor
                .package
                .archive(archive_name)
                .unwrap()
                .objects
                .into_iter()
                .flat_map(|object| object.messages)
            {
                filter_sets += usize::from(message.type_ == FILTER_SET_MESSAGE_TYPE);
                formula_owners +=
                    usize::from(message.type_ == HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE);
            }
        }
        assert_eq!((filter_sets, formula_owners), (2, 2));

        let before = editor.package.to_bytes().unwrap();
        let invalid = HiddenAxes::new([AxisIndex::row(4)]).unwrap();
        assert!(editor.set_table_hidden_axes(table_id, &invalid).is_err());
        assert_eq!(editor.package.to_bytes().unwrap(), before);

        editor
            .set_table_hidden_axes(table_id, &HiddenAxes::empty())
            .unwrap();
        assert!(editor.table_hidden_axes(table_id).unwrap().is_empty());
    }

    #[test]
    fn hidden_axes_follow_table_insertion_and_deletion() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(4, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        editor
            .set_table_hidden_axes(
                table_id,
                &HiddenAxes::new([AxisIndex::row(2), AxisIndex::column(1)]).unwrap(),
            )
            .unwrap();

        editor
            .insert_table_row(table_id, TableRowInsertion::Body { index: 0 })
            .unwrap();
        editor
            .insert_table_column(table_id, TableColumnInsertion::Body { index: 0 })
            .unwrap();
        assert_eq!(
            editor.table_hidden_axes(table_id).unwrap(),
            HiddenAxes::new([AxisIndex::row(3), AxisIndex::column(2),]).unwrap()
        );

        editor
            .remove_table_row(table_id, TableRowDeletion::Body { index: 0 })
            .unwrap();
        editor
            .remove_table_column(table_id, TableColumnDeletion::Body { index: 0 })
            .unwrap();
        assert_eq!(
            editor.table_hidden_axes(table_id).unwrap(),
            HiddenAxes::new([AxisIndex::row(2), AxisIndex::column(1),]).unwrap()
        );

        editor
            .remove_table_row(table_id, TableRowDeletion::Body { index: 1 })
            .unwrap();
        editor
            .remove_table_column(table_id, TableColumnDeletion::Body { index: 0 })
            .unwrap();
        assert!(editor.table_hidden_axes(table_id).unwrap().is_empty());
    }
}
