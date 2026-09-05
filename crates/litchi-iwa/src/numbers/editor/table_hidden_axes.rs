//! Typed hidden-row and hidden-column CRUD for Numbers tables.

use super::*;
use crate::table_hidden_axes::{
    set_table_hidden_axes as set_native_table_hidden_axes,
    table_hidden_axes as read_native_table_hidden_axes,
};
use litchi_iwa_common::table::axis::HiddenAxes;
use litchi_numbers::TableSelector;

impl NumbersEditor {
    /// Read the canonical user-hidden rows and columns of one attached table.
    pub fn table_hidden_axes(&self, selector: TableSelector<'_>) -> Result<HiddenAxes> {
        let table_id = super::selectors::table_id(self, selector)?;
        read_native_table_hidden_axes(&self.package, table_id)
    }

    /// Replace all user-hidden rows and columns transactionally.
    pub fn set_table_hidden_axes(
        &mut self,
        selector: TableSelector<'_>,
        hidden: &HiddenAxes,
    ) -> Result<()> {
        let table_id = super::selectors::table_id(self, selector)?;
        if self.table_hidden_axes(selector)? == *hidden {
            return Ok(());
        }
        let mut staged = self.package.clone();
        set_native_table_hidden_axes(&mut staged, table_id, hidden)?;
        let verified = Self::from_package(staged)?;
        if verified.table_hidden_axes(selector)? != *hidden {
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
    use crate::table_hidden_axes::table_hidden_graph_snapshot;
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

        editor
            .set_table_hidden_axes(test_table_selector(&editor, table_id), &hidden)
            .unwrap();
        assert_eq!(
            editor
                .table_hidden_axes(test_table_selector(&editor, table_id))
                .unwrap(),
            hidden
        );
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
        assert!(
            editor
                .set_table_hidden_axes(test_table_selector(&editor, table_id), &invalid)
                .is_err()
        );
        assert_eq!(editor.package.to_bytes().unwrap(), before);

        editor
            .set_table_hidden_axes(test_table_selector(&editor, table_id), &HiddenAxes::empty())
            .unwrap();
        assert!(
            editor
                .table_hidden_axes(test_table_selector(&editor, table_id))
                .unwrap()
                .is_empty()
        );
    }

    #[test]
    fn legacy_table_info_type_6003_keeps_canonical_model_and_axis_graph() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(4, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let hidden = HiddenAxes::new([AxisIndex::row(2), AxisIndex::column(1)]).unwrap();

        editor
            .set_table_hidden_axes(test_table_selector(&editor, table_id), &hidden)
            .unwrap();

        let graph = table_hidden_graph_snapshot(&editor.package, table_id).unwrap();
        // Source-built Numbers tables use the canonical type-6001 model while
        // retaining the canonical type-6000 TableInfo envelope.
        assert_eq!(graph.model_message_type, 6_001);
        assert_eq!(graph.info_message_type, 6_000);
        assert_eq!(graph.row_count, 4);
        assert_eq!(graph.column_count, 3);

        let info_archive = graph.info_archive.clone();
        let info_object_id = graph.info_object_id;
        let info_message_index = graph.info_message_index;
        editor
            .package
            .update_archive(&info_archive, |archive| {
                let object = archive.object_mut(info_object_id).ok_or_else(|| {
                    Error::InvalidFormat("Numbers table-info owner disappeared".to_owned())
                })?;
                let message = object.messages[info_message_index].clone();
                let data = crate::wire::patch_length_delimited_field(&message.data, 1, true, None)?;
                object.replace_message(info_message_index, RawMessage { type_: 6_003, data })?;
                Ok(())
            })
            .unwrap();

        let legacy_graph = table_hidden_graph_snapshot(&editor.package, table_id).unwrap();
        assert_eq!(legacy_graph.model_message_type, 6_001);
        assert_eq!(legacy_graph.info_message_type, 6_003);
        assert_eq!(legacy_graph.row_count, 4);
        assert_eq!(legacy_graph.column_count, 3);
        assert_eq!(
            read_native_table_hidden_axes(&editor.package, table_id).unwrap(),
            hidden
        );
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
                test_table_selector(&editor, table_id),
                &HiddenAxes::new([AxisIndex::row(2), AxisIndex::column(1)]).unwrap(),
            )
            .unwrap();

        editor
            .insert_table_row(
                test_table_selector(&editor, table_id),
                RowInsertion::Body { index: 0 },
            )
            .unwrap();
        editor
            .insert_table_column(
                test_table_selector(&editor, table_id),
                ColumnInsertion::Body { index: 0 },
            )
            .unwrap();
        assert_eq!(
            editor
                .table_hidden_axes(test_table_selector(&editor, table_id))
                .unwrap(),
            HiddenAxes::new([AxisIndex::row(3), AxisIndex::column(2),]).unwrap()
        );

        editor
            .remove_table_row(
                test_table_selector(&editor, table_id),
                RowDeletion::Body { index: 0 },
            )
            .unwrap();
        editor
            .remove_table_column(
                test_table_selector(&editor, table_id),
                ColumnDeletion::Body { index: 0 },
            )
            .unwrap();
        assert_eq!(
            editor
                .table_hidden_axes(test_table_selector(&editor, table_id))
                .unwrap(),
            HiddenAxes::new([AxisIndex::row(2), AxisIndex::column(1),]).unwrap()
        );

        editor
            .remove_table_row(
                test_table_selector(&editor, table_id),
                RowDeletion::Body { index: 1 },
            )
            .unwrap();
        editor
            .remove_table_column(
                test_table_selector(&editor, table_id),
                ColumnDeletion::Body { index: 0 },
            )
            .unwrap();
        assert!(
            editor
                .table_hidden_axes(test_table_selector(&editor, table_id))
                .unwrap()
                .is_empty()
        );
    }
}
