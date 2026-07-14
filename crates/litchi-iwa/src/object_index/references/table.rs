//! Application-specific object-reference extraction.

use super::super::ObjectIndex;
use crate::archive::RawMessage;

impl ObjectIndex {
    pub(super) fn extract_table_references(&mut self, object_id: u64, raw_msg: &RawMessage) {
        use prost::Message;

        match raw_msg.type_ {
            // TST (Table) types
            6000 | 6001 => {
                // TST.TableModelArchive contains multiple style and data references
                if let Ok(table) = crate::protobuf::tst::TableModelArchive::decode(&*raw_msg.data) {
                    // Extract style references
                    self.extract_reference(object_id, &table.table_style);
                    self.extract_reference(object_id, &table.body_text_style);
                    self.extract_reference(object_id, &table.header_row_text_style);
                    self.extract_reference(object_id, &table.header_column_text_style);
                    self.extract_reference(object_id, &table.footer_row_text_style);
                    self.extract_reference(object_id, &table.body_cell_style);
                    self.extract_reference(object_id, &table.header_row_style);
                    self.extract_reference(object_id, &table.header_column_style);
                    self.extract_reference(object_id, &table.footer_row_style);

                    // Extract optional style references
                    if let Some(ref table_name_style) = table.table_name_style {
                        self.extract_reference(object_id, table_name_style);
                    }
                    if let Some(ref table_name_shape_style) = table.table_name_shape_style {
                        self.extract_reference(object_id, table_name_shape_style);
                    }

                    // Extract data store sub-references
                    // DataStore contains references to column_headers, string_table, style_table, etc.
                    self.extract_reference(object_id, &table.data_store.column_headers);
                    self.extract_reference(object_id, &table.data_store.string_table);
                    self.extract_reference(object_id, &table.data_store.style_table);
                    self.extract_reference(object_id, &table.data_store.formula_table);
                    self.extract_reference(object_id, &table.data_store.format_table);

                    // Optional references
                    if let Some(ref formula_error_table) = table.data_store.formula_error_table {
                        self.extract_reference(object_id, formula_error_table);
                    }
                    if let Some(ref choice_list) =
                        table.data_store.multiple_choice_list_format_table
                    {
                        self.extract_reference(object_id, choice_list);
                    }
                    if let Some(ref merge_map) = table.data_store.merge_region_map {
                        self.extract_reference(object_id, merge_map);
                    }
                }
            },

            6005 | 6201 => {
                // TST.TableDataList - may contain references to other data structures
                // The actual cell data is stored here
            },

            _ => {},
        }
    }
}
