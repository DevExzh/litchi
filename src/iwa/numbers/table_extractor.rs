//! Table Data Extraction from TST Protobuf Messages
//!
//! This module provides utilities for extracting cell data from Numbers table structures.
//! Numbers stores table data in a complex format using Tiles, TableDataList, and Cell messages.
//!
//! ## Architecture
//!
//! - **TableModelArchive**: Contains table metadata and references to data stores
//! - **DataStore**: Contains references to various data tables (strings, formulas, styles)
//! - **TableDataList**: Maps keys to actual cell content (strings, formulas, formats)
//! - **TileStorage**: Contains the actual cells in a sparse tile-based structure
//! - **Tile**: Contains rows of cells with their values
//!
//! ## Example
//!
//! ```rust,ignore
//! use litchi::iwa::numbers::table_extractor::TableDataExtractor;
//! use litchi::iwa::bundle::Bundle;
//! use litchi::iwa::object_index::ObjectIndex;
//!
//! let bundle = Bundle::open("document.numbers")?;
//! let index = ObjectIndex::from_bundle(&bundle)?;
//! let extractor = TableDataExtractor::new(&bundle, &index);
//!
//! let tables = extractor.extract_all_tables()?;
//! for table in tables {
//!     println!("Table: {}", table.name);
//!     println!("{}", table.to_csv());
//! }
//! ```

use super::cell::CellValue;
use super::table::NumbersTable;
use crate::iwa::bundle::Bundle;
use crate::iwa::object_index::{ObjectIndex, ResolvedObject};
use crate::iwa::protobuf::{tsce, tst};
use crate::iwa::{Error, Result};
use prost::Message;
use std::collections::HashMap;

/// Extractor for Numbers table data
pub struct TableDataExtractor<'a> {
    bundle: &'a Bundle,
    object_index: &'a ObjectIndex,
}

impl<'a> TableDataExtractor<'a> {
    /// Create a new table data extractor
    pub fn new(bundle: &'a Bundle, object_index: &'a ObjectIndex) -> Self {
        Self {
            bundle,
            object_index,
        }
    }

    /// Extract all tables from the document
    pub fn extract_all_tables(&self) -> Result<Vec<NumbersTable>> {
        let mut tables = Vec::new();

        // Find all TableModelArchive objects (message type 6000 or 6001)
        let table_entries = self.object_index.find_objects_by_type(6000);
        tables.extend(self.extract_tables_from_entries(table_entries)?);

        let table_entries = self.object_index.find_objects_by_type(6001);
        tables.extend(self.extract_tables_from_entries(table_entries)?);

        Ok(tables)
    }

    /// Extract tables from object index entries
    fn extract_tables_from_entries(
        &self,
        entries: Vec<&crate::iwa::object_index::ObjectIndexEntry>,
    ) -> Result<Vec<NumbersTable>> {
        let mut tables = Vec::new();

        for entry in entries {
            if let Some(resolved) = self.object_index.resolve_object(self.bundle, entry.id)?
                && let Some(table) = self.extract_table_from_object(&resolved)?
            {
                tables.push(table);
            }
        }

        Ok(tables)
    }

    /// Extract a single table from a resolved object
    fn extract_table_from_object(&self, object: &ResolvedObject) -> Result<Option<NumbersTable>> {
        // Find the TableModelArchive message
        for msg in &object.messages {
            if (msg.type_ == 6000 || msg.type_ == 6001)
                && let Ok(table_model) = tst::TableModelArchive::decode(&*msg.data)
            {
                return self.parse_table_model(table_model).map(Some);
            }
        }

        Ok(None)
    }

    /// Parse a TableModelArchive protobuf message
    fn parse_table_model(&self, table_model: tst::TableModelArchive) -> Result<NumbersTable> {
        let mut table = NumbersTable::new(table_model.table_name.clone());
        table.row_count = table_model.number_of_rows as usize;
        table.column_count = table_model.number_of_columns as usize;

        // Extract string table for cell text values
        // string_table is a required field, not Optional
        let string_table =
            self.load_table_data_list(table_model.data_store.string_table.identifier)?;

        // Extract formula table for formula cells
        // formula_table is a required field, not Optional
        let formula_table =
            self.load_table_data_list(table_model.data_store.formula_table.identifier)?;

        // Parse tiles to extract cell data
        self.parse_tiles(
            &table_model.data_store.tiles,
            &string_table,
            &formula_table,
            &mut table,
        )?;

        Ok(table)
    }

    /// Load a TableDataList from an object reference
    fn load_table_data_list(&self, object_id: u64) -> Result<HashMap<u32, String>> {
        let mut result = HashMap::new();

        if let Some(resolved) = self.object_index.resolve_object(self.bundle, object_id)? {
            for msg in &resolved.messages {
                // TableDataList has message types 6005, 6201
                if (msg.type_ == 6005 || msg.type_ == 6201)
                    && let Ok(data_list) = tst::TableDataList::decode(&*msg.data)
                {
                    for entry in data_list.entries {
                        if let Some(ref string_val) = entry.string {
                            result.insert(entry.key, string_val.clone());
                        }
                    }
                }
            }
        }

        Ok(result)
    }

    /// Parse tile storage to extract cells
    fn parse_tiles(
        &self,
        tile_storage: &tst::TileStorage,
        string_table: &HashMap<u32, String>,
        formula_table: &HashMap<u32, String>,
        table: &mut NumbersTable,
    ) -> Result<()> {
        // Resolve each tile reference and parse its contents
        for tile_ref in &tile_storage.tiles {
            // tile is a required field, not Optional
            let tile_reference = &tile_ref.tile;
            self.parse_tile(
                tile_reference.identifier,
                string_table,
                formula_table,
                table,
            )?;
        }

        Ok(())
    }

    /// Parse a single tile object
    fn parse_tile(
        &self,
        tile_id: u64,
        string_table: &HashMap<u32, String>,
        formula_table: &HashMap<u32, String>,
        table: &mut NumbersTable,
    ) -> Result<()> {
        if let Some(resolved) = self.object_index.resolve_object(self.bundle, tile_id)? {
            for msg in &resolved.messages {
                // Tile messages are typically in the TST namespace
                if let Ok(tile) = tst::Tile::decode(&*msg.data) {
                    self.parse_tile_rows(&tile, string_table, formula_table, table)?;
                }
            }
        }

        Ok(())
    }

    /// Parse rows within a tile
    fn parse_tile_rows(
        &self,
        tile: &tst::Tile,
        string_table: &HashMap<u32, String>,
        formula_table: &HashMap<u32, String>,
        table: &mut NumbersTable,
    ) -> Result<()> {
        for row_info in &tile.row_infos {
            self.parse_tile_row(row_info, string_table, formula_table, table)?;
        }

        Ok(())
    }

    /// Parse a single tile row
    fn parse_tile_row(
        &self,
        row_info: &tst::TileRowInfo,
        _string_table: &HashMap<u32, String>,
        _formula_table: &HashMap<u32, String>,
        table: &mut NumbersTable,
    ) -> Result<()> {
        let row_index = row_info.tile_row_index as usize;

        // The cell_storage_buffer contains serialized Cell messages
        // The cell_offsets buffer contains the byte offsets for each cell

        // Parse cell offsets (variable-length encoded)
        let offsets = self.parse_cell_offsets(&row_info.cell_offsets)?;

        // Parse each cell from the storage buffer
        for (col_index, (offset, next_offset)) in
            offsets.iter().zip(offsets.iter().skip(1)).enumerate()
        {
            let cell_data = &row_info.cell_storage_buffer[*offset..*next_offset];

            if let Ok(cell) = tst::Cell::decode(cell_data) {
                let cell_value = self.parse_cell(&cell)?;
                table.set_cell(row_index, col_index, cell_value);
            }
        }

        Ok(())
    }

    /// Parse cell offsets from the offsets buffer
    ///
    /// Offsets are stored as variable-length integers (varints).
    /// Each offset indicates the starting position of a cell in the storage buffer.
    fn parse_cell_offsets(&self, offsets_buffer: &[u8]) -> Result<Vec<usize>> {
        let mut offsets = vec![0]; // First cell always starts at offset 0
        let mut pos = 0;

        while pos < offsets_buffer.len() {
            let (offset, bytes_read) = self.decode_varint(&offsets_buffer[pos..])?;
            pos += bytes_read;

            // Offsets are cumulative
            let prev_offset = *offsets.last().unwrap_or(&0);
            offsets.push(prev_offset + offset);
        }

        Ok(offsets)
    }

    /// Decode a single varint from a byte slice
    ///
    /// Returns (value, bytes_consumed)
    fn decode_varint(&self, data: &[u8]) -> Result<(usize, usize)> {
        let mut result: u64 = 0;
        let mut shift = 0;
        let mut bytes_read = 0;

        for &byte in data.iter().take(10) {
            // Max 10 bytes for u64
            bytes_read += 1;
            result |= u64::from(byte & 0x7F) << shift;

            if byte & 0x80 == 0 {
                return Ok((result as usize, bytes_read));
            }

            shift += 7;
        }

        Err(Error::ParseError("Invalid varint encoding".to_string()))
    }

    /// Parse a single Cell protobuf message into a CellValue
    fn parse_cell(&self, cell: &tst::Cell) -> Result<CellValue> {
        use tst::CellValueType;

        match cell.value_type() {
            CellValueType::EmptyCellValueType => Ok(CellValue::Empty),

            CellValueType::NumberCellValueType => {
                let number = cell.number_value.unwrap_or(0.0);
                Ok(CellValue::Number(number))
            },

            CellValueType::StringCellValueType => {
                // String cells reference the string table
                if let Some(ref string_val) = cell.string_value {
                    Ok(CellValue::Text(string_val.clone()))
                } else {
                    // Try to look up in string table via cell_style or text_style reference
                    // This is a simplified approach; actual implementation may need more logic
                    Ok(CellValue::Empty)
                }
            },

            CellValueType::BoolCellValueType => {
                let bool_val = cell.bool_value.unwrap_or(false);
                Ok(CellValue::Boolean(bool_val))
            },

            CellValueType::DateCellValueType => {
                // Date values are stored as numbers (timestamp)
                let date_num = cell.number_value.unwrap_or(0.0);
                // Convert to string representation
                // Note: Apple's date epoch is different from Unix epoch
                Ok(CellValue::Date(format!("{}", date_num)))
            },

            CellValueType::DurationCellValueType => {
                let duration = cell.number_value.unwrap_or(0.0);
                Ok(CellValue::Duration(duration))
            },

            CellValueType::ErrorCellValueType => Ok(CellValue::Error("ERROR".to_string())),

            CellValueType::ProvidedCellValueType => {
                // Provided values may come from formulas or other sources
                if let Some(ref formula) = cell.formula {
                    // Extract formula string representation
                    let formula_str = self.extract_formula_string(formula)?;
                    Ok(CellValue::Formula(formula_str))
                } else {
                    Ok(CellValue::Empty)
                }
            },

            CellValueType::RichTextCellType => {
                // Rich text requires resolving the richTextPayload reference
                if let Some(ref payload_ref) = cell.rich_text_payload {
                    if let Some(text) = self.extract_rich_text(payload_ref.identifier)? {
                        Ok(CellValue::Text(text))
                    } else {
                        Ok(CellValue::Empty)
                    }
                } else {
                    Ok(CellValue::Empty)
                }
            },
        }
    }

    /// Extract formula string from FormulaArchive
    fn extract_formula_string(&self, _formula: &tsce::FormulaArchive) -> Result<String> {
        // Formula structure in iWork contains AST nodes, not direct text
        // The formula_text field doesn't exist in FormulaArchive
        // A full implementation would reconstruct the formula from ast_node_array
        // For now, return a placeholder indicating formula presence
        Ok("=FORMULA()".to_string())
    }

    /// Extract rich text from a storage reference
    fn extract_rich_text(&self, storage_id: u64) -> Result<Option<String>> {
        if let Some(resolved) = self.object_index.resolve_object(self.bundle, storage_id)? {
            // Look for TSWP.StorageArchive messages
            for msg in &resolved.messages {
                if msg.type_ >= 2001
                    && msg.type_ <= 2022
                    && let Ok(storage) =
                        crate::iwa::protobuf::tswp::StorageArchive::decode(&*msg.data)
                    && !storage.text.is_empty()
                {
                    return Ok(Some(storage.text.join("\n")));
                }
            }
        }

        Ok(None)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_varint_decoding() {
        let extractor = TableDataExtractor {
            bundle: unsafe { &*(std::ptr::null() as *const Bundle) },
            object_index: unsafe { &*(std::ptr::null() as *const ObjectIndex) },
        };

        // Test single-byte varint: 0
        let data = vec![0x00];
        let (value, bytes) = extractor.decode_varint(&data).unwrap();
        assert_eq!(value, 0);
        assert_eq!(bytes, 1);

        // Test single-byte varint: 127
        let data = vec![0x7F];
        let (value, bytes) = extractor.decode_varint(&data).unwrap();
        assert_eq!(value, 127);
        assert_eq!(bytes, 1);

        // Test two-byte varint: 300 = 0b100101100
        // Encoded as: 0xAC 0x02 (10101100 00000010)
        let data = vec![0xAC, 0x02];
        let (value, bytes) = extractor.decode_varint(&data).unwrap();
        assert_eq!(value, 300);
        assert_eq!(bytes, 2);
    }

    #[test]
    fn test_cell_value_parsing() {
        // Test would require actual protobuf messages
        // Placeholder test
    }
}
