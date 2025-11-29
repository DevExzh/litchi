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
    pub fn extract_table_from_object(
        &self,
        object: &ResolvedObject,
    ) -> Result<Option<NumbersTable>> {
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
            let (offset, bytes_read) = Self::decode_varint(&offsets_buffer[pos..])?;
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
    fn decode_varint(data: &[u8]) -> Result<(usize, usize)> {
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
                // String cells can store text directly or reference the string table
                if let Some(ref string_val) = cell.string_value
                    && !string_val.is_empty()
                {
                    Ok(CellValue::Text(string_val.clone()))
                } else {
                    // In some cases, strings are stored via references
                    // The actual string data would be in the string_table
                    // For a production implementation, we would:
                    // 1. Check cell.text_style or cell.cell_style for a reference
                    // 2. Resolve that reference to find the actual string value
                    // 3. Look up the string in the string_table using a key
                    //
                    // For now, return Empty if no direct string value is present
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
    ///
    ///   - Reconstructs formula text from Abstract Syntax Tree
    ///   - Handles operators, functions, cell references, and constants
    ///   - Based on TSCE.ASTNodeArrayArchive protobuf structure
    ///   - Implements reverse-polish notation to infix conversion
    ///
    /// iWork stores formulas as Abstract Syntax Trees (AST) in reverse-polish
    /// notation (postfix). This function reconstructs the formula text by
    /// traversing the AST and converting it to standard infix notation.
    ///
    /// # Performance
    ///
    /// O(n) where n is the number of AST nodes. Uses a stack-based algorithm
    /// for efficient conversion.
    fn extract_formula_string(&self, formula: &tsce::FormulaArchive) -> Result<String> {
        use crate::iwa::protobuf::tsce::ast_node_array_archive::AstNodeType;

        let ast_array = &formula.ast_node_array;

        // Formulas are stored in reverse-polish notation (postfix)
        // We need to convert to infix notation using a stack
        if ast_array.ast_node.is_empty() {
            return Ok("=".to_string());
        }

        // Stack to hold expression parts during reconstruction
        let mut expr_stack: Vec<String> = Vec::new();

        // Process each AST node
        for node in &ast_array.ast_node {
            let ast_node_type = node.ast_node_type();

            match ast_node_type {
                // Arithmetic operators (binary)
                AstNodeType::AdditionNode => {
                    if expr_stack.len() >= 2 {
                        let right = expr_stack.pop().unwrap();
                        let left = expr_stack.pop().unwrap();
                        expr_stack.push(format!("({}+{})", left, right));
                    }
                },
                AstNodeType::SubtractionNode => {
                    if expr_stack.len() >= 2 {
                        let right = expr_stack.pop().unwrap();
                        let left = expr_stack.pop().unwrap();
                        expr_stack.push(format!("({}-{})", left, right));
                    }
                },
                AstNodeType::MultiplicationNode => {
                    if expr_stack.len() >= 2 {
                        let right = expr_stack.pop().unwrap();
                        let left = expr_stack.pop().unwrap();
                        expr_stack.push(format!("({}*{})", left, right));
                    }
                },
                AstNodeType::DivisionNode => {
                    if expr_stack.len() >= 2 {
                        let right = expr_stack.pop().unwrap();
                        let left = expr_stack.pop().unwrap();
                        expr_stack.push(format!("({}/{})", left, right));
                    }
                },
                AstNodeType::PowerNode => {
                    if expr_stack.len() >= 2 {
                        let right = expr_stack.pop().unwrap();
                        let left = expr_stack.pop().unwrap();
                        expr_stack.push(format!("({}^{})", left, right));
                    }
                },

                // Note: Comparison operators are handled differently in Numbers AST
                // They're not separate node types but may be represented through function nodes
                // For simplicity, we skip explicit handling here and rely on function dispatch

                // Constants
                AstNodeType::NumberNode => {
                    if let Some(number) = node.ast_number_node_number {
                        expr_stack.push(number.to_string());
                    }
                },
                AstNodeType::StringNode => {
                    if let Some(ref string) = node.ast_string_node_string {
                        expr_stack.push(format!("\"{}\"", string));
                    }
                },
                AstNodeType::BooleanNode => {
                    if let Some(boolean) = node.ast_boolean_node_boolean {
                        expr_stack.push(if boolean { "TRUE" } else { "FALSE" }.to_string());
                    }
                },

                // Cell references
                AstNodeType::CellReferenceNode => {
                    if let Some(ref cell_ref) = node.ast_local_cell_reference_node_reference {
                        // Convert row/column handles to A1 notation
                        let col_letter = self.column_index_to_letter(cell_ref.column_handle);
                        let row_num = cell_ref.row_handle + 1; // 0-based to 1-based
                        let col_sticky = if cell_ref.column_is_sticky != 0 {
                            "$"
                        } else {
                            ""
                        };
                        let row_sticky = if cell_ref.row_is_sticky != 0 { "$" } else { "" };
                        expr_stack.push(format!(
                            "{}{}{}{}",
                            col_sticky, col_letter, row_sticky, row_num
                        ));
                    } else if let Some(ref cross_ref) =
                        node.ast_cross_table_cell_reference_node_reference
                    {
                        // Cross-table reference
                        let col_letter = self.column_index_to_letter(cross_ref.column_handle);
                        let row_num = cross_ref.row_handle + 1;
                        expr_stack.push(format!("{}::{}{}", "Table", col_letter, row_num));
                    }
                },

                // Functions
                AstNodeType::FunctionNode => {
                    if let Some(function_index) = node.ast_function_node_index {
                        let num_args = node.ast_function_node_num_args.unwrap_or(0);
                        let function_name = self.get_function_name(function_index);

                        // Pop arguments from stack (in reverse order)
                        let mut args = Vec::new();
                        for _ in 0..num_args {
                            if let Some(arg) = expr_stack.pop() {
                                args.push(arg);
                            }
                        }
                        args.reverse();

                        let args_str = args.join(",");
                        expr_stack.push(format!("{}({})", function_name, args_str));
                    }
                },

                // List (for function arguments)
                AstNodeType::ListNode => {
                    if let Some(num_args) = node.ast_list_node_num_args {
                        // Collect arguments
                        let mut args = Vec::new();
                        for _ in 0..num_args {
                            if let Some(arg) = expr_stack.pop() {
                                args.push(arg);
                            }
                        }
                        args.reverse();
                        expr_stack.push(args.join(","));
                    }
                },

                // Unary operators - represented differently in the AST
                // Numbers uses NegationNode instead of UnaryMinusNode
                AstNodeType::NegationNode => {
                    if let Some(operand) = expr_stack.pop() {
                        expr_stack.push(format!("-({})", operand));
                    }
                },

                // Concatenation
                AstNodeType::ConcatenationNode => {
                    if expr_stack.len() >= 2 {
                        let right = expr_stack.pop().unwrap();
                        let left = expr_stack.pop().unwrap();
                        expr_stack.push(format!("({}&{})", left, right));
                    }
                },

                // Other node types - handle gracefully
                _ => {
                    // Unknown or special node types - keep processing
                    // (e.g., whitespace nodes, thunk nodes, etc.)
                },
            }
        }

        // The final result should be on top of the stack
        let result = if expr_stack.is_empty() {
            "=FORMULA()".to_string()
        } else {
            format!("={}", expr_stack.pop().unwrap())
        };

        Ok(result)
    }

    /// Convert column index to Excel-style letter (0 -> A, 1 -> B, ..., 25 -> Z, 26 -> AA)
    fn column_index_to_letter(&self, index: u32) -> String {
        let mut result = String::new();
        let mut idx = index;

        loop {
            let remainder = idx % 26;
            result.insert(0, (b'A' + remainder as u8) as char);
            if idx < 26 {
                break;
            }
            idx = idx / 26 - 1;
        }

        result
    }

    /// Get function name from function index
    /// Based on Numbers built-in function list
    fn get_function_name(&self, index: u32) -> String {
        // Common function indices (based on analysis of Numbers documents)
        // This mapping comes from observing Numbers files and documentation
        match index {
            0 => "SUM".to_string(),
            1 => "AVERAGE".to_string(),
            2 => "COUNT".to_string(),
            3 => "MAX".to_string(),
            4 => "MIN".to_string(),
            5 => "PRODUCT".to_string(),
            6 => "IF".to_string(),
            7 => "AND".to_string(),
            8 => "OR".to_string(),
            9 => "NOT".to_string(),
            10 => "ROUND".to_string(),
            11 => "SQRT".to_string(),
            12 => "ABS".to_string(),
            13 => "CONCATENATE".to_string(),
            14 => "LEFT".to_string(),
            15 => "RIGHT".to_string(),
            16 => "MID".to_string(),
            17 => "LEN".to_string(),
            18 => "UPPER".to_string(),
            19 => "LOWER".to_string(),
            20 => "PROPER".to_string(),
            21 => "TRIM".to_string(),
            22 => "SUBSTITUTE".to_string(),
            23 => "FIND".to_string(),
            24 => "SEARCH".to_string(),
            25 => "NOW".to_string(),
            26 => "TODAY".to_string(),
            27 => "DATE".to_string(),
            28 => "TIME".to_string(),
            29 => "YEAR".to_string(),
            30 => "MONTH".to_string(),
            31 => "DAY".to_string(),
            32 => "HOUR".to_string(),
            33 => "MINUTE".to_string(),
            34 => "SECOND".to_string(),
            35 => "WEEKDAY".to_string(),
            36 => "VLOOKUP".to_string(),
            37 => "HLOOKUP".to_string(),
            38 => "INDEX".to_string(),
            39 => "MATCH".to_string(),
            40 => "CHOOSE".to_string(),
            // More functions exist, but these are the most common
            _ => format!("FUNC{}", index),
        }
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
        // Test single-byte varint: 0
        let data = vec![0x00];
        let (value, bytes) = TableDataExtractor::decode_varint(&data).unwrap();
        assert_eq!(value, 0);
        assert_eq!(bytes, 1);

        // Test single-byte varint: 127
        let data = vec![0x7F];
        let (value, bytes) = TableDataExtractor::decode_varint(&data).unwrap();
        assert_eq!(value, 127);
        assert_eq!(bytes, 1);

        // Test two-byte varint: 300 = 0b100101100
        // Encoded as: 0xAC 0x02 (10101100 00000010)
        let data = vec![0xAC, 0x02];
        let (value, bytes) = TableDataExtractor::decode_varint(&data).unwrap();
        assert_eq!(value, 300);
        assert_eq!(bytes, 2);
    }

    #[test]
    fn test_cell_value_parsing() {
        // Test would require actual protobuf messages
        // Placeholder test
    }
}
