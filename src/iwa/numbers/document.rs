//! Numbers Document Implementation
//!
//! Provides high-level API for working with Apple Numbers spreadsheets.

use std::path::Path;

use crate::iwa::Result;
use crate::iwa::bundle::Bundle;
use crate::iwa::object_index::ObjectIndex;
use crate::iwa::registry::Application;
use crate::iwa::text::TextExtractor;
use super::sheet::NumbersSheet;
use super::table::NumbersTable;

/// High-level interface for Numbers documents
pub struct NumbersDocument {
    /// Underlying bundle
    bundle: Bundle,
    /// Object index for cross-referencing
    object_index: ObjectIndex,
}

impl NumbersDocument {
    /// Open a Numbers document from a path
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::iwa::numbers::NumbersDocument;
    ///
    /// let doc = NumbersDocument::open("spreadsheet.numbers")?;
    /// println!("Loaded Numbers document");
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let bundle = Bundle::open(path)?;
        let object_index = ObjectIndex::from_bundle(&bundle)?;

        Ok(Self {
            bundle,
            object_index,
        })
    }

    /// Open a Numbers document from raw bytes
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::iwa::numbers::NumbersDocument;
    /// use std::fs;
    ///
    /// let data = fs::read("spreadsheet.numbers")?;
    /// let doc = NumbersDocument::from_bytes(&data)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let bundle = Bundle::from_bytes(bytes)?;
        let object_index = ObjectIndex::from_bundle(&bundle)?;

        Ok(Self {
            bundle,
            object_index,
        })
    }

    /// Extract all text content from the document
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::iwa::numbers::NumbersDocument;
    ///
    /// let doc = NumbersDocument::open("spreadsheet.numbers")?;
    /// let text = doc.text()?;
    /// println!("{}", text);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        let mut extractor = TextExtractor::new();
        extractor.extract_from_bundle(&self.bundle)?;
        Ok(extractor.get_text())
    }

    /// Extract sheets from the document
    ///
    /// Numbers documents consist of multiple sheets, each containing tables.
    /// This method parses the document structure and returns all sheets.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::iwa::numbers::NumbersDocument;
    ///
    /// let doc = NumbersDocument::open("spreadsheet.numbers")?;
    /// let sheets = doc.sheets()?;
    ///
    /// for sheet in sheets {
    ///     println!("Sheet: {}", sheet.name);
    ///     for table in &sheet.tables {
    ///         println!("  Table: {} ({}x{})", 
    ///             table.name, table.row_count, table.column_count);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn sheets(&self) -> Result<Vec<NumbersSheet>> {
        let mut sheets = Vec::new();

        // Find sheet archives (message type 2 is TN.SheetArchive, type 1003 in our decoder)
        let sheet_objects = self.bundle.find_objects_by_type(1003);

        if sheet_objects.is_empty() {
            // Try alternate sheet message type (TN.SheetArchive from JSON)
            let alt_sheet_objects = self.bundle.find_objects_by_type(2);
            
            for (index, (_archive_name, object)) in alt_sheet_objects.iter().enumerate() {
                let sheet = self.parse_sheet(index, object)?;
                if !sheet.is_empty() || !sheet.name.is_empty() {
                    sheets.push(sheet);
                }
            }
        } else {
            for (index, (_archive_name, object)) in sheet_objects.iter().enumerate() {
                let sheet = self.parse_sheet(index, object)?;
                if !sheet.is_empty() || !sheet.name.is_empty() {
                    sheets.push(sheet);
                }
            }
        }

        // If no sheets found, try to extract tables directly
        if sheets.is_empty() {
            let tables = self.extract_tables()?;
            if !tables.is_empty() {
                let mut default_sheet = NumbersSheet::new("Sheet 1".to_string(), 0);
                for table in tables {
                    default_sheet.add_table(table);
                }
                sheets.push(default_sheet);
            }
        }

        Ok(sheets)
    }

    /// Parse a single sheet from an object
    fn parse_sheet(&self, index: usize, object: &crate::iwa::archive::ArchiveObject) -> Result<NumbersSheet> {
        // Extract sheet name from decoded messages
        let text_parts = object.extract_text();
        let sheet_name = text_parts.first()
            .cloned()
            .unwrap_or_else(|| format!("Sheet {}", index + 1));

        let mut sheet = NumbersSheet::new(sheet_name, index);

        // Extract tables for this sheet
        // In a full implementation, we would parse the sheet protobuf message
        // to get references to table objects and resolve them
        let tables = self.extract_tables()?;
        for table in tables {
            sheet.add_table(table);
        }

        Ok(sheet)
    }

    /// Extract tables from the document
    fn extract_tables(&self) -> Result<Vec<NumbersTable>> {
        let mut tables = Vec::new();

        // Find table model objects (message type 100 is TST.TableModelArchive)
        let table_objects = self.bundle.find_objects_by_type(100);

        for (_archive_name, object) in table_objects {
            let table = self.parse_table(object)?;
            if !table.is_empty() || !table.name.is_empty() {
                tables.push(table);
            }
        }

        Ok(tables)
    }

    /// Parse a single table from an object
    fn parse_table(&self, object: &crate::iwa::archive::ArchiveObject) -> Result<NumbersTable> {
        use prost::Message;
        
        // Extract table name from decoded messages
        let text_parts = object.extract_text();
        let table_name = text_parts.first()
            .cloned()
            .unwrap_or_else(|| "Table".to_string());

        let mut table = NumbersTable::new(table_name);

        // Parse the TableModelArchive protobuf message
        // TableModelArchive contains:
        // - table_name: string
        // - number_of_rows: uint32
        // - number_of_columns: uint32
        // - data_store: reference to TableDataList
        // - table_id: UUID
        
        if let Some(raw_message) = object.messages.first() {
            // Try to decode as TableModelArchive
            if let Ok(table_model) = crate::iwa::protobuf::tst::TableModelArchive::decode(&*raw_message.data) {
                // Set table dimensions from protobuf fields
                // These are required uint32 fields in the proto, so they're always present
                table.row_count = table_model.number_of_rows as usize;
                table.column_count = table_model.number_of_columns as usize;
                
                // Extract table name if available
                if !table_model.table_name.is_empty() {
                    table.name = table_model.table_name.clone();
                }
                
                // TODO: Parse cell data from data_store reference
                // The data_store field contains a reference to a TableDataList object
                // which stores the actual cell values. We would need to:
                // 1. Extract the data_store object ID
                // 2. Resolve it using the object_index
                // 3. Parse the TableDataList to extract cell values
                //
                // For now, we'll extract any text content found in the object
                // as cell values (this will work for simple cases)
                if !text_parts.is_empty() {
                    // Place text parts as cells in the first column
                    for (idx, text) in text_parts.iter().skip(1).enumerate() {
                        if !text.is_empty() && idx < table.row_count {
                            table.set_cell(idx, 0, super::cell::CellValue::Text(text.clone()));
                        }
                    }
                }
            }
        }

        Ok(table)
    }

    /// Get the underlying bundle
    pub fn bundle(&self) -> &Bundle {
        &self.bundle
    }

    /// Get the object index
    pub fn object_index(&self) -> &ObjectIndex {
        &self.object_index
    }

    /// Get document statistics
    pub fn stats(&self) -> NumbersDocumentStats {
        let total_objects = self.object_index.all_object_ids().len();
        let sheets_result = self.sheets();
        let sheet_count = sheets_result.as_ref().map(|s| s.len()).unwrap_or(0);
        let table_count = sheets_result
            .as_ref()
            .map(|sheets| sheets.iter().map(|s| s.table_count()).sum())
            .unwrap_or(0);

        NumbersDocumentStats {
            total_objects,
            sheet_count,
            table_count,
            application: Application::Numbers,
        }
    }
}

/// Statistics about a Numbers document
#[derive(Debug, Clone)]
pub struct NumbersDocumentStats {
    /// Total number of objects
    pub total_objects: usize,
    /// Number of sheets
    pub sheet_count: usize,
    /// Total number of tables across all sheets
    pub table_count: usize,
    /// Application type (always Numbers)
    pub application: Application,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_numbers_document_open() {
        let doc_path = std::path::Path::new("test.numbers");
        if !doc_path.exists() {
            // Skip test if test file doesn't exist
            return;
        }

        let doc_result = NumbersDocument::open(doc_path);
        assert!(doc_result.is_ok(), "Failed to open Numbers document: {:?}", doc_result.err());

        let doc = doc_result.unwrap();
        assert!(doc.object_index.all_object_ids().len() > 0);
    }

    #[test]
    fn test_numbers_text_extraction() {
        let doc_path = std::path::Path::new("test.numbers");
        if !doc_path.exists() {
            return;
        }

        let doc = NumbersDocument::open(doc_path).unwrap();
        let text_result = doc.text();
        assert!(text_result.is_ok());
    }

    #[test]
    fn test_numbers_sheets() {
        let doc_path = std::path::Path::new("test.numbers");
        if !doc_path.exists() {
            return;
        }

        let doc = NumbersDocument::open(doc_path).unwrap();
        let sheets_result = doc.sheets();
        assert!(sheets_result.is_ok());

        let sheets = sheets_result.unwrap();
        // Document should have at least one sheet (even if implicit)
        assert!(!sheets.is_empty(), "Document should have at least one sheet");
    }
}

