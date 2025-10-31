//! Numbers Document Implementation
//!
//! Provides high-level API for working with Apple Numbers spreadsheets.

use std::path::Path;

use super::sheet::NumbersSheet;
use super::table::NumbersTable;
use crate::iwa::Result;
use crate::iwa::bundle::Bundle;
use crate::iwa::object_index::ObjectIndex;
use crate::iwa::registry::Application;
use crate::iwa::text::TextExtractor;

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

    /// Create a Numbers document from an already-parsed ZIP archive.
    ///
    /// This is used for single-pass parsing where the ZIP archive has already
    /// been parsed during format detection. It avoids double-parsing.
    pub fn from_zip_archive(
        zip_archive: zip::ZipArchive<std::io::Cursor<Vec<u8>>>,
    ) -> Result<Self> {
        let bundle = Bundle::from_zip_archive(zip_archive)?;
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
            let tables = self.extract_all_tables()?;
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
    fn parse_sheet(
        &self,
        index: usize,
        object: &crate::iwa::archive::ArchiveObject,
    ) -> Result<NumbersSheet> {
        use prost::Message;

        // Extract sheet name from decoded messages
        let text_parts = object.extract_text();
        let sheet_name = text_parts
            .first()
            .cloned()
            .unwrap_or_else(|| format!("Sheet {}", index + 1));

        let mut sheet = NumbersSheet::new(sheet_name, index);

        // Parse the SheetArchive protobuf message to get table references
        if let Some(raw_message) = object.messages.first()
            && let Ok(sheet_archive) =
                crate::iwa::protobuf::tn::SheetArchive::decode(&*raw_message.data)
        {
            // Extract table references from drawable_infos
            // Tables in Numbers are stored as drawables
            for drawable_ref in &sheet_archive.drawable_infos {
                if let Ok(table) = self.extract_table_from_drawable(drawable_ref.identifier) {
                    sheet.add_table(table);
                }
            }
        }

        // Fallback: Extract all tables from the document if sheet has none
        if sheet.table_count() == 0 {
            let tables = self.extract_all_tables()?;
            for table in tables {
                sheet.add_table(table);
            }
        }

        Ok(sheet)
    }

    /// Extract all tables from the document
    fn extract_all_tables(&self) -> Result<Vec<NumbersTable>> {
        use super::table_extractor::TableDataExtractor;

        let extractor = TableDataExtractor::new(&self.bundle, &self.object_index);
        extractor.extract_all_tables()
    }

    /// Extract a table from a drawable reference
    fn extract_table_from_drawable(&self, drawable_id: u64) -> Result<NumbersTable> {
        use prost::Message;

        if let Some(resolved) = self
            .object_index
            .resolve_object(&self.bundle, drawable_id)?
        {
            // Look for TableInfoArchive which wraps the table model
            for msg in &resolved.messages {
                if let Ok(table_info) =
                    crate::iwa::protobuf::tst::TableInfoArchive::decode(&*msg.data)
                {
                    // The table_model field contains a reference to the TableModelArchive
                    let table_model_id = table_info.table_model.identifier;
                    return self.extract_table_from_model(table_model_id);
                }
            }
        }

        Err(crate::iwa::Error::ParseError(
            "Could not extract table from drawable".to_string(),
        ))
    }

    /// Extract a table from a TableModelArchive reference
    fn extract_table_from_model(&self, table_model_id: u64) -> Result<NumbersTable> {
        use super::table_extractor::TableDataExtractor;

        let extractor = TableDataExtractor::new(&self.bundle, &self.object_index);

        if let Some(resolved) = self
            .object_index
            .resolve_object(&self.bundle, table_model_id)?
            && let Some(table) = extractor.extract_table_from_object(&resolved)?
        {
            return Ok(table);
        }

        Err(crate::iwa::Error::ParseError(
            "Could not extract table from model".to_string(),
        ))
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
        assert!(
            doc_result.is_ok(),
            "Failed to open Numbers document: {:?}",
            doc_result.err()
        );

        let doc = doc_result.unwrap();
        assert!(!doc.object_index.all_object_ids().is_empty());
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
        assert!(
            !sheets.is_empty(),
            "Document should have at least one sheet"
        );
    }
}
