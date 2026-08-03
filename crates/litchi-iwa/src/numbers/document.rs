//! Numbers Document Implementation
//!
//! Provides high-level API for working with Apple Numbers spreadsheets.

use std::path::Path;
use std::sync::Arc;

use super::sheet::NumbersSheet;
use super::table::NumbersTable;
use crate::bundle::{Bundle, BundleLimits};
use crate::object_index::ObjectIndex;
use crate::registry::{Application, detect_application_from_document};
use crate::text::TextExtractor;
use crate::{Error, Result};

/// High-level interface for Numbers documents
#[derive(Debug, Clone)]
pub struct NumbersDocument {
    state: Arc<NumbersDocumentState>,
}

#[derive(Debug)]
struct NumbersDocumentState {
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
    /// use litchi_iwa::numbers::NumbersDocument;
    ///
    /// let doc = NumbersDocument::open("spreadsheet.numbers")?;
    /// println!("Loaded Numbers document");
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::open_with_limits(path, BundleLimits::default())
    }

    /// Open a Numbers document under caller-selected bundle ingress ceilings.
    pub fn open_with_limits<P: AsRef<Path>>(path: P, limits: BundleLimits) -> Result<Self> {
        let bundle = Bundle::open_with_limits(path, limits)?;
        Self::verify_application(&bundle)?;
        let object_index = ObjectIndex::from_bundle(&bundle)?;

        Ok(Self::from_parts(bundle, object_index))
    }

    /// Open a Numbers document from raw bytes
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_iwa::numbers::NumbersDocument;
    /// use std::fs;
    ///
    /// let data = fs::read("spreadsheet.numbers")?;
    /// let doc = NumbersDocument::from_bytes(&data)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, BundleLimits::default())
    }

    /// Open a Numbers document from bytes under caller-selected ingress
    /// ceilings.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: BundleLimits) -> Result<Self> {
        let bundle = Bundle::from_bytes_with_limits(bytes, limits)?;
        Self::verify_application(&bundle)?;
        let object_index = ObjectIndex::from_bundle(&bundle)?;

        Ok(Self::from_parts(bundle, object_index))
    }

    fn from_parts(bundle: Bundle, object_index: ObjectIndex) -> Self {
        Self {
            state: Arc::new(NumbersDocumentState {
                bundle,
                object_index,
            }),
        }
    }

    /// Capture a cheap immutable snapshot that shares all parsed document state.
    pub fn snapshot(&self) -> Self {
        self.clone()
    }

    /// Create a Numbers document from raw bytes (ZIP archive data).
    ///
    /// This convenience entry point currently performs the same parsing as
    /// [`Self::from_bytes`]; it does not accept a previously parsed archive.
    pub fn from_archive_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes(bytes)
    }

    /// Create a Numbers document from archive bytes under caller-selected
    /// ingress ceilings.
    pub fn from_archive_bytes_with_limits(bytes: &[u8], limits: BundleLimits) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, limits)
    }

    fn verify_application(bundle: &Bundle) -> Result<()> {
        Self::root_document(bundle).map(|_| ())
    }

    fn root_document(bundle: &Bundle) -> Result<crate::protobuf::tn::DocumentArchive> {
        use prost::Message;

        let object = bundle
            .get_archive("Index/Document.iwa")
            .and_then(|archive| archive.object(1))
            .ok_or_else(|| Error::InvalidFormat("Numbers root object 1 is missing".to_owned()))?;
        object
            .messages
            .iter()
            .find(|message| {
                detect_application_from_document(&message.data) == Some(Application::Numbers)
            })
            .and_then(|message| {
                crate::protobuf::tn::DocumentArchive::decode(message.data.as_slice()).ok()
            })
            .ok_or_else(|| {
                Error::InvalidFormat("package does not contain a Numbers root document".to_owned())
            })
    }

    /// Extract all text content from the document
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_iwa::numbers::NumbersDocument;
    ///
    /// let doc = NumbersDocument::open("spreadsheet.numbers")?;
    /// let text = doc.text()?;
    /// println!("{}", text);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        let mut extractor = TextExtractor::new();
        extractor.extract_from_bundle(&self.state.bundle)?;
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
    /// use litchi_iwa::numbers::NumbersDocument;
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
        use super::table_extractor::TableDataExtractor;

        let document = Self::root_document(&self.state.bundle)?;
        let extractor = TableDataExtractor::new(&self.state.bundle, &self.state.object_index);
        let mut sheets = Vec::with_capacity(document.sheets.len());
        for (index, reference) in document.sheets.into_iter().enumerate() {
            let object = self
                .state
                .bundle
                .iter_archives()
                .map(|(_, archive)| archive)
                .find_map(|archive| archive.object(reference.identifier))
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers document references missing sheet object {}",
                        reference.identifier
                    ))
                })?;
            sheets.push(self.parse_sheet(index, object, &extractor)?);
        }
        Ok(sheets)
    }

    /// Parse a single sheet from an object
    fn parse_sheet(
        &self,
        index: usize,
        object: &crate::archive::ArchiveObject,
        extractor: &super::table_extractor::TableDataExtractor<'_>,
    ) -> Result<NumbersSheet> {
        use prost::Message;

        let sheet_archive = object
            .messages
            .iter()
            .find_map(|message| {
                crate::protobuf::tn::SheetArchive::decode(message.data.as_slice())
                    .ok()
                    .or_else(|| {
                        crate::protobuf::tn::FormBasedSheetArchive::decode(message.data.as_slice())
                            .ok()
                            .map(|form| form.super_)
                    })
            })
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers sheet object {:?} has no TN.SheetArchive payload",
                    object.archive_info.identifier
                ))
            })?;
        let mut sheet = NumbersSheet::new(sheet_archive.name, index);
        for drawable_ref in &sheet_archive.drawable_infos {
            if let Some(table) =
                self.extract_table_from_drawable(drawable_ref.identifier, extractor)?
            {
                sheet.add_table(table);
            }
        }

        Ok(sheet)
    }

    /// Extract a table from a drawable reference
    fn extract_table_from_drawable(
        &self,
        drawable_id: u64,
        extractor: &super::table_extractor::TableDataExtractor<'_>,
    ) -> Result<Option<NumbersTable>> {
        use prost::Message;

        let resolved = self
            .state
            .object_index
            .resolve_ref_id(&self.state.bundle, drawable_id)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers sheet references missing drawable object {drawable_id}"
                ))
            })?;
        // Protobuf decoding is permissive, so accept a TableInfoArchive only
        // when its model reference resolves to a typed TableModelArchive.
        for message in resolved.messages {
            let Ok(table_info) =
                crate::protobuf::tst::TableInfoArchive::decode(message.data.as_slice())
            else {
                continue;
            };
            let table_model_id = table_info.table_model.identifier;
            let Some(model) = self
                .state
                .object_index
                .resolve_ref_id(&self.state.bundle, table_model_id)?
            else {
                continue;
            };
            if !model.messages.iter().any(|message| {
                (message.type_ == 6000 || message.type_ == 6001)
                    && crate::protobuf::tst::TableModelArchive::decode(message.data.as_slice())
                        .is_ok()
            }) {
                continue;
            }
            return self
                .extract_table_from_model(table_model_id, extractor)
                .map(Some);
        }

        Ok(None)
    }

    /// Extract a table from a TableModelArchive reference
    fn extract_table_from_model(
        &self,
        table_model_id: u64,
        extractor: &super::table_extractor::TableDataExtractor<'_>,
    ) -> Result<NumbersTable> {
        if let Some(resolved) = self
            .state
            .object_index
            .resolve_ref_id(&self.state.bundle, table_model_id)?
            && let Some(table) = extractor.extract_table_from_object(&resolved)?
        {
            return Ok(table);
        }

        Err(crate::Error::ParseError(
            "Could not extract table from model".to_string(),
        ))
    }

    /// Get the underlying bundle
    pub fn bundle(&self) -> &Bundle {
        &self.state.bundle
    }

    /// Get the object index
    pub fn object_index(&self) -> &ObjectIndex {
        &self.state.object_index
    }

    /// Return a bounded, deterministic validation report for this snapshot.
    pub fn validation_report(&self) -> crate::bundle::BundleValidationReport {
        self.state.bundle.validation_report()
    }

    /// Validate this immutable snapshot without mutating it.
    pub fn validate(&self) -> Result<()> {
        self.validation_report().as_result()
    }

    /// Get document statistics after resolving the document sheets.
    pub fn stats(&self) -> Result<NumbersDocumentStats> {
        let total_objects = self.state.object_index.object_count();
        let sheets = self.sheets()?;
        let sheet_count = sheets.len();
        let table_count = sheets.iter().map(|sheet| sheet.table_count()).sum();

        Ok(NumbersDocumentStats {
            total_objects,
            sheet_count,
            table_count,
            application: Application::Numbers,
        })
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

    fn assert_send_sync<T: Send + Sync>() {}

    #[test]
    fn numbers_documents_are_send_and_sync() {
        assert_send_sync::<NumbersDocument>();
    }

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
        assert!(!doc.object_index().object_ids().is_empty());
        assert!(doc.validate().is_ok());
    }

    #[test]
    fn numbers_from_bytes_with_limits_enforces_input_budget() {
        let limits = BundleLimits::new(1, 10, 100, 100, 100).unwrap();
        let error = NumbersDocument::from_bytes_with_limits(&[0, 1], limits).unwrap_err();
        assert!(error.to_string().contains("iWork bundle input"));
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
}
