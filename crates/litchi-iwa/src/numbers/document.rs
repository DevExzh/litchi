//! Numbers Document Implementation
//!
//! Provides high-level API for working with Apple Numbers spreadsheets.

use std::path::Path;
use std::sync::Arc;

use super::sheet::NumbersSheet;
use super::table::NumbersTable;
use crate::bundle::{Bundle, BundleLimits};
use crate::detect::detect_application_from_document;
use crate::object_index::ObjectIndex;
use crate::registry::Application;
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
    /// Immutable archive-free semantic state shared by document clones.
    semantic_document: litchi_numbers::Document,
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
        let root_document = Self::root_document(&bundle)?;
        let object_index = ObjectIndex::from_bundle(&bundle)?;

        Self::from_parts(bundle, object_index, root_document)
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
        let root_document = Self::root_document(&bundle)?;
        let object_index = ObjectIndex::from_bundle(&bundle)?;

        Self::from_parts(bundle, object_index, root_document)
    }

    fn from_parts(
        bundle: Bundle,
        object_index: ObjectIndex,
        root_document: crate::protobuf::tn::DocumentArchive,
    ) -> Result<Self> {
        let archive_sheets = Self::decode_sheets(&bundle, &object_index, &root_document)?;
        let mut semantic_sheets = Vec::new();
        semantic_sheets.try_reserve(archive_sheets.len()).map_err(|_| {
            crate::Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "Numbers semantic sheets",
                amount: archive_sheets.len(),
            })
        })?;
        for sheet in archive_sheets {
            semantic_sheets.push(sheet.into_semantic()?);
        }
        let semantic_document = litchi_numbers::Document::from_sheets(semantic_sheets)
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "Numbers semantic document is invalid at ingress: {error}"
                ))
            })?;
        Ok(Self {
            state: Arc::new(NumbersDocumentState {
                bundle,
                object_index,
                semantic_document,
            }),
        })
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

    /// Extract archive-aware sheet adapters.
    ///
    /// Numbers documents consist of multiple sheets, each containing tables.
    /// Native comments and other sidecars remain available through this
    /// archive-boundary view. Use [`Self::semantic_sheets`] for the immutable
    /// dependency-free model.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_iwa::numbers::NumbersDocument;
    ///
    /// let doc = NumbersDocument::open("spreadsheet.numbers")?;
    /// let sheets = doc.sheets()?;
    ///
    /// for sheet in doc.sheets()? {
    ///     println!("Sheet: {}", sheet.name());
    ///     for table in sheet.tables() {
    ///         println!("  Table: {} ({}x{})",
    ///             table.name(), table.row_count(), table.column_count());
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn sheets(&self) -> Result<Vec<NumbersSheet>> {
        let root_document = Self::root_document(&self.state.bundle)?;
        Self::decode_sheets(&self.state.bundle, &self.state.object_index, &root_document)
    }

    fn decode_sheets(
        bundle: &Bundle,
        object_index: &ObjectIndex,
        document: &crate::protobuf::tn::DocumentArchive,
    ) -> Result<Vec<NumbersSheet>> {
        use super::table_extractor::TableDataExtractor;

        let extractor = TableDataExtractor::new(bundle, object_index);
        let mut sheets = Vec::with_capacity(document.sheets.len());
        for (index, reference) in document.sheets.iter().enumerate() {
            let object = bundle
                .iter_archives()
                .map(|(_, archive)| archive)
                .find_map(|archive| archive.object(reference.identifier))
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers document references missing sheet object {}",
                        reference.identifier
                    ))
                })?;
            sheets.push(Self::parse_sheet(
                bundle,
                object_index,
                index,
                object,
                &extractor,
            )?);
        }

        Ok(sheets)
    }

    /// Return the immutable semantic sheet snapshot.
    ///
    /// The `Arc` is shared across repeated calls and cheap document snapshots;
    /// native IDs, protobuf records, package state, and comments stay in the
    /// archive-boundary view returned by [`Self::sheets`].
    pub fn semantic_sheets(&self) -> Arc<[litchi_numbers::Sheet]> {
        self.state.semantic_document.shared_sheets()
    }

    /// Borrow the immutable archive-free semantic Numbers document.
    ///
    /// Borrow the immutable archive-free semantic Numbers document.
    #[must_use]
    pub fn semantic_document(&self) -> &litchi_numbers::Document {
        &self.state.semantic_document
    }

    /// Capture a cheap handle to the immutable semantic Numbers snapshot.
    #[must_use]
    pub fn semantic_snapshot(&self) -> litchi_numbers::Document {
        self.state.semantic_document.snapshot()
    }

    /// Parse a single sheet from an object
    fn parse_sheet(
        bundle: &Bundle,
        object_index: &ObjectIndex,
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
            if let Some(table) = Self::extract_table_from_drawable(
                bundle,
                object_index,
                drawable_ref.identifier,
                extractor,
            )? {
                sheet.add_table(table);
            }
        }

        Ok(sheet)
    }

    /// Extract a table from a drawable reference
    fn extract_table_from_drawable(
        bundle: &Bundle,
        object_index: &ObjectIndex,
        drawable_id: u64,
        extractor: &super::table_extractor::TableDataExtractor<'_>,
    ) -> Result<Option<NumbersTable>> {
        use prost::Message;

        let resolved = object_index
            .resolve_ref_id(bundle, drawable_id)?
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
            let Some(model) = object_index.resolve_ref_id(bundle, table_model_id)? else {
                continue;
            };
            if !model.messages.iter().any(|message| {
                (message.type_ == 6000 || message.type_ == 6001)
                    && crate::protobuf::tst::TableModelArchive::decode(message.data.as_slice())
                        .is_ok()
            }) {
                continue;
            }
            return Self::extract_table_from_model(bundle, object_index, table_model_id, extractor)
                .map(Some);
        }

        Ok(None)
    }

    /// Extract a table from a TableModelArchive reference
    fn extract_table_from_model(
        bundle: &Bundle,
        object_index: &ObjectIndex,
        table_model_id: u64,
        extractor: &super::table_extractor::TableDataExtractor<'_>,
    ) -> Result<NumbersTable> {
        if let Some(resolved) = object_index.resolve_ref_id(bundle, table_model_id)?
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
