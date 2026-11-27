//! XLSX workbook-metadata package vocabulary.
//!
//! OPC relationship discovery remains in the compatibility host adapter.
//! These constants keep the owner’s package contract centralized without
//! changing that adapter’s error or facade boundary.

/// Transitional `SpreadsheetML` namespace for workbook metadata.
pub const SPREADSHEETML_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
/// Strict `SpreadsheetML` namespace for workbook metadata.
pub const STRICT_SPREADSHEETML_NAMESPACE: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
/// Transitional relationship type for a workbook sheet-metadata part.
pub const SHEET_METADATA_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
/// Strict relationship type for a workbook sheet-metadata part.
pub const STRICT_SHEET_METADATA_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/sheetMetadata";
/// OPC content type for a workbook sheet-metadata part.
pub const SHEET_METADATA_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml";
