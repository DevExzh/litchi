//! Layered SpreadsheetML workbook future-metadata owner.
///
/// Semantic values live in `model`, bounded XML/MCE conversion in `codec`,
/// and the package contract vocabulary in `package`. OPC discovery remains
/// in the OOXML compatibility host adapter.
mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::{
    FutureMetadata, MetadataBehavior, MetadataBlock, MetadataRecord, MetadataType,
    OpaqueMetadataExtension, WorkbookMetadata,
};
pub use package::{
    SHEET_METADATA_CONTENT_TYPE, SHEET_METADATA_RELATIONSHIP_TYPE, SPREADSHEETML_NAMESPACE,
    STRICT_SHEET_METADATA_RELATIONSHIP_TYPE, STRICT_SPREADSHEETML_NAMESPACE,
};

// Historical names remain canonical at this contextual owner path; the
// litchi_xlsx crate root continues to re-export the same types.
