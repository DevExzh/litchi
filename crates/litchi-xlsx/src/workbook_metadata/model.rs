//! Semantic SpreadsheetML workbook-metadata values.
//!
//! These types are intentionally contextual to the `workbook_metadata`
//! owner, so their canonical names do not repeat the owner prefix.

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MetadataBehavior {
    pub ghost_row: bool,
    pub ghost_column: bool,
    pub edit: bool,
    pub delete: bool,
    pub copy: bool,
    pub paste_all: bool,
    pub paste_formulas: bool,
    pub paste_values: bool,
    pub paste_formats: bool,
    pub paste_comments: bool,
    pub paste_data_validation: bool,
    pub paste_borders: bool,
    pub paste_column_widths: bool,
    pub paste_number_formats: bool,
    pub merge: bool,
    pub split_first: bool,
    pub split_all: bool,
    pub row_column_shift: bool,
    pub clear_all: bool,
    pub clear_formats: bool,
    pub clear_contents: bool,
    pub clear_comments: bool,
    pub assign: bool,
    pub coerce: bool,
    pub cell_metadata: bool,
    pub adjust: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MetadataType {
    pub name: String,
    pub minimum_supported_version: u32,
    pub behavior: MetadataBehavior,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MetadataRecord {
    /// One-based index into `WorkbookMetadata::types`.
    pub type_index: u32,
    /// Zero-based index into the matching future-metadata store.
    pub value_index: u32,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpaqueMetadataExtension {
    pub uri: String,
    /// Deterministically normalized, inert child XML from the extension.
    pub payload_xml: Vec<u8>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MetadataBlock {
    pub records: Vec<MetadataRecord>,
    pub extensions: Vec<OpaqueMetadataExtension>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FutureMetadata {
    pub name: String,
    pub blocks: Vec<MetadataBlock>,
    pub extensions: Vec<OpaqueMetadataExtension>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct WorkbookMetadata {
    pub types: Vec<MetadataType>,
    pub future: Vec<FutureMetadata>,
    pub cell_blocks: Vec<MetadataBlock>,
    pub value_blocks: Vec<MetadataBlock>,
    pub extensions: Vec<OpaqueMetadataExtension>,
}
