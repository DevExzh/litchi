//! Semantic model for the terminal records of a PPT `DocumentContainer`.

/// Position of the optional PowerPoint 12 custom table-style package.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CustomTableStylesPlacement {
    BeforeEndDocument,
    AfterEndDocument,
}

/// Strictly validated terminal structure of a `DocumentContainer`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DocumentStructure {
    pub end_document_child_index: usize,
    pub custom_table_styles: Option<CustomTableStylesPlacement>,
}
