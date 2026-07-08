//! PPT-specific Escher (Office Drawing) functionality.
//!
//! This module re-exports the shared Escher functionality and adds
//! PPT-specific extensions where needed.

// Re-export shared Escher functionality
pub use crate::escher::{
    EscherArrayProperty, EscherContainer, EscherParser, EscherProperties, EscherPropertyId,
    EscherPropertyValue, EscherRecord, EscherRecordType, EscherShape, EscherShapeFactory,
    EscherShapeType, ShapeAnchor, extract_text_from_escher,
};

// Re-export text extraction for backwards compatibility
pub use crate::escher::text::extract_text_from_textbox;
