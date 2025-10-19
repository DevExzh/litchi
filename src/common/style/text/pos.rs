//! Common text properties shared across formats.
//!
//! This module contains text formatting types that are used across
//! different Office file formats (OLE, OOXML, etc.).

/// Vertical text position (superscript/subscript).
///
/// Used to represent vertical alignment of text in both OLE (.doc) 
/// and OOXML (.docx) formats.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum VerticalPosition {
    /// Normal position
    #[default]
    Normal,
    /// Superscript
    Superscript,
    /// Subscript
    Subscript,
}

impl VerticalPosition {
    /// Check if this is a normal (non-superscript, non-subscript) position.
    #[inline]
    pub fn is_normal(&self) -> bool {
        matches!(self, VerticalPosition::Normal)
    }

    /// Check if this is superscript.
    #[inline]
    pub fn is_superscript(&self) -> bool {
        matches!(self, VerticalPosition::Superscript)
    }

    /// Check if this is subscript.
    #[inline]
    pub fn is_subscript(&self) -> bool {
        matches!(self, VerticalPosition::Subscript)
    }
}

