//! Metadata elements for ODF documents.
//!
//! This module provides classes for metadata elements like creator,
//! date, title, and other document properties.

use super::element::{Element, ElementBase};

/// A metadata element
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct Meta {
    element: Element,
}

#[allow(dead_code)]
impl Meta {
    /// Create new metadata
    pub fn new() -> Self {
        Self {
            element: Element::new("office:meta"),
        }
    }

    /// Get the title
    pub fn title(&self) -> Option<&str> {
        self.element.get_attribute("dc:title")
    }

    /// Get the creator
    pub fn creator(&self) -> Option<&str> {
        self.element.get_attribute("dc:creator")
    }
}

impl Default for Meta {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_meta_new() {
        let meta = Meta::new();
        assert!(meta.title().is_none());
        assert!(meta.creator().is_none());
    }

    #[test]
    fn test_meta_default() {
        let meta: Meta = Default::default();
        assert!(meta.title().is_none());
        assert!(meta.creator().is_none());
    }
}
