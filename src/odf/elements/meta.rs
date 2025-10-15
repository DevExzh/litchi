//! Metadata elements for ODF documents.
//!
//! This module provides classes for metadata elements like creator,
//! date, title, and other document properties.

use super::element::{Element, ElementBase};

/// A metadata element
#[derive(Debug, Clone)]
pub struct Meta {
    element: Element,
}

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
