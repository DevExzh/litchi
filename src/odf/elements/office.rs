//! Office document elements.
//!
//! This module provides classes for the main office document elements
//! like document root, body, and other top-level structures.

use super::element::{Element, ElementBase};

/// An office document element
#[derive(Debug, Clone)]
pub struct Document {
    element: Element,
}

impl Document {
    /// Create a new office document
    pub fn new() -> Self {
        Self {
            element: Element::new("office:document"),
        }
    }

    /// Get the document version
    pub fn version(&self) -> Option<&str> {
        self.element.get_attribute("office:version")
    }
}

/// An office body element
#[derive(Debug, Clone)]
pub struct Body {
    element: Element,
}

impl Body {
    /// Create a new office body
    pub fn new() -> Self {
        Self {
            element: Element::new("office:body"),
        }
    }

    /// Get the body type
    pub fn body_type(&self) -> Option<&str> {
        // Check child elements to determine body type
        for child in self.element.children() {
            let tag = child.tag_name();
            if tag == "office:text" {
                return Some("text");
            } else if tag == "office:spreadsheet" {
                return Some("spreadsheet");
            } else if tag == "office:presentation" {
                return Some("presentation");
            }
        }
        None
    }
}
