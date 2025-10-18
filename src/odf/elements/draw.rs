//! Drawing elements for ODF presentations.
//!
//! This module provides classes for drawing elements like shapes, frames,
//! images, and other graphical content.

use super::element::{Element, ElementBase};
use crate::common::Result;

/// A drawing page (slide) element
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct DrawPage {
    element: Element,
}

impl DrawPage {
    /// Create a new drawing page
    #[allow(dead_code)]
    pub fn new() -> Self {
        Self {
            element: Element::new("draw:page"),
        }
    }

    /// Get the page name
    #[allow(dead_code)]
    pub fn name(&self) -> Option<&str> {
        self.element.get_attribute("draw:name")
    }

    /// Set the page name
    #[allow(dead_code)]
    pub fn set_name(&mut self, name: &str) {
        self.element.set_attribute("draw:name", name);
    }

    /// Get the style name
    #[allow(dead_code)]
    pub fn style_name(&self) -> Option<&str> {
        self.element.get_attribute("draw:style-name")
    }
}

/// A text box element
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct TextBox {
    element: Element,
}

#[allow(dead_code)]
impl TextBox {
    /// Create a new text box
    pub fn new() -> Self {
        Self {
            element: Element::new("draw:text-box"),
        }
    }

    /// Get the text content
    pub fn text(&self) -> Result<String> {
        Ok(self.element.get_text_recursive())
    }
}

/// An image element
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct Image {
    element: Element,
}

#[allow(dead_code)]
impl Image {
    /// Create a new image
    pub fn new() -> Self {
        Self {
            element: Element::new("draw:image"),
        }
    }

    /// Get the image href
    pub fn href(&self) -> Option<&str> {
        self.element.get_attribute("xlink:href")
    }
}
