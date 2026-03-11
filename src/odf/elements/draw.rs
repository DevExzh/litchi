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

impl Default for DrawPage {
    fn default() -> Self {
        Self::new()
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

impl Default for TextBox {
    fn default() -> Self {
        Self::new()
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

impl Default for Image {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_draw_page_new() {
        let page = DrawPage::new();
        assert!(page.name().is_none());
        assert!(page.style_name().is_none());
    }

    #[test]
    fn test_draw_page_default() {
        let page: DrawPage = Default::default();
        assert!(page.name().is_none());
    }

    #[test]
    fn test_draw_page_name() {
        let mut page = DrawPage::new();
        page.set_name("Slide1");
        assert_eq!(page.name(), Some("Slide1"));
    }

    #[test]
    fn test_text_box_new() {
        let text_box = TextBox::new();
        assert_eq!(text_box.text().unwrap(), "");
    }

    #[test]
    fn test_text_box_default() {
        let text_box: TextBox = Default::default();
        assert_eq!(text_box.text().unwrap(), "");
    }

    #[test]
    fn test_image_new() {
        let image = Image::new();
        assert!(image.href().is_none());
    }

    #[test]
    fn test_image_default() {
        let image: Image = Default::default();
        assert!(image.href().is_none());
    }
}
