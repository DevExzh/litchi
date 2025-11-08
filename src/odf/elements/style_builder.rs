//! Style builder for creating and modifying ODF styles.
//!
//! This module provides convenient builders for creating styles programmatically.
//!
//! **Note**: These builders are provided as public API for advanced style manipulation.
//! Future versions will integrate these more deeply into document builders.

#![allow(dead_code)] // Public API - style builder utilities for future features

use super::style::Style;

/// Builder for creating text styles.
///
/// # Examples
///
/// ```
/// use litchi::odf::elements::StyleBuilder;
///
/// let style = StyleBuilder::new("MyStyle")
///     .font_size("14pt")
///     .font_weight("bold")
///     .font_family("Arial")
///     .build();
/// ```
#[allow(dead_code)] // Public API for future use
pub struct StyleBuilder {
    name: String,
    family: String,
    font_size: Option<String>,
    font_weight: Option<String>,
    font_family: Option<String>,
    font_style: Option<String>,
    color: Option<String>,
    background_color: Option<String>,
    text_decoration: Option<String>,
    text_align: Option<String>,
}

impl StyleBuilder {
    /// Create a new style builder with a name.
    ///
    /// # Arguments
    ///
    /// * `name` - Name for the style
    pub fn new(name: &str) -> Self {
        Self {
            name: name.to_string(),
            family: "text".to_string(),
            font_size: None,
            font_weight: None,
            font_family: None,
            font_style: None,
            color: None,
            background_color: None,
            text_decoration: None,
            text_align: None,
        }
    }

    /// Set the style family (text, paragraph, table, etc.).
    pub fn family(mut self, family: &str) -> Self {
        self.family = family.to_string();
        self
    }

    /// Set the font size (e.g., "12pt", "14px").
    pub fn font_size(mut self, size: &str) -> Self {
        self.font_size = Some(size.to_string());
        self
    }

    /// Set the font weight (e.g., "bold", "normal").
    pub fn font_weight(mut self, weight: &str) -> Self {
        self.font_weight = Some(weight.to_string());
        self
    }

    /// Set the font family (e.g., "Arial", "Times New Roman").
    pub fn font_family(mut self, family: &str) -> Self {
        self.font_family = Some(family.to_string());
        self
    }

    /// Set the font style (e.g., "italic", "normal").
    pub fn font_style(mut self, style: &str) -> Self {
        self.font_style = Some(style.to_string());
        self
    }

    /// Set the text color (e.g., "#000000", "red").
    pub fn color(mut self, color: &str) -> Self {
        self.color = Some(color.to_string());
        self
    }

    /// Set the background color.
    pub fn background_color(mut self, color: &str) -> Self {
        self.background_color = Some(color.to_string());
        self
    }

    /// Set text decoration (e.g., "underline", "line-through").
    pub fn text_decoration(mut self, decoration: &str) -> Self {
        self.text_decoration = Some(decoration.to_string());
        self
    }

    /// Set text alignment (e.g., "left", "center", "right", "justify").
    pub fn text_align(mut self, align: &str) -> Self {
        self.text_align = Some(align.to_string());
        self
    }

    /// Build the style.
    pub fn build(self) -> Style {
        let mut style = Style::with_name_and_family(&self.name, &self.family);

        // Set text properties
        if let Some(size) = self.font_size {
            style.set_text_property("fo:font-size", &size);
        }
        if let Some(weight) = self.font_weight {
            style.set_text_property("fo:font-weight", &weight);
        }
        if let Some(family) = self.font_family {
            style.set_text_property("style:font-name", &family);
        }
        if let Some(style_val) = self.font_style {
            style.set_text_property("fo:font-style", &style_val);
        }
        if let Some(color) = self.color {
            style.set_text_property("fo:color", &color);
        }
        if let Some(bg_color) = self.background_color {
            style.set_text_property("fo:background-color", &bg_color);
        }
        if let Some(decoration) = self.text_decoration {
            style.set_text_property("style:text-underline-style", &decoration);
        }

        // Set paragraph properties
        if let Some(align) = self.text_align {
            style.set_paragraph_property("fo:text-align", &align);
        }

        style
    }
}

/// Builder for creating paragraph styles.
#[allow(dead_code)] // Public API for future use
pub struct ParagraphStyleBuilder {
    inner: StyleBuilder,
    margin_top: Option<String>,
    margin_bottom: Option<String>,
    margin_left: Option<String>,
    margin_right: Option<String>,
    line_height: Option<String>,
}

impl ParagraphStyleBuilder {
    /// Create a new paragraph style builder.
    pub fn new(name: &str) -> Self {
        Self {
            inner: StyleBuilder::new(name).family("paragraph"),
            margin_top: None,
            margin_bottom: None,
            margin_left: None,
            margin_right: None,
            line_height: None,
        }
    }

    /// Set font size.
    pub fn font_size(mut self, size: &str) -> Self {
        self.inner = self.inner.font_size(size);
        self
    }

    /// Set font weight.
    pub fn font_weight(mut self, weight: &str) -> Self {
        self.inner = self.inner.font_weight(weight);
        self
    }

    /// Set text alignment.
    pub fn text_align(mut self, align: &str) -> Self {
        self.inner = self.inner.text_align(align);
        self
    }

    /// Set top margin.
    pub fn margin_top(mut self, margin: &str) -> Self {
        self.margin_top = Some(margin.to_string());
        self
    }

    /// Set bottom margin.
    pub fn margin_bottom(mut self, margin: &str) -> Self {
        self.margin_bottom = Some(margin.to_string());
        self
    }

    /// Set left margin.
    pub fn margin_left(mut self, margin: &str) -> Self {
        self.margin_left = Some(margin.to_string());
        self
    }

    /// Set right margin.
    pub fn margin_right(mut self, margin: &str) -> Self {
        self.margin_right = Some(margin.to_string());
        self
    }

    /// Set line height.
    pub fn line_height(mut self, height: &str) -> Self {
        self.line_height = Some(height.to_string());
        self
    }

    /// Build the paragraph style.
    pub fn build(self) -> Style {
        let mut style = self.inner.build();

        // Set paragraph-specific properties
        if let Some(margin) = self.margin_top {
            style.set_paragraph_property("fo:margin-top", &margin);
        }
        if let Some(margin) = self.margin_bottom {
            style.set_paragraph_property("fo:margin-bottom", &margin);
        }
        if let Some(margin) = self.margin_left {
            style.set_paragraph_property("fo:margin-left", &margin);
        }
        if let Some(margin) = self.margin_right {
            style.set_paragraph_property("fo:margin-right", &margin);
        }
        if let Some(height) = self.line_height {
            style.set_paragraph_property("fo:line-height", &height);
        }

        style
    }
}
