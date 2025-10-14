//! Styles and formatting for Excel files.
//!
//! This module provides parsing and management of cell styles,
//! number formats, and other formatting information.

use std::collections::HashMap;

use crate::sheet::Result;

/// Number format information
#[derive(Debug, Clone)]
pub struct NumberFormat {
    /// Format ID
    pub id: u32,
    /// Format code (e.g., "General", "0.00", etc.)
    pub code: String,
}

/// Font information
#[derive(Debug, Clone, Default)]
pub struct Font {
    /// Font name
    pub name: Option<String>,
    /// Font size
    pub size: Option<f64>,
    /// Bold flag
    pub bold: bool,
    /// Italic flag
    pub italic: bool,
    /// Color (RGB or theme color)
    pub color: Option<String>,
}

/// Fill information
#[derive(Debug, Clone)]
pub enum Fill {
    /// Pattern fill
    Pattern {
        /// Pattern type (e.g., "solid", "gray125")
        pattern_type: String,
        /// Foreground color
        fg_color: Option<String>,
        /// Background color
        bg_color: Option<String>,
    },
    /// Gradient fill (simplified)
    Gradient,
}

/// Border information
#[derive(Debug, Clone, Default)]
pub struct Border {
    /// Left border style
    pub left: Option<BorderStyle>,
    /// Right border style
    pub right: Option<BorderStyle>,
    /// Top border style
    pub top: Option<BorderStyle>,
    /// Bottom border style
    pub bottom: Option<BorderStyle>,
}

/// Border style information
#[derive(Debug, Clone)]
pub struct BorderStyle {
    /// Style name (e.g., "thin", "medium", "thick")
    pub style: String,
    /// Color
    pub color: Option<String>,
}

/// Cell style information
#[derive(Debug, Clone, Default)]
pub struct CellStyle {
    /// Number format ID
    pub num_fmt_id: Option<u32>,
    /// Font ID
    pub font_id: Option<u32>,
    /// Fill ID
    pub fill_id: Option<u32>,
    /// Border ID
    pub border_id: Option<u32>,
    /// Alignment information (simplified)
    pub alignment: Option<Alignment>,
}

/// Alignment information
#[derive(Debug, Clone)]
pub struct Alignment {
    /// Horizontal alignment
    pub horizontal: Option<String>,
    /// Vertical alignment
    pub vertical: Option<String>,
    /// Wrap text flag
    pub wrap_text: bool,
}

/// Styles collection
#[derive(Debug, Default)]
pub struct Styles {
    /// Number formats
    pub number_formats: HashMap<u32, NumberFormat>,
    /// Fonts
    pub fonts: Vec<Font>,
    /// Fills
    pub fills: Vec<Fill>,
    /// Borders
    pub borders: Vec<Border>,
    /// Cell style formats
    pub cell_styles: Vec<CellStyle>,
    /// Cell style XFs
    pub cell_xfs: Vec<CellStyle>,
}

impl Styles {
    /// Create a new empty styles collection.
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse styles from xl/styles.xml content.
    pub fn parse(_content: &str) -> Result<Self> {
        // TODO: Implement proper styles parsing
        // For now, return empty styles
        Ok(Styles::default())
    }

    /// Get a number format by ID.
    pub fn get_number_format(&self, id: u32) -> Option<&NumberFormat> {
        self.number_formats.get(&id)
    }

    /// Get a font by ID.
    pub fn get_font(&self, id: usize) -> Option<&Font> {
        self.fonts.get(id)
    }

    /// Get a cell style by ID.
    pub fn get_cell_style(&self, id: usize) -> Option<&CellStyle> {
        self.cell_xfs.get(id)
    }
}
