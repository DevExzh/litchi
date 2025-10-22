//! Styles and formatting for Excel files.
//!
//! This module provides comprehensive parsing and management of cell styles,
//! number formats, fonts, fills, borders, and other formatting information
//! from XLSX files.
//!
//! # Architecture
//!
//! The styles module is organized into several components:
//!
//! - `parser`: XML parsing logic for styles.xml
//! - `number_format`: Number format definitions and detection
//! - `font`: Font information
//! - `fill`: Fill patterns and colors
//! - `border`: Border styles
//! - `alignment`: Cell alignment information
//! - `cell_style`: Cell style format records
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi::ooxml::xlsx::Styles;
//!
//! let styles_xml = std::fs::read_to_string("xl/styles.xml")?;
//! let styles = Styles::parse(&styles_xml)?;
//!
//! // Get a cell style
//! if let Some(style) = styles.get_cell_style(0) {
//!     println!("Style has font ID: {:?}", style.font_id);
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

mod alignment;
mod border;
mod cell_style;
mod fill;
mod font;
mod number_format;
mod parser;

pub use alignment::Alignment;
pub use border::{Border, BorderStyle};
pub use cell_style::CellStyle;
pub use fill::Fill;
pub use font::Font;
pub use number_format::NumberFormat;

use std::collections::HashMap;

use crate::ooxml::error::Result;

/// Styles collection for an Excel workbook.
///
/// Contains all the formatting information including number formats,
/// fonts, fills, borders, and cell styles.
#[derive(Debug, Default)]
pub struct Styles {
    /// Custom number formats (ID -> format code)
    pub number_formats: HashMap<u32, NumberFormat>,
    /// Font definitions
    pub fonts: Vec<Font>,
    /// Fill patterns and colors
    pub fills: Vec<Fill>,
    /// Border styles
    pub borders: Vec<Border>,
    /// Cell style formats (used as templates)
    pub cell_styles: Vec<CellStyle>,
    /// Cell format records (cellXfs - the actual styles applied to cells)
    pub cell_xfs: Vec<CellStyle>,
}

impl Styles {
    /// Create a new empty styles collection.
    #[inline]
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse styles from xl/styles.xml content.
    ///
    /// # Example
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Styles;
    ///
    /// let xml_content = std::fs::read_to_string("xl/styles.xml")?;
    /// let styles = Styles::parse(&xml_content)?;
    /// println!("Loaded {} fonts", styles.fonts.len());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn parse(content: &str) -> Result<Self> {
        parser::parse_styles(content)
    }

    /// Get a number format by ID.
    ///
    /// Returns both built-in and custom number formats.
    #[inline]
    pub fn get_number_format(&self, id: u32) -> Option<&NumberFormat> {
        self.number_formats.get(&id)
    }

    /// Get a font by ID (index).
    #[inline]
    pub fn get_font(&self, id: usize) -> Option<&Font> {
        self.fonts.get(id)
    }

    /// Get a fill by ID (index).
    #[inline]
    pub fn get_fill(&self, id: usize) -> Option<&Fill> {
        self.fills.get(id)
    }

    /// Get a border by ID (index).
    #[inline]
    pub fn get_border(&self, id: usize) -> Option<&Border> {
        self.borders.get(id)
    }

    /// Get a cell style by ID (index).
    ///
    /// This returns the actual cell format (from cellXfs) that is
    /// referenced by cells in the workbook.
    #[inline]
    pub fn get_cell_style(&self, id: usize) -> Option<&CellStyle> {
        self.cell_xfs.get(id)
    }

    /// Get the number of fonts defined.
    #[inline]
    pub fn font_count(&self) -> usize {
        self.fonts.len()
    }

    /// Get the number of fills defined.
    #[inline]
    pub fn fill_count(&self) -> usize {
        self.fills.len()
    }

    /// Get the number of borders defined.
    #[inline]
    pub fn border_count(&self) -> usize {
        self.borders.len()
    }

    /// Get the number of cell styles defined.
    #[inline]
    pub fn cell_style_count(&self) -> usize {
        self.cell_xfs.len()
    }
}
