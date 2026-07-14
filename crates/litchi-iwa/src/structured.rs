//! Structured Data Extraction from iWork Documents
//!
//! This module provides utilities for extracting structured content such as:
//! - Tables from Numbers spreadsheets
//! - Slides from Keynote presentations  
//! - Sections and paragraphs from Pages documents

mod data;
mod extract;
mod section;
mod slide;
mod table;

#[cfg(test)]
mod model_tests;

pub use data::StructuredData;
pub use extract::{
    extract_all, extract_chart_metadata, extract_sections, extract_shape_text, extract_slides,
    extract_tables,
};
pub use section::Section;
pub use slide::Slide;
pub use table::{CellValue, Table};
