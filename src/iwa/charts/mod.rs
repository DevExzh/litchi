//! Chart Support for iWork Documents
//!
//! This module provides support for extracting metadata and content from
//! charts in iWork documents (Numbers, Pages, Keynote).
//!
//! Charts in iWork contain:
//! - Chart titles and axis labels
//! - Data series names
//! - Legend text
//! - Grid data (row/column names and values)

pub mod metadata_extractor;

pub use metadata_extractor::{ChartMetadataExtractor, ChartMetadata};

