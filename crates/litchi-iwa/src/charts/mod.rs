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

mod archive;
mod data;
mod direction;
mod kind;
pub mod metadata_extractor;
pub(crate) mod source;

pub use archive::IWorkChartArchive;
pub use data::ChartData;
pub use direction::ChartSeriesDirection;
pub use kind::ChartKind;
pub use metadata_extractor::{ChartMetadata, ChartMetadataExtractor};
