//! Chart Metadata Extraction
//!
//! Charts in iWork documents (TSCH namespace) contain rich metadata including
//! titles, labels, data series information, and grid data. This module extracts
//! that information for analysis and export.
//!
//! ## Chart Structure
//!
//! - **ChartArchive**: Main chart object with references to styles and data
//! - **ChartGridArchive**: Contains the actual chart data (rows, columns, values)
//! - **ChartMediatorArchive**: Links chart to data sources
//!
//! ## Example
//!
//! ```rust,ignore
//! use litchi_iwa::charts::ChartMetadataExtractor;
//!
//! let extractor = ChartMetadataExtractor::new(&bundle, &index);
//! let charts = extractor.extract_all_charts()?;
//!
//! for chart in charts {
//!     println!("Chart with {} series", chart.series_count);
//!     println!("Row names: {:?}", chart.row_names);
//!     println!("Column names: {:?}", chart.column_names);
//! }
//! ```

use crate::Result;
use crate::bundle::Bundle;
use crate::charts::options::read_chart_non_style_title;
use crate::charts::{ChartKind, IWorkChartArchive};
use crate::object_index::{ObjectIndex, ResolvedObject};
use crate::protobuf::tsch;
use prost::Message;

const LEGACY_CHART_MESSAGE_TYPE: u32 = 5_000;
const CHART_DRAWABLE_MESSAGE_TYPE: u32 = 5_021;
const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;

/// Metadata extracted from a chart
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartMetadata {
    /// Object identifier of the chart drawable.
    pub object_id: u64,
    /// Chart title (if present)
    pub title: Option<String>,
    /// Row names from the chart grid
    pub row_names: Vec<String>,
    /// Column names from the chart grid
    pub column_names: Vec<String>,
    /// Number of data series
    pub series_count: usize,
    /// Strongly typed native chart kind.
    pub chart_type: ChartKind,
    /// Whether chart contains default/sample data
    pub contains_default_data: bool,
}

impl ChartMetadata {
    /// Borrow all text content from the chart without allocating or cloning.
    pub fn all_text(&self) -> impl Iterator<Item = &str> {
        self.title
            .iter()
            .map(String::as_str)
            .chain(self.row_names.iter().map(String::as_str))
            .chain(self.column_names.iter().map(String::as_str))
    }

    /// Check if chart has any meaningful content
    pub fn has_content(&self) -> bool {
        self.title.is_some() || !self.row_names.is_empty() || !self.column_names.is_empty()
    }
}

/// Extractor for chart metadata
pub struct ChartMetadataExtractor<'a> {
    bundle: &'a Bundle,
    object_index: &'a ObjectIndex,
}

impl<'a> ChartMetadataExtractor<'a> {
    /// Create a new chart metadata extractor
    pub fn new(bundle: &'a Bundle, object_index: &'a ObjectIndex) -> Self {
        Self {
            bundle,
            object_index,
        }
    }

    /// Extract metadata from all charts in the document
    pub fn extract_all_charts(&self) -> Result<Vec<ChartMetadata>> {
        let mut charts = Vec::new();

        for chart_type in [LEGACY_CHART_MESSAGE_TYPE, CHART_DRAWABLE_MESSAGE_TYPE] {
            let chart_entries = self.object_index.find_objects_by_type(chart_type);

            for entry in chart_entries {
                if let Some(resolved) = self.object_index.resolve(self.bundle, entry.id())?
                    && let Some(metadata) = self.extract_chart_metadata(&resolved)?
                {
                    charts.push(metadata);
                }
            }
        }

        Ok(charts)
    }

    /// Extract metadata from a single chart object
    fn extract_chart_metadata(&self, object: &ResolvedObject) -> Result<Option<ChartMetadata>> {
        for message in &object.messages {
            match message.type_ {
                CHART_DRAWABLE_MESSAGE_TYPE => {
                    let drawable = IWorkChartArchive::decode(&message.data)?;
                    if let Some(chart) = drawable.chart {
                        return Ok(Some(self.parse_chart(object.id().get(), &chart)?));
                    }
                },
                LEGACY_CHART_MESSAGE_TYPE => {
                    let chart = tsch::pre_uff::ChartInfoArchive::decode(message.data.as_slice())?;
                    return Ok(Some(Self::parse_legacy_chart(object.id().get(), &chart)));
                },
                _ => {},
            }
        }

        Ok(None)
    }

    /// Parse a ChartArchive to extract metadata
    fn parse_chart(&self, object_id: u64, chart: &tsch::ChartArchive) -> Result<ChartMetadata> {
        let grid = chart.grid.as_ref();
        Ok(ChartMetadata {
            object_id,
            title: self.extract_chart_title(chart)?,
            row_names: grid.map_or_else(Vec::new, |grid| grid.row_name.clone()),
            column_names: grid.map_or_else(Vec::new, |grid| grid.column_name.clone()),
            series_count: grid.map_or(0, |grid| grid.grid_row.len()),
            chart_type: ChartKind::from_raw(
                chart
                    .chart_type
                    .unwrap_or(tsch::ChartType::UndefinedChartType as i32),
            ),
            contains_default_data: chart.contains_default_data.unwrap_or(false),
        })
    }

    fn parse_legacy_chart(
        object_id: u64,
        chart: &tsch::pre_uff::ChartInfoArchive,
    ) -> ChartMetadata {
        let grid = chart.chart_model.inline_grid.as_ref();
        ChartMetadata {
            object_id,
            title: None,
            row_names: grid.map_or_else(Vec::new, |grid| grid.row_name.clone()),
            column_names: grid.map_or_else(Vec::new, |grid| grid.column_name.clone()),
            series_count: grid.map_or(0, |grid| grid.value_row.len()),
            chart_type: ChartKind::from_raw(chart.chart_type),
            contains_default_data: false,
        }
    }

    /// Extract the title from the chart's native non-style extension.
    fn extract_chart_title(&self, chart: &tsch::ChartArchive) -> Result<Option<String>> {
        let Some(reference) = chart.chart_non_style.as_ref() else {
            return Ok(None);
        };
        let Some(resolved) = self
            .object_index
            .resolve_object(self.bundle, reference.identifier)?
        else {
            return Ok(None);
        };
        let mut messages = resolved
            .messages
            .iter()
            .filter(|message| message.type_ == CHART_NON_STYLE_MESSAGE_TYPE);
        let Some(message) = messages.next() else {
            return Err(crate::Error::InvalidFormat(format!(
                "chart non-style {} must have exactly one payload",
                reference.identifier
            )));
        };
        if messages.next().is_some() {
            return Err(crate::Error::InvalidFormat(format!(
                "chart non-style {} must have exactly one payload",
                reference.identifier
            )));
        }
        read_chart_non_style_title(message.data.as_slice())
    }

    /// Extract metadata from a specific chart by object ID
    pub fn extract_chart_by_id(&self, chart_id: u64) -> Result<Option<ChartMetadata>> {
        if let Some(resolved) = self.object_index.resolve_object(self.bundle, chart_id)? {
            return self.extract_chart_metadata(&resolved);
        }

        Ok(None)
    }

    /// Get all chart titles in the document
    pub fn get_all_chart_titles(&self) -> Result<Vec<String>> {
        let charts = self.extract_all_charts()?;
        Ok(charts.into_iter().filter_map(|c| c.title).collect())
    }

    /// Get total number of charts in the document
    pub fn chart_count(&self) -> Result<usize> {
        let mut count = 0;

        for chart_type in [LEGACY_CHART_MESSAGE_TYPE, CHART_DRAWABLE_MESSAGE_TYPE] {
            count += self.object_index.find_objects_by_type(chart_type).len();
        }

        Ok(count)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_chart_metadata_creation() {
        let metadata = ChartMetadata {
            object_id: 42,
            title: None,
            row_names: Vec::new(),
            column_names: Vec::new(),
            series_count: 0,
            chart_type: ChartKind::Undefined,
            contains_default_data: false,
        };
        assert_eq!(metadata.title, None);
        assert_eq!(metadata.series_count, 0);
        assert!(!metadata.has_content());
    }

    #[test]
    fn test_chart_metadata_with_content() {
        let metadata = ChartMetadata {
            object_id: 42,
            title: Some("Sales Chart".to_owned()),
            row_names: vec!["Q1".to_owned(), "Q2".to_owned()],
            column_names: vec!["Revenue".to_owned()],
            series_count: 2,
            chart_type: ChartKind::Column2d,
            contains_default_data: false,
        };

        assert!(metadata.has_content());
        let all_text = metadata.all_text().collect::<Vec<_>>();
        assert_eq!(all_text.len(), 4);
        assert!(all_text.contains(&"Sales Chart"));
    }

    #[test]
    fn test_chart_type_display() {
        let metadata = ChartMetadata {
            object_id: 42,
            title: None,
            row_names: vec![],
            column_names: vec![],
            series_count: 0,
            chart_type: ChartKind::Bar2d,
            contains_default_data: false,
        };

        assert_eq!(metadata.chart_type, ChartKind::Bar2d);
    }
}
