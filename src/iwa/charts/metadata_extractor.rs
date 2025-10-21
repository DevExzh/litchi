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
//! use litchi::iwa::charts::ChartMetadataExtractor;
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

use crate::iwa::Result;
use crate::iwa::bundle::Bundle;
use crate::iwa::object_index::{ObjectIndex, ResolvedObject};
use crate::iwa::protobuf::tsch;
use prost::Message;

/// Metadata extracted from a chart
#[derive(Debug, Clone)]
pub struct ChartMetadata {
    /// Chart title (if present)
    pub title: Option<String>,
    /// Row names from the chart grid
    pub row_names: Vec<String>,
    /// Column names from the chart grid
    pub column_names: Vec<String>,
    /// Number of data series
    pub series_count: usize,
    /// Chart type (as string representation)
    pub chart_type: String,
    /// Whether chart contains default/sample data
    pub contains_default_data: bool,
}

impl ChartMetadata {
    /// Create a new empty chart metadata
    pub fn new() -> Self {
        Self {
            title: None,
            row_names: Vec::new(),
            column_names: Vec::new(),
            series_count: 0,
            chart_type: String::from("unknown"),
            contains_default_data: false,
        }
    }

    /// Get all text content from the chart
    pub fn all_text(&self) -> Vec<String> {
        let mut text = Vec::new();

        if let Some(ref title) = self.title {
            text.push(title.clone());
        }

        text.extend(self.row_names.clone());
        text.extend(self.column_names.clone());

        text
    }

    /// Check if chart has any meaningful content
    pub fn has_content(&self) -> bool {
        self.title.is_some() || !self.row_names.is_empty() || !self.column_names.is_empty()
    }
}

impl Default for ChartMetadata {
    fn default() -> Self {
        Self::new()
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

        // Find all ChartArchive objects (message types 5000, 5004, 5021)
        for chart_type in [5000u32, 5004, 5021] {
            let chart_entries = self.object_index.find_objects_by_type(chart_type);

            for entry in chart_entries {
                if let Some(resolved) = self.object_index.resolve_object(self.bundle, entry.id)?
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
        for msg in &object.messages {
            if (msg.type_ == 5000 || msg.type_ == 5004 || msg.type_ == 5021)
                && let Ok(chart) = tsch::ChartArchive::decode(&*msg.data)
            {
                return Ok(Some(self.parse_chart(&chart)?));
            }
        }

        Ok(None)
    }

    /// Parse a ChartArchive to extract metadata
    fn parse_chart(&self, chart: &tsch::ChartArchive) -> Result<ChartMetadata> {
        let mut metadata = ChartMetadata::new();

        // Extract chart type
        if let Some(chart_type) = chart.chart_type {
            metadata.chart_type = self.chart_type_to_string(chart_type);
        }

        // Check if contains default data
        metadata.contains_default_data = chart.contains_default_data.unwrap_or(false);

        // Extract grid data (row names, column names)
        if let Some(ref grid) = chart.grid {
            metadata.row_names = grid.row_name.clone();
            metadata.column_names = grid.column_name.clone();
            metadata.series_count = grid.grid_row.len();
        }

        // Try to extract chart title from paragraph styles or other sources
        // Chart titles are typically stored in separate text storage objects
        // referenced by the chart's style properties
        if let Some(title) = self.extract_chart_title(chart)? {
            metadata.title = Some(title);
        }

        Ok(metadata)
    }

    /// Convert chart type enum to string
    fn chart_type_to_string(&self, chart_type: i32) -> String {
        use crate::iwa::protobuf::tsch::ChartType;

        match ChartType::try_from(chart_type) {
            Ok(ChartType::UndefinedChartType) => "Undefined".to_string(),
            Ok(ChartType::ColumnChartType2D) => "Column".to_string(),
            Ok(ChartType::BarChartType2D) => "Bar".to_string(),
            Ok(ChartType::LineChartType2D) => "Line".to_string(),
            Ok(ChartType::AreaChartType2D) => "Area".to_string(),
            Ok(ChartType::PieChartType2D) => "Pie".to_string(),
            Ok(ChartType::StackedColumnChartType2D) => "Stacked Column".to_string(),
            Ok(ChartType::StackedBarChartType2D) => "Stacked Bar".to_string(),
            Ok(ChartType::StackedAreaChartType2D) => "Stacked Area".to_string(),
            Ok(ChartType::ScatterChartType2D) => "Scatter".to_string(),
            Ok(ChartType::MixedChartType2D) => "Mixed".to_string(),
            Ok(ChartType::TwoAxisChartType2D) => "Two Axis".to_string(),
            Ok(ChartType::ColumnChartType3D) => "Column 3D".to_string(),
            Ok(ChartType::BarChartType3D) => "Bar 3D".to_string(),
            Ok(ChartType::LineChartType3D) => "Line 3D".to_string(),
            Ok(ChartType::AreaChartType3D) => "Area 3D".to_string(),
            Ok(ChartType::PieChartType3D) => "Pie 3D".to_string(),
            Ok(ChartType::StackedColumnChartType3D) => "Stacked Column 3D".to_string(),
            Ok(ChartType::StackedBarChartType3D) => "Stacked Bar 3D".to_string(),
            Ok(ChartType::StackedAreaChartType3D) => "Stacked Area 3D".to_string(),
            Ok(ChartType::BubbleChartType2D) => "Bubble".to_string(),
            _ => format!("Unknown({})", chart_type),
        }
    }

    /// Extract chart title from text storage references
    ///
    /// Chart titles are typically stored in TSWP.StorageArchive objects
    /// that are referenced by the chart or its style properties.
    fn extract_chart_title(&self, chart: &tsch::ChartArchive) -> Result<Option<String>> {
        // Check paragraph styles for title text
        for para_ref in &chart.paragraph_styles {
            if let Some(text) = self.extract_text_from_style_ref(para_ref.identifier)?
                && !text.is_empty()
            {
                return Ok(Some(text));
            }
        }

        // Check chart style for title references
        if let Some(ref chart_style_ref) = chart.chart_style
            && let Some(text) = self.extract_text_from_style_ref(chart_style_ref.identifier)?
            && !text.is_empty()
        {
            return Ok(Some(text));
        }

        Ok(None)
    }

    /// Extract text from a style reference
    ///
    /// Styles may reference text storage objects that contain the actual text.
    fn extract_text_from_style_ref(&self, style_id: u64) -> Result<Option<String>> {
        if let Some(_resolved) = self.object_index.resolve_object(self.bundle, style_id)? {
            // Look for associated text storages through dependencies
            if let Some(deps) = self.object_index.get_dependencies(style_id) {
                for &dep_id in deps {
                    if let Some(text) = self.extract_text_from_storage(dep_id)? {
                        return Ok(Some(text));
                    }
                }
            }
        }

        Ok(None)
    }

    /// Extract text from a TSWP.StorageArchive object
    fn extract_text_from_storage(&self, storage_id: u64) -> Result<Option<String>> {
        if let Some(resolved) = self.object_index.resolve_object(self.bundle, storage_id)? {
            for msg in &resolved.messages {
                // TSWP storage types
                if msg.type_ >= 2001
                    && msg.type_ <= 2022
                    && let Ok(storage) =
                        crate::iwa::protobuf::tswp::StorageArchive::decode(&*msg.data)
                    && !storage.text.is_empty()
                {
                    return Ok(Some(storage.text.join(" ")));
                }
            }
        }

        Ok(None)
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

        for chart_type in [5000u32, 5004, 5021] {
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
        let metadata = ChartMetadata::new();
        assert_eq!(metadata.title, None);
        assert_eq!(metadata.series_count, 0);
        assert!(!metadata.has_content());
    }

    #[test]
    fn test_chart_metadata_with_content() {
        let mut metadata = ChartMetadata::new();
        metadata.title = Some("Sales Chart".to_string());
        metadata.row_names = vec!["Q1".to_string(), "Q2".to_string()];
        metadata.column_names = vec!["Revenue".to_string()];

        assert!(metadata.has_content());
        let all_text = metadata.all_text();
        assert_eq!(all_text.len(), 4);
        assert!(all_text.contains(&"Sales Chart".to_string()));
    }

    #[test]
    fn test_chart_type_display() {
        let metadata = ChartMetadata {
            title: None,
            row_names: vec![],
            column_names: vec![],
            series_count: 0,
            chart_type: "Bar".to_string(),
            contains_default_data: false,
        };

        assert_eq!(metadata.chart_type, "Bar");
    }
}
