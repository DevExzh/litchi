//! Structured Data Extraction from iWork Documents
//!
//! This module provides utilities for extracting structured content such as:
//! - Tables from Numbers spreadsheets
//! - Slides from Keynote presentations  
//! - Sections and paragraphs from Pages documents

use std::collections::HashMap;

use crate::Result;
use crate::bundle::Bundle;
use crate::charts::metadata_extractor::ChartMetadataExtractor;
use crate::numbers::table_extractor::TableDataExtractor;
use crate::object_index::ObjectIndex;
use crate::shapes::text_extractor::ShapeTextExtractor;

/// Represents a table extracted from a Numbers document
#[derive(Debug, Clone)]
pub struct Table {
    name: String,
    row_count: usize,
    column_count: usize,
    cells: HashMap<(usize, usize), CellValue>,
}

impl Table {
    /// Create a new empty table
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            row_count: 0,
            column_count: 0,
            cells: HashMap::new(),
        }
    }

    /// Create a table with dimensions declared by the source document.
    fn with_dimensions(name: impl Into<String>, row_count: usize, column_count: usize) -> Self {
        Self {
            name: name.into(),
            row_count,
            column_count,
            cells: HashMap::new(),
        }
    }

    /// Borrow the table name.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Return the number of addressable rows.
    pub const fn row_count(&self) -> usize {
        self.row_count
    }

    /// Return the number of addressable columns.
    pub const fn column_count(&self) -> usize {
        self.column_count
    }

    /// Return the table dimensions as `(rows, columns)`.
    pub const fn dimensions(&self) -> (usize, usize) {
        (self.row_count, self.column_count)
    }

    /// Get a cell value at the specified position
    pub fn get_cell(&self, row: usize, col: usize) -> Option<&CellValue> {
        self.cells.get(&(row, col))
    }

    /// Iterate over materialized cells without exposing the backing map.
    pub fn iter_cells(&self) -> impl Iterator<Item = ((usize, usize), &CellValue)> + '_ {
        self.cells
            .iter()
            .map(|(position, value)| (*position, value))
    }

    /// Return the number of materialized cells, including explicit empty cells.
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    /// Set a cell value at the specified position
    pub fn set_cell(&mut self, row: usize, col: usize, value: CellValue) {
        self.cells.insert((row, col), value);
        self.row_count = self.row_count.max(row + 1);
        self.column_count = self.column_count.max(col + 1);
    }

    /// Convert table to CSV format
    pub fn to_csv(&self) -> String {
        let mut csv = String::new();
        for row in 0..self.row_count {
            for col in 0..self.column_count {
                if col > 0 {
                    csv.push(',');
                }
                if let Some(cell) = self.get_cell(row, col) {
                    csv.push_str(&cell.to_string());
                }
            }
            csv.push('\n');
        }
        csv
    }
}

/// Represents a cell value in a table
#[derive(Debug, Clone)]
pub enum CellValue {
    /// Text/string value
    Text(String),
    /// Numeric value
    Number(f64),
    /// Boolean value
    Boolean(bool),
    /// Seconds since Numbers' 2001-01-01 UTC epoch.
    Date(f64),
    /// Duration in seconds.
    Duration(f64),
    /// Formula (stored as string)
    Formula(String),
    /// Spreadsheet error value.
    Error(String),
    /// Empty cell
    Empty,
}

impl CellValue {
    /// Check if cell is empty
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }
}

impl std::fmt::Display for CellValue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            CellValue::Text(s) => write!(f, "{}", s),
            CellValue::Number(n) => write!(f, "{}", n),
            CellValue::Boolean(b) => write!(f, "{}", b),
            CellValue::Date(d) => write!(f, "{}", d),
            CellValue::Duration(d) => write!(f, "{}", d),
            CellValue::Formula(formula) => write!(f, "{}", formula),
            CellValue::Error(error) => write!(f, "ERROR: {}", error),
            CellValue::Empty => Ok(()),
        }
    }
}

/// Represents a slide in a Keynote presentation
#[derive(Debug, Clone)]
pub struct Slide {
    /// Slide index (0-based)
    pub index: usize,
    /// Slide title
    pub title: Option<String>,
    /// Text content on the slide
    pub text_content: Vec<String>,
    /// Notes associated with the slide
    pub notes: Option<String>,
}

impl Slide {
    /// Create a new slide
    pub fn new(index: usize) -> Self {
        Self {
            index,
            title: None,
            text_content: Vec::new(),
            notes: None,
        }
    }

    /// Get all text from the slide (title + content + notes)
    pub fn all_text(&self) -> Vec<String> {
        let mut all = Vec::new();
        if let Some(ref title) = self.title {
            all.push(title.clone());
        }
        all.extend(self.text_content.clone());
        if let Some(ref notes) = self.notes {
            all.push(notes.clone());
        }
        all
    }
}

/// Represents a section in a Pages document
#[derive(Debug, Clone)]
pub struct Section {
    /// Section index (0-based)
    pub index: usize,
    /// Section heading
    pub heading: Option<String>,
    /// Paragraphs in this section
    pub paragraphs: Vec<String>,
}

impl Section {
    /// Create a new section
    pub fn new(index: usize) -> Self {
        Self {
            index,
            heading: None,
            paragraphs: Vec::new(),
        }
    }

    /// Get all text from the section
    pub fn all_text(&self) -> Vec<String> {
        let mut all = Vec::new();
        if let Some(ref heading) = self.heading {
            all.push(heading.clone());
        }
        all.extend(self.paragraphs.clone());
        all
    }
}

/// Extract tables from a Numbers document
///
/// Uses the TableDataExtractor to parse complete table structures including
/// cell values, formulas, and formatting information.
pub fn extract_tables(bundle: &Bundle, object_index: &ObjectIndex) -> Result<Vec<Table>> {
    let extractor = TableDataExtractor::new(bundle, object_index);
    let numbers_tables = extractor.extract_all_tables()?;

    // Convert NumbersTable to our Table type for compatibility
    let tables = numbers_tables
        .into_iter()
        .map(|nt| {
            let (row_count, column_count) = nt.dimensions();
            let table_name = nt.name().to_owned();
            let (cells, _) = nt.into_parts();
            let mut table = Table::with_dimensions(table_name, row_count, column_count);

            // Convert cells from NumbersTable format to our CellValue format
            for ((row, col), cell) in cells {
                let cell_value = convert_numbers_cell_to_structured(cell);
                table.set_cell(row, col, cell_value);
            }

            table
        })
        .collect();

    Ok(tables)
}

/// Convert Numbers CellValue to structured CellValue
fn convert_numbers_cell_to_structured(cell: crate::numbers::CellValue) -> CellValue {
    use crate::numbers::CellValue as NC;

    match cell {
        NC::Empty => CellValue::Empty,
        NC::Text(s) => CellValue::Text(s),
        NC::Number(n) => CellValue::Number(n),
        NC::Boolean(b) => CellValue::Boolean(b),
        NC::Date(d) => CellValue::Date(d),
        NC::Duration(value) => CellValue::Duration(value),
        NC::Formula(f) => CellValue::Formula(f),
        NC::Error(e) => CellValue::Error(e),
    }
}

/// Extract slides from a Keynote presentation
pub fn extract_slides(bundle: &Bundle, object_index: &ObjectIndex) -> Result<Vec<Slide>> {
    use prost::Message;

    let mut slides = Vec::new();
    let Some(document_object) = bundle_object(bundle, 1) else {
        return Ok(slides);
    };
    let Some(document) = document_object.messages.iter().find_map(|message| {
        crate::protobuf::kn::DocumentArchive::decode(message.data.as_slice()).ok()
    }) else {
        return Ok(slides);
    };
    let Some(show_object) = bundle_object(bundle, document.show.identifier) else {
        return Ok(slides);
    };
    let Some(show) = show_object
        .messages
        .iter()
        .find_map(|message| crate::protobuf::kn::ShowArchive::decode(message.data.as_slice()).ok())
    else {
        return Ok(slides);
    };

    for node_reference in show.slide_tree.slides {
        let Some(node_object) = bundle_object(bundle, node_reference.identifier) else {
            continue;
        };
        let Some(node) = node_object.messages.iter().find_map(|message| {
            crate::protobuf::kn::SlideNodeArchive::decode(message.data.as_slice()).ok()
        }) else {
            continue;
        };
        let Some(slide_reference) = node.slide else {
            continue;
        };
        let Some(slide_object) = bundle_object(bundle, slide_reference.identifier) else {
            continue;
        };
        let Some(archive) = slide_object.messages.iter().find_map(|message| {
            crate::protobuf::kn::SlideArchive::decode(message.data.as_slice()).ok()
        }) else {
            continue;
        };

        let index = slides.len();
        let mut slide = Slide::new(index);
        slide.title = archive.name.filter(|name| !name.is_empty());
        let title_placeholder = archive
            .title_placeholder
            .as_ref()
            .map(|reference| reference.identifier);
        let body_placeholder = archive
            .body_placeholder
            .as_ref()
            .map(|reference| reference.identifier);

        if let Some(identifier) = title_placeholder
            && let Some(text) = drawable_text(bundle, object_index, identifier)?
        {
            slide.title = Some(text);
        }
        if let Some(identifier) = body_placeholder
            && let Some(text) = drawable_text(bundle, object_index, identifier)?
        {
            slide.text_content.push(text);
        }
        for drawable in archive.owned_drawables {
            if Some(drawable.identifier) == title_placeholder
                || Some(drawable.identifier) == body_placeholder
            {
                continue;
            }
            if let Some(text) = drawable_text(bundle, object_index, drawable.identifier)? {
                slide.text_content.push(text);
            }
        }
        if let Some(note) = archive.note
            && let Some(note_object) = object_index.resolve_ref_id(bundle, note.identifier)?
        {
            for message in note_object.messages {
                let Ok(note) = crate::protobuf::kn::NoteArchive::decode(message.data.as_slice())
                else {
                    continue;
                };
                if let Some(storage) = object_index
                    .resolve_ref_id(bundle, note.contained_storage.identifier)?
                    .and_then(|object| {
                        object.messages.iter().find_map(|message| {
                            crate::protobuf::tswp::StorageArchive::decode(message.data.as_slice())
                                .ok()
                        })
                    })
                {
                    let text = storage.text.concat();
                    if !text.is_empty() {
                        slide.notes = Some(text);
                    }
                }
                break;
            }
        }
        slides.push(slide);
    }

    Ok(slides)
}

fn bundle_object(bundle: &Bundle, identifier: u64) -> Option<&crate::archive::ArchiveObject> {
    bundle
        .iter_archives()
        .map(|(_, archive)| archive)
        .find_map(|archive| archive.object(identifier))
}

fn drawable_text(
    bundle: &Bundle,
    object_index: &ObjectIndex,
    identifier: u64,
) -> Result<Option<String>> {
    use prost::Message;

    let Some(drawable) = object_index.resolve_ref_id(bundle, identifier)? else {
        return Ok(None);
    };
    let storage_id = drawable.messages.iter().find_map(|message| {
        crate::protobuf::kn::PlaceholderArchive::decode(message.data.as_slice())
            .ok()
            .and_then(|placeholder| placeholder.super_.owned_storage)
            .or_else(|| {
                crate::protobuf::tswp::ShapeInfoArchive::decode(message.data.as_slice())
                    .ok()
                    .and_then(|shape| shape.owned_storage)
            })
            .map(|reference| reference.identifier)
    });
    let Some(storage_id) = storage_id else {
        return Ok(None);
    };
    let Some(storage_object) = object_index.resolve_ref_id(bundle, storage_id)? else {
        return Ok(None);
    };
    for message in storage_object.messages {
        if let Ok(storage) = crate::protobuf::tswp::StorageArchive::decode(message.data.as_slice())
        {
            let text = storage.text.concat();
            return Ok((!text.is_empty()).then_some(text));
        }
    }
    Ok(None)
}

/// Extract the main body section from a Pages document.
///
/// Current Pages packages store the main text reference directly in
/// `TP.DocumentArchive.body_storage`. `TP.SectionArchive` describes page and
/// header/footer properties; it is not a container for the document body.
pub fn extract_sections(bundle: &Bundle, object_index: &ObjectIndex) -> Result<Vec<Section>> {
    use prost::Message;

    let Some(document_object) = bundle
        .get_archive("Index/Document.iwa")
        .and_then(|archive| archive.object(1))
    else {
        return Ok(Vec::new());
    };
    let Some(document) = document_object
        .messages
        .iter()
        .find(|message| message.type_ == 10000)
        .and_then(|message| {
            crate::protobuf::tp::DocumentArchive::decode(message.data.as_slice()).ok()
        })
    else {
        return Ok(Vec::new());
    };

    let mut section = Section::new(0);
    if let Some(reference) = document.body_storage {
        let object = object_index
            .resolve_ref_id(bundle, reference.identifier)?
            .ok_or_else(|| {
                crate::Error::InvalidFormat(format!(
                    "Pages body storage object {} is missing",
                    reference.identifier
                ))
            })?;
        let storage = object
            .messages
            .iter()
            .filter(|message| message.type_ == 2001 || message.type_ == 2022)
            .find_map(|message| {
                crate::protobuf::tswp::StorageArchive::decode(message.data.as_slice()).ok()
            })
            .ok_or_else(|| {
                crate::Error::InvalidFormat(format!(
                    "Pages body object {} has no text storage payload",
                    reference.identifier
                ))
            })?;
        let text = storage.text.concat();
        if !text.is_empty() {
            section.paragraphs.push(text);
        }
    }

    Ok(vec![section])
}

/// Extract all structured data from a document based on its type
///
/// This function uses specialized extractors for each content type:
/// - TableDataExtractor for Numbers tables with full cell parsing
/// - ShapeTextExtractor for text in shapes and text boxes
/// - ChartMetadataExtractor for chart data
pub fn extract_all(bundle: &Bundle, object_index: &ObjectIndex) -> Result<StructuredData> {
    let tables = extract_tables(bundle, object_index)?;
    let slides = extract_slides(bundle, object_index)?;
    let sections = extract_sections(bundle, object_index)?;

    Ok(StructuredData {
        tables,
        slides,
        sections,
    })
}

/// Extract text from shapes and text boxes
///
/// This extracts text content from TSD.ShapeArchive objects, including
/// text boxes, callouts, and grouped shapes.
pub fn extract_shape_text(bundle: &Bundle, object_index: &ObjectIndex) -> Result<Vec<String>> {
    let extractor = ShapeTextExtractor::new(bundle, object_index);
    extractor.extract_all_shape_text()
}

/// Extract chart metadata
///
/// Returns metadata from all charts in the document, including titles,
/// row/column names, and data series information.
pub fn extract_chart_metadata(
    bundle: &Bundle,
    object_index: &ObjectIndex,
) -> Result<Vec<crate::charts::ChartMetadata>> {
    let extractor = ChartMetadataExtractor::new(bundle, object_index);
    extractor.extract_all_charts()
}

/// Container for all structured data extracted from a document
#[derive(Debug, Clone)]
pub struct StructuredData {
    /// Tables (primarily from Numbers)
    pub tables: Vec<Table>,
    /// Slides (from Keynote)
    pub slides: Vec<Slide>,
    /// Sections (from Pages)
    pub sections: Vec<Section>,
}

impl StructuredData {
    /// Check if any structured data was extracted
    pub fn is_empty(&self) -> bool {
        self.tables.is_empty() && self.slides.is_empty() && self.sections.is_empty()
    }

    /// Get summary statistics
    pub fn summary(&self) -> String {
        format!(
            "Tables: {}, Slides: {}, Sections: {}",
            self.tables.len(),
            self.slides.len(),
            self.sections.len()
        )
    }

    /// Extract all text from all structured elements
    pub fn all_text(&self) -> Vec<String> {
        let mut all_text = Vec::new();

        // Add table names
        for table in &self.tables {
            all_text.push(format!("Table: {}", table.name()));
        }

        // Add slide content
        for slide in &self.slides {
            all_text.extend(slide.all_text());
        }

        // Add section content
        for section in &self.sections {
            all_text.extend(section.all_text());
        }

        all_text
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_table_creation() {
        let mut table = Table::new("Test Table".to_string());
        assert_eq!(table.name(), "Test Table");
        assert_eq!(table.row_count(), 0);
        assert_eq!(table.column_count(), 0);

        table.set_cell(0, 0, CellValue::Text("Header 1".to_string()));
        table.set_cell(0, 1, CellValue::Text("Header 2".to_string()));
        table.set_cell(1, 0, CellValue::Number(42.0));
        table.set_cell(1, 1, CellValue::Boolean(true));

        assert_eq!(table.row_count(), 2);
        assert_eq!(table.column_count(), 2);
        assert_eq!(table.dimensions(), (2, 2));
        assert_eq!(table.cell_count(), 4);
        assert_eq!(table.iter_cells().count(), 4);

        let csv = table.to_csv();
        assert!(csv.contains("Header 1"));
        assert!(csv.contains("42"));
    }

    #[test]
    fn test_cell_value() {
        let text_cell = CellValue::Text("Hello".to_string());
        assert_eq!(text_cell.to_string(), "Hello");
        assert!(!text_cell.is_empty());

        let empty_cell = CellValue::Empty;
        assert_eq!(empty_cell.to_string(), "");
        assert!(empty_cell.is_empty());

        let number_cell = CellValue::Number(std::f64::consts::PI);
        assert_eq!(number_cell.to_string(), "3.141592653589793");
    }

    #[test]
    fn test_slide_creation() {
        let mut slide = Slide::new(0);
        assert_eq!(slide.index, 0);
        assert_eq!(slide.title, None);

        slide.title = Some("Introduction".to_string());
        slide.text_content.push("Point 1".to_string());
        slide.text_content.push("Point 2".to_string());
        slide.notes = Some("Speaker notes".to_string());

        let all_text = slide.all_text();
        assert_eq!(all_text.len(), 4);
        assert_eq!(all_text[0], "Introduction");
        assert_eq!(all_text[3], "Speaker notes");
    }

    #[test]
    fn test_section_creation() {
        let mut section = Section::new(0);
        section.heading = Some("Chapter 1".to_string());
        section.paragraphs.push("First paragraph.".to_string());
        section.paragraphs.push("Second paragraph.".to_string());

        let all_text = section.all_text();
        assert_eq!(all_text.len(), 3);
        assert_eq!(all_text[0], "Chapter 1");
    }

    #[test]
    fn test_structured_data() {
        let mut table = Table::new("Data".to_string());
        table.set_cell(0, 0, CellValue::Text("A".to_string()));

        let mut slide = Slide::new(0);
        slide.title = Some("Title".to_string());

        let mut section = Section::new(0);
        section.heading = Some("Heading".to_string());

        let data = StructuredData {
            tables: vec![table],
            slides: vec![slide],
            sections: vec![section],
        };

        assert!(!data.is_empty());
        let summary = data.summary();
        assert!(summary.contains("Tables: 1"));
        assert!(summary.contains("Slides: 1"));
        assert!(summary.contains("Sections: 1"));
    }
}
