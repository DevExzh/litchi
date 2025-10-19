//! Structured Data Extraction from iWork Documents
//!
//! This module provides utilities for extracting structured content such as:
//! - Tables from Numbers spreadsheets
//! - Slides from Keynote presentations  
//! - Sections and paragraphs from Pages documents

use std::collections::HashMap;

use crate::iwa::bundle::Bundle;
use crate::iwa::object_index::ObjectIndex;
use crate::iwa::Result;

/// Represents a table extracted from a Numbers document
#[derive(Debug, Clone)]
pub struct Table {
    /// Table name
    pub name: String,
    /// Number of rows
    pub row_count: usize,
    /// Number of columns
    pub column_count: usize,
    /// Cell data (row, column) -> value
    pub cells: HashMap<(usize, usize), CellValue>,
}

impl Table {
    /// Create a new empty table
    pub fn new(name: String) -> Self {
        Self {
            name,
            row_count: 0,
            column_count: 0,
            cells: HashMap::new(),
        }
    }

    /// Get a cell value at the specified position
    pub fn get_cell(&self, row: usize, col: usize) -> Option<&CellValue> {
        self.cells.get(&(row, col))
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
    /// Date value (as string)
    Date(String),
    /// Formula (stored as string)
    Formula(String),
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
            CellValue::Formula(formula) => write!(f, "{}", formula),
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
pub fn extract_tables(bundle: &Bundle, _object_index: &ObjectIndex) -> Result<Vec<Table>> {
    let mut tables = Vec::new();

    // Find all table model objects (message type 100)
    let table_objects = bundle.find_objects_by_type(100);

    for (_archive_name, object) in table_objects {
        for decoded_msg in &object.decoded_messages {
            // Try to downcast to TableModelWrapper
            if let Some(table_msg) = decoded_msg.extract_text().first() {
                let table = Table::new(table_msg.clone());
                // In a full implementation, we would parse the table structure
                // and extract cell data. For now, we just create a table with the name.
                tables.push(table);
            }
        }
    }

    Ok(tables)
}

/// Extract slides from a Keynote presentation
pub fn extract_slides(bundle: &Bundle, _object_index: &ObjectIndex) -> Result<Vec<Slide>> {
    let mut slides = Vec::new();

    // Find all slide objects (message type 1102 based on our decoder map)
    let slide_objects = bundle.find_objects_by_type(1102);

    for (index, (_archive_name, object)) in slide_objects.iter().enumerate() {
        let mut slide = Slide::new(index);

        // Extract text content from the slide
        let text_parts = object.extract_text();
        if !text_parts.is_empty() {
            slide.title = text_parts.first().cloned();
            slide.text_content = text_parts.into_iter().skip(1).collect();
        }

        slides.push(slide);
    }

    Ok(slides)
}

/// Extract sections from a Pages document
pub fn extract_sections(bundle: &Bundle, _object_index: &ObjectIndex) -> Result<Vec<Section>> {
    let mut sections = Vec::new();

    // In a full implementation, we would identify section boundaries
    // For now, we'll treat all text storage objects as potential sections
    let text_objects = bundle.find_objects_by_type(2022); // Common TSWP storage type

    for (index, (_archive_name, object)) in text_objects.iter().enumerate() {
        let mut section = Section::new(index);

        // Extract text content
        let text_parts = object.extract_text();
        if !text_parts.is_empty() {
            section.heading = text_parts.first().cloned();
            section.paragraphs = text_parts.into_iter().skip(1).collect();
        }

        if section.heading.is_some() || !section.paragraphs.is_empty() {
            sections.push(section);
        }
    }

    Ok(sections)
}

/// Extract all structured data from a document based on its type
pub fn extract_all(
    bundle: &Bundle,
    object_index: &ObjectIndex,
) -> Result<StructuredData> {
    let tables = extract_tables(bundle, object_index)?;
    let slides = extract_slides(bundle, object_index)?;
    let sections = extract_sections(bundle, object_index)?;

    Ok(StructuredData {
        tables,
        slides,
        sections,
    })
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
            all_text.push(format!("Table: {}", table.name));
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
        assert_eq!(table.name, "Test Table");
        assert_eq!(table.row_count, 0);
        assert_eq!(table.column_count, 0);

        table.set_cell(0, 0, CellValue::Text("Header 1".to_string()));
        table.set_cell(0, 1, CellValue::Text("Header 2".to_string()));
        table.set_cell(1, 0, CellValue::Number(42.0));
        table.set_cell(1, 1, CellValue::Boolean(true));

        assert_eq!(table.row_count, 2);
        assert_eq!(table.column_count, 2);

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

        let number_cell = CellValue::Number(3.14);
        assert_eq!(number_cell.to_string(), "3.14");
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

