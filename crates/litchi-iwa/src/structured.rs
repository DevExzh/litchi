//! Structured Data Extraction from iWork Documents
//!
//! This module provides utilities for extracting structured content such as:
//! - Tables from Numbers spreadsheets
//! - Slides from Keynote presentations  
//! - Sections and paragraphs from Pages documents

use crate::Result;
use crate::bundle::Bundle;
use crate::numbers::table_extractor::TableDataExtractor;
use crate::object_index::ObjectIndex;
use litchi_keynote::Slide;
use litchi_numbers::Table;
use litchi_pages::{Section, SectionType};

/// Extract tables from a Numbers document
///
/// Uses the TableDataExtractor to parse complete table structures including
/// cell values, formulas, and formatting information.
pub fn extract_tables(bundle: &Bundle, object_index: &ObjectIndex) -> Result<Vec<Table>> {
    let extractor = TableDataExtractor::new(bundle, object_index);
    let numbers_tables = extractor.extract_all_tables()?;

    numbers_tables
        .into_iter()
        .map(crate::numbers::table::NumbersTable::into_semantic_table)
        .collect()
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
        let mut slide = Slide::builder(index);
        slide.set_title(archive.name.filter(|name| !name.is_empty()));
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
            slide.set_title(Some(text));
        }
        if let Some(identifier) = body_placeholder
            && let Some(text) = drawable_text(bundle, object_index, identifier)?
        {
            slide.push_text(text);
        }
        for drawable in archive.owned_drawables {
            if Some(drawable.identifier) == title_placeholder
                || Some(drawable.identifier) == body_placeholder
            {
                continue;
            }
            if let Some(text) = drawable_text(bundle, object_index, drawable.identifier)? {
                slide.push_text(text);
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
                        slide.set_notes(Some(text));
                    }
                }
                break;
            }
        }
        slides.push(slide.build());
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

    let mut section = Section::new(0, SectionType::Body);
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
/// This function uses specialized extractors for Numbers tables, Keynote
/// slides, and Pages sections. Shape and chart readers remain separate
/// format-facing modules because they require native archive context.
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
        let mut builder = Table::builder("Test Table", litchi_numbers::Dimensions::new(2, 2));
        assert_eq!(builder.name(), "Test Table");

        assert!(
            builder
                .set(
                    litchi_numbers::Position::new(0, 0),
                    litchi_numbers::cell::Value::Text("Header 1".to_owned())
                )
                .is_ok()
        );
        assert!(
            builder
                .set(
                    litchi_numbers::Position::new(0, 1),
                    litchi_numbers::cell::Value::Text("Header 2".to_owned())
                )
                .is_ok()
        );
        assert!(
            builder
                .set(
                    litchi_numbers::Position::new(1, 0),
                    litchi_numbers::cell::Value::Number(42.0)
                )
                .is_ok()
        );
        assert!(
            builder
                .set(
                    litchi_numbers::Position::new(1, 1),
                    litchi_numbers::cell::Value::Boolean(true)
                )
                .is_ok()
        );

        let table = builder.finish().expect("valid table builder");
        assert_eq!(table.name(), "Test Table");
        assert_eq!(table.row_count(), 2);
        assert_eq!(table.column_count(), 2);
        assert_eq!(table.dimensions().rows(), 2);
        assert_eq!(table.dimensions().columns(), 2);
        assert_eq!(table.cell_count(), 4);
        assert_eq!(table.iter_cells().count(), 4);

        let csv = table.to_csv();
        assert!(csv.contains("Header 1"));
        assert!(csv.contains("42"));
    }

    #[test]
    fn test_cell_value() {
        let text_cell = litchi_numbers::cell::Value::Text("Hello".to_owned());
        assert_eq!(text_cell.to_string(), "Hello");
        assert!(!text_cell.is_empty());

        let empty_cell = litchi_numbers::cell::Value::Empty;
        assert_eq!(empty_cell.to_string(), "");
        assert!(empty_cell.is_empty());

        let number_cell = litchi_numbers::cell::Value::Number(std::f64::consts::PI);
        assert_eq!(number_cell.to_string(), "3.141592653589793");
    }

    #[test]
    fn test_slide_creation() {
        let builder = Slide::builder(0);
        let slide = builder.build();
        assert_eq!(slide.index(), 0);
        assert_eq!(slide.title(), None);

        let mut builder = Slide::builder(0);
        builder.set_title(Some("Introduction".to_string()));
        builder.push_text("Point 1".to_string());
        builder.push_text("Point 2".to_string());
        builder.set_notes(Some("Speaker notes".to_string()));
        let slide = builder.build();

        let all_text = slide.all_text();
        assert_eq!(all_text.len(), 4);
        assert_eq!(all_text[0], "Introduction");
        assert_eq!(all_text[3], "Speaker notes");
    }

    #[test]
    fn test_section_creation() {
        let mut section = Section::new(0, SectionType::Body);
        section.heading = Some("Chapter 1".to_string());
        section.paragraphs.push("First paragraph.".to_string());
        section.paragraphs.push("Second paragraph.".to_string());

        let all_text = section.all_text();
        assert_eq!(all_text.len(), 3);
        assert_eq!(all_text[0], "Chapter 1");
    }

    #[test]
    fn test_structured_data() {
        let mut table_builder = Table::builder("Data", litchi_numbers::Dimensions::new(1, 1));
        assert!(
            table_builder
                .set(
                    litchi_numbers::Position::new(0, 0),
                    litchi_numbers::cell::Value::Text("A".to_owned())
                )
                .is_ok()
        );
        let table = table_builder.finish().expect("valid table builder");

        let mut slide_builder = Slide::builder(0);
        slide_builder.set_title(Some("Title".to_string()));
        let slide = slide_builder.build();

        let mut section = Section::new(0, SectionType::Body);
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
