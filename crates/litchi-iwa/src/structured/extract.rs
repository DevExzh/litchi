//! Application-specific structured-data extraction.

use super::{CellValue, Section, Slide, StructuredData, Table};
use crate::Result;
use crate::bundle::Bundle;
use crate::charts::metadata_extractor::ChartMetadataExtractor;
use crate::numbers::table_extractor::TableDataExtractor;
use crate::object_index::ObjectIndex;
use crate::shapes::text_extractor::ShapeTextExtractor;

/// Extract tables from a Numbers document.
pub fn extract_tables(bundle: &Bundle, object_index: &ObjectIndex) -> Result<Vec<Table>> {
    let extractor = TableDataExtractor::new(bundle, object_index);
    let numbers_tables = extractor.extract_all_tables()?;

    Ok(numbers_tables
        .into_iter()
        .map(|numbers_table| {
            let mut table = Table::new(numbers_table.name);
            table.row_count = numbers_table.row_count;
            table.column_count = numbers_table.column_count;
            table.cells = numbers_table
                .cells
                .into_iter()
                .map(|(position, cell)| (position, convert_numbers_cell(cell)))
                .collect();
            table
        })
        .collect())
}

fn convert_numbers_cell(cell: crate::numbers::CellValue) -> CellValue {
    use crate::numbers::CellValue as NumbersCell;

    match cell {
        NumbersCell::Empty => CellValue::Empty,
        NumbersCell::Text(value) => CellValue::Text(value),
        NumbersCell::Number(value) => CellValue::Number(value),
        NumbersCell::Boolean(value) => CellValue::Boolean(value),
        NumbersCell::Date(value) => CellValue::Date(value),
        NumbersCell::Duration(_) => CellValue::Empty,
        NumbersCell::Formula(value) => CellValue::Formula(value),
        NumbersCell::Error(value) => CellValue::Text(format!("ERROR: {value}")),
    }
}

/// Extract slides from a Keynote presentation.
pub fn extract_slides(bundle: &Bundle, _object_index: &ObjectIndex) -> Result<Vec<Slide>> {
    Ok(bundle
        .find_objects_by_type(1102)
        .iter()
        .enumerate()
        .map(|(index, (_, object))| {
            let mut slide = Slide::new(index);
            let text = object.extract_text();
            slide.title = text.first().cloned();
            slide.text_content = text.into_iter().skip(1).collect();
            slide
        })
        .collect())
}

/// Extract sections from a Pages document.
pub fn extract_sections(bundle: &Bundle, object_index: &ObjectIndex) -> Result<Vec<Section>> {
    let mut sections = extract_marked_sections(bundle, object_index)?;
    if sections.is_empty() {
        sections = extract_storage_sections(bundle);
    }
    Ok(sections)
}

fn extract_marked_sections(bundle: &Bundle, object_index: &ObjectIndex) -> Result<Vec<Section>> {
    let mut sections = Vec::new();
    for (index, (_, section_object)) in bundle.find_objects_by_type(10011).iter().enumerate() {
        let mut section = Section::new(index);
        if let Some(section_id) = section_object.archive_info.identifier
            && let Some(reference_ids) = object_index.get_dependencies(section_id)
        {
            for &reference_id in reference_ids {
                if let Ok(Some(object)) = object_index.resolve_object(bundle, reference_id) {
                    append_section_text(&mut section, storage_text(&object.messages));
                }
            }
        }
        append_section_text(&mut section, section_object.extract_text());
        if section.heading.is_some() || !section.paragraphs.is_empty() {
            sections.push(section);
        }
    }
    Ok(sections)
}

fn storage_text(messages: &[crate::archive::RawMessage]) -> Vec<String> {
    use prost::Message;

    messages
        .iter()
        .filter(|message| matches!(message.type_, 2001 | 2022))
        .flat_map(|message| {
            crate::protobuf::tswp::StorageArchive::decode(&*message.data)
                .map(|storage| storage.text)
                .unwrap_or_default()
        })
        .collect()
}

fn extract_storage_sections(bundle: &Bundle) -> Vec<Section> {
    bundle
        .find_objects_by_type(2022)
        .iter()
        .enumerate()
        .filter_map(|(index, (_, object))| {
            let mut section = Section::new(index);
            append_section_text(&mut section, object.extract_text());
            (section.heading.is_some() || !section.paragraphs.is_empty()).then_some(section)
        })
        .collect()
}

fn append_section_text(section: &mut Section, text: Vec<String>) {
    let mut text = text.into_iter();
    if section.heading.is_none() {
        section.heading = text.next();
    }
    section.paragraphs.extend(text);
}

/// Extract all structured data from a document.
pub fn extract_all(bundle: &Bundle, object_index: &ObjectIndex) -> Result<StructuredData> {
    Ok(StructuredData {
        tables: extract_tables(bundle, object_index)?,
        slides: extract_slides(bundle, object_index)?,
        sections: extract_sections(bundle, object_index)?,
    })
}

/// Extract text from shapes and text boxes.
pub fn extract_shape_text(bundle: &Bundle, object_index: &ObjectIndex) -> Result<Vec<String>> {
    ShapeTextExtractor::new(bundle, object_index).extract_all_shape_text()
}

/// Extract chart metadata.
pub fn extract_chart_metadata(
    bundle: &Bundle,
    object_index: &ObjectIndex,
) -> Result<Vec<crate::charts::ChartMetadata>> {
    ChartMetadataExtractor::new(bundle, object_index).extract_all_charts()
}
