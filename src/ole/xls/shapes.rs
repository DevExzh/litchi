//! Shape extraction for Excel workbooks.
//!
//! Excel workbooks store drawing objects in MsoDrawing and MsoDrawingGroup records.
//! This module provides access to these shapes using the shared Escher infrastructure.

use crate::ole::escher::{EscherShape, EscherShapeFactory, extract_text_from_escher};
use crate::ole::xls::records::RecordIter;
use std::io::Cursor;

/// Shape information extracted from an Excel workbook.
#[derive(Debug, Clone)]
pub struct XlsShape {
    /// Shape type (rectangle, ellipse, line, etc.)
    pub shape_type: crate::ole::escher::EscherShapeType,
    /// Shape ID
    pub shape_id: u32,
    /// Text content extracted from the shape (if any)
    pub text: Option<String>,
    /// Whether this is a group shape
    pub is_group: bool,
    /// Child shapes (for group shapes)
    pub children: Vec<XlsShape>,
}

impl XlsShape {
    /// Create an XlsShape from an EscherShape.
    fn from_escher(escher_shape: &EscherShape) -> Self {
        Self {
            shape_type: escher_shape.shape_type,
            shape_id: escher_shape.shape_id.unwrap_or(0),
            text: escher_shape.text.clone(),
            is_group: escher_shape.is_group,
            children: escher_shape
                .children
                .iter()
                .map(Self::from_escher)
                .collect(),
        }
    }
}

/// Extract shapes from Excel workbook MsoDrawing records.
///
/// # Arguments
///
/// * `workbook_data` - The raw workbook stream data
///
/// # Returns
///
/// A vector of shapes found in the workbook.
pub fn extract_shapes_from_workbook(workbook_data: &[u8]) -> std::io::Result<Vec<XlsShape>> {
    let mut all_shapes = Vec::new();
    let mut drawing_data = Vec::new();

    // Parse BIFF records to find MsoDrawing records
    let cursor = Cursor::new(workbook_data);
    if let Ok(mut record_iter) = RecordIter::new(cursor) {
        while let Some(Ok(record)) = record_iter.next() {
            match record.header.record_type {
                // MsoDrawing (0x00EC) - contains Escher data
                0x00EC => {
                    drawing_data.extend_from_slice(&record.data);
                },
                // MsoDrawingGroup (0x00EB) - contains drawing group Escher data
                0x00EB => {
                    // Parse drawing group data if needed
                    if !record.data.is_empty() {
                        let shapes = EscherShapeFactory::extract_shapes_from_drawing(&record.data)?;
                        all_shapes.extend(shapes.iter().map(XlsShape::from_escher));
                    }
                },
                _ => {},
            }
        }
    }

    // Parse accumulated drawing data
    if !drawing_data.is_empty() {
        let shapes = EscherShapeFactory::extract_shapes_from_drawing(&drawing_data)?;
        all_shapes.extend(shapes.iter().map(XlsShape::from_escher));
    }

    Ok(all_shapes)
}

/// Extract text from all shapes in an Excel workbook.
///
/// # Arguments
///
/// * `workbook_data` - The raw workbook stream data
///
/// # Returns
///
/// A string containing all text extracted from shapes.
pub fn extract_shape_text_from_workbook(workbook_data: &[u8]) -> std::io::Result<String> {
    let mut all_text = String::new();
    let mut drawing_data = Vec::new();

    let cursor = Cursor::new(workbook_data);
    if let Ok(mut record_iter) = RecordIter::new(cursor) {
        while let Some(Ok(record)) = record_iter.next() {
            if record.header.record_type == 0x00EC || record.header.record_type == 0x00EB {
                drawing_data.extend_from_slice(&record.data);
            }
        }
    }

    if !drawing_data.is_empty() {
        all_text = extract_text_from_escher(&drawing_data)?;
    }

    Ok(all_text)
}

/// Count shapes in an Excel workbook.
///
/// # Arguments
///
/// * `workbook_data` - The raw workbook stream data
///
/// # Returns
///
/// The number of shapes found.
pub fn count_shapes_in_workbook(workbook_data: &[u8]) -> usize {
    let mut drawing_data = Vec::new();

    let cursor = Cursor::new(workbook_data);
    if let Ok(mut record_iter) = RecordIter::new(cursor) {
        while let Some(Ok(record)) = record_iter.next() {
            if record.header.record_type == 0x00EC || record.header.record_type == 0x00EB {
                drawing_data.extend_from_slice(&record.data);
            }
        }
    }

    if drawing_data.is_empty() {
        0
    } else {
        EscherShapeFactory::count_shapes_in_drawing(&drawing_data)
    }
}
