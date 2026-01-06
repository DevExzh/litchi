//! Shape extraction for Word documents.
//!
//! Word documents store drawing objects (shapes, images, text boxes) in the OfficeArt
//! format (Escher) within the "Data" stream. This module provides access to these shapes.

use crate::ole::escher::{EscherShape, EscherShapeFactory, extract_text_from_escher};
use crate::ole::file::OleFile;
use std::io::{Read, Seek};

/// Shape information extracted from a Word document.
#[derive(Debug, Clone)]
pub struct DocShape {
    /// Shape type (rectangle, ellipse, line, etc.)
    pub shape_type: crate::ole::escher::EscherShapeType,
    /// Shape ID
    pub shape_id: u32,
    /// Text content extracted from the shape (if any)
    pub text: Option<String>,
    /// Whether this is a group shape
    pub is_group: bool,
    /// Child shapes (for group shapes)
    pub children: Vec<DocShape>,
}

impl DocShape {
    /// Create a DocShape from an EscherShape.
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

/// Extract all shapes from a Word document's Data stream.
///
/// # Arguments
///
/// * `ole` - The OLE file containing the document
///
/// # Returns
///
/// A vector of shapes found in the document, or an empty vector if no shapes exist.
pub fn extract_shapes<R: Read + Seek>(ole: &mut OleFile<R>) -> std::io::Result<Vec<DocShape>> {
    // Try to open the Data stream (where drawings are stored)
    let data_stream = match ole.open_stream(&["Data"]) {
        Ok(stream) => stream,
        Err(_) => return Ok(Vec::new()), // No Data stream = no shapes
    };

    // Parse Escher data and extract shapes
    let escher_shapes = EscherShapeFactory::extract_shapes_from_drawing(&data_stream)?;

    // Convert to DocShape
    let shapes = escher_shapes.iter().map(DocShape::from_escher).collect();

    Ok(shapes)
}

/// Extract text from all shapes in a Word document.
///
/// # Arguments
///
/// * `ole` - The OLE file containing the document
///
/// # Returns
///
/// A string containing all text extracted from shapes, or an empty string if no text found.
pub fn extract_shape_text<R: Read + Seek>(ole: &mut OleFile<R>) -> std::io::Result<String> {
    let data_stream = match ole.open_stream(&["Data"]) {
        Ok(stream) => stream,
        Err(_) => return Ok(String::new()),
    };

    extract_text_from_escher(&data_stream)
}

/// Count the number of shapes in a Word document.
///
/// # Arguments
///
/// * `ole` - The OLE file containing the document
///
/// # Returns
///
/// The number of shapes found, or 0 if no shapes exist.
pub fn count_shapes<R: Read + Seek>(ole: &mut OleFile<R>) -> std::io::Result<usize> {
    let data_stream = match ole.open_stream(&["Data"]) {
        Ok(stream) => stream,
        Err(_) => return Ok(0),
    };

    Ok(EscherShapeFactory::count_shapes_in_drawing(&data_stream))
}
