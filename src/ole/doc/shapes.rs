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

#[cfg(test)]
mod tests {
    use super::*;

    fn create_test_doc_shape(
        shape_type: crate::ole::escher::EscherShapeType,
        shape_id: u32,
        text: Option<String>,
        is_group: bool,
        children: Vec<DocShape>,
    ) -> DocShape {
        DocShape {
            shape_type,
            shape_id,
            text,
            is_group,
            children,
        }
    }

    #[test]
    fn test_doc_shape_creation() {
        let doc_shape = create_test_doc_shape(
            crate::ole::escher::EscherShapeType::Rectangle,
            100,
            Some("Shape text".to_string()),
            false,
            vec![],
        );

        assert_eq!(doc_shape.shape_id, 100);
        assert_eq!(doc_shape.text, Some("Shape text".to_string()));
        assert!(!doc_shape.is_group);
        assert!(doc_shape.children.is_empty());
    }

    #[test]
    fn test_doc_shape_group() {
        let child = create_test_doc_shape(
            crate::ole::escher::EscherShapeType::Rectangle,
            101,
            None,
            false,
            vec![],
        );

        let parent = create_test_doc_shape(
            crate::ole::escher::EscherShapeType::Group,
            100,
            None,
            true,
            vec![child],
        );

        assert!(parent.is_group);
        assert_eq!(parent.children.len(), 1);
        assert_eq!(parent.children[0].shape_id, 101);
    }

    #[test]
    fn test_doc_shape_clone() {
        let doc_shape = create_test_doc_shape(
            crate::ole::escher::EscherShapeType::Ellipse,
            50,
            Some("Clonable".to_string()),
            false,
            vec![],
        );
        let cloned = doc_shape.clone();

        assert_eq!(cloned.shape_id, doc_shape.shape_id);
        assert_eq!(cloned.text, doc_shape.text);
        assert_eq!(cloned.is_group, doc_shape.is_group);
    }

    #[test]
    fn test_doc_shape_debug() {
        let doc_shape = create_test_doc_shape(
            crate::ole::escher::EscherShapeType::Rectangle,
            1,
            Some("Debug test".to_string()),
            false,
            vec![],
        );
        let debug_str = format!("{:?}", doc_shape);

        assert!(debug_str.contains("DocShape"));
    }

    #[test]
    fn test_nested_groups() {
        let inner_child = create_test_doc_shape(
            crate::ole::escher::EscherShapeType::Rectangle,
            3,
            Some("Inner".to_string()),
            false,
            vec![],
        );

        let middle = create_test_doc_shape(
            crate::ole::escher::EscherShapeType::Group,
            2,
            None,
            true,
            vec![inner_child],
        );

        let outer = create_test_doc_shape(
            crate::ole::escher::EscherShapeType::Group,
            1,
            None,
            true,
            vec![middle],
        );

        assert!(outer.is_group);
        assert_eq!(outer.children.len(), 1);
        assert!(outer.children[0].is_group);
        assert_eq!(outer.children[0].children.len(), 1);
        assert_eq!(outer.children[0].children[0].shape_id, 3);
    }

    #[test]
    fn test_doc_shape_variants() {
        use crate::ole::escher::EscherShapeType;

        let shape_types = vec![
            EscherShapeType::Rectangle,
            EscherShapeType::Ellipse,
            EscherShapeType::Line,
            EscherShapeType::Group,
            EscherShapeType::Picture,
            EscherShapeType::TextBox,
            EscherShapeType::Polygon,
            EscherShapeType::AutoShape,
            EscherShapeType::Connector,
            EscherShapeType::Unknown,
        ];

        for (i, shape_type) in shape_types.iter().enumerate() {
            let doc_shape = create_test_doc_shape(*shape_type, i as u32, None, false, vec![]);
            assert_eq!(doc_shape.shape_id, i as u32);
            assert_eq!(doc_shape.shape_type, *shape_type);
        }
    }

    #[test]
    fn test_doc_shape_empty_text() {
        let doc_shape = create_test_doc_shape(
            crate::ole::escher::EscherShapeType::TextBox,
            1,
            None,
            false,
            vec![],
        );
        assert!(doc_shape.text.is_none());
    }

    #[test]
    fn test_doc_shape_unicode_text() {
        let doc_shape = create_test_doc_shape(
            crate::ole::escher::EscherShapeType::TextBox,
            1,
            Some("Unicode: 你好世界 🎉".to_string()),
            false,
            vec![],
        );
        assert_eq!(doc_shape.text.unwrap(), "Unicode: 你好世界 🎉");
    }

    #[test]
    fn test_deeply_nested_groups() {
        let level4 = create_test_doc_shape(
            crate::ole::escher::EscherShapeType::Rectangle,
            4,
            None,
            false,
            vec![],
        );
        let level3 = create_test_doc_shape(
            crate::ole::escher::EscherShapeType::Group,
            3,
            None,
            true,
            vec![level4],
        );
        let level2 = create_test_doc_shape(
            crate::ole::escher::EscherShapeType::Group,
            2,
            None,
            true,
            vec![level3],
        );
        let level1 = create_test_doc_shape(
            crate::ole::escher::EscherShapeType::Group,
            1,
            None,
            true,
            vec![level2],
        );

        assert!(level1.is_group);
        assert!(level1.children[0].is_group);
        assert!(level1.children[0].children[0].is_group);
        assert!(!level1.children[0].children[0].children[0].is_group);
        assert_eq!(level1.children[0].children[0].children[0].shape_id, 4);
    }

    #[test]
    fn test_multiple_children() {
        let children: Vec<DocShape> = (1..=5)
            .map(|i| {
                create_test_doc_shape(
                    crate::ole::escher::EscherShapeType::Rectangle,
                    i,
                    Some(format!("Child {}", i)),
                    false,
                    vec![],
                )
            })
            .collect();

        let parent = create_test_doc_shape(
            crate::ole::escher::EscherShapeType::Group,
            0,
            None,
            true,
            children,
        );

        assert_eq!(parent.children.len(), 5);
        for (i, child) in parent.children.iter().enumerate() {
            assert_eq!(child.shape_id, (i + 1) as u32);
            assert_eq!(child.text, Some(format!("Child {}", i + 1)));
        }
    }
}
