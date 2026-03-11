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

#[cfg(test)]
mod tests {
    use super::*;

    fn create_test_xls_shape(
        shape_type: crate::ole::escher::EscherShapeType,
        shape_id: u32,
        text: Option<String>,
        is_group: bool,
        children: Vec<XlsShape>,
    ) -> XlsShape {
        XlsShape {
            shape_type,
            shape_id,
            text,
            is_group,
            children,
        }
    }

    #[test]
    fn test_xls_shape_creation() {
        let shape = create_test_xls_shape(
            crate::ole::escher::EscherShapeType::Rectangle,
            100,
            Some("Shape text".to_string()),
            false,
            vec![],
        );

        assert_eq!(shape.shape_id, 100);
        assert_eq!(shape.text, Some("Shape text".to_string()));
        assert!(!shape.is_group);
        assert!(shape.children.is_empty());
    }

    #[test]
    fn test_xls_shape_group() {
        let child = create_test_xls_shape(
            crate::ole::escher::EscherShapeType::Rectangle,
            101,
            None,
            false,
            vec![],
        );

        let parent = create_test_xls_shape(
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
    fn test_xls_shape_clone() {
        let shape = create_test_xls_shape(
            crate::ole::escher::EscherShapeType::Ellipse,
            50,
            Some("Clonable".to_string()),
            false,
            vec![],
        );
        let cloned = shape.clone();

        assert_eq!(cloned.shape_id, shape.shape_id);
        assert_eq!(cloned.text, shape.text);
        assert_eq!(cloned.is_group, shape.is_group);
    }

    #[test]
    fn test_xls_shape_debug() {
        let shape = create_test_xls_shape(
            crate::ole::escher::EscherShapeType::Rectangle,
            1,
            Some("Debug test".to_string()),
            false,
            vec![],
        );
        let debug_str = format!("{:?}", shape);

        assert!(debug_str.contains("XlsShape"));
    }

    #[test]
    fn test_nested_groups() {
        let inner_child = create_test_xls_shape(
            crate::ole::escher::EscherShapeType::Rectangle,
            3,
            Some("Inner".to_string()),
            false,
            vec![],
        );

        let middle = create_test_xls_shape(
            crate::ole::escher::EscherShapeType::Group,
            2,
            None,
            true,
            vec![inner_child],
        );

        let outer = create_test_xls_shape(
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
    fn test_xls_shape_variants() {
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
            let shape = create_test_xls_shape(*shape_type, i as u32, None, false, vec![]);
            assert_eq!(shape.shape_id, i as u32);
            assert_eq!(shape.shape_type, *shape_type);
        }
    }

    #[test]
    fn test_xls_shape_empty_text() {
        let shape = create_test_xls_shape(
            crate::ole::escher::EscherShapeType::TextBox,
            1,
            None,
            false,
            vec![],
        );
        assert!(shape.text.is_none());
    }

    #[test]
    fn test_xls_shape_unicode_text() {
        let shape = create_test_xls_shape(
            crate::ole::escher::EscherShapeType::TextBox,
            1,
            Some("Unicode: 你好世界 🎉".to_string()),
            false,
            vec![],
        );
        assert_eq!(shape.text.unwrap(), "Unicode: 你好世界 🎉");
    }

    #[test]
    fn test_extract_shapes_empty_data() {
        let result = extract_shapes_from_workbook(b"");
        assert!(result.is_ok());
        assert!(result.unwrap().is_empty());
    }

    #[test]
    fn test_extract_shape_text_empty_data() {
        let result = extract_shape_text_from_workbook(b"");
        assert!(result.is_ok());
        assert_eq!(result.unwrap(), "");
    }

    #[test]
    fn test_count_shapes_empty_data() {
        let count = count_shapes_in_workbook(b"");
        assert_eq!(count, 0);
    }

    #[test]
    fn test_count_shapes_invalid_data() {
        // Random data that isn't valid BIFF
        let data = vec![0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07];
        let count = count_shapes_in_workbook(&data);
        assert_eq!(count, 0);
    }

    #[test]
    fn test_multiple_children() {
        let children: Vec<XlsShape> = (1..=5)
            .map(|i| {
                create_test_xls_shape(
                    crate::ole::escher::EscherShapeType::Rectangle,
                    i,
                    Some(format!("Child {}", i)),
                    false,
                    vec![],
                )
            })
            .collect();

        let parent = create_test_xls_shape(
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

    #[test]
    fn test_deeply_nested_groups() {
        let level4 = create_test_xls_shape(
            crate::ole::escher::EscherShapeType::Rectangle,
            4,
            None,
            false,
            vec![],
        );
        let level3 = create_test_xls_shape(
            crate::ole::escher::EscherShapeType::Group,
            3,
            None,
            true,
            vec![level4],
        );
        let level2 = create_test_xls_shape(
            crate::ole::escher::EscherShapeType::Group,
            2,
            None,
            true,
            vec![level3],
        );
        let level1 = create_test_xls_shape(
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
}
