//! Shape extraction for Word documents.
//!
//! Word documents store drawing objects (shapes, images, text boxes) as
//! OfficeArtContent in the FIB-selected table stream. This module projects
//! those records together with Word's separate textbox stories.

use crate::doc::Document;
use crate::doc::parts::fib::FileInformationBlock;
use litchi_cfb::OleFile;
use litchi_odraw::{
    Record, RecordKind,
    shape::{self, Kind, Shape},
};
use std::{
    collections::HashMap,
    io::{Read, Seek},
};

/// FIB index (into `FileInformationBlock::get_table_pointer`) of the
/// `fcDggInfo`/`lcbDggInfo` pair holding the document's OfficeArtContent.
pub(crate) const FIB_INDEX_DGG_INFO: usize = 50;

/// Size of an OfficeArt record header in bytes.
const RECORD_HEADER_LEN: usize = 8;

/// Shape information extracted from a Word document.
#[derive(Debug, Clone)]
pub struct DocShape {
    /// Shape type (rectangle, ellipse, line, etc.)
    pub shape_type: Kind,
    /// Shape ID
    pub shape_id: u32,
    /// Text content extracted from the shape (if any)
    pub text: Option<String>,
    /// Whether this is a group shape
    pub is_group: bool,
    /// Child shapes (for group shapes)
    pub children: Vec<DocShape>,
    /// Fill color as (R, G, B), when the shape sets an explicit `fillColor`.
    pub fill_color: Option<(u8, u8, u8)>,
    /// Line color as (R, G, B), when the shape sets an explicit `lineColor`.
    pub line_color: Option<(u8, u8, u8)>,
    /// The raw MSOSPT preset-geometry value ([MS-ODRAW] 2.4.24) from the
    /// shape's OfficeArtFSP record, when the shape has preset geometry.
    pub native_shape_type: Option<u16>,
    text_link: bool,
}

impl DocShape {
    /// Project a host-neutral OfficeArt shape into Word's drawing facade.
    fn from_odraw(shape: &Shape<'_>) -> std::io::Result<Self> {
        // In Word, OfficeArtClientTextbox contains only a TXID into the Word
        // textbox story. Text itself is resolved by `Document::text_boxes` and
        // cannot be decoded from OfficeArt bytes in isolation.
        let text_link = if let Some(textbox) = shape.textbox() {
            let _: &[u8; 4] = textbox
                .data()
                .try_into()
                .map_err(|_| invalid_data("Word OfficeArtClientTextbox payload is not one TXID"))?;
            true
        } else {
            false
        };

        let children = shape
            .children()
            .iter()
            .map(Self::from_odraw)
            .collect::<std::io::Result<_>>()?;

        Ok(Self {
            shape_type: shape.kind(),
            shape_id: shape.id(),
            text: None,
            is_group: matches!(shape.kind(), Kind::Group | Kind::Table),
            children,
            fill_color: shape.props().get_fill_color(),
            line_color: shape.props().get_line_color(),
            native_shape_type: Some(shape.native_kind().raw()),
            text_link,
        })
    }
}

fn invalid_data(error: impl ToString) -> std::io::Error {
    std::io::Error::new(std::io::ErrorKind::InvalidData, error.to_string())
}

fn project(data: &[u8]) -> std::io::Result<Vec<DocShape>> {
    shape::parse(data)
        .map_err(invalid_data)?
        .iter()
        .map(DocShape::from_odraw)
        .collect()
}

fn count_doc_tree(shapes: &[DocShape]) -> std::io::Result<usize> {
    let mut count = 0usize;
    let mut pending = shapes.iter().rev().collect::<Vec<_>>();
    while let Some(shape) = pending.pop() {
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid_data("OfficeArt shape count overflow"))?;
        pending.extend(shape.children.iter().rev());
    }
    Ok(count)
}

fn drawing_shapes<R: Read + Seek>(ole: &mut OleFile<R>) -> std::io::Result<Vec<DocShape>> {
    let word_document = ole
        .open_stream(&["WordDocument"])
        .map_err(|error| invalid_data(format!("cannot open WordDocument stream: {error}")))?;
    let fib = FileInformationBlock::parse(&word_document).map_err(invalid_data)?;
    let table_stream_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_stream = ole.open_stream(&[table_stream_name]).map_err(|error| {
        invalid_data(format!("cannot open {table_stream_name} stream: {error}"))
    })?;
    extract_dgg_shapes(&fib, &table_stream)
}

fn text_by_shape<R: Read + Seek>(ole: &mut OleFile<R>) -> std::io::Result<HashMap<u32, String>> {
    let document = Document::from_ole(ole).map_err(invalid_data)?;
    let mut texts = HashMap::new();
    for text_box in document
        .text_boxes()
        .into_iter()
        .chain(document.header_text_boxes())
    {
        if texts.insert(text_box.shape_id, text_box.text).is_some() {
            return Err(invalid_data(format!(
                "duplicate Word textbox shape identifier {}",
                text_box.shape_id
            )));
        }
    }
    Ok(texts)
}

fn apply_text(shapes: &mut [DocShape], texts: &mut HashMap<u32, String>) -> std::io::Result<()> {
    for shape in shapes {
        if shape.text_link {
            shape.text = Some(texts.remove(&shape.shape_id).ok_or_else(|| {
                invalid_data(format!(
                    "Word shape {} links to a missing textbox story",
                    shape.shape_id
                ))
            })?);
        }
        apply_text(&mut shape.children, texts)?;
    }
    Ok(())
}

fn has_text_links(shapes: &[DocShape]) -> bool {
    let mut pending = shapes.iter().rev().collect::<Vec<_>>();
    while let Some(shape) = pending.pop() {
        if shape.text_link {
            return true;
        }
        pending.extend(shape.children.iter().rev());
    }
    false
}

/// Extract shapes from the document's drawing group (fcDggInfo).
///
/// The OfficeArtContent in the table stream holds one OfficeArtWordDrawing
/// per story; each consists of a `dgglbl` byte followed by an
/// OfficeArtDgContainer with the story's floating shapes.
pub(crate) fn extract_dgg_shapes(
    fib: &FileInformationBlock,
    table_stream: &[u8],
) -> std::io::Result<Vec<DocShape>> {
    let Some((offset, length)) = fib.get_table_pointer(FIB_INDEX_DGG_INFO) else {
        return Ok(Vec::new());
    };
    if length == 0 {
        return Ok(Vec::new());
    }
    let start = usize::try_from(offset).map_err(invalid_data)?;
    let length = usize::try_from(length).map_err(invalid_data)?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| invalid_data("Word drawing-group extent overflow"))?;
    let dgg_info = table_stream
        .get(start..end)
        .ok_or_else(|| invalid_data("Word drawing group extends past the table stream"))?;

    let mut shapes = Vec::new();
    // The OfficeArtContent starts with the OfficeArtDggContainer; each
    // OfficeArtWordDrawing after it is a dgglbl byte + OfficeArtDgContainer.
    let (first, first_size) = Record::parse(dgg_info, 0).map_err(invalid_data)?;
    if first.kind() != RecordKind::DggContainer {
        return Err(invalid_data(
            "Word OfficeArtContent does not start with a DggContainer",
        ));
    }
    let mut offset = first_size;
    while offset < dgg_info.len() {
        let label = *dgg_info
            .get(offset)
            .ok_or_else(|| invalid_data("truncated OfficeArtWordDrawing label"))?;
        if label > 1 {
            return Err(invalid_data(format!(
                "OfficeArtWordDrawing has invalid story label {label}"
            )));
        }
        let record_offset = offset
            .checked_add(1)
            .ok_or_else(|| invalid_data("Word drawing offset overflow"))?;
        if dgg_info.len().saturating_sub(record_offset) < RECORD_HEADER_LEN {
            return Err(invalid_data("truncated OfficeArtWordDrawing"));
        }
        let (record, record_size) = Record::parse(dgg_info, record_offset).map_err(invalid_data)?;
        if record.kind() != RecordKind::DgContainer {
            return Err(invalid_data(
                "OfficeArtWordDrawing does not contain a DgContainer",
            ));
        }
        let record_end = record_offset
            .checked_add(record_size)
            .ok_or_else(|| invalid_data("Word drawing extent overflow"))?;
        shapes.extend(project(&dgg_info[record_offset..record_end])?);
        offset = record_end;
    }
    Ok(shapes)
}

/// Extracts all floating shapes from the document's OfficeArt content.
///
/// # Arguments
///
/// * `ole` - The OLE file containing the document
///
/// # Returns
///
/// A vector of shapes found in the document, or an empty vector if no shapes exist.
pub fn extract_shapes<R: Read + Seek>(ole: &mut OleFile<R>) -> std::io::Result<Vec<DocShape>> {
    extract_drawing_shapes(ole)
}

/// Extract all floating shapes from a Word document's drawing group
/// (the `fcDggInfo` OfficeArtContent in the table stream).
///
/// # Arguments
///
/// * `ole` - The OLE file containing the document
///
/// # Returns
///
/// A vector of floating shapes found in the document, or an empty vector if
/// the document has no drawing group.
pub fn extract_drawing_shapes<R: Read + Seek>(
    ole: &mut OleFile<R>,
) -> std::io::Result<Vec<DocShape>> {
    let mut shapes = drawing_shapes(ole)?;
    if !has_text_links(&shapes) {
        return Ok(shapes);
    }
    let mut texts = text_by_shape(ole)?;
    apply_text(&mut shapes, &mut texts)?;
    if !texts.is_empty() {
        return Err(invalid_data(
            "Word textbox story references a shape absent from OfficeArtContent",
        ));
    }
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
    let document = Document::from_ole(ole).map_err(invalid_data)?;
    Ok(document
        .text_boxes()
        .into_iter()
        .chain(document.header_text_boxes())
        .map(|text_box| text_box.text)
        .filter(|text| !text.is_empty())
        .collect::<Vec<_>>()
        .join("\n"))
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
    count_doc_tree(&drawing_shapes(ole)?)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn create_test_doc_shape(
        shape_type: Kind,
        shape_id: u32,
        text: Option<String>,
        is_group: bool,
        children: Vec<DocShape>,
    ) -> DocShape {
        let text_link = text.is_some();
        DocShape {
            shape_type,
            shape_id,
            text,
            is_group,
            children,
            fill_color: None,
            line_color: None,
            native_shape_type: None,
            text_link,
        }
    }

    #[test]
    fn test_doc_shape_creation() {
        let doc_shape = create_test_doc_shape(
            Kind::Rectangle,
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
        let child = create_test_doc_shape(Kind::Rectangle, 101, None, false, vec![]);

        let parent = create_test_doc_shape(Kind::Group, 100, None, true, vec![child]);

        assert!(parent.is_group);
        assert_eq!(parent.children.len(), 1);
        assert_eq!(parent.children[0].shape_id, 101);
    }

    #[test]
    fn test_doc_shape_clone() {
        let doc_shape = create_test_doc_shape(
            Kind::Ellipse,
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
            Kind::Rectangle,
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
        let inner_child =
            create_test_doc_shape(Kind::Rectangle, 3, Some("Inner".to_string()), false, vec![]);

        let middle = create_test_doc_shape(Kind::Group, 2, None, true, vec![inner_child]);

        let outer = create_test_doc_shape(Kind::Group, 1, None, true, vec![middle]);

        assert!(outer.is_group);
        assert_eq!(outer.children.len(), 1);
        assert!(outer.children[0].is_group);
        assert_eq!(outer.children[0].children.len(), 1);
        assert_eq!(outer.children[0].children[0].shape_id, 3);
    }

    #[test]
    fn test_doc_shape_variants() {
        let shape_types = vec![
            Kind::Rectangle,
            Kind::Ellipse,
            Kind::Line,
            Kind::Group,
            Kind::Picture,
            Kind::TextBox,
            Kind::Polygon,
            Kind::AutoShape,
            Kind::Connector,
            Kind::Unknown,
        ];

        for (i, shape_type) in shape_types.iter().enumerate() {
            let doc_shape = create_test_doc_shape(*shape_type, i as u32, None, false, vec![]);
            assert_eq!(doc_shape.shape_id, i as u32);
            assert_eq!(doc_shape.shape_type, *shape_type);
        }
    }

    #[test]
    fn test_doc_shape_empty_text() {
        let doc_shape = create_test_doc_shape(Kind::TextBox, 1, None, false, vec![]);
        assert!(doc_shape.text.is_none());
    }

    #[test]
    fn test_doc_shape_unicode_text() {
        let doc_shape = create_test_doc_shape(
            Kind::TextBox,
            1,
            Some("Unicode: 你好世界 🎉".to_string()),
            false,
            vec![],
        );
        assert_eq!(doc_shape.text.unwrap(), "Unicode: 你好世界 🎉");
    }

    #[test]
    fn test_deeply_nested_groups() {
        let level4 = create_test_doc_shape(Kind::Rectangle, 4, None, false, vec![]);
        let level3 = create_test_doc_shape(Kind::Group, 3, None, true, vec![level4]);
        let level2 = create_test_doc_shape(Kind::Group, 2, None, true, vec![level3]);
        let level1 = create_test_doc_shape(Kind::Group, 1, None, true, vec![level2]);

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
                    Kind::Rectangle,
                    i,
                    Some(format!("Child {}", i)),
                    false,
                    vec![],
                )
            })
            .collect();

        let parent = create_test_doc_shape(Kind::Group, 0, None, true, children);

        assert_eq!(parent.children.len(), 5);
        for (i, child) in parent.children.iter().enumerate() {
            assert_eq!(child.shape_id, (i + 1) as u32);
            assert_eq!(child.text, Some(format!("Child {}", i + 1)));
        }
    }
}
