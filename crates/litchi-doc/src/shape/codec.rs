//! DOC-specific OfficeArt and textbox-story decoding.

use super::model::Shape;
use crate::Document;
use crate::parts::fib::FileInformationBlock;
use litchi_cfb::OleFile;
use litchi_odraw::{Record, RecordKind};

use std::{
    collections::HashMap,
    io::{Read, Seek},
};

/// FIB index (into `FileInformationBlock::get_table_pointer`) of the
/// `fcDggInfo`/`lcbDggInfo` pair holding the document's OfficeArtContent.
pub(crate) const FIB_INDEX_DGG_INFO: usize = 50;

/// Size of an OfficeArt record header in bytes.
const RECORD_HEADER_LEN: usize = 8;

fn project(data: &[u8]) -> std::io::Result<Vec<Shape>> {
    litchi_odraw::shape::parse(data)
        .map_err(invalid_data)?
        .iter()
        .map(Shape::from_office_art)
        .collect()
}

fn count_tree(shapes: &[Shape]) -> std::io::Result<usize> {
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

fn drawing_shapes<R: Read + Seek>(ole: &mut OleFile<R>) -> std::io::Result<Vec<Shape>> {
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

fn apply_text(shapes: &mut [Shape], texts: &mut HashMap<u32, String>) -> std::io::Result<()> {
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

fn has_text_links(shapes: &[Shape]) -> bool {
    let mut pending = shapes.iter().rev().collect::<Vec<_>>();
    while let Some(shape) = pending.pop() {
        if shape.text_link {
            return true;
        }
        pending.extend(shape.children.iter().rev());
    }
    false
}

/// Extract shapes from the document's drawing group (`fcDggInfo`).
///
/// The OfficeArtContent in the table stream holds one OfficeArtWordDrawing
/// per story; each consists of a `dgglbl` byte followed by an
/// OfficeArtDgContainer with the story's floating shapes.
pub(crate) fn extract_dgg_shapes(
    fib: &FileInformationBlock,
    table_stream: &[u8],
) -> std::io::Result<Vec<Shape>> {
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

/// Extract all floating shapes from a Word document's OfficeArt content.
pub fn extract_shapes<R: Read + Seek>(ole: &mut OleFile<R>) -> std::io::Result<Vec<Shape>> {
    extract_drawing_shapes(ole)
}

/// Extract all floating shapes from a Word document's drawing group.
pub fn extract_drawing_shapes<R: Read + Seek>(ole: &mut OleFile<R>) -> std::io::Result<Vec<Shape>> {
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
pub fn count_shapes<R: Read + Seek>(ole: &mut OleFile<R>) -> std::io::Result<usize> {
    count_tree(&drawing_shapes(ole)?)
}

fn invalid_data(error: impl ToString) -> std::io::Error {
    std::io::Error::new(std::io::ErrorKind::InvalidData, error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shape::Kind;

    fn create_shape(
        shape_type: Kind,
        shape_id: u32,
        text: Option<String>,
        is_group: bool,
        children: Vec<Shape>,
    ) -> Shape {
        let text_link = text.is_some();
        Shape {
            shape_type,
            shape_id,
            text,
            is_group,
            children,
            fill_color: None,
            line_color: None,
            native_shape_type: None,
            group_bounds: None,
            unknown_records: Vec::new(),
            text_link,
        }
    }

    #[test]
    fn shape_creation_preserves_contextual_fields() {
        let shape = create_shape(
            Kind::Rectangle,
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
    fn nested_groups_keep_source_order() {
        let child = create_shape(Kind::Rectangle, 101, None, false, vec![]);
        let parent = create_shape(Kind::Group, 100, None, true, vec![child]);

        assert!(parent.is_group);
        assert_eq!(parent.children.len(), 1);
        assert_eq!(parent.children[0].shape_id, 101);
    }

    #[test]
    fn shape_clone_is_owned() {
        let shape = create_shape(
            Kind::Ellipse,
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
    fn shape_debug_is_contextual() {
        let shape = create_shape(
            Kind::Rectangle,
            1,
            Some("Debug test".to_string()),
            false,
            vec![],
        );
        let debug = format!("{shape:?}");

        assert!(debug.contains("Shape"));
        assert!(!debug.contains("DocShape"));
    }

    #[test]
    fn deeply_nested_groups_keep_all_children() {
        let level4 = create_shape(Kind::Rectangle, 4, None, false, vec![]);
        let level3 = create_shape(Kind::Group, 3, None, true, vec![level4]);
        let level2 = create_shape(Kind::Group, 2, None, true, vec![level3]);
        let level1 = create_shape(Kind::Group, 1, None, true, vec![level2]);

        assert!(level1.is_group);
        assert!(level1.children[0].is_group);
        assert!(level1.children[0].children[0].is_group);
        assert!(!level1.children[0].children[0].children[0].is_group);
        assert_eq!(level1.children[0].children[0].children[0].shape_id, 4);
    }

    #[test]
    fn multiple_children_keep_their_order() {
        let children: Vec<Shape> = (1..=5)
            .map(|id| {
                create_shape(
                    Kind::Rectangle,
                    id,
                    Some(format!("Child {id}")),
                    false,
                    vec![],
                )
            })
            .collect();
        let parent = create_shape(Kind::Group, 0, None, true, children);

        assert_eq!(parent.children.len(), 5);
        for (index, child) in parent.children.iter().enumerate() {
            assert_eq!(child.shape_id, (index + 1) as u32);
            assert_eq!(child.text, Some(format!("Child {}", index + 1)));
        }
    }
}
