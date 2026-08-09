//! DOC-specific `OfficeArt` and textbox-story decoding.

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
/// `fcDggInfo`/`lcbDggInfo` pair holding the document's `OfficeArtContent`.
pub(crate) const FIB_INDEX_DGG_INFO: usize = 50;

/// Size of an `OfficeArt` record header in bytes.
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
/// The `OfficeArtContent` in the table stream holds one `OfficeArtWordDrawing`
/// per story; each consists of a `dgglbl` byte followed by an
/// `OfficeArtDgContainer` with the story's floating shapes.
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

/// Extract all floating shapes from a Word document's `OfficeArt` content.
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
    use crate::shape::{Anchor, Bounds, Flags, Kind, PropertyTable, ShapeId};
    use litchi_odraw::write::{self, Atom, Container as OutContainer, ShapeBuilder};

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
            anchor: None,
            client_anchor: None,
            flags: Flags::empty(),
            unknown_records: Vec::new(),
            unknown_properties: Vec::new(),
            office_art: Vec::new().into_boxed_slice(),
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

    fn officeart_drawing() -> Vec<u8> {
        fn shape(kind: litchi_odraw::shape::Native, id: u32, child: bool) -> Vec<u8> {
            let mut body = Vec::new();
            let mut flags = Flags::HAVE_ANCHOR | Flags::HAVE_SPT;
            if child {
                flags |= Flags::CHILD;
            }
            ShapeBuilder::new(kind, id)
                .with_flags(flags)
                .write(&mut body)
                .expect("write shape atom");
            if child {
                write::child_anchor(&mut body, 10, 20, 110, 70).expect("write child anchor");
            } else {
                let mut payload = Vec::new();
                for coordinate in [10_i32, 20, 110, 70] {
                    payload.extend_from_slice(&coordinate.to_le_bytes());
                }
                write::atom(&mut body, 0, Atom::ClientAnchor, &payload).expect("write host anchor");
            }
            let mut record = Vec::new();
            write::container(&mut record, 0, OutContainer::Sp, &body)
                .expect("write shape container");
            record
        }

        let mut group_header_body = Vec::new();
        write::spgr(&mut group_header_body, 0, 0, 1000, 500).expect("write group bounds");
        ShapeBuilder::new(litchi_odraw::shape::Native::FREEFORM, 3)
            .with_flags(Flags::GROUP | Flags::HAVE_ANCHOR)
            .write(&mut group_header_body)
            .expect("write group shape");
        let mut group_anchor = Vec::new();
        for coordinate in [100_i32, 200, 500, 400] {
            group_anchor.extend_from_slice(&coordinate.to_le_bytes());
        }
        write::atom(&mut group_header_body, 0, Atom::ClientAnchor, &group_anchor)
            .expect("write group host anchor");
        let mut group_header = Vec::new();
        write::container(&mut group_header, 0, OutContainer::Sp, &group_header_body)
            .expect("write group header");

        let mut group_body = group_header;
        group_body.extend_from_slice(&shape(litchi_odraw::shape::Native::ELLIPSE, 4, true));
        let mut group = Vec::new();
        write::container(&mut group, 0, OutContainer::Spgr, &group_body).expect("write group");

        let mut patriarch_body = Vec::new();
        write::spgr(&mut patriarch_body, 0, 0, 1000, 500).expect("write patriarch bounds");
        ShapeBuilder::new(litchi_odraw::shape::Native::FREEFORM, 1)
            .with_flags(Flags::GROUP | Flags::PATRIARCH)
            .write(&mut patriarch_body)
            .expect("write patriarch shape");
        let mut patriarch = Vec::new();
        write::container(&mut patriarch, 0, OutContainer::Sp, &patriarch_body)
            .expect("write patriarch");

        let mut root_body = patriarch;
        root_body.extend_from_slice(&shape(litchi_odraw::shape::Native::RECTANGLE, 2, false));
        root_body.extend_from_slice(&group);
        let mut root = Vec::new();
        write::container(&mut root, 0, OutContainer::Spgr, &root_body).expect("write root group");

        let mut drawing_body = Vec::new();
        write::dg(&mut drawing_body, 5, 5).expect("write drawing atom");
        drawing_body.extend_from_slice(&root);
        let mut bytes = Vec::new();
        write::container(&mut bytes, 0, OutContainer::Dg, &drawing_body).expect("write drawing");
        bytes
    }

    #[test]
    fn projection_exposes_group_identity_and_both_anchor_spaces() {
        let bytes = officeart_drawing();
        let office_shapes = litchi_odraw::shape::parse(&bytes).expect("parse drawing");
        let shapes = office_shapes
            .iter()
            .map(Shape::from_office_art)
            .collect::<std::io::Result<Vec<_>>>()
            .expect("project drawing");

        let rectangle = &shapes[0];
        assert_eq!(rectangle.identity().raw(), 2);
        assert!(rectangle.shape_flags().contains(Flags::HAVE_ANCHOR));
        assert_eq!(rectangle.anchor(), None);
        assert_eq!(rectangle.client_anchor().unwrap().payload().len(), 16);

        let group = shapes[1].group().expect("typed group projection");
        assert_eq!(group.identity(), ShapeId::from_raw(3));
        assert_eq!(group.shape_id(), 3);
        assert_eq!(group.bounds(), Some(&Bounds::new(0, 0, 1000, 500)));
        assert_eq!(group.children().len(), 1);

        let child = &group.children()[0];
        assert_eq!(child.identity(), ShapeId::from_raw(4));
        assert_eq!(child.anchor(), Some(&Anchor::new(10, 20, 110, 70)));
        assert!(child.client_anchor().is_none());
        assert_eq!(group.find(ShapeId::from_raw(4)).map(Shape::id), Some(4));

        let (drawing_record, _) = Record::parse(&bytes, 0).expect("parse drawing record");
        let drawing = litchi_odraw::Container::try_new(drawing_record).expect("drawing container");
        let root_record = drawing
            .find(RecordKind::SpgrContainer)
            .expect("scan drawing records")
            .expect("root group record");
        let root = litchi_odraw::Container::try_new(root_record).expect("root group container");
        let group_record = root
            .find(RecordKind::SpgrContainer)
            .expect("scan nested group records")
            .expect("nested group record");
        let root_start = group_record
            .data_offset(&bytes)
            .expect("root group belongs to source bytes")
            .checked_sub(RECORD_HEADER_LEN)
            .expect("root group has a header");
        let root_end = root_start
            .checked_add(RECORD_HEADER_LEN)
            .and_then(|offset| {
                offset.checked_add(usize::try_from(group_record.len()).expect("record length fits"))
            })
            .expect("root group extent");
        assert_eq!(group.office_art_bytes(), &bytes[root_start..root_end]);
    }

    #[test]
    fn projection_retains_exact_standalone_container_bytes() {
        let mut body = Vec::new();
        ShapeBuilder::new(litchi_odraw::shape::Native::RECTANGLE, 9)
            .with_flags(Flags::HAVE_ANCHOR | Flags::HAVE_SPT)
            .write(&mut body)
            .expect("write shape atom");
        let extension = Atom::unknown(0xF123, 0).expect("unknown record kind");
        write::atom(&mut body, 0, extension, &[0xAA, 0xBB, 0xCC]).expect("write unknown record");
        write::atom(&mut body, 0, Atom::ClientAnchor, &[7, 0, 0, 0]).expect("write host anchor");

        let mut source = Vec::new();
        write::container(&mut source, 0, OutContainer::Sp, &body).expect("write shape container");
        let office_shapes = litchi_odraw::shape::parse(&source).expect("parse shape");
        let shape = Shape::from_office_art(&office_shapes[0]).expect("project shape");

        assert_eq!(shape.office_art_bytes(), source.as_slice());
        let (record, consumed) =
            Record::parse(shape.office_art_bytes(), 0).expect("reparse retained shape container");
        assert_eq!(record.kind(), RecordKind::SpContainer);
        assert_eq!(consumed, source.len());
        assert_eq!(shape.unknown_records[0].data(), &[0xAA, 0xBB, 0xCC]);
    }

    #[test]
    fn projection_retains_unknown_primary_secondary_and_tertiary_properties() {
        let mut body = Vec::new();
        ShapeBuilder::new(litchi_odraw::shape::Native::RECTANGLE, 9)
            .with_flags(Flags::HAVE_ANCHOR | Flags::HAVE_SPT)
            .write(&mut body)
            .expect("write shape atom");

        let primary = [0xFE, 0xBF, 0x03, 0x00, 0x00, 0x00, 0xAA, 0xBB, 0xCC];
        write::atom(&mut body, 1, Atom::Opt, &primary).expect("write primary property");

        let secondary = [0xFD, 0x3F, 0xF9, 0xFF, 0xFF, 0xFF];
        write::atom(&mut body, 1, Atom::SecondaryOpt, &secondary)
            .expect("write secondary property");

        let tertiary = [0xFC, 0xBF, 0x02, 0x00, 0x00, 0x00, 0x10, 0x20];
        write::atom(&mut body, 1, Atom::TertiaryOpt, &tertiary).expect("write tertiary property");
        write::atom(&mut body, 0, Atom::ClientAnchor, &[7, 0, 0, 0]).expect("write host anchor");

        let mut bytes = Vec::new();
        write::container(&mut bytes, 0, OutContainer::Sp, &body).expect("write shape container");
        let office_shapes = litchi_odraw::shape::parse(&bytes).expect("parse shape");
        let shape = Shape::from_office_art(&office_shapes[0]).expect("project shape");

        assert_eq!(shape.unknown_properties().len(), 3);
        assert_eq!(
            shape.unknown_properties()[0].table(),
            PropertyTable::Primary
        );
        assert_eq!(shape.unknown_properties()[0].raw_id(), 0x3FFE);
        assert_eq!(shape.unknown_properties()[0].bytes(), &primary[..]);
        assert_eq!(shape.unknown_properties()[0].data(), &[0xAA, 0xBB, 0xCC]);
        assert!(shape.unknown_properties()[0].is_complex());
        assert_eq!(
            shape.unknown_properties()[1].table(),
            PropertyTable::Secondary
        );
        assert_eq!(shape.unknown_properties()[1].raw_value(), -7);
        assert_eq!(
            shape.unknown_properties()[2].table(),
            PropertyTable::Tertiary
        );
        assert_eq!(shape.client_anchor().unwrap().index(), Some(7));
    }
}
