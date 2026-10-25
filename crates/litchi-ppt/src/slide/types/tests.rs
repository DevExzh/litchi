#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::*;
use crate::consts::RecordType;
use crate::odraw::ShapeExt as _;
use crate::records::Record;
use crate::slide::SlideData;

const ROOT_SHAPE_FLAGS: u32 = 0x0A00;
const CHILD_SHAPE_FLAGS: u32 = 0x0A02;

fn drawing(children: &[u8]) -> Vec<u8> {
    use crate::officeart_wire::{record_type, write_atom, write_container};

    let mut body = Vec::new();
    write_atom(
        &mut body,
        0,
        0,
        record_type::DG,
        &[0; 8],
        Default::default(),
    )
    .unwrap();
    body.extend_from_slice(children);
    let mut drawing = Vec::new();
    write_container(
        &mut drawing,
        0,
        record_type::DG_CONTAINER,
        &body,
        Default::default(),
    )
    .unwrap();
    drawing
}

fn create_frame_escher_drawing(
    blip_id: u32,
    interactive_action: Option<u8>,
    external_object_id: Option<u32>,
) -> Vec<u8> {
    use crate::officeart_wire::{
        PropertyBuilder, ShapeBuilder, record_type, write_atom, write_container,
    };

    let mut shape_children = Vec::new();
    ShapeBuilder::new(75, 42)
        .with_flags(ROOT_SHAPE_FLAGS)
        .write(&mut shape_children, Default::default())
        .unwrap();
    let mut properties = PropertyBuilder::new();
    properties.add_simple(0x4104, blip_id.cast_signed());
    properties
        .write(&mut shape_children, Default::default())
        .unwrap();
    write_client_anchor(&mut shape_children, 10, 20, 210, 120).unwrap();

    let mut client_data_children = Vec::new();
    if let Some(external_object_id_value) = external_object_id {
        write_atom(
            &mut client_data_children,
            0,
            0,
            3009,
            &external_object_id_value.to_le_bytes(),
            Default::default(),
        )
        .unwrap();
    }
    if let Some(action) = interactive_action {
        let mut interactive_atom = [0u8; 16];
        interactive_atom[8] = action;
        let mut interactive_children = Vec::new();
        write_atom(
            &mut interactive_children,
            0,
            0,
            4083,
            &interactive_atom,
            Default::default(),
        )
        .unwrap();
        write_container(
            &mut client_data_children,
            0,
            4082,
            &interactive_children,
            Default::default(),
        )
        .unwrap();
    }
    if !client_data_children.is_empty() {
        write_container(
            &mut shape_children,
            0,
            record_type::CLIENT_DATA,
            &client_data_children,
            Default::default(),
        )
        .unwrap();
    }

    let mut shape_container = Vec::new();
    write_container(
        &mut shape_container,
        0,
        record_type::SP_CONTAINER,
        &shape_children,
        Default::default(),
    )
    .unwrap();

    drawing(&shape_container)
}

fn create_picture_escher_drawing(blip_id: u32) -> Vec<u8> {
    create_frame_escher_drawing(blip_id, None, None)
}

fn create_autoshape_escher_drawing() -> Vec<u8> {
    use crate::officeart_wire::{
        PropertyBuilder, ShapeBuilder, record_type, write_atom, write_container,
    };

    let mut shape_children = Vec::new();
    ShapeBuilder::new(13, 44)
        .with_flags(ROOT_SHAPE_FLAGS)
        .write(&mut shape_children, Default::default())
        .unwrap();
    let mut properties = PropertyBuilder::new();
    properties.add_simple(0x0147, 32_768);
    properties.add_simple(0x0149, -123);
    properties
        .write(&mut shape_children, Default::default())
        .unwrap();
    write_client_anchor(&mut shape_children, 11, 22, 211, 122).unwrap();

    let utf16: Vec<u8> = "Arrow label"
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect();
    let mut embedded_text = Vec::new();
    write_atom(&mut embedded_text, 0, 0, 4000, &utf16, Default::default()).unwrap();
    write_container(
        &mut shape_children,
        0,
        record_type::CLIENT_TEXTBOX,
        &embedded_text,
        Default::default(),
    )
    .unwrap();

    let mut shape_container = Vec::new();
    write_container(
        &mut shape_container,
        0,
        record_type::SP_CONTAINER,
        &shape_children,
        Default::default(),
    )
    .unwrap();

    drawing(&shape_container)
}

fn create_freeform_escher_drawing() -> Vec<u8> {
    use crate::officeart_wire::{PropertyBuilder, ShapeBuilder, record_type, write_container};

    let mut shape_children = Vec::new();
    ShapeBuilder::new(0, 45)
        .with_flags(ROOT_SHAPE_FLAGS)
        .write(&mut shape_children, Default::default())
        .unwrap();

    let mut vertices = Vec::new();
    vertices.extend_from_slice(&2u16.to_le_bytes());
    vertices.extend_from_slice(&2u16.to_le_bytes());
    vertices.extend_from_slice(&8u16.to_le_bytes());
    for (x, y) in [(0i32, 0i32), (21600, 21600)] {
        vertices.extend_from_slice(&x.to_le_bytes());
        vertices.extend_from_slice(&y.to_le_bytes());
    }
    let mut properties = PropertyBuilder::new();
    properties.add_simple(0x0140, 0);
    properties.add_simple(0x0141, 0);
    properties.add_simple(0x0142, 21600);
    properties.add_simple(0x0143, 21600);
    properties.add_simple(0x0144, 4);
    properties.add_complex(0x0145, &vertices);
    let segments = [
        2, 0, 2, 0, 2, 0, // IMsoArray header
        0x00, 0x40, // moveTo
        0x00, 0x80, // end
    ];
    properties.add_complex(0x0146, &segments);
    properties
        .write(&mut shape_children, Default::default())
        .unwrap();
    write_client_anchor(&mut shape_children, 5, 6, 105, 206).unwrap();

    let mut shape_container = Vec::new();
    write_container(
        &mut shape_container,
        0,
        record_type::SP_CONTAINER,
        &shape_children,
        Default::default(),
    )
    .unwrap();
    drawing(&shape_container)
}

fn create_animated_escher_drawing() -> Vec<u8> {
    use crate::animation::{
        AnimationInfo, LegacyAnimationAtom, LegacyAnimationBuild, LegacyAnimationEffect,
        write_animation_info,
    };
    use crate::officeart_wire::{ShapeBuilder, record_type, write_container};

    let atom = LegacyAnimationAtom {
        build_type: LegacyAnimationBuild::OneBuild,
        effect: LegacyAnimationEffect::Fade,
        order_id: 2,
        ..LegacyAnimationAtom::default()
    };
    let mut info = AnimationInfo::new();
    info.legacy_atom = Some(atom);
    let (animation, _) = write_animation_info(&info).unwrap();

    let mut shape_children = Vec::new();
    ShapeBuilder::new(1, 88)
        .with_flags(ROOT_SHAPE_FLAGS)
        .write(&mut shape_children, Default::default())
        .unwrap();
    write_client_anchor(&mut shape_children, 10, 20, 210, 120).unwrap();
    write_container(
        &mut shape_children,
        0,
        record_type::CLIENT_DATA,
        &animation,
        Default::default(),
    )
    .unwrap();

    let mut shape_container = Vec::new();
    write_container(
        &mut shape_container,
        0,
        record_type::SP_CONTAINER,
        &shape_children,
        Default::default(),
    )
    .unwrap();
    drawing(&shape_container)
}

fn create_placeholder_escher_drawing(round_trip_records: &[u8]) -> Vec<u8> {
    use crate::officeart_wire::{ShapeBuilder, record_type, write_atom, write_container};

    let mut shape_children = Vec::new();
    ShapeBuilder::new(202, 43)
        .with_flags(ROOT_SHAPE_FLAGS)
        .write(&mut shape_children, Default::default())
        .unwrap();
    write_client_anchor(&mut shape_children, 15, 25, 315, 125).unwrap();

    let utf16: Vec<u8> = "Slide title"
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect();
    let mut embedded_text = Vec::new();
    write_atom(&mut embedded_text, 0, 0, 4000, &utf16, Default::default()).unwrap();
    write_container(
        &mut shape_children,
        0,
        record_type::CLIENT_TEXTBOX,
        &embedded_text,
        Default::default(),
    )
    .unwrap();

    let mut placeholder_data = Vec::new();
    placeholder_data.extend_from_slice(&3u32.to_le_bytes());
    placeholder_data.push(13); // native slide title placeholder
    placeholder_data.push(2); // quarter size
    placeholder_data.extend_from_slice(&0u16.to_le_bytes());
    let mut client_data_children = Vec::new();
    write_atom(
        &mut client_data_children,
        0,
        0,
        3011,
        &placeholder_data,
        Default::default(),
    )
    .unwrap();
    client_data_children.extend_from_slice(round_trip_records);
    write_container(
        &mut shape_children,
        0,
        record_type::CLIENT_DATA,
        &client_data_children,
        Default::default(),
    )
    .unwrap();

    let mut shape_container = Vec::new();
    write_container(
        &mut shape_container,
        0,
        record_type::SP_CONTAINER,
        &shape_children,
        Default::default(),
    )
    .unwrap();

    drawing(&shape_container)
}

fn create_round_trip_placeholder_escher_drawing(
    shape_type: u16,
    round_trip_records: &[u8],
) -> Vec<u8> {
    use crate::officeart_wire::{ShapeBuilder, record_type, write_container};

    let mut shape_children = Vec::new();
    ShapeBuilder::new(shape_type, 46)
        .with_flags(ROOT_SHAPE_FLAGS)
        .write(&mut shape_children, Default::default())
        .unwrap();
    write_client_anchor(&mut shape_children, 20, 30, 220, 130).unwrap();
    write_container(
        &mut shape_children,
        0,
        record_type::CLIENT_DATA,
        round_trip_records,
        Default::default(),
    )
    .unwrap();

    let mut shape_container = Vec::new();
    write_container(
        &mut shape_container,
        0,
        record_type::SP_CONTAINER,
        &shape_children,
        Default::default(),
    )
    .unwrap();
    drawing(&shape_container)
}

fn create_table_escher_drawing() -> Vec<u8> {
    use crate::officeart_wire::{
        ShapeBuilder, record_type, write_atom, write_child_anchor, write_container, write_spgr,
    };

    fn shape_container(children: &[u8]) -> Vec<u8> {
        let mut container = Vec::new();
        write_container(
            &mut container,
            0,
            record_type::SP_CONTAINER,
            children,
            Default::default(),
        )
        .unwrap();
        container
    }

    fn table_cell(shape_id: u32, text: &str, left: i32, top: i32) -> Vec<u8> {
        let mut children = Vec::new();
        ShapeBuilder::new(1, shape_id)
            .with_flags(CHILD_SHAPE_FLAGS)
            .write(&mut children, Default::default())
            .unwrap();
        write_child_anchor(
            &mut children,
            left,
            top,
            left + 100,
            top + 50,
            Default::default(),
        )
        .unwrap();

        let utf16: Vec<u8> = text.encode_utf16().flat_map(u16::to_le_bytes).collect();
        let mut embedded_text = Vec::new();
        write_atom(&mut embedded_text, 0, 0, 4000, &utf16, Default::default()).unwrap();
        write_container(
            &mut children,
            0,
            record_type::CLIENT_TEXTBOX,
            &embedded_text,
            Default::default(),
        )
        .unwrap();
        shape_container(&children)
    }

    let mut patriarch_children = Vec::new();
    write_spgr(&mut patriarch_children, 0, 0, 0, 0, Default::default()).unwrap();
    ShapeBuilder::new(0, 1)
        .with_flags(0x0005)
        .write(&mut patriarch_children, Default::default())
        .unwrap();
    let patriarch = shape_container(&patriarch_children);

    let mut table_header_children = Vec::new();
    write_spgr(
        &mut table_header_children,
        0,
        0,
        200,
        100,
        Default::default(),
    )
    .unwrap();
    ShapeBuilder::new(0, 10)
        .with_flags(0x0201)
        .write(&mut table_header_children, Default::default())
        .unwrap();
    let mut table_properties = Vec::new();
    table_properties.extend_from_slice(&0x039Fu16.to_le_bytes());
    table_properties.extend_from_slice(&1i32.to_le_bytes());
    write_atom(
        &mut table_header_children,
        3,
        1,
        0xF122,
        &table_properties,
        Default::default(),
    )
    .unwrap();
    write_client_anchor(&mut table_header_children, 20, 30, 220, 130).unwrap();
    let table_header = shape_container(&table_header_children);

    let mut table_children = table_header;
    for (shape_id, text, left, top) in [
        (11, "A1", 0, 0),
        (12, "B1", 100, 0),
        (13, "A2", 0, 50),
        (14, "B2", 100, 50),
    ] {
        table_children.extend_from_slice(&table_cell(shape_id, text, left, top));
    }
    let mut table_group = Vec::new();
    write_container(
        &mut table_group,
        0,
        record_type::SPGR_CONTAINER,
        &table_children,
        Default::default(),
    )
    .unwrap();

    let mut root_group_children = patriarch;
    root_group_children.extend_from_slice(&table_group);
    let mut root_group = Vec::new();
    write_container(
        &mut root_group,
        0,
        record_type::SPGR_CONTAINER,
        &root_group_children,
        Default::default(),
    )
    .unwrap();

    drawing(&root_group)
}

// Helper function to create a test record
fn create_test_record(record_type: RecordType, data: Vec<u8>, children: Vec<Record>) -> Record {
    Record {
        record_type,
        record_type_raw: record_type as u16,
        version: 0,
        instance: 0,
        data_length: u32::try_from(data.len()).unwrap(),
        data: data.into(),
        children,
    }
}

fn record_bytes(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    data.extend_from_slice(&kind.to_le_bytes());
    data.extend_from_slice(&u32::try_from(payload.len()).unwrap().to_le_bytes());
    data.extend_from_slice(payload);
    data
}

fn write_client_anchor(
    data: &mut Vec<u8>,
    left: i32,
    top: i32,
    right: i32,
    bottom: i32,
) -> std::io::Result<()> {
    crate::Anchor::full(left, top, right, bottom)
        .map_err(|error| std::io::Error::new(std::io::ErrorKind::InvalidData, error.to_string()))?
        .write_to(data)
}

fn prog_tags_record(version: u8, blob_payload: &[u8]) -> Record {
    let tag_name: Vec<u8> = format!("___PPT{version}")
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect();
    let name = record_bytes(0, 0, 4026, &tag_name);
    let blob = record_bytes(0, 0, 0x138b, blob_payload);
    let mut tag_payload = name;
    tag_payload.extend_from_slice(&blob);
    let tag = record_bytes(0x0f, 0, 0x138a, &tag_payload);
    create_test_record(RecordType::ProgTags, tag, Vec::new())
}

// Helper function to create a basic slide record without children
fn create_basic_slide_record() -> Record {
    create_test_record(RecordType::Slide, vec![0u8; 8], Vec::new())
}

// Helper function to create a slide with PPDrawing
fn create_slide_with_drawing() -> Record {
    let dg = record_bytes(0, 0, 0xf008, &[0; 8]);
    let ppdrawing = create_test_record(
        RecordType::PPDrawing,
        record_bytes(0x0f, 0, 0xf002, &dg),
        Vec::new(),
    );
    create_test_record(RecordType::Slide, vec![0u8; 8], vec![ppdrawing])
}

// Helper function to create a slide with text
fn create_slide_with_text() -> Record {
    // Create a TextCharsAtom with "Test" in UTF-16 LE
    let text_data = vec![
        0x54, 0x00, // 'T'
        0x65, 0x00, // 'e'
        0x73, 0x00, // 's'
        0x74, 0x00, // 't'
    ];
    let text_atom = create_test_record(RecordType::TextCharsAtom, text_data, Vec::new());
    create_test_record(RecordType::Slide, vec![0u8; 8], vec![text_atom])
}

// Helper function to create SlideData
fn create_slide_data(record: Record, persist_id: u32, doc_data: &[u8]) -> SlideData<'_> {
    SlideData::new_for_test(persist_id, 0, record, doc_data)
}

#[test]
fn test_slide_creation() {
    let doc_data = vec![0u8; 1024];
    let record = create_basic_slide_record();
    let slide_data = create_slide_data(record, 256, &doc_data);

    let slide = Slide::from_slide_data(slide_data, 1);

    assert_eq!(slide.slide_number(), 1);
    assert_eq!(slide.persist_id(), 256);
}

#[test]
fn test_slide_number_accessor() {
    let doc_data = vec![0u8; 512];
    let record = create_basic_slide_record();
    let slide_data = create_slide_data(record, 100, &doc_data);

    let slide = Slide::from_slide_data(slide_data, 5);

    assert_eq!(slide.slide_number(), 5);
}

#[test]
fn exposes_inert_slide_library_synchronization_metadata() {
    let server: Vec<u8> = "server-id"
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect();
    let url: Vec<u8> = "http://example.com/library"
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect();
    let server_record = record_bytes(0, 0, 4026, &server);
    let url_record = record_bytes(0, 1, 4026, &url);
    let mut times = Vec::new();
    for fields in [
        [2026u16, 7, 4, 16, 12, 30, 45, 500],
        [2025u16, 1, 3, 2, 8, 0, 0, 0],
    ] {
        times.extend(fields.into_iter().flat_map(u16::to_le_bytes));
    }
    let atom = record_bytes(0, 0, 0x3715, &times);
    let container = record_bytes(0x0f, 0, 0x3714, &[server_record, url_record, atom].concat());
    let sync = Record::parse(&container, 0, &crate::SourceLimits::default())
        .unwrap()
        .0;
    let mut slide_record = create_test_record(RecordType::Slide, Vec::new(), vec![sync]);
    // The owner validates a real SlideContainer root (MS-PPT: container
    // records carry recVer 0xF).
    slide_record.version = 0x0f;
    let doc_data = vec![0u8; 32];
    let slide = Slide::from_slide_data(create_slide_data(slide_record, 256, &doc_data), 1);

    let sync_info = slide.sync_info().unwrap().unwrap();
    assert_eq!(sync_info.server_slide_id().as_str(), "server-id");
    assert_eq!(
        sync_info.slide_library_url().as_str(),
        "http://example.com/library"
    );
    assert_eq!(sync_info.server_modified().year(), 2026);
    assert_eq!(sync_info.client_inserted().year(), 2025);
    assert!(std::ptr::eq(sync_info, slide.sync_info().unwrap().unwrap()));
}

#[test]
fn exposes_direct_powerpoint12_slide_master_references() {
    let composite = Record {
        version: 0,
        instance: 0,
        record_type: RecordType::RoundTripCompositeMasterId12Atom,
        record_type_raw: 0x041d,
        data_length: 4,
        data: 17u32.to_le_bytes().into(),
        children: Vec::new(),
    };
    let mut content_data = Vec::new();
    content_data.extend_from_slice(&23u32.to_le_bytes());
    content_data.extend_from_slice(&5u16.to_le_bytes());
    content_data.extend_from_slice(&9u16.to_le_bytes());
    let content = Record {
        version: 0,
        instance: 7,
        record_type: RecordType::RoundTripContentMasterId12Atom,
        record_type_raw: 0x0422,
        data_length: 8,
        data: content_data.into(),
        children: Vec::new(),
    };
    let slide_record = create_test_record(RecordType::Slide, Vec::new(), vec![composite, content]);
    let doc_data = vec![0u8; 32];
    let slide = Slide::from_slide_data(create_slide_data(slide_record, 256, &doc_data), 1);

    let metadata = slide.powerpoint12_round_trip_metadata().unwrap();
    assert_eq!(metadata.composite_master_id, Some(17));
    let content_master = metadata.content_master.unwrap();
    assert_eq!(content_master.record_instance, 7);
    assert_eq!(content_master.main_master_id, 23);
    assert_eq!(content_master.layout_instance_id, 5);
    assert_eq!(content_master.unused, 9);
}

#[test]
fn test_persist_id_accessor() {
    let doc_data = vec![0u8; 512];
    let record = create_basic_slide_record();
    let slide_data = create_slide_data(record, 999, &doc_data);

    let slide = Slide::from_slide_data(slide_data, 1);

    assert_eq!(slide.persist_id(), 999);
}

#[test]
fn test_has_drawing_without_ppdrawing() {
    let doc_data = vec![0u8; 1024];
    let record = create_basic_slide_record();
    let slide_data = create_slide_data(record, 256, &doc_data);

    let slide = Slide::from_slide_data(slide_data, 1);

    assert!(!slide.has_drawing());
}

#[test]
fn test_has_drawing_with_ppdrawing() {
    let doc_data = vec![0u8; 1024];
    let record = create_slide_with_drawing();
    let slide_data = create_slide_data(record, 256, &doc_data);

    let slide = Slide::from_slide_data(slide_data, 1);

    assert!(slide.has_drawing());
}

#[test]
fn test_record_accessor() {
    let doc_data = vec![0u8; 1024];
    let record = create_basic_slide_record();
    let slide_data = create_slide_data(record, 256, &doc_data);

    let slide = Slide::from_slide_data(slide_data, 1);

    let rec = slide.record();
    assert_eq!(rec.record_type, RecordType::Slide);
}

#[test]
fn test_shapes_empty_slide() {
    let doc_data = vec![0u8; 1024];
    let record = create_basic_slide_record();
    let slide_data = create_slide_data(record, 256, &doc_data);

    let slide = Slide::from_slide_data(slide_data, 1);

    let shapes = slide.shapes().unwrap();
    assert_eq!(shapes.len(), 0);
}

#[test]
fn test_shapes_lazy_loading() {
    let doc_data = vec![0u8; 1024];
    let record = create_slide_with_drawing();
    let slide_data = create_slide_data(record, 256, &doc_data);

    let slide = Slide::from_slide_data(slide_data, 1);

    // First call should initialize
    let shapes1 = slide.shapes().unwrap();
    // Second call should return cached value
    let shapes2 = slide.shapes().unwrap();

    // Both should return the same reference
    assert_eq!(shapes1.len(), shapes2.len());
}

#[test]
fn test_shape_count_empty() {
    let doc_data = vec![0u8; 1024];
    let record = create_basic_slide_record();
    let slide_data = create_slide_data(record, 256, &doc_data);

    let slide = Slide::from_slide_data(slide_data, 1);

    assert_eq!(slide.shape_count().unwrap(), 0);
}

#[test]
fn test_text_extraction_empty_slide() {
    let doc_data = vec![0u8; 1024];
    let record = create_basic_slide_record();
    let slide_data = create_slide_data(record, 256, &doc_data);

    let slide = Slide::from_slide_data(slide_data, 1);

    let text = slide.text().unwrap();
    assert_eq!(text, "");
}

#[test]
fn test_text_extraction_with_text_chars_atom() {
    let doc_data = vec![0u8; 1024];
    let record = create_slide_with_text();
    let slide_data = create_slide_data(record, 256, &doc_data);

    let slide = Slide::from_slide_data(slide_data, 1);

    let text = slide.text().unwrap();
    assert_eq!(text, "Test");
}

#[test]
fn test_text_lazy_loading() {
    let doc_data = vec![0u8; 1024];
    let record = create_slide_with_text();
    let slide_data = create_slide_data(record, 256, &doc_data);

    let slide = Slide::from_slide_data(slide_data, 1);

    // First call should extract text
    let text1 = slide.text().unwrap();
    // Second call should return cached value
    let text2 = slide.text().unwrap();

    assert_eq!(text1, text2);
    assert_eq!(text1, "Test");
}

#[test]
fn test_text_extraction_with_nested_records() {
    let doc_data = vec![0u8; 1024];

    // Create nested structure: Slide -> SlideContainer -> TextCharsAtom
    let text_data = vec![
        0x41, 0x00, // 'A'
        0x42, 0x00, // 'B'
    ];
    let text_atom = create_test_record(RecordType::TextCharsAtom, text_data, Vec::new());

    let container = create_test_record(RecordType::SlideAtom, vec![0u8; 8], vec![text_atom]);

    let slide_record = create_test_record(RecordType::Slide, vec![0u8; 8], vec![container]);

    let slide_data = create_slide_data(slide_record, 256, &doc_data);
    let slide = Slide::from_slide_data(slide_data, 1);

    let text = slide.text().unwrap();
    assert_eq!(text, "AB");
}

#[test]
fn test_text_extraction_multiple_text_atoms() {
    let doc_data = vec![0u8; 1024];

    // Create multiple TextCharsAtom records
    let text1_data = vec![
        0x48, 0x00, // 'H'
        0x69, 0x00, // 'i'
    ];
    let text1 = create_test_record(RecordType::TextCharsAtom, text1_data, Vec::new());

    let text2_data = vec![
        0x42, 0x00, // 'B'
        0x79, 0x00, // 'y'
        0x65, 0x00, // 'e'
    ];
    let text2 = create_test_record(RecordType::TextCharsAtom, text2_data, Vec::new());

    let slide_record = create_test_record(RecordType::Slide, vec![0u8; 8], vec![text1, text2]);

    let slide_data = create_slide_data(slide_record, 256, &doc_data);
    let slide = Slide::from_slide_data(slide_data, 1);

    let text = slide.text().unwrap();
    // Both text atoms should be extracted and joined
    assert!(text.contains("Hi"));
    assert!(text.contains("Bye"));
}

#[test]
fn test_slide_with_different_text_atom_types() {
    let doc_data = vec![0u8; 1024];

    // Create TextBytesAtom (ASCII/ANSI encoding)
    let text_bytes = vec![0x54, 0x65, 0x78, 0x74]; // "Text" in ASCII
    let text_bytes_atom = create_test_record(RecordType::TextBytesAtom, text_bytes, Vec::new());

    let slide_record = create_test_record(RecordType::Slide, vec![0u8; 8], vec![text_bytes_atom]);

    let slide_data = create_slide_data(slide_record, 256, &doc_data);
    let slide = Slide::from_slide_data(slide_data, 1);

    let text = slide.text().unwrap();
    assert_eq!(text, "Text");
}

#[test]
fn test_multiple_slide_numbers() {
    let doc_data = vec![0u8; 1024];

    let records: Vec<_> = (0..5).map(|_| create_basic_slide_record()).collect();

    let slides: Vec<_> = records
        .into_iter()
        .enumerate()
        .map(|(i, record)| {
            let slide_data = create_slide_data(record, 100 + u32::try_from(i).unwrap(), &doc_data);
            Slide::from_slide_data(slide_data, i + 1)
        })
        .collect();

    // Verify slide numbers are correctly assigned
    for (i, slide) in slides.iter().enumerate() {
        assert_eq!(slide.slide_number(), i + 1);
        assert_eq!(slide.persist_id(), 100 + u32::try_from(i).unwrap());
    }
}

#[test]
fn test_convert_escher_to_shape_enum_with_unknown_type() {
    // This tests that unknown shape types are filtered out
    // We can't easily construct EscherShape objects in tests without
    // implementing complex test data, but we can test the None path
    // through indirect testing via shapes()

    let doc_data = vec![0u8; 1024];
    let record = create_basic_slide_record();
    let slide_data = create_slide_data(record, 256, &doc_data);

    let slide = Slide::from_slide_data(slide_data, 1);

    // Should return empty vec for slide without PPDrawing
    let shapes = slide.shapes().unwrap();
    assert_eq!(shapes.len(), 0);
}

#[test]
fn referenced_picture_frame_is_exposed_as_picture_shape() {
    let doc_data = vec![0u8; 32];
    let ppdrawing = create_test_record(
        RecordType::PPDrawing,
        create_picture_escher_drawing(7),
        Vec::new(),
    );
    let record = create_test_record(RecordType::Slide, Vec::new(), vec![ppdrawing]);
    let slide = Slide::from_slide_data(create_slide_data(record, 256, &doc_data), 1);

    let shapes = slide.shapes().unwrap();
    assert_eq!(shapes.len(), 1);
    let picture = shapes[0].as_picture().expect("picture frame");
    assert_eq!(picture.properties.id, 42);
    assert_eq!(picture.blip_id().map(litchi_odraw::image::Id::get), Some(7));
    assert_eq!(picture.properties.x, 10);
    assert_eq!(picture.properties.y, 20);
    assert_eq!(picture.properties.width, 200);
    assert_eq!(picture.properties.height, 100);
}

#[test]
fn autoshape_preserves_native_type_and_sparse_adjustments() {
    use crate::shapes::{Shape, autoshape::AutoShapeType};

    let doc_data = vec![0u8; 32];
    let ppdrawing = create_test_record(
        RecordType::PPDrawing,
        create_autoshape_escher_drawing(),
        Vec::new(),
    );
    let record = create_test_record(RecordType::Slide, Vec::new(), vec![ppdrawing]);
    let slide = Slide::from_slide_data(create_slide_data(record, 256, &doc_data), 1);

    let shapes = slide.shapes().unwrap();
    assert_eq!(shapes.len(), 1);
    let autoshape = shapes[0].as_autoshape().expect("auto shape");
    assert_eq!(autoshape.id(), 44);
    assert_eq!(autoshape.auto_shape_type(), AutoShapeType::Arrow);
    assert_eq!(autoshape.adjustments(), &[32_768, 0, -123]);
    assert_eq!(autoshape.bounds(), (11, 22, 200, 100));
    assert_eq!(autoshape.text(), "Arrow label");
    assert!(autoshape.has_text());
    assert_eq!(Shape::text(autoshape).unwrap(), "Arrow label");
}

#[test]
fn parsed_autoshape_text_is_source_bound_before_exposure() {
    use crate::shapes::{
        ShapeEnum,
        shape::{Mutation, MutationError},
    };

    let doc_data = vec![0u8; 32];
    let ppdrawing = create_test_record(
        RecordType::PPDrawing,
        create_autoshape_escher_drawing(),
        Vec::new(),
    );
    let record = create_test_record(RecordType::Slide, Vec::new(), vec![ppdrawing]);
    let slide = Slide::from_slide_data(create_slide_data(record, 256, &doc_data), 1);

    let shapes = slide.shapes();
    assert!(shapes.is_ok());
    let Ok(parsed_shapes) = shapes else {
        return;
    };
    assert_eq!(parsed_shapes.len(), 1);
    let Some(ShapeEnum::AutoShape(mut autoshape)) = parsed_shapes.first().cloned() else {
        return;
    };
    assert_eq!(
        autoshape.set_text("replacement".to_string()),
        Err(MutationError::SourceBound {
            mutation: Mutation::Text,
        })
    );
    assert_eq!(autoshape.text(), "Arrow label");
}

#[test]
fn non_primitive_shape_with_vertices_is_exposed_as_freeform_autoshape() {
    use crate::shapes::{
        Shape,
        autoshape::AutoShapeType,
        geometry::{GeometryRect, ShapePathType},
    };

    let doc_data = vec![0u8; 32];
    let ppdrawing = create_test_record(
        RecordType::PPDrawing,
        create_freeform_escher_drawing(),
        Vec::new(),
    );
    let record = create_test_record(RecordType::Slide, Vec::new(), vec![ppdrawing]);
    let slide = Slide::from_slide_data(create_slide_data(record, 256, &doc_data), 1);

    let shapes = slide.shapes().unwrap();
    assert_eq!(shapes.len(), 1);
    let freeform = shapes[0].as_autoshape().expect("freeform auto shape");
    assert_eq!(freeform.auto_shape_type(), AutoShapeType::Custom(0));
    assert_eq!(freeform.properties().id, 45);
    assert_eq!(freeform.bounds(), (5, 6, 100, 200));
    let geometry = freeform.geometry().expect("freeform geometry");
    assert_eq!(
        geometry.coordinate_space(),
        Some(GeometryRect::new(0, 0, 21600, 21600))
    );
    assert_eq!(geometry.path_type(), Some(ShapePathType::Complex));
    assert_eq!(geometry.vertices(), &[(0, 0), (21600, 21600)]);
    assert_eq!(geometry.segment_info(), &[0x4000, 0x8000]);
}

#[test]
fn ole_frame_is_distinguished_from_an_ordinary_picture() {
    use crate::shapes::{PictureFrameKind, ShapeType};

    let doc_data = vec![0u8; 32];
    let ppdrawing = create_test_record(
        RecordType::PPDrawing,
        create_frame_escher_drawing(8, Some(5), Some(77)),
        Vec::new(),
    );
    let record = create_test_record(RecordType::Slide, Vec::new(), vec![ppdrawing]);
    let slide = Slide::from_slide_data(create_slide_data(record, 256, &doc_data), 1);

    let shapes = slide.shapes().unwrap();
    assert_eq!(shapes[0].shape_type(), ShapeType::Object);
    let object = shapes[0].as_object_frame().expect("OLE frame");
    assert_eq!(object.frame_kind(), PictureFrameKind::OleObject);
    assert_eq!(object.external_object_id(), Some(77));
    assert_eq!(object.blip_id().map(litchi_odraw::image::Id::get), Some(8));
    assert_eq!(object.properties.x, 10);
    assert_eq!(object.properties.width, 200);
}

#[test]
fn media_frame_preserves_preview_and_external_object_references() {
    use crate::shapes::{PictureFrameKind, ShapeType};

    let doc_data = vec![0u8; 32];
    let ppdrawing = create_test_record(
        RecordType::PPDrawing,
        create_frame_escher_drawing(9, Some(6), Some(88)),
        Vec::new(),
    );
    let record = create_test_record(RecordType::Slide, Vec::new(), vec![ppdrawing]);
    let slide = Slide::from_slide_data(create_slide_data(record, 256, &doc_data), 1);

    let shapes = slide.shapes().unwrap();
    assert_eq!(shapes[0].shape_type(), ShapeType::Media);
    let media = shapes[0].as_media_frame().expect("media frame");
    assert_eq!(media.frame_kind(), PictureFrameKind::Media);
    assert_eq!(media.external_object_id(), Some(88));
    assert_eq!(media.blip_id().map(litchi_odraw::image::Id::get), Some(9));
    assert_eq!(media.properties.y, 20);
    assert_eq!(media.properties.height, 100);
}

#[test]
fn external_object_reference_alone_marks_an_ole_frame() {
    use crate::shapes::{PictureFrameKind, ShapeType};

    let doc_data = vec![0u8; 32];
    let ppdrawing = create_test_record(
        RecordType::PPDrawing,
        create_frame_escher_drawing(10, None, Some(99)),
        Vec::new(),
    );
    let record = create_test_record(RecordType::Slide, Vec::new(), vec![ppdrawing]);
    let slide = Slide::from_slide_data(create_slide_data(record, 256, &doc_data), 1);

    let shapes = slide.shapes().unwrap();
    assert_eq!(shapes[0].shape_type(), ShapeType::Object);
    let object = shapes[0].as_object_frame().expect("OLE frame");
    assert_eq!(object.frame_kind(), PictureFrameKind::OleObject);
    assert_eq!(object.external_object_id(), Some(99));
}

#[test]
fn placeholder_client_data_is_exposed_with_text_and_geometry() {
    use crate::shapes::{PlaceholderSize, PlaceholderType, Shape};

    let doc_data = vec![0u8; 32];
    let ppdrawing = create_test_record(
        RecordType::PPDrawing,
        create_placeholder_escher_drawing(&[]),
        Vec::new(),
    );
    let record = create_test_record(RecordType::Slide, Vec::new(), vec![ppdrawing]);
    let slide = Slide::from_slide_data(create_slide_data(record, 256, &doc_data), 1);

    let shapes = slide.shapes().unwrap();
    assert_eq!(shapes.len(), 1);
    let placeholder = shapes[0].as_placeholder().expect("title placeholder");
    assert_eq!(placeholder.id(), 43);
    assert_eq!(placeholder.placeholder_type(), PlaceholderType::Title);
    assert_eq!(placeholder.placeholder_size(), PlaceholderSize::Quarter);
    assert_eq!(placeholder.index(), Some(3));
    assert_eq!(placeholder.bounds(), (15, 25, 300, 100));
    assert_eq!(shapes[0].text().unwrap(), "Slide title");
    assert!(placeholder.has_text());
}

#[test]
fn powerpoint12_header_footer_placeholder_is_exposed_with_new_identity() {
    use crate::shapes::{PlaceholderSize, PlaceholderType};
    use crate::{HeaderFooterPlaceholder, NewPlaceholder, ShapeMetadata};

    let records = [
        record_bytes(0, 0, 0x0420, &[10]),
        record_bytes(0, 0, 0x0bdd, &[26]),
    ]
    .concat();
    let doc_data = vec![0u8; 32];
    let ppdrawing = create_test_record(
        RecordType::PPDrawing,
        create_round_trip_placeholder_escher_drawing(202, &records),
        Vec::new(),
    );
    let slide_record = create_test_record(RecordType::Slide, Vec::new(), vec![ppdrawing]);
    let slide = Slide::from_slide_data(create_slide_data(slide_record, 256, &doc_data), 1);

    let shapes = slide.shapes().unwrap();
    let placeholder = shapes[0].as_placeholder().expect("header placeholder");
    assert_eq!(placeholder.placeholder_type(), PlaceholderType::Header);
    assert_eq!(placeholder.placeholder_size(), PlaceholderSize::Half);
    assert_eq!(placeholder.index(), None);
    assert_eq!(
        shapes[0].powerpoint12_shape_metadata(),
        Some(&ShapeMetadata {
            header_footer: Some(HeaderFooterPlaceholder::Header),
            new_placeholder: Some(NewPlaceholder::Picture),
            ..ShapeMetadata::default()
        })
    );
}

#[test]
fn legacy_placeholder_identity_precedes_powerpoint12_round_trip_identity() {
    use crate::HeaderFooterPlaceholder;
    use crate::shapes::PlaceholderType;

    let footer = record_bytes(0, 0, 0x0420, &[9]);
    let doc_data = vec![0u8; 32];
    let ppdrawing = create_test_record(
        RecordType::PPDrawing,
        create_placeholder_escher_drawing(&footer),
        Vec::new(),
    );
    let slide_record = create_test_record(RecordType::Slide, Vec::new(), vec![ppdrawing]);
    let slide = Slide::from_slide_data(create_slide_data(slide_record, 256, &doc_data), 1);

    let shapes = slide.shapes().unwrap();
    let placeholder = shapes[0].as_placeholder().expect("title placeholder");
    assert_eq!(placeholder.placeholder_type(), PlaceholderType::Title);
    assert_eq!(
        shapes[0]
            .powerpoint12_shape_metadata()
            .and_then(|metadata| metadata.header_footer),
        Some(HeaderFooterPlaceholder::Footer)
    );
}

#[test]
fn new_placeholder_identity_is_inert_on_non_placeholder_shapes() {
    use crate::NewPlaceholder;

    let picture = record_bytes(0, 0, 0x0bdd, &[26]);
    let doc_data = vec![0u8; 32];
    let ppdrawing = create_test_record(
        RecordType::PPDrawing,
        create_round_trip_placeholder_escher_drawing(1, &picture),
        Vec::new(),
    );
    let slide_record = create_test_record(RecordType::Slide, Vec::new(), vec![ppdrawing]);
    let slide = Slide::from_slide_data(create_slide_data(slide_record, 256, &doc_data), 1);

    let shapes = slide.shapes().unwrap();
    assert!(shapes[0].as_autoshape().is_some());
    assert_eq!(
        shapes[0]
            .powerpoint12_shape_metadata()
            .and_then(|metadata| metadata.new_placeholder),
        Some(NewPlaceholder::Picture)
    );
}

fn slide_with_round_trip_atoms(shape_type: u16, records: &[u8]) -> Slide<'static> {
    static DOC_DATA: [u8; 32] = [0; 32];
    let ppdrawing = create_test_record(
        RecordType::PPDrawing,
        create_round_trip_placeholder_escher_drawing(shape_type, records),
        Vec::new(),
    );
    let slide_record = create_test_record(RecordType::Slide, Vec::new(), vec![ppdrawing]);
    Slide::from_slide_data(create_slide_data(slide_record, 256, &DOC_DATA), 1)
}

fn tolerated(slide: &Slide<'_>) -> Vec<(crate::Defect, u32, crate::DefectLocation)> {
    slide
        .tolerance_report()
        .unwrap()
        .defects()
        .iter()
        .map(|defect| (defect.defect, defect.context, defect.location))
        .collect()
}

/// Duplicated round-trip placeholder atoms keep the first occurrence and
/// undefined placeholder IDs read as absent; both are recorded (F5.11).
#[test]
fn repairs_duplicate_or_undefined_powerpoint12_placeholder_atoms() {
    use crate::{Defect, DefectLocation, HeaderFooterPlaceholder, NewPlaceholder, ShapeDefect};

    let location = DefectLocation {
        persist_id: Some(256),
        shape_id: Some(46),
    };
    let duplicate_hf = [
        record_bytes(0, 0, 0x0420, &[7]),
        record_bytes(0, 0, 0x0420, &[8]),
    ]
    .concat();
    let slide = slide_with_round_trip_atoms(202, &duplicate_hf);
    let metadata = slide.shapes().unwrap()[0].powerpoint12_shape_metadata();
    assert_eq!(
        metadata.and_then(|metadata| metadata.header_footer),
        Some(HeaderFooterPlaceholder::Date)
    );
    assert_eq!(
        tolerated(&slide),
        [(
            Defect::Shape(ShapeDefect::DuplicateRoundTripAtom),
            0x0420,
            location
        )]
    );

    let duplicate_new = [
        record_bytes(0, 0, 0x0bdd, &[25]),
        record_bytes(0, 0, 0x0bdd, &[26]),
    ]
    .concat();
    let slide = slide_with_round_trip_atoms(202, &duplicate_new);
    assert_eq!(
        slide.shapes().unwrap()[0]
            .powerpoint12_shape_metadata()
            .and_then(|metadata| metadata.new_placeholder),
        Some(NewPlaceholder::VerticalObject)
    );
    assert_eq!(
        tolerated(&slide),
        [(
            Defect::Shape(ShapeDefect::DuplicateRoundTripAtom),
            0x0bdd,
            location
        )]
    );

    for (atom, context) in [
        (record_bytes(0, 0, 0x0420, &[6]), 0x0420_0006),
        (record_bytes(0, 0, 0x0420, &[11]), 0x0420_000b),
        (record_bytes(0, 0, 0x0bdd, &[24]), 0x0bdd_0018),
        (record_bytes(0, 0, 0x0bdd, &[27]), 0x0bdd_001b),
    ] {
        let slide = slide_with_round_trip_atoms(202, &atom);
        let shapes = slide.shapes().unwrap();
        assert_eq!(shapes.len(), 1);
        assert_eq!(shapes[0].powerpoint12_shape_metadata(), None);
        assert_eq!(
            tolerated(&slide),
            [(
                Defect::Shape(ShapeDefect::UnknownRoundTripPlaceholderId),
                context,
                location
            )]
        );
    }

    // The strict projection still rejects every one of them.
    for malformed in [
        duplicate_hf,
        duplicate_new,
        record_bytes(0, 0, 0x0420, &[6]),
    ] {
        let drawing = create_round_trip_placeholder_escher_drawing(202, &malformed);
        let shapes = litchi_odraw::shape::parse(&drawing, Default::default()).unwrap();
        assert!(shapes[0].powerpoint12_shape_metadata().is_err());
    }
}

/// A round-trip atom with a malformed header or size drops its shape alone
/// (F5.11); a truncated atom breaks `ClientData` framing and stays fatal.
#[test]
fn isolates_malformed_powerpoint12_placeholder_atoms_per_shape() {
    use crate::{Defect, DefectLocation, ShapeDefect};

    for malformed in [
        record_bytes(1, 0, 0x0420, &[7]),
        record_bytes(0, 1, 0x0420, &[7]),
        record_bytes(0, 0, 0x0420, &[]),
    ] {
        let slide = slide_with_round_trip_atoms(202, &malformed);
        assert!(slide.shapes().unwrap().is_empty());
        assert_eq!(
            tolerated(&slide),
            [(
                Defect::Shape(ShapeDefect::UnreadableShape),
                0,
                DefectLocation {
                    persist_id: Some(256),
                    shape_id: Some(46),
                }
            )]
        );
    }

    let mut truncated = record_bytes(0, 0, 0x0420, &[7]);
    truncated[4..8].copy_from_slice(&2u32.to_le_bytes());
    let slide = slide_with_round_trip_atoms(202, &truncated);
    assert!(slide.shapes().is_err());
    assert!(slide.tolerance_report().is_err());
}

#[test]
fn accepts_every_powerpoint12_placeholder_identity() {
    use crate::{HeaderFooterPlaceholder, NewPlaceholder};

    for (id, expected) in [
        (7, HeaderFooterPlaceholder::Date),
        (8, HeaderFooterPlaceholder::SlideNumber),
        (9, HeaderFooterPlaceholder::Footer),
        (10, HeaderFooterPlaceholder::Header),
    ] {
        let atom = record_bytes(0, 0, 0x0420, &[id]);
        let drawing = create_round_trip_placeholder_escher_drawing(202, &atom);
        let shapes = litchi_odraw::shape::parse(&drawing, Default::default()).unwrap();
        assert_eq!(
            shapes[0]
                .powerpoint12_shape_metadata()
                .unwrap()
                .and_then(|metadata| metadata.header_footer),
            Some(expected)
        );
    }

    for (id, expected) in [
        (25, NewPlaceholder::VerticalObject),
        (26, NewPlaceholder::Picture),
    ] {
        let atom = record_bytes(0, 0, 0x0bdd, &[id]);
        let drawing = create_round_trip_placeholder_escher_drawing(1, &atom);
        let shapes = litchi_odraw::shape::parse(&drawing, Default::default()).unwrap();
        assert_eq!(
            shapes[0]
                .powerpoint12_shape_metadata()
                .unwrap()
                .and_then(|metadata| metadata.new_placeholder),
            Some(expected)
        );
    }
}

#[test]
fn exposes_powerpoint12_shape_id_and_custom_layout_checksums() {
    use crate::ShapeChecksums;

    let mut checksums = Vec::new();
    checksums.extend_from_slice(&0u32.to_le_bytes());
    checksums.extend_from_slice(&u32::MAX.to_le_bytes());
    let records = [
        record_bytes(0, 0, 0x041f, &u32::MAX.to_le_bytes()),
        record_bytes(0, 0, 0x0426, &checksums),
    ]
    .concat();
    let doc_data = vec![0u8; 32];
    let ppdrawing = create_test_record(
        RecordType::PPDrawing,
        create_round_trip_placeholder_escher_drawing(1, &records),
        Vec::new(),
    );
    let slide_record = create_test_record(RecordType::Slide, Vec::new(), vec![ppdrawing]);
    let slide = Slide::from_slide_data(create_slide_data(slide_record, 256, &doc_data), 1);
    let metadata = slide.shapes().unwrap()[0]
        .powerpoint12_shape_metadata()
        .unwrap();

    assert_eq!(metadata.shape_id, Some(u32::MAX));
    assert_eq!(
        metadata.custom_layout_checksums,
        Some(ShapeChecksums {
            shape: 0,
            text: u32::MAX,
        })
    );
}

#[test]
fn rejects_malformed_or_duplicate_powerpoint12_shape_round_trip_atoms() {
    let duplicate_id = [
        record_bytes(0, 0, 0x041f, &1u32.to_le_bytes()),
        record_bytes(0, 0, 0x041f, &2u32.to_le_bytes()),
    ]
    .concat();
    let checksum = [0u8; 8];
    let duplicate_checksums = [
        record_bytes(0, 0, 0x0426, &checksum),
        record_bytes(0, 0, 0x0426, &checksum),
    ]
    .concat();
    let mut truncated_id = record_bytes(0, 0, 0x041f, &[0; 3]);
    truncated_id[4..8].copy_from_slice(&4u32.to_le_bytes());
    let mut truncated_checksums = record_bytes(0, 0, 0x0426, &[0; 7]);
    truncated_checksums[4..8].copy_from_slice(&8u32.to_le_bytes());

    for malformed in [
        record_bytes(1, 0, 0x041f, &0u32.to_le_bytes()),
        record_bytes(0, 1, 0x041f, &0u32.to_le_bytes()),
        record_bytes(0, 0, 0x041f, &[0; 3]),
        record_bytes(0, 0, 0x041f, &[0; 5]),
        record_bytes(1, 0, 0x0426, &checksum),
        record_bytes(0, 1, 0x0426, &checksum),
        record_bytes(0, 0, 0x0426, &[0; 7]),
        record_bytes(0, 0, 0x0426, &[0; 9]),
        truncated_id,
        truncated_checksums,
        duplicate_id,
        duplicate_checksums,
    ] {
        let drawing = create_round_trip_placeholder_escher_drawing(1, &malformed);
        let shapes = litchi_odraw::shape::parse(&drawing, Default::default()).unwrap();
        assert!(shapes[0].powerpoint12_shape_metadata().is_err());
    }
}

#[test]
fn table_group_is_exposed_with_grid_and_text() {
    let doc_data = vec![0u8; 32];
    let ppdrawing = create_test_record(
        RecordType::PPDrawing,
        create_table_escher_drawing(),
        Vec::new(),
    );
    let record = create_test_record(RecordType::Slide, Vec::new(), vec![ppdrawing]);
    let slide = Slide::from_slide_data(create_slide_data(record, 256, &doc_data), 1);

    let shapes = slide.shapes().unwrap();
    assert_eq!(shapes.len(), 1);
    let table = shapes[0].as_table().expect("table group");
    assert_eq!(table.id(), 10);
    assert_eq!(table.rows(), 2);
    assert_eq!(table.columns(), 2);
    assert_eq!(table.cell(0, 0), Some("A1"));
    assert_eq!(table.cell(0, 1), Some("B1"));
    assert_eq!(table.cell(1, 0), Some("A2"));
    assert_eq!(table.cell(1, 1), Some("B2"));
    assert_eq!((table.left(), table.top()), (20, 30));
    assert_eq!((table.width(), table.height()), (200, 100));
}

#[test]
fn test_extract_text_recursive_depth() {
    let doc_data = vec![0u8; 1024];

    // Create deeply nested structure
    let text_data = vec![0x58, 0x00]; // 'X'
    let text_atom = create_test_record(RecordType::TextCharsAtom, text_data, Vec::new());

    let level3 = create_test_record(RecordType::SlideAtom, vec![], vec![text_atom]);

    let level2 = create_test_record(RecordType::SlideAtom, vec![], vec![level3]);

    let level1 = create_test_record(RecordType::Slide, vec![], vec![level2]);

    let slide_data = create_slide_data(level1, 256, &doc_data);
    let slide = Slide::from_slide_data(slide_data, 1);

    let text = slide.text().unwrap();
    assert_eq!(text, "X");
}

#[test]
fn test_slide_with_whitespace_only_text() {
    let doc_data = vec![0u8; 1024];

    // Create TextCharsAtom with only whitespace
    let text_data = vec![
        0x20, 0x00, // space
        0x20, 0x00, // space
        0x09, 0x00, // tab
    ];
    let text_atom = create_test_record(RecordType::TextCharsAtom, text_data, Vec::new());

    let slide_record = create_test_record(RecordType::Slide, vec![], vec![text_atom]);

    let slide_data = create_slide_data(slide_record, 256, &doc_data);
    let slide = Slide::from_slide_data(slide_data, 1);

    let text = slide.text().unwrap();
    // Whitespace-only text should be filtered out
    assert_eq!(text, "");
}

#[test]
fn test_slide_zero_based_vs_one_based_numbering() {
    let doc_data = vec![0u8; 1024];
    let record = create_basic_slide_record();

    // Test that slide_number is 1-based (display number)
    let slide_data = create_slide_data(record, 256, &doc_data);
    let slide = Slide::from_slide_data(slide_data, 1);

    assert_eq!(slide.slide_number(), 1); // 1-based for user display
}

#[test]
fn test_shape_count_matches_shapes_len() {
    let doc_data = vec![0u8; 1024];
    let record = create_slide_with_drawing();
    let slide_data = create_slide_data(record, 256, &doc_data);

    let slide = Slide::from_slide_data(slide_data, 1);

    let shape_count = slide.shape_count().unwrap();
    let shapes_len = slide.shapes().unwrap().len();

    assert_eq!(shape_count, shapes_len);
}

#[test]
fn test_text_and_shapes_independent_caching() {
    let doc_data = vec![0u8; 1024];

    // Create slide with both text and PPDrawing
    let text_data = vec![0x41, 0x00]; // 'A'
    let text_atom = create_test_record(RecordType::TextCharsAtom, text_data, Vec::new());

    let dg = record_bytes(0, 0, 0xf008, &[0; 8]);
    let ppdrawing = create_test_record(
        RecordType::PPDrawing,
        record_bytes(0x0f, 0, 0xf002, &dg),
        Vec::new(),
    );

    let slide_record = create_test_record(RecordType::Slide, vec![], vec![text_atom, ppdrawing]);

    let slide_data = create_slide_data(slide_record, 256, &doc_data);
    let slide = Slide::from_slide_data(slide_data, 1);

    // Access text first
    let text = slide.text().unwrap();
    assert_eq!(text, "A");

    // Then access shapes - should work independently
    let shapes = slide.shapes().unwrap();
    assert_eq!(shapes.len(), 0);

    // Access again to verify both caches work
    let text2 = slide.text().unwrap();
    let shapes2 = slide.shapes().unwrap();

    assert_eq!(text, text2);
    assert_eq!(shapes.len(), shapes2.len());
}

#[test]
fn test_slide_with_cstring_record() {
    let doc_data = vec![0u8; 1024];

    // CString records contain UTF-16LE text.
    let cstring_data = "Hi😀".encode_utf16().flat_map(u16::to_le_bytes).collect();
    let cstring = create_test_record(RecordType::CString, cstring_data, Vec::new());

    let slide_record = create_test_record(RecordType::Slide, vec![], vec![cstring]);

    let slide_data = create_slide_data(slide_record, 256, &doc_data);
    let slide = Slide::from_slide_data(slide_data, 1);

    let text = slide.text().unwrap();
    assert_eq!(text, "Hi😀");
}

#[test]
fn test_large_persist_id() {
    let doc_data = vec![0u8; 1024];
    let record = create_basic_slide_record();

    // Test with large persist ID
    let large_id = u32::MAX - 1;
    let slide_data = create_slide_data(record, large_id, &doc_data);
    let slide = Slide::from_slide_data(slide_data, 1);

    assert_eq!(slide.persist_id(), large_id);
}

#[test]
fn test_slide_with_empty_data() {
    let doc_data = vec![0u8; 0]; // Empty document data
    let record = create_basic_slide_record();
    let slide_data = create_slide_data(record, 256, &doc_data);

    let slide = Slide::from_slide_data(slide_data, 1);

    // Should still work with basic accessors
    assert_eq!(slide.slide_number(), 1);
    assert_eq!(slide.persist_id(), 256);
    assert!(!slide.has_drawing());
}

#[test]
fn exposes_inert_shape_animations_from_the_slide() {
    let doc_data = vec![0u8; 1024];
    let ppdrawing = create_test_record(
        RecordType::PPDrawing,
        create_animated_escher_drawing(),
        Vec::new(),
    );
    let slide_record = create_test_record(RecordType::Slide, Vec::new(), vec![ppdrawing]);
    let slide = Slide::from_slide_data(create_slide_data(slide_record, 256, &doc_data), 1);

    let animations = slide.animations().unwrap();
    assert_eq!(animations.len(), 1);
    assert_eq!(animations[0].shape_id, 88);
    let atom = animations[0].animation.legacy_atom.as_ref().unwrap();
    assert_eq!(atom.effect, crate::animation::LegacyAnimationEffect::Fade);
    assert_eq!(atom.order_id, 2);
}

#[test]
fn exposes_powerpoint_2002_animation_extension_from_programmable_tags() {
    use crate::animation::{
        BuildList, ExtendedTimeNode, TimeNodeAtom, TimeNodeKind, write_build_list,
        write_extended_time_node,
    };
    use crate::officeart_wire::{write_atom, write_container};
    use crate::writer::comments::{SlideComment, build_slide_comments};

    let timing = ExtendedTimeNode {
        atom: TimeNodeAtom {
            node_type: Some(TimeNodeKind::Sequential),
            duration_ms: Some(2_000),
            ..TimeNodeAtom::default()
        },
        ..ExtendedTimeNode::default()
    };
    let comment = SlideComment::new("Ada Lovelace", "Animate this", 12, 34);
    let mut extension_data = build_slide_comments(&[comment]).unwrap();
    extension_data.extend(write_extended_time_node(&timing).unwrap());
    extension_data.extend(write_build_list(&BuildList::new()).unwrap());

    let tag_name: Vec<u8> = "___PPT10"
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect();
    let mut prog_binary_children = Vec::new();
    write_atom(
        &mut prog_binary_children,
        0,
        0,
        4026,
        &tag_name,
        Default::default(),
    )
    .unwrap();
    write_atom(
        &mut prog_binary_children,
        0,
        0,
        RecordType::BinaryTagData.as_u16(),
        &extension_data,
        Default::default(),
    )
    .unwrap();
    let mut prog_tags_children = Vec::new();
    write_container(
        &mut prog_tags_children,
        0,
        RecordType::ProgBinaryTag.as_u16(),
        &prog_binary_children,
        Default::default(),
    )
    .unwrap();
    let mut slide_children = Vec::new();
    write_container(
        &mut slide_children,
        0,
        RecordType::ProgTags.as_u16(),
        &prog_tags_children,
        Default::default(),
    )
    .unwrap();
    let mut bytes = Vec::new();
    write_container(
        &mut bytes,
        0,
        RecordType::Slide.as_u16(),
        &slide_children,
        Default::default(),
    )
    .unwrap();

    let (record, consumed) = Record::parse(&bytes, 0, &crate::SourceLimits::default()).unwrap();
    assert_eq!(consumed, bytes.len());
    let doc_data = Vec::new();
    let slide = Slide::from_slide_data(create_slide_data(record, 256, &doc_data), 1);
    let extension = slide.animation_extension().unwrap().unwrap();
    assert_eq!(extension.time_node, Some(timing));
    assert_eq!(extension.build_list, Some(BuildList::new()));
    let comments = slide.comments().unwrap();
    assert_eq!(comments.len(), 1);
    assert_eq!(comments[0].author, "Ada Lovelace");
    assert_eq!(comments[0].text, "Animate this");
    assert_eq!(slide.text().unwrap(), "");
}

#[test]
fn ignores_powerpoint10_slide_settings_in_other_tag_versions() {
    let flags = record_bytes(0, 0, RecordType::SlideFlags10Atom.as_u16(), &[3, 0, 0, 0]);
    let slide_record = create_test_record(
        RecordType::Slide,
        Vec::new(),
        vec![prog_tags_record(9, &flags)],
    );
    let doc_data = Vec::new();
    let slide = Slide::from_slide_data(create_slide_data(slide_record, 256, &doc_data), 1);

    assert_eq!(slide.animation_extension().unwrap(), None);
}

#[test]
fn truncated_comment_atoms_are_rejected_without_panicking() {
    let mut child = Vec::new();
    child.extend(0u16.to_le_bytes());
    child.extend(RecordType::Comment2000Atom.as_u16().to_le_bytes());
    child.extend(28u32.to_le_bytes());
    child.push(0);

    let mut data = Vec::new();
    data.extend(0x000Fu16.to_le_bytes());
    data.extend(RecordType::Comment2000.as_u16().to_le_bytes());
    data.extend(u32::try_from(child.len()).unwrap().to_le_bytes());
    data.extend(child);

    let extension = prog_tags_record(10, &data);
    let slide = create_test_record(RecordType::Slide, Vec::new(), vec![extension]);
    assert!(
        crate::comments::parse_slide_comments(&slide, &crate::SourceLimits::default()).is_err()
    );
}

#[test]
fn test_slide_text_extraction_preserves_order() {
    let doc_data = vec![0u8; 1024];

    // Create multiple text atoms in specific order
    let text1 = create_test_record(
        RecordType::TextCharsAtom,
        vec![0x31, 0x00], // '1'
        Vec::new(),
    );

    let text2 = create_test_record(
        RecordType::TextCharsAtom,
        vec![0x32, 0x00], // '2'
        Vec::new(),
    );

    let text3 = create_test_record(
        RecordType::TextCharsAtom,
        vec![0x33, 0x00], // '3'
        Vec::new(),
    );

    let slide_record = create_test_record(RecordType::Slide, vec![], vec![text1, text2, text3]);

    let slide_data = create_slide_data(slide_record, 256, &doc_data);
    let slide = Slide::from_slide_data(slide_data, 1);

    let text = slide.text().unwrap();
    // Text should be extracted in order and joined with newlines
    assert!(text.contains('1'));
    assert!(text.contains('2'));
    assert!(text.contains('3'));
    // Verify order is preserved
    let pos1 = text.find('1').unwrap();
    let pos2 = text.find('2').unwrap();
    let pos3 = text.find('3').unwrap();
    assert!(pos1 < pos2);
    assert!(pos2 < pos3);
}

// ---------------------------------------------------------------------------
// L-PPT-R1: outline text references resolved into shapes (finding F-10).
// ---------------------------------------------------------------------------

const OUTLINE_TEXT_REF_ATOM: u16 = 0x0f9e;

fn utf16(text: &str) -> Vec<u8> {
    text.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

/// One slide's slide-list text: a title body and a two-paragraph body.
fn outline_bodies() -> Vec<crate::OutlineTextBody> {
    let bytes = [
        record_bytes(0, 0, 3999, &6u32.to_le_bytes()),
        record_bytes(0, 0, 4000, &utf16("Resolved title")),
        record_bytes(0, 1, 3999, &5u32.to_le_bytes()),
        record_bytes(0, 0, 4000, &utf16("Body\rtext")),
    ]
    .concat();
    let mut records = Vec::new();
    let mut offset = 0;
    while offset < bytes.len() {
        let (record, consumed) =
            Record::parse(&bytes, offset, &crate::SourceLimits::default()).unwrap();
        records.push(record);
        offset += consumed;
    }
    crate::OutlineTextBody::collect(&records.iter().collect::<Vec<_>>()).unwrap()
}

/// A shape container whose `ClientTextbox` holds `textbox_records`, with a
/// `PlaceholderAtom` when `placeholder` names a position and kind.
fn outline_shape(
    native_kind: u16,
    id: u32,
    placeholder: Option<(i32, u8)>,
    textbox_records: &[u8],
) -> Vec<u8> {
    use crate::officeart_wire::{ShapeBuilder, record_type, write_atom, write_container};

    let mut children = Vec::new();
    ShapeBuilder::new(native_kind, id)
        .with_flags(ROOT_SHAPE_FLAGS)
        .write(&mut children, Default::default())
        .unwrap();
    write_client_anchor(&mut children, 10, 10, 110, 60).unwrap();
    if let Some((position, kind)) = placeholder {
        let mut data = Vec::new();
        data.extend_from_slice(&position.to_le_bytes());
        data.push(kind);
        data.push(0);
        data.extend_from_slice(&0u16.to_le_bytes());
        let mut client_data = Vec::new();
        write_atom(&mut client_data, 0, 0, 3011, &data, Default::default()).unwrap();
        write_container(
            &mut children,
            0,
            record_type::CLIENT_DATA,
            &client_data,
            Default::default(),
        )
        .unwrap();
    }
    write_container(
        &mut children,
        0,
        record_type::CLIENT_TEXTBOX,
        textbox_records,
        Default::default(),
    )
    .unwrap();
    let mut container = Vec::new();
    write_container(
        &mut container,
        0,
        record_type::SP_CONTAINER,
        &children,
        Default::default(),
    )
    .unwrap();
    container
}

fn outline_ref(index: i32) -> Vec<u8> {
    record_bytes(0, 0, OUTLINE_TEXT_REF_ATOM, &index.to_le_bytes())
}

/// Shapes 100-106: a resolved title placeholder, a dangling, a negative and
/// a malformed reference, an auto shape and a text box that reference the
/// body, and a placeholder that holds inline text beside its reference.
fn outline_drawing() -> Vec<u8> {
    let inline_title = record_bytes(0, 0, 4000, &utf16("Inline"));
    let ruler = record_bytes(0, 0, 4006, &0u32.to_le_bytes());
    let shapes = [
        outline_shape(1, 100, Some((0, 13)), &outline_ref(0)),
        outline_shape(1, 101, Some((1, 14)), &outline_ref(7)),
        outline_shape(1, 102, Some((2, 14)), &outline_ref(-2)),
        outline_shape(
            1,
            103,
            Some((3, 14)),
            &record_bytes(0, 0, OUTLINE_TEXT_REF_ATOM, &[1, 0]),
        ),
        outline_shape(1, 104, None, &outline_ref(1)),
        outline_shape(202, 105, None, &[outline_ref(1), ruler].concat()),
        outline_shape(
            1,
            106,
            Some((4, 13)),
            &[outline_ref(0), inline_title].concat(),
        ),
    ]
    .concat();
    drawing(&shapes)
}

fn unresolved_defects(report: &crate::leniency::ToleranceReport) -> Vec<(u32, Option<u32>)> {
    report
        .defects()
        .iter()
        .filter(|defect| {
            defect.defect
                == crate::Defect::Shape(crate::leniency::ShapeDefect::UnresolvedOutlineTextRef)
        })
        .map(|defect| (defect.context, defect.location.shape_id))
        .collect()
}

#[test]
fn outline_references_supply_text_and_record_what_they_cannot_resolve() {
    let bodies = outline_bodies();
    let drawing = outline_drawing();
    let mut report = crate::leniency::ToleranceReport::default();
    let shapes = Slide::parse_shape_enums(
        &drawing,
        &bodies,
        crate::SourceLimits::default(),
        &mut report,
    )
    .unwrap();

    let texts = shapes
        .iter()
        .map(|shape| shape.text().unwrap())
        .collect::<Vec<_>>();
    assert_eq!(
        texts,
        [
            "Resolved title",
            "",
            "",
            "",
            "Body\rtext",
            "Body\rtext",
            // The reference wins over inline text beside it, as in POI and
            // LibreOffice.
            "Resolved title",
        ]
    );
    assert!(shapes[0].as_placeholder().is_some());
    assert!(shapes[4].as_autoshape().is_some());
    let textbox = shapes[5].as_textbox().expect("text box kind");
    assert_eq!(textbox.outline_text_refs().len(), 1);
    // The text box keeps its own ruler; the body has none.
    assert!(textbox.text_ruler().is_some());
    assert_eq!(
        unresolved_defects(&report),
        [
            (7, Some(101)),
            (0xffff_fffe, Some(102)),
            (u32::MAX, Some(103))
        ]
    );

    // Slide text resolves the same references in drawing order (each text
    // box ends with a line feed; `Slide::text` trims the whole).
    assert_eq!(
        crate::odraw::text_from_slide_drawing(&drawing, &bodies, crate::SourceLimits::default())
            .unwrap(),
        "Resolved title\nBody\rtext\nBody\rtext\nResolved title\n"
    );
    // Without a slide, only inline text is read.
    assert_eq!(
        crate::odraw::text_from_drawing(&drawing, crate::SourceLimits::default()).unwrap(),
        "Inline\n"
    );
}

#[test]
fn outline_resolution_never_changes_which_shapes_are_read() {
    let drawing = outline_drawing();
    let identity = |shapes: &[ShapeEnum<'static>]| {
        shapes
            .iter()
            .map(|shape| (shape.shape_type(), shape.text().is_ok()))
            .collect::<Vec<_>>()
    };
    let mut resolved_report = crate::leniency::ToleranceReport::default();
    let resolved = Slide::parse_shape_enums(
        &drawing,
        &outline_bodies(),
        crate::SourceLimits::default(),
        &mut resolved_report,
    )
    .unwrap();
    let mut plain_report = crate::leniency::ToleranceReport::default();
    let plain = Slide::parse_shape_enums(
        &drawing,
        &[],
        crate::SourceLimits::default(),
        &mut plain_report,
    )
    .unwrap();
    assert_eq!(identity(&resolved), identity(&plain));
    // Without bodies every reference is unresolved, and every shape is kept.
    assert_eq!(unresolved_defects(&plain_report).len(), 7);
    assert!(
        plain
            .iter()
            .take(6)
            .all(|shape| shape.text().unwrap().is_empty())
    );
}

#[test]
fn outline_resolution_keeps_the_shapes_of_every_corpus_slide() {
    for path in [
        "../../test-data/ole/ppt/text_shapes.ppt",
        "../../test-data/ole/ppt/empty_textbox.ppt",
        "../../test-data/poi/test-data/slideshow/41246-1.ppt",
        "../../test-data/poi/test-data/slideshow/ole2-embedding-2003.ppt",
        "../../test-data/poi/test-data/slideshow/WithLinks.ppt",
    ] {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(path);
        let package = crate::Package::open(&path).unwrap();
        let presentation = package.presentation().unwrap();
        for slide in presentation.slides().unwrap() {
            let Some(ppdrawing) = slide.record().find_child(RecordType::PPDrawing) else {
                continue;
            };
            // The source-backed text editor's view of the slide.
            let plain = Slide::parse_shape_enums(
                &ppdrawing.data,
                &[],
                crate::SourceLimits::default(),
                &mut crate::leniency::ToleranceReport::default(),
            )
            .unwrap();
            let resolved = slide.shapes().unwrap();
            assert_eq!(
                resolved
                    .iter()
                    .map(ShapeEnum::shape_type)
                    .collect::<Vec<_>>(),
                plain.iter().map(ShapeEnum::shape_type).collect::<Vec<_>>(),
                "{}",
                path.display()
            );
        }
    }
}

// ---------------------------------------------------------------------------
// L-PPT-R3 (name and alternative text) and L-PPT-R4 (stable shape index).
// ---------------------------------------------------------------------------

fn utf16z(text: &str) -> Vec<u8> {
    let mut bytes = utf16(text);
    bytes.extend_from_slice(&[0, 0]);
    bytes
}

/// A non-group shape container: `native_kind` and `id`, the flags of a root
/// or group-child shape, an anchor unless `anchored` is false, and the
/// complex `wzName`/`wzDescription` payloads given.
fn identity_shape(
    native_kind: u16,
    id: u32,
    child: bool,
    anchored: bool,
    name: Option<&[u8]>,
    description: Option<&[u8]>,
) -> Vec<u8> {
    use crate::officeart_wire::{
        PropertyBuilder, ShapeBuilder, record_type, write_child_anchor, write_container,
    };

    let mut flags = 0x0800 | if child { 0x0002 } else { 0 };
    if anchored {
        flags |= 0x0200;
    }
    let mut children = Vec::new();
    ShapeBuilder::new(native_kind, id)
        .with_flags(flags)
        .write(&mut children, Default::default())
        .unwrap();
    if name.is_some() || description.is_some() {
        let mut properties = PropertyBuilder::new();
        if let Some(name) = name {
            properties.add_complex(0x0380, name);
        }
        if let Some(description) = description {
            properties.add_complex(0x0381, description);
        }
        properties.write(&mut children, Default::default()).unwrap();
    }
    if anchored {
        if child {
            write_child_anchor(&mut children, 0, 0, 100, 50, Default::default()).unwrap();
        } else {
            write_client_anchor(&mut children, 10, 10, 110, 60).unwrap();
        }
    }
    let mut container = Vec::new();
    write_container(
        &mut container,
        0,
        record_type::SP_CONTAINER,
        &children,
        Default::default(),
    )
    .unwrap();
    container
}

/// Root group: an unanchored line (dropped), a named rectangle with
/// alternative text, an unknown kind (dropped), a group whose first child is
/// an unanchored line (dropped) and whose second is named, a rectangle whose
/// name has no terminator and a trailing odd byte, and one whose name is
/// empty.
fn identity_drawing() -> Vec<u8> {
    use crate::officeart_wire::{ShapeBuilder, record_type, write_container, write_spgr};

    let header = |id: u32, flags: u32, anchored: bool| {
        let mut children = Vec::new();
        write_spgr(&mut children, 0, 0, 100, 100, Default::default()).unwrap();
        ShapeBuilder::new(0, id)
            .with_flags(flags)
            .write(&mut children, Default::default())
            .unwrap();
        if anchored {
            write_client_anchor(&mut children, 0, 0, 100, 100).unwrap();
        }
        let mut container = Vec::new();
        write_container(
            &mut container,
            0,
            record_type::SP_CONTAINER,
            &children,
            Default::default(),
        )
        .unwrap();
        container
    };

    let mut group_children = header(20, 0x0201, true);
    group_children.extend(identity_shape(20, 21, true, false, None, None));
    group_children.extend(identity_shape(
        1,
        22,
        true,
        true,
        Some(&utf16z("Inner")),
        None,
    ));
    let mut group = Vec::new();
    write_container(
        &mut group,
        0,
        record_type::SPGR_CONTAINER,
        &group_children,
        Default::default(),
    )
    .unwrap();

    let mut odd = utf16("Odd");
    odd.push(b'!');
    let mut root_children = header(1, 0x0005, false);
    root_children.extend(identity_shape(20, 2, false, false, None, None));
    root_children.extend(identity_shape(
        1,
        3,
        false,
        true,
        Some(&utf16z("Logo")),
        Some(&utf16z("Company logo")),
    ));
    root_children.extend(identity_shape(300, 4, false, true, None, None));
    root_children.extend(group);
    root_children.extend(identity_shape(1, 5, false, true, Some(&odd), None));
    root_children.extend(identity_shape(1, 6, false, true, Some(&[0, 0]), None));
    let mut root = Vec::new();
    write_container(
        &mut root,
        0,
        record_type::SPGR_CONTAINER,
        &root_children,
        Default::default(),
    )
    .unwrap();
    drawing(&root)
}

fn identity_slide(doc_data: &[u8]) -> Slide<'_> {
    let ppdrawing = create_test_record(RecordType::PPDrawing, identity_drawing(), Vec::new());
    let record = create_test_record(RecordType::Slide, Vec::new(), vec![ppdrawing]);
    Slide::from_slide_data(create_slide_data(record, 256, doc_data), 1)
}

#[test]
fn shapes_keep_their_name_and_alternative_text() {
    let doc_data = vec![0u8; 32];
    let slide = identity_slide(&doc_data);
    let shapes = slide.shapes().unwrap();
    let names = shapes.iter().map(ShapeEnum::name).collect::<Vec<_>>();
    assert_eq!(names, [Some("Logo"), None, Some("Odd"), None]);
    assert_eq!(shapes[0].description(), Some("Company logo"));
    assert_eq!(
        crate::shapes::Shape::properties(shapes[0].as_autoshape().unwrap())
            .name
            .as_deref(),
        Some("Logo")
    );
    assert!(
        shapes[1..]
            .iter()
            .all(|shape| shape.description().is_none())
    );
    let group = shapes[1].as_group().expect("group");
    assert_eq!(group.children().len(), 1);
    assert_eq!(group.children()[0].name(), Some("Inner"));
}

#[test]
fn shapes_keep_their_source_position_when_siblings_are_dropped() {
    let doc_data = vec![0u8; 32];
    let slide = identity_slide(&doc_data);
    let shapes = slide.shapes().unwrap();
    // The unanchored line (position 0) and the unknown kind (position 2)
    // are not returned; the others keep their drawing positions.
    assert_eq!(
        shapes
            .iter()
            .map(|shape| (shape.id(), shape.source_index()))
            .collect::<Vec<_>>(),
        [(3, Some(1)), (20, Some(3)), (5, Some(4)), (6, Some(5))]
    );
    let inner = &shapes[1].as_group().unwrap().children()[0];
    assert_eq!((inner.id(), inner.source_index()), (22, Some(1)));

    // A second read of the same drawing reports the same positions.
    let again = identity_slide(&doc_data);
    assert_eq!(
        again
            .shapes()
            .unwrap()
            .iter()
            .map(ShapeEnum::source_index)
            .collect::<Vec<_>>(),
        [Some(1), Some(3), Some(4), Some(5)]
    );

    // A shape built by a caller has no source position.
    let detached: ShapeEnum<'static> =
        ShapeEnum::Line(crate::shapes::shape_enum::LineShape::new(9, 0, 0, 1, 1));
    assert_eq!(detached.source_index(), None);
    assert_eq!(detached.id(), 9);
}

#[test]
fn shape_by_id_finds_group_children_and_only_returned_shapes() {
    let doc_data = vec![0u8; 32];
    let slide = identity_slide(&doc_data);
    assert_eq!(slide.shape_by_id(3).unwrap().unwrap().name(), Some("Logo"));
    assert_eq!(
        slide.shape_by_id(22).unwrap().unwrap().name(),
        Some("Inner")
    );
    assert!(slide.shape_by_id(20).unwrap().unwrap().as_group().is_some());
    // The dropped line and an unknown identifier find nothing.
    assert!(slide.shape_by_id(2).unwrap().is_none());
    assert!(slide.shape_by_id(21).unwrap().is_none());
    assert!(slide.shape_by_id(999).unwrap().is_none());
}

#[test]
fn corpus_shapes_report_their_names_and_positions() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/ppt/SampleShow.ppt");
    let package = crate::Package::open(&path).unwrap();
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let identities = |slide: &Slide<'_>| {
        slide
            .shapes()
            .unwrap()
            .iter()
            .map(|shape| {
                (
                    shape.id(),
                    shape.source_index(),
                    shape.name().map(str::to_owned),
                )
            })
            .collect::<Vec<_>>()
    };
    assert_eq!(
        identities(&slides[0]),
        [
            (3074, Some(0), Some("The Title".to_owned())),
            (3075, Some(1), Some("Another Subtitle".to_owned())),
        ]
    );
    let notes = slides[0].speaker_notes().unwrap().expect("notes page");
    assert_eq!(
        notes
            .shapes()
            .unwrap()
            .iter()
            .map(|shape| shape.name().map(str::to_owned))
            .collect::<Vec<_>>(),
        [
            Some("Slide Image Placeholder 1".to_owned()),
            Some("Notes Placeholder 2".to_owned()),
            Some("Slide Number Placeholder 3".to_owned()),
        ]
    );
    assert_eq!(
        slides[0].shape_by_id(3075).unwrap().unwrap().name(),
        Some("Another Subtitle")
    );
}
