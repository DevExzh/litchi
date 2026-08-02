//! Round-trip tests for the DOC primitive drawing-shape writer.
//!
//! Writes a .doc containing text, an inline picture, a floating picture, and
//! floating primitive shapes (rectangle, ellipse, rounded rectangle), then
//! re-opens it with the crate's own reader and shape extraction APIs.
use litchi_doc::shapes::extract_drawing_shapes;
use litchi_doc::writer::{DocDrawingShape, DocPicture, DocShapeKind, DocWriter, FloatingPosition};
use litchi_doc::{Package, ShapeHorizontalOrigin, ShapeTextWrap, ShapeVerticalOrigin};
use litchi_odraw::shape::Kind;
use std::io::{Cursor, Write};
use std::path::PathBuf;

const CRC32_POLYNOMIAL: u32 = 0xEDB8_8320;

fn crc32(bytes: &[u8]) -> u32 {
    let mut crc = u32::MAX;
    for &byte in bytes {
        crc ^= u32::from(byte);
        for _ in 0..8 {
            crc = if crc & 1 != 0 {
                (crc >> 1) ^ CRC32_POLYNOMIAL
            } else {
                crc >> 1
            };
        }
    }
    !crc
}

/// Build a minimal but fully valid RGB PNG of the given pixel dimensions.
fn make_png(width: u32, height: u32) -> Vec<u8> {
    let mut png = vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
    let mut chunk = |chunk_type: &[u8; 4], payload: &[u8]| {
        png.extend_from_slice(&(payload.len() as u32).to_be_bytes());
        png.extend_from_slice(chunk_type);
        png.extend_from_slice(payload);
        let mut crc_input = Vec::with_capacity(4 + payload.len());
        crc_input.extend_from_slice(chunk_type);
        crc_input.extend_from_slice(payload);
        png.extend_from_slice(&crc32(&crc_input).to_be_bytes());
    };
    let mut ihdr = Vec::with_capacity(13);
    ihdr.extend_from_slice(&width.to_be_bytes());
    ihdr.extend_from_slice(&height.to_be_bytes());
    ihdr.extend_from_slice(&[8, 2, 0, 0, 0]);
    chunk(b"IHDR", &ihdr);
    let scanlines = vec![0u8; (width as usize * 3 + 1) * height as usize];
    let mut encoder = flate2::write::ZlibEncoder::new(Vec::new(), flate2::Compression::fast());
    encoder.write_all(&scanlines).unwrap();
    chunk(b"IDAT", &encoder.finish().unwrap());
    chunk(b"IEND", &[]);
    png
}

fn jpeg_fixture() -> Vec<u8> {
    let path =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/images/jpg/abstract4.jpg");
    std::fs::read(path).expect("read JPEG fixture")
}

/// Insertion order and expected shape ids:
/// group=1024, inline PNG=1025, floating JPEG=1026, rectangle=1027,
/// ellipse=1028, rounded rectangle=1029.
const RECTANGLE_SPID: u32 = 1027;
const ELLIPSE_SPID: u32 = 1028;
const ROUND_RECT_SPID: u32 = 1029;

fn write_doc_with_shapes(jpeg_bytes: &[u8]) -> Vec<u8> {
    let mut writer = DocWriter::new();
    writer.add_paragraph("before shapes").unwrap();
    writer
        .insert_picture(DocPicture::new(make_png(32, 16)).unwrap())
        .unwrap();
    writer
        .insert_floating_picture(
            DocPicture::new(jpeg_bytes.to_vec()).unwrap(),
            FloatingPosition::new(1440, 720),
        )
        .unwrap();
    writer
        .insert_floating_shape(
            DocDrawingShape::new(DocShapeKind::Rectangle, 2880, 1440)
                .unwrap()
                .with_fill(0xFF, 0x00, 0x00)
                .with_line(0x00, 0x00, 0xFF),
            FloatingPosition::new(2000, 1000)
                .with_origins(ShapeHorizontalOrigin::Page, ShapeVerticalOrigin::Paragraph)
                .with_text_wrap(ShapeTextWrap::Square),
        )
        .unwrap();
    writer
        .insert_floating_shape(
            DocDrawingShape::new(DocShapeKind::Ellipse, 1440, 1440)
                .unwrap()
                .with_line(0x00, 0x80, 0x00),
            FloatingPosition::new(4000, 2000),
        )
        .unwrap();
    writer
        .insert_floating_shape(
            DocDrawingShape::new(DocShapeKind::RoundRectangle, 2160, 1080).unwrap(),
            FloatingPosition::new(6000, 3000).with_text_wrap(ShapeTextWrap::TopAndBottom),
        )
        .unwrap();
    writer.add_paragraph("after shapes").unwrap();

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    cursor.into_inner()
}

#[test]
fn primitive_shapes_round_trip_through_doc_reader() {
    let doc_bytes = write_doc_with_shapes(&jpeg_fixture());

    let mut package = Package::from_reader(Cursor::new(&doc_bytes)).unwrap();
    let document = package.document().unwrap();

    // Text and both pictures survive alongside the shapes.
    let text = document.text().unwrap();
    assert!(text.contains("before shapes"));
    assert!(text.contains("after shapes"));
    let picture_runs: Vec<_> = document
        .paragraphs()
        .unwrap()
        .into_iter()
        .flat_map(|paragraph| paragraph.runs().unwrap())
        .filter(|run| run.image().is_some())
        .collect();
    assert_eq!(picture_runs.len(), 2);
    let floating_image = picture_runs[1].image().unwrap();
    assert_eq!(
        document.image_data(floating_image).unwrap().data().unwrap(),
        jpeg_fixture().as_slice()
    );

    // Four floating anchors: the picture plus three shapes, in document order.
    let positions = document.shape_positions();
    assert_eq!(positions.len(), 4);
    let cps: Vec<u32> = positions.iter().map(|anchor| anchor.cp).collect();
    assert!(cps.windows(2).all(|pair| pair[0] < pair[1]));
    // "before shapes" (13) + CR (1) + inline paragraph (2) = 16; each floating
    // anchor paragraph adds 2 CPs.
    assert_eq!(cps, vec![16, 18, 20, 22]);

    let rectangle = &positions[1].spa;
    assert_eq!(rectangle.shape_id, RECTANGLE_SPID);
    assert_eq!((rectangle.left, rectangle.top), (2000, 1000));
    assert_eq!((rectangle.width(), rectangle.height()), (2880, 1440));
    assert_eq!(rectangle.horizontal_origin, ShapeHorizontalOrigin::Page);
    assert_eq!(rectangle.vertical_origin, ShapeVerticalOrigin::Paragraph);
    assert_eq!(rectangle.wrap, ShapeTextWrap::Square);

    let ellipse = &positions[2].spa;
    assert_eq!(ellipse.shape_id, ELLIPSE_SPID);
    assert_eq!((ellipse.left, ellipse.top), (4000, 2000));
    assert_eq!((ellipse.width(), ellipse.height()), (1440, 1440));

    let round_rect = &positions[3].spa;
    assert_eq!(round_rect.shape_id, ROUND_RECT_SPID);
    assert_eq!(round_rect.wrap, ShapeTextWrap::TopAndBottom);

    // Shape ids are unique across the whole document.
    let mut lids: Vec<u32> = positions.iter().map(|anchor| anchor.spa.shape_id).collect();
    lids.sort_unstable();
    lids.dedup();
    assert_eq!(lids.len(), 4);
}

#[test]
fn primitive_shapes_extract_with_type_geometry_and_colors() {
    let doc_bytes = write_doc_with_shapes(&jpeg_fixture());

    let mut ole = litchi_cfb::OleFile::open(Cursor::new(&doc_bytes)).unwrap();
    let shapes = extract_drawing_shapes(&mut ole).unwrap();

    // The drawing group exposes the floating picture frame and the three
    // primitive shapes (the group header is not a user shape).
    assert_eq!(shapes.len(), 4);

    let picture = shapes
        .iter()
        .find(|shape| shape.shape_type == Kind::Picture)
        .expect("floating picture frame must be extracted");
    assert_eq!(picture.shape_id, 1026);

    let rectangle = shapes
        .iter()
        .find(|shape| shape.shape_id == RECTANGLE_SPID)
        .expect("rectangle must be extracted");
    assert_eq!(rectangle.shape_type, Kind::Rectangle);
    assert_eq!(rectangle.native_shape_type, Some(0x0001));
    assert_eq!(rectangle.fill_color, Some((0xFF, 0x00, 0x00)));
    assert_eq!(rectangle.line_color, Some((0x00, 0x00, 0xFF)));

    let ellipse = shapes
        .iter()
        .find(|shape| shape.shape_id == ELLIPSE_SPID)
        .expect("ellipse must be extracted");
    assert_eq!(ellipse.shape_type, Kind::Ellipse);
    assert_eq!(ellipse.native_shape_type, Some(0x0003));
    assert_eq!(ellipse.fill_color, None, "ellipse has no fill");
    assert_eq!(ellipse.line_color, Some((0x00, 0x80, 0x00)));

    let round_rect = shapes
        .iter()
        .find(|shape| shape.shape_id == ROUND_RECT_SPID)
        .expect("rounded rectangle must be extracted");
    assert_eq!(round_rect.native_shape_type, Some(0x0002));

    // All extracted shape ids are unique and collide with no picture spid.
    let mut spids: Vec<u32> = shapes.iter().map(|shape| shape.shape_id).collect();
    spids.sort_unstable();
    spids.dedup();
    assert_eq!(spids.len(), shapes.len());
}

#[test]
fn genuine_word_floating_pictures_extract_from_dgg_info() {
    // FloatingPictures.doc is a genuine Word 97 file whose floating pictures
    // live in the table-stream OfficeArtContent (fcDggInfo).
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/doc/FloatingPictures.doc");
    let mut ole = litchi_cfb::OleFile::open(std::fs::File::open(path).unwrap()).unwrap();
    let shapes = extract_drawing_shapes(&mut ole).unwrap();

    assert_eq!(shapes.len(), 4);
    let mut spids: Vec<u32> = shapes.iter().map(|shape| shape.shape_id).collect();
    spids.sort_unstable();
    assert_eq!(spids, vec![1028, 1029, 1030, 1031]);
    assert!(shapes.iter().all(|shape| shape.shape_type == Kind::Picture));
}

#[test]
fn document_with_only_shapes_has_no_picture_data() {
    let mut writer = DocWriter::new();
    writer.add_paragraph("shapes only").unwrap();
    writer
        .insert_floating_shape(
            DocDrawingShape::new(DocShapeKind::Rectangle, 1440, 720).unwrap(),
            FloatingPosition::new(720, 720),
        )
        .unwrap();

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let doc_bytes = cursor.into_inner();

    let mut package = Package::from_reader(Cursor::new(&doc_bytes)).unwrap();
    let document = package.document().unwrap();
    assert_eq!(document.text().unwrap(), "shapes only\r\u{8}\r");
    assert_eq!(document.shape_positions().len(), 1);

    let mut ole = litchi_cfb::OleFile::open(Cursor::new(&doc_bytes)).unwrap();
    let shapes = extract_drawing_shapes(&mut ole).unwrap();
    assert_eq!(shapes.len(), 1);
    assert_eq!(shapes[0].shape_type, Kind::Rectangle);
    // Default style: no fill, black line.
    assert_eq!(shapes[0].fill_color, None);
    assert_eq!(shapes[0].line_color, Some((0, 0, 0)));
}
