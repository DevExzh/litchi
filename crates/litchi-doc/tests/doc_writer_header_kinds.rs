//! Round-trip tests for even/first-page header text boxes and header pictures.
//!
//! Covers the full header kind set (odd/even/first-page), the automatic
//! section flags (DOP fFacingPages, SEP sprmSFTitlePage), and the header
//! picture pipeline (PICF block + PlcfSpaHdr + header-drawing picture frame).
use litchi_doc::shape::extract_drawing_shapes;
use litchi_doc::writer::{DocPicture, DocWriter, FloatingPosition};
use litchi_doc::writer::{Kind as DrawingKind, Shape as DrawingShape};
use litchi_doc::{
    DocHeaderKind, HeaderFooterParagraph, HeaderFooterType, Package, ShapeHorizontalOrigin,
    ShapeTextWrap, ShapeVerticalOrigin,
};
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

fn write_doc_with_even_and_first_boxes() -> Vec<u8> {
    let mut writer = DocWriter::new();
    writer.add_paragraph("Body text").unwrap();
    writer
        .set_odd_header_paragraphs(vec![HeaderFooterParagraph::plain("OddH")])
        .unwrap();
    writer
        .insert_header_text_box(
            DocHeaderKind::Even,
            DrawingShape::new(DrawingKind::Rectangle, 2000, 1000).unwrap(),
            FloatingPosition::new(1000, 500),
            "Even box",
        )
        .unwrap();
    writer
        .insert_header_text_box(
            DocHeaderKind::FirstPage,
            DrawingShape::new(DrawingKind::Ellipse, 3000, 1500).unwrap(),
            FloatingPosition::new(2000, 800)
                .with_origins(ShapeHorizontalOrigin::Page, ShapeVerticalOrigin::Paragraph)
                .with_text_wrap(ShapeTextWrap::Square),
            "First\nBox",
        )
        .unwrap();

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    cursor.into_inner()
}

#[test]
fn even_and_first_page_header_text_boxes_round_trip() {
    let doc_bytes = write_doc_with_even_and_first_boxes();
    let mut package = Package::from_reader(Cursor::new(&doc_bytes)).unwrap();
    let document = package.document().unwrap();

    assert!(document.text().unwrap().contains("Body text"));

    // Both header text boxes with their kinds, in story order.
    let boxes = document.header_text_boxes();
    assert_eq!(boxes.len(), 2);
    assert_eq!(boxes[0].shape_id, 2049);
    assert_eq!(boxes[0].text, "Even box\r");
    assert_eq!(boxes[0].header_kind, Some(HeaderFooterType::EvenPageHeader));
    assert_eq!(boxes[1].shape_id, 2050);
    assert_eq!(boxes[1].text, "First\rBox\r");
    assert_eq!(
        boxes[1].header_kind,
        Some(HeaderFooterType::FirstPageHeader)
    );

    // Header shape anchors: even story comes first (anchor at CP 0), the
    // first-page story follows the even story (3 CPs) and the odd story (6).
    let positions = document.header_shape_positions();
    assert_eq!(positions.len(), 2);
    assert_eq!(positions[0].cp, 0);
    assert_eq!(positions[0].spa.shape_id, 2049);
    assert_eq!((positions[0].spa.left, positions[0].spa.top), (1000, 500));
    assert_eq!(positions[1].cp, 9);
    assert_eq!(positions[1].spa.shape_id, 2050);
    assert_eq!(
        document.header_story_kind_at_cp(positions[0].cp),
        Some(HeaderFooterType::EvenPageHeader)
    );
    assert_eq!(
        document.header_story_kind_at_cp(positions[1].cp),
        Some(HeaderFooterType::FirstPageHeader)
    );

    // Headers themselves: odd text intact; even/first contain their anchors.
    let headers = document.headers().unwrap();
    let odd = headers
        .iter()
        .find(|h| h.header_footer_type == HeaderFooterType::OddPageHeader)
        .unwrap();
    assert_eq!(odd.text(), "OddH\r\r");
    let even = headers
        .iter()
        .find(|h| h.header_footer_type == HeaderFooterType::EvenPageHeader)
        .unwrap();
    assert_eq!(even.text(), "\u{8}\r\r");
    let first = headers
        .iter()
        .find(|h| h.header_footer_type == HeaderFooterType::FirstPageHeader)
        .unwrap();
    assert_eq!(first.text(), "\u{8}\r\r");

    // Drawing layer: rectangle and ellipse with header-cluster spids.
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(&doc_bytes)).unwrap();
    let shapes = extract_drawing_shapes(&mut ole).unwrap();
    let rectangle = shapes.iter().find(|s| s.shape_id == 2049).unwrap();
    assert_eq!(rectangle.shape_type, Kind::TextBox);
    let ellipse = shapes.iter().find(|s| s.shape_id == 2050).unwrap();
    assert_eq!(ellipse.shape_type, Kind::TextBox);

    // Main stories are untouched.
    assert!(document.text_boxes().is_empty());
    assert!(document.shape_positions().is_empty());
}

#[test]
fn even_and_first_headers_set_section_flags() {
    use litchi_doc::parts::document_properties::DocumentProperties;
    use litchi_doc::parts::fib::FileInformationBlock;

    let doc_bytes = write_doc_with_even_and_first_boxes();
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(&doc_bytes)).unwrap();
    let word_document = ole.open_stream(&["WordDocument"]).unwrap();
    let table_stream = ole.open_stream(&["1Table"]).unwrap();
    let fib = FileInformationBlock::parse(&word_document).unwrap();

    // DOP fFacingPages: enabled automatically by the even-page header.
    let dop = DocumentProperties::parse(&fib, &table_stream)
        .unwrap()
        .expect("DOP must exist");
    assert!(dop.base().facing_pages());

    // SEP sprmSFTitlePage: emitted automatically by the first-page header.
    let title_page_sprm = litchi_doc::sprm_operations::SPRM_S_F_TITLE_PAGE.to_le_bytes();
    assert!(
        word_document
            .windows(title_page_sprm.len())
            .any(|window| window == title_page_sprm),
        "SEPX must contain sprmSFTitlePage"
    );
}

#[test]
fn header_picture_round_trips_with_byte_identical_payload() {
    let png_bytes = make_png(32, 16);

    let mut writer = DocWriter::new();
    writer.add_paragraph("Body").unwrap();
    writer
        .insert_header_picture(
            DocHeaderKind::Odd,
            DocPicture::new(png_bytes.clone()).unwrap(),
            FloatingPosition::new(720, 360),
        )
        .unwrap();
    // A main-story floating picture as well, to prove coexistence.
    writer
        .insert_floating_picture(
            DocPicture::new(jpeg_fixture()).unwrap(),
            FloatingPosition::new(1440, 720),
        )
        .unwrap();

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let doc_bytes = cursor.into_inner();

    let mut package = Package::from_reader(Cursor::new(&doc_bytes)).unwrap();
    let document = package.document().unwrap();

    // Two picture runs (the main-story floating JPEG anchor and the header
    // PNG anchor, both 0x0008 characters); identify by extracted format.
    let picture_runs: Vec<_> = document
        .paragraphs()
        .unwrap()
        .into_iter()
        .flat_map(|paragraph| paragraph.runs().unwrap())
        .filter(|run| run.image().is_some())
        .collect();
    assert_eq!(picture_runs.len(), 2);
    let extracted: Vec<_> = picture_runs
        .iter()
        .map(|run| document.image_data(run.image().unwrap()).unwrap())
        .collect();
    let header_png = extracted
        .iter()
        .find(|image| image.kind() == litchi_odraw::image::Kind::Png)
        .expect("header PNG picture must be extractable");
    assert_eq!(header_png.data().unwrap(), png_bytes.as_slice());
    assert!(
        extracted
            .iter()
            .any(|image| image.kind() == litchi_odraw::image::Kind::Jpeg)
    );

    // Header shape position for the picture (spid 2049 = first header item).
    let header_positions = document.header_shape_positions();
    assert_eq!(header_positions.len(), 1);
    let spa = &header_positions[0].spa;
    assert_eq!(spa.shape_id, 2049);
    assert_eq!((spa.left, spa.top), (720, 360));
    assert_eq!((spa.width(), spa.height()), (480, 240));

    // Drawing layer: the header picture frame plus the main one.
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(&doc_bytes)).unwrap();
    let shapes = extract_drawing_shapes(&mut ole).unwrap();
    let header_pic = shapes.iter().find(|s| s.shape_id == 2049).unwrap();
    assert_eq!(header_pic.shape_type, Kind::Picture);
    assert!(
        shapes
            .iter()
            .any(|s| s.shape_type == Kind::Picture && s.shape_id == 1025)
    );

    // The main story anchor/PLC tables are unaffected.
    assert_eq!(document.shape_positions().len(), 1);
    assert!(document.header_text_boxes().is_empty());
}
