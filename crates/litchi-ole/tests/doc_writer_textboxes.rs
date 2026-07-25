//! Round-trip tests for DOC floating text boxes.
//!
//! Writes a .doc containing text, pictures, a plain shape, and floating text
//! boxes, then re-opens it with the crate's own reader: the textbox story
//! (ccpTxbx + PlcftxbxTxt) and the OfficeArtClientTextbox links must resolve
//! to the right shapes and text.
#![cfg(feature = "imgconv")]

use litchi_ole::doc::shapes::extract_drawing_shapes;
use litchi_ole::doc::writer::{
    DocDrawingShape, DocPicture, DocShapeKind, DocWriter, FloatingPosition,
};
use litchi_ole::doc::{Package, ShapeHorizontalOrigin, ShapeTextWrap, ShapeVerticalOrigin};
use litchi_ole::escher::EscherShapeType;
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
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/images/jpg/abstract4.jpg");
    std::fs::read(path).expect("read JPEG fixture")
}

/// Insertion order and expected spids: group=1024, inline PNG=1025,
/// floating JPEG=1026, rectangle=1027, first text box=1028, second=1029.
const RECTANGLE_SPID: u32 = 1027;
const FIRST_BOX_SPID: u32 = 1028;
const SECOND_BOX_SPID: u32 = 1029;

fn write_doc_with_text_boxes(jpeg_bytes: &[u8]) -> Vec<u8> {
    let mut writer = DocWriter::new();
    writer.add_paragraph("before boxes").unwrap();
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
                .with_fill(0xFF, 0x00, 0x00),
            FloatingPosition::new(2000, 1000),
        )
        .unwrap();
    writer
        .insert_floating_text_box(
            DocDrawingShape::new(DocShapeKind::Rectangle, 3600, 1800)
                .unwrap()
                .with_fill(0xFF, 0xFF, 0xCC)
                .with_line(0x80, 0x00, 0x00),
            FloatingPosition::new(2500, 1200)
                .with_origins(
                    ShapeHorizontalOrigin::Page,
                    ShapeVerticalOrigin::Paragraph,
                )
                .with_text_wrap(ShapeTextWrap::Square),
            "First box",
        )
        .unwrap();
    writer
        .insert_floating_text_box(
            DocDrawingShape::new(DocShapeKind::RoundRectangle, 3600, 1800).unwrap(),
            FloatingPosition::new(5000, 3000),
            "Hello\nWorld",
        )
        .unwrap();
    writer.add_paragraph("after boxes").unwrap();

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    cursor.into_inner()
}

#[test]
fn text_boxes_round_trip_text_and_shapes() {
    let doc_bytes = write_doc_with_text_boxes(&jpeg_fixture());

    let mut package = Package::from_reader(Cursor::new(&doc_bytes)).unwrap();
    let document = package.document().unwrap();

    // Main story text is intact; the textbox story text is reachable through
    // the concatenated full text, like every other subdocument story.
    let text = document.text().unwrap();
    assert!(text.contains("before boxes"));
    assert!(text.contains("after boxes"));
    assert!(text.contains("First box"));

    // Textbox story resolves both boxes with their text.
    let boxes = document.text_boxes();
    assert_eq!(boxes.len(), 2);
    assert_eq!(boxes[0].shape_id, FIRST_BOX_SPID);
    assert_eq!(boxes[0].text, "First box\r");
    assert_eq!(boxes[1].shape_id, SECOND_BOX_SPID);
    assert_eq!(boxes[1].text, "Hello\rWorld\r");

    // Anchors: picture, rectangle, two text boxes, in document order.
    let positions = document.shape_positions();
    assert_eq!(positions.len(), 4);
    let lids: Vec<u32> = positions.iter().map(|anchor| anchor.spa.shape_id).collect();
    assert_eq!(lids, vec![1026, RECTANGLE_SPID, FIRST_BOX_SPID, SECOND_BOX_SPID]);

    // The first text box keeps its geometry, origins, wrap, and colors.
    let box_spa = &positions[2].spa;
    assert_eq!((box_spa.left, box_spa.top), (2500, 1200));
    assert_eq!((box_spa.width(), box_spa.height()), (3600, 1800));
    assert_eq!(box_spa.horizontal_origin, ShapeHorizontalOrigin::Page);
    assert_eq!(box_spa.vertical_origin, ShapeVerticalOrigin::Paragraph);
    assert_eq!(box_spa.wrap, ShapeTextWrap::Square);

    // The drawing layer reports the boxes as text boxes with their colors.
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(&doc_bytes)).unwrap();
    let shapes = extract_drawing_shapes(&mut ole).unwrap();
    assert_eq!(shapes.len(), 4);
    let first = shapes
        .iter()
        .find(|shape| shape.shape_id == FIRST_BOX_SPID)
        .expect("first text box in drawing layer");
    assert_eq!(first.shape_type, EscherShapeType::TextBox);
    assert_eq!(first.fill_color, Some((0xFF, 0xFF, 0xCC)));
    assert_eq!(first.line_color, Some((0x80, 0x00, 0x00)));
    let second = shapes
        .iter()
        .find(|shape| shape.shape_id == SECOND_BOX_SPID)
        .expect("second text box in drawing layer");
    assert_eq!(second.shape_type, EscherShapeType::TextBox);
    // The plain rectangle is still a rectangle.
    let rectangle = shapes
        .iter()
        .find(|shape| shape.shape_id == RECTANGLE_SPID)
        .unwrap();
    assert_eq!(rectangle.shape_type, EscherShapeType::Rectangle);
}

#[test]
fn text_boxes_write_valid_client_textbox_links() {
    use litchi_ole::doc::parts::fib::FileInformationBlock;
    use litchi_ole::escher::EscherRecord;

    const FIB_INDEX_DGG_INFO: usize = 50;
    const FIB_INDEX_PLCF_TXBX_TXT: usize = 56;
    const RECORD_CLIENT_TEXTBOX: u16 = 0xF00D;

    let doc_bytes = write_doc_with_text_boxes(&jpeg_fixture());
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(&doc_bytes)).unwrap();
    let word_document = ole.open_stream(&["WordDocument"]).unwrap();
    let table_stream = ole.open_stream(&["1Table"]).unwrap();
    let fib = FileInformationBlock::parse(&word_document).unwrap();

    // ccpTxbx covers the story: "First box\r\r" (11) + "Hello\rWorld\r\r" (13)
    // + story-final CR (1) = 25.
    let (story_start, story_end) = fib.get_textbox_range().unwrap();
    assert_eq!(story_end - story_start, 25);

    // PlcftxbxTxt: two real entries + one spare.
    let (plc_offset, plc_len) = fib.get_table_pointer(FIB_INDEX_PLCF_TXBX_TXT).unwrap();
    assert_eq!(plc_len as usize, 4 * 4 + 3 * 22);
    let plc = &table_stream[plc_offset as usize..(plc_offset + plc_len) as usize];
    let cps: Vec<u32> = (0..4)
        .map(|i| u32::from_le_bytes(plc[i * 4..i * 4 + 4].try_into().unwrap()))
        .collect();
    assert_eq!(cps, vec![0, 11, 24, 25]);

    // The DggInfo shapes carry ClientTextbox records with sequential TXIDs.
    let (dgg_offset, dgg_len) = fib.get_table_pointer(FIB_INDEX_DGG_INFO).unwrap();
    assert!(dgg_len > 0);
    let dgg = &table_stream[dgg_offset as usize..(dgg_offset + dgg_len) as usize];
    fn collect_txids(data: &[u8], txids: &mut Vec<u32>) {
        let mut offset = 0;
        while offset + 8 <= data.len() {
            let Ok((record, size)) = EscherRecord::parse(data, offset) else {
                break;
            };
            if record.record_type_raw == RECORD_CLIENT_TEXTBOX {
                txids.push(u32::from_le_bytes(record.data[0..4].try_into().unwrap()));
            }
            if record.version == 0xF {
                collect_txids(record.data, txids);
            }
            offset += size;
        }
    }
    let mut txids = Vec::new();
    // Top level: DggContainer, then a dgglbl byte before the DgContainer.
    let (first, first_size) = EscherRecord::parse(dgg, 0).unwrap();
    collect_txids(first.data, &mut txids);
    collect_txids(&dgg[first_size + 1..], &mut txids);
    assert_eq!(txids, vec![0x0001_0000, 0x0002_0000]);
}

#[test]
fn genuine_word_text_boxes_parse_via_crate_reader() {
    // saved-by-table.doc is a genuine Word 97 file with 16 text boxes whose
    // story lives in compressed pieces.
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/document/saved-by-table.doc");
    let mut package = Package::open(path).unwrap();
    let document = package.document().unwrap();

    let boxes = document.text_boxes();
    assert_eq!(boxes.len(), 16);
    let spids: Vec<u32> = boxes.iter().map(|b| b.shape_id).collect();
    // Lids of the 16 real FTXBXS entries (the spare is skipped).
    let expected: Vec<u32> = [0x407, 0x40b, 0x40e]
        .into_iter()
        .chain(0x412..=0x41e)
        .collect();
    assert_eq!(spids, expected);
    assert!(
        boxes.iter().all(|b| !b.text.is_empty()),
        "every genuine text box has story text"
    );
}

#[test]
fn document_without_text_boxes_has_empty_story() {
    let mut writer = DocWriter::new();
    writer.add_paragraph("no boxes").unwrap();
    writer
        .insert_floating_shape(
            DocDrawingShape::new(DocShapeKind::Ellipse, 1440, 720).unwrap(),
            FloatingPosition::new(720, 720),
        )
        .unwrap();

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let doc_bytes = cursor.into_inner();

    let mut package = Package::from_reader(Cursor::new(&doc_bytes)).unwrap();
    let document = package.document().unwrap();
    assert!(document.text_boxes().is_empty());
    assert_eq!(document.shape_positions().len(), 1);
    assert_eq!(document.text().unwrap(), "no boxes\r\u{8}\r");
}
