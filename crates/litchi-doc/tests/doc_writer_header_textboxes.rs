//! Round-trip tests for DOC header-story text boxes (letterhead/watermark).
//!
//! Writes a .doc with main-story content plus a header containing text and a
//! floating text box, then re-opens it with the crate's own reader: header
//! text, header text box text/shape metadata, and both drawing layers must
//! resolve correctly without disturbing the main stories.
use litchi_doc::parts::headers::HeaderFooterType;
use litchi_doc::shape::{extract_drawing_shapes, extract_shape_text};
use litchi_doc::writer::{FloatingPosition, Writer};
use litchi_doc::writer::{Kind as DrawingKind, Shape as DrawingShape};
use litchi_doc::{
    HeaderFooterParagraph, HeaderKind, Package, ShapeHorizontalOrigin, ShapeTextWrap,
    ShapeVerticalOrigin,
};
use litchi_odraw::{Record, shape::Kind};
use std::io::Cursor;

/// The first header-story shape id: header shapes use their own cluster
/// starting at 2049 (the Main Document cluster is 1024-based).
const HEADER_BOX_SPID: u32 = 2049;
/// Main-story shape id for the coexisting main text box.
const MAIN_BOX_SPID: u32 = 1025;

fn write_doc_with_header_text_box() -> Vec<u8> {
    let mut writer = Writer::new();
    writer.add_paragraph("Body text").unwrap();
    writer
        .set_odd_header_paragraphs(vec![HeaderFooterParagraph::plain("LETTERHEAD")])
        .unwrap();
    // Main-story text box for coexistence.
    writer
        .insert_floating_text_box(
            DrawingShape::new(DrawingKind::Rectangle, 2000, 1000).unwrap(),
            FloatingPosition::new(1440, 1440),
            "Main box",
        )
        .unwrap();
    // Header text box (the watermark).
    writer
        .insert_header_text_box(
            HeaderKind::Odd,
            DrawingShape::new(DrawingKind::Rectangle, 4000, 2000)
                .unwrap()
                .with_fill(0xEE, 0xEE, 0xFF)
                .with_line(0x40, 0x40, 0x40),
            FloatingPosition::new(1000, 500)
                .with_origins(ShapeHorizontalOrigin::Page, ShapeVerticalOrigin::Paragraph)
                .with_text_wrap(ShapeTextWrap::None)
                .behind_text(true),
            "Watermark\nDraft",
        )
        .unwrap();
    writer.add_paragraph("More body").unwrap();

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    cursor.into_inner()
}

#[test]
fn header_text_box_round_trips_through_doc_reader() {
    let doc_bytes = write_doc_with_header_text_box();

    let mut package = Package::from_reader(Cursor::new(&doc_bytes)).unwrap();
    let document = package.document().unwrap();

    // Main story intact.
    let text = document.text().unwrap();
    assert!(text.contains("Body text"));
    assert!(text.contains("More body"));

    // Header text via the existing header API, with the 0x0008 anchor.
    let headers = document.headers().unwrap();
    let odd_header = headers
        .iter()
        .find(|header| header.header_footer_type == HeaderFooterType::OddPageHeader)
        .expect("odd header must exist");
    // Header text via the existing header API: the anchor paragraph plus the
    // guard paragraph mark the writer emits between header stories.
    assert_eq!(odd_header.text(), "LETTERHEAD\r\u{8}\r\r");

    // Header text box text and metadata.
    let header_boxes = document.header_text_boxes();
    assert_eq!(header_boxes.len(), 1);
    assert_eq!(header_boxes[0].shape_id, HEADER_BOX_SPID);
    assert_eq!(header_boxes[0].text, "Watermark\rDraft\r");

    // The main-story text box is unaffected.
    let main_boxes = document.text_boxes();
    assert_eq!(main_boxes.len(), 1);
    assert_eq!(main_boxes[0].shape_id, MAIN_BOX_SPID);
    assert_eq!(main_boxes[0].text, "Main box\r");

    // Header shape position: the anchor is the second odd-header paragraph,
    // after "LETTERHEAD\r" (11 story CPs).
    let header_positions = document.header_shape_positions();
    assert_eq!(header_positions.len(), 1);
    let anchor = &header_positions[0];
    assert_eq!(anchor.cp, 11);
    let spa = &anchor.spa;
    assert_eq!(spa.shape_id, HEADER_BOX_SPID);
    assert_eq!((spa.left, spa.top), (1000, 500));
    assert_eq!((spa.width(), spa.height()), (4000, 2000));
    assert_eq!(spa.horizontal_origin, ShapeHorizontalOrigin::Page);
    assert_eq!(spa.vertical_origin, ShapeVerticalOrigin::Paragraph);
    assert_eq!(spa.wrap, ShapeTextWrap::None);
    assert!(spa.below_text);

    // Main shape positions only list the main text box.
    let main_positions = document.shape_positions();
    assert_eq!(main_positions.len(), 1);
    assert_eq!(main_positions[0].spa.shape_id, MAIN_BOX_SPID);

    // Both drawing layers extract: the header text box and the main one.
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(&doc_bytes)).unwrap();
    let shapes = extract_drawing_shapes(&mut ole).unwrap();
    let header_box = shapes
        .iter()
        .find(|shape| shape.shape_id == HEADER_BOX_SPID)
        .expect("header text box in drawing layer");
    assert_eq!(header_box.shape_type, Kind::TextBox);
    assert_eq!(header_box.text.as_deref(), Some("Watermark\rDraft\r"));
    assert_eq!(header_box.fill_color, Some((0xEE, 0xEE, 0xFF)));
    assert_eq!(header_box.line_color, Some((0x40, 0x40, 0x40)));
    let main_box = shapes
        .iter()
        .find(|shape| shape.shape_id == MAIN_BOX_SPID)
        .expect("main text box in drawing layer");
    assert_eq!(main_box.shape_type, Kind::TextBox);
    assert_eq!(main_box.text.as_deref(), Some("Main box\r"));

    let mut ole = litchi_cfb::OleFile::open(Cursor::new(&doc_bytes)).unwrap();
    assert_eq!(
        extract_shape_text(&mut ole).unwrap(),
        "Main box\r\nWatermark\rDraft\r"
    );
}

#[test]
fn header_text_box_writes_two_drawing_layers() {
    use litchi_doc::parts::fib::FileInformationBlock;
    const FIB_INDEX_DGG_INFO: usize = 50;
    const RECORD_DG_CONTAINER: u16 = 0xF002;

    let doc_bytes = write_doc_with_header_text_box();
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(&doc_bytes)).unwrap();
    let word_document = ole.open_stream(&["WordDocument"]).unwrap();
    let table_stream = ole.open_stream(&["1Table"]).unwrap();
    let fib = FileInformationBlock::parse(&word_document).unwrap();

    // ccpHdrTxbx covers the story: "Watermark\rDraft\r\r" (17) + final CR = 18.
    let (story_start, story_end) = fib.get_header_textbox_range().unwrap();
    assert_eq!(story_end - story_start, 18);

    // DggInfo contains two OfficeArtWordDrawing elements with dgglbl 0 and 1.
    let (dgg_offset, dgg_len) = fib.get_table_pointer(FIB_INDEX_DGG_INFO).unwrap();
    assert!(dgg_len > 0);
    let dgg = &table_stream[dgg_offset as usize..(dgg_offset + dgg_len) as usize];
    let (_first, first_size) = Record::parse(dgg, 0).unwrap();
    let mut dgglbls = Vec::new();
    let mut offset = first_size;
    while offset + 9 <= dgg.len() {
        dgglbls.push(dgg[offset]);
        let Ok((record, size)) = Record::parse(dgg, offset + 1) else {
            break;
        };
        if record.raw_kind() != RECORD_DG_CONTAINER {
            break;
        }
        offset += 1 + size;
    }
    assert_eq!(dgglbls, vec![0, 1]);
}

#[test]
fn document_with_header_but_no_boxes_has_empty_header_story_tables() {
    let mut writer = Writer::new();
    writer.add_paragraph("plain").unwrap();
    writer.set_odd_header("just a header");

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let doc_bytes = cursor.into_inner();

    let mut package = Package::from_reader(Cursor::new(&doc_bytes)).unwrap();
    let document = package.document().unwrap();
    assert!(document.header_text_boxes().is_empty());
    assert!(document.header_shape_positions().is_empty());
    assert!(document.text_boxes().is_empty());
}
