//! Round-trip tests for the DOC inline picture writer.
//!
//! Writes a new .doc with PNG and JPEG pictures via `DocWriter::insert_picture`
//! and re-opens it with the crate's own DOC reader, asserting image count,
//! format, dimensions, and byte-identity of the embedded payloads.
#![cfg(feature = "imgconv")]

use litchi_ole::doc::image::PictureFields;
use litchi_ole::doc::writer::{DocPicture, DocWriter, FloatingPosition, PictureFormat};
use litchi_ole::doc::{
    Package, Run, ShapeHorizontalOrigin, ShapeTextWrap, ShapeVerticalOrigin,
};
use litchi_imgconv::BlipType;
use std::io::{Cursor, Write};
use std::path::PathBuf;

/// CRC-32 (ISO 3309, reflected polynomial) used by PNG chunks.
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
    const PNG_COLOR_TYPE_RGB: u8 = 2;
    const PNG_BIT_DEPTH_8: u8 = 8;

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
    ihdr.extend_from_slice(&[PNG_BIT_DEPTH_8, PNG_COLOR_TYPE_RGB, 0, 0, 0]);
    chunk(b"IHDR", &ihdr);

    // Raw scanlines: one filter byte (none) plus RGB triplets per row.
    let bytes_per_pixel = 3usize;
    let mut scanlines = vec![0u8; (width as usize * bytes_per_pixel + 1) * height as usize];
    for (index, byte) in scanlines.iter_mut().enumerate() {
        *byte = (index % 251) as u8;
    }
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

fn picture_runs(document: &litchi_ole::doc::Document) -> Vec<Run> {
    document
        .paragraphs()
        .unwrap()
        .into_iter()
        .flat_map(|paragraph| paragraph.runs().unwrap())
        .filter(|run| run.image().is_some())
        .collect()
}

#[test]
fn doc_picture_new_sniffs_format_and_dimensions() {
    let png = DocPicture::new(make_png(32, 16)).unwrap();
    assert_eq!(png.format(), PictureFormat::Png);
    // 96 DPI: 32 px * 1440 twips/inch / 96 dpi = 480 twips.
    assert_eq!(png.width_twips(), 480);
    assert_eq!(png.height_twips(), 240);

    let jpeg = DocPicture::new(jpeg_fixture()).unwrap();
    assert_eq!(jpeg.format(), PictureFormat::Jpeg);
    assert!(jpeg.width_twips() > 0);
    assert!(jpeg.height_twips() > 0);
}

#[test]
fn png_and_jpeg_pictures_round_trip_through_doc_reader() {
    let png_bytes = make_png(32, 16);
    let jpeg_bytes = jpeg_fixture();
    let png_picture = DocPicture::new(png_bytes.clone()).unwrap();
    let jpeg_picture = DocPicture::new(jpeg_bytes.clone()).unwrap();
    let expected_dims = [
        (png_picture.width_twips(), png_picture.height_twips()),
        (jpeg_picture.width_twips(), jpeg_picture.height_twips()),
    ];

    let mut writer = DocWriter::new();
    writer.add_paragraph("before pictures").unwrap();
    writer.insert_picture(png_picture).unwrap();
    writer.insert_picture(jpeg_picture).unwrap();
    writer.add_paragraph("after pictures").unwrap();

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();

    let mut package = Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    let document = package.document().unwrap();

    // Surrounding text survives and the picture characters stay invisible to text().
    let text = document.text().unwrap();
    assert!(text.contains("before pictures"));
    assert!(text.contains("after pictures"));

    let runs = picture_runs(&document);
    assert_eq!(runs.len(), 2, "expected exactly two picture runs");

    let expected = [
        (BlipType::Png, png_bytes.as_slice()),
        (BlipType::Jpeg, jpeg_bytes.as_slice()),
    ];
    let data_stream = document.data_stream().expect("Data stream must exist");
    for (run, ((blip_type, payload), (width_twips, height_twips))) in
        runs.iter().zip(expected.iter().zip(expected_dims.iter()))
    {
        let image = run.image().unwrap();
        let extracted = document.image_data(image).unwrap();
        assert_eq!(extracted.blip_type(), Some(*blip_type));
        assert_eq!(
            extracted.raw_data(),
            *payload,
            "embedded payload must be byte-identical"
        );

        // PICF goal dimensions match the writer's display dimensions.
        let picf = PictureFields::try_parse(data_stream, image.pic_offset() as usize).unwrap();
        assert_eq!(picf.dxa_goal as u32, *width_twips);
        assert_eq!(picf.dya_goal as u32, *height_twips);
    }
}

#[test]
fn document_without_pictures_keeps_empty_data_stream() {
    let mut writer = DocWriter::new();
    writer.add_paragraph("plain text").unwrap();

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();

    let mut package = Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    let document = package.document().unwrap();
    assert_eq!(document.text().unwrap(), "plain text\r");
    assert!(picture_runs(&document).is_empty());
    assert!(document.shape_positions().is_empty());
}

/// Character position of the floating anchor in the shared test document:
/// "before pictures" (15) + CR (1) + inline picture paragraph (2) = 18.
const FLOATING_ANCHOR_CP: u32 = 18;

fn write_doc_with_inline_and_floating(jpeg_bytes: &[u8]) -> (Vec<u8>, u32, u32) {
    let png_picture = DocPicture::new(make_png(32, 16)).unwrap();
    let jpeg_picture = DocPicture::new(jpeg_bytes.to_vec()).unwrap();
    let floating_dims = (jpeg_picture.width_twips(), jpeg_picture.height_twips());

    let mut writer = DocWriter::new();
    writer.add_paragraph("before pictures").unwrap();
    writer.insert_picture(png_picture).unwrap();
    writer
        .insert_floating_picture(
            jpeg_picture,
            FloatingPosition::new(1440, 720)
                .with_origins(
                    ShapeHorizontalOrigin::Page,
                    ShapeVerticalOrigin::Paragraph,
                )
                .with_text_wrap(ShapeTextWrap::Square)
                .lock_anchor(true),
        )
        .unwrap();
    writer.add_paragraph("after pictures").unwrap();

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    (cursor.into_inner(), floating_dims.0, floating_dims.1)
}

#[test]
fn inline_and_floating_pictures_round_trip_through_doc_reader() {
    let jpeg_bytes = jpeg_fixture();
    let (doc_bytes, floating_width, floating_height) = write_doc_with_inline_and_floating(&jpeg_bytes);

    let mut package = Package::from_reader(Cursor::new(&doc_bytes)).unwrap();
    let document = package.document().unwrap();

    let text = document.text().unwrap();
    assert!(text.contains("before pictures"));
    assert!(text.contains("after pictures"));

    // Two picture runs: the inline 0x0001 anchor and the floating 0x0008 anchor.
    let runs = picture_runs(&document);
    assert_eq!(runs.len(), 2);
    assert_eq!(runs[0].text().unwrap(), "\u{0001}");
    assert_eq!(runs[1].text().unwrap(), "\u{0008}");

    // The floating picture's bytes and format survive the round trip.
    let floating_image = runs[1].image().unwrap();
    let extracted = document.image_data(floating_image).unwrap();
    assert_eq!(extracted.blip_type(), Some(BlipType::Jpeg));
    assert_eq!(extracted.raw_data(), jpeg_bytes.as_slice());

    // The SPA carries the anchor CP, position rectangle, origins, and wrap.
    let positions = document.shape_positions();
    assert_eq!(positions.len(), 1);
    let anchor = &positions[0];
    assert_eq!(anchor.cp, FLOATING_ANCHOR_CP);
    let spa = &anchor.spa;
    assert_eq!((spa.left, spa.top), (1440, 720));
    assert_eq!(spa.width() as u32, floating_width);
    assert_eq!(spa.height() as u32, floating_height);
    assert_eq!(spa.horizontal_origin, ShapeHorizontalOrigin::Page);
    assert_eq!(spa.vertical_origin, ShapeVerticalOrigin::Paragraph);
    assert_eq!(spa.wrap, ShapeTextWrap::Square);
    assert!(spa.anchor_locked);
    assert!(!spa.below_text);
    // Inline picture spid is 1025, so the floating shape id is 1026.
    assert_eq!(spa.shape_id, 1026);
}

#[test]
fn floating_picture_writes_valid_dgg_info() {
    use litchi_ole::doc::parts::fib::FileInformationBlock;
    use litchi_ole::escher::EscherRecord;

    const FIB_INDEX_DGG_INFO: usize = 50;
    const RECORD_DGG_CONTAINER: u16 = 0xF000;
    const RECORD_BSTORE_CONTAINER: u16 = 0xF001;
    const RECORD_DG_CONTAINER: u16 = 0xF002;
    const RECORD_SPGR_CONTAINER: u16 = 0xF003;
    const RECORD_BSE: u16 = 0xF007;
    const RECORD_SP: u16 = 0xF00A;
    const RECORD_CLIENT_ANCHOR: u16 = 0xF010;

    let jpeg_bytes = jpeg_fixture();
    let (doc_bytes, _, _) = write_doc_with_inline_and_floating(&jpeg_bytes);

    let mut ole = litchi_cfb::OleFile::open(Cursor::new(&doc_bytes)).unwrap();
    let word_document = ole.open_stream(&["WordDocument"]).unwrap();
    let table_stream = ole.open_stream(&["1Table"]).unwrap();
    let fib = FileInformationBlock::parse(&word_document).unwrap();

    // fcDggInfo must point at a non-empty OfficeArtContent.
    let (dgg_offset, dgg_len) = fib.get_table_pointer(FIB_INDEX_DGG_INFO).unwrap();
    assert!(dgg_len > 0);
    let dgg = &table_stream[dgg_offset as usize..(dgg_offset + dgg_len) as usize];

    // Top level: DggContainer, dgglbl byte, DgContainer.
    let (dgg_container, dgg_container_size) = EscherRecord::parse(dgg, 0).unwrap();
    assert_eq!(dgg_container.record_type_raw, RECORD_DGG_CONTAINER);
    let dgglbl = dgg[dgg_container_size];
    assert_eq!(dgglbl, 0, "dgglbl 0 = Main Document drawing");
    let (dg_container, _) = EscherRecord::parse(dgg, dgg_container_size + 1).unwrap();
    assert_eq!(dg_container.record_type_raw, RECORD_DG_CONTAINER);

    // The BStoreContainer holds one BSE whose embedded BLIP is the JPEG.
    let mut offset = 0;
    let mut bse_count = 0;
    while offset < dgg_container.data.len() {
        let (record, size) = EscherRecord::parse(dgg_container.data, offset).unwrap();
        if record.record_type_raw == RECORD_BSTORE_CONTAINER {
            let mut bse_offset = 0;
            while bse_offset < record.data.len() {
                let (bse, bse_size) = EscherRecord::parse(record.data, bse_offset).unwrap();
                assert_eq!(bse.record_type_raw, RECORD_BSE);
                assert!(
                    bse.data
                        .windows(jpeg_bytes.len())
                        .any(|window| window == jpeg_bytes.as_slice()),
                    "embedded BLIP must contain the original JPEG bytes"
                );
                bse_count += 1;
                bse_offset += bse_size;
            }
        }
        offset += size;
    }
    assert_eq!(bse_count, 1);

    // The drawing holds the group shape plus one picture-frame shape whose
    // spid matches the SPA lid; its ClientAnchor indexes the PlcfSpa aCP.
    fn walk_shapes(
        data: &[u8],
        spids: &mut Vec<u32>,
        client_anchors: &mut Vec<u32>,
        spgr_containers: &mut usize,
    ) {
        let mut offset = 0;
        while offset + 8 <= data.len() {
            let Ok((record, size)) = EscherRecord::parse(data, offset) else {
                break;
            };
            match record.record_type_raw {
                RECORD_SPGR_CONTAINER => *spgr_containers += 1,
                RECORD_SP => {
                    let spid = u32::from_le_bytes(record.data[0..4].try_into().unwrap());
                    spids.push(spid);
                },
                RECORD_CLIENT_ANCHOR => {
                    let index = u32::from_le_bytes(record.data[0..4].try_into().unwrap());
                    client_anchors.push(index);
                },
                _ => {},
            }
            if record.version == 0xF {
                walk_shapes(record.data, spids, client_anchors, spgr_containers);
            }
            offset += size;
        }
    }
    let mut spids = Vec::new();
    let mut client_anchors = Vec::new();
    let mut spgr_containers = 0;
    walk_shapes(dg_container.data, &mut spids, &mut client_anchors, &mut spgr_containers);

    assert_eq!(spgr_containers, 1);
    assert_eq!(spids, vec![1024, 1026], "group shape + floating picture");
    assert_eq!(client_anchors, vec![0]);
}

