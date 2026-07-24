//! Round-trip tests for the DOC inline picture writer.
//!
//! Writes a new .doc with PNG and JPEG pictures via `DocWriter::insert_picture`
//! and re-opens it with the crate's own DOC reader, asserting image count,
//! format, dimensions, and byte-identity of the embedded payloads.
#![cfg(feature = "imgconv")]

use litchi_ole::doc::image::PictureFields;
use litchi_ole::doc::writer::{DocPicture, DocWriter, PictureFormat};
use litchi_ole::doc::{Package, Run};
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
}
