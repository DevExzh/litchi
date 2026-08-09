#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::*;
use crate::{Error, ImageLimit};

fn record(version: u8, instance: u16, kind: u16, body: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(8 + body.len());
    bytes.extend_from_slice(&(u16::from(version) | (instance << 4)).to_le_bytes());
    bytes.extend_from_slice(&kind.to_le_bytes());
    bytes.extend_from_slice(&u32::try_from(body.len()).unwrap().to_le_bytes());
    bytes.extend_from_slice(body);
    bytes
}

fn png_body(two: bool, data: &[u8]) -> Vec<u8> {
    let mut body = vec![1; 16];
    if two {
        body.extend_from_slice(&[2; 16]);
    }
    body.push(0xFF);
    body.extend_from_slice(data);
    body
}

#[test]
fn parses_borrowed_two_uid_bitmap_and_retains_jpeg_flavor() {
    let bytes = record(0, 0x6E1, 0xF01E, &png_body(true, b"png"));
    let blip = Blip::parse(&bytes).expect("valid PNG BLIP");
    let Blip::Png(bitmap) = blip else {
        panic!("expected PNG")
    };
    assert_eq!(bitmap.uids().second(), Some(Uid::new([2; 16])));
    assert_eq!(bitmap.data(), b"png");
    assert_eq!(bitmap.data().as_ptr(), bytes[8 + 33..].as_ptr());

    let jpeg = record(0, 0x46A, 0xF02A, &png_body(false, b"jpeg"));
    let Blip::Jpeg(jpeg_bitmap) = Blip::parse(&jpeg).expect("valid alternate JPEG") else {
        panic!("expected JPEG")
    };
    assert_eq!(jpeg_bitmap.jpeg_flavor(), Some(JpegFlavor::Alternate));
    assert_eq!(jpeg_bitmap.record().raw_kind(), 0xF02A);
}

#[test]
fn validates_metafile_sizes_and_compression() {
    let mut body = vec![0; 16];
    body.extend_from_slice(&3u32.to_le_bytes());
    body.extend_from_slice(&[0; 24]);
    body.extend_from_slice(&3u32.to_le_bytes());
    body.push(0xFE);
    body.push(0xFE);
    body.extend_from_slice(b"wmf");
    let bytes = record(0, 0x216, 0xF01B, &body);
    let Blip::Wmf(meta) = Blip::parse(&bytes).expect("valid WMF") else {
        panic!("expected WMF")
    };
    assert_eq!(meta.header().compression, Compression::None);
    assert_eq!(meta.data(), b"wmf");

    let mut invalid = bytes;
    invalid[8 + 16 + 28] = 4;
    assert!(matches!(
        Blip::parse(&invalid),
        Err(Error::ImageSizeMismatch {
            field: "cbSave",
            ..
        })
    ));
}

#[test]
fn preserves_unknown_file_block_kinds() {
    let bytes = record(7, 0x123, 0xF020, b"future");
    let Blip::Opaque(opaque) = Blip::parse(&bytes).expect("opaque BLIP") else {
        panic!("expected opaque BLIP")
    };
    assert_eq!(opaque.raw_kind(), 0xF020);
    assert_eq!(opaque.data(), b"future");
}

#[test]
fn lazily_validates_direct_and_fbse_store_blocks() {
    let png = record(0, 0x6E0, 0xF01E, &png_body(false, b"x"));
    let mut fbse_body = Vec::new();
    fbse_body.extend_from_slice(&[Kind::Png.raw(), Kind::Pict.raw()]);
    fbse_body.extend_from_slice(&[1; 16]);
    fbse_body.extend_from_slice(&0u16.to_le_bytes());
    fbse_body.extend_from_slice(&u32::try_from(png.len()).unwrap().to_le_bytes());
    fbse_body.extend_from_slice(&1u32.to_le_bytes());
    fbse_body.extend_from_slice(&0u32.to_le_bytes());
    fbse_body.extend_from_slice(&[0; 4]);
    fbse_body.extend_from_slice(&png);
    let fbse = record(2, u16::from(Kind::Png.raw()), 0xF007, &fbse_body);

    let mut body = fbse;
    body.extend_from_slice(&png);
    let store_record = record(0x0F, 2, 0xF001, &body);
    let store = Store::parse(&store_record).expect("valid store");
    assert_eq!(store.len(), 2);
    let Some(Block::Entry(entry)) = store.get(Id::new(1).unwrap()).unwrap() else {
        panic!("expected FBSE")
    };
    assert_eq!(entry.win(), Kind::Png);
    assert_eq!(entry.mac(), Kind::Pict);
    assert!(matches!(
        store.get(Id::new(2).unwrap()).unwrap(),
        Some(Block::Blip(Blip::Png(_)))
    ));
}

#[test]
fn resolves_delay_offset_zero_and_rejects_missing_context() {
    let png = record(0, 0x6E0, 0xF01E, &png_body(false, b"x"));
    let mut fbse_body = Vec::new();
    fbse_body.extend_from_slice(&[Kind::Png.raw(), Kind::Png.raw()]);
    fbse_body.extend_from_slice(&[1; 16]);
    fbse_body.extend_from_slice(&0u16.to_le_bytes());
    fbse_body.extend_from_slice(&u32::try_from(png.len()).unwrap().to_le_bytes());
    fbse_body.extend_from_slice(&1u32.to_le_bytes());
    fbse_body.extend_from_slice(&0u32.to_le_bytes());
    fbse_body.extend_from_slice(&[0; 4]);
    let fbse = record(2, u16::from(Kind::Png.raw()), 0xF007, &fbse_body);
    let store_record = record(0x0F, 1, 0xF001, &fbse);
    let store = Store::parse(&store_record).unwrap();
    let id = Id::new(1).unwrap();
    assert_eq!(
        store.resolve(id, Context::new()).unwrap_err(),
        Error::MissingDelayStore
    );
    assert!(matches!(
        store
            .resolve(id, Context::new().with_delay(Delay::new(&png)))
            .unwrap(),
        Some(Blip::Png(_))
    ));
    let mut wrong_uid = png;
    wrong_uid[8] ^= 0xFF;
    assert!(matches!(
        store.resolve(id, Context::new().with_delay(Delay::new(&wrong_uid))),
        Err(Error::MalformedImage { .. })
    ));
}

#[test]
fn detects_store_count_mismatch_on_iteration() {
    let bytes = record(0x0F, 1, 0xF001, &[]);
    assert!(matches!(
        Store::parse(&bytes),
        Err(Error::MalformedImage { .. })
    ));
}

#[test]
fn bounds_headerless_delay_block_count_and_fuses_after_error() {
    let png = record(0, 0x6E0, 0xF01E, &png_body(false, b"x"));
    let mut bytes = png.clone();
    bytes.extend_from_slice(&png);
    let mut blocks = Delay::with_limits(
        &bytes,
        Limits {
            max_store_entries: 1,
            ..Limits::default()
        },
    )
    .iter();

    assert!(blocks.next().is_some_and(|block| block.is_ok()));
    assert!(matches!(
        blocks.next(),
        Some(Err(Error::ImageLimitExceeded {
            limit: ImageLimit::StoreEntries,
            maximum: 1,
        }))
    ));
    assert!(blocks.next().is_none());
}
