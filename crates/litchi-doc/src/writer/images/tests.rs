//! Focused DOC image writer layout and validation tests.

use super::codec::*;
use super::model::*;
use super::validation::*;
use litchi_odraw::image::Kind as BlipKind;
use litchi_odraw::image::write::digest;

/// Minimal byte sequences recognised by the format sniffer.
fn png_bytes() -> Vec<u8> {
    let mut data = vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
    data.extend_from_slice(&PNG_IHDR_LEN.to_be_bytes());
    data.extend_from_slice(b"IHDR");
    data.extend_from_slice(&32u32.to_be_bytes()); // width
    data.extend_from_slice(&16u32.to_be_bytes()); // height
    data.extend_from_slice(&[8, 2, 0, 0, 0]); // bit depth, RGB, compression, filter, interlace
    data.extend_from_slice(&[0; 4]); // CRC (not validated here)
    data
}

fn jpeg_bytes() -> Vec<u8> {
    let mut data = vec![0xFF, 0xD8, 0xFF, 0xE0];
    data.extend_from_slice(&16u16.to_be_bytes()); // APP0 length
    data.extend_from_slice(b"JFIF\0");
    data.extend_from_slice(&[0; 9]); // rest of APP0
    data.extend_from_slice(&[0xFF, 0xC0]);
    data.extend_from_slice(&17u16.to_be_bytes()); // SOF0 length
    data.push(8); // precision
    data.extend_from_slice(&24u16.to_be_bytes()); // height
    data.extend_from_slice(&48u16.to_be_bytes()); // width
    data
}

#[test]
fn sniff_kind_detects_png_and_jpeg() {
    assert_eq!(sniff_kind(&png_bytes()).unwrap(), BlipKind::Png);
    assert_eq!(sniff_kind(&jpeg_bytes()).unwrap(), BlipKind::Jpeg);
    assert!(sniff_kind(b"not an image").is_err());
}

#[test]
fn png_pixel_dimensions_reads_ihdr() {
    assert_eq!(png_pixel_dimensions(&png_bytes()), Some((32, 16)));
    assert_eq!(png_pixel_dimensions(&png_bytes()[..20]), None);
}

#[test]
fn jpeg_pixel_dimensions_reads_sof0() {
    assert_eq!(jpeg_pixel_dimensions(&jpeg_bytes()), Some((48, 24)));
    assert_eq!(jpeg_pixel_dimensions(&jpeg_bytes()[..10]), None);
}

#[test]
fn doc_picture_new_derives_dimensions_at_96_dpi() {
    let picture = Picture::new(png_bytes()).unwrap();
    assert_eq!(picture.kind(), BlipKind::Png);
    assert_eq!(picture.width_twips(), 32 * TWIPS_PER_INCH / ASSUMED_DPI);
    assert_eq!(picture.height_twips(), 16 * TWIPS_PER_INCH / ASSUMED_DPI);
}

#[test]
fn doc_picture_rejects_unknown_format_and_bad_dimensions() {
    assert!(Picture::new(b"garbage".to_vec()).is_err());
    assert!(Picture::from_parts(png_bytes(), 0, 100).is_err());
    assert!(Picture::from_parts(png_bytes(), 100, MAX_PICF_DIMENSION_TWIPS + 1).is_err());
}

/// Parse an Escher record header, returning (version, instance, type, payload).
fn parse_record_header(block: &[u8], offset: usize) -> (u16, u16, u16, u32) {
    let ver_inst = u16::from_le_bytes(block[offset..offset + 2].try_into().unwrap());
    let record_type = u16::from_le_bytes(block[offset + 2..offset + 4].try_into().unwrap());
    let length = u32::from_le_bytes(block[offset + 4..offset + 8].try_into().unwrap());
    (ver_inst & 0xF, ver_inst >> 4, record_type, length)
}

#[test]
fn picture_block_layout_matches_ms_doc() {
    let picture = Picture::new(png_bytes()).unwrap();
    let mut block = Vec::new();
    write_picture_block(&picture, FIRST_PICTURE_SHAPE_ID, &mut block).unwrap();

    // PICF header.
    let lcb = i32::from_le_bytes(block[0..4].try_into().unwrap()) as usize;
    assert_eq!(lcb, block.len());
    assert_eq!(
        i16::from_le_bytes(block[4..6].try_into().unwrap()),
        PICF_CB_HEADER
    );
    assert_eq!(
        i16::from_le_bytes(block[6..8].try_into().unwrap()),
        PICF_MM_SHAPE
    );
    // Block type byte (grf low byte) 0x00 + mm 0x64 is what the reader
    // recognises as a Word 2000 picture.
    assert_eq!(block[0x0E], 0);
    let dxa_goal = i16::from_le_bytes(block[0x1C..0x1E].try_into().unwrap());
    let dya_goal = i16::from_le_bytes(block[0x1E..0x20].try_into().unwrap());
    assert_eq!(dxa_goal as u32, picture.width_twips());
    assert_eq!(dya_goal as u32, picture.height_twips());

    // SpContainer follows the PICF.
    let (ver, _inst, record_type, len) = parse_record_header(&block, PICF_HEADER_LEN);
    assert_eq!((ver, record_type), (VERSION_CONTAINER, RECORD_SP_CONTAINER));
    let bse_offset = PICF_HEADER_LEN + RECORD_HEADER_LEN + len as usize;

    // BSE record with the embedded BLIP.
    let (ver, inst, record_type, len) = parse_record_header(&block, bse_offset);
    assert_eq!((ver, record_type), (2, RECORD_BSE));
    assert_eq!(inst, u16::from(BlipKind::Png.raw()));
    let bse = &block[bse_offset + RECORD_HEADER_LEN..];
    assert_eq!(bse[0], BlipKind::Png.raw()); // btWin32
    assert_eq!(bse[1], BlipKind::Png.raw()); // btMacOS
    let blip_size = u32::from_le_bytes(bse[20..24].try_into().unwrap()) as usize;
    let c_ref = u32::from_le_bytes(bse[24..28].try_into().unwrap());
    let fo_delay = u32::from_le_bytes(bse[28..32].try_into().unwrap());
    assert_eq!(c_ref, 1);
    assert_eq!(fo_delay, BSE_NO_DELAY_STREAM); // embedded
    assert_eq!(blip_size + BSE_HEADER_LEN, len as usize);

    // Embedded BLIP record.
    let blip = &bse[BSE_HEADER_LEN..BSE_HEADER_LEN + blip_size];
    let parsed = litchi_odraw::image::Blip::parse(blip).unwrap();
    assert_eq!(parsed.kind(), BlipKind::Png);
    assert_eq!(parsed.record().instance(), 0x6e0);
    // The BLIP UID matches the BSE rgbUid, and the payload is byte-identical.
    assert_eq!(parsed.uids().unwrap().effective(), digest(picture.data()));
    assert_eq!(parsed.data(), picture.data());
}

#[test]
fn picture_uid_is_the_md4_digest_of_blip_data() {
    let picture = Picture::new(png_bytes()).unwrap();
    assert_eq!(
        digest(picture.data()),
        picture_blip(&picture).unwrap().uid()
    );
    // RFC 1320 test vector.
    assert_eq!(
        digest(b"abc").bytes(),
        [
            0xa4, 0x48, 0x01, 0x7a, 0xaf, 0x21, 0xd8, 0x52, 0x5f, 0xc1, 0x0a, 0xe8, 0x7a, 0xa6,
            0x72, 0x9d,
        ]
    );
    assert_eq!(
        digest(b"12345678901234567890123456789012345678901234567890123456789012345678901234567890")
            .bytes(),
        [
            0xe3, 0x3b, 0x4d, 0xdc, 0x9c, 0x38, 0xf2, 0x19, 0x9c, 0x3e, 0x7b, 0x16, 0x4f, 0xcc,
            0x05, 0x36,
        ]
    );
}

#[test]
fn picture_block_bse_parses_with_crate_reader() {
    use litchi_odraw::Record;
    use litchi_odraw::image::{Blip, Context, Entry, Storage};

    let picture = Picture::new(jpeg_bytes()).unwrap();
    let mut block = Vec::new();
    write_picture_block(&picture, FIRST_PICTURE_SHAPE_ID, &mut block).unwrap();

    let (_ver, _inst, _record_type, sp_len) = parse_record_header(&block, PICF_HEADER_LEN);
    let bse_offset = PICF_HEADER_LEN + RECORD_HEADER_LEN + sp_len as usize;
    let (record, _) = Record::parse(&block, bse_offset).unwrap();
    let bse = Entry::parse(record).unwrap();
    assert_eq!(bse.kind().unwrap(), BlipKind::Jpeg);
    assert!(matches!(
        bse.storage().unwrap(),
        Storage::Embedded(Blip::Jpeg(_))
    ));
    let blip = bse.resolve(Context::new()).unwrap().unwrap();
    assert_eq!(blip.data(), picture.data());
}

// ── Floating pictures ──

use crate::parts::spa::{ShapeHorizontalOrigin, ShapeTextWrap, ShapeVerticalOrigin, ShapeWrapSide};

fn floating_shapes<'a>(
    png: &'a Picture,
    jpeg: &'a Picture,
    positions: &'a [FloatingPosition],
) -> Vec<FloatingShapeInfo<'a>> {
    vec![
        FloatingShapeInfo {
            anchor_cp: 12,
            shape_id: FIRST_PICTURE_SHAPE_ID,
            content: FloatingShapeContent::Picture(png),
            width_twips: png.width_twips(),
            height_twips: png.height_twips(),
            position: &positions[0],
            text: None,
        },
        FloatingShapeInfo {
            anchor_cp: 30,
            shape_id: FIRST_PICTURE_SHAPE_ID + 1,
            content: FloatingShapeContent::Picture(jpeg),
            width_twips: jpeg.width_twips(),
            height_twips: jpeg.height_twips(),
            position: &positions[1],
            text: None,
        },
    ]
}

fn sample_positions() -> [FloatingPosition; 2] {
    [
        FloatingPosition::new(1440, 720)
            .with_origins(ShapeHorizontalOrigin::Page, ShapeVerticalOrigin::Paragraph)
            .with_text_wrap(ShapeTextWrap::Square)
            .lock_anchor(true),
        FloatingPosition::new(2880, 1440)
            .with_text_wrap(ShapeTextWrap::None)
            .behind_text(true),
    ]
}

#[test]
fn floating_position_builder_defaults() {
    let position = FloatingPosition::new(100, 200);
    assert_eq!(position.horizontal_origin, ShapeHorizontalOrigin::Page);
    assert_eq!(position.vertical_origin, ShapeVerticalOrigin::Page);
    assert_eq!(position.wrap, ShapeTextWrap::Square);
    assert_eq!(position.wrap_side, ShapeWrapSide::Both);
    assert!(!position.behind_text);
    assert!(!position.anchor_locked);
}

#[test]
fn plcf_spa_layout_matches_ms_doc() {
    let png = Picture::new(png_bytes()).unwrap();
    let jpeg = Picture::new(jpeg_bytes()).unwrap();
    let positions = sample_positions();
    let shapes = floating_shapes(&png, &jpeg, &positions);

    let plcf = build_plcf_spa(&shapes, 500);
    let anchors = crate::parts::spa::parse_plcf_spa(&plcf).unwrap();
    assert_eq!(anchors.len(), 2);
    assert_eq!(anchors[0].cp, 12);
    assert_eq!(anchors[1].cp, 30);

    let first = &anchors[0].spa;
    assert_eq!(first.shape_id, FIRST_PICTURE_SHAPE_ID);
    assert_eq!((first.left, first.top), (1440, 720));
    assert_eq!(first.width() as u32, png.width_twips());
    assert_eq!(first.height() as u32, png.height_twips());
    assert_eq!(first.horizontal_origin, ShapeHorizontalOrigin::Page);
    assert_eq!(first.vertical_origin, ShapeVerticalOrigin::Paragraph);
    assert_eq!(first.wrap, ShapeTextWrap::Square);
    assert!(first.anchor_locked);

    let second = &anchors[1].spa;
    assert_eq!(second.shape_id, FIRST_PICTURE_SHAPE_ID + 1);
    assert_eq!(second.wrap, ShapeTextWrap::None);
    assert!(second.below_text);

    // The final CP is ccpText and exceeds every anchor CP.
    let final_cp = u32::from_le_bytes(plcf[8..12].try_into().unwrap());
    assert_eq!(final_cp, 500);
}

/// Collect all Escher records in a byte slice as (offset, ver, instance, type, len).
/// Handles the OfficeArtWordDrawing dgglbl byte between top-level records.
fn collect_records(data: &[u8], start: usize, end: usize) -> Vec<(usize, u16, u16, u16, u32)> {
    fn walk(data: &[u8], start: usize, end: usize, records: &mut Vec<(usize, u16, u16, u16, u32)>) {
        let mut offset = start;
        while offset + RECORD_HEADER_LEN <= end {
            let (ver, inst, record_type, len) = parse_record_header(data, offset);
            records.push((offset, ver, inst, record_type, len));
            if ver == VERSION_CONTAINER {
                walk(
                    data,
                    offset + RECORD_HEADER_LEN,
                    offset + RECORD_HEADER_LEN + len as usize,
                    records,
                );
            }
            offset += RECORD_HEADER_LEN + len as usize;
        }
    }

    let mut records = Vec::new();
    let (_ver, _inst, _record_type, len) = parse_record_header(data, start);
    let mut drawing_offset = start + RECORD_HEADER_LEN + len as usize;
    walk(data, start, drawing_offset, &mut records);
    // Each OfficeArtWordDrawing: dgglbl byte followed by a DgContainer.
    while drawing_offset + 1 + RECORD_HEADER_LEN <= end {
        let dgglbl = data[drawing_offset];
        assert!(dgglbl <= 1, "dgglbl must be 0 (main) or 1 (header)");
        let (_ver, _inst, _record_type, len) = parse_record_header(data, drawing_offset + 1);
        walk(
            data,
            drawing_offset + 1,
            drawing_offset + 1 + RECORD_HEADER_LEN + len as usize,
            &mut records,
        );
        drawing_offset += 1 + RECORD_HEADER_LEN + len as usize;
    }
    records
}

#[test]
fn dgg_info_layout_matches_ms_odraw() {
    let png = Picture::new(png_bytes()).unwrap();
    let jpeg = Picture::new(jpeg_bytes()).unwrap();
    let positions = sample_positions();
    let shapes = floating_shapes(&png, &jpeg, &positions);

    let dgg = build_dgg_info(&shapes, &[], 3).unwrap();
    let records = collect_records(&dgg, 0, dgg.len());

    // Top level: DggContainer, then dgglbl + DgContainer.
    let (off, ver, _inst, record_type, len) = records[0];
    assert_eq!(
        (off, ver, record_type),
        (0, VERSION_CONTAINER, RECORD_DGG_CONTAINER)
    );
    let dg_container_offset = off + RECORD_HEADER_LEN + len as usize;
    assert_eq!(dgg[dg_container_offset], DGGLBL_MAIN_DOCUMENT);

    // Dgg: cluster table and counts.
    let dgg_record = records
        .iter()
        .find(|record| record.3 == RECORD_DGG)
        .unwrap();
    let dgg_payload = &dgg[dgg_record.0 + RECORD_HEADER_LEN..];
    let spid_max = u32::from_le_bytes(dgg_payload[0..4].try_into().unwrap());
    assert_eq!(spid_max, 2 * SHAPE_IDS_PER_CLUSTER);
    assert_eq!(
        u32::from_le_bytes(dgg_payload[4..8].try_into().unwrap()),
        2 // one OfficeArtIDCL plus one, per [MS-ODRAW] 2.2.47
    );
    // cspSaved = 2 shapes + group; cdgSaved = 1 drawing.
    assert_eq!(
        u32::from_le_bytes(dgg_payload[8..12].try_into().unwrap()),
        3
    );
    assert_eq!(
        u32::from_le_bytes(dgg_payload[12..16].try_into().unwrap()),
        1
    );

    // BStoreContainer holds one BSE per picture.
    let bstore = records
        .iter()
        .find(|record| record.3 == RECORD_BSTORE_CONTAINER)
        .unwrap();
    assert_eq!(bstore.2, 2);
    let bses: Vec<_> = records
        .iter()
        .filter(|record| record.3 == RECORD_BSE)
        .collect();
    assert_eq!(bses.len(), 2);

    // Dg: shape count includes the group shape; spidCur is next free.
    let dg = records.iter().find(|record| record.3 == RECORD_DG).unwrap();
    let dg_payload = &dgg[dg.0 + RECORD_HEADER_LEN..];
    assert_eq!(u32::from_le_bytes(dg_payload[0..4].try_into().unwrap()), 3);
    assert_eq!(
        u32::from_le_bytes(dg_payload[4..8].try_into().unwrap()),
        FIRST_PICTURE_SHAPE_ID + 2
    );

    // Group shape with fGroup|fPatriarch, then one picture shape per picture.
    let sps: Vec<_> = records
        .iter()
        .filter(|record| record.3 == RECORD_SP)
        .collect();
    assert_eq!(sps.len(), 3);
    let (group_spid, group_flags) = (
        u32::from_le_bytes(dgg[sps[0].0 + 8..sps[0].0 + 12].try_into().unwrap()),
        u32::from_le_bytes(dgg[sps[0].0 + 12..sps[0].0 + 16].try_into().unwrap()),
    );
    assert_eq!(group_spid, GROUP_SHAPE_ID);
    assert_eq!(group_flags, SP_FLAG_GROUP | SP_FLAG_PATRIARCH);
    for (index, sp) in sps[1..].iter().enumerate() {
        let spid = u32::from_le_bytes(dgg[sp.0 + 8..sp.0 + 12].try_into().unwrap());
        assert_eq!(spid, FIRST_PICTURE_SHAPE_ID + index as u32);
        assert_eq!(sp.2, SHAPE_TYPE_PICTURE_FRAME);
    }

    // OPT pib references the BSEs 1-based, in order.
    let opts: Vec<_> = records
        .iter()
        .filter(|record| record.3 == RECORD_OPT)
        .collect();
    assert_eq!(opts.len(), 2);
    for (index, opt) in opts.iter().enumerate() {
        let prop = u16::from_le_bytes(dgg[opt.0 + 8..opt.0 + 10].try_into().unwrap());
        let value = u32::from_le_bytes(dgg[opt.0 + 10..opt.0 + 14].try_into().unwrap());
        assert_eq!(prop, OPT_PIB_BLIP_INDEX);
        assert_eq!(value, index as u32 + OPT_PIB_FIRST_BSE);
    }

    // ClientAnchor records index into the PlcfSpa aCP array, in order.
    let anchors: Vec<_> = records
        .iter()
        .filter(|record| record.3 == RECORD_CLIENT_ANCHOR)
        .collect();
    assert_eq!(anchors.len(), 2);
    for (index, anchor) in anchors.iter().enumerate() {
        let value = u32::from_le_bytes(dgg[anchor.0 + 8..anchor.0 + 12].try_into().unwrap());
        assert_eq!(value, index as u32);
    }

    // Spids across the drawing are unique.
    let mut spids: Vec<u32> = sps
        .iter()
        .map(|sp| u32::from_le_bytes(dgg[sp.0 + 8..sp.0 + 12].try_into().unwrap()))
        .collect();
    spids.sort_unstable();
    spids.dedup();
    assert_eq!(spids.len(), 3);
}

#[test]
fn dgg_info_bse_blips_parse_with_crate_reader() {
    use litchi_odraw::Record;
    use litchi_odraw::image::{Context, Entry};

    let png = Picture::new(png_bytes()).unwrap();
    let jpeg = Picture::new(jpeg_bytes()).unwrap();
    let positions = sample_positions();
    let shapes = floating_shapes(&png, &jpeg, &positions);

    let dgg = build_dgg_info(&shapes, &[], 2).unwrap();
    let records = collect_records(&dgg, 0, dgg.len());
    let bses: Vec<_> = records
        .iter()
        .filter(|record| record.3 == RECORD_BSE)
        .collect();
    assert_eq!(bses.len(), 2);

    let expected = [(BlipKind::Png, png.data()), (BlipKind::Jpeg, jpeg.data())];
    for (bse_record, (kind, payload)) in bses.iter().zip(expected.iter()) {
        let (record, _) = Record::parse(&dgg, bse_record.0).unwrap();
        let bse = Entry::parse(record).unwrap();
        assert_eq!(bse.kind().unwrap(), *kind);
        let blip = bse.resolve(Context::new()).unwrap().unwrap();
        assert_eq!(blip.data(), *payload);
    }
}

#[test]
fn spid_max_rounds_up_to_next_cluster() {
    assert_eq!(spid_max(FIRST_PICTURE_SHAPE_ID), 2 * SHAPE_IDS_PER_CLUSTER);
    assert_eq!(
        spid_max(FIRST_PICTURE_SHAPE_ID + 1),
        2 * SHAPE_IDS_PER_CLUSTER
    );
    assert_eq!(
        spid_max(FIRST_PICTURE_SHAPE_ID + SHAPE_IDS_PER_CLUSTER),
        3 * SHAPE_IDS_PER_CLUSTER
    );
}

#[test]
fn dgg_info_with_header_drawing_uses_own_cluster() {
    use crate::writer::shapes::Shape;

    let png = Picture::new(png_bytes()).unwrap();
    let jpeg = Picture::new(jpeg_bytes()).unwrap();
    let positions = sample_positions();
    let main_shapes = floating_shapes(&png, &jpeg, &positions);

    let rect = Shape::new(crate::writer::shapes::Kind::Rectangle, 1440, 720).unwrap();
    let position = FloatingPosition::new(720, 360);
    let header_shapes = vec![FloatingShapeInfo {
        anchor_cp: 0,
        shape_id: HEADER_FIRST_SHAPE_ID,
        content: FloatingShapeContent::Primitive(&rect),
        width_twips: rect.width_twips(),
        height_twips: rect.height_twips(),
        position: &position,
        text: Some("hdr"),
    }];

    let dgg = build_dgg_info(&main_shapes, &header_shapes, 2).unwrap();
    let records = collect_records(&dgg, 0, dgg.len());

    // Two drawings: main (dgglbl 0, Dg instance 1) and header (dgglbl 1,
    // Dg instance 2). The Dgg cluster table has two entries, while cidcl
    // is the entry count plus one.
    let dgg_record = records
        .iter()
        .find(|record| record.3 == RECORD_DGG)
        .unwrap();
    let dgg_payload = &dgg[dgg_record.0 + RECORD_HEADER_LEN..];
    assert_eq!(u32::from_le_bytes(dgg_payload[4..8].try_into().unwrap()), 3);
    assert_eq!(
        u32::from_le_bytes(dgg_payload[0..4].try_into().unwrap()),
        3 * SHAPE_IDS_PER_CLUSTER
    );
    let dg_records: Vec<_> = records
        .iter()
        .filter(|record| record.3 == RECORD_DG)
        .collect();
    assert_eq!(dg_records.len(), 2);
    assert_eq!(dg_records[0].2, 1);
    assert_eq!(dg_records[1].2, 2);

    // dgglbl bytes: main then header.
    let dg_container_offsets: Vec<usize> = records
        .iter()
        .filter(|record| record.3 == RECORD_DG_CONTAINER)
        .map(|record| record.0)
        .collect();
    assert_eq!(dg_container_offsets.len(), 2);
    assert_eq!(dgg[dg_container_offsets[0] - 1], 0);
    assert_eq!(dgg[dg_container_offsets[1] - 1], 1);

    // The header shape uses the header cluster spid and its own TXID
    // numbering (0x10000 for the first header text box).
    let sps: Vec<_> = records
        .iter()
        .filter(|record| record.3 == RECORD_SP)
        .collect();
    let spids: Vec<u32> = sps
        .iter()
        .map(|sp| u32::from_le_bytes(dgg[sp.0 + 8..sp.0 + 12].try_into().unwrap()))
        .collect();
    assert!(spids.contains(&HEADER_FIRST_SHAPE_ID));

    let txid_records: Vec<_> = records.iter().filter(|record| record.3 == 0xF00D).collect();
    assert_eq!(txid_records.len(), 1);
    let txid = u32::from_le_bytes(
        dgg[txid_records[0].0 + 8..txid_records[0].0 + 12]
            .try_into()
            .unwrap(),
    );
    assert_eq!(txid, 0x0001_0000);
}
