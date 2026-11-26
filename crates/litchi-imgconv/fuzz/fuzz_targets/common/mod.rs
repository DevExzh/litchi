//! Shared budgets and wrapper construction for the metafile fuzz targets.

#![allow(dead_code)]

/// A raster-sized conversion budget: a 256x256 render, 1 MiB of encoded
/// output and 20 000 metafile records.
pub fn small_limits() -> litchi_imgconv::Limits {
    litchi_imgconv::Limits {
        max_encoded_bytes: 1 << 20,
        max_uncompressed_bytes: 4 << 20,
        max_width: 256,
        max_height: 256,
        max_pixels: 256 * 256,
        max_output_bytes: 1 << 20,
        max_records: 20_000,
        max_objects: 1_024,
        max_state_depth: 32,
        max_path_points: 65_536,
        max_svg_elements: 20_000,
    }
}

/// The budget of the targets that rasterize: the rasterizer's cost grows with
/// the canvas area, and the sanitizer-instrumented build runs it much slower
/// than a release build, so the canvas is a quarter of [`small_limits`].
pub fn raster_limits() -> litchi_imgconv::Limits {
    litchi_imgconv::Limits {
        max_width: 128,
        max_height: 128,
        max_pixels: 128 * 128,
        ..small_limits()
    }
}

/// Asserts what every successful SVG conversion promises: the output stays
/// within `max_output_bytes`, is UTF-8 and parses as an XML document whose
/// root element is `svg`. A viewer rejects anything less.
pub fn assert_well_formed_svg(bytes: &[u8], max_output_bytes: usize) {
    assert!(
        bytes.len() <= max_output_bytes,
        "SVG exceeds the output limit"
    );
    let text = std::str::from_utf8(bytes).expect("SVG output is not UTF-8");
    let document = roxmltree::Document::parse(text)
        .unwrap_or_else(|error| panic!("SVG output is not well-formed XML: {error}"));
    assert_eq!(document.root_element().tag_name().name(), "svg");
}

/// The `OfficeArt` traversal budget: 20 000 records, 1 MiB per record body and
/// 256 image files per scanner.
pub fn art_limits() -> litchi_odraw::image::Limits {
    litchi_odraw::image::Limits {
        records: litchi_odraw::Limits {
            max_record_bytes: 1 << 20,
            max_depth: 32,
            max_records: 20_000,
        },
        max_blip_bytes: 1 << 20,
        max_store_entries: 256,
        max_files: 256,
    }
}

/// Wraps `payload` in a well-formed EMF stream whose only content record is an
/// `EMR_COMMENT_EMFPLUS` (MS-EMF 2.3.2.7) carrying it, so arbitrary bytes reach
/// the EMF+ playback that a dual-mode metafile drives.
pub fn wrap_emfplus(payload: &[u8]) -> Vec<u8> {
    const EMFPLUS_COMMENT_IDENTIFIER: u32 = 0x2B46_4D45;
    const EMR_COMMENT: u32 = 0x0000_0046;
    const EMR_EOF: u32 = 0x0000_000E;
    const HEADER_SIZE: usize = 88;
    const EOF_SIZE: usize = 20;

    // EMR_COMMENT: Type, Size, DataSize, then the identifier and payload,
    // padded to a four-byte boundary like any EMF record.
    let data_size = 4 + payload.len();
    let unpadded = 8 + 4 + data_size;
    let comment_size = unpadded + (4 - unpadded % 4) % 4;
    let file_size = HEADER_SIZE + comment_size + EOF_SIZE;

    let mut emf = Vec::with_capacity(file_size);

    // EMR_HEADER (MS-EMF 2.3.4.2), 88 bytes.
    for value in [
        1u32,
        HEADER_SIZE as u32,
        0,           // rclBounds.left
        0,           // rclBounds.top
        256,         // rclBounds.right
        256,         // rclBounds.bottom
        0,           // rclFrame.left
        0,           // rclFrame.top
        25_400,      // rclFrame.right, 254 mm in 0.01 mm units
        25_400,      // rclFrame.bottom
        0x464D_4520, // "EMF "
        0x0001_0000, // version
        file_size as u32,
        3,   // nRecords: header, comment, EOF
        1,   // nHandles, with the reserved high half zero
        0,   // nDescription
        0,   // offDescription
        0,   // nPalEntries
        256, // szlDevice.cx
        256, // szlDevice.cy
        254, // szlMillimeters.cx
        254, // szlMillimeters.cy
    ] {
        emf.extend_from_slice(&value.to_le_bytes());
    }

    // EMR_COMMENT_EMFPLUS (MS-EMF 2.3.3.3, MS-EMFPLUS 2.3.1.1).
    for value in [EMR_COMMENT, comment_size as u32, data_size as u32] {
        emf.extend_from_slice(&value.to_le_bytes());
    }
    emf.extend_from_slice(&EMFPLUS_COMMENT_IDENTIFIER.to_le_bytes());
    emf.extend_from_slice(payload);
    emf.resize(file_size - EOF_SIZE, 0);

    // EMR_EOF (MS-EMF 2.3.4.1).
    for value in [
        EMR_EOF,
        EOF_SIZE as u32,
        0,
        EOF_SIZE as u32,
        EOF_SIZE as u32,
    ] {
        emf.extend_from_slice(&value.to_le_bytes());
    }
    debug_assert_eq!(emf.len(), file_size);
    emf
}
