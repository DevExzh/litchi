//! Real host-origin VBA sources vendored from Apache POI (Apache-2.0).
//!
//! POI ships each source beside its macro-enabled Office host fixture. These
//! tests exercise the MS-OVBA codec directly, with finite input and output
//! budgets, so the runtime crate does not depend on an OOXML ZIP reader.

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "fixture setup and assertions intentionally panic on failure"
)]

use litchi_vba::{Limits, codec};

const FIXTURE_LIMITS: Limits = Limits {
    max_cfb_bytes: 32 * 1024,
    max_compressed_stream_bytes: 16 * 1024,
    max_decompressed_stream_bytes: 32 * 1024,
    max_modules: 32,
    max_string_bytes: 16 * 1024,
    max_total_source_bytes: 32 * 1024,
};

const WORD: &[u8] = include_bytes!("../../../test-data/poi/test-data/document/SimpleMacro.vba");
const EXCEL: &[u8] = include_bytes!("../../../test-data/poi/test-data/spreadsheet/SimpleMacro.vba");
const POWERPOINT: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/slideshow/SimpleMacro.vba");

#[test]
fn real_word_vba_source_is_bounded_and_lossless() {
    assert_fixture(WORD, "This is a macro word processing document");
}

#[test]
fn real_excel_vba_source_is_bounded_and_lossless() {
    assert_fixture(EXCEL, "This is a macro workbook");
}

#[test]
fn real_powerpoint_vba_source_is_bounded_and_lossless() {
    assert_fixture(POWERPOINT, "This is a macro slideshow");
}

fn assert_fixture(source: &[u8], expected_source: &str) {
    assert!(source.len() <= FIXTURE_LIMITS.max_decompressed_stream_bytes);
    assert!(
        std::str::from_utf8(source)
            .unwrap()
            .contains(expected_source)
    );

    let compressed = codec::encode(source, &FIXTURE_LIMITS).unwrap();
    assert!(compressed.len() <= FIXTURE_LIMITS.max_compressed_stream_bytes);
    assert_eq!(codec::decode(&compressed, &FIXTURE_LIMITS).unwrap(), source);
    assert_eq!(codec::encode(source, &FIXTURE_LIMITS).unwrap(), compressed);
}
