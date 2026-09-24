#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_wrap,
    clippy::let_underscore_must_use,
    clippy::manual_midpoint,
    clippy::map_unwrap_or,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    clippy::bool_assert_comparison,
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    clippy::decimal_bitwise_operands,
    clippy::default_trait_access,
    clippy::doc_markdown,
    clippy::expect_used,
    clippy::field_reassign_with_default,
    clippy::float_cmp,
    clippy::implicit_clone,
    clippy::items_after_statements,
    clippy::manual_let_else,
    clippy::manual_repeat_n,
    clippy::manual_string_new,
    clippy::match_wildcard_for_single_variants,
    clippy::needless_raw_string_hashes,
    clippy::redundant_closure_for_method_calls,
    clippy::shadow_unrelated,
    clippy::similar_names,
    clippy::uninlined_format_args,
    clippy::unreadable_literal,
    clippy::unwrap_used,
    reason = "integration-test fixtures favor explicit wire values and concise panic-driven assertions over production-style ergonomics"
)]

use std::path::PathBuf;

use litchi_cfb::OleWriter;
use litchi_ole_common::property_set::{
    CodePage, DOCUMENT_SUMMARY_INFORMATION_FMTID, Section, Stream, Value,
};
use std::io::Cursor;

mod common;

fn property_blob() -> Vec<u8> {
    let mut bytes = vec![0; 56];
    let put = |bytes: &mut [u8], offset: usize, value: u32| {
        bytes[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
    };
    put(&mut bytes, 0, 48);
    put(&mut bytes, 4, 8);
    put(&mut bytes, 8, 3);
    put(&mut bytes, 12, 44);
    put(&mut bytes, 16, 4);
    put(&mut bytes, 20, 48);
    put(&mut bytes, 24, 0);
    put(&mut bytes, 28, 52);
    put(&mut bytes, 32, 0xA5A5_5A5A);
    put(&mut bytes, 36, 0);
    put(&mut bytes, 40, 54);
    bytes[44..47].copy_from_slice(&[1, 2, 3]);
    bytes[47] = 0xCC;
    bytes[48..52].copy_from_slice(&[4, 5, 6, 7]);
    bytes
}

#[test]
fn doc_facade_discovers_unsigned_signature_state() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let mut doc =
        litchi_doc::Package::open(root.join("test-data/ole/doc/documentProperties.doc")).unwrap();
    assert!(doc.signatures().unwrap().is_empty());
}

#[test]
fn public_property_set_facade_reads_and_edits_vba_signature_payloads() {
    let raw_blob = property_blob();
    assert_eq!(raw_blob.len(), 56);
    assert_eq!(&raw_blob[..4], &48u32.to_le_bytes());
    let mut section = Section::new(DOCUMENT_SUMMARY_INFORMATION_FMTID);
    section.set_page(CodePage::Utf16Le);
    section
        .add(
            litchi_ole_common::property_set::document_summary::DIGITAL_SIGNATURE,
            Value::Blob(raw_blob.clone()),
        )
        .unwrap();
    let stream = Stream::new(section).to_bytes().unwrap();
    let mut expected_blob_property = vec![0x41, 0, 0, 0, 56, 0, 0, 0];
    expected_blob_property.extend_from_slice(&raw_blob);
    assert!(
        stream
            .windows(expected_blob_property.len())
            .any(|window| window == expected_blob_property),
        "PIDDSI must serialize VT_BLOB with one 56-byte size before raw cb=48 DigSigBlob"
    );

    let mut writer = OleWriter::new();
    const DOP_INDEX: usize = 31;
    const POINTER_COUNT: usize = 136;
    let pointer_end = 154 + POINTER_COUNT * 8;
    let mut word = vec![0u8; pointer_end + 4];
    word[0..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
    // FibBase.csw and cslw are fixed MS-DOC counts.
    word[32..34].copy_from_slice(&0x000eu16.to_le_bytes());
    word[62..64].copy_from_slice(&0x0016u16.to_le_bytes());
    word[2..4].copy_from_slice(&0x00c1u16.to_le_bytes());
    word[152..154].copy_from_slice(&(POINTER_COUNT as u16).to_le_bytes());
    word[pointer_end..pointer_end + 2].copy_from_slice(&2u16.to_le_bytes());
    word[pointer_end + 2..pointer_end + 4].copy_from_slice(&0x0101u16.to_le_bytes());
    let table = vec![0u8; 500];
    let dop_pointer = 154 + DOP_INDEX * 8;
    word[dop_pointer..dop_pointer + 4].copy_from_slice(&0u32.to_le_bytes());
    word[dop_pointer + 4..dop_pointer + 8].copy_from_slice(&500u32.to_le_bytes());
    writer.create_stream(&["WordDocument"], &word).unwrap();
    writer.create_stream(&["0Table"], &table).unwrap();
    writer
        .create_stream(&["\u{0005}DocumentSummaryInformation"], &stream)
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();

    let bytes = common::with_valid_word97_dop(output.into_inner());
    let mut package = litchi_doc::Package::from_reader(Cursor::new(bytes.clone())).unwrap();
    let package_signature = package.vba_signature().unwrap().unwrap();
    assert_eq!(package_signature.info().signature(), [1, 2, 3]);

    let source = litchi_doc::package::Snapshot::from_bytes(bytes).unwrap();
    let signature = source.vba_signature().unwrap().unwrap();
    assert_eq!(signature.kind(), litchi_doc::vba_signature::Kind::Property);
    assert_eq!(signature.info().signature(), [1, 2, 3]);

    let mut no_op = source.transaction().unwrap();
    assert!(!no_op.edit_vba_signature(|_| Ok(())).unwrap());
    let no_op_commit = no_op.commit().unwrap();
    assert!(!no_op_commit.changed());
    assert_eq!(no_op_commit.snapshot().bytes(), source.bytes());

    let mut transaction = source.transaction().unwrap();
    assert!(
        transaction
            .edit_vba_signature(|edit| edit.set_signature([9, 8]).map(|_| ()))
            .unwrap()
    );
    let commit = transaction.commit().unwrap();
    let edited = commit.snapshot().vba_signature().unwrap().unwrap();
    assert_eq!(edited.info().signature(), [9, 8]);
    assert_eq!(edited.info().certificate_store(), [4, 5, 6, 7]);
    let reverted = commit.patch().revert(commit.snapshot()).unwrap();
    assert_eq!(reverted.bytes(), source.bytes());
}
