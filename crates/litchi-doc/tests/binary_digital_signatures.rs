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
        litchi_doc::raw::Reader::open(root.join("test-data/ole/doc/documentProperties.doc"))
            .unwrap();
    assert!(
        doc.signatures(&litchi_sign::Policy::strict())
            .unwrap()
            .is_empty()
    );
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

    let bytes = common::with_stream(
        common::authored_document("signed body"),
        "\u{0005}DocumentSummaryInformation",
        stream,
    );
    let mut package = litchi_doc::raw::Reader::from_reader(Cursor::new(bytes.clone())).unwrap();
    let package_signature = package
        .vba_signature(litchi_ole_common::vba_signature::Limits::default())
        .unwrap()
        .unwrap();
    assert_eq!(package_signature.info().signature(), [1, 2, 3]);

    let package = litchi_doc::Package::from_bytes(&bytes).unwrap();
    let source = package.property_sets().unwrap();
    let signature = source.vba_signature().unwrap().unwrap();
    assert_eq!(signature.kind(), litchi_doc::vba_signature::Kind::Property);
    assert_eq!(signature.info().signature(), [1, 2, 3]);

    let mut no_op = package.edit_property_sets().unwrap();
    assert!(
        !no_op
            .edit_vba_signature(litchi_ole_common::vba_signature::Limits::default(), |_| Ok(
                ()
            ))
            .unwrap()
    );
    let no_op_commit = no_op.commit().unwrap();
    assert!(!no_op_commit.diagnostics().changed());
    assert_eq!(no_op_commit.package().to_bytes().unwrap(), source.bytes());

    let mut transaction = package.edit_property_sets().unwrap();
    assert!(
        transaction
            .edit_vba_signature(
                litchi_ole_common::vba_signature::Limits::default(),
                |edit| edit.set_signature([9, 8]).map(|_| ())
            )
            .unwrap()
    );
    let commit = transaction.commit().unwrap();
    let edited = commit
        .package()
        .property_sets()
        .unwrap()
        .vba_signature()
        .unwrap()
        .unwrap();
    assert_eq!(edited.info().signature(), [9, 8]);
    assert_eq!(edited.info().certificate_store(), [4, 5, 6, 7]);
    let reverted = commit
        .package()
        .apply_property_sets(&commit.patch().inverse())
        .unwrap();
    assert_eq!(reverted.package().to_bytes().unwrap(), source.bytes());
}
