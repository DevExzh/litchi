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

use litchi_cfb::{OleFile, OleWriter};
use litchi_doc::embedded_object::Limits;
use litchi_doc::writer::Writer;
use litchi_doc::{Editor, Package, WriteOptions};
use std::fs;
use std::io::Cursor;
use std::path::PathBuf;

fn base_doc() -> Vec<u8> {
    let mut writer = Writer::new();
    writer.add_paragraph("before embedded objects").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn object_cfb(marker: &[u8]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_stream(&["CONTENTS"], marker).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn inert_picf(marker: u32) -> Vec<u8> {
    let mut value = 12u32.to_le_bytes().to_vec();
    value.extend_from_slice(&marker.to_le_bytes());
    value.extend_from_slice(&[0; 4]);
    value
}

fn options(id: u32) -> WriteOptions {
    WriteOptions::new(id, object_cfb(&id.to_le_bytes()), inert_picf(id))
}

#[test]
fn managed_objects_add_reorder_remove_and_reopen_transactionally() {
    let original = base_doc();
    let mut editor = Editor::open(original, Limits::default()).unwrap();
    editor.add(options(11)).unwrap();
    editor.add(options(22)).unwrap();
    assert_eq!(
        editor
            .objects()
            .unwrap()
            .iter()
            .map(|value| value.storage_id)
            .collect::<Vec<_>>(),
        vec![11, 22]
    );
    editor.reorder(&[22, 11]).unwrap();
    assert_eq!(
        editor
            .objects()
            .unwrap()
            .iter()
            .map(|value| value.storage_id)
            .collect::<Vec<_>>(),
        vec![22, 11]
    );
    let bytes = editor.finish().unwrap();

    let mut reopened = Editor::open(bytes, Limits::default()).unwrap();
    assert_eq!(
        reopened
            .objects()
            .unwrap()
            .iter()
            .map(|value| value.storage_id)
            .collect::<Vec<_>>(),
        vec![22, 11]
    );
    reopened.remove(22).unwrap();
    let bytes = reopened.finish().unwrap();
    let ole = OleFile::open(Cursor::new(bytes.clone())).unwrap();
    assert!(!ole.exists(&["ObjectPool", "_22"]));
    assert!(ole.exists(&["ObjectPool", "_11"]));
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let document = package.document().unwrap();
    assert!(document.text().unwrap().contains("before embedded objects"));
}

#[test]
fn malformed_add_and_reorder_leave_editor_state_unchanged() {
    let mut editor = Editor::open(base_doc(), Limits::default()).unwrap();
    editor.add(options(1)).unwrap();
    editor.add(options(2)).unwrap();
    let before = editor.objects().unwrap();
    let mut invalid = options(3);
    invalid.picture_data[0..4].copy_from_slice(&999u32.to_le_bytes());
    assert!(editor.add(invalid).is_err());
    assert!(editor.reorder(&[1, 1]).is_err());
    assert_eq!(editor.objects().unwrap(), before);
}

#[test]
fn producer_fixtures_append_only_when_their_layout_is_supported() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let fixtures = [
        root.join("test-data/poi/test-data/document/word_with_embeded.doc"),
        root.join("test-data/libreoffice-core/embeddedobj/qa/cppunit/data/insert-file-config.doc"),
    ];
    let mut supported = 0usize;
    let mut rejected = 0usize;
    for (ordinal, path) in fixtures.into_iter().enumerate() {
        let original = fs::read(&path).unwrap();
        if let Ok(mut editor) = Editor::open(original.clone(), Limits::default()) {
            let id = 2_000_000 + ordinal as u32;
            editor.add(options(id)).unwrap();
            let output = editor.finish().unwrap();
            let reopened = Editor::open(output.clone(), Limits::default()).unwrap();
            assert!(
                reopened
                    .objects()
                    .unwrap()
                    .iter()
                    .any(|value| value.storage_id == id)
            );
            let mut package = Package::from_reader(Cursor::new(output)).unwrap();
            assert!(!package.document().unwrap().text().unwrap().is_empty());
            supported += 1;
        } else {
            // Construction owns no external state and therefore cannot mutate the fixture.
            assert_eq!(fs::read(&path).unwrap(), original);
            rejected += 1;
        }
    }
    assert_eq!(supported + rejected, 2);
}
