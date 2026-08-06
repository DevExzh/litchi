//! Focused regression tests for the inert annotation-bookmark owner.

use super::{Editor, Snapshot, Tag, TagId, Tags, TransactionError, parse, parse_bytes, to_bytes};
use crate::parts::fib::FileInformationBlock;
use litchi_cfb::OleWriter;
use std::io::Cursor;

const POINTER: usize = 154 + super::FIB_INDEX * 8;

fn sample() -> Tags {
    Tags::try_new(vec![
        Tag::new(TagId::new(0x8000_0001)),
        Tag::new(TagId::new(0x0000_0042)),
    ])
    .expect("sample tags are valid")
}

#[test]
fn round_trip_preserves_opaque_tag_ids() {
    let value = sample();
    let bytes = to_bytes(&value).expect("sample serializes");
    assert_eq!(parse_bytes(&bytes).expect("sample parses"), value);
    assert_eq!(parse_bytes(&bytes).unwrap().to_bytes().unwrap(), bytes);
    assert_eq!(value.entries()[0].id().raw(), 0x8000_0001);
}

#[test]
fn parses_fib_range_and_rejects_invalid_atnbe_shapes() {
    let payload = to_bytes(&sample()).unwrap();
    let offset = 4usize;
    let mut table = vec![0xa5; offset];
    table.extend_from_slice(&payload);
    let fib = fib_with_pointer(offset, payload.len());
    assert_eq!(parse(&fib, &table).unwrap(), Some(sample()));

    let mut wrong_extend = payload.clone();
    wrong_extend[0..2].copy_from_slice(&0u16.to_le_bytes());
    assert!(parse_bytes(&wrong_extend).is_err());

    let mut wrong_extra = payload.clone();
    wrong_extra[4..6].copy_from_slice(&0u16.to_le_bytes());
    assert!(parse_bytes(&wrong_extra).is_err());

    let mut wrong_count = payload.clone();
    wrong_count[2..4].copy_from_slice(&0x3ffdu16.to_le_bytes());
    assert!(parse_bytes(&wrong_count).is_err());

    let mut wrong_string = payload.clone();
    wrong_string[6..8].copy_from_slice(&1u16.to_le_bytes());
    assert!(parse_bytes(&wrong_string).is_err());

    let mut wrong_class = payload.clone();
    wrong_class[8..10].copy_from_slice(&0u16.to_le_bytes());
    assert!(parse_bytes(&wrong_class).is_err());

    let mut wrong_old_tag = payload.clone();
    wrong_old_tag[16..20].copy_from_slice(&0i32.to_le_bytes());
    assert!(parse_bytes(&wrong_old_tag).is_err());

    let mut duplicate = payload.clone();
    duplicate[22..26].copy_from_slice(&0x8000_0001u32.to_le_bytes());
    assert!(parse_bytes(&duplicate).is_err());

    assert!(parse_bytes(&payload[..payload.len() - 1]).is_err());
    let mut trailing = payload;
    trailing.push(0);
    assert!(parse_bytes(&trailing).is_err());
}

#[test]
fn transaction_is_atomic_and_reversible() {
    let source = Snapshot::new(sample()).unwrap();
    let mut transaction = source.edit();
    assert!(matches!(
        transaction.replace_entry(99, Tag::new(TagId::new(9))),
        Err(TransactionError::Invalid(_))
    ));
    assert_eq!(transaction.snapshot(), &source);

    transaction
        .replace_entry(0, Tag::new(TagId::new(0x1234)))
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert!(!commit.patch().is_noop());
    assert_eq!(
        commit.patch().inverse().apply(commit.snapshot()).unwrap(),
        source
    );
    assert!(matches!(
        commit.patch().apply(&Snapshot::empty()),
        Err(TransactionError::Conflict)
    ));
}

#[test]
fn package_noop_returns_exact_original_cfb_bytes() {
    let original = write_doc(b"opaque prefix", Some(&to_bytes(&sample()).unwrap()));
    let committed = Editor::open(original.clone()).unwrap().commit().unwrap();
    assert_eq!(committed.snapshot().finish().unwrap(), original);
    assert_eq!(committed.package_patch().after(), original.as_slice());
}

#[test]
fn package_edit_appends_and_clear_only_removes_the_fib_range() {
    let original = write_doc(b"opaque prefix", None);
    let mut editor = Editor::open(original).unwrap();
    let committed = editor.set(sample()).unwrap();
    let written = committed.snapshot().finish().unwrap();
    let reopened = Editor::open(written.clone()).unwrap();
    assert_eq!(reopened.value(), Some(&sample()));

    let mut ole = litchi_cfb::OleFile::open(Cursor::new(written)).unwrap();
    let table = ole.open_stream(&["0Table"]).unwrap();
    assert!(table.starts_with(b"opaque prefix"));
    drop(table);
    drop(ole);

    let mut editor = Editor::open(reopened.finish().unwrap()).unwrap();
    let committed = editor.clear().unwrap();
    let finished = committed.snapshot().finish().unwrap();
    let reopened = Editor::open(finished.clone()).unwrap();
    assert!(reopened.value().is_none());

    let mut ole = litchi_cfb::OleFile::open(Cursor::new(finished)).unwrap();
    let word = ole.open_stream(&["WordDocument"]).unwrap();
    assert_eq!(&word[POINTER..POINTER + 8], &[0; 8]);
    let table = ole.open_stream(&["0Table"]).unwrap();
    assert!(table.starts_with(b"opaque prefix"));
}

fn fib_with_pointer(offset: usize, length: usize) -> FileInformationBlock {
    let mut word = vec![0u8; POINTER + 8];
    word[0..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
    word[2..4].copy_from_slice(&0x00c1u16.to_le_bytes());
    word[152..154].copy_from_slice(&38u16.to_le_bytes());
    word[POINTER..POINTER + 4].copy_from_slice(&u32::try_from(offset).unwrap().to_le_bytes());
    word[POINTER + 4..POINTER + 8].copy_from_slice(&u32::try_from(length).unwrap().to_le_bytes());
    FileInformationBlock::parse(&word).unwrap()
}

fn write_doc(prefix: &[u8], payload: Option<&[u8]>) -> Vec<u8> {
    let mut table_stream = prefix.to_vec();
    let mut word = vec![0u8; POINTER + 8];
    word[0..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
    word[2..4].copy_from_slice(&0x00c1u16.to_le_bytes());
    word[152..154].copy_from_slice(&38u16.to_le_bytes());
    if let Some(payload) = payload {
        let offset = table_stream.len();
        table_stream.extend_from_slice(payload);
        word[POINTER..POINTER + 4].copy_from_slice(&u32::try_from(offset).unwrap().to_le_bytes());
        word[POINTER + 4..POINTER + 8]
            .copy_from_slice(&u32::try_from(payload.len()).unwrap().to_le_bytes());
    }

    let mut writer = OleWriter::new();
    writer.create_stream(&["WordDocument"], &word).unwrap();
    writer.create_stream(&["0Table"], &table_stream).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}
