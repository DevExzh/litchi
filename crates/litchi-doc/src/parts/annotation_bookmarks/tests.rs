//! Focused regression tests for the inert annotation-bookmark owner.

use super::{
    Editor, MAX_ENTRIES, Snapshot, Tag, TagId, Tags, TransactionError, parse, parse_bytes, to_bytes,
};
use crate::package::Error as PackageError;
use crate::parts::fib::FileInformationBlock;
use crate::parts::protection::{ProtectionAuthorization, ProtectionPolicy};
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

fn table_with_count(count: usize) -> Vec<u8> {
    let mut data = Vec::with_capacity(6 + count * 12);
    data.extend_from_slice(&0xFFFFu16.to_le_bytes());
    data.extend_from_slice(&(count as u16).to_le_bytes());
    data.extend_from_slice(&10u16.to_le_bytes());
    for tag in 0..count {
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0x0100u16.to_le_bytes());
        data.extend_from_slice(&(tag as u32).to_le_bytes());
        data.extend_from_slice(&(-1i32).to_le_bytes());
    }
    data
}

#[test]
fn accepts_maximum_count_and_rejects_correctly_sized_next_count() {
    let maximum = table_with_count(MAX_ENTRIES);
    let tags = parse_bytes(&maximum).expect("maximum annotation-bookmark table parses");
    assert_eq!(tags.len(), MAX_ENTRIES);

    let too_many = table_with_count(MAX_ENTRIES + 1);
    assert_eq!(too_many.len(), maximum.len() + 12);
    assert!(parse_bytes(&too_many).is_err());
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
    wrong_count[2..4].copy_from_slice(&0x3ffcu16.to_le_bytes());
    let error = parse_bytes(&wrong_count).unwrap_err();
    assert!(error.to_string().contains("0x3FFB"));

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

#[test]
fn protected_package_allows_noop_but_requires_explicit_authorization_for_metadata_edit() {
    let original = write_protected_doc();
    let committed = Editor::open(original.clone()).unwrap().commit().unwrap();
    assert_eq!(committed.snapshot().finish().unwrap(), original);

    let mut editor = Editor::open(original.clone()).unwrap();
    let error = editor.set(sample()).unwrap_err();
    assert!(matches!(error, PackageError::ProtectionDenied(_)));
    assert_eq!(editor.finish().unwrap(), original);

    let authorization =
        ProtectionAuthorization::audited("test-suite", "approved metadata repair").unwrap();
    let policy = ProtectionPolicy::allow_protected(authorization);
    let mut editor = Editor::open_with_policy(original.clone(), policy.clone()).unwrap();
    let committed = editor.set(sample()).unwrap();
    assert!(!committed.patch().is_noop());
    assert!(committed.snapshot().value().is_some());
    assert!(committed.package_patch().apply(&original).is_err());
    assert_eq!(
        committed
            .package_patch()
            .apply_with_policy(&original, policy)
            .unwrap(),
        committed.snapshot().finish().unwrap()
    );
}

fn fib_with_pointer(offset: usize, length: usize) -> FileInformationBlock {
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
    word[POINTER..POINTER + 4].copy_from_slice(&u32::try_from(offset).unwrap().to_le_bytes());
    word[POINTER + 4..POINTER + 8].copy_from_slice(&u32::try_from(length).unwrap().to_le_bytes());
    FileInformationBlock::parse(&word).unwrap()
}

fn write_doc(prefix: &[u8], payload: Option<&[u8]>) -> Vec<u8> {
    const DOP_INDEX: usize = 31;
    const POINTER_COUNT: usize = 136;
    let mut table_stream = prefix.to_vec();
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
    if let Some(payload) = payload {
        let offset = table_stream.len();
        table_stream.extend_from_slice(payload);
        word[POINTER..POINTER + 4].copy_from_slice(&u32::try_from(offset).unwrap().to_le_bytes());
        word[POINTER + 4..POINTER + 8]
            .copy_from_slice(&u32::try_from(payload.len()).unwrap().to_le_bytes());
    }
    let dop_offset = table_stream.len();
    table_stream.extend_from_slice(
        &crate::parts::document_properties::DocumentProperties::writer_bytes(
            false, false, false, true,
        ),
    );
    let dop_pointer = 154 + DOP_INDEX * 8;
    word[dop_pointer..dop_pointer + 4]
        .copy_from_slice(&u32::try_from(dop_offset).unwrap().to_le_bytes());
    word[dop_pointer + 4..dop_pointer + 8].copy_from_slice(&594u32.to_le_bytes());

    let mut writer = OleWriter::new();
    writer.create_stream(&["WordDocument"], &word).unwrap();
    writer.create_stream(&["0Table"], &table_stream).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn write_protected_doc() -> Vec<u8> {
    const DOP_INDEX: usize = 31;
    let pointer_count = 136usize;
    let dop_offset = 16usize;
    let mut table_stream = vec![0x5a; dop_offset];
    let mut dop = crate::parts::document_properties::DocumentProperties::writer_bytes(
        false, false, false, true,
    );
    dop[6] = 0x10;
    table_stream.extend_from_slice(&dop);
    let pointer_end = 154 + pointer_count * 8;
    let mut word = vec![0u8; pointer_end + 4];
    word[0..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
    // FibBase.csw and cslw are fixed MS-DOC counts.
    word[32..34].copy_from_slice(&0x000eu16.to_le_bytes());
    word[62..64].copy_from_slice(&0x0016u16.to_le_bytes());
    word[2..4].copy_from_slice(&0x00c1u16.to_le_bytes());
    word[152..154].copy_from_slice(&(pointer_count as u16).to_le_bytes());
    word[pointer_end..pointer_end + 2].copy_from_slice(&2u16.to_le_bytes());
    word[pointer_end + 2..pointer_end + 4].copy_from_slice(&0x0101u16.to_le_bytes());
    let pointer = 154 + DOP_INDEX * 8;
    word[pointer..pointer + 4].copy_from_slice(&(dop_offset as u32).to_le_bytes());
    word[pointer + 4..pointer + 8].copy_from_slice(&(dop.len() as u32).to_le_bytes());

    let mut writer = OleWriter::new();
    writer.create_stream(&["WordDocument"], &word).unwrap();
    writer.create_stream(&["0Table"], &table_stream).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}
