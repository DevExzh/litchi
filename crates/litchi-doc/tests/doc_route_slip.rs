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

//! DOC route-slip facade, lifecycle, selector, and protection tests.

use litchi_cfb::OleFile;
use litchi_doc::parts::fib::FileInformationBlock;
use litchi_doc::route_slip::record::{Error as RecordError, Snapshot};
use litchi_doc::route_slip::{
    DeliveryOption, EditKind, Error, Metadata, NarrowString, Protection, Recipient,
    RecipientSelectionError, RecipientSelector, parse, to_bytes,
};
use litchi_doc::{Builder, Package, SourceLimits, raw::Reader};
use std::io::Cursor;

fn base_doc() -> Vec<u8> {
    let mut writer = Builder::new();
    writer.add_paragraph("route-slip source").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.build().unwrap().write_to(&mut output).unwrap();
    output.into_inner()
}

fn recipient(name: &[u8]) -> Recipient {
    Recipient::try_new(vec![0, 0xff, name[0]], NarrowString::new(name.to_vec())).unwrap()
}

fn route(protection: Protection) -> Metadata {
    Metadata::try_new(
        true,
        true,
        true,
        protection,
        0,
        DeliveryOption::Serial,
        NarrowString::new(vec![0x80, b's', b'u', b'b']),
        NarrowString::new(vec![0x81, b'm', b's', b'g']),
        NarrowString::new(vec![0x82, b'o', b'k']),
        NarrowString::new(vec![0x83, b't', b'i', b't', b'l', b'e']),
        vec![recipient(b"Alice"), recipient(b"Bob")],
    )
    .unwrap()
}

fn route_payload(bytes: &[u8]) -> Vec<u8> {
    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    let word = ole.open_stream(&["WordDocument"]).unwrap();
    let fib = FileInformationBlock::parse(&word).unwrap();
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table = ole.open_stream(&[table_name]).unwrap();
    let (offset, length) = fib.get_table_pointer(70).unwrap();
    table[offset as usize..(offset + length) as usize].to_vec()
}

#[test]
fn package_edit_round_trips_document_facade_and_exact_fib_payload() {
    let original = base_doc();
    let package = Package::from_bytes(&original).unwrap();
    let mut edit = package.edit_route_slip().unwrap();
    assert!(edit.metadata().expect("route-slip metadata").is_none());

    let source = route(Protection::Off);
    edit.set(source.clone()).unwrap();
    assert_eq!(edit.metadata().expect("route-slip metadata"), Some(&source));

    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.patch().before_bytes(), original.as_slice());
    let bytes = commit.patch().after_bytes().to_vec();
    assert_eq!(commit.package().to_bytes().unwrap(), bytes);
    assert_eq!(parse_bytes_from_fib(&bytes), source);
    assert_ne!(bytes, original);

    let mut reader = Reader::from_reader(Cursor::new(bytes.clone())).unwrap();
    let document = reader.document().unwrap();
    assert_eq!(document.route_slip().unwrap(), Some(&source));
    assert_eq!(
        route_payload(&bytes),
        to_bytes(&source, &SourceLimits::default()).unwrap()
    );

    // The patch replays onto its source and its inverse restores it.
    let replayed = package.apply_route_slip(commit.patch()).unwrap();
    assert_eq!(replayed.package().to_bytes().unwrap(), bytes);
    let restored = commit
        .package()
        .apply_route_slip(&commit.patch().inverse())
        .unwrap();
    assert_eq!(restored.package().to_bytes().unwrap(), original);
    assert!(matches!(
        commit.package().apply_route_slip(commit.patch()),
        Err(Error::PatchConflict)
    ));
}

#[test]
fn unchanged_commit_publishes_the_source_package() {
    let original = base_doc();
    let package = Package::from_bytes(&original).unwrap();
    let commit = package.edit_route_slip().unwrap().commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), original);
}

#[test]
fn lifecycle_edits_are_snapshot_based_and_clear_the_fib_range() {
    let package = Package::from_bytes(&base_doc()).unwrap();
    let mut edit = package.edit_route_slip().unwrap();
    edit.set(route(Protection::Off)).unwrap();

    edit.set_stage(RecipientSelector::Name(b"Bob")).unwrap();
    assert_eq!(
        edit.metadata()
            .expect("route-slip metadata")
            .expect("route-slip value")
            .stage,
        1
    );

    let added = recipient(b"Carol");
    edit.add_recipient(added.clone()).unwrap();
    assert_eq!(
        edit.metadata()
            .expect("route-slip metadata")
            .expect("route-slip value")
            .recipients
            .last(),
        Some(&added)
    );

    edit.advance_stage().unwrap();
    assert_eq!(
        edit.metadata()
            .expect("route-slip metadata")
            .expect("route-slip value")
            .stage,
        2
    );
    assert!(edit.advance_stage().is_err());
    edit.set_stage(RecipientSelector::Index(0)).unwrap();
    edit.advance_stage().unwrap();
    assert_eq!(
        edit.metadata()
            .expect("route-slip metadata")
            .expect("route-slip value")
            .stage,
        1
    );

    edit.replace_recipient(RecipientSelector::Current, recipient(b"Bobby"))
        .unwrap();
    edit.remove_recipient(RecipientSelector::Index(0)).unwrap();
    assert_eq!(
        edit.metadata()
            .expect("route-slip metadata")
            .expect("route-slip value")
            .recipients[0]
            .name
            .as_bytes(),
        b"Bobby"
    );

    edit.complete().unwrap();
    assert!(edit.metadata().expect("route-slip metadata").is_none());
    let bytes = edit.commit().unwrap().patch().after_bytes().to_vec();
    let mut reader = Reader::from_reader(Cursor::new(bytes.clone())).unwrap();
    assert!(reader.document().unwrap().route_slip().unwrap().is_none());

    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    let word = ole.open_stream(&["WordDocument"]).unwrap();
    let fib = FileInformationBlock::parse(&word).unwrap();
    assert_eq!(fib.get_table_pointer(70), Some((0, 0)));
}

#[test]
fn protected_package_lifecycle_edit_is_rejected_atomically() {
    let package = Package::from_bytes(&base_doc()).unwrap();
    let mut edit = package.edit_route_slip().unwrap();
    edit.set(route(Protection::Annotation)).unwrap();
    let before = edit.route_slip().expect("route-slip snapshot").clone();

    let error = edit.clear().unwrap_err();
    assert!(error.to_string().contains("protected"));
    assert_eq!(edit.route_slip().expect("route-slip snapshot"), &before);
}

#[test]
fn selectors_are_typed_and_protected_edits_leave_the_source_unchanged() {
    let mut duplicate = route(Protection::Off);
    duplicate.recipients[1].name = NarrowString::new(b"Alice".to_vec());
    let snapshot = Snapshot::new(duplicate).unwrap();
    let mut edit = snapshot.edit();

    let error = edit
        .set_stage(RecipientSelector::Name(b"Alice"))
        .unwrap_err();
    assert!(matches!(
        error,
        RecordError::Selection(RecipientSelectionError::AmbiguousName { .. })
    ));
    let error = edit.set_stage(RecipientSelector::Index(99)).unwrap_err();
    assert!(matches!(
        error,
        RecordError::Selection(RecipientSelectionError::IndexOutOfBounds { .. })
    ));

    edit.set_stage(RecipientSelector::Index(1)).unwrap();
    // Dropping the edit discards it; the source never changed.
    drop(edit);
    let edit = snapshot.edit();
    assert_eq!(edit.staged(), &snapshot);

    let protected = Snapshot::new(route(Protection::Annotation)).unwrap();
    let mut protected_edit = protected.edit();
    let error = protected_edit
        .set_stage(RecipientSelector::Index(1))
        .unwrap_err();
    assert!(matches!(
        error,
        RecordError::Protected(Protection::Annotation)
    ));
    assert_eq!(protected_edit.staged(), &protected);

    assert!(Protection::Off.allows(EditKind::Content));
    assert!(Protection::RevisionMark.allows(EditKind::Revision));
    assert!(!Protection::RevisionMark.allows(EditKind::Content));
    assert!(Protection::Annotation.allows(EditKind::Annotation));
    assert!(Protection::Form.allows(EditKind::FormField));
}

#[test]
fn failed_removal_does_not_change_a_single_recipient_snapshot() {
    let mut value = route(Protection::Off);
    value.recipients.truncate(1);
    let snapshot = Snapshot::new(value).unwrap();
    let mut edit = snapshot.edit();
    assert!(edit.remove_recipient(RecipientSelector::Current).is_err());
    assert_eq!(edit.staged(), &snapshot);
}

fn parse_bytes_from_fib(bytes: &[u8]) -> Metadata {
    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    let word = ole.open_stream(&["WordDocument"]).unwrap();
    let fib = FileInformationBlock::parse(&word).unwrap();
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table = ole.open_stream(&[table_name]).unwrap();
    parse(&fib, &table, &SourceLimits::default())
        .unwrap()
        .unwrap()
}
