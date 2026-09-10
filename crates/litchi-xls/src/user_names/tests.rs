use super::codec::{
    BC_USRS_RECORD_TYPE, C_USR_RECORD_TYPE, CB_USR_RECORD_TYPE, USR_CHK_RECORD_TYPE,
    USR_INFO_RECORD_TYPE, frame_record,
};
use super::{Limits, Snapshot, UserEntry};
use crate::revision_records::ShortDtr;
use std::sync::Arc;

fn opened_at() -> ShortDtr {
    ShortDtr::parse(USR_INFO_RECORD_TYPE, &[0xEA, 0x07, 1, 2, 3, 4, 5, 1]).unwrap()
}

fn stream_with_user() -> (Vec<u8>, [u8; 16]) {
    stream_with_user_named("Alice")
}

fn stream_with_user_named(name: &str) -> (Vec<u8>, [u8; 16]) {
    let guid = [0x11; 16];
    let mut user = UserEntry::new(7, guid, opened_at(), name)
        .unwrap()
        .with_unused(0xA5);
    user.string_flags = 0x02;
    let user_payload = user.to_payload().unwrap();
    let mut cbusr = vec![0u8; 512];
    cbusr[..2].copy_from_slice(&(user_payload.len() as u16).to_le_bytes());
    cbusr[510..].copy_from_slice(&[0x34, 0x12]);
    let mut stream = Vec::new();
    stream.extend_from_slice(&frame_record(C_USR_RECORD_TYPE, &1u16.to_le_bytes()).unwrap());
    stream
        .extend_from_slice(&frame_record(USR_CHK_RECORD_TYPE, &[0x00, 0x06, 0xF1, 0xF2]).unwrap());
    stream.extend_from_slice(&frame_record(CB_USR_RECORD_TYPE, &cbusr).unwrap());
    stream.extend_from_slice(&frame_record(BC_USRS_RECORD_TYPE, &9u16.to_le_bytes()).unwrap());
    stream.extend_from_slice(&frame_record(USR_INFO_RECORD_TYPE, &user_payload).unwrap());
    (stream, guid)
}

#[test]
fn timestamp_edit_preserves_nul_in_inert_user_name() {
    let (bytes, guid) = stream_with_user_named("A\0lice");
    let source = Snapshot::parse_with_revision_guids(&bytes, &[guid]).unwrap();
    assert_eq!(source.users()[0].user_name().as_bytes(), b"A\0lice");

    let mut edit = source.edit();
    edit.set_user_opened_at(
        0,
        ShortDtr::parse(USR_INFO_RECORD_TYPE, &[0xEA, 0x07, 2, 3, 4, 5, 6, 2]).unwrap(),
    )
    .unwrap();
    let target = edit.commit().unwrap().into_snapshot();
    assert_eq!(target.users()[0].user_name().as_bytes(), b"A\0lice");
}

#[test]
fn parses_user_names_and_preserves_reserved_fields() {
    let (bytes, guid) = stream_with_user();
    let snapshot = Snapshot::parse(&bytes).unwrap();
    assert_eq!(snapshot.user_count(), 1);
    assert_eq!(snapshot.users()[0].user_id(), 7);
    assert_eq!(snapshot.users()[0].guid(), &guid);
    assert_eq!(snapshot.users()[0].unused(), 0xA5);
    assert_eq!(snapshot.users()[0].string_flags(), 0x02);
    assert_eq!(snapshot.model().user_check().reserved(), 0xF2F1);
    assert_eq!(snapshot.model().briefcase_user_count(), 9);
    assert_eq!(snapshot.model().user_record_sizes()[255], 0x1234);
    assert_eq!(snapshot.finish(), bytes);
}

#[test]
fn metadata_edit_is_source_checked_and_reversible() {
    let (bytes, guid) = stream_with_user();
    let source = Snapshot::parse_with_revision_guids(&bytes, &[guid]).unwrap();
    let mut edit = source.edit();
    edit.set_user_name(0, "Élodie").unwrap();
    edit.set_user_opened_at(
        0,
        ShortDtr::parse(USR_INFO_RECORD_TYPE, &[0xEA, 0x07, 2, 3, 4, 5, 6, 2]).unwrap(),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.patch().apply(&source).unwrap(), *commit.snapshot());
    assert_eq!(commit.patch().revert(commit.snapshot()).unwrap(), source);
    assert_eq!(
        commit.patch().inverse().apply(commit.snapshot()).unwrap(),
        source
    );
    assert_eq!(commit.snapshot().users()[0].user_name(), "Élodie");
    assert_eq!(commit.snapshot().users()[0].string_flags(), 0x02);
    assert_eq!(commit.snapshot().model().user_record_sizes()[255], 0x1234);
    assert_eq!(commit.snapshot().model().user_check().reserved(), 0xF2F1);
}

#[test]
fn semantic_revert_is_an_exact_noop_and_invalid_edit_is_atomic() {
    let (bytes, guid) = stream_with_user();
    let source = Snapshot::parse_with_revision_guids(&bytes, &[guid]).unwrap();
    let mut edit = source.edit();
    edit.set_user_name(0, "Bob").unwrap();
    assert!(edit.is_changed());
    edit.set_user_name(0, "Alice").unwrap();
    assert!(!edit.is_changed());
    let noop = edit.commit().unwrap();
    assert!(!noop.changed());
    assert_eq!(noop.snapshot(), &source);

    let mut rejected = source.edit();
    assert!(rejected.set_user_name(0, "").is_err());
    assert!(!rejected.is_changed());
    assert_eq!(rejected.users()[0].user_name(), "Alice");
}

#[test]
fn collection_edit_updates_cusr_and_cbusr() {
    let (mut bytes, guid) = stream_with_user();
    // The normative reserved tail is zero for a collection-growing edit.
    bytes[18 + 510..18 + 512].fill(0);
    let source = Snapshot::parse_with_revision_guids(&bytes, &[guid]).unwrap();
    let another = UserEntry::new(8, guid, opened_at(), "Bob").unwrap();
    let mut edit = source.edit();
    edit.push(another.clone()).unwrap();
    let added = edit.commit().unwrap().into_snapshot();
    assert_eq!(added.user_count(), 2);
    assert_eq!(added.users()[1].user_name(), "Bob");

    let mut append_again = source.edit();
    append_again.push(another.clone()).unwrap();
    let append_commit = append_again.commit().unwrap();
    assert_eq!(
        append_commit
            .patch()
            .revert(append_commit.snapshot())
            .unwrap(),
        source
    );
    assert_eq!(
        append_commit
            .patch()
            .inverse()
            .apply(append_commit.snapshot())
            .unwrap(),
        source
    );

    let mut remove = added.edit();
    assert_eq!(remove.remove(0).unwrap().user_name(), "Alice");
    let removed = remove.commit().unwrap().into_snapshot();
    assert_eq!(removed.user_count(), 1);
    assert_eq!(removed.users()[0].user_name(), "Bob");
    assert_eq!(removed.model().user_record_sizes()[1], 0);
}

#[test]
fn insertion_requires_revision_guid_closure() {
    let (bytes, guid) = stream_with_user();
    let source = Snapshot::parse(&bytes).unwrap();
    let user = UserEntry::new(8, guid, opened_at(), "Bob").unwrap();
    assert!(source.edit().push(user).is_err());
}

#[test]
fn patch_rejects_same_bytes_with_a_different_guid_closure() {
    let (bytes, guid) = stream_with_user();
    let other_guid = [0x22; 16];
    let source = Snapshot::parse_with_revision_guids(&bytes, &[guid]).unwrap();
    let alternate = Snapshot::parse_with_revision_guids(&bytes, &[guid, other_guid]).unwrap();
    assert_ne!(alternate, source);

    let mut edit = source.edit();
    edit.set_user_name(0, "Bob").unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.patch().apply(&alternate).is_err());
}

#[test]
fn shared_guid_parse_precharges_source_and_guid_limits() {
    let (bytes, guid) = stream_with_user();
    let source = Arc::<[u8]>::from(bytes.clone().into_boxed_slice());
    let source_bound = Limits::default()
        .with_max_stream_bytes(bytes.len() - 1)
        .with_max_revision_guids(0);
    let error =
        Snapshot::parse_shared_with_revision_guids_and_limits(source, &[guid], source_bound)
            .unwrap_err();
    assert!(error.to_string().contains("User Names stream"));

    let source = Arc::<[u8]>::from(bytes.into_boxed_slice());
    let guid_bound = Limits::default().with_max_revision_guids(0);
    let error = Snapshot::parse_shared_with_revision_guids_and_limits(source, &[guid], guid_bound)
        .unwrap_err();
    assert!(error.to_string().contains("Revision Log contains"));
}

#[test]
fn duplicate_revision_guids_are_rejected_without_quadratic_lookup() {
    let (bytes, guid) = stream_with_user();
    let error = Snapshot::parse_with_revision_guids(&bytes, &[guid, guid]).unwrap_err();
    assert!(error.to_string().contains("duplicate RRDHead GUIDs"));
}

#[test]
fn malformed_count_and_size_are_rejected() {
    let (mut bytes, _) = stream_with_user();
    // CUsr declares two users while only one UsrInfo follows.
    bytes[4..6].copy_from_slice(&2u16.to_le_bytes());
    assert!(Snapshot::parse(&bytes).is_err());

    let (mut bytes, _) = stream_with_user();
    // CbUsr's first declared payload length no longer matches UsrInfo.
    let cbusr_payload_offset = 4 + 2 + 4 + 4 + 4;
    bytes[cbusr_payload_offset..cbusr_payload_offset + 2].copy_from_slice(&1u16.to_le_bytes());
    assert!(Snapshot::parse(&bytes).is_err());

    let (mut bytes, _) = stream_with_user();
    // UsrChk.version is restricted to the BIFF2-BIFF8 values in MS-XLS
    // 2.4.338; the reserved word remains opaque when the version is valid.
    let usr_chk_payload_offset = 4 + 2 + 4;
    bytes[usr_chk_payload_offset..usr_chk_payload_offset + 2]
        .copy_from_slice(&0x0700u16.to_le_bytes());
    assert!(Snapshot::parse(&bytes).is_err());

    let (bytes, _) = stream_with_user();
    assert!(
        Snapshot::parse_with_limits(
            &bytes,
            Limits::default().with_max_stream_bytes(bytes.len() - 1),
        )
        .is_err()
    );
    assert!(
        Snapshot::parse_with_limits(
            &bytes,
            Limits::default().with_max_stream_bytes(Limits::MAX_STREAM_BYTES + 1),
        )
        .is_err()
    );
}
