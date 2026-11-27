use super::model::{
    Catalog, Guid, Header, Info, RawRecord, RevisionHeaders, RevisionLog, ShortDateTime, User,
    UserNames,
};
use super::package::store;
use super::{Edit, Patch, Snapshot};
use crate::raw::{Encoder, Kind, Records};
use litchi_opc::{OpcPackage, PackURI};

fn read(package: &OpcPackage) -> super::Result<Snapshot> {
    Snapshot::read(package, crate::raw::Limits::DEFAULT)
}

fn start_edit(package: &OpcPackage) -> Edit {
    crate::Package::from_opc(package.clone())
        .unwrap()
        .edit_shared_workbook()
        .unwrap()
}

fn apply(package: &mut OpcPackage, patch: &Patch) -> super::Result<Snapshot> {
    let published = crate::Package::from_opc(package.clone())?
        .apply_shared_workbook(patch)?
        .into_package();
    *package = published.into_opc();
    Snapshot::read(package, crate::raw::Limits::DEFAULT)
}

const FUTURE_KIND: u16 = 0x3fff;

#[test]
fn snapshot_noop_preserves_exact_owner_bytes() {
    let (mut package, _) = fixture();
    let before = owner_image(&package);
    let snapshot = read(&package).unwrap();
    let commit = start_edit(&package).commit().unwrap();
    assert!(commit.patch().is_noop());
    let applied = apply(&mut package, commit.patch()).unwrap();
    assert_eq!(applied, snapshot);
    assert_eq!(owner_image(&package), before);
}

#[test]
fn typed_metadata_edit_preserves_unknown_revision_header_records() {
    let (mut package, header_part) = fixture();
    let mut bytes = package.get_part(&header_part).unwrap().blob().to_vec();
    let mut writer = Encoder::new(&mut bytes, crate::raw::Limits::DEFAULT);
    writer
        .write_record(Kind::new(FUTURE_KIND).unwrap(), b"future-header")
        .unwrap();
    package.get_part_mut(&header_part).unwrap().set_blob(bytes);

    let snapshot = read(&package).unwrap();
    assert!(
        snapshot
            .headers()
            .unwrap()
            .raw_records()
            .iter()
            .any(|record| record.kind == FUTURE_KIND)
    );
    let mut edit = start_edit(&package);
    let mut info = snapshot.headers().unwrap().info.clone();
    info.revision_id = 7;
    edit.set_info(info).unwrap();
    let commit = edit.commit().unwrap();
    let applied = apply(&mut package, commit.patch()).unwrap();

    assert_eq!(applied.headers().unwrap().info.revision_id, 7);
    assert!(
        applied
            .headers()
            .unwrap()
            .raw_records()
            .iter()
            .any(|record| record.kind == FUTURE_KIND && record.payload == b"future-header")
    );
    let reparsed = Records::new(
        package.get_part(&header_part).unwrap().blob(),
        crate::raw::Limits::DEFAULT,
    )
    .map(Result::unwrap)
    .collect::<Vec<_>>();
    assert_eq!(reparsed.last().unwrap().kind().get(), FUTURE_KIND);
}

#[test]
fn user_crud_and_stale_patches_are_source_checked_and_atomic() {
    let (mut package, _) = fixture();
    let snapshot = read(&package).unwrap();
    let mut edit = start_edit(&package);
    let mut user = snapshot.users().unwrap().users[0].clone();
    user.name = "Alicia".to_string();
    edit.upsert_user(user).unwrap();
    let commit = edit.commit().unwrap();

    let applied = apply(&mut package, commit.patch()).unwrap();
    assert_eq!(applied.users().unwrap().users[0].name, "Alicia");
    let before_stale_apply = owner_image(&package);
    let error = apply(&mut package, commit.patch()).expect_err("stale patch");
    assert!(matches!(
        error,
        crate::package::PackageError::PatchConflict { .. }
    ));
    assert_eq!(owner_image(&package), before_stale_apply);

    let mut invalid = start_edit(&package);
    let before = invalid.catalog().unwrap().clone();
    let mut info = applied.headers().unwrap().info.clone();
    info.version = 0;
    assert!(invalid.set_info(info).is_err());
    assert_eq!(invalid.catalog().unwrap(), &before);
}

#[test]
fn opaque_revision_records_are_borrowed_without_replay() {
    let (package, _) = fixture();
    let snapshot = read(&package).unwrap();
    let records = snapshot.logs()[0].views().collect::<Vec<_>>();
    assert_eq!(records.len(), 1);
    assert_eq!(records[0].kind, FUTURE_KIND);
    assert_eq!(records[0].payload, b"opaque-log");
    assert!(records[0].envelope.is_none());
}

#[test]
fn malformed_biff12_records_are_rejected_without_partial_metadata() {
    assert!(super::codec::parse_log(&[0x80], crate::raw::Limits::DEFAULT).is_err());

    let mut users = Vec::new();
    let mut writer = Encoder::new(&mut users, crate::raw::Limits::DEFAULT);
    writer.write_record(Kind::new(401).unwrap(), &[]).unwrap();
    assert!(super::codec::parse_users(&users, crate::raw::Limits::DEFAULT).is_err());
}

fn fixture() -> (OpcPackage, PackURI) {
    let mut package = crate::Package::new().unwrap().into_opc();

    let guid = Guid::from_bytes([1; 16]);
    let date = ShortDateTime {
        year: 2024,
        month: 1,
        day: 1,
        hour: 9,
        minute: 30,
        second: 0,
        weekday: 1,
    };
    let catalog = Catalog {
        users: Some(UserNames::new(vec![User {
            id: 1,
            guid,
            opened_at: date,
            name: "Alice".to_string(),
        }])),
        headers: Some(RevisionHeaders::new(
            Info {
                guid,
                root_guid: guid,
                revision_id: 0,
                version: 1,
                has_revisions: false,
                no_revision_history: true,
                protected: false,
                revision_history_interval: 0,
            },
            vec![Header {
                guid,
                saved_at: date,
                next_sheet_id: 0xffff,
                revision_min: 0,
                revision_max: 0,
                user_name: "Alice".to_string(),
                relationship_id: "rIdLog".to_string(),
                sheet_ids: vec![1],
                reviewed: Vec::new(),
            }],
        )),
        logs: vec![RevisionLog::new(vec![RawRecord::new(
            FUTURE_KIND,
            b"opaque-log".to_vec(),
        )])],
    };
    store(&mut package, &catalog, crate::raw::Limits::DEFAULT).unwrap();
    let headers = package
        .iter_parts()
        .find(|part| part.content_type() == super::package::HEADERS_CONTENT_TYPE)
        .unwrap()
        .partname()
        .clone();
    (package, headers)
}

fn owner_image(package: &OpcPackage) -> Vec<(String, Vec<u8>)> {
    let mut image = package
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
        .filter(|part| {
            matches!(
                part.content_type(),
                super::package::USERS_CONTENT_TYPE
                    | super::package::HEADERS_CONTENT_TYPE
                    | super::package::LOG_CONTENT_TYPE
            )
        })
        .map(|part| (part.partname().to_string(), part.blob().to_vec()))
        .collect::<Vec<_>>();
    image.sort_by(|left, right| left.0.cmp(&right.0));
    image
}

fn record_bytes(records: &[RawRecord]) -> Vec<u8> {
    let mut bytes = Vec::new();
    let mut writer = Encoder::new(&mut bytes, crate::raw::Limits::DEFAULT);
    for record in records {
        writer
            .write_record(Kind::new(record.kind).unwrap(), &record.payload)
            .unwrap();
    }
    bytes
}

#[test]
fn user_wire_order_matches_count_begin_users_and_end() {
    let mut users = UserNames::new(vec![]);
    bind_test_users(&mut users);
    let bytes = super::codec::write_users(&users, crate::raw::Limits::DEFAULT).unwrap();
    assert_eq!(bytes, [0x8f, 3, 2, 0, 0, 0x91, 3, 0, 0x93, 3, 0]);
    assert!(
        super::codec::parse_users(&bytes, crate::raw::Limits::DEFAULT)
            .unwrap()
            .users
            .is_empty()
    );
    for records in [
        vec![
            RawRecord::new(401, vec![]),
            RawRecord::new(399, vec![0, 0]),
            RawRecord::new(403, vec![]),
        ],
        vec![
            RawRecord::new(399, vec![0, 0]),
            RawRecord::new(403, vec![]),
            RawRecord::new(401, vec![]),
        ],
        vec![
            RawRecord::new(399, vec![0, 0]),
            RawRecord::new(401, vec![]),
            RawRecord::new(403, vec![0]),
        ],
        vec![
            RawRecord::new(399, vec![0, 0]),
            RawRecord::new(401, vec![]),
            RawRecord::new(403, vec![]),
            RawRecord::new(403, vec![]),
        ],
    ] {
        assert!(
            super::codec::parse_users(&record_bytes(&records), crate::raw::Limits::DEFAULT)
                .is_err()
        );
    }
}

#[test]
fn new_users_stay_before_the_end_marker_and_keep_opaque_neighbors() {
    let input = record_bytes(&[
        RawRecord::new(399, vec![0, 0]),
        RawRecord::new(401, vec![]),
        RawRecord::new(FUTURE_KIND, b"before-end".to_vec()),
        RawRecord::new(403, vec![]),
        RawRecord::new(FUTURE_KIND, b"after-end".to_vec()),
    ]);
    let mut users = super::codec::parse_users(&input, crate::raw::Limits::DEFAULT).unwrap();
    bind_test_users(&mut users);
    users.users.push(User {
        id: 1,
        guid: Guid::from_bytes([1; 16]),
        opened_at: ShortDateTime::default(),
        name: "Native derivative".into(),
    });
    let bytes = super::codec::write_users(&users, crate::raw::Limits::DEFAULT).unwrap();
    let records = Records::new(&bytes, crate::raw::Limits::DEFAULT)
        .map(Result::unwrap)
        .collect::<Vec<_>>();
    assert_eq!(
        records.iter().map(|r| r.kind().get()).collect::<Vec<_>>(),
        [399, 401, FUTURE_KIND, 400, 403, FUTURE_KIND]
    );
    assert_eq!(records[2].payload(), b"before-end");
    assert_eq!(records[5].payload(), b"after-end");
    let parsed = super::codec::parse_users(&bytes, crate::raw::Limits::DEFAULT).unwrap();
    assert_eq!(parsed.users, users.users);
    let mut bad = parsed.raw_records().to_vec();
    bad.swap(3, 4);
    assert!(super::codec::parse_users(&record_bytes(&bad), crate::raw::Limits::DEFAULT).is_err());
}

#[test]
fn new_headers_stay_before_the_end_marker_with_opaque_tail_retention() {
    let (package, _) = fixture();
    let snapshot = read(&package).unwrap();
    let mut headers = snapshot.headers().unwrap().clone();
    let mut source = headers.raw_records().to_vec();
    assert_eq!(source.last().unwrap().kind, 403);
    source.push(RawRecord::new(FUTURE_KIND, b"after-end".to_vec()));
    headers =
        super::codec::parse_headers(&record_bytes(&source), crate::raw::Limits::DEFAULT).unwrap();
    let mut added = headers.headers[0].clone();
    added.guid = Guid::from_bytes([2; 16]);
    added.relationship_id = "rIdSecond".into();
    headers.relationship_id = "rIdHeaders".into();
    headers.part_name = "/xl/revisions/revisionHeaders.bin".into();
    headers.info.guid = added.guid;
    headers.headers.push(added);
    let bytes = super::codec::write_headers(&headers, crate::raw::Limits::DEFAULT).unwrap();
    let parsed = super::codec::parse_headers(&bytes, crate::raw::Limits::DEFAULT).unwrap();
    assert_eq!(parsed.headers, headers.headers);
    assert_eq!(
        parsed
            .raw_records()
            .iter()
            .map(|r| r.kind)
            .collect::<Vec<_>>(),
        [398, 411, 411, 403, FUTURE_KIND]
    );
    assert_eq!(parsed.raw_records().last().unwrap().payload, b"after-end");
    let mut bad = parsed.raw_records().to_vec();
    bad.swap(2, 3);
    assert!(super::codec::parse_headers(&record_bytes(&bad), crate::raw::Limits::DEFAULT).is_err());
}

#[test]
fn revision_envelopes_use_record_ids_not_specification_section_numbers() {
    // Only the common RRd envelope is projected; record-specific bodies stay opaque.
    let mut prefix = vec![0; 14];
    prefix[4..8].copy_from_slice(&42u32.to_le_bytes());
    prefix[8..10].copy_from_slice(&8u16.to_le_bytes());
    prefix[10..12].copy_from_slice(&3u16.to_le_bytes());
    prefix[12..14].copy_from_slice(&2u16.to_le_bytes());
    for kind in [405, 407, 409, 412, 413, 414, 415, 416, 417, 418, 419, 421] {
        let r = super::codec::revision_envelope(kind, &prefix).unwrap();
        assert_eq!(r.revision_id, 42);
        assert_eq!(r.revision_type, 8);
        assert!(r.accepted && r.undo_action);
        assert_eq!(r.sheet_id, 2);
        assert!(super::codec::revision_envelope(kind, &prefix[..13]).is_none());
    }
    for kind in [
        771,
        772,
        773,
        774,
        779,
        781,
        782,
        783,
        784,
        785,
        787,
        788,
        400,
        403,
        FUTURE_KIND,
    ] {
        assert!(super::codec::revision_envelope(kind, &prefix).is_none());
    }
}

fn bind_test_users(users: &mut UserNames) {
    users.relationship_id = "rIdUsers".into();
    users.part_name = "/xl/revisions/userNames.bin".into();
}

#[test]
fn revision_append_and_remove_maintain_header_pointers_atomically() {
    let (package, _) = fixture();
    let source = read(&package).unwrap();
    let mut header = source.headers().unwrap().headers[0].clone();
    let guid = Guid::from_bytes([2; 16]);
    header.guid = guid;
    header.relationship_id = "rIdSecond".into();
    let mut log = source.logs()[0].clone();
    log.relationship_id = "rIdSecond".into();
    log.part_name = "/xl/revisions/revisionLog2.bin".into();
    let mut edit = start_edit(&package);
    edit.push_revision(header, log).unwrap();
    assert_eq!(
        edit.catalog().unwrap().headers.as_ref().unwrap().info.guid,
        guid
    );
    let before = edit.catalog().unwrap().clone();
    // The first header still owns the existing user, so removing it refuses.
    assert!(edit.remove_revision(0).is_err());
    assert_eq!(edit.catalog().unwrap(), &before);
    let mut info = before.headers.as_ref().unwrap().info.clone();
    info.root_guid = guid;
    edit.set_info(info).unwrap();
    edit.remove_revision(1).unwrap().unwrap();
    assert_eq!(edit.catalog().unwrap(), source.catalog());
    assert!(edit.remove_revision(0).is_err());
    assert_eq!(edit.catalog().unwrap(), source.catalog());
}
