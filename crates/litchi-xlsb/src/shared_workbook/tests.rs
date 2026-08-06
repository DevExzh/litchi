use super::model::{
    Catalog, Guid, Header, Info, RawRecord, RevisionHeaders, RevisionLog, ShortDateTime, User,
    UserNames,
};
use super::{apply, read, store};
use crate::raw::{Kind, Records, Writer};
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::{BlobPart, OpcPackage, PackURI};

const FUTURE_KIND: u16 = 0x3fff;

#[test]
fn snapshot_noop_preserves_exact_owner_bytes() {
    let (mut package, _) = fixture();
    let before = owner_image(&package);
    let snapshot = read(&package).unwrap();
    let commit = snapshot.edit().commit().unwrap();
    assert!(commit.patch().is_empty());
    let applied = apply(&mut package, commit.patch()).unwrap();
    assert_eq!(applied, snapshot);
    assert_eq!(owner_image(&package), before);
}

#[test]
fn typed_metadata_edit_preserves_unknown_revision_header_records() {
    let (mut package, header_part) = fixture();
    let mut bytes = package.get_part(&header_part).unwrap().blob().to_vec();
    let mut writer = Writer::new(&mut bytes);
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
    let mut edit = snapshot.edit();
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
    let reparsed = Records::new(package.get_part(&header_part).unwrap().blob())
        .map(Result::unwrap)
        .collect::<Vec<_>>();
    assert_eq!(reparsed.last().unwrap().kind().get(), FUTURE_KIND);
}

#[test]
fn user_crud_and_stale_patches_are_source_checked_and_atomic() {
    let (mut package, _) = fixture();
    let snapshot = read(&package).unwrap();
    let mut edit = snapshot.edit();
    let mut user = snapshot.users().unwrap().users[0].clone();
    user.name = "Alicia".to_string();
    edit.upsert_user(user).unwrap();
    let commit = edit.commit().unwrap();

    let applied = apply(&mut package, commit.patch()).unwrap();
    assert_eq!(applied.users().unwrap().users[0].name, "Alicia");
    let before_stale_apply = owner_image(&package);
    assert!(apply(&mut package, commit.patch()).is_err());
    assert_eq!(owner_image(&package), before_stale_apply);

    let mut invalid = applied.edit();
    let before = invalid.catalog().clone();
    let mut info = applied.headers().unwrap().info.clone();
    info.version = 0;
    assert!(invalid.set_info(info).is_err());
    assert_eq!(invalid.catalog(), &before);
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
    assert!(super::codec::parse_log(&[0x80]).is_err());

    let mut users = Vec::new();
    let mut writer = Writer::new(&mut users);
    writer.write_record(Kind::new(401).unwrap(), &[]).unwrap();
    assert!(super::codec::parse_users(&users).is_err());
}

fn fixture() -> (OpcPackage, PackURI) {
    let mut package = OpcPackage::new();
    let workbook = PackURI::new("/xl/workbook.bin").unwrap();
    package.add_part(Box::new(BlobPart::new(
        workbook,
        content_type::XLSB_BIN.to_string(),
        Vec::new(),
    )));
    package
        .rels_mut()
        .get_or_add(relationship_type::OFFICE_DOCUMENT, "xl/workbook.bin");

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
    store(&mut package, &catalog).unwrap();
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
        .iter_parts()
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
