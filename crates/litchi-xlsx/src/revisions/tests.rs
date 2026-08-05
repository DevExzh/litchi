use super::model::{MAX_DEPTH, MAX_PART_BYTES, NS, USERS_REL};
use super::*;
use litchi_opc::{BlobPart, OpcPackage, PackURI};

const LO: &[u8] = include_bytes!(
    "../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/track-changes/simple-cell-changes.xlsx"
);
const POI: &[u8] = include_bytes!(
    "../../../../test-data/poi/test-data/spreadsheet/workbookProtection_workbook_revision_protected.xlsx"
);
fn guid(n: u8) -> String {
    format!("{{00000000-0000-0000-0000-{n:012X}}}")
}
fn log() -> RevisionLog {
    RevisionLog {
        records: vec![RevisionRecord {
            kind: RevisionRecordKind::CellChange,
            revision_id: Some(1),
            sheet_id: Some(1),
            attributes: vec![],
            children: vec![RevisionXmlElement {
                name: "nc".into(),
                attributes: vec![RevisionAttribute {
                    namespace: RevisionAttributeNamespace::Unqualified,
                    name: "r".into(),
                    value: "A1".into(),
                }],
                children: vec![RevisionXmlElement {
                    name: "v".into(),
                    attributes: vec![],
                    children: vec![],
                    text: "=1+1".into(),
                }],
                text: String::new(),
            }],
            text: String::new(),
        }],
    }
}
fn value() -> Revisions {
    let h = RevisionHeader {
        guid: guid(2),
        date_time: "2026-07-17T12:00:00Z".into(),
        max_sheet_id: 2,
        user_name: "Reviewer".into(),
        relationship_id: "rId1".into(),
        min_revision_id: Some(1),
        max_revision_id: Some(1),
        sheet_ids: vec![1],
        trailing_elements: vec![],
    };
    Revisions {
        users_relationship_id: "rId2".into(),
        users_part_name: "/xl/revisions/userNames.xml".into(),
        headers_relationship_id: "rId3".into(),
        headers_part_name: "/xl/revisions/revisionHeaders.xml".into(),
        users: RevisionUsers::default(),
        headers: RevisionHeaders {
            properties: RevisionHeaderProperties {
                guid: h.guid.clone(),
                disk_revisions: Some(true),
                ..Default::default()
            },
            headers: vec![h],
        },
        logs: vec![RevisionLogPart {
            relationship_id: "rId1".into(),
            part_name: "/xl/revisions/revisionLog1.xml".into(),
            log: log(),
        }],
    }
}
fn package() -> OpcPackage {
    let mut p = OpcPackage::new();
    p.rels_mut().add_relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument".into(),
        "xl/workbook.xml".into(),
        "rId1".into(),
        false,
    );
    p.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/workbook.xml").unwrap(),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),
        Vec::new(),
    )));
    p
}
#[test]
fn loads_libreoffice_and_poi_reference_packages() {
    let a = load_workbook_revisions(&OpcPackage::from_bytes(LO).unwrap())
        .unwrap()
        .unwrap();
    assert_eq!(a.logs.len(), 5);
    assert!(
        a.logs
            .iter()
            .flat_map(|l| &l.log.records)
            .any(|r| r.kind == RevisionRecordKind::RowColumn)
    );
    assert!(
        a.logs
            .iter()
            .flat_map(|l| &l.log.records)
            .any(|r| r.kind == RevisionRecordKind::CustomView)
    );
    let b = load_workbook_revisions(&OpcPackage::from_bytes(POI).unwrap())
        .unwrap()
        .unwrap();
    assert_eq!(b.logs.len(), 3);
    assert_eq!(b.logs.iter().map(|l| l.log.records.len()).sum::<usize>(), 2);
}
#[test]
fn strict_writers_are_deterministic_and_round_trip() {
    let v = value();
    let u = write_revision_users(&v.users, RevisionConformance::Strict).unwrap();
    let h = write_revision_headers(&v.headers, RevisionConformance::Strict).unwrap();
    let l = write_revision_log(&v.logs[0].log, RevisionConformance::Strict).unwrap();
    assert_eq!(
        u,
        write_revision_users(&v.users, RevisionConformance::Strict).unwrap()
    );
    assert_eq!(parse_revision_users(&u).unwrap(), v.users);
    assert_eq!(parse_revision_headers(&h).unwrap(), v.headers);
    assert_eq!(parse_revision_log(&l).unwrap(), v.logs[0].log);
}
#[test]
fn mce_fallback() {
    let x = format!(
        r#"<revisions xmlns="{NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:z="urn:z" mc:Ignorable="z"><mc:AlternateContent><mc:Choice Requires="z"><z:x/></mc:Choice><mc:Fallback><rcv guid="{}" action="add"/></mc:Fallback></mc:AlternateContent></revisions>"#,
        guid(3)
    );
    assert_eq!(
        parse_revision_log(x.as_bytes()).unwrap().records[0].kind,
        RevisionRecordKind::CustomView
    );
}
#[test]
fn package_writer_round_trip() {
    let mut p = package();
    let v = value();
    store_workbook_revisions(&mut p, &v, RevisionConformance::Strict).unwrap();
    assert_eq!(load_workbook_revisions(&p).unwrap().unwrap(), v);
}
#[test]
fn malformed_and_caps() {
    for x in [
        format!(r#"<users xmlns="{NS}" count="2"/>"#),
        format!(r#"<headers xmlns="{NS}" guid="bad"/>"#),
        format!(r#"<revisions xmlns="{NS}"><bad/></revisions>"#),
        format!(r#"<!DOCTYPE x><revisions xmlns="{NS}"/>"#),
    ] {
        assert!(if x.contains("<users") {
            parse_revision_users(x.as_bytes()).is_err()
        } else if x.contains("<headers") {
            parse_revision_headers(x.as_bytes()).is_err()
        } else {
            parse_revision_log(x.as_bytes()).is_err()
        });
    }
    assert!(parse_revision_log(&vec![b' '; MAX_PART_BYTES + 1]).is_err());
    let deep = format!(
        r#"<revisions xmlns="{NS}"><rcc>{}{}</rcc></revisions>"#,
        "<x>".repeat(MAX_DEPTH),
        "</x>".repeat(MAX_DEPTH)
    );
    assert!(parse_revision_log(deep.as_bytes()).is_err());
}
#[test]
fn graph_and_reference_errors() {
    let mut v = value();
    v.logs[0].log.records[0].sheet_id = Some(9);
    assert!(
        store_workbook_revisions(&mut package(), &v, RevisionConformance::Transitional).is_err()
    );
    let mut p = package();
    p.get_part_mut(&PackURI::new("/xl/workbook.xml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            USERS_REL.into(),
            "https://invalid.example/users".into(),
            "rId2".into(),
            true,
        );
    assert!(load_workbook_revisions(&p).is_err());
}
