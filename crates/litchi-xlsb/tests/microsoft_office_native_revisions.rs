//! Native BIFF12 revision metadata; no revision replay or collaboration.
use litchi_core::sheet::{Cell as _, CellValue, WorkbookTrait};
use litchi_xlsb::{
    Package,
    shared_workbook::{Catalog, Guid, User},
};
use std::collections::BTreeMap;
fn native() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel/excel-revisions.xlsb"),
    )
    .unwrap()
}
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    r.member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = r.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
fn changed(a: &[u8], b: &[u8]) -> Vec<String> {
    let a = members(a);
    let b = members(b);
    let names: std::collections::BTreeSet<_> = a.keys().chain(b.keys()).collect();
    names
        .into_iter()
        .filter(|n| a.get(*n) != b.get(*n))
        .cloned()
        .collect()
}
fn graph_elements(xml: &[u8]) -> Vec<(String, BTreeMap<String, String>)> {
    use quick_xml::events::Event;
    let mut r = quick_xml::Reader::from_reader(xml);
    let mut out = Vec::new();
    loop {
        match r.read_event().unwrap() {
            Event::Start(e) | Event::Empty(e) => {
                let attrs = e
                    .attributes()
                    .map(|a| {
                        let a = a.unwrap();
                        (
                            String::from_utf8(a.key.as_ref().to_vec()).unwrap(),
                            a.decoded_and_normalized_value(
                                quick_xml::XmlVersion::Explicit1_0,
                                r.decoder(),
                            )
                            .unwrap()
                            .into_owned(),
                        )
                    })
                    .collect();
                out.push((
                    String::from_utf8(e.name().as_ref().to_vec()).unwrap(),
                    attrs,
                ));
            },
            Event::Eof => break,
            Event::Decl(_) | Event::End(_) => {},
            Event::Text(t) => assert!(t.as_ref().iter().all(u8::is_ascii_whitespace)),
            _ => panic!("unexpected graph markup"),
        }
    }
    out.sort();
    out
}
fn content_types(parts: &BTreeMap<String, Vec<u8>>) -> BTreeMap<String, String> {
    let mut defaults = BTreeMap::new();
    let mut overrides = BTreeMap::new();
    for (name, attrs) in graph_elements(&parts["[Content_Types].xml"]) {
        match name.as_str() {
            "Default" => {
                assert_eq!(attrs.len(), 2);
                assert!(
                    defaults
                        .insert(attrs["Extension"].clone(), attrs["ContentType"].clone())
                        .is_none()
                );
            },
            "Override" => {
                assert_eq!(attrs.len(), 2);
                assert!(
                    overrides
                        .insert(attrs["PartName"].clone(), attrs["ContentType"].clone())
                        .is_none()
                );
            },
            "Types" => assert_eq!(
                attrs,
                BTreeMap::from([(
                    "xmlns".into(),
                    "http://schemas.openxmlformats.org/package/2006/content-types".into()
                )])
            ),
            _ => panic!("unexpected content-type element: {name}"),
        }
    }
    parts
        .keys()
        .filter(|n| *n != "[Content_Types].xml")
        .map(|n| {
            let kind = overrides
                .get(&format!("/{n}"))
                .or_else(|| defaults.get(n.rsplit('.').next().unwrap()))
                .unwrap();
            (n.clone(), kind.clone())
        })
        .collect()
}
fn assert_same_payloads_and_graph(before: &[u8], after: &[u8]) {
    let a = members(before);
    let b = members(after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    // A default and per-part overrides can describe the same OPC typing.
    // Check every member's resolved type, not that serialization choice.
    assert_eq!(content_types(&a), content_types(&b));
    for (n, data) in a {
        if n == "[Content_Types].xml" {
            continue;
        }
        if n.ends_with(".rels") {
            assert_eq!(graph_elements(&data), graph_elements(&b[&n]), "{n}");
        } else {
            assert_eq!(data, b[&n], "{n}");
        }
    }
}

fn save(p: &Package) -> Package {
    let f = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
    p.save(f.path()).unwrap();
    Package::open(f.path()).unwrap()
}
fn assert_unrelated(before: &[u8], after: &[u8]) {
    let a = members(before);
    let b = members(after);
    let keep = |n: &String| {
        !n.starts_with("xl/revisions/")
            && n != "[Content_Types].xml"
            && n != "xl/_rels/workbook.bin.rels"
    };
    let ak = a.keys().filter(|n| keep(n)).collect::<Vec<_>>();
    let bk = b.keys().filter(|n| keep(n)).collect::<Vec<_>>();
    assert_eq!(ak, bk);
    for n in ak {
        assert!(a[n] == b[n], "changed unrelated {n}");
    }
}
fn values(p: &Package) {
    let b = p.workbook().unwrap();
    assert_eq!(b.worksheet_names(), ["Sheet1", "Review_014"]);
    let s = b.worksheet(0).unwrap();
    assert_eq!(s.get_cell(1, 1).unwrap().value(), &CellValue::Int(25));
    let CellValue::Formula {
        cached_value: Some(v),
        ..
    } = s.get_cell(2, 1).unwrap().value()
    else {
        panic!("saved formula")
    };
    assert_eq!(**v, CellValue::Float(75.));
    assert_eq!(
        b.worksheet(1).unwrap().get_cell(0, 0).unwrap().value(),
        &CellValue::String("NEW_SHEET_014".into())
    );
}
fn kinds(data: &[u8]) -> Vec<u16> {
    litchi_xlsb::raw::Records::new(data, litchi_xlsb::raw::Limits::DEFAULT)
        .map(|r| r.unwrap().kind().get())
        .collect()
}
fn guid(hex: &str) -> Guid {
    let bytes = hex
        .as_bytes()
        .chunks_exact(2)
        .map(|c| u8::from_str_radix(std::str::from_utf8(c).unwrap(), 16).unwrap())
        .collect::<Vec<_>>();
    Guid::from_bytes(bytes.try_into().unwrap())
}

#[test]
fn native_revision_catalog_retains_headers_user_framing_and_six_record_envelopes() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let s = p.shared_workbook().unwrap();
    let parts = members(&bytes);
    values(&p);
    let users = s.users().unwrap();
    assert!(users.users.is_empty());
    assert_eq!(users.relationship_id, "rId7");
    assert_eq!(users.part_name, "/xl/revisions/userNames.bin");
    assert_eq!(kinds(&parts["xl/revisions/userNames.bin"]), [399, 401, 403]);
    assert_eq!(parts["xl/revisions/userNames.bin"].len(), 11);
    let h = s.headers().unwrap();
    assert_eq!(h.relationship_id, "rId8");
    assert_eq!(h.headers.len(), 3);
    assert_eq!(h.info.revision_id, 6);
    assert_eq!(h.info.version, 1);
    assert!(h.info.has_revisions);
    assert!(!h.info.no_revision_history && !h.info.protected);
    assert_eq!(h.info.revision_history_interval, 30);
    let guids = [
        guid("8993e54a5a96472cb819cc750216fef0"),
        guid("4bcb55e22c8b4203935518ed86d66a4a"),
        guid("8291feaa3ada4ca2acb515d001b28196"),
    ];
    assert_eq!(h.info.guid, guids[2]);
    assert_eq!(h.info.root_guid, guids[2]);
    for (i, header) in h.headers.iter().enumerate() {
        assert_eq!(header.guid, guids[i]);
        assert_eq!(header.user_name, "Ryker Zhu");
        assert_eq!(header.relationship_id, format!("rId{}", i + 1));
        let d = header.saved_at;
        assert_eq!(
            (d.year, d.month, d.day, d.hour, d.weekday),
            (2026, 10, 15, 18, 4)
        );
        assert_eq!((d.minute, d.second), [(54, 12), (54, 52), (55, 15)][i]);
        assert_eq!(header.next_sheet_id, if i == 0 { 2 } else { 3 });
        assert_eq!(header.sheet_ids, if i == 0 { vec![1] } else { vec![1, 2] });
        assert_eq!(
            (header.revision_min, header.revision_max),
            if i == 1 { (1, 6) } else { (0, 0) }
        );
        assert!(header.reviewed.is_empty());
    }
    assert_eq!(s.logs().len(), 3);
    assert_eq!(
        s.logs().iter().map(|l| l.records.len()).collect::<Vec<_>>(),
        [1, 20, 3]
    );
    let envelopes = s.logs()[1]
        .views()
        .filter_map(|r| r.envelope)
        .collect::<Vec<_>>();
    assert_eq!(envelopes.len(), 6);
    for (i, e) in envelopes.iter().enumerate() {
        assert_eq!(e.revision_id, i as u32 + 1);
        assert_eq!(e.revision_type, if i == 3 { 5 } else { 8 });
        assert_eq!(e.sheet_id, if i < 3 { 1 } else { 2 });
        assert!(!e.accepted && !e.undo_action && !e.reserved_one && !e.reserved_two);
    }
    let views = s.logs()[2]
        .views()
        .filter_map(|r| r.envelope)
        .collect::<Vec<_>>();
    assert_eq!(
        views
            .iter()
            .map(|e| (e.revision_id, e.revision_type, e.sheet_id))
            .collect::<Vec<_>>(),
        [(0, 44, 65535), (0, 43, 65535)]
    );
    for (i, log) in s.logs().iter().enumerate() {
        assert_eq!(log.relationship_id, format!("rId{}", i + 1));
        assert_eq!(
            log.part_name,
            format!("/xl/revisions/revisionLog{}.bin", i + 1)
        );
        assert_eq!(log.records.last().unwrap().kind, 403);
    }
    for part in s.source_parts() {
        if part.part_name().starts_with("/xl/revisions/") {
            assert_eq!(
                part.bytes(),
                parts[part.part_name().trim_start_matches('/')]
            );
        }
    }
    assert_eq!(
        p.edit_shared_workbook()
            .unwrap()
            .commit()
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        bytes
    );
}

#[test]
fn native_current_user_crud_keeps_revision_parts_and_places_users_before_eof() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let s = p.shared_workbook().unwrap();
    let last = s.headers().unwrap().headers.last().unwrap();
    let user = User {
        id: 42,
        guid: last.guid,
        opened_at: last.saved_at,
        name: "Native reviewer".into(),
    };
    let mut e = p.edit_shared_workbook().unwrap();
    e.upsert_user(user.clone()).unwrap();
    let before = e.catalog().unwrap().clone();
    let mut invalid = user.clone();
    invalid.guid = Guid::from_bytes([0xee; 16]);
    assert!(e.upsert_user(invalid).is_err());
    assert_eq!(e.catalog().unwrap(), &before);
    let c = e.commit().unwrap();
    let after = save(c.package());
    let out = after.to_bytes().unwrap();
    values(&after);
    assert_eq!(
        after
            .shared_workbook()
            .unwrap()
            .users()
            .unwrap()
            .users
            .as_slice(),
        std::slice::from_ref(&user)
    );
    assert_eq!(
        changed(&bytes, &out),
        [
            "xl/revisions/_rels/revisionHeaders.bin.rels",
            "xl/revisions/userNames.bin"
        ]
    );
    assert_eq!(
        graph_elements(&members(&bytes)["xl/revisions/_rels/revisionHeaders.bin.rels"]),
        graph_elements(&members(&out)["xl/revisions/_rels/revisionHeaders.bin.rels"]),
    );
    assert_eq!(
        kinds(&members(&out)["xl/revisions/userNames.bin"]),
        [399, 401, 400, 403]
    );
    assert_same_payloads_and_graph(
        &out,
        &p.apply_shared_workbook(c.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
    );
    assert!(matches!(
        after.apply_shared_workbook(c.patch()),
        Err(litchi_xlsb::package::PackageError::PatchConflict { .. })
    ));
    assert_same_payloads_and_graph(
        &bytes,
        &after
            .apply_shared_workbook(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
    );
    let mut updated = user;
    updated.name = "Reviewer 名".into();
    let mut e = after.edit_shared_workbook().unwrap();
    e.upsert_user(updated.clone()).unwrap();
    let updated_commit = e.commit().unwrap();
    let updated = save(updated_commit.package());
    assert_eq!(
        updated.shared_workbook().unwrap().users().unwrap().users[0].name,
        "Reviewer 名"
    );
    assert_unrelated(&bytes, &updated.to_bytes().unwrap());
    let mut e = updated.edit_shared_workbook().unwrap();
    assert!(e.remove_user(42).unwrap().is_some());
    assert!(e.remove_user(42).unwrap().is_none());
    let removed = e.commit().unwrap();
    let empty = save(removed.package());
    assert!(
        empty
            .shared_workbook()
            .unwrap()
            .users()
            .unwrap()
            .users
            .is_empty()
    );
    assert_same_payloads_and_graph(&bytes, &empty.to_bytes().unwrap());
    assert_same_payloads_and_graph(
        &updated.to_bytes().unwrap(),
        &empty
            .apply_shared_workbook(&removed.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
    );
    assert_eq!(p.to_bytes().unwrap(), bytes);
}

#[test]
fn native_header_append_keeps_a_valid_latest_header_identity() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let s = p.shared_workbook().unwrap();
    let mut e = p.edit_shared_workbook().unwrap();
    let mut info = s.headers().unwrap().info.clone();
    info.revision_history_interval = 45;
    e.set_info(info.clone()).unwrap();
    let mut bad = info;
    bad.version = 0;
    assert!(e.set_info(bad).is_err());
    assert_eq!(
        e.catalog()
            .unwrap()
            .headers
            .as_ref()
            .unwrap()
            .info
            .revision_history_interval,
        45
    );
    let mut header = s.headers().unwrap().headers[1].clone();
    header.user_name = "Native header reviewer".into();
    e.set_revision(1, header.clone(), s.logs()[1].clone())
        .unwrap();
    let c = e.commit().unwrap();
    let edited = save(c.package());
    assert_eq!(
        changed(&bytes, &edited.to_bytes().unwrap()),
        [
            "xl/revisions/_rels/revisionHeaders.bin.rels",
            "xl/revisions/revisionHeaders.bin"
        ]
    );
    assert_eq!(
        graph_elements(&members(&bytes)["xl/revisions/_rels/revisionHeaders.bin.rels"]),
        graph_elements(
            &members(&edited.to_bytes().unwrap())["xl/revisions/_rels/revisionHeaders.bin.rels"]
        ),
    );
    assert_eq!(
        edited.shared_workbook().unwrap().headers().unwrap().headers[1].user_name,
        header.user_name
    );
    assert_unrelated(&bytes, &edited.to_bytes().unwrap());
    assert_same_payloads_and_graph(
        &bytes,
        &edited
            .apply_shared_workbook(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
    );
    let mut header = s.headers().unwrap().headers[2].clone();
    header.guid = Guid::from_bytes([0x42; 16]);
    header.relationship_id = "rId4".into();
    let mut log = s.logs()[0].clone();
    log.relationship_id = "rId4".into();
    log.part_name = "/xl/revisions/revisionLog4.bin".into();
    let mut e = p.edit_shared_workbook().unwrap();
    e.push_revision(header.clone(), log).unwrap();
    let c = e.commit().unwrap();
    let after = save(c.package());
    let snapshot = after.shared_workbook().unwrap();
    assert_eq!(snapshot.headers().unwrap().headers.len(), 4);
    assert_eq!(snapshot.headers().unwrap().info.guid, header.guid);
    assert_eq!(
        snapshot.headers().unwrap().info.root_guid,
        s.headers().unwrap().info.root_guid
    );
    assert_eq!(
        kinds(&members(&after.to_bytes().unwrap())["xl/revisions/revisionHeaders.bin"]),
        [398, 411, 411, 411, 411, 403]
    );
    values(&after);
    assert_unrelated(&bytes, &after.to_bytes().unwrap());
    let mut e = after.edit_shared_workbook().unwrap();
    e.remove_revision(3).unwrap().unwrap();
    let removed = e.commit().unwrap();
    let restored = save(removed.package());
    assert_eq!(
        restored
            .shared_workbook()
            .unwrap()
            .headers()
            .unwrap()
            .info
            .guid,
        s.headers().unwrap().info.guid
    );
    assert_same_payloads_and_graph(&bytes, &restored.to_bytes().unwrap());
    assert_same_payloads_and_graph(
        &after.to_bytes().unwrap(),
        &restored
            .apply_shared_workbook(&removed.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
    );
    let mut e = p.edit_shared_workbook().unwrap();
    e.remove_revision(2).unwrap().unwrap();
    let c = e.commit().unwrap();
    let removed_last = save(c.package());
    let h = removed_last
        .shared_workbook()
        .unwrap()
        .headers()
        .unwrap()
        .clone();
    assert_eq!(h.info.guid, s.headers().unwrap().headers[1].guid);
    assert_eq!(h.info.root_guid, h.info.guid);
    assert_unrelated(&bytes, &removed_last.to_bytes().unwrap());
    assert_eq!(p.to_bytes().unwrap(), bytes);
}

#[test]
fn native_revision_owner_removal_and_inverse_preserve_saved_workbook_values() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let original = p.shared_workbook().unwrap();
    let mut e = p.edit_shared_workbook().unwrap();
    e.replace(Catalog::empty()).unwrap();
    let c = e.commit().unwrap();
    let absent = save(c.package());
    let out = absent.to_bytes().unwrap();
    assert!(absent.shared_workbook().unwrap().is_empty());
    assert!(!members(&out).keys().any(|n| n.starts_with("xl/revisions/")));
    values(&absent);
    assert_unrelated(&bytes, &out);
    assert_same_payloads_and_graph(
        &bytes,
        &absent
            .apply_shared_workbook(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
    );
    let mut e = absent.edit_shared_workbook().unwrap();
    e.replace(original.catalog().clone()).unwrap();
    let recreated = e.commit().unwrap();
    let recreated = save(recreated.package());
    assert_eq!(
        recreated.shared_workbook().unwrap().catalog(),
        original.catalog()
    );
    assert_same_payloads_and_graph(&bytes, &recreated.to_bytes().unwrap());
    assert_eq!(p.to_bytes().unwrap(), bytes);
}
