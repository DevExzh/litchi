//! Native SpreadsheetML revision graphs; changes stay inert.
use litchi_xlsx::{
    Cell, Package, Value,
    revisions::{
        RevisionAttribute, RevisionAttributeNamespace, RevisionRecordKind as Kind, RevisionUser,
        RevisionXmlElement, Revisions,
    },
};
use std::collections::BTreeMap;
const GUIDS: [&str; 3] = [
    "{4AE59389-965A-2C47-B819-CC750216FEF0}",
    "{E255CB4B-8B2C-0342-9355-18ED86D66A4A}",
    "{AAFE9182-DA3A-A24C-ACB5-15D001B28196}",
];
fn native() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel/excel-revisions.xlsx"),
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
    let f = tempfile::NamedTempFile::with_suffix(".xlsx").unwrap();
    p.save(f.path()).unwrap();
    Package::open(f.path()).unwrap()
}
fn model(p: &Package) -> Revisions {
    p.edit_revisions()
        .unwrap()
        .revisions()
        .unwrap()
        .unwrap()
        .clone()
}
// XML attribute order is not semantic. Keep record/child order and every
// namespace, value and text node, while comparing canonical owner recreation.
fn semantic(mut value: Revisions) -> Revisions {
    fn attributes(values: &mut [RevisionAttribute]) {
        values.sort_by(|a, b| {
            a.namespace
                .cmp(&b.namespace)
                .then(a.name.cmp(&b.name))
                .then(a.value.cmp(&b.value))
        });
    }
    fn elements(values: &mut [RevisionXmlElement]) {
        for value in values {
            attributes(&mut value.attributes);
            elements(&mut value.children);
        }
    }
    for user in &mut value.users.users {
        elements(&mut user.extension_elements);
    }
    for header in &mut value.headers.headers {
        elements(&mut header.trailing_elements);
    }
    for log in &mut value.logs {
        for record in &mut log.log.records {
            attributes(&mut record.attributes);
            elements(&mut record.children);
        }
    }
    value
}
fn attr<'a>(attrs: &'a [RevisionAttribute], name: &str) -> &'a str {
    &attrs
        .iter()
        .find(|a| a.namespace == RevisionAttributeNamespace::Unqualified && a.name == name)
        .unwrap()
        .value
}
fn node<'a>(nodes: &'a [RevisionXmlElement], path: &[&str]) -> &'a RevisionXmlElement {
    let first = nodes.iter().find(|n| n.name == path[0]).unwrap();
    if path.len() == 1 {
        first
    } else {
        node(&first.children, &path[1..])
    }
}
fn values(p: &Package) {
    let book = p.workbook();
    let sheet = book.sheet("Sheet1").unwrap().unwrap();
    assert!(book.sheet("Review_014").unwrap().is_some());
    assert!(book.sheet("History").unwrap().is_none());
    let b2 = sheet.cell("B2").unwrap();
    let Some(Cell::Value(Value::Number(v))) = b2.stored() else {
        panic!("B2")
    };
    assert_eq!(v.as_str(), "25");
    let b3 = sheet.cell("B3").unwrap();
    let Some(Cell::Formula(f)) = b3.stored() else {
        panic!("B3")
    };
    assert_eq!(f.text(), "B2*3");
    let Value::Number(v) = f.cached().unwrap().value() else {
        panic!("formula cache")
    };
    assert_eq!(v.as_str(), "75");
}
fn assert_unrelated(before: &[u8], after: &[u8]) {
    let a = members(before);
    let b = members(after);
    let keep = |n: &String| {
        !n.starts_with("xl/revisions/")
            && n != "[Content_Types].xml"
            && n != "xl/_rels/workbook.xml.rels"
    };
    let ak = a.keys().filter(|n| keep(n)).collect::<Vec<_>>();
    let bk = b.keys().filter(|n| keep(n)).collect::<Vec<_>>();
    assert_eq!(ak, bk);
    for n in ak {
        assert!(a[n] == b[n], "changed unrelated {n}");
    }
}

#[test]
fn native_revision_graph_retains_exact_headers_cells_formulas_and_custom_views() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let r = model(&p);
    values(&p);
    assert!(r.users.users.is_empty());
    assert_eq!(r.users_relationship_id, "rId8");
    assert_eq!(r.headers_relationship_id, "rId7");
    assert_eq!(r.users_part_name, "/xl/revisions/userNames.xml");
    assert_eq!(r.headers_part_name, "/xl/revisions/revisionHeaders.xml");
    assert_eq!(r.headers.properties.guid, GUIDS[2]);
    assert_eq!(r.headers.properties.revision_id, Some(6));
    assert_eq!(r.headers.properties.version, Some(3));
    assert_eq!(r.headers.properties.disk_revisions, Some(true));
    assert_eq!(r.headers.properties.last_guid, None);
    assert_eq!(r.headers.properties.track_revisions, None);
    assert_eq!(r.headers.headers.len(), 3);
    for (i, h) in r.headers.headers.iter().enumerate() {
        assert_eq!(h.guid, GUIDS[i]);
        assert_eq!(h.user_name, "Ryker Zhu");
        assert_eq!(
            h.date_time,
            [
                "2026-10-15T18:54:12",
                "2026-10-15T18:54:52",
                "2026-10-15T18:55:15"
            ][i]
        );
        assert_eq!(h.max_sheet_id, if i == 0 { 2 } else { 3 });
        assert_eq!(h.sheet_ids, if i == 0 { vec![1] } else { vec![1, 2] });
        assert_eq!(
            (h.min_revision_id, h.max_revision_id),
            if i == 1 {
                (Some(1), Some(6))
            } else {
                (None, None)
            }
        );
        assert_eq!(h.relationship_id, format!("rId{}", i + 1));
    }
    assert_eq!(
        r.logs
            .iter()
            .map(|l| l.log.records.len())
            .collect::<Vec<_>>(),
        [0, 6, 2]
    );
    let changes = &r.logs[1].log.records;
    for (i, c) in changes.iter().enumerate() {
        assert_eq!(c.revision_id, Some(i as u32 + 1));
        assert_eq!(
            c.kind,
            if i == 3 {
                Kind::InsertSheet
            } else {
                Kind::CellChange
            }
        );
        if i != 3 {
            assert_eq!(c.sheet_id, Some(if i < 3 { 1 } else { 2 }));
        }
    }
    for (i, cell, old, new, leaf) in [(0, "B2", "10", "25", "v"), (1, "B3", "B2*2", "B2*3", "f")] {
        assert_eq!(
            attr(&node(&changes[i].children, &["oc"]).attributes, "r"),
            cell
        );
        assert_eq!(node(&changes[i].children, &["oc", leaf]).text, old);
        assert_eq!(node(&changes[i].children, &["nc", leaf]).text, new);
    }
    assert_eq!(
        node(&changes[2].children, &["oc", "is", "t"]).text,
        "BASELINE_014"
    );
    assert_eq!(
        node(&changes[2].children, &["nc", "is", "t"]).text,
        "UPDATED_014"
    );
    assert_eq!(attr(&changes[3].attributes, "sheetId"), "2");
    assert_eq!(attr(&changes[3].attributes, "sheetPosition"), "1");
    assert_eq!(
        attr(&changes[3].attributes, "name"),
        "[excel-revisions.xlsx]Review_014"
    );
    for (i, cell, text) in [
        (4, "A1", "NEW_SHEET_014"),
        (5, "B1", "REVISION_CREATED_014"),
    ] {
        assert!(!changes[i].children.iter().any(|n| n.name == "oc"));
        assert_eq!(
            attr(&node(&changes[i].children, &["nc"]).attributes, "r"),
            cell
        );
        assert_eq!(node(&changes[i].children, &["nc", "is", "t"]).text, text);
    }
    for (c, action) in r.logs[2].log.records.iter().zip(["delete", "add"]) {
        assert_eq!(c.kind, Kind::CustomView);
        assert_eq!(c.revision_id, None);
        assert_eq!(attr(&c.attributes, "action"), action);
        assert_eq!(
            attr(&c.attributes, "guid"),
            "{DFC82956-C1E8-CD4D-A754-0ADAD5CAF49A}"
        );
    }
    let mut e = p.edit_revisions().unwrap();
    assert!(!e.replace(Some(r)).unwrap());
    assert_eq!(e.commit().unwrap().package().to_bytes().unwrap(), bytes);
}

#[test]
fn native_revision_user_and_header_edits_preserve_logs_values_and_staged_work() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let original = model(&p);
    let user = RevisionUser {
        guid: GUIDS[2].into(),
        name: "Native reviewer".into(),
        id: 42,
        date_time: "2026-10-15T18:55:15".into(),
        extension_elements: vec![],
    };
    let mut e = p.edit_revisions().unwrap();
    assert!(e.set_user(user.clone()).unwrap());
    let staged = e.revisions().unwrap().unwrap().clone();
    let mut bad = user.clone();
    bad.guid = "not-a-guid".into();
    assert!(e.set_user(bad).is_err());
    assert_eq!(e.revisions().unwrap(), Some(&staged));
    e.edit(|r| {
        r.headers.properties.preserve_history = Some(45);
        r.headers.headers[1].user_name = "Native header reviewer".into();
        Ok(())
    })
    .unwrap();
    let expected = e.revisions().unwrap().unwrap().clone();
    let c = e.commit().unwrap();
    let after = save(c.package());
    let output = after.to_bytes().unwrap();
    assert_eq!(model(&after), expected);
    assert_eq!(model(&after).logs, original.logs);
    values(&after);
    assert_unrelated(&bytes, &output);
    assert_same_payloads_and_graph(
        &output,
        &p.apply_revisions(c.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
    );
    assert!(matches!(
        after.apply_revisions(c.patch()),
        Err(litchi_xlsx::Error::PatchConflict { .. })
    ));
    assert_same_payloads_and_graph(
        &bytes,
        &after
            .apply_revisions(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
    );
    let mut user = user;
    user.name = "Reviewer 名".into();
    let mut e = after.edit_revisions().unwrap();
    assert!(e.set_user(user.clone()).unwrap());
    let c = e.commit().unwrap();
    let updated = save(c.package());
    assert_eq!(model(&updated).users.users, [user]);
    assert_unrelated(&output, &updated.to_bytes().unwrap());
    let mut e = updated.edit_revisions().unwrap();
    assert!(e.remove_user(42).unwrap().is_some());
    assert!(e.remove_user(42).unwrap().is_none());
    let c = e.commit().unwrap();
    let absent = save(c.package());
    assert!(model(&absent).users.users.is_empty());
    assert_eq!(model(&absent).logs, original.logs);
    assert_unrelated(&bytes, &absent.to_bytes().unwrap());
    assert_same_payloads_and_graph(
        &updated.to_bytes().unwrap(),
        &absent
            .apply_revisions(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
    );
    assert_eq!(p.to_bytes().unwrap(), bytes);
}

#[test]
fn native_revision_header_log_pairs_append_remove_and_restore_graph_ownership() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let original = model(&p);
    let mut header = original.headers.headers[2].clone();
    header.guid = "{42424242-4242-4242-4242-424242424242}".into();
    header.relationship_id = "rId4".into();
    let mut log = original.logs[0].clone();
    log.relationship_id = "rId4".into();
    log.part_name = "/xl/revisions/revisionLog4.xml".into();
    let mut e = p.edit_revisions().unwrap();
    let mut wrong = log.clone();
    wrong.relationship_id = "rIdWrong".into();
    assert!(e.set_revision(header.clone(), wrong).is_err());
    assert_eq!(e.revisions().unwrap(), Some(&original));
    assert!(e.set_revision(header.clone(), log).unwrap());
    let c = e.commit().unwrap();
    let after = save(c.package());
    let output = after.to_bytes().unwrap();
    let read = model(&after);
    assert_eq!(read.headers.headers.len(), 4);
    assert_eq!(read.logs.len(), 4);
    assert_eq!(read.headers.properties.guid, header.guid);
    assert!(read.logs[3].log.records.is_empty());
    assert!(members(&output).contains_key("xl/revisions/revisionLog4.xml"));
    assert_unrelated(&bytes, &output);
    values(&after);
    assert_same_payloads_and_graph(
        &bytes,
        &after
            .apply_revisions(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
    );
    let mut e = after.edit_revisions().unwrap();
    assert!(e.remove_revision("rId4").unwrap());
    assert!(!e.remove_revision("rId4").unwrap());
    let c = e.commit().unwrap();
    let removed = save(c.package());
    assert_eq!(model(&removed), original);
    assert!(!members(&removed.to_bytes().unwrap()).contains_key("xl/revisions/revisionLog4.xml"));
    assert_unrelated(&bytes, &removed.to_bytes().unwrap());
    assert_same_payloads_and_graph(
        &output,
        &removed
            .apply_revisions(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
    );
    assert_eq!(p.to_bytes().unwrap(), bytes);
}

#[test]
fn native_revision_owner_removal_and_recreation_keep_workbook_payloads() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let original = model(&p);
    let mut e = p.edit_revisions().unwrap();
    assert!(e.replace(None).unwrap());
    let c = e.commit().unwrap();
    let absent = save(c.package());
    let output = absent.to_bytes().unwrap();
    assert!(
        absent
            .edit_revisions()
            .unwrap()
            .revisions()
            .unwrap()
            .is_none()
    );
    assert!(
        !members(&output)
            .keys()
            .any(|n| n.starts_with("xl/revisions/"))
    );
    values(&absent);
    assert_unrelated(&bytes, &output);
    assert_same_payloads_and_graph(
        &bytes,
        &absent
            .apply_revisions(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
    );
    let mut e = absent.edit_revisions().unwrap();
    assert!(e.replace(Some(original.clone())).unwrap());
    let c = e.commit().unwrap();
    let recreated = save(c.package());
    assert_eq!(semantic(model(&recreated)), semantic(original));
    values(&recreated);
    assert_unrelated(&bytes, &recreated.to_bytes().unwrap());
    assert_same_payloads_and_graph(
        &output,
        &recreated
            .apply_revisions(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
    );
    assert_eq!(p.to_bytes().unwrap(), bytes);
}
