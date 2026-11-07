//! Explicit threaded-comment derivatives of native Excel parents; no collaboration services.
use litchi_core::edit::RefusalKind;
use litchi_xlsb::{
    Package,
    comments::threaded::{self as tc, Comment, Comments, Mention, People, Person},
};
use std::{collections::BTreeMap, path::Path};
const A: &str = "{AAAAAAAA-1111-4111-8111-111111111111}";
const B: &str = "{BBBBBBBB-2222-4222-8222-222222222222}";
const ROOT: &str = "{CCCCCCCC-3333-4333-8333-333333333333}";
const REPLY: &str = "{DDDDDDDD-4444-4444-8444-444444444444}";
const MENTION: &str = "{EEEEEEEE-5555-4555-8555-555555555555}";
fn native(file: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(file),
    )
    .unwrap()
}
fn members(p: &Package) -> BTreeMap<String, Vec<u8>> {
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        p.to_bytes().unwrap(),
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
fn people() -> People {
    let mut a = Person::new(A, "Ada & Co");
    a.user_id = Some("ada@example.invalid".into());
    a.provider_id = Some("local".into());
    People {
        persons: vec![a, Person::new(B, "Bob")],
        ..Default::default()
    }
}
fn comments() -> Comments {
    Comments {
        comments: vec![
            Comment {
                cell_ref: Some("B2".into()),
                id: ROOT.into(),
                person_id: A.into(),
                text: Some("🙂 Bob & co.".into()),
                date_time: Some("2026-11-07T10:20:30Z".into()),
                done: Some(false),
                mentions: vec![Mention::new(B, MENTION, 3, 3)],
                extensions: vec![tc::RawXml::new(
                    br#"<x:retained xmlns:x="urn:litchi:thread-extension" value="A &amp; B"/>"#
                        .to_vec(),
                )],
                ..Default::default()
            },
            Comment {
                id: REPLY.into(),
                parent_id: Some(ROOT.into()),
                person_id: B.into(),
                text: Some("Réponse ✓".into()),
                date_time: Some("2026-11-07T10:21:00Z".into()),
                ..Default::default()
            },
        ],
        ..Default::default()
    }
}
fn preserved(native: &Package, changed: &Package, owner_rels: &str) {
    let rels = |r: &litchi_opc::Relationships| {
        let mut v = r
            .iter()
            .filter(|r| {
                !matches!(
                    r.reltype(),
                    tc::package::COMMENTS_RELATIONSHIP_TYPE
                        | tc::package::PERSONS_RELATIONSHIP_TYPE
                )
            })
            .map(|r| {
                (
                    r.r_id().to_owned(),
                    r.reltype().to_owned(),
                    r.target_ref().to_owned(),
                    r.is_external(),
                )
            })
            .collect::<Vec<_>>();
        v.sort();
        v
    };
    assert_eq!(
        rels(native.opc_package().rels()),
        rels(changed.opc_package().rels())
    );
    for part in native.opc_package().iter_parts() {
        let other = changed.opc_package().get_part(part.partname()).unwrap();
        assert_eq!(part.content_type(), other.content_type());
        assert_eq!(rels(part.rels()), rels(other.rels()));
    }
    let a = members(native);
    let b = members(changed);
    for (name, data) in a {
        if !matches!(
            name.as_str(),
            "[Content_Types].xml" | "xl/_rels/workbook.bin.rels"
        ) && name != owner_rels
        {
            assert!(b.get(&name) == Some(&data), "changed native member {name}");
        }
    }
}
fn publish(p: &Package, c: &tc::edit::Commit) -> Package {
    let f = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
    c.package().save(f.path()).unwrap();
    let r = Package::open(f.path()).unwrap();
    assert_eq!(
        r.threaded_comments().unwrap(),
        c.package().threaded_comments().unwrap()
    );
    assert!(members(&r) == members(c.package()));
    assert!(members(p.apply_threaded_comments(c.patch()).unwrap().package()) == members(&r));
    assert!(
        members(
            r.apply_threaded_comments(&c.patch().inverse())
                .unwrap()
                .package()
        ) == members(p)
    );
    r
}
#[test]
fn native_xlsb_thread_lifecycle_preserves_cells_legacy_notes_and_chart_tab_order() {
    for (file, sheet, owner_rels) in [
        (
            "excel-interaction-print.xlsb",
            "/xl/worksheets/sheet1.bin",
            "xl/worksheets/_rels/sheet1.bin.rels",
        ),
        (
            "excel-values-formulas-chart.xlsb",
            "/xl/worksheets/sheet2.bin",
            "xl/worksheets/_rels/sheet2.bin.rels",
        ),
    ] {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        assert!(p.threaded_comments().unwrap().is_empty());
        let before = members(&p);
        let mut e = p.edit_threaded_comments().unwrap();
        e.set_people(people()).unwrap();
        e.set_worksheet(sheet, comments()).unwrap();
        let c = e.commit().unwrap();
        let added = publish(&p, &c);
        preserved(&p, &added, owner_rels);
        let snapshot = added.threaded_comments().unwrap();
        assert_eq!(snapshot.people().unwrap(), &people());
        assert_eq!(snapshot.worksheet(sheet).unwrap(), &comments());
        let threads = tc::group_threads(snapshot.worksheet(sheet).unwrap()).unwrap();
        assert_eq!(threads.len(), 1);
        assert_eq!(threads[0].root.id, ROOT);
        assert_eq!(threads[0].replies[0].id, REPLY);
        assert_eq!(threads[0].root.mentions[0].start_index, 3);
        assert_eq!(threads[0].root.mentions[0].length, 3);
        let identity = snapshot.graph().worksheets[0].clone();
        let persons = snapshot.graph().persons.as_ref().unwrap().clone();
        let mut e = added.edit_threaded_comments().unwrap();
        let mut changed = comments();
        changed.comments[0].done = Some(true);
        changed.comments[1].text = Some("Reviewed reply 日本語".into());
        e.set_worksheet(sheet, changed.clone()).unwrap();
        let mut renamed = people().persons.remove(1);
        renamed.display_name = "Robert".into();
        e.upsert_person(renamed).unwrap();
        let c = e.commit().unwrap();
        let updated = publish(&added, &c);
        preserved(&p, &updated, owner_rels);
        let now = updated.threaded_comments().unwrap();
        assert_eq!(now.worksheet(sheet).unwrap(), &changed);
        assert_eq!(now.graph().worksheets[0].part_name, identity.part_name);
        assert_eq!(
            now.graph().worksheets[0].relationship_id,
            identity.relationship_id
        );
        assert_eq!(
            now.graph().persons.as_ref().unwrap().part_name,
            persons.part_name
        );
        assert_eq!(
            now.graph().persons.as_ref().unwrap().relationship_id,
            persons.relationship_id
        );
        assert_eq!(now.people().unwrap().persons[1].display_name, "Robert");
        let old = members(&added);
        let new = members(&updated);
        let mut expected = vec![
            identity.part_name.trim_start_matches('/').to_owned(),
            persons.part_name.trim_start_matches('/').to_owned(),
        ];
        expected.sort();
        assert_eq!(
            old.keys()
                .filter(|n| old.get(*n) != new.get(*n))
                .cloned()
                .collect::<Vec<_>>(),
            expected
        );
        assert!(matches!(
            updated.apply_threaded_comments(c.patch()),
            Err(litchi_xlsb::package::PackageError::PatchConflict { .. })
        ));
        let mut e = updated.edit_threaded_comments().unwrap();
        assert!(
            matches!(e.remove_person(B),Err(litchi_xlsb::package::PackageError::Refused(r)) if r.kind()==RefusalKind::InvalidValue)
        );
        let removed = e.remove_thread(sheet, ROOT).unwrap().unwrap();
        assert_eq!(removed.replies.len(), 1);
        assert!(e.remove_person(B).unwrap().is_some());
        assert!(e.remove_worksheet(sheet).unwrap());
        assert!(e.remove_people().unwrap());
        let c = e.commit().unwrap();
        let cleared = publish(&updated, &c);
        assert!(cleared.threaded_comments().unwrap().is_empty());
        preserved(&p, &cleared, owner_rels);
        assert_eq!(
            members(&cleared).keys().collect::<Vec<_>>(),
            before.keys().collect::<Vec<_>>()
        );
        assert!(p.to_bytes().unwrap() == bytes);
        assert!(native(file) == bytes);
    }
}
#[test]
fn native_xlsb_invalid_thread_references_preserve_valid_staging() {
    let bytes = native("excel-metadata.xlsb");
    let p = Package::from_bytes(&bytes).unwrap();
    let sheet = "/xl/worksheets/sheet1.bin";
    let mut e = p.edit_threaded_comments().unwrap();
    e.set_people(people()).unwrap();
    e.set_worksheet(sheet, comments()).unwrap();
    // The final three UTF-16 units are valid even though the emoji is one scalar.
    let mut boundary = comments();
    boundary.comments[0].mentions[0].start_index = 9;
    e.set_worksheet(sheet, boundary.clone()).unwrap();
    boundary.comments[0].mentions[0].length = 4;
    assert!(
        matches!(e.set_worksheet(sheet,boundary),Err(litchi_xlsb::package::PackageError::Refused(r)) if r.kind()==RefusalKind::InvalidValue)
    );
    assert_eq!(
        e.worksheet(sheet).unwrap().unwrap().comments[0].mentions[0].length,
        3
    );
    e.set_worksheet(sheet, comments()).unwrap();
    let staged = e.graph().unwrap().clone();
    for case in 0..5 {
        let mut invalid = comments();
        match case {
            0 => invalid.comments[0].mentions[0].length = 100,
            1 => invalid.comments[0].text = None,
            2 => invalid.comments[1].parent_id = Some(MENTION.into()),
            3 => invalid.comments[1].cell_ref = Some("B2".into()),
            _ => invalid.comments[0].person_id = MENTION.into(),
        };
        assert!(
            matches!(e.set_worksheet(sheet,invalid),Err(litchi_xlsb::package::PackageError::Refused(r)) if r.kind()==RefusalKind::InvalidValue&&r.operation()=="edit_threaded_comments"&&r.reason()=="the threaded-comments graph is invalid")
        );
        assert_eq!(e.graph().unwrap(), &staged);
    }
    assert!(
        matches!(e.remove_people(),Err(litchi_xlsb::package::PackageError::Refused(r)) if r.kind()==RefusalKind::InvalidValue)
    );
    assert_eq!(e.graph().unwrap(), &staged);
    let c = e.commit().unwrap();
    let reopened = publish(&p, &c);
    assert_eq!(
        reopened
            .threaded_comments()
            .unwrap()
            .worksheet(sheet)
            .unwrap(),
        &comments()
    );
    assert!(p.to_bytes().unwrap() == bytes);
}

#[test]
fn native_xlsb_chart_tab_cannot_own_a_threaded_comment_part() {
    let bytes = native("excel-values-formulas-chart.xlsb");
    let p = Package::from_bytes(&bytes).unwrap();
    let mut e = p.edit_threaded_comments().unwrap();
    e.set_people(people()).unwrap();
    e.set_worksheet("/xl/chartsheets/sheet1.bin", comments())
        .unwrap();
    assert!(matches!(
        e.clone().commit(),
        Err(litchi_xlsb::package::PackageError::Unrecognized {
            typ: "XLSB threaded-comments package graph",
            val: "threaded-comment operations require an XLSB worksheet"
        })
    ));
    e.reset().unwrap();
    assert!(e.commit().unwrap().patch().is_noop());
    assert!(p.to_bytes().unwrap() == bytes);
}
