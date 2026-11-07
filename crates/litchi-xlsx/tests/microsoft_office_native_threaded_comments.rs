//! Native-parent threaded XML models/readers; XLSX has no public package thread editor.
use litchi_opc::{
    PackURI, TargetMode,
    constants::{content_type as ct, relationship_type as rt},
    part::BlobPart,
};
use litchi_xlsx::{
    Error, Package,
    threaded_comments::{self as tc, Comment, Comments, Mention, People, Person},
};
use std::{collections::BTreeMap, io::Cursor, path::Path};
const A: &str = "{AAAAAAAA-1111-4111-8111-111111111111}";
const B: &str = "{BBBBBBBB-2222-4222-8222-222222222222}";
const ROOT: &str = "{CCCCCCCC-3333-4333-8333-333333333333}";
const REPLY: &str = "{DDDDDDDD-4444-4444-8444-444444444444}";
const MENTION: &str = "{EEEEEEEE-5555-4555-8555-555555555555}";
const THREADS: &str = "xl/threadedComments/threadedComment1.xml";
const PERSONS: &str = "xl/persons/person1.xml";
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
fn assert_original_graph(base: &Package, added: &Package) {
    let base = base.clone().into_plain_opc();
    let added = added.clone().into_plain_opc();
    let relationships = |r: &litchi_opc::Relationships| {
        let mut values = r
            .iter()
            .filter(|r| !matches!(r.reltype(), rt::PERSONS | rt::THREADED_COMMENTS))
            .map(|r| {
                (
                    r.r_id().to_owned(),
                    r.reltype().to_owned(),
                    r.target_ref().to_owned(),
                    r.is_external(),
                )
            })
            .collect::<Vec<_>>();
        values.sort();
        values
    };
    assert_eq!(relationships(base.rels()), relationships(added.rels()));
    for part in base.iter_parts() {
        let other = added.get_part(part.partname()).unwrap();
        assert_eq!(part.content_type(), other.content_type());
        assert_eq!(relationships(part.rels()), relationships(other.rels()));
    }
    let mut expected = base
        .iter_parts()
        .map(|p| p.partname().to_string())
        .collect::<Vec<_>>();
    expected.extend([format!("/{THREADS}"), format!("/{PERSONS}")]);
    expected.sort();
    let mut actual = added
        .iter_parts()
        .map(|p| p.partname().to_string())
        .collect::<Vec<_>>();
    actual.sort();
    assert_eq!(actual, expected);
}
fn people() -> People {
    People {
        persons: vec![
            Person {
                display_name: "Ada & Co".into(),
                id: A.into(),
                user_id: Some("ada@example.invalid".into()),
                provider_id: Some("local".into()),
            },
            Person {
                display_name: "Bob".into(),
                id: B.into(),
                ..Default::default()
            },
        ],
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
                mentions: vec![Mention {
                    mention_person_id: B.into(),
                    mention_id: MENTION.into(),
                    start_index: 3,
                    length: 3,
                }],
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
    }
}
fn graph(comments: Comments) -> tc::Graph {
    tc::Graph {
        persons: Some(tc::PeoplePart {
            relationship_id: "rIdNativePeople065".into(),
            part_name: format!("/{PERSONS}"),
            persons: people(),
        }),
        worksheets: vec![tc::CommentsPart {
            worksheet_part_name: "/xl/worksheets/sheet1.xml".into(),
            relationship_id: "rIdNativeThreads065".into(),
            part_name: format!("/{THREADS}"),
            comments,
        }],
    }
}
fn with_xml(bytes: &[u8], xml: String) -> Package {
    let mut p = Package::from_bytes(bytes).unwrap().into_plain_opc();
    for (name, content, xml) in [
        (
            PERSONS,
            ct::SML_PERSONS,
            tc::write_persons(&people()).unwrap(),
        ),
        (THREADS, ct::SML_THREADED_COMMENTS, xml),
    ] {
        p.try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{name}")).unwrap(),
            content.into(),
            xml.into_bytes(),
        )))
        .unwrap();
    }
    p.get_part_mut(&PackURI::new("/xl/workbook.xml").unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::PERSONS.into(),
            "persons/person1.xml".into(),
            "rIdNativePeople065".into(),
            TargetMode::Internal,
        )
        .unwrap();
    p.get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::THREADED_COMMENTS.into(),
            "../threadedComments/threadedComment1.xml".into(),
            "rIdNativeThreads065".into(),
            TargetMode::Internal,
        )
        .unwrap();
    Package::from_opc(p).unwrap()
}
fn assert_models(comments: &Comments, people: &People) {
    let expected = self::comments();
    assert_eq!(comments.comments.len(), expected.comments.len());
    for (actual, wanted) in comments.comments.iter().zip(&expected.comments) {
        assert_eq!(
            (
                &actual.cell_ref,
                &actual.id,
                &actual.parent_id,
                &actual.person_id,
                &actual.text,
                &actual.date_time,
                actual.done
            ),
            (
                &wanted.cell_ref,
                &wanted.id,
                &wanted.parent_id,
                &wanted.person_id,
                &wanted.text,
                &wanted.date_time,
                wanted.done
            )
        );
        assert_eq!(actual.mentions.len(), wanted.mentions.len());
        for (a, b) in actual.mentions.iter().zip(&wanted.mentions) {
            assert_eq!(
                (&a.mention_person_id, &a.mention_id, a.start_index, a.length),
                (&b.mention_person_id, &b.mention_id, b.start_index, b.length)
            );
        }
    }
    let expected = self::people();
    assert_eq!(people.persons.len(), expected.persons.len());
    for (a, b) in people.persons.iter().zip(&expected.persons) {
        assert_eq!(
            (&a.id, &a.display_name, &a.user_id, &a.provider_id),
            (&b.id, &b.display_name, &b.user_id, &b.provider_id)
        );
    }
    let units = comments.comments[0]
        .text
        .as_ref()
        .unwrap()
        .encode_utf16()
        .collect::<Vec<_>>();
    assert_eq!(String::from_utf16(&units[3..6]).unwrap(), "Bob");

    assert_eq!(
        tc::write_comments(comments).unwrap(),
        tc::write_comments(&self::comments()).unwrap()
    );
    assert_eq!(
        tc::write_persons(people).unwrap(),
        tc::write_persons(&self::people()).unwrap()
    );
    assert_eq!(comments.comments[0].text.as_deref(), Some("🙂 Bob & co."));
    assert_eq!(
        (
            comments.comments[0].mentions[0].start_index,
            comments.comments[0].mentions[0].length
        ),
        (3, 3)
    );
    assert_eq!(comments.comments[1].parent_id.as_deref(), Some(ROOT));
    assert_eq!(comments.comments[0].done, Some(false));
    assert_eq!(comments.comments[1].done, None);
}
#[test]
fn native_xlsx_thread_models_read_on_both_owners_and_survive_ordinary_cell_edits() {
    for file in [
        "excel-metadata.xlsx",
        "excel-recorded-macro.xlsm",
        "excel-interaction-print.xlsx",
    ] {
        let bytes = native(file);
        let base = Package::from_bytes(&bytes).unwrap();
        assert!(base.workbook().people().unwrap().is_none());
        assert!(
            base.workbook()
                .sheet("Sheet1")
                .unwrap()
                .unwrap()
                .threaded_comments()
                .unwrap()
                .is_none()
        );
        tc::validate_graph(&graph(comments())).unwrap();
        let p = with_xml(&bytes, tc::write_comments(&comments()).unwrap());
        assert_original_graph(&base, &p);
        let before = members(&p);
        for (name, data) in members(&base) {
            if !matches!(
                name.as_str(),
                "[Content_Types].xml"
                    | "xl/_rels/workbook.xml.rels"
                    | "xl/worksheets/_rels/sheet1.xml.rels"
            ) {
                assert!(before.get(&name) == Some(&data), "changed {name}");
            }
        }
        let workbook = p.workbook();
        assert_models(
            &workbook
                .sheet("Sheet1")
                .unwrap()
                .unwrap()
                .threaded_comments()
                .unwrap()
                .unwrap(),
            &workbook.people().unwrap().unwrap(),
        );
        let input = p.to_bytes().unwrap();
        let source = litchi_xlsx::SourceBackedWorkbook::from_reader(Cursor::new(&input)).unwrap();
        assert_models(
            &source
                .sheet("Sheet1")
                .unwrap()
                .unwrap()
                .threaded_comments()
                .unwrap()
                .unwrap(),
            &source.people().unwrap().unwrap(),
        );
        let f = tempfile::NamedTempFile::with_suffix(".xlsx").unwrap();
        p.save(f.path()).unwrap();
        assert!(members(&Package::open(f.path()).unwrap()) == before);
        assert!(
            p.workbook()
                .sheet("Sheet1")
                .unwrap()
                .unwrap()
                .cell("Z100")
                .unwrap()
                .is_missing()
        );
        let mut e = p.edit().unwrap();
        e.sheet("Sheet1").unwrap().unwrap().set("Z100", 65).unwrap();
        if file == "excel-interaction-print.xlsx" {
            assert!(matches!(
                e.commit(),
                Err(Error::EditBlocked {
                    reason: litchi_xlsx::EditBlock::ProtectedSheet,
                    ..
                })
            ));
            assert!(members(&p) == before);
        } else {
            let c = e.commit().unwrap();
            let after = members(c.package());
            assert_eq!(after[THREADS], before[THREADS]);
            assert_eq!(after[PERSONS], before[PERSONS]);
            for (name, data) in &before {
                if name.ends_with("vbaProject.bin")
                    || name.contains("comments")
                    || name.contains("vmlDrawing")
                {
                    assert!(after.get(name) == Some(data));
                }
            }
            let q = Package::from_bytes(&c.package().to_bytes().unwrap()).unwrap();
            assert_eq!(
                q.workbook()
                    .sheet("Sheet1")
                    .unwrap()
                    .unwrap()
                    .cell("Z100")
                    .unwrap()
                    .stored(),
                Some(&litchi_xlsx::Cell::Value(litchi_xlsx::Value::Number(
                    litchi_xlsx::cell::Number::new("65").unwrap()
                )))
            );

            assert_models(
                &q.workbook()
                    .sheet("Sheet1")
                    .unwrap()
                    .unwrap()
                    .threaded_comments()
                    .unwrap()
                    .unwrap(),
                &q.workbook().people().unwrap().unwrap(),
            );
            assert!(members(c.package().apply(&c.patch().inverse()).unwrap().package()) == before);
        }
        assert!(p.to_bytes().unwrap() == input);
        assert!(base.to_bytes().unwrap() == bytes);
        assert!(native(file) == bytes);
    }
}
#[test]
fn native_xlsx_mentions_without_text_are_refused_by_model_and_package_readers() {
    let bytes = native("excel-metadata.xlsx");
    let mut invalid = comments();
    invalid.comments[0].text = None;
    let source = tc::write_comments(&comments()).unwrap();
    let start = source.find("<text>").unwrap();
    let end = start + source[start..].find("</text>").unwrap() + 7;
    let xml = format!("{}{}", &source[..start], &source[end..]);
    let p = with_xml(&bytes, xml);
    let input = p.to_bytes().unwrap();
    let error = p
        .workbook()
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .threaded_comments()
        .unwrap_err();
    assert!(matches!(error,Error::Invalid(message) if message.contains("mentions but no text")));
    let source = litchi_xlsx::SourceBackedWorkbook::from_reader(Cursor::new(&input)).unwrap();
    let error = source
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .threaded_comments()
        .unwrap_err();
    assert!(matches!(error,Error::Invalid(message) if message.contains("mentions but no text")));
    assert!(p.to_bytes().unwrap() == input);
    assert!(native("excel-metadata.xlsx") == bytes);
    let mut empty = invalid.clone();
    empty.comments[0].mentions.clear();
    let roundtrip = tc::parse_comments(&tc::write_comments(&empty).unwrap()).unwrap();
    assert_eq!(roundtrip.comments[0].text, None);
    assert!(roundtrip.comments[0].mentions.is_empty());
    let error = tc::validate_comments(&invalid).unwrap_err();
    assert!(
        matches!(error,Error::Invalid(message) if message=="threaded-comment has mentions but no text")
    );
    assert!(
        matches!(tc::write_comments(&invalid),Err(Error::Invalid(message)) if message=="threaded-comment has mentions but no text")
    );
}
#[test]
fn native_xlsx_external_thread_relationships_refuse_without_fetching_targets() {
    let bytes = native("excel-metadata.xlsx");
    let base = with_xml(&bytes, tc::write_comments(&comments()).unwrap());
    for (owner, rid, kind, message) in [
        (
            "/xl/workbook.xml",
            "rIdNativePeople065",
            rt::PERSONS,
            "a persons relationship cannot be external",
        ),
        (
            "/xl/worksheets/sheet1.xml",
            "rIdNativeThreads065",
            rt::THREADED_COMMENTS,
            "a threaded-comments relationship cannot be external",
        ),
    ] {
        let mut opc = base.clone().into_plain_opc();
        let part = opc.get_part_mut(&PackURI::new(owner).unwrap()).unwrap();
        part.rels_mut().remove(rid);
        part.rels_mut()
            .try_add_relationship(
                kind.into(),
                "https://example.invalid/never-fetched.xml".into(),
                rid.into(),
                TargetMode::External,
            )
            .unwrap();
        let p = Package::from_opc(opc).unwrap();
        let input = p.to_bytes().unwrap();
        let error = if kind == rt::PERSONS {
            p.workbook().people().unwrap_err()
        } else {
            p.workbook()
                .sheet("Sheet1")
                .unwrap()
                .unwrap()
                .threaded_comments()
                .unwrap_err()
        };
        assert!(matches!(error,Error::Invalid(m) if m==message));
        let source = litchi_xlsx::SourceBackedWorkbook::from_reader(Cursor::new(&input)).unwrap();
        let error = if kind == rt::PERSONS {
            source.people().unwrap_err()
        } else {
            source
                .sheet("Sheet1")
                .unwrap()
                .unwrap()
                .threaded_comments()
                .unwrap_err()
        };
        assert!(matches!(error,Error::Invalid(m) if m==message));
        assert!(p.to_bytes().unwrap() == input);
    }
    assert!(native("excel-metadata.xlsx") == bytes);
}
