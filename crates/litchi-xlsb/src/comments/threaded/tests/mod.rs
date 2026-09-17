use super::package::{
    COMMENTS_CONTENT_TYPE, COMMENTS_RELATIONSHIP_TYPE, PERSONS_CONTENT_TYPE,
    PERSONS_RELATIONSHIP_TYPE, WORKSHEET_CONTENT_TYPE,
};
use super::{
    Comment, Comments, CommentsPart, Graph, Mention, People, PeoplePart, Person, RawXml, Thread,
    apply, load_graph, parse_comments, parse_persons, read, remove_graph, store_graph,
    validate_comments, validate_graph as validate_model_graph, write_comments, write_persons,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, TargetMode};

const NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
const ALICE: &str = "{11111111-1111-1111-1111-111111111111}";
const ROOT: &str = "{22222222-2222-2222-2222-222222222222}";
const REPLY: &str = "{33333333-3333-3333-3333-333333333333}";
const MENTION: &str = "{44444444-4444-4444-4444-444444444444}";

#[test]
fn parses_and_preserves_unknown_threaded_payloads() {
    let people = parse_persons(format!(
        r#"<tc:personList xmlns:tc="{NS}" xmlns:f="urn:future"><tc:person displayName="Alice" id="{ALICE}"><f:personExt><f:value/></f:personExt></tc:person><f:peopleExt/></tc:personList>"#
    ))
    .unwrap();
    assert_eq!(people.persons[0].display_name, "Alice");
    assert_eq!(people.persons[0].extensions.len(), 1);
    assert_eq!(people.extensions.len(), 1);
    let people_xml = write_persons(&people).unwrap();
    assert!(String::from_utf8_lossy(&people_xml).contains("personExt"));

    let comments = parse_comments(format!(
        r#"<tc:ThreadedComments xmlns:tc="{NS}" xmlns:f="urn:future"><tc:threadedComment ref="B2" id="{ROOT}" personId="{ALICE}"><tc:text>Hello &amp; @Bob</tc:text><tc:mentions><tc:mention mentionpersonId="{ALICE}" mentionId="{MENTION}" startIndex="8" length="4"/></tc:mentions><f:commentExt/></tc:threadedComment><f:rootExt/></tc:ThreadedComments>"#
    ))
    .unwrap();
    assert_eq!(comments.comments[0].text.as_deref(), Some("Hello & @Bob"));
    assert_eq!(comments.comments[0].mentions.len(), 1);
    assert_eq!(comments.comments[0].extensions.len(), 1);
    assert_eq!(comments.extensions.len(), 1);
    let xml = write_comments(&comments).unwrap();
    assert!(String::from_utf8_lossy(&xml).contains("commentExt"));
}

#[test]
fn validates_root_reply_and_mention_graphs() {
    let people = People {
        persons: vec![Person::new(ALICE, "Alice")],
        ..People::default()
    };
    let comments = Comments {
        comments: vec![
            Comment {
                cell_ref: Some("B2".into()),
                id: ROOT.into(),
                person_id: ALICE.into(),
                text: Some("Hello @Bob".into()),
                mentions: vec![Mention::new(ALICE, MENTION, 6, 4)],
                ..Comment::default()
            },
            Comment {
                id: REPLY.into(),
                person_id: ALICE.into(),
                parent_id: Some(ROOT.into()),
                text: Some("Reply".into()),
                ..Comment::default()
            },
        ],
        ..Comments::default()
    };
    let graph = Graph {
        persons: Some(PeoplePart {
            relationship_id: String::new(),
            part_name: String::new(),
            persons: people,
        }),
        worksheets: vec![CommentsPart {
            worksheet_part_name: "/xl/worksheets/sheet1.bin".into(),
            relationship_id: String::new(),
            part_name: String::new(),
            comments: comments.clone(),
        }],
    };
    validate_model_graph(&graph).unwrap();
    validate_comments(&comments).unwrap();
    let threads: Vec<Thread> = super::group_threads(&comments).unwrap();
    assert_eq!(threads.len(), 1);
    assert_eq!(threads[0].replies.len(), 1);
}

#[test]
fn rejects_unsafe_semantic_references() {
    let comments = Comments {
        comments: vec![Comment {
            cell_ref: Some("A1".into()),
            id: ROOT.into(),
            person_id: ALICE.into(),
            text: Some("x".into()),
            mentions: vec![Mention::new(ALICE, MENTION, 1, 1)],
            ..Comment::default()
        }],
        ..Comments::default()
    };
    assert!(
        validate_model_graph(&Graph {
            persons: None,
            worksheets: vec![CommentsPart {
                worksheet_part_name: "/xl/worksheets/sheet1.bin".into(),
                relationship_id: String::new(),
                part_name: String::new(),
                comments,
            }],
        })
        .is_err()
    );
}

#[test]
fn package_crud_is_bounded_and_keeps_legacy_and_unknown_parts() {
    let (mut package, _workbook, worksheet) = fixture();
    let legacy = PackURI::new("/xl/comments1.bin").unwrap();
    package.add_part(Box::new(BlobPart::new(
        legacy.clone(),
        "application/vnd.ms-excel.comments".into(),
        b"legacy".to_vec(),
    )));
    let unknown = PackURI::new("/xl/custom.bin").unwrap();
    package.add_part(Box::new(BlobPart::new(
        unknown.clone(),
        "application/octet-stream".into(),
        b"unknown".to_vec(),
    )));
    package
        .get_part_mut(&worksheet)
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            "urn:future:binding".into(),
            "../custom.bin".into(),
            "rIdFuture".into(),
            TargetMode::Internal,
        )
        .unwrap();

    let graph = Graph {
        persons: Some(PeoplePart {
            relationship_id: String::new(),
            part_name: String::new(),
            persons: People {
                persons: vec![Person::new(ALICE, "Alice")],
                ..People::default()
            },
        }),
        worksheets: vec![CommentsPart {
            worksheet_part_name: worksheet.to_string(),
            relationship_id: String::new(),
            part_name: String::new(),
            comments: Comments {
                comments: vec![Comment {
                    cell_ref: Some("A1".into()),
                    id: ROOT.into(),
                    person_id: ALICE.into(),
                    ..Comment::default()
                }],
                ..Comments::default()
            },
        }],
    };
    store_graph(&mut package, &graph).unwrap();
    super::package::validate_graph(&package).unwrap();
    let loaded = load_graph(&package).unwrap();
    assert_eq!(loaded.worksheets.len(), 1);
    assert_eq!(loaded.persons.unwrap().persons.persons.len(), 1);
    assert_eq!(package.get_part(&legacy).unwrap().blob(), b"legacy");
    assert_eq!(package.get_part(&unknown).unwrap().blob(), b"unknown");
    assert_eq!(
        package
            .get_part(&worksheet)
            .unwrap()
            .rels()
            .get("rIdFuture")
            .unwrap()
            .reltype(),
        "urn:future:binding"
    );

    assert!(remove_graph(&mut package).unwrap());
    assert!(package.get_part(&legacy).is_ok());
    assert!(package.get_part(&unknown).is_ok());
    assert!(
        package
            .iter_parts()
            .all(|part| part.content_type() != COMMENTS_CONTENT_TYPE
                && part.content_type() != PERSONS_CONTENT_TYPE)
    );
}

#[test]
fn rejects_external_or_misowned_relationships() {
    let (mut package, workbook, worksheet) = fixture();
    package
        .get_part_mut(&workbook)
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            PERSONS_RELATIONSHIP_TYPE.into(),
            "https://example.invalid/person.xml".into(),
            "rIdPersons1".into(),
            TargetMode::External,
        )
        .unwrap();
    assert!(super::package::validate_graph(&package).is_err());

    let mut package = fixture().0;
    let target = PackURI::new("/xl/threadedComments/threadedComment1.xml").unwrap();
    package.add_part(Box::new(BlobPart::new(
        target.clone(),
        COMMENTS_CONTENT_TYPE.into(),
        format!(r#"<ThreadedComments xmlns="{NS}"/>"#).into_bytes(),
    )));
    package
        .get_part_mut(&workbook)
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            COMMENTS_RELATIONSHIP_TYPE.into(),
            "../threadedComments/threadedComment1.xml".into(),
            "rIdBad".into(),
            TargetMode::Internal,
        )
        .unwrap();
    assert!(super::package::validate_graph(&package).is_err());
    let _ = worksheet;
}

#[test]
fn snapshot_noop_preserves_exact_owner_bytes() {
    let (mut package, _workbook, worksheet) = fixture();
    store_graph(&mut package, &graph_fixture(&worksheet)).unwrap();
    let before = owner_image(&package);
    let snapshot = read(&package).unwrap();
    assert!(snapshot.source_parts().iter().any(|part| part.has_bytes()));

    let commit = snapshot.edit().commit().unwrap();
    assert!(commit.patch().is_empty());
    let after = apply(&mut package, commit.patch()).unwrap();

    assert_eq!(after, snapshot);
    assert_eq!(owner_image(&package), before);
}

#[test]
fn changed_edit_validates_and_preserves_unknown_xml() {
    let (mut package, _workbook, worksheet) = fixture();
    store_graph(&mut package, &graph_fixture(&worksheet)).unwrap();
    let snapshot = read(&package).unwrap();
    let mut edit = snapshot.edit();
    let mut alice = snapshot.people().unwrap().persons[0].clone();
    alice.display_name = "Alicia".into();
    edit.upsert_person(alice).unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.patch().is_empty());
    assert!(!commit.snapshot().is_source_bound());
    assert!(commit.snapshot().source_parts().is_empty());

    let committed = apply(&mut package, commit.patch()).unwrap();
    assert_eq!(
        committed.people().unwrap().persons[0].display_name,
        "Alicia"
    );
    let person_part = package
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
        .find(|part| part.content_type() == PERSONS_CONTENT_TYPE)
        .unwrap();
    let comments_part = package
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
        .find(|part| part.content_type() == COMMENTS_CONTENT_TYPE)
        .unwrap();
    assert!(String::from_utf8_lossy(person_part.blob()).contains("personExt"));
    assert!(String::from_utf8_lossy(comments_part.blob()).contains("commentExt"));
}

#[test]
fn invalid_edits_and_source_conflicts_are_atomic() {
    let (mut package, _workbook, worksheet) = fixture();
    store_graph(&mut package, &graph_fixture(&worksheet)).unwrap();
    let snapshot = read(&package).unwrap();

    let mut invalid = snapshot.edit();
    let before_graph = invalid.graph().clone();
    assert!(invalid.set_people(People::default()).is_err());
    assert_eq!(invalid.graph(), &before_graph);

    let mut changed = snapshot.edit();
    let mut alice = snapshot.people().unwrap().persons[0].clone();
    alice.display_name = "Other writer".into();
    changed.upsert_person(alice).unwrap();
    let commit = changed.commit().unwrap();

    let mut conflict = package.clone();
    let mut other = snapshot.graph().clone();
    other.persons.as_mut().unwrap().persons.persons[0].display_name = "Conflict".into();
    store_graph(&mut conflict, &other).unwrap();
    let before = owner_image(&conflict);
    assert!(apply(&mut conflict, commit.patch()).is_err());
    assert_eq!(owner_image(&conflict), before);
}

#[test]
fn removing_threads_and_people_cleans_owner_parts_and_relationships() {
    let (mut package, workbook, worksheet) = fixture();
    store_graph(&mut package, &graph_fixture(&worksheet)).unwrap();
    let snapshot = read(&package).unwrap();
    let mut edit = snapshot.edit();
    let removed = edit.remove_thread(worksheet.as_str(), ROOT).unwrap();
    assert!(removed.is_some());
    assert!(edit.remove_worksheet(worksheet.as_str()).unwrap());
    assert!(edit.remove_person(ALICE).unwrap().is_some());
    let commit = edit.commit().unwrap();
    apply(&mut package, commit.patch()).unwrap();

    assert!(
        package
            .iter_parts()
            .all(|part| part.content_type() != COMMENTS_CONTENT_TYPE
                && part.content_type() != PERSONS_CONTENT_TYPE)
    );
    assert!(
        !package
            .get_part(&workbook)
            .unwrap()
            .rels()
            .iter()
            .any(|rel| rel.reltype() == PERSONS_RELATIONSHIP_TYPE)
    );
    assert!(
        !package
            .get_part(&worksheet)
            .unwrap()
            .rels()
            .iter()
            .any(|rel| rel.reltype() == COMMENTS_RELATIONSHIP_TYPE)
    );
}

fn graph_fixture(worksheet: &PackURI) -> Graph {
    let mut alice = Person::new(ALICE, "Alice");
    alice.extensions.push(RawXml::new(
        br#"<f:personExt xmlns:f="urn:future"/>"#.to_vec(),
    ));
    Graph {
        persons: Some(PeoplePart {
            relationship_id: String::new(),
            part_name: String::new(),
            persons: People {
                persons: vec![alice],
                ..People::default()
            },
        }),
        worksheets: vec![CommentsPart {
            worksheet_part_name: worksheet.to_string(),
            relationship_id: String::new(),
            part_name: String::new(),
            comments: Comments {
                comments: vec![Comment {
                    cell_ref: Some("A1".into()),
                    id: ROOT.into(),
                    person_id: ALICE.into(),
                    text: Some("hello".into()),
                    extensions: vec![RawXml::new(
                        br#"<f:commentExt xmlns:f="urn:future"/>"#.to_vec(),
                    )],
                    ..Comment::default()
                }],
                ..Comments::default()
            },
        }],
    }
}

fn owner_image(package: &OpcPackage) -> Vec<(String, Vec<u8>)> {
    let mut image: Vec<_> = package
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
        .filter(|part| {
            part.content_type() == PERSONS_CONTENT_TYPE
                || part.content_type() == COMMENTS_CONTENT_TYPE
        })
        .map(|part| (part.partname().to_string(), part.blob().to_vec()))
        .collect();
    image.sort_by(|left, right| left.0.cmp(&right.0));
    image
}

fn fixture() -> (OpcPackage, PackURI, PackURI) {
    let mut package = OpcPackage::new();
    let workbook = PackURI::new("/xl/workbook.bin").unwrap();
    let worksheet = PackURI::new("/xl/worksheets/sheet1.bin").unwrap();
    package.add_part(Box::new(BlobPart::new(
        workbook.clone(),
        ct::XLSB_BIN.into(),
        Vec::new(),
    )));
    package.add_part(Box::new(BlobPart::new(
        worksheet.clone(),
        WORKSHEET_CONTENT_TYPE.into(),
        Vec::new(),
    )));
    package
        .rels_mut()
        .get_or_add(rt::OFFICE_DOCUMENT, "xl/workbook.bin");
    (package, workbook, worksheet)
}

fn corrupt_zip_member(mut bytes: Vec<u8>, member: &str) -> Vec<u8> {
    let member = member.as_bytes();
    let mut local_found = false;
    let mut cursor = 0_usize;
    while let Some(relative) = bytes[cursor..]
        .windows(4)
        .position(|window| window == b"PK\x03\x04")
    {
        let header = cursor + relative;
        let name_length = u16::from_le_bytes([bytes[header + 26], bytes[header + 27]]) as usize;
        let extra_length = u16::from_le_bytes([bytes[header + 28], bytes[header + 29]]) as usize;
        let name_start = header + 30;
        let data_start = name_start + name_length + extra_length;
        if &bytes[name_start..name_start + name_length] == member {
            assert!(data_start < bytes.len(), "ZIP member has no payload");
            local_found = true;
            break;
        }
        cursor = header + 4;
    }
    assert!(local_found, "ZIP member {member:?} was not found");

    // Preserve structural ZIP admission and make the deferred read fail at
    // verification. A compressed-byte flip can preserve an equivalent stream
    // (for example by changing only a DEFLATE block-final marker), whereas a
    // central-directory CRC mismatch is deterministic.
    let mut cursor = 0_usize;
    while let Some(relative) = bytes[cursor..]
        .windows(4)
        .position(|window| window == b"PK\x01\x02")
    {
        let header = cursor + relative;
        if header + 46 > bytes.len() {
            break;
        }
        let name_length = u16::from_le_bytes([bytes[header + 28], bytes[header + 29]]) as usize;
        let extra_length = u16::from_le_bytes([bytes[header + 30], bytes[header + 31]]) as usize;
        let comment_length = u16::from_le_bytes([bytes[header + 32], bytes[header + 33]]) as usize;
        let name_start = header + 46;
        let name_end = name_start + name_length;
        let record_end = name_end + extra_length + comment_length;
        if record_end > bytes.len() {
            break;
        }
        if &bytes[name_start..name_end] == member {
            bytes[header + 16] ^= 1;
            return bytes;
        }
        cursor = header + 4;
    }
    panic!("central record for ZIP member {member:?} was not found");
}

fn dual_workbook_package() -> (OpcPackage, PackURI, PackURI) {
    let mut package = OpcPackage::new();
    let canonical = PackURI::new("/xl/workbook.bin").unwrap();
    let alternate = PackURI::new("/xl/workbook-alt.bin").unwrap();
    package.add_part(Box::new(BlobPart::new(
        canonical,
        ct::XLSB_BIN.into(),
        b"canonical workbook".to_vec(),
    )));
    package.add_part(Box::new(BlobPart::new(
        alternate.clone(),
        ct::XLSB_BIN.into(),
        b"alternate workbook".to_vec(),
    )));
    package
        .rels_mut()
        .get_or_add(rt::OFFICE_DOCUMENT, "xl/workbook-alt.bin");
    (
        package,
        PackURI::new("/xl/workbook.bin").unwrap(),
        alternate,
    )
}

#[test]
fn deferred_root_workbook_failure_is_not_replaced_by_canonical_fallback() {
    let (package, _canonical, _alternate) = dual_workbook_package();
    let mut bytes = PackageWriter::to_bytes(&package).unwrap();
    bytes = corrupt_zip_member(bytes, "xl/workbook-alt.bin");
    let package = OpcPackage::from_vec(bytes).unwrap();

    let error = super::package::validate_graph(&package).unwrap_err();
    assert!(matches!(
        error,
        crate::package::error::Error::Opc(OpcError::ZipError(_))
    ));
}

#[test]
fn threaded_graph_removal_uses_the_resolved_root_workbook() {
    let (mut package, _canonical, alternate) = dual_workbook_package();
    let persons = PackURI::new("/xl/persons/person1.xml").unwrap();
    package.add_part(Box::new(BlobPart::new(
        persons.clone(),
        PERSONS_CONTENT_TYPE.into(),
        format!(r#"<tc:personList xmlns:tc="{NS}"/>"#).into_bytes(),
    )));
    package
        .get_part_mut(&alternate)
        .unwrap()
        .rels_mut()
        .get_or_add(PERSONS_RELATIONSHIP_TYPE, "persons/person1.xml");

    super::package::validate_graph(&package).unwrap();
    assert!(remove_graph(&mut package).unwrap());
    assert!(package.get_part(&persons).is_err());
    assert!(
        package
            .get_part(&alternate)
            .unwrap()
            .rels()
            .iter()
            .all(|relationship| relationship.reltype() != PERSONS_RELATIONSHIP_TYPE)
    );
}
