use super::package::{
    COMMENTS_CONTENT_TYPE, COMMENTS_RELATIONSHIP_TYPE, PERSONS_CONTENT_TYPE,
    PERSONS_RELATIONSHIP_TYPE, WORKSHEET_CONTENT_TYPE,
};
use super::{
    Comment, Comments, CommentsPart, Graph, Mention, People, PeoplePart, Person, Thread,
    load_graph, parse_comments, parse_persons, remove_graph, store_graph, validate_comments,
    validate_graph as validate_model_graph, write_comments, write_persons,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, TargetMode};

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
