use litchi_ooxml::xlsx::{
    Mention, Person, ThreadedComment, add_threaded_comment, add_threaded_comment_person,
    add_threaded_comment_reply, find_threaded_comment, find_threaded_comment_person,
    load_threaded_comment_graph, remove_threaded_comment, remove_threaded_comment_person,
    reorder_threaded_comment_persons, reorder_threaded_comments, replace_threaded_comment,
    replace_threaded_comment_person, update_threaded_comment, update_threaded_comment_person,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI, Part};

const PERSON_A: &str = "{11111111-1111-4111-8111-111111111111}";
const PERSON_B: &str = "{22222222-2222-4222-8222-222222222222}";
const COMMENT_A: &str = "{33333333-3333-4333-8333-333333333333}";
const COMMENT_B: &str = "{44444444-4444-4444-8444-444444444444}";
const REPLY_A: &str = "{55555555-5555-4555-8555-555555555555}";
const MENTION_A: &str = "{66666666-6666-4666-8666-666666666666}";

fn package_with_legacy_note() -> (OpcPackage, PackURI, PackURI, PackURI) {
    let mut package = OpcPackage::new();
    let workbook_name = PackURI::new("/xl/workbook.xml").unwrap();
    let worksheet_name = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    let comments_name = PackURI::new("/xl/comments1.xml").unwrap();
    let vml_name = PackURI::new("/xl/drawings/vmlDrawing1.vml").unwrap();
    let mut workbook = BlobPart::new(
        workbook_name,
        ct::SML_SHEET_MAIN.into(),
        br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#.to_vec(),
    );
    workbook.relate_to("worksheets/sheet1.xml", rt::WORKSHEET);
    let mut worksheet = BlobPart::new(
        worksheet_name.clone(),
        ct::SML_WORKSHEET.into(),
        br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><legacyDrawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdVml"/></worksheet>"#.to_vec(),
    );
    worksheet.rels_mut().add_relationship(
        rt::COMMENTS.into(),
        "../comments1.xml".into(),
        "rIdLegacyComments".into(),
        false,
    );
    worksheet.rels_mut().add_relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing".into(),
        "../drawings/vmlDrawing1.vml".into(),
        "rIdVml".into(),
        false,
    );
    package.add_part(Box::new(workbook));
    package.add_part(Box::new(worksheet));
    package.add_part(Box::new(BlobPart::new(
        comments_name.clone(),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml".into(),
        br#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Legacy</author></authors><commentList><comment ref="A1" authorId="0"><text><t>note</t></text></comment></commentList></comments>"#.to_vec(),
    )));
    package.add_part(Box::new(BlobPart::new(
        vml_name.clone(),
        "application/vnd.openxmlformats-officedocument.vmlDrawing".into(),
        br#"<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="note"/></xml>"#.to_vec(),
    )));
    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
    (package, worksheet_name, comments_name, vml_name)
}

fn person(id: &str, name: &str) -> Person {
    Person {
        display_name: name.into(),
        id: id.into(),
        user_id: Some(format!("{}@example.test", name.to_ascii_lowercase())),
        provider_id: Some("local-test".into()),
    }
}

fn root(id: &str, person_id: &str, cell_ref: &str, text: &str) -> ThreadedComment {
    ThreadedComment {
        cell_ref: Some(cell_ref.into()),
        id: id.into(),
        parent_id: None,
        person_id: person_id.into(),
        text: Some(text.into()),
        date_time: Some("2026-07-19T12:00:00Z".into()),
        done: Some(false),
        mentions: Vec::new(),
    }
}

fn reply(id: &str, person_id: &str, text: &str) -> ThreadedComment {
    ThreadedComment {
        cell_ref: None,
        id: id.into(),
        parent_id: None,
        person_id: person_id.into(),
        text: Some(text.into()),
        date_time: Some("2026-07-19T12:01:00Z".into()),
        done: None,
        mentions: Vec::new(),
    }
}

#[test]
fn package_crud_keeps_threads_people_and_legacy_notes_consistent() {
    let (mut package, worksheet, legacy_comments, legacy_vml) = package_with_legacy_note();
    let legacy_comments_blob = package.get_part(&legacy_comments).unwrap().blob().to_vec();
    let legacy_vml_blob = package.get_part(&legacy_vml).unwrap().blob().to_vec();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/persons/person.xml").unwrap(),
        "application/octet-stream".into(),
        b"occupied-person-name".to_vec(),
    )));
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/threadedComments/threadedComment1.xml").unwrap(),
        "application/octet-stream".into(),
        b"occupied-comment-name".to_vec(),
    )));
    package
        .get_part_mut(&PackURI::new("/xl/workbook.xml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            "urn:test:occupied-person-rid".into(),
            "occupied-person.bin".into(),
            "rIdPersons1".into(),
            false,
        );
    package
        .get_part_mut(&worksheet)
        .unwrap()
        .rels_mut()
        .add_relationship(
            "urn:test:occupied-thread-rid".into(),
            "occupied-thread.bin".into(),
            "rIdThreadedComments1".into(),
            false,
        );

    let persons_part = add_threaded_comment_person(&mut package, person(PERSON_A, "Ada")).unwrap();
    assert_ne!(persons_part.part_name, "/xl/persons/person.xml");
    add_threaded_comment_person(&mut package, person(PERSON_B, "Grace")).unwrap();
    let mut first = root(COMMENT_A, PERSON_A, "B2", "Hi @Grace");
    first.mentions.push(Mention {
        mention_person_id: PERSON_B.into(),
        mention_id: MENTION_A.into(),
        start_index: 3,
        length: 6,
    });
    let first_part = add_threaded_comment(&mut package, &worksheet, first).unwrap();
    assert_ne!(first_part.part_name, "/xl/threadedComments/threadedComment1.xml");
    add_threaded_comment(
        &mut package,
        &worksheet,
        root(COMMENT_B, PERSON_B, "C3", "Second thread"),
    )
    .unwrap();
    add_threaded_comment_reply(
        &mut package,
        &worksheet,
        "B2",
        COMMENT_A,
        reply(REPLY_A, PERSON_B, "Reply"),
    )
    .unwrap();

    assert_eq!(
        find_threaded_comment(&package, &worksheet, "B2", REPLY_A)
            .unwrap()
            .unwrap()
            .parent_id
            .as_deref(),
        Some(COMMENT_A)
    );
    assert!(remove_threaded_comment(&mut package, &worksheet, "B2", COMMENT_A).is_err());
    assert!(remove_threaded_comment_person(&mut package, PERSON_B).is_err());
    assert!(add_threaded_comment(
        &mut package,
        &worksheet,
        root(
            "{77777777-7777-4777-8777-777777777777}",
            "{88888888-8888-4888-8888-888888888888}",
            "D4",
            "unresolved"
        )
    )
    .is_err());

    update_threaded_comment(&mut package, &worksheet, "B2", COMMENT_A, |comment| {
        comment.text = Some("Updated".into());
        comment.mentions.clear();
        comment.done = Some(true);
    })
    .unwrap();
    let mut replacement = root(COMMENT_B, PERSON_B, "C3", "Replacement");
    replacement.done = Some(true);
    replace_threaded_comment(&mut package, &worksheet, "C3", COMMENT_B, replacement).unwrap();
    reorder_threaded_comments(
        &mut package,
        &worksheet,
        &[COMMENT_B.into(), COMMENT_A.into(), REPLY_A.into()],
    )
    .unwrap();
    reorder_threaded_comment_persons(&mut package, &[PERSON_B.into(), PERSON_A.into()]).unwrap();
    update_threaded_comment_person(&mut package, PERSON_A, |person| {
        person.display_name = "Ada Lovelace".into();
    })
    .unwrap();
    replace_threaded_comment_person(&mut package, PERSON_A, person(PERSON_A, "Ada Lovelace"))
        .unwrap();
    assert_eq!(
        find_threaded_comment_person(&package, PERSON_A)
            .unwrap()
            .unwrap()
            .display_name,
        "Ada Lovelace"
    );

    assert!(remove_threaded_comment(&mut package, &worksheet, "B2", REPLY_A).unwrap());
    assert!(remove_threaded_comment(&mut package, &worksheet, "C3", COMMENT_B).unwrap());
    assert!(remove_threaded_comment_person(&mut package, PERSON_B).unwrap());

    let threaded_part = PackURI::new(&first_part.part_name).unwrap();
    let mut shared_owner = BlobPart::new(
        PackURI::new("/xl/shared-owner.xml").unwrap(),
        "application/xml".into(),
        Vec::new(),
    );
    shared_owner.relate_to(
        &threaded_part.relative_ref("/xl/"),
        "urn:test:shared-threaded-comments",
    );
    package.add_part(Box::new(shared_owner));
    assert!(remove_threaded_comment(&mut package, &worksheet, "B2", COMMENT_A).unwrap());
    assert!(package.get_part(&threaded_part).is_ok());
    assert!(remove_threaded_comment_person(&mut package, PERSON_A).unwrap());

    assert_eq!(package.get_part(&legacy_comments).unwrap().blob(), legacy_comments_blob);
    assert_eq!(package.get_part(&legacy_vml).unwrap().blob(), legacy_vml_blob);
    assert!(load_threaded_comment_graph(&package).unwrap().worksheets.is_empty());
}

#[test]
fn failed_timestamp_and_identity_mutations_are_atomic() {
    let (mut package, worksheet, _, _) = package_with_legacy_note();
    add_threaded_comment_person(&mut package, person(PERSON_A, "Ada")).unwrap();
    let mut oversized = person(PERSON_B, "Oversized");
    oversized.user_id = Some("x".repeat(16_385));
    assert!(add_threaded_comment_person(&mut package, oversized).is_err());
    add_threaded_comment(
        &mut package,
        &worksheet,
        root(COMMENT_A, PERSON_A, "A1", "original"),
    )
    .unwrap();
    let graph_before = load_threaded_comment_graph(&package).unwrap();
    assert!(update_threaded_comment(&mut package, &worksheet, "A1", COMMENT_A, |comment| {
        comment.date_time = Some("not-a-timestamp".into());
    })
    .is_err());
    assert!(update_threaded_comment_person(&mut package, PERSON_A, |person| {
        person.id = PERSON_B.into();
    })
    .is_err());
    let graph_after = load_threaded_comment_graph(&package).unwrap();
    assert_eq!(
        graph_after.worksheets[0].comments.comments[0].text,
        graph_before.worksheets[0].comments.comments[0].text
    );
    assert_eq!(
        graph_after.persons.unwrap().persons.persons[0].id,
        PERSON_A
    );
}
