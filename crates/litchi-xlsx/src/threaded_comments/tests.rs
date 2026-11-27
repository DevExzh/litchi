use super::{
    Comment, Comments, Mention, People, Person, parse_comments, parse_persons, write_comments,
    write_persons,
};

const NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
const PERSON_ID: &str = "{11111111-1111-1111-1111-111111111111}";
const COMMENT_ID: &str = "{22222222-2222-2222-2222-222222222222}";
const BOB_ID: &str = "{33333333-3333-3333-3333-333333333333}";
const REPLY_ID: &str = "{33333333-3333-3333-3333-333333333333}";
const MENTION_ID: &str = "{44444444-4444-4444-4444-444444444444}";

fn assert_compact_xml(xml: &str) {
    assert!(
        !xml.bytes()
            .any(|byte| matches!(byte, b'\n' | b'\r' | b'\t'))
    );
    assert!(!xml.contains("> <"));
}

#[test]
fn parses_prefixed_people_and_threaded_comments() {
    let people = parse_persons(&format!(
        r#"<tc:personList xmlns:tc="{NS}" xmlns:f="urn:foreign">
            <f:person displayName="Ignored" id="ignored"/>
            <tc:person displayName="Alice &amp; Bob" id="{PERSON_ID}" userId="alice" providerId="aad"/>
        </tc:personList>"#
    ))
    .unwrap();
    assert_eq!(people.persons.len(), 1);
    assert_eq!(people.persons[0].display_name, "Alice & Bob");

    let comments = parse_comments(&format!(
        r#"<tc:ThreadedComments xmlns:tc="{NS}" xmlns:f="urn:foreign">
            <f:threadedComment id="ignored" personId="ignored"/>
            <tc:threadedComment ref="B2" id="{COMMENT_ID}" personId="{PERSON_ID}" dT="2026-07-14T10:00:00Z" done="true">
                <tc:text>Hello &amp; @Bob</tc:text>
                <tc:mentions><tc:mention mentionpersonId="{BOB_ID}" mentionId="{MENTION_ID}" startIndex="8" length="4"/></tc:mentions>
            </tc:threadedComment>
        </tc:ThreadedComments>"#
    ))
    .unwrap();
    let comment = &comments.comments[0];
    assert_eq!(comment.cell_ref.as_deref(), Some("B2"));
    assert_eq!(comment.text.as_deref(), Some("Hello & @Bob"));
    assert_eq!(comment.done, Some(true));
    assert_eq!(comment.mentions.len(), 1);
}

#[test]
fn accepts_empty_present_parts() {
    let people = parse_persons(&format!(r#"<personList xmlns="{NS}"/>"#)).unwrap();
    let comments = parse_comments(&format!(r#"<ThreadedComments xmlns="{NS}"/>"#)).unwrap();

    assert!(people.persons.is_empty());
    assert!(comments.comments.is_empty());
}

#[test]
fn rejects_malformed_threaded_parts() {
    let invalid_people = [
        r#"<personList xmlns="urn:foreign"/>"#.to_string(),
        format!(r#"<personList xmlns="{NS}"><person id="x"/></personList>"#),
        format!(
            r#"<personList xmlns="{NS}"><person displayName="A" id="{PERSON_ID}"/><person displayName="B" id="{PERSON_ID}"/></personList>"#
        ),
    ];
    for xml in invalid_people {
        assert!(parse_persons(&xml).is_err(), "accepted {xml}");
    }

    let invalid_comments = [
        r#"<ThreadedComments xmlns="urn:foreign"/>"#.to_string(),
        format!(
            r#"<ThreadedComments xmlns="{NS}"><threadedComment personId="p"/></ThreadedComments>"#
        ),
        format!(
            r#"<ThreadedComments xmlns="{NS}"><threadedComment id="{COMMENT_ID}" personId="{PERSON_ID}" done="yes"/></ThreadedComments>"#
        ),
        format!(
            r#"<ThreadedComments xmlns="{NS}"><threadedComment ref="A0" id="{COMMENT_ID}" personId="{PERSON_ID}"/></ThreadedComments>"#
        ),
        format!(
            r#"<ThreadedComments xmlns="{NS}"><threadedComment id="{COMMENT_ID}" personId="{PERSON_ID}"><mentions/><text>x</text></threadedComment></ThreadedComments>"#
        ),
        format!(
            r#"<ThreadedComments xmlns="{NS}"><threadedComment id="{COMMENT_ID}" personId="{PERSON_ID}"><text>x</text><text>y</text></threadedComment></ThreadedComments>"#
        ),
        format!(
            r#"<ThreadedComments xmlns="{NS}"><threadedComment id="{COMMENT_ID}" personId="{PERSON_ID}"><text>x</text><mentions><mention mentionpersonId="{PERSON_ID}" mentionId="{MENTION_ID}" startIndex="1" length="1"/></mentions></threadedComment></ThreadedComments>"#
        ),
    ];
    for xml in invalid_comments {
        assert!(parse_comments(&xml).is_err(), "accepted {xml}");
    }
}

#[test]
fn writes_schema_valid_people_and_comments() {
    let people = People {
        persons: vec![Person {
            display_name: "Alice & Bob".into(),
            id: PERSON_ID.into(),
            user_id: Some("alice@example.com".into()),
            provider_id: None,
        }],
    };
    let people_xml = write_persons(&people).unwrap();
    assert_compact_xml(&people_xml);
    assert!(people_xml.contains("Alice &amp; Bob"));

    let comments = Comments {
        comments: vec![
            Comment {
                cell_ref: Some("A1".into()),
                id: COMMENT_ID.into(),
                person_id: PERSON_ID.into(),
                text: Some("Hi @Bob".into()),
                mentions: vec![Mention {
                    mention_person_id: PERSON_ID.into(),
                    mention_id: MENTION_ID.into(),
                    start_index: 3,
                    length: 4,
                }],
                ..Default::default()
            },
            Comment {
                id: REPLY_ID.into(),
                person_id: PERSON_ID.into(),
                parent_id: Some(COMMENT_ID.into()),
                ..Default::default()
            },
        ],
    };
    let comments_xml = write_comments(&comments).unwrap();
    assert_compact_xml(&comments_xml);
    assert!(comments_xml.contains("<text>Hi @Bob</text>"));
    assert!(comments_xml.contains(&format!(r#" parentId="{COMMENT_ID}""#)));
}

#[test]
fn rejects_invalid_people_and_comments() {
    let duplicate_people = People {
        persons: vec![
            Person {
                id: PERSON_ID.into(),
                ..Default::default()
            },
            Person {
                id: PERSON_ID.into(),
                ..Default::default()
            },
        ],
    };
    assert!(write_persons(&duplicate_people).is_err());

    let invalid = [
        Comment {
            id: "not-a-guid".into(),
            person_id: PERSON_ID.into(),
            ..Default::default()
        },
        Comment {
            cell_ref: Some("A0".into()),
            id: COMMENT_ID.into(),
            person_id: PERSON_ID.into(),
            ..Default::default()
        },
        Comment {
            id: COMMENT_ID.into(),
            person_id: PERSON_ID.into(),
            parent_id: Some(REPLY_ID.into()),
            ..Default::default()
        },
        Comment {
            id: COMMENT_ID.into(),
            person_id: PERSON_ID.into(),
            text: Some("x".into()),
            mentions: vec![Mention {
                mention_person_id: PERSON_ID.into(),
                mention_id: MENTION_ID.into(),
                start_index: 1,
                length: 1,
            }],
            ..Default::default()
        },
    ];
    for comment in invalid {
        assert!(
            write_comments(&Comments {
                comments: vec![comment]
            })
            .is_err()
        );
    }
}
