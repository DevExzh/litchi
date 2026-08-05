use litchi_xlsx::threaded_comments::{
    Comment, Comments, People, Person, parse_comments, parse_persons, validate_comments,
    validate_people, write_comments, write_persons,
};

#[test]
fn threaded_comment_models_round_trip_through_the_bounded_codecs() {
    let people = People {
        persons: vec![Person {
            display_name: "Ada".into(),
            id: "{11111111-1111-4111-8111-111111111111}".into(),
            user_id: Some("ada@example.test".into()),
            provider_id: Some("local".into()),
        }],
    };
    validate_people(&people).unwrap();
    let people_xml = write_persons(&people).unwrap();
    assert_eq!(parse_persons(&people_xml).unwrap().persons.len(), 1);

    let comments = Comments {
        comments: vec![Comment {
            cell_ref: Some("B2".into()),
            id: "{22222222-2222-4222-8222-222222222222}".into(),
            parent_id: None,
            person_id: people.persons[0].id.clone(),
            text: Some("Hello".into()),
            date_time: Some("2026-07-19T12:00:00Z".into()),
            done: Some(false),
            mentions: Vec::new(),
        }],
    };
    validate_comments(&comments).unwrap();
    let comments_xml = write_comments(&comments).unwrap();
    assert_eq!(parse_comments(&comments_xml).unwrap().comments.len(), 1);
}
