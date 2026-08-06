use super::{Bookmark, Summary};
use crate::consts::RecordType;
use crate::records::Record;

fn root(children: Vec<Record>) -> Record {
    Record {
        version: 0x0f,
        instance: 0,
        record_type: RecordType::Document,
        record_type_raw: RecordType::Document.as_u16(),
        data_length: 0,
        data: Vec::new(),
        children,
    }
}

fn summary() -> Summary {
    Summary {
        id_seed: 43,
        bookmarks: vec![
            Bookmark {
                container_instance: 7,
                id: 41,
                name: "Revenue".into(),
                value: "FY 2026".into(),
            },
            Bookmark {
                container_instance: 0,
                id: 42,
                name: "EmptyValue".into(),
                value: String::new(),
            },
        ],
    }
}

#[test]
fn protocol_shaped_bookmark_summary_roundtrips() {
    let expected = summary();
    let parsed = Summary::parse(&root(vec![expected.to_record().unwrap()]))
        .unwrap()
        .unwrap();
    assert_eq!(parsed, expected);
    assert_eq!(
        parsed.to_record_bytes().unwrap(),
        expected.to_record_bytes().unwrap()
    );
    parsed.validate_text_bookmark_ids([41, 42]).unwrap();
}

#[test]
fn rejects_hostile_ids_names_values_and_seed() {
    let record = summary().to_record().unwrap();
    assert!(Summary::parse(&root(vec![record.clone(), record])).is_err());
    let mut value = summary();
    value.id_seed = 42;
    assert!(value.to_record_bytes().is_err());
    value = summary();
    value.bookmarks[1].id = 41;
    assert!(value.to_record_bytes().is_err());
    value = summary();
    value.bookmarks[0].name.clear();
    assert!(value.to_record_bytes().is_err());
    value = summary();
    value.bookmarks[0].value = "bad\nvalue".into();
    assert!(value.to_record_bytes().is_err());
    value = summary();
    assert!(value.validate_text_bookmark_ids([41]).is_err());
    assert!(value.validate_text_bookmark_ids([41, 42, 99]).is_err());
}
