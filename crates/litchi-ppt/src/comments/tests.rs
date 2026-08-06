use super::{Authors, Catalog, parse_slide_comments};
use crate::consts::RecordType;
use crate::presentation::ParsedSlideComments;
use crate::records::Record;

fn record_bytes(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    data.extend_from_slice(&kind.to_le_bytes());
    data.extend_from_slice(&(payload.len() as u32).to_le_bytes());
    data.extend_from_slice(payload);
    data
}

fn utf16(value: &str) -> Vec<u8> {
    value.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

fn prog_tags_record(version: u8, blob_payload: &[u8]) -> Record {
    let name = record_bytes(0, 0, 4026, &utf16(&format!("___PPT{version}")));
    let blob = record_bytes(0, 0, 0x138b, blob_payload);
    let mut tag_payload = name;
    tag_payload.extend_from_slice(&blob);
    let tag = record_bytes(0x0f, 0, 0x138a, &tag_payload);
    Record {
        record_type: RecordType::ProgTags,
        record_type_raw: 0x1388,
        version: 0x0f,
        instance: 0,
        data_length: tag.len() as u32,
        data: tag,
        children: Vec::new(),
    }
}

fn root(children: Vec<Record>) -> Record {
    Record {
        record_type: RecordType::Document,
        record_type_raw: 1000,
        version: 0x0f,
        instance: 0,
        data_length: 0,
        data: Vec::new(),
        children,
    }
}

fn comment_atom(index: i32) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&index.to_le_bytes());
    for value in [2026, 7, 4, 16, 13, 14, 15, 999] {
        data.extend_from_slice(&u16::try_from(value).unwrap().to_le_bytes());
    }
    data.extend_from_slice(&(-12i32).to_le_bytes());
    data.extend_from_slice(&34i32.to_le_bytes());
    record_bytes(0, 0, 12001, &data)
}

fn comment_container(index: i32) -> Vec<u8> {
    let mut children = record_bytes(0, 0, 4026, &utf16("Ada Lovelace"));
    children.extend_from_slice(&record_bytes(
        0,
        1,
        4026,
        &utf16("First\tline\r\nSecond line"),
    ));
    children.extend_from_slice(&record_bytes(0, 2, 4026, &utf16("AL")));
    children.extend_from_slice(&comment_atom(index));
    record_bytes(0x0f, 0, 12000, &children)
}

fn author_container(seed: i32) -> Vec<u8> {
    let mut children = record_bytes(0, 0, 4026, &utf16("Ada Lovelace"));
    let mut index = 3i32.to_le_bytes().to_vec();
    index.extend_from_slice(&seed.to_le_bytes());
    children.extend_from_slice(&record_bytes(0, 0, 12005, &index));
    record_bytes(0x0f, 0, 12004, &children)
}

#[test]
fn parses_comments_and_author_metadata() {
    let comment_root = root(vec![prog_tags_record(10, &comment_container(7))]);
    let comments = parse_slide_comments(&comment_root).unwrap();
    assert_eq!(comments.len(), 1);
    let comment = &comments[0];
    assert_eq!(comment.author, "Ada Lovelace");
    assert_eq!(comment.text, "First\tline\r\nSecond line");
    assert_eq!(comment.index, 7);
    assert_eq!(comment.year, 2026);
    assert_eq!(comment.day_of_week, 4);
    assert_eq!(comment.millisecond, 999);
    assert_eq!((comment.x, comment.y), (-12, 34));

    let author_root = root(vec![prog_tags_record(10, &author_container(7))]);
    let authors = Authors::parse(&author_root).unwrap();
    let author = authors.find("Ada Lovelace").unwrap();
    assert_eq!(author.color_index, Some(3));
    assert_eq!(author.comment_index_seed, Some(7));
    authors.validate_comments(&comments).unwrap();
}

#[test]
fn rejects_comment_indices_above_author_seed() {
    let comments =
        parse_slide_comments(&root(vec![prog_tags_record(10, &comment_container(8))])).unwrap();
    let authors = Authors::parse(&root(vec![prog_tags_record(10, &author_container(7))])).unwrap();

    assert!(authors.validate_comments(&comments).is_err());
}

#[test]
fn catalog_joins_slide_comments_and_author_seeds() {
    let comments =
        parse_slide_comments(&root(vec![prog_tags_record(10, &comment_container(7))])).unwrap();
    let authors = Authors::parse(&root(vec![prog_tags_record(10, &author_container(7))])).unwrap();
    let catalog = Catalog::from_parts(
        authors,
        vec![ParsedSlideComments {
            slide_number: 2,
            comments,
        }],
    )
    .unwrap();

    assert_eq!(catalog.authors().len(), 1);
    assert_eq!(catalog.slides().len(), 1);
    assert_eq!(catalog.comments().count(), 1);
    assert_eq!(catalog.slide(2).unwrap().comments[0].index, 7);
    assert!(catalog.slide(1).is_none());
}

#[test]
fn ignores_comments_from_other_programmable_tag_versions() {
    let document = root(vec![prog_tags_record(9, &comment_container(7))]);

    assert!(parse_slide_comments(&document).unwrap().is_empty());
    assert!(Authors::parse(&document).unwrap().authors.is_empty());
}

#[test]
fn rejects_malformed_comment_containers() {
    let mut out_of_order = record_bytes(0, 1, 4026, &utf16("text"));
    out_of_order.extend_from_slice(&record_bytes(0, 0, 4026, &utf16("author")));
    out_of_order.extend_from_slice(&comment_atom(0));
    let mut forbidden_text = record_bytes(0, 1, 4026, &[0x0b, 0]);
    forbidden_text.extend_from_slice(&comment_atom(0));
    let mut duplicate_atom = comment_atom(0);
    duplicate_atom.extend_from_slice(&comment_atom(1));
    let malformed = [
        record_bytes(0x0e, 0, 12000, &comment_atom(0)),
        record_bytes(0x0f, 0, 12000, &[]),
        record_bytes(0x0f, 0, 12000, &comment_atom(-1)),
        record_bytes(0x0f, 0, 12000, &out_of_order),
        record_bytes(0x0f, 0, 12000, &forbidden_text),
        record_bytes(0x0f, 0, 12000, &duplicate_atom),
    ];
    for record in malformed {
        let document = root(vec![prog_tags_record(10, &record)]);
        assert!(parse_slide_comments(&document).is_err());
    }
}

#[test]
fn rejects_malformed_comment_authors() {
    let name = record_bytes(0, 0, 4026, &utf16("Ada Lovelace"));
    let mut negative_color = (-1i32).to_le_bytes().to_vec();
    negative_color.extend_from_slice(&0i32.to_le_bytes());
    let atom = record_bytes(0, 0, 12005, &negative_color);
    let mut out_of_order = atom.clone();
    out_of_order.extend_from_slice(&name);
    let malformed = [
        record_bytes(0x0e, 0, 12004, &name),
        record_bytes(0x0f, 0, 12004, &atom),
        record_bytes(0x0f, 0, 12004, &out_of_order),
        record_bytes(0x0f, 0, 12004, &record_bytes(0, 0, 4026, &[0x01, 0])),
    ];
    for record in malformed {
        let document = root(vec![prog_tags_record(10, &record)]);
        assert!(Authors::parse(&document).is_err());
    }
}
