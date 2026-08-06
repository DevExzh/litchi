//! Focused round-trip and malformed-input coverage for document comparison.

use super::*;
use crate::consts::RecordType;
use crate::records::Record;

fn record(version: u16, kind: RecordType, data: Vec<u8>) -> Record {
    Record {
        version,
        instance: 0,
        record_type: kind,
        record_type_raw: kind.as_u16(),
        data_length: data.len() as u32,
        data,
        children: Vec::new(),
    }
}

fn node(diff_type: DiffType, index: bool, children: Vec<DiffNode>) -> DiffNode {
    DiffNode::new(diff_type, index, DiffFlags::for_type(diff_type), children).unwrap()
}

fn record_from_tree(tree: &DiffTree10) -> Record {
    let bytes = tree.to_record_bytes().unwrap();
    Record::parse(&bytes, 0).unwrap().0
}

#[test]
fn diff_enums_are_exhaustive_and_reject_gaps() {
    let values = [
        0, 2, 3, 4, 5, 6, 7, 9, 10, 11, 12, 14, 15, 18, 19, 21, 22, 23,
    ];
    assert!(
        values
            .into_iter()
            .all(|value| DiffType::try_from(value).is_ok())
    );
    assert!(DiffType::try_from(1).is_err());
    assert_eq!(ElementType::try_from(1).unwrap(), ElementType::Shape);
    assert!(ElementType::try_from(0).is_err());
}

#[test]
fn diff_tree_round_trips_and_canonicalizes_ignored_bits() {
    let named_show = node(DiffType::NamedShow, false, vec![]);
    let named_show_list = node(DiffType::NamedShowList, false, vec![named_show]);
    let document = DiffNode::new(
        DiffType::Document,
        false,
        DiffFlags::Document(DocDiffFlags {
            slide_size: true,
            ..Default::default()
        }),
        vec![named_show_list],
    )
    .unwrap();
    let tree = DiffTree10::new("Reviewer".to_string(), document).unwrap();
    let mut bytes = tree.to_record_bytes().unwrap();
    let reviewer_len = 8 + "Reviewer".encode_utf16().count() * 2;
    let doc_flags_offset = 8 + reviewer_len + codec::DIFF_HEADER_SIZE;
    bytes[doc_flags_offset + 3] = 0x80;
    let record = Record::parse(&bytes, 0).unwrap().0;
    let parsed = DiffTree10::parse(&record).unwrap();
    assert_eq!(parsed.reviewer_name(), "Reviewer");
    assert_eq!(parsed.document_diff().ignored_flag_bits(), 0x8000_0000);
    let canonical = parsed.to_record_bytes().unwrap();
    assert_eq!(canonical[doc_flags_offset + 3] & 0x80, 0);
}

#[test]
fn malformed_tag_and_child_order_are_rejected() {
    let document = node(DiffType::Document, false, vec![]);
    let tree = DiffTree10::new("R".to_string(), document).unwrap();
    let mut bytes = tree.to_record_bytes().unwrap();
    let tag_offset = 8 + 8 + 2 + 20;
    bytes[tag_offset..tag_offset + 4].copy_from_slice(&1u32.to_le_bytes());
    let record = Record::parse(&bytes, 0).unwrap().0;
    assert!(DiffTree10::parse(&record).is_err());

    let wrong_child = node(DiffType::Text, false, vec![]);
    assert!(
        DiffNode::new(
            DiffType::Document,
            false,
            DiffFlags::for_type(DiffType::Document),
            vec![wrong_child],
        )
        .is_err()
    );
}

#[test]
fn depth_and_record_count_limits_are_enforced() {
    let text = node(DiffType::Text, false, vec![]);
    let shape = node(DiffType::Shape, false, vec![text]);
    let shape_list = node(DiffType::ShapeList, false, vec![shape]);
    let slide = node(DiffType::Slide, false, vec![shape_list]);
    let slide_list = node(DiffType::SlideList, false, vec![slide]);
    let document = node(DiffType::Document, false, vec![slide_list]);
    let tree = DiffTree10::new("R".to_string(), document).unwrap();
    let record = record_from_tree(&tree);
    assert!(DiffTree10::parse_with_limits(&record, 2, 100).is_err());
    assert!(DiffTree10::parse_with_limits(&record, 32, 3).is_err());
}

#[test]
fn reviewing_toolbar_round_trips_ignored_bits_and_mutation() {
    let parsed =
        ReviewingToolbarStates::parse(&record(0, RecordType::DocToolbarStates10Atom, vec![0xfd]))
            .unwrap();
    assert!(parsed.show_reviewing_toolbar());
    assert!(!parsed.show_reviewing_gallery());
    assert_eq!(parsed.ignored_reserved_bits(), 0xfc);
    assert_eq!(parsed.to_record_bytes()[8], 0xfd);

    let mut created = ReviewingToolbarStates::new(false, false);
    created.set_show_reviewing_gallery(true);
    assert_eq!(created.to_record_bytes()[8], 0x02);
}

#[test]
fn slide_list_table_round_trips_filetime_order_and_entries() {
    let table = SlideListTable10::new(vec![
        SlideCreationEntry::new(7, 0x1122_3344_5566_7788),
        SlideCreationEntry::new(u32::MAX, u64::MAX),
    ])
    .unwrap();
    let bytes = table.to_record_bytes().unwrap();
    let parsed_record = record(0x0f, RecordType::SlideListTable10, bytes[8..].to_vec());
    let parsed = SlideListTable10::parse(&parsed_record).unwrap();
    assert_eq!(parsed, table);
    assert_eq!(bytes[32..36], 0x1122_3344u32.to_le_bytes());
    assert_eq!(bytes[36..40], 0x5566_7788u32.to_le_bytes());
}

#[test]
fn rejects_bad_headers_counts_order_truncation_and_trailing_data() {
    let entry = SlideCreationEntry::new(1, 2);
    let table = SlideListTable10::new(vec![entry]).unwrap();
    let bytes = table.to_record_bytes().unwrap();
    let payload = bytes[8..].to_vec();

    let mut cases = Vec::new();
    let mut negative = payload.clone();
    negative[8..12].copy_from_slice(&(-1i32).to_le_bytes());
    cases.push(negative);
    let mut mismatch = payload.clone();
    mismatch[8..12].copy_from_slice(&2i32.to_le_bytes());
    cases.push(mismatch);
    let mut wrong_child = payload.clone();
    wrong_child[14..16]
        .copy_from_slice(&RecordType::SlideListTableSize10Atom.as_u16().to_le_bytes());
    cases.push(wrong_child);
    cases.push(payload[..payload.len() - 1].to_vec());
    let mut trailing = payload.clone();
    trailing.push(0);
    cases.push(trailing);

    for data in cases {
        assert!(
            SlideListTable10::parse(&record(0x0f, RecordType::SlideListTable10, data,)).is_err()
        );
    }
    assert!(SlideListTable10::parse(&record(0, RecordType::SlideListTable10, payload,)).is_err());
}

#[test]
fn rejects_atom_children_and_oversized_builder() {
    let mut atom = record(0, RecordType::DocToolbarStates10Atom, vec![0]);
    atom.children
        .push(record(0, RecordType::Unknown, Vec::new()));
    assert!(ReviewingToolbarStates::parse(&atom).is_err());

    let oversized = vec![SlideCreationEntry::new(0, 0); validation::MAX_SLIDE_LIST_ENTRIES + 1];
    assert!(SlideListTable10::new(oversized).is_err());
}
