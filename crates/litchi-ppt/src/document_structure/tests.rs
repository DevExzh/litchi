use super::*;
use crate::consts::RecordType;
use crate::records::Record;

fn record(record_type: RecordType, version: u16, data: Vec<u8>) -> Record {
    Record {
        record_type,
        record_type_raw: record_type.as_u16(),
        version,
        instance: 0,
        data_length: data.len() as u32,
        data,
        children: Vec::new(),
    }
}

fn document(children: Vec<Record>) -> Record {
    let mut value = record(RecordType::Document, 0x0f, Vec::new());
    value.children = children;
    value
}

#[test]
fn accepts_both_defined_custom_table_style_placements() {
    let prefix = record(RecordType::DocumentAtom, 0, Vec::new());
    let end = record(RecordType::EndDocument, 0, Vec::new());
    let styles = record(RecordType::RoundTripCustomTableStyles12Atom, 0, Vec::new());
    assert_eq!(
        DocumentStructure::parse(&document(vec![prefix.clone(), styles.clone(), end.clone()]))
            .unwrap()
            .custom_table_styles,
        Some(CustomTableStylesPlacement::BeforeEndDocument)
    );
    assert_eq!(
        DocumentStructure::parse(&document(vec![prefix, end, styles]))
            .unwrap()
            .custom_table_styles,
        Some(CustomTableStylesPlacement::AfterEndDocument)
    );
}

#[test]
fn rejects_missing_duplicate_nonempty_and_nonterminal_end_records() {
    let end = record(RecordType::EndDocument, 0, Vec::new());
    assert!(DocumentStructure::parse(&document(Vec::new())).is_err());
    assert!(DocumentStructure::parse(&document(vec![end.clone(), end.clone()])).is_err());
    assert!(
        DocumentStructure::parse(&document(vec![
            record(RecordType::EndDocument, 0, vec![0],)
        ]))
        .is_err()
    );
    assert!(
        DocumentStructure::parse(&document(vec![
            end,
            record(RecordType::DocumentAtom, 0, Vec::new()),
        ]))
        .is_err()
    );
}
