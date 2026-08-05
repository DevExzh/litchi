//! Regression coverage for the layered BIFF8 RealTimeData owner.

use super::codec::{
    FRT_HEADER_LEN, HIGH_BYTE, REAL_TIME_DATA_RECORD_TYPE, RTD_OPER_BOOLEAN, RTD_OPER_ERROR,
    RTD_OPER_INTEGER, RTD_OPER_LONG_TEXT, RTD_OPER_NUMBER, RTD_OPER_SHORT_TEXT, read_u16, read_u32,
};
use super::model::{Cell, Record, Value};

/// Build an FrtHeader for the RealTimeData record type.
fn frt_header() -> Vec<u8> {
    let mut header = Vec::with_capacity(FRT_HEADER_LEN);
    header.extend_from_slice(&REAL_TIME_DATA_RECORD_TYPE.to_le_bytes());
    header.extend_from_slice(&[0u8; 10]); // grbitFrt + reserved
    header
}

/// Build a compressed XLUnicodeStringSegmentedRTD from sub-strings.
fn segmented_topic(segments: &[&str]) -> Vec<u8> {
    // cch is the size of the complete compressed rgb field, including
    // the one-byte count prefix for every sub-string.
    let cch: usize = segments.iter().map(|segment| 1 + segment.len()).sum();
    let mut out = Vec::new();
    out.extend_from_slice(&(cch as u32).to_le_bytes());
    out.push(0u8); // fHighByte = 0
    for segment in segments {
        out.push(segment.len() as u8);
        out.extend_from_slice(segment.as_bytes());
    }
    out
}

fn rtd_cell(row: u16, column: u16, sheet_index: u16) -> [u8; 6] {
    let mut cell = [0u8; 6];
    cell[0..2].copy_from_slice(&row.to_le_bytes());
    cell[2..4].copy_from_slice(&column.to_le_bytes());
    cell[4..6].copy_from_slice(&sheet_index.to_le_bytes());
    cell
}

#[test]
fn parses_text_topic_with_cells() {
    let mut payload = frt_header();
    payload.extend_from_slice(&0u32.to_le_bytes()); // ichSamePrefix
    payload.extend_from_slice(&segmented_topic(&["PROG.ID", "", "STOCK", "MSFT"]));
    payload.extend_from_slice(&RTD_OPER_SHORT_TEXT.to_le_bytes());
    payload.extend_from_slice(&5u32.to_le_bytes()); // cchRTDOperStr
    payload.push(0u8); // compressed
    payload.extend_from_slice(b"58.25");
    payload.extend_from_slice(&rtd_cell(1, 2, 0));
    payload.extend_from_slice(&rtd_cell(3, 4, 1));

    let rtd = Record::parse(&payload, None).expect("parse");
    assert_eq!(rtd.common_prefix_len, 0);
    assert_eq!(rtd.topic_segments, vec!["PROG.ID", "", "STOCK", "MSFT"]);
    assert_eq!(rtd.topic, "PROG.IDSTOCKMSFT");
    assert_eq!(rtd.value, Value::Text("58.25".to_string()));
    assert_eq!(
        rtd.cells,
        vec![
            Cell {
                row: 1,
                column: 2,
                sheet_index: 0
            },
            Cell {
                row: 3,
                column: 4,
                sheet_index: 1
            },
        ]
    );
}

#[test]
fn parses_numeric_boolean_error_and_integer_values() {
    for (kind, body, expected) in [
        (
            RTD_OPER_NUMBER,
            42.5f64.to_le_bytes().to_vec(),
            Value::Number(42.5),
        ),
        (
            RTD_OPER_BOOLEAN,
            1u32.to_le_bytes().to_vec(),
            Value::Boolean(true),
        ),
        (
            RTD_OPER_ERROR,
            0x2Au32.to_le_bytes().to_vec(),
            Value::Error(0x2A),
        ),
        (
            RTD_OPER_INTEGER,
            (-7i32).to_le_bytes().to_vec(),
            Value::Integer(-7),
        ),
    ] {
        let mut payload = frt_header();
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&segmented_topic(&["A", "B", "C"]));
        payload.extend_from_slice(&kind.to_le_bytes());
        payload.extend_from_slice(&body);
        let rtd = Record::parse(&payload, None).expect("parse");
        assert_eq!(rtd.value, expected);
        assert!(rtd.cells.is_empty());
    }
}

#[test]
fn rejects_invalid_boolean() {
    let mut payload = frt_header();
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&segmented_topic(&["A", "B", "C"]));
    payload.extend_from_slice(&RTD_OPER_BOOLEAN.to_le_bytes());
    payload.extend_from_slice(&7u32.to_le_bytes());
    assert!(Record::parse(&payload, None).is_err());
}

#[test]
fn rejects_mismatched_rtd_string_kind() {
    for (kind, char_count) in [(RTD_OPER_SHORT_TEXT, 256u32), (RTD_OPER_LONG_TEXT, 5)] {
        let mut payload = frt_header();
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&segmented_topic(&["A", "B", "C"]));
        payload.extend_from_slice(&kind.to_le_bytes());
        payload.extend_from_slice(&char_count.to_le_bytes());
        payload.push(0);
        payload.extend(std::iter::repeat_n(b'x', char_count as usize));
        assert!(Record::parse(&payload, None).is_err());
    }
}

#[test]
fn reapplies_shared_prefix_from_previous_topic() {
    let mut first = frt_header();
    first.extend_from_slice(&0u32.to_le_bytes());
    first.extend_from_slice(&segmented_topic(&["PROG.ID", "", "STOCK"]));
    first.extend_from_slice(&RTD_OPER_INTEGER.to_le_bytes());
    first.extend_from_slice(&1i32.to_le_bytes());
    let first = Record::parse(&first, None).expect("parse first");
    assert_eq!(first.topic, "PROG.IDSTOCK");

    // Second record shares the "PROG.ID" prefix (7 characters) and only
    // stores the trailing sub-strings.
    let mut second = frt_header();
    second.extend_from_slice(&7u32.to_le_bytes()); // ichSamePrefix
    second.extend_from_slice(&segmented_topic(&["", "BOND", "X"]));
    second.extend_from_slice(&RTD_OPER_SHORT_TEXT.to_le_bytes());
    second.extend_from_slice(&3u32.to_le_bytes());
    second.push(0u8);
    second.extend_from_slice(b"102");
    let second = Record::parse(&second, Some(&first.topic)).expect("parse second");
    assert_eq!(second.common_prefix_len, 7);
    assert_eq!(second.topic_segments, vec!["", "BOND", "X"]);
    assert_eq!(second.topic, "PROG.IDBONDX");
    assert_eq!(second.value, Value::Text("102".to_string()));
}

#[test]
fn rejects_prefix_without_previous_topic() {
    let mut payload = frt_header();
    payload.extend_from_slice(&3u32.to_le_bytes());
    payload.extend_from_slice(&segmented_topic(&["A", "B", "C"]));
    payload.extend_from_slice(&RTD_OPER_INTEGER.to_le_bytes());
    payload.extend_from_slice(&0i32.to_le_bytes());
    assert!(Record::parse(&payload, None).is_err());
}

#[test]
fn rejects_prefix_longer_than_previous_topic() {
    let mut payload = frt_header();
    payload.extend_from_slice(&9u32.to_le_bytes());
    payload.extend_from_slice(&segmented_topic(&["A", "B", "C"]));
    payload.extend_from_slice(&RTD_OPER_INTEGER.to_le_bytes());
    payload.extend_from_slice(&0i32.to_le_bytes());
    assert!(Record::parse(&payload, Some("short")).is_err());
}

#[test]
fn parses_wide_strings() {
    let mut topic = Vec::new();
    topic.extend_from_slice(&6u32.to_le_bytes()); // cch includes 3 two-byte count prefixes
    topic.push(1u8); // fHighByte = 1
    // Three wide substrings: 'A', 'B', and '€'.
    for unit in [u32::from('A') as u16, u32::from('B') as u16, 0x20AC] {
        topic.extend_from_slice(&1u16.to_le_bytes());
        topic.extend_from_slice(&unit.to_le_bytes());
    }

    let mut payload = frt_header();
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&topic);
    payload.extend_from_slice(&RTD_OPER_SHORT_TEXT.to_le_bytes());
    payload.extend_from_slice(&1u32.to_le_bytes());
    payload.push(1u8); // wide RTDOperStr
    payload.extend_from_slice(&0x20ACu16.to_le_bytes());

    let rtd = Record::parse(&payload, None).expect("parse");
    assert_eq!(rtd.topic, "AB€");
    assert_eq!(rtd.value, Value::Text("€".to_string()));
}

#[test]
fn rejects_frt_header_rt_mismatch() {
    let mut payload = frt_header();
    payload[0] = 0x12; // corrupt rt to 0x0912... -> mismatch
    payload[1] = 0x09;
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&segmented_topic(&["A", "B", "C"]));
    payload.extend_from_slice(&RTD_OPER_INTEGER.to_le_bytes());
    payload.extend_from_slice(&0i32.to_le_bytes());
    assert!(Record::parse(&payload, None).is_err());
}

#[test]
fn rejects_unknown_oper_kind_and_ragged_cells() {
    let mut payload = frt_header();
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&segmented_topic(&["A", "B", "C"]));
    payload.extend_from_slice(&0xDEADu32.to_le_bytes());
    assert!(Record::parse(&payload, None).is_err());

    let mut payload = frt_header();
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&segmented_topic(&["A", "B", "C"]));
    payload.extend_from_slice(&RTD_OPER_INTEGER.to_le_bytes());
    payload.extend_from_slice(&0i32.to_le_bytes());
    payload.extend_from_slice(&[0u8; 5]); // not a multiple of 6
    assert!(Record::parse(&payload, None).is_err());

    let mut payload = frt_header();
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&segmented_topic(&["A", "B", "C"]));
    payload.extend_from_slice(&RTD_OPER_INTEGER.to_le_bytes());
    payload.extend_from_slice(&0i32.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&256u16.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    assert!(Record::parse(&payload, None).is_err());
}

#[test]
fn rejects_truncated_payloads() {
    assert!(Record::parse(&[], None).is_err());
    assert!(Record::parse(&frt_header(), None).is_err());
    let mut payload = frt_header();
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&segmented_topic(&["A", "B", "C"]));
    // RTDOper kind present but the 4-byte body is missing.
    payload.extend_from_slice(&RTD_OPER_INTEGER.to_le_bytes());
    assert!(Record::parse(&payload, None).is_err());
}

#[test]
fn rejects_overflowing_and_odd_wire_lengths() {
    assert!(read_u16(&[], usize::MAX).is_err());
    assert!(read_u32(&[], usize::MAX).is_err());

    let mut topic = frt_header();
    topic.extend_from_slice(&0u32.to_le_bytes());
    topic.extend_from_slice(&1u32.to_le_bytes());
    topic.push(HIGH_BYTE);
    topic.extend_from_slice(&1u16.to_le_bytes());
    topic.push(0); // Missing the second byte of the UTF-16 code unit.
    assert!(Record::parse(&topic, None).is_err());

    let mut value = frt_header();
    value.extend_from_slice(&0u32.to_le_bytes());
    value.extend_from_slice(&segmented_topic(&[]));
    value.extend_from_slice(&RTD_OPER_SHORT_TEXT.to_le_bytes());
    value.extend_from_slice(&1u32.to_le_bytes());
    value.push(HIGH_BYTE);
    value.push(0); // Missing the second byte of the UTF-16 code unit.
    assert!(Record::parse(&value, None).is_err());

    let mut huge_topic = frt_header();
    huge_topic.extend_from_slice(&0u32.to_le_bytes());
    huge_topic.extend_from_slice(&u32::MAX.to_le_bytes());
    huge_topic.push(0);
    assert!(Record::parse(&huge_topic, None).is_err());
}

#[test]
fn payload_round_trips() {
    let values = [
        Record {
            common_prefix_len: 0,
            topic_segments: vec![
                "PROG.ID".to_string(),
                String::new(),
                "STOCK".to_string(),
                "MSFT".to_string(),
            ],
            topic: "PROG.IDSTOCKMSFT".to_string(),
            value: Value::Text("58.25".to_string()),
            cells: vec![Cell {
                row: 1,
                column: 2,
                sheet_index: 0,
            }],
        },
        Record {
            common_prefix_len: 0,
            topic_segments: vec!["宽".to_string(), "server".to_string(), "€uro".to_string()],
            topic: "宽server€uro".to_string(),
            value: Value::Number(42.5),
            cells: Vec::new(),
        },
        Record {
            common_prefix_len: 0,
            topic_segments: vec!["A".to_string(), "B".to_string(), "C".to_string()],
            topic: "ABC".to_string(),
            value: Value::Boolean(true),
            cells: vec![
                Cell {
                    row: 0,
                    column: 0,
                    sheet_index: 0,
                },
                Cell {
                    row: 65535,
                    column: 255,
                    sheet_index: 3,
                },
            ],
        },
    ];
    for value in values {
        let payload = value.to_payload().expect("serialize");
        let parsed = Record::parse(&payload, None).expect("re-parse");
        assert_eq!(parsed, value);
    }
}

#[test]
fn compressed_latin1_uses_character_count_not_utf8_byte_count() {
    let value = Record {
        common_prefix_len: 0,
        topic_segments: vec!["PROG".to_string(), "server".to_string(), "é".to_string()],
        topic: "PROGserveré".to_string(),
        value: Value::Text("é".to_string()),
        cells: Vec::new(),
    };

    let payload = value.to_payload().expect("serialize");
    let parsed = Record::parse(&payload, None).expect("re-parse");
    assert_eq!(parsed, value);
}

#[test]
fn payload_round_trips_long_text_variant() {
    let long_text = "x".repeat(300);
    let value = Record {
        common_prefix_len: 0,
        topic_segments: vec!["A".to_string(), "B".to_string(), "C".to_string()],
        topic: "ABC".to_string(),
        value: Value::Text(long_text),
        cells: Vec::new(),
    };
    let payload = value.to_payload().expect("serialize");
    // grbit 0x1000 selects the long-string RTDOperStr form.
    let kind_offset = payload.len() - 4 - 4 - 1 - 300;
    assert_eq!(read_u32(&payload, kind_offset).unwrap(), RTD_OPER_LONG_TEXT);
    let parsed = Record::parse(&payload, None).expect("re-parse");
    assert_eq!(parsed, value);
}

#[test]
fn serialize_promotes_long_compressed_segment_to_wide() {
    let value = Record {
        common_prefix_len: 0,
        topic_segments: vec!["A".to_string(), "B".to_string(), "x".repeat(300)],
        topic: String::new(),
        value: Value::Integer(0),
        cells: Vec::new(),
    };
    // A 300-character compressed segment does not fit the 1-byte count,
    // but the wide encoding holds it.
    let payload = value.to_payload().expect("serialize");
    assert_eq!(payload[FRT_HEADER_LEN + 4 + 4], HIGH_BYTE);
}

#[test]
fn segmented_topic_cch_covers_count_prefixes_and_empty_segments() {
    let value = Record {
        common_prefix_len: 0,
        topic_segments: vec![
            "PROG".to_string(),
            String::new(),
            "A".to_string(),
            String::new(),
        ],
        topic: "PROGA".to_string(),
        value: Value::Integer(7),
        cells: Vec::new(),
    };

    let payload = value.to_payload().expect("serialize");
    let topic_offset = FRT_HEADER_LEN + 4;
    // rgb is [4, PROG, 0, 1, A, 0], so cch is 9 encoded bytes.
    assert_eq!(read_u32(&payload, topic_offset).unwrap(), 9);
    assert_eq!(
        &payload[topic_offset + 5..topic_offset + 5 + 9],
        b"\x04PROG\x00\x01A\x00"
    );
    let parsed = Record::parse(&payload, None).expect("parse");
    assert_eq!(parsed, value);
    assert_eq!(parsed.to_payload().unwrap(), payload);
}
