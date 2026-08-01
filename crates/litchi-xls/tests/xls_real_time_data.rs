//! Round-trip tests for the BIFF8 RealTimeData record (RTD topics).

use litchi_xls::writer::XlsWriter;
use litchi_xls::{XlsRealTimeData, XlsRtdCell, XlsRtdValue, XlsWorkbook};
use std::io::Cursor;

fn rtd_topic(segments: &[&str], value: XlsRtdValue, cells: Vec<XlsRtdCell>) -> XlsRealTimeData {
    XlsRealTimeData {
        common_prefix_len: 0,
        topic_segments: segments.iter().map(|segment| segment.to_string()).collect(),
        topic: segments.concat(),
        value,
        cells,
    }
}

#[test]
fn real_time_data_round_trips_through_writer_and_reader() {
    let topics = vec![
        rtd_topic(
            &["PROG.ID", "", "STOCK", "MSFT"],
            XlsRtdValue::Text("58.25".to_string()),
            vec![XlsRtdCell {
                row: 1,
                column: 2,
                sheet_index: 0,
            }],
        ),
        rtd_topic(
            &["PROG.ID", "", "BOND"],
            XlsRtdValue::Number(102.375),
            vec![
                XlsRtdCell {
                    row: 3,
                    column: 4,
                    sheet_index: 0,
                },
                XlsRtdCell {
                    row: 5,
                    column: 6,
                    sheet_index: 0,
                },
            ],
        ),
        rtd_topic(
            &["OTHER.SERVER", "remote", "FX"],
            XlsRtdValue::Boolean(true),
            Vec::new(),
        ),
        rtd_topic(
            &["OTHER.SERVER", "remote", "RATES"],
            XlsRtdValue::Integer(-7),
            Vec::new(),
        ),
        rtd_topic(
            &["OTHER.SERVER", "remote", "ERR"],
            XlsRtdValue::Error(0x2A),
            Vec::new(),
        ),
    ];

    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Quotes").unwrap();
    writer.write_number(sheet, 1, 2, 58.25).unwrap();
    for topic in &topics {
        writer.add_real_time_data(topic.clone());
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();

    let workbook = XlsWorkbook::new(Cursor::new(output.into_inner())).unwrap();
    assert_eq!(workbook.real_time_data(), topics.as_slice());
}

#[test]
fn real_time_data_prefix_compression_round_trips() {
    let first = rtd_topic(
        &["PROG.ID", "", "STOCK"],
        XlsRtdValue::Integer(1),
        Vec::new(),
    );
    // The second topic shares the "PROG.ID" prefix (7 characters) with the
    // first, so only the trailing sub-strings are stored.
    let mut second = rtd_topic(&["", "BOND"], XlsRtdValue::Integer(2), Vec::new());
    second.common_prefix_len = 7;
    second.topic = "PROG.IDBOND".to_string();

    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Quotes").unwrap();
    writer.write_number(sheet, 0, 0, 1.0).unwrap();
    writer.add_real_time_data(first.clone());
    writer.add_real_time_data(second.clone());
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();

    let workbook = XlsWorkbook::new(Cursor::new(output.into_inner())).unwrap();
    let parsed = workbook.real_time_data();
    assert_eq!(parsed.len(), 2);
    assert_eq!(parsed[0], first);
    assert_eq!(parsed[1].common_prefix_len, 7);
    assert_eq!(
        parsed[1].topic_segments,
        vec![String::new(), "BOND".to_string()]
    );
    assert_eq!(parsed[1].topic, "PROG.IDBOND");
    assert_eq!(parsed[1], second);
}

#[test]
fn workbook_without_real_time_data_has_none() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Plain").unwrap();
    writer.write_string(sheet, 0, 0, "content").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();

    let workbook = XlsWorkbook::new(Cursor::new(output.into_inner())).unwrap();
    assert!(workbook.real_time_data().is_empty());
}
