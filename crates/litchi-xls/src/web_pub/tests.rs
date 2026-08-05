//! Regression coverage for the `WebPub` model and codec layers.

use super::*;

/// Build a compressed WebPubString.
fn web_pub_string(text: &str) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&(text.len() as u16).to_le_bytes());
    out.push(0u8); // fHighByte = 0
    out.extend_from_slice(text.as_bytes());
    out
}

struct WebPubBuilder {
    tws: u8,
    twd: u8,
    has_ref: bool,
    flags: u16,
    style_id: u32,
    source_name: Option<String>,
    file_destination: String,
    div_id: String,
    title: String,
    chart_shape_id: Option<u32>,
    reserved: Vec<u8>,
}

impl WebPubBuilder {
    fn new(tws: u8) -> Self {
        WebPubBuilder {
            tws,
            twd: 0,
            has_ref: tws == 0x04,
            flags: 0,
            style_id: 1,
            source_name: None,
            file_destination: String::new(),
            div_id: String::new(),
            title: String::new(),
            chart_shape_id: None,
            reserved: Vec::new(),
        }
    }

    fn build(self) -> Vec<u8> {
        let mut tail = Vec::new();
        if self.tws > 0x04 && self.tws != 0xFF {
            tail.extend_from_slice(&web_pub_string(self.source_name.as_deref().unwrap_or("")));
        }
        tail.extend_from_slice(&web_pub_string(&self.file_destination));
        tail.extend_from_slice(&web_pub_string(&self.div_id));
        tail.extend_from_slice(&web_pub_string(&self.title));
        if let Some(shape_id) = self.chart_shape_id {
            tail.extend_from_slice(&shape_id.to_le_bytes());
        }
        tail.extend_from_slice(&self.reserved);
        tail.extend_from_slice(&[0u8; 2]); // unused3

        let mut payload = Vec::new();
        payload.extend_from_slice(&WEB_PUB_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&(if self.has_ref { FRT_REF } else { 0 }).to_le_bytes());
        payload.extend_from_slice(&1u16.to_le_bytes()); // rwFirst
        payload.extend_from_slice(&9u16.to_le_bytes()); // rwLast
        payload.extend_from_slice(&2u16.to_le_bytes()); // colFirst
        payload.extend_from_slice(&5u16.to_le_bytes()); // colLast
        payload.push(self.tws);
        payload.push(self.twd);
        payload.extend_from_slice(&self.flags.to_le_bytes());
        payload.extend_from_slice(&[0u8; 4]); // reserved3 + unused2
        payload.extend_from_slice(&self.style_id.to_le_bytes());
        payload.extend_from_slice(&(tail.len() as u32).to_le_bytes());
        payload.extend_from_slice(&tail);
        payload
    }
}

#[test]
fn parses_workbook_publication() {
    let mut builder = WebPubBuilder::new(0x00);
    builder.twd = 0x01;
    builder.flags = AUTO_REPUBLISH | MHTML;
    builder.style_id = 0x1122_3344;
    builder.file_destination = "https://example.com/report.mht".to_string();
    builder.div_id = "bookmarks".to_string();
    builder.title = "Quarterly report".to_string();
    let pub_record = WebPub::parse(&builder.build()).expect("parse");
    assert_eq!(pub_record.source, WebSourceType::Workbook);
    assert_eq!(pub_record.page_type, WebPageType::WorkbookFunctionality);
    assert_eq!(pub_record.range, None);
    assert!(pub_record.auto_republish);
    assert!(pub_record.single_file);
    assert_eq!(pub_record.style_id, 0x1122_3344);
    assert_eq!(pub_record.source_name, None);
    assert_eq!(
        pub_record.file_destination,
        "https://example.com/report.mht"
    );
    assert_eq!(pub_record.div_id, "bookmarks");
    assert_eq!(pub_record.title, "Quarterly report");
    assert_eq!(pub_record.chart_shape_id, None);
    assert!(pub_record.reserved.is_empty());
}

#[test]
fn parses_range_publication_with_ref8() {
    let mut builder = WebPubBuilder::new(0x04);
    builder.file_destination = "C:\\pub\\range.htm".to_string();
    builder.title = "Range".to_string();
    let pub_record = WebPub::parse(&builder.build()).expect("parse");
    assert_eq!(pub_record.source, WebSourceType::Range);
    assert_eq!(
        pub_record.range,
        Some(WebPubRange::new(1, 9, 2, 5).unwrap())
    );
    assert_eq!(pub_record.source_name, None);
}

#[test]
fn parses_chart_publication_with_source_name_and_shape_id() {
    let mut builder = WebPubBuilder::new(0x05);
    builder.source_name = Some("Chart 1".to_string());
    builder.chart_shape_id = Some(0x0401);
    builder.file_destination = "chart.htm".to_string();
    builder.title = "Chart".to_string();
    builder.reserved = vec![0xAA, 0xBB, 0xCC];
    let pub_record = WebPub::parse(&builder.build()).expect("parse");
    assert_eq!(pub_record.source, WebSourceType::Chart);
    assert_eq!(pub_record.source_name.as_deref(), Some("Chart 1"));
    assert_eq!(pub_record.chart_shape_id, Some(0x0401));
    assert_eq!(pub_record.reserved, vec![0xAA, 0xBB, 0xCC]);
}

#[test]
fn parses_named_range_publication() {
    let mut builder = WebPubBuilder::new(0x08);
    builder.source_name = Some("Sales".to_string());
    builder.file_destination = "sales.htm".to_string();
    builder.title = "Sales".to_string();
    let pub_record = WebPub::parse(&builder.build()).expect("parse");
    assert_eq!(pub_record.source, WebSourceType::NamedRange);
    assert_eq!(pub_record.source_name.as_deref(), Some("Sales"));
    assert_eq!(pub_record.chart_shape_id, None);
}

#[test]
fn rejects_ref_flag_mismatch() {
    // tws = Range without fFrtRef.
    let mut builder = WebPubBuilder::new(0x04);
    builder.has_ref = false;
    builder.file_destination = "x.htm".to_string();
    builder.title = "x".to_string();
    assert!(WebPub::parse(&builder.build()).is_err());

    // tws = Workbook with fFrtRef set.
    let mut builder = WebPubBuilder::new(0x00);
    builder.has_ref = true;
    builder.file_destination = "x.htm".to_string();
    builder.title = "x".to_string();
    assert!(WebPub::parse(&builder.build()).is_err());

    let mut builder = WebPubBuilder::new(0x04);
    builder.file_destination = "x.htm".to_string();
    builder.title = "x".to_string();
    let mut payload = builder.build();
    payload[8..10].copy_from_slice(&256u16.to_le_bytes());
    assert!(WebPub::parse(&payload).is_err());
}

#[test]
fn rejects_cb_mismatch() {
    let mut builder = WebPubBuilder::new(0x00);
    builder.file_destination = "x.htm".to_string();
    builder.title = "x".to_string();
    let mut payload = builder.build();
    // Corrupt the cb field (offset 24).
    payload[24] = payload[24].wrapping_add(1);
    assert!(WebPub::parse(&payload).is_err());
}

#[test]
fn rejects_unknown_type_codes_and_rt_mismatch() {
    let mut builder = WebPubBuilder::new(0x09); // not in the tws table
    builder.file_destination = "x.htm".to_string();
    builder.title = "x".to_string();
    assert!(WebPub::parse(&builder.build()).is_err());

    let mut builder = WebPubBuilder::new(0x00);
    builder.twd = 0x04; // not in the twd table
    builder.file_destination = "x.htm".to_string();
    builder.title = "x".to_string();
    assert!(WebPub::parse(&builder.build()).is_err());

    let mut builder = WebPubBuilder::new(0x00);
    builder.file_destination = "x.htm".to_string();
    builder.title = "x".to_string();
    let mut payload = builder.build();
    payload[0] = 0x02; // corrupt rt to 0x0902... -> mismatch
    payload[1] = 0x09;
    assert!(WebPub::parse(&payload).is_err());
}

#[test]
fn rejects_truncated_and_oversized_strings() {
    assert!(WebPub::parse(&[]).is_err());
    assert!(WebPub::parse(&[0u8; 20]).is_err());

    // A WebPubString longer than 255 characters is illegal.
    let mut builder = WebPubBuilder::new(0x00);
    builder.file_destination = "x".repeat(300);
    builder.title = "x".to_string();
    assert!(WebPub::parse(&builder.build()).is_err());
}

#[test]
fn payload_round_trips() {
    let values = [
        WebPub {
            source: WebSourceType::Workbook,
            page_type: WebPageType::WorkbookFunctionality,
            range: None,
            auto_republish: true,
            single_file: true,
            style_id: 0x1122_3344,
            source_name: None,
            file_destination: "https://example.com/report.mht".to_string(),
            div_id: "top".to_string(),
            title: "Quarterly report".to_string(),
            chart_shape_id: None,
            reserved: Vec::new(),
        },
        WebPub {
            source: WebSourceType::Range,
            page_type: WebPageType::ViewOnly,
            range: Some(WebPubRange::new(1, 9, 2, 5).unwrap()),
            auto_republish: false,
            single_file: false,
            style_id: 7,
            source_name: None,
            file_destination: "C:\\pub\\range.htm".to_string(),
            div_id: String::new(),
            title: "Range €".to_string(),
            chart_shape_id: None,
            reserved: vec![0xAA, 0xBB],
        },
        WebPub {
            source: WebSourceType::Chart,
            page_type: WebPageType::ChartFunctionality,
            range: None,
            auto_republish: true,
            single_file: false,
            style_id: 9,
            source_name: Some("Chart 1".to_string()),
            file_destination: "chart.htm".to_string(),
            div_id: "c1".to_string(),
            title: "Chart".to_string(),
            chart_shape_id: Some(0x0401),
            reserved: Vec::new(),
        },
    ];
    for value in values {
        let payload = value.to_payload().expect("serialize");
        let parsed = WebPub::parse(&payload).expect("re-parse");
        assert_eq!(parsed, value);
    }
}

#[test]
fn serialize_rejects_inconsistent_conditional_fields() {
    let mut value = WebPub {
        source: WebSourceType::Workbook,
        page_type: WebPageType::ViewOnly,
        range: Some(WebPubRange::new(0, 1, 0, 1).unwrap()),
        auto_republish: false,
        single_file: false,
        style_id: 1,
        source_name: None,
        file_destination: "x.htm".to_string(),
        div_id: String::new(),
        title: "x".to_string(),
        chart_shape_id: None,
        reserved: Vec::new(),
    };
    assert!(value.to_payload().is_err());
    value.range = None;
    value.source_name = Some("unexpected".to_string());
    assert!(value.to_payload().is_err());
    value.source_name = None;
    value.chart_shape_id = Some(1);
    assert!(value.to_payload().is_err());

    // A WebPubString longer than 255 characters cannot be written.
    value.chart_shape_id = None;
    value.title = "x".repeat(300);
    assert!(value.to_payload().is_err());
}
