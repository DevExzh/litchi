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

#[test]
fn snapshot_noop_and_edits_preserve_opaque_wire_bytes() {
    let mut builder = WebPubBuilder::new(0x00);
    builder.flags = 0x4000;
    builder.file_destination = "https://example.invalid/publish.mht".to_string();
    builder.div_id = "top".to_string();
    builder.title = "Original".to_string();
    builder.reserved = vec![0xAA, 0xBB, 0xCC];
    let mut bytes = builder.build();
    bytes[16..20].copy_from_slice(&[0x11, 0x22, 0x33, 0x44]);
    let trailing = bytes.len() - 2;
    bytes[trailing..].copy_from_slice(&[0xD1, 0xD2]);

    let snapshot = Snapshot::parse(bytes.clone()).expect("snapshot");
    assert_eq!(snapshot.finish(), bytes);

    let noop = snapshot.edit().commit().expect("no-op commit");
    assert!(!noop.changed());
    assert!(noop.patch().is_noop());
    assert_eq!(noop.into_bytes(), bytes);

    let mut transaction = snapshot.edit();
    transaction
        .set_title("A changed publication title")
        .expect("title edit")
        .set_file_destination("https://example.invalid/café.mht")
        .expect("inert URL edit")
        .set_auto_republish(true)
        .expect("flag edit")
        .set_style_id(0x1122_3344)
        .expect("style edit");
    let commit = transaction.commit().expect("commit");
    let after = commit.snapshot().bytes();
    assert_eq!(&after[16..20], &[0x11, 0x22, 0x33, 0x44]);
    assert_eq!(&after[after.len() - 2..], &[0xD1, 0xD2]);
    assert!(after.windows(3).any(|window| window == [0xAA, 0xBB, 0xCC]));
    assert_eq!(
        commit.snapshot().publication().title,
        "A changed publication title"
    );
    assert_eq!(
        commit.snapshot().publication().file_destination,
        "https://example.invalid/café.mht"
    );
    assert_eq!(commit.snapshot().publication().style_id, 0x1122_3344);
    assert!(commit.snapshot().publication().auto_republish);
    assert_eq!(u16::from_le_bytes([after[14], after[15]]) & 0x4000, 0x4000);

    let applied = commit.patch().apply(&snapshot).expect("apply");
    let reverted = commit.patch().inverse().apply(&applied).expect("inverse");
    assert_eq!(reverted.finish(), bytes);
    let reverted_directly = commit
        .patch()
        .revert(commit.snapshot())
        .expect("direct revert");
    assert_eq!(reverted_directly.finish(), bytes);
}

#[test]
fn bounded_conditional_edits_and_invalid_edits_are_atomic() {
    let mut workbook = WebPubBuilder::new(0x00);
    workbook.file_destination = "book.htm".to_string();
    workbook.title = "Book".to_string();
    let snapshot = Snapshot::parse(workbook.build()).expect("snapshot");
    let mut transaction = snapshot.edit();
    let before = transaction.snapshot().expect("candidate").finish();

    assert!(transaction.set_title("x".repeat(256)).is_err());
    assert!(transaction.set_source_name("unexpected").is_err());
    assert!(
        transaction
            .set_range(WebPubRange::new(0, 1, 0, 1).unwrap())
            .is_err()
    );
    assert!(transaction.set_chart_shape_id(7).is_err());
    assert_eq!(
        transaction
            .snapshot()
            .expect("unchanged candidate")
            .finish(),
        before
    );

    let mut range = WebPubBuilder::new(0x04);
    range.file_destination = "range.htm".to_string();
    range.title = "Range".to_string();
    let range_snapshot = Snapshot::parse(range.build()).expect("range snapshot");
    let mut range_edit = range_snapshot.edit();
    range_edit
        .set_range(WebPubRange::new(10, 20, 3, 8).unwrap())
        .expect("range edit");
    assert_eq!(
        range_edit
            .commit()
            .expect("range commit")
            .snapshot()
            .publication()
            .range,
        Some(WebPubRange::new(10, 20, 3, 8).unwrap())
    );

    let mut chart = WebPubBuilder::new(0x05);
    chart.source_name = Some("Chart 1".to_string());
    chart.chart_shape_id = Some(4);
    chart.file_destination = "chart.htm".to_string();
    chart.title = "Chart".to_string();
    let chart_snapshot = Snapshot::parse(chart.build()).expect("chart snapshot");
    let mut chart_edit = chart_snapshot.edit();
    chart_edit
        .set_source_name("Chart 2")
        .expect("source name edit")
        .set_chart_shape_id(8)
        .expect("shape edit");
    let chart_commit = chart_edit.commit().expect("chart commit");
    assert_eq!(
        chart_commit.snapshot().publication().source_name.as_deref(),
        Some("Chart 2")
    );
    assert_eq!(
        chart_commit.snapshot().publication().chart_shape_id,
        Some(8)
    );
}

#[test]
fn patch_rejects_stale_source_without_mutating_it() {
    let mut builder = WebPubBuilder::new(0x00);
    builder.file_destination = "source.htm".to_string();
    builder.title = "Source".to_string();
    let snapshot = Snapshot::parse(builder.build()).expect("snapshot");
    let mut transaction = snapshot.edit();
    transaction.set_title("Changed").expect("edit");
    let commit = transaction.commit().expect("commit");

    let mut stale_bytes = snapshot.finish();
    stale_bytes[16] ^= 1;
    let stale = Snapshot::parse(stale_bytes).expect("stale snapshot");
    let stale_before = stale.finish();
    assert!(commit.patch().apply(&stale).is_err());
    assert_eq!(stale.finish(), stale_before);
}
