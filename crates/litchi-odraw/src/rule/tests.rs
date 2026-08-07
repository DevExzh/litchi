#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::{Kind, Rule, from_record, parse};
use crate::{Error, Record, RecordKind};

fn record(version: u8, instance: u16, kind: u16, body: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(8 + body.len());
    let ver_inst = (u16::from(version) & 0x000F) | ((instance & 0x0FFF) << 4);
    bytes.extend_from_slice(&ver_inst.to_le_bytes());
    bytes.extend_from_slice(&kind.to_le_bytes());
    bytes.extend_from_slice(&(body.len() as u32).to_le_bytes());
    bytes.extend_from_slice(body);
    bytes
}

fn fields(values: &[u32]) -> Vec<u8> {
    let mut body = Vec::with_capacity(values.len() * 4);
    for value in values {
        body.extend_from_slice(&value.to_le_bytes());
    }
    body
}

fn encode(rule: &Rule<'_>) -> Vec<u8> {
    let mut encoded = Vec::new();
    rule.write_to(&mut encoded)
        .expect("Vec writing cannot fail");
    encoded
}

#[test]
fn decodes_and_reencodes_connector_rules() {
    let bytes = record(
        1,
        0,
        RecordKind::ConnectorRule.raw(),
        &fields(&[7, 11, 13, 17, 19, 23]),
    );

    let rule = parse(&bytes).expect("valid connector rule");
    let connector = rule.as_connector().expect("connector variant");
    assert_eq!(connector.rule_id(), 7);
    assert_eq!(connector.start_shape_id(), 11);
    assert_eq!(connector.end_shape_id(), 13);
    assert_eq!(connector.connector_shape_id(), 17);
    assert_eq!(connector.start_connection_site(), 19);
    assert_eq!(connector.end_connection_site(), 23);
    assert_eq!(rule.kind(), Kind::Connector);
    assert_eq!(encode(&rule), bytes);
}

#[test]
fn decodes_arc_and_callout_rules() {
    for (kind, expected) in [
        (RecordKind::ArcRule.raw(), Kind::Arc),
        (RecordKind::CalloutRule.raw(), Kind::Callout),
    ] {
        let bytes = record(0, 0, kind, &fields(&[31, 37]));
        let rule = parse(&bytes).expect("valid two-field rule");

        assert_eq!(rule.kind(), expected);
        assert_eq!(rule.raw_kind(), kind);
        assert_eq!(rule.version(), 0);
        assert_eq!(rule.instance(), 0);
        assert_eq!(
            rule.as_arc()
                .map(|value| (value.rule_id(), value.shape_id())),
            (expected == Kind::Arc).then_some((31, 37))
        );
        assert_eq!(
            rule.as_callout()
                .map(|value| (value.rule_id(), value.shape_id())),
            (expected == Kind::Callout).then_some((31, 37))
        );
        assert_eq!(encode(&rule), bytes);
    }
}

#[test]
fn preserves_unknown_rule_bytes_and_header() {
    let bytes = record(3, 0x0ABC, 0xF1FE, &[0x00, 0xFF, 0x11, 0x22, 0x33]);
    let rule = parse(&bytes).expect("unknown extension record");

    assert_eq!(rule.kind(), Kind::Opaque(0xF1FE));
    let opaque = rule.as_opaque().expect("opaque variant");
    assert_eq!(opaque.raw_kind(), 0xF1FE);
    assert_eq!(opaque.version(), 3);
    assert_eq!(opaque.instance(), 0x0ABC);
    assert_eq!(opaque.data(), &[0x00, 0xFF, 0x11, 0x22, 0x33]);
    assert_eq!(encode(&rule), bytes);
}

#[test]
fn parses_a_preparsed_record_without_copying() {
    let bytes = record(0, 0, RecordKind::ArcRule.raw(), &fields(&[41, 43]));
    let (record, consumed) = Record::parse(&bytes, 0).expect("valid record");
    assert_eq!(consumed, bytes.len());

    let rule = from_record(record).expect("valid arc rule");
    assert_eq!(rule.as_arc().map(|value| value.shape_id()), Some(43));
}

#[test]
fn rejects_malformed_headers_and_non_rules() {
    let connector_body = fields(&[1, 2, 3, 4, 5, 6]);
    assert!(matches!(
        parse(&record(
            0,
            0,
            RecordKind::ConnectorRule.raw(),
            &connector_body
        )),
        Err(Error::MalformedShape {
            reason: "connector rule version is invalid"
        })
    ));

    let wrong_length = record(0, 0, RecordKind::ArcRule.raw(), &fields(&[1]));
    assert!(matches!(
        parse(&wrong_length),
        Err(Error::MalformedShape {
            reason: "arc rule length is invalid"
        })
    ));

    let shape = record(2, 0, RecordKind::Sp.raw(), &fields(&[1, 2]));
    assert!(matches!(
        parse(&shape),
        Err(Error::MalformedShape {
            reason: "record is not an OfficeArt solver rule"
        })
    ));
}

#[test]
fn rejects_truncation_and_top_level_trailing_data() {
    assert!(matches!(
        parse(&[0, 0, 0x12]),
        Err(Error::TruncatedHeader { .. })
    ));

    let mut bytes = record(0, 0, RecordKind::ArcRule.raw(), &fields(&[1, 2]));
    bytes.push(0xAA);
    assert!(matches!(
        parse(&bytes),
        Err(Error::TrailingData { offset }) if offset == 16
    ));
}
