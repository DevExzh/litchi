#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "these tests construct deliberately small wire fixtures"
)]

use super::*;

fn push_u64(output: &mut Vec<u8>, value: usize) {
    output.extend_from_slice(&(value as u64).to_le_bytes());
}

fn push_bytes(output: &mut Vec<u8>, bytes: &[u8]) {
    push_u64(output, bytes.len());
    output.extend_from_slice(bytes);
}

fn closure(record_count: usize) -> Vec<u8> {
    let mut output = CLOSURE_HEADER.to_vec();
    push_u64(&mut output, record_count);
    output
}

fn record_header(output: &mut Vec<u8>, kind: u8, name: &str) {
    output.push(kind);
    push_bytes(output, name.as_bytes());
}

fn absent_member(output: &mut Vec<u8>) {
    output.push(0);
}

/// Append the private closure member layout without using the production
/// encoder.  This keeps these tests useful for malformed-wire admission: the
/// first byte is the member-presence flag, followed by the kind-specific
/// fields and, for relationship-bearing members, a strict boolean and its
/// raw relationship member.
fn member(
    output: &mut Vec<u8>,
    kind: u8,
    presence: u8,
    content_type: Option<&[u8]>,
    payload: &[u8],
    relationship: Option<(u8, &[u8])>,
) {
    output.push(presence);
    if kind == 2 {
        push_bytes(
            output,
            content_type.expect("part members carry a content type"),
        );
    }
    push_bytes(output, payload);
    if kind == 1 || kind == 2 {
        let (present, bytes) = relationship.expect("relationship-bearing members carry rels");
        output.push(present);
        push_bytes(output, bytes);
    }
}

fn single_record(
    kind: u8,
    name: &str,
    before: impl FnOnce(&mut Vec<u8>),
    after: impl FnOnce(&mut Vec<u8>),
) -> Vec<u8> {
    let mut output = closure(1);
    record_header(&mut output, kind, name);
    before(&mut output);
    after(&mut output);
    output
}

fn limits_with_small_parts() -> Limits {
    let mut limits = Limits::standard();
    limits.xml_bytes = 4;
    limits.image_bytes = 8;
    // Keep the enclosing wire fixture below its aggregate admission limit;
    // the tests below exercise the per-kind limit instead.
    limits.total_xml_bytes = 64;
    limits.total_image_bytes = 1024;
    limits
}

fn assert_invalid(result: Result<Vec<ClosureRecord>>) {
    let error = result.expect_err("malformed closure was accepted");
    assert!(
        matches!(error, Error::Invalid(_)),
        "expected Invalid, got {error:?}"
    );
}

fn assert_limit(result: Result<Vec<ClosureRecord>>, resource: &str, maximum: usize, actual: usize) {
    let error = result.expect_err("over-limit closure was accepted");
    match error {
        Error::Limit {
            resource: actual_resource,
            max: actual_maximum,
            actual: actual_value,
        } => {
            assert_eq!(actual_resource, resource);
            assert_eq!(actual_maximum, maximum);
            assert_eq!(actual_value, actual);
        },
        other => panic!("expected Limit for {resource:?}, got {other:?}"),
    }
}

#[test]
fn closure_rejects_non_boolean_member_presence() {
    let bytes = single_record(
        0,
        "[Content_Types].xml",
        |output| member(output, 0, 2, None, b"content-types", None),
        absent_member,
    );

    assert_invalid(decode_closure(&bytes, &Limits::standard()));
}

#[test]
fn closure_rejects_ascii_case_equivalent_part_records() {
    let mut bytes = closure(2);
    record_header(&mut bytes, 2, "/Web/Part.xml");
    member(
        &mut bytes,
        2,
        1,
        Some(b"application/xml"),
        b"one",
        Some((0, b"")),
    );
    absent_member(&mut bytes);
    record_header(&mut bytes, 2, "/web/part.xml");
    member(
        &mut bytes,
        2,
        1,
        Some(b"application/xml"),
        b"two",
        Some((0, b"")),
    );
    absent_member(&mut bytes);

    assert_invalid(decode_closure(&bytes, &Limits::standard()));
}

#[test]
fn closure_rejects_package_root_as_a_part_record() {
    let bytes = single_record(
        2,
        "/",
        |output| {
            member(
                output,
                2,
                1,
                Some(b"application/xml"),
                b"part",
                Some((0, b"")),
            );
        },
        absent_member,
    );

    assert_invalid(decode_closure(&bytes, &Limits::standard()));
}

#[test]
fn closure_enforces_per_kind_and_relationship_quotas() {
    let cases = [
        (0, "[Content_Types].xml", None, b"12345".as_slice(), None),
        (
            1,
            "/",
            None,
            b"12345".as_slice(),
            Some((0, b"12345".as_slice())),
        ),
        (
            2,
            "/part.xml",
            Some(b"application/xml".as_slice()),
            b"12345".as_slice(),
            Some((0, b"".as_slice())),
        ),
        (
            2,
            "/part.bin",
            Some(b"image/png".as_slice()),
            b"123456789".as_slice(),
            Some((0, b"".as_slice())),
        ),
    ];

    for (kind, name, content_type, payload, relationship) in cases {
        let bytes = single_record(
            kind,
            name,
            |output| member(output, kind, 1, content_type, payload, relationship),
            absent_member,
        );
        assert_limit(
            decode_closure(&bytes, &limits_with_small_parts()),
            "Web Extensions durable closure payload",
            if kind == 2 && content_type.is_some_and(|value| value == b"image/png") {
                8
            } else {
                4
            },
            if kind == 2 && content_type.is_some_and(|value| value == b"image/png") {
                9
            } else {
                5
            },
        );
    }
}

#[test]
fn closure_accepts_members_within_per_kind_quotas() {
    let mut bytes = closure(4);
    record_header(&mut bytes, 0, "[Content_Types].xml");
    member(&mut bytes, 0, 1, None, b"1234", None);
    absent_member(&mut bytes);
    record_header(&mut bytes, 1, "/");
    member(&mut bytes, 1, 1, None, b"1234", Some((0, b"1234")));
    absent_member(&mut bytes);
    record_header(&mut bytes, 2, "/part.xml");
    member(
        &mut bytes,
        2,
        1,
        Some(b"application/xml"),
        b"1234",
        Some((0, b"")),
    );
    absent_member(&mut bytes);
    record_header(&mut bytes, 2, "/part.bin");
    member(
        &mut bytes,
        2,
        1,
        Some(b"image/png"),
        b"1234",
        Some((0, b"")),
    );
    absent_member(&mut bytes);

    let records = decode_closure(&bytes, &limits_with_small_parts())
        .expect("members at their configured per-kind quotas should decode");
    assert_eq!(records.len(), 4);
}

#[test]
fn closure_enforces_aggregate_xml_payload_quota() {
    let mut bytes = closure(2);
    record_header(&mut bytes, 2, "/one.xml");
    member(
        &mut bytes,
        2,
        1,
        Some(b"application/xml"),
        b"1234",
        Some((0, b"")),
    );
    absent_member(&mut bytes);
    record_header(&mut bytes, 2, "/two.xml");
    member(
        &mut bytes,
        2,
        1,
        Some(b"application/xml"),
        b"5678",
        Some((0, b"")),
    );
    absent_member(&mut bytes);

    let mut limits = Limits::standard();
    limits.xml_bytes = 4;
    limits.total_xml_bytes = 6;
    // The outer closure itself is admitted; only its XML payload aggregate
    // should exceed the configured six-byte budget.
    limits.total_image_bytes = 512;

    assert_limit(
        decode_closure(&bytes, &limits),
        "Web Extensions durable XML payload bytes",
        6,
        8,
    );
}

#[test]
fn closure_enforces_aggregate_decoded_string_quota() {
    let bytes = single_record(
        2,
        "/part.xml",
        |output| {
            member(output, 2, 1, Some(b"application/xml"), b"x", Some((0, b"")));
        },
        absent_member,
    );

    let mut limits = Limits::standard();
    limits.total_string_bytes = 8;

    assert_limit(
        decode_closure(&bytes, &limits),
        "Web Extensions durable decoded strings",
        8,
        9,
    );
}
