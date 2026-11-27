//! Differential coverage for the fixed-size ODG attribute-value batch.
//!
//! These tests stay below `snapshot` so they can compare the private batch
//! helper with the original ordered `attribute` calls.  The production helper
//! is expected to replay those calls when its fast path encounters a raw,
//! decoding, or expanded-name duplicate error.

use super::*;
use quick_xml::{
    events::{BytesStart, Event},
    reader::NsReader,
};

const DRAW_XMLNS: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const SVG_XMLNS: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";

fn first_empty(xml: &str) -> (NsReader<&[u8]>, BytesStart<'static>) {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let event = reader
        .read_resolved_event()
        .expect("test XML should produce one readable event")
        .1;
    let element = match event {
        Event::Empty(element) => element.into_owned(),
        other => panic!("expected an empty element, got {other:?}"),
    };
    (reader, element)
}

fn named_empty<'a>(xml: &'a str, local: &[u8]) -> (NsReader<&'a [u8]>, BytesStart<'static>) {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    loop {
        let event = reader
            .read_resolved_event()
            .expect("test XML should produce readable events")
            .1;
        match event {
            Event::Empty(element) if element.local_name().as_ref() == local => {
                return (reader, element.into_owned());
            },
            Event::Eof => panic!("test XML did not contain the requested empty element"),
            _ => {},
        }
    }
}

fn ordered_attribute_values<const N: usize>(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    requests: [(&[u8], &[u8]); N],
) -> Result<[Option<String>; N]> {
    let mut values: [Option<String>; N] = std::array::from_fn(|_| None);
    for (slot, (expected, local)) in requests.into_iter().enumerate() {
        values[slot] = attribute(reader, element, expected, local)?;
    }
    Ok(values)
}

fn assert_same_result<const N: usize>(
    left: &Result<[Option<String>; N]>,
    right: &Result<[Option<String>; N]>,
) {
    match (left, right) {
        (Ok(left), Ok(right)) => assert_eq!(left, right),
        (Err(left), Err(right)) => assert_eq!(left.to_string(), right.to_string()),
        (left, right) => panic!("attribute batch result differs: {left:?} vs {right:?}"),
    }
}

#[test]
fn value_batch_matches_ordered_values_and_normalizes_references() {
    let xml = format!(
        r#"<d:rect xmlns:d="{DRAW_XMLNS}" xmlns:s="{SVG_XMLNS}"
            d:name="first&amp;second&#x21;" s:x="1&amp;2" d:missing="unused"/>"#
    );
    let (reader, element) = first_empty(&xml);
    let requests = [(DRAW, b"name".as_slice()), (SVG, b"x".as_slice())];
    let batch = attribute_values(&reader, &element, requests);
    let ordered = ordered_attribute_values(&reader, &element, requests);
    assert_same_result(&batch, &ordered);
    assert_eq!(
        batch.expect("valid values should decode"),
        [Some("first&second!".to_owned()), Some("1&2".to_owned())]
    );
}

#[test]
fn value_batch_preserves_aliases_and_ignores_unbound_default_and_foreign_names() {
    let xml = format!(
        r#"<d:rect xmlns="{DRAW_XMLNS}" xmlns:d="{DRAW_XMLNS}"
            xmlns:a="{DRAW_XMLNS}" xmlns:s="{SVG_XMLNS}" xmlns:f="urn:example:foreign"
            name="unbound" f:name="foreign-name" a:name="alias-name"
            s:x="svg-x" f:x="&bad-foreign;" u:name="&bad-unknown;"
            f:opaque="opaque"/>"#
    );
    let (reader, element) = first_empty(&xml);
    let requests = [
        (DRAW, b"name".as_slice()),
        (SVG, b"x".as_slice()),
        (DRAW, b"layer".as_slice()),
    ];
    let batch = attribute_values(&reader, &element, requests);
    let ordered = ordered_attribute_values(&reader, &element, requests);
    assert_same_result(&batch, &ordered);
    assert_eq!(
        batch.expect("foreign and unbound fields should be ignored"),
        [
            Some("alias-name".to_owned()),
            Some("svg-x".to_owned()),
            None
        ]
    );
}

#[test]
fn value_batch_follows_namespace_rebinding_on_nested_shape() {
    let xml = format!(
        r#"<f:outer xmlns:f="urn:example:foreign" xmlns:d="urn:example:foreign">
            <d:scope xmlns:d="urn:example:foreign:other">
                <d:rect xmlns:d="{DRAW_XMLNS}" d:name="rebound" d:layer="layer"/>
            </d:scope>
        </f:outer>"#
    );
    let (reader, element) = named_empty(&xml, b"rect");
    let requests = [(DRAW, b"name".as_slice()), (DRAW, b"layer".as_slice())];
    let batch = attribute_values(&reader, &element, requests);
    let ordered = ordered_attribute_values(&reader, &element, requests);
    assert_same_result(&batch, &ordered);
    assert_eq!(
        batch.expect("the innermost drawing binding should win"),
        [Some("rebound".to_owned()), Some("layer".to_owned())]
    );
}

#[test]
fn value_batch_supports_duplicate_query_slots_independently() {
    let xml = format!(
        r#"<d:rect xmlns:d="{DRAW_XMLNS}" xmlns:a="{DRAW_XMLNS}"
            a:name="one" d:layer="two"/>"#
    );
    let (reader, element) = first_empty(&xml);
    let requests = [
        (DRAW, b"name".as_slice()),
        (DRAW, b"name".as_slice()),
        (DRAW, b"layer".as_slice()),
    ];
    let batch = attribute_values(&reader, &element, requests);
    let ordered = ordered_attribute_values(&reader, &element, requests);
    assert_same_result(&batch, &ordered);
    assert_eq!(
        batch.expect("one source field can satisfy two query slots"),
        [
            Some("one".to_owned()),
            Some("one".to_owned()),
            Some("two".to_owned())
        ]
    );
}

#[test]
fn value_batch_handles_zero_query_slots() {
    let xml = format!(r#"<d:rect xmlns:d="{DRAW_XMLNS}" malformed/>"#);
    let (reader, element) = first_empty(&xml);
    let requests: [(&[u8], &[u8]); 0] = [];
    let batch = attribute_values(&reader, &element, requests);
    let ordered = ordered_attribute_values(&reader, &element, requests);
    assert_same_result(&batch, &ordered);
    assert_eq!(batch.expect("an empty query is valid"), []);
}

#[test]
fn value_batch_replays_ordered_error_for_reversed_invalid_values() {
    let xml = format!(
        r#"<d:rect xmlns:d="{DRAW_XMLNS}" xmlns:s="{SVG_XMLNS}"
            s:y="&bad-y;" s:x="&bad-x;"/>"#
    );
    let (reader, element) = first_empty(&xml);
    let requests = [(SVG, b"x".as_slice()), (SVG, b"y".as_slice())];
    let batch = attribute_values(&reader, &element, requests);
    let ordered = ordered_attribute_values(&reader, &element, requests);
    assert!(
        batch.is_err(),
        "an invalid referenced value must be rejected"
    );
    assert_same_result(&batch, &ordered);
}

#[test]
fn value_batch_decodes_before_reporting_expanded_name_duplicate() {
    let xml = format!(
        r#"<d:rect xmlns:d="{DRAW_XMLNS}" xmlns:a="{DRAW_XMLNS}"
            d:name="first" a:name="&bad-second;"/>"#
    );
    let (reader, element) = first_empty(&xml);
    let requests = [(DRAW, b"name".as_slice())];
    let batch = attribute_values(&reader, &element, requests);
    let ordered = ordered_attribute_values(&reader, &element, requests);
    assert!(
        batch.is_err(),
        "the invalid second alias value must be rejected"
    );
    assert_same_result(&batch, &ordered);
}

#[test]
fn value_batch_replays_raw_syntax_and_duplicate_errors() {
    for xml in [
        format!(r#"<d:rect xmlns:d="{DRAW_XMLNS}" d:name="name" malformed/>"#),
        format!(r#"<d:rect xmlns:d="{DRAW_XMLNS}" d:name="one" d:name="two"/>"#),
    ] {
        let (reader, element) = first_empty(&xml);
        let requests = [(DRAW, b"name".as_slice())];
        let batch = attribute_values(&reader, &element, requests);
        let ordered = ordered_attribute_values(&reader, &element, requests);
        assert!(batch.is_err(), "malformed raw attributes must be rejected");
        assert_same_result(&batch, &ordered);
    }
}

#[test]
fn value_batch_matches_ordered_values_with_large_irrelevant_inventory() {
    let mut xml = format!(
        r#"<d:rect xmlns:d="{DRAW_XMLNS}" xmlns:s="{SVG_XMLNS}" xmlns:f="urn:example:foreign" "#
    );
    for index in 0..256 {
        xml.push_str(&format!("f:opaque{index}=\"unrelated-{index}\" "));
    }
    xml.push_str(r#"d:name="retained" s:x="2cm"/>"#);
    let (reader, element) = first_empty(&xml);
    let requests = [
        (DRAW, b"name".as_slice()),
        (SVG, b"x".as_slice()),
        (DRAW, b"style-name".as_slice()),
    ];
    let batch = attribute_values(&reader, &element, requests);
    let ordered = ordered_attribute_values(&reader, &element, requests);
    assert_same_result(&batch, &ordered);
    assert_eq!(
        batch.expect("irrelevant attributes should not affect selected values"),
        [Some("retained".to_owned()), Some("2cm".to_owned()), None]
    );
}
