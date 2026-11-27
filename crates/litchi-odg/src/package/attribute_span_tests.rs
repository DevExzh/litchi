//! Differential coverage for the fixed shape-attribute source-span batch.
//!
//! These tests live below `snapshot` so they can compare the private batched
//! helper with the original ordered helper without widening the production
//! visibility of either function.

use super::*;
use quick_xml::{events::Event, reader::NsReader};

fn ordered_shape_attribute_source_spans(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    tag: &[u8],
    tag_start: usize,
) -> Result<ShapeAttributeSpans> {
    Ok([
        attribute_source_span(reader, element, tag, tag_start, DRAW, b"control")?,
        shape_name_span(reader, element, tag, tag_start)?,
        attribute_source_span(reader, element, tag, tag_start, DRAW, b"layer")?,
        attribute_source_span(reader, element, tag, tag_start, SVG, b"x")?,
        attribute_source_span(reader, element, tag, tag_start, SVG, b"y")?,
        attribute_source_span(reader, element, tag, tag_start, SVG, b"width")?,
        attribute_source_span(reader, element, tag, tag_start, SVG, b"height")?,
        attribute_source_span(reader, element, tag, tag_start, SVG, b"x1")?,
        attribute_source_span(reader, element, tag, tag_start, SVG, b"y1")?,
        attribute_source_span(reader, element, tag, tag_start, SVG, b"x2")?,
        attribute_source_span(reader, element, tag, tag_start, SVG, b"y2")?,
        attribute_source_span(reader, element, tag, tag_start, SVG, b"viewBox")?,
        attribute_source_span(reader, element, tag, tag_start, DRAW, b"points")?,
        attribute_source_span(reader, element, tag, tag_start, DRAW, b"transform")?,
        attribute_source_span(reader, element, tag, tag_start, SVG, b"d")?,
        attribute_source_span(reader, element, tag, tag_start, DRAW, b"style-name")?,
    ])
}

fn compare_first_shape(
    xml: &str,
    tag_start: usize,
) -> (
    Vec<u8>,
    Result<ShapeAttributeSpans>,
    Result<ShapeAttributeSpans>,
    Result<ShapeAttributeSpans>,
) {
    compare_first_shape_with_tag(xml, None, tag_start)
}

fn compare_first_shape_with_tag(
    xml: &str,
    replacement_tag: Option<&[u8]>,
    tag_start: usize,
) -> (
    Vec<u8>,
    Result<ShapeAttributeSpans>,
    Result<ShapeAttributeSpans>,
    Result<ShapeAttributeSpans>,
) {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let (start, element, end) = loop {
        let start = usize::try_from(reader.buffer_position())
            .expect("test XML position should fit in usize");
        let (_, event) = reader
            .read_resolved_event()
            .expect("test XML should produce a readable event");
        match event {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"rect" =>
            {
                let end = usize::try_from(reader.buffer_position())
                    .expect("test XML position should fit in usize");
                break (start, element.into_owned(), end);
            },
            Event::Eof => panic!("test XML did not contain a draw:rect"),
            _ => {},
        }
    };
    let source_tag = xml
        .as_bytes()
        .get(start..end)
        .expect("event span should be within the source");
    let tag = replacement_tag.unwrap_or(source_tag).to_vec();
    let batch = shape_attribute_source_spans_batch(&reader, &element, &tag, tag_start);
    let wrapped = shape_attribute_source_spans(&reader, &element, &tag, tag_start);
    let ordered = ordered_shape_attribute_source_spans(&reader, &element, &tag, tag_start);
    (tag, batch, wrapped, ordered)
}

fn span_bytes<'a>(tag: &'a [u8], span: &Range<usize>, tag_start: usize) -> &'a [u8] {
    let start = span
        .start
        .checked_sub(tag_start)
        .expect("span should include the supplied tag start");
    let end = span
        .end
        .checked_sub(tag_start)
        .expect("span should include the supplied tag start");
    &tag[start..end]
}

fn assert_same_result(left: &Result<ShapeAttributeSpans>, right: &Result<ShapeAttributeSpans>) {
    match (left, right) {
        (Ok(left), Ok(right)) => assert_eq!(left, right),
        (Err(left), Err(right)) => assert_eq!(left.to_string(), right.to_string()),
        (left, right) => panic!("result shape differs: {left:?} vs {right:?}"),
    }
}

const DRAW_XMLNS: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const SVG_XMLNS: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";

#[test]
fn batch_matches_ordered_spans_with_aliases_and_spacing() {
    let xml = format!(
        r#"<d:rect xmlns:d="{DRAW_XMLNS}" xmlns:s="{SVG_XMLNS}" xmlns:f="urn:example:foreign"
            f:name="foreign-name" d:control = 'control' s:x="x" d:name = "name"
            s:y = 'y' d:layer="layer" s:width = "width" s:height='height'
            s:x1 = 'x1' s:y1="y1" s:x2 = "x2" s:y2='y2'
            s:viewBox = "view box" f:viewBox="foreign-view" d:points='points'
            d:transform = "transform" s:d='path' d:style-name = 'style'/>"#
    );
    let (tag, batch, wrapped, ordered) = compare_first_shape(&xml, 0);
    assert_same_result(&batch, &ordered);
    assert_same_result(&wrapped, &ordered);
    assert_same_result(&batch, &wrapped);
    let spans = batch.expect("valid shape attributes should produce spans");
    for (span, expected) in spans.iter().zip([
        "control",
        "name",
        "layer",
        "x",
        "y",
        "width",
        "height",
        "x1",
        "y1",
        "x2",
        "y2",
        "view box",
        "points",
        "transform",
        "path",
        "style",
    ]) {
        let span = span.as_ref().expect("each requested field is present");
        assert_eq!(span_bytes(&tag, span, 0), expected.as_bytes());
    }
}

#[test]
fn batch_returns_none_for_unmatched_and_unbound_same_local_names() {
    let xml = format!(
        r#"<d:rect xmlns:d="{DRAW_XMLNS}" xmlns:f="urn:example:foreign"
            name="unbound" f:name="foreign-name" f:x="foreign-x" f:style-name="foreign-style"
            f:opaque="preserve"/>"#
    );
    let (_tag, batch, wrapped, ordered) = compare_first_shape(&xml, 0);
    assert_same_result(&batch, &ordered);
    assert_same_result(&wrapped, &ordered);
    let spans = batch.expect("unmatched fields should be accepted");
    assert!(spans.iter().all(Option::is_none));
}

#[test]
fn batch_follows_namespace_rebinding_on_nested_shape() {
    let xml = format!(
        r#"<f:outer xmlns:f="urn:example:foreign"><d:scope xmlns:d="urn:example:foreign">
            <d:rect xmlns:d="{DRAW_XMLNS}" d:name="rebound" d:layer="layer"/>
        </d:scope></f:outer>"#
    );
    let (tag, batch, wrapped, ordered) = compare_first_shape(&xml, 37);
    assert_same_result(&batch, &ordered);
    assert_same_result(&wrapped, &ordered);
    let spans = batch.expect("rebound drawing attributes should resolve");
    assert_eq!(
        span_bytes(&tag, &spans[1].clone().expect("name span"), 37),
        b"rebound"
    );
    assert_eq!(
        spans[2].as_ref().map(|span| span_bytes(&tag, span, 37)),
        Some(b"layer".as_slice())
    );
}

#[test]
fn batch_and_ordered_helper_preserve_malformed_attribute_error_display() {
    let xml = format!(r#"<d:rect xmlns:d="{DRAW_XMLNS}" d:name="name" malformed/>"#);
    let (_tag, batch, wrapped, ordered) = compare_first_shape(&xml, 0);
    assert!(batch.is_err(), "batch must retain raw attribute validation");
    assert_same_result(&wrapped, &ordered);
    assert_same_result(&batch, &wrapped);
}

#[test]
fn batch_falls_back_for_raw_duplicate_attribute_error() {
    let xml = format!(
        r#"<d:rect xmlns:d="{DRAW_XMLNS}" xmlns:f="urn:example:foreign"
            d:name="name"
            f:opaque="one" f:opaque="two"/>"#
    );
    let (_tag, batch, wrapped, ordered) = compare_first_shape(&xml, 0);
    assert!(batch.is_err(), "duplicate aliases must be rejected");
    assert!(
        ordered.is_err(),
        "the ordered helper must reject duplicates"
    );
    assert_same_result(&wrapped, &ordered);
    assert_same_result(&batch, &wrapped);
}

#[test]
fn batch_falls_back_for_reverse_order_semantic_duplicate_fields() {
    let xml = format!(
        r#"<d:rect xmlns:d="{DRAW_XMLNS}" xmlns:a="{DRAW_XMLNS}"
            d:layer="first-layer" a:layer="second-layer"
            d:name="first-name" a:name="second-name"/>"#
    );
    let (_tag, batch, wrapped, ordered) = compare_first_shape(&xml, 0);
    assert!(batch.is_err(), "duplicate aliases must be rejected");
    assert!(
        ordered.is_err(),
        "the ordered helper must reject duplicate aliases"
    );
    assert_same_result(&wrapped, &ordered);
    assert_same_result(&batch, &wrapped);
}

#[test]
fn batch_error_is_replayed_in_order_when_raw_tag_disagrees_with_event() {
    let xml = format!(
        r#"<d:rect xmlns:d="{DRAW_XMLNS}" xmlns:a="{DRAW_XMLNS}"
            d:control="control" d:name="name" d:layer="first-layer" a:layer="second-layer"/>"#
    );
    let replacement_tag = format!(r#"<d:rect xmlns:d="{DRAW_XMLNS}" d:name="name"/>"#);
    let (_tag, batch, wrapped, ordered) =
        compare_first_shape_with_tag(&xml, Some(replacement_tag.as_bytes()), 0);
    assert!(batch.is_err(), "the batch sees the duplicate later field");
    assert!(
        ordered.is_err(),
        "the ordered helper sees the missing first field"
    );
    assert_ne!(
        batch.as_ref().unwrap_err().to_string(),
        ordered.as_ref().unwrap_err().to_string(),
        "the mismatched raw tag must exercise distinct direct error paths"
    );
    assert_same_result(&wrapped, &ordered);
}

#[test]
fn batch_matches_ordered_helper_for_large_unrelated_inventory_and_start_tag() {
    let mut xml = format!(
        r#"<d:rect xmlns:d="{DRAW_XMLNS}" xmlns:s="{SVG_XMLNS}" xmlns:f="urn:example:foreign" "#
    );
    for index in 0..128 {
        write!(xml, "f:opaque{index}='unrelated-{index}' ").expect("String writes cannot fail");
    }
    xml.push_str(r#"d:name='large-inventory'><ignored/></d:rect>"#);
    let (tag, batch, wrapped, ordered) = compare_first_shape(&xml, 91);
    assert_same_result(&batch, &ordered);
    assert_same_result(&wrapped, &ordered);
    let spans = batch.expect("large unrelated inventories remain valid");
    assert_eq!(
        span_bytes(&tag, &spans[1].clone().expect("name span"), 91),
        b"large-inventory"
    );
}
