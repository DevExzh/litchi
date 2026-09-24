#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "these tests construct deliberately small durable intent fixtures"
)]

use std::ops::Range;
use std::sync::Arc;

use super::super::super::model::{ExtKind, Reference, Store};
use super::*;

fn pane(id: &str, relationship_id: &str, image: Option<(&str, &[u8])>) -> Pane {
    let reference = Reference::new(format!("{id}-reference"), "1.0", Store::Omex)
        .expect("test reference should be valid");
    let add_in = AddIn::new(id, reference).expect("test add-in should be valid");
    let mut pane = Pane::new(add_in);
    pane.relationship_id = relationship_id.to_owned();
    if let Some((part_name, data)) = image {
        pane = pane
            .embed(part_name, "image/png", Arc::new(data.to_vec()))
            .expect("test image should be valid");
    }
    pane
}

fn custom_intent(
    additions: Vec<Pane>,
    operations: Vec<(OwnerSelector, CustomFunctionOperation)>,
) -> WebIntent {
    WebIntent::CustomFunctions(CustomFunctionEdit::ensure(additions, operations))
}

fn encode_test_intent(intent: &WebIntent) -> Vec<u8> {
    encode_intent(intent, MAX_INTENT_BYTES, Limits::standard()).expect("test intent should encode")
}

fn read_u64(bytes: &[u8], cursor: &mut usize) -> usize {
    let end = cursor
        .checked_add(8)
        .expect("wire offset should not overflow");
    let value = u64::from_le_bytes(
        bytes
            .get(*cursor..end)
            .expect("test wire should contain an integer")
            .try_into()
            .expect("integer should have eight bytes"),
    );
    *cursor = end;
    usize::try_from(value).expect("test integer should fit this platform")
}

fn skip_fixed(bytes: &[u8], cursor: &mut usize, length: usize) {
    let end = cursor
        .checked_add(length)
        .expect("wire offset should not overflow");
    assert!(
        bytes.get(*cursor..end).is_some(),
        "test wire should contain field"
    );
    *cursor = end;
}

fn skip_bytes(bytes: &[u8], cursor: &mut usize) -> Range<usize> {
    let length = read_u64(bytes, cursor);
    let start = *cursor;
    skip_fixed(bytes, cursor, length);
    start..*cursor
}

/// Locate the add-in XML payloads in a custom-function intent.  The scanner
/// deliberately understands only the stable private wire envelope needed by
/// the malformed-later-field test; it does not call the production decoder.
fn custom_add_in_spans(bytes: &[u8]) -> Vec<Range<usize>> {
    assert!(bytes.starts_with(INTENT_HEADER));
    let mut cursor = INTENT_HEADER.len();
    assert_eq!(
        bytes[cursor], 2,
        "fixture should be a custom-function intent"
    );
    cursor += 1;
    skip_fixed(bytes, &mut cursor, 1); // graph mode
    let addition_count = read_u64(bytes, &mut cursor);
    let mut spans = Vec::with_capacity(addition_count);
    for _ in 0..addition_count {
        skip_bytes(bytes, &mut cursor); // dock state
        skip_fixed(bytes, &mut cursor, 1 + 8 + 4 + 1); // flags, width, row, lock
        skip_bytes(bytes, &mut cursor); // task-pane relationship ID
        spans.push(skip_bytes(bytes, &mut cursor)); // add-in XML

        let resource_count = read_u64(bytes, &mut cursor);
        for _ in 0..resource_count {
            skip_bytes(bytes, &mut cursor); // snapshot relationship ID
            let target_kind = bytes
                .get(cursor)
                .copied()
                .expect("test resource should contain a target kind");
            cursor += 1;
            match target_kind {
                0 => {
                    skip_bytes(bytes, &mut cursor); // part name
                    skip_bytes(bytes, &mut cursor); // content type
                    skip_bytes(bytes, &mut cursor); // image bytes
                },
                1 => {
                    skip_bytes(bytes, &mut cursor); // external target
                },
                other => panic!("unexpected test resource target kind {other}"),
            }
        }
        let extension_present = bytes
            .get(cursor)
            .copied()
            .expect("test pane should contain extension presence");
        cursor += 1;
        if extension_present == 1 {
            skip_bytes(bytes, &mut cursor);
        } else {
            assert_eq!(extension_present, 0);
        }
    }
    spans
}

fn assert_limit<T>(result: Result<T>, resource: &str, maximum: usize, actual: usize) {
    match result {
        Err(Error::Limit {
            resource: actual_resource,
            max: actual_maximum,
            actual: actual_value,
        }) => {
            assert_eq!(actual_resource, resource);
            assert_eq!(actual_maximum, maximum);
            assert_eq!(actual_value, actual);
        },
        Ok(_) => panic!("over-limit durable intent was accepted"),
        Err(other) => panic!("expected Limit for {resource:?}, got {other:?}"),
    }
}

fn image_data(pane: &Pane) -> &Arc<Vec<u8>> {
    let resource = pane
        .snapshot_resources
        .first()
        .expect("test pane should contain one image");
    match &resource.target {
        SnapshotTarget::Internal { data, .. } => data,
        SnapshotTarget::External { .. } => panic!("test pane should contain an internal image"),
    }
}

#[test]
fn intent_enforces_aggregate_xml_before_parsing_a_later_field() {
    let intent = custom_intent(
        vec![
            pane("first", "rIdPane1", None),
            pane("second", "rIdPane2", None),
        ],
        Vec::new(),
    );
    let bytes = encode_test_intent(&intent);
    let spans = custom_add_in_spans(&bytes);
    assert_eq!(spans.len(), 2);
    let total_xml = spans.iter().map(Range::len).sum::<usize>();
    let per_item_xml = spans.iter().map(Range::len).max().expect("two XML fields");
    assert!(total_xml > per_item_xml);

    let mut within_limits = Limits::standard();
    within_limits.xml_bytes = per_item_xml;
    within_limits.total_xml_bytes = total_xml;
    let decoded = decode_intent(&bytes, &within_limits).expect("aggregate XML budget should fit");
    assert_eq!(decoded, intent);

    let mut malformed_later = bytes.clone();
    for byte in &mut malformed_later[spans[1].clone()] {
        *byte = b'!';
    }
    let mut aggregate_limits = within_limits;
    aggregate_limits.total_xml_bytes = total_xml - 1;
    // The second payload is malformed, but its aggregate admission must fail
    // before the XML parser gets to inspect those bytes.
    assert_limit(
        decode_intent(&malformed_later, &aggregate_limits),
        "aggregate web extension XML bytes",
        total_xml - 1,
        total_xml,
    );
}

#[test]
fn intent_deduplicates_case_folded_shared_snapshot_resources() {
    let image = b"1234";
    let intent = custom_intent(
        vec![
            pane("first", "rIdPane1", Some(("/media/shared.png", image))),
            pane("second", "rIdPane2", Some(("/MEDIA/SHARED.PNG", image))),
        ],
        Vec::new(),
    );
    let bytes = encode_test_intent(&intent);
    let mut limits = Limits::standard();
    limits.image_bytes = image.len();
    limits.total_image_bytes = image.len();

    let decoded = decode_intent(&bytes, &limits)
        .expect("case-folded references to one physical image should share the image budget");
    let WebIntent::CustomFunctions(edit) = decoded else {
        panic!("test intent should decode as custom functions");
    };
    assert_eq!(edit.additions.len(), 2);
    assert!(Arc::ptr_eq(
        image_data(&edit.additions[0]),
        image_data(&edit.additions[1]),
    ));
}

#[test]
fn intent_aggregates_distinct_snapshot_resources() {
    let image = b"1234";
    let intent = custom_intent(
        vec![
            pane("first", "rIdPane1", Some(("/media/one.png", image))),
            pane("second", "rIdPane2", Some(("/media/two.png", image))),
        ],
        Vec::new(),
    );
    let bytes = encode_test_intent(&intent);

    let mut positive_limits = Limits::standard();
    positive_limits.image_bytes = image.len();
    positive_limits.total_image_bytes = image.len() * 2;
    decode_intent(&bytes, &positive_limits)
        .expect("two distinct images should fit their combined image budget");

    let mut aggregate_limits = positive_limits;
    aggregate_limits.total_image_bytes = image.len() * 2 - 1;
    assert_limit(
        decode_intent(&bytes, &aggregate_limits),
        "aggregate web extension image bytes",
        image.len() * 2 - 1,
        image.len() * 2,
    );
}

#[test]
fn intent_rejects_add_in_extension_in_pane_slot_before_retention() {
    let mut addition = pane("wrong-kind", "rIdPane1", None);
    addition.extension_list =
        Some(ExtList::empty(ExtKind::AddIn).expect("test add-in extension should be valid"));
    let bytes = encode_test_intent(&custom_intent(vec![addition], Vec::new()));

    let error = decode_intent(&bytes, &Limits::standard())
        .expect_err("an add-in extLst must not be accepted in a pane slot");
    assert!(matches!(
        error,
        Error::Invalid(message) if message.contains(TASK_PANES_NAMESPACE)
    ));
}

fn parser_string_cost(pane: &Pane) -> usize {
    let limits = Limits::standard();
    let xml = write_add_in_with(&pane.add_in, Conformance::Transitional, &limits)
        .expect("test add-in should encode");
    let mut budget = OperationBudget::default();
    parse_add_in_with_budget(&xml, &limits, &mut budget).expect("test add-in should parse");
    budget.string_bytes
}

#[test]
fn intent_charges_xml_derived_and_direct_strings_cumulatively() {
    let first = pane("first", "rIdPane1", None);
    let second = pane("second", "rIdPane2", None);
    let operation_id = "x".repeat(257);
    let intent = custom_intent(
        vec![first.clone(), second.clone()],
        vec![(
            OwnerSelector::PaneIndex(0),
            CustomFunctionOperation::InsertId {
                index: 0,
                id: operation_id.clone(),
            },
        )],
    );
    let bytes = encode_test_intent(&intent);

    // The parser budget accounts for retained XML-derived strings.  The
    // pane's dock and relationship IDs are direct wire strings, and the
    // operation ID is decoded after both XML documents.  Leave exactly one
    // byte short of the combined pre-operation ledger so the final direct
    // string fails only after the XML contributions were admitted.
    let base_string_bytes = [first, second]
        .iter()
        .map(|pane| parser_string_cost(pane) + pane.dock_state().len() + pane.relationship_id.len())
        .sum::<usize>();
    let mut positive_limits = Limits::standard();
    positive_limits.total_string_bytes = base_string_bytes + operation_id.len();
    decode_intent(&bytes, &positive_limits)
        .expect("XML-derived and direct strings should fit at the exact total");

    let mut aggregate_limits = positive_limits;
    aggregate_limits.total_string_bytes = base_string_bytes + operation_id.len() - 1;
    assert_limit(
        decode_intent(&bytes, &aggregate_limits),
        "Web Extensions durable decoded strings",
        aggregate_limits.total_string_bytes,
        base_string_bytes + operation_id.len(),
    );
}
