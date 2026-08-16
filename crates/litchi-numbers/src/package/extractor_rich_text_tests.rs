fn rich_text_reference(identifier: u64) -> Vec<u8> {
    let mut reference = Vec::new();
    append_varint_field(&mut reference, 1, identifier)
        .unwrap_or_else(|error| panic!("reference should encode: {error}"));
    reference
}

fn rich_text_payload(identifier: u64) -> Vec<u8> {
    let reference = rich_text_reference(identifier);
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 1, &reference)
        .unwrap_or_else(|error| panic!("storage reference should encode: {error}"));
    // CellID is a required message in the native schema.  An empty message is
    // the generated decoder's default and is sufficient for this envelope
    // projection: the extractor only needs to identify the storage edge.
    append_length_delimited_field(&mut payload, 3, &[])
        .unwrap_or_else(|error| panic!("cell owner should encode: {error}"));
    payload
}

fn assert_rich_text_payload_rejected(source: &[u8]) {
    assert!(
        super::preflight_rich_text_payload(source).is_err(),
        "malformed rich-text payload was accepted: {source:02x?}"
    );
}

#[test]
fn rich_text_projection_matches_generated_oracle_and_reports_bounded_work() {
    let source = rich_text_payload(42);
    let oracle = tst::RichTextPayloadArchive::decode(source.as_slice())
        .unwrap_or_else(|error| panic!("generated oracle rejected valid payload: {error}"));
    let (storage, report) = super::preflight_rich_text_payload(&source)
        .unwrap_or_else(|error| panic!("strict rich-text projection rejected payload: {error}"));

    assert_eq!(storage, oracle.storage.identifier);
    assert_eq!(report.scanned_bytes(), source.len());
    assert_eq!(report.fields(), 2);
    assert_eq!(report.messages(), 1);
    assert_eq!(report.max_depth(), 0);
}

#[test]
fn rich_text_projection_preserves_unknown_fields_and_matches_oracle() {
    let mut source = rich_text_payload(7);
    // Unknown length-delimited bytes are intentionally opaque to this narrow
    // projection.  Include bytes which are not a valid nested protobuf so the
    // test proves the scanner does not recurse into unrelated fields.
    append_length_delimited_field(&mut source, 100, &[0xff, 0x80, 0x00])
        .unwrap_or_else(|error| panic!("unknown field should encode: {error}"));
    append_varint_field(&mut source, 101, u64::MAX)
        .unwrap_or_else(|error| panic!("unknown scalar should encode: {error}"));

    let oracle = tst::RichTextPayloadArchive::decode(source.as_slice())
        .unwrap_or_else(|error| panic!("generated oracle rejected unknown fields: {error}"));
    let (storage, report) = super::preflight_rich_text_payload(&source)
        .unwrap_or_else(|error| panic!("unknown fields changed projection semantics: {error}"));

    assert_eq!(storage, oracle.storage.identifier);
    assert_eq!(report.scanned_bytes(), source.len());
    assert_eq!(report.fields(), 4);
    assert_eq!(report.messages(), 1);
}

#[test]
fn rich_text_projection_rejects_missing_duplicate_and_wrong_wire_edges() {
    let reference = rich_text_reference(42);
    let mut missing_storage = Vec::new();
    append_length_delimited_field(&mut missing_storage, 3, &[])
        .unwrap_or_else(|error| panic!("cell owner should encode: {error}"));
    assert_rich_text_payload_rejected(&missing_storage);

    let mut missing_cell = Vec::new();
    append_length_delimited_field(&mut missing_cell, 1, &reference)
        .unwrap_or_else(|error| panic!("storage reference should encode: {error}"));
    assert_rich_text_payload_rejected(&missing_cell);

    let mut duplicate_storage = rich_text_payload(42);
    append_length_delimited_field(&mut duplicate_storage, 1, &reference)
        .unwrap_or_else(|error| panic!("duplicate storage should encode: {error}"));
    assert_rich_text_payload_rejected(&duplicate_storage);

    let mut duplicate_cell = rich_text_payload(42);
    append_length_delimited_field(&mut duplicate_cell, 3, &[])
        .unwrap_or_else(|error| panic!("duplicate cell owner should encode: {error}"));
    assert_rich_text_payload_rejected(&duplicate_cell);

    let mut wrong_storage_wire = vec![0x08, 0x2a];
    append_length_delimited_field(&mut wrong_storage_wire, 3, &[])
        .unwrap_or_else(|error| panic!("cell owner should encode: {error}"));
    assert_rich_text_payload_rejected(&wrong_storage_wire);

    let mut wrong_cell_wire = rich_text_payload(42);
    // Replace the canonical field-3 length-delimited key with a varint key.
    wrong_cell_wire.truncate(wrong_cell_wire.len() - 2);
    wrong_cell_wire.extend_from_slice(&[0x18, 0x01]);
    assert_rich_text_payload_rejected(&wrong_cell_wire);
}

#[test]
fn rich_text_projection_rejects_invalid_local_references() {
    let mut zero = Vec::new();
    append_varint_field(&mut zero, 1, 0)
        .unwrap_or_else(|error| panic!("zero reference should encode: {error}"));
    let mut zero_payload = Vec::new();
    append_length_delimited_field(&mut zero_payload, 1, &zero)
        .unwrap_or_else(|error| panic!("storage reference should encode: {error}"));
    append_length_delimited_field(&mut zero_payload, 3, &[])
        .unwrap_or_else(|error| panic!("cell owner should encode: {error}"));
    assert_rich_text_payload_rejected(&zero_payload);

    let mut external = rich_text_reference(42);
    append_varint_field(&mut external, 3, 1)
        .unwrap_or_else(|error| panic!("external flag should encode: {error}"));
    let mut external_payload = Vec::new();
    append_length_delimited_field(&mut external_payload, 1, &external)
        .unwrap_or_else(|error| panic!("storage reference should encode: {error}"));
    append_length_delimited_field(&mut external_payload, 3, &[])
        .unwrap_or_else(|error| panic!("cell owner should encode: {error}"));
    assert_rich_text_payload_rejected(&external_payload);

    let mut duplicate_identifier = rich_text_reference(42);
    append_varint_field(&mut duplicate_identifier, 1, 43)
        .unwrap_or_else(|error| panic!("duplicate identifier should encode: {error}"));
    let mut duplicate_payload = Vec::new();
    append_length_delimited_field(&mut duplicate_payload, 1, &duplicate_identifier)
        .unwrap_or_else(|error| panic!("storage reference should encode: {error}"));
    append_length_delimited_field(&mut duplicate_payload, 3, &[])
        .unwrap_or_else(|error| panic!("cell owner should encode: {error}"));
    assert_rich_text_payload_rejected(&duplicate_payload);

    let noncanonical_value = [0x08, 0x81, 0x00];
    let mut noncanonical_payload = Vec::new();
    append_length_delimited_field(&mut noncanonical_payload, 1, &noncanonical_value)
        .unwrap_or_else(|error| panic!("storage reference should encode: {error}"));
    append_length_delimited_field(&mut noncanonical_payload, 3, &[])
        .unwrap_or_else(|error| panic!("cell owner should encode: {error}"));
    assert_rich_text_payload_rejected(&noncanonical_payload);
}

#[test]
fn rich_text_projection_rejects_noncanonical_selected_framing() {
    let reference = rich_text_reference(42);

    // Field 1's key (0x0a) is encoded as the redundant two-byte varint
    // 0x8a 0x00.  The selected key must be canonical even when its value is
    // otherwise valid.
    let mut noncanonical_key = vec![0x8a, 0x00, reference.len() as u8];
    noncanonical_key.extend_from_slice(&reference);
    append_length_delimited_field(&mut noncanonical_key, 3, &[])
        .unwrap_or_else(|error| panic!("cell owner should encode: {error}"));
    assert_rich_text_payload_rejected(&noncanonical_key);

    // Field 1's length is encoded as the redundant two-byte varint 0x82 0x00.
    let mut noncanonical_length = vec![0x0a, 0x82, 0x00];
    noncanonical_length.extend_from_slice(&reference);
    append_length_delimited_field(&mut noncanonical_length, 3, &[])
        .unwrap_or_else(|error| panic!("cell owner should encode: {error}"));
    assert_rich_text_payload_rejected(&noncanonical_length);

    // The nested local-reference key is likewise selected and must be
    // canonical.  Its key value is 0x08, redundantly encoded as 0x88 0x00.
    let noncanonical_nested_key = [0x88, 0x00, 0x2a];
    let mut nested_key_payload = Vec::new();
    append_length_delimited_field(&mut nested_key_payload, 1, &noncanonical_nested_key)
        .unwrap_or_else(|error| panic!("storage reference should encode: {error}"));
    append_length_delimited_field(&mut nested_key_payload, 3, &[])
        .unwrap_or_else(|error| panic!("cell owner should encode: {error}"));
    assert_rich_text_payload_rejected(&nested_key_payload);
}

#[test]
fn rich_text_projection_rejects_truncated_and_unbalanced_wire() {
    assert_rich_text_payload_rejected(&[0x0a]);
    assert_rich_text_payload_rejected(&[0x0a, 0x02, 0x08]);

    let mut unbalanced_group = rich_text_payload(42);
    unbalanced_group.extend_from_slice(&[0x1b]);
    assert_rich_text_payload_rejected(&unbalanced_group);

    let mut trailing = rich_text_payload(42);
    trailing.push(0xff);
    assert_rich_text_payload_rejected(&trailing);
}

#[test]
fn rich_text_projection_reports_unknown_field_work_without_descending() {
    let opaque = vec![0xff; 4_096];
    let mut source = rich_text_payload(99);
    append_length_delimited_field(&mut source, 127, &opaque)
        .unwrap_or_else(|error| panic!("opaque unknown field should encode: {error}"));

    let (storage, report) = super::preflight_rich_text_payload(&source)
        .unwrap_or_else(|error| panic!("opaque unknown field should be skipped: {error}"));
    assert_eq!(storage, 99);
    assert_eq!(report.scanned_bytes(), source.len());
    assert_eq!(report.fields(), 3);
    assert_eq!(report.messages(), 1);
    assert_eq!(report.max_depth(), 0);
}

#[test]
fn rich_text_projection_budget_is_inclusive_for_fields_and_work() {
    let source = rich_text_payload(99);
    let (_, report) = super::preflight_rich_text_payload(&source)
        .unwrap_or_else(|error| panic!("valid source rejected: {error}"));

    let mut exact = ProjectionBudget::new(SemanticLimits::default());
    exact.payload_fields = crate::MAX_REFERENCES - report.fields();
    exact.payload_work = super::MAX_PAYLOAD_WORK - report.scanned_bytes();
    exact
        .charge_wire_preflight(report)
        .unwrap_or_else(|error| panic!("inclusive rich-text budget rejected: {error}"));
    assert_eq!(exact.payload_fields, crate::MAX_REFERENCES);
    assert_eq!(exact.payload_work, super::MAX_PAYLOAD_WORK);

    let mut fields_over = ProjectionBudget::new(SemanticLimits::default());
    fields_over.payload_fields = crate::MAX_REFERENCES - report.fields() + 1;
    let before_fields = fields_over.payload_fields;
    assert!(matches!(
        fields_over.charge_wire_preflight(report),
        Err(Error::SemanticLimit {
            kind: SemanticLimitKind::Objects,
            observed,
            maximum: crate::MAX_REFERENCES,
            ..
        }) if observed == crate::MAX_REFERENCES + 1
    ));
    assert_eq!(fields_over.payload_fields, before_fields);

    let mut work_over = ProjectionBudget::new(SemanticLimits::default());
    work_over.payload_work = super::MAX_PAYLOAD_WORK - report.scanned_bytes() + 1;
    assert!(matches!(
        work_over.charge_wire_preflight(report),
        Err(Error::SemanticLimit {
            kind: SemanticLimitKind::FormulaWork,
            observed,
            maximum: super::MAX_PAYLOAD_WORK,
            ..
        }) if observed == super::MAX_PAYLOAD_WORK + 1
    ));
}

#[test]
fn rich_text_projection_failure_is_atomic_and_source_is_unchanged() {
    let source = rich_text_payload(42);
    let mut malformed = source.clone();
    // A duplicate selected field appears after a complete valid prefix.  The
    // projection must not publish the prefix's storage identifier or report.
    append_length_delimited_field(&mut malformed, 1, &rich_text_reference(43))
        .unwrap_or_else(|error| panic!("duplicate storage should encode: {error}"));
    let before = malformed.clone();
    assert_rich_text_payload_rejected(&malformed);
    assert_eq!(malformed, before);

    // A failed attempt must not contaminate any state used by a subsequent
    // source.  This models the caller's candidate-then-publication boundary.
    let (storage, report) = super::preflight_rich_text_payload(&source)
        .unwrap_or_else(|error| panic!("valid source was contaminated: {error}"));
    assert_eq!(storage, 42);
    assert_eq!(report.fields(), 2);
    assert_eq!(report.messages(), 1);
}
