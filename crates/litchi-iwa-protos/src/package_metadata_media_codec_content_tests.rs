use super::*;

fn varint(output: &mut Vec<u8>, mut value: u64) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output.push(byte);
        if value == 0 {
            return;
        }
    }
}

fn varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    varint(output, u64::from(number) << 3);
    varint(output, value);
}

fn bytes_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) {
    varint(output, (u64::from(number) << 3) | 2);
    varint(
        output,
        u64::try_from(payload.len()).expect("test payload fits in a varint"),
    );
    output.extend_from_slice(payload);
}

fn fixed64_field(output: &mut Vec<u8>, number: u32, value: u64) {
    varint(output, (u64::from(number) << 3) | 1);
    output.extend_from_slice(&value.to_le_bytes());
}

fn digest(seed: u8) -> [u8; SHA1_DIGEST_BYTES] {
    let mut value = [0; SHA1_DIGEST_BYTES];
    for (index, byte) in value.iter_mut().enumerate() {
        *byte = seed.wrapping_add(u8::try_from(index).expect("digest index fits"));
    }
    value
}

fn rich_data_info(identifier: u64, digest: &[u8], length: u64, unknown: bool) -> Vec<u8> {
    let mut output = Vec::new();
    varint_field(&mut output, DATA_IDENTIFIER_FIELD, identifier);
    bytes_field(&mut output, DATA_DIGEST_FIELD, digest);
    bytes_field(&mut output, DATA_PREFERRED_NAME_FIELD, b"clip.m4a");
    bytes_field(&mut output, DATA_FILE_NAME_FIELD, b"clip.m4a");
    bytes_field(&mut output, 5, b"resource://clip");
    bytes_field(&mut output, 6, b"bookmark");
    bytes_field(&mut output, 7, b"https://example.invalid/clip");
    varint_field(&mut output, 8, 1);
    varint_field(&mut output, 9, 2);
    let mut attributes = Vec::new();
    varint_field(&mut attributes, 1, 7);
    bytes_field(&mut output, 10, &attributes);
    bytes_field(&mut output, 11, &[0x08, 0x01]);
    bytes_field(&mut output, 12, b"old-digest");
    bytes_field(&mut output, 13, &[0x10, 0x02]);
    varint_field(&mut output, 14, length);
    varint_field(&mut output, 15, 0);
    varint_field(&mut output, 16, 3);
    fixed64_field(&mut output, 17, 0x0102_0304_0506_0708);
    varint_field(&mut output, DATA_MATERIALIZED_LENGTH_FIELD, length);
    bytes_field(&mut output, 99, b"pasteboard-path");
    if unknown {
        bytes_field(&mut output, 27, b"future-data-info");
    }
    output
}

fn owner(object_identifier: u64, count: u32) -> Vec<u8> {
    let mut output = Vec::new();
    varint_field(
        &mut output,
        OWNER_OBJECT_IDENTIFIER_FIELD,
        object_identifier,
    );
    varint_field(&mut output, OWNER_COUNT_FIELD, u64::from(count));
    output
}

fn component(data_identifier: u64) -> Vec<u8> {
    component_with_identity(9, b"Document", data_identifier)
}

fn component_with_identity(
    identifier: u64,
    preferred_locator: &[u8],
    data_identifier: u64,
) -> Vec<u8> {
    let mut reference = Vec::new();
    varint_field(&mut reference, DATA_IDENTIFIER_FIELD, data_identifier);
    bytes_field(&mut reference, OWNER_COUNT_FIELD, &owner(700, 2));

    let mut output = Vec::new();
    varint_field(&mut output, COMPONENT_IDENTIFIER_FIELD, identifier);
    bytes_field(
        &mut output,
        COMPONENT_PREFERRED_LOCATOR_FIELD,
        preferred_locator,
    );
    bytes_field(&mut output, COMPONENT_DATA_REFERENCE_FIELD, &reference);
    output
}

fn source_with_data_info(data_info: &[u8], root_unknown: bool) -> Vec<u8> {
    let mut output = Vec::new();
    varint_field(&mut output, ROOT_LAST_IDENTIFIER_FIELD, 10);
    bytes_field(&mut output, ROOT_COMPONENT_FIELD, &component(41));
    bytes_field(&mut output, ROOT_DATA_INFO_FIELD, data_info);
    if root_unknown {
        varint_field(&mut output, 25, 0xcafe);
    }
    output
}

fn source_with_versioned_component(data_info: &[u8]) -> Vec<u8> {
    let mut output = source_with_data_info(data_info, false);
    bytes_field(&mut output, ROOT_VERSIONED_COMPONENT_FIELD, &component(41));
    output
}

fn source_with_many_records(count: u64) -> Vec<u8> {
    let mut output = Vec::new();
    varint_field(&mut output, ROOT_LAST_IDENTIFIER_FIELD, 10_000);
    for identifier in 1..=count {
        let old_digest = digest(identifier as u8);
        let data_info = rich_data_info(identifier, &old_digest, identifier + 1, false);
        bytes_field(&mut output, ROOT_DATA_INFO_FIELD, &data_info);
        let component_identifier = 100 + identifier;
        let locator = format!("Document-{identifier}");
        let component =
            component_with_identity(component_identifier, locator.as_bytes(), identifier);
        bytes_field(&mut output, ROOT_COMPONENT_FIELD, &component);
    }
    output
}

fn rewrite_options(source: &[u8]) -> DecodeOptions {
    let baseline = DecodeOptions::for_source(source);
    DecodeOptions::new(
        baseline.max_message_bytes(),
        baseline.max_fields(),
        source.len().saturating_mul(64).max(1),
        baseline.max_components(),
        baseline.max_data_records(),
        baseline.max_owners(),
        baseline.max_digest_bytes(),
        baseline.max_name_bytes(),
        baseline.max_depth(),
    )
    .with_max_output_bytes(source.len().saturating_mul(4).max(1))
}

#[test]
fn content_replacement_updates_only_digest_and_length_and_preserves_owner_envelope() {
    let old_digest = digest(0x10);
    let new_digest = digest(0xb0);
    let data_info = rich_data_info(41, &old_digest, 127, false);
    let source = source_with_data_info(&data_info, true);
    let component_bytes = {
        let mut value = Vec::new();
        bytes_field(&mut value, ROOT_COMPONENT_FIELD, &component(41));
        value
    };
    let replacement = DataInfoContentReplacement::new(41, &old_digest, &new_digest, 127, 128);
    let replacements = [replacement];
    let batch = MediaRewriteBatch::empty().with_content_replacements(&replacements);
    let output = rewrite_package_metadata_media(&source, batch, rewrite_options(&source))
        .expect("rich native optional fields are a safe selected shape");

    assert_eq!(output.report().data_replacements(), 1);
    assert_eq!(output.report().output_bytes(), output.bytes().len());
    assert!(
        output
            .bytes()
            .windows(component_bytes.len())
            .any(|window| { window == component_bytes.as_slice() })
    );
    let mut facts = Vec::new();
    visit_package_metadata_media(
        output.bytes(),
        DecodeOptions::for_source(output.bytes()),
        &mut Facts { values: &mut facts },
    )
    .expect("rewritten metadata remains strictly inspectable");
    assert_eq!(facts, vec![(41, new_digest, Some(128))]);
}

struct Facts<'a> {
    values: &'a mut Vec<(u64, [u8; SHA1_DIGEST_BYTES], Option<u64>)>,
}

impl PackageMetadataMediaVisitor for Facts<'_> {
    fn visit_data_info(&mut self, snapshot: DataInfoSnapshot<'_>) -> Result<(), DecodeError> {
        let mut digest = [0; SHA1_DIGEST_BYTES];
        digest.copy_from_slice(snapshot.digest());
        self.values.push((
            snapshot.identifier(),
            digest,
            snapshot.materialized_length(),
        ));
        Ok(())
    }
}

#[test]
fn content_replacement_is_exactly_reversible_and_noop_preserves_source() {
    let old_digest = digest(0x20);
    let new_digest = digest(0x90);
    let source = source_with_data_info(&rich_data_info(41, &old_digest, 127, false), false);
    let forward = [DataInfoContentReplacement::new(
        41,
        &old_digest,
        &new_digest,
        127,
        16_384,
    )];
    let changed = rewrite_package_metadata_media_content_replacements(
        &source,
        &forward,
        rewrite_options(&source),
    )
    .expect("content replacement");
    let inverse = [DataInfoContentReplacement::new(
        41,
        &new_digest,
        &old_digest,
        16_384,
        127,
    )];
    let restored = rewrite_package_metadata_media_content_replacements(
        changed.bytes(),
        &inverse,
        rewrite_options(changed.bytes()),
    )
    .expect("inverse content replacement");
    assert_eq!(restored.bytes(), source.as_slice());

    let noop = [DataInfoContentReplacement::new(
        41,
        &old_digest,
        &old_digest,
        127,
        127,
    )];
    let unchanged = rewrite_package_metadata_media_content_replacements(
        &source,
        &noop,
        rewrite_options(&source),
    )
    .expect("no-op compare-and-set");
    assert_eq!(unchanged.bytes(), source.as_slice());
}

#[test]
fn content_replacement_rejects_stale_duplicate_and_opaque_selected_records() {
    let old_digest = digest(0x30);
    let new_digest = digest(0x40);
    let source = source_with_data_info(&rich_data_info(41, &old_digest, 9, false), false);
    let stale_digest = digest(0x31);
    let stale = [DataInfoContentReplacement::new(
        41,
        &stale_digest,
        &new_digest,
        9,
        10,
    )];
    let error = rewrite_package_metadata_media_content_replacements(
        &source,
        &stale,
        rewrite_options(&source),
    )
    .expect_err("stale digest must fail closed");
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::DataInfoContentMismatch)
    );

    let stale_length = [DataInfoContentReplacement::new(
        41,
        &old_digest,
        &new_digest,
        10,
        11,
    )];
    let error = rewrite_package_metadata_media_content_replacements(
        &source,
        &stale_length,
        rewrite_options(&source),
    )
    .expect_err("stale length must fail closed");
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::DataInfoContentMismatch)
    );

    let duplicate = [
        DataInfoContentReplacement::new(41, &old_digest, &new_digest, 9, 10),
        DataInfoContentReplacement::new(41, &old_digest, &new_digest, 9, 11),
    ];
    let error = rewrite_package_metadata_media_content_replacements(
        &source,
        &duplicate,
        rewrite_options(&source),
    )
    .expect_err("duplicate replacements must fail atomically");
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::DuplicateOperation)
    );

    let opaque_source = source_with_data_info(&rich_data_info(41, &old_digest, 9, true), false);
    let opaque = [DataInfoContentReplacement::new(
        41,
        &old_digest,
        &new_digest,
        9,
        10,
    )];
    let error = rewrite_package_metadata_media_content_replacements(
        &opaque_source,
        &opaque,
        rewrite_options(&opaque_source),
    )
    .expect_err("unknown selected fields must fail closed");
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::UnknownSelectedRecord)
    );
}

#[test]
fn content_replacement_exposes_finite_preflight_requirements() {
    let old_digest = digest(0x50);
    let new_digest = digest(0x60);
    let source = source_with_data_info(&rich_data_info(41, &old_digest, 2, false), false);
    let replacements = [DataInfoContentReplacement::new(
        41,
        &old_digest,
        &new_digest,
        2,
        3,
    )];
    let prepared = prepare_package_metadata_media_content_replacements(
        &source,
        &replacements,
        rewrite_options(&source),
    )
    .expect("content replacement should preflight");
    let requirements = prepared.execution_requirements();
    assert_eq!(requirements.data_replacements(), 1);
    assert_eq!(requirements.allocations(), 1);
    assert_eq!(requirements.retained_bytes(), requirements.output_bytes());
    let output = prepared
        .execute(requirements.exact_limits())
        .expect("exact finite limits should execute");
    assert_eq!(output.report().data_replacements(), 1);
}

#[test]
fn exact_limits_cover_data_and_owner_addition_topology_and_rescans() {
    let old_digest = digest(0x58);
    let addition_digest = digest(0x68);
    let source = source_with_data_info(&rich_data_info(41, &old_digest, 2, false), false);
    let data_additions = [DataInfoAddition::new(42, &addition_digest, "new.m4a")
        .with_file_name("new.m4a")
        .with_materialized_length(3)];
    let owner_additions = [DataReferenceOwnerAddition::new(
        ComponentSelector::new(9, "Document"),
        42,
        701,
        1,
    )];
    let batch = MediaRewriteBatch::new(&data_additions, &[], &owner_additions, &[]);
    let prepared = prepare_package_metadata_media_rewrite(&source, batch, rewrite_options(&source))
        .expect("additions should preflight with finite bounds");
    let requirements = prepared.execution_requirements();
    assert_eq!(requirements.data_records(), 2);
    assert_eq!(requirements.owners(), 2);
    let output = prepared
        .execute(requirements.exact_limits())
        .expect("exact limits must admit every addition and verification pass");
    assert_eq!(output.report().data_additions(), 1);
    assert_eq!(output.report().owner_additions(), 1);
    let report =
        inspect_package_metadata_media(output.bytes(), DecodeOptions::for_source(output.bytes()))
            .expect("candidate remains strictly inspectable");
    assert_eq!(report.data_records(), 2);
    assert_eq!(report.owners(), 2);
}

#[test]
fn many_data_additions_are_fully_charged_by_exact_field_requirements() {
    let old_digest = digest(0x5a);
    let addition_digest = digest(0x6a);
    let source = source_with_data_info(&rich_data_info(41, &old_digest, 2, false), false);
    let additions: Vec<_> = (0..16)
        .map(|index| {
            DataInfoAddition::new(100 + index, &addition_digest, "new.m4a")
                .with_file_name("new.m4a")
                .with_materialized_length(3)
        })
        .collect();
    let batch = MediaRewriteBatch::new(&additions, &[], &[], &[]);
    let options = rewrite_options(&source)
        .with_max_work_bytes(source.len().saturating_mul(256))
        .with_max_output_bytes(source.len().saturating_mul(16));
    let prepared = prepare_package_metadata_media_rewrite(&source, batch, options)
        .expect("many additions should preflight with explicit field charges");
    let requirements = prepared.execution_requirements();
    assert!(requirements.fields() >= additions.len().saturating_mul(6));
    let output = prepared
        .execute(requirements.exact_limits())
        .expect("exact field and topology limits must admit every addition");
    let report =
        inspect_package_metadata_media(output.bytes(), DecodeOptions::for_source(output.bytes()))
            .expect("many additions remain inspectable");
    assert_eq!(report.data_records(), additions.len() + 1);
}

#[test]
fn current_and_versioned_component_namespaces_are_distinct_for_owner_updates() {
    let old_digest = digest(0x70);
    let source = source_with_versioned_component(&rich_data_info(41, &old_digest, 2, false));
    let selector = ComponentSelector::new(9, "Document");
    let update = [DataReferenceOwnerCountUpdate::new(selector, 41, 700, 2, 3)];
    let batch = MediaRewriteBatch::empty().with_owner_count_updates(&update);
    let output = rewrite_package_metadata_media(&source, batch, rewrite_options(&source))
        .expect("versioned and current component namespaces are independently valid");
    assert_eq!(output.report().components(), 2);
    assert_eq!(output.report().owner_updates(), 1);
}

#[test]
fn duplicate_identity_audit_charges_work_before_identity_sorting() {
    let digest = digest(0x80);
    let source = source_with_data_info(&rich_data_info(41, &digest, 2, false), false);
    let options = DecodeOptions::for_source(&source).with_max_work_bytes(1);
    let error = inspect_package_metadata_media(&source, options)
        .expect_err("duplicate audit must be bounded before entering identity sorting");
    assert!(matches!(
        error.resource_limit(),
        Some(DecodeLimit::Work { .. })
    ));
}

#[test]
fn for_source_admits_native_scale_unique_media_records() {
    let source = source_with_many_records(32);
    let report = inspect_package_metadata_media(&source, DecodeOptions::for_source(&source))
        .expect("bounded uniqueness audit admits native-scale metadata");
    assert_eq!(report.data_records(), 32);
    assert_eq!(report.components(), 32);
    assert_eq!(report.data_references(), 32);
    assert_eq!(report.owners(), 32);
}

#[test]
fn component_identity_hash_collisions_are_sorted_before_duplicate_check() {
    let mut identities = vec![
        ComponentIdentity {
            versioned: false,
            identifier: 7,
            locator: "A",
            locator_hash: 1,
        },
        ComponentIdentity {
            versioned: false,
            identifier: 7,
            locator: "B",
            locator_hash: 1,
        },
        ComponentIdentity {
            versioned: false,
            identifier: 7,
            locator: "A",
            locator_hash: 1,
        },
    ];
    let source = vec![0; 128];
    let options = DecodeOptions::for_source(&source);
    let mut state = ScanState::new(&source, options).expect("test scan state");
    let error = finish_component_identities(&mut identities, options, &mut state)
        .expect_err("A, B, A must remain a duplicate after a forced hash collision");
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::DuplicateComponent)
    );
}
