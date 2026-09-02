use litchi_iwa_protos::package_metadata_media_codec::{
    ComponentDataReferenceSnapshot, ComponentSelector, ComponentSnapshot, DataInfoAddition,
    DataInfoRemoval, DataInfoSnapshot, DataReferenceOwnerAddition, DataReferenceOwnerCountUpdate,
    DataReferenceOwnerRemoval, DecodeError, DecodeLimit, DecodeOptions, InvalidReason,
    MediaRewriteBatch, OwnerSnapshot, PackageMetadataMediaVisitor, inspect_package_metadata_media,
    prepare_package_metadata_media_rewrite, rewrite_package_metadata_media,
    visit_package_metadata_media,
};

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

fn field_varint(output: &mut Vec<u8>, field: u32, value: u64) {
    varint(output, u64::from(field) << 3);
    varint(output, value);
}

fn field_bytes(output: &mut Vec<u8>, field: u32, payload: &[u8]) {
    varint(output, (u64::from(field) << 3) | 2);
    varint(
        output,
        u64::try_from(payload.len()).expect("test payload length fits in a varint"),
    );
    output.extend_from_slice(payload);
}

fn owner(object_identifier: u64, count: u32, unknown: bool) -> Vec<u8> {
    let mut output = Vec::new();
    field_varint(&mut output, 1, object_identifier);
    field_varint(&mut output, 2, u64::from(count));
    if unknown {
        field_varint(&mut output, 31, 0xfeed);
    }
    output
}

fn data_reference(data_identifier: u64, owners: &[(u64, u32)], unknown: bool) -> Vec<u8> {
    let mut output = Vec::new();
    field_varint(&mut output, 1, data_identifier);
    for &(object_identifier, count) in owners {
        field_bytes(&mut output, 2, &owner(object_identifier, count, unknown));
    }
    if unknown {
        field_varint(&mut output, 30, 0xbeef);
    }
    output
}

fn component(
    identifier: u64,
    preferred_locator: &str,
    locator: Option<&str>,
    references: &[(u64, &[(u64, u32)], bool)],
    unknown: bool,
) -> Vec<u8> {
    let mut output = Vec::new();
    field_varint(&mut output, 1, identifier);
    field_bytes(&mut output, 2, preferred_locator.as_bytes());
    if let Some(locator) = locator {
        field_bytes(&mut output, 3, locator.as_bytes());
    }
    for &(data_identifier, owners, reference_unknown) in references {
        field_bytes(
            &mut output,
            7,
            &data_reference(data_identifier, owners, reference_unknown),
        );
    }
    if unknown {
        field_varint(&mut output, 29, 0xabcd);
    }
    output
}

fn data_info(
    identifier: u64,
    digest: &[u8],
    preferred_file_name: &str,
    file_name: Option<&str>,
    materialized_length: Option<u64>,
    unknown: bool,
) -> Vec<u8> {
    let mut output = Vec::new();
    field_varint(&mut output, 1, identifier);
    field_bytes(&mut output, 2, digest);
    field_bytes(&mut output, 3, preferred_file_name.as_bytes());
    if let Some(file_name) = file_name {
        field_bytes(&mut output, 4, file_name.as_bytes());
    }
    if let Some(materialized_length) = materialized_length {
        field_varint(&mut output, 18, materialized_length);
    }
    if unknown {
        field_bytes(&mut output, 27, b"future-data-info");
    }
    output
}

fn package(
    last_object_identifier: u64,
    components: &[Vec<u8>],
    data_infos: &[Vec<u8>],
    unknown_root: bool,
) -> Vec<u8> {
    let mut output = Vec::new();
    field_varint(&mut output, 1, last_object_identifier);
    for component in components {
        field_bytes(&mut output, 3, component);
    }
    for data_info in data_infos {
        field_bytes(&mut output, 4, data_info);
    }
    if unknown_root {
        field_varint(&mut output, 25, 0xcafe);
    }
    output
}

fn digest(seed: u8) -> [u8; 20] {
    let mut digest = [0; 20];
    for (index, byte) in digest.iter_mut().enumerate() {
        *byte = seed.wrapping_add(u8::try_from(index).expect("digest index fits"));
    }
    digest
}

fn options(source: &[u8]) -> DecodeOptions {
    DecodeOptions::for_source(source)
}

fn canonical_source() -> Vec<u8> {
    let first_digest = digest(0x10);
    let second_digest = digest(0x80);
    let first_owners = [(91, 1), (92, 2)];
    let second_owners = [(93, 1)];
    let components = [
        component(
            7,
            "Index/Document.iwa",
            None,
            &[(41, &first_owners, false), (42, &second_owners, false)],
            false,
        ),
        component(
            8,
            "Index/Metadata.iwa",
            Some("Index/Metadata.iwa"),
            &[],
            false,
        ),
    ];
    let data_infos = [
        data_info(
            41,
            &first_digest,
            "intro.m4a",
            Some("intro.m4a"),
            Some(4_096),
            false,
        ),
        data_info(42, &second_digest, "outro.m4a", None, None, false),
    ];
    package(200, &components, &data_infos, false)
}

fn source_with_unknowns() -> Vec<u8> {
    let first_digest = digest(0x20);
    let owners = [(91, 1)];
    let components = [component(
        7,
        "Index/Document.iwa",
        None,
        &[(41, &owners, true)],
        true,
    )];
    let data_infos = [data_info(
        41,
        &first_digest,
        "intro.m4a",
        Some("intro.m4a"),
        Some(4_096),
        true,
    )];
    package(200, &components, &data_infos, true)
}

#[derive(Default)]
struct Facts {
    data_infos: Vec<(u64, Vec<u8>, String, Option<String>, Option<u64>, bool)>,
    components: Vec<(u64, String, Option<String>, bool)>,
    references: Vec<(u64, u64, usize, bool)>,
    owners: Vec<(u64, u64, u32, bool)>,
}

impl PackageMetadataMediaVisitor for Facts {
    fn visit_data_info(&mut self, data_info: DataInfoSnapshot<'_>) -> Result<(), DecodeError> {
        self.data_infos.push((
            data_info.identifier(),
            data_info.digest().to_vec(),
            data_info.preferred_file_name().to_owned(),
            data_info.file_name().map(str::to_owned),
            data_info.materialized_length(),
            data_info.has_unknown_fields(),
        ));
        Ok(())
    }

    fn visit_component(&mut self, component: ComponentSnapshot<'_>) -> Result<(), DecodeError> {
        self.components.push((
            component.identifier(),
            component.preferred_locator().to_owned(),
            component.locator().map(str::to_owned),
            component.has_unknown_fields(),
        ));
        Ok(())
    }

    fn visit_data_reference(
        &mut self,
        component: ComponentSnapshot<'_>,
        data_reference: ComponentDataReferenceSnapshot<'_>,
    ) -> Result<(), DecodeError> {
        self.references.push((
            component.identifier(),
            data_reference.data_identifier(),
            data_reference.owner_count(),
            data_reference.has_unknown_fields(),
        ));
        Ok(())
    }

    fn visit_owner(
        &mut self,
        component: ComponentSnapshot<'_>,
        data_reference: ComponentDataReferenceSnapshot<'_>,
        owner: OwnerSnapshot<'_>,
    ) -> Result<(), DecodeError> {
        self.owners.push((
            component.identifier(),
            data_reference.data_identifier(),
            owner.count(),
            owner.has_unknown_fields(),
        ));
        Ok(())
    }
}

#[test]
fn canonical_inspection_streams_borrowed_data_and_owner_facts() {
    let source = canonical_source();
    let mut facts = Facts::default();
    let report = visit_package_metadata_media(&source, options(&source), &mut facts).unwrap();

    assert_eq!(report.input_bytes(), source.len());
    assert_eq!(report.components(), 2);
    assert_eq!(report.data_records(), 2);
    assert_eq!(report.owners(), 3);
    assert_eq!(report.unknown_records(), 0);
    assert!(!report.data_metadata_map_present());
    assert_eq!(facts.components[0].0, 7);
    assert_eq!(facts.components[0].1, "Index/Document.iwa");
    assert_eq!(facts.components[0].2, None);
    assert_eq!(facts.components[1].2.as_deref(), Some("Index/Metadata.iwa"));
    assert_eq!(facts.references, [(7, 41, 2, false), (7, 42, 1, false)]);
    assert_eq!(
        facts.owners,
        [(7, 41, 1, false), (7, 41, 2, false), (7, 42, 1, false)]
    );
    assert_eq!(facts.data_infos[0].0, 41);
    assert_eq!(facts.data_infos[0].1, digest(0x10));
    assert_eq!(facts.data_infos[0].2, "intro.m4a");
    assert_eq!(facts.data_infos[0].3.as_deref(), Some("intro.m4a"));
    assert_eq!(facts.data_infos[0].4, Some(4_096));
}

#[test]
fn inspection_preserves_unknown_facts_without_materializing_them() {
    let source = source_with_unknowns();
    let mut facts = Facts::default();
    let report = visit_package_metadata_media(&source, options(&source), &mut facts).unwrap();

    assert!(report.unknown_records() >= 4);
    assert!(facts.components[0].3);
    assert!(facts.references[0].3);
    assert!(facts.owners[0].3);
    assert!(facts.data_infos[0].5);
}

#[test]
fn empty_batch_inspection_is_exact_and_does_not_allocate_output() {
    let source = canonical_source();
    let report = inspect_package_metadata_media(&source, options(&source)).unwrap();
    assert_eq!(report.input_bytes(), source.len());
    assert_eq!(report.data_records(), 2);
    assert!(MediaRewriteBatch::empty().is_empty());
}

#[test]
fn duplicate_data_info_and_component_and_owner_records_are_rejected() {
    let first_digest = digest(0x10);
    let first = data_info(41, &first_digest, "intro.m4a", None, None, false);
    let duplicate_data = package(200, &[], &[first.clone(), first], false);
    let error =
        inspect_package_metadata_media(&duplicate_data, options(&duplicate_data)).unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::DuplicateDataInfo)
    );

    let duplicate_component = package(
        200,
        &[
            component(7, "Index/Document.iwa", None, &[], false),
            component(7, "Index/Document.iwa", None, &[], false),
        ],
        &[],
        false,
    );
    let error = inspect_package_metadata_media(&duplicate_component, options(&duplicate_component))
        .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::DuplicateComponent)
    );

    let duplicate_owner = [(91, 1), (91, 2)];
    let duplicate_owner_source = package(
        200,
        &[component(
            7,
            "Index/Document.iwa",
            None,
            &[(41, &duplicate_owner, false)],
            false,
        )],
        &[data_info(41, &first_digest, "intro.m4a", None, None, false)],
        false,
    );
    let error =
        inspect_package_metadata_media(&duplicate_owner_source, options(&duplicate_owner_source))
            .unwrap_err();
    assert_eq!(error.invalid_reason(), Some(InvalidReason::DuplicateOwner));
}

#[test]
fn malformed_required_identifiers_digest_names_and_counts_fail_closed() {
    let malformed_cases = [
        (
            package(
                200,
                &[],
                &[data_info(0, &digest(0x10), "intro.m4a", None, None, false)],
                false,
            ),
            InvalidReason::InvalidIdentifier,
        ),
        (
            package(
                200,
                &[],
                &[data_info(41, &[1, 2], "intro.m4a", None, None, false)],
                false,
            ),
            InvalidReason::InvalidDigest,
        ),
        (
            package(
                200,
                &[],
                &[data_info(
                    41,
                    &digest(0x10),
                    "../intro.m4a",
                    None,
                    None,
                    false,
                )],
                false,
            ),
            InvalidReason::InvalidName,
        ),
        (
            package(
                200,
                &[component(
                    7,
                    "Index/Document.iwa",
                    None,
                    &[(41, &[(91, 0)], false)],
                    false,
                )],
                &[data_info(41, &digest(0x10), "intro.m4a", None, None, false)],
                false,
            ),
            InvalidReason::InvalidIdentifier,
        ),
    ];

    for (source, expected) in malformed_cases {
        let error = inspect_package_metadata_media(&source, options(&source)).unwrap_err();
        assert_eq!(error.invalid_reason(), Some(expected));
    }

    let empty_name = package(
        200,
        &[],
        &[data_info(41, &digest(0x10), "", None, None, false)],
        false,
    );
    let error = inspect_package_metadata_media(&empty_name, options(&empty_name)).unwrap_err();
    assert!(matches!(
        error.resource_limit(),
        Some(DecodeLimit::NameBytes { maximum: 4096, .. })
    ));
}

#[test]
fn missing_required_fields_and_malformed_wire_fail_before_callbacks() {
    let mut missing_data_identifier = Vec::new();
    field_bytes(&mut missing_data_identifier, 2, &digest(0x10));
    field_bytes(&mut missing_data_identifier, 3, b"intro.m4a");
    let source = package(200, &[], &[missing_data_identifier], false);
    let mut facts = Facts::default();
    let error = visit_package_metadata_media(&source, options(&source), &mut facts).unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::MissingRequiredField)
    );
    assert!(facts.data_infos.is_empty());

    let mut malformed = canonical_source();
    malformed[0] = 0x09;
    let error = inspect_package_metadata_media(&malformed, options(&malformed)).unwrap_err();
    assert_eq!(error.invalid_reason(), Some(InvalidReason::MalformedWire));
}

#[test]
fn every_structural_limit_is_reported_as_a_typed_refusal() {
    let source = canonical_source();
    let baseline = options(&source);
    let cases = [
        (
            DecodeOptions::new(
                source.len() - 1,
                baseline.max_fields(),
                baseline.max_work_bytes(),
                baseline.max_components(),
                baseline.max_data_records(),
                baseline.max_owners(),
                baseline.max_digest_bytes(),
                baseline.max_name_bytes(),
                baseline.max_depth(),
            ),
            DecodeLimit::Bytes {
                observed: source.len(),
                maximum: source.len() - 1,
            },
        ),
        (
            DecodeOptions::new(
                baseline.max_message_bytes(),
                1,
                baseline.max_work_bytes(),
                baseline.max_components(),
                baseline.max_data_records(),
                baseline.max_owners(),
                baseline.max_digest_bytes(),
                baseline.max_name_bytes(),
                baseline.max_depth(),
            ),
            DecodeLimit::Fields {
                observed: 0,
                maximum: 1,
            },
        ),
        (
            DecodeOptions::new(
                baseline.max_message_bytes(),
                baseline.max_fields(),
                1,
                baseline.max_components(),
                baseline.max_data_records(),
                baseline.max_owners(),
                baseline.max_digest_bytes(),
                baseline.max_name_bytes(),
                baseline.max_depth(),
            ),
            DecodeLimit::Work {
                observed: 0,
                maximum: 1,
            },
        ),
        (
            DecodeOptions::new(
                baseline.max_message_bytes(),
                baseline.max_fields(),
                baseline.max_work_bytes(),
                1,
                baseline.max_data_records(),
                baseline.max_owners(),
                baseline.max_digest_bytes(),
                baseline.max_name_bytes(),
                baseline.max_depth(),
            ),
            DecodeLimit::Components {
                observed: 0,
                maximum: 1,
            },
        ),
        (
            DecodeOptions::new(
                baseline.max_message_bytes(),
                baseline.max_fields(),
                baseline.max_work_bytes(),
                baseline.max_components(),
                1,
                baseline.max_owners(),
                baseline.max_digest_bytes(),
                baseline.max_name_bytes(),
                baseline.max_depth(),
            ),
            DecodeLimit::DataRecords {
                observed: 0,
                maximum: 1,
            },
        ),
        (
            DecodeOptions::new(
                baseline.max_message_bytes(),
                baseline.max_fields(),
                baseline.max_work_bytes(),
                baseline.max_components(),
                baseline.max_data_records(),
                2,
                baseline.max_digest_bytes(),
                baseline.max_name_bytes(),
                baseline.max_depth(),
            ),
            DecodeLimit::Owners {
                observed: 0,
                maximum: 2,
            },
        ),
        (
            DecodeOptions::new(
                baseline.max_message_bytes(),
                baseline.max_fields(),
                baseline.max_work_bytes(),
                baseline.max_components(),
                baseline.max_data_records(),
                baseline.max_owners(),
                19,
                baseline.max_name_bytes(),
                baseline.max_depth(),
            ),
            DecodeLimit::DigestBytes {
                observed: 0,
                maximum: 19,
            },
        ),
        (
            DecodeOptions::new(
                baseline.max_message_bytes(),
                baseline.max_fields(),
                baseline.max_work_bytes(),
                baseline.max_components(),
                baseline.max_data_records(),
                baseline.max_owners(),
                baseline.max_digest_bytes(),
                3,
                baseline.max_depth(),
            ),
            DecodeLimit::NameBytes {
                observed: 0,
                maximum: 3,
            },
        ),
        (
            DecodeOptions::new(
                baseline.max_message_bytes(),
                baseline.max_fields(),
                baseline.max_work_bytes(),
                baseline.max_components(),
                baseline.max_data_records(),
                baseline.max_owners(),
                baseline.max_digest_bytes(),
                baseline.max_name_bytes(),
                1,
            ),
            DecodeLimit::Nesting {
                observed: 0,
                maximum: 1,
            },
        ),
    ];

    for (limited, expected) in cases {
        let error = inspect_package_metadata_media(&source, limited).unwrap_err();
        let actual = error.resource_limit().expect("limit should be reported");
        assert_eq!(
            core::mem::discriminant(&actual),
            core::mem::discriminant(&expected)
        );
        match (actual, expected) {
            (
                DecodeLimit::Bytes { maximum: left, .. },
                DecodeLimit::Bytes { maximum: right, .. },
            )
            | (
                DecodeLimit::Fields { maximum: left, .. },
                DecodeLimit::Fields { maximum: right, .. },
            )
            | (DecodeLimit::Work { maximum: left, .. }, DecodeLimit::Work { maximum: right, .. })
            | (
                DecodeLimit::Components { maximum: left, .. },
                DecodeLimit::Components { maximum: right, .. },
            )
            | (
                DecodeLimit::DataRecords { maximum: left, .. },
                DecodeLimit::DataRecords { maximum: right, .. },
            )
            | (
                DecodeLimit::Owners { maximum: left, .. },
                DecodeLimit::Owners { maximum: right, .. },
            )
            | (
                DecodeLimit::DigestBytes { maximum: left, .. },
                DecodeLimit::DigestBytes { maximum: right, .. },
            )
            | (
                DecodeLimit::NameBytes { maximum: left, .. },
                DecodeLimit::NameBytes { maximum: right, .. },
            ) => assert_eq!(left, right),
            (
                DecodeLimit::Nesting { maximum: left, .. },
                DecodeLimit::Nesting { maximum: right, .. },
            ) => assert_eq!(left, right),
            _ => unreachable!("discriminants were checked above"),
        }
    }
}

#[test]
fn selectors_are_borrowed_and_constructor_values_are_strictly_typed() {
    let selector = ComponentSelector::new(7, "Index/Document.iwa");
    assert_eq!(selector.identifier(), 7);
    assert_eq!(selector.locator(), "Index/Document.iwa");

    let digest = digest(0x33);
    let addition = DataInfoAddition::new(50, &digest, "new.m4a")
        .with_file_name("new.m4a")
        .with_materialized_length(5);
    assert_eq!(addition.identifier(), 50);
    assert_eq!(addition.digest(), digest);
    assert_eq!(addition.preferred_file_name(), "new.m4a");
    assert_eq!(addition.file_name(), Some("new.m4a"));
    assert_eq!(addition.materialized_length(), Some(5));

    let removal = DataInfoRemoval::new(50);
    assert_eq!(removal.identifier(), 50);
}

fn rewrite_options(source: &[u8]) -> DecodeOptions {
    let baseline = options(source);
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

fn inspect_facts(source: &[u8]) -> Facts {
    let mut facts = Facts::default();
    visit_package_metadata_media(source, options(source), &mut facts)
        .expect("rewriter output should remain inspectable");
    facts
}

#[test]
fn data_info_addition_is_canonical_and_exact_removal_is_its_inverse() {
    let source = canonical_source();
    let new_digest = digest(0x44);
    let additions = [DataInfoAddition::new(50, &new_digest, "middle.m4a")
        .with_file_name("middle.m4a")
        .with_materialized_length(7_654)];
    let batch = MediaRewriteBatch::new(&additions, &[], &[], &[]);

    let output = rewrite_package_metadata_media(&source, batch, rewrite_options(&source)).unwrap();
    let report = output.report();
    assert_eq!(report.input_bytes(), source.len());
    assert_eq!(report.data_additions(), 1);
    assert_eq!(report.data_removals(), 0);
    assert_eq!(report.owner_additions(), 0);
    assert_eq!(report.owner_removals(), 0);
    assert_eq!(report.output_bytes(), output.bytes().len());
    assert!(output.bytes().starts_with(&source));

    let added = output.into_bytes();
    let facts = inspect_facts(&added);
    assert_eq!(facts.data_infos.len(), 3);
    assert_eq!(facts.data_infos[2].0, 50);
    assert_eq!(facts.data_infos[2].1, new_digest);
    assert_eq!(facts.data_infos[2].2, "middle.m4a");
    assert_eq!(facts.data_infos[2].3.as_deref(), Some("middle.m4a"));
    assert_eq!(facts.data_infos[2].4, Some(7_654));

    let removals = [DataInfoRemoval::new(50)];
    let inverse = MediaRewriteBatch::new(&[], &removals, &[], &[]);
    let restored =
        rewrite_package_metadata_media(&added, inverse, rewrite_options(&added)).unwrap();
    assert_eq!(restored.bytes(), source);
}

#[test]
fn owner_addition_inserts_a_distinct_owner_and_exact_removal_drops_it() {
    let source = canonical_source();
    let selector = ComponentSelector::new(7, "Index/Document.iwa");
    let additions = [DataReferenceOwnerAddition::new(selector, 41, 94, 3)];
    let batch = MediaRewriteBatch::new(&[], &[], &additions, &[]);

    let output = rewrite_package_metadata_media(&source, batch, rewrite_options(&source)).unwrap();
    let report = output.report();
    assert_eq!(report.owner_additions(), 1);
    let added = output.into_bytes();
    let facts = inspect_facts(&added);
    assert_eq!(facts.references, [(7, 41, 3, false), (7, 42, 1, false)]);
    assert_eq!(facts.owners.len(), 4);
    assert_eq!(facts.owners[2], (7, 41, 3, false));

    let removals = [DataReferenceOwnerRemoval::new(selector, 41, 94, 3)];
    let inverse = MediaRewriteBatch::new(&[], &[], &[], &removals);
    let restored =
        rewrite_package_metadata_media(&added, inverse, rewrite_options(&added)).unwrap();
    assert_eq!(restored.bytes(), source);
}

#[test]
fn owner_addition_creates_a_missing_parent_and_last_owner_removal_drops_parent() {
    let source = canonical_source();
    let selector = ComponentSelector::new(8, "Index/Metadata.iwa");
    let additions = [DataReferenceOwnerAddition::new(selector, 42, 99, 1)];
    let batch = MediaRewriteBatch::new(&[], &[], &additions, &[]);
    let output = rewrite_package_metadata_media(&source, batch, rewrite_options(&source)).unwrap();
    let added = output.into_bytes();
    let facts = inspect_facts(&added);
    assert!(facts.references.contains(&(8, 42, 1, false)));
    assert!(facts.owners.contains(&(8, 42, 1, false)));

    let removals = [DataReferenceOwnerRemoval::new(selector, 42, 99, 1)];
    let inverse = MediaRewriteBatch::new(&[], &[], &[], &removals);
    let removed = rewrite_package_metadata_media(&added, inverse, rewrite_options(&added)).unwrap();
    let facts = inspect_facts(removed.bytes());
    assert!(!facts.references.contains(&(8, 42, 0, false)));
    assert!(!facts.owners.contains(&(8, 42, 1, false)));
}

#[test]
fn owner_count_update_is_exact_and_byte_preserving_with_an_exact_inverse() {
    let source = canonical_source();
    let selector = ComponentSelector::new(7, "Index/Document.iwa");
    let updates = [DataReferenceOwnerCountUpdate::new(selector, 41, 92, 2, 4)];
    let batch = MediaRewriteBatch::new(&[], &[], &[], &[]).with_owner_updates(&updates);

    let output = rewrite_package_metadata_media(&source, batch, rewrite_options(&source)).unwrap();
    assert_eq!(output.report().owner_additions(), 0);
    assert_eq!(output.report().owner_removals(), 0);
    assert_eq!(output.report().owner_updates(), 1);
    let updated = output.into_bytes();
    let facts = inspect_facts(&updated);
    assert_eq!(
        facts.owners,
        [(7, 41, 1, false), (7, 41, 4, false), (7, 42, 1, false)]
    );

    let inverse_updates = [DataReferenceOwnerCountUpdate::new(selector, 41, 92, 4, 2)];
    let inverse =
        MediaRewriteBatch::new(&[], &[], &[], &[]).with_owner_count_updates(&inverse_updates);
    let restored =
        rewrite_package_metadata_media(&updated, inverse, rewrite_options(&updated)).unwrap();
    assert_eq!(restored.bytes(), source);
}

#[test]
fn owner_count_update_rejects_stale_zero_duplicate_conflicting_and_opaque_requests() {
    let source = canonical_source();
    let selector = ComponentSelector::new(7, "Index/Document.iwa");

    let zero_expected = [DataReferenceOwnerCountUpdate::new(selector, 41, 91, 0, 2)];
    let error = rewrite_package_metadata_media(
        &source,
        MediaRewriteBatch::new(&[], &[], &[], &[]).with_owner_updates(&zero_expected),
        rewrite_options(&source),
    )
    .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::InvalidIdentifier)
    );

    let zero_new = [DataReferenceOwnerCountUpdate::new(selector, 41, 91, 1, 0)];
    let error = rewrite_package_metadata_media(
        &source,
        MediaRewriteBatch::new(&[], &[], &[], &[]).with_owner_updates(&zero_new),
        rewrite_options(&source),
    )
    .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::InvalidIdentifier)
    );

    let stale = [DataReferenceOwnerCountUpdate::new(selector, 41, 91, 9, 2)];
    let error = rewrite_package_metadata_media(
        &source,
        MediaRewriteBatch::new(&[], &[], &[], &[]).with_owner_updates(&stale),
        rewrite_options(&source),
    )
    .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::ExistingOwnerCollision)
    );

    let duplicate = [
        DataReferenceOwnerCountUpdate::new(selector, 41, 91, 1, 2),
        DataReferenceOwnerCountUpdate::new(selector, 41, 91, 1, 3),
    ];
    let error = rewrite_package_metadata_media(
        &source,
        MediaRewriteBatch::new(&[], &[], &[], &[]).with_owner_updates(&duplicate),
        rewrite_options(&source),
    )
    .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::ConflictingOperation)
    );

    let additions = [DataReferenceOwnerAddition::new(selector, 41, 91, 1)];
    let conflicts_with_addition = [DataReferenceOwnerCountUpdate::new(selector, 41, 91, 1, 2)];
    let error = rewrite_package_metadata_media(
        &source,
        MediaRewriteBatch::new(&[], &[], &additions, &[])
            .with_owner_updates(&conflicts_with_addition),
        rewrite_options(&source),
    )
    .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::DuplicateOperation)
    );

    let removals = [DataReferenceOwnerRemoval::new(selector, 41, 91, 1)];
    let conflicts_with_removal = [DataReferenceOwnerCountUpdate::new(selector, 41, 91, 1, 2)];
    let error = rewrite_package_metadata_media(
        &source,
        MediaRewriteBatch::new(&[], &[], &[], &removals)
            .with_owner_updates(&conflicts_with_removal),
        rewrite_options(&source),
    )
    .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::DuplicateOperation)
    );

    let opaque = source_with_unknowns();
    let opaque_updates = [DataReferenceOwnerCountUpdate::new(selector, 41, 91, 1, 2)];
    let error = rewrite_package_metadata_media(
        &opaque,
        MediaRewriteBatch::new(&[], &[], &[], &[]).with_owner_updates(&opaque_updates),
        rewrite_options(&opaque),
    )
    .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::UnknownSelectedRecord)
    );
}

#[test]
fn untouched_unknown_records_are_byte_preserved_while_selected_unknowns_refuse_mutation() {
    let source = source_with_unknowns();
    let new_digest = digest(0xa0);
    let additions = [DataInfoAddition::new(50, &new_digest, "new.m4a")];
    let add_batch = MediaRewriteBatch::new(&additions, &[], &[], &[]);
    let output =
        rewrite_package_metadata_media(&source, add_batch, rewrite_options(&source)).unwrap();
    assert!(output.bytes().starts_with(&source));
    assert!(inspect_facts(output.bytes()).data_infos[0].5);

    let selector = ComponentSelector::new(7, "Index/Document.iwa");
    let owner_additions = [DataReferenceOwnerAddition::new(selector, 41, 94, 1)];
    let error = rewrite_package_metadata_media(
        &source,
        MediaRewriteBatch::new(&[], &[], &owner_additions, &[]),
        rewrite_options(&source),
    )
    .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::UnknownSelectedRecord)
    );

    let removals = [DataInfoRemoval::new(41)];
    let error = rewrite_package_metadata_media(
        &source,
        MediaRewriteBatch::new(&[], &removals, &[], &[]),
        rewrite_options(&source),
    )
    .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::UnknownSelectedRecord)
    );

    let owner_removals = [DataReferenceOwnerRemoval::new(selector, 41, 91, 1)];
    let error = rewrite_package_metadata_media(
        &source,
        MediaRewriteBatch::new(&[], &[], &[], &owner_removals),
        rewrite_options(&source),
    )
    .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::UnknownSelectedRecord)
    );
}

#[test]
fn rewrite_rejects_duplicate_operations_collisions_and_reference_violations_atomically() {
    let source = canonical_source();
    let source_before = source.clone();
    let first_digest = digest(0x31);
    let additions = [
        DataInfoAddition::new(50, &first_digest, "one.m4a"),
        DataInfoAddition::new(50, &first_digest, "two.m4a"),
    ];
    let error = rewrite_package_metadata_media(
        &source,
        MediaRewriteBatch::new(&additions, &[], &[], &[]),
        rewrite_options(&source),
    )
    .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::DuplicateOperation)
    );
    assert_eq!(source, source_before);

    let collision = [DataInfoAddition::new(41, &first_digest, "collision.m4a")];
    let error = rewrite_package_metadata_media(
        &source,
        MediaRewriteBatch::new(&collision, &[], &[], &[]),
        rewrite_options(&source),
    )
    .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::ExistingDataCollision)
    );

    let referenced = [DataInfoRemoval::new(41)];
    let error = rewrite_package_metadata_media(
        &source,
        MediaRewriteBatch::new(&[], &referenced, &[], &[]),
        rewrite_options(&source),
    )
    .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::DataInfoReferenced)
    );

    let selector = ComponentSelector::new(7, "Index/Document.iwa");
    let existing_owner = [DataReferenceOwnerAddition::new(selector, 41, 91, 1)];
    let error = rewrite_package_metadata_media(
        &source,
        MediaRewriteBatch::new(&[], &[], &existing_owner, &[]),
        rewrite_options(&source),
    )
    .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::ExistingOwnerCollision)
    );

    let mismatched = [DataReferenceOwnerRemoval::new(selector, 41, 91, 7)];
    let error = rewrite_package_metadata_media(
        &source,
        MediaRewriteBatch::new(&[], &[], &[], &mismatched),
        rewrite_options(&source),
    )
    .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::ExistingOwnerCollision)
    );

    let missing_component = [DataReferenceOwnerRemoval::new(
        ComponentSelector::new(77, "Index/Missing.iwa"),
        41,
        91,
        1,
    )];
    let error = rewrite_package_metadata_media(
        &source,
        MediaRewriteBatch::new(&[], &[], &[], &missing_component),
        rewrite_options(&source),
    )
    .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::ComponentNotFound)
    );

    let missing_data = [DataReferenceOwnerAddition::new(selector, 999, 94, 1)];
    let error = rewrite_package_metadata_media(
        &source,
        MediaRewriteBatch::new(&[], &[], &missing_data, &[]),
        rewrite_options(&source),
    )
    .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::DataInfoNotFound)
    );
}

#[test]
fn malformed_additions_and_source_duplicates_fail_closed_before_publication() {
    let source = canonical_source();
    let invalid_digest_value = digest(0x31);
    let invalid_id = [DataInfoAddition::new(0, &invalid_digest_value, "new.m4a")];
    let error = rewrite_package_metadata_media(
        &source,
        MediaRewriteBatch::new(&invalid_id, &[], &[], &[]),
        rewrite_options(&source),
    )
    .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::InvalidIdentifier)
    );

    let invalid_digest = [DataInfoAddition::new(50, &[1, 2], "new.m4a")];
    let error = rewrite_package_metadata_media(
        &source,
        MediaRewriteBatch::new(&invalid_digest, &[], &[], &[]),
        rewrite_options(&source),
    )
    .unwrap_err();
    assert_eq!(error.invalid_reason(), Some(InvalidReason::InvalidDigest));

    let invalid_path = [DataInfoAddition::new(
        50,
        &invalid_digest_value,
        "../new.m4a",
    )];
    let error = rewrite_package_metadata_media(
        &source,
        MediaRewriteBatch::new(&invalid_path, &[], &[], &[]),
        rewrite_options(&source),
    )
    .unwrap_err();
    assert_eq!(error.invalid_reason(), Some(InvalidReason::InvalidName));

    let zero_count = [DataReferenceOwnerAddition::new(
        ComponentSelector::new(7, "Index/Document.iwa"),
        41,
        95,
        0,
    )];
    let error = rewrite_package_metadata_media(
        &source,
        MediaRewriteBatch::new(&[], &[], &zero_count, &[]),
        rewrite_options(&source),
    )
    .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::InvalidIdentifier)
    );

    let mut duplicate = canonical_source();
    let duplicate_record = data_info(41, &digest(0x10), "intro.m4a", None, None, false);
    field_bytes(&mut duplicate, 4, &duplicate_record);
    let error = rewrite_package_metadata_media(
        &duplicate,
        MediaRewriteBatch::empty(),
        rewrite_options(&duplicate),
    )
    .unwrap_err();
    assert_eq!(
        error.invalid_reason(),
        Some(InvalidReason::DuplicateDataInfo)
    );
}

#[test]
fn prepare_execute_charges_exact_resources_and_refuses_tight_limits_without_output() {
    let source = canonical_source();
    let new_digest = digest(0xd0);
    let additions = [DataInfoAddition::new(50, &new_digest, "new.m4a")];
    let batch = MediaRewriteBatch::new(&additions, &[], &[], &[]);
    let prepared = prepare_package_metadata_media_rewrite(&source, batch, rewrite_options(&source))
        .expect("valid rewrite should prepare");
    let requirements = prepared.execution_requirements();
    assert_eq!(prepared.output_size(), requirements.output_bytes());
    assert!(requirements.output_bytes() > source.len());
    assert!(requirements.fields() > 0);
    assert!(requirements.work_bytes() > 0);
    assert_eq!(requirements.allocations(), 1);
    assert_eq!(requirements.retained_bytes(), requirements.output_bytes());
    assert_eq!(requirements.scratch_bytes(), 0);

    let exact = requirements.exact_limits();
    let output = prepared
        .execute(exact)
        .expect("exact limits should execute");
    assert_eq!(output.bytes().len(), requirements.output_bytes());
    assert_eq!(output.report().allocations(), 1);

    let prepared = prepare_package_metadata_media_rewrite(&source, batch, rewrite_options(&source))
        .expect("valid rewrite should prepare again");
    let mut too_small = exact;
    too_small.max_output_bytes = requirements.output_bytes().saturating_sub(1);
    let error = prepared.execute(too_small).unwrap_err();
    assert!(matches!(
        error.resource_limit(),
        Some(DecodeLimit::OutputBytes { .. })
    ));

    let prepared = prepare_package_metadata_media_rewrite(&source, batch, rewrite_options(&source))
        .expect("valid rewrite should prepare again");
    let mut too_few_fields = exact;
    too_few_fields.max_fields = requirements.fields().saturating_sub(1);
    let error = prepared.execute(too_few_fields).unwrap_err();
    assert!(matches!(
        error.resource_limit(),
        Some(DecodeLimit::Fields { .. })
    ));

    let prepared = prepare_package_metadata_media_rewrite(&source, batch, rewrite_options(&source))
        .expect("valid rewrite should prepare again");
    let mut too_little_work = exact;
    too_little_work.max_work_bytes = requirements.work_bytes().saturating_sub(1);
    let error = prepared.execute(too_little_work).unwrap_err();
    assert!(matches!(
        error.resource_limit(),
        Some(DecodeLimit::Work { .. })
    ));

    let mut output_limited_options = rewrite_options(&source);
    output_limited_options = output_limited_options.with_max_output_bytes(source.len());
    let error = prepare_package_metadata_media_rewrite(&source, batch, output_limited_options)
        .expect_err("output limit should refuse preparation");
    assert!(matches!(
        error.resource_limit(),
        Some(DecodeLimit::OutputBytes { .. })
    ));
}

#[test]
fn execution_allocation_and_retention_limits_fail_before_candidate_allocation() {
    let source = canonical_source();
    let digest = digest(0xe0);
    let additions = [DataInfoAddition::new(50, &digest, "new.m4a")];
    let batch = MediaRewriteBatch::new(&additions, &[], &[], &[]);
    let prepared = prepare_package_metadata_media_rewrite(&source, batch, rewrite_options(&source))
        .expect("valid rewrite should prepare");
    let requirements = prepared.execution_requirements();
    let mut no_allocations = requirements.exact_limits();
    no_allocations.max_allocations = 0;
    let error = prepared.execute(no_allocations).unwrap_err();
    assert_eq!(error.invalid_reason(), Some(InvalidReason::Verification));

    let prepared = prepare_package_metadata_media_rewrite(&source, batch, rewrite_options(&source))
        .expect("valid rewrite should prepare again");
    let mut no_retention = requirements.exact_limits();
    no_retention.max_retained_bytes = requirements.retained_bytes().saturating_sub(1);
    let error = prepared.execute(no_retention).unwrap_err();
    assert_eq!(error.invalid_reason(), Some(InvalidReason::Verification));
}
