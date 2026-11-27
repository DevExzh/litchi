#![no_main]

//! Bounded fuzzing for the PackageMetadata media-metadata seam.
//!
//! The target keeps the source-owned representation in control: malformed
//! inputs are accepted as ordinary fuzz cases, while every successful scan is
//! replayed through both the report-only and streaming visitor entry points.
//! The visitor stores only counters, so a wide package cannot turn fuzzing
//! into an input-sized `Vec<Vec<u8>>` allocation.  A small valid seed also
//! drives the complete source-preserving mutation chain, including prepared
//! and one-shot execution under exact finite resource limits.

use std::{hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::package_metadata_media_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 16 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_COMPONENTS: usize = 4 * 1024;
const MAX_DATA_RECORDS: usize = 8 * 1024;
const MAX_OWNERS: usize = 16 * 1024;
const MAX_DIGEST_BYTES: usize = 20;
const MAX_NAME_BYTES: usize = 4096;
const MAX_DEPTH: u32 = 64;

// A valid empty root, a complete one-owner closure, and malformed/ambiguous
// records are kept hot even when a campaign starts with arbitrary bytes.
const FIXED_CASES: &[&[u8]] = &[
    &[],
    &[0x08, 0x0a],
    &[
        0x08, 0x0a, 0x1a, 0x20, 0x08, 0x07, 0x12, 0x12, 0x49, 0x6e, 0x64, 0x65, 0x78, 0x2f, 0x44,
        0x6f, 0x63, 0x75, 0x6d, 0x65, 0x6e, 0x74, 0x2e, 0x69, 0x77, 0x61, 0x3a, 0x08, 0x08, 0x0b,
        0x12, 0x04, 0x08, 0x2a, 0x10, 0x01, 0x22, 0x24, 0x08, 0x0b, 0x12, 0x14, 0x30, 0x31, 0x32,
        0x33, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x30, 0x31, 0x32, 0x33, 0x34, 0x35, 0x36, 0x37,
        0x38, 0x39, 0x1a, 0x0a, 0x61, 0x75, 0x64, 0x69, 0x6f, 0x2e, 0x61, 0x69, 0x66, 0x66,
    ],
    // Two references share one owner and a second data item has a materialized
    // filename/length.  This keeps owner multiplicity and optional DataInfo
    // fields reachable without embedding package bytes.
    &[
        0x08, 0x0a, 0x1a, 0x30, 0x08, 0x07, 0x12, 0x12, 0x49, 0x6e, 0x64, 0x65, 0x78, 0x2f, 0x44,
        0x6f, 0x63, 0x75, 0x6d, 0x65, 0x6e, 0x74, 0x2e, 0x69, 0x77, 0x61, 0x3a, 0x0e, 0x08, 0x0b,
        0x12, 0x04, 0x08, 0x2a, 0x10, 0x02, 0x12, 0x04, 0x08, 0x2b, 0x10, 0x01, 0x3a, 0x08, 0x08,
        0x0c, 0x12, 0x04, 0x08, 0x2a, 0x10, 0x01, 0x22, 0x24, 0x08, 0x0b, 0x12, 0x14, 0x30, 0x31,
        0x32, 0x33, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x30, 0x31, 0x32, 0x33, 0x34, 0x35, 0x36,
        0x37, 0x38, 0x39, 0x1a, 0x0a, 0x61, 0x75, 0x64, 0x69, 0x6f, 0x2e, 0x61, 0x69, 0x66, 0x66,
        0x22, 0x31, 0x08, 0x0c, 0x12, 0x14, 0x61, 0x62, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68, 0x69,
        0x6a, 0x6b, 0x6c, 0x6d, 0x6e, 0x6f, 0x70, 0x71, 0x72, 0x73, 0x74, 0x1a, 0x09, 0x66, 0x69,
        0x72, 0x73, 0x74, 0x2e, 0x77, 0x61, 0x76, 0x22, 0x09, 0x66, 0x69, 0x72, 0x73, 0x74, 0x2e,
        0x77, 0x61, 0x76, 0x90, 0x01, 0x7b,
    ],
    // Unknown root/data fields must be preserved or rejected according to the
    // selected operation; the scan itself must remain atomic and source-free.
    &[0x08, 0x0a, 0xa0, 0x06, 0x81, 0x00],
    &[
        0x08, 0x0a, 0x22, 0x28, 0x08, 0x0b, 0x12, 0x14, 0x30, 0x31, 0x32, 0x33, 0x34, 0x35, 0x36,
        0x37, 0x38, 0x39, 0x30, 0x31, 0x32, 0x33, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x1a, 0x0a,
        0x61, 0x75, 0x64, 0x69, 0x6f, 0x2e, 0x61, 0x69, 0x66, 0x66, 0xa0, 0x06, 0x81, 0x00,
    ],
    // Duplicate DataInfo identity and duplicate owner selectors are strict
    // refusals, not permissive last-write-wins records.
    &[
        0x08, 0x0a, 0x22, 0x24, 0x08, 0x0b, 0x12, 0x14, 0x30, 0x31, 0x32, 0x33, 0x34, 0x35, 0x36,
        0x37, 0x38, 0x39, 0x30, 0x31, 0x32, 0x33, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x1a, 0x0a,
        0x61, 0x75, 0x64, 0x69, 0x6f, 0x2e, 0x61, 0x69, 0x66, 0x66, 0x22, 0x24, 0x08, 0x0b, 0x12,
        0x14, 0x30, 0x31, 0x32, 0x33, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x30, 0x31, 0x32, 0x33,
        0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x1a, 0x0a, 0x61, 0x75, 0x64, 0x69, 0x6f, 0x2e, 0x61,
        0x69, 0x66, 0x66,
    ],
    &[
        0x08, 0x0a, 0x1a, 0x26, 0x08, 0x07, 0x12, 0x12, 0x49, 0x6e, 0x64, 0x65, 0x78, 0x2f, 0x44,
        0x6f, 0x63, 0x75, 0x6d, 0x65, 0x6e, 0x74, 0x2e, 0x69, 0x77, 0x61, 0x3a, 0x0e, 0x08, 0x0b,
        0x12, 0x04, 0x08, 0x2a, 0x10, 0x01, 0x12, 0x04, 0x08, 0x2a, 0x10, 0x01,
    ],
    &[0x08, 0x01, 0x22, 0x80],
    &[0x08, 0x01, 0x12, 0x01, 0x00],
    &[0x88, 0x00, 0x01],
];

#[derive(Default)]
struct Facts {
    data_records: usize,
    components: usize,
    references: usize,
    owners: usize,
    unknown_data: usize,
}

impl codec::PackageMetadataMediaVisitor for Facts {
    fn visit_data_info(
        &mut self,
        data_info: codec::DataInfoSnapshot<'_>,
    ) -> Result<(), codec::DecodeError> {
        self.data_records = self.data_records.saturating_add(1);
        black_box((
            data_info.identifier(),
            data_info.digest(),
            data_info.preferred_file_name(),
            data_info.file_name(),
            data_info.materialized_length(),
            data_info.raw(),
            data_info.has_unknown_fields(),
        ));
        self.unknown_data += usize::from(data_info.has_unknown_fields());
        Ok(())
    }

    fn visit_component(
        &mut self,
        component: codec::ComponentSnapshot<'_>,
    ) -> Result<(), codec::DecodeError> {
        self.components = self.components.saturating_add(1);
        black_box((
            component.identifier(),
            component.preferred_locator(),
            component.locator(),
            component.effective_locator(),
            component.is_versioned(),
            component.has_unknown_fields(),
        ));
        Ok(())
    }

    fn visit_data_reference(
        &mut self,
        component: codec::ComponentSnapshot<'_>,
        data_reference: codec::ComponentDataReferenceSnapshot<'_>,
    ) -> Result<(), codec::DecodeError> {
        self.references = self.references.saturating_add(1);
        black_box((
            component.identifier(),
            data_reference.data_identifier(),
            data_reference.owner_count(),
            data_reference.raw(),
            data_reference.has_unknown_fields(),
        ));
        Ok(())
    }

    fn visit_owner(
        &mut self,
        component: codec::ComponentSnapshot<'_>,
        data_reference: codec::ComponentDataReferenceSnapshot<'_>,
        owner: codec::OwnerSnapshot<'_>,
    ) -> Result<(), codec::DecodeError> {
        self.owners = self.owners.saturating_add(1);
        black_box((
            component.identifier(),
            data_reference.data_identifier(),
            owner.object_identifier(),
            owner.count(),
            owner.raw(),
            owner.has_unknown_fields(),
        ));
        Ok(())
    }
}

fuzz_target!(|data: &[u8]| {
    if let Some(source) = normalize_input(data) {
        exercise_source(&source);
        exercise_map_source(&source);
    }

    static FIXED: OnceLock<()> = OnceLock::new();
    FIXED.get_or_init(|| {
        for source in FIXED_CASES {
            exercise_source(source);
        }
        exercise_transition_chain(FIXED_CASES[2]);
    });
});

#[derive(Default)]
struct MapFacts {
    entries: usize,
    first_identifier: Option<u64>,
}

impl codec::DataMetadataMapVisitor for MapFacts {
    fn visit_entry(
        &mut self,
        entry: codec::DataMetadataMapEntry,
    ) -> Result<(), codec::DecodeError> {
        self.entries += 1;
        self.first_identifier.get_or_insert(entry.data_identifier());
        black_box((
            entry.metadata_object_identifier(),
            entry.has_unknown_fields(),
        ));
        Ok(())
    }
}

fn exercise_map_source(source: &[u8]) {
    let finite = options(source);
    match codec::DataMetadataMapSource::from_source(MAP_OBJECT_IDENTIFIER, source, finite) {
        Ok(witness) => {
            assert_eq!(witness.payload(), source);
            let mut facts = MapFacts::default();
            witness
                .visit_entries(finite, &mut facts)
                .expect("validated map must remain visitable under the same limits");
            assert_eq!(facts.entries, witness.entries());
            let identifier = facts.first_identifier.unwrap_or(1);
            assert_eq!(
                witness
                    .contains_data_identifier(identifier, finite)
                    .expect("validated map lookup must obey the same finite profile"),
                facts.first_identifier.is_some()
            );
        },
        Err(error) => {
            black_box((error.resource_limit(), error.invalid_reason()));
        },
    }
}

fn normalize_input(data: &[u8]) -> Option<Vec<u8>> {
    if let Some(encoded) = data.strip_prefix(b"hex:") {
        return decode_hex(encoded);
    }
    (data.len() <= MAX_INPUT_BYTES).then(|| data.to_vec())
}

fn decode_hex(encoded: &[u8]) -> Option<Vec<u8>> {
    if encoded.len() > MAX_INPUT_BYTES.saturating_mul(2).saturating_add(16) {
        return None;
    }
    let mut output = Vec::with_capacity(encoded.len() / 2);
    let mut high = None;
    for byte in encoded.iter().copied() {
        if byte.is_ascii_whitespace() {
            continue;
        }
        let nibble = hex_nibble(byte)?;
        if let Some(high_nibble) = high.take() {
            output.push((high_nibble << 4) | nibble);
            if output.len() > MAX_INPUT_BYTES {
                return None;
            }
        } else {
            high = Some(nibble);
        }
    }
    high.is_none().then_some(output)
}

fn hex_nibble(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn append_varint(output: &mut Vec<u8>, mut value: u64) {
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

fn append_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    append_varint(output, u64::from(number) << 3);
    append_varint(output, value);
}

fn append_bytes_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) {
    append_varint(output, (u64::from(number) << 3) | 2);
    append_varint(output, payload.len() as u64);
    output.extend_from_slice(payload);
}

fn options(source: &[u8]) -> codec::DecodeOptions {
    codec::DecodeOptions::new(
        MAX_INPUT_BYTES.max(source.len()),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_COMPONENTS,
        MAX_DATA_RECORDS,
        MAX_OWNERS,
        MAX_DIGEST_BYTES,
        MAX_NAME_BYTES,
        MAX_DEPTH,
    )
    .with_max_output_bytes(MAX_OUTPUT_BYTES.max(source.len()))
}

fn exercise_source(source: &[u8]) {
    let original = source.to_vec();
    let decode_options = options(source);
    let inspected = codec::inspect_package_metadata_media(source, decode_options);
    assert_eq!(
        source,
        original.as_slice(),
        "inspection modified its source"
    );

    match inspected {
        Ok(report) => {
            assert_eq!(report.input_bytes(), source.len());
            assert!(report.fields() <= MAX_FIELDS);
            assert!(report.work_bytes() <= MAX_WORK_BYTES);
            assert!(report.max_depth() <= MAX_DEPTH);
            assert!(report.components() <= MAX_COMPONENTS);
            assert!(report.data_records() <= MAX_DATA_RECORDS);
            assert!(report.owners() <= MAX_OWNERS);

            let mut facts = Facts::default();
            let visited = codec::visit_package_metadata_media(source, decode_options, &mut facts)
                .expect("report-only and streaming scans must agree");
            assert_eq!(visited, report);
            assert_eq!(facts.components, report.components());
            assert_eq!(facts.data_records, report.data_records());
            assert_eq!(facts.owners, report.owners());
            // DecodeReport intentionally exposes only bounded owner/data
            // totals. Keep the streamed reference count opaque while still
            // making the visitor work observable to the optimizer.
            black_box((facts, report.data_metadata_map_present()));
        },
        Err(error) => {
            black_box((error.resource_limit(), error.invalid_reason()));
        },
    }
    assert_eq!(source, original.as_slice(), "scan modified its source");
}

const TRANSITION_COMPONENT: codec::ComponentSelector<'static> =
    codec::ComponentSelector::new(7, "Index/Document.iwa");
const TRANSITION_DATA_IDENTIFIER: u64 = 11;
const TRANSITION_OWNER_IDENTIFIER: u64 = 42;
const ADDED_DATA_IDENTIFIER: u64 = 43;
const ADDED_OWNER_IDENTIFIER: u64 = 94;
const ADDED_DATA_DIGEST: &[u8; 20] = b"01234567890123456789";
const MAP_OBJECT_IDENTIFIER: u64 = 80;
const MAP_METADATA_OBJECT_IDENTIFIER: u64 = 81;
const SURVIVING_OWNER_IDENTIFIER: u64 = 43;

#[derive(Clone, Copy)]
enum RewriteLimit {
    OutputBytes,
    Fields,
    WorkBytes,
    Components,
    DataRecords,
    Owners,
    Allocations,
    RetainedBytes,
    ScratchBytes,
}

fn exercise_transition_chain(source: &[u8]) {
    let empty_batch = codec::MediaRewriteBatch::empty();
    let _ = run_rewrite(source, empty_batch, options(source));

    let data_addition =
        codec::DataInfoAddition::new(ADDED_DATA_IDENTIFIER, ADDED_DATA_DIGEST, "fuzz-added.m4a")
            .with_file_name("fuzz-added.m4a")
            .with_materialized_length(128);
    let data_additions = [data_addition];
    let data_batch = codec::MediaRewriteBatch::new(&data_additions, &[], &[], &[]);
    let data_added = run_rewrite(source, data_batch, options(source))
        .expect("the complete valid seed must accept DataInfo addition");

    let owner_addition = codec::DataReferenceOwnerAddition::new(
        TRANSITION_COMPONENT,
        ADDED_DATA_IDENTIFIER,
        ADDED_OWNER_IDENTIFIER,
        2,
    );
    let owner_additions = [owner_addition];
    let owner_batch = codec::MediaRewriteBatch::new(&[], &[], &owner_additions, &[]);
    let owner_added = run_rewrite(&data_added, owner_batch, options(&data_added))
        .expect("the complete valid seed must accept owner addition");

    let owner_removal = codec::DataReferenceOwnerRemoval::new(
        TRANSITION_COMPONENT,
        ADDED_DATA_IDENTIFIER,
        ADDED_OWNER_IDENTIFIER,
        2,
    );
    let owner_removals = [owner_removal];
    let owner_remove_batch = codec::MediaRewriteBatch::new(&[], &[], &[], &owner_removals);
    let owner_removed = run_rewrite(&owner_added, owner_remove_batch, options(&owner_added))
        .expect("the complete valid seed must accept owner removal");

    let data_removal = codec::DataInfoRemoval::new(ADDED_DATA_IDENTIFIER);
    let data_removals = [data_removal];
    let data_remove_batch = codec::MediaRewriteBatch::new(&[], &data_removals, &[], &[]);
    let data_removed = run_rewrite(&owner_removed, data_remove_batch, options(&owner_removed))
        .expect("the complete valid seed must accept DataInfo removal");

    let existing_owner_removal = codec::DataReferenceOwnerRemoval::new(
        TRANSITION_COMPONENT,
        TRANSITION_DATA_IDENTIFIER,
        TRANSITION_OWNER_IDENTIFIER,
        1,
    );
    let existing_owner_removals = [existing_owner_removal];
    let existing_owner_batch =
        codec::MediaRewriteBatch::new(&[], &[], &[], &existing_owner_removals);
    let existing_owner_removed = run_rewrite(source, existing_owner_batch, options(source))
        .expect("the complete valid seed must accept existing-owner removal");

    let existing_data_removal = codec::DataInfoRemoval::new(TRANSITION_DATA_IDENTIFIER);
    let existing_data_removals = [existing_data_removal];
    let existing_data_batch = codec::MediaRewriteBatch::new(&[], &existing_data_removals, &[], &[]);
    let existing_data_removed = run_rewrite(
        &existing_owner_removed,
        existing_data_batch,
        options(&existing_owner_removed),
    )
    .expect("an unreferenced existing DataInfo must be removable");

    let final_report = codec::inspect_package_metadata_media(&data_removed, options(&data_removed))
        .expect("the transition chain must produce a valid package");
    assert_eq!(final_report.data_records(), 1);
    let final_existing_report = codec::inspect_package_metadata_media(
        &existing_data_removed,
        options(&existing_data_removed),
    )
    .expect("the existing-owner transition must produce a valid package");
    assert_eq!(final_existing_report.data_records(), 0);
    black_box((data_removed, existing_data_removed));

    exercise_data_metadata_map_transactions();
}

fn source_with_data_metadata_map(extra_owner: bool) -> Vec<u8> {
    let mut owner = Vec::new();
    append_varint_field(&mut owner, 1, TRANSITION_OWNER_IDENTIFIER);
    append_varint_field(&mut owner, 2, 1);

    let mut reference = Vec::new();
    append_varint_field(&mut reference, 1, TRANSITION_DATA_IDENTIFIER);
    append_bytes_field(&mut reference, 2, &owner);
    if extra_owner {
        let mut surviving_owner = Vec::new();
        append_varint_field(&mut surviving_owner, 1, SURVIVING_OWNER_IDENTIFIER);
        append_varint_field(&mut surviving_owner, 2, 1);
        append_bytes_field(&mut reference, 2, &surviving_owner);
    }

    let mut component = Vec::new();
    append_varint_field(&mut component, 1, TRANSITION_COMPONENT.identifier());
    append_bytes_field(&mut component, 2, TRANSITION_COMPONENT.locator().as_bytes());
    append_bytes_field(&mut component, 7, &reference);

    let mut data_info = Vec::new();
    append_varint_field(&mut data_info, 1, TRANSITION_DATA_IDENTIFIER);
    append_bytes_field(&mut data_info, 2, ADDED_DATA_DIGEST);
    append_bytes_field(&mut data_info, 3, b"audio.aiff");

    let mut root_map_reference = Vec::new();
    append_varint_field(&mut root_map_reference, 1, MAP_OBJECT_IDENTIFIER);

    let mut source = Vec::new();
    append_varint_field(&mut source, 1, 10);
    append_bytes_field(&mut source, 3, &component);
    append_bytes_field(&mut source, 4, &data_info);
    append_bytes_field(&mut source, 10, &root_map_reference);
    source
}

fn data_metadata_map_payload(
    data_identifier: u64,
    metadata_object_identifier: u64,
    unknown_reference_field: bool,
) -> Vec<u8> {
    let mut metadata_reference = Vec::new();
    append_varint_field(&mut metadata_reference, 1, metadata_object_identifier);
    if unknown_reference_field {
        append_varint_field(&mut metadata_reference, 99, 1);
    }

    let mut entry = Vec::new();
    append_varint_field(&mut entry, 1, data_identifier);
    append_bytes_field(&mut entry, 2, &metadata_reference);

    let mut payload = Vec::new();
    append_bytes_field(&mut payload, 1, &entry);
    payload
}

fn map_witness<'source>(
    object_identifier: u64,
    payload: &'source [u8],
) -> Result<codec::DataMetadataMapSource<'source>, codec::DecodeError> {
    codec::DataMetadataMapSource::from_source(object_identifier, payload, options(payload))
}

fn final_owner_batch<'source>(
    map_payload: &'source [u8],
    data_removals: &'source [codec::DataInfoRemoval],
    owner_removals: &'source [codec::DataReferenceOwnerRemoval<'source>],
) -> Result<codec::MediaRewriteBatch<'source>, codec::DecodeError> {
    Ok(
        codec::MediaRewriteBatch::new(&[], data_removals, &[], owner_removals)
            .with_data_metadata_map_source(map_witness(MAP_OBJECT_IDENTIFIER, map_payload)?),
    )
}

fn assert_refused<'source>(
    source: &'source [u8],
    batch: codec::MediaRewriteBatch<'source>,
    label: &str,
) {
    let original = source.to_vec();
    assert!(
        run_rewrite(source, batch, options(source)).is_none(),
        "{label} must refuse before publication"
    );
    assert_eq!(source, original.as_slice(), "{label} modified its source");
}

fn exercise_data_metadata_map_transactions() {
    let source = source_with_data_metadata_map(false);
    let original = source.clone();
    let owner_removal = codec::DataReferenceOwnerRemoval::new(
        TRANSITION_COMPONENT,
        TRANSITION_DATA_IDENTIFIER,
        TRANSITION_OWNER_IDENTIFIER,
        1,
    );
    let owner_removals = [owner_removal];
    let data_removal = codec::DataInfoRemoval::new(TRANSITION_DATA_IDENTIFIER);
    let data_removals = [data_removal];
    let absent_map = Vec::new();
    let valid_batch = final_owner_batch(absent_map.as_slice(), &data_removals, &owner_removals)
        .expect("an empty DataMetadataMap witness must remain valid");
    let output = run_rewrite(&source, valid_batch, options(&source))
        .expect("an absent map key must admit the atomic final-owner/DataInfo removal");
    let output_report = codec::inspect_package_metadata_media(&output, options(&output))
        .expect("the final-owner candidate must remain inspectable");
    assert_eq!(output_report.data_records(), 0);
    assert_eq!(output_report.owners(), 0);
    assert!(output_report.data_metadata_map_present());
    assert_eq!(
        source,
        original.as_slice(),
        "valid rewrite modified its source"
    );

    let present_map = data_metadata_map_payload(
        TRANSITION_DATA_IDENTIFIER,
        MAP_METADATA_OBJECT_IDENTIFIER,
        false,
    );
    let present_map_original = present_map.clone();
    assert_refused(
        &source,
        final_owner_batch(present_map.as_slice(), &data_removals, &owner_removals)
            .expect("a known DataMetadataMap key must remain structurally valid"),
        "a present DataMetadataMap key",
    );
    assert_eq!(present_map, present_map_original, "present map was mutated");

    let stale_owner_removal = codec::DataReferenceOwnerRemoval::new(
        TRANSITION_COMPONENT,
        TRANSITION_DATA_IDENTIFIER,
        TRANSITION_OWNER_IDENTIFIER,
        2,
    );
    let stale_owner_removals = [stale_owner_removal];
    let wrong_count_map = Vec::new();
    assert_refused(
        &source,
        final_owner_batch(
            wrong_count_map.as_slice(),
            &data_removals,
            &stale_owner_removals,
        )
        .expect("an empty DataMetadataMap witness must remain valid"),
        "a stale final-owner count",
    );

    let partial_source = source_with_data_metadata_map(true);
    let partial_map = Vec::new();
    assert_refused(
        &partial_source,
        final_owner_batch(partial_map.as_slice(), &data_removals, &owner_removals)
            .expect("an empty DataMetadataMap witness must remain valid"),
        "a source with a surviving owner",
    );

    let unknown_map = data_metadata_map_payload(
        TRANSITION_DATA_IDENTIFIER + 1,
        MAP_METADATA_OBJECT_IDENTIFIER,
        true,
    );
    let unknown_map_original = unknown_map.clone();
    let original = source.clone();
    match final_owner_batch(unknown_map.as_slice(), &data_removals, &owner_removals) {
        Ok(batch) => assert_refused(
            &source,
            batch,
            "an unknown nested map reference with an absent key",
        ),
        Err(error) => {
            black_box((error.resource_limit(), error.invalid_reason()));
        },
    }
    assert_eq!(
        source,
        original.as_slice(),
        "unknown map case modified its source"
    );
    assert_eq!(unknown_map, unknown_map_original, "unknown map was mutated");
}

fn run_rewrite<'source>(
    source: &'source [u8],
    batch: codec::MediaRewriteBatch<'source>,
    decode_options: codec::DecodeOptions,
) -> Option<Vec<u8>> {
    let original = source.to_vec();
    let prepared =
        match codec::prepare_package_metadata_media_rewrite(source, batch, decode_options) {
            Ok(prepared) => prepared,
            Err(error) => {
                black_box((
                    error.resource_limit(),
                    error.invalid_reason(),
                    error.allocation_request(),
                ));
                return None;
            },
        };
    let requirements = prepared.execution_requirements();
    let exact_limits = requirements.exact_limits();
    let output = match prepared.execute(exact_limits) {
        Ok(output) => output,
        Err(error) => {
            black_box((
                error.resource_limit(),
                error.invalid_reason(),
                error.allocation_request(),
            ));
            return None;
        },
    };
    assert_eq!(output.bytes().len(), requirements.output_bytes());
    assert_eq!(output.report().output_bytes(), requirements.output_bytes());
    assert_eq!(output.report().fields(), requirements.fields());
    assert_eq!(output.report().work_bytes(), requirements.work_bytes());
    assert_eq!(output.report().allocations(), requirements.allocations());
    assert_eq!(
        output.report().retained_bytes(),
        requirements.retained_bytes()
    );
    assert_eq!(
        output.report().scratch_bytes(),
        requirements.scratch_bytes()
    );

    let one_shot = codec::rewrite_package_metadata_media(source, batch, decode_options)
        .expect("prepared and one-shot rewrites must accept the same valid batch");
    assert_eq!(one_shot.bytes(), output.bytes());
    assert_eq!(source, original.as_slice(), "rewrite modified its source");
    exercise_limit_failures(source, batch, decode_options, requirements);
    black_box(output.report());
    Some(output.into_bytes())
}

fn exercise_limit_failures<'source>(
    source: &'source [u8],
    batch: codec::MediaRewriteBatch<'source>,
    decode_options: codec::DecodeOptions,
    requirements: codec::RewriteExecutionRequirements,
) {
    const LIMITS: [RewriteLimit; 9] = [
        RewriteLimit::OutputBytes,
        RewriteLimit::Fields,
        RewriteLimit::WorkBytes,
        RewriteLimit::Components,
        RewriteLimit::DataRecords,
        RewriteLimit::Owners,
        RewriteLimit::Allocations,
        RewriteLimit::RetainedBytes,
        RewriteLimit::ScratchBytes,
    ];
    let original = source.to_vec();
    for kind in LIMITS {
        let Some(limits) = one_below_limits(requirements, kind) else {
            continue;
        };
        let prepared = codec::prepare_package_metadata_media_rewrite(source, batch, decode_options)
            .expect("exact limits must remain a valid preflight");
        let error = prepared
            .execute(limits)
            .expect_err("one-below execution limits must refuse publication");
        black_box((
            error.resource_limit(),
            error.invalid_reason(),
            error.allocation_request(),
        ));
        assert_eq!(
            source,
            original.as_slice(),
            "failed rewrite modified its source"
        );
    }
}

fn one_below_limits(
    requirements: codec::RewriteExecutionRequirements,
    kind: RewriteLimit,
) -> Option<codec::RewriteExecutionLimits> {
    let mut limits = requirements.exact_limits();
    let required = match kind {
        RewriteLimit::OutputBytes => &mut limits.max_output_bytes,
        RewriteLimit::Fields => &mut limits.max_fields,
        RewriteLimit::WorkBytes => &mut limits.max_work_bytes,
        RewriteLimit::Components => &mut limits.max_components,
        RewriteLimit::DataRecords => &mut limits.max_data_records,
        RewriteLimit::Owners => &mut limits.max_owners,
        RewriteLimit::Allocations => &mut limits.max_allocations,
        RewriteLimit::RetainedBytes => &mut limits.max_retained_bytes,
        RewriteLimit::ScratchBytes => &mut limits.max_scratch_bytes,
    };
    *required = required.checked_sub(1)?;
    Some(limits)
}
