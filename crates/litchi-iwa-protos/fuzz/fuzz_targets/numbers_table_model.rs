#![no_main]

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::numbers_table_cell_storage_codec::{
    DataStoreSnapshot, DecodeError, DecodeOptions, HeaderRecord, ReferenceRecord, StorageVisitor,
    TableDataListEntrySnapshot, TableModelSnapshot, TileReferenceRecord, TileRowInfoSnapshot,
    decode_data_store_compatibility_with_report, decode_data_store_compatibility_with_visitor,
    decode_data_store_dense_native_with_report, decode_data_store_with_report,
    decode_data_store_with_visitor, decode_table_data_list_segment_type_with_report,
    decode_table_data_list_segment_with_report, decode_table_data_list_type_with_report,
    decode_table_model_compatibility_with_report, decode_table_model_compatibility_with_visitor,
    decode_table_model_with_compatibility_data_store_with_report, decode_table_model_with_report,
    decode_table_model_with_visitor,
};
use litchi_iwa_protos::tsce;
use litchi_iwa_protos::tsp::{self, Reference};
use litchi_iwa_protos::tst::{
    DataStore, HeaderStorage, TableDataList, TableDataListSegment, TableModelArchive, TableRbTree,
    TileStorage,
};
use prost::Message as _;

// Inputs are skipped rather than truncated, so every decode path receives one
// unchanged caller-owned source. The codec's aggregate ledger is bounded by
// the same finite profile used by the other Numbers wire targets.
const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_REFERENCES: usize = 1_024;
const MAX_TEXT_BYTES: usize = 64 * 1024;
const MAX_RECURSION: u32 = 64;
const MAX_RETAINED_CALLBACKS: usize = MAX_FIELDS;
const CONTROL_BYTES: usize = 8;

#[derive(Clone, Copy)]
struct SourceRange {
    start: usize,
    end: usize,
}

impl SourceRange {
    fn new(source: &[u8]) -> Self {
        let start = source.as_ptr() as usize;
        let end = start
            .checked_add(source.len())
            .expect("bounded source pointer range");
        Self { start, end }
    }

    fn assert_borrowed(self, bytes: &[u8]) {
        if bytes.is_empty() {
            return;
        }
        let start = bytes.as_ptr() as usize;
        let end = start
            .checked_add(bytes.len())
            .expect("bounded borrowed payload range");
        assert!(
            start >= self.start && end <= self.end,
            "decoded payload did not borrow from its source"
        );
    }
}

#[derive(Default)]
struct CallbackFacts {
    tile_references: usize,
    rows: usize,
    header_buckets: usize,
    headers: usize,
    entries: usize,
    segments: usize,
}

impl CallbackFacts {
    fn count(slot: &mut usize) {
        *slot = slot
            .checked_add(1)
            .expect("bounded callback count overflowed");
    }

    fn assert_bounded(&self) {
        assert!(self.tile_references <= MAX_RETAINED_CALLBACKS);
        assert!(self.rows <= MAX_RETAINED_CALLBACKS);
        assert!(self.header_buckets <= MAX_RETAINED_CALLBACKS);
        assert!(self.headers <= MAX_RETAINED_CALLBACKS);
        assert!(self.entries <= MAX_RETAINED_CALLBACKS);
        assert!(self.segments <= MAX_RETAINED_CALLBACKS);
    }
}

struct Collector {
    source: SourceRange,
    facts: CallbackFacts,
}

impl Collector {
    fn new(source: &[u8]) -> Self {
        Self {
            source: SourceRange::new(source),
            facts: CallbackFacts::default(),
        }
    }
}

impl StorageVisitor for Collector {
    fn visit_tile_reference(&mut self, record: TileReferenceRecord<'_>) -> Result<(), DecodeError> {
        self.source.assert_borrowed(record.raw());
        CallbackFacts::count(&mut self.facts.tile_references);
        Ok(())
    }

    fn visit_tile_row(&mut self, row: TileRowInfoSnapshot<'_>) -> Result<(), DecodeError> {
        self.source
            .assert_borrowed(row.cell_storage_buffer_pre_bnc());
        self.source.assert_borrowed(row.cell_offsets_pre_bnc());
        if let Some(bytes) = row.cell_storage_buffer() {
            self.source.assert_borrowed(bytes);
        }
        if let Some(bytes) = row.cell_offsets() {
            self.source.assert_borrowed(bytes);
        }
        CallbackFacts::count(&mut self.facts.rows);
        Ok(())
    }

    fn visit_header_bucket(&mut self, record: ReferenceRecord<'_>) -> Result<(), DecodeError> {
        self.source.assert_borrowed(record.raw());
        CallbackFacts::count(&mut self.facts.header_buckets);
        Ok(())
    }

    fn visit_header_record(&mut self, record: HeaderRecord<'_>) -> Result<(), DecodeError> {
        self.source.assert_borrowed(record.raw());
        CallbackFacts::count(&mut self.facts.headers);
        Ok(())
    }

    fn visit_list_entry(
        &mut self,
        entry: TableDataListEntrySnapshot<'_>,
    ) -> Result<(), DecodeError> {
        if let Some(value) = entry.string_value() {
            self.source.assert_borrowed(value.as_bytes());
        }
        for payload in [
            entry.formula(),
            entry.format(),
            entry.custom_format(),
            entry.import_warning_set(),
            entry.cell_spec(),
        ]
        .into_iter()
        .flatten()
        {
            self.source.assert_borrowed(payload);
        }
        CallbackFacts::count(&mut self.facts.entries);
        Ok(())
    }

    fn visit_list_segment(&mut self, record: ReferenceRecord<'_>) -> Result<(), DecodeError> {
        self.source.assert_borrowed(record.raw());
        CallbackFacts::count(&mut self.facts.segments);
        Ok(())
    }
}

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };

    // Arbitrary input is still valuable for malformed-wire and resource-limit
    // coverage. Independent synthetic recipes keep valid model/store paths
    // reachable even when a random byte source is not a protobuf payload.
    exercise_model(&source);
    exercise_store(&source);
    exercise_segment_routes(&source);

    let model = match control(data, 0) % 5 {
        0 => canonical_model(data),
        1 => duplicate_model_field(data),
        2 => wrong_wire_model(data),
        3 => truncated_model(data),
        _ => invalid_utf8_model(data),
    };
    exercise_model(&model);

    let store = match control(data, 1) % 4 {
        0 => canonical_store(data),
        1 => duplicate_store_field(data),
        2 => malformed_reference_store(data),
        _ => malformed_group_store(data),
    };
    exercise_store(&store);

    let (root, segment) = canonical_list_routes(data);
    exercise_segment_routes(&root);
    exercise_segment_routes(&segment);
    assert_route_shape_rejection(&root, &segment);
});

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

fn options() -> DecodeOptions {
    DecodeOptions::new(
        MAX_INPUT_BYTES,
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        MAX_REFERENCES,
        MAX_TEXT_BYTES,
    )
}

fn exercise_model(source: &[u8]) {
    let before = source.to_vec();
    let strict = decode_table_model_with_report(source, options());
    assert_eq!(source, before.as_slice(), "model decode modified source");

    let mut strict_visitor = Collector::new(source);
    let streamed = decode_table_model_with_visitor(source, options(), &mut strict_visitor);
    assert_eq!(source, before.as_slice(), "model visitor modified source");
    strict_visitor.facts.assert_bounded();
    assert_same_result(
        strict.clone(),
        streamed,
        "strict table-model report/visitor",
    );

    let mut compatibility_visitor = Collector::new(source);
    let compatibility = decode_table_model_compatibility_with_report(source, options());
    assert_eq!(
        source,
        before.as_slice(),
        "compatibility model modified source"
    );
    let compatibility_streamed = decode_table_model_compatibility_with_visitor(
        source,
        options(),
        &mut compatibility_visitor,
    );
    assert_eq!(
        source,
        before.as_slice(),
        "compatibility model visitor modified source"
    );
    compatibility_visitor.facts.assert_bounded();
    assert_same_result(
        compatibility.clone(),
        compatibility_streamed,
        "compatibility table-model report/visitor",
    );

    if let Ok((snapshot, _report)) = compatibility {
        assert_model_borrows(source, snapshot);
    }

    let dense = decode_table_model_with_compatibility_data_store_with_report(source, options());
    assert_eq!(
        source,
        before.as_slice(),
        "dense-native model modified source"
    );
    if let Ok((snapshot, _report)) = dense {
        assert_model_borrows(source, snapshot);
    }

    if let Ok((snapshot, _report)) = strict {
        assert_model_borrows(source, snapshot);
        // Fields 39/93 are intentionally opaque in the strict projection;
        // a wire-valid payload can therefore be outside Prost's typed schema.
        // Use Prost as a semantic oracle whenever it accepts that complete
        // generated shape, while retaining strict coverage for opaque-only
        // successes.
        if let Ok(oracle) = TableModelArchive::decode(source) {
            assert_table_model_matches(snapshot, &oracle);
        }
        exercise_store(snapshot.base_data_store());
    }
}

fn exercise_store(source: &[u8]) {
    let before = source.to_vec();
    let strict = decode_data_store_with_report(source, options());
    assert_eq!(source, before.as_slice(), "store decode modified source");
    let mut strict_visitor = Collector::new(source);
    let streamed = decode_data_store_with_visitor(source, options(), &mut strict_visitor);
    assert_eq!(source, before.as_slice(), "store visitor modified source");
    strict_visitor.facts.assert_bounded();
    assert_same_result(strict.clone(), streamed, "strict data-store report/visitor");

    let compatibility = decode_data_store_compatibility_with_report(source, options());
    assert_eq!(
        source,
        before.as_slice(),
        "compatibility store modified source"
    );
    let mut compatibility_visitor = Collector::new(source);
    let compatibility_streamed =
        decode_data_store_compatibility_with_visitor(source, options(), &mut compatibility_visitor);
    assert_eq!(
        source,
        before.as_slice(),
        "compatibility store visitor modified source"
    );
    compatibility_visitor.facts.assert_bounded();
    assert_same_result(
        compatibility.clone(),
        compatibility_streamed,
        "compatibility data-store report/visitor",
    );

    if let Ok((snapshot, _report)) = compatibility {
        assert_store_borrows(source, snapshot);
    }

    let dense = decode_data_store_dense_native_with_report(source, options());
    assert_eq!(
        source,
        before.as_slice(),
        "dense-native store modified source"
    );
    if let Ok((snapshot, _report)) = dense {
        assert_store_borrows(source, snapshot);
    }

    if let Ok((snapshot, _report)) = strict {
        assert_store_borrows(source, snapshot);
        // Row/header trees and sidecar payloads remain source-owned opaque
        // routes in this projection. A strict wire success may consequently
        // be rejected by Prost's typed nested schema; compare the semantic
        // oracle for the subset Prost can decode.
        if let Ok(oracle) = DataStore::decode(source) {
            assert_data_store_matches(snapshot, &oracle);
        }
    }
}

fn exercise_segment_routes(source: &[u8]) {
    // The table extractor probes a scalar envelope before deciding whether a
    // payload is a root list or a referenced segment. Exercise both probes on
    // arbitrary bytes and on the complete synthetic routes below; a failed
    // probe must leave the caller-owned source untouched.
    exercise_segment_route(source, false);
    exercise_segment_route(source, true);
}

fn exercise_segment_route(source: &[u8], segment: bool) {
    let before = source.to_vec();
    let probe = if segment {
        decode_table_data_list_segment_type_with_report(source, options())
    } else {
        decode_table_data_list_type_with_report(source, options())
    };
    assert_eq!(
        source,
        before.as_slice(),
        "list route probe modified source"
    );

    let Ok((snapshot, report)) = probe else {
        return;
    };
    assert_storage_report(report, source.len());

    // A successful scalar probe is only an admission decision; run the full
    // route as well so valid synthetic entries exercise the selected segment
    // decoder rather than stopping at the routing envelope.
    if segment {
        if let Ok((full_snapshot, full_report)) =
            decode_table_data_list_segment_with_report(source, options())
        {
            assert_storage_report(full_report, source.len());
            assert_eq!(full_snapshot.list_type(), snapshot.list_type());
            // The scalar probe intentionally skips repeated entry payloads.
            // Only ask Prost to act as an oracle after the full strict route
            // accepts; arbitrary fuzz bytes may contain a valid envelope and
            // a malformed repeated child that the probe correctly leaves
            // opaque.
            if let Ok(oracle) = TableDataListSegment::decode(source) {
                assert_eq!(snapshot.list_type(), oracle.list_type);
            }
        }
    } else if let Ok((full_snapshot, full_report)) =
        litchi_iwa_protos::numbers_table_cell_storage_codec::decode_table_data_list_with_report(
            source,
            options(),
        )
    {
        assert_storage_report(full_report, source.len());
        assert_eq!(full_snapshot.list_type(), snapshot.list_type());
        if let Ok(oracle) = TableDataList::decode(source) {
            assert_eq!(snapshot.list_type(), oracle.list_type);
        }
    }
}

fn assert_route_shape_rejection(root: &[u8], segment: &[u8]) {
    let root_as_segment = decode_table_data_list_segment_type_with_report(root, options());
    assert!(
        root_as_segment.is_err(),
        "root envelope was admitted as a segment route"
    );
    let segment_as_root = decode_table_data_list_type_with_report(segment, options());
    assert!(
        segment_as_root.is_err(),
        "segment envelope was admitted as a root route"
    );
}

fn assert_same_result<T: Copy + PartialEq>(
    report: Result<
        (
            T,
            litchi_iwa_protos::numbers_table_cell_storage_codec::DecodeReport,
        ),
        DecodeError,
    >,
    visitor: Result<
        (
            T,
            litchi_iwa_protos::numbers_table_cell_storage_codec::DecodeReport,
        ),
        DecodeError,
    >,
    label: &str,
) {
    match (report, visitor) {
        (Ok((report_snapshot, report_usage)), Ok((visitor_snapshot, visitor_usage))) => {
            assert!(
                report_snapshot == visitor_snapshot,
                "{label} snapshot mismatch"
            );
            assert_eq!(report_usage, visitor_usage, "{label} report mismatch");
        },
        (Err(_), Err(_)) => {},
        (Ok(_), Err(error)) => panic!("{label} visitor rejected report success: {error:?}"),
        (Err(error), Ok(_)) => panic!("{label} visitor accepted report rejection: {error:?}"),
    }
}

fn assert_model_borrows(source: &[u8], snapshot: TableModelSnapshot<'_>) {
    let range = SourceRange::new(source);
    range.assert_borrowed(snapshot.table_id().as_bytes());
    range.assert_borrowed(snapshot.table_name().as_bytes());
    range.assert_borrowed(snapshot.base_data_store());
    range.assert_optional(snapshot.conditional_style_formula_owner_id());
    range.assert_optional(snapshot.spill_owner());
}

fn assert_table_model_matches(snapshot: TableModelSnapshot<'_>, oracle: &TableModelArchive) {
    assert_eq!(snapshot.table_id(), oracle.table_id.as_str());
    assert_eq!(snapshot.table_name(), oracle.table_name.as_str());
    assert_eq!(snapshot.number_of_rows(), oracle.number_of_rows);
    assert_eq!(snapshot.number_of_columns(), oracle.number_of_columns);

    // The handwritten projection preserves proto2 presence for these routes,
    // while generated required references have default values when omitted.
    // Compare every route that the strict snapshot actually admitted; this
    // keeps the oracle semantic without treating generated defaults as source
    // presence evidence.
    assert_reference_if_present(snapshot.table_style(), &oracle.table_style);
    assert_reference_if_present(snapshot.body_text_style(), &oracle.body_text_style);
    assert_reference_if_present(
        snapshot.header_row_text_style(),
        &oracle.header_row_text_style,
    );
    assert_reference_if_present(
        snapshot.header_column_text_style(),
        &oracle.header_column_text_style,
    );
    assert_reference_if_present(
        snapshot.footer_row_text_style(),
        &oracle.footer_row_text_style,
    );
    assert_reference_if_present(snapshot.body_cell_style(), &oracle.body_cell_style);
    assert_reference_if_present(snapshot.header_row_style(), &oracle.header_row_style);
    assert_reference_if_present(snapshot.header_column_style(), &oracle.header_column_style);
    assert_reference_if_present(snapshot.footer_row_style(), &oracle.footer_row_style);
    assert_optional_reference(
        snapshot.table_name_style(),
        oracle.table_name_style.as_ref(),
    );
    assert_optional_reference(
        snapshot.table_name_shape_style(),
        oracle.table_name_shape_style.as_ref(),
    );
    assert_optional_reference(
        snapshot.hidden_state_formula_owner_for_columns(),
        oracle.hidden_state_formula_owner_for_columns.as_ref(),
    );
    assert_optional_reference(
        snapshot.hidden_state_formula_owner_for_rows(),
        oracle.hidden_state_formula_owner_for_rows.as_ref(),
    );
    assert_optional_reference(snapshot.pivot_owner(), oracle.pivot_owner.as_ref());
    assert_optional_reference(snapshot.category_owner(), oracle.category_owner.as_ref());

    if let Some(raw) = snapshot.conditional_style_formula_owner_id() {
        let expected = oracle
            .conditional_style_formula_owner_id
            .as_ref()
            .expect("strict conditional-owner presence disagreed with Prost");
        let decoded = tsp::CfuuidArchive::decode(raw).unwrap_or_else(|error| {
            panic!("strict conditional-owner acceptance disagreed with Prost: {error}")
        });
        assert_eq!(&decoded, expected);
    }
    if let Some(raw) = snapshot.spill_owner() {
        let expected = oracle
            .spill_owner
            .as_ref()
            .expect("strict spill-owner presence disagreed with Prost");
        let decoded = tsce::SpillOwnerArchive::decode(raw).unwrap_or_else(|error| {
            panic!("strict spill-owner acceptance disagreed with Prost: {error}")
        });
        assert_eq!(&decoded, expected);
    }
}

fn assert_store_borrows(source: &[u8], snapshot: DataStoreSnapshot<'_>) {
    let range = SourceRange::new(source);
    range.assert_borrowed(snapshot.row_headers());
    range.assert_borrowed(snapshot.tiles());
    range.assert_borrowed(snapshot.row_tile_tree());
    range.assert_borrowed(snapshot.column_tile_tree());
}

fn assert_data_store_matches(snapshot: DataStoreSnapshot<'_>, oracle: &DataStore) {
    let row_headers = HeaderStorage::decode(snapshot.row_headers()).unwrap_or_else(|error| {
        panic!("strict row-header acceptance disagreed with Prost: {error}")
    });
    assert_eq!(&row_headers, &oracle.row_headers);
    let tiles = TileStorage::decode(snapshot.tiles()).unwrap_or_else(|error| {
        panic!("strict tile-storage acceptance disagreed with Prost: {error}")
    });
    assert_eq!(&tiles, &oracle.tiles);
    let row_tile_tree = TableRbTree::decode(snapshot.row_tile_tree())
        .unwrap_or_else(|error| panic!("strict row-tree acceptance disagreed with Prost: {error}"));
    assert_eq!(&row_tile_tree, &oracle.row_tile_tree);
    let column_tile_tree =
        TableRbTree::decode(snapshot.column_tile_tree()).unwrap_or_else(|error| {
            panic!("strict column-tree acceptance disagreed with Prost: {error}")
        });
    assert_eq!(&column_tile_tree, &oracle.column_tile_tree);

    assert_reference(snapshot.column_headers(), &oracle.column_headers);
    assert_reference(snapshot.string_table(), &oracle.string_table);
    assert_reference(snapshot.style_table(), &oracle.style_table);
    assert_reference(snapshot.formula_table(), &oracle.formula_table);
    assert_eq!(snapshot.next_row_strip_id(), oracle.next_row_strip_id);
    assert_eq!(snapshot.next_column_strip_id(), oracle.next_column_strip_id);
    assert_reference(
        snapshot.format_table_pre_bnc(),
        &oracle.format_table_pre_bnc,
    );
    assert_optional_reference(
        snapshot.formula_error_table(),
        oracle.formula_error_table.as_ref(),
    );
    assert_optional_reference(
        snapshot.merge_region_map(),
        oracle.merge_region_map.as_ref(),
    );
    assert_eq!(
        snapshot.storage_version_pre_bnc(),
        oracle.storage_version_pre_bnc
    );
    assert_optional_reference(
        snapshot.deprecated_custom_format_table(),
        oracle.deprecated_custom_format_table.as_ref(),
    );
    assert_optional_reference(
        snapshot.multiple_choice_list_format_table(),
        oracle.multiple_choice_list_format_table.as_ref(),
    );
    assert_optional_reference(snapshot.rich_text_table(), oracle.rich_text_table.as_ref());
    assert_optional_reference(
        snapshot.conditional_style_table(),
        oracle.conditionalstyletable.as_ref(),
    );
    assert_optional_reference(
        snapshot.comment_storage_table(),
        oracle.comment_storage_table.as_ref(),
    );
    assert_optional_reference(
        snapshot.import_warning_set_table(),
        oracle.import_warning_set_table.as_ref(),
    );
    assert_optional_reference(
        snapshot.control_cell_spec_table(),
        oracle.control_cell_spec_table.as_ref(),
    );
    assert_optional_reference(snapshot.format_table(), oracle.format_table.as_ref());
}

fn assert_reference_if_present(
    snapshot: Option<litchi_iwa_protos::numbers_table_cell_storage_codec::ReferenceSnapshot>,
    oracle: &Reference,
) {
    if let Some(snapshot) = snapshot {
        assert_reference(snapshot, oracle);
    }
}

fn assert_optional_reference(
    snapshot: Option<litchi_iwa_protos::numbers_table_cell_storage_codec::ReferenceSnapshot>,
    oracle: Option<&Reference>,
) {
    match (snapshot, oracle) {
        (Some(snapshot), Some(oracle)) => assert_reference(snapshot, oracle),
        (None, None) => {},
        // Generated required fields cannot carry source presence. For an
        // omitted optional strict field, only reject a generated value when
        // the strict decoder also retained one; this arm documents that the
        // generated default is not evidence that the source contained it.
        (None, Some(_)) => {},
        (Some(_), None) => panic!("strict optional reference presence disagreed with Prost"),
    }
}

fn assert_reference(
    snapshot: litchi_iwa_protos::numbers_table_cell_storage_codec::ReferenceSnapshot,
    oracle: &Reference,
) {
    assert_eq!(snapshot.identifier(), oracle.identifier);
    assert_eq!(snapshot.deprecated_type(), oracle.deprecated_type);
    assert_eq!(
        snapshot.deprecated_is_external(),
        oracle.deprecated_is_external
    );
}

fn assert_storage_report(
    report: litchi_iwa_protos::numbers_table_cell_storage_codec::DecodeReport,
    source_bytes: usize,
) {
    assert_eq!(report.source_bytes(), source_bytes);
    assert!(report.fields() <= MAX_FIELDS);
    assert!(report.work_bytes() <= MAX_WORK_BYTES);
    assert!(report.max_depth() <= MAX_RECURSION);
    assert!(report.references() <= MAX_REFERENCES);
    assert!(report.text_bytes() <= MAX_TEXT_BYTES);
}

impl SourceRange {
    fn assert_optional(self, bytes: Option<&[u8]>) {
        if let Some(bytes) = bytes {
            self.assert_borrowed(bytes);
        }
    }
}

// --- Small canonical-wire recipes -------------------------------------------------

fn canonical_model(data: &[u8]) -> Vec<u8> {
    let store = canonical_store(data);
    let mut model = Vec::new();
    bytes_field(&mut model, 1, b"table-fuzz");
    bytes_field(&mut model, 4, &store);
    varint_field(&mut model, 6, u64::from(control(data, 2)));
    varint_field(&mut model, 7, u64::from(control(data, 3)));
    bytes_field(&mut model, 8, b"Table fuzz");
    for (field, id) in [
        (3, 11),
        (18, 12),
        (19, 13),
        (20, 14),
        (21, 15),
        (24, 16),
        (25, 17),
        (26, 18),
        (27, 19),
        (30, 20),
        (36, 21),
        (34, 22),
        (35, 23),
        (85, 24),
        (86, 25),
    ] {
        bytes_field(&mut model, field, &reference(id));
    }
    bytes_field(&mut model, 39, &[]);
    bytes_field(&mut model, 93, &[]);
    model
}

fn canonical_store(data: &[u8]) -> Vec<u8> {
    let mut headers = Vec::new();
    varint_field(&mut headers, 1, u64::from(control(data, 4) % 8));

    let mut store = Vec::new();
    bytes_field(&mut store, 1, &headers);
    bytes_field(&mut store, 2, &reference(2));
    bytes_field(&mut store, 3, &tile_storage(data));
    bytes_field(&mut store, 4, &reference(3));
    bytes_field(&mut store, 5, &reference(4));
    bytes_field(&mut store, 6, &reference(5));
    varint_field(&mut store, 7, u64::from(control(data, 5)));
    varint_field(&mut store, 8, u64::from(control(data, 6)));
    bytes_field(&mut store, 9, &[]);
    bytes_field(&mut store, 10, &[]);
    bytes_field(&mut store, 11, &reference(6));
    for (field, id) in [
        (12, 7),
        (13, 8),
        (15, 9),
        (16, 10),
        (17, 11),
        (18, 12),
        (19, 13),
        (20, 14),
        (21, 15),
        (22, 16),
    ] {
        bytes_field(&mut store, field, &reference(id));
    }
    varint_field(&mut store, 14, 0);
    store
}

fn tile_storage(data: &[u8]) -> Vec<u8> {
    let mut tile = Vec::new();
    varint_field(&mut tile, 2, u64::from(control(data, 7)));
    varint_field(&mut tile, 3, u64::from(control(data, 0) & 1));
    tile
}

fn duplicate_model_field(data: &[u8]) -> Vec<u8> {
    let mut model = canonical_model(data);
    bytes_field(&mut model, 1, b"duplicate");
    model
}

fn wrong_wire_model(data: &[u8]) -> Vec<u8> {
    let mut model = canonical_model(data);
    // Field 6 is a required varint; append a length-delimited occurrence.
    model.push(0x32);
    model.push(0x00);
    model
}

fn truncated_model(data: &[u8]) -> Vec<u8> {
    let mut model = canonical_model(data);
    model.push(0x22);
    model.push(0x80);
    model
}

fn invalid_utf8_model(data: &[u8]) -> Vec<u8> {
    let mut model = canonical_model(data);
    bytes_field(&mut model, 1, &[0xff]);
    model
}

fn duplicate_store_field(data: &[u8]) -> Vec<u8> {
    let mut store = canonical_store(data);
    bytes_field(&mut store, 4, &reference(99));
    store
}

fn malformed_reference_store(data: &[u8]) -> Vec<u8> {
    let mut store = canonical_store(data);
    let mut invalid = Vec::new();
    varint_field(&mut invalid, 3, 1);
    bytes_field(&mut store, 4, &invalid);
    store
}

fn malformed_group_store(data: &[u8]) -> Vec<u8> {
    let mut store = canonical_store(data);
    // Start-group without an end-group must be rejected by both strict paths.
    // Keep the group at an unknown top-level field so this case reaches group
    // skipping rather than being rejected as a duplicate selected field.
    put_varint(&mut store, (90_u64 << 3) | 3);
    varint_field(&mut store, 91, u64::from(control(data, 1)));
    store
}

fn canonical_list_routes(data: &[u8]) -> (Vec<u8>, Vec<u8>) {
    let list_type = u64::from(control(data, 0) % 12 + 1);
    let entry = canonical_list_entry(data);

    let mut root = Vec::new();
    varint_field(&mut root, 1, list_type);
    varint_field(&mut root, 2, u64::from(control(data, 1)));
    bytes_field(&mut root, 3, &entry);
    bytes_field(&mut root, 4, &reference(0x30 + u64::from(control(data, 2))));
    varint_field(&mut root, 5, u64::from(control(data, 3) & 1));

    let mut range = Vec::new();
    varint_field(&mut range, 1, u64::from(control(data, 4)));
    varint_field(&mut range, 2, 1);
    let mut segment = Vec::new();
    varint_field(&mut segment, 1, list_type);
    bytes_field(&mut segment, 2, &range);
    bytes_field(&mut segment, 3, &entry);

    (root, segment)
}

fn canonical_list_entry(data: &[u8]) -> Vec<u8> {
    let mut entry = Vec::new();
    varint_field(&mut entry, 1, u64::from(control(data, 5)));
    varint_field(&mut entry, 2, 1);
    bytes_field(&mut entry, 3, b"route");
    bytes_field(
        &mut entry,
        4,
        &reference(0x40 + u64::from(control(data, 6))),
    );
    entry
}

fn reference(identifier: u64) -> Vec<u8> {
    let mut output = Vec::new();
    varint_field(&mut output, 1, identifier);
    output
}

fn bytes_field(output: &mut Vec<u8>, field: u32, value: &[u8]) {
    put_varint(output, (u64::from(field) << 3) | 2);
    put_varint(
        output,
        u64::try_from(value.len()).expect("bounded table-model payload length"),
    );
    output.extend_from_slice(value);
}

fn varint_field(output: &mut Vec<u8>, field: u32, value: u64) {
    put_varint(output, u64::from(field) << 3);
    put_varint(output, value);
}

fn put_varint(output: &mut Vec<u8>, mut value: u64) {
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

fn control(data: &[u8], index: usize) -> u8 {
    data.get(index % CONTROL_BYTES).copied().unwrap_or_default()
}
