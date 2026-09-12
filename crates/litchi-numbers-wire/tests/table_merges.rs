//! Independent contract tests for the bounded native merge reader.
//!
//! The fixtures are handwritten protobuf wire messages.  They intentionally
//! avoid the generated `TST` and `TSCE` types so that the selected-path reader
//! is tested against the wire contract rather than against its own encoder.

use litchi_iwa_common::table::merge::Region;
use litchi_iwa_common::{Error, LimitKind, WireLimits};
use litchi_numbers_wire::table_merges::{ReadLimits, read_table_merges};

const TABLE_ID: &str = "00112233445566778899aabbccddeeff";
// `formula_owner_uuid_for_table` reverses the table UUID bytes before it is
// lowered to the four uint32 words in a native CFUUID archive.
const TABLE_OWNER_WORDS: [u32; 4] = [0x3322_1100, 0x7766_5544, 0xbbaa_9988, 0xffee_ddcc];

fn varint(mut value: u64) -> Vec<u8> {
    let mut output = Vec::new();
    loop {
        let byte = u8::try_from(value & 0x7f).expect("a varint chunk fits in a byte");
        value >>= 7;
        if value == 0 {
            output.push(byte);
            return output;
        }
        output.push(byte | 0x80);
    }
}

fn field_raw(number: u32, wire_type: u8, payload: &[u8]) -> Vec<u8> {
    let mut output = varint((u64::from(number) << 3) | u64::from(wire_type));
    output.extend_from_slice(payload);
    output
}

fn field_varint(number: u32, value: u64) -> Vec<u8> {
    field_raw(number, 0, &varint(value))
}

fn field_bytes(number: u32, payload: &[u8]) -> Vec<u8> {
    let mut output = field_raw(
        number,
        2,
        &varint(u64::try_from(payload.len()).expect("fixture length fits in u64")),
    );
    output.extend_from_slice(payload);
    output
}

fn with_unknown(mut source: Vec<u8>) -> Vec<u8> {
    // Unknown fields are retained as opaque bytes.  The payload is
    // deliberately not a valid nested message, so a reader that recursively
    // interprets unknown data would reject this otherwise valid fixture.
    source.extend(field_bytes(100, &[0xde, 0xad, 0xbe, 0xef]));
    source
}

fn cfuuid(words: [u32; 4]) -> Vec<u8> {
    cfuuid_with_uuid_bytes(words, None)
}

fn cfuuid_with_uuid_bytes(words: [u32; 4], uuid_bytes: Option<&[u8]>) -> Vec<u8> {
    let mut output = Vec::new();
    if let Some(uuid_bytes) = uuid_bytes {
        output.extend(field_bytes(1, uuid_bytes));
    }
    for (index, word) in words.into_iter().enumerate() {
        output.extend(field_varint(
            u32::try_from(index + 2).expect("four UUID words fit in u32"),
            u64::from(word),
        ));
    }
    output
}

fn absolute_range(begin: u32, end: Option<u32>) -> Vec<u8> {
    let mut output = field_varint(1, u64::from(begin));
    if let Some(end) = end {
        output.extend(field_varint(2, u64::from(end)));
    }
    output
}

fn sticky_bits(values: [bool; 4]) -> Vec<u8> {
    let mut output = Vec::new();
    for (index, value) in values.into_iter().enumerate() {
        output.extend(field_varint(
            u32::try_from(index + 1).expect("four sticky fields fit in u32"),
            u64::from(value),
        ));
    }
    output
}

fn range_node(
    region: Region,
    table_words: Option<[u32; 4]>,
    sticky: [bool; 4],
    include_relative_range: bool,
    preserve_rectangular: Option<bool>,
    unknown: bool,
) -> Vec<u8> {
    range_node_with_uuid_bytes(
        region,
        table_words,
        sticky,
        include_relative_range,
        preserve_rectangular,
        unknown,
        None,
    )
}

fn range_node_with_uuid_bytes(
    region: Region,
    table_words: Option<[u32; 4]>,
    sticky: [bool; 4],
    include_relative_range: bool,
    preserve_rectangular: Option<bool>,
    unknown: bool,
    uuid_bytes: Option<&[u8]>,
) -> Vec<u8> {
    range_node_with_extra_cross_fields(
        region,
        table_words,
        sticky,
        include_relative_range,
        preserve_rectangular,
        unknown,
        uuid_bytes,
        &[],
    )
}

#[allow(clippy::too_many_arguments)]
fn range_node_with_extra_cross_fields(
    region: Region,
    table_words: Option<[u32; 4]>,
    sticky: [bool; 4],
    include_relative_range: bool,
    preserve_rectangular: Option<bool>,
    unknown: bool,
    uuid_bytes: Option<&[u8]>,
    extra_cross_fields: &[u8],
) -> Vec<u8> {
    let mut node = field_varint(1, 67); // AST_COLON_TRACT_NODE
    if let Some(table_words) = table_words {
        let mut cross_extra = field_bytes(1, &cfuuid_with_uuid_bytes(table_words, uuid_bytes));
        cross_extra.extend_from_slice(extra_cross_fields);
        node.extend(field_bytes(28, &cross_extra));
    }
    node.extend(field_bytes(33, &sticky_bits(sticky)));

    let mut tract = Vec::new();
    if include_relative_range {
        // The relative range uses an int32 begin coordinate.  Zero is enough
        // to make the range present and therefore unsupported for a merge.
        tract.extend(field_bytes(1, &field_varint(1, 0)));
    }
    tract.extend(field_bytes(
        3,
        &absolute_range(region.column(), Some(region.end_column())),
    ));
    tract.extend(field_bytes(
        4,
        &absolute_range(region.row(), Some(region.end_row())),
    ));
    if let Some(value) = preserve_rectangular {
        tract.extend(field_varint(5, u64::from(value)));
    }
    node.extend(field_bytes(40, &tract));
    if unknown {
        node = with_unknown(node);
    }
    node
}

fn function_node(identifier: u32, argument_count: u32, unknown: bool) -> Vec<u8> {
    let mut node = field_varint(1, 16); // AST_FUNCTION_NODE
    node.extend(field_varint(2, u64::from(identifier)));
    node.extend(field_varint(3, u64::from(argument_count)));
    if unknown {
        node = with_unknown(node);
    }
    node
}

fn formula(
    region: Region,
    table_words: Option<[u32; 4]>,
    function: (u32, u32),
    sticky: [bool; 4],
    include_relative_range: bool,
    preserve_rectangular: Option<bool>,
    unknown: bool,
    extra_node: Option<Vec<u8>>,
) -> Vec<u8> {
    formula_with_uuid_bytes(
        region,
        table_words,
        function,
        sticky,
        include_relative_range,
        preserve_rectangular,
        unknown,
        extra_node,
        None,
    )
}

fn formula_with_uuid_bytes(
    region: Region,
    table_words: Option<[u32; 4]>,
    function: (u32, u32),
    sticky: [bool; 4],
    include_relative_range: bool,
    preserve_rectangular: Option<bool>,
    unknown: bool,
    extra_node: Option<Vec<u8>>,
    uuid_bytes: Option<&[u8]>,
) -> Vec<u8> {
    let range = range_node_with_extra_cross_fields(
        region,
        table_words,
        sticky,
        include_relative_range,
        preserve_rectangular,
        unknown,
        uuid_bytes,
        &[],
    );
    let function = function_node(function.0, function.1, unknown);
    let mut ast = field_bytes(1, &range);
    ast.extend(field_bytes(1, &function));
    if let Some(extra_node) = extra_node {
        ast.extend(field_bytes(1, &extra_node));
    }
    let mut output = field_bytes(1, &ast);
    if unknown {
        output = with_unknown(output);
    }
    output
}

#[allow(clippy::too_many_arguments)]
fn formula_with_extra_cross_fields(
    region: Region,
    table_words: Option<[u32; 4]>,
    function: (u32, u32),
    sticky: [bool; 4],
    include_relative_range: bool,
    preserve_rectangular: Option<bool>,
    unknown: bool,
    uuid_bytes: Option<&[u8]>,
    extra_cross_fields: &[u8],
    extra_node: Option<Vec<u8>>,
) -> Vec<u8> {
    let range = range_node_with_extra_cross_fields(
        region,
        table_words,
        sticky,
        include_relative_range,
        preserve_rectangular,
        unknown,
        uuid_bytes,
        extra_cross_fields,
    );
    let function = function_node(function.0, function.1, unknown);
    let mut ast = field_bytes(1, &range);
    ast.extend(field_bytes(1, &function));
    if let Some(extra_node) = extra_node {
        ast.extend(field_bytes(1, &extra_node));
    }
    let mut output = field_bytes(1, &ast);
    if unknown {
        output = with_unknown(output);
    }
    output
}

fn pair(index: u32, formula: &[u8], unknown: bool) -> Vec<u8> {
    let mut output = field_varint(1, u64::from(index));
    output.extend(field_bytes(2, formula));
    if unknown {
        output = with_unknown(output);
    }
    output
}

fn formula_store(next_index: u32, pairs: &[Vec<u8>], unknown: bool) -> Vec<u8> {
    let mut output = field_varint(2, u64::from(next_index));
    for pair in pairs {
        output.extend(field_bytes(3, pair));
    }
    if unknown {
        output = with_unknown(output);
    }
    output
}

fn merge_owner(store: Option<&[u8]>, unknown: bool) -> Vec<u8> {
    merge_owner_with_uid(&cfuuid(TABLE_OWNER_WORDS), store, unknown)
}

fn merge_owner_with_uid(uid: &[u8], store: Option<&[u8]>, unknown: bool) -> Vec<u8> {
    let mut output = field_bytes(1, uid);
    if let Some(store) = store {
        output.extend(field_bytes(2, store));
    }
    if unknown {
        output = with_unknown(output);
    }
    output
}

fn table_model(rows: u32, columns: u32, owner: Option<&[u8]>, unknown: bool) -> Vec<u8> {
    let mut output = field_bytes(1, TABLE_ID.as_bytes());
    output.extend(field_varint(6, u64::from(rows)));
    output.extend(field_varint(7, u64::from(columns)));
    if let Some(owner) = owner {
        output.extend(field_bytes(47, owner));
    }
    if unknown {
        output = with_unknown(output);
    }
    output
}

fn one_region_formula(region: Region) -> Vec<u8> {
    formula(
        region,
        Some(TABLE_OWNER_WORDS),
        (168, 1),
        [true; 4],
        false,
        Some(true),
        false,
        None,
    )
}

fn model_with_regions(regions: &[Region]) -> Vec<u8> {
    let pairs: Vec<_> = regions
        .iter()
        .enumerate()
        .map(|(index, region)| pair(index as u32, &one_region_formula(*region), false))
        .collect();
    let store = formula_store(regions.len() as u32, &pairs, false);
    let owner = merge_owner(Some(&store), false);
    table_model(16, 16, Some(&owner), false)
}

fn default_limits() -> ReadLimits {
    ReadLimits::default()
}

fn assert_invalid(source: &[u8]) {
    let failure = read_table_merges(source, default_limits()).expect_err("fixture must be invalid");
    assert!(
        matches!(failure.error(), Error::InvalidFormat(_)),
        "expected InvalidFormat, got {failure:?}"
    );
    assert!(
        failure.attempted().fields > 0 || failure.attempted().input_bytes > 0,
        "invalid input must report attempted work: {failure:?}"
    );
}

fn assert_limit(source: &[u8], limits: ReadLimits, kind: LimitKind, maximum: usize) {
    let failure = read_table_merges(source, limits).expect_err("fixture must exceed its limit");
    assert!(
        matches!(
            failure.error(),
            Error::LimitExceeded {
                kind: actual_kind,
                observed,
                limit: actual_limit,
            } if *actual_kind == kind && *actual_limit <= maximum && *observed > *actual_limit
        ),
        "expected {kind:?} to exceed configured limit {maximum}, got {failure:?}"
    );
}

#[test]
fn reads_non_overlapping_regions_in_source_order() {
    let first = Region::new(1, 2, 2, 3).unwrap();
    let second = Region::new(8, 0, 1, 2).unwrap();
    let source = model_with_regions(&[first, second]);
    let before = source.clone();

    let result = read_table_merges(&source, default_limits()).expect("valid merge graph");
    assert_eq!(source, before, "reader must borrow without rewriting input");
    assert_eq!(result.regions, [first, second]);
    assert!(result.report.input_bytes() >= source.len());
    assert!(result.report.fields() > 0);
    assert!(result.report.work() > 0);
}

#[test]
fn absent_owner_and_empty_store_are_empty_successes() {
    let without_owner = table_model(4, 4, None, false);
    let result = read_table_merges(&without_owner, default_limits()).expect("no merge owner");
    assert!(result.regions.is_empty());
    assert!(result.report.fields() > 0);

    let owner_without_store = merge_owner(None, false);
    let model = table_model(4, 4, Some(&owner_without_store), false);
    let result = read_table_merges(&model, default_limits()).expect("empty merge owner");
    assert!(result.regions.is_empty());

    let empty_store = formula_store(0, &[], false);
    let owner_with_empty_store = merge_owner(Some(&empty_store), false);
    let model = table_model(4, 4, Some(&owner_with_empty_store), false);
    let result = read_table_merges(&model, default_limits()).expect("empty formula store");
    assert!(result.regions.is_empty());
}

#[test]
fn rejects_missing_required_merge_owner_id() {
    let store = formula_store(1, &[], false);
    let owner_without_id = field_bytes(2, &store);

    assert_invalid(&table_model(8, 8, Some(&owner_without_id), false));
}

#[test]
fn rejects_missing_required_empty_formula_store_index() {
    let empty_store = Vec::new();
    let owner = merge_owner(Some(&empty_store), false);

    assert_invalid(&table_model(8, 8, Some(&owner), false));
}

#[test]
fn unknown_fields_remain_opaque_at_every_selected_level() {
    let region = Region::new(2, 3, 2, 2).unwrap();
    let encoded_formula = formula(
        region,
        Some(TABLE_OWNER_WORDS),
        (168, 1),
        [true; 4],
        false,
        Some(true),
        true,
        None,
    );
    let encoded_pair = pair(0, &encoded_formula, true);
    let encoded_store = formula_store(1, &[encoded_pair], true);
    let encoded_owner = merge_owner(Some(&encoded_store), true);
    let source = table_model(8, 8, Some(&encoded_owner), true);

    let result = read_table_merges(&source, default_limits())
        .expect("canonically framed unknown fields are opaque");
    assert_eq!(result.regions, [region]);
}

#[test]
fn accepts_valid_cross_table_display_names() {
    let region = Region::new(1, 1, 2, 2).unwrap();
    let mut cross_fields = field_bytes(2, b"Sheet 1");
    cross_fields.extend(field_bytes(3, b"Table 1"));
    cross_fields.extend(field_bytes(4, b"Summary"));
    cross_fields.extend(field_bytes(5, b"Revenue"));
    let encoded = formula_with_extra_cross_fields(
        region,
        Some(TABLE_OWNER_WORDS),
        (168, 1),
        [true; 4],
        false,
        Some(true),
        false,
        None,
        &cross_fields,
        None,
    );
    let pair = pair(0, &encoded, false);
    let store = formula_store(1, &[pair], false);
    let owner = merge_owner(Some(&store), false);
    let result = read_table_merges(&table_model(8, 8, Some(&owner), false), default_limits())
        .expect("valid cross-table display names remain accepted");
    assert_eq!(result.regions, [region]);
}

#[test]
fn rejects_malformed_cross_table_display_names() {
    let region = Region::new(1, 1, 2, 2).unwrap();
    let invalid_wire_type = field_varint(2, 1);
    let invalid_utf8 = field_bytes(2, &[0xff]);
    for cross_fields in [invalid_wire_type, invalid_utf8] {
        let encoded = formula_with_extra_cross_fields(
            region,
            Some(TABLE_OWNER_WORDS),
            (168, 1),
            [true; 4],
            false,
            Some(true),
            false,
            None,
            &cross_fields,
            None,
        );
        let pair = pair(0, &encoded, false);
        let store = formula_store(1, &[pair], false);
        let owner = merge_owner(Some(&store), false);
        assert_invalid(&table_model(8, 8, Some(&owner), false));
    }
}

#[test]
fn rejects_matching_cfuuid_words_with_uuid_bytes_present() {
    let region = Region::new(1, 1, 2, 2).unwrap();
    for uuid_bytes in [Vec::new(), vec![0xde, 0xad, 0xbe, 0xef]] {
        let encoded = formula_with_uuid_bytes(
            region,
            Some(TABLE_OWNER_WORDS),
            (168, 1),
            [true; 4],
            false,
            Some(true),
            false,
            None,
            Some(&uuid_bytes),
        );
        let pair = pair(0, &encoded, false);
        let store = formula_store(1, &[pair], false);
        let owner = merge_owner(Some(&store), false);
        assert_invalid(&table_model(8, 8, Some(&owner), false));
    }
}

#[test]
fn rejects_truncated_owner_cfuuid_payload() {
    // The nested CFUUID word is a truncated varint.  The outer owner field
    // remains correctly length-delimited, so only a nested schema-aware scan
    // can reject this known field.
    let truncated_uuid = field_raw(2, 0, &[0x80]);
    let owner = merge_owner_with_uid(&truncated_uuid, None, false);
    assert_invalid(&table_model(8, 8, Some(&owner), false));
}

#[test]
fn rejects_missing_and_malformed_merge_formula_shape() {
    let region = Region::new(1, 1, 2, 2).unwrap();
    let valid_range = range_node(
        region,
        Some(TABLE_OWNER_WORDS),
        [true; 4],
        false,
        Some(true),
        false,
    );
    let valid_function = function_node(168, 1, false);

    let cases = [
        // A merge formula is exactly a colon tract followed by the native
        // merge function.
        formula_with_nodes(vec![valid_range.clone()]),
        formula_with_nodes(vec![valid_range.clone(), function_node(167, 1, false)]),
        formula_with_nodes(vec![
            range_node(
                region,
                Some(TABLE_OWNER_WORDS),
                [true, false, true, true],
                false,
                Some(true),
                false,
            ),
            valid_function.clone(),
        ]),
        formula_with_nodes(vec![
            range_node(
                region,
                Some(TABLE_OWNER_WORDS),
                [true; 4],
                true,
                Some(true),
                false,
            ),
            valid_function.clone(),
        ]),
        formula_with_nodes(vec![
            range_node(
                region,
                Some(TABLE_OWNER_WORDS),
                [true; 4],
                false,
                Some(false),
                false,
            ),
            valid_function.clone(),
        ]),
        formula_with_nodes(vec![
            range_node(region, None, [true; 4], false, Some(true), false),
            valid_function.clone(),
        ]),
        formula_with_nodes(vec![
            valid_range,
            valid_function,
            field_varint(1, 17), // NumberNode: an extra postfix operand.
        ]),
    ];

    for formula in cases {
        let pair = pair(0, &formula, false);
        let store = formula_store(1, &[pair], false);
        let owner = merge_owner(Some(&store), false);
        assert_invalid(&table_model(8, 8, Some(&owner), false));
    }
}

fn formula_with_nodes(nodes: Vec<Vec<u8>>) -> Vec<u8> {
    let mut ast = Vec::new();
    for node in nodes {
        ast.extend(field_bytes(1, &node));
    }
    field_bytes(1, &ast)
}

#[test]
fn rejects_duplicate_and_out_of_range_formula_indexes() {
    let region = Region::new(1, 1, 2, 2).unwrap();
    let encoded = one_region_formula(region);

    let duplicate = [pair(0, &encoded, false), pair(0, &encoded, false)];
    let duplicate_store = formula_store(2, &duplicate, false);
    let duplicate_owner = merge_owner(Some(&duplicate_store), false);
    assert_invalid(&table_model(8, 8, Some(&duplicate_owner), false));

    let out_of_range = [pair(2, &encoded, false)];
    let out_of_range_store = formula_store(2, &out_of_range, false);
    let out_of_range_owner = merge_owner(Some(&out_of_range_store), false);
    assert_invalid(&table_model(8, 8, Some(&out_of_range_owner), false));
}

#[test]
fn rejects_foreign_formula_owner_and_regions_outside_table_bounds() {
    let region = Region::new(1, 1, 2, 2).unwrap();
    let foreign_words = [1, 2, 3, 4];
    let foreign_formula = formula(
        region,
        Some(foreign_words),
        (168, 1),
        [true; 4],
        false,
        Some(true),
        false,
        None,
    );
    let foreign_pair = pair(0, &foreign_formula, false);
    let foreign_store = formula_store(1, &[foreign_pair], false);
    let foreign_owner = merge_owner(Some(&foreign_store), false);
    assert_invalid(&table_model(8, 8, Some(&foreign_owner), false));

    let out_of_bounds = Region::new(7, 7, 2, 2).unwrap();
    let pair = pair(0, &one_region_formula(out_of_bounds), false);
    let store = formula_store(1, &[pair], false);
    let owner = merge_owner(Some(&store), false);
    assert_invalid(&table_model(8, 8, Some(&owner), false));
}

#[test]
fn rejects_overlapping_regions_without_returning_a_partial_projection() {
    let first = Region::new(1, 1, 3, 3).unwrap();
    let second = Region::new(3, 3, 2, 2).unwrap();
    let source = model_with_regions(&[first, second]);
    let failure = read_table_merges(&source, default_limits()).expect_err("overlap is invalid");
    assert!(
        matches!(failure.error(), Error::InvalidFormat(message) if message.contains("overlap"))
    );
}

#[test]
fn rejects_noncanonical_wire_framing() {
    let mut source = field_bytes(1, TABLE_ID.as_bytes());
    source.extend(field_varint(7, 8));
    // Number-of-rows is a varint.  The value zero encoded with two bytes is
    // noncanonical and must not be silently accepted by the selected scan.
    source.extend(field_raw(6, 0, &[0x80, 0x00]));
    assert_invalid(&source);
}

#[test]
fn exact_input_and_field_limits_are_inclusive() {
    let source = model_with_regions(&[Region::new(1, 1, 2, 2).unwrap()]);
    let baseline = read_table_merges(&source, default_limits()).expect("valid merge graph");

    let exact_wire = WireLimits::default()
        .with_input_bytes(baseline.report.input_bytes())
        .unwrap()
        .with_fields(baseline.report.fields())
        .unwrap();
    let exact_limits = ReadLimits {
        wire: exact_wire,
        ..default_limits()
    };
    let exact = read_table_merges(&source, exact_limits).expect("exact report limits fit");
    assert_eq!(exact.report, baseline.report);

    let input_limit = baseline.report.input_bytes() - 1;
    let input_wire = WireLimits::default().with_input_bytes(input_limit).unwrap();
    assert_limit(
        &source,
        ReadLimits {
            wire: input_wire,
            ..default_limits()
        },
        LimitKind::InputBytes,
        input_limit,
    );

    let field_limit = baseline.report.fields() - 1;
    let field_wire = WireLimits::default().with_fields(field_limit).unwrap();
    assert_limit(
        &source,
        ReadLimits {
            wire: field_wire,
            ..default_limits()
        },
        LimitKind::Fields,
        field_limit,
    );
}

#[test]
fn region_and_overlap_limits_fail_before_unbounded_growth() {
    let first = Region::new(1, 1, 2, 2).unwrap();
    let second = Region::new(5, 5, 2, 2).unwrap();
    let source = model_with_regions(&[first, second]);

    assert_limit(
        &source,
        ReadLimits {
            max_regions: 1,
            ..default_limits()
        },
        LimitKind::TableCells,
        1,
    );
    assert_limit(
        &source,
        ReadLimits {
            max_overlap_checks: 0,
            ..default_limits()
        },
        LimitKind::RewriteWork,
        0,
    );

    let invalid_regions = ReadLimits {
        max_regions: WireLimits::MAX_FIELDS + 1,
        ..default_limits()
    };
    let failure = read_table_merges(&source, invalid_regions).expect_err("invalid region limit");
    assert!(matches!(
        failure.error(),
        Error::InvalidLimit {
            field: "table merge regions",
            value,
            maximum: WireLimits::MAX_FIELDS,
        } if *value == WireLimits::MAX_FIELDS + 1
    ));
}
