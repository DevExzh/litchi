//! Independent contract tests for the bounded Numbers FormulaArchive envelope.
//!
//! The fixtures below are handwritten protobuf wire messages.  They do not use
//! generated builders, so schema admission, canonical scalar checks, and the
//! scalar-route decision are tested independently of the compatibility decoder.

use litchi_iwa_common::{Error, LimitKind, WireLimits};
use litchi_numbers_wire::formula_envelope::{
    AttemptedFormulaEnvelopeCost, FormulaEnvelopeFailure, FormulaEnvelopeLimits,
    preflight_formula_envelope,
};

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

fn field_varint(number: u32, value: u64) -> Vec<u8> {
    field_raw(number, 0, &varint(value))
}

fn field_raw(number: u32, wire_type: u8, payload: &[u8]) -> Vec<u8> {
    let mut output = varint((u64::from(number) << 3) | u64::from(wire_type));
    output.extend_from_slice(payload);
    output
}

fn field_fixed64(number: u32, value: u64) -> Vec<u8> {
    field_raw(number, 1, &value.to_le_bytes())
}

fn field_fixed64_raw(number: u32, payload: &[u8]) -> Vec<u8> {
    field_raw(number, 1, payload)
}

fn field_bytes(number: u32, payload: &[u8]) -> Vec<u8> {
    let mut output = varint((u64::from(number) << 3) | 2);
    output.extend(varint(
        u64::try_from(payload.len()).expect("fixture length fits in u64"),
    ));
    output.extend_from_slice(payload);
    output
}

fn field_bytes_with_length_prefix(number: u32, encoded_length: &[u8], payload: &[u8]) -> Vec<u8> {
    let mut output = varint((u64::from(number) << 3) | 2);
    output.extend_from_slice(encoded_length);
    output.extend_from_slice(payload);
    output
}

fn node(kind: u64, extra: &[u8]) -> Vec<u8> {
    let mut output = field_varint(1, kind);
    output.extend_from_slice(extra);
    output
}

fn function_node(identifier: u64, argument_count: u64) -> Vec<u8> {
    let mut fields = field_varint(2, identifier);
    fields.extend(field_varint(3, argument_count));
    node(16, &fields)
}

fn formula_with_node(node_payload: &[u8]) -> Vec<u8> {
    let ast = field_bytes(1, node_payload);
    field_bytes(1, &ast)
}

fn formula_with_nodes(nodes: &[&[u8]]) -> Vec<u8> {
    let mut ast = Vec::new();
    for node_payload in nodes {
        ast.extend(field_bytes(1, node_payload));
    }
    field_bytes(1, &ast)
}

fn append_root_field(source: &mut Vec<u8>, field: &[u8]) {
    source.extend_from_slice(field);
}

fn valid_number_formula() -> Vec<u8> {
    formula_with_node(&node(17, &[]))
}

fn valid_limits() -> FormulaEnvelopeLimits {
    FormulaEnvelopeLimits {
        max_fields: WireLimits::MAX_FIELDS,
        max_input_bytes: WireLimits::MAX_INPUT_BYTES,
        max_work: WireLimits::MAX_REWRITE_WORK,
        base_fields: 0,
        base_work: 0,
    }
}

fn assert_invalid(failure: &FormulaEnvelopeFailure) -> &Error {
    failure.error()
}

fn assert_limit(error: &Error, kind: LimitKind) {
    assert!(
        matches!(error, Error::LimitExceeded { kind: actual, .. } if *actual == kind),
        "expected {kind:?} limit error, got {error:?}"
    );
}

fn assert_invalid_limit(error: &Error, field: &'static str, value: usize, maximum: usize) {
    assert!(
        matches!(
            error,
            Error::InvalidLimit {
                field: actual,
                value: actual_value,
                maximum: actual_maximum,
            } if *actual == field && *actual_value == value && *actual_maximum == maximum
        ),
        "expected invalid limit {field}={value} (max {maximum}), got {error:?}"
    );
}

fn assert_attempted_positive(attempted: AttemptedFormulaEnvelopeCost) {
    assert!(
        attempted.work() > 0,
        "a failed scan must charge source work"
    );
    assert!(attempted.fields() > 0, "a failed scan must visit a field");
}

#[test]
fn empty_and_minimal_envelopes_report_root_and_scalar_counters() {
    let empty =
        preflight_formula_envelope(&[], valid_limits()).expect("empty archive is legacy-valid");
    assert!(!empty.root_ast_present());
    assert!(!empty.scalar_visitor_eligible());
    assert_eq!(empty.fields(), 0);
    assert_eq!(empty.scanned_bytes(), 0);
    assert_eq!(empty.scalar_visitor_node_count(), 0);
    assert_eq!(empty.lazy_traversal_entry_count(), 0);
    assert_eq!(empty.cost(), AttemptedFormulaEnvelopeCost::default());

    let source = valid_number_formula();
    let report =
        preflight_formula_envelope(&source, valid_limits()).expect("minimal number is valid");
    let node_payload = node(17, &[]);
    let ast_payload = field_bytes(1, &node_payload);
    assert!(report.root_ast_present());
    assert!(report.scalar_visitor_eligible());
    assert_eq!(report.scalar_visitor_node_count(), 1);
    assert_eq!(report.lazy_traversal_entry_count(), 0);
    assert_eq!(report.fields(), 3);
    assert_eq!(
        report.scanned_bytes(),
        source.len() + ast_payload.len() + node_payload.len()
    );
    assert_eq!(report.cost().fields(), report.fields());
    assert_eq!(report.cost().work(), report.scanned_bytes());
}

#[test]
fn unknown_fields_are_opaque_but_conservatively_leave_scalar_route() {
    let mut source = valid_number_formula();
    // The payload is intentionally not a nested message.  Unknown fields are
    // retained as opaque bytes and are not recursively interpreted.
    append_root_field(&mut source, &field_bytes(100, &[0xff, 0x80]));

    let report = preflight_formula_envelope(&source, valid_limits())
        .expect("unknown opaque fields must not invalidate a valid archive");
    assert!(report.root_ast_present());
    assert!(!report.scalar_visitor_eligible());
    assert_eq!(report.scalar_visitor_node_count(), 1);
    assert_eq!(report.lazy_traversal_entry_count(), 0);
    assert_eq!(report.fields(), 4);
}

#[test]
fn known_unreferenced_fields_are_still_strictly_validated() {
    let mut source = valid_number_formula();
    // Root field 7 is a known UUID envelope requiring both lower and upper
    // halves.  Nothing in the AST references it, but the envelope is still
    // part of the admitted source and must not be deferred past preflight.
    append_root_field(&mut source, &field_bytes(7, &[]));

    let failure = preflight_formula_envelope(&source, valid_limits())
        .expect_err("a malformed unreferenced known envelope must fail");
    let error = assert_invalid(&failure);
    assert!(matches!(error, Error::InvalidFormat(message) if message.contains("required field")));
}

#[test]
fn required_and_duplicate_fields_fail_closed() {
    let local_reference = node(27, &field_bytes(15, &[]));
    let missing_required = formula_with_node(&local_reference);
    let failure = preflight_formula_envelope(&missing_required, valid_limits())
        .expect_err("local reference envelope requires all four coordinates");
    assert!(matches!(
        assert_invalid(&failure),
        Error::InvalidFormat(message) if message.contains("required field")
    ));

    let mut duplicate_root = valid_number_formula();
    append_root_field(&mut duplicate_root, &field_varint(2, 1));
    append_root_field(&mut duplicate_root, &field_varint(2, 2));
    let failure = preflight_formula_envelope(&duplicate_root, valid_limits())
        .expect_err("known singular root fields must not use last-value-wins semantics");
    assert!(matches!(
        assert_invalid(&failure),
        Error::InvalidFormat(message) if message.contains("occurs more than once")
    ));

    let duplicate_node = {
        let mut payload = field_varint(1, 17);
        payload.extend(field_varint(1, 17));
        payload
    };
    let failure = preflight_formula_envelope(&formula_with_node(&duplicate_node), valid_limits())
        .expect_err("AST node type is singular");
    assert!(matches!(
        assert_invalid(&failure),
        Error::InvalidFormat(message) if message.contains("occurs more than once")
    ));
}

#[test]
fn wrong_wire_types_noncanonical_values_and_invalid_utf8_are_rejected() {
    let wrong_root = field_varint(1, 1);
    let failure = preflight_formula_envelope(&wrong_root, valid_limits())
        .expect_err("FormulaArchive field 1 must be length-delimited");
    assert!(
        matches!(assert_invalid(&failure), Error::InvalidFormat(message) if message.contains("wrong") || message.contains("wire type"))
    );

    let noncanonical_length = field_bytes_with_length_prefix(1, &[0x81, 0x00], &[0]);
    let failure = preflight_formula_envelope(&noncanonical_length, valid_limits())
        .expect_err("known length-delimited fields must use canonical lengths");
    assert!(matches!(
        assert_invalid(&failure),
        Error::InvalidFormat(message) if message.contains("noncanonical")
    ));

    let overlong_type = vec![0x08, 0x91, 0x00];
    let failure = preflight_formula_envelope(&formula_with_node(&overlong_type), valid_limits())
        .expect_err("known scalar varints must use their canonical width");
    assert!(matches!(
        assert_invalid(&failure),
        Error::InvalidFormat(message) if message.contains("noncanonical")
    ));

    let bad_bool = node(17, &field_varint(5, 2));
    let failure = preflight_formula_envelope(&formula_with_node(&bad_bool), valid_limits())
        .expect_err("boolean fields accept only zero and one");
    assert!(matches!(
        assert_invalid(&failure),
        Error::InvalidFormat(message) if message.contains("noncanonical scalar value")
    ));

    let bad_u32 = node(17, &field_varint(2, u64::from(u32::MAX) + 1));
    let failure = preflight_formula_envelope(&formula_with_node(&bad_u32), valid_limits())
        .expect_err("u32 fields must not carry values outside their domain");
    assert!(matches!(
        assert_invalid(&failure),
        Error::InvalidFormat(message) if message.contains("noncanonical scalar value")
    ));

    let bad_int32 = node(17, &field_varint(9, 0x8000_0000));
    let failure = preflight_formula_envelope(&formula_with_node(&bad_int32), valid_limits())
        .expect_err("int32 fields reject the gap between signed ranges");
    assert!(matches!(
        assert_invalid(&failure),
        Error::InvalidFormat(message) if message.contains("noncanonical scalar value")
    ));

    let bad_utf8 = node(19, &field_bytes(6, &[0xff]));
    let failure = preflight_formula_envelope(&formula_with_node(&bad_utf8), valid_limits())
        .expect_err("known UTF-8 fields must be valid before lazy rendering");
    assert!(matches!(
        assert_invalid(&failure),
        Error::InvalidFormat(message) if message.contains("not valid UTF-8")
    ));

    let valid_fixed64 = node(17, &field_fixed64(4, 0));
    let report = preflight_formula_envelope(&formula_with_node(&valid_fixed64), valid_limits())
        .expect("a complete fixed64 number field is valid");
    assert!(report.scalar_visitor_eligible());

    let malformed_fixed64 = node(17, &field_fixed64_raw(4, &[0; 7]));
    let failure =
        preflight_formula_envelope(&formula_with_node(&malformed_fixed64), valid_limits())
            .expect_err("known fixed64 fields must contain exactly eight bytes");
    assert!(matches!(assert_invalid(&failure), Error::InvalidFormat(_)));
}

#[test]
fn scalar_eligibility_is_conservative_for_strings_unknown_functions_and_local_refs() {
    let string = node(19, &field_bytes(6, b"hello"));
    let report = preflight_formula_envelope(&formula_with_node(&string), valid_limits())
        .expect("valid string node remains readable");
    assert!(!report.scalar_visitor_eligible());
    assert_eq!(report.scalar_visitor_node_count(), 1);

    let supported_function = function_node(15, 1);
    let number = node(17, &[]);
    let report = preflight_formula_envelope(
        &formula_with_nodes(&[number.as_slice(), supported_function.as_slice()]),
        valid_limits(),
    )
    .expect("supported function identifiers remain on the scalar route");
    assert!(report.scalar_visitor_eligible());
    assert_eq!(report.scalar_visitor_node_count(), 2);

    let unknown_function = function_node(999, 1);
    let number = node(17, &[]);
    let report = preflight_formula_envelope(
        &formula_with_nodes(&[number.as_slice(), unknown_function.as_slice()]),
        valid_limits(),
    )
    .expect("unknown function identifiers remain compatibility-readable");
    assert!(!report.scalar_visitor_eligible());

    let mut local = Vec::new();
    for field in 1..=4 {
        local.extend(field_varint(field, 0));
    }
    let local_reference = node(27, &field_bytes(15, &local));
    let report = preflight_formula_envelope(&formula_with_node(&local_reference), valid_limits())
        .expect("a complete local reference envelope is structurally valid");
    assert!(!report.scalar_visitor_eligible());
    assert_eq!(report.scalar_visitor_node_count(), 1);
}

#[test]
fn aggregate_limits_are_inclusive_and_failures_expose_attempted_cost() {
    let source = valid_number_formula();
    let baseline = preflight_formula_envelope(&source, valid_limits()).expect("baseline is valid");

    let exact = FormulaEnvelopeLimits {
        max_fields: baseline.fields(),
        max_input_bytes: baseline.scanned_bytes(),
        max_work: baseline.scanned_bytes(),
        base_fields: 0,
        base_work: 0,
    };
    let exact_report = preflight_formula_envelope(&source, exact).expect("limits are inclusive");
    assert_eq!(exact_report, baseline);

    let mut one_short_fields = exact;
    one_short_fields.max_fields -= 1;
    let failure = preflight_formula_envelope(&source, one_short_fields)
        .expect_err("one field below the exact aggregate must fail");
    let attempted = failure.attempted();
    assert_limit(assert_invalid(&failure), LimitKind::Fields);
    assert_attempted_positive(attempted);

    let mut one_short_work = exact;
    one_short_work.max_work -= 1;
    let failure = preflight_formula_envelope(&source, one_short_work)
        .expect_err("one byte below the exact aggregate work must fail");
    assert_limit(assert_invalid(&failure), LimitKind::RewriteWork);
    assert_attempted_positive(failure.attempted());

    let mut one_short_input = exact;
    one_short_input.max_input_bytes -= 1;
    let failure = preflight_formula_envelope(&source, one_short_input)
        .expect_err("one byte below the scanned input must fail");
    assert_limit(assert_invalid(&failure), LimitKind::InputBytes);
    assert!(failure.attempted().work() > 0);

    let mut cumulative_fields = exact;
    cumulative_fields.base_fields = 1;
    let failure = preflight_formula_envelope(&source, cumulative_fields)
        .expect_err("base field usage must count toward the aggregate ceiling");
    assert_limit(assert_invalid(&failure), LimitKind::Fields);

    let mut cumulative_work = exact;
    cumulative_work.base_work = 1;
    let failure = preflight_formula_envelope(&source, cumulative_work)
        .expect_err("base work usage must count toward the aggregate ceiling");
    assert_limit(assert_invalid(&failure), LimitKind::RewriteWork);
}

#[test]
fn invalid_aggregate_offsets_and_ceilings_fail_before_scanning() {
    let source = valid_number_formula();

    let fields_at_max = FormulaEnvelopeLimits {
        max_fields: WireLimits::MAX_FIELDS,
        max_input_bytes: WireLimits::MAX_INPUT_BYTES,
        max_work: WireLimits::MAX_REWRITE_WORK,
        base_fields: usize::MAX,
        base_work: 0,
    };
    let failure = preflight_formula_envelope(&source, fields_at_max)
        .expect_err("base fields above the configured ceiling must be rejected");
    assert_invalid_limit(
        assert_invalid(&failure),
        "formula envelope base fields",
        usize::MAX,
        WireLimits::MAX_FIELDS,
    );

    let work_at_max = FormulaEnvelopeLimits {
        max_fields: WireLimits::MAX_FIELDS,
        max_input_bytes: WireLimits::MAX_INPUT_BYTES,
        max_work: WireLimits::MAX_REWRITE_WORK,
        base_fields: 0,
        base_work: usize::MAX,
    };
    let failure = preflight_formula_envelope(&source, work_at_max)
        .expect_err("base work above the configured ceiling must be rejected");
    assert_invalid_limit(
        assert_invalid(&failure),
        "formula envelope base work",
        usize::MAX,
        WireLimits::MAX_REWRITE_WORK,
    );

    let oversized_fields = FormulaEnvelopeLimits {
        max_fields: usize::MAX,
        max_input_bytes: WireLimits::MAX_INPUT_BYTES,
        max_work: WireLimits::MAX_REWRITE_WORK,
        base_fields: 0,
        base_work: 0,
    };
    let failure = preflight_formula_envelope(&source, oversized_fields)
        .expect_err("field ceilings above the common hard cap must be rejected");
    assert_invalid_limit(
        assert_invalid(&failure),
        "formula envelope fields",
        usize::MAX,
        WireLimits::MAX_FIELDS,
    );

    let oversized_input = FormulaEnvelopeLimits {
        max_fields: WireLimits::MAX_FIELDS,
        max_input_bytes: usize::MAX,
        max_work: WireLimits::MAX_REWRITE_WORK,
        base_fields: 0,
        base_work: 0,
    };
    let failure = preflight_formula_envelope(&source, oversized_input)
        .expect_err("input ceilings above the common hard cap must be rejected");
    assert_invalid_limit(
        assert_invalid(&failure),
        "formula envelope input bytes",
        usize::MAX,
        WireLimits::MAX_INPUT_BYTES,
    );

    let oversized_work = FormulaEnvelopeLimits {
        max_fields: WireLimits::MAX_FIELDS,
        max_input_bytes: WireLimits::MAX_INPUT_BYTES,
        max_work: usize::MAX,
        base_fields: 0,
        base_work: 0,
    };
    let failure = preflight_formula_envelope(&source, oversized_work)
        .expect_err("work ceilings above the common hard cap must be rejected");
    assert_invalid_limit(
        assert_invalid(&failure),
        "formula envelope work",
        usize::MAX,
        WireLimits::MAX_REWRITE_WORK,
    );
}

#[test]
fn zero_ceilings_reject_nonempty_input_but_preserve_empty_default() {
    let zero = FormulaEnvelopeLimits {
        max_fields: 0,
        max_input_bytes: 0,
        max_work: 0,
        base_fields: 0,
        base_work: 0,
    };
    let empty = preflight_formula_envelope(&[], zero)
        .expect("the historical empty archive remains valid under zero budgets");
    assert!(!empty.root_ast_present());
    assert_eq!(empty.cost(), AttemptedFormulaEnvelopeCost::default());

    let source = valid_number_formula();

    let mut no_input = valid_limits();
    no_input.max_input_bytes = 0;
    let failure = preflight_formula_envelope(&source, no_input)
        .expect_err("a nonempty source must charge input bytes before scanning");
    assert_limit(assert_invalid(&failure), LimitKind::InputBytes);
    assert_eq!(failure.attempted().fields(), 0);
    assert_eq!(failure.attempted().work(), source.len());

    let mut no_fields = valid_limits();
    no_fields.max_fields = 0;
    let failure = preflight_formula_envelope(&source, no_fields)
        .expect_err("a nonempty source must visit at least one field");
    assert_limit(assert_invalid(&failure), LimitKind::Fields);
    assert_eq!(failure.attempted().fields(), 1);
    assert_eq!(failure.attempted().work(), source.len());

    let mut no_work = valid_limits();
    no_work.max_work = 0;
    let failure = preflight_formula_envelope(&source, no_work)
        .expect_err("a nonempty source must charge wire work before scanning");
    assert_limit(assert_invalid(&failure), LimitKind::RewriteWork);
    assert_eq!(failure.attempted().fields(), 0);
    assert_eq!(failure.attempted().work(), source.len());
}
