#![no_main]

use std::hint::black_box;

use libfuzzer_sys::fuzz_target;
use litchi_numbers_wire::formula_envelope::{FormulaEnvelopeLimits, preflight_formula_envelope};

const MAX_SOURCE_BYTES: usize = 1_024;
const MAX_FIELDS: usize = 256;
const MAX_WORK: usize = 8_192;

/// Append one canonical protobuf varint.
fn append_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn append_key(output: &mut Vec<u8>, number: u32, wire_type: u8) {
    append_varint(output, (u64::from(number) << 3) | u64::from(wire_type));
}

fn append_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    append_key(output, number, 0);
    append_varint(output, value);
}

fn append_fixed64_field(output: &mut Vec<u8>, number: u32, value: u64) {
    append_key(output, number, 1);
    output.extend_from_slice(&value.to_le_bytes());
}

fn append_bytes_field(output: &mut Vec<u8>, number: u32, value: &[u8]) {
    append_key(output, number, 2);
    append_varint(output, value.len() as u64);
    output.extend_from_slice(value);
}

fn node(kind: u32) -> Vec<u8> {
    let mut output = Vec::with_capacity(3);
    append_varint_field(&mut output, 1, u64::from(kind));
    output
}

fn number_node(value: f64) -> Vec<u8> {
    let mut output = node(17);
    append_fixed64_field(&mut output, 4, value.to_bits());
    output
}

fn boolean_node(value: bool) -> Vec<u8> {
    let mut output = node(18);
    append_varint_field(&mut output, 5, u64::from(value));
    output
}

fn string_node(value: &[u8]) -> Vec<u8> {
    let mut output = node(19);
    append_bytes_field(&mut output, 6, value);
    output
}

fn date_node(value: f64) -> Vec<u8> {
    let mut output = node(20);
    append_fixed64_field(&mut output, 7, value.to_bits());
    output
}

fn duration_node(value: f64, unit: u32) -> Vec<u8> {
    let mut output = node(21);
    append_fixed64_field(&mut output, 8, value.to_bits());
    append_varint_field(&mut output, 9, u64::from(unit));
    output
}

fn function_node(identifier: u32, arguments: u32) -> Vec<u8> {
    let mut output = node(16);
    append_varint_field(&mut output, 2, u64::from(identifier));
    append_varint_field(&mut output, 3, u64::from(arguments));
    output
}

fn array_node(columns: u32, rows: u32) -> Vec<u8> {
    let mut output = node(24);
    append_varint_field(&mut output, 11, u64::from(columns));
    append_varint_field(&mut output, 12, u64::from(rows));
    output
}

fn unknown_function_node(name: &[u8], arguments: u32) -> Vec<u8> {
    let mut output = node(31);
    append_bytes_field(&mut output, 17, name);
    append_varint_field(&mut output, 18, u64::from(arguments));
    output
}

fn local_reference_node() -> Vec<u8> {
    let mut reference = Vec::new();
    append_varint_field(&mut reference, 1, 2);
    append_varint_field(&mut reference, 2, 3);
    append_varint_field(&mut reference, 3, 1);
    append_varint_field(&mut reference, 4, 0);

    let mut output = node(27);
    append_bytes_field(&mut output, 15, &reference);
    output
}

fn cross_reference_node() -> Vec<u8> {
    let mut cfuuid = Vec::new();
    append_bytes_field(&mut cfuuid, 1, &[1, 2, 3, 4]);

    let mut reference = Vec::new();
    append_varint_field(&mut reference, 1, 2);
    append_varint_field(&mut reference, 2, 3);
    append_varint_field(&mut reference, 3, 1);
    append_varint_field(&mut reference, 4, 0);
    append_bytes_field(&mut reference, 5, &cfuuid);
    append_bytes_field(&mut reference, 6, b" ");
    append_bytes_field(&mut reference, 9, b" ");

    let mut output = node(28);
    append_bytes_field(&mut output, 16, &reference);
    output
}

fn thunk_node(child: Vec<u8>) -> Vec<u8> {
    let mut array = Vec::new();
    append_bytes_field(&mut array, 1, &child);
    let mut output = node(26);
    append_bytes_field(&mut output, 14, &array);
    output
}

fn formula(nodes: &[Vec<u8>]) -> Vec<u8> {
    let mut array = Vec::new();
    for node in nodes {
        append_bytes_field(&mut array, 1, node);
    }
    let mut output = Vec::new();
    append_bytes_field(&mut output, 1, &array);
    output
}

fn formula_with_array(array: &[u8]) -> Vec<u8> {
    let mut output = Vec::new();
    append_bytes_field(&mut output, 1, array);
    output
}

fn finite_value(cursor: &mut Cursor<'_>) -> f64 {
    let value = f64::from_bits(cursor.u64());
    if value.is_finite() {
        value
    } else {
        f64::from(cursor.byte()) / 4.0
    }
}

fn bounded_ascii(cursor: &mut Cursor<'_>, maximum: usize) -> Vec<u8> {
    let length = cursor.bounded(maximum.saturating_add(1));
    (0..length)
        .map(|_| b'a'.saturating_add(cursor.byte() % 26))
        .collect()
}

fn rich_formula(cursor: &mut Cursor<'_>) -> Vec<u8> {
    let number = number_node(finite_value(cursor));
    let boolean = boolean_node(cursor.byte() & 1 != 0);
    let string = if cursor.byte() & 1 == 0 {
        string_node(b"Caf\xC3\xA9")
    } else {
        let value = bounded_ascii(cursor, 16);
        string_node(&value)
    };
    let date = date_node(finite_value(cursor));
    let duration = duration_node(finite_value(cursor), u32::from(cursor.byte() % 4));
    let function = function_node([15, 30, 84, 88, 168][cursor.bounded(5)], 2);
    let array = array_node(2, 1);
    let unknown_function = unknown_function_node(b"NATIVE", 1);
    let local = local_reference_node();
    let cross = cross_reference_node();
    let thunk = thunk_node(number_node(1.0));
    formula(&[
        number,
        boolean,
        string,
        date,
        duration,
        function,
        array,
        unknown_function,
        local,
        cross,
        thunk,
    ])
}

fn deep_thunk(depth: usize) -> Vec<u8> {
    let mut current = number_node(1.0);
    for _ in 0..depth {
        current = thunk_node(current);
    }
    formula(&[current])
}

fn inspect(source: &[u8], limits: FormulaEnvelopeLimits) {
    match preflight_formula_envelope(source, limits) {
        Ok(report) => {
            black_box((
                report.fields(),
                report.scanned_bytes(),
                report.root_ast_present(),
                report.scalar_visitor_eligible(),
                report.scalar_visitor_node_count(),
                report.lazy_traversal_entry_count(),
                report.cost(),
                report.wire_preflight(),
            ));
        },
        Err(failure) => {
            let attempted = failure.attempted();
            black_box((attempted.fields(), attempted.work(), failure.error()));
        },
    }
}

fn inspect_with_budget_variants(source: &[u8], selector: u8) {
    let source_length = source.len();
    let generous = FormulaEnvelopeLimits {
        max_fields: MAX_FIELDS,
        max_input_bytes: MAX_SOURCE_BYTES,
        max_work: MAX_WORK,
        base_fields: 0,
        base_work: 0,
    };
    inspect(source, generous);

    // A one-field allowance exercises failures after the source charge for
    // nested and repeated messages without allowing unbounded traversal.
    inspect(
        source,
        FormulaEnvelopeLimits {
            max_fields: 1,
            ..generous
        },
    );

    // Charge the source itself beyond the aggregate work ceiling.  The
    // attempted cost must still expose those bytes to the caller.
    inspect(
        source,
        FormulaEnvelopeLimits {
            max_work: source_length.saturating_sub(1).max(1),
            ..generous
        },
    );

    // Leave only a small residual after a synthetic prior scan, then vary the
    // residual from input so both exact and over-limit field/work boundaries
    // are reached by the same envelope.
    let residual_fields = usize::from(selector % 4).saturating_add(1);
    let residual_work = source_length.saturating_add(usize::from(selector % 3));
    inspect(
        source,
        FormulaEnvelopeLimits {
            max_fields: residual_fields,
            max_work: residual_work.max(1),
            base_fields: residual_fields.saturating_sub(1),
            base_work: source_length.saturating_sub(1),
            ..generous
        },
    );

    // A byte ceiling below the bounded source length reaches the common wire
    // scanner's input limit while keeping all allocations finite.
    inspect(
        source,
        FormulaEnvelopeLimits {
            max_input_bytes: source_length.saturating_sub(1).max(1),
            ..generous
        },
    );
}

fn malformed_cases(base: &[u8]) -> Vec<Vec<u8>> {
    let mut cases = Vec::with_capacity(12);

    // Non-empty archives must carry the required root AST-node array.
    cases.push(vec![0x10, 0x01]);

    // Known field one with the wrong wire type.
    cases.push(vec![0x08, 0x01]);

    // A valid root followed by a duplicate singular host-column field.
    let mut duplicate_root = base.to_vec();
    append_varint_field(&mut duplicate_root, 2, 1);
    append_varint_field(&mut duplicate_root, 2, 2);
    cases.push(duplicate_root);

    // The AST node's required type is absent.
    let mut missing_node_type = Vec::new();
    append_fixed64_field(&mut missing_node_type, 4, 1.0f64.to_bits());
    let mut missing_array = Vec::new();
    append_bytes_field(&mut missing_array, 1, &missing_node_type);
    cases.push(formula_with_array(&missing_array));

    // A singular AST-node type appears twice.
    let mut duplicate_node_type = node(17);
    append_varint_field(&mut duplicate_node_type, 1, 17);
    cases.push(formula(&[duplicate_node_type]));

    // A known UTF-8 field contains an invalid byte sequence.
    cases.push(formula(&[string_node(&[0xff, 0xfe])]));

    // A known scalar varint uses a noncanonical two-byte representation.
    let mut noncanonical_scalar = node(16);
    append_key(&mut noncanonical_scalar, 2, 0);
    noncanonical_scalar.extend_from_slice(&[0x80, 0x00]);
    cases.push(formula(&[noncanonical_scalar]));

    // The local-reference message is missing required children.
    let mut incomplete_reference = Vec::new();
    append_varint_field(&mut incomplete_reference, 1, 1);
    let mut incomplete_node = node(27);
    append_bytes_field(&mut incomplete_node, 15, &incomplete_reference);
    cases.push(formula(&[incomplete_node]));

    // A length-delimited field declares more bytes than remain in the input.
    cases.push(vec![0x0a, 0x04, 0x01]);

    // Unknown fields remain opaque but make scalar admission conservative.
    let mut unknown_root = base.to_vec();
    append_varint_field(&mut unknown_root, 100, 7);
    cases.push(unknown_root);

    // Truncated prefixes cover every framing boundary without large inputs.
    for length in 0..=base.len().min(4) {
        cases.push(base[..length].to_vec());
    }

    cases
}

fn arbitrary_source(data: &[u8]) -> &[u8] {
    &data[..data.len().min(MAX_SOURCE_BYTES)]
}

struct Cursor<'source> {
    source: &'source [u8],
    offset: usize,
}

impl<'source> Cursor<'source> {
    const fn new(source: &'source [u8]) -> Self {
        Self { source, offset: 0 }
    }

    fn byte(&mut self) -> u8 {
        let value = self.source.get(self.offset).copied().unwrap_or(0);
        self.offset = self.offset.saturating_add(1);
        value
    }

    fn bounded(&mut self, maximum: usize) -> usize {
        if maximum == 0 {
            0
        } else {
            usize::from(self.byte()) % maximum
        }
    }

    fn u64(&mut self) -> u64 {
        u64::from_le_bytes([
            self.byte(),
            self.byte(),
            self.byte(),
            self.byte(),
            self.byte(),
            self.byte(),
            self.byte(),
            self.byte(),
        ])
    }
}

fuzz_target!(|data: &[u8]| {
    let source = arbitrary_source(data);
    let mut cursor = Cursor::new(data);
    let scalar = formula(&[number_node(finite_value(&mut cursor))]);
    let rich = rich_formula(&mut cursor);

    inspect_with_budget_variants(source, cursor.byte());
    inspect_with_budget_variants(&[], cursor.byte());
    inspect_with_budget_variants(&scalar, cursor.byte());
    inspect_with_budget_variants(&rich, cursor.byte());
    inspect_with_budget_variants(&deep_thunk(usize::from(cursor.byte() % 8)), cursor.byte());

    for malformed in malformed_cases(&scalar) {
        inspect_with_budget_variants(&malformed, cursor.byte());
    }
});
