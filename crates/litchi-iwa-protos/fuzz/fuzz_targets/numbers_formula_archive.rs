#![no_main]

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::numbers_formula_codec::{
    BinaryOperator, DecodeError, DecodeOptions, FormulaContext, FormulaNode, FormulaVisitor,
    LocalPrecedent, decode_formula_archive_with_visitor, inspect_formula_archive,
};
use litchi_iwa_protos::tsce::FormulaArchive;
use litchi_iwa_protos::tsce::ast_node_array_archive::{
    AstColumnCoordinateArchive, AstNodeArchive, AstRowCoordinateArchive,
};
use prost::Message as _;
use std::hint::black_box;

// The target deliberately skips inputs above the same finite profile used by
// the strict formula reader. Inputs are never truncated: every check below
// sees one unchanged source slice.
const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_NODES: usize = 2 * 1024;
const MAX_TEXT_BYTES: usize = 64 * 1024;
const MAX_RECURSION: u32 = 32;
const MAX_RETAINED_FACTS: usize = MAX_NODES;
const CONTROL_BYTES: usize = 8;

const OWNER: u32 = 7;
const HOST_ROW: u32 = 4;
const HOST_COLUMN: u32 = 5;
const TABLE_ROWS: u32 = 20;
const TABLE_COLUMNS: u32 = 20;

#[derive(Default)]
struct Facts {
    nodes: Vec<FormulaNode>,
    precedents: Vec<LocalPrecedent>,
    text_bytes: usize,
    text_values: usize,
    unsupported: usize,
    ranges: usize,
}

impl Facts {
    fn bounded() -> Self {
        Self {
            nodes: Vec::with_capacity(MAX_RETAINED_FACTS),
            precedents: Vec::with_capacity(MAX_RETAINED_FACTS),
            ..Self::default()
        }
    }

    fn assert_empty(&self) {
        assert!(
            self.nodes.is_empty(),
            "failed decode published formula nodes"
        );
        assert!(
            self.precedents.is_empty(),
            "failed decode published formula precedents"
        );
        assert_eq!(self.text_values, 0);
        assert_eq!(self.unsupported, 0);
        assert_eq!(self.ranges, 0);
    }
}

impl FormulaVisitor for Facts {
    fn visit_node(&mut self, node: FormulaNode) -> Result<(), DecodeError> {
        assert!(
            self.nodes.len() < MAX_RETAINED_FACTS,
            "formula callback exceeded the bounded node staging profile"
        );
        self.nodes.push(node);
        Ok(())
    }

    fn visit_precedent(&mut self, precedent: LocalPrecedent) -> Result<(), DecodeError> {
        assert!(
            self.precedents.len() < MAX_RETAINED_FACTS,
            "formula callback exceeded the bounded precedent staging profile"
        );
        self.precedents.push(precedent);
        Ok(())
    }

    fn visit_text(&mut self, value: &str) -> Result<(), DecodeError> {
        self.text_values = self.text_values.saturating_add(1);
        self.text_bytes = self.text_bytes.saturating_add(value.len());
        Ok(())
    }

    fn visit_unsupported_local(
        &mut self,
        _node: litchi_iwa_protos::numbers_formula_codec::UnsupportedLocal,
    ) -> Result<(), DecodeError> {
        self.unsupported = self.unsupported.saturating_add(1);
        Ok(())
    }

    fn visit_range(
        &mut self,
        _range: litchi_iwa_protos::numbers_formula_codec::FormulaWriteRange,
    ) -> Result<(), DecodeError> {
        self.ranges = self.ranges.saturating_add(1);
        Ok(())
    }
}

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };

    exercise_source(&source);

    // Direct mutations are useful even when libFuzzer has not yet discovered
    // a valid FormulaArchive. They stay independent so a malformed candidate
    // cannot prevent the aggregate, depth, or duplicate cases from running.
    match control(data, 0) % 6 {
        0 => exercise_source(&duplicate_root_formula()),
        1 => exercise_source(&duplicate_node_formula()),
        2 => exercise_source(&missing_node_type_formula()),
        3 => exercise_source(&wrong_wire_formula()),
        4 => exercise_source(&noncanonical_formula()),
        _ => exercise_source(&invalid_utf8_formula()),
    }

    if control(data, 1) & 1 == 0 {
        exercise_source(&aggregate_formula(data));
    } else {
        // The generated Prost value is intentionally used only as a bounded
        // oracle for a strict rejection. This exercises nested AST recursion
        // without allowing the permissive generated decoder to define the
        // acceptance policy.
        let deep = deep_formula(data);
        exercise_generated_decoder(&deep);
        exercise_source(&deep);
    }
});

fn context() -> FormulaContext {
    FormulaContext::new(OWNER, HOST_ROW, HOST_COLUMN, TABLE_ROWS, TABLE_COLUMNS)
}

fn options() -> DecodeOptions {
    DecodeOptions::new(
        MAX_INPUT_BYTES,
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        MAX_NODES,
        MAX_TEXT_BYTES,
    )
}

fn exercise_source(source: &[u8]) {
    assert!(
        source.len() <= MAX_INPUT_BYTES,
        "target constructed an input outside its finite profile"
    );
    let before = source.to_vec();
    let inspect = inspect_formula_archive(source, context(), options());

    assert_eq!(
        source,
        before.as_slice(),
        "formula inspection modified source"
    );

    let mut facts = Facts::bounded();
    let streamed = decode_formula_archive_with_visitor(source, context(), options(), &mut facts);

    assert_eq!(source, before.as_slice(), "formula decode modified source");
    assert_report_bounds(inspect.as_ref().ok());
    assert_report_bounds(streamed.as_ref().ok());

    match (inspect, streamed) {
        (Err(_), Err(_)) => {
            // A visitor is allowed to observe a valid prefix before a later
            // wire error, but this target's visitor cannot publish a partial
            // result. The strict reader currently performs its aggregate
            // callback admission before invoking it, so failed callbacks are
            // expected to leave the facts empty.
            facts.assert_empty();
        },
        (Ok(_), Err(error)) => {
            // `inspect_formula_archive` reports the first pass only. The
            // visitor entry point additionally authorizes a second callback
            // pass, so an exact first-pass boundary may be rejected solely by
            // the aggregate fields/work/text ceiling.
            assert!(
                error.resource_limit().is_some(),
                "strict inspector and visitor disagreed on formula validity"
            );
            facts.assert_empty();
        },
        (Err(_), Ok(_)) => {
            panic!("visitor accepted a FormulaArchive rejected by inspection");
        },
        (Ok(inspected), Ok(decoded)) => {
            assert_eq!(
                decoded, inspected,
                "formula preflight and callback reports diverged"
            );
            assert_eq!(facts.nodes.len(), decoded.node_count());
            assert_eq!(facts.precedents.len(), decoded.precedent_count());
            assert_eq!(facts.unsupported, 0);
            assert_eq!(facts.ranges, 0);
            assert_formula_parity(source, &facts);
            assert_precedent_parity(&facts);
        },
    }
}

fn assert_report_bounds(report: Option<&litchi_iwa_protos::numbers_formula_codec::DecodeReport>) {
    let Some(report) = report else {
        return;
    };
    assert!(report.bytes() <= MAX_INPUT_BYTES);
    assert!(report.fields() <= MAX_FIELDS);
    assert!(report.work() <= MAX_WORK_BYTES);
    assert!(report.max_depth() <= MAX_RECURSION);
    assert!(report.node_count() <= MAX_NODES);
    assert!(report.text_bytes() <= MAX_TEXT_BYTES);
    assert_eq!(report.allocations(), 0);
}

fn assert_formula_parity(source: &[u8], facts: &Facts) {
    let generated = FormulaArchive::decode(source)
        .unwrap_or_else(|error| panic!("strict acceptance disagreed with Prost: {error}"));
    assert_eq!(generated.ast_node_array.ast_node.len(), facts.nodes.len());

    for (node, expected) in facts
        .nodes
        .iter()
        .copied()
        .zip(generated.ast_node_array.ast_node.iter())
    {
        assert_node_matches(node, expected);
    }
}

fn assert_precedent_parity(facts: &Facts) {
    let mut index = 0usize;
    for node in facts.nodes.iter().copied() {
        let coordinate = match node {
            FormulaNode::LocalCell { coordinate, .. }
            | FormulaNode::CellReference { coordinate } => Some(coordinate),
            _ => None,
        };
        let Some(coordinate) = coordinate else {
            continue;
        };
        let precedent = facts
            .precedents
            .get(index)
            .unwrap_or_else(|| panic!("missing precedent callback for formula node {index}"));
        assert_eq!(precedent.owner(), OWNER);
        assert_eq!(precedent.coordinate().row(), coordinate.row());
        assert_eq!(precedent.coordinate().column(), coordinate.column());
        index += 1;
    }
    assert_eq!(index, facts.precedents.len());
}

fn assert_node_matches(node: FormulaNode, expected: &AstNodeArchive) {
    let kind = expected.ast_node_type;
    match node {
        FormulaNode::Binary(operator) => {
            assert_eq!(kind, binary_kind(operator));
        },
        FormulaNode::Negation => assert_eq!(kind, 13),
        FormulaNode::PlusSign => assert_eq!(kind, 14),
        FormulaNode::Percent => assert_eq!(kind, 15),
        FormulaNode::Function {
            identifier,
            argument_count,
        } => {
            assert_eq!(kind, 16);
            assert_eq!(expected.ast_function_node_index, Some(identifier));
            assert_eq!(
                expected.ast_function_node_num_args.unwrap_or_default(),
                argument_count
            );
        },
        FormulaNode::Number { bits } => {
            assert_eq!(kind, 17);
            assert_eq!(
                expected.ast_number_node_number.map(|value| value.to_bits()),
                Some(bits)
            );
        },
        FormulaNode::Boolean(value) => {
            assert_eq!(kind, 18);
            assert_eq!(expected.ast_boolean_node_boolean, Some(value));
        },
        FormulaNode::Empty => assert_eq!(kind, 22),
        FormulaNode::Token(value) => {
            assert_eq!(kind, 23);
            assert_eq!(expected.ast_token_node_boolean, Some(value));
        },
        FormulaNode::LocalCell {
            coordinate,
            row_is_sticky,
            column_is_sticky,
        } => {
            assert_eq!(kind, 27);
            let reference = expected
                .ast_local_cell_reference_node_reference
                .as_ref()
                .expect("strict local reference had no generated payload");
            assert_eq!(reference.row_handle, coordinate.row());
            assert_eq!(reference.column_handle, coordinate.column());
            assert_eq!(reference.row_is_sticky, row_is_sticky);
            assert_eq!(reference.column_is_sticky, column_is_sticky);
        },
        FormulaNode::CellReference { coordinate } => {
            assert_eq!(kind, 36);
            let column = expected
                .ast_column
                .as_ref()
                .expect("strict cell reference had no generated column");
            let row = expected
                .ast_row
                .as_ref()
                .expect("strict cell reference had no generated row");
            assert_axis(column, coordinate.column(), true);
            assert_axis(row, coordinate.row(), false);
        },
        FormulaNode::Colon => assert_eq!(kind, 29),
        FormulaNode::ColonWithUids => assert_eq!(kind, 45),
        FormulaNode::AppendWhitespace => assert_eq!(kind, 32),
        FormulaNode::PrependWhitespace => assert_eq!(kind, 33),
        FormulaNode::ResolvedCellReference { .. } | FormulaNode::ResolvedRange { .. } => {
            panic!("evaluator decoder emitted a resolved-owner node")
        },
    }
}

fn assert_axis(axis: &impl AxisValue, coordinate: u32, column: bool) {
    let expected = axis.coordinate();
    assert_eq!(expected, coordinate as i32);
    assert_eq!(axis.absolute(), false, "unexpected absolute {column} axis");
}

trait AxisValue {
    fn coordinate(&self) -> i32;
    fn absolute(&self) -> bool;
}

impl AxisValue for AstColumnCoordinateArchive {
    fn coordinate(&self) -> i32 {
        self.column
    }

    fn absolute(&self) -> bool {
        self.absolute.unwrap_or(false)
    }
}

impl AxisValue for AstRowCoordinateArchive {
    fn coordinate(&self) -> i32 {
        self.row
    }

    fn absolute(&self) -> bool {
        self.absolute.unwrap_or(false)
    }
}

fn binary_kind(operator: BinaryOperator) -> i32 {
    match operator {
        BinaryOperator::Add => 1,
        BinaryOperator::Subtract => 2,
        BinaryOperator::Multiply => 3,
        BinaryOperator::Divide => 4,
        BinaryOperator::Power => 5,
        BinaryOperator::Concatenate => 6,
        BinaryOperator::GreaterThan => 7,
        BinaryOperator::GreaterThanOrEqual => 8,
        BinaryOperator::LessThan => 9,
        BinaryOperator::LessThanOrEqual => 10,
        BinaryOperator::Equal => 11,
        BinaryOperator::NotEqual => 12,
    }
}

fn exercise_generated_decoder(source: &[u8]) {
    assert!(source.len() <= MAX_INPUT_BYTES);
    let before = source.to_vec();
    let generated = FormulaArchive::decode(source);
    assert_eq!(
        source,
        before.as_slice(),
        "generated formula decode modified source"
    );
    if let Ok(formula) = generated {
        black_box(formula.ast_node_array.ast_node.len());
    }
}

fn aggregate_formula(data: &[u8]) -> Vec<u8> {
    // Around this boundary the one-pass inspector can still succeed while
    // the visitor's two-pass aggregate authorization must refuse. The count
    // is bounded independently of fuzz input length.
    let count = 1_400 + usize::from(control(data, 2)) % 400;
    let mut nodes = Vec::with_capacity(count);
    for index in 0..count {
        nodes.push(number_node((index as f64) + 0.5));
    }
    formula(&nodes)
}

fn deep_formula(data: &[u8]) -> Vec<u8> {
    let depth = MAX_RECURSION as usize + 8 + usize::from(control(data, 3) % 8);
    let mut array = Vec::new();
    bytes_field(&mut array, 1, &number_node(1.0));
    for _ in 0..depth {
        let mut node = Vec::new();
        varint_field(&mut node, 1, 26);
        bytes_field(&mut node, 14, &array);
        let mut nested = Vec::new();
        bytes_field(&mut nested, 1, &node);
        array = nested;
    }
    let mut source = Vec::new();
    bytes_field(&mut source, 1, &array);
    source
}

fn duplicate_root_formula() -> Vec<u8> {
    let root = formula(&[number_node(1.0)]);
    let mut duplicate = root.clone();
    duplicate.extend_from_slice(&root);
    duplicate
}

fn duplicate_node_formula() -> Vec<u8> {
    let mut node = number_node(1.0);
    varint_field(&mut node, 1, 17);
    formula(&[node])
}

fn missing_node_type_formula() -> Vec<u8> {
    let mut node = Vec::new();
    fixed64_field(&mut node, 4, 1.0f64.to_bits());
    formula(&[node])
}

fn wrong_wire_formula() -> Vec<u8> {
    let mut node = Vec::new();
    varint_field(&mut node, 1, 17);
    varint_field(&mut node, 4, 0);
    formula(&[node])
}

fn noncanonical_formula() -> Vec<u8> {
    // The node type is a known scalar, so its overlong varint must be rejected
    // before generated Prost gets a chance to normalize it.
    let node = vec![0x08, 0x91, 0x00, 0x21, 0, 0, 0, 0, 0, 0, 0, 0];
    formula(&[node])
}

fn invalid_utf8_formula() -> Vec<u8> {
    let mut node = Vec::new();
    varint_field(&mut node, 1, 19);
    bytes_field(&mut node, 6, &[0xff]);
    formula(&[node])
}

fn number_node(value: f64) -> Vec<u8> {
    let mut node = Vec::new();
    varint_field(&mut node, 1, 17);
    fixed64_field(&mut node, 4, value.to_bits());
    node
}

fn formula(nodes: &[Vec<u8>]) -> Vec<u8> {
    let mut array = Vec::new();
    for node in nodes {
        bytes_field(&mut array, 1, node);
    }
    let mut source = Vec::new();
    bytes_field(&mut source, 1, &array);
    source
}

fn varint_field(output: &mut Vec<u8>, field: u32, value: u64) {
    put_varint(output, (u64::from(field) << 3) | 0);
    put_varint(output, value);
}

fn fixed64_field(output: &mut Vec<u8>, field: u32, value: u64) {
    put_varint(output, (u64::from(field) << 3) | 1);
    output.extend_from_slice(&value.to_le_bytes());
}

fn bytes_field(output: &mut Vec<u8>, field: u32, value: &[u8]) {
    put_varint(output, (u64::from(field) << 3) | 2);
    put_varint(
        output,
        u64::try_from(value.len()).expect("bounded formula payload length"),
    );
    output.extend_from_slice(value);
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

fn control(data: &[u8], index: usize) -> u8 {
    data.get(index % CONTROL_BYTES).copied().unwrap_or_default()
}
