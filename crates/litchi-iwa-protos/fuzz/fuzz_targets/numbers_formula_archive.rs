#![no_main]

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::numbers_formula_codec::{
    BinaryOperator, DecodeError, DecodeOptions, FormulaContext, FormulaNode, FormulaRenderEvent,
    FormulaRenderVisitor, FormulaVisitor, LocalPrecedent, decode_formula_archive_for_render,
    decode_formula_archive_with_visitor, inspect_formula_archive,
};
use litchi_iwa_protos::tsce::ast_node_array_archive::{
    AstColumnCoordinateArchive, AstNodeArchive, AstRowCoordinateArchive,
};
use litchi_iwa_protos::tsce::{AstNodeArrayArchive, FormulaArchive};
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
const MAX_RENDER_RECURSION: u32 = 64;
const MAX_RENDER_EVENTS: usize = MAX_FIELDS.saturating_mul(16);
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

// The compatibility renderer is intentionally exercised through a bounded,
// source-order event sink.  It stores offsets into the caller's bytes rather
// than cloning strings, which makes the borrow and source-atomicity contract
// observable without turning the fuzz target into an unbounded renderer.
const EVENT_BEGIN_ARRAY: u8 = 1;
const EVENT_END_ARRAY: u8 = 2;
const EVENT_THUNK_BEGIN: u8 = 3;
const EVENT_THUNK_END: u8 = 4;
const EVENT_BINARY: u8 = 5;
const EVENT_NEGATION: u8 = 6;
const EVENT_PLUS_SIGN: u8 = 7;
const EVENT_PERCENT: u8 = 8;
const EVENT_NUMBER: u8 = 9;
const EVENT_STRING: u8 = 10;
const EVENT_BOOLEAN: u8 = 11;
const EVENT_TOKEN: u8 = 12;
const EVENT_DATE: u8 = 13;
const EVENT_DURATION: u8 = 14;
const EVENT_EMPTY: u8 = 15;
const EVENT_FUNCTION: u8 = 16;
const EVENT_LIST: u8 = 17;
const EVENT_ARRAY: u8 = 18;
const EVENT_UNKNOWN_FUNCTION: u8 = 19;
const EVENT_CELL: u8 = 20;
const EVENT_LOCAL: u8 = 21;
const EVENT_CROSS: u8 = 22;
const EVENT_COLON: u8 = 23;
const EVENT_COLON_UIDS: u8 = 24;
const EVENT_COLON_TRACT: u8 = 25;
const EVENT_CATEGORY: u8 = 26;
const EVENT_REFERENCE_ERROR: u8 = 27;
const EVENT_APPEND_WHITESPACE: u8 = 28;
const EVENT_PREPEND_WHITESPACE: u8 = 29;
const EVENT_IGNORED: u8 = 30;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct RenderEventFact {
    tag: u8,
    first: u64,
    second: u64,
    text_start: usize,
    text_len: usize,
}

struct RenderFacts {
    source_start: usize,
    source_end: usize,
    events: Vec<RenderEventFact>,
    max_render_depth: u32,
}

impl RenderFacts {
    fn bounded(source: &[u8]) -> Self {
        let start = source.as_ptr() as usize;
        Self {
            source_start: start,
            source_end: start.saturating_add(source.len()),
            events: Vec::with_capacity(MAX_RENDER_EVENTS),
            max_render_depth: 0,
        }
    }

    fn assert_empty(&self) {
        assert!(
            self.events.is_empty(),
            "failed render decode published partial events"
        );
    }

    fn borrowed_range(&self, value: &str) -> (usize, usize) {
        if value.is_empty() {
            return (0, 0);
        }
        let start = value.as_ptr() as usize;
        let end = start
            .checked_add(value.len())
            .expect("borrowed render string pointer overflow");
        assert!(
            start >= self.source_start && end <= self.source_end,
            "render callback copied or synthesized string data"
        );
        (start - self.source_start, value.len())
    }

    fn push(
        &mut self,
        tag: u8,
        first: u64,
        second: u64,
        text: Option<&str>,
    ) -> Result<(), DecodeError> {
        if self.events.len() >= MAX_RENDER_EVENTS {
            return Err(DecodeError::allocation(MAX_RENDER_EVENTS));
        }
        let (text_start, text_len) = text
            .map(|value| self.borrowed_range(value))
            .unwrap_or((0, 0));
        self.events.push(RenderEventFact {
            tag,
            first,
            second,
            text_start,
            text_len,
        });
        Ok(())
    }
}

impl FormulaRenderVisitor for RenderFacts {
    fn visit(&mut self, event: FormulaRenderEvent<'_>) -> Result<(), DecodeError> {
        match event {
            FormulaRenderEvent::BeginArray { depth } => {
                self.max_render_depth = self.max_render_depth.max(depth);
                self.push(EVENT_BEGIN_ARRAY, u64::from(depth), 0, None)
            },
            FormulaRenderEvent::EndArray => self.push(EVENT_END_ARRAY, 0, 0, None),
            FormulaRenderEvent::ThunkBegin => self.push(EVENT_THUNK_BEGIN, 0, 0, None),
            FormulaRenderEvent::ThunkEnd => self.push(EVENT_THUNK_END, 0, 0, None),
            FormulaRenderEvent::Binary(operator) => {
                self.push(EVENT_BINARY, u64::from(binary_code(operator)), 0, None)
            },
            FormulaRenderEvent::Negation => self.push(EVENT_NEGATION, 0, 0, None),
            FormulaRenderEvent::PlusSign => self.push(EVENT_PLUS_SIGN, 0, 0, None),
            FormulaRenderEvent::Percent => self.push(EVENT_PERCENT, 0, 0, None),
            FormulaRenderEvent::Number { value } => {
                self.push(EVENT_NUMBER, value.to_bits(), 0, None)
            },
            FormulaRenderEvent::String(value) => self.push(EVENT_STRING, 0, 0, Some(value)),
            FormulaRenderEvent::Boolean(value) => {
                self.push(EVENT_BOOLEAN, u64::from(value), 0, None)
            },
            FormulaRenderEvent::Token(value) => self.push(EVENT_TOKEN, u64::from(value), 0, None),
            FormulaRenderEvent::Date { value } => self.push(EVENT_DATE, value.to_bits(), 0, None),
            FormulaRenderEvent::Duration { value } => {
                self.push(EVENT_DURATION, value.to_bits(), 0, None)
            },
            FormulaRenderEvent::EmptyArgument => self.push(EVENT_EMPTY, 0, 0, None),
            FormulaRenderEvent::Function {
                identifier,
                argument_count,
            } => self.push(
                EVENT_FUNCTION,
                u64::from(identifier),
                u64::from(argument_count),
                None,
            ),
            FormulaRenderEvent::List { argument_count } => {
                self.push(EVENT_LIST, u64::from(argument_count), 0, None)
            },
            FormulaRenderEvent::Array { columns, rows } => {
                self.push(EVENT_ARRAY, u64::from(columns), u64::from(rows), None)
            },
            FormulaRenderEvent::UnknownFunction {
                name,
                argument_count,
            } => self.push(EVENT_UNKNOWN_FUNCTION, u64::from(argument_count), 0, name),
            FormulaRenderEvent::CellReference(_) => self.push(EVENT_CELL, 0, 0, None),
            FormulaRenderEvent::LocalCellReference(_) => self.push(EVENT_LOCAL, 0, 0, None),
            FormulaRenderEvent::CrossTableCellReference(_) => self.push(EVENT_CROSS, 0, 0, None),
            FormulaRenderEvent::Colon => self.push(EVENT_COLON, 0, 0, None),
            FormulaRenderEvent::ColonWithUids => self.push(EVENT_COLON_UIDS, 0, 0, None),
            FormulaRenderEvent::ColonTract(_) => self.push(EVENT_COLON_TRACT, 0, 0, None),
            FormulaRenderEvent::CategoryReference(_) => self.push(EVENT_CATEGORY, 0, 0, None),
            FormulaRenderEvent::ReferenceError => self.push(EVENT_REFERENCE_ERROR, 0, 0, None),
            FormulaRenderEvent::AppendWhitespace => self.push(EVENT_APPEND_WHITESPACE, 0, 0, None),
            FormulaRenderEvent::PrependWhitespace => {
                self.push(EVENT_PREPEND_WHITESPACE, 0, 0, None)
            },
            FormulaRenderEvent::Ignored { raw_node_type } => {
                self.push(EVENT_IGNORED, u64::from(raw_node_type), 0, None)
            },
        }
    }
}

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };

    exercise_source(&source);
    exercise_render_source(&source);

    // Direct mutations are useful even when libFuzzer has not yet discovered
    // a valid FormulaArchive. They stay independent so a malformed candidate
    // cannot prevent the aggregate, depth, or duplicate cases from running.
    match control(data, 0) % 6 {
        0 => {
            let candidate = duplicate_root_formula();
            exercise_source(&candidate);
            exercise_render_source(&candidate);
        },
        1 => {
            let candidate = duplicate_node_formula();
            exercise_source(&candidate);
            exercise_render_source(&candidate);
        },
        2 => {
            let candidate = missing_node_type_formula();
            exercise_source(&candidate);
            exercise_render_source(&candidate);
        },
        3 => {
            let candidate = wrong_wire_formula();
            exercise_source(&candidate);
            exercise_render_source(&candidate);
        },
        4 => {
            let candidate = noncanonical_formula();
            exercise_source(&candidate);
            exercise_render_source(&candidate);
        },
        _ => {
            let candidate = invalid_utf8_formula();
            exercise_source(&candidate);
            exercise_render_source(&candidate);
        },
    }

    if control(data, 1) & 1 == 0 {
        let aggregate = aggregate_formula(data);
        exercise_source(&aggregate);
        exercise_render_source(&aggregate);
    } else {
        // The generated Prost value is intentionally used only as a bounded
        // oracle for a strict rejection. This exercises nested AST recursion
        // without allowing the permissive generated decoder to define the
        // acceptance policy.
        let deep = deep_formula(data);
        exercise_generated_decoder(&deep);
        exercise_source(&deep);
        exercise_render_source(&deep);
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

fn render_options_with(
    max_bytes: usize,
    max_fields: usize,
    max_work: usize,
    wire_recursion: u32,
    max_nodes: usize,
    max_text_bytes: usize,
    render_recursion: u32,
) -> DecodeOptions {
    DecodeOptions::new(
        max_bytes,
        max_fields,
        max_work,
        wire_recursion,
        max_nodes,
        max_text_bytes,
    )
    .with_opaque_unknown_fields(true)
    .with_unknown_functions(true)
    .with_render_recursion_limit(render_recursion)
}

fn render_options() -> DecodeOptions {
    render_options_with(
        MAX_INPUT_BYTES,
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        MAX_NODES,
        MAX_TEXT_BYTES,
        MAX_RENDER_RECURSION,
    )
}

fn exercise_render_source(source: &[u8]) {
    assert!(
        source.len() <= MAX_INPUT_BYTES,
        "render target constructed an input outside its finite profile"
    );
    let before = source.to_vec();
    let probe = decode_formula_archive_for_render(source, context(), render_options(), &mut ());
    assert_eq!(
        source,
        before.as_slice(),
        "formula render inspection modified source"
    );

    let Ok(inspected) = probe else {
        // The render decoder performs complete callback admission before it
        // invokes the sink. Re-run malformed inputs with the bounded sink to
        // make that no-partial-publication property explicit.
        let mut facts = RenderFacts::bounded(source);
        assert!(
            decode_formula_archive_for_render(source, context(), render_options(), &mut facts)
                .is_err()
        );
        facts.assert_empty();
        assert_eq!(source, before.as_slice(), "failed render modified source");
        return;
    };

    assert_render_report_bounds(&inspected);
    let mut facts = RenderFacts::bounded(source);
    let decoded =
        decode_formula_archive_for_render(source, context(), render_options(), &mut facts)
            .unwrap_or_else(|error| panic!("render probe and callback pass disagreed: {error:?}"));
    assert_eq!(decoded, inspected);
    assert_render_report_bounds(&decoded);
    assert!(
        !facts.events.is_empty(),
        "successful render emitted no events"
    );
    assert!(facts.max_render_depth <= MAX_RENDER_RECURSION);
    assert_generated_render_parity(source, &facts, &decoded);
    assert_render_limit_probes(source, &decoded, facts.max_render_depth);
    assert_eq!(source, before.as_slice(), "formula render modified source");
}

fn assert_render_report_bounds(report: &litchi_iwa_protos::numbers_formula_codec::DecodeReport) {
    assert!(report.bytes() <= MAX_INPUT_BYTES);
    assert!(report.fields() <= MAX_FIELDS);
    assert!(report.work() <= MAX_WORK_BYTES);
    assert!(report.max_depth() <= MAX_RECURSION);
    assert!(report.node_count() <= MAX_NODES);
    assert!(report.text_bytes() <= MAX_TEXT_BYTES);
    assert_eq!(report.allocations(), 0);
}

fn assert_render_limit_probes(
    source: &[u8],
    report: &litchi_iwa_protos::numbers_formula_codec::DecodeReport,
    render_depth: u32,
) {
    let exact = |bytes, fields, work, wire_depth, nodes, text, logical_depth| {
        render_options_with(bytes, fields, work, wire_depth, nodes, text, logical_depth)
    };
    let probe = |options: DecodeOptions| {
        let before = source.to_vec();
        let mut facts = RenderFacts::bounded(source);
        let error = decode_formula_archive_for_render(source, context(), options, &mut facts)
            .expect_err("max-minus-one render limit unexpectedly succeeded");
        assert!(
            error.resource_limit().is_some(),
            "render max-minus-one did not return a typed resource limit: {error:?}"
        );
        facts.assert_empty();
        assert_eq!(source, before.as_slice(), "render limit modified source");
    };

    let bytes = report.bytes();
    if bytes > 0 {
        probe(exact(
            bytes - 1,
            report.fields(),
            report.work(),
            report.max_depth(),
            report.node_count(),
            report.text_bytes(),
            render_depth,
        ));
    }
    let fields = report.fields();
    if fields > 0 {
        probe(exact(
            bytes,
            fields - 1,
            report.work(),
            report.max_depth(),
            report.node_count(),
            report.text_bytes(),
            render_depth,
        ));
    }
    let work = report.work();
    if work > 0 {
        probe(exact(
            bytes,
            fields,
            work - 1,
            report.max_depth(),
            report.node_count(),
            report.text_bytes(),
            render_depth,
        ));
    }
    let text = report.text_bytes();
    if text > 0 {
        probe(exact(
            bytes,
            fields,
            report.work(),
            report.max_depth(),
            report.node_count(),
            text - 1,
            render_depth,
        ));
    }
    let wire_depth = report.max_depth();
    if wire_depth > 0 {
        probe(exact(
            bytes,
            fields,
            report.work(),
            wire_depth - 1,
            report.node_count(),
            text,
            render_depth,
        ));
    }
    let nodes = report.node_count();
    if nodes > 0 {
        probe(exact(
            bytes,
            fields,
            report.work(),
            report.max_depth(),
            nodes - 1,
            text,
            render_depth,
        ));
    }
    if render_depth > 0 {
        probe(exact(
            bytes,
            fields,
            report.work(),
            report.max_depth(),
            nodes,
            text,
            render_depth - 1,
        ));
    }
}

#[derive(Default)]
struct GeneratedRenderSummary {
    nodes: usize,
    nested_arrays: usize,
    numbers: usize,
    strings: usize,
    dates: usize,
    durations: usize,
    functions: usize,
    lists: usize,
    arrays: usize,
    unknown_functions: usize,
}

fn assert_generated_render_parity(
    source: &[u8],
    facts: &RenderFacts,
    report: &litchi_iwa_protos::numbers_formula_codec::DecodeReport,
) {
    let Ok(generated) = FormulaArchive::decode(source) else {
        // Opaque unknown fields are intentionally accepted by the render
        // adapter even when a future generated Prost schema rejects them.
        // The strict codec report and borrowed event sink remain authoritative
        // for that forward-compatible case.
        return;
    };
    let mut summary = GeneratedRenderSummary::default();
    summarize_generated_array(&generated.ast_node_array, &mut summary);
    assert_eq!(summary.nodes, report.node_count());
    assert_eq!(
        count_render_events(facts, EVENT_BEGIN_ARRAY),
        summary.nested_arrays + 1
    );
    assert_eq!(
        count_render_events(facts, EVENT_END_ARRAY),
        summary.nested_arrays + 1
    );
    assert_eq!(
        count_render_events(facts, EVENT_THUNK_BEGIN),
        count_kind_with_thunk(&generated.ast_node_array)
    );
    assert_eq!(
        count_render_events(facts, EVENT_THUNK_END),
        count_kind_with_thunk(&generated.ast_node_array)
    );
    assert_eq!(count_render_events(facts, EVENT_NUMBER), summary.numbers);
    assert_eq!(count_render_events(facts, EVENT_STRING), summary.strings);
    assert_eq!(count_render_events(facts, EVENT_DATE), summary.dates);
    assert_eq!(
        count_render_events(facts, EVENT_DURATION),
        summary.durations
    );
    assert_eq!(
        count_render_events(facts, EVENT_FUNCTION),
        summary.functions
    );
    assert_eq!(count_render_events(facts, EVENT_LIST), summary.lists);
    assert_eq!(count_render_events(facts, EVENT_ARRAY), summary.arrays);
    assert_eq!(
        count_render_events(facts, EVENT_UNKNOWN_FUNCTION),
        summary.unknown_functions
    );
}

fn summarize_generated_array(array: &AstNodeArrayArchive, summary: &mut GeneratedRenderSummary) {
    for node in &array.ast_node {
        summary.nodes = summary
            .nodes
            .checked_add(1)
            .expect("bounded generated formula node count");
        match node.ast_node_type {
            16 if node.ast_function_node_index.is_some() => summary.functions += 1,
            17 if node.ast_number_node_number.is_some() => summary.numbers += 1,
            19 if node.ast_string_node_string.is_some() => summary.strings += 1,
            20 if node.ast_date_node_date_num.is_some() => summary.dates += 1,
            21 if node.ast_duration_node_unit_num.is_some() => summary.durations += 1,
            24 => summary.arrays += 1,
            25 if node.ast_list_node_num_args.is_some() => summary.lists += 1,
            31 => summary.unknown_functions += 1,
            _ => {},
        }
        if let Some(nested) = &node.ast_thunk_node_array {
            summary.nested_arrays += 1;
            summarize_generated_array(nested, summary);
        }
    }
}

fn count_kind_with_thunk(array: &AstNodeArrayArchive) -> usize {
    array
        .ast_node
        .iter()
        .map(|node| {
            usize::from(node.ast_node_type == 26 && node.ast_thunk_node_array.is_some())
                + node
                    .ast_thunk_node_array
                    .as_ref()
                    .map(count_kind_with_thunk)
                    .unwrap_or_default()
        })
        .sum()
}

fn count_render_events(facts: &RenderFacts, tag: u8) -> usize {
    facts.events.iter().filter(|event| event.tag == tag).count()
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
            assert_formula_report_parity(inspected, decoded);
            assert_eq!(facts.nodes.len(), decoded.node_count());
            assert_eq!(facts.precedents.len(), decoded.precedent_count());
            assert_eq!(facts.unsupported, decoded.unsupported_local_count());
            assert_eq!(facts.ranges, decoded.range_count());
            assert_formula_parity(source, &facts);
            assert_precedent_parity(&facts);
        },
    }
}

fn assert_formula_report_parity(
    inspected: litchi_iwa_protos::numbers_formula_codec::DecodeReport,
    decoded: litchi_iwa_protos::numbers_formula_codec::DecodeReport,
) {
    // The visitor entry point performs the same strict walk twice: once for
    // callback admission and once for publication. Fields, wire work, and
    // text are aggregate costs across both passes; semantic counts are
    // admitted during preflight and therefore remain single-pass values.
    assert_eq!(decoded.bytes(), inspected.bytes());
    assert_eq!(
        decoded.fields(),
        inspected
            .fields()
            .checked_mul(2)
            .expect("bounded formula field report must not overflow")
    );
    assert_eq!(
        decoded.work(),
        inspected
            .work()
            .checked_mul(2)
            .expect("bounded formula work report must not overflow")
    );
    assert_eq!(decoded.max_depth(), inspected.max_depth());
    assert_eq!(
        decoded.text_bytes(),
        inspected
            .text_bytes()
            .checked_mul(2)
            .expect("bounded formula text report must not overflow")
    );
    assert_eq!(decoded.node_count(), inspected.node_count());
    assert_eq!(decoded.precedent_count(), inspected.precedent_count());
    assert_eq!(decoded.range_count(), inspected.range_count());
    assert_eq!(
        decoded.evaluator_supported(),
        inspected.evaluator_supported()
    );
    assert_eq!(
        decoded.unsupported_local_count(),
        inspected.unsupported_local_count()
    );
    assert_eq!(decoded.allocations(), inspected.allocations());
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
    // Range endpoints are reported as precedents without a corresponding
    // scalar `FormulaNode`; their count is checked against the decoder report
    // by the caller, so only coordinate-bearing nodes are paired here.
    assert!(index <= facts.precedents.len());
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
        FormulaNode::LocalCellReference {
            coordinate,
            row_is_sticky,
            column_is_sticky,
        } => {
            assert_eq!(kind, 36);
            let column = expected
                .ast_column
                .as_ref()
                .expect("strict local cell reference had no generated column");
            let row = expected
                .ast_row
                .as_ref()
                .expect("strict local cell reference had no generated row");
            assert_eq!(column.coordinate(), coordinate.column() as i32);
            assert_eq!(column.absolute(), column_is_sticky);
            assert_eq!(row.coordinate(), coordinate.row() as i32);
            assert_eq!(row.absolute(), row_is_sticky);
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
        FormulaNode::LocalRange { .. } => assert_eq!(kind, 67),
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
    assert!(!axis.absolute(), "unexpected absolute {column} axis");
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

fn binary_code(operator: BinaryOperator) -> u8 {
    u8::try_from(binary_kind(operator)).expect("formula binary operator has a bounded kind")
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
    put_varint(output, u64::from(field) << 3);
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
