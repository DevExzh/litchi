//! Independent wire-contract tests for selected Numbers table sidecars.
//!
//! Every envelope in this file is handwritten protobuf wire.  The fixtures
//! exercise the strict storage and comment codecs through the public sidecar
//! reader, while the budget below records the aggregate passes and refuses
//! selected fallible copies.  This keeps the tests independent of generated
//! package models and makes missing, empty, and malformed sidecars explicit.

use litchi_iwa_common::formula::render::FormulaRenderBudget;
use litchi_iwa_common::table::cell::value::Value;
use litchi_iwa_protos::comment_storage_codec::{self, DecodeLimit as CommentDecodeLimit};
use litchi_iwa_protos::numbers_formula_codec;
use litchi_iwa_protos::numbers_table_cell_storage_codec as storage;
use litchi_numbers_wire::cell_value::{CellValueSource, ValueSource};
use litchi_numbers_wire::formula_envelope::{
    AttemptedFormulaEnvelopeCost, FormulaEnvelopeLimits, FormulaEnvelopeReport,
};
use litchi_numbers_wire::formula_render::{
    FormulaEventRenderBudget, FormulaTablePrefix, ReferenceResolver,
};
use litchi_numbers_wire::table_data_list::{
    Message, TABLE_DATA_LIST_MESSAGE_KIND, TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND,
};
use litchi_numbers_wire::table_sidecars::{
    COMMENT_STORAGE_MESSAGE_KIND, CellSidecarResolver, SidecarAllocation, SidecarIssue,
    SidecarKind, SidecarReadBudget, SidecarReference, SidecarTables, SidecarValue,
    read_comment_storage, read_sidecar_list, render_formula,
};

fn varint(mut value: u64) -> Vec<u8> {
    let mut output = Vec::new();
    loop {
        let byte = u8::try_from(value & 0x7f).expect("varint chunk fits in a byte");
        value >>= 7;
        if value == 0 {
            output.push(byte);
            return output;
        }
        output.push(byte | 0x80);
    }
}

fn field_key(number: u32, wire_type: u8) -> Vec<u8> {
    varint((u64::from(number) << 3) | u64::from(wire_type))
}

fn field_varint(number: u32, value: u64) -> Vec<u8> {
    let mut output = field_key(number, 0);
    output.extend(varint(value));
    output
}

fn field_fixed64(number: u32, value: u64) -> Vec<u8> {
    let mut output = field_key(number, 1);
    output.extend(value.to_le_bytes());
    output
}

fn field_bytes(number: u32, payload: &[u8]) -> Vec<u8> {
    let mut output = field_key(number, 2);
    output.extend(varint(
        u64::try_from(payload.len()).expect("fixture payload length fits in u64"),
    ));
    output.extend_from_slice(payload);
    output
}

fn reference(identifier: u64) -> Vec<u8> {
    field_varint(1, identifier)
}

fn list_root(list_type: SidecarKind, entries: &[Vec<u8>]) -> Vec<u8> {
    list_root_with_segments(list_type, entries, &[])
}

fn list_root_with_segments(
    list_type: SidecarKind,
    entries: &[Vec<u8>],
    segments: &[u64],
) -> Vec<u8> {
    let mut output = field_varint(
        1,
        u64::try_from(list_type.list_type()).expect("positive type"),
    );
    output.extend(field_varint(2, 91));
    for entry in entries {
        output.extend(field_bytes(3, entry));
    }
    for segment in segments {
        output.extend(field_bytes(4, &reference(*segment)));
    }
    output
}

fn sidecar_segment(
    list_type: SidecarKind,
    location: u32,
    length: u32,
    entries: &[Vec<u8>],
) -> Vec<u8> {
    let mut range = field_varint(1, u64::from(location));
    range.extend(field_varint(2, u64::from(length)));
    let mut output = field_varint(
        1,
        u64::try_from(list_type.list_type()).expect("positive type"),
    );
    output.extend(field_bytes(2, &range));
    for entry in entries {
        output.extend(field_bytes(3, entry));
    }
    output
}

fn list_entry(key: u32, ref_count: u32, payload_field: u32, payload: &[u8]) -> Vec<u8> {
    let mut output = field_varint(1, u64::from(key));
    output.extend(field_varint(2, u64::from(ref_count)));
    output.extend(field_bytes(payload_field, payload));
    output
}

fn string_entry(key: u32, value: &str) -> Vec<u8> {
    list_entry(key, 1, 3, value.as_bytes())
}

fn formula_node(kind: u64) -> Vec<u8> {
    field_varint(1, kind)
}

fn number_formula() -> Vec<u8> {
    // FormulaArchive.ast_node_array.ast_node[0] = NumberNode.  NumberNode's
    // scalar payload is the fixed64 field 4; retaining it makes this a real
    // scalar formula rather than the compatibility FORMULA() placeholder.
    let mut node = formula_node(17);
    node.extend(field_fixed64(4, 0.0_f64.to_bits()));
    field_bytes(1, &field_bytes(1, &node))
}

fn formula_entry(key: u32, source: &[u8]) -> Vec<u8> {
    list_entry(key, 1, 5, source)
}

fn rich_text_entry(key: u32, identifier: u64) -> Vec<u8> {
    list_entry(key, 1, 9, &reference(identifier))
}

fn comment_entry(key: u32, ref_count: u32, identifier: u64) -> Vec<u8> {
    list_entry(key, ref_count, 10, &reference(identifier))
}

fn malformed_later_entry(mut source: Vec<u8>) -> Vec<u8> {
    // The outer repeated field is framed correctly, but the nested entry is
    // truncated.  The strict full-list pass must reject it after validating
    // the preceding entry.
    source.extend(field_bytes(3, &[0x80]));
    source
}

fn unknown_field(mut source: Vec<u8>) -> Vec<u8> {
    source.extend(field_bytes(100, &[0xde, 0xad, 0xbe, 0xef]));
    source
}

fn comment_storage(
    text: Option<&str>,
    author: Option<u64>,
    replies: &[u64],
    uuid: Option<(u64, u64)>,
) -> Vec<u8> {
    let mut output = Vec::new();
    if let Some(text) = text {
        output.extend(field_bytes(1, text.as_bytes()));
    }
    if let Some(author) = author {
        output.extend(field_bytes(3, &reference(author)));
    }
    for reply in replies {
        output.extend(field_bytes(4, &reference(*reply)));
    }
    if let Some((lower, upper)) = uuid {
        let mut value = field_varint(1, lower);
        value.extend(field_varint(2, upper));
        output.extend(field_bytes(5, &value));
    }
    output
}

fn comment_storage_with_unknown(text: &str) -> Vec<u8> {
    unknown_field(comment_storage(Some(text), Some(77), &[78], Some((9, 10))))
}

fn comment_storage_options(
    source: &[u8],
    max_references: usize,
) -> comment_storage_codec::DecodeOptions {
    comment_storage_codec::DecodeOptions::new(
        source.len().max(1),
        256,
        source.len().saturating_mul(64).max(1),
        16,
        max_references,
        64 * 1024,
    )
}

fn storage_options(source: &[u8]) -> storage::DecodeOptions {
    storage::DecodeOptions::new(
        source.len().max(1),
        4096,
        source.len().saturating_mul(64).max(1),
        64,
        4096,
        64 * 1024,
    )
}

#[derive(Debug, PartialEq, Eq)]
enum HarnessError {
    Issue(SidecarIssue),
    Refused {
        kind: SidecarKind,
        target: SidecarAllocation,
        amount: usize,
    },
    Render(String),
    OutputLimit {
        observed: usize,
        maximum: usize,
    },
    Allocation {
        resource: &'static str,
        amount: usize,
    },
}

#[derive(Debug, Default)]
struct Budget {
    max_entries: usize,
    refuse: Option<(SidecarKind, SidecarAllocation)>,
    list_reports: Vec<(SidecarKind, storage::DecodeReport)>,
    probe_reports: Vec<(SidecarKind, storage::DecodeReport)>,
    comment_reports: Vec<comment_storage_codec::DecodeReport>,
    formula_reports: Vec<numbers_formula_codec::DecodeReport>,
    formula_envelope_reports: Vec<FormulaEnvelopeReport>,
    formula_envelope_failed_costs: Vec<AttemptedFormulaEnvelopeCost>,
    retained: Vec<(SidecarKind, SidecarAllocation, usize)>,
    entries: Vec<(SidecarKind, u32, usize)>,
    max_comment_references: usize,
    maximum_output: usize,
    charged_output: usize,
}

impl Budget {
    fn generous() -> Self {
        Self {
            max_entries: 4096,
            max_comment_references: 4096,
            maximum_output: usize::MAX,
            ..Self::default()
        }
    }
}

impl SidecarReadBudget for Budget {
    type Error = HarnessError;

    fn list_options(
        &mut self,
        _kind: SidecarKind,
        source: &[u8],
    ) -> Result<storage::DecodeOptions, Self::Error> {
        Ok(storage_options(source))
    }

    fn charge_list_report(
        &mut self,
        kind: SidecarKind,
        report: storage::DecodeReport,
    ) -> Result<(), Self::Error> {
        self.list_reports.push((kind, report));
        Ok(())
    }

    fn charge_list_probe_report(
        &mut self,
        kind: SidecarKind,
        report: storage::DecodeReport,
    ) -> Result<(), Self::Error> {
        self.probe_reports.push((kind, report));
        Ok(())
    }

    fn charge_retained(
        &mut self,
        kind: SidecarKind,
        target: SidecarAllocation,
        amount: usize,
    ) -> Result<(), Self::Error> {
        if self.refuse == Some((kind, target)) {
            return Err(HarnessError::Refused {
                kind,
                target,
                amount,
            });
        }
        self.retained.push((kind, target, amount));
        Ok(())
    }

    fn charge_entry(
        &mut self,
        kind: SidecarKind,
        key: u32,
        source_bytes: usize,
    ) -> Result<(), Self::Error> {
        self.entries.push((kind, key, source_bytes));
        Ok(())
    }

    fn max_entries(&self, _kind: SidecarKind) -> usize {
        self.max_entries
    }

    fn map_issue(&mut self, issue: SidecarIssue) -> Self::Error {
        HarnessError::Issue(issue)
    }

    fn comment_options(
        &mut self,
        source: &[u8],
    ) -> Result<comment_storage_codec::DecodeOptions, Self::Error> {
        Ok(comment_storage_options(source, self.max_comment_references))
    }

    fn charge_comment_report(
        &mut self,
        report: comment_storage_codec::DecodeReport,
    ) -> Result<(), Self::Error> {
        self.comment_reports.push(report);
        Ok(())
    }

    fn formula_options(
        &mut self,
        source: &[u8],
    ) -> Result<numbers_formula_codec::DecodeOptions, Self::Error> {
        Ok(numbers_formula_codec::DecodeOptions::new(
            source.len().max(1),
            4096,
            source.len().saturating_mul(64).max(1),
            32,
            4096,
            64 * 1024,
        )
        .with_opaque_unknown_fields(true)
        .with_unknown_functions(true)
        .with_render_recursion_limit(256))
    }

    fn charge_formula_report(
        &mut self,
        report: numbers_formula_codec::DecodeReport,
    ) -> Result<(), Self::Error> {
        self.formula_reports.push(report);
        Ok(())
    }

    fn formula_envelope_limits(
        &mut self,
        _source: &[u8],
    ) -> Result<FormulaEnvelopeLimits, Self::Error> {
        Ok(FormulaEnvelopeLimits {
            max_fields: 4096,
            max_input_bytes: 64 * 1024,
            max_work: 64 * 1024,
            base_fields: 0,
            base_work: 0,
        })
    }

    fn charge_formula_envelope_report(
        &mut self,
        report: FormulaEnvelopeReport,
    ) -> Result<(), Self::Error> {
        self.formula_envelope_reports.push(report);
        Ok(())
    }

    fn retain_formula_envelope_cost(&mut self, cost: AttemptedFormulaEnvelopeCost) {
        self.formula_envelope_failed_costs.push(cost);
    }

    fn map_formula_envelope_error(&mut self, error: litchi_iwa_common::Error) -> Self::Error {
        HarnessError::Render(error.to_string())
    }
}

impl FormulaRenderBudget for Budget {
    type Error = HarnessError;

    fn output_limit(&self, observed: usize) -> Self::Error {
        HarnessError::OutputLimit {
            observed,
            maximum: self.maximum_output,
        }
    }

    fn allocation(&self, resource: &'static str, amount: usize) -> Self::Error {
        HarnessError::Allocation { resource, amount }
    }

    fn invalid(&self, message: &'static str) -> Self::Error {
        HarnessError::Render(message.to_owned())
    }

    fn check(&self, amount: usize) -> Result<(), Self::Error> {
        if amount > self.maximum_output {
            Err(self.output_limit(amount))
        } else {
            Ok(())
        }
    }

    fn check_structure(&self, _nodes: usize, _parts: usize) -> Result<(), Self::Error> {
        Ok(())
    }

    fn charge(&mut self, amount: usize) -> Result<(), Self::Error> {
        let observed = self
            .charged_output
            .checked_add(amount)
            .ok_or_else(|| self.output_limit(usize::MAX))?;
        if observed > self.maximum_output {
            return Err(self.output_limit(observed));
        }
        self.charged_output = observed;
        Ok(())
    }
}

impl FormulaEventRenderBudget for Budget {
    fn check_render_depth(&self, _depth: usize) -> Result<(), Self::Error> {
        Ok(())
    }

    fn parse_error(&self, message: String) -> Self::Error {
        HarnessError::Render(message)
    }

    fn invalid_format(&self, message: String) -> Self::Error {
        HarnessError::Render(message)
    }
}

#[derive(Debug, Default)]
struct Resolver;

impl ReferenceResolver for Resolver {
    fn table_prefix(
        &self,
        _id: &numbers_formula_codec::FormulaRenderCfuuid,
    ) -> Option<FormulaTablePrefix<'_>> {
        None
    }

    fn category_name(
        &self,
        _id: litchi_numbers_wire::formula_render::FormulaCategoryId,
    ) -> Option<&str> {
        None
    }

    fn function_name(&self, _index: u32) -> Option<&str> {
        None
    }
}

#[derive(Debug, PartialEq, Eq)]
enum ResolverError {
    Missing { kind: SidecarKind, key: u32 },
    InvalidScalar,
}

#[derive(Debug, Default)]
struct CellResolver {
    rich_text_calls: Vec<SidecarReference>,
    formula_calls: Vec<(u32, usize)>,
    comment_calls: Vec<SidecarReference>,
}

impl CellSidecarResolver for CellResolver {
    type Error = ResolverError;

    fn rich_text(&mut self, reference: SidecarReference) -> Result<String, Self::Error> {
        self.rich_text_calls.push(reference);
        Ok("rich text".to_owned())
    }

    fn formula(
        &mut self,
        key: u32,
        source: &[u8],
        _row: u32,
        _column: u32,
    ) -> Result<String, Self::Error> {
        self.formula_calls.push((key, source.len()));
        Ok("=formula".to_owned())
    }

    fn comment(
        &mut self,
        reference: SidecarReference,
    ) -> Result<litchi_iwa_common::table::read::Comment, Self::Error> {
        self.comment_calls.push(reference);
        Ok(litchi_iwa_common::table::read::Comment::new("comment"))
    }

    fn retain_text(
        &mut self,
        _kind: SidecarKind,
        _key: u32,
        source: &str,
    ) -> Result<String, Self::Error> {
        Ok(source.to_owned())
    }

    fn missing(&mut self, kind: SidecarKind, key: u32) -> Self::Error {
        ResolverError::Missing { kind, key }
    }

    fn invalid_scalar(&mut self) -> Self::Error {
        ResolverError::InvalidScalar
    }
}

fn read_list(
    kind: SidecarKind,
    source: &[u8],
    budget: &mut Budget,
) -> Result<litchi_numbers_wire::table_sidecars::SidecarList, HarnessError> {
    read_sidecar_list(
        kind,
        700,
        [Message::new(TABLE_DATA_LIST_MESSAGE_KIND, source)],
        |_, _| Ok(None::<Vec<Message<'_>>>),
        budget,
    )
}

#[test]
fn all_supported_sidecar_list_types_project_their_canonical_payload() {
    let formula = number_formula();
    let cases = [
        (SidecarKind::Strings, string_entry(7, "seven")),
        (SidecarKind::Formulas, formula_entry(7, &formula)),
        (SidecarKind::FormulaErrors, string_entry(7, "#VALUE!")),
        (SidecarKind::RichTextPayloads, rich_text_entry(7, 7007)),
        (SidecarKind::Comments, comment_entry(7, 1, 8008)),
    ];

    for (kind, entry) in cases {
        let source = list_root(kind, &[entry]);
        let mut budget = Budget::generous();
        let list = read_list(kind, &source, &mut budget).expect("canonical sidecar list");
        assert_eq!(list.kind(), kind);
        assert_eq!(list.len(), 1);
        match kind {
            SidecarKind::Strings | SidecarKind::FormulaErrors => {
                assert_eq!(
                    list.get(7).and_then(SidecarValue::text),
                    Some(if kind == SidecarKind::Strings {
                        "seven"
                    } else {
                        "#VALUE!"
                    })
                );
            },
            SidecarKind::Formulas => assert_eq!(
                list.get(7).and_then(SidecarValue::formula),
                Some(formula.as_slice())
            ),
            SidecarKind::RichTextPayloads => assert_eq!(
                list.get(7)
                    .and_then(SidecarValue::reference)
                    .map(SidecarReference::identifier),
                Some(7007)
            ),
            SidecarKind::Comments => assert_eq!(
                list.get(7)
                    .and_then(SidecarValue::reference)
                    .map(SidecarReference::identifier),
                Some(8008)
            ),
        }
        assert_eq!(budget.list_reports.len(), 1);
        assert_eq!(budget.probe_reports.len(), 1);
    }
}

#[test]
fn sidecar_lists_accept_unknown_fields_but_reject_a_malformed_later_entry() {
    let valid = unknown_field(list_root(
        SidecarKind::Strings,
        &[string_entry(4, "opaque")],
    ));
    let mut budget = Budget::generous();
    let list = read_list(SidecarKind::Strings, &valid, &mut budget)
        .expect("unknown fields are opaque and source-preserving");
    assert_eq!(list.get(4).and_then(SidecarValue::text), Some("opaque"));

    let malformed = malformed_later_entry(list_root(
        SidecarKind::Strings,
        &[string_entry(4, "prefix")],
    ));
    let mut budget = Budget::generous();
    let error = read_list(SidecarKind::Strings, &malformed, &mut budget)
        .expect_err("the complete strict pass must reject the later malformed entry");
    assert!(matches!(
        error,
        HarnessError::Issue(SidecarIssue::StorageDecode {
            kind: SidecarKind::Strings,
            ..
        })
    ));
}

#[test]
fn comments_sidecar_requires_a_live_reference_count() {
    let source = list_root(SidecarKind::Comments, &[comment_entry(3, 0, 44)]);
    let mut budget = Budget::generous();
    let error = read_list(SidecarKind::Comments, &source, &mut budget)
        .expect_err("comment-storage list entries with no references are invalid");
    assert!(matches!(
        error,
        HarnessError::Issue(SidecarIssue::InvalidEntry {
            kind: SidecarKind::Comments
        })
    ));
}

#[test]
fn formula_source_copy_refusal_is_reported_before_retention() {
    let source = list_root(
        SidecarKind::Formulas,
        &[formula_entry(2, &number_formula())],
    );
    let mut budget = Budget::generous();
    budget.refuse = Some((SidecarKind::Formulas, SidecarAllocation::FormulaBytes));
    let error = read_list(SidecarKind::Formulas, &source, &mut budget)
        .expect_err("formula source retention is fallible");
    assert!(matches!(
        error,
        HarnessError::Refused {
            kind: SidecarKind::Formulas,
            target: SidecarAllocation::FormulaBytes,
            amount
        } if amount > 0
    ));
    assert!(
        budget
            .retained
            .iter()
            .all(|(_, target, _)| *target != SidecarAllocation::FormulaBytes)
    );
    assert_eq!(budget.formula_envelope_reports.len(), 1);
    assert!(budget.formula_envelope_failed_costs.is_empty());
}

#[test]
fn malformed_formula_preflight_reports_attempted_cost_without_publishing_source() {
    let malformed = [0x0a, 0x01, 0x80];
    let source = list_root(SidecarKind::Formulas, &[formula_entry(2, &malformed)]);
    let mut budget = Budget::generous();
    let error = read_list(SidecarKind::Formulas, &source, &mut budget)
        .expect_err("malformed formula envelopes fail before source retention");
    assert!(
        matches!(error, HarnessError::Render(message) if message.contains("Numbers FormulaArchive") || message.contains("wire"))
    );
    assert!(budget.formula_envelope_reports.is_empty());
    assert_eq!(budget.formula_envelope_failed_costs.len(), 1);
    assert!(budget.formula_envelope_failed_costs[0].fields() > 0);
    assert!(budget.formula_envelope_failed_costs[0].work() > 0);
    assert!(
        budget
            .retained
            .iter()
            .all(|(_, target, _)| *target != SidecarAllocation::FormulaBytes)
    );
}

#[test]
fn later_segment_failures_outrank_an_earlier_formula_conversion_error() {
    let malformed_formula = [0x0a, 0x01, 0x80];
    let first_segment = sidecar_segment(
        SidecarKind::Formulas,
        1,
        1,
        &[formula_entry(1, &malformed_formula)],
    );
    let root = list_root_with_segments(SidecarKind::Formulas, &[], &[41, 42]);

    let mut budget = Budget::generous();
    let missing = read_sidecar_list(
        SidecarKind::Formulas,
        700,
        [Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &root)],
        |segment_id, _budget| {
            Ok((segment_id == 41).then(|| {
                vec![Message::new(
                    TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND,
                    &first_segment,
                )]
            }))
        },
        &mut budget,
    )
    .expect_err("missing later segment is structural even after conversion failed");
    assert!(matches!(
        missing,
        HarnessError::Issue(SidecarIssue::Coordinator(
            litchi_numbers_wire::table_data_list::CoordinatorIssue::MissingSegment {
                object_id: 700,
                segment_id: 42
            }
        ))
    ));
    assert_eq!(budget.formula_envelope_failed_costs.len(), 1);

    let duplicate_segment = sidecar_segment(
        SidecarKind::Formulas,
        2,
        1,
        &[formula_entry(2, &number_formula())],
    );
    let mut budget = Budget::generous();
    let duplicate = read_sidecar_list(
        SidecarKind::Formulas,
        700,
        [Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &root)],
        |segment_id, _budget| {
            Ok(match segment_id {
                41 => Some(vec![Message::new(
                    TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND,
                    &first_segment,
                )]),
                42 => Some(vec![
                    Message::new(TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND, &duplicate_segment),
                    Message::new(TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND, &duplicate_segment),
                ]),
                _ => None,
            })
        },
        &mut budget,
    )
    .expect_err("duplicate later segment payload is structural");
    assert!(matches!(
        duplicate,
        HarnessError::Issue(SidecarIssue::Coordinator(
            litchi_numbers_wire::table_data_list::CoordinatorIssue::DuplicateSegmentPayload {
                segment_id: 42
            }
        ))
    ));
    assert_eq!(budget.formula_envelope_failed_costs.len(), 1);

    let corrupt_segment = [0x0a, 0x01, 0x80];
    let mut budget = Budget::generous();
    let corrupt = read_sidecar_list(
        SidecarKind::Formulas,
        700,
        [Message::new(TABLE_DATA_LIST_MESSAGE_KIND, &root)],
        |segment_id, _budget| {
            Ok(match segment_id {
                41 => Some(vec![Message::new(
                    TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND,
                    &first_segment,
                )]),
                42 => Some(vec![Message::new(
                    TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND,
                    &corrupt_segment,
                )]),
                _ => None,
            })
        },
        &mut budget,
    )
    .expect_err("a corrupt later segment must not be hidden by conversion failure");
    assert!(matches!(
        corrupt,
        HarnessError::Issue(SidecarIssue::StorageDecode {
            kind: SidecarKind::Formulas,
            ..
        })
    ));
    assert_eq!(budget.formula_envelope_failed_costs.len(), 1);
}

#[test]
fn comment_storage_reads_replies_and_preserves_unknown_wire() {
    let source = comment_storage_with_unknown("root text");
    let root = SidecarReference::new(700).expect("nonzero root identity");
    let mut budget = Budget::generous();
    let storage = read_comment_storage(
        root,
        [Message::new(COMMENT_STORAGE_MESSAGE_KIND, &source)],
        &mut budget,
    )
    .expect("comment payload with one reply");

    assert_eq!(storage.text(), "root text");
    assert_eq!(
        storage.author_reference().map(SidecarReference::identifier),
        Some(77)
    );
    assert_eq!(
        storage
            .replies()
            .iter()
            .map(|reference| reference.identifier())
            .collect::<Vec<_>>(),
        [78]
    );
    assert_eq!(
        storage
            .storage_uuid()
            .map(|uuid| (uuid.lower(), uuid.upper())),
        Some((9, 10))
    );
    assert_eq!(budget.comment_reports.len(), 1);
    assert_eq!(budget.comment_reports[0].reply_references(), 1);
}

#[test]
fn comment_storage_rejects_zero_author_and_zero_uuid() {
    let root = SidecarReference::new(700).expect("nonzero root identity");

    let author_zero = comment_storage(Some("root"), Some(0), &[], Some((1, 2)));
    let mut budget = Budget::generous();
    let error = read_comment_storage(
        root,
        [Message::new(COMMENT_STORAGE_MESSAGE_KIND, &author_zero)],
        &mut budget,
    )
    .expect_err("a present zero author reference is malformed");
    assert!(matches!(
        error,
        HarnessError::Issue(SidecarIssue::ZeroReference {
            kind: SidecarKind::Comments
        })
    ));

    let uuid_zero = comment_storage(Some("root"), Some(11), &[], Some((0, 0)));
    let mut budget = Budget::generous();
    let error = read_comment_storage(
        root,
        [Message::new(COMMENT_STORAGE_MESSAGE_KIND, &uuid_zero)],
        &mut budget,
    )
    .expect_err("a present all-zero storage UUID is not a valid identity");
    assert!(matches!(error, HarnessError::Issue(SidecarIssue::ZeroUuid)));
}

#[test]
fn comment_reply_reference_limit_is_strict_and_source_is_not_published() {
    let source = comment_storage(Some("root"), None, &[11, 12], Some((3, 4)));
    let root = SidecarReference::new(700).expect("nonzero root identity");
    let mut budget = Budget::generous();
    budget.max_comment_references = 1;
    let error = read_comment_storage(
        root,
        [Message::new(COMMENT_STORAGE_MESSAGE_KIND, &source)],
        &mut budget,
    )
    .expect_err("reply references consume the strict aggregate reference budget");
    assert!(matches!(
        error,
        HarnessError::Issue(SidecarIssue::CommentDecode(error))
            if matches!(error.resource_limit(), Some(CommentDecodeLimit::References { maximum: 1, .. }))
    ));
    assert!(budget.comment_reports.is_empty());
}

#[test]
fn formula_render_charges_successful_preflight_only_and_maps_failures() {
    let source = number_formula();
    let mut budget = Budget::generous();
    let rendered = render_formula(&source, 1, 0, 0, 10, 10, &Resolver, &mut budget)
        .expect("minimal formula renders");
    assert_eq!(rendered, "=0");
    assert_eq!(budget.formula_reports.len(), 1);
    assert!(budget.formula_reports[0].bytes() >= source.len());

    let malformed = [0x0a, 0x01, 0x80];
    let mut budget = Budget::generous();
    let error = render_formula(&malformed, 1, 0, 0, 10, 10, &Resolver, &mut budget)
        .expect_err("truncated formula wire fails during preflight");
    assert!(matches!(
        error,
        HarnessError::Issue(SidecarIssue::FormulaDecode(_))
    ));
    assert!(budget.formula_reports.is_empty());
}

#[test]
fn sidecar_materialization_keeps_missing_string_and_rich_text_empty_but_is_strict_for_formula_and_comment()
 {
    let tables = SidecarTables::new();
    let mut resolver = CellResolver::default();
    let missing_string = tables
        .materialize_cell(
            CellValueSource {
                value: ValueSource::Text(17),
                comment_identifier: None,
            },
            0,
            0,
            &mut resolver,
        )
        .expect("missing string retains legacy empty-cell compatibility");
    assert!(matches!(missing_string, (Value::Empty, None)));

    let missing_rich_text = tables
        .materialize_cell(
            CellValueSource {
                value: ValueSource::RichText(18),
                comment_identifier: None,
            },
            0,
            0,
            &mut resolver,
        )
        .expect("missing rich-text sidecar retains legacy empty-cell compatibility");
    assert!(matches!(missing_rich_text, (Value::Empty, None)));

    let formula_error = tables
        .materialize_cell(
            CellValueSource {
                value: ValueSource::Formula(19),
                comment_identifier: None,
            },
            0,
            0,
            &mut resolver,
        )
        .expect_err("formula identifiers are strict sidecar references");
    assert_eq!(
        formula_error,
        ResolverError::Missing {
            kind: SidecarKind::Formulas,
            key: 19,
        }
    );

    let comment_error = tables
        .materialize_cell(
            CellValueSource {
                value: ValueSource::Empty,
                comment_identifier: Some(20),
            },
            0,
            0,
            &mut resolver,
        )
        .expect_err("comment identifiers are strict sidecar references");
    assert_eq!(
        comment_error,
        ResolverError::Missing {
            kind: SidecarKind::Comments,
            key: 20,
        }
    );
}

#[test]
fn materialization_resolves_present_rich_text_and_formula_sidecars() {
    let mut tables = SidecarTables::new();
    tables
        .insert(
            litchi_numbers_wire::table_sidecars::SidecarList::from_entries(
                SidecarKind::RichTextPayloads,
                vec![(
                    4,
                    SidecarValue::Reference(SidecarReference::new(44).unwrap()),
                )],
            )
            .unwrap(),
        )
        .unwrap();
    tables
        .insert(
            litchi_numbers_wire::table_sidecars::SidecarList::from_entries(
                SidecarKind::Formulas,
                vec![(
                    5,
                    SidecarValue::Formula(number_formula().into_boxed_slice()),
                )],
            )
            .unwrap(),
        )
        .unwrap();

    let mut resolver = CellResolver::default();
    let rich = tables
        .materialize_cell(
            CellValueSource {
                value: ValueSource::RichText(4),
                comment_identifier: None,
            },
            1,
            2,
            &mut resolver,
        )
        .unwrap();
    assert!(matches!(rich, (Value::Text(text), None) if text == "rich text"));

    let formula = tables
        .materialize_cell(
            CellValueSource {
                value: ValueSource::Formula(5),
                comment_identifier: None,
            },
            1,
            2,
            &mut resolver,
        )
        .unwrap();
    assert!(matches!(formula, (Value::Formula(text), None) if text == "=formula"));
    assert_eq!(
        resolver.rich_text_calls,
        [SidecarReference::new(44).unwrap()]
    );
    assert_eq!(resolver.formula_calls, [(5, number_formula().len())]);
}
