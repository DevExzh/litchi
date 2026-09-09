//! Generated-free Numbers FormulaArchive retention and rendering.
//!
//! This private adapter deliberately keeps the historical generated
//! `TSCE.FormulaArchive` representation out of production. Formula list
//! entries are strictly preflighted and retained as bounded wire bytes; the
//! compatibility renderer consumes borrowed events from the neutral formula
//! codec only when a cell references an entry.

use super::table_extractor::{FormulaReferenceMaps, FormulaReferenceName, ProjectionBudget};
use crate::{Error, Result};
use litchi_iwa_common::formula::render::{FormulaExpr, FormulaRenderBudget, FormulaRenderer};
use litchi_iwa_common::wire::{
    WireDescent, parse_wire_view_with_limits, preflight_wire_tree_with_limits,
};
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_protos::numbers_formula_codec;
use std::fmt::Write as _;

const MAX_PAYLOAD_WORK: usize = WireLimits::MAX_REWRITE_WORK;
const MAX_FORMULA_RENDER_STRUCTURE_NODES: usize = litchi_numbers::MAX_REFERENCES;

fn allocation_error(resource: &'static str, amount: usize) -> Error {
    Error::IwaCommon(litchi_iwa_common::Error::Allocation { resource, amount })
}
/// Formula bytes retained by the table sidecar.
///
/// The table-list projection must reject malformed formula payloads even when
/// no cell eventually references them, but decoding every archive into the
/// generated repeated-node representation makes an unrelated formula entry an
/// eager allocation. Retain one bounded owned wire copy after a strict
/// envelope preflight and decode it only at the cell that needs rendering.
/// The compatibility renderer consumes the retained bytes through the
/// generated-free event codec, so no generated archive is staged in
/// production.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct FormulaArchiveBytes {
    bytes: Box<[u8]>,
    /// Whether the complete wire preflight found only nodes represented by
    /// the compact generated-free scalar visitor.  This is conservative: a
    /// false value keeps the lossless generated-free compatibility renderer.
    scalar_visitor_eligible: bool,
    /// Exact number of AST nodes observed by the wire preflight.  The scalar
    /// visitor charges this bound before retaining any decoded nodes.
    scalar_visitor_node_count: usize,
    /// Number of non-AST repeated entries traversed by the compatibility
    /// reader. AST-node entries are accounted separately by
    /// `scalar_visitor_node_count` and the render work budget; this counter
    /// preserves the legacy admission charge for the reader's repeated
    /// metadata walk without retaining generated vectors.
    lazy_traversal_entry_count: usize,
}

impl FormulaArchiveBytes {
    pub(super) fn from_wire(source: &[u8], budget: &mut ProjectionBudget) -> Result<Self> {
        // Charge before either preflight or allocation.  A malformed
        // candidate therefore retains the same monotonic wire cost as the
        // previous eager decoder, while no partial owned value can escape.
        budget.charge_formula_wire(source.len())?;
        let (scalar_visitor_eligible, scalar_visitor_node_count, lazy_traversal_entry_count) =
            preflight_formula_archive_envelope(source, budget)?;

        let mut owned = Vec::new();
        owned
            .try_reserve_exact(source.len())
            .map_err(|_| allocation_error("Numbers formula wire", source.len()))?;
        owned.extend_from_slice(source);
        Ok(Self {
            bytes: owned.into_boxed_slice(),
            scalar_visitor_eligible,
            scalar_visitor_node_count,
            lazy_traversal_entry_count,
        })
    }
}

/// Validate the complete FormulaArchive wire tree without materializing its
/// repeated AST representation.
///
/// The generic wire scanner is deliberately schema-directed only for nested
/// message fields.  Unknown fields remain opaque, matching Prost's forward
/// compatible behavior, while known scalar/string fields still have their
/// native wire type and UTF-8 checked.  Canonical field framing is required at
/// every level so a malformed, unreferenced formula cannot hide behind
/// deferred cell rendering.
fn preflight_formula_archive_envelope(
    source: &[u8],
    budget: &mut ProjectionBudget,
) -> Result<(bool, usize, usize)> {
    // Prost accepts an empty proto2 message even when its schema marks field
    // 1 as required.  The former eager FormulaArchive decoder therefore
    // admitted the serialized default archive, whose empty AST rendered as
    // `=`.  Preserve that compatibility case while retaining the required
    // root field check for every non-empty archive.
    if source.is_empty() {
        return Ok((false, 0, 0));
    }

    // The wire preflight report is aggregate: every selected child message is
    // charged in addition to its parent.  A per-message input/field ceiling
    // would reject an otherwise small, valid formula as soon as the first
    // ASTNode is descended.  Keep the scanner finite with the same package
    // ceilings used by `FormulaEnvelopeCost`; the latter folds the current
    // package offsets into the admission decision.
    let max_fields = litchi_numbers::MAX_REFERENCES
        .saturating_sub(budget.payload_fields)
        .clamp(1, WireLimits::MAX_FIELDS);
    let max_input_bytes = MAX_PAYLOAD_WORK
        .saturating_sub(budget.payload_work)
        .clamp(1, WireLimits::MAX_INPUT_BYTES);
    let limits = WireLimits::default()
        .with_input_bytes(max_input_bytes)
        .and_then(|limits| limits.with_fields(max_fields))
        .and_then(|limits| limits.with_nesting(WireLimits::MAX_NESTING))
        .and_then(|limits| limits.with_rewrite_work(MAX_PAYLOAD_WORK))
        .map_err(map_formula_envelope_wire_error)?;

    let mut attempted = match FormulaEnvelopeCost::new(source.len(), budget) {
        Ok(cost) => cost,
        Err(error) => {
            budget.retain_formula_preflight_cost(0, source.len());
            return Err(map_formula_envelope_wire_error(error));
        },
    };
    let mut root_ast_present = false;
    let mut root_ast_count = 0usize;
    let mut scalar_visitor_eligible = true;
    let mut scalar_visitor_node_count = 0usize;
    let mut lazy_traversal_entry_count = 0usize;
    let mut root_known_fields = [0u32; 9];
    let mut root_known_field_count = 0usize;
    let preflight = preflight_wire_tree_with_limits(source, limits, |visit| {
        let field = visit.field();
        attempted.charge_field()?;
        field.validate_canonical_framing()?;
        let schema = formula_envelope_field(visit.path(), field.number());
        // Unknown fields are intentionally retained as opaque wire data by
        // the envelope preflight, but they are not represented by the compact
        // scalar visitor.  Keep these archives on the full compatibility
        // renderer so a future extension cannot silently disappear on the
        // generated-free route.  The strict wire walk and bounded fallback
        // still apply exactly as before.
        if schema.wire_type.is_none() {
            scalar_visitor_eligible = false;
        }
        if let Some(expected_wire_type) = schema.wire_type
            && field.wire_type() != expected_wire_type
        {
            return Err(litchi_iwa_common::Error::InvalidFormat(format!(
                "Numbers FormulaArchive field {} has wire type {}, expected {}",
                field.number(),
                field.wire_type(),
                expected_wire_type
            )));
        }
        if visit.path().is_empty()
            && schema.wire_type.is_some()
            && !formula_field_is_repeated(visit.path(), field.number())
        {
            if root_known_fields[..root_known_field_count].contains(&field.number()) {
                return Err(litchi_iwa_common::Error::InvalidFormat(format!(
                    "Numbers FormulaArchive field {} occurs more than once",
                    field.number()
                )));
            }
            root_known_fields[root_known_field_count] = field.number();
            root_known_field_count += 1;
        }
        if visit.path().is_empty() && field.number() == 1 {
            root_ast_count = root_ast_count.saturating_add(1);
            if root_ast_count > 1 {
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "Numbers FormulaArchive AST node array occurs more than once".to_owned(),
                ));
            }
            root_ast_present = true;
        }
        if schema.nested {
            attempted.charge_nested(field.payload().len())?;
            let mut nested_path = [0u32; WireLimits::MAX_NESTING + 1];
            let Some(path_end) = visit.path().len().checked_add(1) else {
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "Numbers FormulaArchive nesting path overflows".to_owned(),
                ));
            };
            if path_end > nested_path.len() {
                return Err(litchi_iwa_common::Error::LimitExceeded {
                    kind: LimitKind::Nesting,
                    observed: path_end,
                    limit: WireLimits::MAX_NESTING,
                });
            }
            nested_path[..visit.path().len()].copy_from_slice(visit.path());
            nested_path[visit.path().len()] = field.number();
            require_formula_fields(
                field.payload(),
                &nested_path[..path_end],
                schema.required_fields,
            )?;
        }
        if formula_ast_array_path(visit.path()) && field.number() == 1 {
            required_formula_ast_node_type(field.payload())?;
        }
        // Preserve the legacy admission charge for every repeated native
        // vector. AST node entries are charged separately by the formula-render
        // work budget, so count only the other entries (UID lists, ranges,
        // lambda identifiers, and similar vectors) before lazy compatibility
        // traversal begins.
        if formula_field_is_repeated(visit.path(), field.number())
            && !(formula_ast_array_path(visit.path()) && field.number() == 1)
        {
            lazy_traversal_entry_count =
                lazy_traversal_entry_count.checked_add(1).ok_or_else(|| {
                    litchi_iwa_common::Error::InvalidFormat(
                        "Numbers FormulaArchive repeated-entry count overflows host usize"
                            .to_owned(),
                    )
                })?;
        }
        if formula_ast_node_path(visit.path()) {
            if field.number() == 1 {
                scalar_visitor_node_count =
                    scalar_visitor_node_count.checked_add(1).ok_or_else(|| {
                        litchi_iwa_common::Error::InvalidFormat(
                            "Numbers FormulaArchive AST node count overflows host usize".to_owned(),
                        )
                    })?;
                let node_type = litchi_iwa_common::decode_varint_from_bytes(field.payload())
                    .ok()
                    .and_then(|(value, width)| (width == field.payload().len()).then_some(value))
                    .and_then(|value| u32::try_from(value).ok());
                if node_type.is_none_or(|value| {
                    // The standalone LocalCellReferenceNode (27) is a
                    // compatibility-only node: unlike the nested local
                    // reference in CellReferenceNode, the scalar codec
                    // does not preserve its sticky flags.
                    value == 27 || !numbers_formula_codec::is_scalar_visitor_node_type(value)
                }) {
                    scalar_visitor_eligible = false;
                }
            } else if field.number() == 2 {
                // The streaming evaluator admits only the five function
                // identifiers below.  Unknown/native function IDs are valid
                // compatibility-renderer input, but attempting the scalar
                // visitor first would allocate its full node vector only to
                // fall back after the codec rejects the function.
                let function_identifier =
                    litchi_iwa_common::decode_varint_from_bytes(field.payload())
                        .ok()
                        .and_then(|(value, width)| {
                            (width == field.payload().len()).then_some(value)
                        })
                        .and_then(|value| u32::try_from(value).ok());
                if function_identifier.is_none_or(|value| !matches!(value, 15 | 30 | 84 | 88 | 168))
                {
                    scalar_visitor_eligible = false;
                }
            } else if !matches!(field.number(), 2 | 3 | 4 | 5 | 10 | 15 | 25 | 42 | 43) {
                // These are the only direct AST-node fields the compact
                // visitor can represent. Decimal128 sidecars (fields 42/43)
                // are validated by the scalar codec and do not change the
                // rendered number. Nested messages are checked by the
                // surrounding schema walk; any other direct field (strings,
                // dates, arrays, thunks, ranges, UIDs, or owner metadata)
                // falls back to the lossless compatibility renderer.
                scalar_visitor_eligible = false;
            }
        }
        if schema.utf8 {
            std::str::from_utf8(field.payload()).map_err(|_error| {
                litchi_iwa_common::Error::InvalidFormat(
                    "Numbers FormulaArchive string field is not valid UTF-8".to_owned(),
                )
            })?;
        }
        validate_formula_scalar(field, schema.scalar)?;
        Ok(if schema.nested {
            WireDescent::Descend
        } else {
            WireDescent::Skip
        })
    });
    match preflight {
        Ok(report) if root_ast_present => {
            debug_assert_eq!(attempted.fields, report.fields());
            debug_assert_eq!(attempted.work, report.scanned_bytes());
            budget.charge_wire_preflight(report)?;
            Ok((
                scalar_visitor_eligible,
                scalar_visitor_node_count,
                lazy_traversal_entry_count,
            ))
        },
        Ok(report) => {
            debug_assert_eq!(attempted.fields, report.fields());
            debug_assert_eq!(attempted.work, report.scanned_bytes());
            budget.retain_formula_preflight_cost(report.fields(), report.scanned_bytes());
            Err(Error::InvalidFormat(
                "malformed Numbers formula payload".to_owned(),
            ))
        },
        Err(error) => {
            budget.retain_formula_preflight_cost(attempted.fields, attempted.work);
            Err(map_formula_envelope_wire_error(error))
        },
    }
}

#[derive(Debug, Clone, Copy)]
struct FormulaEnvelopeCost {
    base_fields: usize,
    base_work: usize,
    fields: usize,
    work: usize,
}

impl FormulaEnvelopeCost {
    fn new(source_len: usize, budget: &ProjectionBudget) -> litchi_iwa_common::Result<Self> {
        let mut cost = Self {
            base_fields: budget.payload_fields,
            base_work: budget.payload_work,
            fields: 0,
            work: 0,
        };
        cost.charge_work(source_len)?;
        Ok(cost)
    }

    fn charge_field(&mut self) -> litchi_iwa_common::Result<()> {
        self.fields =
            self.fields
                .checked_add(1)
                .ok_or(litchi_iwa_common::Error::LimitExceeded {
                    kind: LimitKind::Fields,
                    observed: usize::MAX,
                    limit: litchi_numbers::MAX_REFERENCES,
                })?;
        let observed = self.base_fields.saturating_add(self.fields);
        if observed > litchi_numbers::MAX_REFERENCES {
            return Err(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed,
                limit: litchi_numbers::MAX_REFERENCES,
            });
        }
        Ok(())
    }

    fn charge_work(&mut self, amount: usize) -> litchi_iwa_common::Result<()> {
        self.work =
            self.work
                .checked_add(amount)
                .ok_or(litchi_iwa_common::Error::LimitExceeded {
                    kind: LimitKind::RewriteWork,
                    observed: usize::MAX,
                    limit: MAX_PAYLOAD_WORK,
                })?;
        let observed = self.base_work.saturating_add(self.work);
        if observed > MAX_PAYLOAD_WORK {
            return Err(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::RewriteWork,
                observed,
                limit: MAX_PAYLOAD_WORK,
            });
        }
        Ok(())
    }

    fn charge_nested(&mut self, bytes: usize) -> litchi_iwa_common::Result<()> {
        self.charge_work(bytes)
    }
}

fn required_formula_ast_node_type(source: &[u8]) -> litchi_iwa_common::Result<()> {
    let limits = WireLimits::default()
        .with_input_bytes(source.len().max(1))
        .and_then(|limits| limits.with_fields(source.len().clamp(1, WireLimits::MAX_FIELDS)))
        .and_then(|limits| limits.with_nesting(1))?;
    let view = parse_wire_view_with_limits(source, limits)?;
    let mut found = false;
    for field in view.fields() {
        field.validate_canonical_framing()?;
        if field.number() == 1 {
            if found {
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "Numbers FormulaArchive AST node type occurs more than once".to_owned(),
                ));
            }
            found = true;
            if field.wire_type() != 0 {
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "Numbers FormulaArchive AST node type has the wrong wire type".to_owned(),
                ));
            }
        }
    }
    if !found {
        return Err(litchi_iwa_common::Error::InvalidFormat(
            "Numbers FormulaArchive AST node is missing required node type".to_owned(),
        ));
    }
    Ok(())
}

/// Check the required direct children of a schema-known nested message.
///
/// Prost's generated proto2 decoder does not consistently enforce required
/// fields on all of the deferred formula messages.  Keep this check local to
/// the nested payload selected by the schema so unknown opaque fields remain
/// untouched and repeated fields retain their normal semantics.
fn require_formula_fields(
    source: &[u8],
    message_path: &[u32],
    required_fields: &[u32],
) -> litchi_iwa_common::Result<()> {
    let limits = WireLimits::default()
        .with_input_bytes(source.len().max(1))
        .and_then(|limits| limits.with_fields(source.len().clamp(1, WireLimits::MAX_FIELDS)))
        .and_then(|limits| limits.with_nesting(1))?;
    let view = parse_wire_view_with_limits(source, limits)?;
    let mut seen = [false; 8];
    let mut known_fields = [0u32; 64];
    let mut known_field_count = 0usize;
    for field in view.fields() {
        field.validate_canonical_framing()?;
        let schema = formula_envelope_field(message_path, field.number());
        if schema.wire_type.is_some() && !formula_field_is_repeated(message_path, field.number()) {
            if known_fields[..known_field_count].contains(&field.number()) {
                if required_fields.contains(&field.number()) {
                    return Err(litchi_iwa_common::Error::InvalidFormat(format!(
                        "Numbers FormulaArchive required field {} occurs more than once",
                        field.number()
                    )));
                }
                return Err(litchi_iwa_common::Error::InvalidFormat(format!(
                    "Numbers FormulaArchive field {} occurs more than once",
                    field.number()
                )));
            }
            if known_field_count == known_fields.len() {
                return Err(litchi_iwa_common::Error::LimitExceeded {
                    kind: LimitKind::Fields,
                    observed: known_field_count.saturating_add(1),
                    limit: known_fields.len(),
                });
            }
            known_fields[known_field_count] = field.number();
            known_field_count += 1;
        }
        let Some(index) = required_fields
            .iter()
            .position(|required| *required == field.number())
        else {
            continue;
        };
        if seen[index] {
            return Err(litchi_iwa_common::Error::InvalidFormat(format!(
                "Numbers FormulaArchive required field {} occurs more than once",
                field.number()
            )));
        }
        seen[index] = true;
    }
    for (index, field_number) in required_fields.iter().enumerate() {
        if !seen[index] {
            return Err(litchi_iwa_common::Error::InvalidFormat(format!(
                "Numbers FormulaArchive required field {} is missing",
                field_number
            )));
        }
    }
    Ok(())
}

/// Return whether a schema-known field is allowed to occur more than once.
///
/// Prost accepts duplicate singular fields with last-value-wins semantics.
/// The formula archive ingress is stricter: only fields declared `repeated` in
/// the native schema may repeat. The path distinguishes repeated AST nodes
/// from the singular field-1 envelopes used by several nearby messages.
fn formula_field_is_repeated(path: &[u32], number: u32) -> bool {
    // Nested formula messages are visited with the complete path from the
    // FormulaArchive root (for example `[1, 1, 39, 1, 6]` for
    // CategoryReferenceArchive.CatRefUidList).  The repeated-field table is
    // expressed relative to the AST node, so strip that prefix before
    // matching the message-specific suffix.  Root AST-node arrays remain
    // addressable by their original path.
    let suffix = formula_ast_node_prefix_len(path).map_or(path, |prefix_len| &path[prefix_len..]);
    if number == 1 {
        formula_ast_array_path(path)
            || suffix == [38]
            || matches!(suffix, [38, 1, 1] | [38, 1, 2])
            || suffix == [39, 1, 6]
            || suffix == [40]
            || suffix == [45]
    } else {
        suffix == [40] && (2..=4).contains(&number)
    }
}

fn validate_canonical_formula_varint(
    field: litchi_iwa_common::wire::WireFieldView<'_>,
) -> litchi_iwa_common::Result<()> {
    let payload = field.payload();
    let (value, width) = litchi_iwa_common::decode_varint_from_bytes(payload).map_err(|error| {
        litchi_iwa_common::Error::InvalidFormat(format!(
            "Numbers FormulaArchive field {} has an invalid scalar varint: {error}",
            field.number()
        ))
    })?;
    if width != payload.len() || width != litchi_iwa_common::varint::encoded_len(value) {
        return Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "Numbers FormulaArchive field {} has a noncanonical scalar varint",
            field.number()
        )));
    }
    Ok(())
}

fn validate_formula_scalar(
    field: litchi_iwa_common::wire::WireFieldView<'_>,
    scalar: FormulaScalar,
) -> litchi_iwa_common::Result<()> {
    match scalar {
        FormulaScalar::Unknown => Ok(()),
        FormulaScalar::Varint => validate_canonical_formula_varint(field),
        FormulaScalar::Bool | FormulaScalar::U32 | FormulaScalar::Int32 | FormulaScalar::SInt32 => {
            let payload = field.payload();
            let (value, width) =
                litchi_iwa_common::decode_varint_from_bytes(payload).map_err(|error| {
                    litchi_iwa_common::Error::InvalidFormat(format!(
                        "Numbers FormulaArchive field {} has an invalid scalar varint: {error}",
                        field.number()
                    ))
                })?;
            if width != payload.len() || width != litchi_iwa_common::varint::encoded_len(value) {
                return Err(litchi_iwa_common::Error::InvalidFormat(format!(
                    "Numbers FormulaArchive field {} has a noncanonical scalar varint",
                    field.number()
                )));
            }
            let valid = match scalar {
                FormulaScalar::Bool => value <= 1,
                FormulaScalar::U32 | FormulaScalar::SInt32 => value <= u64::from(u32::MAX),
                FormulaScalar::Int32 => {
                    i32::try_from(value).is_ok() || value >= 0xffff_ffff_8000_0000
                },
                FormulaScalar::Unknown | FormulaScalar::Varint => true,
            };
            if !valid {
                return Err(litchi_iwa_common::Error::InvalidFormat(format!(
                    "Numbers FormulaArchive field {} has a noncanonical scalar value",
                    field.number()
                )));
            }
            Ok(())
        },
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaScalar {
    Unknown,
    Varint,
    Bool,
    U32,
    Int32,
    SInt32,
}

#[derive(Debug, Clone, Copy)]
struct FormulaEnvelopeField {
    wire_type: Option<u8>,
    nested: bool,
    utf8: bool,
    scalar: FormulaScalar,
    required_fields: &'static [u32],
}

const FORMULA_REQUIRED_NONE: &[u32] = &[];
const FORMULA_REQUIRED_AST_STICKY_BITS: &[u32] = &[1, 2, 3, 4];
const FORMULA_REQUIRED_AST_UID_TRACT: &[u32] = &[1, 2];
const FORMULA_REQUIRED_AST_UID: &[u32] = &[1, 2];
const FORMULA_REQUIRED_AST_CATEGORY_REFERENCE: &[u32] = &[1];
const FORMULA_REQUIRED_CATEGORY_REFERENCE: &[u32] = &[1, 2, 3, 4];
const FORMULA_REQUIRED_PRESERVE_FLAGS: &[u32] = &[1, 2];
const FORMULA_REQUIRED_LOCAL_REFERENCE: &[u32] = &[1, 2, 3, 4];
const FORMULA_REQUIRED_CROSS_REFERENCE: &[u32] = &[1, 2, 3, 4, 5];
const FORMULA_REQUIRED_COORDINATE: &[u32] = &[1];
const FORMULA_REQUIRED_UID_COORDINATE: &[u32] = &[1, 2, 3, 4];
const FORMULA_REQUIRED_AST_CATEGORY_LEVELS: &[u32] = &[1, 2];
const FORMULA_REQUIRED_CROSS_EXTRA: &[u32] = &[1];
const FORMULA_REQUIRED_RANGE_BEGIN: &[u32] = &[1];

const fn formula_envelope_scalar(wire_type: u8) -> FormulaEnvelopeField {
    FormulaEnvelopeField {
        wire_type: Some(wire_type),
        nested: false,
        utf8: false,
        scalar: if wire_type == 0 {
            FormulaScalar::Varint
        } else {
            FormulaScalar::Unknown
        },
        required_fields: FORMULA_REQUIRED_NONE,
    }
}

const fn formula_envelope_bool() -> FormulaEnvelopeField {
    FormulaEnvelopeField {
        wire_type: Some(0),
        nested: false,
        utf8: false,
        scalar: FormulaScalar::Bool,
        required_fields: FORMULA_REQUIRED_NONE,
    }
}

const fn formula_envelope_u32() -> FormulaEnvelopeField {
    FormulaEnvelopeField {
        wire_type: Some(0),
        nested: false,
        utf8: false,
        scalar: FormulaScalar::U32,
        required_fields: FORMULA_REQUIRED_NONE,
    }
}

const fn formula_envelope_int32() -> FormulaEnvelopeField {
    FormulaEnvelopeField {
        wire_type: Some(0),
        nested: false,
        utf8: false,
        scalar: FormulaScalar::Int32,
        required_fields: FORMULA_REQUIRED_NONE,
    }
}

const fn formula_envelope_sint32() -> FormulaEnvelopeField {
    FormulaEnvelopeField {
        wire_type: Some(0),
        nested: false,
        utf8: false,
        scalar: FormulaScalar::SInt32,
        required_fields: FORMULA_REQUIRED_NONE,
    }
}

const fn formula_envelope_nested() -> FormulaEnvelopeField {
    FormulaEnvelopeField {
        wire_type: Some(2),
        nested: true,
        utf8: false,
        scalar: FormulaScalar::Unknown,
        required_fields: FORMULA_REQUIRED_NONE,
    }
}

const fn formula_envelope_nested_required(required_fields: &'static [u32]) -> FormulaEnvelopeField {
    FormulaEnvelopeField {
        wire_type: Some(2),
        nested: true,
        utf8: false,
        scalar: FormulaScalar::Unknown,
        required_fields,
    }
}

const fn formula_envelope_utf8() -> FormulaEnvelopeField {
    FormulaEnvelopeField {
        wire_type: Some(2),
        nested: false,
        utf8: true,
        scalar: FormulaScalar::Unknown,
        required_fields: FORMULA_REQUIRED_NONE,
    }
}

const FORMULA_ENVELOPE_UNKNOWN: FormulaEnvelopeField = FormulaEnvelopeField {
    wire_type: None,
    nested: false,
    utf8: false,
    scalar: FormulaScalar::Unknown,
    required_fields: FORMULA_REQUIRED_NONE,
};

/// Return the schema shape for a FormulaArchive field.  ASTNodeArray/ASTNode
/// paths are recognized structurally, so thunk arrays can recurse to any
/// bounded depth without treating arbitrary length-delimited strings/bytes as
/// messages.
fn formula_envelope_field(path: &[u32], number: u32) -> FormulaEnvelopeField {
    if path.is_empty() {
        return match number {
            1 => formula_envelope_nested(),
            2..=3 => formula_envelope_u32(),
            4..=5 => formula_envelope_bool(),
            6 => formula_envelope_nested(),
            7..=9 => formula_envelope_nested_required(FORMULA_REQUIRED_AST_UID),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        };
    }
    if path == [6] {
        return if (1..=5).contains(&number) {
            formula_envelope_bool()
        } else {
            FORMULA_ENVELOPE_UNKNOWN
        };
    }
    if matches!(path, [7] | [8] | [9]) {
        return if (1..=2).contains(&number) {
            formula_envelope_scalar(0)
        } else {
            FORMULA_ENVELOPE_UNKNOWN
        };
    }
    if formula_ast_array_path(path) {
        return if number == 1 {
            formula_envelope_nested()
        } else {
            FORMULA_ENVELOPE_UNKNOWN
        };
    }
    let Some(node_prefix_len) = formula_ast_node_prefix_len(path) else {
        return FORMULA_ENVELOPE_UNKNOWN;
    };
    formula_ast_suffix_field(&path[node_prefix_len..], number)
}

/// `ASTNodeArrayArchive` appears at `[1]`, then at a repeated `1,14` suffix
/// for every thunk (`ASTNodeArray.ast_node`, `ASTNode.thunk_array`).
fn formula_ast_array_path(path: &[u32]) -> bool {
    if path == [1] {
        return true;
    }
    if path.len() < 3 || path[0] != 1 {
        return false;
    }
    let mut index = 1;
    while index + 1 < path.len() && path[index..index + 2] == [1, 14] {
        index += 2;
    }
    index == path.len()
}

/// Return whether `path` identifies an ASTNodeArchive itself rather than one
/// of its nested coordinate/reference messages. The scalar visitor eligibility
/// pass uses this to distinguish direct node fields from nested wire fields.
fn formula_ast_node_path(path: &[u32]) -> bool {
    formula_ast_node_prefix_len(path) == Some(path.len())
}

/// Return the ASTNode path prefix length.  A node is reached through `1,1`,
/// then zero or more `14,1` thunk/array edges.  The remainder identifies a
/// child message (for example `[16, 5]` is a cross-table CFUUID).
fn formula_ast_node_prefix_len(path: &[u32]) -> Option<usize> {
    if path.len() < 2 || path[..2] != [1, 1] {
        return None;
    }
    let mut index = 2;
    while index + 1 < path.len() && path[index..index + 2] == [14, 1] {
        index += 2;
    }
    Some(index)
}

fn formula_ast_suffix_field(suffix: &[u32], number: u32) -> FormulaEnvelopeField {
    match suffix {
        // TSCE.ASTNodeArchive
        [] => match number {
            1 | 9 | 47 => formula_envelope_int32(),
            2..=3 | 11..=13 | 18 | 22..=24 | 37 | 46 => formula_envelope_u32(),
            5 | 10 | 19..=20 | 29 | 36 => formula_envelope_bool(),
            42..=43 => formula_envelope_scalar(0),
            4 | 7 | 8 => formula_envelope_scalar(1),
            6 | 17 | 21 | 25 | 34 | 35 => formula_envelope_utf8(),
            14 | 40 | 45 => formula_envelope_nested(),
            15 => formula_envelope_nested_required(FORMULA_REQUIRED_LOCAL_REFERENCE),
            16 => formula_envelope_nested_required(FORMULA_REQUIRED_CROSS_REFERENCE),
            26 | 27 => formula_envelope_nested_required(FORMULA_REQUIRED_COORDINATE),
            28 => formula_envelope_nested_required(FORMULA_REQUIRED_CROSS_EXTRA),
            30 => formula_envelope_nested_required(FORMULA_REQUIRED_UID_COORDINATE),
            33 | 41 => formula_envelope_nested_required(FORMULA_REQUIRED_AST_STICKY_BITS),
            38 => formula_envelope_nested_required(&[2]),
            39 => formula_envelope_nested_required(FORMULA_REQUIRED_AST_CATEGORY_REFERENCE),
            44 => formula_envelope_nested_required(FORMULA_REQUIRED_AST_CATEGORY_LEVELS),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.ASTNodeArrayArchive (ASTNode.thunk_array)
        [14] => {
            if number == 1 {
                formula_envelope_nested()
            } else {
                FORMULA_ENVELOPE_UNKNOWN
            }
        },
        // Local and cross-table cell reference messages are deliberately
        // separate: only the cross-table variant has field 5 (CFUUID) and
        // fields 6..9 (whitespace strings).
        [15] => match number {
            1..=4 => formula_envelope_u32(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        [16] => match number {
            1..=4 => formula_envelope_u32(),
            5 => formula_envelope_nested(),
            6..=9 => formula_envelope_utf8(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSP.CFUUIDArchive
        [16, 5] | [28, 1] => match number {
            1 => formula_envelope_scalar(2),
            2..=5 => formula_envelope_u32(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.ASTColumn/RowCoordinateArchive
        [26] | [27] => match number {
            1 => formula_envelope_sint32(),
            2 => formula_envelope_bool(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.ASTStickyBits
        [33] | [41] | [38, 2] => match number {
            1..=4 => formula_envelope_bool(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.ASTCrossTableReferenceExtraInfoArchive
        [28] => match number {
            1 => formula_envelope_nested(),
            2..=5 => formula_envelope_utf8(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.ASTUidCoordinateArchive
        [30] => match number {
            1 | 2 => formula_envelope_nested_required(FORMULA_REQUIRED_AST_UID),
            3..=4 => formula_envelope_bool(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSP.UUID
        [30, 1]
        | [30, 2]
        | [38, 1, 1, 1]
        | [38, 1, 2, 1]
        | [39, 1, 1]
        | [39, 1, 2]
        | [39, 1, 9]
        | [39, 1, 10]
        | [39, 1, 6, 1] => {
            if (1..=2).contains(&number) {
                formula_envelope_scalar(0)
            } else {
                FORMULA_ENVELOPE_UNKNOWN
            }
        },
        // TSCE.ASTUidTractList
        [38] => match number {
            1 => formula_envelope_nested_required(FORMULA_REQUIRED_AST_UID_TRACT),
            2 => formula_envelope_nested_required(FORMULA_REQUIRED_AST_STICKY_BITS),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.ASTUidTract
        [38, 1] => match number {
            1..=2 => formula_envelope_nested(),
            3 | 5 => formula_envelope_bool(),
            4 => formula_envelope_int32(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.ASTUidList
        [38, 1, 1] | [38, 1, 2] => {
            if number == 1 {
                formula_envelope_nested_required(FORMULA_REQUIRED_AST_UID)
            } else {
                FORMULA_ENVELOPE_UNKNOWN
            }
        },
        // TSCE.ASTCategoryReferenceArchive
        [39] => {
            if number == 1 {
                formula_envelope_nested_required(FORMULA_REQUIRED_CATEGORY_REFERENCE)
            } else {
                FORMULA_ENVELOPE_UNKNOWN
            }
        },
        // TSCE.CategoryReferenceArchive
        [39, 1] => match number {
            1 | 2 | 9 | 10 => formula_envelope_nested_required(FORMULA_REQUIRED_AST_UID),
            3 | 13 => formula_envelope_u32(),
            4 => formula_envelope_sint32(),
            8 => formula_envelope_int32(),
            11 | 12 | 14 => formula_envelope_bool(),
            6 => formula_envelope_nested(),
            7 => formula_envelope_nested_required(FORMULA_REQUIRED_PRESERVE_FLAGS),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.CategoryReferenceArchive.CatRefUidList
        [39, 1, 6] => {
            if number == 1 {
                formula_envelope_nested_required(FORMULA_REQUIRED_AST_UID)
            } else {
                FORMULA_ENVELOPE_UNKNOWN
            }
        },
        // TSCE.PreserveColumnRowFlagsArchive
        [39, 1, 7] => match number {
            1..=4 => formula_envelope_bool(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.ASTCategoryLevels
        [44] => match number {
            1..=3 => formula_envelope_u32(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.ASTColonTractArchive and its four range message variants.
        [40] => match number {
            1..=4 => formula_envelope_nested_required(FORMULA_REQUIRED_RANGE_BEGIN),
            5 => formula_envelope_bool(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        [40, 1] | [40, 2] => match number {
            1 | 2 => formula_envelope_int32(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        [40, 3] | [40, 4] => match number {
            1 | 2 => formula_envelope_u32(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.ASTLambdaIdentsListArchive
        [45] => match number {
            1 | 3 | 4 => formula_envelope_utf8(),
            2 => formula_envelope_u32(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        _ => FORMULA_ENVELOPE_UNKNOWN,
    }
}

fn map_formula_envelope_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::InputBytes,
            observed,
            limit,
        } => formula_limit_error(LimitKind::InputBytes, observed, limit),
        litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Fields,
            observed,
            limit,
        } => formula_limit_error(LimitKind::Fields, observed, limit),
        litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Nesting,
            observed,
            limit,
        } => formula_limit_error(LimitKind::Nesting, observed, limit),
        litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::RewriteWork,
            observed,
            limit,
        } => formula_limit_error(LimitKind::RewriteWork, observed, limit),
        litchi_iwa_common::Error::Allocation { resource, amount } => {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation { resource, amount })
        },
        litchi_iwa_common::Error::InvalidFormat(message) => {
            Error::InvalidFormat(format!("malformed Numbers formula payload: {message}"))
        },
        other => Error::InvalidFormat(format!("malformed Numbers formula payload: {other}")),
    }
}

fn formula_limit_error(kind: LimitKind, observed: usize, limit: usize) -> Error {
    Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
        kind,
        observed,
        limit,
    })
}

fn formula_output_limit_error(observed: usize, budget: &ProjectionBudget) -> Error {
    formula_limit_error(
        LimitKind::OutputBytes,
        observed,
        budget.max_output_text_bytes,
    )
}

impl FormulaRenderBudget for ProjectionBudget {
    type Error = Error;

    fn output_limit(&self, observed: usize) -> Self::Error {
        formula_output_limit_error(observed, self)
    }

    fn allocation(&self, resource: &'static str, amount: usize) -> Self::Error {
        allocation_error(resource, amount)
    }

    fn invalid(&self, message: &'static str) -> Self::Error {
        Error::ParseError(message.to_owned())
    }

    fn check(&self, amount: usize) -> std::result::Result<(), Self::Error> {
        self.check_output_text(amount)
    }

    fn check_structure(&self, nodes: usize, parts: usize) -> std::result::Result<(), Self::Error> {
        let max_nodes = MAX_FORMULA_RENDER_STRUCTURE_NODES;
        let max_parts = max_nodes.saturating_mul(8);
        if nodes > max_nodes {
            return Err(formula_limit_error(
                LimitKind::RewriteWork,
                nodes,
                max_nodes,
            ));
        }
        if parts > max_parts {
            return Err(formula_limit_error(
                LimitKind::RewriteWork,
                parts,
                max_parts,
            ));
        }
        Ok(())
    }

    fn charge(&mut self, amount: usize) -> std::result::Result<(), Self::Error> {
        self.charge_output_text(amount)
    }
}

/// Compact source-order node sink used for the scalar generated-free formula
/// subset. The node vector is bounded by the exact wire-preflight node count;
/// no generated repeated-message tree is retained on this route.
struct ScalarFormulaVisitor {
    nodes: Vec<numbers_formula_codec::FormulaNode>,
    node_limit: usize,
    unsupported: bool,
}

impl ScalarFormulaVisitor {
    fn with_capacity(node_limit: usize) -> Result<Self> {
        let mut nodes = Vec::new();
        nodes
            .try_reserve_exact(node_limit)
            .map_err(|_| allocation_error("Numbers scalar formula nodes", node_limit))?;
        Ok(Self {
            nodes,
            node_limit,
            unsupported: false,
        })
    }
}

impl numbers_formula_codec::FormulaVisitor for ScalarFormulaVisitor {
    fn visit_node(
        &mut self,
        node: numbers_formula_codec::FormulaNode,
    ) -> std::result::Result<(), numbers_formula_codec::DecodeError> {
        if matches!(
            node,
            numbers_formula_codec::FormulaNode::LocalCellReference { .. }
                | numbers_formula_codec::FormulaNode::LocalRange { .. }
                | numbers_formula_codec::FormulaNode::CellReference { .. }
                | numbers_formula_codec::FormulaNode::ResolvedCellReference { .. }
                | numbers_formula_codec::FormulaNode::ResolvedRange { .. }
        ) {
            self.unsupported = true;
            return Ok(());
        }
        if self.nodes.len() >= self.node_limit {
            return Err(numbers_formula_codec::DecodeError::allocation(
                self.node_limit.saturating_add(1),
            ));
        }
        self.nodes.push(node);
        Ok(())
    }
}

/// Try the compact generated-free formula visitor. `None` means that the
/// archive is valid but outside the scalar subset (or cannot be resolved
/// against this table's bounds), so the lossless compatibility renderer must be
/// used. The strict archive preflight has already charged the complete wire
/// walk, keeping this bounded fallback probe from admitting an unbounded tree.
fn render_scalar_formula(
    formula: &FormulaArchiveBytes,
    host_row: usize,
    host_column: usize,
    row_count: usize,
    column_count: usize,
    budget: &mut ProjectionBudget,
) -> Result<Option<String>> {
    let owner = 1;
    let host_row = match u32::try_from(host_row) {
        Ok(value) => value,
        Err(_) => return Ok(None),
    };
    let host_column = match u32::try_from(host_column) {
        Ok(value) => value,
        Err(_) => return Ok(None),
    };
    let rows = match u32::try_from(row_count) {
        Ok(value) if value != 0 => value,
        _ => return Ok(None),
    };
    let columns = match u32::try_from(column_count) {
        Ok(value) if value != 0 => value,
        _ => return Ok(None),
    };
    if host_row >= rows || host_column >= columns {
        return Ok(None);
    }

    // The wire preflight derives this exact bound without retaining any AST
    // nodes. Charge it before constructing the visitor so a rejected formula
    // cannot allocate a node vector beyond the package render-work budget.
    budget.charge_formula_render_work(formula.scalar_visitor_node_count)?;

    let max_fields = budget.remaining_payload_fields();
    let max_work = budget.remaining_payload_work();
    let max_text_bytes = budget.remaining_staging_text_bytes();
    // The neutral scalar evaluator has a deliberately smaller wire-depth
    // ceiling than the compatibility renderer's recursive thunk surface.
    let scalar_depth = u32::try_from(budget.max_formula_render_depth)
        .unwrap_or(u32::MAX)
        .min(32);
    let options = numbers_formula_codec::DecodeOptions::new(
        formula.bytes.len(),
        max_fields,
        max_work,
        scalar_depth,
        formula.scalar_visitor_node_count,
        max_text_bytes,
    );
    let context =
        numbers_formula_codec::FormulaContext::new(owner, host_row, host_column, rows, columns);
    let mut visitor = ScalarFormulaVisitor::with_capacity(formula.scalar_visitor_node_count)?;
    let report = match numbers_formula_codec::decode_formula_archive_with_visitor(
        formula.bytes.as_ref(),
        context,
        options,
        &mut visitor,
    ) {
        Ok(report) => report,
        Err(error) => {
            if let Some(numbers_formula_codec::DecodeLimit::Allocation { requested }) =
                error.resource_limit()
            {
                return Err(allocation_error("Numbers scalar formula nodes", requested));
            }
            // The strict envelope admitted every scalar-eligible wire field.
            // A codec error is therefore a semantic/resource refusal, not a
            // compatibility probe. Ending the operation avoids replaying a
            // failed, unreported traversal under the same residual budget.
            return Err(map_formula_render_decode_error(error));
        },
    };
    budget.charge_formula_decode_report(report)?;
    if visitor.unsupported {
        return Ok(None);
    }
    debug_assert_eq!(report.node_count(), formula.scalar_visitor_node_count);
    Ok(Some(render_scalar_formula_nodes(&visitor.nodes, budget)?))
}

fn render_scalar_formula_nodes(
    nodes: &[numbers_formula_codec::FormulaNode],
    budget: &mut ProjectionBudget,
) -> Result<String> {
    if nodes.is_empty() {
        return retain_text("=", budget);
    }
    let mut renderer = FormulaRenderer::default();
    let mut stack = Vec::new();
    stack
        .try_reserve_exact(nodes.len())
        .map_err(|_| allocation_error("Numbers scalar formula expression stack", nodes.len()))?;
    for node in nodes {
        use numbers_formula_codec::{BinaryOperator, FormulaNode};
        let expression = match *node {
            FormulaNode::Binary(operator) => {
                let (symbol, operation) = match operator {
                    BinaryOperator::Add => ("+", "addition"),
                    BinaryOperator::Subtract => ("-", "subtraction"),
                    BinaryOperator::Multiply => ("*", "multiplication"),
                    BinaryOperator::Divide => ("/", "division"),
                    BinaryOperator::Power => ("^", "power"),
                    BinaryOperator::Concatenate => ("&", "concatenation"),
                    BinaryOperator::GreaterThan => (">", "greater than"),
                    BinaryOperator::GreaterThanOrEqual => (">=", "greater than or equal"),
                    BinaryOperator::LessThan => ("<", "less than"),
                    BinaryOperator::LessThanOrEqual => ("<=", "less than or equal"),
                    BinaryOperator::Equal => ("=", "equality"),
                    BinaryOperator::NotEqual => ("<>", "inequality"),
                };
                Some(render_binary(
                    &mut stack,
                    &mut renderer,
                    symbol,
                    operation,
                    true,
                    budget,
                )?)
            },
            FormulaNode::Negation => stack
                .pop()
                .map(|operand| renderer.unary("-(", operand, ")", budget))
                .transpose()?,
            FormulaNode::Percent => {
                let operand = stack.pop().ok_or_else(|| {
                    Error::ParseError(
                        "Numbers formula percent operator is missing an operand".to_owned(),
                    )
                })?;
                Some(renderer.unary("(", operand, ")%", budget)?)
            },
            FormulaNode::Function {
                identifier,
                argument_count,
            } => {
                let arguments = pop_formula_arguments(&mut stack, argument_count, "function")?;
                Some(renderer.comma_joined(
                    Some(fallible_function_name(identifier, &renderer, budget)?),
                    arguments,
                    "(",
                    ")",
                    budget,
                )?)
            },
            FormulaNode::Number { bits } => {
                let value = fallible_formula_display(f64::from_bits(bits), &renderer, budget)?;
                Some(renderer.owned_expr(value, budget)?)
            },
            FormulaNode::Boolean(value) | FormulaNode::Token(value) => {
                Some(renderer.static_expr(if value { "TRUE" } else { "FALSE" }, budget)?)
            },
            FormulaNode::Empty => Some(renderer.static_expr("", budget)?),
            FormulaNode::LocalCell {
                coordinate,
                row_is_sticky,
                column_is_sticky,
            } => {
                let column = FormulaColumn(coordinate.column());
                let row = checked_formula_row_number(coordinate.row())?;
                let value = fallible_formula_format(&renderer, budget, |output| {
                    write!(
                        output,
                        "{}{column}{}{row}",
                        if column_is_sticky != 0 { "$" } else { "" },
                        if row_is_sticky != 0 { "$" } else { "" },
                    )
                })?;
                Some(renderer.owned_expr(value, budget)?)
            },
            FormulaNode::Colon | FormulaNode::ColonWithUids => Some(render_binary(
                &mut stack,
                &mut renderer,
                ":",
                "range",
                false,
                budget,
            )?),
            FormulaNode::PlusSign
            | FormulaNode::AppendWhitespace
            | FormulaNode::PrependWhitespace => None,
            FormulaNode::LocalCellReference { .. }
            | FormulaNode::LocalRange { .. }
            | FormulaNode::CellReference { .. }
            | FormulaNode::ResolvedCellReference { .. }
            | FormulaNode::ResolvedRange { .. } => {
                return Err(Error::InvalidFormat(
                    "Numbers scalar formula visitor received an owner-bearing node".to_owned(),
                ));
            },
        };
        if let Some(expression) = expression {
            stack.push(expression);
        }
    }
    let Some(root) = stack.pop() else {
        // Keep parity with the historical compatibility renderer: an archive
        // containing only ignored postfix markers (or a missing negation
        // operand) falls through to its FORMULA() placeholder.
        return retain_text("=FORMULA()", budget);
    };
    renderer.render(root, budget)
}

fn render_formula_compatibility(
    formula: &FormulaArchiveBytes,
    host_row: usize,
    host_column: usize,
    row_count: usize,
    column_count: usize,
    formula_references: &FormulaReferenceMaps,
    budget: &mut ProjectionBudget,
) -> Result<String> {
    let owner = 1;
    let host_row = u32::try_from(host_row)
        .map_err(|_| Error::ParseError("Numbers formula host row exceeds u32".to_owned()))?;
    let host_column = u32::try_from(host_column)
        .map_err(|_| Error::ParseError("Numbers formula host column exceeds u32".to_owned()))?;
    let rows = u32::try_from(row_count)
        .map_err(|_| Error::ParseError("Numbers formula row count exceeds u32".to_owned()))?;
    let columns = u32::try_from(column_count)
        .map_err(|_| Error::ParseError("Numbers formula column count exceeds u32".to_owned()))?;
    // Keep raw wire recursion independent from the semantic AST depth. The
    // codec validates the former while the visitor enforces the latter.
    let raw_depth = u32::try_from(WireLimits::MAX_NESTING).unwrap_or(u32::MAX);
    let render_depth = u32::try_from(budget.max_formula_render_depth).unwrap_or(u32::MAX);
    let max_fields = budget.remaining_payload_fields();
    let max_work = budget.remaining_payload_work();
    let max_text_bytes = budget.remaining_staging_text_bytes();
    let options = numbers_formula_codec::DecodeOptions::new(
        formula.bytes.len(),
        max_fields,
        max_work,
        raw_depth,
        formula.scalar_visitor_node_count,
        max_text_bytes,
    )
    .with_opaque_unknown_fields(true)
    .with_render_recursion_limit(render_depth);
    let context =
        numbers_formula_codec::FormulaContext::new(owner, host_row, host_column, rows, columns);
    let mut visitor =
        CompatibilityFormulaVisitor::new(host_row, host_column, formula_references, budget);
    let decoded = numbers_formula_codec::decode_formula_archive_for_render(
        formula.bytes.as_ref(),
        context,
        options,
        &mut visitor,
    );
    let report = match decoded {
        Ok(report) => report,
        Err(error) => {
            return Err(visitor
                .error
                .take()
                .unwrap_or_else(|| map_formula_render_decode_error(error)));
        },
    };
    visitor.budget.charge_formula_decode_report(report)?;
    if let Some(error) = visitor.error.take() {
        return Err(error);
    }
    debug_assert_eq!(report.node_count(), formula.scalar_visitor_node_count);
    visitor.finish()
}

fn map_formula_render_decode_error(error: numbers_formula_codec::DecodeError) -> Error {
    use numbers_formula_codec::DecodeLimit;
    match error.resource_limit() {
        Some(DecodeLimit::Bytes { observed, maximum }) => {
            formula_limit_error(LimitKind::InputBytes, observed, maximum)
        },
        Some(DecodeLimit::Fields { observed, maximum }) => {
            formula_limit_error(LimitKind::Fields, observed, maximum)
        },
        Some(DecodeLimit::Work { observed, maximum }) => {
            formula_limit_error(LimitKind::RewriteWork, observed, maximum)
        },
        Some(DecodeLimit::Nesting { observed, maximum }) => {
            formula_limit_error(LimitKind::Nesting, observed as usize, maximum as usize)
        },
        Some(DecodeLimit::Nodes { observed, maximum }) => {
            formula_limit_error(LimitKind::RewriteWork, observed, maximum)
        },
        Some(DecodeLimit::Text { observed, maximum }) => {
            formula_limit_error(LimitKind::OutputBytes, observed, maximum)
        },
        Some(DecodeLimit::Allocation { requested }) => {
            allocation_error("Numbers formula compatibility traversal", requested)
        },
        None => Error::InvalidFormat("malformed Numbers formula payload".to_owned()),
    }
}

fn retain_text(value: &str, budget: &mut ProjectionBudget) -> Result<String> {
    budget.charge_output_text(value.len())?;
    let mut retained = String::new();
    retained
        .try_reserve_exact(value.len())
        .map_err(|_| allocation_error("Numbers rendered formula", value.len()))?;
    retained.push_str(value);
    Ok(retained)
}

/// Render one retained formula archive with the same scalar-first/compatibility
/// selection used by the Numbers package extractor.  Formula wire admission,
/// node work, lazy traversal, and rendered output all debit the caller's
/// aggregate table projection budget.
pub(super) fn render_formula_string(
    formula: &FormulaArchiveBytes,
    host_row: usize,
    host_column: usize,
    row_count: usize,
    column_count: usize,
    formula_references: &FormulaReferenceMaps,
    budget: &mut ProjectionBudget,
) -> Result<String> {
    let render_work_before_scalar = budget.formula_render_work;
    if formula.scalar_visitor_eligible
        && let Some(rendered) = render_scalar_formula(
            formula,
            host_row,
            host_column,
            row_count,
            column_count,
            budget,
        )?
    {
        return Ok(rendered);
    }
    let nodes_precharged = budget
        .formula_render_work
        .saturating_sub(render_work_before_scalar)
        >= formula.scalar_visitor_node_count;
    if !nodes_precharged {
        budget.charge_formula_render_work(formula.scalar_visitor_node_count)?;
    }
    budget.charge_formula_lazy_work(formula.lazy_traversal_entry_count)?;
    render_formula_compatibility(
        formula,
        host_row,
        host_column,
        row_count,
        column_count,
        formula_references,
        budget,
    )
}

struct CompatibilityFormulaArray {
    expressions: Vec<FormulaExpr>,
    had_node: bool,
    is_thunk: bool,
}

struct CompatibilityFormulaVisitor<'references, 'budget> {
    host_row: u32,
    host_column: u32,
    formula_references: &'references FormulaReferenceMaps,
    budget: &'budget mut ProjectionBudget,
    renderer: FormulaRenderer,
    arrays: Vec<CompatibilityFormulaArray>,
    root_expression: Option<FormulaExpr>,
    pending_thunk: bool,
    error: Option<Error>,
}

impl<'references, 'budget> CompatibilityFormulaVisitor<'references, 'budget> {
    fn new(
        host_row: u32,
        host_column: u32,
        formula_references: &'references FormulaReferenceMaps,
        budget: &'budget mut ProjectionBudget,
    ) -> Self {
        Self {
            host_row,
            host_column,
            formula_references,
            budget,
            renderer: FormulaRenderer::default(),
            arrays: Vec::new(),
            root_expression: None,
            pending_thunk: false,
            error: None,
        }
    }

    fn fail(&mut self, error: Error) -> numbers_formula_codec::DecodeError {
        self.error = Some(error);
        numbers_formula_codec::DecodeError::allocation(0)
    }

    fn current_array_mut(&mut self) -> std::result::Result<&mut CompatibilityFormulaArray, Error> {
        self.arrays.last_mut().ok_or_else(|| {
            Error::ParseError("Numbers formula event stream has no active array".to_owned())
        })
    }

    fn mark_node(&mut self) -> std::result::Result<(), Error> {
        self.current_array_mut()?.had_node = true;
        Ok(())
    }

    fn push_expression(&mut self, expression: FormulaExpr) -> std::result::Result<(), Error> {
        let array = self.current_array_mut()?;
        array.expressions.try_reserve(1).map_err(|_error| {
            allocation_error(
                "Numbers formula compatibility expression stack",
                array.expressions.len().saturating_add(1),
            )
        })?;
        array.expressions.push(expression);
        Ok(())
    }

    fn pop_binary(
        &mut self,
        operation: &str,
    ) -> std::result::Result<(FormulaExpr, FormulaExpr), Error> {
        let array = self.current_array_mut()?;
        pop_binary_operands(&mut array.expressions, operation)
    }

    fn pop_arguments(
        &mut self,
        count: u32,
        node_kind: &str,
    ) -> std::result::Result<Vec<FormulaExpr>, Error> {
        let array = self.current_array_mut()?;
        pop_formula_arguments(&mut array.expressions, count, node_kind)
    }

    fn begin_array(&mut self, depth: u32) -> std::result::Result<(), Error> {
        // The codec reports logical AST-array depth (root = 1, each thunk
        // adds one). Check that depth directly so the package limit remains
        // identical to the legacy renderer's semantic recursion bound.
        let semantic_depth = usize::try_from(depth).unwrap_or(usize::MAX);
        self.budget.check_formula_render_depth(semantic_depth)?;
        let is_thunk = self.pending_thunk;
        self.pending_thunk = false;
        self.arrays.try_reserve(1).map_err(|_error| {
            allocation_error(
                "Numbers formula compatibility arrays",
                self.arrays.len().saturating_add(1),
            )
        })?;
        self.arrays.push(CompatibilityFormulaArray {
            expressions: Vec::new(),
            had_node: false,
            is_thunk,
        });
        Ok(())
    }

    fn end_array(&mut self) -> std::result::Result<(), Error> {
        let array = self.arrays.pop().ok_or_else(|| {
            Error::ParseError("Numbers formula event stream ended an inactive array".to_owned())
        })?;
        let expression = if let Some(expression) = array.expressions.last().copied() {
            expression
        } else if array.is_thunk {
            self.renderer
                .static_expr(if array.had_node { "FORMULA()" } else { "" }, self.budget)?
        } else if array.had_node {
            self.renderer.static_expr("FORMULA()", self.budget)?
        } else {
            // The root empty archive is handled by `finish`; retaining no
            // expression here preserves the legacy decoder's `=` result.
            return Ok(());
        };
        if self.arrays.is_empty() {
            self.root_expression = Some(expression);
        } else if array.is_thunk {
            self.push_expression(expression)?;
        } else {
            return Err(Error::ParseError(
                "Numbers formula event stream nested an unexpected array".to_owned(),
            ));
        }
        Ok(())
    }

    fn render_event(
        &mut self,
        event: numbers_formula_codec::FormulaRenderEvent<'_>,
    ) -> std::result::Result<Option<FormulaExpr>, Error> {
        use numbers_formula_codec::{BinaryOperator, FormulaRenderEvent};
        let expression = match event {
            FormulaRenderEvent::Binary(operator) => {
                let (symbol, operation) = match operator {
                    BinaryOperator::Add => ("+", "addition"),
                    BinaryOperator::Subtract => ("-", "subtraction"),
                    BinaryOperator::Multiply => ("*", "multiplication"),
                    BinaryOperator::Divide => ("/", "division"),
                    BinaryOperator::Power => ("^", "power"),
                    BinaryOperator::Concatenate => ("&", "concatenation"),
                    BinaryOperator::GreaterThan => (">", "greater than"),
                    BinaryOperator::GreaterThanOrEqual => (">=", "greater than or equal"),
                    BinaryOperator::LessThan => ("<", "less than"),
                    BinaryOperator::LessThanOrEqual => ("<=", "less than or equal"),
                    BinaryOperator::Equal => ("=", "equality"),
                    BinaryOperator::NotEqual => ("<>", "inequality"),
                };
                let (left, right) = self.pop_binary(operation)?;
                Some(
                    self.renderer
                        .binary(left, symbol, right, true, self.budget)?,
                )
            },
            FormulaRenderEvent::Negation => self
                .current_array_mut()?
                .expressions
                .pop()
                .map(|operand| self.renderer.unary("-(", operand, ")", self.budget))
                .transpose()?,
            FormulaRenderEvent::Percent => {
                let operand = self.current_array_mut()?.expressions.pop().ok_or_else(|| {
                    Error::ParseError(
                        "Numbers formula percent operator is missing an operand".to_owned(),
                    )
                })?;
                Some(self.renderer.unary("(", operand, ")%", self.budget)?)
            },
            FormulaRenderEvent::Number { value } => Some(self.renderer.owned_expr(
                fallible_formula_display(value, &self.renderer, self.budget)?,
                self.budget,
            )?),
            FormulaRenderEvent::String(value) => Some(self.renderer.owned_expr(
                formula_string_literal(value, &self.renderer, self.budget)?,
                self.budget,
            )?),
            FormulaRenderEvent::Boolean(value) | FormulaRenderEvent::Token(value) => Some(
                self.renderer
                    .static_expr(if value { "TRUE" } else { "FALSE" }, self.budget)?,
            ),
            FormulaRenderEvent::Date { value } => {
                let days = value / 86_400.0;
                Some(self.renderer.owned_expr(
                    fallible_formula_format(&self.renderer, self.budget, |output| {
                        write!(output, "(DATE(2001,1,1)+{days})")
                    })?,
                    self.budget,
                )?)
            },
            FormulaRenderEvent::Duration { value } => Some(self.renderer.owned_expr(
                fallible_formula_display(value, &self.renderer, self.budget)?,
                self.budget,
            )?),
            FormulaRenderEvent::EmptyArgument => Some(self.renderer.static_expr("", self.budget)?),
            FormulaRenderEvent::Function {
                identifier,
                argument_count,
            } => {
                let arguments = self.pop_arguments(argument_count, "function")?;
                Some(self.renderer.comma_joined(
                    Some(fallible_function_name(
                        identifier,
                        &self.renderer,
                        self.budget,
                    )?),
                    arguments,
                    "(",
                    ")",
                    self.budget,
                )?)
            },
            FormulaRenderEvent::List { argument_count } => {
                let arguments = self.pop_arguments(argument_count, "list")?;
                Some(
                    self.renderer
                        .comma_joined(None, arguments, "", "", self.budget)?,
                )
            },
            FormulaRenderEvent::Array { columns, rows } => {
                let count = columns.checked_mul(rows).ok_or_else(|| {
                    Error::ParseError("Numbers formula array size overflow".to_owned())
                })?;
                let values = self.pop_arguments(count, "array")?;
                let columns = usize::try_from(columns).map_err(|_| {
                    Error::ParseError("Numbers formula array width exceeds usize".to_owned())
                })?;
                Some(self.renderer.array(values, columns, self.budget)?)
            },
            FormulaRenderEvent::UnknownFunction {
                name,
                argument_count,
            } => {
                let arguments = self.pop_arguments(argument_count, "unknown function")?;
                Some(self.renderer.comma_joined(
                    Some(fallible_formula_owned(
                        name.unwrap_or("UNKNOWN"),
                        &self.renderer,
                        self.budget,
                    )?),
                    arguments,
                    "(",
                    ")",
                    self.budget,
                )?)
            },
            FormulaRenderEvent::CellReference(reference) => {
                Some(self.render_cell_reference(&reference)?)
            },
            FormulaRenderEvent::LocalCellReference(reference) => {
                Some(self.render_standalone_local_cell_reference(reference)?)
            },
            FormulaRenderEvent::CrossTableCellReference(reference) => {
                Some(self.render_cross_table_cell_reference(reference)?)
            },
            FormulaRenderEvent::Colon => Some(self.render_binary("colon", ":", false)?),
            FormulaRenderEvent::ColonWithUids => Some(self.render_binary("range", ":", false)?),
            FormulaRenderEvent::ColonTract(tract) => Some(self.render_colon_tract(&tract)?),
            FormulaRenderEvent::CategoryReference(category) => {
                Some(self.render_category_reference(category)?)
            },
            FormulaRenderEvent::ReferenceError => {
                Some(self.renderer.static_expr("#REF!", self.budget)?)
            },
            FormulaRenderEvent::Ignored { .. }
            | FormulaRenderEvent::PlusSign
            | FormulaRenderEvent::AppendWhitespace
            | FormulaRenderEvent::PrependWhitespace
            | FormulaRenderEvent::BeginArray { .. }
            | FormulaRenderEvent::EndArray
            | FormulaRenderEvent::ThunkBegin
            | FormulaRenderEvent::ThunkEnd => None,
        };
        Ok(expression)
    }

    fn render_binary(
        &mut self,
        operation: &str,
        operator: &'static str,
        wrapped: bool,
    ) -> std::result::Result<FormulaExpr, Error> {
        let (left, right) = self.pop_binary(operation)?;
        self.renderer
            .binary(left, operator, right, wrapped, self.budget)
    }

    fn render_cell_reference(
        &mut self,
        reference: &numbers_formula_codec::FormulaRenderCellReference,
    ) -> std::result::Result<FormulaExpr, Error> {
        if let Some(coordinates) = reference.coordinates {
            let column_absolute = coordinates.column.absolute;
            let row_absolute = coordinates.row.absolute;
            let column = FormulaColumn(resolve_formula_coordinate(
                self.host_column as usize,
                coordinates.column.coordinate,
                column_absolute,
                "column",
            )?);
            let row = checked_formula_row_number(resolve_formula_coordinate(
                self.host_row as usize,
                coordinates.row.coordinate,
                row_absolute,
                "row",
            )?)?;
            let prefix = reference.cross_table_extra.as_ref().and_then(|extra| {
                formula_render_prefix_parts(&extra.table_id, self.formula_references)
            });
            return self.renderer.owned_expr(
                fallible_formula_format(&self.renderer, self.budget, |output| {
                    if reference.cross_table_extra.is_some() {
                        write_formula_reference_prefix(output, prefix)?;
                    }
                    write!(
                        output,
                        "{}{column}{}{row}",
                        if column_absolute { "$" } else { "" },
                        if row_absolute { "$" } else { "" },
                    )
                })?,
                self.budget,
            );
        }
        if let Some(local) = reference.local {
            return self.render_local_cell_reference(Some(local));
        }
        if let Some(cross) = reference.cross_table {
            return self.render_cross_table_cell_reference(Some(cross));
        }
        self.renderer.owned_expr(
            fallible_formula_owned("#REF!", &self.renderer, self.budget)?,
            self.budget,
        )
    }

    fn render_local_cell_reference(
        &mut self,
        reference: Option<numbers_formula_codec::FormulaRenderLocalCellReference>,
    ) -> std::result::Result<FormulaExpr, Error> {
        let Some(reference) = reference else {
            return self.renderer.owned_expr(
                fallible_formula_owned("#REF!", &self.renderer, self.budget)?,
                self.budget,
            );
        };
        let column = FormulaColumn(reference.column_handle);
        let row = checked_formula_row_number(reference.row_handle)?;
        self.renderer.owned_expr(
            fallible_formula_format(&self.renderer, self.budget, |output| {
                write!(
                    output,
                    "{}{column}{}{row}",
                    if reference.column_is_sticky != 0 {
                        "$"
                    } else {
                        ""
                    },
                    if reference.row_is_sticky != 0 {
                        "$"
                    } else {
                        ""
                    },
                )
            })?,
            self.budget,
        )
    }

    /// A standalone LocalCellReferenceNode is rendered by the legacy
    /// generated path without consulting its sticky flags.  Keep that quirk
    /// distinct from CellReferenceNode's nested-local fallback, which does
    /// preserve the flags.
    fn render_standalone_local_cell_reference(
        &mut self,
        reference: Option<numbers_formula_codec::FormulaRenderLocalCellReference>,
    ) -> std::result::Result<FormulaExpr, Error> {
        let Some(reference) = reference else {
            return self.renderer.owned_expr(
                fallible_formula_owned("#REF!", &self.renderer, self.budget)?,
                self.budget,
            );
        };
        let column = FormulaColumn(reference.column_handle);
        let row = checked_formula_row_number(reference.row_handle)?;
        self.renderer.owned_expr(
            fallible_formula_format(&self.renderer, self.budget, |output| {
                write!(output, "{column}{row}")
            })?,
            self.budget,
        )
    }

    fn render_cross_table_cell_reference(
        &mut self,
        reference: Option<numbers_formula_codec::FormulaRenderCrossTableCellReference>,
    ) -> std::result::Result<FormulaExpr, Error> {
        let Some(reference) = reference else {
            return self.renderer.owned_expr(
                fallible_formula_owned("#REF!", &self.renderer, self.budget)?,
                self.budget,
            );
        };
        let prefix = formula_render_prefix_parts(&reference.table_id, self.formula_references);
        let column = FormulaColumn(reference.column_handle);
        let row = checked_formula_row_number(reference.row_handle)?;
        self.renderer.owned_expr(
            fallible_formula_format(&self.renderer, self.budget, |output| {
                write_formula_reference_prefix(output, prefix)?;
                write!(output, "{column}{row}")
            })?,
            self.budget,
        )
    }

    fn render_colon_tract(
        &mut self,
        tract: &numbers_formula_codec::FormulaRenderColonTract,
    ) -> std::result::Result<FormulaExpr, Error> {
        let whole_rows = tract.relative_column.count == 0
            && tract.absolute_column.count == 1
            && tract.absolute_column.first_begin == Some(i16::MAX as i64)
            && tract.absolute_column.first_end.is_none();
        let whole_columns = tract.relative_row.count == 0
            && tract.absolute_row.count == 1
            && tract.absolute_row.first_begin == Some(i32::MAX as i64)
            && tract.absolute_row.first_end.is_none();
        let has_columns =
            !whole_rows && (tract.relative_column.count != 0 || tract.absolute_column.count != 0);
        let has_rows =
            !whole_columns && (tract.relative_row.count != 0 || tract.absolute_row.count != 0);
        if !has_columns && !has_rows {
            return Err(Error::ParseError(
                "Numbers formula colon tract has no row or column coordinates".to_owned(),
            ));
        }
        let columns = if has_columns {
            Some((
                resolve_render_colon_axis(
                    &tract.relative_column,
                    &tract.absolute_column,
                    tract.sticky.begin_column_is_absolute,
                    false,
                    self.host_column as usize,
                    "column",
                )?,
                resolve_render_colon_axis(
                    &tract.relative_column,
                    &tract.absolute_column,
                    tract.sticky.end_column_is_absolute,
                    true,
                    self.host_column as usize,
                    "column",
                )?,
            ))
        } else {
            None
        };
        let rows = if has_rows {
            Some((
                resolve_render_colon_axis(
                    &tract.relative_row,
                    &tract.absolute_row,
                    tract.sticky.begin_row_is_absolute,
                    false,
                    self.host_row as usize,
                    "row",
                )?,
                resolve_render_colon_axis(
                    &tract.relative_row,
                    &tract.absolute_row,
                    tract.sticky.end_row_is_absolute,
                    true,
                    self.host_row as usize,
                    "row",
                )?,
            ))
        } else {
            None
        };
        let rows = rows
            .map(|(begin, end)| {
                Ok::<(u64, u64), Error>((
                    checked_formula_row_number(begin)?,
                    checked_formula_row_number(end)?,
                ))
            })
            .transpose()?;
        let prefix = tract.cross_table_extra.as_ref().and_then(|extra| {
            formula_render_prefix_parts(&extra.table_id, self.formula_references)
        });
        self.renderer.owned_expr(
            fallible_formula_format(&self.renderer, self.budget, |output| {
                if tract.cross_table_extra.is_some() {
                    write_formula_reference_prefix(output, prefix)?;
                }
                if let Some((begin, end)) = columns {
                    if tract.sticky.begin_column_is_absolute {
                        output.write_char('$')?;
                    }
                    write!(output, "{}", FormulaColumn(begin))?;
                    if let Some((row_begin, _)) = rows {
                        if tract.sticky.begin_row_is_absolute {
                            output.write_char('$')?;
                        }
                        write!(output, "{row_begin}")?;
                    }
                    output.write_char(':')?;
                    if tract.sticky.end_column_is_absolute {
                        output.write_char('$')?;
                    }
                    write!(output, "{}", FormulaColumn(end))?;
                    if let Some((_, row_end)) = rows {
                        if tract.sticky.end_row_is_absolute {
                            output.write_char('$')?;
                        }
                        write!(output, "{row_end}")?;
                    }
                } else if let Some((begin, end)) = rows {
                    if tract.sticky.begin_row_is_absolute {
                        output.write_char('$')?;
                    }
                    write!(output, "{begin}")?;
                    output.write_char(':')?;
                    if tract.sticky.end_row_is_absolute {
                        output.write_char('$')?;
                    }
                    write!(output, "{end}")?;
                }
                Ok(())
            })?,
            self.budget,
        )
    }

    fn render_category_reference(
        &mut self,
        category: Option<numbers_formula_codec::FormulaRenderCategoryReference>,
    ) -> std::result::Result<FormulaExpr, Error> {
        let category_uid = category.and_then(|category| {
            category
                .absolute_group_uid
                .or(category.relative_group_uid)
                .or(category.last_group_uid)
        });
        let Some(category_uid) = category_uid else {
            return self.renderer.owned_expr(
                fallible_formula_owned("#CATEGORY!", &self.renderer, self.budget)?,
                self.budget,
            );
        };
        let Some(label) = self
            .formula_references
            .categories
            .get(&[category_uid.lower, category_uid.upper])
        else {
            return self.renderer.owned_expr(
                fallible_formula_owned("#CATEGORY!", &self.renderer, self.budget)?,
                self.budget,
            );
        };
        self.renderer.owned_expr(
            render_category_label_checked(label, &self.renderer, self.budget)?,
            self.budget,
        )
    }

    fn finish(mut self) -> Result<String> {
        if let Some(error) = self.error.take() {
            return Err(error);
        }
        if !self.arrays.is_empty() {
            return Err(Error::ParseError(
                "Numbers formula event stream left an active array".to_owned(),
            ));
        }
        match self.root_expression {
            Some(expression) => self.renderer.render(expression, self.budget),
            None => retain_text("=", self.budget),
        }
    }
}

impl numbers_formula_codec::FormulaRenderVisitor for CompatibilityFormulaVisitor<'_, '_> {
    fn visit(
        &mut self,
        event: numbers_formula_codec::FormulaRenderEvent<'_>,
    ) -> std::result::Result<(), numbers_formula_codec::DecodeError> {
        let result = match event {
            numbers_formula_codec::FormulaRenderEvent::BeginArray { depth } => {
                self.begin_array(depth).map(|_| None)
            },
            numbers_formula_codec::FormulaRenderEvent::EndArray => self.end_array().map(|_| None),
            numbers_formula_codec::FormulaRenderEvent::ThunkBegin => self.mark_node().map(|_| {
                self.pending_thunk = true;
                None
            }),
            numbers_formula_codec::FormulaRenderEvent::ThunkEnd => Ok(None),
            event => self.mark_node().and_then(|()| self.render_event(event)),
        };
        match result {
            Ok(Some(expression)) => match self.push_expression(expression) {
                Ok(()) => Ok(()),
                Err(error) => Err(self.fail(error)),
            },
            Ok(None) => Ok(()),
            Err(error) => Err(self.fail(error)),
        }
    }
}

fn formula_render_prefix_parts<'a>(
    owner: &numbers_formula_codec::FormulaRenderCfuuid,
    references: &'a FormulaReferenceMaps,
) -> FormulaPrefix<'a> {
    let key = [owner.word0?, owner.word1?, owner.word2?, owner.word3?];
    references.owners.get(&key)
}

fn render_category_label_checked(
    label: &str,
    renderer: &FormulaRenderer,
    budget: &ProjectionBudget,
) -> Result<String> {
    let escaped_extra = label
        .bytes()
        .filter(|byte| *byte == b'\\' || *byte == b']')
        .count();
    let required = "#CATEGORY!["
        .len()
        .checked_add(label.len())
        .and_then(|length| length.checked_add(escaped_extra))
        .and_then(|length| length.checked_add(1))
        .ok_or_else(|| formula_output_limit_error(usize::MAX, budget))?;
    renderer.check_additional_owned(required, budget)?;
    let mut output = String::new();
    output
        .try_reserve_exact(required)
        .map_err(|_error| allocation_error("Numbers formula category text", required))?;
    output.push_str("#CATEGORY![");
    for character in label.chars() {
        if character == '\\' || character == ']' {
            output.push('\\');
        }
        output.push(character);
    }
    output.push(']');
    Ok(output)
}

fn resolve_render_colon_axis(
    relative: &numbers_formula_codec::FormulaRenderRangeSummary,
    absolute: &numbers_formula_codec::FormulaRenderRangeSummary,
    is_absolute: bool,
    is_end: bool,
    host: usize,
    axis: &str,
) -> Result<u32> {
    let summary = if is_absolute { absolute } else { relative };
    let stored = if is_end {
        summary.first_end.or(summary.first_begin)
    } else {
        summary.first_begin
    }
    .ok_or_else(|| {
        Error::ParseError(format!(
            "Numbers formula colon tract has no {} {} coordinate",
            if is_absolute { "absolute" } else { "relative" },
            axis
        ))
    })?;
    if is_absolute {
        u32::try_from(stored).map_err(|_| {
            Error::ParseError(format!(
                "Numbers formula colon tract absolute {axis} coordinate is out of range"
            ))
        })
    } else {
        let stored = i32::try_from(stored).map_err(|_| {
            Error::ParseError(format!(
                "Numbers formula colon tract relative {axis} coordinate is out of range"
            ))
        })?;
        resolve_formula_coordinate(host, stored, false, axis)
    }
}
fn fallible_formula_owned(
    value: &str,
    renderer: &FormulaRenderer,
    budget: &ProjectionBudget,
) -> Result<String> {
    renderer.check_additional_owned(value.len(), budget)?;
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|_error| allocation_error("Numbers formula owned text", value.len()))?;
    owned.push_str(value);
    Ok(owned)
}

fn fallible_formula_display(
    value: impl std::fmt::Display,
    renderer: &FormulaRenderer,
    budget: &ProjectionBudget,
) -> Result<String> {
    #[derive(Default)]
    struct Counter {
        bytes: usize,
    }
    impl std::fmt::Write for Counter {
        fn write_str(&mut self, value: &str) -> std::fmt::Result {
            self.bytes = self.bytes.checked_add(value.len()).ok_or(std::fmt::Error)?;
            Ok(())
        }
    }
    let mut counter = Counter::default();
    write!(&mut counter, "{value}")
        .map_err(|_error| formula_output_limit_error(usize::MAX, budget))?;
    renderer.check_additional_owned(counter.bytes, budget)?;
    let mut output = String::new();
    output
        .try_reserve_exact(counter.bytes)
        .map_err(|_error| allocation_error("Numbers formula owned text", counter.bytes))?;
    write!(&mut output, "{value}")
        .map_err(|_error| Error::InvalidFormat("Numbers formula formatting failed".to_owned()))?;
    Ok(output)
}

fn fallible_formula_format(
    renderer: &FormulaRenderer,
    budget: &ProjectionBudget,
    write_value: impl Fn(&mut dyn std::fmt::Write) -> std::fmt::Result,
) -> Result<String> {
    #[derive(Default)]
    struct Counter(usize);
    impl std::fmt::Write for Counter {
        fn write_str(&mut self, value: &str) -> std::fmt::Result {
            self.0 = self.0.checked_add(value.len()).ok_or(std::fmt::Error)?;
            Ok(())
        }
    }
    let mut counter = Counter::default();
    write_value(&mut counter).map_err(|_error| formula_output_limit_error(usize::MAX, budget))?;
    renderer.check_additional_owned(counter.0, budget)?;
    let mut output = String::new();
    output
        .try_reserve_exact(counter.0)
        .map_err(|_error| allocation_error("Numbers formula owned text", counter.0))?;
    write_value(&mut output)
        .map_err(|_error| Error::InvalidFormat("Numbers formula formatting failed".to_owned()))?;
    Ok(output)
}

#[derive(Clone, Copy)]
struct FormulaColumn(u32);

impl std::fmt::Display for FormulaColumn {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let mut bytes = [0_u8; 7];
        let mut cursor = bytes.len();
        let mut value = self.0;
        loop {
            cursor -= 1;
            bytes[cursor] = b'A' + u8::try_from(value % 26).map_err(|_error| std::fmt::Error)?;
            if value < 26 {
                break;
            }
            value = value / 26 - 1;
        }
        formatter
            .write_str(std::str::from_utf8(&bytes[cursor..]).map_err(|_error| std::fmt::Error)?)
    }
}

fn fallible_function_name(
    index: u32,
    renderer: &FormulaRenderer,
    budget: &ProjectionBudget,
) -> Result<String> {
    if let Some(name) = super::function_map::function_name(index) {
        fallible_formula_owned(name, renderer, budget)
    } else {
        fallible_formula_format(renderer, budget, |output| write!(output, "FUNC{index}"))
    }
}

type FormulaPrefix<'a> = Option<&'a FormulaReferenceName>;

fn write_formula_reference_prefix(
    output: &mut dyn std::fmt::Write,
    prefix: FormulaPrefix<'_>,
) -> std::fmt::Result {
    if let Some(name) = prefix {
        write!(output, "{}::{}::", name.sheet, name.table)
    } else {
        output.write_str("Table::")
    }
}

fn render_binary(
    stack: &mut Vec<FormulaExpr>,
    renderer: &mut FormulaRenderer,
    operator: &'static str,
    operation: &str,
    wrapped: bool,
    budget: &ProjectionBudget,
) -> Result<FormulaExpr> {
    let (left, right) = pop_binary_operands(stack, operation)?;
    renderer.binary(left, operator, right, wrapped, budget)
}

fn formula_string_literal(
    value: &str,
    renderer: &FormulaRenderer,
    budget: &ProjectionBudget,
) -> Result<String> {
    let quote_count = value.bytes().filter(|byte| *byte == b'"').count();
    let length = value
        .len()
        .checked_add(quote_count)
        .and_then(|length| length.checked_add(2))
        .ok_or_else(|| allocation_error("Numbers formula string literal", usize::MAX))?;
    renderer.check_additional_owned(length, budget)?;
    let mut literal = String::new();
    literal
        .try_reserve_exact(length)
        .map_err(|_error| allocation_error("Numbers formula string literal", length))?;
    literal.push('"');
    for character in value.chars() {
        if character == '"' {
            literal.push('"');
        }
        literal.push(character);
    }
    literal.push('"');
    Ok(literal)
}

fn resolve_formula_coordinate(host: usize, stored: i32, absolute: bool, axis: &str) -> Result<u32> {
    let coordinate = if absolute {
        i64::from(stored)
    } else {
        i64::try_from(host)
            .map_err(|_error| {
                Error::ParseError(format!("Numbers formula host {axis} exceeds i64"))
            })?
            .checked_add(i64::from(stored))
            .ok_or_else(|| Error::ParseError(format!("Numbers formula {axis} overflow")))?
    };
    u32::try_from(coordinate).map_err(|_error| {
        Error::ParseError(format!(
            "Numbers formula {axis} coordinate {coordinate} is out of range"
        ))
    })
}

fn checked_formula_row_number(row: u32) -> Result<u64> {
    u64::from(row)
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("Numbers formula row coordinate overflow".to_owned()))
}

fn pop_binary_operands<T>(stack: &mut Vec<T>, operation: &str) -> Result<(T, T)> {
    let right = stack.pop().ok_or_else(|| {
        Error::ParseError(format!(
            "Malformed Numbers formula: {operation} is missing its right operand"
        ))
    })?;
    let left = stack.pop().ok_or_else(|| {
        Error::ParseError(format!(
            "Malformed Numbers formula: {operation} is missing its left operand"
        ))
    })?;
    Ok((left, right))
}

fn pop_formula_arguments<T>(stack: &mut Vec<T>, count: u32, node_kind: &str) -> Result<Vec<T>> {
    let count = usize::try_from(count).map_err(|_| {
        Error::ParseError(format!(
            "Numbers formula {node_kind} argument count exceeds usize"
        ))
    })?;
    let start = stack.len().checked_sub(count).ok_or_else(|| {
        Error::ParseError(format!(
            "Malformed Numbers formula: {node_kind} requires {count} arguments but only {} are available",
            stack.len()
        ))
    })?;
    let mut arguments = Vec::new();
    arguments
        .try_reserve_exact(count)
        .map_err(|_| allocation_error("Numbers formula arguments", count))?;
    arguments.extend(stack.drain(start..));
    Ok(arguments)
}
