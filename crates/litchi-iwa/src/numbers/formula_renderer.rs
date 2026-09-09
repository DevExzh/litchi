//! Generated-free Numbers FormulaArchive retention and rendering.
//!
//! This private adapter deliberately keeps the historical generated
//! `TSCE.FormulaArchive` representation out of production. Formula list
//! entries are strictly preflighted and retained as bounded wire bytes; the
//! compatibility renderer consumes borrowed events from the neutral formula
//! codec only when a cell references an entry.

use super::table_extractor::{FormulaReferenceMaps, ProjectionBudget};
use crate::{Error, Result};
use litchi_iwa_common::formula::render::FormulaRenderBudget;
use litchi_iwa_common::wire::{
    WireDescent, parse_wire_view_with_limits, preflight_wire_tree_with_limits,
};
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_protos::numbers_formula_codec;
use litchi_numbers_wire::formula_render::{
    self as shared_formula_render, FormulaCategoryId, FormulaEventRenderBudget,
    FormulaRenderCodecVisitor, FormulaTablePrefix, ReferenceResolver,
};

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
    formula_references: &FormulaReferenceMaps,
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
    Ok(Some(shared_formula_render::render_scalar_formula_nodes(
        &visitor.nodes,
        formula_references,
        budget,
    )?))
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
    // codec validates the former while the shared visitor enforces the latter.
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
        FormulaRenderCodecVisitor::new(host_row, host_column, formula_references, budget);
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
                .take_error()
                .unwrap_or_else(|| map_formula_render_decode_error(error)));
        },
    };
    visitor.budget_mut().charge_formula_decode_report(report)?;
    if let Some(error) = visitor.take_error() {
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
            formula_references,
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
impl FormulaEventRenderBudget for ProjectionBudget {
    fn check_render_depth(&self, depth: usize) -> std::result::Result<(), Self::Error> {
        self.check_formula_render_depth(depth)
    }

    fn parse_error(&self, message: String) -> Self::Error {
        Error::ParseError(message)
    }

    fn invalid_format(&self, message: String) -> Self::Error {
        Error::InvalidFormat(message)
    }
}

impl ReferenceResolver for FormulaReferenceMaps {
    fn table_prefix(
        &self,
        id: &numbers_formula_codec::FormulaRenderCfuuid,
    ) -> Option<FormulaTablePrefix<'_>> {
        let key = [id.word0?, id.word1?, id.word2?, id.word3?];
        let name = self.owners.get(&key)?;
        Some(FormulaTablePrefix {
            sheet: name.sheet.as_str(),
            table: name.table.as_str(),
        })
    }

    fn category_name(&self, id: FormulaCategoryId) -> Option<&str> {
        self.categories
            .get(&[id.lower, id.upper])
            .map(String::as_str)
    }

    fn function_name(&self, index: u32) -> Option<&str> {
        super::function_map::function_name(index)
    }
}
