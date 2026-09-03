//! Table Data Extraction from TST Protobuf Messages
//!
//! This module provides utilities for extracting cell data from Numbers table structures.
//! Numbers stores table data in a complex format using Tiles, TableDataList, and Cell messages.
//!
//! ## Architecture
//!
//! - **TableModelArchive**: Contains table metadata and references to data stores
//! - **DataStore**: Contains references to various data tables (strings, formulas, styles)
//! - **TableDataList**: Maps keys to actual cell content (strings, formulas, formats)
//! - **TileStorage**: Contains the actual cells in a sparse tile-based structure
//! - **Tile**: Contains rows of cells with their values
//!
//! ## Public boundary
//!
//! `TableDataExtractor`, `Components`, and `Index` are implementation details
//! of [`crate::Package`]. Applications should parse a native package through
//! [`crate::Package`] and consume its archive-free [`crate::Document`] result;
//! this decoder is intentionally not part of the public API.

use super::Components;
use super::names;
use super::table::Table;
use super::{
    Error, Result, SemanticLimitKind, SemanticLimits, SemanticPath, TABLE_MODEL_MESSAGE_TYPE,
    table_info_decode_options,
};
use super::{Index, Resolved};
use crate::DEFAULT_MAX_TEXT_BYTES;
use crate::cell::FiniteF64;
use crate::cell::Value as CellValue;
use crate::cell::wire::{BncCellView, CachedScalar, StoredValue};
use litchi_iwa_common::comment::{AuthorId, Comment, StorageId, Uuid};
use litchi_iwa_common::wire::{
    WireDescent, parse_wire_view_with_limits, preflight_wire_tree_with_limits,
};
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_protos::comment_storage_codec;
use litchi_iwa_protos::group_node_category_codec::{self, CategoryValueView, GroupNodeView};
use litchi_iwa_protos::table_info_codec;
#[cfg(test)]
use litchi_iwa_protos::tsce;
use litchi_iwa_protos::{numbers_formula_codec, numbers_table_cell_storage_codec, tst};
#[cfg(test)]
use prost::Message as _;
use std::borrow::Cow;
use std::cell::RefCell;
use std::collections::{HashMap, HashSet};
use std::fmt::Write as _;
use std::sync::Arc;

type CompactTable<T> = Box<[(u32, T)]>;
type StringTable = CompactTable<String>;
type FormulaTable = CompactTable<FormulaArchiveBytes>;
type FormulaErrorTable = CompactTable<String>;
type CommentTable = CompactTable<Comment>;
type FormulaOwnerKey = [u32; 4];
type FormulaCategoryKey = [u64; 2];
const RICH_TEXT_PAYLOAD_MESSAGE_TYPE: u32 = 6_218;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;

const TILE_MESSAGE_TYPE: u32 = 6_002;
const MAX_TABLE_ROWS: usize = 1 << 20;
const MAX_TABLE_COLUMNS: usize = 1 << 14;
const MAX_ADDRESSABLE_CELLS: usize = 1 << 24;
const MAX_TABLE_MATERIALIZED_CELLS: usize = 1 << 20;
const MAX_FORMULA_CATEGORY_DEPTH: usize = 64;
const MAX_FORMULA_WORK: usize = crate::MAX_REFERENCES;
const MAX_FORMULA_WIRE_BYTES: usize = DEFAULT_MAX_TEXT_BYTES;
const MAX_PAYLOAD_WORK: usize = WireLimits::MAX_REWRITE_WORK;

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
struct FormulaArchiveBytes {
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
    fn from_wire(source: &[u8], budget: &mut ProjectionBudget) -> Result<Self> {
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
    let max_fields = crate::MAX_REFERENCES
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
        if let Some(expected_wire_type) = schema.wire_type {
            if field.wire_type() != expected_wire_type {
                return Err(litchi_iwa_common::Error::InvalidFormat(format!(
                    "Numbers FormulaArchive field {} has wire type {}, expected {}",
                    field.number(),
                    field.wire_type(),
                    expected_wire_type
                )));
            }
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
                if node_type
                    .is_none_or(|value| !numbers_formula_codec::is_scalar_visitor_node_type(value))
                {
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
            Err(Error::MalformedPayload {
                path: SemanticPath::StructuredTables,
            })
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
                    limit: crate::MAX_REFERENCES,
                })?;
        let observed = self.base_fields.saturating_add(self.fields);
        if observed > crate::MAX_REFERENCES {
            return Err(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed,
                limit: crate::MAX_REFERENCES,
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
        } => Error::SemanticLimit {
            kind: SemanticLimitKind::FormulaWireBytes,
            observed,
            maximum: limit,
            path: SemanticPath::StructuredTables,
        },
        litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Fields,
            observed,
            limit,
        } => Error::SemanticLimit {
            kind: SemanticLimitKind::Objects,
            observed,
            maximum: limit,
            path: SemanticPath::StructuredTables,
        },
        litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Nesting,
            observed,
            limit,
        } => Error::SemanticLimit {
            kind: SemanticLimitKind::FormulaDepth,
            observed,
            maximum: limit,
            path: SemanticPath::StructuredTables,
        },
        litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::RewriteWork,
            observed,
            limit,
        } => Error::SemanticLimit {
            kind: SemanticLimitKind::FormulaWork,
            observed,
            maximum: limit,
            path: SemanticPath::StructuredTables,
        },
        litchi_iwa_common::Error::Allocation { resource, amount } => {
            Error::Common(litchi_iwa_common::Error::Allocation { resource, amount })
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => Error::MalformedPayload {
            path: SemanticPath::StructuredTables,
        },
        other => Error::Common(other),
    }
}

fn checked_formula_work_product(source_len: usize, passes: usize, maximum: usize) -> Result<usize> {
    source_len
        .checked_mul(passes)
        .ok_or_else(|| formula_semantic_limit(SemanticLimitKind::FormulaWork, usize::MAX, maximum))
}

fn table_cell_decode_options(
    source: &[u8],
    max_references: usize,
    max_text_bytes: usize,
    max_fields: usize,
    max_work: usize,
) -> numbers_table_cell_storage_codec::DecodeOptions {
    numbers_table_cell_storage_codec::DecodeOptions::new(
        source.len().max(1),
        max_fields,
        max_work,
        u32::try_from(WireLimits::MAX_NESTING).unwrap_or(u32::MAX),
        max_references,
        max_text_bytes,
    )
}

/// Route one list candidate using only its scalar envelope.
///
/// The Numbers storage codec performs the bounded handwritten/Buffa parity
/// pass; this adapter charges its report before the caller decides whether
/// entry values may be staged.  In particular, a wrong-list candidate cannot
/// resolve rich text/comments, and a referenced segment is accepted solely by
/// the bytes found in that object (the object message type is not schema
/// evidence).
fn probe_table_data_list_type(
    source: &[u8],
    segment: bool,
    budget: &mut ProjectionBudget,
) -> Result<i32> {
    let options = table_cell_decode_options(
        source,
        usize::MAX,
        usize::MAX,
        budget.remaining_payload_fields(),
        budget.remaining_payload_work(),
    );
    let field_offset = budget.payload_fields;
    let work_offset = budget.payload_work;
    let decoded = if segment {
        numbers_table_cell_storage_codec::decode_table_data_list_segment_type_with_report(
            source, options,
        )
    } else {
        numbers_table_cell_storage_codec::decode_table_data_list_type_with_report(source, options)
    }
    .map_err(|error| {
        map_table_cell_codec_error_with_offsets(error, 0, field_offset, work_offset, 0)
    })?;
    let (snapshot, report) = decoded;
    budget.charge_decode_work(report)?;
    Ok(snapshot.list_type())
}

fn map_table_cell_codec_error(error: numbers_table_cell_storage_codec::DecodeError) -> Error {
    map_table_cell_codec_error_with_reference_offset(error, 0)
}

fn map_table_cell_codec_error_with_reference_offset(
    error: numbers_table_cell_storage_codec::DecodeError,
    reference_offset: usize,
) -> Error {
    map_table_cell_codec_error_with_offsets(error, reference_offset, 0, 0, 0)
}

fn map_table_cell_codec_error_with_offsets(
    error: numbers_table_cell_storage_codec::DecodeError,
    reference_offset: usize,
    field_offset: usize,
    work_offset: usize,
    text_offset: usize,
) -> Error {
    let Some(limit) = error.resource_limit() else {
        return Error::InvalidFormat("Numbers table storage projection is invalid".to_owned());
    };
    map_table_cell_decode_limit_with_offsets(
        limit,
        reference_offset,
        field_offset,
        work_offset,
        text_offset,
    )
}

fn map_table_cell_decode_limit_with_reference_offset(
    limit: numbers_table_cell_storage_codec::DecodeLimit,
    reference_offset: usize,
) -> Error {
    map_table_cell_decode_limit_with_offsets(limit, reference_offset, 0, 0, 0)
}

fn map_table_cell_decode_limit_with_offsets(
    limit: numbers_table_cell_storage_codec::DecodeLimit,
    reference_offset: usize,
    field_offset: usize,
    work_offset: usize,
    text_offset: usize,
) -> Error {
    use numbers_table_cell_storage_codec::DecodeLimit;
    let (kind, observed, maximum) = match limit {
        DecodeLimit::Bytes { observed, maximum } => {
            (SemanticLimitKind::FormulaWireBytes, observed, maximum)
        },
        DecodeLimit::References { observed, maximum } => (
            SemanticLimitKind::References,
            observed.saturating_add(reference_offset),
            maximum.saturating_add(reference_offset),
        ),
        DecodeLimit::Text { observed, maximum } => (
            SemanticLimitKind::TextBytes,
            observed.saturating_add(text_offset),
            maximum.saturating_add(text_offset),
        ),
        DecodeLimit::Fields { observed, maximum } => (
            SemanticLimitKind::Objects,
            observed.saturating_add(field_offset),
            maximum.saturating_add(field_offset),
        ),
        DecodeLimit::Work { observed, maximum } => (
            SemanticLimitKind::FormulaWork,
            observed.saturating_add(work_offset),
            maximum.saturating_add(work_offset),
        ),
        DecodeLimit::Nesting { observed, maximum } => (
            SemanticLimitKind::FormulaDepth,
            observed as usize,
            maximum as usize,
        ),
        DecodeLimit::Allocation { requested } => {
            return Error::Common(litchi_iwa_common::Error::Allocation {
                resource: "Numbers table storage projection",
                amount: requested,
            });
        },
        _ => {
            return Error::InvalidFormat("Numbers table storage projection is invalid".to_owned());
        },
    };
    Error::SemanticLimit {
        kind,
        observed,
        maximum,
        path: SemanticPath::Package,
    }
}

fn map_comment_storage_codec_error(
    error: comment_storage_codec::DecodeError,
    reference_offset: usize,
    field_offset: usize,
    work_offset: usize,
    output_text_offset: usize,
) -> Error {
    let Some(limit) = error.resource_limit() else {
        return Error::MalformedPayload {
            path: SemanticPath::StructuredTables,
        };
    };
    use comment_storage_codec::DecodeLimit;
    let (kind, observed, maximum) = match limit {
        DecodeLimit::Bytes { observed, maximum } => {
            (SemanticLimitKind::FormulaWireBytes, observed, maximum)
        },
        DecodeLimit::References { observed, maximum } => (
            SemanticLimitKind::References,
            observed.saturating_add(reference_offset),
            maximum.saturating_add(reference_offset),
        ),
        DecodeLimit::Text { observed, maximum } => (
            SemanticLimitKind::OutputTextBytes,
            observed.saturating_add(output_text_offset),
            maximum.saturating_add(output_text_offset),
        ),
        DecodeLimit::Fields { observed, maximum } => (
            SemanticLimitKind::Objects,
            observed.saturating_add(field_offset),
            maximum.saturating_add(field_offset),
        ),
        DecodeLimit::Work { observed, maximum } => (
            SemanticLimitKind::FormulaWork,
            observed.saturating_add(work_offset),
            maximum.saturating_add(work_offset),
        ),
        DecodeLimit::Nesting { observed, maximum } => (
            SemanticLimitKind::FormulaDepth,
            observed as usize,
            maximum as usize,
        ),
        _ => {
            return Error::MalformedPayload {
                path: SemanticPath::StructuredTables,
            };
        },
    };
    Error::SemanticLimit {
        kind,
        observed,
        maximum,
        path: SemanticPath::Package,
    }
}

struct CellTables<'a> {
    strings: &'a StringTable,
    formulas: &'a FormulaTable,
    formula_errors: &'a FormulaErrorTable,
    rich_text: &'a StringTable,
    comments: Option<&'a CommentTable>,
    formula_references: &'a FormulaReferenceMaps,
}

struct ParsedCell {
    value: CellValue,
    comment_identifier: Option<u32>,
}

#[derive(Debug)]
struct CellBudget {
    remaining: usize,
}

impl CellBudget {
    fn new() -> Self {
        Self {
            remaining: MAX_TABLE_MATERIALIZED_CELLS,
        }
    }

    fn check(&self, requested: usize) -> Result<()> {
        if requested > self.remaining {
            return Err(Error::Common(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::MaterializedCells,
                observed: MAX_TABLE_MATERIALIZED_CELLS
                    .saturating_sub(self.remaining)
                    .saturating_add(requested),
                limit: MAX_TABLE_MATERIALIZED_CELLS,
            }));
        }
        Ok(())
    }

    fn consume(&mut self, materialized: usize) -> Result<()> {
        self.check(materialized)?;
        self.remaining -= materialized;
        Ok(())
    }
}

#[derive(Debug, Clone, Copy)]
struct ProjectionBudget {
    references: usize,
    payload_fields: usize,
    payload_work: usize,
    staging_text_bytes: usize,
    formula_wire_bytes: usize,
    materialized_cells: usize,
    output_text_bytes: usize,
    formula_render_work: usize,
    max_materialized_cells: usize,
    max_output_text_bytes: usize,
    max_formula_render_work: usize,
    max_formula_render_depth: usize,
    max_references: usize,
}

impl ProjectionBudget {
    const fn new(limits: SemanticLimits) -> Self {
        Self {
            references: 0,
            payload_fields: 0,
            payload_work: 0,
            staging_text_bytes: 0,
            formula_wire_bytes: 0,
            materialized_cells: 0,
            output_text_bytes: 0,
            formula_render_work: 0,
            max_materialized_cells: limits.max_materialized_cells(),
            max_output_text_bytes: limits.max_output_text_bytes(),
            max_formula_render_work: limits.max_formula_render_work(),
            max_formula_render_depth: limits.max_formula_render_depth(),
            max_references: limits.max_references(),
        }
    }

    fn charge_materialized_cells(&mut self, amount: usize) -> Result<()> {
        self.materialized_cells = projection_charge(
            self.materialized_cells,
            amount,
            self.max_materialized_cells,
            SemanticLimitKind::MaterializedCells,
        )?;
        Ok(())
    }

    fn check_materialized_cells(&self, amount: usize) -> Result<()> {
        projection_charge(
            self.materialized_cells,
            amount,
            self.max_materialized_cells,
            SemanticLimitKind::MaterializedCells,
        )?;
        Ok(())
    }

    const fn remaining_materialized_cells(&self) -> usize {
        self.max_materialized_cells
            .saturating_sub(self.materialized_cells)
    }

    const fn remaining_output_text_bytes(&self) -> usize {
        self.max_output_text_bytes
            .saturating_sub(self.output_text_bytes)
    }

    const fn remaining_references(&self) -> usize {
        self.max_references.saturating_sub(self.references)
    }

    fn charge_references(&mut self, amount: usize) -> Result<()> {
        self.references = projection_charge(
            self.references,
            amount,
            self.max_references,
            SemanticLimitKind::References,
        )?;
        Ok(())
    }

    const fn remaining_payload_fields(&self) -> usize {
        crate::MAX_REFERENCES.saturating_sub(self.payload_fields)
    }

    const fn remaining_payload_work(&self) -> usize {
        MAX_PAYLOAD_WORK.saturating_sub(self.payload_work)
    }

    const fn remaining_staging_text_bytes(&self) -> usize {
        DEFAULT_MAX_TEXT_BYTES.saturating_sub(self.staging_text_bytes)
    }

    const fn remaining_formula_wire_bytes(&self) -> usize {
        MAX_FORMULA_WIRE_BYTES.saturating_sub(self.formula_wire_bytes)
    }

    fn charge_staging_text(&mut self, amount: usize) -> Result<()> {
        self.staging_text_bytes = projection_charge(
            self.staging_text_bytes,
            amount,
            DEFAULT_MAX_TEXT_BYTES,
            SemanticLimitKind::TextBytes,
        )?;
        Ok(())
    }

    fn charge_formula_wire(&mut self, amount: usize) -> Result<()> {
        self.formula_wire_bytes = projection_charge(
            self.formula_wire_bytes,
            amount,
            MAX_FORMULA_WIRE_BYTES,
            SemanticLimitKind::FormulaWireBytes,
        )?;
        Ok(())
    }

    /// Retain only the non-transactional portion of a failed formula-map
    /// candidate. Formula discovery may have scanned wire and spent work
    /// before a later field, name, or retained-entry check rejects the map;
    /// those counters must remain monotonic, while the map's entries and text
    /// stay unpublished with the rest of the rejected candidate.
    fn retain_formula_map_cost(&mut self, work_items: usize, wire_bytes: usize) {
        self.payload_work = self
            .payload_work
            .saturating_add(work_items)
            .min(MAX_PAYLOAD_WORK);
        self.formula_wire_bytes = self
            .formula_wire_bytes
            .saturating_add(wire_bytes)
            .min(MAX_FORMULA_WIRE_BYTES);
    }

    fn charge_wire_preflight(
        &mut self,
        report: litchi_iwa_common::wire::WirePreflight,
    ) -> Result<()> {
        self.payload_fields = projection_charge(
            self.payload_fields,
            report.fields(),
            crate::MAX_REFERENCES,
            SemanticLimitKind::Objects,
        )?;
        self.payload_work = projection_charge(
            self.payload_work,
            report.scanned_bytes(),
            MAX_PAYLOAD_WORK,
            SemanticLimitKind::FormulaWork,
        )?;
        Ok(())
    }

    /// Retain the fields/work spent by a rejected raw formula preflight.
    ///
    /// Formula bytes are owned transactionally, but the wire walk itself is
    /// an aggregate admission cost. Saturation records a failed boundary
    /// without publishing any staged formula value.
    fn retain_formula_preflight_cost(&mut self, fields: usize, work: usize) {
        self.payload_fields = self
            .payload_fields
            .saturating_add(fields)
            .min(crate::MAX_REFERENCES);
        self.payload_work = self.payload_work.saturating_add(work).min(MAX_PAYLOAD_WORK);
    }

    fn charge_decode_report(
        &mut self,
        report: numbers_table_cell_storage_codec::DecodeReport,
    ) -> Result<()> {
        // Admit the complete report as one transaction. A referenced segment
        // can cross any one of the aggregate ceilings after an earlier root
        // has already consumed part of it; publishing references/fields before
        // work fails would leave the next segment with a budget that no longer
        // describes the successful projection.
        let mut next = *self;
        next.charge_references(report.references())?;
        next.payload_fields = projection_charge(
            next.payload_fields,
            report.fields(),
            crate::MAX_REFERENCES,
            SemanticLimitKind::Objects,
        )?;
        next.payload_work = projection_charge(
            next.payload_work,
            report.work_bytes(),
            MAX_PAYLOAD_WORK,
            SemanticLimitKind::FormulaWork,
        )?;
        *self = next;
        Ok(())
    }

    /// Atomically admit one selected root/segment report and its staged text.
    ///
    /// `DecodeReport::text_bytes` is charged separately from generic reports
    /// because model/tile reports use different semantic text owners. List
    /// roots and segments, however, publish both counters together; keeping
    /// this operation transactional prevents a text-limit refusal from
    /// retaining the report's references, fields, or work.
    fn charge_table_list_decode_report(
        &mut self,
        report: numbers_table_cell_storage_codec::DecodeReport,
    ) -> Result<()> {
        let mut next = *self;
        next.charge_decode_report(report)?;
        next.charge_staging_text(report.text_bytes())?;
        *self = next;
        Ok(())
    }

    fn charge_comment_decode_report(
        &mut self,
        report: comment_storage_codec::DecodeReport,
    ) -> Result<()> {
        let mut next = *self;
        next.charge_references(report.references())?;
        next.payload_fields = projection_charge(
            next.payload_fields,
            report.fields(),
            crate::MAX_REFERENCES,
            SemanticLimitKind::Objects,
        )?;
        next.payload_work = projection_charge(
            next.payload_work,
            report.work_bytes(),
            MAX_PAYLOAD_WORK,
            SemanticLimitKind::FormulaWork,
        )?;
        *self = next;
        Ok(())
    }

    fn charge_decode_work(
        &mut self,
        report: numbers_table_cell_storage_codec::DecodeReport,
    ) -> Result<()> {
        let mut next = *self;
        next.payload_fields = projection_charge(
            next.payload_fields,
            report.fields(),
            crate::MAX_REFERENCES,
            SemanticLimitKind::Objects,
        )?;
        next.payload_work = projection_charge(
            next.payload_work,
            report.work_bytes(),
            MAX_PAYLOAD_WORK,
            SemanticLimitKind::FormulaWork,
        )?;
        *self = next;
        Ok(())
    }

    /// Admit non-AST repeated entries before the lazy FormulaArchive
    /// compatibility traversal. This preserves the legacy per-decode work
    /// charge without retaining generated vectors in the cell extractor.
    fn charge_formula_lazy_work(&mut self, amount: usize) -> Result<()> {
        self.payload_work = projection_charge(
            self.payload_work,
            amount,
            MAX_PAYLOAD_WORK,
            SemanticLimitKind::FormulaWork,
        )?;
        Ok(())
    }

    fn check_output_text(&self, amount: usize) -> Result<()> {
        projection_charge(
            self.output_text_bytes,
            amount,
            self.max_output_text_bytes,
            SemanticLimitKind::OutputTextBytes,
        )?;
        Ok(())
    }

    fn charge_output_text(&mut self, amount: usize) -> Result<()> {
        self.output_text_bytes = projection_charge(
            self.output_text_bytes,
            amount,
            self.max_output_text_bytes,
            SemanticLimitKind::OutputTextBytes,
        )?;
        Ok(())
    }

    /// Admit one owned comment copy before it is attached to a cell.
    ///
    /// The comment sidecar owns one decoded value, but every cell reference
    /// needs an owned value in the current table representation.  Account for
    /// that copy's text, reply identities, and object slot independently of
    /// the source-side comment decode so repeated cell references cannot
    /// amplify memory past the package projection limits.
    fn charge_comment_materialization(&mut self, comment: &Comment) -> Result<()> {
        self.charge_materialized_cells(1)?;
        self.charge_output_text(comment.text.len())?;
        self.charge_references(comment.reply_ids.len())?;
        Ok(())
    }

    fn charge_formula_render_work(&mut self, amount: usize) -> Result<()> {
        let charged = projection_charge(
            self.formula_render_work,
            amount,
            self.max_formula_render_work,
            SemanticLimitKind::FormulaRenderWork,
        );
        match charged {
            Ok(observed) => self.formula_render_work = observed,
            Err(error) => {
                // Work already performed by a rejected candidate cannot be
                // reclaimed. Saturating the counter makes repeated hostile
                // candidates fail before receiving a fresh allowance.
                self.formula_render_work = self.max_formula_render_work;
                return Err(error);
            },
        }
        Ok(())
    }

    fn commit_attempt(&mut self, candidate: Self, published: bool) {
        if published {
            *self = candidate;
        } else {
            // Retained cells and text are transactional, but CPU work is a
            // package-wide admission cost even when the candidate is rejected.
            self.payload_work = self.payload_work.max(candidate.payload_work);
            self.formula_wire_bytes = self.formula_wire_bytes.max(candidate.formula_wire_bytes);
            self.formula_render_work = self.formula_render_work.max(candidate.formula_render_work);
        }
    }

    fn check_formula_render_depth(&self, depth: usize) -> Result<()> {
        if depth > self.max_formula_render_depth {
            return Err(Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaRenderDepth,
                observed: depth,
                maximum: self.max_formula_render_depth,
                path: SemanticPath::StructuredTables,
            });
        }
        Ok(())
    }
}

trait ListValueConverter<T> {
    fn convert(
        &mut self,
        entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
        budget: &mut ProjectionBudget,
    ) -> Result<T>;
}

impl<T, F> ListValueConverter<T> for F
where
    F: for<'source> FnMut(
        numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'source>,
        &mut ProjectionBudget,
    ) -> Result<T>,
{
    fn convert(
        &mut self,
        entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
        budget: &mut ProjectionBudget,
    ) -> Result<T> {
        self(entry, budget)
    }
}

/// Stage only final semantic values while the strict list codec walks the
/// source.  The visitor never returns a conversion/allocation error to the
/// codec: callbacks can run before a later wire or Buffa parity failure, so a
/// candidate-local semantic error is retained and the complete source is
/// still traversed.
struct TypedListVisitor<'converter, T, C> {
    converter: &'converter mut C,
    projection_budget: &'converter mut ProjectionBudget,
    stage_semantics: bool,
    expected_list_type: i32,
    values: Vec<(u32, T)>,
    keys: HashSet<u32>,
    segment_ids: Vec<u64>,
    segment_id_set: HashSet<u64>,
    segment: bool,
    segment_key_min: Option<u32>,
    segment_key_max: Option<u32>,
    structural_error: Option<Error>,
    semantic_error: Option<Error>,
}

/// Stage typed reply identities while the strict comment-storage codec walks
/// the source.  Reply callbacks can run before a later wire or Buffa parity
/// failure, so the staged prefix is only published after the complete payload
/// has decoded successfully.
struct CommentReplyVisitor {
    root_storage_id: u64,
    reply_ids: Vec<u64>,
    seen_reply_ids: HashSet<u64>,
    semantic_error: Option<Error>,
}

impl CommentReplyVisitor {
    fn new(root_storage_id: u64) -> Self {
        Self {
            root_storage_id,
            reply_ids: Vec::new(),
            seen_reply_ids: HashSet::new(),
            semantic_error: None,
        }
    }

    fn record_semantic_error(&mut self, error: Error) {
        if self.semantic_error.is_none() {
            self.semantic_error = Some(error);
        }
    }

    fn take_parts(self) -> (Vec<u64>, Option<Error>) {
        (self.reply_ids, self.semantic_error)
    }
}

impl comment_storage_codec::CommentStorageVisitor for CommentReplyVisitor {
    fn visit_reply(
        &mut self,
        reply: comment_storage_codec::ReferenceRecord<'_>,
    ) -> std::result::Result<(), comment_storage_codec::DecodeError> {
        // Keep validating the enclosing wire after a candidate-local semantic
        // failure, allowing a later wire error to retain precedence.
        if self.semantic_error.is_some() {
            return Ok(());
        }
        let identifier = reply.identifier();
        // A direct self-reference or duplicate reply would make the native
        // thread ambiguous. Keep this check alongside collection so malformed
        // graph edges cannot escape through the typed sidecar.
        if identifier == self.root_storage_id || self.seen_reply_ids.contains(&identifier) {
            self.record_semantic_error(Error::MalformedPayload {
                path: SemanticPath::StructuredTables,
            });
            return Ok(());
        }
        if self.reply_ids.try_reserve(1).is_err() {
            self.record_semantic_error(allocation_error(
                "Numbers comment replies",
                self.reply_ids.len().saturating_add(1),
            ));
            return Ok(());
        }
        if self.seen_reply_ids.try_reserve(1).is_err() {
            self.record_semantic_error(allocation_error(
                "Numbers comment reply identities",
                self.seen_reply_ids.len().saturating_add(1),
            ));
            return Ok(());
        }
        self.seen_reply_ids.insert(identifier);
        self.reply_ids.push(identifier);
        Ok(())
    }
}

impl<'converter, T, C> TypedListVisitor<'converter, T, C> {
    fn new(
        converter: &'converter mut C,
        projection_budget: &'converter mut ProjectionBudget,
        expected_list_type: i32,
        segment: bool,
        stage_semantics: bool,
    ) -> Self {
        Self {
            converter,
            projection_budget,
            stage_semantics,
            expected_list_type,
            values: Vec::new(),
            keys: HashSet::new(),
            segment_ids: Vec::new(),
            segment_id_set: HashSet::new(),
            segment,
            segment_key_min: None,
            segment_key_max: None,
            structural_error: None,
            semantic_error: None,
        }
    }

    fn record_structural_error(&mut self, error: Error) {
        if self.structural_error.is_none() {
            self.structural_error = Some(error);
        }
    }

    fn record_semantic_error(&mut self, error: Error) {
        if self.semantic_error.is_none() {
            self.semantic_error = Some(error);
        }
    }

    fn take_parts(
        self,
    ) -> (
        Vec<(u32, T)>,
        HashSet<u32>,
        Vec<u64>,
        Option<Error>,
        Option<Error>,
    ) {
        (
            self.values,
            self.keys,
            self.segment_ids,
            self.structural_error,
            self.semantic_error,
        )
    }

    fn take_segment_bounds(&self) -> (Option<u32>, Option<u32>) {
        (self.segment_key_min, self.segment_key_max)
    }
}

impl<T, C> numbers_table_cell_storage_codec::StorageVisitor for TypedListVisitor<'_, T, C>
where
    C: ListValueConverter<T>,
{
    fn visit_list_entry(
        &mut self,
        entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
    ) -> std::result::Result<(), numbers_table_cell_storage_codec::DecodeError> {
        if self.segment {
            match &mut self.segment_key_min {
                Some(minimum) => *minimum = (*minimum).min(entry.key()),
                minimum @ None => *minimum = Some(entry.key()),
            }
            match &mut self.segment_key_max {
                Some(maximum) => *maximum = (*maximum).max(entry.key()),
                maximum @ None => *maximum = Some(entry.key()),
            }
        }
        if self.structural_error.is_some() {
            return Ok(());
        }
        if !entry_matches_list_type_snapshot(&entry, self.expected_list_type) {
            self.record_structural_error(Error::InvalidFormat(
                "Numbers table-data-list entry has no selected payload".to_owned(),
            ));
            return Ok(());
        }
        if self.keys.contains(&entry.key()) {
            self.record_structural_error(Error::InvalidFormat(
                "Numbers table sidecar contains duplicate keys".to_owned(),
            ));
            return Ok(());
        }
        if self.keys.try_reserve(1).is_err() {
            // Keep walking after a fallible staging failure. The error is
            // candidate-local and must not prevent a later structural error
            // from winning publication.
            self.record_semantic_error(allocation_error(
                "Numbers table-list entry keys",
                self.keys.len().saturating_add(1),
            ));
            return Ok(());
        }
        self.keys.insert(entry.key());
        // Once a semantic conversion/allocation error has been retained, keep
        // checking the wire-level shape and duplicate-key invariants above,
        // but do not invoke the converter or allocate another retained value.
        // Non-admitted candidates use the same structural path while never
        // staging semantic values or charging references/text.
        if self.semantic_error.is_some() || !self.stage_semantics {
            return Ok(());
        }
        if self.values.try_reserve(1).is_err() {
            self.record_semantic_error(allocation_error(
                "Numbers table-list entries",
                self.values.len().saturating_add(1),
            ));
            return Ok(());
        }
        match self.converter.convert(entry, self.projection_budget) {
            Ok(value) => self.values.push((entry.key(), value)),
            Err(error) => self.record_semantic_error(error),
        }
        Ok(())
    }

    fn visit_list_segment(
        &mut self,
        reference: numbers_table_cell_storage_codec::ReferenceRecord<'_>,
    ) -> std::result::Result<(), numbers_table_cell_storage_codec::DecodeError> {
        if !self.stage_semantics {
            return Ok(());
        }
        let identifier = reference.reference().identifier();
        if self.segment_id_set.contains(&identifier) {
            self.record_structural_error(Error::InvalidFormat(format!(
                "Numbers table-data-list repeats segment object {identifier}"
            )));
            return Ok(());
        }
        if self.segment_id_set.try_reserve(1).is_err() {
            self.record_semantic_error(allocation_error(
                "Numbers table-list segment identities",
                self.segment_id_set.len().saturating_add(1),
            ));
            return Ok(());
        }
        if self.segment_ids.try_reserve(1).is_err() {
            self.record_semantic_error(allocation_error(
                "Numbers table-list segment identities",
                self.segment_ids.len().saturating_add(1),
            ));
            return Ok(());
        }
        self.segment_id_set.insert(identifier);
        self.segment_ids.push(identifier);
        Ok(())
    }
}

struct TileRowVisitor<'a, 'tables> {
    row_origin: usize,
    tile_size: usize,
    row_count: usize,
    column_count: usize,
    budget: &'a mut CellBudget,
    cell_tables: &'a CellTables<'tables>,
    projection_budget: &'a mut ProjectionBudget,
    table: &'a mut Table,
    materialized_cells: usize,
    semantic_error: Option<Error>,
}

impl<'a, 'tables> numbers_table_cell_storage_codec::StorageVisitor for TileRowVisitor<'a, 'tables> {
    fn visit_tile_row(
        &mut self,
        row: numbers_table_cell_storage_codec::TileRowInfoSnapshot<'_>,
    ) -> std::result::Result<(), numbers_table_cell_storage_codec::DecodeError> {
        self.materialized_cells = self
            .materialized_cells
            .saturating_add(usize::try_from(row.cell_count()).unwrap_or(usize::MAX));
        if self.semantic_error.is_some() {
            return Ok(());
        }
        if let Err(error) = self
            .projection_budget
            .check_materialized_cells(self.materialized_cells)
        {
            self.semantic_error = Some(error);
            return Ok(());
        }
        if let Err(error) = TableDataExtractor::parse_tile_row(
            &row,
            self.row_origin,
            self.tile_size,
            self.row_count,
            self.column_count,
            self.budget,
            self.cell_tables,
            self.projection_budget,
            self.table,
        ) {
            self.semantic_error = Some(error);
        }
        Ok(())
    }
}

/// Stage tile-storage references until the complete strict envelope has
/// decoded successfully. A later wire error must not leave parsed cells in a
/// candidate table, so this visitor never resolves or publishes a tile.
struct TileStorageReferenceStage {
    references: Vec<(u32, u64)>,
    allocation_error: Option<Error>,
}

impl TileStorageReferenceStage {
    fn new() -> Self {
        Self {
            references: Vec::new(),
            allocation_error: None,
        }
    }

    fn take(self) -> Result<Vec<(u32, u64)>> {
        if let Some(error) = self.allocation_error {
            return Err(error);
        }
        Ok(self.references)
    }
}

impl numbers_table_cell_storage_codec::StorageVisitor for TileStorageReferenceStage {
    fn visit_tile_reference(
        &mut self,
        record: numbers_table_cell_storage_codec::TileReferenceRecord<'_>,
    ) -> std::result::Result<(), numbers_table_cell_storage_codec::DecodeError> {
        if self.allocation_error.is_some() {
            return Ok(());
        }
        if self.references.try_reserve(1).is_err() {
            self.allocation_error = Some(allocation_error(
                "Numbers tile-storage references",
                self.references.len().saturating_add(1),
            ));
            return Ok(());
        }
        self.references
            .push((record.tile_id(), record.reference().identifier()));
        Ok(())
    }
}

fn projection_charge(
    current: usize,
    amount: usize,
    maximum: usize,
    kind: SemanticLimitKind,
) -> Result<usize> {
    let observed = current.checked_add(amount).ok_or(Error::SemanticLimit {
        kind,
        observed: usize::MAX,
        maximum,
        path: SemanticPath::StructuredTables,
    })?;
    if observed > maximum {
        return Err(Error::SemanticLimit {
            kind,
            observed,
            maximum,
            path: SemanticPath::StructuredTables,
        });
    }
    Ok(observed)
}

fn allocation_error(resource: &'static str, amount: usize) -> Error {
    Error::Common(litchi_iwa_common::Error::Allocation { resource, amount })
}

fn table_limit_error(observed: usize, maximum: usize) -> Error {
    Error::SemanticLimit {
        kind: SemanticLimitKind::Tables,
        observed,
        maximum,
        path: SemanticPath::StructuredTables,
    }
}

fn checked_next_table_count(current: usize, maximum: usize) -> Result<usize> {
    let observed = current.checked_add(1).ok_or_else(|| {
        Error::InvalidFormat("Numbers table count overflows host usize".to_owned())
    })?;
    if observed > maximum {
        return Err(table_limit_error(observed, maximum));
    }
    Ok(observed)
}

fn decode_projected_legacy_candidate<T>(
    data: &[u8],
    admit: impl FnOnce() -> Result<()>,
    parse: impl FnOnce(&[u8]) -> Result<T>,
) -> Result<Option<T>> {
    match has_legacy_table_model_wire_shape(data) {
        Ok(true) => {},
        Ok(false) => return Ok(None),
        Err(error) => return Err(error),
    }
    admit()?;
    parse(data).map(Some).map_err(|error| match error {
        Error::InvalidFormat(_) => Error::MalformedPayload {
            path: SemanticPath::StructuredTables,
        },
        other => other,
    })
}

/// Require the parse-relevant required fields that distinguish a historical
/// type-6000 `TableModelArchive` from its modern `TableInfoArchive` owner.
///
/// Prost supplies defaults for absent proto2 required fields, so a successful
/// generated decode alone is not evidence that the payload is a table model.
/// In particular, protobuf's permissive unknown-field handling lets ordinary
/// table-info metadata decode into a mostly default model. Scanning only the
/// flat envelope retains the compatibility fallback without allocating a
/// second field index. Once this shape is present, the candidate is admitted
/// and every model-decoding failure must remain observable to the caller.
fn has_legacy_table_model_wire_shape(data: &[u8]) -> Result<bool> {
    const DATA_STORE: u8 = 1 << 0;
    const ROW_COUNT: u8 = 1 << 1;
    const COLUMN_COUNT: u8 = 1 << 2;
    const TABLE_NAME: u8 = 1 << 3;
    const REQUIRED: u8 = DATA_STORE | ROW_COUNT | COLUMN_COUNT | TABLE_NAME;

    let mut present = 0_u8;
    let mut ambiguous = false;
    let preflight = preflight_wire_tree_with_limits(data, WireLimits::default(), |visit| {
        let field = visit.field();
        let required = match field.number() {
            4 => Some((DATA_STORE, 2)),
            6 => Some((ROW_COUNT, 0)),
            7 => Some((COLUMN_COUNT, 0)),
            8 => Some((TABLE_NAME, 2)),
            _ => None,
        };
        if let Some((bit, expected_wire_type)) = required {
            if field.wire_type() != expected_wire_type || present & bit != 0 {
                ambiguous = true;
            } else {
                present |= bit;
            }
        }
        Ok(WireDescent::Skip)
    });
    if let Err(error) = preflight {
        return match error {
            litchi_iwa_common::Error::InvalidFormat(_) if present == REQUIRED => {
                Err(Error::MalformedPayload {
                    path: SemanticPath::StructuredTables,
                })
            },
            litchi_iwa_common::Error::InvalidFormat(_) => Ok(false),
            common_error @ (litchi_iwa_common::Error::LimitExceeded { .. }
            | litchi_iwa_common::Error::Allocation { .. }
            | litchi_iwa_common::Error::InvalidLimit { .. }) => Err(Error::Common(common_error)),
        };
    }
    if present != REQUIRED {
        return Ok(false);
    }
    if ambiguous {
        return Err(Error::MalformedPayload {
            path: SemanticPath::StructuredTables,
        });
    }
    Ok(true)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TableModelCompatibilityShape {
    Sparse,
    DenseWithGeneratedDefaultStyles,
}

fn table_model_compatibility_shape(data: &[u8]) -> Option<TableModelCompatibilityShape> {
    // `prost::Message::encode_to_vec` emits proto2 required message fields
    // even when their generated value is the empty/default message. A few
    // historical and archive-free Numbers producers therefore serialize all
    // native style references as `Reference { identifier: 0 }` while still
    // carrying the selected DataStore routes. Those payloads are the
    // compatibility shape even though the DataStore itself is populated;
    // strict reference projection quite correctly rejects identifier zero.
    const DEFAULT_STYLE_REFERENCE_FIELDS: [(u32, u16); 9] = [
        (3, 1 << 0),
        (18, 1 << 1),
        (19, 1 << 2),
        (20, 1 << 3),
        (21, 1 << 4),
        (24, 1 << 5),
        (25, 1 << 6),
        (26, 1 << 7),
        (27, 1 << 8),
    ];
    const ALL_DEFAULT_STYLE_REFERENCES: u16 = (1 << DEFAULT_STYLE_REFERENCE_FIELDS.len()) - 1;
    let mut table_id = false;
    let mut base_data_store = false;
    let mut dense_data_store = false;
    let mut number_of_rows = false;
    let mut number_of_columns = false;
    let mut table_name = false;
    let mut default_style_references = 0_u16;
    let mut invalid_default_style_shape = false;
    let mut invalid_shape = false;
    let root_scan = preflight_wire_tree_with_limits(data, WireLimits::default(), |visit| {
        // A sparse compatibility model omits the native DataStore metadata
        // envelopes. Complete native models carry row headers, column
        // headers, and style metadata (fields 1, 2, and 5), and must continue
        // through the strict storage codec. Descend only into the selected
        // base-data-store envelope so this admission gate can distinguish the
        // two shapes without decoding or indexing it.
        if visit.path() == [4] {
            let field = visit.field();
            // Prost emits an empty required message as the canonical
            // default scalar (`08 00`) in the synthetic/legacy sparse
            // envelopes. Treat that representation (and a truly omitted
            // payload) as sparse; any other metadata bytes make this a
            // complete native store and keep it on the strict route.
            let default_metadata =
                field.wire_type() == 2 && matches!(field.payload(), [] | [0x08, 0x00]);
            if matches!(field.number(), 1 | 2 | 5) && !default_metadata {
                dense_data_store = true;
            }
            return Ok(WireDescent::Skip);
        }
        if !visit.path().is_empty() {
            return Ok(WireDescent::Skip);
        }
        let field = visit.field();
        if let Some((_, bit)) = DEFAULT_STYLE_REFERENCE_FIELDS
            .iter()
            .find(|(number, _)| *number == field.number())
        {
            // Treat only the exact canonical empty Reference as the
            // generated proto2 default. Duplicates, wrong wire types,
            // malformed framing, and non-zero references stay on the
            // strict/native admission path.
            if field.wire_type() != 2
                || default_style_references & *bit != 0
                || field.validate_canonical_framing().is_err()
                || field.payload() != [0x08, 0x00]
            {
                // Style references are outside the historical sparse-shape
                // admission gate, so a real reference must not invalidate
                // that pre-existing route. It only disqualifies the new,
                // narrower generated-default-style compatibility shape.
                invalid_default_style_shape = true;
            } else {
                default_style_references |= *bit;
            }
            return Ok(WireDescent::Skip);
        }
        match field.number() {
            1 => {
                // The compatibility extractor does not consume table_id, but
                // a present value still has to retain the strict scalar
                // framing/UTF-8 contract before sparse fallback is admitted.
                if field.wire_type() != 2
                    || table_id
                    || field.validate_canonical_framing().is_err()
                    || std::str::from_utf8(field.payload()).is_err()
                {
                    invalid_shape = true;
                } else {
                    table_id = true;
                }
            },
            4 => {
                if field.wire_type() != 2
                    || base_data_store
                    || field.validate_canonical_framing().is_err()
                {
                    invalid_shape = true;
                } else {
                    base_data_store = true;
                    return Ok(WireDescent::Descend);
                }
            },
            6 => {
                let payload = field.payload();
                let canonical = litchi_iwa_common::decode_varint_from_bytes(payload).is_ok_and(
                    |(value, width)| {
                        width == payload.len()
                            && width == litchi_iwa_common::varint::encoded_len(value)
                            && u32::try_from(value).is_ok()
                    },
                );
                if field.wire_type() != 0
                    || number_of_rows
                    || field.validate_canonical_framing().is_err()
                    || !canonical
                {
                    invalid_shape = true;
                } else {
                    number_of_rows = true;
                }
            },
            7 => {
                let payload = field.payload();
                let canonical = litchi_iwa_common::decode_varint_from_bytes(payload).is_ok_and(
                    |(value, width)| {
                        width == payload.len()
                            && width == litchi_iwa_common::varint::encoded_len(value)
                            && u32::try_from(value).is_ok()
                    },
                );
                if field.wire_type() != 0
                    || number_of_columns
                    || field.validate_canonical_framing().is_err()
                    || !canonical
                {
                    invalid_shape = true;
                } else {
                    number_of_columns = true;
                }
            },
            8 => {
                // The archive-free compatibility projection mirrors
                // generated protobuf behavior for a repeated singular name:
                // retain the last valid value.  Keep that historical
                // compatibility behavior for sparse model ingress; the names
                // transaction performs its own strict table-name projection
                // before any mutation is published.
                if field.wire_type() != 2
                    || field.validate_canonical_framing().is_err()
                    || std::str::from_utf8(field.payload()).is_err()
                {
                    invalid_shape = true;
                } else {
                    table_name = true;
                }
            },
            _ => {},
        }
        Ok(WireDescent::Skip)
    });
    // Sparse compatibility defaults for omitted dimensions would otherwise
    // manufacture a valid-looking 0x0 table from a merely truncated model.
    // An explicit 0x0 pair remains valid and retains historical compatibility.
    if root_scan.is_err()
        || !base_data_store
        || !number_of_rows
        || !number_of_columns
        || !table_name
        || invalid_shape
    {
        return None;
    }
    if !dense_data_store {
        Some(TableModelCompatibilityShape::Sparse)
    } else if !invalid_default_style_shape
        && default_style_references == ALL_DEFAULT_STYLE_REFERENCES
    {
        Some(TableModelCompatibilityShape::DenseWithGeneratedDefaultStyles)
    } else {
        None
    }
}

fn sparse_table_model_compatibility_shape(data: &[u8]) -> bool {
    table_model_compatibility_shape(data) == Some(TableModelCompatibilityShape::Sparse)
}

fn compact_table<T>(entries: impl IntoIterator<Item = (u32, T)>) -> Result<CompactTable<T>> {
    let mut compacted = Vec::new();
    let entries = entries.into_iter();
    let (lower_bound, _) = entries.size_hint();
    compacted
        .try_reserve(lower_bound)
        .map_err(|_| allocation_error("Numbers table sidecar entries", lower_bound))?;
    for entry in entries {
        compacted
            .try_reserve(1)
            .map_err(|_| allocation_error("Numbers table sidecar entries", compacted.len() + 1))?;
        compacted.push(entry);
    }
    compact_table_vec(compacted)
}

fn compact_table_vec<T>(mut compacted: Vec<(u32, T)>) -> Result<CompactTable<T>> {
    compacted.sort_unstable_by_key(|(key, _)| *key);
    if compacted.windows(2).any(|pair| pair[0].0 == pair[1].0) {
        return Err(Error::InvalidFormat(
            "Numbers table sidecar contains duplicate keys".to_owned(),
        ));
    }
    Ok(compacted.into_boxed_slice())
}

fn compact_table_get<T>(table: &[(u32, T)], key: u32) -> Option<&T> {
    table
        .binary_search_by_key(&key, |(entry_key, _)| *entry_key)
        .ok()
        .map(|index| &table[index].1)
}

fn nonzero_reference_id(identifier: u64) -> Option<u64> {
    (identifier != 0).then_some(identifier)
}

fn retain_text(value: &str, budget: &mut ProjectionBudget) -> Result<String> {
    budget.charge_output_text(value.len())?;
    let mut retained = String::new();
    retained
        .try_reserve_exact(value.len())
        .map_err(|_| allocation_error("Numbers retained semantic text", value.len()))?;
    retained.push_str(value);
    Ok(retained)
}

/// Clone one comment without relying on infallible collection allocation.
///
/// Comment values are stored in a sidecar and copied into every referencing
/// cell.  Keep this boundary fallible even though the source comment itself
/// has already passed semantic limits: a hostile package can reference one
/// large comment from many cells and otherwise turn a derived `Clone` into an
/// unbounded allocation path.
fn try_clone_comment(comment: &Comment) -> Result<Comment> {
    let mut text = String::new();
    text.try_reserve_exact(comment.text.len())
        .map_err(|_| allocation_error("Numbers materialized comment text", comment.text.len()))?;
    text.push_str(&comment.text);

    let mut reply_ids = Vec::new();
    reply_ids
        .try_reserve_exact(comment.reply_ids.len())
        .map_err(|_| {
            allocation_error(
                "Numbers materialized comment replies",
                comment.reply_ids.len(),
            )
        })?;
    reply_ids.extend(comment.reply_ids.iter().copied());

    Ok(Comment {
        text,
        creation_date_seconds: comment.creation_date_seconds,
        author_id: comment.author_id,
        reply_ids: reply_ids.into_boxed_slice(),
        storage_uuid: comment.storage_uuid,
    })
}

fn retained_table_text(
    table: &[(u32, String)],
    identifier: u32,
    budget: &mut ProjectionBudget,
) -> Result<Option<String>> {
    compact_table_get(table, identifier)
        .map(|value| retain_text(value, budget))
        .transpose()
}

fn checked_table_dimensions(row_count: u32, column_count: u32) -> Result<(usize, usize)> {
    let row_count = usize::try_from(row_count).map_err(|_| {
        Error::InvalidFormat("Numbers table row count does not fit the host usize".to_owned())
    })?;
    let column_count = usize::try_from(column_count).map_err(|_| {
        Error::InvalidFormat("Numbers table column count does not fit the host usize".to_owned())
    })?;

    if row_count > MAX_TABLE_ROWS {
        return Err(Error::Common(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::TableRows,
            observed: row_count,
            limit: MAX_TABLE_ROWS,
        }));
    }
    if column_count > MAX_TABLE_COLUMNS {
        return Err(Error::Common(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::TableColumns,
            observed: column_count,
            limit: MAX_TABLE_COLUMNS,
        }));
    }

    let addressable_cells = row_count.checked_mul(column_count).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table dimensions overflow host address space: {row_count}x{column_count}"
        ))
    })?;
    if addressable_cells > MAX_ADDRESSABLE_CELLS {
        return Err(Error::Common(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::TableCells,
            observed: addressable_cells,
            limit: MAX_ADDRESSABLE_CELLS,
        }));
    }

    Ok((row_count, column_count))
}

fn validate_table_row(row: usize, row_count: usize) -> Result<()> {
    if row >= row_count {
        return Err(Error::InvalidFormat(format!(
            "Numbers tile row {row} is outside the declared table height {row_count}"
        )));
    }
    Ok(())
}

fn validate_table_column(column: usize, column_count: usize) -> Result<()> {
    if column >= column_count {
        return Err(Error::InvalidFormat(format!(
            "Numbers cell column {column} is outside the declared table width {column_count}"
        )));
    }
    Ok(())
}

fn record_first_list_error(slot: &mut Option<Error>, error: Error) {
    if slot.is_none() {
        *slot = Some(error);
    }
}

fn entry_matches_list_type_snapshot(
    entry: &numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
    list_type: i32,
) -> bool {
    let present = [
        entry.string_value().is_some(),
        entry.reference().is_some(),
        entry.formula().is_some(),
        entry.format().is_some(),
        entry.custom_format().is_some(),
        entry.rich_text_payload().is_some(),
        entry.comment_storage().is_some(),
        entry.import_warning_set().is_some(),
        entry.cell_spec().is_some(),
    ];
    if present.into_iter().filter(|value| *value).count() != 1 {
        return false;
    }
    match list_type {
        value
            if value == tst::table_data_list::ListType::String as i32
                || value == tst::table_data_list::ListType::FormulaError as i32 =>
        {
            entry.string_value().is_some()
        },
        value if value == tst::table_data_list::ListType::Formula as i32 => {
            entry.formula().is_some()
        },
        value if value == tst::table_data_list::ListType::RichTextPayload as i32 => {
            entry.rich_text_payload().is_some()
        },
        value if value == tst::table_data_list::ListType::CommentStorage as i32 => {
            entry.comment_storage().is_some()
        },
        _ => true,
    }
}

#[derive(Debug, Clone)]
struct FormulaReferenceName {
    sheet: Arc<String>,
    table: Arc<String>,
}

#[derive(Debug, Clone, Default)]
struct FormulaReferenceMaps {
    owners: HashMap<FormulaOwnerKey, FormulaReferenceName>,
    categories: HashMap<FormulaCategoryKey, String>,
}

/// Borrowed table-model values needed by semantic extraction.
///
/// The strict storage codec validates the complete selected model/datastore
/// envelope but intentionally does not retain a generated archive.  The
/// table name is supplied by the separate strict name projection, while the
/// datastore snapshot supplies the sidecar references and tile-storage bytes
/// consumed below.
#[derive(Clone, Copy)]
struct ProjectedTableModel<'source> {
    table_name: &'source str,
    number_of_rows: u32,
    number_of_columns: u32,
    compatibility_defaults: bool,
    compatibility_data_store: bool,
    base_data_store: numbers_table_cell_storage_codec::DataStoreSnapshot<'source>,
}

/// Extractor for Numbers table data
pub(super) struct TableDataExtractor<'a> {
    bundle: &'a Components,
    object_index: &'a Index,
    projection_budget: RefCell<ProjectionBudget>,
    max_tables: usize,
    retain_comments: bool,
    document_projection: bool,
}

impl<'a> TableDataExtractor<'a> {
    /// Return whether the index contains a candidate table-model object.
    ///
    /// This cheap type probe lets generic structured extraction avoid building
    /// formula-reference sidecars for Pages and Keynote packages.
    pub(super) fn has_table_models(object_index: &Index) -> bool {
        [TABLE_MODEL_MESSAGE_TYPE, 6_000]
            .into_iter()
            .any(|message_type| {
                object_index
                    .iter_entries_by_type(message_type)
                    .next()
                    .is_some()
            })
    }

    /// Create a new table data extractor
    pub(super) fn new(
        bundle: &'a Components,
        object_index: &'a Index,
        limits: SemanticLimits,
    ) -> Self {
        Self {
            bundle,
            object_index,
            projection_budget: RefCell::new(ProjectionBudget::new(limits)),
            max_tables: limits.max_tables(),
            retain_comments: true,
            document_projection: false,
        }
    }

    pub(super) fn without_comments(mut self) -> Self {
        self.retain_comments = false;
        self.document_projection = true;
        self
    }

    pub(super) fn native_projection(mut self) -> Self {
        self.document_projection = true;
        self
    }

    /// Charge semantic text retained outside table projection, such as rooted
    /// sheet names, against the same package-wide output budget.
    pub(super) fn charge_output_text(&self, amount: usize) -> Result<()> {
        self.projection_budget
            .borrow_mut()
            .charge_output_text(amount)
    }

    /// Merge a rooted reference admission into the table-sidecar budget.
    pub(super) fn charge_references(&self, amount: usize) -> Result<()> {
        self.projection_budget
            .borrow_mut()
            .charge_references(amount)
    }

    /// Extract all tables from the document
    pub(super) fn extract_all_tables(&self) -> Result<Vec<Table>> {
        let max_tables = self.max_tables;
        let mut tables = Vec::new();
        self.for_each_table(max_tables, |table| {
            let next_count = checked_next_table_count(tables.len(), max_tables)?;
            tables
                .try_reserve(1)
                .map_err(|_| allocation_error("Numbers extracted table results", next_count))?;
            tables.push(table);
            Ok(())
        })?;
        Ok(tables)
    }

    /// Extract all tables directly into the canonical Numbers semantic model.
    ///
    /// The archive adapter's builder is consumed one table at a time. Its
    /// sparse cell and header buffers move into the leaf table, so this path
    /// avoids first allocating a `Vec<Table>` only to convert every
    /// element into a second result vector for structured extraction.
    pub(super) fn extract_all_semantic_tables(
        &self,
        max_tables: usize,
    ) -> Result<Vec<crate::Table>> {
        let max_tables = max_tables.min(self.max_tables);
        let mut tables = Vec::new();
        self.for_each_table(max_tables, |table| {
            let next_count = checked_next_table_count(tables.len(), max_tables)?;
            tables
                .try_reserve(1)
                .map_err(|_| allocation_error("Numbers semantic table results", next_count))?;
            tables.push(table.into_semantic_table()?);
            Ok(())
        })?;
        Ok(tables)
    }

    fn for_each_table(
        &self,
        max_tables: usize,
        mut visit: impl FnMut(Table) -> Result<()>,
    ) -> Result<()> {
        let max_tables = max_tables.min(self.max_tables);
        let mut seen_objects = HashSet::new();
        let mut table_count = 0usize;

        // Real packages index TableModelArchive as 6001. Older generated
        // fixtures may store the same payload under 6000, so the object
        // adapter accepts 6000 only when its payload passes model extraction;
        // a genuine TableInfoArchive is ignored rather than mis-decoded.
        for message_type in [TABLE_MODEL_MESSAGE_TYPE, 6_000] {
            for entry in self.object_index.iter_entries_by_type(message_type) {
                if seen_objects.contains(&entry.id()) {
                    continue;
                }
                let next_seen = seen_objects.len().checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat(
                        "Numbers structured table identity count overflows host usize".to_owned(),
                    )
                })?;
                seen_objects.try_reserve(1).map_err(|_error| {
                    allocation_error("Numbers structured table identities", next_seen)
                })?;
                seen_objects.insert(entry.id());
                // Candidate admission is deliberately checked before protobuf
                // decoding. Once the caller-selected table budget is full, a
                // later malformed canonical candidate cannot force another
                // potentially large model allocation merely to choose an error.
                if message_type == TABLE_MODEL_MESSAGE_TYPE && table_count >= max_tables {
                    checked_next_table_count(table_count, max_tables)?;
                }
                if let Some(resolved) = self.object_index.resolve_ref(self.bundle, entry.id())?
                    && let Some(table) =
                        self.extract_table_candidate(&resolved, message_type, || {
                            checked_next_table_count(table_count, max_tables).map(|_| ())
                        })?
                {
                    table_count = checked_next_table_count(table_count, max_tables)?;
                    visit(table)?;
                }
            }
        }
        Ok(())
    }

    /// Strictly project the model and its embedded datastore without
    /// constructing generated Prost values.
    ///
    /// The model codec already walks the datastore once while validating the
    /// model envelope.  The second datastore pass is deliberately retained
    /// only to recover the borrowed sidecar/tile routes needed by extraction;
    /// its references are not charged twice, while its actual field/work
    /// inspection remains part of the aggregate projection budget.
    fn project_table_model<'source>(
        &self,
        source: &'source [u8],
        budget: &mut ProjectionBudget,
        path: SemanticPath,
    ) -> Result<ProjectedTableModel<'source>> {
        let mut compatibility_defaults = !self.document_projection;
        let mut compatibility_data_store = compatibility_defaults;
        let options = table_cell_decode_options(
            source,
            budget.remaining_references(),
            MAX_FORMULA_WIRE_BYTES,
            budget.remaining_payload_fields(),
            budget.remaining_payload_work(),
        );
        let decoded = if compatibility_defaults {
            numbers_table_cell_storage_codec::decode_table_model_compatibility_with_report(
                source, options,
            )
        } else {
            match numbers_table_cell_storage_codec::decode_table_model_with_report(source, options)
            {
                Ok(decoded) => Ok(decoded),
                Err(error) if error.resource_limit().is_some() => Err(error),
                Err(error) => {
                    // Classification is deliberately deferred until the
                    // strict route has failed. It is an admission probe, not
                    // an unconditional walk of every native model.
                    if let Some(shape) = table_model_compatibility_shape(source) {
                        compatibility_defaults = true;
                        // Generated proto2 default style references affect the
                        // root projection only. A populated native DataStore
                        // must retain strict selected-sidecar/tile validation.
                        compatibility_data_store = shape == TableModelCompatibilityShape::Sparse;
                        let compatibility_options = table_cell_decode_options(
                            source,
                            usize::MAX,
                            MAX_FORMULA_WIRE_BYTES,
                            budget.remaining_payload_fields(),
                            budget.remaining_payload_work(),
                        );
                        numbers_table_cell_storage_codec::decode_table_model_compatibility_with_report(
                            source,
                            compatibility_options,
                        )
                    } else {
                        // A dense native model remains strict at the root, but
                        // malformed unselected DataStore metadata (for example,
                        // row-header ownership) must not make package ingress
                        // fail. Retry only the nested compatibility projection;
                        // selected sidecars and tile storage remain strict below.
                        match numbers_table_cell_storage_codec::decode_table_model_with_compatibility_data_store_with_report(source, options)
                        {
                            Ok(decoded) => {
                                compatibility_data_store = true;
                                Ok(decoded)
                            },
                            Err(_compatibility_error) => Err(error),
                        }
                    }
                },
            }
        };
        let (model, report) = decoded.map_err(|error| {
            map_table_cell_codec_error_with_offsets(
                error,
                budget.references,
                budget.payload_fields,
                budget.payload_work,
                0,
            )
        })?;
        if compatibility_defaults {
            // Generated compatibility defaults do not represent newly
            // retained references. The sidecar/list/tile routes charge the
            // references they actually resolve below; retain this envelope
            // pass's bounded wire/work cost without charging its borrowed
            // default/reference fields a second time.
            budget.charge_decode_work(report)?;
        } else {
            budget.charge_decode_report(report)?;
        }

        // Decode the display name only after the model envelope has passed its
        // wire gate. Compatibility models use generated defaults for omitted
        // fields, while native models retain the independent strict names
        // projection.
        let table_name = if compatibility_defaults {
            model.table_name()
        } else {
            names::preflight_table_name(source)
                .map_err(|error| super::map_sheet_preflight_error(error, path, 0))?
        };
        budget.charge_output_text(table_name.len())?;

        let data_store_source = model.base_data_store();
        let data_store_options = table_cell_decode_options(
            data_store_source,
            usize::MAX,
            usize::MAX,
            budget.remaining_payload_fields(),
            budget.remaining_payload_work(),
        );
        let data_store_decoded = if compatibility_data_store {
            numbers_table_cell_storage_codec::decode_data_store_compatibility_with_report(
                data_store_source,
                data_store_options,
            )
        } else {
            numbers_table_cell_storage_codec::decode_data_store_with_report(
                data_store_source,
                data_store_options,
            )
        };
        let (base_data_store, data_store_report) = data_store_decoded.map_err(|error| {
            map_table_cell_codec_error_with_offsets(
                error,
                budget.references,
                budget.payload_fields,
                budget.payload_work,
                0,
            )
        })?;
        // The model pass is authoritative for selected reference admission;
        // this recovery pass only adds the second scan's fields/work cost.
        budget.charge_decode_work(data_store_report)?;

        Ok(ProjectedTableModel {
            table_name,
            number_of_rows: model.number_of_rows(),
            number_of_columns: model.number_of_columns(),
            compatibility_defaults,
            compatibility_data_store,
            base_data_store,
        })
    }

    /// Project one raw model payload into a candidate-local semantic table.
    fn parse_projected_table_payload(&self, source: &[u8], path: SemanticPath) -> Result<Table> {
        let mut candidate_budget = *self.projection_budget.borrow();
        let projected = match self.project_table_model(source, &mut candidate_budget, path) {
            Ok(projected) => projected,
            Err(error) => {
                self.projection_budget
                    .borrow_mut()
                    .commit_attempt(candidate_budget, false);
                return Err(error);
            },
        };
        self.parse_projected_table_model(projected, true, Some(candidate_budget))
    }

    /// Extract a single table from a resolved object
    fn extract_table_candidate(
        &self,
        object: &Resolved<'_>,
        candidate_type: u32,
        legacy_admit: impl FnOnce() -> Result<()>,
    ) -> Result<Option<Table>> {
        if candidate_type == TABLE_MODEL_MESSAGE_TYPE {
            let mut messages = object
                .messages
                .iter()
                .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE);
            let Some(message) = messages.next() else {
                return Err(Error::InvalidFormat(
                    "Numbers canonical table candidate has no canonical payload".to_owned(),
                ));
            };
            if messages.next().is_some() {
                return Err(Error::InvalidFormat(
                    "Numbers canonical table candidate has duplicate canonical payloads".to_owned(),
                ));
            }
            return self
                .parse_projected_table_payload(
                    message.data.as_slice(),
                    SemanticPath::StructuredTables,
                )
                .map(Some);
        }

        // Protobuf is permissive, and legacy fixtures used 6000 for a model.
        // Decode only the primary candidate payload. A secondary canonical
        // payload must not promote an object classified as legacy metadata.
        let Some(message) = object
            .messages
            .iter()
            .next()
            .filter(|message| message.type_ == 6_000)
        else {
            return Err(Error::InvalidFormat(
                "Numbers legacy table candidate has no primary legacy payload".to_owned(),
            ));
        };
        if object
            .messages
            .iter()
            .skip(1)
            .any(|candidate| candidate.type_ == 6_000)
        {
            return Err(Error::InvalidFormat(
                "Numbers legacy table candidate has duplicate legacy payloads".to_owned(),
            ));
        }
        decode_projected_legacy_candidate(&message.data, legacy_admit, |source| {
            self.parse_projected_table_payload(source, SemanticPath::StructuredTables)
        })
    }

    /// Extract a table model reached through a schema-proven `TableInfo` edge.
    ///
    /// Rooted ownership is stricter than the global compatibility scan:
    /// canonical payloads are authoritative, while a native projection
    /// accepts one historical 6000 payload beside a canonical 6001 payload
    /// only when their bytes are identical. Conflicting mixed payloads and
    /// duplicate canonical payloads fail; compatibility mode preserves its
    /// canonical-first treatment of mixed legacy candidates.
    pub(super) fn extract_reachable_table_from_object(
        &self,
        object: &Resolved<'_>,
        path: SemanticPath,
    ) -> Result<Table> {
        let mut typed = object
            .messages
            .iter()
            .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE);
        let canonical = typed.next();
        if canonical.is_some() && typed.next().is_some() {
            return Err(Error::InvalidFormat(format!(
                "Numbers {path} table model contains duplicate canonical payloads"
            )));
        }
        let mut legacy = object
            .messages
            .iter()
            .filter(|message| message.type_ == 6_000);
        let legacy_first = legacy.next();
        let legacy_duplicate = legacy.next().is_some();
        if legacy_duplicate && (self.document_projection || canonical.is_none()) {
            return Err(Error::InvalidFormat(format!(
                "Numbers {path} table model contains duplicate legacy payloads"
            )));
        }
        let message = match (canonical, legacy_first) {
            // Generated compatibility archives can repeat canonical model
            // bytes under the historical model type. Canonical is
            // authoritative for compatibility extraction; native extraction
            // accepts this alias only when the bytes are identical and rejects
            // conflicting payloads.
            (Some(message), Some(legacy))
                if !self.document_projection || message.data == legacy.data =>
            {
                message
            },
            (Some(_), Some(_)) => {
                return Err(Error::InvalidFormat(format!(
                    "Numbers {path} table model has ambiguous payload ownership"
                )));
            },
            (Some(message), None) | (None, Some(message)) => message,
            (None, None) => {
                return Err(Error::InvalidFormat(format!(
                    "Numbers {path} table model has no recognized payload"
                )));
            },
        };
        if !self.document_projection {
            return self.parse_projected_table_payload(message.data.as_slice(), path);
        }
        let mut candidate_budget = *self.projection_budget.borrow();
        let projected =
            match self.project_table_model(message.data.as_slice(), &mut candidate_budget, path) {
                Ok(projected) => projected,
                Err(error) => {
                    self.projection_budget
                        .borrow_mut()
                        .commit_attempt(candidate_budget, false);
                    return Err(error);
                },
            };
        self.parse_projected_table_model(projected, true, Some(candidate_budget))
    }

    fn parse_projected_table_model(
        &self,
        table_model: ProjectedTableModel<'_>,
        name_precharged: bool,
        initial_budget: Option<ProjectionBudget>,
    ) -> Result<Table> {
        let data_store = table_model.base_data_store;
        let compatibility_defaults = table_model.compatibility_defaults;
        self.parse_table_model_parts(
            table_model.table_name,
            table_model.number_of_rows,
            table_model.number_of_columns,
            data_store.string_table().identifier(),
            data_store.formula_table().identifier(),
            data_store
                .formula_error_table()
                .and_then(|reference| nonzero_reference_id(reference.identifier())),
            data_store
                .rich_text_table()
                .and_then(|reference| nonzero_reference_id(reference.identifier())),
            data_store
                .comment_storage_table()
                .and_then(|reference| nonzero_reference_id(reference.identifier())),
            name_precharged,
            initial_budget,
            move |extractor, cell_tables, budget, table| {
                extractor.parse_projected_tiles(
                    data_store.tiles(),
                    compatibility_defaults,
                    cell_tables,
                    budget,
                    table,
                )
            },
        )
    }

    fn parse_table_model_parts<F>(
        &self,
        table_name: &str,
        number_of_rows: u32,
        number_of_columns: u32,
        string_table_id: u64,
        formula_table_id: u64,
        formula_error_table_id: Option<u64>,
        rich_text_table_id: Option<u64>,
        comment_storage_table_id: Option<u64>,
        name_precharged: bool,
        initial_budget: Option<ProjectionBudget>,
        parse_tiles: F,
    ) -> Result<Table>
    where
        F: FnOnce(&Self, &CellTables<'_>, &mut ProjectionBudget, &mut Table) -> Result<()>,
    {
        // Projection is transactional at the table boundary. A rejected legacy
        // candidate must not consume retained-cell or retained-text capacity
        // that belongs to a later schema-proven table. Formula work remains a
        // monotonic package-wide cost across successful and rejected attempts.
        let mut projection_budget =
            initial_budget.unwrap_or_else(|| *self.projection_budget.borrow());
        let result = (|| {
            let (row_count, column_count) =
                checked_table_dimensions(number_of_rows, number_of_columns)?;
            if !name_precharged {
                projection_budget.charge_output_text(table_name.len())?;
            }
            let mut table = Table::with_dimensions(table_name, row_count, column_count)?;

            // Extract string table for cell text values
            // string_table is a required field, not Optional
            let string_table = self.load_string_table(string_table_id, &mut projection_budget)?;

            // Extract formula table for formula cells
            // formula_table is a required field, not Optional
            let formula_table =
                self.load_formula_table(formula_table_id, &mut projection_budget)?;
            let formula_references = if formula_table.is_empty() {
                None
            } else {
                let mut cost = FormulaReferenceBudget::new(
                    projection_budget.remaining_references(),
                    MAX_FORMULA_WORK.min(projection_budget.remaining_payload_work()),
                    projection_budget.remaining_formula_wire_bytes(),
                    projection_budget.remaining_staging_text_bytes(),
                );
                let references =
                    match build_formula_reference_maps(self.bundle, self.object_index, &mut cost) {
                        Ok(references) => references,
                        Err(error) => {
                            projection_budget
                                .retain_formula_map_cost(cost.work_items, cost.wire_bytes);
                            return Err(error);
                        },
                    };
                projection_budget.charge_references(cost.retained_entries)?;
                projection_budget.charge_formula_wire(cost.wire_bytes)?;
                projection_budget.charge_staging_text(cost.text_bytes)?;
                projection_budget.payload_work = projection_charge(
                    projection_budget.payload_work,
                    cost.work_items,
                    MAX_PAYLOAD_WORK,
                    SemanticLimitKind::FormulaWork,
                )?;
                Some(references)
            };
            let empty_formula_references = FormulaReferenceMaps::default();
            let formula_error_table = match formula_error_table_id {
                Some(identifier) => {
                    self.load_formula_error_table(identifier, &mut projection_budget)?
                },
                None => Box::default(),
            };

            let rich_text_table = match rich_text_table_id {
                Some(identifier) => {
                    self.load_rich_text_table(identifier, &mut projection_budget)?
                },
                None => Box::default(),
            };
            let comment_table = if self.retain_comments {
                match comment_storage_table_id {
                    Some(identifier) => {
                        Some(self.load_comment_table(identifier, &mut projection_budget)?)
                    },
                    None => None,
                }
            } else {
                None
            };

            // Parse tiles to extract cell data
            let cell_tables = CellTables {
                strings: &string_table,
                formulas: &formula_table,
                formula_errors: &formula_error_table,
                rich_text: &rich_text_table,
                comments: comment_table.as_ref(),
                formula_references: formula_references
                    .as_ref()
                    .unwrap_or(&empty_formula_references),
            };
            parse_tiles(self, &cell_tables, &mut projection_budget, &mut table)?;

            Ok(table)
        })();

        self.projection_budget
            .borrow_mut()
            .commit_attempt(projection_budget, result.is_ok());
        result
    }

    /// Parse a generated model for the compatibility projection. Rooted
    /// semantic extraction uses the strict borrowed projection below when
    /// `document_projection` is enabled; compatibility callers retain the
    /// keeping every production caller on the strict borrowed projection.
    #[cfg(test)]
    fn parse_table_model(
        &self,
        table_model: tst::TableModelArchive,
        name_precharged: bool,
        initial_budget: Option<ProjectionBudget>,
    ) -> Result<Table> {
        let tst::TableModelArchive {
            table_name,
            number_of_rows,
            number_of_columns,
            base_data_store,
            ..
        } = table_model;
        let string_table_id = base_data_store.string_table.identifier;
        let formula_table_id = base_data_store.formula_table.identifier;
        let formula_error_table_id = base_data_store
            .formula_error_table
            .as_ref()
            .map(|reference| reference.identifier);
        let rich_text_table_id = base_data_store
            .rich_text_table
            .as_ref()
            .map(|reference| reference.identifier);
        let comment_storage_table_id = base_data_store
            .comment_storage_table
            .as_ref()
            .map(|reference| reference.identifier);
        self.parse_table_model_parts(
            &table_name,
            number_of_rows,
            number_of_columns,
            string_table_id,
            formula_table_id,
            formula_error_table_id,
            rich_text_table_id,
            comment_storage_table_id,
            name_precharged,
            initial_budget,
            move |extractor, cell_tables, budget, table| {
                extractor.parse_generated_tiles(&base_data_store.tiles, cell_tables, budget, table)
            },
        )
    }

    /// Load a TableDataList from an object reference
    fn load_string_table(
        &self,
        object_id: u64,
        budget: &mut ProjectionBudget,
    ) -> Result<StringTable> {
        // Compatibility envelopes materialize an omitted/empty required
        // reference as identifier zero.  Zero is a valid generated default,
        // not an object lookup; retain an empty sidecar for that route.
        if object_id == 0 {
            return Ok(Box::default());
        }
        let mut converter =
            |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
             _budget: &mut ProjectionBudget| {
                let value = entry.string_value().ok_or_else(|| {
                    Error::InvalidFormat("Numbers string entry has no string value".to_owned())
                })?;
                let mut retained = String::new();
                retained
                    .try_reserve_exact(value.len())
                    .map_err(|_| allocation_error("Numbers string sidecar", value.len()))?;
                retained.push_str(value);
                Ok(retained)
            };
        self.load_table_data_list_entries(
            object_id,
            tst::table_data_list::ListType::String,
            budget,
            &mut converter,
        )
    }

    fn load_formula_table(
        &self,
        object_id: u64,
        budget: &mut ProjectionBudget,
    ) -> Result<FormulaTable> {
        if object_id == 0 {
            return Ok(Box::default());
        }
        let mut converter =
            |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
             budget: &mut ProjectionBudget| {
                let value = entry.formula().ok_or_else(|| {
                    Error::InvalidFormat("Numbers formula entry has no formula payload".to_owned())
                })?;
                FormulaArchiveBytes::from_wire(value, budget)
            };
        self.load_table_data_list_entries(
            object_id,
            tst::table_data_list::ListType::Formula,
            budget,
            &mut converter,
        )
    }

    fn load_formula_error_table(
        &self,
        object_id: u64,
        budget: &mut ProjectionBudget,
    ) -> Result<FormulaErrorTable> {
        let mut converter =
            |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
             _budget: &mut ProjectionBudget| {
                let value = entry.string_value().ok_or_else(|| {
                    Error::InvalidFormat(
                        "Numbers formula-error entry has no string value".to_owned(),
                    )
                })?;
                let mut retained = String::new();
                retained
                    .try_reserve_exact(value.len())
                    .map_err(|_| allocation_error("Numbers formula-error sidecar", value.len()))?;
                retained.push_str(value);
                Ok(retained)
            };
        self.load_table_data_list_entries(
            object_id,
            tst::table_data_list::ListType::FormulaError,
            budget,
            &mut converter,
        )
    }

    fn load_rich_text_table(
        &self,
        object_id: u64,
        budget: &mut ProjectionBudget,
    ) -> Result<StringTable> {
        let mut converter =
            |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
             budget: &mut ProjectionBudget| {
                let payload_reference = entry.rich_text_payload().ok_or_else(|| {
                    Error::InvalidFormat(
                        "Numbers rich-text entry has no payload reference".to_owned(),
                    )
                })?;
                let payload_object = self
                    .object_index
                    .resolve_ref_id(self.bundle, payload_reference.identifier())?
                    .ok_or_else(|| {
                        Error::InvalidFormat("Numbers rich-text payload is missing".to_owned())
                    })?;
                let mut messages = payload_object
                    .messages
                    .iter()
                    .filter(|message| message.type_ == RICH_TEXT_PAYLOAD_MESSAGE_TYPE);
                let payload_message = messages.next().ok_or_else(|| {
                    Error::InvalidFormat(
                        "Numbers rich-text payload has no canonical message".to_owned(),
                    )
                })?;
                if messages.next().is_some() {
                    return Err(Error::InvalidFormat(
                        "Numbers rich-text payload has duplicate canonical messages".to_owned(),
                    ));
                }
                let (projected_storage, payload_report) =
                    preflight_rich_text_payload(&payload_message.data)?;
                budget.charge_wire_preflight(payload_report)?;
                budget.charge_references(1)?;
                // The bounded preflight above is the authoritative projection of this
                // payload.  Do not decode the complete RichTextPayloadArchive here:
                // it is a tiny envelope whose only parse-relevant value is the local
                // storage reference, and a generated decode would allocate an entire
                // archive before the text-wire budgets have been applied.
                self.extract_rich_text(projected_storage, budget)
            };
        self.load_table_data_list_entries(
            object_id,
            tst::table_data_list::ListType::RichTextPayload,
            budget,
            &mut converter,
        )
    }

    fn load_comment_table(
        &self,
        object_id: u64,
        budget: &mut ProjectionBudget,
    ) -> Result<CommentTable> {
        let mut converter =
            |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
             budget: &mut ProjectionBudget| {
                if entry.ref_count() == 0 {
                    return Err(Error::InvalidFormat(format!(
                        "Numbers comment entry {} has a zero reference count",
                        entry.key()
                    )));
                }
                let storage_id = entry
                    .comment_storage()
                    .map(|reference| reference.identifier())
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers comment entry {} has no storage reference",
                            entry.key()
                        ))
                    })?;
                let storage_object = self
                    .object_index
                    .resolve_ref_id(self.bundle, storage_id)?
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers comment storage object {storage_id} is missing"
                        ))
                    })?;
                let mut payload_count = 0usize;
                // Decode the first candidate into the borrowed semantic
                // snapshot and staged replies. Later candidates still take
                // the exact strict decode path, but use the no-op visitor so
                // duplicate payloads do not allocate a generated reply
                // collection merely to be discarded.
                let mut first_candidate = None;
                for message in storage_object
                    .messages
                    .iter()
                    .filter(|message| message.type_ == 3056)
                {
                    payload_count += 1;
                    let source = message.data.as_slice();
                    let options = comment_storage_codec::DecodeOptions::new(
                        source.len().max(1),
                        budget.remaining_payload_fields(),
                        budget.remaining_payload_work(),
                        64,
                        budget.remaining_references(),
                        budget.remaining_output_text_bytes(),
                    );
                    let reference_offset = budget.references;
                    let field_offset = budget.payload_fields;
                    let work_offset = budget.payload_work;
                    let output_text_offset = budget.output_text_bytes;

                    if first_candidate.is_none() {
                        let mut visitor = CommentReplyVisitor::new(storage_id);
                        let (comment, report) =
                            comment_storage_codec::decode_comment_storage_archive_with_visitor(
                                source,
                                options,
                                &mut visitor,
                            )
                            .map_err(|error| {
                                map_comment_storage_codec_error(
                                    error,
                                    reference_offset,
                                    field_offset,
                                    work_offset,
                                    output_text_offset,
                                )
                            })?;
                        budget.charge_comment_decode_report(report)?;
                        budget.charge_output_text(report.text_bytes())?;
                        let (raw_reply_ids, semantic_error) = visitor.take_parts();
                        // The codec streams every validated `replies` field,
                        // and its report is the authoritative cardinality.
                        // Keep this check at the publication boundary so a
                        // future visitor change cannot silently publish a
                        // truncated reply collection after a successful
                        // decode.  An allocation failure remains the
                        // visitor's typed semantic error and is handled
                        // below, preserving its precedence.
                        if semantic_error.is_none()
                            && raw_reply_ids.len() != report.reply_references()
                        {
                            return Err(Error::MalformedPayload {
                                path: SemanticPath::StructuredTables,
                            });
                        }
                        first_candidate = Some((comment, raw_reply_ids, semantic_error));
                    } else {
                        let (_comment, report) =
                            comment_storage_codec::decode_comment_storage_archive_with_visitor(
                                source,
                                options,
                                &mut (),
                            )
                            .map_err(|error| {
                                map_comment_storage_codec_error(
                                    error,
                                    reference_offset,
                                    field_offset,
                                    work_offset,
                                    output_text_offset,
                                )
                            })?;
                        budget.charge_comment_decode_report(report)?;
                        budget.charge_output_text(report.text_bytes())?;
                    }
                }
                if payload_count == 0 {
                    return Err(Error::InvalidFormat(format!(
                        "Object {storage_id} has no TSD comment-storage payload"
                    )));
                }
                if payload_count != 1 {
                    return Err(Error::InvalidFormat(format!(
                        "Object {storage_id} has multiple TSD comment-storage payloads"
                    )));
                }
                let Some((comment, raw_reply_ids, semantic_error)) = first_candidate else {
                    return Err(Error::MalformedPayload {
                        path: SemanticPath::StructuredTables,
                    });
                };
                if let Some(error) = semantic_error {
                    return Err(error);
                }
                let source_text = comment.text().unwrap_or_default();
                let mut text = String::new();
                text.try_reserve_exact(source_text.len())
                    .map_err(|_| allocation_error("Numbers comment text", source_text.len()))?;
                text.push_str(source_text);
                let mut reply_ids = Vec::new();
                reply_ids
                    .try_reserve_exact(raw_reply_ids.len())
                    .map_err(|_| {
                        allocation_error("Numbers comment replies", raw_reply_ids.len())
                    })?;
                for reply in raw_reply_ids {
                    reply_ids.push(StorageId::from_raw(reply).map_err(map_comment_error)?);
                }
                Ok(Comment {
                    text,
                    creation_date_seconds: comment.creation_date().map(|date| date.seconds()),
                    author_id: comment
                        .author()
                        .map(|author| {
                            AuthorId::from_raw(author.identifier()).map_err(map_comment_error)
                        })
                        .transpose()?,
                    reply_ids: reply_ids.into_boxed_slice(),
                    storage_uuid: comment
                        .storage_uuid()
                        .map(|uuid| {
                            Uuid::from_parts(uuid.lower(), uuid.upper()).map_err(map_comment_error)
                        })
                        .transpose()?,
                })
            };
        self.load_table_data_list_entries(
            object_id,
            tst::table_data_list::ListType::CommentStorage,
            budget,
            &mut converter,
        )
    }

    fn load_table_data_list_entries<T, C>(
        &self,
        object_id: u64,
        list_type: tst::table_data_list::ListType,
        budget: &mut ProjectionBudget,
        converter: &mut C,
    ) -> Result<CompactTable<T>>
    where
        C: ListValueConverter<T>,
    {
        let resolved = self
            .object_index
            .resolve_ref_id(self.bundle, object_id)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers table-data-list object {object_id} is missing"
                ))
            })?;
        let expected = list_type as i32;
        let mut selected_values = None;
        let mut selected_keys = None;
        let mut segment_ids = None;
        let mut structural_error = None;
        let mut semantic_error = None;

        for message in resolved
            .messages
            .iter()
            .filter(|message| message.type_ == 6005 || message.type_ == 6201)
        {
            let list_type_probe = probe_table_data_list_type(&message.data, false, budget)?;
            let duplicate_candidate = selected_values.is_some();
            let admitting_candidate = !duplicate_candidate && list_type_probe == expected;
            let options = table_cell_decode_options(
                &message.data,
                if admitting_candidate {
                    budget.remaining_references()
                } else {
                    usize::MAX
                },
                if admitting_candidate {
                    budget.remaining_staging_text_bytes()
                } else {
                    usize::MAX
                },
                budget.remaining_payload_fields(),
                budget.remaining_payload_work(),
            );
            let reference_offset = budget.references;
            let field_offset = budget.payload_fields;
            let work_offset = budget.payload_work;
            let text_offset = budget.staging_text_bytes;
            let mut visitor =
                TypedListVisitor::new(converter, budget, expected, false, admitting_candidate);
            let decoded = numbers_table_cell_storage_codec::decode_table_data_list_with_visitor(
                &message.data,
                options,
                &mut visitor,
            )
            .map_err(|error| {
                map_table_cell_codec_error_with_offsets(
                    error,
                    reference_offset,
                    field_offset,
                    work_offset,
                    text_offset,
                )
            })?;
            let (snapshot, report) = decoded;
            let (values, keys, references, callback_structural, callback_semantic) =
                visitor.take_parts();
            if snapshot.list_type() != expected {
                // A candidate for another list type is still strictly walked,
                // but its references and text are not admitted to this table.
                budget.charge_decode_work(report)?;
                continue;
            }
            if selected_values.is_some() {
                // A duplicate selected root is structurally inspected but is
                // never admitted to the aggregate reference/text budgets.
                budget.charge_decode_work(report)?;
                record_first_list_error(
                    &mut structural_error,
                    Error::InvalidFormat(format!(
                        "Object {object_id} has multiple Numbers {list_type:?} TableDataList payloads"
                    )),
                );
                continue;
            }
            budget.charge_table_list_decode_report(report)?;
            if let Some(error) = callback_structural {
                record_first_list_error(&mut structural_error, error);
            }
            if let Some(error) = callback_semantic {
                record_first_list_error(&mut semantic_error, error);
            }
            selected_values = Some(values);
            selected_keys = Some(keys);
            segment_ids = Some(references);
        }

        let mut values = selected_values.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Object {object_id} has no Numbers {list_type:?} TableDataList payload"
            ))
        })?;
        let mut keys = selected_keys.unwrap_or_default();
        let segment_ids = segment_ids.unwrap_or_default();

        for segment_id in segment_ids {
            let segment_object = match self.object_index.resolve_ref_id(self.bundle, segment_id) {
                Ok(Some(object)) => object,
                Ok(None) => {
                    record_first_list_error(
                        &mut structural_error,
                        Error::InvalidFormat(format!(
                            "Numbers table-data-list segment object {segment_id} is missing"
                        )),
                    );
                    continue;
                },
                Err(error) => {
                    record_first_list_error(&mut structural_error, error);
                    continue;
                },
            };
            let mut segment_count = 0usize;
            for segment_message in segment_object
                .messages
                .iter()
                .filter(|message| message.type_ == 6011)
            {
                segment_count = segment_count.saturating_add(1);
                let list_type_probe =
                    probe_table_data_list_type(&segment_message.data, true, budget)?;
                let admitting_segment = segment_count == 1 && list_type_probe == expected;
                let options = table_cell_decode_options(
                    &segment_message.data,
                    if admitting_segment {
                        budget.remaining_references()
                    } else {
                        usize::MAX
                    },
                    if admitting_segment {
                        budget.remaining_staging_text_bytes()
                    } else {
                        usize::MAX
                    },
                    budget.remaining_payload_fields(),
                    budget.remaining_payload_work(),
                );
                let reference_offset = budget.references;
                let field_offset = budget.payload_fields;
                let work_offset = budget.payload_work;
                let text_offset = budget.staging_text_bytes;
                let mut visitor =
                    TypedListVisitor::new(converter, budget, expected, true, admitting_segment);
                let decoded =
                    numbers_table_cell_storage_codec::decode_table_data_list_segment_with_visitor(
                        &segment_message.data,
                        options,
                        &mut visitor,
                    )
                    .map_err(|error| {
                        map_table_cell_codec_error_with_offsets(
                            error,
                            reference_offset,
                            field_offset,
                            work_offset,
                            text_offset,
                        )
                    })?;
                let (snapshot, report) = decoded;
                let bounds = visitor.take_segment_bounds();
                let (
                    segment_values,
                    _segment_keys,
                    _segment_refs,
                    callback_structural,
                    callback_semantic,
                ) = visitor.take_parts();
                if segment_count > 1 {
                    // Duplicate segment payloads are fully wire-checked but
                    // cannot contribute references or text to the admitted
                    // segment's aggregate budget.
                    budget.charge_decode_work(report)?;
                    record_first_list_error(
                        &mut structural_error,
                        Error::InvalidFormat(format!(
                            "Object {segment_id} has multiple Numbers TableDataListSegment payloads"
                        )),
                    );
                    continue;
                }
                if snapshot.list_type() != expected {
                    budget.charge_decode_work(report)?;
                    record_first_list_error(
                        &mut structural_error,
                        Error::InvalidFormat(format!(
                            "Numbers table-data-list segment {segment_id} has list type {}, expected {list_type:?}",
                            snapshot.list_type()
                        )),
                    );
                    continue;
                }
                budget.charge_table_list_decode_report(report)?;
                if let Some(error) = callback_structural {
                    record_first_list_error(&mut structural_error, error);
                }
                if let Some(error) = callback_semantic {
                    record_first_list_error(&mut semantic_error, error);
                }
                let end = snapshot
                    .key_range_location()
                    .checked_add(snapshot.key_range_length());
                let Some(end) = end else {
                    record_first_list_error(
                        &mut structural_error,
                        Error::InvalidFormat(format!(
                            "Numbers table-data-list segment {segment_id} key range overflows"
                        )),
                    );
                    continue;
                };
                if let (Some(minimum), Some(maximum)) = bounds
                    && (minimum < snapshot.key_range_location() || maximum >= end)
                {
                    record_first_list_error(
                        &mut structural_error,
                        Error::InvalidFormat(format!(
                            "Numbers table-data-list segment {segment_id} contains an entry outside its key range"
                        )),
                    );
                    continue;
                }
                for (key, value) in segment_values {
                    if keys.contains(&key) {
                        record_first_list_error(
                            &mut structural_error,
                            Error::InvalidFormat(format!(
                                "Numbers {list_type:?} table {object_id} repeats entry key {key} across root and segments"
                            )),
                        );
                        continue;
                    }
                    if keys.try_reserve(1).is_err() {
                        record_first_list_error(
                            &mut semantic_error,
                            allocation_error("Numbers table-list entry keys", keys.len() + 1),
                        );
                        continue;
                    }
                    keys.insert(key);
                    if values.try_reserve(1).is_err() {
                        record_first_list_error(
                            &mut semantic_error,
                            allocation_error("Numbers table-list entries", values.len() + 1),
                        );
                        continue;
                    }
                    values.push((key, value));
                }
            }
            if segment_count == 0 {
                record_first_list_error(
                    &mut structural_error,
                    Error::InvalidFormat(format!(
                        "Object {segment_id} has no Numbers TableDataListSegment payload"
                    )),
                );
            }
        }
        if let Some(error) = structural_error {
            return Err(error);
        }
        if let Some(error) = semantic_error {
            return Err(error);
        }
        compact_table_vec(values)
    }

    /// Strictly decode tile-storage metadata, stage its references, then
    /// resolve tiles only after the enclosing payload has passed wire parity.
    fn parse_projected_tiles(
        &self,
        source: &[u8],
        compatibility_defaults: bool,
        cell_tables: &CellTables<'_>,
        projection_budget: &mut ProjectionBudget,
        table: &mut Table,
    ) -> Result<()> {
        // Prost materializes an omitted TileStorage message as its generated
        // default. Preserve that historical compatibility behavior while
        // keeping the native strict path's required-envelope validation.
        if compatibility_defaults && source.is_empty() {
            return Ok(());
        }
        let options = table_cell_decode_options(
            source,
            crate::MAX_REFERENCES,
            usize::MAX,
            projection_budget.remaining_payload_fields(),
            projection_budget.remaining_payload_work(),
        );
        let reference_offset = projection_budget.references;
        let field_offset = projection_budget.payload_fields;
        let work_offset = projection_budget.payload_work;
        let mut visitor = TileStorageReferenceStage::new();
        let (snapshot, report) =
            numbers_table_cell_storage_codec::decode_tile_storage_with_visitor(
                source,
                options,
                &mut visitor,
            )
            .map_err(|error| {
                map_table_cell_codec_error_with_offsets(
                    error,
                    reference_offset,
                    field_offset,
                    work_offset,
                    0,
                )
            })?;
        projection_budget.charge_decode_work(report)?;
        let references = visitor.take()?;
        let tile_size = usize::try_from(snapshot.tile_size().unwrap_or(256)).map_err(|_| {
            Error::InvalidFormat("Numbers tile size does not fit the host usize".to_owned())
        })?;
        self.parse_tiles_from_references(
            tile_size,
            references,
            cell_tables,
            projection_budget,
            table,
        )
    }

    fn parse_tiles_from_references(
        &self,
        tile_size: usize,
        references: Vec<(u32, u64)>,
        cell_tables: &CellTables<'_>,
        projection_budget: &mut ProjectionBudget,
        table: &mut Table,
    ) -> Result<()> {
        if tile_size == 0 {
            return Err(Error::InvalidFormat(
                "Numbers table declares a zero tile size".to_owned(),
            ));
        }
        let tile_count = if table.row_count() == 0 {
            0
        } else {
            (table.row_count() - 1) / tile_size + 1
        };
        let mut seen_tile_ids = HashSet::new();
        seen_tile_ids
            .try_reserve(references.len())
            .map_err(|_| allocation_error("Numbers tile keys", references.len()))?;
        let mut budget = CellBudget::new();
        // Resolve each tile reference and parse its contents only after the
        // strict tile-storage envelope has completed successfully.
        for (tile_id, tile_reference) in references {
            let tile_key = usize::try_from(tile_id).map_err(|_| {
                Error::InvalidFormat("Numbers tile key does not fit the host usize".to_owned())
            })?;
            if tile_key >= tile_count {
                return Err(Error::InvalidFormat(format!(
                    "Numbers tile key {tile_key} is outside the declared table height {}",
                    table.row_count()
                )));
            }
            if !seen_tile_ids.insert(tile_id) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers table repeats tile key {tile_key}"
                )));
            }
            let row_origin = tile_key
                .checked_mul(tile_size)
                .ok_or_else(|| Error::ParseError("Numbers tile row origin overflow".to_owned()))?;
            self.parse_tile(
                tile_reference,
                row_origin,
                tile_size,
                table.row_count(),
                table.column_count(),
                &mut budget,
                cell_tables,
                projection_budget,
                table,
            )?;
        }
        Ok(())
    }

    /// Test-only generated fixture adapter. Production receives only the
    /// borrowed strict tile-storage snapshot above.
    #[cfg(test)]
    fn parse_generated_tiles(
        &self,
        tile_storage: &tst::TileStorage,
        cell_tables: &CellTables<'_>,
        projection_budget: &mut ProjectionBudget,
        table: &mut Table,
    ) -> Result<()> {
        let tile_size = usize::try_from(tile_storage.tile_size.unwrap_or(256)).map_err(|_| {
            Error::InvalidFormat("Numbers tile size does not fit the host usize".to_owned())
        })?;
        if tile_size == 0 {
            return Err(Error::InvalidFormat(
                "Numbers table declares a zero tile size".to_owned(),
            ));
        }
        let tile_count = if table.row_count() == 0 {
            0
        } else {
            (table.row_count() - 1) / tile_size + 1
        };
        let mut seen_tile_ids = HashSet::new();
        seen_tile_ids
            .try_reserve(tile_storage.tiles.len())
            .map_err(|_| allocation_error("Numbers tile keys", tile_storage.tiles.len()))?;
        let mut budget = CellBudget::new();
        // Resolve each tile reference and parse its contents
        for tile_ref in &tile_storage.tiles {
            let tile_key = usize::try_from(tile_ref.tileid).map_err(|_| {
                Error::InvalidFormat("Numbers tile key does not fit the host usize".to_owned())
            })?;
            if tile_key >= tile_count {
                return Err(Error::InvalidFormat(format!(
                    "Numbers tile key {tile_key} is outside the declared table height {}",
                    table.row_count()
                )));
            }
            if !seen_tile_ids.insert(tile_ref.tileid) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers table repeats tile key {tile_key}"
                )));
            }
            let row_origin = tile_key
                .checked_mul(tile_size)
                .ok_or_else(|| Error::ParseError("Numbers tile row origin overflow".to_owned()))?;
            // tile is a required field, not Optional
            let tile_reference = &tile_ref.tile;
            self.parse_tile(
                tile_reference.identifier,
                row_origin,
                tile_size,
                table.row_count(),
                table.column_count(),
                &mut budget,
                cell_tables,
                projection_budget,
                table,
            )?;
        }

        Ok(())
    }

    /// Parse a single tile object
    fn parse_tile(
        &self,
        tile_id: u64,
        row_origin: usize,
        tile_size: usize,
        row_count: usize,
        column_count: usize,
        budget: &mut CellBudget,
        cell_tables: &CellTables<'_>,
        projection_budget: &mut ProjectionBudget,
        table: &mut Table,
    ) -> Result<()> {
        let resolved = self
            .object_index
            .resolve_ref_id(self.bundle, tile_id)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers tile object {tile_id} referenced by table is missing"
                ))
            })?;
        let mut decoded = false;
        for msg in resolved.messages {
            if msg.type_ != TILE_MESSAGE_TYPE {
                continue;
            }
            if decoded {
                return Err(Error::InvalidFormat(format!(
                    "Numbers tile object {tile_id} contains multiple tile payloads"
                )));
            }
            let options = table_cell_decode_options(
                &msg.data,
                projection_budget.remaining_references(),
                projection_budget.remaining_output_text_bytes(),
                projection_budget.remaining_payload_fields(),
                projection_budget.remaining_payload_work(),
            );
            let reference_offset = projection_budget.references;
            let field_offset = projection_budget.payload_fields;
            let work_offset = projection_budget.payload_work;
            let text_offset = projection_budget.output_text_bytes;
            let (materialized_cells, semantic_error, report) = {
                let mut visitor = TileRowVisitor {
                    row_origin,
                    tile_size,
                    row_count,
                    column_count,
                    budget,
                    cell_tables,
                    projection_budget,
                    table,
                    materialized_cells: 0,
                    semantic_error: None,
                };
                let (_, report) = numbers_table_cell_storage_codec::decode_tile_with_visitor(
                    &msg.data,
                    options,
                    &mut visitor,
                )
                .map_err(|error| {
                    map_table_cell_codec_error_with_offsets(
                        error,
                        reference_offset,
                        field_offset,
                        work_offset,
                        text_offset,
                    )
                })?;
                (
                    visitor.materialized_cells,
                    visitor.semantic_error.take(),
                    report,
                )
            };
            projection_budget.charge_decode_report(report)?;
            projection_budget.charge_materialized_cells(materialized_cells)?;
            if let Some(error) = semantic_error {
                return Err(error);
            }
            decoded = true;
        }

        if !decoded {
            return Err(Error::InvalidFormat(format!(
                "Numbers tile object {tile_id} has no tile payload"
            )));
        }

        Ok(())
    }

    /// Parse a single tile row
    fn parse_tile_row(
        row_info: &numbers_table_cell_storage_codec::TileRowInfoSnapshot<'_>,
        row_origin: usize,
        tile_size: usize,
        row_count: usize,
        column_count: usize,
        budget: &mut CellBudget,
        cell_tables: &CellTables<'_>,
        projection_budget: &mut ProjectionBudget,
        table: &mut Table,
    ) -> Result<()> {
        let tile_row_index = usize::try_from(row_info.tile_row_index()).map_err(|_| {
            Error::InvalidFormat("Numbers tile row index does not fit the host usize".to_owned())
        })?;
        if tile_row_index >= tile_size {
            return Err(Error::InvalidFormat(format!(
                "Numbers tile row {} is outside tile size {tile_size}",
                row_info.tile_row_index()
            )));
        }
        let row_index = row_origin
            .checked_add(tile_row_index)
            .ok_or_else(|| Error::ParseError("Numbers tile row index overflow".to_owned()))?;
        validate_table_row(row_index, row_count)?;

        // The cell_storage_buffer contains serialized Cell messages
        // The cell_offsets buffer contains the byte offsets for each cell

        let (cell_storage, cell_offsets) =
            match (row_info.cell_storage_buffer(), row_info.cell_offsets()) {
                (Some(storage), Some(offsets)) => (storage, offsets),
                _ => (
                    row_info.cell_storage_buffer_pre_bnc(),
                    row_info.cell_offsets_pre_bnc(),
                ),
            };

        let expected_cells = usize::try_from(row_info.cell_count()).map_err(|_| {
            Error::InvalidFormat("Numbers cell count does not fit the host usize".to_owned())
        })?;
        budget.check(expected_cells)?;
        // The strict borrowed-row projection charged the aggregate declared
        // cell count before row-level cell decoding.
        let cells = Self::parse_cell_offsets(
            cell_offsets,
            cell_storage.len(),
            row_info.has_wide_offsets().unwrap_or(false),
            expected_cells,
            column_count,
        )?;
        budget.consume(cells.len())?;

        for (column_index, range) in cells {
            validate_table_column(column_index, column_count)?;
            let parsed = Self::parse_cell_storage(
                &cell_storage[range],
                cell_tables,
                projection_budget,
                row_index,
                column_index,
                row_count,
                column_count,
            )?;
            table.try_set_cell(row_index, column_index, parsed.value)?;
            if let Some(identifier) = parsed.comment_identifier
                && let Some(comments) = cell_tables.comments
            {
                let comment = compact_table_get(comments, identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers comment table has no entry {identifier} referenced by cell ({row_index}, {column_index})"
                    ))
                })?;
                // Admit and materialize the owned copy as one transaction. A
                // sidecar comment may be referenced by many cells; charging
                // each copy before insertion prevents repeated text/reply
                // cloning from bypassing the projection budget. Keep the
                // caller's budget unchanged if the fallible clone or table
                // insertion fails.
                let mut candidate_budget = *projection_budget;
                candidate_budget.charge_comment_materialization(comment)?;
                let materialized = try_clone_comment(comment)?;
                table.try_set_comment(row_index, column_index, materialized)?;
                *projection_budget = candidate_budget;
            }
        }

        Ok(())
    }

    /// Parse cell offsets from the offsets buffer
    ///
    /// The offset table is an array of little-endian `u16` values. `0xffff`
    /// marks a missing column; wide rows store offsets in four-byte units.
    /// Native producers may pad the table past the semantic table width, but
    /// every padded slot must retain the missing-column sentinel.
    fn parse_cell_offsets(
        offsets_buffer: &[u8],
        storage_length: usize,
        wide_offsets: bool,
        expected_cells: usize,
        column_count: usize,
    ) -> Result<Vec<(usize, std::ops::Range<usize>)>> {
        if !offsets_buffer.len().is_multiple_of(2) {
            return Err(Error::ParseError(
                "Numbers cell offset table has an odd byte length".to_string(),
            ));
        }

        let slot_count = offsets_buffer.len() / 2;
        if expected_cells > slot_count {
            return Err(Error::ParseError(format!(
                "Numbers row declares {expected_cells} cells but has only {slot_count} offset slots"
            )));
        }
        if expected_cells > column_count {
            return Err(Error::InvalidFormat(format!(
                "Numbers row declares {expected_cells} cells but table width is {column_count}"
            )));
        }
        if let Some((column, _bytes)) = offsets_buffer
            .chunks_exact(2)
            .enumerate()
            .skip(column_count)
            .find(|(_column, bytes)| u16::from_le_bytes([bytes[0], bytes[1]]) != u16::MAX)
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers cell offset at column {column} is outside the declared table width {column_count}"
            )));
        }

        let present_cells = offsets_buffer
            .chunks_exact(2)
            .take(column_count)
            .filter(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]) != u16::MAX)
            .count();
        if present_cells != expected_cells {
            return Err(Error::ParseError(format!(
                "Numbers row declares {expected_cells} cells but has {present_cells} offsets"
            )));
        }

        let width = if wide_offsets { 4usize } else { 1usize };
        let mut cells = Vec::new();
        cells
            .try_reserve_exact(expected_cells)
            .map_err(|_| allocation_error("Numbers cell ranges", expected_cells))?;
        let mut previous = None;
        for (column, bytes) in offsets_buffer
            .chunks_exact(2)
            .take(column_count)
            .enumerate()
        {
            let raw_offset = u16::from_le_bytes([bytes[0], bytes[1]]);
            if raw_offset == u16::MAX {
                continue;
            }
            let offset = usize::from(raw_offset)
                .checked_mul(width)
                .ok_or_else(|| Error::ParseError("Numbers cell offset overflow".to_string()))?;
            if offset >= storage_length {
                return Err(Error::ParseError(format!(
                    "Numbers cell offset {offset} exceeds storage length {storage_length}"
                )));
            }
            if let Some((previous_column, previous_offset)) = previous {
                if offset <= previous_offset {
                    return Err(Error::ParseError(format!(
                        "Numbers cell offsets are not strictly increasing: {previous_offset} then {offset}"
                    )));
                }
                cells.push((previous_column, previous_offset..offset));
            }
            previous = Some((column, offset));
        }
        if let Some((column, start)) = previous {
            if storage_length <= start {
                return Err(Error::ParseError(format!(
                    "Numbers cell offset range ends at {storage_length} after {start}"
                )));
            }
            cells.push((column, start..storage_length));
        }
        Ok(cells)
    }

    fn parse_cell_storage(
        data: &[u8],
        cell_tables: &CellTables<'_>,
        projection_budget: &mut ProjectionBudget,
        row: usize,
        column: usize,
        row_count: usize,
        column_count: usize,
    ) -> Result<ParsedCell> {
        let version = *data
            .first()
            .ok_or_else(|| Error::ParseError("Empty Numbers cell storage".to_string()))?;
        match version {
            0..=4 => Self::parse_pre_bnc_cell(
                data,
                cell_tables,
                projection_budget,
                row,
                column,
                row_count,
                column_count,
            ),
            5 => Self::parse_bnc_cell(
                data,
                cell_tables,
                projection_budget,
                row,
                column,
                row_count,
                column_count,
            ),
            other => Err(Error::ParseError(format!(
                "Unsupported Numbers cell storage version {other}"
            ))),
        }
    }

    fn parse_bnc_cell(
        data: &[u8],
        cell_tables: &CellTables<'_>,
        projection_budget: &mut ProjectionBudget,
        row: usize,
        column: usize,
        row_count: usize,
        column_count: usize,
    ) -> Result<ParsedCell> {
        let cell = BncCellView::parse(data).map_err(|error| {
            Error::ParseError(format!(
                "Numbers BNC cell ({row}, {column}) is invalid: {error}"
            ))
        })?;
        let comment_identifier = cell.comment_identifier();

        if let StoredValue::Formula(identifier) = cell.stored_value() {
            let formula = compact_table_get(cell_tables.formulas, identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers formula table has no entry {identifier} referenced by cell ({row}, {column})"
                ))
            })?;
            let rendered = Self::extract_formula_string(
                formula,
                row,
                column,
                row_count,
                column_count,
                cell_tables.formula_references,
                projection_budget,
            )
            .map_err(|error| {
                Error::ParseError(format!(
                    "Numbers formula {identifier} at cell ({row}, {column}) is invalid: {error}"
                ))
            })?;
            return Ok(ParsedCell {
                value: CellValue::Formula(rendered),
                comment_identifier,
            });
        }

        let zero = finite_zero()?;
        let scalar = cell.cached_scalar();
        let value = match cell.stored_value() {
            StoredValue::Empty => CellValue::Empty,
            StoredValue::Number => match scalar {
                Some(CachedScalar::Number(value)) => CellValue::Number(value),
                Some(
                    CachedScalar::Boolean(_) | CachedScalar::Date(_) | CachedScalar::Duration(_),
                ) => {
                    return Err(Error::InvalidFormat(format!(
                        "Numbers numeric BNC cell ({row}, {column}) has a mismatched scalar encoding"
                    )));
                },
                Some(CachedScalar::Unsupported(_)) | None => CellValue::Number(zero),
            },
            StoredValue::Text(identifier) => {
                retained_table_text(cell_tables.strings, identifier, projection_budget)?
                    .map_or(CellValue::Empty, CellValue::Text)
            },
            StoredValue::RichText(identifier) => {
                retained_table_text(cell_tables.rich_text, identifier, projection_budget)?
                    .map_or(CellValue::Empty, CellValue::Text)
            },
            StoredValue::Date => match scalar {
                Some(CachedScalar::Date(value)) => CellValue::Date(value),
                Some(_) | None => CellValue::Date(zero),
            },
            StoredValue::Boolean => match scalar {
                Some(CachedScalar::Boolean(value)) => CellValue::Boolean(value),
                Some(_) | None => CellValue::Boolean(false),
            },
            StoredValue::Duration => match scalar {
                Some(CachedScalar::Duration(value)) => CellValue::Duration(value),
                Some(_) | None => CellValue::Duration(zero),
            },
            StoredValue::Error => {
                let error = cell
                    .formula_error_identifier()
                    .and_then(|id| compact_table_get(cell_tables.formula_errors, id))
                    .map_or("FORMULA", String::as_str);
                CellValue::Error(retain_text(error, projection_budget)?)
            },
            StoredValue::Formula(_) => {
                return Err(Error::InvalidFormat(format!(
                    "Numbers formula BNC cell ({row}, {column}) reached scalar decoding"
                )));
            },
            StoredValue::Unsupported(other) => {
                return Err(Error::ParseError(format!(
                    "Unsupported Numbers BNC cell type {other}"
                )));
            },
        };
        Ok(ParsedCell {
            value,
            comment_identifier,
        })
    }

    fn parse_pre_bnc_cell(
        data: &[u8],
        cell_tables: &CellTables<'_>,
        projection_budget: &mut ProjectionBudget,
        row: usize,
        column: usize,
        row_count: usize,
        column_count: usize,
    ) -> Result<ParsedCell> {
        let version = data
            .first()
            .copied()
            .ok_or_else(|| Error::ParseError("Empty Numbers pre-BNC cell payload".to_owned()))?;
        let header_length = if version <= 1 { 8 } else { 12 };
        if data.len() < header_length {
            return Err(Error::ParseError(
                "Truncated Numbers pre-BNC cell header".to_string(),
            ));
        }
        let cell_type = data[if version == 4 { 1 } else { 2 }];
        let flags = if version <= 1 {
            u32::from(u16::from_le_bytes([data[4], data[5]]))
        } else {
            read_u32_le(&data[4..8])?
        };
        let mut cursor = header_length;
        let mut number: Option<FiniteF64> = None;
        let mut date: Option<FiniteF64> = None;
        let mut string_id = None;
        let mut rich_text_id = None;
        let mut formula_id = None;
        let mut formula_error_id = None;
        let mut comment_identifier = None;

        for (flag, size) in [
            (0x000002, 4),
            (0x000080, 4),
            (0x000400, 4),
            (0x000800, 4),
            (0x000004, 4),
            (0x000008, 4),
            (0x000100, 4),
            (0x000200, 4),
            (0x001000, 4),
            (0x002000, 4),
            (0x000010, 4),
            (0x000020, 8),
            (0x000040, 8),
            (0x010000, 4),
            (0x080000, 4),
            (0x020000, 4),
            (0x040000, 4),
            (0x100000, 4),
            (0x200000, 4),
            (0x400000, 4),
            (0x800000, 4),
        ] {
            if flags & flag == 0 {
                continue;
            }
            let field = take_field(data, &mut cursor, size)?;
            match flag {
                0x000008 => formula_id = Some(read_u32_le(field)?),
                0x000100 => formula_error_id = Some(read_u32_le(field)?),
                0x001000 => comment_identifier = Some(read_u32_le(field)?),
                0x000200 => rich_text_id = Some(read_u32_le(field)?),
                0x000010 => string_id = Some(read_u32_le(field)?),
                0x000020 => number = Some(read_f64_le(field)?),
                0x000040 => date = Some(read_f64_le(field)?),
                _ => {},
            }
        }

        if let Some(identifier) = formula_id {
            let formula = compact_table_get(cell_tables.formulas, identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers formula table has no entry {identifier} referenced by cell ({row}, {column})"
                ))
            })?;
            let rendered = Self::extract_formula_string(
                formula,
                row,
                column,
                row_count,
                column_count,
                cell_tables.formula_references,
                projection_budget,
            )
            .map_err(|error| {
                Error::ParseError(format!(
                    "Numbers formula {identifier} at cell ({row}, {column}) is invalid: {error}"
                ))
            })?;
            return Ok(ParsedCell {
                value: CellValue::Formula(rendered),
                comment_identifier,
            });
        }

        let zero = finite_zero()?;
        let value = match cell_type {
            0 => CellValue::Empty,
            2 => CellValue::Number(number.unwrap_or(zero)),
            3 => match string_id {
                Some(identifier) => {
                    retained_table_text(cell_tables.strings, identifier, projection_budget)?
                        .map_or(CellValue::Empty, CellValue::Text)
                },
                None => CellValue::Empty,
            },
            5 => CellValue::Date(date.unwrap_or(zero)),
            6 => CellValue::Boolean(number.unwrap_or(zero).get() != 0.0),
            7 => CellValue::Duration(number.unwrap_or(zero)),
            8 => {
                let error = formula_error_id
                    .and_then(|id| compact_table_get(cell_tables.formula_errors, id))
                    .map_or("FORMULA", String::as_str);
                CellValue::Error(retain_text(error, projection_budget)?)
            },
            9 => match rich_text_id {
                Some(identifier) => {
                    retained_table_text(cell_tables.rich_text, identifier, projection_budget)?
                        .map_or(CellValue::Empty, CellValue::Text)
                },
                None => CellValue::Empty,
            },
            other => {
                return Err(Error::ParseError(format!(
                    "Unsupported Numbers pre-BNC cell type {other}"
                )));
            },
        };
        Ok(ParsedCell {
            value,
            comment_identifier,
        })
    }

    /// Render a formula through the bounded, non-copying expression arena.
    fn extract_formula_string(
        formula: &FormulaArchiveBytes,
        host_row: usize,
        host_column: usize,
        row_count: usize,
        column_count: usize,
        formula_references: &FormulaReferenceMaps,
        projection_budget: &mut ProjectionBudget,
    ) -> Result<String> {
        let render_work_before_scalar = projection_budget.formula_render_work;
        if formula.scalar_visitor_eligible
            && let Some(rendered) = render_scalar_formula(
                formula,
                host_row,
                host_column,
                row_count,
                column_count,
                projection_budget,
            )?
        {
            return Ok(rendered);
        }
        // The scalar probe charges its exact node bound before allocating its
        // visitor. If it declines after that charge, carry the same admission
        // into the compatibility renderer; otherwise admit the preflighted
        // bound before its lazy traversal starts.
        let nodes_precharged = projection_budget
            .formula_render_work
            .saturating_sub(render_work_before_scalar)
            >= formula.scalar_visitor_node_count;
        if !nodes_precharged {
            projection_budget.charge_formula_render_work(formula.scalar_visitor_node_count)?;
        }
        projection_budget.charge_formula_lazy_work(formula.lazy_traversal_entry_count)?;
        render_formula_compatibility(
            formula,
            host_row,
            host_column,
            row_count,
            column_count,
            formula_references,
            projection_budget,
        )
    }

    /// Test-only reference renderer retained for differential coverage while
    /// the streaming FormulaArchive reader is migrated independently.
    ///
    ///   - Reconstructs formula text from Abstract Syntax Tree
    ///   - Handles operators, functions, cell references, and constants
    ///   - Based on TSCE.ASTNodeArrayArchive protobuf structure
    ///   - Implements reverse-polish notation to infix conversion
    ///
    /// iWork stores formulas as Abstract Syntax Trees (AST) in reverse-polish
    /// notation (postfix). This function reconstructs the formula text by
    /// traversing the AST and converting it to standard infix notation.
    ///
    /// # Performance
    ///
    /// O(n) where n is the number of AST nodes. Uses a stack-based algorithm
    /// for efficient conversion.
    #[cfg(test)]
    fn extract_formula_string_reference(
        formula: &tsce::FormulaArchive,
        host_row: usize,
        host_column: usize,
        formula_references: &FormulaReferenceMaps,
    ) -> Result<String> {
        use litchi_iwa_protos::tsce::ast_node_array_archive::AstNodeType;

        let ast_array = &formula.ast_node_array;

        // Formulas are stored in reverse-polish notation (postfix)
        // We need to convert to infix notation using a stack
        if ast_array.ast_node.is_empty() {
            return Ok("=".to_string());
        }

        // Stack to hold expression parts during reconstruction
        let mut expr_stack: Vec<String> = Vec::new();

        // Process each AST node
        for node in &ast_array.ast_node {
            let ast_node_type = node.ast_node_type();

            match ast_node_type {
                // Arithmetic operators (binary)
                AstNodeType::AdditionNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "addition")?;
                    expr_stack.push(format!("({}+{})", left, right));
                },
                AstNodeType::SubtractionNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "subtraction")?;
                    expr_stack.push(format!("({}-{})", left, right));
                },
                AstNodeType::MultiplicationNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "multiplication")?;
                    expr_stack.push(format!("({}*{})", left, right));
                },
                AstNodeType::DivisionNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "division")?;
                    expr_stack.push(format!("({}/{})", left, right));
                },
                AstNodeType::PowerNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "power")?;
                    expr_stack.push(format!("({}^{})", left, right));
                },
                AstNodeType::GreaterThanNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "greater than")?;
                    expr_stack.push(format!("({left}>{right})"));
                },
                AstNodeType::GreaterThanOrEqualToNode => {
                    let (left, right) =
                        pop_binary_operands(&mut expr_stack, "greater than or equal")?;
                    expr_stack.push(format!("({left}>={right})"));
                },
                AstNodeType::LessThanNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "less than")?;
                    expr_stack.push(format!("({left}<{right})"));
                },
                AstNodeType::LessThanOrEqualToNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "less than or equal")?;
                    expr_stack.push(format!("({left}<={right})"));
                },
                AstNodeType::EqualToNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "equality")?;
                    expr_stack.push(format!("({left}={right})"));
                },
                AstNodeType::NotEqualToNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "inequality")?;
                    expr_stack.push(format!("({left}<>{right})"));
                },

                // Constants
                AstNodeType::NumberNode => {
                    if let Some(number) = node.ast_number_node_number {
                        expr_stack.push(number.to_string());
                    }
                },
                AstNodeType::StringNode => {
                    if let Some(ref string) = node.ast_string_node_string {
                        expr_stack.push(format!("\"{}\"", string.replace('"', "\"\"")));
                    }
                },
                AstNodeType::BooleanNode => {
                    if let Some(boolean) = node.ast_boolean_node_boolean {
                        expr_stack.push(if boolean { "TRUE" } else { "FALSE" }.to_string());
                    }
                },
                AstNodeType::TokenNode => {
                    if let Some(boolean) = node.ast_token_node_boolean {
                        expr_stack.push(if boolean { "TRUE" } else { "FALSE" }.to_owned());
                    }
                },
                AstNodeType::DateNode => {
                    if let Some(seconds) = node.ast_date_node_date_num {
                        expr_stack.push(format!("(DATE(2001,1,1)+{})", seconds / 86_400.0));
                    }
                },
                AstNodeType::DurationNode => {
                    if let Some(value) = node.ast_duration_node_unit_num {
                        expr_stack.push(value.to_string());
                    }
                },
                AstNodeType::EmptyArgumentNode => expr_stack.push(String::new()),

                // Cell references
                AstNodeType::CellReferenceNode => {
                    if let (Some(ast_column), Some(ast_row)) = (&node.ast_column, &node.ast_row) {
                        let column = resolve_formula_coordinate(
                            host_column,
                            ast_column.column,
                            ast_column.absolute.unwrap_or(false),
                            "column",
                        )?;
                        let row = resolve_formula_coordinate(
                            host_row,
                            ast_row.row,
                            ast_row.absolute.unwrap_or(false),
                            "row",
                        )?;
                        let column_absolute = ast_column.absolute.unwrap_or(false);
                        let row_absolute = ast_row.absolute.unwrap_or(false);
                        let prefix = node
                            .ast_cross_table_reference_extra_info
                            .as_ref()
                            .map(|extra| {
                                formula_reference_prefix(&extra.table_id, formula_references)
                            })
                            .unwrap_or_default();
                        expr_stack.push(format!(
                            "{prefix}{}{}{}{}",
                            if column_absolute { "$" } else { "" },
                            Self::column_index_to_letter(column),
                            if row_absolute { "$" } else { "" },
                            checked_formula_row_number(row)?
                        ));
                    } else if let Some(ref cell_ref) = node.ast_local_cell_reference_node_reference
                    {
                        // Convert row/column handles to A1 notation
                        let col_letter = Self::column_index_to_letter(cell_ref.column_handle);
                        let row_num = checked_formula_row_number(cell_ref.row_handle)?;
                        let col_sticky = if cell_ref.column_is_sticky != 0 {
                            "$"
                        } else {
                            ""
                        };
                        let row_sticky = if cell_ref.row_is_sticky != 0 { "$" } else { "" };
                        expr_stack.push(format!(
                            "{}{}{}{}",
                            col_sticky, col_letter, row_sticky, row_num
                        ));
                    } else if let Some(ref cross_ref) =
                        node.ast_cross_table_cell_reference_node_reference
                    {
                        // Cross-table reference
                        let col_letter = Self::column_index_to_letter(cross_ref.column_handle);
                        let row_num = checked_formula_row_number(cross_ref.row_handle)?;
                        let prefix =
                            formula_reference_prefix(&cross_ref.table_id, formula_references);
                        expr_stack.push(format!("{prefix}{col_letter}{row_num}"));
                    } else {
                        expr_stack.push("#REF!".to_owned());
                    }
                },
                AstNodeType::LocalCellReferenceNode => {
                    if let Some(cell_ref) = &node.ast_local_cell_reference_node_reference {
                        let col_letter = Self::column_index_to_letter(cell_ref.column_handle);
                        let row_num = checked_formula_row_number(cell_ref.row_handle)?;
                        expr_stack.push(format!("{}{}", col_letter, row_num));
                    } else {
                        expr_stack.push("#REF!".to_owned());
                    }
                },
                AstNodeType::CrossTableCellReferenceNode => {
                    if let Some(cell_ref) = &node.ast_cross_table_cell_reference_node_reference {
                        let col_letter = Self::column_index_to_letter(cell_ref.column_handle);
                        let prefix =
                            formula_reference_prefix(&cell_ref.table_id, formula_references);
                        let row_num = checked_formula_row_number(cell_ref.row_handle)?;
                        expr_stack.push(format!("{prefix}{col_letter}{row_num}"));
                    } else {
                        expr_stack.push("#REF!".to_owned());
                    }
                },

                // Functions
                AstNodeType::FunctionNode => {
                    if let Some(function_index) = node.ast_function_node_index {
                        let num_args = node.ast_function_node_num_args.unwrap_or(0);
                        let function_name = Self::get_function_name(function_index);

                        // Pop arguments from stack (in reverse order)
                        let args = pop_formula_arguments(&mut expr_stack, num_args, "function")?;

                        let args_str = args.join(",");
                        expr_stack.push(format!("{}({})", function_name, args_str));
                    }
                },

                // List (for function arguments)
                AstNodeType::ListNode => {
                    if let Some(num_args) = node.ast_list_node_num_args {
                        // Collect arguments
                        let args = pop_formula_arguments(&mut expr_stack, num_args, "list")?;
                        expr_stack.push(args.join(","));
                    }
                },
                AstNodeType::ArrayNode => {
                    let columns = node.ast_array_node_num_col.unwrap_or(0);
                    let rows = node.ast_array_node_num_row.unwrap_or(0);
                    let count = columns.checked_mul(rows).ok_or_else(|| {
                        Error::ParseError("Numbers formula array size overflow".to_owned())
                    })?;
                    let values = pop_formula_arguments(&mut expr_stack, count, "array")?;
                    let columns = usize::try_from(columns).map_err(|_| {
                        Error::ParseError("Numbers formula array width exceeds usize".to_owned())
                    })?;
                    let rendered = if columns == 0 {
                        String::new()
                    } else {
                        values
                            .chunks(columns)
                            .map(|row| row.join(","))
                            .collect::<Vec<_>>()
                            .join(";")
                    };
                    expr_stack.push(format!("{{{rendered}}}"));
                },
                AstNodeType::ThunkNode => {
                    if let Some(array) = &node.ast_thunk_node_array {
                        let nested = tsce::FormulaArchive {
                            ast_node_array: array.clone(),
                            ..Default::default()
                        };
                        let rendered = Self::extract_formula_string_reference(
                            &nested,
                            host_row,
                            host_column,
                            formula_references,
                        )?;
                        expr_stack.push(rendered.trim_start_matches('=').to_owned());
                    }
                },

                // Unary operators - represented differently in the AST
                // Numbers uses NegationNode instead of UnaryMinusNode
                AstNodeType::NegationNode => {
                    if let Some(operand) = expr_stack.pop() {
                        expr_stack.push(format!("-({})", operand));
                    }
                },
                AstNodeType::PercentNode => {
                    let operand = expr_stack.pop().ok_or_else(|| {
                        Error::ParseError(
                            "Numbers formula percent operator is missing an operand".to_owned(),
                        )
                    })?;
                    expr_stack.push(format!("({operand})%"));
                },

                // Concatenation
                AstNodeType::ConcatenationNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "concatenation")?;
                    expr_stack.push(format!("({}&{})", left, right));
                },
                AstNodeType::ColonNode | AstNodeType::ColonNodeWithUids => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "range")?;
                    expr_stack.push(format!("{left}:{right}"));
                },
                AstNodeType::ColonTractNode => {
                    expr_stack.push(render_colon_tract(
                        node,
                        host_row,
                        host_column,
                        formula_references,
                    )?);
                },
                AstNodeType::ReferenceErrorNode | AstNodeType::ReferenceErrorWithUids => {
                    expr_stack.push("#REF!".to_owned());
                },
                AstNodeType::CategoryRefNode => {
                    expr_stack.push(render_category_reference(node, formula_references));
                },
                AstNodeType::UnknownFunctionNode => {
                    let count = node.ast_unknown_function_node_num_args.unwrap_or(0);
                    let arguments =
                        pop_formula_arguments(&mut expr_stack, count, "unknown function")?;
                    let name = node
                        .ast_unknown_function_node_string
                        .as_deref()
                        .unwrap_or("UNKNOWN");
                    expr_stack.push(format!("{name}({})", arguments.join(",")));
                },
                AstNodeType::PlusSignNode
                | AstNodeType::BeginThunkNode
                | AstNodeType::EndThunkNode
                | AstNodeType::AppendWhitespaceNode
                | AstNodeType::PrependWhitespaceNode => {},

                // Other node types - handle gracefully
                _ => {
                    // Unknown or special node types - keep processing
                    // (e.g., whitespace nodes, thunk nodes, etc.)
                },
            }
        }

        // The final result should be on top of the stack
        let result = expr_stack
            .pop()
            .map_or_else(|| "=FORMULA()".to_string(), |value| format!("={value}"));

        Ok(result)
    }

    /// Convert column index to Excel-style letter (0 -> A, 1 -> B, ..., 25 -> Z, 26 -> AA)
    fn column_index_to_letter(index: u32) -> String {
        let mut result = String::new();
        let mut idx = index;

        loop {
            let remainder = idx % 26;
            result.insert(0, (b'A' + remainder as u8) as char);
            if idx < 26 {
                break;
            }
            idx = idx / 26 - 1;
        }

        result
    }

    /// Get function name from function index
    /// Based on Numbers built-in function list
    fn get_function_name(index: u32) -> String {
        super::function_map::function_name(index)
            .map(str::to_owned)
            .unwrap_or_else(|| format!("FUNC{index}"))
    }

    /// Extract rich text from a storage reference
    fn extract_rich_text(&self, storage_id: u64, budget: &mut ProjectionBudget) -> Result<String> {
        if storage_id == 0 {
            return Err(Error::InvalidFormat(
                "Numbers rich-text storage has a null reference".to_owned(),
            ));
        }
        let resolved = self
            .object_index
            .resolve_ref_id(self.bundle, storage_id)?
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers rich-text storage is missing".to_owned())
            })?;
        let mut messages = resolved
            .messages
            .iter()
            .filter(|message| message.type_ == STORAGE_MESSAGE_TYPE);
        let message = messages.next().ok_or_else(|| {
            Error::InvalidFormat("Numbers rich-text storage has no canonical message".to_owned())
        })?;
        if messages.next().is_some() {
            return Err(Error::InvalidFormat(
                "Numbers rich-text storage has duplicate canonical messages".to_owned(),
            ));
        }
        let remaining = if self.document_projection {
            budget.remaining_staging_text_bytes()
        } else {
            budget.remaining_output_text_bytes()
        };
        if remaining == 0 {
            return Err(Error::SemanticLimit {
                kind: SemanticLimitKind::OutputTextBytes,
                observed: 1,
                maximum: 0,
                path: SemanticPath::Package,
            });
        }
        let limits = litchi_iwa_text_wire::Limits::new(
            message.data.len().max(1),
            message.data.len().max(1),
            message.data.len().max(1),
            remaining,
        )
        .map_err(|_error| {
            Error::InvalidFormat("Numbers rich-text limits are invalid".to_owned())
        })?;
        let validated_text_len = if self.document_projection {
            let rewrite_limits = litchi_iwa_text_wire::RewriteLimits::new(
                message.data.len().max(1),
                budget.remaining_payload_fields().max(1),
                8,
                crate::MAX_REFERENCES,
                remaining,
                crate::MAX_REFERENCES,
                budget.remaining_references().max(1),
                message.data.len().max(1),
                budget.remaining_payload_work().max(1),
            )
            .map_err(|_error| {
                Error::InvalidFormat("Numbers rich-text limits are invalid".to_owned())
            })?;
            let validation =
                litchi_iwa_text_wire::validate_storage_with_limits(&message.data, rewrite_limits)
                    .map_err(map_rich_text_rewrite_error)?;
            budget.payload_fields = projection_charge(
                budget.payload_fields,
                validation.fields(),
                crate::MAX_REFERENCES,
                SemanticLimitKind::Objects,
            )?;
            budget.charge_references(validation.reference_occurrences())?;
            let second_pass_work = message
                .data
                .len()
                .checked_mul(2)
                .and_then(|work| work.checked_add(validation.utf8_len().saturating_mul(2)))
                .ok_or_else(|| {
                    formula_semantic_limit(
                        SemanticLimitKind::FormulaWork,
                        usize::MAX,
                        MAX_PAYLOAD_WORK,
                    )
                })?;
            budget.payload_work = projection_charge(
                budget.payload_work,
                validation
                    .validation_work()
                    .saturating_add(second_pass_work),
                MAX_PAYLOAD_WORK,
                SemanticLimitKind::FormulaWork,
            )?;
            Some(validation.utf8_len())
        } else {
            None
        };
        let storage = litchi_iwa_text_wire::from_bytes_with_limits(&message.data, limits)
            .map_err(map_rich_text_error)?;
        if self.document_projection {
            if validated_text_len != Some(storage.len()) {
                return Err(Error::InvalidFormat(
                    "Numbers rich-text storage failed strict text parity".to_owned(),
                ));
            }
            budget.charge_staging_text(storage.len())?;
        } else {
            budget.charge_output_text(storage.len())?;
        }
        Ok(storage.into_text())
    }
}

fn preflight_rich_text_payload(
    source: &[u8],
) -> Result<(u64, litchi_iwa_common::wire::WirePreflight)> {
    let limits = WireLimits::default()
        .with_input_bytes(source.len().clamp(1, WireLimits::MAX_INPUT_BYTES))?
        .with_fields(source.len().clamp(1, WireLimits::MAX_FIELDS))?
        .with_nesting(1)?;
    let mut storage = None;
    let mut cell = false;
    let report = preflight_wire_tree_with_limits(source, limits, |visit| {
        match visit.field().number() {
            1 => {
                if storage.is_some() || visit.field().wire_type() != 2 {
                    return Err(litchi_iwa_common::Error::InvalidFormat(
                        "invalid rich-text storage reference".to_owned(),
                    ));
                }
                visit.field().validate_canonical_framing()?;
                storage = Some(names::preflight_local_reference(visit.field().payload())?);
            },
            3 => {
                if cell || visit.field().wire_type() != 2 {
                    return Err(litchi_iwa_common::Error::InvalidFormat(
                        "invalid rich-text cell owner".to_owned(),
                    ));
                }
                visit.field().validate_canonical_framing()?;
                cell = true;
            },
            _ => {},
        }
        Ok(WireDescent::Skip)
    })?;
    if !cell {
        return Err(Error::InvalidFormat(
            "Numbers rich-text payload has no cell owner".to_owned(),
        ));
    }
    let storage = storage.ok_or_else(|| {
        Error::InvalidFormat("Numbers rich-text payload has no storage reference".to_owned())
    })?;
    Ok((storage, report))
}

fn map_rich_text_error(error: litchi_iwa_text_wire::Error) -> Error {
    match error {
        litchi_iwa_text_wire::Error::TooManyTextBytes { actual, limit } => Error::SemanticLimit {
            kind: SemanticLimitKind::OutputTextBytes,
            observed: actual,
            maximum: limit,
            path: SemanticPath::Package,
        },
        litchi_iwa_text_wire::Error::TooManyFragments { actual, limit } => Error::SemanticLimit {
            kind: SemanticLimitKind::Objects,
            observed: actual,
            maximum: limit,
            path: SemanticPath::Package,
        },
        litchi_iwa_text_wire::Error::Common(error) => Error::Common(error),
        _ => Error::InvalidFormat("Numbers rich-text storage is invalid".to_owned()),
    }
}

fn map_rich_text_rewrite_error(error: litchi_iwa_text_wire::RewriteError) -> Error {
    match error {
        litchi_iwa_text_wire::RewriteError::LimitExceeded {
            resource,
            observed,
            limit,
        } => Error::SemanticLimit {
            kind: if resource.contains("reference") {
                SemanticLimitKind::References
            } else if resource.contains("field") || resource.contains("entry") {
                SemanticLimitKind::Objects
            } else if resource.contains("text") {
                SemanticLimitKind::TextBytes
            } else if resource.contains("work") {
                SemanticLimitKind::FormulaWork
            } else {
                SemanticLimitKind::FormulaWireBytes
            },
            observed,
            maximum: limit,
            path: SemanticPath::Package,
        },
        litchi_iwa_text_wire::RewriteError::Allocation { amount, .. } => {
            allocation_error("Numbers rich-text storage validation", amount)
        },
        _ => Error::InvalidFormat("Numbers rich-text storage is invalid".to_owned()),
    }
}

fn map_comment_error(_error: litchi_iwa_common::comment::Error) -> Error {
    Error::InvalidFormat("Numbers comment metadata is invalid".to_owned())
}

#[derive(Debug)]
struct FormulaReferenceBudget {
    retained_entries: usize,
    maximum_retained_entries: usize,
    work_items: usize,
    wire_bytes: usize,
    text_bytes: usize,
    maximum_work: usize,
    maximum_wire_bytes: usize,
    maximum_text_bytes: usize,
}

impl FormulaReferenceBudget {
    const fn new(
        maximum_retained_entries: usize,
        maximum_work: usize,
        maximum_wire_bytes: usize,
        maximum_text_bytes: usize,
    ) -> Self {
        Self {
            retained_entries: 0,
            maximum_retained_entries,
            work_items: 0,
            wire_bytes: 0,
            text_bytes: 0,
            maximum_work: if maximum_work > MAX_FORMULA_WORK {
                MAX_FORMULA_WORK
            } else {
                maximum_work
            },
            maximum_wire_bytes,
            maximum_text_bytes,
        }
    }

    fn charge_retained_entry(&mut self) -> Result<()> {
        self.retained_entries = self.retained_entries.checked_add(1).ok_or_else(|| {
            formula_semantic_limit(
                SemanticLimitKind::References,
                usize::MAX,
                self.maximum_retained_entries,
            )
        })?;
        if self.retained_entries > self.maximum_retained_entries {
            return Err(formula_semantic_limit(
                SemanticLimitKind::References,
                self.retained_entries,
                self.maximum_retained_entries,
            ));
        }
        Ok(())
    }

    fn charge_work(&mut self, amount: usize) -> Result<()> {
        self.ensure_work_capacity(amount)?;
        self.work_items += amount;
        Ok(())
    }

    /// Retain work that was performed while a candidate was being rejected.
    ///
    /// `charge_work` is intentionally transactional for callers that publish
    /// a successful candidate: it leaves the counter unchanged when the
    /// requested amount crosses the limit.  A formula-category preflight is
    /// different. Its wire walk has already spent the work by the time it
    /// reports a depth/field/resource error, so the attempted amount must be
    /// admitted to the candidate budget even though entries and text from
    /// that candidate are later discarded.
    fn retain_attempted_work(&mut self, amount: usize) {
        self.work_items = self
            .work_items
            .saturating_add(amount)
            .min(self.maximum_work);
    }

    fn ensure_work_capacity(&self, additional: usize) -> Result<()> {
        let observed = self.work_items.checked_add(additional).ok_or_else(|| {
            formula_semantic_limit(
                SemanticLimitKind::FormulaWork,
                usize::MAX,
                self.maximum_work,
            )
        })?;
        if observed > self.maximum_work {
            return Err(formula_semantic_limit(
                SemanticLimitKind::FormulaWork,
                observed,
                self.maximum_work,
            ));
        }
        Ok(())
    }

    fn charge_wire_bytes(&mut self, bytes: usize) -> Result<()> {
        self.charge_work(bytes)?;
        self.wire_bytes = self.wire_bytes.checked_add(bytes).ok_or_else(|| {
            formula_semantic_limit(
                SemanticLimitKind::FormulaWireBytes,
                usize::MAX,
                self.maximum_wire_bytes,
            )
        })?;
        if self.wire_bytes > self.maximum_wire_bytes {
            return Err(formula_semantic_limit(
                SemanticLimitKind::FormulaWireBytes,
                self.wire_bytes,
                self.maximum_wire_bytes,
            ));
        }
        Ok(())
    }

    fn charge_table_info_payload(&mut self, source_len: usize) -> Result<()> {
        // `table_info_codec` charges the source once for each strict/Buffa
        // pass, with a minimum work allowance of one for an empty payload.
        // `charge_wire_bytes` already contributes the first source-length
        // pass to work, so charge the remaining three here.
        self.charge_wire_bytes(source_len)?;
        self.charge_work(checked_formula_work_product(source_len, 3, self.maximum_work)?.max(1))?;
        Ok(())
    }

    fn charge_text(&mut self, bytes: usize) -> Result<()> {
        self.text_bytes = self.text_bytes.checked_add(bytes).ok_or_else(|| {
            formula_semantic_limit(
                SemanticLimitKind::TextBytes,
                usize::MAX,
                self.maximum_text_bytes,
            )
        })?;
        if self.text_bytes > self.maximum_text_bytes {
            return Err(formula_semantic_limit(
                SemanticLimitKind::TextBytes,
                self.text_bytes,
                self.maximum_text_bytes,
            ));
        }
        Ok(())
    }
}

fn formula_semantic_limit(kind: SemanticLimitKind, observed: usize, maximum: usize) -> Error {
    Error::SemanticLimit {
        kind,
        observed,
        maximum,
        path: SemanticPath::StructuredTables,
    }
}

fn build_formula_reference_maps(
    bundle: &Components,
    object_index: &Index,
    budget: &mut FormulaReferenceBudget,
) -> Result<FormulaReferenceMaps> {
    let mut result = FormulaReferenceMaps::default();
    result
        .categories
        .try_reserve(1)
        .map_err(|_error| allocation_error("Numbers formula categories", 1))?;
    let mut grand_total = String::new();
    grand_total
        .try_reserve_exact("Grand Total".len())
        .map_err(|_error| allocation_error("Numbers formula categories", "Grand Total".len()))?;
    grand_total.push_str("Grand Total");
    result.categories.insert([1, 0], grand_total);
    let mut table_info_names = HashMap::<u64, FormulaReferenceName>::new();
    let root_message = bundle
        .get_archive("Index/Document.iwa")
        .and_then(|archive| archive.object(1))
        .and_then(|object| object.messages.iter().find(|message| message.type_ == 1));

    if let Some(root) = root_message {
        budget.charge_wire_bytes(root.data.len())?;
        budget.charge_work(checked_formula_work_product(
            root.data.len(),
            7,
            budget.maximum_work,
        )?)?;
        let options = litchi_iwa_protos::numbers_sheet_order_codec::DecodeOptions::new(
            root.data.len().max(1),
            root.data.len().saturating_mul(2).max(1),
            budget.maximum_work.max(1),
            2,
            budget.maximum_retained_entries.saturating_add(1).max(1),
        );
        let sheet_order =
            litchi_iwa_protos::numbers_sheet_order_codec::decode_document_sheet_order(
                &root.data, options,
            )
            .map_err(|_error| Error::MalformedPayload {
                path: SemanticPath::Document,
            })?;
        for sheet_reference in sheet_order.sheet_references() {
            budget.charge_work(1)?;
            let Some(sheet_object) =
                object_index.resolve_ref_id(bundle, sheet_reference.identifier())?
            else {
                continue;
            };
            let Some(sheet_message) = sheet_object.messages.iter().find(|message| {
                message.type_ == super::SHEET_MESSAGE_TYPE
                    || message.type_ == super::FORM_BASED_SHEET_MESSAGE_TYPE
            }) else {
                continue;
            };
            budget.charge_wire_bytes(sheet_message.data.len())?;
            budget.charge_work(checked_formula_work_product(
                sheet_message.data.len(),
                7,
                budget.maximum_work,
            )?)?;
            let (sheet_name, drawables) = names::preflight_sheet_payload(
                sheet_message.type_,
                &sheet_message.data,
                budget.maximum_retained_entries,
            )
            .map_err(|_error| Error::MalformedPayload {
                path: SemanticPath::StructuredTables,
            })?;
            let mut cached_sheet_name = None::<Arc<String>>;
            for drawable in drawables {
                budget.charge_work(1)?;
                let Some(drawable_object) = object_index.resolve_ref_id(bundle, drawable)? else {
                    continue;
                };
                let table_name =
                    formula_table_name(bundle, object_index, drawable_object.messages, budget)?;
                if let Some(table) = table_name {
                    let is_new = !table_info_names.contains_key(&drawable);
                    if is_new {
                        budget.charge_retained_entry()?;
                    }
                    if is_new {
                        table_info_names.try_reserve(1).map_err(|_error| {
                            allocation_error(
                                "Numbers formula table names",
                                table_info_names.len() + 1,
                            )
                        })?;
                    }
                    let retained_sheet_name = if let Some(name) = &cached_sheet_name {
                        Arc::clone(name)
                    } else {
                        budget.charge_text(sheet_name.len())?;
                        let mut name = String::new();
                        name.try_reserve_exact(sheet_name.len()).map_err(|_error| {
                            allocation_error("Numbers formula sheet name", sheet_name.len())
                        })?;
                        name.push_str(sheet_name);
                        let name = Arc::new(name);
                        cached_sheet_name = Some(Arc::clone(&name));
                        name
                    };
                    table_info_names.insert(
                        drawable,
                        FormulaReferenceName {
                            sheet: retained_sheet_name,
                            table: Arc::new(table),
                        },
                    );
                }
            }
        }
    }

    for (_, archive) in bundle.iter_archives() {
        for object in &archive.objects {
            for message in &object.messages {
                if message.type_ == 6383 {
                    collect_formula_category_payload(
                        message.data.as_slice(),
                        &mut result.categories,
                        budget,
                    )?;
                    continue;
                }
                if message.type_ != 4008 {
                    continue;
                }
                budget.charge_wire_bytes(message.data.len())?;
                // The owner preflight selects two deferred submessages: the
                // nested UUID and the local table reference.  Charge the
                // complete selected-tree cost even when the candidate is
                // malformed; otherwise a stream of rejected owners can make
                // the nested scans effectively free.  The root source bytes
                // are already charged above, so the accumulator starts with
                // that same scan and adds selected descendants/fields.
                let mut owner_cost = 0usize;
                let owner_preflight = preflight_formula_owner(&message.data, &mut owner_cost);
                budget.charge_work(owner_cost)?;
                let (key, table_identifier, _report) = match owner_preflight {
                    Ok(projection) => projection,
                    Err(
                        error @ Error::SemanticLimit {
                            kind: SemanticLimitKind::FormulaWork,
                            ..
                        },
                    ) => return Err(error),
                    Err(_) => continue,
                };
                let Some(name) = table_info_names.get(&table_identifier) else {
                    continue;
                };
                if !result.owners.contains_key(&key) {
                    budget.charge_retained_entry()?;
                    result.owners.try_reserve(1).map_err(|_error| {
                        allocation_error("Numbers formula owners", result.owners.len() + 1)
                    })?;
                }
                result.owners.insert(key, name.clone());
            }
        }
    }
    Ok(result)
}

fn formula_table_name(
    bundle: &Components,
    object_index: &Index,
    messages: &[litchi_iwa_core::RawMessage],
    budget: &mut FormulaReferenceBudget,
) -> Result<Option<String>> {
    // A malformed TableInfo candidate remains best-effort and is skipped for
    // compatibility.  Once two candidates pass the strict model-reference
    // projection, however, selecting whichever happens to occur first would
    // make formula prefixes depend on archive message order.  Count all
    // valid candidates before publishing one name and fail closed on a
    // duplicate.
    let mut valid_candidates = 0usize;
    let mut selected_name = None;
    for table_info_message in messages {
        if table_info_message.type_ != 6_000 && table_info_message.type_ != 6_003 {
            continue;
        }
        let source = table_info_message.data.as_slice();
        // `table_info_codec` bounds its strict and Buffa traversals with
        // `source.len().saturating_mul(4).max(1)` work. The wire charge also
        // contributes one source byte of work, so precharge the remaining
        // three passes before the best-effort decode. Malformed TableInfo
        // remains skippable, but its bounded inspection is never free.
        budget.charge_table_info_payload(source.len())?;
        let Ok(model_reference) = table_info_codec::decode_table_model_reference(
            source,
            table_info_decode_options(source),
        ) else {
            continue;
        };
        valid_candidates = valid_candidates.saturating_add(1);
        if valid_candidates > 1 {
            return Err(Error::InvalidFormat(
                "Numbers formula table has duplicate valid TableInfo candidates".to_owned(),
            ));
        }
        let Some(model_object) =
            object_index.resolve_ref_id(bundle, model_reference.identifier().get())?
        else {
            continue;
        };
        let mut canonical = model_object
            .messages
            .iter()
            .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE);
        let model_message = if let Some(message) = canonical.next() {
            if canonical.next().is_some() {
                return Err(Error::InvalidFormat(
                    "Numbers formula table model has duplicate canonical payloads".to_owned(),
                ));
            }
            Some(message)
        } else {
            let mut legacy = model_object
                .messages
                .iter()
                .filter(|message| message.type_ == 6_000);
            let message = legacy.next();
            if legacy.next().is_some() {
                return Err(Error::InvalidFormat(
                    "Numbers formula table model has duplicate legacy payloads".to_owned(),
                ));
            }
            message
        };
        if let Some(message) = model_message {
            budget.charge_wire_bytes(message.data.len())?;
            budget.charge_work(checked_formula_work_product(
                message.data.len(),
                7,
                budget.maximum_work,
            )?)?;
            let name = names::preflight_table_name(&message.data).map_err(|_error| {
                Error::MalformedPayload {
                    path: SemanticPath::StructuredTables,
                }
            })?;
            budget.charge_text(name.len())?;
            let mut retained = String::new();
            retained
                .try_reserve_exact(name.len())
                .map_err(|_error| allocation_error("Numbers formula table name", name.len()))?;
            retained.push_str(name);
            selected_name = Some(retained);
        }
    }
    Ok(selected_name)
}

fn charge_formula_preflight_work(
    work: &mut usize,
    amount: usize,
    maximum: usize,
) -> litchi_iwa_common::Result<()> {
    *work = work
        .checked_add(amount)
        .ok_or(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Fields,
            observed: usize::MAX,
            limit: maximum,
        })?;
    if *work > maximum {
        return Err(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Fields,
            observed: *work,
            limit: maximum,
        });
    }
    Ok(())
}

fn formula_projection_wire_type(
    field: litchi_iwa_common::wire::WireFieldView<'_>,
    expected: u8,
) -> litchi_iwa_common::Result<()> {
    if field.wire_type() == expected {
        Ok(())
    } else {
        Err(litchi_iwa_common::Error::InvalidFormat(
            "formula category projection field has the wrong wire type".to_owned(),
        ))
    }
}

fn preflight_formula_uuid(
    source: &[u8],
    work: &mut usize,
    maximum_work: usize,
) -> litchi_iwa_common::Result<()> {
    charge_formula_preflight_work(work, 1, maximum_work)?;
    let limits = WireLimits::default()
        .with_input_bytes(source.len().clamp(1, WireLimits::MAX_INPUT_BYTES))?
        .with_fields(maximum_work.clamp(1, WireLimits::MAX_FIELDS))?
        .with_nesting(1)?;
    preflight_wire_tree_with_limits(source, limits, |visit| {
        charge_formula_preflight_work(work, 1, maximum_work)?;
        let field = visit.field();
        if visit.path().is_empty() && matches!(field.number(), 1 | 2) {
            formula_projection_wire_type(field, 0)?;
        }
        Ok(WireDescent::Skip)
    })?;
    Ok(())
}

fn preflight_formula_cell_value(
    source: &[u8],
    work: &mut usize,
    maximum_work: usize,
) -> litchi_iwa_common::Result<()> {
    charge_formula_preflight_work(work, 1, maximum_work)?;
    let input_bytes = source
        .len()
        .checked_mul(2)
        .ok_or(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Fields,
            observed: usize::MAX,
            limit: maximum_work,
        })?
        .clamp(1, WireLimits::MAX_INPUT_BYTES);
    let limits = WireLimits::default()
        .with_input_bytes(input_bytes)?
        .with_fields(maximum_work.clamp(1, WireLimits::MAX_FIELDS))?
        .with_nesting(1)?;
    preflight_wire_tree_with_limits(source, limits, |visit| {
        charge_formula_preflight_work(work, 1, maximum_work)?;
        let field = visit.field();
        if visit.path().is_empty() && matches!(field.number(), 2..=5) {
            formula_projection_wire_type(field, 2)?;
            charge_formula_preflight_work(work, 1, maximum_work)?;
            return Ok(WireDescent::Descend);
        }
        let expected_wire_type = match (visit.path(), field.number()) {
            ([2], 1) => Some(0),
            ([3 | 4], 1) => Some(1),
            ([5], 1) => Some(2),
            _ => None,
        };
        if let Some(expected) = expected_wire_type {
            formula_projection_wire_type(field, expected)?;
        }
        if visit.path() == [5]
            && field.number() == 1
            && std::str::from_utf8(field.payload()).is_err()
        {
            return Err(litchi_iwa_common::Error::InvalidFormat(
                "formula category projection string is not UTF-8".to_owned(),
            ));
        }
        Ok(WireDescent::Skip)
    })?;
    Ok(())
}

fn preflight_formula_category_payload(
    source: &[u8],
    budget: &mut FormulaReferenceBudget,
) -> Result<Option<usize>> {
    // Charge source bytes before inspecting their framing so a package cannot
    // multiply malformed-candidate scan work without consuming a hard budget.
    let work_before_wire = budget.work_items;
    if let Err(error) = budget.charge_wire_bytes(source.len()) {
        // If the wire charge failed at the work boundary, `charge_work` did
        // not mutate its counter. The bytes were nevertheless inspected, so
        // retain that attempted work before returning the semantic/resource
        // error. Wire-limit failures have already charged their work and must
        // not be counted twice.
        if budget.work_items == work_before_wire {
            budget.retain_attempted_work(source.len());
        }
        return Err(error);
    }
    let remaining_work = budget.maximum_work.saturating_sub(budget.work_items);
    if remaining_work == 0 {
        budget.retain_attempted_work(1);
        let observed = budget.work_items.checked_add(1).ok_or_else(|| {
            formula_semantic_limit(SemanticLimitKind::FormulaWork, usize::MAX, MAX_FORMULA_WORK)
        })?;
        return Err(formula_semantic_limit(
            SemanticLimitKind::FormulaWork,
            observed,
            MAX_FORMULA_WORK,
        ));
    }
    let category_passes = MAX_FORMULA_CATEGORY_DEPTH.checked_add(1).ok_or_else(|| {
        formula_semantic_limit(
            SemanticLimitKind::FormulaWork,
            usize::MAX,
            budget.maximum_work,
        )
    })?;
    let input_bytes =
        checked_formula_work_product(source.len(), category_passes, budget.maximum_work)?
            .clamp(1, WireLimits::MAX_INPUT_BYTES);
    let fields = remaining_work.clamp(1, WireLimits::MAX_FIELDS);
    let limits = WireLimits::default()
        .with_input_bytes(input_bytes)?
        .with_fields(fields)?
        .with_nesting(MAX_FORMULA_CATEGORY_DEPTH)?;
    let work_before_projection = budget.work_items;
    let mut group_nodes = 1usize;
    let mut projection_work = 1usize;
    let preflight = preflight_wire_tree_with_limits(source, limits, |visit| {
        let field = visit.field();
        charge_formula_preflight_work(&mut projection_work, 1, remaining_work)?;
        if !visit.path().iter().all(|path_field| *path_field == 3) {
            return Err(litchi_iwa_common::Error::InvalidFormat(
                "formula category topology preflight left the child path".to_owned(),
            ));
        }
        match field.number() {
            1 => {
                formula_projection_wire_type(field, 2)?;
                preflight_formula_uuid(field.payload(), &mut projection_work, remaining_work)?;
                Ok(WireDescent::Skip)
            },
            3 => {
                formula_projection_wire_type(field, 2)?;
                let observed_depth = visit.path().len().checked_add(1).ok_or(
                    litchi_iwa_common::Error::LimitExceeded {
                        kind: LimitKind::Nesting,
                        observed: usize::MAX,
                        limit: MAX_FORMULA_CATEGORY_DEPTH,
                    },
                )?;
                if observed_depth > MAX_FORMULA_CATEGORY_DEPTH {
                    return Err(litchi_iwa_common::Error::LimitExceeded {
                        kind: LimitKind::Nesting,
                        observed: observed_depth,
                        limit: MAX_FORMULA_CATEGORY_DEPTH,
                    });
                }
                group_nodes =
                    group_nodes
                        .checked_add(1)
                        .ok_or(litchi_iwa_common::Error::LimitExceeded {
                            kind: LimitKind::Fields,
                            observed: usize::MAX,
                            limit: MAX_FORMULA_WORK,
                        })?;
                charge_formula_preflight_work(&mut projection_work, 1, remaining_work)?;
                Ok(WireDescent::Descend)
            },
            7 => {
                formula_projection_wire_type(field, 2)?;
                preflight_formula_cell_value(
                    field.payload(),
                    &mut projection_work,
                    remaining_work,
                )?;
                Ok(WireDescent::Skip)
            },
            _ => Ok(WireDescent::Skip),
        }
    });
    match preflight {
        Ok(_report) => {
            budget.charge_work(projection_work)?;
            Ok(Some(group_nodes))
        },
        Err(litchi_iwa_common::Error::InvalidFormat(_)) => {
            budget.charge_work(projection_work)?;
            Ok(None)
        },
        Err(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Nesting,
            observed,
            ..
        }) => {
            // The scanner may reject the next descent before it produces a
            // report. Preserve the work already spent on the visited prefix;
            // only the retained category entries/text are transactional.
            budget.retain_attempted_work(projection_work);
            Err(formula_semantic_limit(
                SemanticLimitKind::FormulaDepth,
                observed,
                MAX_FORMULA_CATEGORY_DEPTH,
            ))
        },
        Err(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Fields,
            observed,
            ..
        }) => {
            // `charge_formula_preflight_work` increments its local counter
            // before reporting an over-bound field. Keep that attempted work
            // in the candidate budget even though no category map is
            // published.
            budget.retain_attempted_work(projection_work);
            let observed = work_before_projection
                .checked_add(observed)
                .ok_or_else(|| {
                    formula_semantic_limit(
                        SemanticLimitKind::FormulaWork,
                        usize::MAX,
                        MAX_FORMULA_WORK,
                    )
                })?;
            Err(formula_semantic_limit(
                SemanticLimitKind::FormulaWork,
                observed,
                MAX_FORMULA_WORK,
            ))
        },
        Err(error) => {
            // Input/resource failures from the bounded wire scanner likewise
            // have consumed the visited prefix. Retain its work before
            // surfacing the error; callers only merge work/wire from a
            // rejected FormulaReferenceBudget, never entries or text.
            budget.retain_attempted_work(projection_work);
            Err(Error::Common(error))
        },
    }
}

fn collect_formula_category_payload(
    source: &[u8],
    names: &mut HashMap<FormulaCategoryKey, String>,
    budget: &mut FormulaReferenceBudget,
) -> Result<()> {
    let Some(expected_nodes) = preflight_formula_category_payload(source, budget)? else {
        return Ok(());
    };
    let decode_options = group_node_category_codec::DecodeOptions::new(
        source.len().max(1),
        u32::try_from(MAX_FORMULA_CATEGORY_DEPTH + 3).unwrap_or(u32::MAX),
    );
    let Ok(group_node) = group_node_category_codec::decode_group_node(source, decode_options)
    else {
        return Ok(());
    };
    collect_formula_category_names_with_budget(&group_node, expected_nodes, names, budget)
}

fn collect_formula_category_names_with_budget(
    root_node: &GroupNodeView<'_>,
    expected_nodes: usize,
    names: &mut HashMap<FormulaCategoryKey, String>,
    budget: &mut FormulaReferenceBudget,
) -> Result<()> {
    let mut visited = 1usize;
    retain_formula_category_name(root_node, names, budget)?;
    let mut pending = Vec::new();
    pending
        .try_reserve(1)
        .map_err(|_error| allocation_error("Numbers formula category traversal", 1))?;
    pending.push(root_node.children());
    while let Some(children) = pending.last_mut() {
        let Some(child_result) = children.next() else {
            pending.pop();
            continue;
        };
        let child_node = child_result.map_err(|_error| {
            Error::InvalidFormat(
                "Numbers formula category projection diverged from its wire preflight".to_owned(),
            )
        })?;
        visited = visited.checked_add(1).ok_or_else(|| {
            formula_semantic_limit(SemanticLimitKind::FormulaWork, usize::MAX, MAX_FORMULA_WORK)
        })?;
        if visited > expected_nodes {
            return Err(Error::InvalidFormat(
                "Numbers formula category projection exceeded its wire preflight".to_owned(),
            ));
        }
        retain_formula_category_name(&child_node, names, budget)?;
        pending.try_reserve(1).map_err(|_error| {
            allocation_error("Numbers formula category traversal", pending.len() + 1)
        })?;
        pending.push(child_node.children());
    }
    if visited != expected_nodes {
        return Err(Error::InvalidFormat(
            "Numbers formula category projection did not reach its preflighted nodes".to_owned(),
        ));
    }
    Ok(())
}

fn retain_formula_category_name(
    node: &GroupNodeView<'_>,
    names: &mut HashMap<FormulaCategoryKey, String>,
    budget: &mut FormulaReferenceBudget,
) -> Result<()> {
    let key = node
        .group_uid()
        .map_err(formula_category_projection_error)?
        .map_or([0, 0], |uid| [uid.lower(), uid.upper()]);
    let Some(value) = node
        .category_value()
        .map_err(formula_category_projection_error)?
    else {
        return Ok(());
    };
    let Some(label) = group_cell_value_label(&value).map_err(formula_category_projection_error)?
    else {
        return Ok(());
    };
    let is_new = !names.contains_key(&key);
    if is_new {
        budget.charge_retained_entry()?;
        names
            .try_reserve(1)
            .map_err(|_error| allocation_error("Numbers formula categories", names.len() + 1))?;
    }
    budget.charge_text(label.len())?;
    let label = match label {
        Cow::Owned(label) => label,
        Cow::Borrowed(label) => {
            let mut retained = String::new();
            retained.try_reserve_exact(label.len()).map_err(|_error| {
                allocation_error("Numbers formula category label", label.len())
            })?;
            retained.push_str(label);
            retained
        },
    };
    names.insert(key, label);
    Ok(())
}

fn group_cell_value_label<'source>(
    value: &CategoryValueView<'source>,
) -> std::result::Result<Option<Cow<'source, str>>, group_node_category_codec::DecodeError> {
    if let Some(string) = value.string()? {
        return Ok(Some(Cow::Borrowed(string)));
    }
    if let Some(number) = value.number()? {
        return Ok(Some(Cow::Owned(number.to_string())));
    }
    if let Some(boolean) = value.boolean()? {
        return Ok(Some(Cow::Borrowed(if boolean { "TRUE" } else { "FALSE" })));
    }
    Ok(value.date()?.map(|date| Cow::Owned(date.to_string())))
}

fn formula_category_projection_error(_error: group_node_category_codec::DecodeError) -> Error {
    Error::InvalidFormat(
        "Numbers formula category projection diverged from its wire preflight".to_owned(),
    )
}

#[cfg(test)]
fn render_category_reference(
    node: &tsce::ast_node_array_archive::AstNodeArchive,
    references: &FormulaReferenceMaps,
) -> String {
    let category_uid = node
        .ast_category_ref
        .as_ref()
        .map(|ast| &ast.category_ref)
        .and_then(|category| {
            category
                .absolute_group_uid
                .as_ref()
                .or(category.relative_group_uid.as_ref())
                .or_else(|| category.group_uids.as_ref()?.uid.last())
        });
    category_uid
        .and_then(|uid| references.categories.get(&formula_category_key(uid)))
        .map(|label| {
            let escaped = label.replace('\\', "\\\\").replace(']', "\\]");
            format!("#CATEGORY![{escaped}]")
        })
        .unwrap_or_else(|| "#CATEGORY!".to_owned())
}

fn find_bundle_object(
    bundle: &Components,
    identifier: u64,
) -> Option<&litchi_iwa_core::ArchiveObject> {
    bundle
        .iter_archives()
        .map(|(_, archive)| archive)
        .find_map(|archive| archive.object(identifier))
}

fn formula_owner_key(owner: &litchi_iwa_protos::tsp::Uuid) -> FormulaOwnerKey {
    [
        owner.lower as u32,
        (owner.lower >> 32) as u32,
        owner.upper as u32,
        (owner.upper >> 32) as u32,
    ]
}

fn checked_formula_owner_work(work: &mut usize, amount: usize) -> litchi_iwa_common::Result<()> {
    *work = work
        .checked_add(amount)
        .ok_or(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::RewriteWork,
            observed: usize::MAX,
            limit: MAX_FORMULA_WORK,
        })?;
    Ok(())
}

fn map_formula_owner_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::RewriteWork,
            observed,
            limit,
        } => formula_semantic_limit(SemanticLimitKind::FormulaWork, observed, limit),
        other => Error::Common(other),
    }
}

fn preflight_formula_owner(
    source: &[u8],
    // Aggregate selected-tree work is reported through this side channel so
    // malformed preflights can retain the cost even though they return no
    // `WirePreflight` report.
    work: &mut usize,
) -> Result<(FormulaOwnerKey, u64, litchi_iwa_common::wire::WirePreflight)> {
    // Keep the root-message scan charge outside the wire preflight result so
    // it survives a parse error.  `WirePreflight` is only returned on
    // success, while malformed owner candidates are intentionally skipped by
    // the compatibility scan below.
    checked_formula_owner_work(work, source.len()).map_err(map_formula_owner_wire_error)?;
    let limits = WireLimits::default()
        .with_input_bytes(source.len().clamp(1, WireLimits::MAX_INPUT_BYTES))?
        .with_fields(source.len().clamp(1, WireLimits::MAX_FIELDS))?
        .with_nesting(2)?;
    let mut owner_key = None;
    let mut table = None;
    let report = preflight_wire_tree_with_limits(source, limits, |visit| {
        // Count every field before validating its schema role.  This keeps
        // the bounded work charge for a malformed field that aborts the
        // preflight, including malformed known fields and malformed trailing
        // descendants reached before the error.
        checked_formula_owner_work(work, 1)?;
        if visit.path().is_empty() && visit.field().number() == 1 {
            if owner_key.is_some() || visit.field().wire_type() != 2 {
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "invalid formula owner".into(),
                ));
            }
            visit.field().validate_canonical_framing()?;
            let mut lower = None;
            let mut upper = None;
            // The nested UUID message is scanned independently from the
            // owner root.  Charge its bounded byte walk before entering it so
            // malformed nested wire cannot discard the cost on error.
            checked_formula_owner_work(work, visit.field().payload().len())?;
            let nested = preflight_wire_tree_with_limits(
                visit.field().payload(),
                WireLimits::default()
                    .with_input_bytes(
                        visit
                            .field()
                            .payload()
                            .len()
                            .clamp(1, WireLimits::MAX_INPUT_BYTES),
                    )?
                    .with_fields(8)?
                    .with_nesting(1)?,
                |uuid| {
                    checked_formula_owner_work(work, 1)?;
                    if !uuid.path().is_empty() || uuid.field().wire_type() != 0 {
                        return Err(litchi_iwa_common::Error::InvalidFormat(
                            "invalid formula owner UUID".into(),
                        ));
                    }
                    uuid.field().validate_canonical_key()?;
                    let (value, length) =
                        litchi_iwa_common::varint::decode_varint_from_bytes(uuid.field().payload())
                            .map_err(|_error| {
                                litchi_iwa_common::Error::InvalidFormat(
                                    "invalid formula owner UUID".into(),
                                )
                            })?;
                    if length != uuid.field().payload().len() {
                        return Err(litchi_iwa_common::Error::InvalidFormat(
                            "invalid formula owner UUID".into(),
                        ));
                    }
                    let canonical_length = if value == 0 {
                        1
                    } else {
                        usize::try_from((64 - value.leading_zeros()).div_ceil(7)).map_err(
                            |_error| {
                                litchi_iwa_common::Error::InvalidFormat(
                                    "invalid formula owner UUID".into(),
                                )
                            },
                        )?
                    };
                    if length != canonical_length {
                        return Err(litchi_iwa_common::Error::InvalidFormat(
                            "invalid formula owner UUID".into(),
                        ));
                    }
                    match uuid.field().number() {
                        1 if lower.replace(value).is_none() => {},
                        2 if upper.replace(value).is_none() => {},
                        _ => {
                            return Err(litchi_iwa_common::Error::InvalidFormat(
                                "invalid formula owner UUID".into(),
                            ));
                        },
                    }
                    Ok(WireDescent::Skip)
                },
            )?;
            let _ = nested;
            let lower = lower.ok_or_else(|| {
                litchi_iwa_common::Error::InvalidFormat("missing formula owner UUID".into())
            })?;
            let upper = upper.ok_or_else(|| {
                litchi_iwa_common::Error::InvalidFormat("missing formula owner UUID".into())
            })?;
            owner_key = Some([
                lower as u32,
                (lower >> 32) as u32,
                upper as u32,
                (upper >> 32) as u32,
            ]);
        } else if visit.path().is_empty() && visit.field().number() == 11 {
            if table.is_some() || visit.field().wire_type() != 2 {
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "invalid formula owner table".into(),
                ));
            }
            visit.field().validate_canonical_framing()?;
            // `names::preflight_local_reference` performs its own bounded
            // wire walk but exposes only the decoded identifier.  Charge the
            // selected payload before invoking it so malformed local
            // references retain their scan cost as well.
            checked_formula_owner_work(work, visit.field().payload().len())?;
            table = Some(names::preflight_local_reference(visit.field().payload())?);
        }
        Ok(WireDescent::Skip)
    })
    .map_err(map_formula_owner_wire_error)?;
    Ok((
        owner_key
            .ok_or_else(|| Error::InvalidFormat("Numbers formula owner has no UUID".to_owned()))?,
        table
            .ok_or_else(|| Error::InvalidFormat("Numbers formula owner has no table".to_owned()))?,
        report,
    ))
}

fn formula_category_key(category: &litchi_iwa_protos::tsp::Uuid) -> FormulaCategoryKey {
    [category.lower, category.upper]
}

#[cfg(test)]
fn cfuuid_key(owner: &litchi_iwa_protos::tsp::CfuuidArchive) -> Option<FormulaOwnerKey> {
    Some([
        owner.uuid_w0?,
        owner.uuid_w1?,
        owner.uuid_w2?,
        owner.uuid_w3?,
    ])
}

#[cfg(test)]
fn formula_reference_prefix(
    owner: &litchi_iwa_protos::tsp::CfuuidArchive,
    references: &FormulaReferenceMaps,
) -> String {
    cfuuid_key(owner)
        .and_then(|key| references.owners.get(&key))
        .map(|name| format!("{}::{}::", name.sheet, name.table))
        .unwrap_or_else(|| "Table::".to_owned())
}

type FormulaExpr = usize;

#[derive(Debug)]
enum FormulaPart {
    Static(&'static str),
    Owned(String),
    Expr(FormulaExpr),
}

#[derive(Debug)]
struct FormulaNode {
    parts: std::ops::Range<usize>,
    rendered_len: usize,
}

#[derive(Debug, Default)]
struct FormulaRenderer {
    nodes: Vec<FormulaNode>,
    parts: Vec<FormulaPart>,
    owned_bytes: usize,
}

impl FormulaRenderer {
    fn check_additional_owned(&self, additional: usize, budget: &ProjectionBudget) -> Result<()> {
        let retained = self
            .owned_bytes
            .checked_add(additional)
            .and_then(|bytes| bytes.checked_add(1))
            .ok_or_else(|| formula_output_limit_error(usize::MAX, budget))?;
        budget.check_output_text(retained)
    }

    fn static_expr(
        &mut self,
        value: &'static str,
        budget: &ProjectionBudget,
    ) -> Result<FormulaExpr> {
        self.fixed([FormulaPart::Static(value)], budget)
    }

    fn owned_expr(&mut self, value: String, budget: &ProjectionBudget) -> Result<FormulaExpr> {
        self.fixed([FormulaPart::Owned(value)], budget)
    }

    fn binary(
        &mut self,
        left: FormulaExpr,
        operator: &'static str,
        right: FormulaExpr,
        wrapped: bool,
        budget: &ProjectionBudget,
    ) -> Result<FormulaExpr> {
        if wrapped {
            self.fixed(
                [
                    FormulaPart::Static("("),
                    FormulaPart::Expr(left),
                    FormulaPart::Static(operator),
                    FormulaPart::Expr(right),
                    FormulaPart::Static(")"),
                ],
                budget,
            )
        } else {
            self.fixed(
                [
                    FormulaPart::Expr(left),
                    FormulaPart::Static(operator),
                    FormulaPart::Expr(right),
                ],
                budget,
            )
        }
    }

    fn unary(
        &mut self,
        prefix: &'static str,
        expression: FormulaExpr,
        suffix: &'static str,
        budget: &ProjectionBudget,
    ) -> Result<FormulaExpr> {
        self.fixed(
            [
                FormulaPart::Static(prefix),
                FormulaPart::Expr(expression),
                FormulaPart::Static(suffix),
            ],
            budget,
        )
    }

    fn comma_joined(
        &mut self,
        function_prefix: Option<String>,
        arguments: Vec<FormulaExpr>,
        open: &'static str,
        close: &'static str,
        budget: &ProjectionBudget,
    ) -> Result<FormulaExpr> {
        let part_count = arguments
            .len()
            .checked_mul(2)
            .and_then(|count| count.checked_add(3))
            .ok_or_else(|| formula_output_limit_error(usize::MAX, budget))?;
        let mut parts = Vec::new();
        parts
            .try_reserve_exact(part_count)
            .map_err(|_error| allocation_error("Numbers formula render parts", part_count))?;
        if let Some(label) = function_prefix {
            parts.push(FormulaPart::Owned(label));
        }
        parts.push(FormulaPart::Static(open));
        for (index, argument) in arguments.into_iter().enumerate() {
            if index != 0 {
                parts.push(FormulaPart::Static(","));
            }
            parts.push(FormulaPart::Expr(argument));
        }
        parts.push(FormulaPart::Static(close));
        self.dynamic(parts, budget)
    }

    fn array(
        &mut self,
        values: Vec<FormulaExpr>,
        columns: usize,
        budget: &ProjectionBudget,
    ) -> Result<FormulaExpr> {
        let part_count = values
            .len()
            .checked_mul(2)
            .and_then(|count| count.checked_add(2))
            .ok_or_else(|| formula_output_limit_error(usize::MAX, budget))?;
        let mut parts = Vec::new();
        parts
            .try_reserve_exact(part_count)
            .map_err(|_error| allocation_error("Numbers formula array parts", part_count))?;
        parts.push(FormulaPart::Static("{"));
        for (index, value) in values.into_iter().enumerate() {
            if index != 0 {
                parts.push(FormulaPart::Static(
                    if columns != 0 && index % columns == 0 {
                        ";"
                    } else {
                        ","
                    },
                ));
            }
            parts.push(FormulaPart::Expr(value));
        }
        parts.push(FormulaPart::Static("}"));
        self.dynamic(parts, budget)
    }

    fn fixed<const N: usize>(
        &mut self,
        parts: [FormulaPart; N],
        budget: &ProjectionBudget,
    ) -> Result<FormulaExpr> {
        let (rendered_len, owned_bytes) = self.measure(&parts, budget)?;
        self.reserve_node(N)?;
        let start = self.parts.len();
        self.parts.extend(parts);
        self.push_node(start, rendered_len, owned_bytes)
    }

    fn dynamic(
        &mut self,
        parts: Vec<FormulaPart>,
        budget: &ProjectionBudget,
    ) -> Result<FormulaExpr> {
        let (rendered_len, owned_bytes) = self.measure(&parts, budget)?;
        self.reserve_node(parts.len())?;
        let start = self.parts.len();
        self.parts.extend(parts);
        self.push_node(start, rendered_len, owned_bytes)
    }

    fn measure(&self, parts: &[FormulaPart], budget: &ProjectionBudget) -> Result<(usize, usize)> {
        let mut rendered_len = 0usize;
        let mut owned_bytes = 0usize;
        for part in parts {
            let part_len = match part {
                FormulaPart::Static(value) => value.len(),
                FormulaPart::Owned(value) => {
                    owned_bytes = owned_bytes
                        .checked_add(value.len())
                        .ok_or_else(|| formula_output_limit_error(usize::MAX, budget))?;
                    value.len()
                },
                FormulaPart::Expr(expression) => {
                    self.nodes
                        .get(*expression)
                        .ok_or_else(|| {
                            Error::ParseError(
                                "Numbers formula renderer contains an invalid expression"
                                    .to_owned(),
                            )
                        })?
                        .rendered_len
                },
            };
            rendered_len = rendered_len
                .checked_add(part_len)
                .ok_or_else(|| formula_output_limit_error(usize::MAX, budget))?;
        }
        let retained_owned = self
            .owned_bytes
            .checked_add(owned_bytes)
            .and_then(|bytes| bytes.checked_add(1))
            .ok_or_else(|| formula_output_limit_error(usize::MAX, budget))?;
        budget.check_output_text(retained_owned)?;
        let output_len = rendered_len
            .checked_add(1)
            .ok_or_else(|| formula_output_limit_error(usize::MAX, budget))?;
        budget.check_output_text(output_len)?;
        Ok((rendered_len, owned_bytes))
    }

    fn reserve_node(&mut self, part_count: usize) -> Result<()> {
        self.nodes.try_reserve_exact(1).map_err(|_error| {
            allocation_error(
                "Numbers formula render nodes",
                self.nodes.len().saturating_add(1),
            )
        })?;
        self.parts.try_reserve_exact(part_count).map_err(|_error| {
            allocation_error(
                "Numbers formula render parts",
                self.parts.len().saturating_add(part_count),
            )
        })?;
        Ok(())
    }

    fn push_node(
        &mut self,
        start: usize,
        rendered_len: usize,
        owned_bytes: usize,
    ) -> Result<FormulaExpr> {
        self.owned_bytes = self
            .owned_bytes
            .checked_add(owned_bytes)
            .ok_or_else(|| allocation_error("Numbers formula owned text", usize::MAX))?;
        let end = self.parts.len();
        let expression = self.nodes.len();
        self.nodes.push(FormulaNode {
            parts: start..end,
            rendered_len,
        });
        Ok(expression)
    }

    fn render(&self, expression: FormulaExpr, budget: &mut ProjectionBudget) -> Result<String> {
        let node = self.nodes.get(expression).ok_or_else(|| {
            Error::ParseError("Numbers formula has no renderable expression".to_owned())
        })?;
        let output_len = node
            .rendered_len
            .checked_add(1)
            .ok_or_else(|| formula_output_limit_error(usize::MAX, budget))?;
        budget.charge_output_text(output_len)?;

        let mut output = String::new();
        output
            .try_reserve_exact(output_len)
            .map_err(|_error| allocation_error("Numbers rendered formula", output_len))?;
        output.push('=');

        let mut pending = Vec::new();
        self.push_parts_reversed(&mut pending, node.parts.clone())?;
        while let Some(part) = pending.pop() {
            match part {
                FormulaPart::Static(value) => output.push_str(value),
                FormulaPart::Owned(value) => output.push_str(value),
                FormulaPart::Expr(child) => {
                    let child_node = self.nodes.get(*child).ok_or_else(|| {
                        Error::ParseError(
                            "Numbers formula renderer contains an invalid child".to_owned(),
                        )
                    })?;
                    self.push_parts_reversed(&mut pending, child_node.parts.clone())?;
                },
            }
        }
        debug_assert_eq!(output.len(), output_len);
        Ok(output)
    }

    fn push_parts_reversed<'a>(
        &'a self,
        pending: &mut Vec<&'a FormulaPart>,
        range: std::ops::Range<usize>,
    ) -> Result<()> {
        let count = range.len();
        pending.try_reserve_exact(count).map_err(|_error| {
            allocation_error(
                "Numbers formula render stack",
                pending.len().saturating_add(count),
            )
        })?;
        pending.extend(self.parts[range].iter().rev());
        Ok(())
    }
}

fn formula_output_limit_error(observed: usize, budget: &ProjectionBudget) -> Error {
    Error::SemanticLimit {
        kind: SemanticLimitKind::OutputTextBytes,
        observed,
        maximum: budget.max_output_text_bytes,
        path: SemanticPath::StructuredTables,
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

    let options = numbers_formula_codec::DecodeOptions::new(
        formula.bytes.len(),
        crate::MAX_REFERENCES,
        MAX_PAYLOAD_WORK,
        32,
        formula.scalar_visitor_node_count,
        DEFAULT_MAX_TEXT_BYTES,
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
            return Ok(None);
        },
    };
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
    let options = numbers_formula_codec::DecodeOptions::new(
        formula.bytes.len(),
        crate::MAX_REFERENCES.saturating_mul(2),
        MAX_PAYLOAD_WORK.saturating_mul(2),
        raw_depth,
        formula.scalar_visitor_node_count,
        DEFAULT_MAX_TEXT_BYTES.saturating_mul(2),
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
    if let Some(error) = visitor.error.take() {
        return Err(error);
    }
    debug_assert_eq!(report.node_count(), formula.scalar_visitor_node_count);
    visitor.finish()
}

fn map_formula_render_decode_error(error: numbers_formula_codec::DecodeError) -> Error {
    use numbers_formula_codec::DecodeLimit;
    match error.resource_limit() {
        Some(DecodeLimit::Bytes { observed, maximum }) => Error::SemanticLimit {
            kind: SemanticLimitKind::FormulaWireBytes,
            observed,
            maximum,
            path: SemanticPath::StructuredTables,
        },
        Some(DecodeLimit::Fields { observed, maximum }) => Error::SemanticLimit {
            kind: SemanticLimitKind::Objects,
            observed,
            maximum,
            path: SemanticPath::StructuredTables,
        },
        Some(DecodeLimit::Work { observed, maximum }) => Error::SemanticLimit {
            kind: SemanticLimitKind::FormulaWork,
            observed,
            maximum,
            path: SemanticPath::StructuredTables,
        },
        Some(DecodeLimit::Nesting { observed, maximum }) => Error::SemanticLimit {
            kind: SemanticLimitKind::FormulaRenderDepth,
            observed: observed as usize,
            maximum: maximum as usize,
            path: SemanticPath::StructuredTables,
        },
        Some(DecodeLimit::Nodes { observed, maximum }) => Error::SemanticLimit {
            kind: SemanticLimitKind::FormulaRenderWork,
            observed,
            maximum,
            path: SemanticPath::StructuredTables,
        },
        Some(DecodeLimit::Text { observed, maximum }) => Error::SemanticLimit {
            kind: SemanticLimitKind::TextBytes,
            observed,
            maximum,
            path: SemanticPath::StructuredTables,
        },
        Some(DecodeLimit::Allocation { requested }) => {
            allocation_error("Numbers formula compatibility traversal", requested)
        },
        None => Error::MalformedPayload {
            path: SemanticPath::StructuredTables,
        },
    }
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

#[cfg(test)]
fn render_formula(
    formula: &tsce::FormulaArchive,
    host_row: usize,
    host_column: usize,
    formula_references: &FormulaReferenceMaps,
    budget: &mut ProjectionBudget,
) -> Result<String> {
    render_formula_with_node_charge(
        formula,
        host_row,
        host_column,
        formula_references,
        budget,
        false,
    )
}

#[cfg(test)]
fn render_formula_precharged(
    formula: &tsce::FormulaArchive,
    host_row: usize,
    host_column: usize,
    formula_references: &FormulaReferenceMaps,
    budget: &mut ProjectionBudget,
) -> Result<String> {
    render_formula_with_node_charge(
        formula,
        host_row,
        host_column,
        formula_references,
        budget,
        true,
    )
}

#[cfg(test)]
fn render_formula_with_node_charge(
    formula: &tsce::FormulaArchive,
    host_row: usize,
    host_column: usize,
    formula_references: &FormulaReferenceMaps,
    budget: &mut ProjectionBudget,
    nodes_precharged: bool,
) -> Result<String> {
    let ast = &formula.ast_node_array;
    if ast.ast_node.is_empty() {
        return retain_text("=", budget);
    }

    let mut renderer = FormulaRenderer::default();
    let root = match render_formula_ast_array(
        ast,
        host_row,
        host_column,
        formula_references,
        budget,
        &mut renderer,
        1,
        nodes_precharged,
    )? {
        Some(root) => root,
        None => renderer.static_expr("FORMULA()", budget)?,
    };
    renderer.render(root, budget)
}

#[allow(
    clippy::too_many_lines,
    reason = "the exhaustive AST match preserves native node semantics"
)]
#[cfg(test)]
fn render_formula_ast_array(
    ast: &tsce::AstNodeArrayArchive,
    host_row: usize,
    host_column: usize,
    formula_references: &FormulaReferenceMaps,
    budget: &mut ProjectionBudget,
    renderer: &mut FormulaRenderer,
    depth: usize,
    nodes_precharged: bool,
) -> Result<Option<FormulaExpr>> {
    use litchi_iwa_protos::tsce::ast_node_array_archive::AstNodeType;

    budget.check_formula_render_depth(depth)?;
    if !nodes_precharged {
        budget.charge_formula_render_work(ast.ast_node.len())?;
    }
    let mut stack = Vec::new();
    stack
        .try_reserve_exact(ast.ast_node.len())
        .map_err(|_error| {
            allocation_error("Numbers formula expression stack", ast.ast_node.len())
        })?;

    for node in &ast.ast_node {
        let expression = match node.ast_node_type() {
            AstNodeType::AdditionNode => Some(render_binary(
                &mut stack, renderer, "+", "addition", true, budget,
            )?),
            AstNodeType::SubtractionNode => Some(render_binary(
                &mut stack,
                renderer,
                "-",
                "subtraction",
                true,
                budget,
            )?),
            AstNodeType::MultiplicationNode => Some(render_binary(
                &mut stack,
                renderer,
                "*",
                "multiplication",
                true,
                budget,
            )?),
            AstNodeType::DivisionNode => Some(render_binary(
                &mut stack, renderer, "/", "division", true, budget,
            )?),
            AstNodeType::PowerNode => Some(render_binary(
                &mut stack, renderer, "^", "power", true, budget,
            )?),
            AstNodeType::GreaterThanNode => Some(render_binary(
                &mut stack,
                renderer,
                ">",
                "greater than",
                true,
                budget,
            )?),
            AstNodeType::GreaterThanOrEqualToNode => Some(render_binary(
                &mut stack,
                renderer,
                ">=",
                "greater than or equal",
                true,
                budget,
            )?),
            AstNodeType::LessThanNode => Some(render_binary(
                &mut stack,
                renderer,
                "<",
                "less than",
                true,
                budget,
            )?),
            AstNodeType::LessThanOrEqualToNode => Some(render_binary(
                &mut stack,
                renderer,
                "<=",
                "less than or equal",
                true,
                budget,
            )?),
            AstNodeType::EqualToNode => Some(render_binary(
                &mut stack, renderer, "=", "equality", true, budget,
            )?),
            AstNodeType::NotEqualToNode => Some(render_binary(
                &mut stack,
                renderer,
                "<>",
                "inequality",
                true,
                budget,
            )?),
            AstNodeType::NumberNode => node
                .ast_number_node_number
                .map(|number| {
                    let value = fallible_formula_display(number, renderer, budget)?;
                    renderer.owned_expr(value, budget)
                })
                .transpose()?,
            AstNodeType::StringNode => node
                .ast_string_node_string
                .as_deref()
                .map(|value| formula_string_literal(value, renderer, budget))
                .transpose()?
                .map(|value| renderer.owned_expr(value, budget))
                .transpose()?,
            AstNodeType::BooleanNode => node
                .ast_boolean_node_boolean
                .map(|value| renderer.static_expr(if value { "TRUE" } else { "FALSE" }, budget))
                .transpose()?,
            AstNodeType::TokenNode => node
                .ast_token_node_boolean
                .map(|value| renderer.static_expr(if value { "TRUE" } else { "FALSE" }, budget))
                .transpose()?,
            AstNodeType::DateNode => node
                .ast_date_node_date_num
                .map(|seconds| {
                    let days = seconds / 86_400.0;
                    let value = fallible_formula_format(renderer, budget, |output| {
                        write!(output, "(DATE(2001,1,1)+{days})")
                    })?;
                    renderer.owned_expr(value, budget)
                })
                .transpose()?,
            AstNodeType::DurationNode => node
                .ast_duration_node_unit_num
                .map(|value| {
                    let value = fallible_formula_display(value, renderer, budget)?;
                    renderer.owned_expr(value, budget)
                })
                .transpose()?,
            AstNodeType::EmptyArgumentNode => Some(renderer.static_expr("", budget)?),
            AstNodeType::CellReferenceNode => Some(renderer.owned_expr(
                render_cell_reference_checked(
                    node,
                    host_row,
                    host_column,
                    formula_references,
                    renderer,
                    budget,
                )?,
                budget,
            )?),
            AstNodeType::LocalCellReferenceNode => Some(
                renderer.owned_expr(
                    node.ast_local_cell_reference_node_reference
                        .as_ref()
                        .map_or_else(
                            || fallible_formula_owned("#REF!", renderer, budget),
                            |cell| {
                                let column = FormulaColumn(cell.column_handle);
                                let row = checked_formula_row_number(cell.row_handle)?;
                                fallible_formula_format(renderer, budget, |output| {
                                    write!(output, "{column}{row}")
                                })
                            },
                        )?,
                    budget,
                )?,
            ),
            AstNodeType::CrossTableCellReferenceNode => Some(
                renderer.owned_expr(
                    node.ast_cross_table_cell_reference_node_reference
                        .as_ref()
                        .map_or_else(
                            || fallible_formula_owned("#REF!", renderer, budget),
                            |cell| {
                                let prefix = formula_reference_prefix_parts(
                                    &cell.table_id,
                                    formula_references,
                                );
                                let column = FormulaColumn(cell.column_handle);
                                let row = checked_formula_row_number(cell.row_handle)?;
                                fallible_formula_format(renderer, budget, |output| {
                                    write_formula_reference_prefix(output, prefix)?;
                                    write!(output, "{column}{row}")
                                })
                            },
                        )?,
                    budget,
                )?,
            ),
            AstNodeType::FunctionNode => {
                if let Some(index) = node.ast_function_node_index {
                    let arguments = pop_formula_arguments(
                        &mut stack,
                        node.ast_function_node_num_args.unwrap_or(0),
                        "function",
                    )?;
                    Some(renderer.comma_joined(
                        Some(fallible_function_name(index, renderer, budget)?),
                        arguments,
                        "(",
                        ")",
                        budget,
                    )?)
                } else {
                    None
                }
            },
            AstNodeType::ListNode => {
                if let Some(count) = node.ast_list_node_num_args {
                    let arguments = pop_formula_arguments(&mut stack, count, "list")?;
                    Some(renderer.comma_joined(None, arguments, "", "", budget)?)
                } else {
                    None
                }
            },
            AstNodeType::ArrayNode => {
                let column_count = node.ast_array_node_num_col.unwrap_or(0);
                let rows = node.ast_array_node_num_row.unwrap_or(0);
                let count = column_count.checked_mul(rows).ok_or_else(|| {
                    Error::ParseError("Numbers formula array size overflow".to_owned())
                })?;
                let values = pop_formula_arguments(&mut stack, count, "array")?;
                let column_count_usize = usize::try_from(column_count).map_err(|_error| {
                    Error::ParseError("Numbers formula array width exceeds usize".to_owned())
                })?;
                Some(renderer.array(values, column_count_usize, budget)?)
            },
            AstNodeType::ThunkNode => {
                if let Some(nested) = &node.ast_thunk_node_array {
                    let nested_expression = match render_formula_ast_array(
                        nested,
                        host_row,
                        host_column,
                        formula_references,
                        budget,
                        renderer,
                        depth.checked_add(1).ok_or(Error::SemanticLimit {
                            kind: SemanticLimitKind::FormulaRenderDepth,
                            observed: usize::MAX,
                            maximum: budget.max_formula_render_depth,
                            path: SemanticPath::StructuredTables,
                        })?,
                        nodes_precharged,
                    )? {
                        Some(expression) => expression,
                        None => renderer.static_expr(
                            if nested.ast_node.is_empty() {
                                ""
                            } else {
                                "FORMULA()"
                            },
                            budget,
                        )?,
                    };
                    Some(nested_expression)
                } else {
                    None
                }
            },
            AstNodeType::NegationNode => stack
                .pop()
                .map(|operand| renderer.unary("-(", operand, ")", budget))
                .transpose()?,
            AstNodeType::PercentNode => {
                let operand = stack.pop().ok_or_else(|| {
                    Error::ParseError(
                        "Numbers formula percent operator is missing an operand".to_owned(),
                    )
                })?;
                Some(renderer.unary("(", operand, ")%", budget)?)
            },
            AstNodeType::ConcatenationNode => Some(render_binary(
                &mut stack,
                renderer,
                "&",
                "concatenation",
                true,
                budget,
            )?),
            AstNodeType::ColonNode | AstNodeType::ColonNodeWithUids => Some(render_binary(
                &mut stack, renderer, ":", "range", false, budget,
            )?),
            AstNodeType::ColonTractNode => Some(renderer.owned_expr(
                render_colon_tract_checked(
                    node,
                    host_row,
                    host_column,
                    formula_references,
                    renderer,
                    budget,
                )?,
                budget,
            )?),
            AstNodeType::ReferenceErrorNode | AstNodeType::ReferenceErrorWithUids => {
                Some(renderer.static_expr("#REF!", budget)?)
            },
            AstNodeType::CategoryRefNode => Some(renderer.owned_expr(
                render_category_reference_checked(node, formula_references, renderer, budget)?,
                budget,
            )?),
            AstNodeType::UnknownFunctionNode => {
                let arguments = pop_formula_arguments(
                    &mut stack,
                    node.ast_unknown_function_node_num_args.unwrap_or(0),
                    "unknown function",
                )?;
                Some(
                    renderer.comma_joined(
                        Some(fallible_formula_owned(
                            node.ast_unknown_function_node_string
                                .as_deref()
                                .unwrap_or("UNKNOWN"),
                            renderer,
                            budget,
                        )?),
                        arguments,
                        "(",
                        ")",
                        budget,
                    )?,
                )
            },
            AstNodeType::PlusSignNode
            | AstNodeType::BeginThunkNode
            | AstNodeType::EndThunkNode
            | AstNodeType::AppendWhitespaceNode
            | AstNodeType::PrependWhitespaceNode
            | AstNodeType::UidReferenceNode
            | AstNodeType::LetBindNode
            | AstNodeType::VarNode
            | AstNodeType::EndScopeNode
            | AstNodeType::LambdaNode
            | AstNodeType::BeginLambdaThunkNode
            | AstNodeType::EndLambdaThunkNode
            | AstNodeType::LinkedCellRefNode
            | AstNodeType::LinkedColumnRefNode
            | AstNodeType::LinkedRowRefNode
            | AstNodeType::ViewTractRefNode
            | AstNodeType::IntersectionNode
            | AstNodeType::SpillRangeNode => None,
        };
        if let Some(rendered_expression) = expression {
            stack.push(rendered_expression);
        }
    }
    Ok(stack.pop())
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

#[cfg(test)]
fn formula_reference_prefix_parts<'a>(
    owner: &litchi_iwa_protos::tsp::CfuuidArchive,
    references: &'a FormulaReferenceMaps,
) -> FormulaPrefix<'a> {
    cfuuid_key(owner).and_then(|key| references.owners.get(&key))
}

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

#[cfg(test)]
fn render_category_reference_checked(
    node: &tsce::ast_node_array_archive::AstNodeArchive,
    references: &FormulaReferenceMaps,
    renderer: &FormulaRenderer,
    budget: &ProjectionBudget,
) -> Result<String> {
    let category_uid = node
        .ast_category_ref
        .as_ref()
        .map(|ast| &ast.category_ref)
        .and_then(|category| {
            category
                .absolute_group_uid
                .as_ref()
                .or(category.relative_group_uid.as_ref())
                .or_else(|| category.group_uids.as_ref()?.uid.last())
        });
    let Some(label) =
        category_uid.and_then(|uid| references.categories.get(&formula_category_key(uid)))
    else {
        return fallible_formula_owned("#CATEGORY!", renderer, budget);
    };
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

#[cfg(test)]
fn render_cell_reference(
    node: &tsce::ast_node_array_archive::AstNodeArchive,
    host_row: usize,
    host_column: usize,
    formula_references: &FormulaReferenceMaps,
) -> Result<String> {
    if let (Some(ast_column), Some(ast_row)) = (&node.ast_column, &node.ast_row) {
        let column = resolve_formula_coordinate(
            host_column,
            ast_column.column,
            ast_column.absolute.unwrap_or(false),
            "column",
        )?;
        let row = resolve_formula_coordinate(
            host_row,
            ast_row.row,
            ast_row.absolute.unwrap_or(false),
            "row",
        )?;
        let prefix = node
            .ast_cross_table_reference_extra_info
            .as_ref()
            .map(|extra| formula_reference_prefix(&extra.table_id, formula_references))
            .unwrap_or_default();
        return Ok(format!(
            "{prefix}{}{}{}{}",
            if ast_column.absolute.unwrap_or(false) {
                "$"
            } else {
                ""
            },
            TableDataExtractor::column_index_to_letter(column),
            if ast_row.absolute.unwrap_or(false) {
                "$"
            } else {
                ""
            },
            checked_formula_row_number(row)?
        ));
    }
    if let Some(cell) = &node.ast_local_cell_reference_node_reference {
        return Ok(format!(
            "{}{}{}{}",
            if cell.column_is_sticky != 0 { "$" } else { "" },
            TableDataExtractor::column_index_to_letter(cell.column_handle),
            if cell.row_is_sticky != 0 { "$" } else { "" },
            checked_formula_row_number(cell.row_handle)?
        ));
    }
    if let Some(cell) = &node.ast_cross_table_cell_reference_node_reference {
        return Ok(format!(
            "{}{}{}",
            formula_reference_prefix(&cell.table_id, formula_references),
            TableDataExtractor::column_index_to_letter(cell.column_handle),
            checked_formula_row_number(cell.row_handle)?
        ));
    }
    Ok("#REF!".to_owned())
}

#[cfg(test)]
fn render_cell_reference_checked(
    node: &tsce::ast_node_array_archive::AstNodeArchive,
    host_row: usize,
    host_column: usize,
    formula_references: &FormulaReferenceMaps,
    renderer: &FormulaRenderer,
    budget: &ProjectionBudget,
) -> Result<String> {
    if let (Some(ast_column), Some(ast_row)) = (&node.ast_column, &node.ast_row) {
        let column = FormulaColumn(resolve_formula_coordinate(
            host_column,
            ast_column.column,
            ast_column.absolute.unwrap_or(false),
            "column",
        )?);
        let row = resolve_formula_coordinate(
            host_row,
            ast_row.row,
            ast_row.absolute.unwrap_or(false),
            "row",
        )?;
        let row = checked_formula_row_number(row)?;
        let prefix = node
            .ast_cross_table_reference_extra_info
            .as_ref()
            .and_then(|extra| formula_reference_prefix_parts(&extra.table_id, formula_references));
        return fallible_formula_format(renderer, budget, |output| {
            if node.ast_cross_table_reference_extra_info.is_some() {
                write_formula_reference_prefix(output, prefix)?;
            }
            write!(
                output,
                "{}{column}{}{row}",
                if ast_column.absolute.unwrap_or(false) {
                    "$"
                } else {
                    ""
                },
                if ast_row.absolute.unwrap_or(false) {
                    "$"
                } else {
                    ""
                }
            )
        });
    }
    if let Some(cell) = &node.ast_local_cell_reference_node_reference {
        let column = FormulaColumn(cell.column_handle);
        let row = checked_formula_row_number(cell.row_handle)?;
        return fallible_formula_format(renderer, budget, |output| {
            write!(
                output,
                "{}{column}{}{row}",
                if cell.column_is_sticky != 0 { "$" } else { "" },
                if cell.row_is_sticky != 0 { "$" } else { "" }
            )
        });
    }
    if let Some(cell) = &node.ast_cross_table_cell_reference_node_reference {
        let prefix = formula_reference_prefix_parts(&cell.table_id, formula_references);
        let column = FormulaColumn(cell.column_handle);
        let row = checked_formula_row_number(cell.row_handle)?;
        return fallible_formula_format(renderer, budget, |output| {
            write_formula_reference_prefix(output, prefix)?;
            write!(output, "{column}{row}")
        });
    }
    fallible_formula_owned("#REF!", renderer, budget)
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

#[cfg(test)]
fn render_colon_tract(
    node: &tsce::ast_node_array_archive::AstNodeArchive,
    host_row: usize,
    host_column: usize,
    formula_references: &FormulaReferenceMaps,
) -> Result<String> {
    let tract = node.ast_colon_tract.as_ref().ok_or_else(|| {
        Error::ParseError("Numbers formula colon tract is missing its coordinates".to_owned())
    })?;
    let sticky = node.ast_sticky_bits.as_ref().ok_or_else(|| {
        Error::ParseError("Numbers formula colon tract is missing its sticky bits".to_owned())
    })?;
    let prefix = node
        .ast_cross_table_reference_extra_info
        .as_ref()
        .map(|extra| formula_reference_prefix(&extra.table_id, formula_references))
        .unwrap_or_default();
    // Numbers uses these maximum-handle sentinels for the unbounded axis of
    // whole-row and whole-column references (for example `1:2` and `B:C`).
    let whole_rows = tract.relative_column.is_empty()
        && tract.absolute_column.len() == 1
        && tract.absolute_column[0].range_begin == i16::MAX as u32
        && tract.absolute_column[0].range_end.is_none();
    let whole_columns = tract.relative_row.is_empty()
        && tract.absolute_row.len() == 1
        && tract.absolute_row[0].range_begin == i32::MAX as u32
        && tract.absolute_row[0].range_end.is_none();
    let has_columns =
        !whole_rows && (!tract.relative_column.is_empty() || !tract.absolute_column.is_empty());
    let has_rows =
        !whole_columns && (!tract.relative_row.is_empty() || !tract.absolute_row.is_empty());
    match (has_columns, has_rows) {
        (true, true) => {
            let (begin_column, end_column) = resolve_colon_axis(
                &tract.relative_column,
                &tract.absolute_column,
                sticky.begin_column_is_absolute,
                sticky.end_column_is_absolute,
                host_column,
                "column",
            )?;
            let (begin_row, end_row) = resolve_colon_axis(
                &tract.relative_row,
                &tract.absolute_row,
                sticky.begin_row_is_absolute,
                sticky.end_row_is_absolute,
                host_row,
                "row",
            )?;
            Ok(format!(
                "{prefix}{}{}{}{}:{}{}{}{}",
                if sticky.begin_column_is_absolute {
                    "$"
                } else {
                    ""
                },
                TableDataExtractor::column_index_to_letter(begin_column),
                if sticky.begin_row_is_absolute {
                    "$"
                } else {
                    ""
                },
                checked_formula_row_number(begin_row)?,
                if sticky.end_column_is_absolute {
                    "$"
                } else {
                    ""
                },
                TableDataExtractor::column_index_to_letter(end_column),
                if sticky.end_row_is_absolute { "$" } else { "" },
                checked_formula_row_number(end_row)?,
            ))
        },
        (false, true) => {
            let (begin, end) = resolve_colon_axis(
                &tract.relative_row,
                &tract.absolute_row,
                sticky.begin_row_is_absolute,
                sticky.end_row_is_absolute,
                host_row,
                "row",
            )?;
            Ok(format!(
                "{prefix}{}{}:{}{}",
                if sticky.begin_row_is_absolute {
                    "$"
                } else {
                    ""
                },
                checked_formula_row_number(begin)?,
                if sticky.end_row_is_absolute { "$" } else { "" },
                checked_formula_row_number(end)?,
            ))
        },
        (true, false) => {
            let (begin, end) = resolve_colon_axis(
                &tract.relative_column,
                &tract.absolute_column,
                sticky.begin_column_is_absolute,
                sticky.end_column_is_absolute,
                host_column,
                "column",
            )?;
            Ok(format!(
                "{prefix}{}{}:{}{}",
                if sticky.begin_column_is_absolute {
                    "$"
                } else {
                    ""
                },
                TableDataExtractor::column_index_to_letter(begin),
                if sticky.end_column_is_absolute {
                    "$"
                } else {
                    ""
                },
                TableDataExtractor::column_index_to_letter(end),
            ))
        },
        (false, false) => Err(Error::ParseError(
            "Numbers formula colon tract has no row or column coordinates".to_owned(),
        )),
    }
}

#[cfg(test)]
fn render_colon_tract_checked(
    node: &tsce::ast_node_array_archive::AstNodeArchive,
    host_row: usize,
    host_column: usize,
    formula_references: &FormulaReferenceMaps,
    renderer: &FormulaRenderer,
    budget: &ProjectionBudget,
) -> Result<String> {
    let tract = node.ast_colon_tract.as_ref().ok_or_else(|| {
        Error::ParseError("Numbers formula colon tract is missing its coordinates".to_owned())
    })?;
    let sticky = node.ast_sticky_bits.as_ref().ok_or_else(|| {
        Error::ParseError("Numbers formula colon tract is missing its sticky bits".to_owned())
    })?;
    let prefix = node
        .ast_cross_table_reference_extra_info
        .as_ref()
        .and_then(|extra| formula_reference_prefix_parts(&extra.table_id, formula_references));
    let has_prefix = node.ast_cross_table_reference_extra_info.is_some();
    let whole_rows = tract.relative_column.is_empty()
        && tract.absolute_column.len() == 1
        && tract.absolute_column[0].range_begin == i16::MAX as u32
        && tract.absolute_column[0].range_end.is_none();
    let whole_columns = tract.relative_row.is_empty()
        && tract.absolute_row.len() == 1
        && tract.absolute_row[0].range_begin == i32::MAX as u32
        && tract.absolute_row[0].range_end.is_none();
    let has_columns =
        !whole_rows && (!tract.relative_column.is_empty() || !tract.absolute_column.is_empty());
    let has_rows =
        !whole_columns && (!tract.relative_row.is_empty() || !tract.absolute_row.is_empty());
    let (begin_column, end_column) = if has_columns {
        let (begin, end) = resolve_colon_axis(
            &tract.relative_column,
            &tract.absolute_column,
            sticky.begin_column_is_absolute,
            sticky.end_column_is_absolute,
            host_column,
            "column",
        )?;
        (Some(FormulaColumn(begin)), Some(FormulaColumn(end)))
    } else {
        (None, None)
    };
    let (begin_row, end_row) = if has_rows {
        let (begin, end) = resolve_colon_axis(
            &tract.relative_row,
            &tract.absolute_row,
            sticky.begin_row_is_absolute,
            sticky.end_row_is_absolute,
            host_row,
            "row",
        )?;
        (
            Some(checked_formula_row_number(begin)?),
            Some(checked_formula_row_number(end)?),
        )
    } else {
        (None, None)
    };
    if !has_columns && !has_rows {
        return Err(Error::ParseError(
            "Numbers formula colon tract has no row or column coordinates".to_owned(),
        ));
    }
    fallible_formula_format(renderer, budget, |output| {
        if has_prefix {
            write_formula_reference_prefix(output, prefix)?;
        }
        if let Some(column) = begin_column {
            if sticky.begin_column_is_absolute {
                output.write_char('$')?;
            }
            write!(output, "{column}")?;
        }
        if let Some(row) = begin_row {
            if sticky.begin_row_is_absolute {
                output.write_char('$')?;
            }
            write!(output, "{row}")?;
        }
        output.write_char(':')?;
        if let Some(column) = end_column {
            if sticky.end_column_is_absolute {
                output.write_char('$')?;
            }
            write!(output, "{column}")?;
        }
        if let Some(row) = end_row {
            if sticky.end_row_is_absolute {
                output.write_char('$')?;
            }
            write!(output, "{row}")?;
        }
        Ok(())
    })
}

#[cfg(test)]
fn resolve_colon_axis(
    relative: &[tsce::ast_node_array_archive::ast_colon_tract_archive::AstColonTractRelativeRangeArchive],
    absolute: &[tsce::ast_node_array_archive::ast_colon_tract_archive::AstColonTractAbsoluteRangeArchive],
    begin_is_absolute: bool,
    end_is_absolute: bool,
    host: usize,
    axis: &str,
) -> Result<(u32, u32)> {
    let resolve = |is_absolute: bool, is_end: bool| -> Result<u32> {
        if is_absolute {
            let range = absolute.first().ok_or_else(|| {
                Error::ParseError(format!(
                    "Numbers formula colon tract has no absolute {axis} coordinate"
                ))
            })?;
            Ok(if is_end {
                range.range_end.unwrap_or(range.range_begin)
            } else {
                range.range_begin
            })
        } else {
            let range = relative.first().ok_or_else(|| {
                Error::ParseError(format!(
                    "Numbers formula colon tract has no relative {axis} coordinate"
                ))
            })?;
            let stored = if is_end {
                range.range_end.unwrap_or(range.range_begin)
            } else {
                range.range_begin
            };
            resolve_formula_coordinate(host, stored, false, axis)
        }
    };
    Ok((
        resolve(begin_is_absolute, false)?,
        resolve(end_is_absolute, true)?,
    ))
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
    Ok(stack.split_off(start))
}

fn take_field<'a>(data: &'a [u8], cursor: &mut usize, length: usize) -> Result<&'a [u8]> {
    let end = cursor
        .checked_add(length)
        .ok_or_else(|| Error::ParseError("Numbers cell field offset overflow".to_string()))?;
    let field = data.get(*cursor..end).ok_or_else(|| {
        Error::ParseError(format!(
            "Truncated Numbers cell field at offset {} (need {length} bytes)",
            *cursor
        ))
    })?;
    *cursor = end;
    Ok(field)
}

fn read_u32_le(data: &[u8]) -> Result<u32> {
    let bytes: [u8; 4] = data
        .try_into()
        .map_err(|_| Error::ParseError("Expected a four-byte Numbers field".to_string()))?;
    Ok(u32::from_le_bytes(bytes))
}

fn read_f64_le(data: &[u8]) -> Result<FiniteF64> {
    let bytes: [u8; 8] = data
        .try_into()
        .map_err(|_| Error::ParseError("Expected an eight-byte Numbers field".to_string()))?;
    FiniteF64::new(f64::from_le_bytes(bytes)).map_err(|_| {
        Error::ParseError("Numbers scalar field must contain a finite value".to_string())
    })
}

fn finite_zero() -> Result<FiniteF64> {
    FiniteF64::new(0.0).map_err(|_| {
        Error::InvalidFormat("Numbers zero scalar is unexpectedly non-finite".to_string())
    })
}

#[cfg(test)]
fn decode_legacy_table_candidate<T>(
    data: &[u8],
    admit: impl FnOnce() -> Result<()>,
    parse: impl FnOnce(tst::TableModelArchive) -> Result<T>,
) -> Result<Option<T>> {
    match has_legacy_table_model_wire_shape(data) {
        Ok(true) => {},
        Ok(false) => return Ok(None),
        Err(error) => return Err(error),
    }
    admit()?;
    let table_model = tst::TableModelArchive::decode(data).map_err(Error::protobuf)?;
    parse(table_model).map(Some)
}

#[cfg(test)]
mod tests {
    use super::{
        CellBudget, CellTables, Error, FormulaArchiveBytes, FormulaReferenceBudget,
        FormulaReferenceMaps, FormulaRenderer, MAX_FORMULA_CATEGORY_DEPTH, MAX_FORMULA_WIRE_BYTES,
        MAX_FORMULA_WORK, MAX_PAYLOAD_WORK, ProjectionBudget, Table, TableDataExtractor,
        TileRowVisitor, collect_formula_category_payload, decode_legacy_table_candidate,
        formula_table_name, has_legacy_table_model_wire_shape,
        map_table_cell_decode_limit_with_offsets,
        map_table_cell_decode_limit_with_reference_offset, preflight_formula_category_payload,
        preflight_formula_owner, render_formula, render_formula_ast_array,
        render_formula_compatibility, render_formula_precharged,
    };
    use crate::cell::Value as CellValue;
    use crate::cell::wire::{BncCell, decimal128_le};
    use crate::package::{
        Components, Index, ReadOptions, SemanticPath, compatibility_tables_from_bytes_with_options,
    };
    use crate::{
        DEFAULT_MAX_TEXT_BYTES, Package, PackageSemanticLimits as SemanticLimits, SemanticLimitKind,
    };
    use litchi_iwa_archive::Limits;
    use litchi_iwa_common::comment::{Comment, StorageId};
    use litchi_iwa_common::wire::{append_length_delimited_field, append_varint_field};
    use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
    use litchi_iwa_protos::tsce::ast_node_array_archive::ast_colon_tract_archive::AstColonTractAbsoluteRangeArchive;
    use litchi_iwa_protos::tsce::ast_node_array_archive::{
        AstColonTractArchive, AstColumnCoordinateArchive, AstCrossTableCellReferenceNodeArchive,
        AstLocalCellReferenceNodeArchive, AstNodeArchive, AstNodeType, AstRowCoordinateArchive,
        AstStickyBits,
    };
    use litchi_iwa_protos::{numbers_table_cell_storage_codec, tn, tsce, tsd, tsp, tst};
    use prost::Message as _;
    use std::collections::HashMap;
    use std::path::PathBuf;

    const TEST_DECIMAL_FLAG: u32 = 0x0000_0001;

    fn formula_node(kind: AstNodeType) -> AstNodeArchive {
        AstNodeArchive {
            ast_node_type: kind as i32,
            ..Default::default()
        }
    }

    fn number_node(value: f64) -> AstNodeArchive {
        AstNodeArchive {
            ast_number_node_number: Some(value),
            ..formula_node(AstNodeType::NumberNode)
        }
    }

    fn formula(nodes: Vec<AstNodeArchive>) -> tsce::FormulaArchive {
        tsce::FormulaArchive {
            ast_node_array: tsce::AstNodeArrayArchive { ast_node: nodes },
            ..Default::default()
        }
    }

    fn compatibility_render(
        input: &tsce::FormulaArchive,
        host_row: usize,
        host_column: usize,
        row_count: usize,
        column_count: usize,
        references: &FormulaReferenceMaps,
    ) -> super::Result<String> {
        let source = input.encode_to_vec();
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let raw = FormulaArchiveBytes::from_wire(&source, &mut budget)?;
        budget.charge_formula_render_work(raw.scalar_visitor_node_count)?;
        budget.charge_formula_lazy_work(raw.lazy_traversal_entry_count)?;
        render_formula_compatibility(
            &raw,
            host_row,
            host_column,
            row_count,
            column_count,
            references,
            &mut budget,
        )
    }

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            ..Default::default()
        }
    }

    fn formula_owner_wire(uuid: &[u8], table: &[u8]) -> Vec<u8> {
        let mut source = Vec::new();
        append_length_delimited_field(&mut source, 1, uuid).expect("owner UUID field");
        append_length_delimited_field(&mut source, 11, table).expect("owner table field");
        source
    }

    fn formula_owner_uuid(lower: u64, upper: u64) -> Vec<u8> {
        tsp::Uuid { lower, upper }.encode_to_vec()
    }

    fn archive_object(identifier: u64, messages: Vec<RawMessage>) -> super::Result<ArchiveObject> {
        ArchiveObject::new(identifier, messages)
            .map_err(|error| Error::InvalidFormat(error.to_string()))
    }

    fn compatibility_package(objects: Vec<ArchiveObject>) -> super::Result<Vec<u8>> {
        let archive = Archive { objects }
            .to_bytes()
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let stream = SnappyStream::compress(&archive)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        Ok(litchi_iwa_archive::package::to_bytes(
            [("Index/Document.iwa", stream.as_slice())],
            Limits::default(),
        )?)
    }

    fn legacy_model(name: &str, sidecar_id: u64) -> tst::TableModelArchive {
        tst::TableModelArchive {
            table_name: name.to_owned(),
            number_of_rows: 1,
            number_of_columns: 1,
            base_data_store: tst::DataStore {
                string_table: reference(sidecar_id),
                formula_table: reference(sidecar_id),
                ..Default::default()
            },
            ..Default::default()
        }
    }

    fn sparse_model_wire(
        names: &[&str],
        dimensions: Option<(u32, u32)>,
        data_store: &[u8],
    ) -> Vec<u8> {
        let mut source = Vec::new();
        append_length_delimited_field(&mut source, 4, data_store)
            .expect("sparse model data-store field");
        if let Some((rows, columns)) = dimensions {
            append_varint_field(&mut source, 6, u64::from(rows)).expect("sparse model rows");
            append_varint_field(&mut source, 7, u64::from(columns)).expect("sparse model columns");
        }
        for name in names {
            append_length_delimited_field(&mut source, 8, name.as_bytes())
                .expect("sparse model name field");
        }
        source
    }

    #[test]
    fn sparse_shape_requires_explicit_dimensions_and_accepts_legacy_duplicate_name()
    -> super::Result<()> {
        let minimal = sparse_model_wire(&["minimal"], None, &[]);
        assert!(!super::sparse_table_model_compatibility_shape(&minimal));

        // Zero dimensions are valid only when both scalar fields are
        // explicitly present; omitted fields must not be synthesized by the
        // compatibility defaults into a 0x0 table.
        let explicit_zero = sparse_model_wire(&["empty"], Some((0, 0)), &[]);
        assert!(super::sparse_table_model_compatibility_shape(
            &explicit_zero
        ));

        let duplicate = sparse_model_wire(&["first", "last"], Some((1, 1)), &[]);
        assert!(super::sparse_table_model_compatibility_shape(&duplicate));
        Ok(())
    }

    #[test]
    fn compatibility_sparse_projection_retains_last_name() -> super::Result<()> {
        let source = sparse_model_wire(&["first", "last"], Some((0, 0)), &[]);
        with_list_extractor(Vec::new(), false, |extractor| {
            let mut budget = ProjectionBudget::new(SemanticLimits::default());
            let projected = extractor.project_table_model(
                &source,
                &mut budget,
                SemanticPath::StructuredTables,
            )?;
            assert!(projected.compatibility_defaults);
            assert_eq!(projected.table_name, "last");
            Ok(())
        })
    }

    #[test]
    fn sparse_ingress_fallback_retains_last_duplicate_name() -> super::Result<()> {
        // An empty DataStore forces the strict model route to fail. Sparse
        // compatibility ingress mirrors generated protobuf last-wins name
        // behavior; the names transaction's independent strict projection
        // rejects this source before publishing a rename.
        let source = sparse_model_wire(&["first", "last"], Some((1, 1)), &[]);
        with_list_extractor(Vec::new(), true, |extractor| {
            let mut budget = ProjectionBudget::new(SemanticLimits::default());
            let result =
                extractor.project_table_model(&source, &mut budget, SemanticPath::StructuredTables);
            let projected = result?;
            assert!(projected.compatibility_defaults);
            assert_eq!(projected.table_name, "last");
            Ok(())
        })
    }

    #[test]
    fn dense_native_fallback_rejects_duplicate_name_after_model_failure() -> super::Result<()> {
        // A non-default required HeaderStorage envelope identifies a native
        // model even when the selected DataStore routes are incomplete. The
        // dense fallback must retain the strict names projection and reject
        // the duplicate instead of applying sparse last-wins behavior.
        let mut data_store = Vec::new();
        append_length_delimited_field(&mut data_store, 1, &[0x08, 0x01])
            .expect("native header-storage envelope");
        let source = sparse_model_wire(&["first", "last"], Some((1, 1)), &data_store);
        with_list_extractor(Vec::new(), true, |extractor| {
            let mut budget = ProjectionBudget::new(SemanticLimits::default());
            let result =
                extractor.project_table_model(&source, &mut budget, SemanticPath::StructuredTables);
            assert!(result.is_err(), "native duplicate name was admitted");
            Ok(())
        })
    }

    #[test]
    fn native_sparse_fallback_does_not_retry_after_model_limit() -> super::Result<()> {
        let source = sparse_model_wire(&["limited"], Some((1, 1)), &[]);
        with_list_extractor(Vec::new(), true, |extractor| {
            let mut budget = ProjectionBudget::new(SemanticLimits::default());
            // Make the first strict model walk hit its aggregate work ceiling.
            // A resource limit must remain authoritative; it must not trigger
            // the sparse compatibility retry.
            budget.payload_work = MAX_PAYLOAD_WORK;
            let result =
                extractor.project_table_model(&source, &mut budget, SemanticPath::StructuredTables);
            assert!(
                matches!(
                    &result,
                    Err(Error::SemanticLimit {
                        kind: SemanticLimitKind::FormulaWork,
                        ..
                    })
                ),
                "model limit was retried or mapped incorrectly"
            );
            Ok(())
        })
    }

    #[test]
    fn native_sparse_fallback_keeps_malformed_tiles_strict() -> super::Result<()> {
        let mut data_store = Vec::new();
        append_length_delimited_field(&mut data_store, 3, &[0xff])
            .expect("malformed tile-storage field");
        append_length_delimited_field(&mut data_store, 4, &reference(90).encode_to_vec())
            .expect("string-table reference");
        append_length_delimited_field(&mut data_store, 6, &reference(91).encode_to_vec())
            .expect("formula-table reference");
        let source = sparse_model_wire(&["tiles"], Some((1, 1)), &data_store);
        let result = with_list_extractor(
            vec![
                archive_object(
                    90,
                    vec![list_message(
                        tst::table_data_list::ListType::String,
                        Vec::new(),
                        Vec::new(),
                    )],
                )?,
                archive_object(
                    91,
                    vec![list_message(
                        tst::table_data_list::ListType::Formula,
                        Vec::new(),
                        Vec::new(),
                    )],
                )?,
            ],
            true,
            |extractor| {
                extractor.parse_projected_table_payload(&source, SemanticPath::StructuredTables)
            },
        );
        assert!(
            matches!(&result, Err(Error::InvalidFormat(_))),
            "malformed tile-storage payload was admitted: {result:?}"
        );
        Ok(())
    }

    fn native_fixture() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/numbers/basic.numbers")
    }

    fn empty_list(list_type: tst::table_data_list::ListType) -> RawMessage {
        RawMessage {
            type_: 6_005,
            data: tst::TableDataList {
                list_type: list_type as i32,
                next_list_id: 1,
                ..Default::default()
            }
            .encode_to_vec(),
        }
    }

    fn string_list_entry(key: u32, value: &str) -> tst::table_data_list::ListEntry {
        tst::table_data_list::ListEntry {
            key,
            refcount: 1,
            string: Some(value.to_owned()),
            ..Default::default()
        }
    }

    fn list_message(
        list_type: tst::table_data_list::ListType,
        entries: Vec<tst::table_data_list::ListEntry>,
        segments: Vec<u64>,
    ) -> RawMessage {
        RawMessage {
            type_: 6_005,
            data: tst::TableDataList {
                list_type: list_type as i32,
                next_list_id: 1,
                entries,
                segments: segments.into_iter().map(reference).collect(),
                ..Default::default()
            }
            .encode_to_vec(),
        }
    }

    fn segment_message(
        list_type: tst::table_data_list::ListType,
        location: u32,
        length: u32,
        entries: Vec<tst::table_data_list::ListEntry>,
    ) -> RawMessage {
        RawMessage {
            type_: 6_011,
            data: tst::TableDataListSegment {
                list_type: list_type as i32,
                key_range: tsp::Range { location, length },
                entries,
            }
            .encode_to_vec(),
        }
    }

    fn nested_unknown_group_prefix(depth: usize) -> Vec<u8> {
        fn varint(output: &mut Vec<u8>, mut value: u64) {
            while value >= 0x80 {
                output.push((value as u8 & 0x7f) | 0x80);
                value >>= 7;
            }
            output.push(value as u8);
        }

        let mut prefix = Vec::new();
        for _ in 0..depth {
            varint(&mut prefix, (90_u64 << 3) | 3);
        }
        for _ in 0..depth {
            varint(&mut prefix, (90_u64 << 3) | 4);
        }
        prefix
    }

    fn with_list_extractor<T>(
        objects: Vec<ArchiveObject>,
        document_projection: bool,
        visit: impl FnOnce(&TableDataExtractor<'_>) -> super::Result<T>,
    ) -> super::Result<T> {
        let bytes = compatibility_package(objects)?;
        let components = Components::from_bytes(&bytes, Limits::default())?;
        let index = Index::from_components(&components, SemanticLimits::MAX_OBJECTS)?;
        let extractor = TableDataExtractor::new(&components, &index, SemanticLimits::default());
        if document_projection {
            visit(&extractor.without_comments())
        } else {
            visit(&extractor)
        }
    }

    fn load_string_list(
        extractor: &TableDataExtractor<'_>,
        object_id: u64,
    ) -> super::Result<(super::CompactTable<String>, ProjectionBudget)> {
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let table = load_string_list_with_budget(extractor, object_id, &mut budget)?;
        Ok((table, budget))
    }

    fn load_string_list_with_budget(
        extractor: &TableDataExtractor<'_>,
        object_id: u64,
        budget: &mut ProjectionBudget,
    ) -> super::Result<super::CompactTable<String>> {
        let mut converter =
            |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
             _budget: &mut ProjectionBudget| {
                let value = entry.string_value().ok_or_else(|| {
                    Error::InvalidFormat("test string entry has no value".to_owned())
                })?;
                let mut retained = String::new();
                retained
                    .try_reserve_exact(value.len())
                    .map_err(|_| super::allocation_error("test string", value.len()))?;
                retained.push_str(value);
                Ok(retained)
            };
        let table = extractor.load_table_data_list_entries(
            object_id,
            tst::table_data_list::ListType::String,
            budget,
            &mut converter,
        )?;
        Ok(table)
    }

    fn tile_row(
        tile_row_index: u32,
        cell_count: u32,
        pre_bnc_storage: Vec<u8>,
        pre_bnc_offsets: Vec<u8>,
        modern_storage: Option<Vec<u8>>,
        modern_offsets: Option<Vec<u8>>,
        has_wide_offsets: Option<bool>,
    ) -> tst::TileRowInfo {
        tst::TileRowInfo {
            tile_row_index,
            cell_count,
            cell_storage_buffer_pre_bnc: pre_bnc_storage,
            cell_offsets_pre_bnc: pre_bnc_offsets,
            cell_storage_buffer: modern_storage,
            cell_offsets: modern_offsets,
            has_wide_offsets,
            ..Default::default()
        }
    }

    fn tile_source(rows: Vec<tst::TileRowInfo>, num_cells: u32, num_rows: u32) -> Vec<u8> {
        tst::Tile {
            max_column: 1,
            max_row: 1,
            num_cells,
            numrows: num_rows,
            row_infos: rows,
            ..Default::default()
        }
        .encode_to_vec()
    }

    fn decode_tile_scalars(source: &[u8]) -> numbers_table_cell_storage_codec::TileSnapshot {
        let options = numbers_table_cell_storage_codec::DecodeOptions::new(
            source.len().max(1),
            usize::MAX,
            usize::MAX,
            16,
            usize::MAX,
            usize::MAX,
        );
        numbers_table_cell_storage_codec::decode_tile_with_visitor(source, options, &mut ())
            .unwrap_or_else(|error| panic!("tile projection failed: {error:?}"))
            .0
    }

    fn run_tile_visitor(
        source: &[u8],
        column_count: usize,
        limits: SemanticLimits,
    ) -> super::Result<(
        Result<
            numbers_table_cell_storage_codec::DecodeReport,
            numbers_table_cell_storage_codec::DecodeError,
        >,
        Table,
        ProjectionBudget,
        Option<Error>,
        usize,
    )> {
        let mut table = Table::with_dimensions("tile", 2, column_count)?;
        let strings: Box<[(u32, String)]> = Box::default();
        let formulas: Box<[(u32, FormulaArchiveBytes)]> = Box::default();
        let formula_errors: Box<[(u32, String)]> = Box::default();
        let rich_text: Box<[(u32, String)]> = Box::default();
        let comments: Box<[(u32, Comment)]> = Box::default();
        let formula_references = FormulaReferenceMaps::default();
        let cell_tables = CellTables {
            strings: &strings,
            formulas: &formulas,
            formula_errors: &formula_errors,
            rich_text: &rich_text,
            comments: Some(&comments),
            formula_references: &formula_references,
        };
        let mut cell_budget = CellBudget::new();
        let mut projection_budget = ProjectionBudget::new(limits);
        let options = numbers_table_cell_storage_codec::DecodeOptions::new(
            source.len().max(1),
            usize::MAX,
            usize::MAX,
            16,
            usize::MAX,
            usize::MAX,
        );
        let (materialized_cells, semantic_error, report) = {
            let mut visitor = TileRowVisitor {
                row_origin: 0,
                tile_size: 2,
                row_count: table.row_count(),
                column_count: table.column_count(),
                budget: &mut cell_budget,
                cell_tables: &cell_tables,
                projection_budget: &mut projection_budget,
                table: &mut table,
                materialized_cells: 0,
                semantic_error: None,
            };
            let decode_result = numbers_table_cell_storage_codec::decode_tile_with_visitor(
                source,
                options,
                &mut visitor,
            )
            .map(|(_, report)| report);
            (
                visitor.materialized_cells,
                visitor.semantic_error.take(),
                decode_result,
            )
        };
        Ok((
            report,
            table,
            projection_budget,
            semantic_error,
            materialized_cells,
        ))
    }

    fn parse_projected_rows(source: &[u8], column_count: usize) -> super::Result<Table> {
        let (report, table, mut projection_budget, semantic_error, materialized_cells) =
            run_tile_visitor(source, column_count, SemanticLimits::default())?;
        let report = report
            .map_err(|error| Error::InvalidFormat(format!("tile projection failed: {error:?}")))?;
        projection_budget.charge_decode_report(report)?;
        projection_budget.charge_materialized_cells(materialized_cells)?;
        if let Some(error) = semantic_error {
            return Err(error);
        }
        Ok(table)
    }

    #[test]
    fn table_info_wire_is_a_legacy_classification_miss() -> super::Result<()> {
        let table_info = tst::TableInfoArchive::default().encode_to_vec();
        assert!(!has_legacy_table_model_wire_shape(&table_info)?);
        assert!(
            decode_legacy_table_candidate(
                &table_info,
                || panic!("table-info false positive invoked admission"),
                |_model| -> super::Result<()> {
                    panic!("table-info false positive reached model extraction")
                },
            )?
            .is_none()
        );
        Ok(())
    }

    #[test]
    fn legacy_table_info_false_positive_is_ignored_when_table_budget_is_full() {
        let bytes = compatibility_package(vec![
            archive_object(
                1,
                vec![RawMessage {
                    type_: 1,
                    data: tn::DocumentArchive::default().encode_to_vec(),
                }],
            )
            .expect("document archive"),
            archive_object(
                10,
                vec![RawMessage {
                    type_: 6_000,
                    data: tst::TableInfoArchive::default().encode_to_vec(),
                }],
            )
            .expect("table-info archive"),
        ])
        .expect("compatibility package");
        let components = Components::from_bytes(&bytes, Limits::default()).expect("components");
        let index =
            Index::from_components(&components, SemanticLimits::MAX_OBJECTS).expect("object index");
        let extractor = TableDataExtractor::new(&components, &index, SemanticLimits::default());

        let tables = extractor
            .extract_all_semantic_tables(0)
            .expect("table-info false positive is not a table");
        assert!(tables.is_empty());
        let budget = *extractor.projection_budget.borrow();
        assert_eq!(budget.references, 0);
        assert_eq!(budget.payload_fields, 0);
        assert_eq!(budget.payload_work, 0);
        assert_eq!(budget.staging_text_bytes, 0);
        assert_eq!(budget.materialized_cells, 0);
        assert_eq!(budget.output_text_bytes, 0);
        assert_eq!(budget.formula_render_work, 0);
    }

    #[test]
    fn legacy_candidate_admission_happens_before_decode_or_parse() {
        let encoded = legacy_model("over-budget", 90).encode_to_vec();
        let mut admission_called = false;
        let mut parse_called = false;
        let result = decode_legacy_table_candidate(
            &encoded,
            || {
                admission_called = true;
                Err(super::table_limit_error(1, 0))
            },
            |_model| {
                parse_called = true;
                Ok(())
            },
        );

        assert!(matches!(
            result,
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::Tables,
                observed: 1,
                maximum: 0,
                path: SemanticPath::StructuredTables,
            })
        ));
        assert!(admission_called);
        assert!(!parse_called);
    }

    #[test]
    fn legacy_model_at_full_table_budget_leaves_projection_budget_unchanged() {
        let bytes = compatibility_package(vec![
            archive_object(
                1,
                vec![RawMessage {
                    type_: 1,
                    data: tn::DocumentArchive::default().encode_to_vec(),
                }],
            )
            .expect("document archive"),
            archive_object(
                10,
                vec![RawMessage {
                    type_: 6_000,
                    data: legacy_model("over-budget", 90).encode_to_vec(),
                }],
            )
            .expect("legacy table archive"),
        ])
        .expect("compatibility package");
        let components = Components::from_bytes(&bytes, Limits::default()).expect("components");
        let index =
            Index::from_components(&components, SemanticLimits::MAX_OBJECTS).expect("object index");
        let extractor = TableDataExtractor::new(&components, &index, SemanticLimits::default());
        let before = *extractor.projection_budget.borrow();

        let error = extractor
            .extract_all_semantic_tables(0)
            .expect_err("over-budget legacy model");
        assert!(matches!(
            &error,
            Error::SemanticLimit {
                kind: SemanticLimitKind::Tables,
                observed: 1,
                maximum: 0,
                path: SemanticPath::StructuredTables,
            }
        ));
        let after = *extractor.projection_budget.borrow();
        assert_eq!(before.references, after.references);
        assert_eq!(before.payload_fields, after.payload_fields);
        assert_eq!(before.payload_work, after.payload_work);
        assert_eq!(before.staging_text_bytes, after.staging_text_bytes);
        assert_eq!(before.materialized_cells, after.materialized_cells);
        assert_eq!(before.output_text_bytes, after.output_text_bytes);
        assert_eq!(before.formula_render_work, after.formula_render_work);
    }

    #[test]
    fn malformed_schema_shaped_legacy_candidate_wins_when_table_budget_is_available() {
        let malformed = [0x22, 0x01, 0xff, 0x30, 0x01, 0x38, 0x01, 0x42, 0x01, b'x'];
        let bytes = compatibility_package(vec![
            archive_object(
                1,
                vec![RawMessage {
                    type_: 1,
                    data: tn::DocumentArchive::default().encode_to_vec(),
                }],
            )
            .expect("document archive"),
            archive_object(
                10,
                vec![RawMessage {
                    type_: 6_000,
                    data: malformed.to_vec(),
                }],
            )
            .expect("malformed legacy archive"),
        ])
        .expect("compatibility package");
        let components = Components::from_bytes(&bytes, Limits::default()).expect("components");
        let index =
            Index::from_components(&components, SemanticLimits::MAX_OBJECTS).expect("object index");
        let extractor = TableDataExtractor::new(&components, &index, SemanticLimits::default());

        let error = extractor
            .extract_all_semantic_tables(1)
            .expect_err("malformed schema-shaped legacy payload");
        assert!(matches!(
            error,
            Error::MalformedPayload {
                path: SemanticPath::StructuredTables,
            }
        ));
    }

    #[test]
    fn admitted_legacy_candidate_preserves_common_allocation_error() -> super::Result<()> {
        let encoded = legacy_model("model", 90).encode_to_vec();
        assert!(has_legacy_table_model_wire_shape(&encoded)?);
        let result: super::Result<Option<()>> = decode_legacy_table_candidate(
            &encoded,
            || Ok(()),
            |_model| {
                Err(Error::Common(litchi_iwa_common::Error::Allocation {
                    resource: "Numbers retained semantic text",
                    amount: 5,
                }))
            },
        );
        let error = result
            .err()
            .ok_or_else(|| Error::InvalidFormat("allocation error was swallowed".to_owned()))?;
        assert!(matches!(
            &error,
            Error::Common(litchi_iwa_common::Error::Allocation {
                resource: "Numbers retained semantic text",
                amount: 5,
            })
        ));
        assert_eq!(
            error.to_string(),
            "IWA wire allocation failed for Numbers retained semantic text: 5"
        );
        Ok(())
    }

    #[test]
    fn table_cell_codec_allocation_limit_maps_to_typed_common_error() {
        let error = map_table_cell_decode_limit_with_reference_offset(
            numbers_table_cell_storage_codec::DecodeLimit::Allocation { requested: 17 },
            0,
        );
        assert!(matches!(
            error,
            Error::Common(litchi_iwa_common::Error::Allocation {
                resource: "Numbers table storage projection",
                amount: 17,
            })
        ));
    }

    #[test]
    fn table_cell_codec_aggregate_limits_include_prior_projection_cost() {
        let fields = map_table_cell_decode_limit_with_offsets(
            numbers_table_cell_storage_codec::DecodeLimit::Fields {
                observed: 4,
                maximum: 5,
            },
            0,
            11,
            0,
            0,
        );
        assert!(matches!(
            fields,
            Error::SemanticLimit {
                kind: SemanticLimitKind::Objects,
                observed: 15,
                maximum: 16,
                ..
            }
        ));

        let work = map_table_cell_decode_limit_with_offsets(
            numbers_table_cell_storage_codec::DecodeLimit::Work {
                observed: 7,
                maximum: 8,
            },
            0,
            0,
            13,
            0,
        );
        assert!(matches!(
            work,
            Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaWork,
                observed: 20,
                maximum: 21,
                ..
            }
        ));

        let text = map_table_cell_decode_limit_with_offsets(
            numbers_table_cell_storage_codec::DecodeLimit::Text {
                observed: 9,
                maximum: 10,
            },
            0,
            0,
            0,
            17,
        );
        assert!(matches!(
            text,
            Error::SemanticLimit {
                kind: SemanticLimitKind::TextBytes,
                observed: 26,
                maximum: 27,
                ..
            }
        ));
    }

    #[test]
    fn admitted_legacy_decode_failure_reports_exact_content_free_path() -> super::Result<()> {
        // Required model fields 4, 6, 7, and 8 are present with their schema
        // wire types, but the nested DataStore payload is malformed.
        let encoded = [0x22, 0x01, 0xff, 0x30, 0x01, 0x38, 0x01, 0x42, 0x01, b'x'];
        assert!(has_legacy_table_model_wire_shape(&encoded)?);
        let result = decode_legacy_table_candidate(
            &encoded,
            || Ok(()),
            |_model| -> super::Result<()> {
                panic!("malformed admitted model reached semantic extraction")
            },
        );
        let error = result.err().ok_or_else(|| {
            Error::InvalidFormat("admitted model decode error was swallowed".to_owned())
        })?;
        assert!(matches!(
            &error,
            Error::MalformedPayload {
                path: SemanticPath::StructuredTables,
            }
        ));
        assert_eq!(
            error.to_string(),
            "malformed Numbers payload at structured tables"
        );
        Ok(())
    }

    #[test]
    fn admitted_legacy_duplicate_required_field_fails_closed() -> super::Result<()> {
        let mut encoded = legacy_model("model", 90).encode_to_vec();
        encoded.extend_from_slice(&[0x30, 0x01]);
        let result = decode_legacy_table_candidate(
            &encoded,
            || Ok(()),
            |_model| -> super::Result<()> {
                panic!("ambiguous admitted model reached semantic extraction")
            },
        );
        assert!(matches!(
            result,
            Err(Error::MalformedPayload {
                path: SemanticPath::StructuredTables,
            })
        ));
        Ok(())
    }

    #[test]
    fn admitted_legacy_malformed_trailing_wire_fails_closed() -> super::Result<()> {
        let mut encoded = legacy_model("model", 90).encode_to_vec();
        encoded.push(0xff);
        assert!(matches!(
            decode_legacy_table_candidate(
                &encoded,
                || Ok(()),
                |_model| -> super::Result<()> {
                    panic!("malformed admitted model reached semantic extraction")
                },
            ),
            Err(Error::MalformedPayload {
                path: SemanticPath::StructuredTables,
            })
        ));
        Ok(())
    }

    #[test]
    fn admitted_legacy_name_limit_reports_exact_content_free_error() -> super::Result<()> {
        let sidecars = archive_object(
            90,
            [
                tst::table_data_list::ListType::String,
                tst::table_data_list::ListType::Formula,
            ]
            .into_iter()
            .map(|list_type| RawMessage {
                type_: 6_005,
                data: tst::TableDataList {
                    list_type: list_type as i32,
                    next_list_id: 1,
                    ..Default::default()
                }
                .encode_to_vec(),
            })
            .collect(),
        )?;
        let bytes = compatibility_package(vec![
            archive_object(
                1,
                vec![RawMessage {
                    type_: 1,
                    data: tn::DocumentArchive::default().encode_to_vec(),
                }],
            )?,
            sidecars,
            archive_object(
                10,
                vec![RawMessage {
                    type_: 6_000,
                    data: legacy_model("\u{e9}", 90).encode_to_vec(),
                }],
            )?,
        ])?;
        let semantic = SemanticLimits::default()
            .with_projection_limits(SemanticLimits::MAX_MATERIALIZED_CELLS, 1)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let result = compatibility_tables_from_bytes_with_options(
            &bytes,
            ReadOptions::new(Limits::default(), semantic),
        );
        let error = result
            .err()
            .ok_or_else(|| Error::InvalidFormat("text limit error was swallowed".to_owned()))?;
        assert!(matches!(
            &error,
            Error::SemanticLimit {
                kind: SemanticLimitKind::OutputTextBytes,
                observed: 2,
                maximum: 1,
                path: SemanticPath::StructuredTables,
            }
        ));
        assert_eq!(
            error.to_string(),
            "Numbers semantic output text bytes limit exceeded at structured tables: observed 2, maximum 1"
        );
        Ok(())
    }

    #[test]
    fn padded_missing_cell_offset_slots_are_accepted() {
        let offsets = [
            0, 0, // column 0 starts at byte 0
            0xff, 0xff, // native tile-width padding
            0xff, 0xff,
        ];
        let cells = TableDataExtractor::parse_cell_offsets(&offsets, 1, false, 1, 1)
            .unwrap_or_else(|error| panic!("missing padded slots were rejected: {error}"));
        assert_eq!(cells, vec![(0, 0..1)]);
    }

    #[test]
    fn tile_projection_accepts_missing_slots_inside_the_table_width() -> super::Result<()> {
        let cell = BncCell::minimal().encode();
        let source = tile_source(
            vec![tile_row(
                0,
                1,
                cell.clone(),
                vec![0, 0, 0xff, 0xff],
                Some(cell),
                Some(vec![0, 0, 0xff, 0xff]),
                None,
            )],
            1,
            1,
        );
        let table = parse_projected_rows(&source, 2)?;
        assert_eq!(table.get_cell(0, 0), Some(&CellValue::Empty));
        assert_eq!(table.get_cell(0, 1), None);
        Ok(())
    }

    #[test]
    fn tile_projection_rejects_sparse_rows_with_a_false_declared_cell_count() {
        let cell = BncCell::minimal().encode();
        let source = tile_source(
            vec![tile_row(
                0,
                2,
                cell.clone(),
                vec![0, 0, 0xff, 0xff],
                Some(cell),
                Some(vec![0, 0, 0xff, 0xff]),
                None,
            )],
            2,
            1,
        );
        let error = parse_projected_rows(&source, 2)
            .expect_err("a missing slot cannot satisfy the declared occupied-cell count");
        assert!(matches!(error, Error::ParseError(message) if message.contains("has 1 offsets")));
    }

    #[test]
    fn populated_cell_offset_slots_outside_table_width_are_rejected() {
        let offsets = [0, 0, 0, 0];
        let error = match TableDataExtractor::parse_cell_offsets(&offsets, 1, false, 1, 1) {
            Ok(cells) => panic!("populated padded slot produced cells: {cells:?}"),
            Err(error) => error,
        };
        assert!(matches!(
            error,
            Error::InvalidFormat(message) if message.contains("outside the declared table width")
        ));
    }

    #[test]
    fn tile_rows_use_pre_bnc_buffers_when_modern_buffers_are_absent() -> super::Result<()> {
        let source = tile_source(
            vec![tile_row(0, 1, vec![0; 8], vec![0, 0], None, None, None)],
            1,
            1,
        );
        let table = parse_projected_rows(&source, 1)?;
        assert_eq!(table.get_cell(0, 0), Some(&CellValue::Empty));
        Ok(())
    }

    #[test]
    fn tile_rows_use_modern_buffers_only_when_both_are_present() -> super::Result<()> {
        let modern = BncCell::minimal().encode();
        let source = tile_source(
            vec![tile_row(
                0,
                1,
                vec![0xff],
                vec![0xff],
                Some(modern),
                Some(vec![0, 0]),
                None,
            )],
            1,
            1,
        );
        let table = parse_projected_rows(&source, 1)?;
        assert_eq!(table.get_cell(0, 0), Some(&CellValue::Empty));
        Ok(())
    }

    #[test]
    fn tile_rows_fall_back_to_both_pre_bnc_buffers_for_partial_modern_storage() -> super::Result<()>
    {
        let source = tile_source(
            vec![tile_row(
                0,
                1,
                vec![0; 8],
                vec![0, 0],
                Some(vec![0xff]),
                None,
                None,
            )],
            1,
            1,
        );
        let table = parse_projected_rows(&source, 1)?;
        assert_eq!(table.get_cell(0, 0), Some(&CellValue::Empty));
        Ok(())
    }

    #[test]
    fn tile_rows_honor_wide_offset_units() -> super::Result<()> {
        let cell = BncCell::minimal().encode();
        let mut storage = cell.clone();
        storage.extend_from_slice(&cell);
        let source = tile_source(
            vec![tile_row(
                0,
                2,
                Vec::new(),
                vec![0xff, 0xff, 0xff, 0xff],
                Some(storage),
                Some(vec![0, 0, 3, 0]),
                Some(true),
            )],
            2,
            1,
        );
        let table = parse_projected_rows(&source, 2)?;
        assert_eq!(table.get_cell(0, 0), Some(&CellValue::Empty));
        assert_eq!(table.get_cell(0, 1), Some(&CellValue::Empty));
        Ok(())
    }

    #[test]
    fn repeated_comment_cells_charge_owned_materialization_before_insertion() -> super::Result<()> {
        let mut cell = BncCell::minimal();
        cell.set_comment_identifier(Some(7));
        let cell = cell.encode();
        let cell_end = u16::try_from(cell.len())
            .map_err(|_| Error::InvalidFormat("test comment cell is too large".to_owned()))?;
        let mut storage = cell.clone();
        storage.extend_from_slice(&cell);
        let row_source = tile_row(
            0,
            2,
            Vec::new(),
            Vec::new(),
            Some(storage),
            Some(vec![
                0,
                0,
                cell_end.to_le_bytes()[0],
                cell_end.to_le_bytes()[1],
            ]),
            Some(false),
        )
        .encode_to_vec();
        let options = numbers_table_cell_storage_codec::DecodeOptions::new(
            row_source.len().max(1),
            usize::MAX,
            usize::MAX,
            16,
            usize::MAX,
            usize::MAX,
        );
        let row = numbers_table_cell_storage_codec::decode_tile_row_info(&row_source, options)
            .map_err(|error| Error::InvalidFormat(format!("comment row failed: {error:?}")))?;

        let strings: Box<[(u32, String)]> = Box::default();
        let formulas: Box<[(u32, FormulaArchiveBytes)]> = Box::default();
        let formula_errors: Box<[(u32, String)]> = Box::default();
        let rich_text: Box<[(u32, String)]> = Box::default();
        let comments: Box<[(u32, Comment)]> = vec![(
            7,
            Comment {
                text: "copy".to_owned(),
                creation_date_seconds: None,
                author_id: None,
                reply_ids: vec![StorageId::new(9).unwrap()].into_boxed_slice(),
                storage_uuid: None,
            },
        )]
        .into_boxed_slice();
        let formula_references = FormulaReferenceMaps::default();
        let cell_tables = CellTables {
            strings: &strings,
            formulas: &formulas,
            formula_errors: &formula_errors,
            rich_text: &rich_text,
            comments: Some(&comments),
            formula_references: &formula_references,
        };
        let mut table = Table::with_dimensions("comments", 1, 2)?;
        let mut cell_budget = CellBudget::new();
        let limits = SemanticLimits::default()
            .with_projection_limits(3, 8)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut projection_budget = ProjectionBudget::new(limits);

        TableDataExtractor::parse_tile_row(
            &row,
            0,
            2,
            1,
            2,
            &mut cell_budget,
            &cell_tables,
            &mut projection_budget,
            &mut table,
        )?;

        assert_eq!(table.comment_count(), 2);
        assert_eq!(
            table.get_comment(0, 0).map(|comment| comment.text.as_str()),
            Some("copy")
        );
        assert_eq!(
            table
                .get_comment(0, 1)
                .map(|comment| comment.reply_ids.len()),
            Some(1)
        );
        assert_eq!(projection_budget.materialized_cells, 2);
        assert_eq!(projection_budget.output_text_bytes, 8);
        assert_eq!(projection_budget.references, 2);
        let error = projection_budget
            .charge_materialized_cells(2)
            .expect_err("cell and comment materializations must share the bound");
        assert!(matches!(
            error,
            Error::SemanticLimit {
                kind: SemanticLimitKind::MaterializedCells,
                observed: 4,
                maximum: 3,
                ..
            }
        ));
        Ok(())
    }

    #[test]
    fn comment_text_budget_rejects_a_later_copy_before_map_insertion() -> super::Result<()> {
        let mut cell = BncCell::minimal();
        cell.set_comment_identifier(Some(7));
        let cell = cell.encode();
        let cell_end = u16::try_from(cell.len())
            .map_err(|_| Error::InvalidFormat("test comment cell is too large".to_owned()))?;
        let mut storage = cell.clone();
        storage.extend_from_slice(&cell);
        let row_source = tile_row(
            0,
            2,
            Vec::new(),
            Vec::new(),
            Some(storage),
            Some(vec![
                0,
                0,
                cell_end.to_le_bytes()[0],
                cell_end.to_le_bytes()[1],
            ]),
            Some(false),
        )
        .encode_to_vec();
        let options = numbers_table_cell_storage_codec::DecodeOptions::new(
            row_source.len().max(1),
            usize::MAX,
            usize::MAX,
            16,
            usize::MAX,
            usize::MAX,
        );
        let row = numbers_table_cell_storage_codec::decode_tile_row_info(&row_source, options)
            .map_err(|error| Error::InvalidFormat(format!("comment row failed: {error:?}")))?;
        let strings: Box<[(u32, String)]> = Box::default();
        let formulas: Box<[(u32, FormulaArchiveBytes)]> = Box::default();
        let formula_errors: Box<[(u32, String)]> = Box::default();
        let rich_text: Box<[(u32, String)]> = Box::default();
        let comments: Box<[(u32, Comment)]> = vec![(
            7,
            Comment {
                text: "copy".to_owned(),
                creation_date_seconds: None,
                author_id: None,
                reply_ids: vec![StorageId::new(9).unwrap()].into_boxed_slice(),
                storage_uuid: None,
            },
        )]
        .into_boxed_slice();
        let formula_references = FormulaReferenceMaps::default();
        let cell_tables = CellTables {
            strings: &strings,
            formulas: &formulas,
            formula_errors: &formula_errors,
            rich_text: &rich_text,
            comments: Some(&comments),
            formula_references: &formula_references,
        };
        let mut table = Table::with_dimensions("comments", 1, 2)?;
        let mut cell_budget = CellBudget::new();
        let limits = SemanticLimits::default()
            .with_projection_limits(16, 5)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut projection_budget = ProjectionBudget::new(limits);
        let error = TableDataExtractor::parse_tile_row(
            &row,
            0,
            2,
            1,
            2,
            &mut cell_budget,
            &cell_tables,
            &mut projection_budget,
            &mut table,
        )
        .expect_err("second owned comment copy must exceed text budget");

        assert!(matches!(
            error,
            Error::SemanticLimit {
                kind: SemanticLimitKind::OutputTextBytes,
                observed: 8,
                maximum: 5,
                ..
            }
        ));
        assert_eq!(table.comment_count(), 1);
        assert_eq!(projection_budget.materialized_cells, 1);
        assert_eq!(projection_budget.output_text_bytes, 4);
        assert_eq!(projection_budget.references, 1);
        Ok(())
    }

    #[test]
    fn tile_row_projection_is_atomic_when_a_later_row_is_malformed() -> super::Result<()> {
        let valid = tile_row(0, 1, vec![0; 8], vec![0, 0], None, None, None).encode_to_vec();
        let mut malformed =
            tile_row(1, 1, vec![0; 8], vec![0, 0], None, None, None).encode_to_vec();
        malformed.pop();
        let mut source = Vec::new();
        append_varint_field(&mut source, 1, 1)?;
        append_varint_field(&mut source, 2, 1)?;
        append_varint_field(&mut source, 3, 2)?;
        append_varint_field(&mut source, 4, 2)?;
        append_length_delimited_field(&mut source, 5, &valid)?;
        append_length_delimited_field(&mut source, 5, &malformed)?;

        let options = numbers_table_cell_storage_codec::DecodeOptions::new(
            source.len().max(1),
            usize::MAX,
            usize::MAX,
            16,
            usize::MAX,
            usize::MAX,
        );
        assert!(
            numbers_table_cell_storage_codec::decode_tile_with_visitor(&source, options, &mut ())
                .is_err()
        );
        Ok(())
    }

    #[test]
    fn later_malformed_wire_overrides_prior_semantic_tile_error() -> super::Result<()> {
        let mut invalid_cell = vec![0; 8];
        invalid_cell[2] = 99;
        let first = tile_row(0, 1, invalid_cell, vec![0, 0], None, None, None).encode_to_vec();
        let mut malformed =
            tile_row(1, 1, vec![0; 8], vec![0, 0], None, None, None).encode_to_vec();
        malformed.pop();
        let mut source = Vec::new();
        append_varint_field(&mut source, 1, 1)?;
        append_varint_field(&mut source, 2, 1)?;
        append_varint_field(&mut source, 3, 2)?;
        append_varint_field(&mut source, 4, 2)?;
        append_length_delimited_field(&mut source, 5, &first)?;
        append_length_delimited_field(&mut source, 5, &malformed)?;

        let (decode_result, table, _budget, semantic_error, materialized_cells) =
            run_tile_visitor(&source, 1, SemanticLimits::default())?;
        assert!(decode_result.is_err());
        assert!(
            matches!(semantic_error, Some(Error::ParseError(message)) if message.contains("Unsupported Numbers pre-BNC cell type 99"))
        );
        assert_eq!(materialized_cells, 1);
        assert_eq!(table.cell_count(), 0);
        Ok(())
    }

    #[test]
    fn aggregate_cell_limit_precedes_retained_semantic_error_without_row_growth()
    -> super::Result<()> {
        let mut invalid_cell = vec![0; 8];
        invalid_cell[2] = 99;
        let first = tile_row(0, 1, invalid_cell, vec![0, 0], None, None, None).encode_to_vec();
        let second = tile_row(1, 3, vec![0; 8], vec![0, 0], None, None, None).encode_to_vec();
        let source = tile_source(
            vec![
                tst::TileRowInfo::decode(first.as_slice()).map_err(Error::protobuf)?,
                tst::TileRowInfo::decode(second.as_slice()).map_err(Error::protobuf)?,
            ],
            4,
            2,
        );
        let limits = SemanticLimits::default()
            .with_projection_limits(1, SemanticLimits::MAX_OUTPUT_TEXT_BYTES)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let (decode_result, table, mut budget, semantic_error, materialized_cells) =
            run_tile_visitor(&source, 1, limits)?;
        let report = decode_result.map_err(|error| Error::InvalidFormat(format!("{error:?}")))?;
        assert!(matches!(semantic_error, Some(Error::ParseError(_))));
        assert_eq!(materialized_cells, 4);
        assert_eq!(table.cell_count(), 0);
        budget.charge_decode_report(report)?;
        let error = budget
            .charge_materialized_cells(materialized_cells)
            .err()
            .ok_or_else(|| {
                Error::InvalidFormat("aggregate cell limit was not enforced".to_owned())
            })?;
        assert!(matches!(
            error,
            Error::SemanticLimit {
                kind: SemanticLimitKind::MaterializedCells,
                observed: 4,
                maximum: 1,
                path: SemanticPath::StructuredTables,
            }
        ));
        Ok(())
    }

    #[test]
    fn tile_rows_ignore_declared_counts_when_records_are_valid() -> super::Result<()> {
        let source = tile_source(
            vec![tile_row(0, 1, vec![0; 8], vec![0, 0], None, None, None)],
            99,
            77,
        );
        let projected = decode_tile_scalars(&source);
        assert_eq!(projected.num_cells(), 99);
        assert_eq!(projected.num_rows(), 77);
        let table = parse_projected_rows(&source, 1)?;
        assert_eq!(table.get_cell(0, 0), Some(&CellValue::Empty));
        Ok(())
    }

    #[test]
    fn numeric_type_nine_bnc_cell_is_not_misclassified_as_empty_rich_text() {
        let strings: Box<[(u32, String)]> = Box::default();
        let formulas: Box<[(u32, FormulaArchiveBytes)]> = Box::default();
        let formula_errors: Box<[(u32, String)]> = Box::default();
        let rich_text: Box<[(u32, String)]> = Box::default();
        let comments: Box<[(u32, Comment)]> = Box::default();
        let formula_references = FormulaReferenceMaps::default();
        let tables = CellTables {
            strings: &strings,
            formulas: &formulas,
            formula_errors: &formula_errors,
            rich_text: &rich_text,
            comments: Some(&comments),
            formula_references: &formula_references,
        };

        let mut encoded = vec![5, 9, 0, 0, 0, 0, 0, 0];
        encoded.extend_from_slice(&TEST_DECIMAL_FLAG.to_le_bytes());
        encoded.extend_from_slice(
            &decimal128_le(-1_234.5)
                .unwrap_or_else(|error| panic!("test decimal did not encode: {error}")),
        );
        let round_tripped = BncCell::parse(&encoded)
            .unwrap_or_else(|error| panic!("type-nine cell did not parse: {error}"))
            .encode();

        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let parsed =
            TableDataExtractor::parse_bnc_cell(&round_tripped, &tables, &mut budget, 2, 3, 10, 10)
                .unwrap_or_else(|error| panic!("type-nine cell did not extract: {error}"));
        let CellValue::Number(value) = parsed.value else {
            panic!("type-nine decimal was not extracted as a number");
        };
        assert_eq!(value.get(), -1_234.5);
    }

    #[test]
    fn arena_formula_renderer_matches_reference_output() -> super::Result<()> {
        let mut string = formula_node(AstNodeType::StringNode);
        string.ast_string_node_string = Some("a\"b".to_owned());
        let input = formula(vec![
            number_node(1.0),
            number_node(2.0),
            formula_node(AstNodeType::AdditionNode),
            string,
            formula_node(AstNodeType::ConcatenationNode),
        ]);
        let references = FormulaReferenceMaps::default();
        let expected =
            TableDataExtractor::extract_formula_string_reference(&input, 0, 0, &references)?;
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let actual = render_formula(&input, 0, 0, &references, &mut budget)?;
        assert_eq!(actual, expected);
        assert_eq!(actual, "=((1+2)&\"a\"\"b\")");
        Ok(())
    }

    #[test]
    fn formula_row_one_based_conversion_preserves_u32_max() -> super::Result<()> {
        let references = FormulaReferenceMaps::default();
        let nodes = [
            AstNodeArchive {
                ast_node_type: AstNodeType::LocalCellReferenceNode as i32,
                ast_local_cell_reference_node_reference: Some(AstLocalCellReferenceNodeArchive {
                    row_handle: u32::MAX,
                    column_handle: 0,
                    row_is_sticky: 0,
                    column_is_sticky: 0,
                }),
                ..Default::default()
            },
            AstNodeArchive {
                ast_node_type: AstNodeType::CrossTableCellReferenceNode as i32,
                ast_cross_table_cell_reference_node_reference: Some(
                    AstCrossTableCellReferenceNodeArchive {
                        row_handle: u32::MAX,
                        column_handle: 0,
                        row_is_sticky: 0,
                        column_is_sticky: 0,
                        table_id: Default::default(),
                        ..Default::default()
                    },
                ),
                ..Default::default()
            },
        ];

        for (node, expected) in nodes
            .into_iter()
            .zip(["=A4294967296", "=Table::A4294967296"])
        {
            let input = formula(vec![node]);
            let mut budget = ProjectionBudget::new(SemanticLimits::default());
            assert_eq!(
                render_formula(&input, 0, 0, &references, &mut budget)?,
                expected
            );
        }

        let colon = formula(vec![AstNodeArchive {
            ast_node_type: AstNodeType::ColonTractNode as i32,
            ast_colon_tract: Some(AstColonTractArchive {
                absolute_row: vec![AstColonTractAbsoluteRangeArchive {
                    range_begin: u32::MAX,
                    range_end: Some(u32::MAX),
                }],
                ..Default::default()
            }),
            ast_sticky_bits: Some(AstStickyBits {
                begin_row_is_absolute: true,
                begin_column_is_absolute: false,
                end_row_is_absolute: true,
                end_column_is_absolute: false,
            }),
            ..Default::default()
        }]);
        let expected = "=$4294967296:$4294967296";
        assert_eq!(
            TableDataExtractor::extract_formula_string_reference(&colon, 0, 0, &references,)?,
            expected
        );
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert_eq!(
            render_formula(&colon, 0, 0, &references, &mut budget)?,
            expected
        );
        Ok(())
    }

    #[test]
    fn skewed_concatenation_uses_linear_arena_storage() -> super::Result<()> {
        const VALUES: usize = 4_096;
        let mut nodes = Vec::new();
        nodes
            .try_reserve_exact(VALUES * 2 - 1)
            .map_err(|_| super::allocation_error("test formula nodes", VALUES * 2 - 1))?;
        nodes.push(number_node(1.0));
        for _ in 1..VALUES {
            nodes.push(number_node(1.0));
            nodes.push(formula_node(AstNodeType::ConcatenationNode));
        }
        let input = formula(nodes);
        let references = FormulaReferenceMaps::default();
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let mut renderer = FormulaRenderer::default();
        let root = render_formula_ast_array(
            &input.ast_node_array,
            0,
            0,
            &references,
            &mut budget,
            &mut renderer,
            1,
            false,
        )?
        .unwrap_or_else(|| panic!("skewed formula did not produce an expression"));
        assert_eq!(renderer.nodes.len(), VALUES * 2 - 1);
        assert_eq!(renderer.parts.len(), VALUES + (VALUES - 1) * 5);
        let output = renderer.render(root, &mut budget)?;
        assert_eq!(output.len(), VALUES * 4 - 2);
        Ok(())
    }

    #[test]
    fn formula_work_text_and_depth_limits_are_inclusive() -> super::Result<()> {
        let references = FormulaReferenceMaps::default();
        let input = formula(vec![
            number_node(1.0),
            number_node(2.0),
            formula_node(AstNodeType::AdditionNode),
        ]);

        let exact_limits = SemanticLimits::default()
            .with_formula_render_limits(3, 64)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?
            .with_projection_limits(crate::MAX_MATERIALIZED_CELLS, 6)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut exact = ProjectionBudget::new(exact_limits);
        assert_eq!(
            render_formula(&input, 0, 0, &references, &mut exact)?,
            "=(1+2)"
        );

        let tight_work = SemanticLimits::default()
            .with_formula_render_limits(2, 64)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut work_budget = ProjectionBudget::new(tight_work);
        assert!(matches!(
            render_formula(&input, 0, 0, &references, &mut work_budget),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaRenderWork,
                observed: 3,
                maximum: 2,
                ..
            })
        ));

        let tight_text = SemanticLimits::default()
            .with_projection_limits(crate::MAX_MATERIALIZED_CELLS, 5)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut text_budget = ProjectionBudget::new(tight_text);
        assert!(matches!(
            render_formula(&input, 0, 0, &references, &mut text_budget),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::OutputTextBytes,
                observed: 6,
                maximum: 5,
                ..
            })
        ));

        let mut nested = tsce::AstNodeArrayArchive {
            ast_node: vec![number_node(1.0)],
        };
        for _ in 1..=3 {
            let mut thunk = formula_node(AstNodeType::ThunkNode);
            thunk.ast_thunk_node_array = Some(nested);
            nested = tsce::AstNodeArrayArchive {
                ast_node: vec![thunk],
            };
        }
        let nested = tsce::FormulaArchive {
            ast_node_array: nested,
            ..Default::default()
        };
        let depth_limits = SemanticLimits::default()
            .with_formula_render_limits(4, 3)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut depth_budget = ProjectionBudget::new(depth_limits);
        assert!(matches!(
            render_formula(&nested, 0, 0, &references, &mut depth_budget),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaRenderDepth,
                observed: 4,
                maximum: 3,
                ..
            })
        ));
        Ok(())
    }

    #[test]
    fn projection_budget_is_package_aggregate() -> super::Result<()> {
        let limits = SemanticLimits::default()
            .with_projection_limits(3, 5)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut budget = ProjectionBudget::new(limits);
        budget.charge_materialized_cells(2)?;
        budget.charge_materialized_cells(1)?;
        assert!(matches!(
            budget.charge_materialized_cells(1),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::MaterializedCells,
                observed: 4,
                maximum: 3,
                ..
            })
        ));
        budget.charge_output_text(2)?;
        budget.charge_output_text(3)?;
        assert!(matches!(
            budget.charge_output_text(1),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::OutputTextBytes,
                observed: 6,
                maximum: 5,
                ..
            })
        ));
        Ok(())
    }

    #[test]
    fn rooted_table_call_graph_enforces_one_cumulative_reference_budget() -> super::Result<()> {
        fn extract(max_references: usize) -> super::Result<usize> {
            let package = Package::open(native_fixture())?;
            let limits = SemanticLimits::new(
                SemanticLimits::MAX_OBJECTS,
                SemanticLimits::MAX_SHEETS,
                SemanticLimits::MAX_TABLES,
                max_references,
            )
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
            let extractor =
                TableDataExtractor::new(&package.state.components, &package.state.index, limits)
                    .without_comments();

            // The fixture has one rooted sheet, drawable, and TableInfo edge.
            // Production charges these before entering the table-model and
            // sidecar call graph; mirror that exact prefix here.
            extractor.charge_references(3)?;
            let entry = package
                .state
                .index
                .iter_entries_by_type(super::TABLE_MODEL_MESSAGE_TYPE)
                .next()
                .ok_or_else(|| Error::InvalidFormat("native table model is missing".to_owned()))?;
            let object = package
                .state
                .index
                .resolve_ref(&package.state.components, entry.id())?
                .ok_or_else(|| Error::InvalidFormat("native table model is missing".to_owned()))?;
            extractor.extract_reachable_table_from_object(
                &object,
                SemanticPath::Drawable { sheet: 0, index: 0 },
            )?;
            Ok(extractor.projection_budget.borrow().references)
        }

        let exact = extract(SemanticLimits::MAX_REFERENCES)?;
        assert!(
            exact > 3,
            "table/list/rich/tile projection must charge references"
        );
        assert_eq!(extract(exact)?, exact);
        let tight = extract(exact - 1);
        assert!(
            matches!(
                tight,
                Err(Error::SemanticLimit {
                    kind: SemanticLimitKind::References,
                    observed,
                    maximum,
                    ..
                }) if observed == exact && maximum == exact - 1
            ),
            "unexpected tight cumulative result: {tight:?}; exact={exact}"
        );
        Ok(())
    }

    #[test]
    fn strict_selected_list_rejects_conflicting_entry_payloads_in_both_modes() {
        let mut conflicting = string_list_entry(1, "selected");
        conflicting.formula = Some(formula(Vec::new()));
        for document_projection in [false, true] {
            let result = with_list_extractor(
                vec![
                    archive_object(
                        90,
                        vec![list_message(
                            tst::table_data_list::ListType::String,
                            vec![conflicting.clone()],
                            Vec::new(),
                        )],
                    )
                    .expect("list object"),
                ],
                document_projection,
                |extractor| load_string_list(extractor, 90).map(|_| ()),
            );
            assert!(
                matches!(&result, Err(Error::InvalidFormat(message)) if message.contains("no selected payload")),
                "unexpected selected-list result: {result:?}"
            );
        }
    }

    #[test]
    fn wrong_list_candidates_charge_work_not_references_or_text_in_both_modes() {
        for document_projection in [false, true] {
            let result = with_list_extractor(
                vec![
                    archive_object(
                        90,
                        vec![list_message(
                            tst::table_data_list::ListType::Formula,
                            Vec::new(),
                            Vec::new(),
                        )],
                    )
                    .expect("list object"),
                ],
                document_projection,
                |extractor| load_string_list(extractor, 90),
            );
            let error = result.expect_err("wrong list must be rejected");
            assert!(matches!(error, Error::InvalidFormat(_)));
        }
    }

    #[test]
    fn wrong_list_candidate_does_not_consume_exhausted_admission_budgets() -> super::Result<()> {
        let mut wrong = list_message(
            tst::table_data_list::ListType::Formula,
            vec![tst::table_data_list::ListEntry {
                key: 1,
                refcount: 1,
                reference: Some(reference(99)),
                ..Default::default()
            }],
            Vec::new(),
        );
        let mut deep_prefix = nested_unknown_group_prefix(60);
        deep_prefix.extend_from_slice(&wrong.data);
        wrong.data = deep_prefix;
        let selected = list_message(
            tst::table_data_list::ListType::String,
            Vec::new(),
            Vec::new(),
        );

        for document_projection in [false, true] {
            let result = with_list_extractor(
                vec![archive_object(90, vec![wrong.clone(), selected.clone()])?],
                document_projection,
                |extractor| {
                    let mut budget = ProjectionBudget::new(SemanticLimits::default());
                    budget.references = budget.max_references;
                    budget.staging_text_bytes = DEFAULT_MAX_TEXT_BYTES;
                    let table = load_string_list_with_budget(extractor, 90, &mut budget)?;
                    assert!(table.is_empty());
                    assert_eq!(budget.references, budget.max_references);
                    assert_eq!(budget.staging_text_bytes, DEFAULT_MAX_TEXT_BYTES);
                    assert!(budget.payload_work > 0);
                    Ok(())
                },
            );
            result?;
        }
        Ok(())
    }

    #[test]
    fn duplicate_roots_and_segments_do_not_admit_reference_or_text_budget() -> super::Result<()> {
        let duplicate_entry = tst::table_data_list::ListEntry {
            key: 2,
            refcount: 1,
            string: Some("duplicate".to_owned()),
            reference: Some(reference(99)),
            ..Default::default()
        };
        for document_projection in [false, true] {
            let result = with_list_extractor(
                vec![archive_object(
                    90,
                    vec![
                        list_message(
                            tst::table_data_list::ListType::String,
                            vec![string_list_entry(1, "admitted")],
                            Vec::new(),
                        ),
                        list_message(
                            tst::table_data_list::ListType::String,
                            vec![duplicate_entry.clone()],
                            Vec::new(),
                        ),
                    ],
                )?],
                document_projection,
                |extractor| {
                    let mut budget = ProjectionBudget::new(SemanticLimits::default());
                    let result = load_string_list_with_budget(extractor, 90, &mut budget);
                    assert!(matches!(result, Err(Error::InvalidFormat(_))));
                    assert_eq!(budget.references, 0);
                    assert_eq!(budget.staging_text_bytes, "admitted".len());
                    Ok(())
                },
            );
            result?;

            let result = with_list_extractor(
                vec![
                    archive_object(
                        90,
                        vec![list_message(
                            tst::table_data_list::ListType::String,
                            Vec::new(),
                            vec![91],
                        )],
                    )?,
                    archive_object(
                        91,
                        vec![
                            segment_message(
                                tst::table_data_list::ListType::String,
                                1,
                                2,
                                vec![string_list_entry(1, "admitted")],
                            ),
                            segment_message(
                                tst::table_data_list::ListType::String,
                                1,
                                2,
                                vec![duplicate_entry.clone()],
                            ),
                        ],
                    )?,
                ],
                document_projection,
                |extractor| {
                    let mut budget = ProjectionBudget::new(SemanticLimits::default());
                    let result = load_string_list_with_budget(extractor, 90, &mut budget);
                    assert!(matches!(result, Err(Error::InvalidFormat(_))));
                    assert_eq!(budget.references, 1);
                    assert_eq!(budget.staging_text_bytes, "admitted".len());
                    Ok(())
                },
            );
            result?;
        }
        Ok(())
    }

    #[test]
    fn structural_entry_errors_win_after_a_semantic_conversion_error() -> super::Result<()> {
        for document_projection in [false, true] {
            let result =
                with_list_extractor(
                    vec![archive_object(
                        90,
                        vec![list_message(
                            tst::table_data_list::ListType::String,
                            vec![
                                string_list_entry(1, "first"),
                                string_list_entry(2, "later"),
                                string_list_entry(2, "later-duplicate"),
                            ],
                            Vec::new(),
                        )],
                    )?],
                    document_projection,
                    |extractor| {
                        let mut budget = ProjectionBudget::new(SemanticLimits::default());
                        let mut calls = 0usize;
                        let mut converter =
                        |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
                         _budget: &mut ProjectionBudget| {
                            calls += 1;
                            if entry.key() == 1 {
                                return Err(Error::InvalidFormat(
                                    "synthetic semantic conversion failure".to_owned(),
                                ));
                            }
                            Ok(entry.string_value().unwrap_or_default().to_owned())
                        };
                        let result = extractor.load_table_data_list_entries(
                            90,
                            tst::table_data_list::ListType::String,
                            &mut budget,
                            &mut converter,
                        );
                        assert!(matches!(
                            result,
                            Err(Error::InvalidFormat(message))
                                if message.contains("duplicate keys")
                        ));
                        assert_eq!(calls, 1);
                        Ok(())
                    },
                );
            result?;
        }
        Ok(())
    }

    #[test]
    fn table_info_precharge_is_exactly_the_codec_four_pass_bound() -> super::Result<()> {
        let source_len = 7;
        let mut exact = FormulaReferenceBudget::new(
            crate::MAX_REFERENCES,
            source_len * 4,
            source_len,
            DEFAULT_MAX_TEXT_BYTES,
        );
        exact.charge_table_info_payload(source_len)?;
        assert_eq!(exact.wire_bytes, source_len);
        assert_eq!(exact.work_items, source_len * 4);

        let mut one_short = FormulaReferenceBudget::new(
            crate::MAX_REFERENCES,
            source_len * 4 - 1,
            source_len,
            DEFAULT_MAX_TEXT_BYTES,
        );
        assert!(matches!(
            one_short.charge_table_info_payload(source_len),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaWork,
                observed,
                maximum,
                path: SemanticPath::StructuredTables,
            }) if observed == source_len * 4 && maximum == source_len * 4 - 1
        ));
        assert_eq!(one_short.wire_bytes, source_len);
        assert_eq!(one_short.work_items, source_len);
        Ok(())
    }

    #[test]
    fn formula_owner_preflight_charges_nested_work_on_success_and_error() -> super::Result<()> {
        let uuid = formula_owner_uuid(0x11, 0x22);
        let table = reference(7).encode_to_vec();
        let source = formula_owner_wire(&uuid, &table);

        let mut valid_work = 0;
        let (_, _, report) = preflight_formula_owner(&source, &mut valid_work)?;
        assert!(valid_work >= source.len() + uuid.len() + table.len());
        assert!(valid_work >= report.scanned_bytes() + report.fields());

        // The UUID payload is scanned before its unknown field is rejected.
        // The owner candidate remains best-effort, but its nested scan must
        // still consume bounded work.
        let mut malformed_uuid = uuid.clone();
        append_varint_field(&mut malformed_uuid, 3, 1)?;
        let malformed_source = formula_owner_wire(&malformed_uuid, &table);
        let mut malformed_work = 0;
        assert!(preflight_formula_owner(&malformed_source, &mut malformed_work).is_err());
        assert!(malformed_work >= malformed_source.len() + malformed_uuid.len());

        // Local-reference semantic failure is also charged before the helper
        // is invoked, even though the compatibility caller discards it.
        let malformed_table = tsp::Reference {
            identifier: 7,
            deprecated_is_external: Some(true),
            ..Default::default()
        }
        .encode_to_vec();
        let malformed_source = formula_owner_wire(&uuid, &malformed_table);
        let mut local_error_work = 0;
        assert!(preflight_formula_owner(&malformed_source, &mut local_error_work).is_err());
        assert!(local_error_work >= malformed_source.len() + uuid.len() + malformed_table.len());
        Ok(())
    }

    #[test]
    fn formula_owner_preflight_rejects_noncanonical_known_framing() -> super::Result<()> {
        let uuid = formula_owner_uuid(0x11, 0x22);
        let table = reference(7).encode_to_vec();

        let mut overlong_owner_key = vec![0x8a, 0x00, uuid.len() as u8];
        overlong_owner_key.extend_from_slice(&uuid);
        append_length_delimited_field(&mut overlong_owner_key, 11, &table)?;
        let mut work = 0;
        assert!(preflight_formula_owner(&overlong_owner_key, &mut work).is_err());

        let mut overlong_table_length = Vec::new();
        append_length_delimited_field(&mut overlong_table_length, 1, &uuid)?;
        overlong_table_length.extend_from_slice(&[0x5a, 0x80 | table.len() as u8, 0x00]);
        overlong_table_length.extend_from_slice(&table);
        let mut work = 0;
        assert!(preflight_formula_owner(&overlong_table_length, &mut work).is_err());

        // Replace the canonical lower/upper UUID payload with an overlong key
        // on the lower field while retaining a canonical upper field.
        let overlong_uuid_key = vec![0x88, 0x00, 0x11, 0x10, 0x22];
        let source = formula_owner_wire(&overlong_uuid_key, &table);
        let mut work = 0;
        assert!(preflight_formula_owner(&source, &mut work).is_err());
        Ok(())
    }

    #[test]
    fn formula_table_name_rejects_duplicate_valid_candidates_order_independently()
    -> super::Result<()> {
        let model_one = archive_object(
            90,
            vec![RawMessage {
                type_: super::TABLE_MODEL_MESSAGE_TYPE,
                data: legacy_model("one", 1).encode_to_vec(),
            }],
        )?;
        let model_two = archive_object(
            91,
            vec![RawMessage {
                type_: super::TABLE_MODEL_MESSAGE_TYPE,
                data: legacy_model("two", 1).encode_to_vec(),
            }],
        )?;
        let bytes = compatibility_package(vec![model_one, model_two])?;
        let components = Components::from_bytes(&bytes, Limits::default())?;
        let index = Index::from_components(&components, SemanticLimits::MAX_OBJECTS)?;
        let table_info = |model_id| RawMessage {
            type_: 6_000,
            data: tst::TableInfoArchive {
                super_: tsd::DrawableArchive::default(),
                table_model: reference(model_id),
                ..Default::default()
            }
            .encode_to_vec(),
        };

        for messages in [
            vec![table_info(90), table_info(91)],
            vec![table_info(91), table_info(90)],
        ] {
            let mut budget = FormulaReferenceBudget::new(
                crate::MAX_REFERENCES,
                MAX_FORMULA_WORK,
                MAX_FORMULA_WIRE_BYTES,
                DEFAULT_MAX_TEXT_BYTES,
            );
            assert!(matches!(
                formula_table_name(&components, &index, &messages, &mut budget),
                Err(Error::InvalidFormat(message))
                    if message.contains("duplicate valid TableInfo candidates")
            ));
        }

        let valid = table_info(90);
        let malformed = RawMessage {
            type_: 6_003,
            data: vec![0xff],
        };
        let mut budget = FormulaReferenceBudget::new(
            crate::MAX_REFERENCES,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
        let name = formula_table_name(
            &components,
            &index,
            &[malformed.clone(), valid.clone()],
            &mut budget,
        )?;
        assert_eq!(name.as_deref(), Some("one"));
        let mut budget = FormulaReferenceBudget::new(
            crate::MAX_REFERENCES,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
        let name = formula_table_name(&components, &index, &[valid, malformed], &mut budget)?;
        assert_eq!(name.as_deref(), Some("one"));
        Ok(())
    }

    #[test]
    fn failed_formula_map_publishes_only_monotonic_work_and_wire() -> super::Result<()> {
        let mut published = ProjectionBudget::new(SemanticLimits::default());
        published.charge_output_text(3)?;
        published.charge_staging_text(2)?;

        let mut failed_map = FormulaReferenceBudget::new(
            crate::MAX_REFERENCES,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
        failed_map.charge_wire_bytes(5)?;
        failed_map.charge_text(11)?;
        failed_map.charge_retained_entry()?;

        let mut rejected = published;
        rejected.retain_formula_map_cost(failed_map.work_items, failed_map.wire_bytes);
        rejected.charge_output_text(19)?;
        rejected.charge_staging_text(13)?;
        rejected.charge_references(1)?;
        published.commit_attempt(rejected, false);

        assert_eq!(published.payload_work, failed_map.work_items);
        assert_eq!(published.formula_wire_bytes, failed_map.wire_bytes);
        assert_eq!(published.output_text_bytes, 3);
        assert_eq!(published.staging_text_bytes, 2);
        assert_eq!(published.references, 0);
        Ok(())
    }

    #[test]
    fn rejected_formula_candidate_keeps_wire_work_monotonic() {
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let mut candidate = budget;
        candidate.formula_wire_bytes = 17;
        budget.commit_attempt(candidate, false);
        assert_eq!(budget.formula_wire_bytes, 17);
    }

    #[test]
    fn raw_formula_wire_copy_charges_exact_cost_before_lazy_decode() -> super::Result<()> {
        let source = formula(vec![number_node(1.0)]).encode_to_vec();
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let raw = FormulaArchiveBytes::from_wire(&source, &mut budget)?;

        assert_eq!(raw.bytes.as_ref(), source.as_slice());
        assert_eq!(budget.formula_wire_bytes, source.len());
        assert!(budget.payload_fields > 0);
        assert!(budget.payload_work >= source.len());

        // The generated archive remains a test-only differential oracle; the
        // production cell path traverses the retained bytes through the
        // generated-free compatibility reader.
        let decoded = tsce::FormulaArchive::decode(raw.bytes.as_ref()).map_err(Error::protobuf)?;
        assert_eq!(decoded.ast_node_array.ast_node.len(), 1);
        Ok(())
    }

    #[test]
    fn raw_empty_formula_archive_preserves_default_prost_renderer() -> super::Result<()> {
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let raw = FormulaArchiveBytes::from_wire(&[], &mut budget)?;
        let decoded = tsce::FormulaArchive::decode(raw.bytes.as_ref()).map_err(Error::protobuf)?;

        assert!(decoded.ast_node_array.ast_node.is_empty());
        assert_eq!(
            render_formula(
                &decoded,
                0,
                0,
                &FormulaReferenceMaps::default(),
                &mut budget,
            )?,
            "="
        );
        assert_eq!(budget.formula_wire_bytes, 0);
        Ok(())
    }

    #[test]
    fn scalar_formula_visitor_matches_renderer_for_postfix_arithmetic() -> super::Result<()> {
        let input = formula(vec![
            number_node(1.0),
            number_node(2.0),
            formula_node(AstNodeType::AdditionNode),
        ]);
        let source = input.encode_to_vec();
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let raw = FormulaArchiveBytes::from_wire(&source, &mut budget)?;
        assert!(raw.scalar_visitor_eligible);

        let references = FormulaReferenceMaps::default();
        let actual = TableDataExtractor::extract_formula_string(
            &raw,
            0,
            0,
            10,
            10,
            &references,
            &mut budget,
        )?;
        let mut reference_budget = ProjectionBudget::new(SemanticLimits::default());
        let expected = render_formula(&input, 0, 0, &references, &mut reference_budget)?;
        assert_eq!(actual, expected);
        assert_eq!(actual, "=(1+2)");
        assert_eq!(budget.formula_render_work, 3);
        Ok(())
    }

    #[test]
    fn scalar_formula_visitor_accepts_matching_decimal128_number_sidecars() -> super::Result<()> {
        // Writer-canonical number nodes carry the decimal128 representation
        // beside the fixed64 value. Those fields are metadata for the same
        // scalar and must not force a compatibility traversal.
        let input = formula(vec![AstNodeArchive {
            ast_number_node_decimal_low: Some(75),
            ast_number_node_decimal_high: Some(0x303c_0000_0000_0000),
            ..number_node(0.75)
        }]);
        let source = input.encode_to_vec();
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let raw = FormulaArchiveBytes::from_wire(&source, &mut budget)?;
        assert!(raw.scalar_visitor_eligible);

        let actual = TableDataExtractor::extract_formula_string(
            &raw,
            0,
            0,
            10,
            10,
            &FormulaReferenceMaps::default(),
            &mut budget,
        )?;
        let mut reference_budget = ProjectionBudget::new(SemanticLimits::default());
        let expected = render_formula(
            &input,
            0,
            0,
            &FormulaReferenceMaps::default(),
            &mut reference_budget,
        )?;
        assert_eq!(actual, expected);
        assert_eq!(actual, "=0.75");
        assert_eq!(budget.formula_render_work, 1);
        Ok(())
    }

    #[test]
    fn scalar_formula_visitor_keeps_date_array_and_thunk_fallbacks() -> super::Result<()> {
        for kind in [
            AstNodeType::DateNode,
            AstNodeType::ArrayNode,
            AstNodeType::ThunkNode,
        ] {
            let source = formula(vec![formula_node(kind)]).encode_to_vec();
            let mut budget = ProjectionBudget::new(SemanticLimits::default());
            let raw = FormulaArchiveBytes::from_wire(&source, &mut budget)?;
            assert!(
                !raw.scalar_visitor_eligible,
                "{kind:?} must retain compatibility fallback"
            );
        }
        Ok(())
    }

    #[test]
    fn scalar_formula_visitor_preserves_local_sticky_coordinates() -> super::Result<()> {
        let local = AstNodeArchive {
            ast_node_type: AstNodeType::LocalCellReferenceNode as i32,
            ast_local_cell_reference_node_reference: Some(AstLocalCellReferenceNodeArchive {
                row_handle: 2,
                column_handle: 3,
                row_is_sticky: 1,
                column_is_sticky: 1,
            }),
            ..Default::default()
        };
        let input = formula(vec![local]);
        let source = input.encode_to_vec();
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let raw = FormulaArchiveBytes::from_wire(&source, &mut budget)?;
        assert!(raw.scalar_visitor_eligible);
        let actual = TableDataExtractor::extract_formula_string(
            &raw,
            0,
            0,
            10,
            10,
            &FormulaReferenceMaps::default(),
            &mut budget,
        )?;
        assert_eq!(actual, "=$D$3");
        Ok(())
    }

    #[test]
    fn scalar_formula_visitor_falls_back_for_string_nodes() -> super::Result<()> {
        let mut string = formula_node(AstNodeType::StringNode);
        string.ast_string_node_string = Some("a\"b".to_owned());
        let input = formula(vec![string]);
        let source = input.encode_to_vec();
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let raw = FormulaArchiveBytes::from_wire(&source, &mut budget)?;
        assert!(!raw.scalar_visitor_eligible);
        let actual = TableDataExtractor::extract_formula_string(
            &raw,
            0,
            0,
            10,
            10,
            &FormulaReferenceMaps::default(),
            &mut budget,
        )?;
        assert_eq!(actual, "=\"a\"\"b\"");
        Ok(())
    }

    #[test]
    fn compatibility_renderer_matches_legacy_fallback_node_families() -> super::Result<()> {
        let owner = tsp::CfuuidArchive {
            uuid_w0: Some(1),
            uuid_w1: Some(2),
            uuid_w2: Some(3),
            uuid_w3: Some(4),
            ..Default::default()
        };
        let category_uid = tsp::Uuid {
            lower: 11,
            upper: 22,
        };
        let mut references = FormulaReferenceMaps::default();
        references.owners.insert(
            [1, 2, 3, 4],
            super::FormulaReferenceName {
                sheet: std::sync::Arc::new("Sheet".to_owned()),
                table: std::sync::Arc::new("Table".to_owned()),
            },
        );
        references
            .categories
            .insert([11, 22], "Category".to_owned());

        let mut string = formula_node(AstNodeType::StringNode);
        string.ast_string_node_string = Some("a\"b".to_owned());
        let mut date = formula_node(AstNodeType::DateNode);
        date.ast_date_node_date_num = Some(86_400.0);
        let mut duration = formula_node(AstNodeType::DurationNode);
        duration.ast_duration_node_unit_num = Some(2.5);
        let mut array = formula_node(AstNodeType::ArrayNode);
        array.ast_array_node_num_col = Some(2);
        array.ast_array_node_num_row = Some(1);
        let mut thunk = formula_node(AstNodeType::ThunkNode);
        thunk.ast_thunk_node_array = Some(tsce::AstNodeArrayArchive {
            ast_node: vec![number_node(5.0)],
        });
        let mut unknown_function = formula_node(AstNodeType::UnknownFunctionNode);
        unknown_function.ast_unknown_function_node_string = Some("MYFN".to_owned());
        unknown_function.ast_unknown_function_node_num_args = Some(1);
        let cross = AstNodeArchive {
            ast_node_type: AstNodeType::CrossTableCellReferenceNode as i32,
            ast_cross_table_cell_reference_node_reference: Some(
                AstCrossTableCellReferenceNodeArchive {
                    row_handle: 1,
                    column_handle: 2,
                    row_is_sticky: 0,
                    column_is_sticky: 0,
                    table_id: owner,
                    ..Default::default()
                },
            ),
            ..Default::default()
        };
        let cell = AstNodeArchive {
            ast_node_type: AstNodeType::CellReferenceNode as i32,
            ast_column: Some(AstColumnCoordinateArchive {
                column: 1,
                absolute: Some(false),
            }),
            ast_row: Some(AstRowCoordinateArchive {
                row: 1,
                absolute: Some(false),
            }),
            ..Default::default()
        };
        let category = AstNodeArchive {
            ast_node_type: AstNodeType::CategoryRefNode as i32,
            ast_category_ref: Some(tsce::ast_node_array_archive::AstCategoryReferenceArchive {
                category_ref: tsce::CategoryReferenceArchive {
                    group_by_uid: category_uid,
                    column_uid: category_uid,
                    aggregate_type: 1,
                    group_level: 0,
                    group_uids: Some(tsce::category_reference_archive::CatRefUidList {
                        uid: vec![category_uid],
                    }),
                    ..Default::default()
                },
            }),
            ..Default::default()
        };
        let colon = AstNodeArchive {
            ast_node_type: AstNodeType::ColonTractNode as i32,
            ast_colon_tract: Some(AstColonTractArchive {
                relative_column: vec![
                    tsce::ast_node_array_archive::ast_colon_tract_archive::AstColonTractRelativeRangeArchive {
                        range_begin: 0,
                        range_end: Some(1),
                    },
                ],
                relative_row: vec![
                    tsce::ast_node_array_archive::ast_colon_tract_archive::AstColonTractRelativeRangeArchive {
                        range_begin: 0,
                        range_end: Some(1),
                    },
                ],
                preserve_rectangular: Some(true),
                ..Default::default()
            }),
            ast_sticky_bits: Some(AstStickyBits {
                begin_row_is_absolute: false,
                begin_column_is_absolute: false,
                end_row_is_absolute: false,
                end_column_is_absolute: false,
            }),
            ..Default::default()
        };
        let whole_row = AstNodeArchive {
            ast_node_type: AstNodeType::ColonTractNode as i32,
            ast_colon_tract: Some(AstColonTractArchive {
                absolute_column: vec![AstColonTractAbsoluteRangeArchive {
                    range_begin: i16::MAX as u32,
                    range_end: None,
                }],
                absolute_row: vec![AstColonTractAbsoluteRangeArchive {
                    range_begin: 0,
                    range_end: Some(1),
                }],
                preserve_rectangular: Some(true),
                ..Default::default()
            }),
            ast_sticky_bits: Some(AstStickyBits {
                begin_row_is_absolute: true,
                begin_column_is_absolute: false,
                end_row_is_absolute: true,
                end_column_is_absolute: false,
            }),
            ..Default::default()
        };
        let whole_column = AstNodeArchive {
            ast_node_type: AstNodeType::ColonTractNode as i32,
            ast_colon_tract: Some(AstColonTractArchive {
                absolute_column: vec![AstColonTractAbsoluteRangeArchive {
                    range_begin: 0,
                    range_end: Some(1),
                }],
                absolute_row: vec![AstColonTractAbsoluteRangeArchive {
                    range_begin: i32::MAX as u32,
                    range_end: None,
                }],
                preserve_rectangular: Some(true),
                ..Default::default()
            }),
            ast_sticky_bits: Some(AstStickyBits {
                begin_row_is_absolute: false,
                begin_column_is_absolute: true,
                end_row_is_absolute: false,
                end_column_is_absolute: true,
            }),
            ..Default::default()
        };
        let mut nonfinite_date = formula_node(AstNodeType::DateNode);
        nonfinite_date.ast_date_node_date_num = Some(f64::NAN);
        let unknown_enum = AstNodeArchive {
            ast_node_type: 999,
            ..Default::default()
        };

        let cases = vec![
            ("string", formula(vec![string])),
            ("date", formula(vec![date])),
            ("duration", formula(vec![duration])),
            (
                "array",
                formula(vec![number_node(1.0), number_node(2.0), array]),
            ),
            ("thunk", formula(vec![thunk])),
            (
                "unknown function",
                formula(vec![number_node(1.0), unknown_function]),
            ),
            ("cross-table reference", formula(vec![cross])),
            ("cell reference", formula(vec![cell])),
            ("colon range", formula(vec![colon])),
            ("whole row", formula(vec![whole_row])),
            ("whole column", formula(vec![whole_column])),
            ("nonfinite date", formula(vec![nonfinite_date])),
            (
                "unknown enum",
                formula(vec![number_node(1.0), number_node(2.0), unknown_enum]),
            ),
            ("category", formula(vec![category])),
        ];

        for (name, input) in cases {
            let mut reference_budget = ProjectionBudget::new(SemanticLimits::default());
            let expected = render_formula(&input, 0, 0, &references, &mut reference_budget)?;
            let actual = compatibility_render(&input, 0, 0, 10, 10, &references)?;
            assert_eq!(actual, expected, "fallback parity failed for {name}");
        }
        Ok(())
    }

    #[test]
    fn compatibility_renderer_rejects_malformed_postfix_atomically() -> super::Result<()> {
        let input = formula(vec![formula_node(AstNodeType::AdditionNode)]);
        let source = input.encode_to_vec();
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let raw = FormulaArchiveBytes::from_wire(&source, &mut budget)?;
        let output_before = budget.output_text_bytes;
        let error = TableDataExtractor::extract_formula_string(
            &raw,
            0,
            0,
            10,
            10,
            &FormulaReferenceMaps::default(),
            &mut budget,
        )
        .expect_err("missing postfix operands must fail");
        assert!(matches!(error, Error::ParseError(_)));
        assert_eq!(budget.output_text_bytes, output_before);
        Ok(())
    }

    #[test]
    fn compatibility_renderer_depth_limit_is_exact() -> super::Result<()> {
        let mut thunk = formula_node(AstNodeType::ThunkNode);
        thunk.ast_thunk_node_array = Some(tsce::AstNodeArrayArchive {
            ast_node: vec![number_node(1.0)],
        });
        let input = formula(vec![thunk]);
        let source = input.encode_to_vec();

        let exact_limits = SemanticLimits::default()
            .with_formula_render_limits(2, 2)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut exact_budget = ProjectionBudget::new(exact_limits);
        let exact_raw = FormulaArchiveBytes::from_wire(&source, &mut exact_budget)?;
        let actual = TableDataExtractor::extract_formula_string(
            &exact_raw,
            0,
            0,
            10,
            10,
            &FormulaReferenceMaps::default(),
            &mut exact_budget,
        )?;
        assert_eq!(actual, "=1");

        let tight_limits = SemanticLimits::default()
            .with_formula_render_limits(2, 1)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut tight_budget = ProjectionBudget::new(tight_limits);
        let tight_raw = FormulaArchiveBytes::from_wire(&source, &mut tight_budget)?;
        let error = TableDataExtractor::extract_formula_string(
            &tight_raw,
            0,
            0,
            10,
            10,
            &FormulaReferenceMaps::default(),
            &mut tight_budget,
        )
        .expect_err("nested thunk must exceed the one-level render depth");
        assert!(matches!(
            error,
            Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaRenderDepth,
                observed: 2,
                maximum: 1,
                ..
            }
        ));
        assert_eq!(tight_budget.output_text_bytes, 0);
        Ok(())
    }

    #[test]
    fn scalar_formula_visitor_falls_back_for_opaque_nested_fields() -> super::Result<()> {
        // The strict envelope accepts an unknown field as opaque, matching
        // Prost's forward-compatible behavior.  It is nevertheless outside
        // the generated-free scalar representation and must not be admitted
        // to that route merely because the known number fields are scalar.
        let mut node = number_node(1.25).encode_to_vec();
        append_varint_field(&mut node, 48, 1)?;
        let mut ast = Vec::new();
        append_length_delimited_field(&mut ast, 1, &node)?;
        let mut source = Vec::new();
        append_length_delimited_field(&mut source, 1, &ast)?;

        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let raw = FormulaArchiveBytes::from_wire(&source, &mut budget)?;
        assert!(!raw.scalar_visitor_eligible);

        let actual = TableDataExtractor::extract_formula_string(
            &raw,
            0,
            0,
            10,
            10,
            &FormulaReferenceMaps::default(),
            &mut budget,
        )?;
        let input = formula(vec![number_node(1.25)]);
        let mut reference_budget = ProjectionBudget::new(SemanticLimits::default());
        let expected = render_formula(
            &input,
            0,
            0,
            &FormulaReferenceMaps::default(),
            &mut reference_budget,
        )?;
        assert_eq!(actual, expected);
        assert_eq!(actual, "=1.25");
        Ok(())
    }

    #[test]
    fn scalar_formula_preflight_skips_unknown_function_reservation() -> super::Result<()> {
        let mut function = formula_node(AstNodeType::FunctionNode);
        function.ast_function_node_index = Some(999);
        function.ast_function_node_num_args = Some(1);
        let source = formula(vec![number_node(1.0), function]).encode_to_vec();
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let raw = FormulaArchiveBytes::from_wire(&source, &mut budget)?;

        // The compatibility renderer retains unknown function IDs;
        // do not allocate the scalar visitor only to discover that the strict
        // evaluator rejects this native function.
        assert!(!raw.scalar_visitor_eligible);
        Ok(())
    }

    #[test]
    fn compatibility_formula_fallback_node_guard_is_inclusive_and_one_over() -> super::Result<()> {
        let mut first = formula_node(AstNodeType::StringNode);
        first.ast_string_node_string = Some("a".to_owned());
        let mut second = formula_node(AstNodeType::StringNode);
        second.ast_string_node_string = Some("b".to_owned());
        let source = formula(vec![first, second]).encode_to_vec();

        let mut probe_budget = ProjectionBudget::new(SemanticLimits::default());
        let probe = FormulaArchiveBytes::from_wire(&source, &mut probe_budget)?;
        assert_eq!(probe.scalar_visitor_node_count, 2);

        let exact_limits = SemanticLimits::default()
            .with_formula_render_limits(2, 64)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut exact_budget = ProjectionBudget::new(exact_limits);
        let exact_raw = FormulaArchiveBytes::from_wire(&source, &mut exact_budget)?;
        assert_eq!(
            TableDataExtractor::extract_formula_string(
                &exact_raw,
                0,
                0,
                10,
                10,
                &FormulaReferenceMaps::default(),
                &mut exact_budget,
            )?,
            "=\"b\""
        );
        assert_eq!(exact_budget.formula_render_work, 2);

        let tight_limits = SemanticLimits::default()
            .with_formula_render_limits(1, 64)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut tight_budget = ProjectionBudget::new(tight_limits);
        let tight_raw = FormulaArchiveBytes::from_wire(&source, &mut tight_budget)?;
        let error = TableDataExtractor::extract_formula_string(
            &tight_raw,
            0,
            0,
            10,
            10,
            &FormulaReferenceMaps::default(),
            &mut tight_budget,
        )
        .expect_err("one-over compatibility fallback must fail before staging");
        assert!(matches!(
            error,
            Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaRenderWork,
                observed: 2,
                maximum: 1,
                ..
            }
        ));
        assert_eq!(tight_budget.formula_render_work, 1);
        assert_eq!(tight_budget.output_text_bytes, 0);
        Ok(())
    }

    #[test]
    fn compatibility_formula_fallback_charges_non_ast_repeated_entries_before_decode()
    -> super::Result<()> {
        let uuid = tsp::Uuid { lower: 1, upper: 2 };
        let category = tsce::CategoryReferenceArchive {
            group_by_uid: uuid,
            column_uid: uuid,
            aggregate_type: 1,
            group_level: 0,
            group_uids: Some(tsce::category_reference_archive::CatRefUidList {
                uid: vec![uuid; 3],
            }),
            ..Default::default()
        };
        let node = AstNodeArchive {
            ast_node_type: AstNodeType::CategoryRefNode as i32,
            ast_category_ref: Some(tsce::ast_node_array_archive::AstCategoryReferenceArchive {
                category_ref: category,
            }),
            ..Default::default()
        };
        let source = formula(vec![node]).encode_to_vec();

        let mut probe_budget = ProjectionBudget::new(SemanticLimits::default());
        let probe = FormulaArchiveBytes::from_wire(&source, &mut probe_budget)?;
        assert!(!probe.scalar_visitor_eligible);
        assert_eq!(probe.scalar_visitor_node_count, 1);
        assert_eq!(probe.lazy_traversal_entry_count, 3);

        let exact_limits = SemanticLimits::default()
            .with_formula_render_limits(4, 64)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut exact_budget = ProjectionBudget::new(exact_limits);
        let exact_raw = FormulaArchiveBytes::from_wire(&source, &mut exact_budget)?;
        let pre_render_work = exact_budget.payload_work;
        let actual = TableDataExtractor::extract_formula_string(
            &exact_raw,
            0,
            0,
            10,
            10,
            &FormulaReferenceMaps::default(),
            &mut exact_budget,
        )?;
        assert_eq!(actual, "=#CATEGORY!");
        assert_eq!(exact_budget.formula_render_work, 1);
        assert_eq!(
            exact_budget.payload_work,
            pre_render_work + exact_raw.lazy_traversal_entry_count
        );

        let mut tight_budget = ProjectionBudget::new(SemanticLimits::default());
        let tight_raw = FormulaArchiveBytes::from_wire(&source, &mut tight_budget)?;
        tight_budget.payload_work = MAX_PAYLOAD_WORK - 2;
        let error = TableDataExtractor::extract_formula_string(
            &tight_raw,
            0,
            0,
            10,
            10,
            &FormulaReferenceMaps::default(),
            &mut tight_budget,
        )
        .expect_err("one repeated-entry overage must fail before compatibility staging");
        assert!(matches!(
            error,
            Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaWork,
                ..
            }
        ));
        if let Error::SemanticLimit {
            observed, maximum, ..
        } = error
        {
            assert_eq!(observed, MAX_PAYLOAD_WORK + 1);
            assert_eq!(maximum, MAX_PAYLOAD_WORK);
        } else {
            unreachable!("formula-work limit already matched above");
        }
        assert_eq!(tight_budget.output_text_bytes, 0);
        Ok(())
    }

    #[test]
    fn scalar_formula_visitor_preserves_ignored_and_missing_negation_parity() -> super::Result<()> {
        for node in [
            formula_node(AstNodeType::PlusSignNode),
            formula_node(AstNodeType::NegationNode),
        ] {
            let input = formula(vec![node]);
            let source = input.encode_to_vec();
            let mut budget = ProjectionBudget::new(SemanticLimits::default());
            let raw = FormulaArchiveBytes::from_wire(&source, &mut budget)?;
            assert!(raw.scalar_visitor_eligible);
            let actual = TableDataExtractor::extract_formula_string(
                &raw,
                0,
                0,
                10,
                10,
                &FormulaReferenceMaps::default(),
                &mut budget,
            )?;
            let mut reference_budget = ProjectionBudget::new(SemanticLimits::default());
            let expected = render_formula(
                &input,
                0,
                0,
                &FormulaReferenceMaps::default(),
                &mut reference_budget,
            )?;
            assert_eq!(actual, expected);
            assert_eq!(actual, "=FORMULA()");
        }
        Ok(())
    }

    #[test]
    fn compatibility_stream_matches_prost_for_legacy_edge_shapes() -> super::Result<()> {
        fn render_pair(
            input: tsce::FormulaArchive,
            references: &FormulaReferenceMaps,
        ) -> super::Result<String> {
            let source = input.encode_to_vec();
            let mut probe_budget = ProjectionBudget::new(SemanticLimits::default());
            let raw = FormulaArchiveBytes::from_wire(&source, &mut probe_budget)?;

            let mut expected_budget = ProjectionBudget::new(SemanticLimits::default());
            expected_budget.charge_formula_render_work(raw.scalar_visitor_node_count)?;
            let expected =
                render_formula_precharged(&input, 0, 0, references, &mut expected_budget)?;

            let mut actual_budget = ProjectionBudget::new(SemanticLimits::default());
            actual_budget.charge_formula_render_work(raw.scalar_visitor_node_count)?;
            actual_budget.charge_formula_lazy_work(raw.lazy_traversal_entry_count)?;
            let actual =
                render_formula_compatibility(&raw, 0, 0, 10, 10, references, &mut actual_budget)?;
            assert_eq!(actual, expected);
            Ok(actual)
        }

        // Prost's enum accessor maps an unknown raw value to its default
        // AdditionNode.  The generated-free compatibility path must retain
        // that historical behavior after the strict wire preflight.
        let unknown_type = AstNodeArchive {
            ast_node_type: 999,
            ..Default::default()
        };
        assert_eq!(
            render_pair(
                formula(vec![number_node(1.0), number_node(2.0), unknown_type]),
                &FormulaReferenceMaps::default(),
            )?,
            "=(1+2)"
        );

        // Missing scalar payloads are ignored by the legacy renderer,
        // leaving the nonempty archive's FORMULA() placeholder.
        assert_eq!(
            render_pair(
                formula(vec![formula_node(AstNodeType::NumberNode)]),
                &FormulaReferenceMaps::default(),
            )?,
            "=FORMULA()"
        );

        let empty_thunk = AstNodeArchive {
            ast_node_type: AstNodeType::ThunkNode as i32,
            ast_thunk_node_array: Some(tsce::AstNodeArrayArchive::default()),
            ..Default::default()
        };
        assert_eq!(
            render_pair(formula(vec![empty_thunk]), &FormulaReferenceMaps::default())?,
            "="
        );
        let nonempty_thunk = AstNodeArchive {
            ast_node_type: AstNodeType::ThunkNode as i32,
            ast_thunk_node_array: Some(tsce::AstNodeArrayArchive {
                ast_node: vec![number_node(7.0)],
            }),
            ..Default::default()
        };
        assert_eq!(
            render_pair(
                formula(vec![nonempty_thunk]),
                &FormulaReferenceMaps::default()
            )?,
            "=7"
        );
        let ignored_thunk = AstNodeArchive {
            ast_node_type: AstNodeType::ThunkNode as i32,
            ast_thunk_node_array: Some(tsce::AstNodeArrayArchive {
                ast_node: vec![formula_node(AstNodeType::PlusSignNode)],
            }),
            ..Default::default()
        };
        assert_eq!(
            render_pair(
                formula(vec![ignored_thunk]),
                &FormulaReferenceMaps::default()
            )?,
            "=FORMULA()"
        );

        let sticky_local = AstLocalCellReferenceNodeArchive {
            row_handle: 2,
            column_handle: 3,
            row_is_sticky: 1,
            column_is_sticky: 1,
        };
        let standalone_local = AstNodeArchive {
            ast_node_type: AstNodeType::LocalCellReferenceNode as i32,
            ast_local_cell_reference_node_reference: Some(sticky_local),
            ..Default::default()
        };
        assert_eq!(
            render_pair(
                formula(vec![standalone_local]),
                &FormulaReferenceMaps::default()
            )?,
            "=D3"
        );
        let nested_local = AstNodeArchive {
            ast_node_type: AstNodeType::CellReferenceNode as i32,
            ast_local_cell_reference_node_reference: Some(sticky_local),
            ..Default::default()
        };
        assert_eq!(
            render_pair(
                formula(vec![nested_local]),
                &FormulaReferenceMaps::default()
            )?,
            "=$D$3"
        );

        let category_uid = tsp::Uuid {
            lower: 0x11,
            upper: 0x22,
        };
        let relative_uid = tsp::Uuid {
            lower: 0x33,
            upper: 0x44,
        };
        let category = tsce::CategoryReferenceArchive {
            group_by_uid: category_uid,
            column_uid: category_uid,
            aggregate_type: 1,
            group_level: 0,
            relative_group_uid: Some(relative_uid),
            absolute_group_uid: Some(category_uid),
            group_uids: Some(tsce::category_reference_archive::CatRefUidList {
                uid: vec![relative_uid],
            }),
            ..Default::default()
        };
        let category_node = AstNodeArchive {
            ast_node_type: AstNodeType::CategoryRefNode as i32,
            ast_category_ref: Some(tsce::ast_node_array_archive::AstCategoryReferenceArchive {
                category_ref: category,
            }),
            ..Default::default()
        };
        let mut category_references = FormulaReferenceMaps::default();
        category_references.categories.insert(
            [category_uid.lower, category_uid.upper],
            "Absolute".to_owned(),
        );
        category_references.categories.insert(
            [relative_uid.lower, relative_uid.upper],
            "Relative".to_owned(),
        );
        assert_eq!(
            render_pair(formula(vec![category_node]), &category_references)?,
            "=#CATEGORY![Absolute]"
        );

        // The compatibility renderer intentionally preserves Prost's
        // historical nonfinite fixed64 formatting; the strict scalar path
        // declines the nonfinite value before this fallback is entered.
        assert_eq!(
            render_pair(
                formula(vec![number_node(f64::NAN)]),
                &FormulaReferenceMaps::default(),
            )?,
            "=NaN"
        );
        Ok(())
    }

    #[test]
    fn scalar_formula_visitor_charges_node_bound_before_rendering() -> super::Result<()> {
        let input = formula(vec![number_node(1.0), number_node(2.0)]);
        let source = input.encode_to_vec();
        let limits = SemanticLimits::default()
            .with_formula_render_limits(1, 64)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut budget = ProjectionBudget::new(limits);
        let raw = FormulaArchiveBytes::from_wire(&source, &mut budget)?;
        let error = TableDataExtractor::extract_formula_string(
            &raw,
            0,
            0,
            10,
            10,
            &FormulaReferenceMaps::default(),
            &mut budget,
        )
        .expect_err("scalar node allocation must not precede render-work admission");
        assert!(matches!(
            error,
            Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaRenderWork,
                observed: 2,
                maximum: 1,
                ..
            }
        ));
        Ok(())
    }

    #[test]
    fn raw_nonempty_formula_archive_requires_root_ast_field() -> super::Result<()> {
        // A non-empty message containing only an optional host field was
        // accepted by Prost, but is not a valid FormulaArchive envelope for
        // the strict sidecar preflight.
        let mut source = Vec::new();
        append_varint_field(&mut source, 2, 1)?;
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert!(matches!(
            FormulaArchiveBytes::from_wire(&source, &mut budget),
            Err(Error::MalformedPayload {
                path: SemanticPath::StructuredTables,
            })
        ));
        assert_eq!(budget.formula_wire_bytes, source.len());
        Ok(())
    }

    #[test]
    fn raw_formula_archive_ast_nodes_require_field_one_type() -> super::Result<()> {
        for node in [
            // ASTNodeArchive with no required node type.
            {
                let mut node = Vec::new();
                append_varint_field(&mut node, 2, 1)?;
                node
            },
            // ASTNodeArchive field one has the wrong wire type.
            {
                let mut node = Vec::new();
                append_length_delimited_field(&mut node, 1, b"not-a-varint")?;
                node
            },
        ] {
            let mut array = Vec::new();
            append_length_delimited_field(&mut array, 1, &node)?;
            let mut source = Vec::new();
            append_length_delimited_field(&mut source, 1, &array)?;
            let mut budget = ProjectionBudget::new(SemanticLimits::default());
            assert!(matches!(
                FormulaArchiveBytes::from_wire(&source, &mut budget),
                Err(Error::MalformedPayload {
                    path: SemanticPath::StructuredTables,
                })
            ));
            assert_eq!(budget.formula_wire_bytes, source.len());
        }
        Ok(())
    }

    #[test]
    fn raw_formula_archive_rejects_duplicate_root_and_ast_node_type_fields() -> super::Result<()> {
        let mut duplicate_root = Vec::new();
        append_length_delimited_field(&mut duplicate_root, 1, &[])?;
        append_length_delimited_field(&mut duplicate_root, 1, &[])?;
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert!(FormulaArchiveBytes::from_wire(&duplicate_root, &mut budget).is_err());

        let mut duplicate_type = Vec::new();
        append_varint_field(&mut duplicate_type, 1, 17)?;
        append_varint_field(&mut duplicate_type, 1, 17)?;
        let mut node_array = Vec::new();
        append_length_delimited_field(&mut node_array, 1, &duplicate_type)?;
        let mut duplicate_node = Vec::new();
        append_length_delimited_field(&mut duplicate_node, 1, &node_array)?;
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert!(FormulaArchiveBytes::from_wire(&duplicate_node, &mut budget).is_err());
        Ok(())
    }

    #[test]
    fn raw_formula_archive_rejects_duplicate_known_singular_fields() -> super::Result<()> {
        let mut duplicate_host = formula(vec![number_node(1.0)]).encode_to_vec();
        append_varint_field(&mut duplicate_host, 2, 1)?;
        append_varint_field(&mut duplicate_host, 2, 1)?;
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert!(FormulaArchiveBytes::from_wire(&duplicate_host, &mut budget).is_err());

        let mut duplicate_node = Vec::new();
        append_varint_field(&mut duplicate_node, 1, 17)?;
        append_varint_field(&mut duplicate_node, 2, 1)?;
        append_varint_field(&mut duplicate_node, 2, 1)?;
        let mut node_array = Vec::new();
        append_length_delimited_field(&mut node_array, 1, &duplicate_node)?;
        let mut duplicate_nested = Vec::new();
        append_length_delimited_field(&mut duplicate_nested, 1, &node_array)?;
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert!(FormulaArchiveBytes::from_wire(&duplicate_nested, &mut budget).is_err());

        // Repeated ASTNodeArchive entries remain valid and retain their full
        // generated-renderer semantics.
        let repeated = formula(vec![number_node(1.0), number_node(2.0)]).encode_to_vec();
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert!(FormulaArchiveBytes::from_wire(&repeated, &mut budget).is_ok());
        Ok(())
    }

    #[test]
    fn raw_formula_archive_rejects_missing_required_range_fields() -> super::Result<()> {
        let mut node = Vec::new();
        append_varint_field(&mut node, 1, 67)?;
        let mut empty_range = Vec::new();
        append_length_delimited_field(&mut empty_range, 1, &[])?;
        append_length_delimited_field(&mut node, 40, &empty_range)?;
        let mut node_array = Vec::new();
        append_length_delimited_field(&mut node_array, 1, &node)?;
        let mut source = Vec::new();
        append_length_delimited_field(&mut source, 1, &node_array)?;
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert!(FormulaArchiveBytes::from_wire(&source, &mut budget).is_err());
        Ok(())
    }

    #[test]
    fn raw_formula_archive_recurses_uid_tract_and_category_uuid_paths() -> super::Result<()> {
        // ASTUidTractList.sticky_bits is a required nested message. A scalar
        // at field 2 must not be treated as an opaque unknown field.
        let mut bad_tract_list = Vec::new();
        append_varint_field(&mut bad_tract_list, 2, 1)?;
        let mut tract_node = Vec::new();
        append_varint_field(&mut tract_node, 1, 48)?;
        append_length_delimited_field(&mut tract_node, 38, &bad_tract_list)?;
        let mut tract_array = Vec::new();
        append_length_delimited_field(&mut tract_array, 1, &tract_node)?;
        let mut bad_tract_source = Vec::new();
        append_length_delimited_field(&mut bad_tract_source, 1, &tract_array)?;
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert!(FormulaArchiveBytes::from_wire(&bad_tract_source, &mut budget).is_err());

        // CatRefUidList.uid entries are repeated TSP.UUID messages. A
        // length-delimited UUID with no required lower/upper fields must be
        // rejected after the new recursive descent.
        let mut bad_uuid_list = Vec::new();
        append_length_delimited_field(&mut bad_uuid_list, 1, &[])?;
        let mut category = Vec::new();
        append_length_delimited_field(&mut category, 6, &bad_uuid_list)?;
        let mut category_node = Vec::new();
        append_varint_field(&mut category_node, 1, 66)?;
        append_length_delimited_field(&mut category_node, 39, &{
            let mut wrapper = Vec::new();
            append_length_delimited_field(&mut wrapper, 1, &category)?;
            wrapper
        })?;
        let mut category_array = Vec::new();
        append_length_delimited_field(&mut category_array, 1, &category_node)?;
        let mut bad_category_source = Vec::new();
        append_length_delimited_field(&mut bad_category_source, 1, &category_array)?;
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert!(FormulaArchiveBytes::from_wire(&bad_category_source, &mut budget).is_err());
        Ok(())
    }

    #[test]
    fn raw_formula_archive_checks_recursive_required_messages_and_deep_thunks() -> super::Result<()>
    {
        let uuid = tsp::Uuid { lower: 1, upper: 2 }.encode_to_vec();
        let mut category_ref = Vec::new();
        append_length_delimited_field(&mut category_ref, 1, &uuid)?;
        append_length_delimited_field(&mut category_ref, 2, &uuid)?;
        append_varint_field(&mut category_ref, 3, 1)?;
        append_varint_field(&mut category_ref, 4, 0)?;

        // CategoryReferenceArchive's fields 1..4 are all required. Keep a
        // complete nested category as a control so the recursive checks do
        // not reject a valid category reference.
        let mut category_wrapper = Vec::new();
        append_length_delimited_field(&mut category_wrapper, 1, &category_ref)?;
        let mut category_node = Vec::new();
        append_varint_field(&mut category_node, 1, 66)?;
        append_length_delimited_field(&mut category_node, 39, &category_wrapper)?;
        let mut category_array = Vec::new();
        append_length_delimited_field(&mut category_array, 1, &category_node)?;
        let mut valid_category = Vec::new();
        append_length_delimited_field(&mut valid_category, 1, &category_array)?;
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert!(FormulaArchiveBytes::from_wire(&valid_category, &mut budget).is_ok());

        // The ASTCategoryReferenceArchive wrapper itself has a required
        // category_ref field. An empty wrapper must not be accepted merely
        // because the enclosing AST node and array are structurally valid.
        let mut missing_category_node = Vec::new();
        append_varint_field(&mut missing_category_node, 1, 66)?;
        append_length_delimited_field(&mut missing_category_node, 39, &[])?;
        let mut missing_category_array = Vec::new();
        append_length_delimited_field(&mut missing_category_array, 1, &missing_category_node)?;
        let mut missing_category = Vec::new();
        append_length_delimited_field(&mut missing_category, 1, &missing_category_array)?;
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert!(FormulaArchiveBytes::from_wire(&missing_category, &mut budget).is_err());

        let mut incomplete_category_ref = Vec::new();
        append_length_delimited_field(&mut incomplete_category_ref, 1, &uuid)?;
        let mut incomplete_wrapper = Vec::new();
        append_length_delimited_field(&mut incomplete_wrapper, 1, &incomplete_category_ref)?;
        let mut incomplete_node = Vec::new();
        append_varint_field(&mut incomplete_node, 1, 66)?;
        append_length_delimited_field(&mut incomplete_node, 39, &incomplete_wrapper)?;
        let mut incomplete_array = Vec::new();
        append_length_delimited_field(&mut incomplete_array, 1, &incomplete_node)?;
        let mut incomplete_source = Vec::new();
        append_length_delimited_field(&mut incomplete_source, 1, &incomplete_array)?;
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert!(FormulaArchiveBytes::from_wire(&incomplete_source, &mut budget).is_err());

        // The same malformed category UUID must be rejected through more
        // than one thunk/ASTNodeArray edge, not only at the root node path.
        let mut bad_uuid = Vec::new();
        append_length_delimited_field(&mut bad_uuid, 1, &[])?;
        let mut bad_uuid_list = Vec::new();
        append_length_delimited_field(&mut bad_uuid_list, 1, &bad_uuid)?;
        let mut bad_category_ref = Vec::new();
        append_length_delimited_field(&mut bad_category_ref, 1, &bad_uuid_list)?;
        append_length_delimited_field(&mut bad_category_ref, 2, &uuid)?;
        append_varint_field(&mut bad_category_ref, 3, 1)?;
        append_varint_field(&mut bad_category_ref, 4, 0)?;
        let mut bad_category_wrapper = Vec::new();
        append_length_delimited_field(&mut bad_category_wrapper, 1, &bad_category_ref)?;
        let mut leaf = Vec::new();
        append_varint_field(&mut leaf, 1, 66)?;
        append_length_delimited_field(&mut leaf, 39, &bad_category_wrapper)?;
        let mut deep_array = Vec::new();
        append_length_delimited_field(&mut deep_array, 1, &leaf)?;
        for _ in 0..3 {
            let mut thunk = Vec::new();
            append_varint_field(&mut thunk, 1, 26)?;
            append_length_delimited_field(&mut thunk, 14, &deep_array)?;
            deep_array.clear();
            append_length_delimited_field(&mut deep_array, 1, &thunk)?;
        }
        let mut deep_source = Vec::new();
        append_length_delimited_field(&mut deep_source, 1, &deep_array)?;
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert!(FormulaArchiveBytes::from_wire(&deep_source, &mut budget).is_err());
        Ok(())
    }

    #[test]
    fn raw_formula_archive_requires_preserve_flags_required_fields() -> super::Result<()> {
        let uuid = tsp::Uuid { lower: 1, upper: 2 }.encode_to_vec();
        let mut category_ref = Vec::new();
        append_length_delimited_field(&mut category_ref, 1, &uuid)?;
        append_length_delimited_field(&mut category_ref, 2, &uuid)?;
        append_varint_field(&mut category_ref, 3, 1)?;
        append_varint_field(&mut category_ref, 4, 0)?;

        let mut missing_flags = Vec::new();
        append_varint_field(&mut missing_flags, 1, 1)?;
        let mut malformed_category_ref = category_ref.clone();
        append_length_delimited_field(&mut malformed_category_ref, 7, &missing_flags)?;
        let mut malformed_category_wrapper = Vec::new();
        append_length_delimited_field(&mut malformed_category_wrapper, 1, &malformed_category_ref)?;
        let mut malformed_node = Vec::new();
        append_varint_field(&mut malformed_node, 1, 66)?;
        append_length_delimited_field(&mut malformed_node, 39, &malformed_category_wrapper)?;
        let mut malformed_array = Vec::new();
        append_length_delimited_field(&mut malformed_array, 1, &malformed_node)?;
        let mut malformed_source = Vec::new();
        append_length_delimited_field(&mut malformed_source, 1, &malformed_array)?;
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert!(FormulaArchiveBytes::from_wire(&malformed_source, &mut budget).is_err());

        let mut complete_flags = missing_flags;
        append_varint_field(&mut complete_flags, 2, 0)?;
        let mut valid_category_ref = category_ref;
        append_length_delimited_field(&mut valid_category_ref, 7, &complete_flags)?;
        let mut valid_category_wrapper = Vec::new();
        append_length_delimited_field(&mut valid_category_wrapper, 1, &valid_category_ref)?;
        let mut valid_node = Vec::new();
        append_varint_field(&mut valid_node, 1, 66)?;
        append_length_delimited_field(&mut valid_node, 39, &valid_category_wrapper)?;
        let mut valid_array = Vec::new();
        append_length_delimited_field(&mut valid_array, 1, &valid_node)?;
        let mut valid_source = Vec::new();
        append_length_delimited_field(&mut valid_source, 1, &valid_array)?;
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert!(FormulaArchiveBytes::from_wire(&valid_source, &mut budget).is_ok());
        Ok(())
    }

    #[test]
    fn raw_formula_archive_enforces_canonical_boolean_and_integer_scalars() -> super::Result<()> {
        let mut sticky = Vec::new();
        append_varint_field(&mut sticky, 1, 0)?;
        append_varint_field(&mut sticky, 2, 0)?;
        append_varint_field(&mut sticky, 3, 0)?;
        append_varint_field(&mut sticky, 4, 0)?;
        let mut node = Vec::new();
        append_varint_field(&mut node, 1, 17)?;
        append_length_delimited_field(&mut node, 33, &sticky)?;
        let mut array = Vec::new();
        append_length_delimited_field(&mut array, 1, &node)?;
        let mut source = Vec::new();
        append_length_delimited_field(&mut source, 1, &array)?;
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert!(FormulaArchiveBytes::from_wire(&source, &mut budget).is_ok());

        let mut bad_bool = Vec::new();
        append_varint_field(&mut bad_bool, 1, 2)?;
        append_varint_field(&mut bad_bool, 2, 0)?;
        append_varint_field(&mut bad_bool, 3, 0)?;
        append_varint_field(&mut bad_bool, 4, 0)?;
        let mut bad_node = Vec::new();
        append_varint_field(&mut bad_node, 1, 17)?;
        append_length_delimited_field(&mut bad_node, 33, &bad_bool)?;
        let mut bad_array = Vec::new();
        append_length_delimited_field(&mut bad_array, 1, &bad_node)?;
        let mut bad_source = Vec::new();
        append_length_delimited_field(&mut bad_source, 1, &bad_array)?;
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert!(FormulaArchiveBytes::from_wire(&bad_source, &mut budget).is_err());

        // Column coordinates are sint32 and must not carry a value outside
        // the representable u32 zig-zag domain.
        let mut coordinate = Vec::new();
        append_varint_field(&mut coordinate, 1, u64::from(u32::MAX) + 1)?;
        let mut coordinate_node = Vec::new();
        append_varint_field(&mut coordinate_node, 1, 36)?;
        append_length_delimited_field(&mut coordinate_node, 26, &coordinate)?;
        let mut coordinate_array = Vec::new();
        append_length_delimited_field(&mut coordinate_array, 1, &coordinate_node)?;
        let mut coordinate_source = Vec::new();
        append_length_delimited_field(&mut coordinate_source, 1, &coordinate_array)?;
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert!(FormulaArchiveBytes::from_wire(&coordinate_source, &mut budget).is_err());
        Ok(())
    }

    #[test]
    fn raw_formula_archive_rejects_noncanonical_known_scalar_but_keeps_opaque_varints()
    -> super::Result<()> {
        let mut node = Vec::new();
        // ASTNode.type = 17, encoded with an overlong scalar value.
        node.extend_from_slice(&[0x08, 0x91, 0x00]);
        let mut node_array = Vec::new();
        append_length_delimited_field(&mut node_array, 1, &node)?;
        let mut known_bad = Vec::new();
        append_length_delimited_field(&mut known_bad, 1, &node_array)?;
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert!(FormulaArchiveBytes::from_wire(&known_bad, &mut budget).is_err());

        let mut opaque = known_bad;
        // Replace the node with a canonical type, then append an unknown
        // overlong varint. Unknown scalar payloads remain opaque by design.
        opaque.clear();
        let mut node = Vec::new();
        append_varint_field(&mut node, 1, 17)?;
        let mut node_array = Vec::new();
        append_length_delimited_field(&mut node_array, 1, &node)?;
        append_length_delimited_field(&mut opaque, 1, &node_array)?;
        opaque.extend_from_slice(&[0xd0, 0x05, 0x80, 0x00]);
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        assert!(FormulaArchiveBytes::from_wire(&opaque, &mut budget).is_ok());
        Ok(())
    }

    #[test]
    fn raw_formula_wire_budget_failure_is_atomic() -> super::Result<()> {
        let source = formula(vec![number_node(1.0)]).encode_to_vec();
        assert!(source.len() > 1);
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        budget.formula_wire_bytes = MAX_FORMULA_WIRE_BYTES - source.len() + 1;
        let fields_before = budget.payload_fields;
        let work_before = budget.payload_work;

        let error = FormulaArchiveBytes::from_wire(&source, &mut budget)
            .expect_err("one byte over the formula wire budget must fail");
        assert!(matches!(
            error,
            Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaWireBytes,
                observed,
                maximum,
                ..
            } if observed == MAX_FORMULA_WIRE_BYTES + 1 && maximum == MAX_FORMULA_WIRE_BYTES
        ));
        assert_eq!(
            budget.formula_wire_bytes,
            MAX_FORMULA_WIRE_BYTES - source.len() + 1
        );
        assert_eq!(budget.payload_fields, fields_before);
        assert_eq!(budget.payload_work, work_before);
        Ok(())
    }

    #[test]
    fn malformed_unreferenced_formula_keeps_attempted_preflight_cost() -> super::Result<()> {
        let source = [0x0a, 0x01, 0xff];
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let error = FormulaArchiveBytes::from_wire(&source, &mut budget)
            .expect_err("malformed unreferenced formula must be rejected");

        assert!(matches!(
            error,
            Error::MalformedPayload {
                path: SemanticPath::StructuredTables,
            }
        ));
        assert_eq!(budget.formula_wire_bytes, source.len());
        assert!(budget.payload_fields > 0);
        assert!(budget.payload_work >= source.len());
        Ok(())
    }

    #[test]
    fn formula_reference_wire_cost_is_merged_into_projection_budget() -> super::Result<()> {
        let formula_list = RawMessage {
            type_: 6_005,
            data: tst::TableDataList {
                list_type: tst::table_data_list::ListType::Formula as i32,
                next_list_id: 1,
                entries: vec![tst::table_data_list::ListEntry {
                    key: 1,
                    refcount: 1,
                    formula: Some(tsce::FormulaArchive::default()),
                    ..Default::default()
                }],
                ..Default::default()
            }
            .encode_to_vec(),
        };
        let string_list = empty_list(tst::table_data_list::ListType::String);
        let bytes = compatibility_package(vec![
            archive_object(
                1,
                vec![RawMessage {
                    type_: 99_999,
                    data: Vec::new(),
                }],
            )?,
            archive_object(90, vec![string_list, formula_list])?,
            // The payload is deliberately malformed for the category
            // projection. Its bytes still belong to the strict wire budget
            // before the best-effort category decode is discarded.
            archive_object(
                100,
                vec![RawMessage {
                    type_: 6_383,
                    data: vec![0x08, 0x01],
                }],
            )?,
        ])?;
        let components = Components::from_bytes(&bytes, Limits::default())?;
        let index = Index::from_components(&components, SemanticLimits::MAX_OBJECTS)?;
        let extractor = TableDataExtractor::new(&components, &index, SemanticLimits::default());
        let model = tst::TableModelArchive {
            table_name: "wire".to_owned(),
            base_data_store: tst::DataStore {
                string_table: reference(90),
                formula_table: reference(90),
                ..Default::default()
            },
            ..Default::default()
        };

        let mut candidate = ProjectionBudget::new(SemanticLimits::default());
        candidate.formula_wire_bytes = MAX_FORMULA_WIRE_BYTES - 4;
        extractor.parse_table_model(model.clone(), false, Some(candidate))?;
        assert_eq!(
            extractor.projection_budget.borrow().formula_wire_bytes,
            MAX_FORMULA_WIRE_BYTES
        );

        let mut second_candidate = ProjectionBudget::new(SemanticLimits::default());
        second_candidate.formula_wire_bytes = MAX_FORMULA_WIRE_BYTES - 2;
        let error = extractor
            .parse_table_model(model, false, Some(second_candidate))
            .expect_err("the second formula-map build must see the aggregate wire ceiling");
        assert!(matches!(
            error,
            Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaWireBytes,
                observed: 2,
                maximum: 0,
                path: SemanticPath::StructuredTables,
            }
        ));
        Ok(())
    }

    #[test]
    fn list_root_and_segment_order_duplicates_and_ranges_are_strict() -> super::Result<()> {
        let ordered = with_list_extractor(
            vec![
                archive_object(
                    90,
                    vec![list_message(
                        tst::table_data_list::ListType::String,
                        vec![string_list_entry(2, "root")],
                        vec![91],
                    )],
                )?,
                archive_object(
                    91,
                    vec![segment_message(
                        tst::table_data_list::ListType::String,
                        1,
                        1,
                        vec![string_list_entry(1, "segment")],
                    )],
                )?,
            ],
            true,
            |extractor| load_string_list(extractor, 90).map(|(table, _)| table),
        )?;
        assert_eq!(
            ordered.as_ref(),
            &[(1, "segment".to_owned()), (2, "root".to_owned())]
        );

        let duplicate_segment_key = with_list_extractor(
            vec![
                archive_object(
                    90,
                    vec![list_message(
                        tst::table_data_list::ListType::String,
                        Vec::new(),
                        vec![91],
                    )],
                )?,
                archive_object(
                    91,
                    vec![segment_message(
                        tst::table_data_list::ListType::String,
                        1,
                        2,
                        vec![string_list_entry(1, "a"), string_list_entry(1, "b")],
                    )],
                )?,
            ],
            true,
            |extractor| load_string_list(extractor, 90).map(|_| ()),
        );
        assert!(matches!(
            duplicate_segment_key,
            Err(Error::InvalidFormat(_))
        ));

        let outside_range = with_list_extractor(
            vec![
                archive_object(
                    90,
                    vec![list_message(
                        tst::table_data_list::ListType::String,
                        Vec::new(),
                        vec![91],
                    )],
                )?,
                archive_object(
                    91,
                    vec![segment_message(
                        tst::table_data_list::ListType::String,
                        10,
                        1,
                        vec![string_list_entry(1, "outside")],
                    )],
                )?,
            ],
            false,
            |extractor| load_string_list(extractor, 90).map(|_| ()),
        );
        assert!(matches!(outside_range, Err(Error::InvalidFormat(_))));

        let duplicate_segment_id = with_list_extractor(
            vec![
                archive_object(
                    90,
                    vec![list_message(
                        tst::table_data_list::ListType::String,
                        Vec::new(),
                        vec![91, 91],
                    )],
                )?,
                archive_object(
                    91,
                    vec![segment_message(
                        tst::table_data_list::ListType::String,
                        1,
                        1,
                        vec![string_list_entry(1, "segment")],
                    )],
                )?,
            ],
            true,
            |extractor| load_string_list(extractor, 90).map(|_| ()),
        );
        assert!(matches!(duplicate_segment_id, Err(Error::InvalidFormat(_))));

        let duplicate_root = with_list_extractor(
            vec![archive_object(
                90,
                vec![list_message(
                    tst::table_data_list::ListType::String,
                    vec![string_list_entry(1, "a"), string_list_entry(1, "b")],
                    Vec::new(),
                )],
            )?],
            false,
            |extractor| load_string_list(extractor, 90).map(|_| ()),
        );
        assert!(matches!(duplicate_root, Err(Error::InvalidFormat(_))));

        for document_projection in [false, true] {
            let mut malformed = list_message(
                tst::table_data_list::ListType::String,
                vec![string_list_entry(1, "callback-before-wire-error")],
                Vec::new(),
            );
            malformed.data.push(0);
            let result = with_list_extractor(
                vec![archive_object(90, vec![malformed])?],
                document_projection,
                |extractor| load_string_list(extractor, 90).map(|_| ()),
            );
            assert!(matches!(result, Err(Error::InvalidFormat(_))));
        }
        Ok(())
    }

    #[test]
    fn wrong_list_candidate_rolls_back_retention_but_keeps_decode_work() -> super::Result<()> {
        let bytes = compatibility_package(vec![
            archive_object(
                90,
                vec![empty_list(tst::table_data_list::ListType::Formula)],
            )?,
            archive_object(
                91,
                vec![empty_list(tst::table_data_list::ListType::Formula)],
            )?,
            archive_object(
                92,
                vec![
                    empty_list(tst::table_data_list::ListType::String),
                    empty_list(tst::table_data_list::ListType::Formula),
                ],
            )?,
        ])?;
        let components = Components::from_bytes(&bytes, Limits::default())?;
        let index = Index::from_components(&components, SemanticLimits::MAX_OBJECTS)?;
        let extractor = TableDataExtractor::new(&components, &index, SemanticLimits::default())
            .without_comments();

        let invalid = tst::TableModelArchive {
            table_name: "rejected".to_owned(),
            number_of_rows: 1,
            number_of_columns: 1,
            base_data_store: tst::DataStore {
                string_table: reference(90),
                formula_table: reference(91),
                ..Default::default()
            },
            ..Default::default()
        };
        assert!(matches!(
            extractor.parse_table_model(invalid, false, None),
            Err(Error::InvalidFormat(_))
        ));
        let rejected = *extractor.projection_budget.borrow();
        assert_eq!(rejected.references, 0);
        assert_eq!(rejected.materialized_cells, 0);
        assert_eq!(rejected.output_text_bytes, 0);
        assert!(rejected.payload_work > 0);

        let valid = tst::TableModelArchive {
            table_name: "accepted".to_owned(),
            number_of_rows: 1,
            number_of_columns: 1,
            base_data_store: tst::DataStore {
                string_table: reference(92),
                formula_table: reference(92),
                ..Default::default()
            },
            ..Default::default()
        };
        let table = extractor.parse_table_model(valid, false, None)?;
        assert_eq!(table.name(), "accepted");
        let published = *extractor.projection_budget.borrow();
        assert!(published.payload_work > rejected.payload_work);
        assert_eq!(published.output_text_bytes, "accepted".len());
        Ok(())
    }

    #[test]
    fn rejected_attempt_rolls_back_retained_values_but_not_formula_work() -> super::Result<()> {
        let limits = SemanticLimits::default()
            .with_projection_limits(3, 5)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?
            .with_formula_render_limits(2, 1)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut published = ProjectionBudget::new(limits);
        published.charge_materialized_cells(1)?;
        published.charge_output_text(1)?;

        let mut rejected = published;
        rejected.charge_materialized_cells(2)?;
        rejected.charge_output_text(4)?;
        rejected.charge_formula_render_work(1)?;
        published.commit_attempt(rejected, false);

        assert_eq!(published.materialized_cells, 1);
        assert_eq!(published.output_text_bytes, 1);
        assert_eq!(published.formula_render_work, 1);

        let mut over_budget = published;
        assert!(matches!(
            over_budget.charge_formula_render_work(2),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaRenderWork,
                observed: 3,
                maximum: 2,
                ..
            })
        ));
        published.commit_attempt(over_budget, false);
        assert_eq!(published.formula_render_work, 2);
        Ok(())
    }

    #[test]
    fn formula_category_walk_is_lazy_iterative_and_bounded() -> super::Result<()> {
        let mut deep_wire = Vec::new();
        for _ in 0..=MAX_FORMULA_CATEGORY_DEPTH {
            let mut parent = Vec::new();
            append_length_delimited_field(&mut parent, 3, &deep_wire)?;
            deep_wire = parent;
        }
        let mut depth_names = HashMap::new();
        let mut depth_budget = FormulaReferenceBudget::new(
            crate::MAX_REFERENCES,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
        let depth_result =
            collect_formula_category_payload(&deep_wire, &mut depth_names, &mut depth_budget);
        assert!(
            matches!(
                &depth_result,
                Err(Error::SemanticLimit {
                    kind: SemanticLimitKind::FormulaDepth,
                    observed,
                    maximum: MAX_FORMULA_CATEGORY_DEPTH,
                    path: SemanticPath::StructuredTables,
                }) if *observed == MAX_FORMULA_CATEGORY_DEPTH + 1
            ),
            "unexpected depth result: {depth_result:?}"
        );

        let category = |lower, value: &str, child| tst::group_by_archive::GroupNodeArchive {
            group_uid: tsp::Uuid { lower, upper: 0 },
            group_cell_value: Some(tsce::CellValueArchive {
                string_value: Some(tsce::StringCellValueArchive {
                    value: value.to_owned(),
                    ..Default::default()
                }),
                ..Default::default()
            }),
            child,
            ..Default::default()
        };
        let shallow = category(1, "root", vec![category(2, "child", Vec::new())]);
        let mut tight_names = HashMap::new();
        let mut tight_budget = FormulaReferenceBudget::new(
            1,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
        assert!(matches!(
            collect_formula_category_payload(
                &shallow.encode_to_vec(),
                &mut tight_names,
                &mut tight_budget,
            ),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::References,
                observed: 2,
                maximum: 1,
                path: SemanticPath::StructuredTables,
            })
        ));

        let empty_fanout = tst::group_by_archive::GroupNodeArchive {
            child: vec![tst::group_by_archive::GroupNodeArchive::default(); 32],
            ..Default::default()
        };
        let mut fanout_names = HashMap::new();
        let mut false_positive_budget = FormulaReferenceBudget::new(
            1,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
        collect_formula_category_payload(
            &empty_fanout.encode_to_vec(),
            &mut fanout_names,
            &mut false_positive_budget,
        )?;
        assert_eq!(false_positive_budget.retained_entries, 0);

        let one_child = tst::group_by_archive::GroupNodeArchive {
            child: vec![tst::group_by_archive::GroupNodeArchive::default()],
            ..Default::default()
        };
        let mut work_names = HashMap::new();
        let mut full_work_budget = FormulaReferenceBudget::new(
            crate::MAX_REFERENCES,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
        full_work_budget.work_items = MAX_FORMULA_WORK - 1;
        assert!(matches!(
            collect_formula_category_payload(
                &one_child.encode_to_vec(),
                &mut work_names,
                &mut full_work_budget,
            ),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaWork,
                observed,
                maximum: MAX_FORMULA_WORK,
                path: SemanticPath::StructuredTables,
            }) if observed > MAX_FORMULA_WORK
        ));

        let boolean_wrapper = [0x08, 0x01];
        let mut cell_value = Vec::new();
        append_length_delimited_field(&mut cell_value, 2, &boolean_wrapper)?;
        let mut nested_projection = Vec::new();
        append_length_delimited_field(&mut nested_projection, 7, &cell_value)?;
        let mut nested_names = HashMap::new();
        let mut nested_work_budget = FormulaReferenceBudget::new(
            crate::MAX_REFERENCES,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
        nested_work_budget.work_items = MAX_FORMULA_WORK - 5;
        assert!(matches!(
            collect_formula_category_payload(
                &nested_projection,
                &mut nested_names,
                &mut nested_work_budget,
            ),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaWork,
                observed,
                maximum: MAX_FORMULA_WORK,
                path: SemanticPath::StructuredTables,
            }) if observed > MAX_FORMULA_WORK
        ));

        let mut string_wrapper = Vec::new();
        append_length_delimited_field(&mut string_wrapper, 1, b"valid")?;
        let mut malformed_cell = Vec::new();
        append_length_delimited_field(&mut malformed_cell, 5, &string_wrapper)?;
        malformed_cell.extend_from_slice(&[0x20, 0x01]);
        let mut malformed_projection = Vec::new();
        append_length_delimited_field(&mut malformed_projection, 7, &malformed_cell)?;
        let mut malformed_names = HashMap::new();
        let mut malformed_budget = FormulaReferenceBudget::new(
            crate::MAX_REFERENCES,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
        collect_formula_category_payload(
            &malformed_projection,
            &mut malformed_names,
            &mut malformed_budget,
        )?;
        assert!(malformed_names.is_empty());
        let malformed_work = malformed_budget.work_items;
        assert!(malformed_work > 0);
        malformed_budget.work_items = MAX_FORMULA_WORK - malformed_work;
        collect_formula_category_payload(
            &malformed_projection,
            &mut malformed_names,
            &mut malformed_budget,
        )?;
        assert_eq!(malformed_budget.work_items, MAX_FORMULA_WORK);
        assert!(matches!(
            collect_formula_category_payload(&[], &mut malformed_names, &mut malformed_budget),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaWork,
                observed,
                maximum: MAX_FORMULA_WORK,
                path: SemanticPath::StructuredTables,
            }) if observed == MAX_FORMULA_WORK + 1
        ));

        let duplicate = category(1, "first", vec![category(1, "second", Vec::new())]);
        let mut duplicate_names = HashMap::new();
        let mut duplicate_budget = FormulaReferenceBudget::new(
            1,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
        collect_formula_category_payload(
            &duplicate.encode_to_vec(),
            &mut duplicate_names,
            &mut duplicate_budget,
        )?;
        assert_eq!(
            duplicate_names.get(&[1, 0]).map(String::as_str),
            Some("second")
        );
        assert_eq!(duplicate_budget.retained_entries, 1);

        duplicate_names.insert([1, 0], "Grand Total".to_owned());
        let localized = category(1, "Localized Total", Vec::new());
        collect_formula_category_payload(
            &localized.encode_to_vec(),
            &mut duplicate_names,
            &mut duplicate_budget,
        )?;
        assert_eq!(
            duplicate_names.get(&[1, 0]).map(String::as_str),
            Some("Localized Total")
        );
        assert_eq!(duplicate_budget.retained_entries, 1);
        Ok(())
    }

    #[test]
    fn formula_category_preflight_work_boundary_is_monotonic_without_retention() {
        // One empty category payload spends exactly one unit of the local
        // preflight work budget. The second attempt is one unit over the
        // aggregate ceiling and must retain the already-spent boundary work.
        let mut names = HashMap::new();
        let mut exact_budget = FormulaReferenceBudget::new(
            crate::MAX_REFERENCES,
            1,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
        collect_formula_category_payload(&[], &mut names, &mut exact_budget)
            .expect("the exact formula-work boundary must be admitted");
        assert_eq!(exact_budget.work_items, 1);
        assert_eq!(exact_budget.retained_entries, 0);
        assert_eq!(exact_budget.text_bytes, 0);
        assert!(names.is_empty());

        let error = collect_formula_category_payload(&[], &mut names, &mut exact_budget)
            .expect_err("one unit over the formula-work boundary must fail");
        assert!(matches!(
            error,
            Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaWork,
                observed: 2,
                maximum: MAX_FORMULA_WORK,
                path: SemanticPath::StructuredTables,
            }
        ));
        assert_eq!(exact_budget.work_items, 1);
        assert_eq!(exact_budget.retained_entries, 0);
        assert_eq!(exact_budget.text_bytes, 0);
        assert!(names.is_empty());

        // A field that crosses the local preflight allowance fails after the
        // source wire charge. Its attempted projection work must still
        // consume the candidate budget, but no map entry or label is visible.
        let mut one_group = Vec::new();
        append_length_delimited_field(&mut one_group, 3, &[])
            .expect("test category wire must encode");
        let mut names = HashMap::new();
        let mut rejected_budget = FormulaReferenceBudget::new(
            crate::MAX_REFERENCES,
            4,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
        let error = collect_formula_category_payload(&one_group, &mut names, &mut rejected_budget)
            .expect_err("the category preflight must cross its exact work boundary");
        assert!(matches!(
            error,
            Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaWork,
                observed: 5,
                maximum: MAX_FORMULA_WORK,
                path: SemanticPath::StructuredTables,
            }
        ));
        assert_eq!(rejected_budget.wire_bytes, one_group.len());
        assert_eq!(rejected_budget.work_items, 4);
        assert_eq!(rejected_budget.retained_entries, 0);
        assert_eq!(rejected_budget.text_bytes, 0);
        assert!(names.is_empty());
    }

    #[test]
    fn formula_category_depth_boundary_retains_prefix_work_without_map_text() {
        let nested_wire = |depth: usize| -> Vec<u8> {
            let mut source = Vec::new();
            for _ in 0..depth {
                let mut parent = Vec::new();
                append_length_delimited_field(&mut parent, 3, &source)
                    .expect("test category wire must encode");
                source = parent;
            }
            source
        };

        let exact_source = nested_wire(MAX_FORMULA_CATEGORY_DEPTH);
        let mut exact_budget = FormulaReferenceBudget::new(
            crate::MAX_REFERENCES,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
        let expected_nodes = preflight_formula_category_payload(&exact_source, &mut exact_budget)
            .expect("the exact category-depth boundary must be admitted");
        assert_eq!(expected_nodes, Some(MAX_FORMULA_CATEGORY_DEPTH + 1));
        assert_eq!(exact_budget.wire_bytes, exact_source.len());
        assert_eq!(
            exact_budget.work_items,
            exact_source.len() + (MAX_FORMULA_CATEGORY_DEPTH * 2 + 1)
        );

        let over_source = nested_wire(MAX_FORMULA_CATEGORY_DEPTH + 1);
        let mut over_names = HashMap::new();
        let mut over_budget = FormulaReferenceBudget::new(
            crate::MAX_REFERENCES,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
        let error =
            collect_formula_category_payload(&over_source, &mut over_names, &mut over_budget)
                .expect_err("one level over the category-depth boundary must fail");
        assert!(matches!(
            error,
            Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaDepth,
                observed,
                maximum: MAX_FORMULA_CATEGORY_DEPTH,
                path: SemanticPath::StructuredTables,
            } if observed == MAX_FORMULA_CATEGORY_DEPTH + 1
        ));
        assert_eq!(over_budget.wire_bytes, over_source.len());
        assert_eq!(
            over_budget.work_items,
            over_source.len() + (MAX_FORMULA_CATEGORY_DEPTH * 2 + 2)
        );
        assert_eq!(over_budget.retained_entries, 0);
        assert_eq!(over_budget.text_bytes, 0);
        assert!(over_names.is_empty());
    }

    #[test]
    fn formula_category_wire_bytes_are_aggregate_and_inclusive() {
        let mut names = HashMap::new();
        let mut budget = FormulaReferenceBudget::new(
            crate::MAX_REFERENCES,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
        budget.wire_bytes = MAX_FORMULA_WIRE_BYTES;

        assert!(matches!(
            collect_formula_category_payload(&[0x08], &mut names, &mut budget),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaWireBytes,
                observed,
                maximum: MAX_FORMULA_WIRE_BYTES,
                path: SemanticPath::StructuredTables,
            }) if observed == MAX_FORMULA_WIRE_BYTES + 1
        ));
    }

    include!("extractor_rich_text_tests.rs");
}
