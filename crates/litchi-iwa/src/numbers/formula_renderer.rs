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
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_protos::numbers_formula_codec;
use litchi_numbers_wire::formula_envelope::{
    self as shared_formula_envelope, FormulaEnvelopeLimits,
};
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
        let report = match shared_formula_envelope::preflight_formula_envelope(
            source,
            FormulaEnvelopeLimits {
                max_fields: litchi_numbers::MAX_REFERENCES,
                max_input_bytes: MAX_PAYLOAD_WORK
                    .saturating_sub(budget.payload_work)
                    .clamp(1, WireLimits::MAX_INPUT_BYTES),
                max_work: MAX_PAYLOAD_WORK,
                base_fields: budget.payload_fields,
                base_work: budget.payload_work,
            },
        ) {
            Ok(report) if source.is_empty() || report.root_ast_present() => {
                if let Some(wire) = report.wire_preflight() {
                    budget.charge_wire_preflight(wire)?;
                }
                report
            },
            Ok(report) => {
                budget.retain_formula_preflight_cost(report.fields(), report.scanned_bytes());
                return Err(Error::InvalidFormat(
                    "malformed Numbers formula payload".to_owned(),
                ));
            },
            Err(failure) => {
                let (error, attempted) = failure.into_parts();
                budget.retain_formula_preflight_cost(attempted.fields(), attempted.work());
                return Err(map_formula_envelope_wire_error(error));
            },
        };

        let mut owned = Vec::new();
        owned
            .try_reserve_exact(source.len())
            .map_err(|_| allocation_error("Numbers formula wire", source.len()))?;
        owned.extend_from_slice(source);
        Ok(Self {
            bytes: owned.into_boxed_slice(),
            scalar_visitor_eligible: report.scalar_visitor_eligible(),
            scalar_visitor_node_count: report.scalar_visitor_node_count(),
            lazy_traversal_entry_count: report.lazy_traversal_entry_count(),
        })
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
