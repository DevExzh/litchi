//! Bounded, source-authoritative rewrites for formula-owner removal.
//!
//! The package editors own graph discovery and archive mutation.  This
//! module only understands the three dependency payload transitions that are
//! needed after a formula-owner family has been selected: removing owner
//! references from a calculation engine, pruning inline cell edges, and
//! pruning the same edges from a cell-record tile.  Every rewrite starts with
//! the borrowed strict projection and verifies the resulting bytes with that
//! projection before publishing them to a package.

use super::*;

/// Accounting for one source-preserving dependency rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaDependencyRewriteReport {
    source: DecodeReport,
    result: DecodeReport,
    output_bytes: usize,
    removed_values: usize,
    rewrite_work_bytes: usize,
}

impl FormulaDependencyRewriteReport {
    /// Aggregate source validation consumption.
    #[must_use]
    pub const fn source(self) -> DecodeReport {
        self.source
    }

    /// Aggregate result validation consumption.
    #[must_use]
    pub const fn result(self) -> DecodeReport {
        self.result
    }

    /// Encoded output size after the rewrite.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Number of owner references, map entries, or expanded-edge tuples
    /// removed by the rewrite.
    #[must_use]
    pub const fn removed_values(self) -> usize {
        self.removed_values
    }

    /// Finite source-copy/rewrite work charged by the prepared operation.
    #[must_use]
    pub const fn rewrite_work_bytes(self) -> usize {
        self.rewrite_work_bytes
    }
}

/// Resource requirements known before a rewrite output is allocated.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaDependencyRewriteRequirements {
    source: DecodeReport,
    result_fields: usize,
    result_max_depth: u32,
    result_references: usize,
    result_reference_bytes: usize,
    result_text_bytes: usize,
    output_bytes: usize,
    rewrite_work_bytes: usize,
}

impl FormulaDependencyRewriteRequirements {
    /// Aggregate source validation consumption.
    #[must_use]
    pub const fn source(self) -> DecodeReport {
        self.source
    }

    /// Conservative result field count checked before candidate allocation.
    #[must_use]
    pub const fn result_fields(self) -> usize {
        self.result_fields
    }

    /// Conservative result nesting depth checked before candidate allocation.
    #[must_use]
    pub const fn result_max_depth(self) -> u32 {
        self.result_max_depth
    }

    /// Conservative result reference count checked before candidate allocation.
    #[must_use]
    pub const fn result_references(self) -> usize {
        self.result_references
    }

    /// Conservative result reference bytes checked before candidate allocation.
    #[must_use]
    pub const fn result_reference_bytes(self) -> usize {
        self.result_reference_bytes
    }

    /// Conservative result text bytes checked before candidate allocation.
    #[must_use]
    pub const fn result_text_bytes(self) -> usize {
        self.result_text_bytes
    }

    /// Exact output bytes for the engine rewrite, or the source-size upper
    /// bound for edge pruning (edge pruning never grows a payload).
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Finite rewrite work upper bound.
    #[must_use]
    pub const fn rewrite_work_bytes(self) -> usize {
        self.rewrite_work_bytes
    }
}

/// Prepared calculation-engine owner removal borrowing the source payload.
pub struct CalculationEngineOwnerRemovalPlan<'source> {
    source: &'source [u8],
    owner_ids: Vec<u64>,
    internal_owner_ids: Vec<u32>,
    formula_count: u64,
    preparation_work_bytes: usize,
    requirements: FormulaDependencyRewriteRequirements,
}

impl fmt::Debug for CalculationEngineOwnerRemovalPlan<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("CalculationEngineOwnerRemovalPlan")
            .field("requirements", &self.requirements)
            .field("owner_ids", &"<redacted>")
            .field("internal_owner_ids", &"<redacted>")
            .field("formula_count", &self.formula_count)
            .finish()
    }
}

impl CalculationEngineOwnerRemovalPlan<'_> {
    /// Return the finite requirements computed by preparation.
    #[must_use]
    pub const fn requirements(&self) -> FormulaDependencyRewriteRequirements {
        self.requirements
    }
}

/// Prepared inline FormulaOwner cell-edge pruning borrowing the source.
pub struct FormulaOwnerCellEdgesRemovalPlan<'source> {
    source: &'source [u8],
    removed_internal_owner_ids: Vec<u32>,
    preparation_work_bytes: usize,
    requirements: FormulaDependencyRewriteRequirements,
}

impl fmt::Debug for FormulaOwnerCellEdgesRemovalPlan<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("FormulaOwnerCellEdgesRemovalPlan")
            .field("requirements", &self.requirements)
            .field("removed_internal_owner_ids", &"<redacted>")
            .finish()
    }
}

impl FormulaOwnerCellEdgesRemovalPlan<'_> {
    /// Return the finite requirements computed by preparation.
    #[must_use]
    pub const fn requirements(&self) -> FormulaDependencyRewriteRequirements {
        self.requirements
    }
}

/// Prepared CellRecordTile edge pruning borrowing the source.
pub struct CellRecordTileEdgesRemovalPlan<'source> {
    source: &'source [u8],
    removed_internal_owner_ids: Vec<u32>,
    preparation_work_bytes: usize,
    requirements: FormulaDependencyRewriteRequirements,
}

impl fmt::Debug for CellRecordTileEdgesRemovalPlan<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("CellRecordTileEdgesRemovalPlan")
            .field("requirements", &self.requirements)
            .field("removed_internal_owner_ids", &"<redacted>")
            .finish()
    }
}

impl CellRecordTileEdgesRemovalPlan<'_> {
    /// Return the finite requirements computed by preparation.
    #[must_use]
    pub const fn requirements(&self) -> FormulaDependencyRewriteRequirements {
        self.requirements
    }
}

/// One aggregate ceiling for every validation, search, scratch-copy, and
/// output-copy pass made by a prepared rewrite.  The regular codec budget is
/// intentionally scoped to one message; this ledger closes the gap created by
/// a rewrite's several bounded nested passes.
#[derive(Debug, Clone, Copy)]
struct WorkBudget {
    maximum: usize,
    observed: usize,
}

#[derive(Default)]
struct InternalOwnerEdgeCount {
    count: usize,
}

// The scan multiplier is a fixed envelope for execution-only passes.  It is
// deliberately independent of the number of dependency records: each nested
// message is a disjoint slice of the root payload, and every strict decoder
// charges at most two message visits for a slice.  The engine execution fits
// within 27 root-sized units (root/tracker/map/reference/patch scans and
// copies, owner-entry and reference validation, plus source/result strict
// validation); the two edge paths fit within 18 units (three path scans and
// copies, expanded-edge staging, and source/result validation).  Four units of
// headroom cover envelope re-encoding and fixed bookkeeping. Membership
// searches are charged separately from this source-sized envelope.
const EXECUTION_SCAN_PASSES: usize = 32;

impl DependencyVisitor for InternalOwnerEdgeCount {
    fn visit_expanded_edge_component(
        &mut self,
        component: ExpandedEdgeComponent,
    ) -> Result<(), DecodeError> {
        if component.kind() == ExpandedEdgeKind::InternalOwner {
            self.count = self.count.checked_add(1).ok_or_else(DecodeError::invalid)?;
        }
        Ok(())
    }
}

impl WorkBudget {
    const fn new(maximum: usize) -> Self {
        Self {
            maximum,
            observed: 0,
        }
    }

    fn charge(&mut self, amount: usize) -> Result<(), DecodeError> {
        let observed = self
            .observed
            .checked_add(amount)
            .ok_or_else(DecodeError::invalid)?;
        if observed > self.maximum {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed,
                maximum: self.maximum,
            }));
        }
        self.observed = observed;
        Ok(())
    }

    fn charge_report(&mut self, report: DecodeReport) -> Result<(), DecodeError> {
        self.charge(report.work_bytes())
    }

    const fn observed(self) -> usize {
        self.observed
    }
}

fn begin_scan(
    source: &[u8],
    options: DecodeOptions,
    work: &mut WorkBudget,
) -> Result<wire::Budget, DecodeError> {
    let mut budget = wire::Budget::new(source, options)?;
    budget.message(source, 1)?;
    // Reserve the full raw scan before a rewrite candidate can allocate.
    work.charge(source.len())?;
    Ok(budget)
}

fn finish_scan(
    work: &mut WorkBudget,
    budget: wire::Budget,
    source_bytes: usize,
) -> Result<(), DecodeError> {
    let report = budget.report();
    let nested_work = report
        .work_bytes()
        .checked_sub(source_bytes)
        .ok_or_else(DecodeError::invalid)?;
    work.charge(nested_work)
}

/// Prepare a calculation-engine owner-family removal.
pub fn prepare_calculation_engine_owner_removal<'source>(
    source: &'source [u8],
    owner_ids: &[u64],
    internal_owner_ids: &[u32],
    formula_count: u64,
    options: DecodeOptions,
) -> Result<CalculationEngineOwnerRemovalPlan<'source>, DecodeError> {
    let mut work = WorkBudget::new(options.max_work_bytes);
    let (engine, source_report) = decode_calculation_engine_with_report(source, options)?;
    work.charge_report(source_report)?;
    let (tracker, tracker_report) =
        decode_dependency_tracker_with_report(engine.dependency_tracker(), options)?;
    work.charge_report(tracker_report)?;
    let previous_count = tracker.number_of_formulas().unwrap_or_default();
    previous_count
        .checked_sub(formula_count)
        .ok_or_else(DecodeError::invalid)?;

    validate_target_width(
        owner_ids.len(),
        internal_owner_ids.len(),
        options,
        &mut work,
    )?;
    let owner_ids = copy_sorted_ids(owner_ids, &mut work)?;
    let internal_owner_ids = copy_sorted_ids(internal_owner_ids, &mut work)?;
    let (output_bytes, membership_queries) = measure_calculation_engine_output(
        source,
        &owner_ids,
        &internal_owner_ids,
        formula_count,
        options,
        &mut work,
    )?;
    let execution_work_bytes = execution_work_upper_bound(
        source.len(),
        output_bytes,
        owner_ids.len().saturating_add(internal_owner_ids.len()),
        membership_queries,
        EXECUTION_SCAN_PASSES,
    )?;
    let preparation_work_bytes = work.observed();
    let total_work_bytes = preparation_work_bytes
        .checked_add(execution_work_bytes)
        .ok_or_else(DecodeError::invalid)?;
    enforce_work(total_work_bytes, options)?;
    let result_fields = source_report
        .fields()
        .checked_add(1)
        .ok_or_else(DecodeError::invalid)?;
    let requirements = requirements(
        source_report,
        result_fields,
        output_bytes,
        total_work_bytes,
        options,
    )?;
    Ok(CalculationEngineOwnerRemovalPlan {
        source,
        owner_ids,
        internal_owner_ids,
        formula_count,
        preparation_work_bytes,
        requirements,
    })
}

/// Execute a prepared calculation-engine owner-family removal.
pub fn execute_calculation_engine_owner_removal(
    plan: CalculationEngineOwnerRemovalPlan<'_>,
    options: DecodeOptions,
) -> Result<(Vec<u8>, FormulaDependencyRewriteReport), DecodeError> {
    enforce_requirements(plan.requirements, options)?;
    if plan.preparation_work_bytes > options.max_work_bytes {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: plan.preparation_work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    let mut work = WorkBudget::new(options.max_work_bytes - plan.preparation_work_bytes);
    let (data, removed_values) = assemble_calculation_engine_owner_removal(
        plan.source,
        &plan.owner_ids,
        &plan.internal_owner_ids,
        plan.formula_count,
        options,
        &mut work,
    )?;
    if data.len() != plan.requirements.output_bytes {
        return Err(DecodeError::invalid());
    }
    let (_, result) = decode_calculation_engine_with_report(&data, options)?;
    work.charge_report(result)?;
    if work.observed() > plan.requirements.rewrite_work_bytes - plan.preparation_work_bytes {
        return Err(DecodeError::invalid());
    }
    Ok((
        data,
        FormulaDependencyRewriteReport {
            source: plan.requirements.source,
            result,
            output_bytes: plan.requirements.output_bytes,
            removed_values,
            rewrite_work_bytes: plan.requirements.rewrite_work_bytes,
        },
    ))
}

/// Prepare and execute a calculation-engine owner-family removal in one
/// bounded operation.
pub fn rewrite_calculation_engine_owner_removal(
    source: &[u8],
    owner_ids: &[u64],
    internal_owner_ids: &[u32],
    formula_count: u64,
    options: DecodeOptions,
) -> Result<(Vec<u8>, FormulaDependencyRewriteReport), DecodeError> {
    execute_calculation_engine_owner_removal(
        prepare_calculation_engine_owner_removal(
            source,
            owner_ids,
            internal_owner_ids,
            formula_count,
            options,
        )?,
        options,
    )
}

/// Prepare inline FormulaOwner cell-edge pruning.
pub fn prepare_formula_owner_cell_edges_removal<'source>(
    source: &'source [u8],
    removed_internal_owner_ids: &[u32],
    options: DecodeOptions,
) -> Result<FormulaOwnerCellEdgesRemovalPlan<'source>, DecodeError> {
    let mut work = WorkBudget::new(options.max_work_bytes);
    let mut edge_count = InternalOwnerEdgeCount::default();
    let (_, source_report) =
        decode_formula_owner_dependencies_with_visitor(source, options, &mut edge_count)?;
    work.charge_report(source_report)?;
    validate_target_width(0, removed_internal_owner_ids.len(), options, &mut work)?;
    let removed_internal_owner_ids = copy_sorted_ids(removed_internal_owner_ids, &mut work)?;
    let preparation_work_bytes = work.observed();
    let execution_work_bytes = execution_work_upper_bound(
        source.len(),
        source.len(),
        removed_internal_owner_ids.len(),
        edge_count.count,
        EXECUTION_SCAN_PASSES,
    )?;
    let total_work_bytes = preparation_work_bytes
        .checked_add(execution_work_bytes)
        .ok_or_else(DecodeError::invalid)?;
    enforce_work(total_work_bytes, options)?;
    let requirements = requirements(
        source_report,
        source_report.fields(),
        source.len(),
        total_work_bytes,
        options,
    )?;
    Ok(FormulaOwnerCellEdgesRemovalPlan {
        source,
        removed_internal_owner_ids,
        preparation_work_bytes,
        requirements,
    })
}

/// Execute prepared inline FormulaOwner cell-edge pruning.
pub fn execute_formula_owner_cell_edges_removal(
    plan: FormulaOwnerCellEdgesRemovalPlan<'_>,
    options: DecodeOptions,
) -> Result<(Vec<u8>, FormulaDependencyRewriteReport), DecodeError> {
    enforce_requirements(plan.requirements, options)?;
    if plan.preparation_work_bytes > options.max_work_bytes {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: plan.preparation_work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    let mut work = WorkBudget::new(options.max_work_bytes - plan.preparation_work_bytes);
    let (data, removed_values) = rewrite_formula_owner_cell_edges_inner(
        plan.source,
        &plan.removed_internal_owner_ids,
        options,
        &mut work,
    )?;
    if data.len() > plan.requirements.output_bytes {
        return Err(DecodeError::invalid());
    }
    let output_bytes = data.len();
    let (_, result) = decode_formula_owner_dependencies_with_report(&data, options)?;
    work.charge_report(result)?;
    if work.observed()
        > plan
            .requirements
            .rewrite_work_bytes
            .checked_sub(plan.preparation_work_bytes)
            .ok_or_else(DecodeError::invalid)?
    {
        return Err(DecodeError::invalid());
    }
    Ok((
        data,
        FormulaDependencyRewriteReport {
            source: plan.requirements.source,
            result,
            output_bytes,
            removed_values,
            rewrite_work_bytes: plan.requirements.rewrite_work_bytes,
        },
    ))
}

/// Prepare and execute inline FormulaOwner cell-edge pruning in one bounded
/// operation.
pub fn rewrite_formula_owner_cell_edges(
    source: &[u8],
    removed_internal_owner_ids: &[u32],
    options: DecodeOptions,
) -> Result<(Vec<u8>, FormulaDependencyRewriteReport), DecodeError> {
    execute_formula_owner_cell_edges_removal(
        prepare_formula_owner_cell_edges_removal(source, removed_internal_owner_ids, options)?,
        options,
    )
}

/// Prepare CellRecordTile expanded-edge pruning.
pub fn prepare_cell_record_tile_edges_removal<'source>(
    source: &'source [u8],
    removed_internal_owner_ids: &[u32],
    options: DecodeOptions,
) -> Result<CellRecordTileEdgesRemovalPlan<'source>, DecodeError> {
    let mut work = WorkBudget::new(options.max_work_bytes);
    let mut edge_count = InternalOwnerEdgeCount::default();
    let (_, source_report) =
        decode_cell_record_tile_with_visitor(source, options, &mut edge_count)?;
    work.charge_report(source_report)?;
    validate_target_width(0, removed_internal_owner_ids.len(), options, &mut work)?;
    let removed_internal_owner_ids = copy_sorted_ids(removed_internal_owner_ids, &mut work)?;
    let preparation_work_bytes = work.observed();
    let execution_work_bytes = execution_work_upper_bound(
        source.len(),
        source.len(),
        removed_internal_owner_ids.len(),
        edge_count.count,
        EXECUTION_SCAN_PASSES,
    )?;
    let total_work_bytes = preparation_work_bytes
        .checked_add(execution_work_bytes)
        .ok_or_else(DecodeError::invalid)?;
    enforce_work(total_work_bytes, options)?;
    let requirements = requirements(
        source_report,
        source_report.fields(),
        source.len(),
        total_work_bytes,
        options,
    )?;
    Ok(CellRecordTileEdgesRemovalPlan {
        source,
        removed_internal_owner_ids,
        preparation_work_bytes,
        requirements,
    })
}

/// Execute prepared CellRecordTile expanded-edge pruning.
pub fn execute_cell_record_tile_edges_removal(
    plan: CellRecordTileEdgesRemovalPlan<'_>,
    options: DecodeOptions,
) -> Result<(Vec<u8>, FormulaDependencyRewriteReport), DecodeError> {
    enforce_requirements(plan.requirements, options)?;
    if plan.preparation_work_bytes > options.max_work_bytes {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: plan.preparation_work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    let mut work = WorkBudget::new(options.max_work_bytes - plan.preparation_work_bytes);
    let (data, removed_values) = rewrite_cell_record_tile_edges_inner(
        plan.source,
        &plan.removed_internal_owner_ids,
        options,
        &mut work,
    )?;
    if data.len() > plan.requirements.output_bytes {
        return Err(DecodeError::invalid());
    }
    let output_bytes = data.len();
    let (_, result) = decode_cell_record_tile_with_report(&data, options)?;
    work.charge_report(result)?;
    if work.observed()
        > plan
            .requirements
            .rewrite_work_bytes
            .checked_sub(plan.preparation_work_bytes)
            .ok_or_else(DecodeError::invalid)?
    {
        return Err(DecodeError::invalid());
    }
    Ok((
        data,
        FormulaDependencyRewriteReport {
            source: plan.requirements.source,
            result,
            output_bytes,
            removed_values,
            rewrite_work_bytes: plan.requirements.rewrite_work_bytes,
        },
    ))
}

/// Prepare and execute CellRecordTile expanded-edge pruning in one bounded
/// operation.
pub fn rewrite_cell_record_tile_edges(
    source: &[u8],
    removed_internal_owner_ids: &[u32],
    options: DecodeOptions,
) -> Result<(Vec<u8>, FormulaDependencyRewriteReport), DecodeError> {
    execute_cell_record_tile_edges_removal(
        prepare_cell_record_tile_edges_removal(source, removed_internal_owner_ids, options)?,
        options,
    )
}

fn requirements(
    source: DecodeReport,
    result_fields: usize,
    output_bytes: usize,
    rewrite_work_bytes: usize,
    options: DecodeOptions,
) -> Result<FormulaDependencyRewriteRequirements, DecodeError> {
    if source.source_bytes() > options.max_message_bytes {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: source.source_bytes(),
            maximum: options.max_message_bytes,
        }));
    }
    if output_bytes > options.max_message_bytes {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: output_bytes,
            maximum: options.max_message_bytes,
        }));
    }
    if result_fields > options.max_fields {
        return Err(DecodeError::limited(DecodeLimit::Fields {
            observed: result_fields,
            maximum: options.max_fields,
        }));
    }
    if source.max_depth() > options.recursion_limit {
        return Err(DecodeError::limited(DecodeLimit::Nesting {
            observed: source.max_depth(),
            maximum: options.recursion_limit,
        }));
    }
    if source.references() > options.max_references {
        return Err(DecodeError::limited(DecodeLimit::References {
            observed: source.references(),
            maximum: options.max_references,
        }));
    }
    if source.text_bytes() > options.max_text_bytes {
        return Err(DecodeError::limited(DecodeLimit::Text {
            observed: source.text_bytes(),
            maximum: options.max_text_bytes,
        }));
    }
    Ok(FormulaDependencyRewriteRequirements {
        source,
        result_fields,
        result_max_depth: source.max_depth(),
        result_references: source.references(),
        result_reference_bytes: source.reference_bytes(),
        result_text_bytes: source.text_bytes(),
        output_bytes,
        rewrite_work_bytes,
    })
}

fn enforce_requirements(
    requirements: FormulaDependencyRewriteRequirements,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    if requirements.source.source_bytes() > options.max_message_bytes {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: requirements.source.source_bytes(),
            maximum: options.max_message_bytes,
        }));
    }
    if requirements.output_bytes > options.max_message_bytes {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: requirements.output_bytes,
            maximum: options.max_message_bytes,
        }));
    }
    if requirements.source.fields() > options.max_fields {
        return Err(DecodeError::limited(DecodeLimit::Fields {
            observed: requirements.source.fields(),
            maximum: options.max_fields,
        }));
    }
    if requirements.result_fields > options.max_fields {
        return Err(DecodeError::limited(DecodeLimit::Fields {
            observed: requirements.result_fields,
            maximum: options.max_fields,
        }));
    }
    if requirements.source.max_depth() > options.recursion_limit
        || requirements.result_max_depth > options.recursion_limit
    {
        let observed = requirements
            .source
            .max_depth()
            .max(requirements.result_max_depth);
        return Err(DecodeError::limited(DecodeLimit::Nesting {
            observed,
            maximum: options.recursion_limit,
        }));
    }
    if requirements.source.references() > options.max_references
        || requirements.result_references > options.max_references
    {
        let observed = requirements
            .source
            .references()
            .max(requirements.result_references);
        return Err(DecodeError::limited(DecodeLimit::References {
            observed,
            maximum: options.max_references,
        }));
    }
    if requirements.source.reference_bytes() > options.max_message_bytes
        || requirements.result_reference_bytes > options.max_message_bytes
    {
        let observed = requirements
            .source
            .reference_bytes()
            .max(requirements.result_reference_bytes);
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed,
            maximum: options.max_message_bytes,
        }));
    }
    if requirements.source.text_bytes() > options.max_text_bytes
        || requirements.result_text_bytes > options.max_text_bytes
    {
        let observed = requirements
            .source
            .text_bytes()
            .max(requirements.result_text_bytes);
        return Err(DecodeError::limited(DecodeLimit::Text {
            observed,
            maximum: options.max_text_bytes,
        }));
    }
    enforce_work(requirements.rewrite_work_bytes, options)?;
    Ok(())
}

fn enforce_work(observed: usize, options: DecodeOptions) -> Result<(), DecodeError> {
    if observed > options.max_work_bytes {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed,
            maximum: options.max_work_bytes,
        }));
    }
    Ok(())
}

fn validate_target_width(
    owner_ids: usize,
    internal_owner_ids: usize,
    options: DecodeOptions,
    work: &mut WorkBudget,
) -> Result<(), DecodeError> {
    let observed = owner_ids
        .checked_add(internal_owner_ids)
        .ok_or_else(DecodeError::invalid)?;
    if observed > options.max_fields {
        return Err(DecodeError::limited(DecodeLimit::Fields {
            observed,
            maximum: options.max_fields,
        }));
    }
    work.charge(observed)?;
    Ok(())
}

fn copy_sorted_ids<T: Copy + Ord>(
    source: &[T],
    work: &mut WorkBudget,
) -> Result<Vec<T>, DecodeError> {
    work.charge(source.len())?;
    let mut output = Vec::new();
    output.try_reserve_exact(source.len()).map_err(|_| {
        DecodeError::limited(DecodeLimit::Allocation {
            requested: source.len(),
        })
    })?;
    output.extend_from_slice(source);
    output.sort_unstable();
    output.dedup();
    work.charge(sort_work(source.len())?)?;
    Ok(output)
}

fn contains_u64(values: &[u64], value: u64, work: &mut WorkBudget) -> Result<bool, DecodeError> {
    work.charge(search_work(values.len())?)?;
    Ok(values.binary_search(&value).is_ok())
}

fn contains_u32(values: &[u32], value: u32, work: &mut WorkBudget) -> Result<bool, DecodeError> {
    work.charge(search_work(values.len())?)?;
    Ok(values.binary_search(&value).is_ok())
}

fn sort_work(length: usize) -> Result<usize, DecodeError> {
    length
        .checked_mul(ceil_log2(length.saturating_add(1)))
        .ok_or_else(DecodeError::invalid)
}

fn search_work(length: usize) -> Result<usize, DecodeError> {
    ceil_log2(length.saturating_add(1))
        .checked_add(1)
        .ok_or_else(DecodeError::invalid)
}

fn ceil_log2(value: usize) -> usize {
    if value <= 1 {
        0
    } else {
        usize::BITS as usize - (value - 1).leading_zeros() as usize
    }
}

fn execution_work_upper_bound(
    source_bytes: usize,
    output_bytes: usize,
    target_count: usize,
    membership_queries: usize,
    scan_multiplier: usize,
) -> Result<usize, DecodeError> {
    let scans = source_bytes
        .checked_mul(scan_multiplier)
        .ok_or_else(DecodeError::invalid)?;
    let searches = membership_queries
        .checked_mul(search_work(target_count)?)
        .ok_or_else(DecodeError::invalid)?;
    scans
        .checked_add(searches)
        .and_then(|value| value.checked_add(output_bytes))
        .ok_or_else(DecodeError::invalid)
}

fn measure_calculation_engine_output(
    source: &[u8],
    owner_ids: &[u64],
    internal_owner_ids: &[u32],
    formula_count: u64,
    options: DecodeOptions,
    work: &mut WorkBudget,
) -> Result<(usize, usize), DecodeError> {
    let mut budget = begin_scan(source, options, work)?;
    let mut remaining = source;
    let mut output_bytes = 0usize;
    let mut membership_queries = 0usize;
    loop {
        let before = remaining.len();
        let Some(field) = wire::next_field(&mut remaining, &mut budget, 1)? else {
            break;
        };
        let raw_len = before
            .checked_sub(remaining.len())
            .ok_or_else(DecodeError::invalid)?;
        let encoded_len = if field.number == 2 {
            let tracker = field.bytes()?;
            let (payload_len, queries) = measure_tracker_output(
                tracker,
                owner_ids,
                internal_owner_ids,
                formula_count,
                options,
                work,
            )?;
            membership_queries = membership_queries
                .checked_add(queries)
                .ok_or_else(DecodeError::invalid)?;
            encoded_length_delimited_field_length(field.number, payload_len)?
        } else {
            raw_len
        };
        output_bytes = output_bytes
            .checked_add(encoded_len)
            .ok_or_else(DecodeError::invalid)?;
    }
    finish_scan(work, budget, source.len())?;
    if output_bytes > options.max_message_bytes {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: output_bytes,
            maximum: options.max_message_bytes,
        }));
    }
    Ok((output_bytes, membership_queries))
}

fn measure_tracker_output(
    source: &[u8],
    owner_ids: &[u64],
    internal_owner_ids: &[u32],
    formula_count: u64,
    options: DecodeOptions,
    work: &mut WorkBudget,
) -> Result<(usize, usize), DecodeError> {
    let (tracker, tracker_report) = decode_dependency_tracker_with_report(source, options)?;
    work.charge_report(tracker_report)?;
    let previous_count = tracker.number_of_formulas().unwrap_or_default();
    let replacement_count = previous_count
        .checked_sub(formula_count)
        .ok_or_else(DecodeError::invalid)?;
    let mut budget = begin_scan(source, options, work)?;
    let mut remaining = source;
    let mut output_bytes = 0usize;
    let mut membership_queries = 0usize;
    let mut count_present = false;
    loop {
        let before = remaining.len();
        let Some(field) = wire::next_field(&mut remaining, &mut budget, 1)? else {
            break;
        };
        let raw_len = before
            .checked_sub(remaining.len())
            .ok_or_else(DecodeError::invalid)?;
        match field.number {
            3 => {
                let map = field.bytes()?;
                let (payload, queries) =
                    measure_owner_map_output(map, internal_owner_ids, options, work)?;
                membership_queries = membership_queries
                    .checked_add(queries)
                    .ok_or_else(DecodeError::invalid)?;
                output_bytes = output_bytes
                    .checked_add(encoded_length_delimited_field_length(3, payload)?)
                    .ok_or_else(DecodeError::invalid)?;
            },
            5 => {
                let _ = field.varint()?;
                count_present = true;
                output_bytes = output_bytes
                    .checked_add(encoded_varint_field_length(5, replacement_count)?)
                    .ok_or_else(DecodeError::invalid)?;
            },
            6 => {
                let reference = field.bytes()?;
                membership_queries = membership_queries
                    .checked_add(1)
                    .ok_or_else(DecodeError::invalid)?;
                let identifier = reference_identifier(reference, options, work)?;
                if !contains_u64(owner_ids, identifier, work)? {
                    output_bytes = output_bytes
                        .checked_add(encoded_length_delimited_field_length(6, reference.len())?)
                        .ok_or_else(DecodeError::invalid)?;
                }
            },
            _ => {
                output_bytes = output_bytes
                    .checked_add(raw_len)
                    .ok_or_else(DecodeError::invalid)?;
            },
        }
    }
    finish_scan(work, budget, source.len())?;
    if !count_present {
        output_bytes = output_bytes
            .checked_add(encoded_varint_field_length(5, replacement_count)?)
            .ok_or_else(DecodeError::invalid)?;
    }
    if output_bytes > options.max_message_bytes {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: output_bytes,
            maximum: options.max_message_bytes,
        }));
    }
    Ok((output_bytes, membership_queries))
}

fn measure_owner_map_output(
    source: &[u8],
    internal_owner_ids: &[u32],
    options: DecodeOptions,
    work: &mut WorkBudget,
) -> Result<(usize, usize), DecodeError> {
    let mut budget = begin_scan(source, options, work)?;
    let mut remaining = source;
    let mut output_bytes = 0usize;
    let mut membership_queries = 0usize;
    loop {
        let before = remaining.len();
        let Some(field) = wire::next_field(&mut remaining, &mut budget, 1)? else {
            break;
        };
        let raw_len = before
            .checked_sub(remaining.len())
            .ok_or_else(DecodeError::invalid)?;
        if field.number == 1 {
            let entry = field.bytes()?;
            membership_queries = membership_queries
                .checked_add(1)
                .ok_or_else(DecodeError::invalid)?;
            let entry_id = owner_map_entry_internal_id(entry, options, work)?;
            if !contains_u32(internal_owner_ids, entry_id, work)? {
                output_bytes = output_bytes
                    .checked_add(encoded_length_delimited_field_length(1, entry.len())?)
                    .ok_or_else(DecodeError::invalid)?;
            }
        } else {
            output_bytes = output_bytes
                .checked_add(raw_len)
                .ok_or_else(DecodeError::invalid)?;
        }
    }
    finish_scan(work, budget, source.len())?;
    Ok((output_bytes, membership_queries))
}

fn assemble_calculation_engine_owner_removal(
    source: &[u8],
    owner_ids: &[u64],
    internal_owner_ids: &[u32],
    formula_count: u64,
    options: DecodeOptions,
    work: &mut WorkBudget,
) -> Result<(Vec<u8>, usize), DecodeError> {
    let mut removed_values = 0usize;
    let data = rewrite_message(source, options, work, |field, _raw, work| {
        if field.number != 2 {
            return Ok(FieldAction::Keep);
        }
        let tracker = field.bytes()?;
        let (tracker, removed) = rewrite_tracker(
            tracker,
            owner_ids,
            internal_owner_ids,
            formula_count,
            options,
            work,
        )?;
        removed_values = removed_values
            .checked_add(removed)
            .ok_or_else(DecodeError::invalid)?;
        Ok(FieldAction::LengthDelimited(tracker))
    })?;
    Ok((data, removed_values))
}

fn rewrite_tracker(
    source: &[u8],
    owner_ids: &[u64],
    internal_owner_ids: &[u32],
    formula_count: u64,
    options: DecodeOptions,
    work: &mut WorkBudget,
) -> Result<(Vec<u8>, usize), DecodeError> {
    let (tracker, tracker_report) = decode_dependency_tracker_with_report(source, options)?;
    work.charge_report(tracker_report)?;
    let previous_count = tracker.number_of_formulas().unwrap_or_default();
    let replacement_count = previous_count
        .checked_sub(formula_count)
        .ok_or_else(DecodeError::invalid)?;
    let mut removed_values = 0usize;
    let data = rewrite_message(source, options, work, |field, _raw, work| {
        match field.number {
            3 => {
                let map = field.bytes()?;
                let (map, removed) = rewrite_owner_map(map, internal_owner_ids, options, work)?;
                removed_values = removed_values
                    .checked_add(removed)
                    .ok_or_else(DecodeError::invalid)?;
                Ok(FieldAction::LengthDelimited(map))
            },
            6 => {
                let reference = field.bytes()?;
                let identifier = reference_identifier(reference, options, work)?;
                if contains_u64(owner_ids, identifier, work)? {
                    removed_values = removed_values
                        .checked_add(1)
                        .ok_or_else(DecodeError::invalid)?;
                    Ok(FieldAction::Remove)
                } else {
                    Ok(FieldAction::Keep)
                }
            },
            _ => Ok(FieldAction::Keep),
        }
    })?;
    let data = patch_or_append_varint(&data, 5, replacement_count, options, work)?;
    Ok((data, removed_values))
}

fn rewrite_owner_map(
    source: &[u8],
    internal_owner_ids: &[u32],
    options: DecodeOptions,
    work: &mut WorkBudget,
) -> Result<(Vec<u8>, usize), DecodeError> {
    let mut removed_values = 0usize;
    let data = rewrite_message(source, options, work, |field, _raw, work| {
        if field.number != 1 {
            return Ok(FieldAction::Keep);
        }
        let entry = field.bytes()?;
        let internal_owner_id = owner_map_entry_internal_id(entry, options, work)?;
        if contains_u32(internal_owner_ids, internal_owner_id, work)? {
            removed_values = removed_values
                .checked_add(1)
                .ok_or_else(DecodeError::invalid)?;
            Ok(FieldAction::Remove)
        } else {
            Ok(FieldAction::Keep)
        }
    })?;
    Ok((data, removed_values))
}

fn owner_map_entry_internal_id(
    source: &[u8],
    options: DecodeOptions,
    work: &mut WorkBudget,
) -> Result<u32, DecodeError> {
    let mut budget = begin_scan(source, options, work)?;
    let child_depth = 2;
    let mut internal_owner_id = None;
    let mut owner_id = false;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, &mut budget, 1)? {
        match field.number {
            1 => set_once(
                &mut internal_owner_id,
                wire::canonical_u32(field.varint()?)?,
            )?,
            2 => {
                let payload = field.bytes()?;
                decode_cfuuid_in(payload, &mut budget, child_depth)?;
                if owner_id {
                    return Err(DecodeError::invalid());
                }
                owner_id = true;
            },
            _ => {},
        }
    }
    finish_scan(work, budget, source.len())?;
    if !owner_id {
        return Err(DecodeError::invalid());
    }
    internal_owner_id.ok_or_else(DecodeError::invalid)
}

fn reference_identifier(
    source: &[u8],
    options: DecodeOptions,
    work: &mut WorkBudget,
) -> Result<u64, DecodeError> {
    let mut budget = begin_scan(source, options, work)?;
    let identifier = wire::decode_reference(source, &mut budget, 1)?.identifier();
    finish_scan(work, budget, source.len())?;
    Ok(identifier)
}

fn rewrite_formula_owner_cell_edges_inner(
    source: &[u8],
    removed_internal_owner_ids: &[u32],
    options: DecodeOptions,
    work: &mut WorkBudget,
) -> Result<(Vec<u8>, usize), DecodeError> {
    let mut removed_values = 0usize;
    let mut leaf = |expanded_edges: &[u8], work: &mut WorkBudget| {
        let (rewritten, removed) =
            prune_expanded_edges(expanded_edges, removed_internal_owner_ids, options, work)?;
        removed_values = removed_values
            .checked_add(removed)
            .ok_or_else(DecodeError::invalid)?;
        Ok(rewritten)
    };
    let data = rewrite_length_delimited_path(source, &[4, 1, 6], options, work, &mut leaf)?;
    Ok((data, removed_values))
}

fn rewrite_cell_record_tile_edges_inner(
    source: &[u8],
    removed_internal_owner_ids: &[u32],
    options: DecodeOptions,
    work: &mut WorkBudget,
) -> Result<(Vec<u8>, usize), DecodeError> {
    let mut removed_values = 0usize;
    let mut leaf = |expanded_edges: &[u8], work: &mut WorkBudget| {
        let (rewritten, removed) =
            prune_expanded_edges(expanded_edges, removed_internal_owner_ids, options, work)?;
        removed_values = removed_values
            .checked_add(removed)
            .ok_or_else(DecodeError::invalid)?;
        Ok(rewritten)
    };
    let data = rewrite_length_delimited_path(source, &[4, 6], options, work, &mut leaf)?;
    Ok((data, removed_values))
}

fn prune_expanded_edges(
    source: &[u8],
    removed_internal_owner_ids: &[u32],
    options: DecodeOptions,
    work: &mut WorkBudget,
) -> Result<(Vec<u8>, usize), DecodeError> {
    let values = collect_expanded_edges(source, options, work)?;
    let mut remove = Vec::new();
    work.charge(values[4].len())?;
    remove.try_reserve_exact(values[4].len()).map_err(|_| {
        DecodeError::limited(DecodeLimit::Allocation {
            requested: values[4].len(),
        })
    })?;
    let mut removed_values = 0usize;
    for internal_owner_id in &values[4] {
        let should_remove = contains_u32(removed_internal_owner_ids, *internal_owner_id, work)?;
        remove.push(should_remove);
        removed_values = removed_values
            .checked_add(usize::from(should_remove))
            .ok_or_else(DecodeError::invalid)?;
    }
    if removed_values == 0 {
        return Ok((copy_bytes(source, work)?, 0));
    }
    let mut cursors = [0usize; 3];
    let data = rewrite_message(source, options, work, |field, _raw, work| {
        let index = match field.number {
            3 => 0,
            4 => 1,
            5 => 2,
            _ => return Ok(FieldAction::Keep),
        };
        if let Ok(value) = field.varint() {
            let position = cursors[index];
            cursors[index] = position.checked_add(1).ok_or_else(DecodeError::invalid)?;
            if remove.get(position).copied().unwrap_or(false) {
                return Ok(FieldAction::Remove);
            }
            let _ = value;
            return Ok(FieldAction::Keep);
        }
        let payload = field.bytes()?;
        let mut packed = payload;
        let mut retained = Vec::new();
        while !packed.is_empty() {
            let value = wire::take_canonical_varint(&mut packed)?;
            let position = cursors[index];
            cursors[index] = position.checked_add(1).ok_or_else(DecodeError::invalid)?;
            if !remove.get(position).copied().unwrap_or(false) {
                append_varint(&mut retained, value, work)?;
            }
        }
        if retained == payload {
            Ok(FieldAction::Keep)
        } else {
            Ok(FieldAction::LengthDelimited(retained))
        }
    })?;
    if cursors != [values[4].len(); 3] {
        return Err(DecodeError::invalid());
    }
    Ok((data, removed_values))
}

fn collect_expanded_edges(
    source: &[u8],
    options: DecodeOptions,
    work: &mut WorkBudget,
) -> Result<[Vec<u32>; 5], DecodeError> {
    let mut values: [Vec<u32>; 5] = core::array::from_fn(|_| Vec::new());
    let mut budget = begin_scan(source, options, work)?;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, &mut budget, 1)? {
        if !(1..=5).contains(&field.number) {
            continue;
        }
        let index = usize::try_from(field.number - 1).map_err(|_| DecodeError::invalid())?;
        if let Ok(value) = field.varint() {
            push_value(&mut values[index], wire::canonical_u32(value)?, work)?;
        } else {
            let mut packed = field.bytes()?;
            while !packed.is_empty() {
                push_value(
                    &mut values[index],
                    wire::canonical_u32(wire::take_canonical_varint(&mut packed)?)?,
                    work,
                )?;
            }
        }
    }
    finish_scan(work, budget, source.len())?;
    if values[0].len() != values[1].len()
        || values[2].len() != values[3].len()
        || values[2].len() != values[4].len()
    {
        return Err(DecodeError::invalid());
    }
    Ok(values)
}

fn push_value(values: &mut Vec<u32>, value: u32, work: &mut WorkBudget) -> Result<(), DecodeError> {
    work.charge(1)?;
    values.try_reserve(1).map_err(|_| {
        DecodeError::limited(DecodeLimit::Allocation {
            requested: values.len().saturating_add(1),
        })
    })?;
    values.push(value);
    Ok(())
}

#[derive(Debug)]
enum FieldAction {
    Keep,
    Remove,
    LengthDelimited(Vec<u8>),
    Varint(u64),
}

fn rewrite_message<F>(
    source: &[u8],
    options: DecodeOptions,
    work: &mut WorkBudget,
    mut action: F,
) -> Result<Vec<u8>, DecodeError>
where
    F: FnMut(wire::Field<'_>, &[u8], &mut WorkBudget) -> Result<FieldAction, DecodeError>,
{
    let mut budget = begin_scan(source, options, work)?;
    let mut remaining = source;
    let mut output = Vec::new();
    while !remaining.is_empty() {
        let before = remaining.len();
        let field =
            wire::next_field(&mut remaining, &mut budget, 1)?.ok_or_else(DecodeError::invalid)?;
        let after = remaining.len();
        let start = source
            .len()
            .checked_sub(before)
            .ok_or_else(DecodeError::invalid)?;
        let end = source
            .len()
            .checked_sub(after)
            .ok_or_else(DecodeError::invalid)?;
        let raw = source.get(start..end).ok_or_else(DecodeError::invalid)?;
        match action(field, raw, work)? {
            FieldAction::Keep => append_bytes(&mut output, raw, options, work)?,
            FieldAction::Remove => {},
            FieldAction::LengthDelimited(payload) => {
                append_length_delimited(&mut output, field.number, &payload, options, work)?;
            },
            FieldAction::Varint(value) => {
                append_varint_field(&mut output, field.number, value, options, work)?;
            },
        }
    }
    finish_scan(work, budget, source.len())?;
    Ok(output)
}

fn rewrite_length_delimited_path<F>(
    source: &[u8],
    path: &[u32],
    options: DecodeOptions,
    work: &mut WorkBudget,
    leaf: &mut F,
) -> Result<Vec<u8>, DecodeError>
where
    F: FnMut(&[u8], &mut WorkBudget) -> Result<Vec<u8>, DecodeError>,
{
    let field_number = *path.first().ok_or_else(DecodeError::invalid)?;
    let remainder = &path[1..];
    rewrite_message(source, options, work, |field, _raw, work| {
        if field.number != field_number {
            return Ok(FieldAction::Keep);
        }
        let payload = field.bytes()?;
        let rewritten = if remainder.is_empty() {
            leaf(payload, work)?
        } else {
            rewrite_length_delimited_path(payload, remainder, options, work, leaf)?
        };
        if rewritten == payload {
            Ok(FieldAction::Keep)
        } else {
            Ok(FieldAction::LengthDelimited(rewritten))
        }
    })
}

fn patch_or_append_varint(
    source: &[u8],
    field_number: u32,
    replacement: u64,
    options: DecodeOptions,
    work: &mut WorkBudget,
) -> Result<Vec<u8>, DecodeError> {
    let mut present = false;
    let mut output = rewrite_message(source, options, work, |field, _raw, _work| {
        if field.number != field_number {
            return Ok(FieldAction::Keep);
        }
        let _ = field.varint()?;
        if present {
            return Err(DecodeError::invalid());
        }
        present = true;
        Ok(FieldAction::Varint(replacement))
    })?;
    if !present {
        append_varint_field(&mut output, field_number, replacement, options, work)?;
    }
    Ok(output)
}

fn field_encoded_length(field_number: u32, wire_type: u8) -> Result<usize, DecodeError> {
    let tag = (u64::from(field_number) << 3) | u64::from(wire_type);
    Ok(encoded_varint_len(tag))
}

fn encoded_length_delimited_field_length(
    field_number: u32,
    payload_length: usize,
) -> Result<usize, DecodeError> {
    field_encoded_length(field_number, 2)?
        .checked_add(encoded_varint_len(
            u64::try_from(payload_length).map_err(|_| DecodeError::invalid())?,
        ))
        .and_then(|value| value.checked_add(payload_length))
        .ok_or_else(DecodeError::invalid)
}

fn encoded_varint_field_length(field_number: u32, value: u64) -> Result<usize, DecodeError> {
    field_encoded_length(field_number, 0)?
        .checked_add(encoded_varint_len(value))
        .ok_or_else(DecodeError::invalid)
}

fn encoded_varint_len(value: u64) -> usize {
    if value == 0 {
        1
    } else {
        (u64::BITS as usize - value.leading_zeros() as usize).div_ceil(7)
    }
}

fn append_varint(
    output: &mut Vec<u8>,
    mut value: u64,
    work: &mut WorkBudget,
) -> Result<(), DecodeError> {
    let length = encoded_varint_len(value);
    work.charge(length)?;
    output.try_reserve_exact(length).map_err(|_| {
        DecodeError::limited(DecodeLimit::Allocation {
            requested: output.len().saturating_add(length),
        })
    })?;
    loop {
        let mut byte = u8::try_from(value & 0x7f).map_err(|_| DecodeError::invalid())?;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output.push(byte);
        if value == 0 {
            return Ok(());
        }
    }
}

fn append_length_delimited(
    output: &mut Vec<u8>,
    field_number: u32,
    payload: &[u8],
    options: DecodeOptions,
    work: &mut WorkBudget,
) -> Result<(), DecodeError> {
    let additional = encoded_length_delimited_field_length(field_number, payload.len())?;
    ensure_output_capacity(output, additional, options, work)?;
    append_varint_unchecked(output, (u64::from(field_number) << 3) | 2);
    append_varint_unchecked(
        output,
        u64::try_from(payload.len()).map_err(|_| DecodeError::invalid())?,
    );
    output.extend_from_slice(payload);
    Ok(())
}

fn append_varint_field(
    output: &mut Vec<u8>,
    field_number: u32,
    value: u64,
    options: DecodeOptions,
    work: &mut WorkBudget,
) -> Result<(), DecodeError> {
    let additional = encoded_varint_field_length(field_number, value)?;
    ensure_output_capacity(output, additional, options, work)?;
    append_varint_unchecked(output, u64::from(field_number) << 3);
    append_varint_unchecked(output, value);
    Ok(())
}

fn append_bytes(
    output: &mut Vec<u8>,
    bytes: &[u8],
    options: DecodeOptions,
    work: &mut WorkBudget,
) -> Result<(), DecodeError> {
    ensure_output_capacity(output, bytes.len(), options, work)?;
    output.extend_from_slice(bytes);
    Ok(())
}

fn append_varint_unchecked(output: &mut Vec<u8>, mut value: u64) {
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

fn ensure_output_capacity(
    output: &mut Vec<u8>,
    additional: usize,
    options: DecodeOptions,
    work: &mut WorkBudget,
) -> Result<(), DecodeError> {
    let requested = output
        .len()
        .checked_add(additional)
        .ok_or_else(DecodeError::invalid)?;
    if requested > options.max_message_bytes {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: requested,
            maximum: options.max_message_bytes,
        }));
    }
    work.charge(additional)?;
    if output.capacity() < requested {
        output
            .try_reserve_exact(requested - output.len())
            .map_err(|_| DecodeError::limited(DecodeLimit::Allocation { requested }))?;
    }
    Ok(())
}

fn copy_bytes(source: &[u8], work: &mut WorkBudget) -> Result<Vec<u8>, DecodeError> {
    work.charge(source.len())?;
    let mut output = Vec::new();
    output.try_reserve_exact(source.len()).map_err(|_| {
        DecodeError::limited(DecodeLimit::Allocation {
            requested: source.len(),
        })
    })?;
    output.extend_from_slice(source);
    Ok(output)
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "Focused dependency-removal fixtures use exact wire construction."
)]
mod tests {
    use super::*;

    fn varint(output: &mut Vec<u8>, mut value: u64) {
        loop {
            let mut byte = u8::try_from(value & 0x7f).unwrap();
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

    fn key(output: &mut Vec<u8>, field: u32, wire_type: u8) {
        varint(output, (u64::from(field) << 3) | u64::from(wire_type));
    }

    fn v(output: &mut Vec<u8>, field: u32, value: u64) {
        key(output, field, 0);
        varint(output, value);
    }

    fn b(output: &mut Vec<u8>, field: u32, payload: &[u8]) {
        key(output, field, 2);
        varint(output, u64::try_from(payload.len()).unwrap());
        output.extend_from_slice(payload);
    }

    fn reference(identifier: u64) -> Vec<u8> {
        let mut output = Vec::new();
        v(&mut output, 1, identifier);
        output
    }

    fn uuid(seed: u64) -> Vec<u8> {
        let mut output = Vec::new();
        v(&mut output, 1, seed);
        v(&mut output, 2, seed + 1);
        output
    }

    fn cfuuid(seed: u64) -> Vec<u8> {
        let mut output = Vec::new();
        v(&mut output, 2, seed as u32 as u64);
        v(&mut output, 3, (seed >> 32) as u32 as u64);
        v(&mut output, 4, (seed + 1) as u32 as u64);
        v(&mut output, 5, ((seed + 1) >> 32) as u32 as u64);
        output
    }

    fn owner_map_entry(internal_owner_id: u32, seed: u64) -> Vec<u8> {
        let mut output = Vec::new();
        v(&mut output, 1, u64::from(internal_owner_id));
        b(&mut output, 2, &cfuuid(seed));
        output
    }

    fn tracker() -> Vec<u8> {
        let mut map = Vec::new();
        b(&mut map, 1, &owner_map_entry(7, 70));
        b(&mut map, 1, &owner_map_entry(8, 80));
        v(&mut map, 2, 999);
        let mut output = Vec::new();
        b(&mut output, 3, &map);
        v(&mut output, 5, 3);
        b(&mut output, 6, &reference(101));
        b(&mut output, 6, &reference(102));
        b(&mut output, 6, &reference(103));
        output
    }

    fn tracker_with_owner_count(count: usize) -> Vec<u8> {
        let mut map = Vec::new();
        let mut output = Vec::new();
        for index in 0..count {
            let internal = u32::try_from(index + 1).unwrap();
            b(&mut map, 1, &owner_map_entry(internal, 70 + index as u64));
        }
        b(&mut output, 3, &map);
        v(&mut output, 5, u64::try_from(count).unwrap());
        for index in 0..count {
            b(
                &mut output,
                6,
                &reference(u64::try_from(index + 1).unwrap()),
            );
        }
        output
    }

    fn engine() -> Vec<u8> {
        let mut output = Vec::new();
        v(&mut output, 90, 9000);
        b(&mut output, 2, &tracker());
        b(&mut output, 91, b"unknown");
        output
    }

    fn engine_with_owner_count(count: usize) -> Vec<u8> {
        let mut output = Vec::new();
        v(&mut output, 90, 9000);
        b(&mut output, 2, &tracker_with_owner_count(count));
        b(&mut output, 91, b"unknown");
        output
    }

    fn expanded_edges() -> Vec<u8> {
        let mut output = Vec::new();
        v(&mut output, 1, 1);
        v(&mut output, 1, 2);
        v(&mut output, 2, 3);
        v(&mut output, 2, 4);
        v(&mut output, 3, 10);
        v(&mut output, 3, 11);
        v(&mut output, 4, 20);
        v(&mut output, 4, 21);
        v(&mut output, 5, 7);
        v(&mut output, 5, 8);
        output
    }

    fn expanded_edges_for(owners: &[u32]) -> Vec<u8> {
        let mut output = Vec::new();
        v(&mut output, 1, 1);
        v(&mut output, 1, 2);
        v(&mut output, 2, 3);
        v(&mut output, 2, 4);
        for (index, owner) in owners.iter().copied().enumerate() {
            let index = u32::try_from(index).unwrap();
            v(&mut output, 3, u64::from(10 + index));
            v(&mut output, 4, u64::from(20 + index));
            v(&mut output, 5, u64::from(owner));
        }
        b(&mut output, 99, b"edge-unknown");
        output
    }

    fn packed(output: &mut Vec<u8>, field: u32, values: &[u32]) {
        let mut payload = Vec::new();
        for value in values {
            varint(&mut payload, u64::from(*value));
        }
        b(output, field, &payload);
    }

    fn packed_expanded_edges() -> Vec<u8> {
        let mut output = Vec::new();
        packed(&mut output, 1, &[1, 2]);
        packed(&mut output, 2, &[3, 4]);
        packed(&mut output, 3, &[10, 11]);
        packed(&mut output, 4, &[20, 21]);
        packed(&mut output, 5, &[7, 8]);
        b(&mut output, 99, b"packed-unknown");
        output
    }

    fn malformed_expanded_edges() -> Vec<u8> {
        let mut output = Vec::new();
        v(&mut output, 1, 1);
        v(&mut output, 2, 2);
        v(&mut output, 3, 10);
        v(&mut output, 3, 11);
        v(&mut output, 4, 20);
        v(&mut output, 5, 7);
        v(&mut output, 5, 8);
        output
    }

    fn cell_record(edges: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        v(&mut output, 1, 1);
        v(&mut output, 2, 2);
        b(&mut output, 6, edges);
        output
    }

    fn owner(edges: &[u8]) -> Vec<u8> {
        let mut cells = Vec::new();
        b(&mut cells, 1, &cell_record(edges));
        let mut output = Vec::new();
        b(&mut output, 1, &uuid(10));
        v(&mut output, 2, 7);
        b(&mut output, 4, &cells);
        output
    }

    fn tile(edges: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        v(&mut output, 1, 7);
        v(&mut output, 2, 0);
        v(&mut output, 3, 0);
        b(&mut output, 4, &cell_record(edges));
        output
    }

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(
            source.len().saturating_add(128),
            20_000,
            source.len().saturating_mul(64).max(1),
            64,
            20_000,
            20_000,
        )
    }

    fn options_with_work(source: &[u8], work: usize) -> DecodeOptions {
        options_with_limits(source, 20_000, work, 64, 20_000, 20_000)
    }

    fn options_with_limits(
        source: &[u8],
        fields: usize,
        work: usize,
        recursion: u32,
        references: usize,
        text: usize,
    ) -> DecodeOptions {
        DecodeOptions::new(
            source.len().saturating_add(128),
            fields,
            work,
            recursion,
            references,
            text,
        )
    }

    #[derive(Default)]
    struct EdgeValues {
        values: [Vec<u32>; 5],
    }

    impl DependencyVisitor for EdgeValues {
        fn visit_expanded_edge_component(
            &mut self,
            component: ExpandedEdgeComponent,
        ) -> Result<(), DecodeError> {
            let index = match component.kind() {
                ExpandedEdgeKind::LocalRow => 0,
                ExpandedEdgeKind::LocalColumn => 1,
                ExpandedEdgeKind::ExternalRow => 2,
                ExpandedEdgeKind::ExternalColumn => 3,
                ExpandedEdgeKind::InternalOwner => 4,
            };
            self.values[index].push(component.value());
            Ok(())
        }
    }

    fn owner_edge_values(source: &[u8]) -> [Vec<u32>; 5] {
        let mut values = EdgeValues::default();
        decode_formula_owner_dependencies_with_visitor(source, options(source), &mut values)
            .unwrap();
        values.values
    }

    #[test]
    fn engine_owner_removal_prunes_refs_map_and_count_preserving_unknowns() {
        let source = engine();
        let (rewritten, report) =
            rewrite_calculation_engine_owner_removal(&source, &[102], &[7], 1, options(&source))
                .unwrap();
        assert_eq!(report.removed_values(), 2);
        assert!(rewritten.starts_with(&[0xD0, 0x05]));
        assert!(rewritten.windows(7).any(|window| window == b"unknown"));
        let engine = decode_calculation_engine(&rewritten, options(&rewritten)).unwrap();
        let tracker =
            decode_dependency_tracker(engine.dependency_tracker(), options(&rewritten)).unwrap();
        assert_eq!(tracker.number_of_formulas(), Some(2));
    }

    #[test]
    fn engine_owner_removal_keeps_surviving_map_and_reference_families() {
        let source = engine_with_owner_count(3);
        let (rewritten, report) =
            rewrite_calculation_engine_owner_removal(&source, &[2], &[2], 1, options(&source))
                .unwrap();
        assert_eq!(report.removed_values(), 2);
        assert!(
            rewritten
                .windows(owner_map_entry(1, 70).len())
                .any(|window| window == owner_map_entry(1, 70))
        );
        assert!(
            rewritten
                .windows(owner_map_entry(3, 72).len())
                .any(|window| window == owner_map_entry(3, 72))
        );
        assert!(
            !rewritten
                .windows(owner_map_entry(2, 71).len())
                .any(|window| window == owner_map_entry(2, 71))
        );
        assert!(
            rewritten
                .windows(reference(1).len())
                .any(|window| window == reference(1))
        );
        assert!(
            rewritten
                .windows(reference(3).len())
                .any(|window| window == reference(3))
        );
        assert!(
            !rewritten
                .windows(reference(2).len())
                .any(|window| window == reference(2))
        );
    }

    #[test]
    fn inline_and_tiled_edge_removal_keeps_parallel_edge_arrays_valid() {
        let edges = expanded_edges();
        let owner_source = owner(&edges);
        let (owner_rewritten, owner_report) =
            rewrite_formula_owner_cell_edges(&owner_source, &[7], options(&owner_source)).unwrap();
        assert_eq!(owner_report.removed_values(), 1);
        decode_formula_owner_dependencies(&owner_rewritten, options(&owner_rewritten)).unwrap();
        assert!(owner_rewritten.len() < owner_source.len());

        let tile_source = tile(&edges);
        let (tile_rewritten, tile_report) =
            rewrite_cell_record_tile_edges(&tile_source, &[7], options(&tile_source)).unwrap();
        assert_eq!(tile_report.removed_values(), 1);
        decode_cell_record_tile(&tile_rewritten, options(&tile_rewritten)).unwrap();
        assert!(tile_rewritten.len() < tile_source.len());
    }

    #[test]
    fn malformed_owner_map_entry_is_rejected_before_rewrite() {
        let mut map = Vec::new();
        let mut malformed = Vec::new();
        v(&mut malformed, 1, 7);
        b(&mut map, 1, &malformed);
        let mut tracker = Vec::new();
        b(&mut tracker, 3, &map);
        v(&mut tracker, 5, 0);
        let mut source = Vec::new();
        b(&mut source, 2, &tracker);
        assert!(
            rewrite_calculation_engine_owner_removal(&source, &[], &[], 0, options(&source))
                .is_err()
        );
    }

    #[test]
    fn mixed_and_all_expanded_edge_removals_preserve_parallel_arrays_and_unknowns() {
        let edges = expanded_edges_for(&[7, 8, 9]);
        let source = owner(&edges);
        let (mixed, mixed_report) =
            rewrite_formula_owner_cell_edges(&source, &[8], options(&source)).unwrap();
        assert_eq!(mixed_report.removed_values(), 1);
        assert!(
            mixed
                .windows(b"edge-unknown".len())
                .any(|window| window == b"edge-unknown")
        );
        let values = owner_edge_values(&mixed);
        assert_eq!(values[0], vec![1, 2]);
        assert_eq!(values[1], vec![3, 4]);
        assert_eq!(values[2], vec![10, 12]);
        assert_eq!(values[3], vec![20, 22]);
        assert_eq!(values[4], vec![7, 9]);

        let (all, all_report) =
            rewrite_formula_owner_cell_edges(&source, &[7, 8, 9], options(&source)).unwrap();
        assert_eq!(all_report.removed_values(), 3);
        let values = owner_edge_values(&all);
        assert_eq!(values[0], vec![1, 2]);
        assert_eq!(values[1], vec![3, 4]);
        assert!(values[2].is_empty());
        assert!(values[3].is_empty());
        assert!(values[4].is_empty());
    }

    #[test]
    fn packed_expanded_edge_removal_preserves_packed_input_and_unknowns() {
        let edges = packed_expanded_edges();
        let source = owner(&edges);
        let (rewritten, report) =
            rewrite_formula_owner_cell_edges(&source, &[7], options(&source)).unwrap();
        assert_eq!(report.removed_values(), 1);
        assert!(
            rewritten
                .windows(b"packed-unknown".len())
                .any(|window| window == b"packed-unknown")
        );
        let values = owner_edge_values(&rewritten);
        assert_eq!(values[0], vec![1, 2]);
        assert_eq!(values[1], vec![3, 4]);
        assert_eq!(values[2], vec![11]);
        assert_eq!(values[3], vec![21]);
        assert_eq!(values[4], vec![8]);
    }

    #[test]
    fn malformed_expanded_edge_cardinality_fails_before_rewrite() {
        let edges = malformed_expanded_edges();
        let source = owner(&edges);
        assert!(rewrite_formula_owner_cell_edges(&source, &[7], options(&source)).is_err());

        let tile_source = tile(&edges);
        assert!(rewrite_cell_record_tile_edges(&tile_source, &[7], options(&tile_source)).is_err());
    }

    #[test]
    fn aggregate_work_limits_refuse_before_candidate_and_accept_exact_requirement() {
        let source = engine();
        let high = options(&source);
        let plan =
            prepare_calculation_engine_owner_removal(&source, &[102], &[7], 1, high).unwrap();
        let required = plan.requirements().rewrite_work_bytes();
        assert!(required > 0);

        let zero = rewrite_calculation_engine_owner_removal(
            &source,
            &[102],
            &[7],
            1,
            options_with_work(&source, 0),
        )
        .unwrap_err();
        assert!(matches!(
            zero.resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));

        let under = execute_calculation_engine_owner_removal(
            plan,
            options_with_work(&source, required - 1),
        )
        .unwrap_err();
        assert!(matches!(
            under.resource_limit(),
            Some(DecodeLimit::Work {
                observed,
                maximum
            }) if observed == required && maximum == required - 1
        ));

        let exact_options = options_with_work(&source, required);
        let exact_plan =
            prepare_calculation_engine_owner_removal(&source, &[102], &[7], 1, exact_options)
                .unwrap();
        let (_rewritten, report) =
            execute_calculation_engine_owner_removal(exact_plan, exact_options).unwrap();
        assert_eq!(report.rewrite_work_bytes(), required);
    }

    #[test]
    fn edge_rewrites_accept_exact_aggregate_work_and_reject_one_byte_under() {
        let owner_source = owner(&expanded_edges_for(&[7, 8, 9]));
        let high = options(&owner_source);
        let owner_plan =
            prepare_formula_owner_cell_edges_removal(&owner_source, &[8], high).unwrap();
        let owner_required = owner_plan.requirements().rewrite_work_bytes();
        let owner_under = execute_formula_owner_cell_edges_removal(
            owner_plan,
            options_with_work(&owner_source, owner_required - 1),
        )
        .unwrap_err();
        assert!(matches!(
            owner_under.resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));
        let owner_exact_options = options_with_work(&owner_source, owner_required);
        let owner_exact_plan =
            prepare_formula_owner_cell_edges_removal(&owner_source, &[8], owner_exact_options)
                .unwrap();
        execute_formula_owner_cell_edges_removal(owner_exact_plan, owner_exact_options).unwrap();

        let tile_source = tile(&expanded_edges_for(&[7, 8, 9]));
        let high = options(&tile_source);
        let tile_plan = prepare_cell_record_tile_edges_removal(&tile_source, &[8], high).unwrap();
        let tile_required = tile_plan.requirements().rewrite_work_bytes();
        let tile_under = execute_cell_record_tile_edges_removal(
            tile_plan,
            options_with_work(&tile_source, tile_required - 1),
        )
        .unwrap_err();
        assert!(matches!(
            tile_under.resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));
        let tile_exact_options = options_with_work(&tile_source, tile_required);
        let tile_exact_plan =
            prepare_cell_record_tile_edges_removal(&tile_source, &[8], tile_exact_options).unwrap();
        execute_cell_record_tile_edges_removal(tile_exact_plan, tile_exact_options).unwrap();
    }

    #[test]
    fn larger_nested_rewrites_stay_within_prepared_work_upper_bounds() {
        let owners = (1..=128).collect::<Vec<_>>();
        let owner_source = owner(&expanded_edges_for(&owners));
        let high = options_with_limits(
            &owner_source,
            20_000,
            owner_source.len().saturating_mul(256),
            64,
            20_000,
            20_000,
        );
        let plan = prepare_formula_owner_cell_edges_removal(&owner_source, &[64], high).unwrap();
        let required = plan.requirements().rewrite_work_bytes();
        let exact_options = options_with_work(&owner_source, required);
        let exact_plan =
            prepare_formula_owner_cell_edges_removal(&owner_source, &[64], exact_options).unwrap();
        execute_formula_owner_cell_edges_removal(exact_plan, exact_options).unwrap();

        let engine_source = engine_with_owner_count(128);
        let high = options_with_limits(
            &engine_source,
            20_000,
            engine_source.len().saturating_mul(256),
            64,
            20_000,
            20_000,
        );
        let plan = prepare_calculation_engine_owner_removal(&engine_source, &[64], &[64], 1, high)
            .unwrap();
        let required = plan.requirements().rewrite_work_bytes();
        let exact_options = options_with_work(&engine_source, required);
        let exact_plan = prepare_calculation_engine_owner_removal(
            &engine_source,
            &[64],
            &[64],
            1,
            exact_options,
        )
        .unwrap();
        execute_calculation_engine_owner_removal(exact_plan, exact_options).unwrap();
    }

    #[test]
    fn prepared_subordinate_resource_limits_are_checked_before_candidate() {
        let source = engine();
        let high = options(&source);
        let plan =
            prepare_calculation_engine_owner_removal(&source, &[102], &[7], 1, high).unwrap();
        let requirements = plan.requirements();

        let fields = execute_calculation_engine_owner_removal(
            plan,
            options_with_limits(
                &source,
                requirements.result_fields() - 1,
                requirements.rewrite_work_bytes(),
                64,
                20_000,
                20_000,
            ),
        )
        .unwrap_err();
        assert!(matches!(
            fields.resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));

        let plan =
            prepare_calculation_engine_owner_removal(&source, &[102], &[7], 1, high).unwrap();
        let requirements = plan.requirements();
        let nesting = execute_calculation_engine_owner_removal(
            plan,
            options_with_limits(
                &source,
                20_000,
                requirements.rewrite_work_bytes(),
                requirements.result_max_depth().saturating_sub(1).max(1),
                20_000,
                20_000,
            ),
        )
        .unwrap_err();
        assert!(matches!(
            nesting.resource_limit(),
            Some(DecodeLimit::Nesting { .. })
        ));

        let plan =
            prepare_calculation_engine_owner_removal(&source, &[102], &[7], 1, high).unwrap();
        let requirements = plan.requirements();
        let references = execute_calculation_engine_owner_removal(
            plan,
            options_with_limits(
                &source,
                20_000,
                requirements.rewrite_work_bytes(),
                64,
                requirements.result_references() - 1,
                20_000,
            ),
        )
        .unwrap_err();
        assert!(matches!(
            references.resource_limit(),
            Some(DecodeLimit::References { .. })
        ));

        let owner_source = owner(&expanded_edges_for(&[7, 8, 9]));
        let plan =
            prepare_formula_owner_cell_edges_removal(&owner_source, &[8], options(&owner_source))
                .unwrap();
        let requirements = plan.requirements();
        let fields = execute_formula_owner_cell_edges_removal(
            plan,
            options_with_limits(
                &owner_source,
                requirements.result_fields() - 1,
                requirements.rewrite_work_bytes(),
                64,
                20_000,
                20_000,
            ),
        )
        .unwrap_err();
        assert!(matches!(
            fields.resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));

        let tile_source = tile(&expanded_edges_for(&[7, 8, 9]));
        let plan =
            prepare_cell_record_tile_edges_removal(&tile_source, &[8], options(&tile_source))
                .unwrap();
        let requirements = plan.requirements();
        let fields = execute_cell_record_tile_edges_removal(
            plan,
            options_with_limits(
                &tile_source,
                requirements.result_fields() - 1,
                requirements.rewrite_work_bytes(),
                64,
                20_000,
                20_000,
            ),
        )
        .unwrap_err();
        assert!(matches!(
            fields.resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));
    }
}
