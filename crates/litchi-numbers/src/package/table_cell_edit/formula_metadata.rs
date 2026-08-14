//! Exact, generated-free rewrites of the complete Numbers formula dependency graph.
//!
//! This leaf does not allocate or delete IWA objects. Existing owner-to-tile
//! references stay authoritative; a tiled topology is accepted only when its
//! caller-supplied assignments represent the complete final graph exactly.

use core::{fmt, mem::size_of};

use litchi_iwa_common::{
    WireLimits,
    varint::{encode_varint_into, encoded_len},
    wire::{
        NestedFieldEdit, NestedFieldReplacement, WireView, patch_nested_fields_batched_with_limits,
    },
};
use litchi_iwa_protos::numbers_table_cell_dependency_codec as dependency;

const ENGINE_COUNT_PATH: [u32; 2] = [2, 5];
const OWNER_CELL_PATH: [u32; 1] = [4];
const OWNER_RANGE_PATH: [u32; 1] = [5];
const OWNER_TILED_CELL_PATH: [u32; 1] = [13];
const STRICT_WORK_MULTIPLIER: usize = 16;

#[derive(Debug, Clone, Copy)]
pub(super) struct SourceMessage<'a> {
    pub(super) object_id: u64,
    pub(super) payload: &'a [u8],
    /// Exact ArchiveInfo object-reference identifiers, in source order.
    pub(super) object_references: &'a [u64],
}

#[derive(Debug, Clone, Copy)]
pub(super) struct CellTileSource<'a> {
    pub(super) message: SourceMessage<'a>,
    /// Whether this tile and its reference already exist in the source owner.
    /// A false value describes a caller-allocated empty canonical tile that is
    /// published only by the final artifact.
    pub(super) source_present: bool,
    pub(super) tile_column_begin: u32,
    pub(super) tile_row_begin: u32,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct RangeTileSource<'a> {
    pub(super) message: SourceMessage<'a>,
    /// Expected native `to_owner_id`, used to group the complete range set in
    /// one bounded pass before any tile is emitted.
    pub(super) target_owner: u32,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct SourceOwner<'source, 'plan> {
    pub(super) message: SourceMessage<'source>,
    pub(super) internal_owner: u32,
    pub(super) uid_lower: u64,
    pub(super) uid_upper: u64,
    pub(super) cell_tiles: &'plan [CellTileSource<'source>],
    pub(super) range_tiles: &'plan [RangeTileSource<'source>],
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct Precedent {
    pub(super) target_owner: u32,
    pub(super) row: u32,
    pub(super) column: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct Range {
    pub(super) target_owner: u32,
    pub(super) top: u32,
    pub(super) left: u32,
    pub(super) bottom: u32,
    pub(super) right: u32,
}

/// One final formula host. Host owner and target-qualified facts are distinct.
#[derive(Debug, Clone, Copy)]
pub(super) struct FormulaHost<'a> {
    pub(super) owner: u32,
    pub(super) row: u32,
    pub(super) column: u32,
    pub(super) precedents: &'a [Precedent],
    pub(super) ranges: &'a [Range],
    raw_cell_record: Option<&'a [u8]>,
    raw_range_records: &'a [Option<&'a [u8]>],
    cell_tile_object_id: Option<u64>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct LogicalHost {
    pub(super) owner: u32,
    pub(super) row: u32,
    pub(super) column: u32,
    pub(super) precedents: Vec<Precedent>,
    pub(super) ranges: Vec<Range>,
    pub(super) is_in_cycle: bool,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct LogicalGraph<'graph> {
    pub(super) formula_count: u64,
    /// Sorted unique dependency hosts. A final formula with no precedent,
    /// range, or cycle marker is intentionally absent.
    pub(super) hosts: &'graph [LogicalHost],
}

impl<'a> FormulaHost<'a> {
    pub(super) const fn authored(
        owner: u32,
        row: u32,
        column: u32,
        precedents: &'a [Precedent],
        ranges: &'a [Range],
    ) -> Self {
        Self {
            owner,
            row,
            column,
            precedents,
            ranges,
            raw_cell_record: None,
            raw_range_records: &[],
            cell_tile_object_id: None,
        }
    }
}

#[derive(Debug, Clone, Copy)]
struct CompleteGraph<'source, 'plan> {
    pub(super) engine: SourceMessage<'source>,
    pub(super) owners: &'plan [SourceOwner<'source, 'plan>],
    pub(super) hosts: &'plan [FormulaHost<'source>],
    table_owner: u32,
    formula_count: u64,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct HostKey {
    pub(super) owner: u32,
    pub(super) row: u32,
    pub(super) column: u32,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct HostChange<'a> {
    /// Authoritative old formula presence from the BNC/formula-list index.
    pub(super) old: Option<HostKey>,
    /// Authoritative new formula presence and complete AST-derived facts.
    pub(super) new: Option<FormulaHost<'a>>,
    /// Existing destination cell-record tile for `new`, or `None` for inline.
    pub(super) cell_tile_object_id: Option<u64>,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct SourceGraph<'source, 'plan> {
    pub(super) engine: SourceMessage<'source>,
    pub(super) owners: &'plan [SourceOwner<'source, 'plan>],
    /// Resolved table owner whose BNC/formula-list hosts are authoritative.
    pub(super) table_owner: u32,
    /// Complete authoritative source formula-host index derived from BNC and
    /// FormulaList linkage, sorted and unique. It includes zero-edge formulas.
    pub(super) existing_formula_hosts: &'plan [HostKey],
    /// Sorted changed-host overlay. Exact formula/cache no-ops are excluded.
    pub(super) changes: &'plan [HostChange<'plan>],
}

#[derive(Debug, Clone, Copy)]
pub(super) struct Limits {
    pub(super) max_source_bytes: usize,
    pub(super) max_output_bytes: usize,
    pub(super) max_fields: usize,
    pub(super) max_work_bytes: usize,
    pub(super) max_references: usize,
    pub(super) max_messages: usize,
    pub(super) max_hosts: usize,
    pub(super) max_precedents: usize,
    pub(super) max_ranges: usize,
    pub(super) max_retained_bytes: usize,
    pub(super) max_scratch_bytes: usize,
    pub(super) max_allocations: usize,
    pub(super) recursion_limit: u32,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
/// Governed resource upper bounds for the complete source-decode and rewrite.
///
/// Allocation, work, and scratch values are deliberately admitted upper bounds,
/// not allocator telemetry. They are fixed before the governed phase begins.
pub(super) struct Report {
    pub(super) source_bytes: usize,
    pub(super) output_bytes: usize,
    pub(super) fields: usize,
    pub(super) strict_work_bytes: usize,
    pub(super) graph_work_bytes: usize,
    pub(super) references: usize,
    pub(super) reference_bytes: usize,
    pub(super) text_bytes: usize,
    pub(super) max_depth: u32,
    pub(super) hosts: usize,
    pub(super) precedents: usize,
    pub(super) ranges: usize,
    pub(super) changed_messages: usize,
    pub(super) objects: usize,
    pub(super) retained_elements: usize,
    pub(super) allocations: usize,
    pub(super) retained_bytes: usize,
    pub(super) peak_scratch_bytes: usize,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct ExecutionRequirements {
    pub(super) output_bytes: usize,
    pub(super) fields: usize,
    pub(super) work_bytes: usize,
    pub(super) references: usize,
    pub(super) allocations: usize,
    pub(super) peak_scratch_bytes: usize,
    pub(super) retained_bytes: usize,
    pub(super) retained_elements: usize,
    pub(super) objects: usize,
    pub(super) message_edits: usize,
    pub(super) hosts: usize,
    pub(super) precedents: usize,
    pub(super) ranges: usize,
}

pub(super) struct PreparedGraph<'source> {
    engine: SourceMessage<'source>,
    owners: Vec<PreparedOwner<'source>>,
    final_hosts: Vec<OwnedHost<'source>>,
    logical_hosts: Vec<LogicalHost>,
    cell_assignments: Vec<Vec<CellTileSource<'source>>>,
    table_formula_count: u64,
    engine_formula_count: u64,
    table_owner: u32,
    prepare_report: Report,
    requirements: ExecutionRequirements,
    preflight: Preflight,
}

struct PreparedOwner<'source> {
    message: SourceMessage<'source>,
    internal_owner: u32,
    uid_lower: u64,
    uid_upper: u64,
    range_tiles: Vec<RangeTileSource<'source>>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct MessageEdit {
    pub(super) object_id: u64,
    /// `None` proves the complete candidate equals the exact source payload.
    pub(super) payload: Option<Vec<u8>>,
    /// Exact, unchanged ArchiveInfo object-reference identifiers.
    pub(super) object_references: Vec<u64>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct Artifact {
    pub(super) engine: MessageEdit,
    pub(super) owners: Vec<MessageEdit>,
    pub(super) cell_tiles: Vec<MessageEdit>,
    pub(super) range_tiles: Vec<MessageEdit>,
    pub(super) report: Report,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum Error {
    InvalidGraph,
    Limit {
        resource: &'static str,
        observed: usize,
        maximum: usize,
    },
    Allocation {
        requested: usize,
    },
    StrictDependency,
    Wire,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("invalid or over-budget Numbers formula dependency graph")
    }
}

impl std::error::Error for Error {}

/// Strictly prepare the complete final graph without allocating candidate
/// payload or output bytes. Callers may validate [`LogicalGraph`] and stop
/// before [`PreparedGraph::execute`] when cache or cycle policy refuses it.
pub(super) fn prepare_graph<'source>(
    graph: SourceGraph<'source, '_>,
    limits: Limits,
) -> Result<PreparedGraph<'source>, Error> {
    let admission = SourceAdmission::new(graph, limits)?;
    let uid_index = OwnerUidIndex::new(graph.owners)?;
    let mut source = decode_source_hosts(graph, &uid_index, limits)?;
    source.report.source_bytes = admission.source_bytes;
    source.report.allocations = admission.allocations;
    source.report.peak_scratch_bytes = admission.peak_scratch_bytes;
    source.report.graph_work_bytes = admission.graph_work_bytes;
    let source_report = source.report;
    let engine_formula_count = final_formula_count(source.formula_count, graph.changes)?;
    let table_formula_count =
        final_table_formula_count(graph.existing_formula_hosts, graph.changes)?;
    let final_hosts = overlay_changes(
        source.hosts,
        graph.changes,
        graph.owners,
        graph.existing_formula_hosts,
        limits,
    )?;
    let logical_hosts = logical_hosts(
        &final_hosts,
        graph.existing_formula_hosts,
        graph.changes,
        limits,
    )?;
    let borrowed = borrow_hosts(&final_hosts, limits)?;
    let cell_assignments = prepare_cell_assignments(graph.owners, &final_hosts)?;
    let owners = prepare_owner_envelopes(graph.owners)?;
    let mut prepared_owners = try_vec(owners.len())?;
    for (owner, assignments) in owners.iter().zip(&cell_assignments) {
        prepared_owners.push(SourceOwner {
            message: owner.message,
            internal_owner: owner.internal_owner,
            uid_lower: owner.uid_lower,
            uid_upper: owner.uid_upper,
            cell_tiles: assignments,
            range_tiles: &owner.range_tiles,
        });
    }
    let live_source_bytes = rewrite_live_bytes(
        &final_hosts,
        &logical_hosts,
        &owners,
        &uid_index,
        &borrowed,
        &cell_assignments,
        &prepared_owners,
    )?;
    let remaining = remaining_limits(limits, source_report, live_source_bytes)?;
    let complete = CompleteGraph {
        engine: graph.engine,
        owners: &prepared_owners,
        hosts: &borrowed,
        table_owner: graph.table_owner,
        formula_count: engine_formula_count,
    };
    let preflight = Preflight::new(complete, remaining)?;
    let retained_bytes = prepared_graph_retained(
        &final_hosts,
        &logical_hosts,
        &owners,
        &cell_assignments,
        &preflight,
    )?;
    let retained_elements = prepared_graph_elements(
        &final_hosts,
        &logical_hosts,
        &owners,
        &cell_assignments,
        &preflight,
    )?;
    let mut prepare_report = source_report;
    prepare_report.retained_bytes = retained_bytes;
    prepare_report.retained_elements = retained_elements;
    prepare_report.peak_scratch_bytes = prepare_report.peak_scratch_bytes.max(checked_add(
        live_source_bytes,
        preflight.plan_scratch_bytes(),
    )?);
    let mut requirements = preflight.requirements()?;
    prepare_report.objects = requirements.objects;
    // Prepared state is retained (not retained-output charged again), but it
    // remains live while execute allocates its transient candidate buffers.
    // Expose that overlap on the independent peak-scratch axis.
    requirements.peak_scratch_bytes = checked_add(
        requirements.peak_scratch_bytes,
        checked_add(
            retained_bytes,
            checked_add(
                checked_mul(borrowed.capacity(), size_of::<FormulaHost<'static>>())?,
                checked_mul(
                    prepared_owners.capacity(),
                    size_of::<SourceOwner<'static, 'static>>(),
                )?,
            )?,
        )?,
    )?;
    requirements.ensure(Limits {
        max_scratch_bytes: limits.max_scratch_bytes,
        ..remaining
    })?;
    ensure_limit(
        "prepared graph retained bytes",
        prepare_report.retained_bytes,
        limits.max_retained_bytes,
    )?;
    ensure_limit(
        "prepared graph peak scratch",
        prepare_report.peak_scratch_bytes,
        limits.max_scratch_bytes,
    )?;
    drop(prepared_owners);
    drop(borrowed);
    Ok(PreparedGraph {
        engine: graph.engine,
        owners,
        final_hosts,
        logical_hosts,
        cell_assignments,
        table_formula_count,
        engine_formula_count,
        table_owner: graph.table_owner,
        prepare_report,
        requirements,
        preflight,
    })
}

impl<'source> PreparedGraph<'source> {
    pub(super) fn logical_view(&self) -> LogicalGraph<'_> {
        LogicalGraph {
            formula_count: self.table_formula_count,
            hosts: &self.logical_hosts,
        }
    }

    pub(super) const fn prepare_report(&self) -> Report {
        self.prepare_report
    }

    pub(super) const fn execution_requirements(&self) -> ExecutionRequirements {
        self.requirements
    }

    /// Allocate, splice, and strictly reopen candidates only after all
    /// logical consumers have accepted the prepared graph.
    pub(super) fn execute(self, limits: Limits) -> Result<Artifact, Error> {
        self.requirements.ensure(limits)?;
        let PreparedGraph {
            engine,
            owners,
            final_hosts,
            logical_hosts: _,
            cell_assignments,
            table_formula_count: _,
            engine_formula_count,
            table_owner,
            prepare_report: _,
            requirements: _,
            preflight,
        } = self;
        let borrowed = borrow_hosts(&final_hosts, limits)?;
        let mut prepared_owners = try_vec(owners.len())?;
        for (owner, assignments) in owners.iter().zip(&cell_assignments) {
            prepared_owners.push(SourceOwner {
                message: owner.message,
                internal_owner: owner.internal_owner,
                uid_lower: owner.uid_lower,
                uid_upper: owner.uid_upper,
                cell_tiles: assignments,
                range_tiles: &owner.range_tiles,
            });
        }
        execute_complete_graph(
            CompleteGraph {
                engine,
                owners: &prepared_owners,
                hosts: &borrowed,
                table_owner,
                formula_count: engine_formula_count,
            },
            limits,
            preflight,
        )
    }
}

impl ExecutionRequirements {
    fn ensure(self, limits: Limits) -> Result<(), Error> {
        ensure_limit(
            "execution output bytes",
            self.output_bytes,
            limits.max_output_bytes,
        )?;
        ensure_limit("execution fields", self.fields, limits.max_fields)?;
        ensure_limit("execution work", self.work_bytes, limits.max_work_bytes)?;
        ensure_limit(
            "execution references",
            self.references,
            limits.max_references,
        )?;
        ensure_limit(
            "execution allocations",
            self.allocations,
            limits.max_allocations,
        )?;
        ensure_limit(
            "execution peak scratch",
            self.peak_scratch_bytes,
            limits.max_scratch_bytes,
        )?;
        ensure_limit(
            "execution retained bytes",
            self.retained_bytes,
            limits.max_retained_bytes,
        )?;
        ensure_limit("execution hosts", self.hosts, limits.max_hosts)?;
        ensure_limit(
            "execution precedents",
            self.precedents,
            limits.max_precedents,
        )?;
        ensure_limit("execution ranges", self.ranges, limits.max_ranges)
    }
}

/// Test-only convenience wrapper. Production must cross the logical
/// validation barrier before execution.
#[cfg(test)]
fn rewrite_graph(graph: SourceGraph<'_, '_>, limits: Limits) -> Result<Artifact, Error> {
    prepare_graph(graph, limits)?.execute(limits)
}

#[derive(Debug, Clone, Copy)]
struct OwnerUidEntry {
    lower: u64,
    upper: u64,
    internal_owner: u32,
}

#[derive(Debug, Clone, Copy)]
struct ArchiveOwnerReference {
    object_id: u64,
    owner_order: usize,
    occurrence: Option<usize>,
}

struct OwnerUidIndex {
    entries: Vec<OwnerUidEntry>,
}

impl OwnerUidIndex {
    fn new(owners: &[SourceOwner<'_, '_>]) -> Result<Self, Error> {
        let mut entries = try_vec(owners.len())?;
        for owner in owners {
            entries.push(OwnerUidEntry {
                lower: owner.uid_lower,
                upper: owner.uid_upper,
                internal_owner: owner.internal_owner,
            });
        }
        entries.sort_unstable_by_key(|entry| entry.internal_owner);
        if entries
            .windows(2)
            .any(|pair| pair[0].internal_owner == pair[1].internal_owner)
        {
            return Err(Error::InvalidGraph);
        }
        entries.sort_unstable_by_key(|entry| (entry.lower, entry.upper));
        if entries
            .windows(2)
            .any(|pair| (pair[0].lower, pair[0].upper) == (pair[1].lower, pair[1].upper))
        {
            return Err(Error::InvalidGraph);
        }
        Ok(Self { entries })
    }

    fn internal_owner(&self, lower: u64, upper: u64) -> Option<u32> {
        self.entries
            .binary_search_by_key(&(lower, upper), |entry| (entry.lower, entry.upper))
            .ok()
            .map(|index| self.entries[index].internal_owner)
    }
}

fn remaining_limits(
    limits: Limits,
    report: Report,
    live_source_bytes: usize,
) -> Result<Limits, Error> {
    let used_work = checked_add(report.strict_work_bytes, report.graph_work_bytes)?;
    Ok(Limits {
        max_source_bytes: limits.max_source_bytes,
        max_output_bytes: limits.max_output_bytes,
        max_fields: limits
            .max_fields
            .checked_sub(report.fields)
            .ok_or(Error::InvalidGraph)?,
        max_work_bytes: limits
            .max_work_bytes
            .checked_sub(used_work)
            .ok_or(Error::InvalidGraph)?,
        max_references: limits
            .max_references
            .checked_sub(report.references)
            .ok_or(Error::InvalidGraph)?,
        max_messages: limits.max_messages,
        max_hosts: limits.max_hosts,
        max_precedents: limits.max_precedents,
        max_ranges: limits.max_ranges,
        max_retained_bytes: limits.max_retained_bytes,
        max_scratch_bytes: limits
            .max_scratch_bytes
            .checked_sub(live_source_bytes)
            .ok_or(Error::InvalidGraph)?,
        max_allocations: limits
            .max_allocations
            .checked_sub(report.allocations)
            .ok_or(Error::InvalidGraph)?,
        recursion_limit: limits.recursion_limit,
    })
}

fn owned_hosts_retained(hosts: &[OwnedHost<'_>]) -> Result<usize, Error> {
    hosts.iter().try_fold(
        checked_mul(hosts.len(), size_of::<OwnedHost<'static>>())?,
        |sum, host| {
            checked_add(
                sum,
                checked_add(
                    checked_mul(host.precedents.capacity(), size_of::<Precedent>())?,
                    checked_add(
                        checked_mul(host.ranges.capacity(), size_of::<Range>())?,
                        checked_mul(
                            host.raw_range_records.capacity(),
                            size_of::<Option<&[u8]>>(),
                        )?,
                    )?,
                )?,
            )
        },
    )
}

fn owned_hosts_vec_retained(hosts: &Vec<OwnedHost<'_>>) -> Result<usize, Error> {
    checked_add(
        owned_hosts_retained(hosts)?,
        checked_mul(
            hosts.capacity().saturating_sub(hosts.len()),
            size_of::<OwnedHost<'static>>(),
        )?,
    )
}

fn rewrite_live_bytes(
    final_hosts: &Vec<OwnedHost<'_>>,
    logical_hosts: &Vec<LogicalHost>,
    owners: &Vec<PreparedOwner<'_>>,
    uid_index: &OwnerUidIndex,
    borrowed: &Vec<FormulaHost<'_>>,
    cell_assignments: &Vec<Vec<CellTileSource<'_>>>,
    prepared_owners: &Vec<SourceOwner<'_, '_>>,
) -> Result<usize, Error> {
    let nested_assignments = cell_assignments
        .iter()
        .try_fold(0usize, |sum, assignments| {
            checked_add(
                sum,
                checked_mul(assignments.capacity(), size_of::<CellTileSource<'static>>())?,
            )
        })?;
    let logical_bytes = logical_hosts.iter().try_fold(
        checked_mul(logical_hosts.capacity(), size_of::<LogicalHost>())?,
        |sum, host| {
            checked_add(
                sum,
                checked_add(
                    checked_mul(host.precedents.capacity(), size_of::<Precedent>())?,
                    checked_mul(host.ranges.capacity(), size_of::<Range>())?,
                )?,
            )
        },
    )?;
    let owner_bytes = owners.iter().try_fold(
        checked_mul(owners.capacity(), size_of::<PreparedOwner<'static>>())?,
        |sum, owner| {
            checked_add(
                sum,
                checked_mul(
                    owner.range_tiles.capacity(),
                    size_of::<RangeTileSource<'static>>(),
                )?,
            )
        },
    )?;
    checked_add(
        checked_add(
            checked_add(owned_hosts_vec_retained(final_hosts)?, logical_bytes)?,
            owner_bytes,
        )?,
        checked_add(
            checked_mul(uid_index.entries.capacity(), size_of::<OwnerUidEntry>())?,
            checked_add(
                checked_mul(borrowed.capacity(), size_of::<FormulaHost<'static>>())?,
                checked_add(
                    checked_add(
                        checked_mul(
                            cell_assignments.capacity(),
                            size_of::<Vec<CellTileSource<'static>>>(),
                        )?,
                        nested_assignments,
                    )?,
                    checked_mul(
                        prepared_owners.capacity(),
                        size_of::<SourceOwner<'static, 'static>>(),
                    )?,
                )?,
            )?,
        )?,
    )
}

fn prepared_graph_retained(
    final_hosts: &Vec<OwnedHost<'_>>,
    logical_hosts: &Vec<LogicalHost>,
    owners: &Vec<PreparedOwner<'_>>,
    cell_assignments: &Vec<Vec<CellTileSource<'_>>>,
    preflight: &Preflight,
) -> Result<usize, Error> {
    let logical_bytes = logical_hosts.iter().try_fold(
        checked_mul(logical_hosts.capacity(), size_of::<LogicalHost>())?,
        |sum, host| {
            checked_add(
                sum,
                checked_add(
                    checked_mul(host.precedents.capacity(), size_of::<Precedent>())?,
                    checked_mul(host.ranges.capacity(), size_of::<Range>())?,
                )?,
            )
        },
    )?;
    let assignment_bytes = cell_assignments.iter().try_fold(
        checked_mul(
            cell_assignments.capacity(),
            size_of::<Vec<CellTileSource<'static>>>(),
        )?,
        |sum, assignments| {
            checked_add(
                sum,
                checked_mul(assignments.capacity(), size_of::<CellTileSource<'static>>())?,
            )
        },
    )?;
    let owner_bytes = owners.iter().try_fold(
        checked_mul(owners.capacity(), size_of::<PreparedOwner<'static>>())?,
        |sum, owner| {
            checked_add(
                sum,
                checked_mul(
                    owner.range_tiles.capacity(),
                    size_of::<RangeTileSource<'static>>(),
                )?,
            )
        },
    )?;
    checked_add(
        owned_hosts_vec_retained(final_hosts)?,
        checked_add(
            logical_bytes,
            checked_add(
                owner_bytes,
                checked_add(
                    assignment_bytes,
                    checked_mul(preflight.object_ids.capacity(), size_of::<u64>())?,
                )?,
            )?,
        )?,
    )
}

fn prepared_graph_elements(
    final_hosts: &[OwnedHost<'_>],
    logical_hosts: &[LogicalHost],
    owners: &[PreparedOwner<'_>],
    cell_assignments: &[Vec<CellTileSource<'_>>],
    preflight: &Preflight,
) -> Result<usize, Error> {
    let nested_facts = logical_hosts.iter().try_fold(0usize, |sum, host| {
        checked_add(sum, checked_add(host.precedents.len(), host.ranges.len())?)
    })?;
    let assignment_elements = cell_assignments
        .iter()
        .try_fold(0usize, |sum, values| checked_add(sum, values.len()))?;
    let owner_elements = owners.iter().try_fold(owners.len(), |sum, owner| {
        checked_add(sum, owner.range_tiles.len())
    })?;
    checked_add(
        checked_add(final_hosts.len(), logical_hosts.len())?,
        checked_add(
            nested_facts,
            checked_add(
                owner_elements,
                checked_add(
                    checked_add(cell_assignments.len(), assignment_elements)?,
                    preflight.object_ids.len(),
                )?,
            )?,
        )?,
    )
}

struct SourceAdmission {
    source_bytes: usize,
    allocations: usize,
    peak_scratch_bytes: usize,
    graph_work_bytes: usize,
}

impl SourceAdmission {
    fn new(graph: SourceGraph<'_, '_>, limits: Limits) -> Result<Self, Error> {
        let owner_messages = checked_add(1, graph.owners.len())?;
        ensure_limit("owner source messages", owner_messages, limits.max_messages)?;
        let mut messages = owner_messages;
        for owner in graph.owners {
            messages = checked_add(
                messages,
                checked_add(
                    owner
                        .cell_tiles
                        .iter()
                        .filter(|tile| tile.source_present)
                        .count(),
                    owner.range_tiles.len(),
                )?,
            )?;
            ensure_limit("aggregate source messages", messages, limits.max_messages)?;
        }
        let archive_references =
            graph
                .owners
                .iter()
                .try_fold(graph.engine.object_references.len(), |sum, owner| {
                    let owner_and_cell = owner
                        .cell_tiles
                        .iter()
                        .filter(|tile| tile.source_present)
                        .try_fold(
                            checked_add(sum, owner.message.object_references.len())?,
                            |total, tile| checked_add(total, tile.message.object_references.len()),
                        )?;
                    owner
                        .range_tiles
                        .iter()
                        .try_fold(owner_and_cell, |total, tile| {
                            checked_add(total, tile.message.object_references.len())
                        })
                })?;
        ensure_limit(
            "aggregate ArchiveInfo references",
            archive_references,
            limits.max_references,
        )?;
        // The full engine ArchiveInfo list may include source-owned sidecars.
        // Before inspecting it, admit a sorted owner index and one binary
        // lookup per aggregate reference, plus the linear tile checks.
        let owner_lookup_depth = comparison_depth(graph.owners.len())?;
        let locality_work = checked_add(
            checked_add(
                checked_mul(sort_work_upper(graph.owners.len())?, 2)?,
                checked_mul(graph.engine.object_references.len(), owner_lookup_depth)?,
            )?,
            archive_references,
        )?;
        ensure_limit(
            "ArchiveInfo locality work",
            locality_work,
            limits.max_work_bytes,
        )?;
        validate_archive_info_references(graph)?;
        let mut bytes = graph.engine.payload.len();
        for owner in graph.owners {
            bytes = checked_add(bytes, owner.message.payload.len())?;
            for tile in owner.cell_tiles.iter().filter(|tile| tile.source_present) {
                bytes = checked_add(bytes, tile.message.payload.len())?;
            }
            for tile in owner.range_tiles {
                bytes = checked_add(bytes, tile.message.payload.len())?;
            }
        }
        ensure_limit("aggregate source bytes", bytes, limits.max_source_bytes)?;
        ensure_limit("changed hosts", graph.changes.len(), limits.max_hosts)?;
        ensure_limit(
            "authoritative formula hosts",
            graph.existing_formula_hosts.len(),
            limits.max_hosts,
        )?;
        if graph
            .existing_formula_hosts
            .iter()
            .any(|host| host.owner != graph.table_owner)
            || graph.changes.iter().any(|change| {
                change_key(*change).map_or(true, |key| key.owner != graph.table_owner)
            })
            || !owner_exists(graph.owners, graph.table_owner)
        {
            return Err(Error::InvalidGraph);
        }
        let changed_fact_bytes = graph.changes.iter().try_fold(0usize, |sum, change| {
            let Some(host) = change.new else {
                return Ok(sum);
            };
            checked_add(
                sum,
                checked_add(
                    checked_mul(host.precedents.len(), 40)?,
                    checked_mul(host.ranges.len(), 96)?,
                )?,
            )
        })?;
        let candidate_upper = checked_add(bytes, changed_fact_bytes)?;
        ensure_limit(
            "source candidate upper",
            candidate_upper,
            limits.max_output_bytes,
        )?;
        let decode_work = checked_add(
            checked_mul(bytes, STRICT_WORK_MULTIPLIER)?,
            checked_mul(candidate_upper, STRICT_WORK_MULTIPLIER)?,
        )?;
        let search_width = checked_add(graph.existing_formula_hosts.len(), graph.owners.len())?;
        let search_factor =
            usize::try_from(usize::BITS.saturating_sub(search_width.leading_zeros()))
                .map_err(|_error| Error::InvalidGraph)?
                .saturating_add(8);
        // Host/fact ceilings include the source and final views.  They share
        // the authored changes, so charge the larger host population plus the
        // actual changed facts once rather than adding both declared host
        // ceilings and the same overlay again.
        let merge_items = checked_add(
            checked_add(
                checked_add(messages, graph.existing_formula_hosts.len())?,
                limits.max_hosts,
            )?,
            checked_add(limits.max_precedents, limits.max_ranges)?,
        )?;
        let merge_work = checked_mul(merge_items, search_factor)?;
        // Source canonicalization sorts the UID index, decoded hosts, each
        // host's precedent/range facts, and the final host overlay. Charging
        // the declared aggregate ceilings is conservative and precedes decode.
        let sort_work = checked_add(
            sort_work_upper(graph.owners.len())?,
            checked_add(
                sort_work_upper(checked_add(limits.max_hosts, limits.max_ranges)?)?,
                checked_add(
                    sort_work_upper(limits.max_precedents)?,
                    sort_work_upper(limits.max_ranges)?,
                )?,
            )?,
        )?;
        let graph_work = checked_add(checked_add(merge_work, sort_work)?, locality_work)?;
        ensure_limit(
            "source decode/merge work",
            checked_add(decode_work, graph_work)?,
            limits.max_work_bytes,
        )?;
        let field_upper = checked_add(bytes / 2 + messages, candidate_upper / 2 + messages)?;
        ensure_limit(
            "aggregate field preauthorization",
            field_upper,
            limits.max_fields,
        )?;
        let selected_references =
            graph
                .owners
                .iter()
                .try_fold(graph.owners.len(), |sum, owner| {
                    checked_add(
                        sum,
                        checked_add(owner.cell_tiles.len(), owner.range_tiles.len())?,
                    )
                })?;
        let reference_upper =
            checked_add(archive_references, checked_mul(selected_references, 3)?)?;
        ensure_limit(
            "aggregate reference preauthorization",
            reference_upper,
            limits.max_references,
        )?;
        let scratch = checked_add(
            checked_mul(
                checked_mul(checked_add(limits.max_hosts, limits.max_ranges)?, 2)?,
                size_of::<OwnedHost<'static>>(),
            )?,
            checked_add(
                checked_mul(limits.max_precedents, size_of::<Precedent>() * 6)?,
                checked_add(
                    checked_mul(
                        limits.max_ranges,
                        checked_add(
                            checked_add(size_of::<Range>(), size_of::<Option<&'static [u8]>>())?,
                            checked_add(
                                size_of::<(Range, Option<&'static [u8]>)>(),
                                size_of::<(HostKey, Range, Option<&'static [u8]>)>(),
                            )?,
                        )?,
                    )?,
                    checked_add(
                        checked_mul(limits.max_messages, size_of::<u64>())?,
                        checked_add(
                            checked_mul(graph.owners.len(), size_of::<OwnerUidEntry>())?,
                            checked_mul(graph.owners.len(), size_of::<ArchiveOwnerReference>())?,
                        )?,
                    )?,
                )?,
            )?,
        )?;
        ensure_limit(
            "source scratch preauthorization",
            scratch,
            limits.max_scratch_bytes,
        )?;
        let allocations = checked_add(
            17,
            checked_add(
                checked_mul(messages, 8)?,
                checked_mul(checked_add(limits.max_hosts, limits.max_ranges)?, 3)?,
            )?,
        )?;
        ensure_limit(
            "source allocation preauthorization",
            allocations,
            limits.max_allocations,
        )?;
        Ok(Self {
            source_bytes: bytes,
            allocations,
            peak_scratch_bytes: scratch,
            graph_work_bytes: graph_work,
        })
    }
}

fn sort_work_upper(items: usize) -> Result<usize, Error> {
    if items < 2 {
        return Ok(0);
    }
    let factor = usize::try_from(usize::BITS.saturating_sub((items - 1).leading_zeros()))
        .map_err(|_error| Error::InvalidGraph)?;
    checked_mul(items, factor)
}

fn comparison_depth(items: usize) -> Result<usize, Error> {
    if items < 2 {
        return Ok(1);
    }
    usize::try_from(usize::BITS - (items - 1).leading_zeros()).map_err(|_| Error::InvalidGraph)
}

fn prepare_cell_assignments<'source>(
    owners: &[SourceOwner<'source, '_>],
    hosts: &[OwnedHost<'_>],
) -> Result<Vec<Vec<CellTileSource<'source>>>, Error> {
    let mut all = try_vec(owners.len())?;
    for owner in owners {
        let owner_start = hosts.partition_point(|host| host.key.owner < owner.internal_owner);
        let owner_end = hosts.partition_point(|host| host.key.owner <= owner.internal_owner);
        if owner.cell_tiles.is_empty() {
            if hosts[owner_start..owner_end]
                .iter()
                .any(|host| host.cell_tile_object_id.is_some())
            {
                return Err(Error::InvalidGraph);
            }
            all.push(Vec::new());
            continue;
        }
        if hosts[owner_start..owner_end]
            .iter()
            .any(|host| host.cell_tile_object_id.is_none())
        {
            return Err(Error::InvalidGraph);
        }
        if hosts[owner_start..owner_end].iter().any(|host| {
            owner
                .cell_tiles
                .iter()
                .filter(|tile| Some(tile.message.object_id) == host.cell_tile_object_id)
                .count()
                != 1
        }) {
            return Err(Error::InvalidGraph);
        }
        let mut assignments = try_vec(owner.cell_tiles.len())?;
        assignments.extend_from_slice(owner.cell_tiles);
        all.push(assignments);
    }
    Ok(all)
}

fn prepare_owner_envelopes<'source>(
    owners: &[SourceOwner<'source, '_>],
) -> Result<Vec<PreparedOwner<'source>>, Error> {
    let mut prepared = try_vec(owners.len())?;
    for owner in owners {
        let mut range_tiles = try_vec(owner.range_tiles.len())?;
        range_tiles.extend_from_slice(owner.range_tiles);
        prepared.push(PreparedOwner {
            message: owner.message,
            internal_owner: owner.internal_owner,
            uid_lower: owner.uid_lower,
            uid_upper: owner.uid_upper,
            range_tiles,
        });
    }
    Ok(prepared)
}

fn execute_complete_graph(
    graph: CompleteGraph<'_, '_>,
    limits: Limits,
    preflight: Preflight,
) -> Result<Artifact, Error> {
    let decode_options = dependency::DecodeOptions::new(
        limits.max_source_bytes.max(limits.max_output_bytes),
        limits.max_fields,
        limits.max_work_bytes,
        limits.recursion_limit,
        limits.max_references,
        1,
    );
    let wire_limits = wire_limits(limits)?;
    let mut report = Report {
        source_bytes: preflight.source_bytes,
        hosts: graph.hosts.len(),
        precedents: preflight.precedents,
        ranges: preflight.ranges,
        graph_work_bytes: preflight.graph_work_bytes,
        peak_scratch_bytes: preflight.peak_scratch_bytes,
        ..Report::default()
    };

    let mut engine_refs = ReferenceCollector::new(graph.owners.len())?;
    let (engine_snapshot, source_engine_report) =
        dependency::decode_calculation_engine_with_visitor(
            graph.engine.payload,
            decode_options,
            &mut engine_refs,
        )
        .map_err(|_error| Error::StrictDependency)?;
    add_decode_report(&mut report, source_engine_report)?;
    validate_engine_refs(graph, &engine_refs.formula_owners)?;
    let count = graph.formula_count;
    let (current_tracker, current_tracker_report) =
        dependency::decode_dependency_tracker_with_report(
            engine_snapshot.dependency_tracker(),
            decode_options,
        )
        .map_err(|_error| Error::StrictDependency)?;
    add_decode_report(&mut report, current_tracker_report)?;
    let engine_candidate = patch_nested_fields_batched_with_limits(
        graph.engine.payload,
        &[NestedFieldEdit::new(
            &ENGINE_COUNT_PATH,
            current_tracker.number_of_formulas().is_some(),
            NestedFieldReplacement::Varint(Some(count)),
        )],
        wire_limits,
    )
    .map_err(|_error| Error::Wire)?;
    let mut candidate_engine_refs = ReferenceCollector::new(graph.owners.len())?;
    let (candidate_engine, candidate_engine_report) =
        dependency::decode_calculation_engine_with_visitor(
            &engine_candidate,
            decode_options,
            &mut candidate_engine_refs,
        )
        .map_err(|_error| Error::StrictDependency)?;
    add_decode_report(&mut report, candidate_engine_report)?;
    let (candidate_tracker, candidate_tracker_report) =
        dependency::decode_dependency_tracker_with_report(
            candidate_engine.dependency_tracker(),
            decode_options,
        )
        .map_err(|_error| Error::StrictDependency)?;
    add_decode_report(&mut report, candidate_tracker_report)?;
    if candidate_tracker.number_of_formulas() != Some(count)
        || candidate_engine_refs.formula_owners != engine_refs.formula_owners
    {
        return Err(Error::InvalidGraph);
    }

    let mut owners = try_vec(graph.owners.len())?;
    let mut cell_tiles = try_vec(preflight.cell_tiles)?;
    let mut range_tiles = try_vec(preflight.range_tiles)?;
    let mut host_cursor = 0usize;
    for owner in graph.owners {
        let host_start = host_cursor;
        while host_cursor < graph.hosts.len()
            && graph.hosts[host_cursor].owner == owner.internal_owner
        {
            host_cursor += 1;
        }
        if owner.internal_owner != graph.table_owner {
            owners.push(unchanged_edit(owner.message)?);
            for tile in owner.cell_tiles {
                cell_tiles.push(unchanged_edit(tile.message)?);
            }
            for tile in owner.range_tiles {
                range_tiles.push(unchanged_edit(tile.message)?);
            }
            continue;
        }
        rewrite_owner(
            graph,
            owner,
            host_start,
            host_cursor,
            limits,
            decode_options,
            wire_limits,
            &mut report,
            &mut owners,
            &mut cell_tiles,
            &mut range_tiles,
        )?;
    }
    if host_cursor != graph.hosts.len() {
        return Err(Error::InvalidGraph);
    }

    let engine = finish_edit(graph.engine, engine_candidate, &mut report)?;
    ensure_limit("output bytes", report.output_bytes, limits.max_output_bytes)?;
    ensure_limit("strict fields", report.fields, limits.max_fields)?;
    ensure_limit(
        "strict work",
        report.strict_work_bytes,
        limits.max_work_bytes,
    )?;
    ensure_limit("references", report.references, limits.max_references)?;
    report.allocations = preflight.allocations;
    report.objects = preflight.messages;
    report.retained_elements = checked_add(preflight.messages, report.references)?;
    report.retained_bytes = owners
        .iter()
        .chain(&cell_tiles)
        .chain(&range_tiles)
        .chain(core::iter::once(&engine))
        .try_fold(0usize, |sum, edit| {
            let payload = edit.payload.as_ref().map_or(0, Vec::capacity);
            let references = checked_mul(edit.object_references.capacity(), size_of::<u64>())?;
            checked_add(sum, checked_add(payload, references)?)
        })?;
    report.retained_bytes = checked_add(
        report.retained_bytes,
        checked_mul(
            checked_add(
                checked_add(owners.capacity(), cell_tiles.capacity())?,
                range_tiles.capacity(),
            )?,
            size_of::<MessageEdit>(),
        )?,
    )?;
    ensure_limit(
        "retained bytes",
        report.retained_bytes,
        limits.max_retained_bytes,
    )?;
    Ok(Artifact {
        engine,
        owners,
        cell_tiles,
        range_tiles,
        report,
    })
}

fn unchanged_edit(source: SourceMessage<'_>) -> Result<MessageEdit, Error> {
    let mut references = try_vec(source.object_references.len())?;
    references.extend_from_slice(source.object_references);
    Ok(MessageEdit {
        object_id: source.object_id,
        payload: None,
        object_references: references,
    })
}

#[derive(Debug)]
struct OwnedHost<'a> {
    key: HostKey,
    precedents: Vec<Precedent>,
    ranges: Vec<Range>,
    cell_tile_object_id: Option<u64>,
    has_cell_record: bool,
    is_in_cycle: bool,
    raw_cell_record: Option<&'a [u8]>,
    raw_range_records: Vec<Option<&'a [u8]>>,
}

fn logical_hosts(
    hosts: &[OwnedHost<'_>],
    existing_formula_hosts: &[HostKey],
    changes: &[HostChange<'_>],
    limits: Limits,
) -> Result<Vec<LogicalHost>, Error> {
    let mut logical = try_vec(hosts.len())?;
    let mut change_cursor = 0usize;
    for host in hosts {
        while change_cursor < changes.len() && change_key(changes[change_cursor])? < host.key {
            change_cursor += 1;
        }
        let source_present = existing_formula_hosts.binary_search(&host.key).is_ok();
        let final_present =
            if change_cursor < changes.len() && change_key(changes[change_cursor])? == host.key {
                changes[change_cursor].new.is_some()
            } else {
                source_present
            };
        if !final_present {
            continue;
        }
        if host.precedents.is_empty() && host.ranges.is_empty() && !host.is_in_cycle {
            continue;
        }
        let mut precedents = try_vec(host.precedents.len())?;
        precedents.extend_from_slice(&host.precedents);
        let mut ranges = try_vec(host.ranges.len())?;
        ranges.extend_from_slice(&host.ranges);
        logical.push(LogicalHost {
            owner: host.key.owner,
            row: host.key.row,
            column: host.key.column,
            precedents,
            ranges,
            is_in_cycle: host.is_in_cycle,
        });
    }
    ensure_limit("logical dependency hosts", logical.len(), limits.max_hosts)?;
    Ok(logical)
}

fn change_key(change: HostChange<'_>) -> Result<HostKey, Error> {
    change
        .old
        .or_else(|| {
            change.new.map(|host| HostKey {
                owner: host.owner,
                row: host.row,
                column: host.column,
            })
        })
        .ok_or(Error::InvalidGraph)
}

struct DecodedSource<'a> {
    hosts: Vec<OwnedHost<'a>>,
    formula_count: u64,
    report: Report,
}

fn decode_source_hosts<'source>(
    graph: SourceGraph<'source, '_>,
    uid_index: &OwnerUidIndex,
    limits: Limits,
) -> Result<DecodedSource<'source>, Error> {
    let mut report = Report::default();
    let (engine, engine_report) = dependency::decode_calculation_engine_with_report(
        graph.engine.payload,
        remaining_decode_options(limits, report)?,
    )
    .map_err(|_error| Error::StrictDependency)?;
    add_decode_report(&mut report, engine_report)?;
    let (tracker, tracker_report) = dependency::decode_dependency_tracker_with_report(
        engine.dependency_tracker(),
        remaining_decode_options(limits, report)?,
    )
    .map_err(|_error| Error::StrictDependency)?;
    add_decode_report(&mut report, tracker_report)?;
    let formula_count = tracker.number_of_formulas().ok_or(Error::InvalidGraph)?;
    if graph
        .existing_formula_hosts
        .windows(2)
        .any(|pair| pair[0] >= pair[1])
    {
        return Err(Error::InvalidGraph);
    }
    let observed_host_bound = observed_source_host_bound(graph, limits)?;
    let mut hosts = try_vec(observed_host_bound)?;
    let mut aggregate_precedents = 0usize;
    let mut aggregate_ranges = 0usize;
    for owner in graph.owners {
        let mut inline = SourceFacts::new(
            owner.internal_owner,
            None,
            uid_index,
            owner.message.payload.len(),
            fact_limits(limits, aggregate_precedents, aggregate_ranges)?,
        )?;
        let (snapshot, decoded) = dependency::decode_formula_owner_dependencies_with_visitor(
            owner.message.payload,
            remaining_decode_options(limits, report)?,
            &mut inline,
        )
        .map_err(|_error| Error::StrictDependency)?;
        add_decode_report(&mut report, decoded)?;
        if snapshot.internal_formula_owner_id() != owner.internal_owner
            || snapshot.formula_owner_uid().lower() != owner.uid_lower
            || snapshot.formula_owner_uid().upper() != owner.uid_upper
        {
            return Err(Error::InvalidGraph);
        }
        attach_cell_raw(&mut inline.hosts, snapshot.cell_dependencies(), 1, limits)?;
        attach_inline_range_raw(&mut inline.hosts, snapshot.range_dependencies(), limits)?;
        add_source_fact_counts(
            &mut aggregate_precedents,
            &mut aggregate_ranges,
            &inline.hosts,
            limits,
        )?;
        append_owned_hosts(&mut hosts, inline, limits)?;
        for tile in owner.cell_tiles {
            if !tile.source_present {
                if !tile.message.payload.is_empty()
                    || !tile.message.object_references.is_empty()
                    || tile.message.object_id == 0
                {
                    return Err(Error::InvalidGraph);
                }
                continue;
            }
            let mut facts = SourceFacts::new(
                owner.internal_owner,
                Some(tile.message.object_id),
                uid_index,
                tile.message.payload.len(),
                fact_limits(limits, aggregate_precedents, aggregate_ranges)?,
            )?;
            let (tile_snapshot, decoded) = dependency::decode_cell_record_tile_with_visitor(
                tile.message.payload,
                remaining_decode_options(limits, report)?,
                &mut facts,
            )
            .map_err(|_error| Error::StrictDependency)?;
            add_decode_report(&mut report, decoded)?;
            if tile_snapshot.internal_owner_id() != owner.internal_owner
                || tile_snapshot.tile_column_begin() != tile.tile_column_begin
                || tile_snapshot.tile_row_begin() != tile.tile_row_begin
            {
                return Err(Error::InvalidGraph);
            }
            attach_cell_raw(&mut facts.hosts, Some(tile.message.payload), 4, limits)?;
            add_source_fact_counts(
                &mut aggregate_precedents,
                &mut aggregate_ranges,
                &facts.hosts,
                limits,
            )?;
            append_owned_hosts(&mut hosts, facts, limits)?;
        }
        for tile in owner.range_tiles {
            let mut facts = SourceRangeFacts::new(
                owner.internal_owner,
                tile.target_owner,
                tile.message.payload.len(),
                fact_limits(limits, aggregate_precedents, aggregate_ranges)?,
            )?;
            let (snapshot, decoded) = dependency::decode_range_precedents_tile_with_visitor(
                tile.message.payload,
                remaining_decode_options(limits, report)?,
                &mut facts,
            )
            .map_err(|_error| Error::StrictDependency)?;
            add_decode_report(&mut report, decoded)?;
            if snapshot.to_owner_id() != tile.target_owner {
                return Err(Error::InvalidGraph);
            }
            if facts.invalid {
                return Err(Error::InvalidGraph);
            }
            attach_tiled_range_raw(&mut facts.ranges, tile.message.payload, limits)?;
            aggregate_ranges = checked_add(aggregate_ranges, facts.ranges.len())?;
            ensure_limit(
                "aggregate source ranges",
                aggregate_ranges,
                limits.max_ranges,
            )?;
            merge_source_ranges(&mut hosts, facts.ranges, limits)?;
        }
    }
    hosts.sort_unstable_by_key(|host| host.key);
    coalesce_hosts(&mut hosts, graph.table_owner, limits)?;
    // Coalescing changes only record packing, never the aggregate fact count.
    ensure_limit(
        "aggregate source precedents",
        aggregate_precedents,
        limits.max_precedents,
    )?;
    ensure_limit(
        "aggregate source ranges",
        aggregate_ranges,
        limits.max_ranges,
    )?;
    let union_lower_bound = checked_add(graph.existing_formula_hosts.len(), hosts.len())?
        .checked_sub(
            hosts
                .iter()
                .filter(|host| {
                    graph
                        .existing_formula_hosts
                        .binary_search(&host.key)
                        .is_ok()
                })
                .count(),
        )
        .ok_or(Error::InvalidGraph)?;
    if usize::try_from(formula_count).map_err(|_| Error::InvalidGraph)? < union_lower_bound {
        return Err(Error::InvalidGraph);
    }
    if hosts.iter().any(|host| {
        host.key.owner == graph.table_owner
            && graph
                .existing_formula_hosts
                .binary_search(&host.key)
                .is_err()
    }) {
        return Err(Error::InvalidGraph);
    }
    // Other owners may contain non-table dependency declarations mirrored
    // between inline and tiled storage. They participate in the global engine
    // count lower bound above, but are opaque to this single-table rewrite and
    // are emitted byte-for-byte unchanged by `execute_complete_graph`.
    hosts.retain(|host| host.key.owner == graph.table_owner);
    ensure_limit(
        "aggregate decoded source fields",
        report.fields,
        limits.max_fields,
    )?;
    ensure_limit(
        "aggregate decoded source work",
        report.strict_work_bytes,
        limits.max_work_bytes,
    )?;
    ensure_limit(
        "aggregate decoded source references",
        report.references,
        limits.max_references,
    )?;
    Ok(DecodedSource {
        hosts,
        formula_count,
        report,
    })
}

fn observed_source_host_bound(graph: SourceGraph<'_, '_>, limits: Limits) -> Result<usize, Error> {
    Ok(graph
        .owners
        .iter()
        .try_fold(0usize, |sum, owner| {
            let mut bound = owner.message.payload.len() / 2 + 1;
            for tile in owner.cell_tiles {
                bound = checked_add(bound, tile.message.payload.len() / 2 + 1)?;
            }
            // A range-only source still materializes one temporary OwnedHost
            // per decoded range record before coalescing.
            for tile in owner.range_tiles {
                bound = checked_add(bound, tile.message.payload.len() / 2 + 1)?;
            }
            checked_add(sum, bound)
        })?
        .min(checked_add(limits.max_hosts, limits.max_ranges)?))
}

fn overlay_changes<'a>(
    hosts: Vec<OwnedHost<'a>>,
    changes: &[HostChange<'_>],
    owners: &[SourceOwner<'_, '_>],
    existing_formula_hosts: &[HostKey],
    limits: Limits,
) -> Result<Vec<OwnedHost<'a>>, Error> {
    let capacity = checked_add(hosts.len(), changes.len())?;
    ensure_limit("final dependency host capacity", capacity, limits.max_hosts)?;
    let mut output = try_vec(capacity)?;
    let mut source = hosts.into_iter().peekable();
    let mut previous = None;
    for change in changes {
        let key = change
            .old
            .or_else(|| {
                change.new.map(|host| HostKey {
                    owner: host.owner,
                    row: host.row,
                    column: host.column,
                })
            })
            .ok_or(Error::InvalidGraph)?;
        if previous.is_some_and(|prior| prior >= key) {
            return Err(Error::InvalidGraph);
        }
        previous = Some(key);
        if let (Some(old), Some(new)) = (change.old, change.new) {
            let new_key = HostKey {
                owner: new.owner,
                row: new.row,
                column: new.column,
            };
            if old != new_key {
                return Err(Error::InvalidGraph);
            }
        }
        if change.old.is_some() != existing_formula_hosts.binary_search(&key).is_ok() {
            return Err(Error::InvalidGraph);
        }
        if !owner_exists(owners, key.owner) {
            return Err(Error::InvalidGraph);
        }
        while source.peek().is_some_and(|host| host.key < key) {
            output.push(source.next().ok_or(Error::InvalidGraph)?);
        }
        let decoded_old = source.peek().is_some_and(|host| host.key == key);
        if change.old.is_none() && decoded_old {
            return Err(Error::InvalidGraph);
        }
        if decoded_old {
            let _ = source.next();
        }
        if let Some(new) = change.new {
            let mut precedents = try_vec(new.precedents.len())?;
            precedents.extend_from_slice(new.precedents);
            let mut ranges = try_vec(new.ranges.len())?;
            ranges.extend_from_slice(new.ranges);
            let mut raw_range_records = try_vec(new.ranges.len())?;
            raw_range_records.resize(new.ranges.len(), None);
            if !precedents.is_empty() || !ranges.is_empty() {
                output.push(OwnedHost {
                    key,
                    precedents,
                    ranges,
                    cell_tile_object_id: change.cell_tile_object_id,
                    has_cell_record: true,
                    is_in_cycle: false,
                    raw_cell_record: None,
                    raw_range_records,
                });
            }
        }
    }
    output.extend(source);
    ensure_limit("final dependency hosts", output.len(), limits.max_hosts)?;
    Ok(output)
}

struct SourceFacts<'source, 'index> {
    owner: u32,
    cell_tile_object_id: Option<u64>,
    arrays: [Vec<u32>; 5],
    hosts: Vec<OwnedHost<'source>>,
    limits: Limits,
    invalid: bool,
    limit_resource: Option<&'static str>,
    uid_index: &'index OwnerUidIndex,
}

impl<'source, 'index> SourceFacts<'source, 'index> {
    fn new(
        owner: u32,
        cell_tile_object_id: Option<u64>,
        uid_index: &'index OwnerUidIndex,
        source_bytes: usize,
        limits: Limits,
    ) -> Result<Self, Error> {
        let precedent_capacity = source_bytes.min(limits.max_precedents);
        let host_capacity = (source_bytes / 2 + 1).min(limits.max_hosts);
        let mut arrays = core::array::from_fn(|_| Vec::new());
        for values in &mut arrays {
            reserve_exact_capacity(values, precedent_capacity)?;
        }
        Ok(Self {
            owner,
            cell_tile_object_id,
            arrays,
            hosts: try_vec(host_capacity)?,
            limits,
            invalid: false,
            limit_resource: None,
            uid_index,
        })
    }
}

impl dependency::DependencyVisitor for SourceFacts<'_, '_> {
    fn visit_expanded_edge_component(
        &mut self,
        component: dependency::ExpandedEdgeComponent,
    ) -> Result<(), dependency::DecodeError> {
        let slot = match component.kind() {
            dependency::ExpandedEdgeKind::LocalRow => 0,
            dependency::ExpandedEdgeKind::LocalColumn => 1,
            dependency::ExpandedEdgeKind::ExternalRow => 2,
            dependency::ExpandedEdgeKind::ExternalColumn => 3,
            dependency::ExpandedEdgeKind::InternalOwner => 4,
        };
        if self.arrays[slot].len() >= self.limits.max_precedents {
            self.limit_resource = Some("source precedents");
            return Ok(());
        }
        self.arrays[slot].push(component.value());
        Ok(())
    }

    fn visit_cell_record(
        &mut self,
        record: dependency::CellRecordSnapshot<'_>,
    ) -> Result<(), dependency::DecodeError> {
        if self.arrays[0].len() != self.arrays[1].len()
            || self.arrays[2].len() != self.arrays[3].len()
            || self.arrays[2].len() != self.arrays[4].len()
        {
            self.invalid = true;
            return Ok(());
        }
        let count = self.arrays[0].len() + self.arrays[2].len();
        let Ok(mut precedents) = try_vec(count) else {
            self.invalid = true;
            return Ok(());
        };
        for index in 0..self.arrays[0].len() {
            precedents.push(Precedent {
                target_owner: self.owner,
                row: self.arrays[0][index],
                column: self.arrays[1][index],
            });
        }
        for index in 0..self.arrays[2].len() {
            precedents.push(Precedent {
                target_owner: self.arrays[4][index],
                row: self.arrays[2][index],
                column: self.arrays[3][index],
            });
        }
        precedents.sort_unstable();
        if precedents.windows(2).any(|pair| pair[0] == pair[1]) {
            self.invalid = true;
            return Ok(());
        }
        if self.hosts.len() >= self.limits.max_hosts {
            self.limit_resource = Some("source hosts");
            return Ok(());
        }
        self.hosts.push(OwnedHost {
            key: HostKey {
                owner: self.owner,
                row: record.row(),
                column: record.column(),
            },
            precedents,
            ranges: Vec::new(),
            cell_tile_object_id: self.cell_tile_object_id,
            has_cell_record: true,
            is_in_cycle: record.is_in_a_cycle().unwrap_or(false),
            raw_cell_record: None,
            raw_range_records: Vec::new(),
        });
        for values in &mut self.arrays {
            values.clear();
        }
        Ok(())
    }

    fn visit_range_back_dependency(
        &mut self,
        record: dependency::RangeBackDependencySnapshot<'_>,
    ) -> Result<(), dependency::DecodeError> {
        let (target_owner, range) =
            if let Some(reference) = record.decoded_internal_range_reference() {
                (reference.owner_id(), reference.range())
            } else if let Some(reference) = record.decoded_range_reference() {
                let Some((lower, upper)) = cfuuid_halves(reference.table_id()) else {
                    self.invalid = true;
                    return Ok(());
                };
                let Some(internal_owner) = self.uid_index.internal_owner(lower, upper) else {
                    self.invalid = true;
                    return Ok(());
                };
                (internal_owner, reference.range())
            } else {
                self.invalid = true;
                return Ok(());
            };
        let Ok(mut ranges) = try_vec(1) else {
            self.invalid = true;
            return Ok(());
        };
        ranges.push(Range {
            target_owner,
            top: range.top_left_row(),
            left: range.top_left_column(),
            bottom: range.bottom_right_row(),
            right: range.bottom_right_column(),
        });
        if self.hosts.len() >= self.limits.max_hosts {
            self.limit_resource = Some("source hosts");
            return Ok(());
        }
        self.hosts.push(OwnedHost {
            key: HostKey {
                owner: self.owner,
                row: record.cell_coord_row(),
                column: record.cell_coord_column(),
            },
            precedents: Vec::new(),
            ranges,
            cell_tile_object_id: self.cell_tile_object_id,
            has_cell_record: false,
            is_in_cycle: false,
            raw_cell_record: None,
            raw_range_records: Vec::new(),
        });
        Ok(())
    }
}

struct SourceRangeFacts<'a> {
    host_owner: u32,
    target_owner: u32,
    ranges: Vec<(HostKey, Range, Option<&'a [u8]>)>,
    invalid: bool,
    maximum: usize,
}

impl<'a> SourceRangeFacts<'a> {
    fn new(
        host_owner: u32,
        target_owner: u32,
        source_bytes: usize,
        limits: Limits,
    ) -> Result<Self, Error> {
        Ok(Self {
            host_owner,
            target_owner,
            ranges: try_vec((source_bytes / 2 + 1).min(limits.max_ranges))?,
            invalid: false,
            maximum: limits.max_ranges,
        })
    }
}

impl dependency::DependencyVisitor for SourceRangeFacts<'_> {
    fn visit_from_to_range(
        &mut self,
        record: dependency::FromToRangeSnapshot<'_>,
    ) -> Result<(), dependency::DecodeError> {
        let from = record.decoded_from_coord();
        let rect = record.decoded_refers_to_rect();
        let (Some(column), Some(row), Some(left), Some(top)) = (
            from.column(),
            from.row(),
            rect.origin().column(),
            rect.origin().row(),
        ) else {
            self.invalid = true;
            return Ok(());
        };
        let width = rect.size().num_columns().unwrap_or(1);
        let height = rect.size().num_rows().unwrap_or(1);
        let (Some(right), Some(bottom)) = (
            width
                .checked_sub(1)
                .and_then(|delta| left.checked_add(delta)),
            height
                .checked_sub(1)
                .and_then(|delta| top.checked_add(delta)),
        ) else {
            self.invalid = true;
            return Ok(());
        };
        if self.ranges.len() >= self.maximum {
            self.invalid = true;
            return Ok(());
        }
        self.ranges.push((
            HostKey {
                owner: self.host_owner,
                row,
                column,
            },
            Range {
                target_owner: self.target_owner,
                top,
                left,
                bottom,
                right,
            },
            None,
        ));
        Ok(())
    }
}

fn borrow_hosts<'a>(
    hosts: &'a [OwnedHost<'a>],
    limits: Limits,
) -> Result<Vec<FormulaHost<'a>>, Error> {
    let mut borrowed = try_vec(hosts.len())?;
    for host in hosts {
        borrowed.push(FormulaHost {
            owner: host.key.owner,
            row: host.key.row,
            column: host.key.column,
            precedents: &host.precedents,
            ranges: &host.ranges,
            raw_cell_record: host.raw_cell_record,
            raw_range_records: &host.raw_range_records,
            cell_tile_object_id: host.cell_tile_object_id,
        });
    }
    ensure_limit("borrowed final hosts", borrowed.len(), limits.max_hosts)?;
    Ok(borrowed)
}

fn append_owned_hosts<'source>(
    destination: &mut Vec<OwnedHost<'source>>,
    source: SourceFacts<'source, '_>,
    limits: Limits,
) -> Result<(), Error> {
    if let Some(resource) = source.limit_resource {
        let maximum = if resource == "source hosts" {
            limits.max_hosts
        } else {
            limits.max_precedents
        };
        return Err(Error::Limit {
            resource,
            observed: maximum.checked_add(1).ok_or(Error::InvalidGraph)?,
            maximum,
        });
    }
    if source.invalid || !source.arrays.iter().all(Vec::is_empty) {
        return Err(Error::InvalidGraph);
    }
    ensure_limit(
        "decoded source hosts",
        checked_add(destination.len(), source.hosts.len())?,
        checked_add(limits.max_hosts, limits.max_ranges)?,
    )?;
    if destination.capacity().saturating_sub(destination.len()) < source.hosts.len() {
        return Err(Error::InvalidGraph);
    }
    destination.extend(source.hosts);
    Ok(())
}

fn fact_limits(limits: Limits, precedents: usize, ranges: usize) -> Result<Limits, Error> {
    Ok(Limits {
        max_precedents: limits
            .max_precedents
            .checked_sub(precedents)
            .ok_or(Error::InvalidGraph)?,
        max_ranges: limits
            .max_ranges
            .checked_sub(ranges)
            .ok_or(Error::InvalidGraph)?,
        ..limits
    })
}

fn add_source_fact_counts(
    precedents: &mut usize,
    ranges: &mut usize,
    hosts: &[OwnedHost<'_>],
    limits: Limits,
) -> Result<(), Error> {
    for host in hosts {
        *precedents = checked_add(*precedents, host.precedents.len())?;
        *ranges = checked_add(*ranges, host.ranges.len())?;
    }
    ensure_limit(
        "aggregate source precedents",
        *precedents,
        limits.max_precedents,
    )?;
    ensure_limit("aggregate source ranges", *ranges, limits.max_ranges)
}

fn merge_source_ranges<'a>(
    hosts: &mut Vec<OwnedHost<'a>>,
    ranges: Vec<(HostKey, Range, Option<&'a [u8]>)>,
    limits: Limits,
) -> Result<(), Error> {
    ensure_limit(
        "decoded source range hosts",
        checked_add(hosts.len(), ranges.len())?,
        checked_add(limits.max_hosts, limits.max_ranges)?,
    )?;
    if hosts.capacity().saturating_sub(hosts.len()) < ranges.len() {
        return Err(Error::InvalidGraph);
    }
    for (key, range, raw) in ranges {
        let mut values = try_vec(1)?;
        values.push(range);
        hosts.push(OwnedHost {
            key,
            precedents: Vec::new(),
            ranges: values,
            cell_tile_object_id: None,
            has_cell_record: false,
            is_in_cycle: false,
            raw_cell_record: None,
            raw_range_records: {
                let mut records = try_vec(1)?;
                records.push(raw);
                records
            },
        });
    }
    Ok(())
}

fn attach_cell_raw<'a>(
    hosts: &mut [OwnedHost<'a>],
    container: Option<&'a [u8]>,
    field_number: u32,
    limits: Limits,
) -> Result<(), Error> {
    let Some(container) = container else {
        if hosts.iter().any(|host| host.has_cell_record) {
            return Err(Error::InvalidGraph);
        }
        return Ok(());
    };
    let view = WireView::parse_with_limits(container, wire_limits(limits)?)
        .map_err(|_error| Error::Wire)?;
    let mut records = hosts.iter_mut().filter(|host| host.has_cell_record);
    for field in view.fields().filter(|field| field.number() == field_number) {
        field
            .validate_canonical_framing()
            .map_err(|_error| Error::Wire)?;
        let host = records.next().ok_or(Error::InvalidGraph)?;
        host.raw_cell_record = Some(field.payload());
    }
    if records.next().is_some() {
        return Err(Error::InvalidGraph);
    }
    Ok(())
}

fn attach_inline_range_raw<'a>(
    hosts: &mut [OwnedHost<'a>],
    container: Option<&'a [u8]>,
    limits: Limits,
) -> Result<(), Error> {
    let expected = hosts.iter().map(|host| host.ranges.len()).sum::<usize>();
    let Some(container) = container else {
        return if expected == 0 {
            Ok(())
        } else {
            Err(Error::InvalidGraph)
        };
    };
    let view = WireView::parse_with_limits(container, wire_limits(limits)?)
        .map_err(|_error| Error::Wire)?;
    let mut raw = view.fields().filter(|field| field.number() == 2);
    for host in hosts.iter_mut().filter(|host| !host.ranges.is_empty()) {
        host.raw_range_records = try_vec(host.ranges.len())?;
        for _ in 0..host.ranges.len() {
            let field = raw.next().ok_or(Error::InvalidGraph)?;
            field
                .validate_canonical_framing()
                .map_err(|_error| Error::Wire)?;
            host.raw_range_records.push(Some(field.payload()));
        }
    }
    if raw.next().is_some() {
        return Err(Error::InvalidGraph);
    }
    Ok(())
}

fn attach_tiled_range_raw<'a>(
    ranges: &mut [(HostKey, Range, Option<&'a [u8]>)],
    container: &'a [u8],
    limits: Limits,
) -> Result<(), Error> {
    let view = WireView::parse_with_limits(container, wire_limits(limits)?)
        .map_err(|_error| Error::Wire)?;
    let mut raw = view.fields().filter(|field| field.number() == 2);
    for (_, _, destination) in ranges {
        let field = raw.next().ok_or(Error::InvalidGraph)?;
        field
            .validate_canonical_framing()
            .map_err(|_error| Error::Wire)?;
        *destination = Some(field.payload());
    }
    if raw.next().is_some() {
        return Err(Error::InvalidGraph);
    }
    Ok(())
}

fn coalesce_hosts(
    hosts: &mut Vec<OwnedHost<'_>>,
    table_owner: u32,
    limits: Limits,
) -> Result<(), Error> {
    let input = core::mem::take(hosts);
    let mut output = try_vec(input.len())?;
    let mut input = input.into_iter();
    while let Some(mut host) = input.next() {
        let group_end = input
            .as_slice()
            .partition_point(|candidate| candidate.key == host.key);
        let (additional_ranges, additional_raw) = input.as_slice()[..group_end].iter().try_fold(
            (0usize, 0usize),
            |(ranges, raw), candidate| {
                Ok::<_, Error>((
                    checked_add(ranges, candidate.ranges.len())?,
                    checked_add(raw, candidate.raw_range_records.len())?,
                ))
            },
        )?;
        reserve_exact_capacity(&mut host.ranges, additional_ranges)?;
        reserve_exact_capacity(&mut host.raw_range_records, additional_raw)?;
        for mut candidate in input.by_ref().take(group_end) {
            if host.has_cell_record && candidate.has_cell_record {
                // Numbers may mirror an opaque non-table declaration inline
                // and in a cell tile. Only an exact semantic mirror is
                // admissible; the two raw record payloads use different
                // enclosing field numbers and both source containers are
                // preserved unchanged. Selected-table duplicates remain an
                // authority conflict.
                let selected_exact_mirror = host.key.owner == table_owner
                    && host.cell_tile_object_id.is_none()
                        != candidate.cell_tile_object_id.is_none();
                if (!selected_exact_mirror && host.key.owner == table_owner)
                    || host.precedents != candidate.precedents
                    || host.is_in_cycle != candidate.is_in_cycle
                    || host.raw_cell_record != candidate.raw_cell_record
                {
                    return Err(Error::InvalidGraph);
                }
                if selected_exact_mirror {
                    host.cell_tile_object_id =
                        host.cell_tile_object_id.or(candidate.cell_tile_object_id);
                }
                candidate.has_cell_record = false;
            }
            if candidate.has_cell_record {
                host.precedents = core::mem::take(&mut candidate.precedents);
                host.cell_tile_object_id = candidate.cell_tile_object_id;
                host.has_cell_record = true;
                host.is_in_cycle = candidate.is_in_cycle;
                host.raw_cell_record = candidate.raw_cell_record;
            }
            host.ranges.extend(candidate.ranges);
            host.raw_range_records.extend(candidate.raw_range_records);
        }
        normalize_range_records(&mut host)?;
        output.push(host);
    }
    ensure_limit("coalesced hosts", output.len(), limits.max_hosts)?;
    *hosts = output;
    Ok(())
}

fn normalize_range_records(host: &mut OwnedHost<'_>) -> Result<(), Error> {
    if host.raw_range_records.len() != host.ranges.len() {
        return Err(Error::InvalidGraph);
    }
    let mut pairs = try_vec(host.ranges.len())?;
    for (range, raw) in host
        .ranges
        .iter()
        .copied()
        .zip(host.raw_range_records.iter().copied())
    {
        pairs.push((range, raw));
    }
    pairs.sort_unstable_by_key(|pair| pair.0);
    if pairs.windows(2).any(|pair| pair[0].0 == pair[1].0) {
        return Err(Error::InvalidGraph);
    }
    host.ranges.clear();
    host.raw_range_records.clear();
    for (range, raw) in pairs {
        host.ranges.push(range);
        host.raw_range_records.push(raw);
    }
    Ok(())
}

fn final_formula_count(source_count: u64, changes: &[HostChange<'_>]) -> Result<u64, Error> {
    let mut final_count = source_count;
    for change in changes {
        match (change.old.is_some(), change.new.is_some()) {
            (false, true) => final_count = final_count.checked_add(1).ok_or(Error::InvalidGraph)?,
            (true, false) => final_count = final_count.checked_sub(1).ok_or(Error::InvalidGraph)?,
            (true, true) => {},
            (false, false) => return Err(Error::InvalidGraph),
        }
    }
    Ok(final_count)
}

fn final_table_formula_count(
    existing_formula_hosts: &[HostKey],
    changes: &[HostChange<'_>],
) -> Result<u64, Error> {
    changes.iter().try_fold(
        u64::try_from(existing_formula_hosts.len()).map_err(|_| Error::InvalidGraph)?,
        |count, change| match (change.old.is_some(), change.new.is_some()) {
            (false, true) => count.checked_add(1).ok_or(Error::InvalidGraph),
            (true, false) => count.checked_sub(1).ok_or(Error::InvalidGraph),
            (true, true) => Ok(count),
            (false, false) => Err(Error::InvalidGraph),
        },
    )
}

#[allow(clippy::too_many_arguments)]
fn rewrite_owner(
    graph: CompleteGraph<'_, '_>,
    owner: &SourceOwner<'_, '_>,
    host_start: usize,
    host_end: usize,
    limits: Limits,
    decode_options: dependency::DecodeOptions,
    wire_limits: WireLimits,
    report: &mut Report,
    owners: &mut Vec<MessageEdit>,
    cell_edits: &mut Vec<MessageEdit>,
    range_edits: &mut Vec<MessageEdit>,
) -> Result<(), Error> {
    let hosts = &graph.hosts[host_start..host_end];
    let mut source_visitor = ReferenceCollector::new(checked_add(
        owner.cell_tiles.len(),
        owner.range_tiles.len(),
    )?)?;
    let (snapshot, source_report) = dependency::decode_formula_owner_dependencies_with_visitor(
        owner.message.payload,
        decode_options,
        &mut source_visitor,
    )
    .map_err(|_error| Error::StrictDependency)?;
    add_decode_report(report, source_report)?;
    if snapshot.internal_formula_owner_id() != owner.internal_owner {
        return Err(Error::InvalidGraph);
    }
    validate_tile_references(owner, &source_visitor)?;

    // Native Numbers publishes selected formula dependency records both
    // inline and, when routes exist, in the assigned CellRecordTile.  The two
    // canonical records must remain exact mirrors.
    let cell_payload = encode_inline_cells(hosts, limits)?;
    if !owner.cell_tiles.is_empty() {
        validate_cell_assignment(owner, graph.hosts, host_start, host_end)?;
        for tile in owner.cell_tiles {
            let assigned_count = graph.hosts[host_start..host_end]
                .iter()
                .filter(|host| host.cell_tile_object_id == Some(tile.message.object_id))
                .count();
            let mut assigned = try_vec(assigned_count)?;
            assigned.extend(
                graph.hosts[host_start..host_end]
                    .iter()
                    .filter(|host| host.cell_tile_object_id == Some(tile.message.object_id))
                    .copied(),
            );
            let new_base;
            let base = if tile.source_present {
                let mut source_facts = FactsVisitor::discard();
                let (tile_snapshot, tile_source_report) =
                    dependency::decode_cell_record_tile_with_visitor(
                        tile.message.payload,
                        decode_options,
                        &mut source_facts,
                    )
                    .map_err(|_error| Error::StrictDependency)?;
                add_decode_report(report, tile_source_report)?;
                if tile_snapshot.internal_owner_id() != owner.internal_owner
                    || tile_snapshot.tile_column_begin() != tile.tile_column_begin
                    || tile_snapshot.tile_row_begin() != tile.tile_row_begin
                {
                    return Err(Error::InvalidGraph);
                }
                tile.message.payload
            } else {
                new_base = encode_empty_cell_tile(
                    owner.internal_owner,
                    tile.tile_column_begin,
                    tile.tile_row_begin,
                    limits,
                )?;
                new_base.as_slice()
            };
            let records = encode_cell_records(&assigned, limits)?;
            let candidate = replace_repeated(base, 4, &records, wire_limits)?;
            let mut facts = FactsVisitor::expect(&assigned);
            let (candidate_snapshot, candidate_report) =
                dependency::decode_cell_record_tile_with_visitor(
                    &candidate,
                    decode_options,
                    &mut facts,
                )
                .map_err(|_error| Error::StrictDependency)?;
            add_decode_report(report, candidate_report)?;
            if candidate_snapshot.internal_owner_id() != owner.internal_owner
                || candidate_snapshot.tile_column_begin() != tile.tile_column_begin
                || candidate_snapshot.tile_row_begin() != tile.tile_row_begin
                || !facts.complete()
            {
                return Err(Error::InvalidGraph);
            }
            cell_edits.push(finish_edit(tile.message, candidate, report)?);
        }
    }
    let tiled_cell_ids: Vec<u64> = owner
        .cell_tiles
        .iter()
        .map(|tile| tile.message.object_id)
        .collect();
    let tiled_cell_payload = encode_reference_list(&tiled_cell_ids, limits)?;
    let owner_references = final_owner_references(owner)?;

    let range_payload = if owner.range_tiles.is_empty() {
        Some(encode_inline_ranges(hosts, limits)?)
    } else {
        rewrite_range_tiles(
            owner,
            hosts,
            limits,
            decode_options,
            wire_limits,
            report,
            range_edits,
        )?;
        None
    };
    let cell_replacement = (!cell_payload.is_empty()).then_some(cell_payload.as_slice());
    let tiled_cell_replacement =
        (!tiled_cell_payload.is_empty()).then_some(tiled_cell_payload.as_slice());
    let range_replacement = range_payload
        .as_deref()
        .filter(|payload| !payload.is_empty());
    let candidate = patch_nested_fields_batched_with_limits(
        owner.message.payload,
        &[
            NestedFieldEdit::new(
                &OWNER_CELL_PATH,
                snapshot.cell_dependencies().is_some(),
                NestedFieldReplacement::LengthDelimited(cell_replacement),
            ),
            NestedFieldEdit::new(
                &OWNER_RANGE_PATH,
                snapshot.range_dependencies().is_some(),
                NestedFieldReplacement::LengthDelimited(range_replacement),
            ),
            NestedFieldEdit::new(
                &OWNER_TILED_CELL_PATH,
                snapshot.tiled_cell_dependencies().is_some(),
                NestedFieldReplacement::LengthDelimited(tiled_cell_replacement),
            ),
        ],
        wire_limits,
    )
    .map_err(|_error| Error::Wire)?;
    let mut candidate_refs = ReferenceCollector::new(checked_add(
        owner.cell_tiles.len(),
        owner.range_tiles.len(),
    )?)?;
    let (candidate_snapshot, candidate_report) =
        dependency::decode_formula_owner_dependencies_with_visitor(
            &candidate,
            decode_options,
            &mut candidate_refs,
        )
        .map_err(|_error| Error::StrictDependency)?;
    add_decode_report(report, candidate_report)?;
    if candidate_snapshot.internal_formula_owner_id() != owner.internal_owner
        || candidate_refs.tiled_cells != tiled_cell_ids
        || candidate_refs.tiled_ranges != source_visitor.tiled_ranges
    {
        return Err(Error::InvalidGraph);
    }
    let mut facts = FactsVisitor::expect(hosts);
    let (_facts_snapshot, facts_report) =
        dependency::decode_formula_owner_dependencies_with_visitor(
            &candidate,
            decode_options,
            &mut facts,
        )
        .map_err(|_error| Error::StrictDependency)?;
    add_decode_report(report, facts_report)?;
    if (owner.cell_tiles.is_empty() && !facts.cells_complete())
        || (owner.range_tiles.is_empty() && !facts.ranges_complete())
    {
        return Err(Error::InvalidGraph);
    }
    owners.push(finish_edit_with_references(
        owner.message,
        candidate,
        owner_references,
        report,
    )?);
    Ok(())
}

fn rewrite_range_tiles(
    owner: &SourceOwner<'_, '_>,
    hosts: &[FormulaHost<'_>],
    limits: Limits,
    decode_options: dependency::DecodeOptions,
    wire_limits: WireLimits,
    report: &mut Report,
    edits: &mut Vec<MessageEdit>,
) -> Result<(), Error> {
    let range_count = hosts
        .iter()
        .try_fold(0usize, |count, host| checked_add(count, host.ranges.len()))?;
    let mut grouped = try_vec(range_count)?;
    for host in hosts {
        if host.raw_range_records.len() != host.ranges.len() {
            return Err(Error::InvalidGraph);
        }
        for (range, raw) in host.ranges.iter().zip(host.raw_range_records) {
            grouped.push(RangeFact::new(host, *range, *raw));
        }
    }
    grouped.sort_unstable_by_key(|fact| fact.sort_key());
    let mut previous_target = None;
    for tile in owner.range_tiles {
        let mut discard = FactsVisitor::discard();
        let (snapshot, source_report) = dependency::decode_range_precedents_tile_with_visitor(
            tile.message.payload,
            decode_options,
            &mut discard,
        )
        .map_err(|_error| Error::StrictDependency)?;
        add_decode_report(report, source_report)?;
        let target = snapshot.to_owner_id();
        if target != tile.target_owner {
            return Err(Error::InvalidGraph);
        }
        if previous_target.is_some_and(|prior| prior >= target) {
            return Err(Error::InvalidGraph);
        }
        previous_target = Some(target);
        let start = grouped.partition_point(|fact| fact.target < target);
        let end = grouped.partition_point(|fact| fact.target <= target);
        let assigned = &grouped[start..end];
        let records = encode_tiled_ranges(assigned, limits)?;
        let candidate = replace_repeated(tile.message.payload, 2, &records, wire_limits)?;
        let mut facts = RangeTileVisitor::new(assigned);
        let (candidate_snapshot, candidate_report) =
            dependency::decode_range_precedents_tile_with_visitor(
                &candidate,
                decode_options,
                &mut facts,
            )
            .map_err(|_error| Error::StrictDependency)?;
        add_decode_report(report, candidate_report)?;
        if candidate_snapshot != snapshot || !facts.complete() {
            return Err(Error::InvalidGraph);
        }
        edits.push(finish_edit(tile.message, candidate, report)?);
    }
    for fact in &grouped {
        if owner
            .range_tiles
            .binary_search_by_key(&fact.target, |tile| tile.target_owner)
            .is_err()
        {
            return Err(Error::InvalidGraph);
        }
    }
    Ok(())
}

struct ReferenceCollector {
    formula_owners: Vec<u64>,
    tiled_cells: Vec<u64>,
    tiled_ranges: Vec<u64>,
}

impl ReferenceCollector {
    fn new(max_references: usize) -> Result<Self, Error> {
        Ok(Self {
            formula_owners: try_vec(max_references)?,
            tiled_cells: try_vec(max_references)?,
            tiled_ranges: try_vec(max_references)?,
        })
    }
}

impl dependency::DependencyVisitor for ReferenceCollector {
    fn visit_formula_owner_dependency(
        &mut self,
        record: dependency::ReferenceRecord<'_>,
    ) -> Result<(), dependency::DecodeError> {
        self.formula_owners.push(record.reference().identifier());
        Ok(())
    }

    fn visit_tiled_cell_dependency(
        &mut self,
        record: dependency::ReferenceRecord<'_>,
    ) -> Result<(), dependency::DecodeError> {
        self.tiled_cells.push(record.reference().identifier());
        Ok(())
    }

    fn visit_tiled_range_dependency(
        &mut self,
        record: dependency::ReferenceRecord<'_>,
    ) -> Result<(), dependency::DecodeError> {
        self.tiled_ranges.push(record.reference().identifier());
        Ok(())
    }
}

struct FactsVisitor<'a> {
    expected: Option<&'a [FormulaHost<'a>]>,
    cell_index: usize,
    range_host: usize,
    range_index: usize,
    component_position: usize,
    invalid: bool,
}

impl<'a> FactsVisitor<'a> {
    fn discard() -> Self {
        Self {
            expected: None,
            cell_index: 0,
            range_host: 0,
            range_index: 0,
            component_position: 0,
            invalid: false,
        }
    }

    fn expect(expected: &'a [FormulaHost<'a>]) -> Self {
        Self {
            expected: Some(expected),
            ..Self::discard()
        }
    }

    fn cells_complete(&self) -> bool {
        self.expected
            .is_some_and(|hosts| self.cell_index == hosts.len())
            && !self.invalid
    }

    fn ranges_complete(&self) -> bool {
        let Some(hosts) = self.expected else {
            return false;
        };
        let mut total = 0usize;
        for host in hosts {
            total = total.saturating_add(host.ranges.len());
        }
        let consumed = hosts[..self.range_host.min(hosts.len())]
            .iter()
            .fold(0usize, |sum, host| sum.saturating_add(host.ranges.len()))
            .saturating_add(self.range_index);
        consumed == total && !self.invalid
    }

    fn complete(&self) -> bool {
        self.cells_complete() && self.component_position == 0
    }
}

impl dependency::DependencyVisitor for FactsVisitor<'_> {
    fn visit_expanded_edge_component(
        &mut self,
        component: dependency::ExpandedEdgeComponent,
    ) -> Result<(), dependency::DecodeError> {
        if let Some(expected) = self.expected {
            let matches = expected
                .get(self.cell_index)
                .and_then(|host| expected_component(host, self.component_position))
                .is_some_and(|(kind, index, value)| {
                    component.kind() == kind
                        && component.index() == index
                        && component.value() == value
                });
            if !matches {
                self.invalid = true;
            }
        }
        self.component_position = self.component_position.saturating_add(1);
        Ok(())
    }

    fn visit_cell_record(
        &mut self,
        record: dependency::CellRecordSnapshot<'_>,
    ) -> Result<(), dependency::DecodeError> {
        let Some(expected) = self.expected else {
            self.component_position = 0;
            return Ok(());
        };
        let Some(host) = expected.get(self.cell_index) else {
            self.invalid = true;
            self.component_position = 0;
            return Ok(());
        };
        if record.row() != host.row
            || record.column() != host.column
            || record.is_in_a_cycle().is_some()
            || record.dirty_self_plus_precedents_count().is_some()
            || record.has_calculated_precedents().is_some()
            || expected_component(host, self.component_position).is_some()
        {
            self.invalid = true;
        }
        self.component_position = 0;
        self.cell_index += 1;
        Ok(())
    }

    fn visit_range_back_dependency(
        &mut self,
        record: dependency::RangeBackDependencySnapshot<'_>,
    ) -> Result<(), dependency::DecodeError> {
        let Some(hosts) = self.expected else {
            return Ok(());
        };
        while self.range_host < hosts.len()
            && self.range_index == hosts[self.range_host].ranges.len()
        {
            self.range_host += 1;
            self.range_index = 0;
        }
        let Some(host) = hosts.get(self.range_host) else {
            self.invalid = true;
            return Ok(());
        };
        let Some(expected) = host.ranges.get(self.range_index) else {
            self.invalid = true;
            return Ok(());
        };
        let matches = record.cell_coord_row() == host.row
            && record.cell_coord_column() == host.column
            && record.decoded_range_reference().is_none()
            && record
                .decoded_internal_range_reference()
                .is_some_and(|actual| internal_range_matches(actual, *expected));
        if !matches {
            self.invalid = true;
        }
        self.range_index += 1;
        Ok(())
    }
}

struct RangeTileVisitor<'a> {
    facts: &'a [RangeFact<'a>],
    seen: usize,
    invalid: bool,
}

impl<'a> RangeTileVisitor<'a> {
    const fn new(facts: &'a [RangeFact<'a>]) -> Self {
        Self {
            facts,
            seen: 0,
            invalid: false,
        }
    }

    fn complete(&self) -> bool {
        !self.invalid && self.seen == self.facts.len()
    }
}

impl dependency::DependencyVisitor for RangeTileVisitor<'_> {
    fn visit_from_to_range(
        &mut self,
        record: dependency::FromToRangeSnapshot<'_>,
    ) -> Result<(), dependency::DecodeError> {
        let Some(expected) = self.facts.get(self.seen).copied() else {
            self.invalid = true;
            return Ok(());
        };
        let from = record.decoded_from_coord();
        let rect = record.decoded_refers_to_rect();
        let width = expected
            .right
            .checked_sub(expected.left)
            .and_then(|v| v.checked_add(1));
        let height = expected
            .bottom
            .checked_sub(expected.top)
            .and_then(|v| v.checked_add(1));
        if from.packed_data().is_some()
            || from.column() != Some(expected.host_column)
            || from.row() != Some(expected.host_row)
            || rect.origin().packed_data().is_some()
            || rect.origin().column() != Some(expected.left)
            || rect.origin().row() != Some(expected.top)
            || rect.size().num_columns() != width
            || rect.size().num_rows() != height
        {
            self.invalid = true;
        }
        self.seen += 1;
        Ok(())
    }
}

fn expected_component(
    host: &FormulaHost<'_>,
    wanted: usize,
) -> Option<(dependency::ExpandedEdgeKind, usize, u32)> {
    let mut position = 0usize;
    for (kind, local) in [
        (dependency::ExpandedEdgeKind::LocalRow, true),
        (dependency::ExpandedEdgeKind::LocalColumn, true),
        (dependency::ExpandedEdgeKind::ExternalRow, false),
        (dependency::ExpandedEdgeKind::ExternalColumn, false),
        (dependency::ExpandedEdgeKind::InternalOwner, false),
    ] {
        let mut index = 0usize;
        for precedent in host.precedents {
            if (precedent.target_owner == host.owner) != local {
                continue;
            }
            let value = match kind {
                dependency::ExpandedEdgeKind::LocalRow
                | dependency::ExpandedEdgeKind::ExternalRow => precedent.row,
                dependency::ExpandedEdgeKind::LocalColumn
                | dependency::ExpandedEdgeKind::ExternalColumn => precedent.column,
                dependency::ExpandedEdgeKind::InternalOwner => precedent.target_owner,
            };
            if position == wanted {
                return Some((kind, index, value));
            }
            position += 1;
            index += 1;
        }
    }
    None
}

fn internal_range_matches(
    actual: dependency::InternalRangeReferenceSnapshot,
    expected: Range,
) -> bool {
    let range = actual.range();
    actual.owner_id() == expected.target_owner
        && range.top_left_column() == expected.left
        && range.top_left_row() == expected.top
        && range.bottom_right_column() == expected.right
        && range.bottom_right_row() == expected.bottom
}

fn cfuuid_halves(uuid: dependency::CfuuidSnapshot<'_>) -> Option<(u64, u64)> {
    let words = uuid.words();
    if let [Some(w0), Some(w1), Some(w2), Some(w3)] = words {
        return Some((
            u64::from(w0) | (u64::from(w1) << 32),
            u64::from(w2) | (u64::from(w3) << 32),
        ));
    }
    let bytes: [u8; 16] = uuid.uuid_bytes()?.try_into().ok()?;
    let value = u128::from_be_bytes(bytes);
    Some((value as u64, (value >> 64) as u64))
}

fn encode_inline_cells(hosts: &[FormulaHost<'_>], limits: Limits) -> Result<Vec<u8>, Error> {
    let records = encode_cell_records(hosts, limits)?;
    encode_repeated_message(1, &records, limits.max_output_bytes)
}

fn encode_cell_records(hosts: &[FormulaHost<'_>], limits: Limits) -> Result<Vec<Vec<u8>>, Error> {
    let mut records = try_vec(hosts.len())?;
    for host in hosts {
        records.push(encode_cell_record(host, limits)?);
    }
    Ok(records)
}

fn encode_cell_record(host: &FormulaHost<'_>, limits: Limits) -> Result<Vec<u8>, Error> {
    if let Some(raw) = host.raw_cell_record {
        let mut output = try_vec_bytes(raw.len())?;
        output.extend_from_slice(raw);
        return Ok(output);
    }
    let mut local_rows = try_vec(host.precedents.len())?;
    let mut local_columns = try_vec(host.precedents.len())?;
    let mut external_rows = try_vec(host.precedents.len())?;
    let mut external_columns = try_vec(host.precedents.len())?;
    let mut external_owners = try_vec(host.precedents.len())?;
    for precedent in host.precedents {
        if precedent.target_owner == host.owner {
            local_rows.push(precedent.row);
            local_columns.push(precedent.column);
        } else {
            external_rows.push(precedent.row);
            external_columns.push(precedent.column);
            external_owners.push(precedent.target_owner);
        }
    }
    let edge_payload_len = [
        &local_rows,
        &local_columns,
        &external_rows,
        &external_columns,
        &external_owners,
    ]
    .iter()
    .enumerate()
    .try_fold(0usize, |total, (index, values)| {
        if values.is_empty() {
            Ok(total)
        } else {
            let payload = values.iter().try_fold(0usize, |sum, value| {
                checked_add(sum, encoded_len(u64::from(*value)))
            })?;
            checked_add(
                total,
                message_field_len(
                    u32::try_from(index + 1).map_err(|_error| Error::InvalidGraph)?,
                    payload,
                )?,
            )
        }
    })?;
    let mut edges = try_vec_bytes(edge_payload_len)?;
    append_packed_u32(&mut edges, 1, &local_rows)?;
    append_packed_u32(&mut edges, 2, &local_columns)?;
    append_packed_u32(&mut edges, 3, &external_rows)?;
    append_packed_u32(&mut edges, 4, &external_columns)?;
    append_packed_u32(&mut edges, 5, &external_owners)?;
    let size = checked_add(
        varint_field_len(1, u64::from(host.column))?,
        varint_field_len(2, u64::from(host.row))?,
    )?;
    let size = if edges.is_empty() {
        size
    } else {
        checked_add(size, message_field_len(6, edges.len())?)?
    };
    ensure_limit("encoded cell record", size, limits.max_output_bytes)?;
    let mut output = try_vec_bytes(size)?;
    append_varint(&mut output, 1, u64::from(host.column));
    append_varint(&mut output, 2, u64::from(host.row));
    if !edges.is_empty() {
        append_message(&mut output, 6, &edges)?;
    }
    Ok(output)
}

fn encode_inline_ranges(hosts: &[FormulaHost<'_>], limits: Limits) -> Result<Vec<u8>, Error> {
    let mut records = try_vec(hosts.iter().map(|host| host.ranges.len()).sum())?;
    for host in hosts {
        if host.raw_range_records.len() != host.ranges.len() {
            return Err(Error::InvalidGraph);
        }
        for (range, raw) in host.ranges.iter().zip(host.raw_range_records) {
            if let Some(raw) = raw {
                let mut retained = try_vec_bytes(raw.len())?;
                retained.extend_from_slice(raw);
                records.push(retained);
            } else {
                records.push(encode_range_back(host, *range, limits)?);
            }
        }
    }
    encode_repeated_message(2, &records, limits.max_output_bytes)
}

fn encode_range_back(
    host: &FormulaHost<'_>,
    range: Range,
    limits: Limits,
) -> Result<Vec<u8>, Error> {
    let coordinates_len = [range.left, range.top, range.right, range.bottom]
        .iter()
        .enumerate()
        .try_fold(0usize, |sum, (index, value)| {
            checked_add(
                sum,
                varint_field_len(
                    u32::try_from(index + 1).map_err(|_error| Error::InvalidGraph)?,
                    u64::from(*value),
                )?,
            )
        })?;
    let mut coordinates = try_vec_bytes(coordinates_len)?;
    append_varint(&mut coordinates, 1, u64::from(range.left));
    append_varint(&mut coordinates, 2, u64::from(range.top));
    append_varint(&mut coordinates, 3, u64::from(range.right));
    append_varint(&mut coordinates, 4, u64::from(range.bottom));
    let internal_len = checked_add(
        varint_field_len(1, u64::from(range.target_owner))?,
        message_field_len(2, coordinates.len())?,
    )?;
    let mut internal = try_vec_bytes(internal_len)?;
    append_varint(&mut internal, 1, u64::from(range.target_owner));
    append_message(&mut internal, 2, &coordinates)?;
    let size = checked_add(
        checked_add(
            varint_field_len(1, u64::from(host.row))?,
            varint_field_len(2, u64::from(host.column))?,
        )?,
        message_field_len(4, internal.len())?,
    )?;
    ensure_limit("encoded range record", size, limits.max_output_bytes)?;
    let mut output = try_vec_bytes(size)?;
    append_varint(&mut output, 1, u64::from(host.row));
    append_varint(&mut output, 2, u64::from(host.column));
    append_message(&mut output, 4, &internal)?;
    Ok(output)
}

fn encode_tiled_ranges(facts: &[RangeFact<'_>], limits: Limits) -> Result<Vec<Vec<u8>>, Error> {
    let mut records = try_vec(facts.len())?;
    for fact in facts {
        records.push(encode_from_to_range(*fact, limits)?);
    }
    Ok(records)
}

fn encode_from_to_range(fact: RangeFact<'_>, limits: Limits) -> Result<Vec<u8>, Error> {
    if let Some(raw) = fact.raw {
        let mut output = try_vec_bytes(raw.len())?;
        output.extend_from_slice(raw);
        return Ok(output);
    }
    let width = fact
        .right
        .checked_sub(fact.left)
        .and_then(|v| v.checked_add(1))
        .ok_or(Error::InvalidGraph)?;
    let height = fact
        .bottom
        .checked_sub(fact.top)
        .and_then(|v| v.checked_add(1))
        .ok_or(Error::InvalidGraph)?;
    let from_len = checked_add(
        varint_field_len(2, u64::from(fact.host_column))?,
        varint_field_len(3, u64::from(fact.host_row))?,
    )?;
    let mut from = try_vec_bytes(from_len)?;
    append_varint(&mut from, 2, u64::from(fact.host_column));
    append_varint(&mut from, 3, u64::from(fact.host_row));
    let origin_len = checked_add(
        varint_field_len(2, u64::from(fact.left))?,
        varint_field_len(3, u64::from(fact.top))?,
    )?;
    let mut origin = try_vec_bytes(origin_len)?;
    append_varint(&mut origin, 2, u64::from(fact.left));
    append_varint(&mut origin, 3, u64::from(fact.top));
    let size_len = checked_add(
        varint_field_len(1, u64::from(width))?,
        varint_field_len(2, u64::from(height))?,
    )?;
    let mut size = try_vec_bytes(size_len)?;
    append_varint(&mut size, 1, u64::from(width));
    append_varint(&mut size, 2, u64::from(height));
    let rect_len = checked_add(
        message_field_len(1, origin.len())?,
        message_field_len(2, size.len())?,
    )?;
    let mut rect = try_vec_bytes(rect_len)?;
    append_message(&mut rect, 1, &origin)?;
    append_message(&mut rect, 2, &size)?;
    let output_len = checked_add(
        message_field_len(1, from.len())?,
        message_field_len(2, rect.len())?,
    )?;
    ensure_limit("encoded tiled range", output_len, limits.max_output_bytes)?;
    let mut output = try_vec_bytes(output_len)?;
    append_message(&mut output, 1, &from)?;
    append_message(&mut output, 2, &rect)?;
    Ok(output)
}

fn replace_repeated(
    source: &[u8],
    number: u32,
    payloads: &[Vec<u8>],
    limits: WireLimits,
) -> Result<Vec<u8>, Error> {
    let view = WireView::parse_with_limits(source, limits).map_err(|_error| Error::Wire)?;
    let inserted = payloads.iter().try_fold(0usize, |sum, payload| {
        checked_add(sum, message_field_len(number, payload.len())?)
    })?;
    let mut removed = 0usize;
    let mut found = false;
    for field in view.fields() {
        if field.number() == number {
            if field.wire_type() != 2 {
                return Err(Error::Wire);
            }
            field
                .validate_canonical_framing()
                .map_err(|_error| Error::Wire)?;
            removed = checked_add(removed, field.raw().len())?;
            found = true;
        }
    }
    let size = source
        .len()
        .checked_sub(removed)
        .and_then(|value| value.checked_add(inserted))
        .ok_or(Error::InvalidGraph)?;
    ensure_limit("wire output", size, limits.max_output_bytes())?;
    let mut output = try_vec_bytes(size)?;
    let mut emitted = false;
    for field in view.fields() {
        if field.number() == number {
            if !emitted {
                append_payloads(&mut output, number, payloads)?;
                emitted = true;
            }
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    if !found {
        append_payloads(&mut output, number, payloads)?;
    }
    if output.len() != size {
        return Err(Error::InvalidGraph);
    }
    Ok(output)
}

fn encode_repeated_message(
    number: u32,
    payloads: &[Vec<u8>],
    maximum: usize,
) -> Result<Vec<u8>, Error> {
    let size = payloads.iter().try_fold(0usize, |sum, payload| {
        checked_add(sum, message_field_len(number, payload.len())?)
    })?;
    ensure_limit("encoded repeated message", size, maximum)?;
    let mut output = try_vec_bytes(size)?;
    append_payloads(&mut output, number, payloads)?;
    Ok(output)
}

fn encode_reference_list(ids: &[u64], limits: Limits) -> Result<Vec<u8>, Error> {
    let mut references = try_vec(ids.len())?;
    for id in ids {
        let size = varint_field_len(1, *id)?;
        let mut reference = try_vec_bytes(size)?;
        append_varint(&mut reference, 1, *id);
        references.push(reference);
    }
    encode_repeated_message(1, &references, limits.max_output_bytes)
}

fn encode_empty_cell_tile(
    internal_owner: u32,
    tile_column_begin: u32,
    tile_row_begin: u32,
    limits: Limits,
) -> Result<Vec<u8>, Error> {
    let size = checked_add(
        varint_field_len(1, u64::from(internal_owner))?,
        checked_add(
            varint_field_len(2, u64::from(tile_column_begin))?,
            varint_field_len(3, u64::from(tile_row_begin))?,
        )?,
    )?;
    ensure_limit("encoded empty cell tile", size, limits.max_output_bytes)?;
    let mut output = try_vec_bytes(size)?;
    append_varint(&mut output, 1, u64::from(internal_owner));
    append_varint(&mut output, 2, u64::from(tile_column_begin));
    append_varint(&mut output, 3, u64::from(tile_row_begin));
    Ok(output)
}

fn final_owner_references(owner: &SourceOwner<'_, '_>) -> Result<Vec<u64>, Error> {
    let new_count = owner
        .cell_tiles
        .iter()
        .filter(|tile| !tile.source_present)
        .count();
    let mut references = try_vec(checked_add(
        owner.message.object_references.len(),
        new_count,
    )?)?;
    references.extend_from_slice(owner.message.object_references);
    for tile in owner.cell_tiles.iter().filter(|tile| !tile.source_present) {
        if references.contains(&tile.message.object_id) {
            return Err(Error::InvalidGraph);
        }
        references.push(tile.message.object_id);
    }
    Ok(references)
}

fn append_payloads(output: &mut Vec<u8>, number: u32, payloads: &[Vec<u8>]) -> Result<(), Error> {
    for payload in payloads {
        append_message(output, number, payload)?;
    }
    Ok(())
}

fn append_packed_u32(output: &mut Vec<u8>, number: u32, values: &[u32]) -> Result<(), Error> {
    if values.is_empty() {
        return Ok(());
    }
    let payload_len = values.iter().try_fold(0usize, |sum, value| {
        checked_add(sum, encoded_len(u64::from(*value)))
    })?;
    encode_key(output, number, 2)?;
    encode_varint_into(
        output,
        u64::try_from(payload_len).map_err(|_error| Error::InvalidGraph)?,
    );
    for value in values {
        encode_varint_into(output, u64::from(*value));
    }
    Ok(())
}

fn append_varint(output: &mut Vec<u8>, number: u32, value: u64) {
    encode_varint_into(output, u64::from(number) << 3);
    encode_varint_into(output, value);
}

fn append_message(output: &mut Vec<u8>, number: u32, payload: &[u8]) -> Result<(), Error> {
    encode_key(output, number, 2)?;
    encode_varint_into(
        output,
        u64::try_from(payload.len()).map_err(|_error| Error::InvalidGraph)?,
    );
    output.extend_from_slice(payload);
    Ok(())
}

fn encode_key(output: &mut Vec<u8>, number: u32, wire_type: u8) -> Result<(), Error> {
    if number == 0 || number > 0x1fff_ffff || wire_type > 5 {
        return Err(Error::InvalidGraph);
    }
    encode_varint_into(output, (u64::from(number) << 3) | u64::from(wire_type));
    Ok(())
}

fn varint_field_len(number: u32, value: u64) -> Result<usize, Error> {
    checked_add(encoded_len(u64::from(number) << 3), encoded_len(value))
}

fn message_field_len(number: u32, payload_len: usize) -> Result<usize, Error> {
    let length = u64::try_from(payload_len).map_err(|_error| Error::InvalidGraph)?;
    checked_add(
        encoded_len((u64::from(number) << 3) | 2),
        checked_add(encoded_len(length), payload_len)?,
    )
}

fn validate_engine_refs(graph: CompleteGraph<'_, '_>, actual: &[u64]) -> Result<(), Error> {
    if actual.len() != graph.owners.len() {
        return Err(Error::InvalidGraph);
    }
    for (reference, owner) in actual.iter().zip(graph.owners) {
        if *reference != owner.message.object_id {
            return Err(Error::InvalidGraph);
        }
    }
    Ok(())
}

fn validate_archive_info_references(graph: SourceGraph<'_, '_>) -> Result<(), Error> {
    let mut semantic = try_vec(graph.owners.len())?;
    for (owner_order, owner) in graph.owners.iter().enumerate() {
        semantic.push(ArchiveOwnerReference {
            object_id: owner.message.object_id,
            owner_order,
            occurrence: None,
        });
    }
    semantic.sort_unstable_by_key(|entry| entry.object_id);
    if semantic
        .windows(2)
        .any(|pair| pair[0].object_id == pair[1].object_id)
    {
        return Err(Error::InvalidGraph);
    }
    for (position, reference) in graph.engine.object_references.iter().enumerate() {
        if let Ok(index) = semantic.binary_search_by_key(reference, |entry| entry.object_id) {
            if semantic[index].occurrence.replace(position).is_some() {
                return Err(Error::InvalidGraph);
            }
        }
    }
    semantic.sort_unstable_by_key(|entry| entry.owner_order);
    let mut previous_position = None;
    for entry in semantic {
        let position = entry.occurrence.ok_or(Error::InvalidGraph)?;
        if previous_position.is_some_and(|previous| previous >= position) {
            return Err(Error::InvalidGraph);
        }
        previous_position = Some(position);
    }
    for owner in graph.owners {
        let source_cell_tiles = owner
            .cell_tiles
            .iter()
            .filter(|tile| tile.source_present)
            .count();
        let expected = checked_add(source_cell_tiles, owner.range_tiles.len())?;
        let mut semantic = try_vec(expected)?;
        for (owner_order, tile) in owner
            .cell_tiles
            .iter()
            .filter(|tile| tile.source_present)
            .enumerate()
        {
            semantic.push(ArchiveOwnerReference {
                object_id: tile.message.object_id,
                owner_order,
                occurrence: None,
            });
            if !tile.message.object_references.is_empty() {
                return Err(Error::InvalidGraph);
            }
        }
        for (index, tile) in owner.range_tiles.iter().enumerate() {
            semantic.push(ArchiveOwnerReference {
                object_id: tile.message.object_id,
                owner_order: checked_add(source_cell_tiles, index)?,
                occurrence: None,
            });
            if !tile.message.object_references.is_empty() {
                return Err(Error::InvalidGraph);
            }
        }
        semantic.sort_unstable_by_key(|entry| entry.object_id);
        if semantic
            .windows(2)
            .any(|pair| pair[0].object_id == pair[1].object_id)
        {
            return Err(Error::InvalidGraph);
        }
        for (position, reference) in owner.message.object_references.iter().enumerate() {
            if let Ok(index) = semantic.binary_search_by_key(reference, |entry| entry.object_id) {
                if semantic[index].occurrence.replace(position).is_some() {
                    return Err(Error::InvalidGraph);
                }
            }
        }
        semantic.sort_unstable_by_key(|entry| entry.owner_order);
        let mut previous_position = None;
        for entry in semantic {
            let position = entry.occurrence.ok_or(Error::InvalidGraph)?;
            if previous_position.is_some_and(|previous| previous >= position) {
                return Err(Error::InvalidGraph);
            }
            previous_position = Some(position);
        }
    }
    Ok(())
}

fn validate_tile_references(
    owner: &SourceOwner<'_, '_>,
    actual: &ReferenceCollector,
) -> Result<(), Error> {
    let source_cells = owner.cell_tiles.iter().filter(|tile| tile.source_present);
    if actual.tiled_cells.len() != source_cells.clone().count()
        || actual.tiled_ranges.len() != owner.range_tiles.len()
    {
        return Err(Error::InvalidGraph);
    }
    for (reference, tile) in actual.tiled_cells.iter().zip(source_cells) {
        if *reference != tile.message.object_id {
            return Err(Error::InvalidGraph);
        }
    }
    for (reference, tile) in actual.tiled_ranges.iter().zip(owner.range_tiles) {
        if *reference != tile.message.object_id {
            return Err(Error::InvalidGraph);
        }
    }
    Ok(())
}

fn validate_cell_assignment(
    owner: &SourceOwner<'_, '_>,
    hosts: &[FormulaHost<'_>],
    start: usize,
    end: usize,
) -> Result<(), Error> {
    if owner
        .cell_tiles
        .windows(2)
        .any(|pair| pair[0].message.object_id == pair[1].message.object_id)
        || hosts[start..end].iter().any(|host| {
            host.owner != owner.internal_owner
                || owner
                    .cell_tiles
                    .iter()
                    .filter(|tile| Some(tile.message.object_id) == host.cell_tile_object_id)
                    .count()
                    != 1
        })
    {
        return Err(Error::InvalidGraph);
    }
    Ok(())
}

fn finish_edit(
    source: SourceMessage<'_>,
    candidate: Vec<u8>,
    report: &mut Report,
) -> Result<MessageEdit, Error> {
    report.output_bytes = checked_add(report.output_bytes, candidate.len())?;
    let payload = if candidate == source.payload {
        None
    } else {
        report.changed_messages = checked_add(report.changed_messages, 1)?;
        Some(candidate)
    };
    let mut references = try_vec(source.object_references.len())?;
    references.extend_from_slice(source.object_references);
    Ok(MessageEdit {
        object_id: source.object_id,
        payload,
        object_references: references,
    })
}

fn finish_edit_with_references(
    source: SourceMessage<'_>,
    candidate: Vec<u8>,
    references: Vec<u64>,
    report: &mut Report,
) -> Result<MessageEdit, Error> {
    report.output_bytes = checked_add(report.output_bytes, candidate.len())?;
    let payload = if candidate == source.payload {
        None
    } else {
        report.changed_messages = checked_add(report.changed_messages, 1)?;
        Some(candidate)
    };
    Ok(MessageEdit {
        object_id: source.object_id,
        payload,
        object_references: references,
    })
}

fn add_decode_report(report: &mut Report, observed: dependency::DecodeReport) -> Result<(), Error> {
    report.fields = checked_add(report.fields, observed.fields())?;
    report.strict_work_bytes = checked_add(report.strict_work_bytes, observed.work_bytes())?;
    report.references = checked_add(report.references, observed.references())?;
    report.reference_bytes = checked_add(report.reference_bytes, observed.reference_bytes())?;
    report.text_bytes = checked_add(report.text_bytes, observed.text_bytes())?;
    report.max_depth = report.max_depth.max(observed.max_depth());
    Ok(())
}

fn remaining_decode_options(
    limits: Limits,
    report: Report,
) -> Result<dependency::DecodeOptions, Error> {
    let fields = limits
        .max_fields
        .checked_sub(report.fields)
        .ok_or(Error::InvalidGraph)?;
    let work = limits
        .max_work_bytes
        .checked_sub(report.strict_work_bytes)
        .ok_or(Error::InvalidGraph)?;
    let references = limits
        .max_references
        .checked_sub(report.references)
        .ok_or(Error::InvalidGraph)?;
    if fields == 0 || work == 0 {
        return Err(Error::Limit {
            resource: "aggregate strict decode",
            observed: report.fields.max(report.strict_work_bytes),
            maximum: limits.max_fields.max(limits.max_work_bytes),
        });
    }
    Ok(dependency::DecodeOptions::new(
        limits.max_source_bytes,
        fields,
        work,
        limits.recursion_limit,
        references,
        1,
    ))
}

#[derive(Debug)]
struct Preflight {
    object_ids: Vec<u64>,
    source_bytes: usize,
    output_upper: usize,
    field_upper: usize,
    strict_work_upper: usize,
    reference_upper: usize,
    retained_upper: usize,
    messages: usize,
    hosts: usize,
    precedents: usize,
    ranges: usize,
    cell_tiles: usize,
    range_tiles: usize,
    allocations: usize,
    peak_scratch_bytes: usize,
    graph_work_bytes: usize,
}

impl Preflight {
    fn requirements(&self) -> Result<ExecutionRequirements, Error> {
        Ok(ExecutionRequirements {
            output_bytes: self.output_upper,
            fields: self.field_upper,
            work_bytes: checked_add(self.strict_work_upper, self.graph_work_bytes)?,
            references: self.reference_upper,
            allocations: self.allocations,
            peak_scratch_bytes: self.peak_scratch_bytes,
            retained_bytes: self.retained_upper,
            retained_elements: checked_add(self.messages, self.reference_upper)?,
            objects: self.messages,
            message_edits: self.messages,
            hosts: self.hosts,
            precedents: self.precedents,
            ranges: self.ranges,
        })
    }

    fn plan_scratch_bytes(&self) -> usize {
        self.object_ids.capacity().saturating_mul(size_of::<u64>())
    }
}

#[derive(Debug, Clone, Copy)]
struct RangeFact<'a> {
    target: u32,
    host_row: u32,
    host_column: u32,
    top: u32,
    left: u32,
    bottom: u32,
    right: u32,
    raw: Option<&'a [u8]>,
}

impl<'a> RangeFact<'a> {
    fn new(host: &FormulaHost<'a>, range: Range, raw: Option<&'a [u8]>) -> Self {
        Self {
            target: range.target_owner,
            host_row: host.row,
            host_column: host.column,
            top: range.top,
            left: range.left,
            bottom: range.bottom,
            right: range.right,
            raw,
        }
    }

    const fn sort_key(self) -> (u32, u32, u32, u32, u32, u32, u32) {
        (
            self.target,
            self.host_row,
            self.host_column,
            self.top,
            self.left,
            self.bottom,
            self.right,
        )
    }
}

impl Preflight {
    fn new(graph: CompleteGraph<'_, '_>, limits: Limits) -> Result<Self, Error> {
        ensure_limit("hosts", graph.hosts.len(), limits.max_hosts)?;
        ensure_limit("messages", 1, limits.max_messages)?;
        if graph.owners.is_empty() {
            return Err(Error::InvalidGraph);
        }
        let messages = graph.owners.iter().try_fold(1usize, |sum, owner| {
            checked_add(
                sum,
                checked_add(
                    1,
                    checked_add(owner.cell_tiles.len(), owner.range_tiles.len())?,
                )?,
            )
        })?;
        ensure_limit("messages", messages, limits.max_messages)?;
        let mut source_bytes = graph.engine.payload.len();
        let mut cell_tiles = 0usize;
        let mut range_tiles = 0usize;
        let mut object_ids = try_vec(messages)?;
        object_ids.push(graph.engine.object_id);
        for owner in graph.owners {
            if owner.message.object_id == 0 {
                return Err(Error::InvalidGraph);
            }
            object_ids.push(owner.message.object_id);
            source_bytes = checked_add(source_bytes, owner.message.payload.len())?;
            source_bytes = owner
                .cell_tiles
                .iter()
                .try_fold(source_bytes, |sum, tile| {
                    checked_add(sum, tile.message.payload.len())
                })?;
            source_bytes = owner
                .range_tiles
                .iter()
                .try_fold(source_bytes, |sum, tile| {
                    checked_add(sum, tile.message.payload.len())
                })?;
            cell_tiles = checked_add(cell_tiles, owner.cell_tiles.len())?;
            range_tiles = checked_add(range_tiles, owner.range_tiles.len())?;
            for tile in owner.cell_tiles {
                if tile.message.object_id == 0 {
                    return Err(Error::InvalidGraph);
                }
                object_ids.push(tile.message.object_id);
            }
            for tile in owner.range_tiles {
                if tile.message.object_id == 0 {
                    return Err(Error::InvalidGraph);
                }
                object_ids.push(tile.message.object_id);
            }
        }
        if graph.engine.object_id == 0 {
            return Err(Error::InvalidGraph);
        }
        object_ids.sort_unstable();
        if object_ids.windows(2).any(|pair| pair[0] == pair[1]) {
            return Err(Error::InvalidGraph);
        }
        ensure_limit("source bytes", source_bytes, limits.max_source_bytes)?;

        let mut precedents = 0usize;
        let mut ranges = 0usize;
        let mut previous_host = None;
        for host in graph.hosts {
            let key = (host.owner, host.row, host.column);
            if previous_host.is_some_and(|previous| previous >= key)
                || !owner_exists(graph.owners, host.owner)
            {
                return Err(Error::InvalidGraph);
            }
            previous_host = Some(key);
            let mut previous_precedent = None;
            for precedent in host.precedents {
                let key = (precedent.target_owner, precedent.row, precedent.column);
                if previous_precedent.is_some_and(|previous| previous >= key)
                    || !owner_exists(graph.owners, precedent.target_owner)
                {
                    return Err(Error::InvalidGraph);
                }
                previous_precedent = Some(key);
            }
            let mut previous_range = None;
            for range in host.ranges {
                let key = (
                    range.target_owner,
                    range.top,
                    range.left,
                    range.bottom,
                    range.right,
                );
                if range.top > range.bottom
                    || range.left > range.right
                    || previous_range.is_some_and(|previous| previous >= key)
                    || !owner_exists(graph.owners, range.target_owner)
                {
                    return Err(Error::InvalidGraph);
                }
                previous_range = Some(key);
            }
            precedents = checked_add(precedents, host.precedents.len())?;
            ranges = checked_add(ranges, host.ranges.len())?;
        }
        ensure_limit("precedents", precedents, limits.max_precedents)?;
        ensure_limit("ranges", ranges, limits.max_ranges)?;

        let encoded_upper = checked_add(
            64,
            checked_add(
                checked_mul(graph.hosts.len(), 64)?,
                checked_add(checked_mul(precedents, 40)?, checked_mul(ranges, 96)?)?,
            )?,
        )?;
        let output_upper = checked_add(source_bytes, encoded_upper)?;
        ensure_limit(
            "output preauthorization",
            output_upper,
            limits.max_output_bytes,
        )?;
        let strict_upper = checked_mul(
            checked_add(source_bytes, output_upper)?,
            STRICT_WORK_MULTIPLIER,
        )?;
        ensure_limit(
            "strict work preauthorization",
            strict_upper,
            limits.max_work_bytes,
        )?;
        let field_upper = checked_mul(checked_add(source_bytes, output_upper)?, 4)?;
        ensure_limit("field preauthorization", field_upper, limits.max_fields)?;
        let reference_upper = checked_mul(
            checked_add(graph.owners.len(), checked_add(cell_tiles, range_tiles)?)?,
            4,
        )?;
        ensure_limit(
            "reference preauthorization",
            reference_upper,
            limits.max_references,
        )?;
        let indexed_items = checked_add(messages, ranges)?;
        let sort_factor =
            usize::try_from(usize::BITS.saturating_sub(indexed_items.leading_zeros()))
                .map_err(|_error| Error::InvalidGraph)?;
        let membership_work = checked_mul(
            checked_add(graph.hosts.len(), checked_add(precedents, ranges)?)?,
            graph.owners.len(),
        )?;
        let graph_work = checked_add(
            checked_add(source_bytes, output_upper)?,
            checked_add(
                checked_mul(indexed_items, sort_factor)?,
                checked_add(checked_mul(ranges, 12)?, membership_work)?,
            )?,
        )?;
        ensure_limit(
            "graph work preauthorization",
            graph_work,
            limits.max_work_bytes,
        )?;
        let allocations = checked_add(
            64,
            checked_add(
                checked_mul(messages, 12)?,
                checked_add(checked_mul(graph.hosts.len(), 10)?, checked_mul(ranges, 6)?)?,
            )?,
        )?;
        ensure_limit("allocation events", allocations, limits.max_allocations)?;
        let peak_scratch_bytes = checked_add(
            checked_add(
                encoded_upper,
                checked_mul(checked_add(precedents, ranges)?, size_of::<u32>() * 5)?,
            )?,
            checked_add(
                checked_mul(checked_mul(limits.max_references, 6)?, size_of::<u64>())?,
                checked_add(
                    checked_mul(messages, size_of::<u64>())?,
                    checked_mul(ranges, size_of::<RangeFact<'static>>())?,
                )?,
            )?,
        )?;
        ensure_limit("peak scratch", peak_scratch_bytes, limits.max_scratch_bytes)?;
        let reference_count =
            graph
                .owners
                .iter()
                .try_fold(graph.engine.object_references.len(), |sum, owner| {
                    let cell_references =
                        owner.cell_tiles.iter().try_fold(0usize, |tile_sum, tile| {
                            checked_add(tile_sum, tile.message.object_references.len())
                        })?;
                    let range_references =
                        owner
                            .range_tiles
                            .iter()
                            .try_fold(0usize, |tile_sum, tile| {
                                checked_add(tile_sum, tile.message.object_references.len())
                            })?;
                    checked_add(
                        sum,
                        checked_add(
                            owner.message.object_references.len(),
                            checked_add(cell_references, range_references)?,
                        )?,
                    )
                })?;
        // Engine MessageEdit is stored inline in Artifact; owner and tile
        // edits are retained in three outer Vec buffers whose exact combined
        // capacity is every message except the engine.
        let outer_edits = messages.checked_sub(1).ok_or(Error::InvalidGraph)?;
        let retained_upper = checked_add(
            checked_add(
                output_upper,
                checked_mul(reference_count, size_of::<u64>())?,
            )?,
            checked_mul(outer_edits, size_of::<MessageEdit>())?,
        )?;
        ensure_limit(
            "retained preauthorization",
            retained_upper,
            limits.max_retained_bytes,
        )?;
        Ok(Self {
            object_ids,
            source_bytes,
            output_upper,
            field_upper,
            strict_work_upper: strict_upper,
            reference_upper,
            retained_upper,
            messages,
            hosts: graph.hosts.len(),
            precedents,
            ranges,
            cell_tiles,
            range_tiles,
            allocations,
            peak_scratch_bytes,
            graph_work_bytes: graph_work,
        })
    }
}

fn owner_exists(owners: &[SourceOwner<'_, '_>], internal_owner: u32) -> bool {
    owners
        .iter()
        .any(|owner| owner.internal_owner == internal_owner)
}

fn wire_limits(limits: Limits) -> Result<WireLimits, Error> {
    WireLimits::default()
        .with_input_bytes(
            limits
                .max_source_bytes
                .max(limits.max_output_bytes)
                .min(WireLimits::MAX_INPUT_BYTES),
        )
        .and_then(|value| {
            value.with_output_bytes(limits.max_output_bytes.min(WireLimits::MAX_OUTPUT_BYTES))
        })
        .and_then(|value| value.with_fields(limits.max_fields.min(WireLimits::MAX_FIELDS)))
        .and_then(|value| {
            value.with_nesting((limits.recursion_limit as usize).min(WireLimits::MAX_NESTING))
        })
        .and_then(|value| {
            value.with_rewrite_work(limits.max_work_bytes.min(WireLimits::MAX_REWRITE_WORK))
        })
        .map_err(|_error| Error::Wire)
}

fn try_vec<T>(capacity: usize) -> Result<Vec<T>, Error> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|_error| Error::Allocation {
            requested: capacity,
        })?;
    if output.capacity() != capacity {
        return Err(Error::Allocation {
            requested: capacity,
        });
    }
    Ok(output)
}

fn reserve_exact_capacity<T>(values: &mut Vec<T>, additional: usize) -> Result<(), Error> {
    let required = checked_add(values.len(), additional)?;
    if required <= values.capacity() {
        return Ok(());
    }
    values
        .try_reserve_exact(additional)
        .map_err(|_error| Error::Allocation {
            requested: additional,
        })?;
    if values.capacity() != required {
        return Err(Error::Allocation {
            requested: additional,
        });
    }
    Ok(())
}

fn try_vec_bytes(capacity: usize) -> Result<Vec<u8>, Error> {
    try_vec(capacity)
}

fn checked_add(left: usize, right: usize) -> Result<usize, Error> {
    left.checked_add(right).ok_or(Error::InvalidGraph)
}

fn checked_mul(left: usize, right: usize) -> Result<usize, Error> {
    left.checked_mul(right).ok_or(Error::InvalidGraph)
}

fn ensure_limit(resource: &'static str, observed: usize, maximum: usize) -> Result<(), Error> {
    if observed > maximum {
        return Err(Error::Limit {
            resource,
            observed,
            maximum,
        });
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use litchi_iwa_protos::numbers_formula_codec::{
        self as formula, BinaryOperator, DecodeError, FormulaWriteCellReference,
        FormulaWriteContext, FormulaWriteDependencyLimits, FormulaWriteDependencyVisitor,
        FormulaWriteNode, FormulaWriteOwnerUid, FormulaWritePrecedent, FormulaWriteRange,
        ResolvedFormulaWriteOwner,
    };

    use super::*;

    #[derive(Default)]
    struct WriteFacts {
        precedents: Vec<FormulaWritePrecedent>,
        ranges: Vec<FormulaWriteRange>,
    }

    impl FormulaWriteDependencyVisitor for WriteFacts {
        fn visit_precedent(&mut self, value: FormulaWritePrecedent) -> Result<(), DecodeError> {
            self.precedents.push(value);
            Ok(())
        }

        fn visit_range(&mut self, value: FormulaWriteRange) -> Result<(), DecodeError> {
            self.ranges.push(value);
            Ok(())
        }
    }

    #[test]
    fn inline_graph_preserves_unknowns_and_writes_external_triplets_and_ranges() {
        let uid = FormulaWriteOwnerUid::from_halves(3, 4);
        let nodes = [
            FormulaWriteNode::ResolvedCellReference {
                owner: None,
                reference: FormulaWriteCellReference::new(1, 2, true, true),
            },
            FormulaWriteNode::ResolvedRange {
                owner: Some(uid),
                start: FormulaWriteCellReference::new(2, 3, true, true),
                end: FormulaWriteCellReference::new(3, 4, true, true),
            },
            FormulaWriteNode::Binary(BinaryOperator::Add),
        ];
        let options = formula::DecodeOptions::new(1 << 20, 10_000, 1 << 20, 16, 10_000, 1 << 20);
        let owners = [ResolvedFormulaWriteOwner::new(uid, 20, 100, 100)];
        let plan = formula::plan_resolved_formula_archive(
            &nodes,
            FormulaWriteContext::new(10, 5, 6, 100, 100),
            &owners,
            FormulaWriteDependencyLimits::new(100, 10),
            options,
        )
        .unwrap();
        let mut facts = WriteFacts::default();
        formula::execute_formula_archive_plan_with_visitor(plan, options, &mut facts).unwrap();
        facts
            .precedents
            .sort_by_key(|value| (value.internal_owner(), value.row(), value.column()));
        facts.precedents.dedup();
        facts.ranges.sort_by_key(|value| {
            (
                value.internal_owner(),
                value.top(),
                value.left(),
                value.bottom(),
                value.right(),
            )
        });

        let mut engine = Vec::new();
        let mut tracker = Vec::new();
        append_varint(&mut tracker, 5, 0);
        for object in [101_u64, 102] {
            let mut reference = Vec::new();
            append_varint(&mut reference, 1, object);
            append_message(&mut tracker, 6, &reference).unwrap();
        }
        append_message(&mut engine, 2, &tracker).unwrap();
        append_varint(&mut engine, 99, 7);
        let owner_10 = owner_payload(10, 91);
        let owner_20 = owner_payload(20, 92);
        let precedents: Vec<_> = facts
            .precedents
            .iter()
            .map(|value| Precedent {
                target_owner: value.internal_owner(),
                row: value.row(),
                column: value.column(),
            })
            .collect();
        let ranges: Vec<_> = facts
            .ranges
            .iter()
            .map(|value| Range {
                target_owner: value.internal_owner(),
                top: value.top(),
                left: value.left(),
                bottom: value.bottom(),
                right: value.right(),
            })
            .collect();
        let host = FormulaHost::authored(10, 5, 6, &precedents, &ranges);
        let source_owners = [
            SourceOwner {
                message: SourceMessage {
                    object_id: 101,
                    payload: &owner_10,
                    object_references: &[],
                },
                internal_owner: 10,
                uid_lower: 10,
                uid_upper: 11,
                cell_tiles: &[],
                range_tiles: &[],
            },
            SourceOwner {
                message: SourceMessage {
                    object_id: 102,
                    payload: &owner_20,
                    object_references: &[],
                },
                internal_owner: 20,
                uid_lower: 20,
                uid_upper: 21,
                cell_tiles: &[],
                range_tiles: &[],
            },
        ];
        let changes = [HostChange {
            old: None,
            new: Some(host),
            cell_tile_object_id: None,
        }];
        let graph = SourceGraph {
            engine: SourceMessage {
                object_id: 100,
                payload: &engine,
                object_references: &[777, 101, 888, 102],
            },
            owners: &source_owners,
            table_owner: source_owners[0].internal_owner,
            existing_formula_hosts: &[],
            changes: &changes,
        };
        let prepared = prepare_graph(graph, generous_limits()).unwrap();
        assert_eq!(prepared.prepare_report().output_bytes, 0);
        assert_eq!(prepared.logical_view().formula_count, 1);
        assert_eq!(prepared.logical_view().hosts.len(), 1);
        assert_eq!(prepared.logical_view().hosts[0].precedents, precedents);
        assert_eq!(prepared.logical_view().hosts[0].ranges, ranges);
        let requirements = prepared.execution_requirements();
        assert!(requirements.output_bytes != 0);
        assert!(requirements.retained_bytes != prepared.prepare_report().retained_bytes);
        // Cache/cycle refusal is represented by dropping the governed plan:
        // no candidate payload or MessageEdit has been allocated.
        drop(prepared);

        let prepared = prepare_graph(graph, generous_limits()).unwrap();
        let mut refused = generous_limits();
        refused.max_output_bytes = requirements.output_bytes - 1;
        assert!(matches!(
            prepared.execute(refused),
            Err(Error::Limit {
                resource: "execution output bytes",
                observed,
                maximum,
            }) if observed == requirements.output_bytes && maximum + 1 == observed
        ));

        let prepared = prepare_graph(graph, generous_limits()).unwrap();
        let mut refused = generous_limits();
        refused.max_retained_bytes = requirements.retained_bytes - 1;
        assert!(matches!(
            prepared.execute(refused),
            Err(Error::Limit {
                resource: "execution retained bytes",
                observed,
                maximum,
            }) if observed == requirements.retained_bytes && maximum + 1 == observed
        ));

        for (resource, observed, lower) in [
            ("execution hosts", requirements.hosts, 0usize),
            (
                "execution precedents",
                requirements.precedents,
                requirements.hosts,
            ),
            ("execution ranges", requirements.ranges, requirements.hosts),
        ] {
            if observed == 0 {
                continue;
            }
            let prepared = prepare_graph(graph, generous_limits()).unwrap();
            let mut refused = generous_limits();
            match resource {
                "execution hosts" => refused.max_hosts = observed - 1,
                "execution precedents" => refused.max_precedents = observed - 1,
                "execution ranges" => refused.max_ranges = observed - 1,
                _ => unreachable!(),
            }
            assert!(matches!(
                prepared.execute(refused),
                Err(Error::Limit {
                    resource: actual,
                    observed: actual_observed,
                    maximum,
                }) if actual == resource
                    && actual_observed == observed
                    && maximum + 1 == actual_observed
                    && actual_observed >= lower
            ));
        }

        let artifact = rewrite_graph(graph, generous_limits()).unwrap();
        let engine_output = artifact.engine.payload.as_deref().unwrap();
        assert!(engine_output.ends_with(&[0x98, 0x06, 0x07]));
        let owner_output = artifact.owners[0].payload.as_deref().unwrap();
        assert!(
            owner_output
                .windows(3)
                .any(|window| window == [0xd8, 0x05, 0x5b])
        );
        let expected_hosts = [host];
        let mut visitor = FactsVisitor::expect(&expected_hosts);
        dependency::decode_formula_owner_dependencies_with_visitor(
            owner_output,
            dependency_options(owner_output.len()),
            &mut visitor,
        )
        .unwrap();
        assert!(visitor.cells_complete());
        assert!(visitor.ranges_complete());
        assert_eq!(artifact.engine.object_references, [777, 101, 888, 102]);
    }

    #[test]
    fn new_tile_dually_publishes_exact_inline_record_and_preauthorizes_output() {
        let mut tracker = Vec::new();
        append_varint(&mut tracker, 5, 0);
        let mut owner_reference = Vec::new();
        append_varint(&mut owner_reference, 1, 101);
        append_message(&mut tracker, 6, &owner_reference).unwrap();
        let mut engine = Vec::new();
        append_message(&mut engine, 2, &tracker).unwrap();
        let owner_payload = owner_payload(10, 91);
        let new_tile = [CellTileSource {
            message: SourceMessage {
                object_id: 202,
                payload: &[],
                object_references: &[],
            },
            source_present: false,
            tile_column_begin: 0,
            tile_row_begin: 0,
        }];
        let owners = [SourceOwner {
            message: SourceMessage {
                object_id: 101,
                payload: &owner_payload,
                object_references: &[],
            },
            internal_owner: 10,
            uid_lower: 10,
            uid_upper: 11,
            cell_tiles: &new_tile,
            range_tiles: &[],
        }];
        let precedents = [Precedent {
            target_owner: 10,
            row: 2,
            column: 1,
        }];
        let host = FormulaHost::authored(10, 2, 2, &precedents, &[]);
        let changes = [HostChange {
            old: None,
            new: Some(host),
            cell_tile_object_id: Some(202),
        }];
        let graph = SourceGraph {
            engine: SourceMessage {
                object_id: 100,
                payload: &engine,
                object_references: &[101],
            },
            owners: &owners,
            table_owner: 10,
            existing_formula_hosts: &[],
            changes: &changes,
        };
        let prepared = prepare_graph(graph, generous_limits()).unwrap();
        assert_eq!(prepared.prepare_report().output_bytes, 0);
        let requirements = prepared.execution_requirements();
        let mut refused = generous_limits();
        refused.max_output_bytes = requirements.output_bytes - 1;
        assert!(matches!(
            prepared.execute(refused),
            Err(Error::Limit {
                resource: "execution output bytes",
                observed,
                maximum,
            }) if observed == requirements.output_bytes && maximum + 1 == observed
        ));

        let artifact = prepare_graph(graph, generous_limits())
            .unwrap()
            .execute(generous_limits())
            .unwrap();
        assert_eq!(artifact.owners[0].object_references, [202]);
        let owner_candidate = artifact.owners[0].payload.as_deref().unwrap();
        let tile_candidate = artifact.cell_tiles[0].payload.as_deref().unwrap();
        let expected = [host];
        let mut owner_facts = FactsVisitor::expect(&expected);
        let (_owner, _) = dependency::decode_formula_owner_dependencies_with_visitor(
            owner_candidate,
            dependency_options(owner_candidate.len()),
            &mut owner_facts,
        )
        .unwrap();
        assert!(owner_facts.cells_complete());
        let mut tile_facts = FactsVisitor::expect(&expected);
        let (tile, _) = dependency::decode_cell_record_tile_with_visitor(
            tile_candidate,
            dependency_options(tile_candidate.len()),
            &mut tile_facts,
        )
        .unwrap();
        assert!(tile_facts.complete());
        assert_eq!(tile.internal_owner_id(), 10);
        assert_eq!(tile.tile_column_begin(), 0);
        assert_eq!(tile.tile_row_begin(), 0);
        let owner_view =
            WireView::parse_with_limits(owner_candidate, WireLimits::default()).unwrap();
        let inline = owner_view
            .fields()
            .find(|field| field.number() == 4)
            .unwrap()
            .payload();
        let inline_view = WireView::parse_with_limits(inline, WireLimits::default()).unwrap();
        let inline_record = inline_view
            .fields()
            .find(|field| field.number() == 1)
            .unwrap()
            .payload();
        let tile_view = WireView::parse_with_limits(tile_candidate, WireLimits::default()).unwrap();
        let tiled_record = tile_view
            .fields()
            .find(|field| field.number() == 4)
            .unwrap()
            .payload();
        assert_eq!(inline_record, tiled_record);

        // The emitted dual representation is itself a strict admissible
        // source: exactly one inline record and one assigned tile mirror.
        let roundtrip_owner_payload = artifact.owners[0].payload.as_deref().unwrap();
        let roundtrip_owner_refs = artifact.owners[0].object_references.as_slice();
        let roundtrip_tile_payload = artifact.cell_tiles[0].payload.as_deref().unwrap();
        let roundtrip_tiles = [CellTileSource {
            message: SourceMessage {
                object_id: 202,
                payload: roundtrip_tile_payload,
                object_references: &[],
            },
            source_present: true,
            tile_column_begin: 0,
            tile_row_begin: 0,
        }];
        let roundtrip_owners = [SourceOwner {
            message: SourceMessage {
                object_id: 101,
                payload: roundtrip_owner_payload,
                object_references: roundtrip_owner_refs,
            },
            internal_owner: 10,
            uid_lower: 10,
            uid_upper: 11,
            cell_tiles: &roundtrip_tiles,
            range_tiles: &[],
        }];
        let existing = [HostKey {
            owner: 10,
            row: 2,
            column: 2,
        }];
        let roundtrip = prepare_graph(
            SourceGraph {
                engine: SourceMessage {
                    object_id: 100,
                    payload: artifact.engine.payload.as_deref().unwrap(),
                    object_references: &[101],
                },
                owners: &roundtrip_owners,
                table_owner: 10,
                existing_formula_hosts: &existing,
                changes: &[],
            },
            generous_limits(),
        )
        .unwrap();
        assert_eq!(roundtrip.logical_view().formula_count, 1);
        assert_eq!(roundtrip.logical_view().hosts.len(), 1);
    }

    #[test]
    fn opaque_mirrored_cell_is_preserved_and_divergence_refuses() {
        let mut tracker = Vec::new();
        append_varint(&mut tracker, 5, 1);
        for object in [101_u64, 102] {
            let mut reference = Vec::new();
            append_varint(&mut reference, 1, object);
            append_message(&mut tracker, 6, &reference).unwrap();
        }
        let mut engine = Vec::new();
        append_message(&mut engine, 2, &tracker).unwrap();

        let selected_payload = owner_payload(10, 91);
        let opaque_host = FormulaHost::authored(20, 0, 0, &[], &[]);
        let mut mirrored_record = encode_cell_record(&opaque_host, generous_limits()).unwrap();
        append_varint(&mut mirrored_record, 99, 7);
        let mut inline_cells = Vec::new();
        append_message(&mut inline_cells, 1, &mirrored_record).unwrap();
        let mut opaque_payload = owner_payload(20, 92);
        append_message(&mut opaque_payload, 4, &inline_cells).unwrap();
        let mut tile_reference = Vec::new();
        append_varint(&mut tile_reference, 1, 202);
        let mut tiled_cells = Vec::new();
        append_message(&mut tiled_cells, 1, &tile_reference).unwrap();
        append_message(&mut opaque_payload, 13, &tiled_cells).unwrap();
        let mut tile_payload = Vec::new();
        append_varint(&mut tile_payload, 1, 20);
        append_varint(&mut tile_payload, 2, 0);
        append_varint(&mut tile_payload, 3, 0);
        append_message(&mut tile_payload, 4, &mirrored_record).unwrap();

        let opaque_tile_refs = [];
        let cell_tiles = [CellTileSource {
            message: SourceMessage {
                object_id: 202,
                payload: &tile_payload,
                object_references: &opaque_tile_refs,
            },
            source_present: true,
            tile_column_begin: 0,
            tile_row_begin: 0,
        }];
        // ArchiveInfo may also retain object references owned by fields this
        // writer does not edit (for example the table/drawable bound to a
        // formula owner).  The dependency-tile references remain an exact,
        // ordered subsequence and every extra reference is preserved raw.
        let owner_refs = [777, 202, 888];
        let owners = [
            SourceOwner {
                message: SourceMessage {
                    object_id: 101,
                    payload: &selected_payload,
                    object_references: &[],
                },
                internal_owner: 10,
                uid_lower: 10,
                uid_upper: 11,
                cell_tiles: &[],
                range_tiles: &[],
            },
            SourceOwner {
                message: SourceMessage {
                    object_id: 102,
                    payload: &opaque_payload,
                    object_references: &owner_refs,
                },
                internal_owner: 20,
                uid_lower: 20,
                uid_upper: 21,
                cell_tiles: &cell_tiles,
                range_tiles: &[],
            },
        ];
        let selected_precedents = [Precedent {
            target_owner: 10,
            row: 0,
            column: 0,
        }];
        let selected_host = FormulaHost::authored(10, 2, 2, &selected_precedents, &[]);
        let changes = [HostChange {
            old: None,
            new: Some(selected_host),
            cell_tile_object_id: None,
        }];
        let engine_refs = [101, 102];
        let graph = SourceGraph {
            engine: SourceMessage {
                object_id: 100,
                payload: &engine,
                object_references: &engine_refs,
            },
            owners: &owners,
            table_owner: 10,
            existing_formula_hosts: &[],
            changes: &changes,
        };
        let prepared = prepare_graph(graph, generous_limits()).unwrap();
        assert_eq!(prepared.logical_view().formula_count, 1);
        assert_eq!(prepared.logical_view().hosts.len(), 1);
        assert_eq!(prepared.logical_view().hosts[0].owner, 10);
        let artifact = prepared.execute(generous_limits()).unwrap();
        assert_eq!(artifact.owners[1].payload, None);
        assert_eq!(artifact.owners[1].object_references, owner_refs);
        assert_eq!(artifact.cell_tiles[0].payload, None);
        assert_eq!(artifact.cell_tiles[0].object_references, opaque_tile_refs);

        let divergent_precedents = [Precedent {
            target_owner: 20,
            row: 1,
            column: 1,
        }];
        let divergent_host = FormulaHost::authored(20, 0, 0, &divergent_precedents, &[]);
        let divergent_record = encode_cell_record(&divergent_host, generous_limits()).unwrap();
        let mut divergent_tile = Vec::new();
        append_varint(&mut divergent_tile, 1, 20);
        append_varint(&mut divergent_tile, 2, 0);
        append_varint(&mut divergent_tile, 3, 0);
        append_message(&mut divergent_tile, 4, &divergent_record).unwrap();
        let bad_cell_tiles = [CellTileSource {
            message: SourceMessage {
                payload: &divergent_tile,
                ..cell_tiles[0].message
            },
            ..cell_tiles[0]
        }];
        let mut bad_owners = owners;
        bad_owners[1].cell_tiles = &bad_cell_tiles;
        let bad_graph = SourceGraph {
            owners: &bad_owners,
            ..graph
        };
        let error = prepare_graph(bad_graph, generous_limits()).err();
        assert!(
            matches!(error, Some(Error::InvalidGraph)),
            "error={error:?}"
        );

        let mut unknown_divergent_record = mirrored_record.clone();
        unknown_divergent_record.pop();
        unknown_divergent_record.push(8);
        let mut unknown_divergent_tile = Vec::new();
        append_varint(&mut unknown_divergent_tile, 1, 20);
        append_varint(&mut unknown_divergent_tile, 2, 0);
        append_varint(&mut unknown_divergent_tile, 3, 0);
        append_message(&mut unknown_divergent_tile, 4, &unknown_divergent_record).unwrap();
        let unknown_bad_tiles = [CellTileSource {
            message: SourceMessage {
                payload: &unknown_divergent_tile,
                ..cell_tiles[0].message
            },
            ..cell_tiles[0]
        }];
        let mut unknown_bad_owners = owners;
        unknown_bad_owners[1].cell_tiles = &unknown_bad_tiles;
        assert!(matches!(
            prepare_graph(
                SourceGraph {
                    owners: &unknown_bad_owners,
                    ..graph
                },
                generous_limits()
            ),
            Err(Error::InvalidGraph)
        ));
    }

    #[test]
    fn repeated_splice_retains_all_unselected_raw_fields() {
        let source = [0x08, 0x01, 0x22, 0x02, 0x08, 0x01, 0x98, 0x06, 0x07];
        let replacement = vec![vec![0x08, 0x02]];
        let output = replace_repeated(&source, 4, &replacement, WireLimits::default()).unwrap();
        assert_eq!(
            output,
            [0x08, 0x01, 0x22, 0x02, 0x08, 0x02, 0x98, 0x06, 0x07]
        );
    }

    #[test]
    fn execution_retained_outer_edits_are_preauthorized_exactly() {
        let mut tracker = Vec::new();
        append_varint(&mut tracker, 5, 0);
        let mut reference = Vec::new();
        append_varint(&mut reference, 1, 101);
        append_message(&mut tracker, 6, &reference).unwrap();
        let mut engine = Vec::new();
        append_message(&mut engine, 2, &tracker).unwrap();
        let owner = owner_payload(10, 91);
        let owners = [SourceOwner {
            message: SourceMessage {
                object_id: 101,
                payload: &owner,
                object_references: &[],
            },
            internal_owner: 10,
            uid_lower: 10,
            uid_upper: 11,
            cell_tiles: &[],
            range_tiles: &[],
        }];
        let graph = SourceGraph {
            engine: SourceMessage {
                object_id: 100,
                payload: &engine,
                object_references: &[101],
            },
            owners: &owners,
            table_owner: owners[0].internal_owner,
            existing_formula_hosts: &[],
            changes: &[],
        };
        let prepared = prepare_graph(graph, generous_limits()).unwrap();
        let requirements = prepared.execution_requirements();
        assert!(requirements.retained_bytes >= size_of::<MessageEdit>());
        let mut refused = generous_limits();
        refused.max_retained_bytes = requirements.retained_bytes - 1;
        assert!(matches!(
            prepared.execute(refused),
            Err(Error::Limit {
                resource: "execution retained bytes",
                observed,
                maximum,
            }) if observed == requirements.retained_bytes && maximum + 1 == observed
        ));
    }

    #[test]
    fn try_vec_capacity_is_exact_and_one_less_limit_refuses() {
        for amount in [0, 1, 4_096, 8_192] {
            let vector = try_vec::<u64>(amount).unwrap();
            assert_eq!(vector.capacity(), amount);
        }
        assert!(matches!(
            ensure_limit("exact", 8_192, 8_191),
            Err(Error::Limit {
                observed: 8_192,
                maximum: 8_191,
                ..
            })
        ));
    }

    #[test]
    fn owner_uid_index_scales_and_rejects_duplicates() {
        let empty = SourceMessage {
            object_id: 0,
            payload: &[],
            object_references: &[],
        };
        for amount in [4_096usize, 8_192] {
            let owners: Vec<_> = (0..amount)
                .map(|index| SourceOwner {
                    message: SourceMessage {
                        object_id: index as u64 + 1,
                        ..empty
                    },
                    internal_owner: index as u32 + 1,
                    uid_lower: index as u64 + 10,
                    uid_upper: index as u64 + 20,
                    cell_tiles: &[],
                    range_tiles: &[],
                })
                .collect();
            let index = OwnerUidIndex::new(&owners).unwrap();
            assert_eq!(index.entries.len(), amount);
            assert_eq!(
                index.internal_owner(amount as u64 + 9, amount as u64 + 19),
                Some(amount as u32)
            );
        }
        let duplicate = [
            SourceOwner {
                message: SourceMessage {
                    object_id: 1,
                    ..empty
                },
                internal_owner: 1,
                uid_lower: 7,
                uid_upper: 8,
                cell_tiles: &[],
                range_tiles: &[],
            },
            SourceOwner {
                message: SourceMessage {
                    object_id: 2,
                    ..empty
                },
                internal_owner: 2,
                uid_lower: 7,
                uid_upper: 8,
                cell_tiles: &[],
                range_tiles: &[],
            },
        ];
        assert!(matches!(
            OwnerUidIndex::new(&duplicate),
            Err(Error::InvalidGraph)
        ));
    }

    #[test]
    fn source_admission_scales_and_uid_scratch_is_exact() {
        let empty = SourceMessage {
            object_id: 0,
            payload: &[],
            object_references: &[],
        };
        let mut work = [0usize; 2];
        for (slot, amount) in [4_096usize, 8_192].into_iter().enumerate() {
            let engine_refs: Vec<_> = (1..=amount as u64).collect();
            let owners: Vec<_> = (0..amount)
                .map(|index| SourceOwner {
                    message: SourceMessage {
                        object_id: index as u64 + 1,
                        ..empty
                    },
                    internal_owner: index as u32 + 1,
                    uid_lower: index as u64 + 10,
                    uid_upper: index as u64 + 20,
                    cell_tiles: &[],
                    range_tiles: &[],
                })
                .collect();
            let graph = SourceGraph {
                engine: SourceMessage {
                    object_references: &engine_refs,
                    ..empty
                },
                owners: &owners,
                table_owner: owners[0].internal_owner,
                existing_formula_hosts: &[],
                changes: &[],
            };
            let limits = scaling_limits(amount);
            let admitted = SourceAdmission::new(graph, limits).unwrap();
            work[slot] = admitted.graph_work_bytes;

            let exact = Limits {
                max_scratch_bytes: admitted.peak_scratch_bytes,
                ..limits
            };
            SourceAdmission::new(graph, exact).unwrap();
            let one_less = Limits {
                max_scratch_bytes: admitted.peak_scratch_bytes - 1,
                ..limits
            };
            assert!(matches!(
                SourceAdmission::new(graph, one_less),
                Err(Error::Limit {
                    resource: "source scratch preauthorization",
                    ..
                })
            ));
        }
        assert!(work[1] * 10 <= work[0] * 22, "work={work:?}");
    }

    #[test]
    fn range_only_sources_are_included_in_host_capacity() {
        let range_payload = [0u8; 100];
        let range_tiles = [RangeTileSource {
            message: SourceMessage {
                object_id: 2,
                payload: &range_payload,
                object_references: &[],
            },
            target_owner: 1,
        }];
        let owners = [SourceOwner {
            message: SourceMessage {
                object_id: 1,
                payload: &[],
                object_references: &[],
            },
            internal_owner: 1,
            uid_lower: 1,
            uid_upper: 2,
            cell_tiles: &[],
            range_tiles: &range_tiles,
        }];
        let graph = SourceGraph {
            engine: SourceMessage {
                object_id: 3,
                payload: &[],
                object_references: &[],
            },
            owners: &owners,
            table_owner: owners[0].internal_owner,
            existing_formula_hosts: &[],
            changes: &[],
        };
        assert_eq!(
            observed_source_host_bound(graph, scaling_limits(100)).unwrap(),
            52
        );
    }

    #[test]
    fn same_host_ranges_coalesce_once_and_scale_with_governed_sort() {
        let mut work = [0usize; 2];
        for (slot, amount) in [4_096usize, 8_192].into_iter().enumerate() {
            let mut hosts = try_vec(amount).unwrap();
            for index in 0..amount {
                let mut ranges = try_vec(1).unwrap();
                ranges.push(Range {
                    target_owner: 1,
                    top: index as u32,
                    left: 0,
                    bottom: index as u32,
                    right: 0,
                });
                let mut raw = try_vec(1).unwrap();
                raw.push(None);
                hosts.push(OwnedHost {
                    key: HostKey {
                        owner: 1,
                        row: 0,
                        column: 0,
                    },
                    precedents: Vec::new(),
                    ranges,
                    cell_tile_object_id: None,
                    has_cell_record: false,
                    is_in_cycle: false,
                    raw_cell_record: None,
                    raw_range_records: raw,
                });
            }
            coalesce_hosts(&mut hosts, 1, scaling_limits(amount)).unwrap();
            assert_eq!(hosts.len(), 1);
            assert_eq!(hosts[0].ranges.len(), amount);
            work[slot] = sort_work_upper(amount).unwrap();
        }
        assert!(work[1] * 10 <= work[0] * 22, "work={work:?}");
    }

    #[test]
    fn archive_info_references_require_exact_locality() {
        let empty = SourceMessage {
            object_id: 0,
            payload: &[],
            object_references: &[],
        };
        let cell_tiles = [CellTileSource {
            message: SourceMessage {
                object_id: 11,
                ..empty
            },
            source_present: true,
            tile_column_begin: 0,
            tile_row_begin: 0,
        }];
        let range_tiles = [RangeTileSource {
            message: SourceMessage {
                object_id: 12,
                ..empty
            },
            target_owner: 1,
        }];
        let good_refs = [11, 12];
        let mut owners = [SourceOwner {
            message: SourceMessage {
                object_id: 10,
                payload: &[],
                object_references: &good_refs,
            },
            internal_owner: 1,
            uid_lower: 1,
            uid_upper: 2,
            cell_tiles: &cell_tiles,
            range_tiles: &range_tiles,
        }];
        let engine_refs = [10];
        let engine = SourceMessage {
            object_id: 9,
            payload: &[],
            object_references: &engine_refs,
        };
        validate_archive_info_references(SourceGraph {
            engine,
            owners: &owners,
            table_owner: owners[0].internal_owner,
            existing_formula_hosts: &[],
            changes: &[],
        })
        .unwrap();

        for hostile in [&[77, 10, 88, 10][..], &[77, 88][..]] {
            assert!(matches!(
                validate_archive_info_references(SourceGraph {
                    engine: SourceMessage {
                        object_references: hostile,
                        ..engine
                    },
                    owners: &owners,
                    table_owner: owners[0].internal_owner,
                    existing_formula_hosts: &[],
                    changes: &[],
                }),
                Err(Error::InvalidGraph)
            ));
        }

        let owner_two_refs = [11, 12];
        let owner_two = SourceOwner {
            message: SourceMessage {
                object_id: 20,
                payload: &[],
                object_references: &owner_two_refs,
            },
            internal_owner: 2,
            uid_lower: 3,
            uid_upper: 4,
            cell_tiles: &cell_tiles,
            range_tiles: &range_tiles,
        };
        let two_owners = [owners[0], owner_two];
        for hostile in [&[20, 10][..], &[77, 20, 88, 10, 99][..]] {
            assert!(matches!(
                validate_archive_info_references(SourceGraph {
                    engine: SourceMessage {
                        object_references: hostile,
                        ..engine
                    },
                    owners: &two_owners,
                    table_owner: two_owners[0].internal_owner,
                    existing_formula_hosts: &[],
                    changes: &[],
                }),
                Err(Error::InvalidGraph)
            ));
        }
        validate_archive_info_references(SourceGraph {
            engine: SourceMessage {
                object_references: &[77, 10, 88, 20, 99],
                ..engine
            },
            owners: &two_owners,
            table_owner: two_owners[0].internal_owner,
            existing_formula_hosts: &[],
            changes: &[],
        })
        .unwrap();

        for bad in [&[][..], &[12][..], &[12, 11][..], &[11, 12, 11][..]] {
            owners[0].message.object_references = bad;
            assert!(matches!(
                validate_archive_info_references(SourceGraph {
                    engine,
                    owners: &owners,
                    table_owner: owners[0].internal_owner,
                    existing_formula_hosts: &[],
                    changes: &[],
                }),
                Err(Error::InvalidGraph)
            ));
        }
        owners[0].message.object_references = &[77, 11, 88, 12, 99];
        validate_archive_info_references(SourceGraph {
            engine,
            owners: &owners,
            table_owner: owners[0].internal_owner,
            existing_formula_hosts: &[],
            changes: &[],
        })
        .unwrap();
        owners[0].message.object_references = &good_refs;
        assert!(matches!(
            validate_archive_info_references(SourceGraph {
                engine: SourceMessage {
                    object_references: &[],
                    ..engine
                },
                owners: &owners,
                table_owner: owners[0].internal_owner,
                existing_formula_hosts: &[],
                changes: &[],
            }),
            Err(Error::InvalidGraph)
        ));
        let extra_tile_refs = [77];
        let bad_cells = [CellTileSource {
            message: SourceMessage {
                object_references: &extra_tile_refs,
                ..cell_tiles[0].message
            },
            ..cell_tiles[0]
        }];
        owners[0].cell_tiles = &bad_cells;
        assert!(matches!(
            validate_archive_info_references(SourceGraph {
                engine,
                owners: &owners,
                table_owner: owners[0].internal_owner,
                existing_formula_hosts: &[],
                changes: &[],
            }),
            Err(Error::InvalidGraph)
        ));
    }

    #[test]
    fn range_scratch_admission_is_exact_and_covers_phase_overlap() {
        let empty = SourceMessage {
            object_id: 0,
            payload: &[],
            object_references: &[],
        };
        let engine_refs = [1];
        let owners = [SourceOwner {
            message: SourceMessage {
                object_id: 1,
                ..empty
            },
            internal_owner: 1,
            uid_lower: 1,
            uid_upper: 2,
            cell_tiles: &[],
            range_tiles: &[],
        }];
        let graph = SourceGraph {
            engine: SourceMessage {
                object_references: &engine_refs,
                ..empty
            },
            owners: &owners,
            table_owner: owners[0].internal_owner,
            existing_formula_hosts: &[],
            changes: &[],
        };
        let limits = Limits {
            max_hosts: 1,
            max_ranges: 8_192,
            max_messages: 2,
            ..scaling_limits(8_192)
        };
        let admitted = SourceAdmission::new(graph, limits).unwrap();
        SourceAdmission::new(
            graph,
            Limits {
                max_scratch_bytes: admitted.peak_scratch_bytes,
                ..limits
            },
        )
        .unwrap();
        assert!(matches!(
            SourceAdmission::new(
                graph,
                Limits {
                    max_scratch_bytes: admitted.peak_scratch_bytes - 1,
                    ..limits
                }
            ),
            Err(Error::Limit {
                resource: "source scratch preauthorization",
                ..
            })
        ));
        let per_range = checked_add(
            checked_add(size_of::<Range>(), size_of::<Option<&'static [u8]>>()).unwrap(),
            checked_add(
                size_of::<(Range, Option<&'static [u8]>)>(),
                size_of::<(HostKey, Range, Option<&'static [u8]>)>(),
            )
            .unwrap(),
        )
        .unwrap();
        assert!(per_range >= 84);
    }

    fn owner_payload(internal_owner: u32, unknown: u64) -> Vec<u8> {
        let mut uuid = Vec::new();
        append_varint(&mut uuid, 1, u64::from(internal_owner));
        append_varint(&mut uuid, 2, u64::from(internal_owner) + 1);
        let mut owner = Vec::new();
        append_message(&mut owner, 1, &uuid).unwrap();
        append_varint(&mut owner, 2, u64::from(internal_owner));
        append_varint(&mut owner, 91, unknown);
        owner
    }

    const fn generous_limits() -> Limits {
        Limits {
            max_source_bytes: 1 << 20,
            max_output_bytes: 1 << 20,
            max_fields: 1_000_000,
            max_work_bytes: 16_000_000,
            max_references: 100,
            max_messages: 100,
            max_hosts: 100,
            max_precedents: 1_000,
            max_ranges: 100,
            max_retained_bytes: 2 << 20,
            max_scratch_bytes: 2 << 20,
            max_allocations: 10_000,
            recursion_limit: 16,
        }
    }

    const fn scaling_limits(amount: usize) -> Limits {
        Limits {
            max_source_bytes: 1 << 20,
            max_output_bytes: 1 << 20,
            max_fields: usize::MAX,
            max_work_bytes: usize::MAX,
            max_references: usize::MAX,
            max_messages: amount + 1,
            max_hosts: amount,
            max_precedents: amount,
            max_ranges: amount,
            max_retained_bytes: usize::MAX,
            max_scratch_bytes: usize::MAX,
            max_allocations: usize::MAX,
            recursion_limit: 16,
        }
    }

    fn dependency_options(bytes: usize) -> dependency::DecodeOptions {
        dependency::DecodeOptions::new(bytes, 10_000, bytes * 32, 16, 100, 1)
    }
}
