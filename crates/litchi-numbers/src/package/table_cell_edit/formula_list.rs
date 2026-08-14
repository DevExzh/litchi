//! Bounded exact formula-list planning for rooted and segmented lists.

use core::{fmt, mem::size_of};

use litchi_iwa_common::{
    WireLimits,
    varint::{encode_varint_into, encoded_len},
    wire::{WireView, parse_wire_fields_with_limits},
};
use litchi_iwa_protos::numbers_table_cell_storage_codec as storage;

const FORMULA_LIST_TYPE: i32 = 3;
const ENTRY_FIELD: u32 = 3;
const NEXT_KEY_FIELD: u32 = 2;
const KEY_FIELD: u32 = 1;
const REF_COUNT_FIELD: u32 = 2;
const FORMULA_FIELD: u32 = 5;

#[derive(Debug, Clone, Copy)]
pub(super) struct SourceMessage<'a> {
    pub(super) object_id: u64,
    pub(super) payload: &'a [u8],
    pub(super) object_references: &'a [u64],
}

#[derive(Debug, Clone, Copy)]
pub(super) struct SourceList<'a> {
    pub(super) root: SourceMessage<'a>,
    pub(super) segments: &'a [SourceMessage<'a>],
    pub(super) expected_entries: usize,
    /// Complete source formula hosts, sorted uniquely by coordinate.
    pub(super) source_hosts: &'a [SourceHost],
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct SourceHost {
    pub(super) row: u32,
    pub(super) column: u32,
    pub(super) key: u32,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct HostDelta<'a> {
    pub(super) row: u32,
    pub(super) column: u32,
    pub(super) old_formula_key: Option<u32>,
    pub(super) new_formula: Option<&'a [u8]>,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct Limits {
    pub(super) max_input_bytes: usize,
    pub(super) max_output_bytes: usize,
    pub(super) max_fields: usize,
    pub(super) max_work: usize,
    pub(super) max_entries: usize,
    pub(super) max_hosts: usize,
    pub(super) max_references: usize,
    pub(super) max_retained_elements: usize,
    pub(super) max_retained_bytes: usize,
    pub(super) max_scratch_bytes: usize,
    pub(super) max_allocations: usize,
}

/// Governed bounds admitted before any formula-list planning allocation.
///
/// Artifact sizes are settled after strict candidate reopen. Work, fields,
/// allocation events, and peak scratch are conservative transaction bounds.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct Report {
    pub(super) input_bytes: usize,
    pub(super) output_bytes: usize,
    pub(super) fields: usize,
    pub(super) work: usize,
    pub(super) references: usize,
    pub(super) entries: usize,
    pub(super) hosts: usize,
    pub(super) allocations: usize,
    pub(super) retained_bytes: usize,
    pub(super) peak_scratch_bytes: usize,
    pub(super) changed_messages: usize,
    pub(super) retained_elements: usize,
    pub(super) assignments: usize,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct MessageEdit {
    pub(super) object_id: u64,
    pub(super) payload: Option<Vec<u8>>,
    pub(super) object_references: Vec<u64>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Assignment {
    pub(super) delta: usize,
    pub(super) key: Option<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct FinalEntry {
    pub(super) key: u32,
    pub(super) ref_count: u32,
    pub(super) formula: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct Artifact {
    pub(super) assignments: Vec<Assignment>,
    pub(super) root: MessageEdit,
    pub(super) segments: Vec<MessageEdit>,
    pub(super) final_entries: Vec<FinalEntry>,
    pub(super) report: Report,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum Error {
    InvalidSource,
    Limit {
        resource: Resource,
        observed: usize,
        maximum: usize,
    },
    Allocation {
        requested: usize,
    },
    Wire,
    Strict,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Resource {
    InputBytes,
    OutputBytes,
    Fields,
    Work,
    Entries,
    Hosts,
    References,
    RetainedElements,
    RetainedBytes,
    ScratchBytes,
    Allocations,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("invalid or over-budget Numbers formula list")
    }
}

impl std::error::Error for Error {}

#[derive(Clone, Copy)]
struct Entry<'a> {
    key: u32,
    initial_count: u32,
    final_count: u32,
    formula: &'a [u8],
    raw: &'a [u8],
    owner: usize,
    source_index: usize,
}

#[derive(Clone, Copy)]
struct NewFormula<'a> {
    bytes: &'a [u8],
    key: u32,
    count: u32,
}

#[derive(Clone, Copy, Default)]
struct Requirements {
    input_bytes: usize,
    fields: usize,
    work: usize,
    execution_fields: usize,
    execution_work: usize,
    entries: usize,
    hosts: usize,
    formula_occurrences: usize,
    output_upper: usize,
    references: usize,
    scratch_bytes: usize,
    retained_upper: usize,
    allocations: usize,
    logical_retained_elements: usize,
    logical_retained_bytes: usize,
    logical_transient_bytes: usize,
    logical_scratch_bytes: usize,
    logical_allocations: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct PlanRequirements {
    pub(super) input_bytes: usize,
    pub(super) output_bytes: usize,
    pub(super) fields: usize,
    pub(super) work: usize,
    pub(super) references: usize,
    pub(super) retained_elements: usize,
    pub(super) retained_bytes: usize,
    pub(super) scratch_bytes: usize,
    pub(super) allocations: usize,
    pub(super) assignments: usize,
    pub(super) changed_messages: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct LogicalReport {
    pub(super) input_bytes: usize,
    pub(super) fields: usize,
    pub(super) work: usize,
    pub(super) references: usize,
    pub(super) output_bytes: usize,
    pub(super) retained_elements: usize,
    pub(super) retained_bytes: usize,
    pub(super) peak_scratch_bytes: usize,
    pub(super) allocations: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct ExecutionRequirements {
    pub(super) input_bytes: usize,
    pub(super) fields: usize,
    pub(super) work: usize,
    pub(super) references: usize,
    pub(super) output_bytes: usize,
    pub(super) retained_elements: usize,
    pub(super) retained_bytes: usize,
    pub(super) peak_scratch_bytes: usize,
    pub(super) allocations: usize,
    pub(super) assignments: usize,
    pub(super) changed_messages: usize,
}

#[derive(Clone, Copy)]
pub(super) struct PreparedPlan<'a> {
    source: SourceList<'a>,
    deltas: &'a [HostDelta<'a>],
    limits: Limits,
    requirements: Requirements,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct LogicalAssignment {
    pub(super) delta: usize,
    pub(super) row: u32,
    pub(super) column: u32,
    pub(super) key: Option<u32>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct LogicalEntry<'a> {
    pub(super) key: u32,
    pub(super) ref_count: u32,
    pub(super) formula: &'a [u8],
}

#[derive(Debug, Clone, Copy)]
pub(super) struct LogicalList<'a> {
    pub(super) assignments: &'a [LogicalAssignment],
    pub(super) entries: &'a [LogicalEntry<'a>],
}

pub(super) struct LogicalPlan<'a> {
    source: SourceList<'a>,
    requirements: Requirements,
    entries: Vec<Entry<'a>>,
    key_order: Vec<usize>,
    owner_order: Vec<usize>,
    new_formulas: Vec<NewFormula<'a>>,
    assignments: Vec<Assignment>,
    logical_assignments: Vec<LogicalAssignment>,
    logical_entries: Vec<LogicalEntry<'a>>,
    next_key: u32,
    report: LogicalReport,
    execution: ExecutionRequirements,
}

impl<'a> PreparedPlan<'a> {
    pub(super) fn requirements(self) -> Result<PlanRequirements, Error> {
        public_requirements(self.deltas, self.requirements)
    }

    #[cfg(test)]
    pub(super) fn execute(self) -> Result<Artifact, Error> {
        let limits = self.limits;
        self.logical()?.execute(limits)
    }

    pub(super) fn logical(self) -> Result<LogicalPlan<'a>, Error> {
        prepare_logical(self.source, self.deltas, self.limits, self.requirements)
    }
}

impl<'a> LogicalPlan<'a> {
    pub(super) fn logical_view(&self) -> LogicalList<'_> {
        LogicalList {
            assignments: &self.logical_assignments,
            entries: &self.logical_entries,
        }
    }

    pub(super) const fn prepare_report(&self) -> LogicalReport {
        self.report
    }

    pub(super) const fn execution_requirements(&self) -> ExecutionRequirements {
        self.execution
    }

    pub(super) fn execute(self, limits: Limits) -> Result<Artifact, Error> {
        self.execution.ensure(limits)?;
        execute_logical(self, limits)
    }
}

impl ExecutionRequirements {
    fn ensure(self, limits: Limits) -> Result<(), Error> {
        ensure(
            Resource::InputBytes,
            self.input_bytes,
            limits.max_input_bytes,
        )?;
        ensure(Resource::Fields, self.fields, limits.max_fields)?;
        ensure(Resource::Work, self.work, limits.max_work)?;
        ensure(Resource::References, self.references, limits.max_references)?;
        ensure(
            Resource::OutputBytes,
            self.output_bytes,
            limits.max_output_bytes,
        )?;
        ensure(
            Resource::RetainedElements,
            self.retained_elements,
            limits.max_retained_elements,
        )?;
        ensure(
            Resource::RetainedBytes,
            self.retained_bytes,
            limits.max_retained_bytes,
        )?;
        ensure(
            Resource::ScratchBytes,
            self.peak_scratch_bytes,
            limits.max_scratch_bytes,
        )?;
        ensure(
            Resource::Allocations,
            self.allocations,
            limits.max_allocations,
        )
    }
}

pub(super) fn preflight_work_upper(
    segments: usize,
    hosts: usize,
    deltas: usize,
) -> Result<usize, Error> {
    segments
        .checked_mul(8)
        .and_then(|work| work.checked_add(hosts.checked_mul(2)?))
        .and_then(|work| work.checked_add(deltas.checked_mul(8)?))
        .and_then(|work| work.checked_add(1))
        .ok_or(Error::InvalidSource)
}

pub(super) fn prepare<'a>(
    source: SourceList<'a>,
    deltas: &'a [HostDelta<'a>],
    limits: Limits,
) -> Result<PreparedPlan<'a>, Error> {
    let requirements = preflight(source, deltas, limits)?;
    admit(requirements, limits)?;
    Ok(PreparedPlan {
        source,
        deltas,
        limits,
        requirements,
    })
}

fn public_requirements(
    deltas: &[HostDelta<'_>],
    requirements: Requirements,
) -> Result<PlanRequirements, Error> {
    Ok(PlanRequirements {
        input_bytes: requirements.input_bytes,
        output_bytes: 0,
        fields: requirements.fields,
        work: requirements.work,
        references: requirements.references,
        retained_elements: requirements.logical_retained_elements,
        retained_bytes: requirements.logical_retained_bytes,
        scratch_bytes: requirements.logical_scratch_bytes,
        allocations: requirements.logical_allocations,
        assignments: deltas.len(),
        changed_messages: 0,
    })
}

#[cfg(test)]
fn plan(
    source: SourceList<'_>,
    deltas: &[HostDelta<'_>],
    limits: Limits,
) -> Result<Artifact, Error> {
    prepare(source, deltas, limits)?.execute()
}

fn preflight(
    source: SourceList<'_>,
    deltas: &[HostDelta<'_>],
    limits: Limits,
) -> Result<Requirements, Error> {
    ensure(
        Resource::References,
        source.segments.len(),
        limits.max_references,
    )?;
    ensure(
        Resource::Entries,
        source.expected_entries,
        limits.max_entries,
    )?;
    ensure(Resource::Hosts, source.source_hosts.len(), limits.max_hosts)?;
    ensure(Resource::Hosts, deltas.len(), limits.max_hosts)?;
    ensure(
        Resource::InputBytes,
        source.root.payload.len(),
        limits.max_input_bytes,
    )?;
    if source.root.object_id == 0
        || source.segments.iter().any(|segment| segment.object_id == 0)
        || source
            .segments
            .iter()
            .any(|segment| !segment.object_references.is_empty())
        || source
            .segments
            .windows(2)
            .any(|pair| pair[0].object_id >= pair[1].object_id)
        || source.root.object_references.len() != source.segments.len()
        || source
            .root
            .object_references
            .iter()
            .zip(source.segments)
            .any(|(reference, segment)| *reference != segment.object_id)
        || source
            .source_hosts
            .windows(2)
            .any(|pair| (pair[0].row, pair[0].column) >= (pair[1].row, pair[1].column))
        || deltas
            .windows(2)
            .any(|pair| (pair[0].row, pair[0].column) >= (pair[1].row, pair[1].column))
    {
        return Err(Error::InvalidSource);
    }
    let input_bytes =
        source
            .segments
            .iter()
            .try_fold(source.root.payload.len(), |total, segment| {
                total
                    .checked_add(segment.payload.len())
                    .ok_or(Error::InvalidSource)
            })?;
    ensure(Resource::InputBytes, input_bytes, limits.max_input_bytes)?;
    let messages = source
        .segments
        .len()
        .checked_add(1)
        .ok_or(Error::InvalidSource)?;
    // A canonical protobuf field consumes at least one byte. This is a safe
    // aggregate upper bound across all nested messages and opaque formulas.
    let fields = input_bytes.checked_mul(8).ok_or(Error::InvalidSource)?;
    let formula_occurrences = deltas
        .iter()
        .filter(|delta| delta.new_formula.is_some())
        .count();
    let final_host_upper = source
        .source_hosts
        .len()
        .checked_add(
            deltas
                .iter()
                .filter(|delta| delta.old_formula_key.is_none() && delta.new_formula.is_some())
                .count(),
        )
        .ok_or(Error::InvalidSource)?;
    ensure(Resource::Hosts, final_host_upper, limits.max_hosts)?;
    let (authored_bytes, largest_authored) =
        deltas
            .iter()
            .try_fold((0usize, 0usize), |(total, largest), delta| {
                let bytes = delta.new_formula.map_or(0, <[u8]>::len);
                ensure(Resource::OutputBytes, bytes, limits.max_output_bytes)?;
                Ok::<_, Error>((
                    total.checked_add(bytes).ok_or(Error::InvalidSource)?,
                    largest.max(bytes),
                ))
            })?;
    let index_count = source
        .expected_entries
        .checked_mul(3)
        .and_then(|count| count.checked_add(source.source_hosts.len()))
        .and_then(|count| count.checked_add(deltas.len().checked_mul(3)?))
        .and_then(|count| count.checked_add(formula_occurrences))
        .ok_or(Error::InvalidSource)?;
    let scratch_bytes = source
        .expected_entries
        .checked_mul(size_of::<Entry<'_>>())
        .and_then(|bytes| {
            index_count
                .checked_mul(size_of::<usize>())
                .and_then(|indexes| bytes.checked_add(indexes))
        })
        .and_then(|bytes| {
            formula_occurrences
                .checked_mul(size_of::<NewFormula<'_>>())
                .and_then(|formulas| bytes.checked_add(formulas))
        })
        .and_then(|bytes| bytes.checked_add(input_bytes.checked_mul(8)?))
        .and_then(|bytes| {
            source
                .expected_entries
                .checked_mul(size_of::<Entry<'_>>())
                .and_then(|reopen| bytes.checked_add(reopen))
        })
        .and_then(|bytes| bytes.checked_add(largest_authored.checked_add(32)?))
        .ok_or(Error::InvalidSource)?;
    let logical_retained_elements = source
        .expected_entries
        .checked_mul(4)
        .and_then(|count| count.checked_add(formula_occurrences.checked_mul(2)?))
        .and_then(|count| count.checked_add(deltas.len().checked_mul(2)?))
        .ok_or(Error::InvalidSource)?;
    let entry_logical_bytes = size_of::<Entry<'_>>()
        .checked_add(
            size_of::<usize>()
                .checked_mul(2)
                .ok_or(Error::InvalidSource)?,
        )
        .and_then(|bytes| bytes.checked_add(size_of::<LogicalEntry<'_>>()))
        .ok_or(Error::InvalidSource)?;
    let logical_retained_bytes = source
        .expected_entries
        .checked_mul(entry_logical_bytes)
        .and_then(|bytes| {
            formula_occurrences
                .checked_mul(
                    size_of::<NewFormula<'_>>().checked_add(size_of::<LogicalEntry<'_>>())?,
                )
                .and_then(|formula_bytes| bytes.checked_add(formula_bytes))
        })
        .and_then(|bytes| {
            deltas
                .len()
                .checked_mul(size_of::<Assignment>().checked_add(size_of::<LogicalAssignment>())?)
                .and_then(|delta_bytes| bytes.checked_add(delta_bytes))
        })
        .ok_or(Error::InvalidSource)?;
    // Logical preparation retains the seven vectors represented by
    // `logical_retained_bytes`. `formula_order` and the source-host key
    // multiset are phase-local and never survive in `LogicalPlan`; charge the
    // largest of those temporary buffers (plus bounded decode staging) once,
    // instead of counting the retained vectors a second time.
    let decode_scratch_bytes = input_bytes
        .checked_mul(8)
        .and_then(|bytes| bytes.checked_add(largest_authored.checked_add(32)?))
        .ok_or(Error::InvalidSource)?;
    let formula_order_bytes = source
        .expected_entries
        .checked_mul(size_of::<usize>())
        .ok_or(Error::InvalidSource)?;
    let count_key_bytes = source
        .source_hosts
        .len()
        .checked_mul(size_of::<u32>())
        .ok_or(Error::InvalidSource)?;
    let logical_transient_bytes = decode_scratch_bytes
        .max(formula_order_bytes)
        .max(count_key_bytes);
    let logical_scratch_bytes = logical_retained_bytes
        .checked_add(logical_transient_bytes)
        .ok_or(Error::InvalidSource)?;
    let logical_allocations = [
        source.expected_entries,
        source.expected_entries,
        source.expected_entries,
        source.expected_entries,
        source.source_hosts.len(),
        formula_occurrences,
        deltas.len(),
        deltas.len(),
        source
            .expected_entries
            .checked_add(formula_occurrences)
            .ok_or(Error::InvalidSource)?,
    ]
    .into_iter()
    .filter(|count| *count != 0)
    .count();
    let archive_references = source.segments.iter().try_fold(
        source.root.object_references.len(),
        |total, segment| {
            total
                .checked_add(segment.object_references.len())
                .ok_or(Error::InvalidSource)
        },
    )?;
    let reference_bytes = archive_references
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)?;
    // Root segment references are independently traversed by strict source,
    // topology, and candidate-reopen passes. Segment payloads may not carry
    // nested list routes.
    let references = archive_references
        .checked_add(
            source
                .segments
                .len()
                .checked_mul(3)
                .ok_or(Error::InvalidSource)?,
        )
        // In the hostile upper bound every source byte can encode a one-byte
        // alternate reference field inside an entry. Source and candidate
        // entry validation each inspect it independently.
        .and_then(|count| count.checked_add(input_bytes.checked_mul(2)?))
        .ok_or(Error::InvalidSource)?;
    let retained_upper = input_bytes
        .checked_add(
            deltas
                .len()
                .checked_mul(size_of::<Assignment>())
                .ok_or(Error::InvalidSource)?,
        )
        .and_then(|bytes| {
            messages
                .checked_mul(size_of::<MessageEdit>())
                .and_then(|fixed| bytes.checked_add(fixed))
        })
        .and_then(|bytes| {
            source
                .expected_entries
                .checked_add(formula_occurrences)?
                .checked_mul(size_of::<FinalEntry>())
                .and_then(|fixed| bytes.checked_add(fixed))
        })
        .and_then(|bytes| bytes.checked_add(input_bytes.checked_mul(2)?))
        .and_then(|bytes| bytes.checked_add(authored_bytes.checked_mul(2)?))
        .and_then(|bytes| bytes.checked_add(reference_bytes))
        .ok_or(Error::InvalidSource)?;
    // All manual passes are linear after compact indexes are sorted. Charge
    // the comparison upper bounds of those sorts before any allocation.
    let delta_sort = sort_work(deltas.len())?
        .checked_mul(2)
        .ok_or(Error::InvalidSource)?;
    let formula_sort = sort_work(formula_occurrences)?;
    let formula_comparison_depth = comparison_depth(formula_occurrences)?;
    let sort_work = sort_work(source.expected_entries)?
        .checked_add(sort_work(source.source_hosts.len())?)
        .and_then(|work| work.checked_add(delta_sort))
        .and_then(|work| work.checked_add(formula_sort))
        .ok_or(Error::InvalidSource)?;
    let formula_byte_passes = formula_comparison_depth
        .checked_mul(2)
        .and_then(|passes| {
            comparison_depth(source.expected_entries)
                .ok()?
                .checked_mul(2)
                .and_then(|depth| passes.checked_add(depth))
        })
        .and_then(|passes| passes.checked_add(1))
        .ok_or(Error::InvalidSource)?;
    let work = input_bytes
        .checked_mul(128)
        .and_then(|work| work.checked_add(sort_work))
        .and_then(|work| work.checked_add(index_count.checked_mul(6)?))
        .and_then(|work| {
            authored_bytes
                .checked_mul(formula_byte_passes)
                .and_then(|bytes| work.checked_add(bytes))
        })
        .ok_or(Error::InvalidSource)?;
    let allocations = input_bytes
        .checked_add(10)
        .and_then(|count| count.checked_add(messages.checked_mul(3)?))
        .and_then(|count| count.checked_add(source.expected_entries))
        .and_then(|count| count.checked_add(formula_occurrences))
        .ok_or(Error::InvalidSource)?;
    let output_upper = input_bytes
        .checked_add(authored_bytes)
        .and_then(|bytes| bytes.checked_add(formula_occurrences.checked_mul(64)?))
        .ok_or(Error::InvalidSource)?;
    ensure(Resource::OutputBytes, output_upper, limits.max_output_bytes)?;
    // Delayed execution decodes the source envelopes once more, rewrites the
    // candidate, and strictly reopens that candidate. Candidate fields can be
    // proportional to authored output even when the source list is empty.
    let execution_bytes = input_bytes
        .checked_add(output_upper)
        .ok_or(Error::InvalidSource)?;
    let execution_fields = execution_bytes.checked_mul(8).ok_or(Error::InvalidSource)?;
    let execution_work = execution_bytes
        .checked_mul(128)
        .and_then(|value| value.checked_add(index_count.checked_mul(6)?))
        .ok_or(Error::InvalidSource)?;
    ensure(Resource::Fields, execution_fields, limits.max_fields)?;
    ensure(Resource::Work, execution_work, limits.max_work)?;
    ensure(Resource::References, references, limits.max_references)?;
    ensure(
        Resource::RetainedElements,
        logical_retained_elements,
        limits.max_retained_elements,
    )?;
    ensure(
        Resource::RetainedBytes,
        logical_retained_bytes,
        limits.max_retained_bytes,
    )?;
    ensure(
        Resource::ScratchBytes,
        logical_scratch_bytes,
        limits.max_scratch_bytes,
    )?;
    ensure(
        Resource::Allocations,
        logical_allocations,
        limits.max_allocations,
    )?;
    Ok(Requirements {
        input_bytes,
        fields,
        work,
        execution_fields,
        execution_work,
        entries: source
            .expected_entries
            .checked_add(formula_occurrences)
            .ok_or(Error::InvalidSource)?,
        hosts: source.source_hosts.len(),
        formula_occurrences,
        output_upper,
        references,
        scratch_bytes,
        retained_upper,
        allocations,
        logical_retained_elements,
        logical_retained_bytes,
        logical_transient_bytes,
        logical_scratch_bytes,
        logical_allocations,
    })
}

fn admit(requirements: Requirements, limits: Limits) -> Result<(), Error> {
    ensure(
        Resource::InputBytes,
        requirements.input_bytes,
        limits.max_input_bytes,
    )?;
    ensure(Resource::Fields, requirements.fields, limits.max_fields)?;
    ensure(Resource::Work, requirements.work, limits.max_work)?;
    ensure(Resource::Entries, requirements.entries, limits.max_entries)?;
    ensure(Resource::Hosts, requirements.hosts, limits.max_hosts)?;
    ensure(
        Resource::Hosts,
        requirements.formula_occurrences,
        limits.max_hosts,
    )?;
    ensure(
        Resource::OutputBytes,
        requirements.output_upper,
        limits.max_output_bytes,
    )?;
    ensure(
        Resource::References,
        requirements.references,
        limits.max_references,
    )?;
    ensure(
        Resource::ScratchBytes,
        requirements.scratch_bytes,
        limits.max_scratch_bytes,
    )?;
    ensure(
        Resource::RetainedBytes,
        requirements.retained_upper,
        limits.max_retained_bytes,
    )?;
    ensure(
        Resource::Allocations,
        requirements.allocations,
        limits.max_allocations,
    )
}

fn prepare_logical<'a>(
    source: SourceList<'a>,
    deltas: &[HostDelta<'a>],
    limits: Limits,
    requirements: Requirements,
) -> Result<LogicalPlan<'a>, Error> {
    let options = storage::DecodeOptions::new(
        limits.max_input_bytes,
        limits.max_fields,
        limits.max_work,
        16,
        limits.max_references,
        limits.max_input_bytes,
    );
    let (root, _root_report) =
        storage::decode_table_data_list_with_report(source.root.payload, options)
            .map_err(|_| Error::Strict)?;
    if root.list_type() != FORMULA_LIST_TYPE || root.next_list_id() == 0 {
        return Err(Error::InvalidSource);
    }
    validate_root_segment_references(source.root.payload, source.segments, limits)?;

    let mut entries = exact_vec(source.expected_entries)?;
    decode_entries(source.root.payload, 0, options, &mut entries, limits)?;
    for (owner, segment) in source.segments.iter().enumerate() {
        let (snapshot, _) =
            storage::decode_table_data_list_segment_with_report(segment.payload, options)
                .map_err(|_| Error::Strict)?;
        if snapshot.list_type() != FORMULA_LIST_TYPE
            || snapshot.key_range_length() == 0
            || contains_field(segment.payload, 4, limits)?
        {
            return Err(Error::InvalidSource);
        }
        let begin = entries.len();
        decode_entries(segment.payload, owner + 1, options, &mut entries, limits)?;
        let end = snapshot
            .key_range_location()
            .checked_add(snapshot.key_range_length())
            .ok_or(Error::InvalidSource)?;
        if entries[begin..]
            .iter()
            .any(|entry| entry.key < snapshot.key_range_location() || entry.key >= end)
        {
            return Err(Error::InvalidSource);
        }
    }
    if entries.len() != source.expected_entries {
        return Err(Error::InvalidSource);
    }

    let mut key_order = exact_vec::<usize>(entries.len())?;
    let mut formula_order = exact_vec::<usize>(entries.len())?;
    let mut owner_order = exact_vec::<usize>(entries.len())?;
    for index in 0..entries.len() {
        key_order.push(index);
        formula_order.push(index);
        owner_order.push(index);
    }
    key_order.sort_unstable_by_key(|index| entries[*index].key);
    formula_order
        .sort_unstable_by(|left, right| entries[*left].formula.cmp(entries[*right].formula));
    owner_order.sort_unstable_by_key(|index| (entries[*index].owner, entries[*index].source_index));
    if key_order
        .windows(2)
        .any(|pair| entries[pair[0]].key >= entries[pair[1]].key)
        || formula_order
            .windows(2)
            .any(|pair| entries[pair[0]].formula == entries[pair[1]].formula)
        || entries
            .iter()
            .any(|entry| entry.key == 0 || entry.initial_count == 0)
    {
        return Err(Error::InvalidSource);
    }
    if key_index(&entries, &key_order, root.next_list_id()).is_some() {
        return Err(Error::InvalidSource);
    }
    validate_counts(&entries, &key_order, source.source_hosts)?;
    validate_delta_hosts(source.source_hosts, deltas)?;

    let mut new_formulas = exact_vec::<NewFormula<'_>>(requirements.formula_occurrences)?;
    for delta in deltas {
        if let Some(bytes) = delta.new_formula {
            if bytes.is_empty() {
                return Err(Error::InvalidSource);
            }
            new_formulas.push(NewFormula {
                bytes,
                key: 0,
                count: 0,
            });
        }
    }
    new_formulas.sort_unstable_by(|left, right| left.bytes.cmp(right.bytes));
    new_formulas.dedup_by(|left, right| left.bytes == right.bytes);
    let mut next_key = root.next_list_id();
    for formula in &mut new_formulas {
        if formula_index(&entries, &formula_order, formula.bytes).is_some() {
            continue;
        }
        if next_key == 0 || key_index(&entries, &key_order, next_key).is_some() {
            return Err(Error::InvalidSource);
        }
        formula.key = next_key;
        next_key = next_key.checked_add(1).ok_or(Error::InvalidSource)?;
    }

    let mut assignments = exact_vec::<Assignment>(deltas.len())?;
    for (delta_index, delta) in deltas.iter().enumerate() {
        if let Some(old_key) = delta.old_formula_key {
            let entry_index =
                key_index(&entries, &key_order, old_key).ok_or(Error::InvalidSource)?;
            entries[entry_index].final_count = entries[entry_index]
                .final_count
                .checked_sub(1)
                .ok_or(Error::InvalidSource)?;
        }
        let key = if let Some(bytes) = delta.new_formula {
            if let Some(entry_index) = formula_index(&entries, &formula_order, bytes) {
                entries[entry_index].final_count = entries[entry_index]
                    .final_count
                    .checked_add(1)
                    .ok_or(Error::InvalidSource)?;
                Some(entries[entry_index].key)
            } else {
                let index = new_formulas
                    .binary_search_by(|formula| formula.bytes.cmp(bytes))
                    .map_err(|_| Error::InvalidSource)?;
                new_formulas[index].count = new_formulas[index]
                    .count
                    .checked_add(1)
                    .ok_or(Error::InvalidSource)?;
                Some(new_formulas[index].key)
            }
        } else {
            None
        };
        assignments.push(Assignment {
            delta: delta_index,
            key,
        });
    }
    // Formula lookup is complete. Keep this allocation phase-local so the
    // retained logical report exactly describes the state crossing the cache
    // validation barrier.
    drop(formula_order);

    let logical_capacity = entries
        .iter()
        .filter(|entry| entry.final_count != 0)
        .count()
        .checked_add(
            new_formulas
                .iter()
                .filter(|formula| formula.key != 0 && formula.count != 0)
                .count(),
        )
        .ok_or(Error::InvalidSource)?;
    ensure(Resource::Entries, logical_capacity, limits.max_entries)?;
    let mut logical_entries = exact_vec::<LogicalEntry<'a>>(logical_capacity)?;
    for entry_index in &key_order {
        let entry = entries[*entry_index];
        if entry.final_count != 0 {
            logical_entries.push(LogicalEntry {
                key: entry.key,
                ref_count: entry.final_count,
                formula: entry.formula,
            });
        }
    }
    for formula in new_formulas
        .iter()
        .filter(|formula| formula.key != 0 && formula.count != 0)
    {
        logical_entries.push(LogicalEntry {
            key: formula.key,
            ref_count: formula.count,
            formula: formula.bytes,
        });
    }
    logical_entries.sort_unstable_by_key(|entry| entry.key);
    let mut logical_assignments = exact_vec::<LogicalAssignment>(assignments.len())?;
    for assignment in &assignments {
        let delta = deltas.get(assignment.delta).ok_or(Error::InvalidSource)?;
        logical_assignments.push(LogicalAssignment {
            delta: assignment.delta,
            row: delta.row,
            column: delta.column,
            key: assignment.key,
        });
    }
    let actual_elements = entries
        .capacity()
        .checked_add(key_order.capacity())
        .and_then(|count| count.checked_add(owner_order.capacity()))
        .and_then(|count| count.checked_add(new_formulas.capacity()))
        .and_then(|count| count.checked_add(assignments.capacity()))
        .and_then(|count| count.checked_add(logical_assignments.capacity()))
        .and_then(|count| count.checked_add(logical_entries.capacity()))
        .ok_or(Error::InvalidSource)?;
    if actual_elements > requirements.logical_retained_elements {
        return Err(Error::Allocation {
            requested: actual_elements,
        });
    }
    let actual_bytes = entries
        .capacity()
        .checked_mul(size_of::<Entry<'_>>())
        .and_then(|bytes| {
            key_order
                .capacity()
                .checked_add(owner_order.capacity())
                .and_then(|count| count.checked_mul(size_of::<usize>()))
                .and_then(|index_bytes| bytes.checked_add(index_bytes))
        })
        .and_then(|bytes| {
            new_formulas
                .capacity()
                .checked_mul(size_of::<NewFormula<'_>>())
                .and_then(|value| bytes.checked_add(value))
        })
        .and_then(|bytes| {
            assignments
                .capacity()
                .checked_mul(size_of::<Assignment>())
                .and_then(|value| bytes.checked_add(value))
        })
        .and_then(|bytes| {
            logical_assignments
                .capacity()
                .checked_mul(size_of::<LogicalAssignment>())
                .and_then(|value| bytes.checked_add(value))
        })
        .and_then(|bytes| {
            logical_entries
                .capacity()
                .checked_mul(size_of::<LogicalEntry<'_>>())
                .and_then(|value| bytes.checked_add(value))
        })
        .ok_or(Error::InvalidSource)?;
    if actual_bytes > requirements.logical_retained_bytes {
        return Err(Error::Allocation {
            requested: actual_bytes,
        });
    }
    let execution_elements = assignments
        .len()
        .checked_add(source.segments.len())
        .and_then(|count| count.checked_add(logical_entries.len()))
        .and_then(|count| count.checked_add(1))
        .ok_or(Error::InvalidSource)?;
    // Execution runs with the exact logical plan still live, then allocates
    // the future Artifact and its independent rewrite/reopen scratch. Do not
    // reuse the logical upper here: duplicate formulas and removals can make
    // the retained plan materially smaller than that admission ceiling.
    let moved_assignment_bytes = assignments
        .capacity()
        .checked_mul(size_of::<Assignment>())
        .ok_or(Error::InvalidSource)?;
    let logical_view_bytes = logical_assignments
        .capacity()
        .checked_mul(size_of::<LogicalAssignment>())
        .and_then(|bytes| {
            logical_entries
                .capacity()
                .checked_mul(size_of::<LogicalEntry<'_>>())
                .and_then(|entry_bytes| bytes.checked_add(entry_bytes))
        })
        .ok_or(Error::InvalidSource)?;
    let execution_plan_bytes = actual_bytes
        .checked_sub(logical_view_bytes)
        .and_then(|bytes| bytes.checked_sub(moved_assignment_bytes))
        .ok_or(Error::InvalidSource)?;
    let execution_transient = requirements
        .input_bytes
        .checked_mul(8)
        .ok_or(Error::InvalidSource)?;
    let execution_peak = execution_plan_bytes
        .checked_add(requirements.retained_upper)
        .and_then(|bytes| bytes.checked_add(execution_transient))
        .ok_or(Error::InvalidSource)?;
    let assignment_count = assignments.len();
    Ok(LogicalPlan {
        source,
        requirements,
        entries,
        key_order,
        owner_order,
        new_formulas,
        assignments,
        logical_assignments,
        logical_entries,
        next_key,
        report: LogicalReport {
            input_bytes: requirements.input_bytes,
            fields: requirements.fields,
            work: requirements.work,
            references: requirements.references,
            output_bytes: 0,
            retained_elements: actual_elements,
            retained_bytes: actual_bytes,
            peak_scratch_bytes: actual_bytes
                .checked_add(requirements.logical_transient_bytes)
                .ok_or(Error::InvalidSource)?,
            allocations: requirements.logical_allocations,
        },
        execution: ExecutionRequirements {
            // Candidate rewrite and strict reopen independently traverse the
            // complete source/candidate closure. Reusing the full conservative
            // per-phase ceiling keeps this delayed phase independently bounded.
            input_bytes: requirements.input_bytes,
            fields: requirements.execution_fields,
            work: requirements.execution_work,
            references: requirements.references,
            output_bytes: requirements.output_upper,
            retained_elements: execution_elements,
            retained_bytes: requirements.retained_upper,
            peak_scratch_bytes: execution_peak,
            allocations: requirements
                .allocations
                .checked_sub(requirements.logical_allocations)
                .ok_or(Error::InvalidSource)?,
            assignments: assignment_count,
            changed_messages: source
                .segments
                .len()
                .checked_add(1)
                .ok_or(Error::InvalidSource)?,
        },
    })
}

fn execute_logical(plan: LogicalPlan<'_>, limits: Limits) -> Result<Artifact, Error> {
    let LogicalPlan {
        source,
        requirements,
        entries,
        key_order,
        owner_order,
        new_formulas,
        assignments,
        logical_assignments: _,
        logical_entries,
        next_key,
        report: _,
        execution,
    } = plan;
    // The public logical view is no longer live once execution begins. Its
    // two index-only vectors are dropped before candidate/output allocation;
    // the execution peak therefore overlaps only the state the writer still
    // consumes.
    drop(logical_entries);
    let options = storage::DecodeOptions::new(
        limits.max_input_bytes,
        limits.max_fields,
        limits.max_work,
        16,
        limits.max_references,
        limits.max_input_bytes,
    );
    let (root, root_report) =
        storage::decode_table_data_list_with_report(source.root.payload, options)
            .map_err(|_| Error::Strict)?;

    let mut root_candidate = rewrite_message(
        source.root.payload,
        0,
        &entries,
        &owner_order,
        &new_formulas,
        Some((root.next_list_id(), next_key)),
        limits,
    )?;
    let mut segment_edits = exact_vec::<MessageEdit>(source.segments.len())?;
    for (owner, segment) in source.segments.iter().enumerate() {
        let candidate = rewrite_message(
            segment.payload,
            owner + 1,
            &entries,
            &owner_order,
            &[],
            None,
            limits,
        )?;
        segment_edits.push(finish_edit(*segment, candidate)?);
    }
    let root_edit = finish_edit(source.root, core::mem::take(&mut root_candidate))?;

    let final_capacity = entries
        .iter()
        .filter(|entry| entry.final_count != 0)
        .count()
        .checked_add(
            new_formulas
                .iter()
                .filter(|formula| formula.key != 0 && formula.count != 0)
                .count(),
        )
        .ok_or(Error::InvalidSource)?;
    ensure(Resource::Entries, final_capacity, limits.max_entries)?;
    let mut final_entries = exact_vec::<FinalEntry>(final_capacity)?;
    for entry_index in &key_order {
        let entry = entries[*entry_index];
        if entry.final_count != 0 {
            final_entries.push(FinalEntry {
                key: entry.key,
                ref_count: entry.final_count,
                formula: exact_copy(entry.formula)?,
            });
        }
    }
    for formula in new_formulas
        .iter()
        .filter(|formula| formula.key != 0 && formula.count != 0)
    {
        final_entries.push(FinalEntry {
            key: formula.key,
            ref_count: formula.count,
            formula: exact_copy(formula.bytes)?,
        });
    }
    final_entries.sort_unstable_by_key(|entry| entry.key);
    reopen_candidates(&root_edit, &segment_edits, source, &final_entries, limits)?;
    let hosts = final_entries.iter().try_fold(0usize, |sum, entry| {
        sum.checked_add(usize::try_from(entry.ref_count).map_err(|_| Error::InvalidSource)?)
            .ok_or(Error::InvalidSource)
    })?;
    ensure(Resource::Hosts, hosts, limits.max_hosts)?;
    let output_bytes = root_edit
        .payload
        .as_ref()
        .map_or(0, Vec::len)
        .checked_add(segment_edits.iter().try_fold(0usize, |sum, edit| {
            sum.checked_add(edit.payload.as_ref().map_or(0, Vec::len))
                .ok_or(Error::InvalidSource)
        })?)
        .ok_or(Error::InvalidSource)?;
    let retained_bytes =
        artifact_retained(&assignments, &root_edit, &segment_edits, &final_entries)?;
    ensure(
        Resource::RetainedBytes,
        retained_bytes,
        requirements.retained_upper.min(limits.max_retained_bytes),
    )?;
    let report = Report {
        input_bytes: execution.input_bytes,
        output_bytes,
        fields: execution.fields,
        work: execution.work,
        references: execution.references,
        entries: final_entries.len(),
        hosts,
        allocations: execution.allocations,
        retained_bytes,
        peak_scratch_bytes: execution.peak_scratch_bytes,
        changed_messages: usize::from(root_edit.payload.is_some())
            + segment_edits
                .iter()
                .filter(|edit| edit.payload.is_some())
                .count(),
        retained_elements: assignments
            .capacity()
            .checked_add(segment_edits.capacity())
            .and_then(|elements| elements.checked_add(final_entries.capacity()))
            .and_then(|elements| elements.checked_add(1))
            .ok_or(Error::InvalidSource)?,
        assignments: assignments.len(),
    };
    let _ = root_report;
    Ok(Artifact {
        assignments,
        root: root_edit,
        segments: segment_edits,
        final_entries,
        report,
    })
}

fn decode_entries<'a>(
    source: &'a [u8],
    owner: usize,
    options: storage::DecodeOptions,
    output: &mut Vec<Entry<'a>>,
    limits: Limits,
) -> Result<(), Error> {
    let view = WireView::parse_with_limits(source, wire_limits(source.len(), limits)?)
        .map_err(|_error| Error::Wire)?;
    let mut source_index = 0usize;
    for field in view.fields().filter(|field| field.number() == ENTRY_FIELD) {
        if field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(|_error| Error::Wire)?;
        let raw = field.payload();
        let entry =
            storage::decode_table_data_list_entry(raw, options).map_err(|_| Error::Strict)?;
        let formula = entry.formula().ok_or(Error::InvalidSource)?;
        if entry.key() == 0
            || entry.ref_count() == 0
            || entry.string_value().is_some()
            || entry.reference().is_some()
            || entry.format().is_some()
            || entry.custom_format().is_some()
            || entry.rich_text_payload().is_some()
            || entry.comment_storage().is_some()
            || entry.import_warning_set().is_some()
            || entry.cell_spec().is_some()
        {
            return Err(Error::InvalidSource);
        }
        if output.len() == output.capacity() {
            return Err(Error::InvalidSource);
        }
        output.push(Entry {
            key: entry.key(),
            initial_count: entry.ref_count(),
            final_count: entry.ref_count(),
            formula,
            raw,
            owner,
            source_index,
        });
        source_index = source_index.checked_add(1).ok_or(Error::InvalidSource)?;
    }
    Ok(())
}

fn contains_field(source: &[u8], number: u32, limits: Limits) -> Result<bool, Error> {
    let view = WireView::parse_with_limits(source, wire_limits(source.len(), limits)?)
        .map_err(|_error| Error::Wire)?;
    Ok(view.fields().any(|field| field.number() == number))
}

fn validate_root_segment_references(
    root: &[u8],
    segments: &[SourceMessage<'_>],
    limits: Limits,
) -> Result<(), Error> {
    struct References<'a> {
        segments: &'a [SourceMessage<'a>],
        index: usize,
        invalid: bool,
    }
    impl storage::StorageVisitor for References<'_> {
        fn visit_list_segment(
            &mut self,
            reference: storage::ReferenceRecord<'_>,
        ) -> Result<(), storage::DecodeError> {
            let reference = reference.reference();
            if reference.deprecated_is_external() == Some(true)
                || self
                    .segments
                    .get(self.index)
                    .is_none_or(|segment| segment.object_id != reference.identifier())
            {
                self.invalid = true;
            }
            self.index = self.index.saturating_add(1);
            Ok(())
        }
    }
    let mut visitor = References {
        segments,
        index: 0,
        invalid: false,
    };
    storage::decode_table_data_list_with_visitor(
        root,
        storage::DecodeOptions::new(
            limits.max_input_bytes,
            limits.max_fields,
            limits.max_work,
            16,
            limits.max_references,
            limits.max_input_bytes,
        ),
        &mut visitor,
    )
    .map_err(|_| Error::Strict)?;
    if visitor.invalid || visitor.index != segments.len() {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn validate_counts(
    entries: &[Entry<'_>],
    key_order: &[usize],
    hosts: &[SourceHost],
) -> Result<(), Error> {
    let mut keys = exact_vec::<u32>(hosts.len())?;
    for host in hosts {
        keys.push(host.key);
    }
    keys.sort_unstable();
    let mut key_cursor = 0usize;
    for entry_index in key_order {
        let entry = entries[*entry_index];
        let start = key_cursor;
        while keys.get(key_cursor) == Some(&entry.key) {
            key_cursor += 1;
        }
        if key_cursor - start
            != usize::try_from(entry.initial_count).map_err(|_| Error::InvalidSource)?
        {
            return Err(Error::InvalidSource);
        }
    }
    if key_cursor != keys.len() {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn validate_delta_hosts(hosts: &[SourceHost], deltas: &[HostDelta<'_>]) -> Result<(), Error> {
    for delta in deltas {
        let source = hosts
            .binary_search_by_key(&(delta.row, delta.column), |host| (host.row, host.column))
            .ok()
            .map(|index| hosts[index].key);
        if source != delta.old_formula_key {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn rewrite_message(
    source: &[u8],
    owner: usize,
    entries: &[Entry<'_>],
    owner_order: &[usize],
    added: &[NewFormula<'_>],
    next: Option<(u32, u32)>,
    limits: Limits,
) -> Result<Vec<u8>, Error> {
    let view = WireView::parse_with_limits(source, wire_limits(source.len(), limits)?)
        .map_err(|_error| Error::Wire)?;
    let owner_indexes = owner_slice(entries, owner_order, owner);
    let mut exact = source.len();
    let mut entry_cursor = 0usize;
    let mut next_seen = 0usize;
    for field in view.fields() {
        if field.number() == ENTRY_FIELD {
            let entry = owner_indexes
                .get(entry_cursor)
                .map(|index| entries[*index])
                .ok_or(Error::InvalidSource)?;
            entry_cursor += 1;
            let replacement = if entry.final_count == 0 {
                0
            } else if entry.final_count == entry.initial_count {
                field.raw().len()
            } else {
                length_field_len(ENTRY_FIELD, rewritten_entry_len(entry)?)?
            };
            exact = exact
                .checked_sub(field.raw().len())
                .and_then(|value| value.checked_add(replacement))
                .ok_or(Error::InvalidSource)?;
        } else if let Some((old, new)) = next {
            if field.number() == NEXT_KEY_FIELD {
                next_seen += 1;
                if old != new {
                    exact = exact
                        .checked_sub(field.raw().len())
                        .and_then(|value| {
                            value.checked_add(varint_len(NEXT_KEY_FIELD, u64::from(new)))
                        })
                        .ok_or(Error::InvalidSource)?;
                }
            }
        }
    }
    if entry_cursor != owner_indexes.len() || next.is_some() && next_seen != 1 {
        return Err(Error::InvalidSource);
    }
    for formula in added
        .iter()
        .filter(|formula| formula.key != 0 && formula.count != 0)
    {
        exact = exact
            .checked_add(length_field_len(ENTRY_FIELD, new_entry_len(*formula)?)?)
            .ok_or(Error::InvalidSource)?;
    }
    ensure(Resource::OutputBytes, exact, limits.max_output_bytes)?;
    let mut output = exact_vec::<u8>(exact)?;
    entry_cursor = 0;
    for field in view.fields() {
        if field.number() == ENTRY_FIELD {
            let entry = entries[*owner_indexes
                .get(entry_cursor)
                .ok_or(Error::InvalidSource)?];
            entry_cursor += 1;
            if entry.final_count == 0 {
                continue;
            }
            if entry.final_count == entry.initial_count {
                extend_exact(&mut output, field.raw())?;
            } else {
                append_length(&mut output, ENTRY_FIELD, &rewrite_entry(entry, limits)?)?;
            }
        } else if let Some((old, new)) = next {
            if field.number() == NEXT_KEY_FIELD && old != new {
                append_varint(&mut output, NEXT_KEY_FIELD, u64::from(new));
            } else {
                extend_exact(&mut output, field.raw())?;
            }
        } else {
            extend_exact(&mut output, field.raw())?;
        }
    }
    for formula in added
        .iter()
        .filter(|formula| formula.key != 0 && formula.count != 0)
    {
        append_new_entry(&mut output, *formula, limits)?;
    }
    if output.len() != exact {
        return Err(Error::InvalidSource);
    }
    Ok(output)
}

fn owner_slice<'a>(entries: &[Entry<'_>], order: &'a [usize], owner: usize) -> &'a [usize] {
    let start = order.partition_point(|index| entries[*index].owner < owner);
    let end = order.partition_point(|index| entries[*index].owner <= owner);
    &order[start..end]
}

fn key_index(entries: &[Entry<'_>], order: &[usize], key: u32) -> Option<usize> {
    order
        .binary_search_by_key(&key, |index| entries[*index].key)
        .ok()
        .map(|slot| order[slot])
}

fn formula_index(entries: &[Entry<'_>], order: &[usize], bytes: &[u8]) -> Option<usize> {
    order
        .binary_search_by(|index| entries[*index].formula.cmp(bytes))
        .ok()
        .map(|slot| order[slot])
}

fn rewritten_entry_len(entry: Entry<'_>) -> Result<usize, Error> {
    let fields =
        parse_wire_fields_with_limits(entry.raw, WireLimits::default()).map_err(|_| Error::Wire)?;
    let mut exact = entry.raw.len();
    let mut seen = false;
    for field in fields {
        if field.number() == REF_COUNT_FIELD {
            if seen || field.wire_type() != 0 {
                return Err(Error::InvalidSource);
            }
            seen = true;
            exact = exact
                .checked_sub(field.raw(entry.raw).map_err(|_| Error::Wire)?.len())
                .and_then(|value| {
                    value.checked_add(varint_len(REF_COUNT_FIELD, u64::from(entry.final_count)))
                })
                .ok_or(Error::InvalidSource)?;
        }
    }
    if seen {
        Ok(exact)
    } else {
        Err(Error::InvalidSource)
    }
}

fn rewrite_entry(entry: Entry<'_>, limits: Limits) -> Result<Vec<u8>, Error> {
    let exact = rewritten_entry_len(entry)?;
    ensure(Resource::OutputBytes, exact, limits.max_output_bytes)?;
    let fields = parse_wire_fields_with_limits(entry.raw, wire_limits(entry.raw.len(), limits)?)
        .map_err(|_| Error::Wire)?;
    let mut output = exact_vec(exact)?;
    for field in fields {
        if field.number() == REF_COUNT_FIELD {
            append_varint(&mut output, REF_COUNT_FIELD, u64::from(entry.final_count));
        } else {
            extend_exact(&mut output, field.raw(entry.raw).map_err(|_| Error::Wire)?)?;
        }
    }
    Ok(output)
}

fn new_entry_len(formula: NewFormula<'_>) -> Result<usize, Error> {
    varint_len(KEY_FIELD, u64::from(formula.key))
        .checked_add(varint_len(REF_COUNT_FIELD, u64::from(formula.count)))
        .and_then(|value| {
            value.checked_add(length_field_len(FORMULA_FIELD, formula.bytes.len()).ok()?)
        })
        .ok_or(Error::InvalidSource)
}

fn append_new_entry(
    output: &mut Vec<u8>,
    formula: NewFormula<'_>,
    limits: Limits,
) -> Result<(), Error> {
    let exact = new_entry_len(formula)?;
    ensure(Resource::OutputBytes, exact, limits.max_output_bytes)?;
    let outer = length_field_len(ENTRY_FIELD, exact)?;
    if output
        .len()
        .checked_add(outer)
        .is_none_or(|length| length > output.capacity())
    {
        return Err(Error::InvalidSource);
    }
    encode_varint_into(output, (u64::from(ENTRY_FIELD) << 3) | 2);
    encode_varint_into(
        output,
        u64::try_from(exact).map_err(|_| Error::InvalidSource)?,
    );
    append_varint(output, KEY_FIELD, u64::from(formula.key));
    append_varint(output, REF_COUNT_FIELD, u64::from(formula.count));
    append_length(output, FORMULA_FIELD, formula.bytes)
}

fn finish_edit(source: SourceMessage<'_>, candidate: Vec<u8>) -> Result<MessageEdit, Error> {
    let mut references = exact_vec(source.object_references.len())?;
    extend_exact(&mut references, source.object_references)?;
    Ok(MessageEdit {
        object_id: source.object_id,
        payload: (candidate != source.payload).then_some(candidate),
        object_references: references,
    })
}

fn reopen_candidates(
    root: &MessageEdit,
    segments: &[MessageEdit],
    source: SourceList<'_>,
    expected: &[FinalEntry],
    limits: Limits,
) -> Result<(), Error> {
    let root_bytes = root.payload.as_deref().unwrap_or(source.root.payload);
    let candidate_options = storage::DecodeOptions::new(
        limits.max_output_bytes,
        limits.max_fields,
        limits.max_work,
        16,
        limits.max_references,
        limits.max_output_bytes,
    );
    let root_snapshot = storage::decode_table_data_list(root_bytes, candidate_options)
        .map_err(|_error| Error::Strict)?;
    if root_snapshot.list_type() != FORMULA_LIST_TYPE {
        return Err(Error::InvalidSource);
    }
    let mut actual = exact_vec::<Entry<'_>>(expected.len())?;
    decode_entries(
        root_bytes,
        0,
        candidate_options,
        &mut actual,
        Limits {
            max_entries: expected.len(),
            max_retained_elements: limits.max_retained_elements,
            ..limits
        },
    )?;
    for (owner, (edit, source_segment)) in segments.iter().zip(source.segments).enumerate() {
        let bytes = edit.payload.as_deref().unwrap_or(source_segment.payload);
        let snapshot = storage::decode_table_data_list_segment(bytes, candidate_options)
            .map_err(|_| Error::Strict)?;
        if snapshot.list_type() != FORMULA_LIST_TYPE {
            return Err(Error::InvalidSource);
        }
        decode_entries(
            bytes,
            owner + 1,
            candidate_options,
            &mut actual,
            Limits {
                max_entries: expected.len(),
                max_retained_elements: limits.max_retained_elements,
                ..limits
            },
        )?;
    }
    if actual.len() != expected.len() {
        return Err(Error::InvalidSource);
    }
    actual.sort_unstable_by_key(|entry| entry.key);
    if actual.iter().zip(expected).any(|(actual, expected)| {
        actual.key != expected.key
            || actual.initial_count != expected.ref_count
            || actual.formula != expected.formula
    }) {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn artifact_retained(
    assignments: &Vec<Assignment>,
    root: &MessageEdit,
    segments: &Vec<MessageEdit>,
    entries: &Vec<FinalEntry>,
) -> Result<usize, Error> {
    assignments
        .capacity()
        .checked_mul(size_of::<Assignment>())
        .and_then(|bytes| {
            bytes.checked_add(segments.capacity().checked_mul(size_of::<MessageEdit>())?)
        })
        .and_then(|bytes| {
            bytes.checked_add(entries.capacity().checked_mul(size_of::<FinalEntry>())?)
        })
        .and_then(|bytes| bytes.checked_add(root.payload.as_ref().map_or(0, Vec::capacity)))
        .and_then(|bytes| {
            bytes.checked_add(
                root.object_references
                    .capacity()
                    .checked_mul(size_of::<u64>())?,
            )
        })
        .and_then(|bytes| {
            segments.iter().try_fold(bytes, |sum, edit| {
                sum.checked_add(edit.payload.as_ref().map_or(0, Vec::capacity))
                    .and_then(|value| {
                        value.checked_add(
                            edit.object_references
                                .capacity()
                                .checked_mul(size_of::<u64>())?,
                        )
                    })
            })
        })
        .and_then(|bytes| {
            entries.iter().try_fold(bytes, |sum, entry| {
                sum.checked_add(entry.formula.capacity())
            })
        })
        .ok_or(Error::InvalidSource)
}

fn exact_vec<T>(capacity: usize) -> Result<Vec<T>, Error> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|_| Error::Allocation {
            requested: capacity,
        })?;
    if size_of::<T>() != 0 && output.capacity() != capacity {
        return Err(Error::Allocation {
            requested: capacity,
        });
    }
    Ok(output)
}

fn exact_copy(source: &[u8]) -> Result<Vec<u8>, Error> {
    let mut output = exact_vec(source.len())?;
    extend_exact(&mut output, source)?;
    Ok(output)
}

fn extend_exact<T: Copy>(output: &mut Vec<T>, values: &[T]) -> Result<(), Error> {
    if output
        .len()
        .checked_add(values.len())
        .is_none_or(|length| length > output.capacity())
    {
        return Err(Error::InvalidSource);
    }
    output.extend_from_slice(values);
    Ok(())
}

fn append_varint(output: &mut Vec<u8>, number: u32, value: u64) {
    encode_varint_into(output, u64::from(number) << 3);
    encode_varint_into(output, value);
}

fn append_length(output: &mut Vec<u8>, number: u32, value: &[u8]) -> Result<(), Error> {
    let needed = length_field_len(number, value.len())?;
    if output
        .len()
        .checked_add(needed)
        .is_none_or(|length| length > output.capacity())
    {
        return Err(Error::InvalidSource);
    }
    encode_varint_into(output, (u64::from(number) << 3) | 2);
    encode_varint_into(
        output,
        u64::try_from(value.len()).map_err(|_| Error::InvalidSource)?,
    );
    output.extend_from_slice(value);
    Ok(())
}

fn varint_len(number: u32, value: u64) -> usize {
    encoded_len(u64::from(number) << 3) + encoded_len(value)
}

fn length_field_len(number: u32, length: usize) -> Result<usize, Error> {
    encoded_len((u64::from(number) << 3) | 2)
        .checked_add(encoded_len(
            u64::try_from(length).map_err(|_| Error::InvalidSource)?,
        ))
        .and_then(|value| value.checked_add(length))
        .ok_or(Error::InvalidSource)
}

fn sort_work(elements: usize) -> Result<usize, Error> {
    if elements < 2 {
        return Ok(elements);
    }
    let log = usize::try_from(usize::BITS - (elements - 1).leading_zeros())
        .map_err(|_| Error::InvalidSource)?;
    elements
        .checked_mul(log)
        .and_then(|work| work.checked_add(elements))
        .ok_or(Error::InvalidSource)
}

fn comparison_depth(elements: usize) -> Result<usize, Error> {
    if elements < 2 {
        return Ok(1);
    }
    usize::try_from(usize::BITS - (elements - 1).leading_zeros()).map_err(|_| Error::InvalidSource)
}

fn wire_limits(input: usize, limits: Limits) -> Result<WireLimits, Error> {
    WireLimits::default()
        .with_input_bytes(input.max(1))
        .map_err(|_error| Error::Wire)?
        .with_output_bytes(limits.max_output_bytes.max(1))
        .map_err(|_error| Error::Wire)
}

fn ensure(resource: Resource, observed: usize, maximum: usize) -> Result<(), Error> {
    if observed > maximum {
        Err(Error::Limit {
            resource,
            observed,
            maximum,
        })
    } else {
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn formula(value: u64) -> Vec<u8> {
        use litchi_iwa_protos::numbers_formula_codec::{
            DecodeOptions, FormulaWriteContext, FormulaWriteDependencyLimits, FormulaWriteNode,
            execute_formula_archive_plan, plan_resolved_formula_archive,
        };
        let options = DecodeOptions::new(10_000, 1_000, 100_000, 16, 10, 1_000);
        let nodes = [FormulaWriteNode::Number(value as f64)];
        let plan = plan_resolved_formula_archive(
            &nodes,
            FormulaWriteContext::new(1, 2, 2, 10, 10),
            &[],
            FormulaWriteDependencyLimits::new(0, 0),
            options,
        )
        .unwrap();
        execute_formula_archive_plan(plan, options).unwrap().0
    }

    fn entry(key: u32, count: u32, formula: &[u8]) -> Vec<u8> {
        let exact = varint_len(KEY_FIELD, u64::from(key))
            + varint_len(REF_COUNT_FIELD, u64::from(count))
            + length_field_len(FORMULA_FIELD, formula.len()).unwrap();
        let mut bytes = exact_vec(exact).unwrap();
        append_varint(&mut bytes, KEY_FIELD, u64::from(key));
        append_varint(&mut bytes, REF_COUNT_FIELD, u64::from(count));
        append_length(&mut bytes, FORMULA_FIELD, formula).unwrap();
        bytes
    }

    fn root(next: u32, entries: &[Vec<u8>]) -> Vec<u8> {
        let exact = varint_len(1, u64::try_from(FORMULA_LIST_TYPE).unwrap())
            + varint_len(NEXT_KEY_FIELD, u64::from(next))
            + entries
                .iter()
                .map(|entry| length_field_len(ENTRY_FIELD, entry.len()).unwrap())
                .sum::<usize>();
        let mut bytes = exact_vec(exact).unwrap();
        append_varint(&mut bytes, 1, u64::try_from(FORMULA_LIST_TYPE).unwrap());
        append_varint(&mut bytes, NEXT_KEY_FIELD, u64::from(next));
        for entry in entries {
            append_length(&mut bytes, ENTRY_FIELD, entry).unwrap();
        }
        bytes
    }

    fn reference(identifier: u64) -> Vec<u8> {
        let exact = varint_len(1, identifier);
        let mut bytes = exact_vec(exact).unwrap();
        append_varint(&mut bytes, 1, identifier);
        bytes
    }

    fn root_with_segment(next: u32, entries: &[Vec<u8>], segment: u64) -> Vec<u8> {
        let reference = reference(segment);
        let mut bytes = root(next, entries);
        let extra = length_field_len(4, reference.len()).unwrap() + varint_len(90, 7);
        bytes.try_reserve_exact(extra).unwrap();
        append_varint(&mut bytes, 90, 7);
        append_length(&mut bytes, 4, &reference).unwrap();
        bytes
    }

    fn segment(location: u32, length: u32, entries: &[Vec<u8>]) -> Vec<u8> {
        let range_len = varint_len(1, u64::from(location)) + varint_len(2, u64::from(length));
        let mut range = exact_vec(range_len).unwrap();
        append_varint(&mut range, 1, u64::from(location));
        append_varint(&mut range, 2, u64::from(length));
        let exact = varint_len(1, u64::try_from(FORMULA_LIST_TYPE).unwrap())
            + length_field_len(2, range.len()).unwrap()
            + varint_len(91, 9)
            + entries
                .iter()
                .map(|entry| length_field_len(ENTRY_FIELD, entry.len()).unwrap())
                .sum::<usize>();
        let mut bytes = exact_vec(exact).unwrap();
        append_varint(&mut bytes, 1, u64::try_from(FORMULA_LIST_TYPE).unwrap());
        append_length(&mut bytes, 2, &range).unwrap();
        append_varint(&mut bytes, 91, 9);
        for entry in entries {
            append_length(&mut bytes, ENTRY_FIELD, entry).unwrap();
        }
        bytes
    }

    fn limits(input: usize) -> Limits {
        let output = input.saturating_mul(16).saturating_add(1_024);
        Limits {
            max_input_bytes: input,
            max_output_bytes: output,
            max_fields: input.saturating_add(output).saturating_mul(8),
            max_work: input
                .saturating_add(output)
                .saturating_mul(128)
                .saturating_add(10_000),
            max_entries: 64,
            max_retained_elements: 512,
            max_hosts: 64,
            max_references: input.saturating_mul(4).saturating_add(128),
            max_retained_bytes: input.saturating_mul(32).saturating_add(10_000),
            max_scratch_bytes: input.saturating_mul(16).saturating_add(10_000),
            max_allocations: input.saturating_add(64),
        }
    }

    fn scale_limits(input: usize, hosts: usize, formula_bytes: usize) -> Limits {
        let output = formula_bytes
            .saturating_mul(hosts)
            .saturating_add(input)
            .saturating_add(hosts.saturating_mul(64));
        Limits {
            max_input_bytes: input,
            max_output_bytes: output,
            max_fields: input.saturating_add(output).saturating_mul(8),
            max_work: usize::MAX / 4,
            max_entries: hosts.saturating_add(8),
            max_retained_elements: hosts.saturating_mul(8).saturating_add(64),
            max_hosts: hosts,
            max_references: input.saturating_mul(3).saturating_add(64),
            max_retained_bytes: formula_bytes
                .saturating_mul(hosts)
                .saturating_add(hosts.saturating_mul(256))
                .saturating_add(input.saturating_mul(8)),
            max_scratch_bytes: formula_bytes
                .saturating_mul(hosts)
                .saturating_add(hosts.saturating_mul(256))
                .saturating_add(input.saturating_mul(16)),
            max_allocations: input
                .saturating_add(hosts.saturating_mul(4))
                .saturating_add(128),
        }
    }

    #[test]
    fn exact_noop_and_wrong_coordinate_key_are_authoritative() {
        let ast = formula(17);
        let root_bytes = root(2, &[entry(1, 1, &ast)]);
        let source = SourceList {
            root: SourceMessage {
                object_id: 41,
                payload: &root_bytes,
                object_references: &[],
            },
            segments: &[],
            expected_entries: 1,
            source_hosts: &[SourceHost {
                row: 2,
                column: 2,
                key: 1,
            }],
        };
        let delta = HostDelta {
            row: 2,
            column: 2,
            old_formula_key: Some(1),
            new_formula: Some(&ast),
        };
        let artifact = plan(source, &[delta], limits(root_bytes.len())).unwrap();
        assert_eq!(artifact.root.payload, None);
        assert_eq!(
            artifact.assignments,
            [Assignment {
                delta: 0,
                key: Some(1)
            }]
        );
        assert_eq!(artifact.report.changed_messages, 0);
        assert_eq!(artifact.report.entries, 1);
        assert_eq!(artifact.report.hosts, 1);
        assert!(artifact.report.allocations != 0);
        assert!(artifact.report.peak_scratch_bytes != 0);

        let wrong = HostDelta { row: 3, ..delta };
        assert!(matches!(
            plan(source, &[wrong], limits(root_bytes.len())),
            Err(Error::InvalidSource)
        ));

        let occupied_next = root(1, &[entry(1, 1, &ast)]);
        assert!(matches!(
            plan(
                SourceList {
                    root: SourceMessage {
                        payload: &occupied_next,
                        ..source.root
                    },
                    segments: &[],
                    expected_entries: 1,
                    source_hosts: source.source_hosts,
                },
                &[],
                limits(occupied_next.len())
            ),
            Err(Error::InvalidSource)
        ));

        let later = formula(19);
        let later_collision = root(2, &[entry(1, 1, &ast), entry(3, 1, &later)]);
        let later_hosts = [
            SourceHost {
                row: 2,
                column: 2,
                key: 1,
            },
            SourceHost {
                row: 2,
                column: 3,
                key: 3,
            },
        ];
        let new_a = formula(20);
        let new_b = formula(21);
        let later_deltas = [
            HostDelta {
                row: 3,
                column: 2,
                old_formula_key: None,
                new_formula: Some(&new_a),
            },
            HostDelta {
                row: 3,
                column: 3,
                old_formula_key: None,
                new_formula: Some(&new_b),
            },
        ];
        assert!(matches!(
            plan(
                SourceList {
                    root: SourceMessage {
                        object_id: 41,
                        payload: &later_collision,
                        object_references: &[],
                    },
                    segments: &[],
                    expected_entries: 2,
                    source_hosts: &later_hosts,
                },
                &later_deltas,
                limits(later_collision.len())
            ),
            Err(Error::InvalidSource)
        ));

        let empty_segment = segment(10, 0, &[]);
        let empty_segments = [SourceMessage {
            object_id: 99,
            payload: &empty_segment,
            object_references: &[],
        }];
        let empty_root = root_with_segment(2, &[entry(1, 1, &ast)], 99);
        assert!(matches!(
            plan(
                SourceList {
                    root: SourceMessage {
                        object_id: 41,
                        payload: &empty_root,
                        object_references: &[99],
                    },
                    segments: &empty_segments,
                    expected_entries: 1,
                    source_hosts: source.source_hosts,
                },
                &[],
                limits(empty_root.len() + empty_segment.len())
            ),
            Err(Error::InvalidSource)
        ));
    }

    #[test]
    fn insertion_is_deterministic_and_limits_preempt_allocation() {
        let old = formula(17);
        let new = formula(18);
        let root = root(2, &[entry(1, 1, &old)]);
        let source = SourceList {
            root: SourceMessage {
                object_id: 41,
                payload: &root,
                object_references: &[],
            },
            segments: &[],
            expected_entries: 1,
            source_hosts: &[SourceHost {
                row: 2,
                column: 2,
                key: 1,
            }],
        };
        let delta = HostDelta {
            row: 3,
            column: 2,
            old_formula_key: None,
            new_formula: Some(&new),
        };
        let artifact = plan(source, &[delta], limits(root.len())).unwrap();
        assert_eq!(artifact.assignments[0].key, Some(2));
        assert_eq!(artifact.final_entries.len(), 2);

        let mut refused = limits(root.len());
        refused.max_allocations = artifact.report.allocations - 1;
        assert!(matches!(
            plan(source, &[delta], refused),
            Err(Error::Limit { .. })
        ));
        let mut refused = limits(root.len());
        refused.max_scratch_bytes = artifact.report.peak_scratch_bytes - 1;
        assert!(matches!(
            plan(source, &[delta], refused),
            Err(Error::Limit { .. })
        ));
    }

    #[test]
    fn segmented_reuse_preserves_routes_unknown_order_and_refcounts() {
        let released = formula(17);
        let reused = formula(18);
        let root = root_with_segment(11, &[entry(1, 1, &released)], 99);
        let segment = segment(10, 10, &[entry(10, 1, &reused)]);
        let segments = [SourceMessage {
            object_id: 99,
            payload: &segment,
            object_references: &[],
        }];
        let hosts = [
            SourceHost {
                row: 2,
                column: 2,
                key: 1,
            },
            SourceHost {
                row: 2,
                column: 3,
                key: 10,
            },
        ];
        let source = SourceList {
            root: SourceMessage {
                object_id: 41,
                payload: &root,
                object_references: &[99],
            },
            segments: &segments,
            expected_entries: 2,
            source_hosts: &hosts,
        };
        let delta = HostDelta {
            row: 2,
            column: 2,
            old_formula_key: Some(1),
            new_formula: Some(&reused),
        };
        let artifact = plan(source, &[delta], limits(root.len() + segment.len())).unwrap();
        assert_eq!(artifact.assignments[0].key, Some(10));
        assert_eq!(artifact.final_entries.len(), 1);
        assert_eq!(artifact.final_entries[0].key, 10);
        assert_eq!(artifact.final_entries[0].ref_count, 2);
        assert_eq!(artifact.root.object_references, [99]);
        assert!(artifact.segments[0].object_references.is_empty());
        let root_candidate = artifact.root.payload.as_deref().unwrap();
        let segment_candidate = artifact.segments[0].payload.as_deref().unwrap();
        assert!(contains_field(root_candidate, 90, limits(root.len() + segment.len())).unwrap());
        assert!(contains_field(segment_candidate, 91, limits(root.len() + segment.len())).unwrap());
        let root_unknown = root_candidate
            .windows(2)
            .position(|pair| pair == [0xd0, 0x05])
            .unwrap();
        let root_route = root_candidate
            .windows(1)
            .position(|byte| byte == [0x22])
            .unwrap();
        assert!(root_unknown < root_route);

        let mut nested = segment.clone();
        let nested_ref = reference(123);
        nested
            .try_reserve_exact(length_field_len(4, nested_ref.len()).unwrap())
            .unwrap();
        append_length(&mut nested, 4, &nested_ref).unwrap();
        let nested_segments = [SourceMessage {
            object_id: 99,
            payload: &nested,
            object_references: &[],
        }];
        assert!(matches!(
            plan(
                SourceList {
                    segments: &nested_segments,
                    ..source
                },
                &[delta],
                limits(root.len() + nested.len())
            ),
            Err(Error::InvalidSource)
        ));
    }

    #[test]
    fn governed_4096_to_8192_scaling_and_every_axis_max_minus_one() {
        fn run(hosts: usize) -> (Artifact, Requirements, Limits) {
            let ast = formula(17);
            let root = root(1, &[]);
            let deltas = (0..hosts)
                .map(|row| HostDelta {
                    row: u32::try_from(row).unwrap(),
                    column: 1,
                    old_formula_key: None,
                    new_formula: Some(ast.as_slice()),
                })
                .collect::<Vec<_>>();
            let source = SourceList {
                root: SourceMessage {
                    object_id: 41,
                    payload: &root,
                    object_references: &[],
                },
                segments: &[],
                expected_entries: 0,
                source_hosts: &[],
            };
            let limits = scale_limits(root.len(), hosts, ast.len());
            let requirements = preflight(source, &deltas, limits).unwrap();
            let artifact = plan(source, &deltas, limits).unwrap();
            (artifact, requirements, limits)
        }

        let (small, _, _) = run(4096);
        let (large, requirements, limits) = run(8192);
        for (small, large) in [
            (small.report.work, large.report.work),
            (small.report.allocations, large.report.allocations),
            (small.report.retained_bytes, large.report.retained_bytes),
            (
                small.report.peak_scratch_bytes,
                large.report.peak_scratch_bytes,
            ),
            (small.report.output_bytes, large.report.output_bytes),
        ] {
            assert!(large.saturating_mul(10) <= small.saturating_mul(22));
        }

        let ast = formula(17);
        let root = root(1, &[]);
        let deltas = (0..8192)
            .map(|row| HostDelta {
                row: u32::try_from(row).unwrap(),
                column: 1,
                old_formula_key: None,
                new_formula: Some(ast.as_slice()),
            })
            .collect::<Vec<_>>();
        let source = SourceList {
            root: SourceMessage {
                object_id: 41,
                payload: &root,
                object_references: &[],
            },
            segments: &[],
            expected_entries: 0,
            source_hosts: &[],
        };
        let mut cases = Vec::new();
        for axis in 0..9 {
            let mut refused = limits;
            match axis {
                0 => refused.max_input_bytes = requirements.input_bytes - 1,
                1 => refused.max_fields = requirements.fields - 1,
                2 => refused.max_work = requirements.work - 1,
                3 => refused.max_allocations = requirements.allocations - 1,
                4 => refused.max_scratch_bytes = requirements.scratch_bytes - 1,
                5 => refused.max_retained_bytes = requirements.retained_upper - 1,
                6 => refused.max_hosts = requirements.formula_occurrences - 1,
                7 => refused.max_entries = requirements.entries - 1,
                8 => refused.max_references = requirements.references - 1,
                _ => unreachable!(),
            }
            cases.push(refused);
        }
        for refused in cases {
            assert!(matches!(
                plan(source, &deltas, refused),
                Err(Error::Limit { .. })
            ));
        }
        let mut work_refused = limits;
        work_refused.max_work = requirements.work - 1;
        assert!(matches!(
            plan(source, &deltas, work_refused),
            Err(Error::Limit { .. })
        ));
        let mut output_refused = limits;
        output_refused.max_output_bytes = requirements
            .input_bytes
            .checked_add(ast.len().checked_mul(8192).unwrap())
            .and_then(|bytes| bytes.checked_add(8192 * 32))
            .unwrap()
            - 1;
        assert!(matches!(
            plan(source, &deltas, output_refused),
            Err(Error::Limit { .. })
        ));
        assert_eq!(large.assignments.len(), 8192);
        assert_eq!(large.final_entries[0].ref_count, 8192);
    }

    #[test]
    fn logical_barrier_is_output_free_scaled_and_execution_preempts_output() {
        fn reports(hosts: usize) -> (LogicalReport, ExecutionRequirements) {
            let ast = formula(17);
            let root = root(1, &[]);
            let deltas = (0..hosts)
                .map(|row| HostDelta {
                    row: u32::try_from(row).unwrap(),
                    column: 1,
                    old_formula_key: None,
                    new_formula: Some(ast.as_slice()),
                })
                .collect::<Vec<_>>();
            let source = SourceList {
                root: SourceMessage {
                    object_id: 41,
                    payload: &root,
                    object_references: &[],
                },
                segments: &[],
                expected_entries: 0,
                source_hosts: &[],
            };
            let logical = prepare(source, &deltas, scale_limits(root.len(), hosts, ast.len()))
                .unwrap()
                .logical()
                .unwrap();
            let report = logical.prepare_report();
            let execution = logical.execution_requirements();
            assert_eq!(report.output_bytes, 0);
            (report, execution)
        }

        let (small_report, small_execution) = reports(4096);
        let (large_report, large_execution) = reports(8192);
        for (small, large) in [
            (small_report.work, large_report.work),
            (small_report.allocations, large_report.allocations),
            (
                small_report.retained_elements,
                large_report.retained_elements,
            ),
            (small_report.retained_bytes, large_report.retained_bytes),
            (
                small_report.peak_scratch_bytes,
                large_report.peak_scratch_bytes,
            ),
            (small_execution.work, large_execution.work),
            (small_execution.allocations, large_execution.allocations),
            (
                small_execution.retained_elements,
                large_execution.retained_elements,
            ),
            (
                small_execution.retained_bytes,
                large_execution.retained_bytes,
            ),
            (
                small_execution.peak_scratch_bytes,
                large_execution.peak_scratch_bytes,
            ),
            (small_execution.output_bytes, large_execution.output_bytes),
        ] {
            assert!(large.saturating_mul(10) <= small.saturating_mul(22));
        }

        let ast = formula(17);
        let root = root(1, &[]);
        let deltas = [HostDelta {
            row: 2,
            column: 1,
            old_formula_key: None,
            new_formula: Some(ast.as_slice()),
        }];
        let source = SourceList {
            root: SourceMessage {
                object_id: 41,
                payload: &root,
                object_references: &[],
            },
            segments: &[],
            expected_entries: 0,
            source_hosts: &[],
        };
        let limits = scale_limits(root.len(), 1, ast.len());
        let requirements = preflight(source, &deltas, limits).unwrap();
        let mut logical_refused = limits;
        logical_refused.max_scratch_bytes = requirements.logical_scratch_bytes - 1;
        assert!(matches!(
            prepare(source, &deltas, logical_refused),
            Err(Error::Limit {
                resource: Resource::ScratchBytes,
                ..
            })
        ));

        let logical = || prepare(source, &deltas, limits).unwrap().logical().unwrap();
        let report = logical().prepare_report();
        for axis in 0..8 {
            let mut refused = limits;
            let expected = match axis {
                0 => {
                    refused.max_input_bytes = report.input_bytes - 1;
                    Resource::InputBytes
                },
                1 => {
                    refused.max_fields = report.fields - 1;
                    Resource::Fields
                },
                2 => {
                    refused.max_work = report.work - 1;
                    Resource::Work
                },
                3 => {
                    refused.max_references = report.references - 1;
                    Resource::References
                },
                4 => {
                    refused.max_retained_elements = report.retained_elements - 1;
                    Resource::RetainedElements
                },
                5 => {
                    refused.max_retained_bytes = report.retained_bytes - 1;
                    Resource::RetainedBytes
                },
                6 => {
                    refused.max_scratch_bytes = report.peak_scratch_bytes - 1;
                    Resource::ScratchBytes
                },
                7 => {
                    refused.max_allocations = report.allocations - 1;
                    Resource::Allocations
                },
                _ => unreachable!(),
            };
            assert!(matches!(
                prepare(source, &deltas, refused),
                Err(Error::Limit { resource, .. }) if resource == expected
            ));
        }

        let execution = logical().execution_requirements();
        let accepted_execution = Limits {
            max_input_bytes: limits.max_input_bytes.max(execution.input_bytes),
            max_output_bytes: limits.max_output_bytes.max(execution.output_bytes),
            max_fields: limits.max_fields.max(execution.fields),
            max_work: limits.max_work.max(execution.work),
            max_references: limits.max_references.max(execution.references),
            max_retained_elements: limits
                .max_retained_elements
                .max(execution.retained_elements),
            max_retained_bytes: limits.max_retained_bytes.max(execution.retained_bytes),
            max_scratch_bytes: limits.max_scratch_bytes.max(execution.peak_scratch_bytes),
            max_allocations: limits.max_allocations.max(execution.allocations),
            ..limits
        };
        for axis in 0..9 {
            let mut refused = accepted_execution;
            let expected = match axis {
                0 => {
                    refused.max_input_bytes = execution.input_bytes - 1;
                    Resource::InputBytes
                },
                1 => {
                    refused.max_fields = execution.fields - 1;
                    Resource::Fields
                },
                2 => {
                    refused.max_work = execution.work - 1;
                    Resource::Work
                },
                3 => {
                    refused.max_references = execution.references - 1;
                    Resource::References
                },
                4 => {
                    refused.max_output_bytes = execution.output_bytes - 1;
                    Resource::OutputBytes
                },
                5 => {
                    refused.max_retained_elements = execution.retained_elements - 1;
                    Resource::RetainedElements
                },
                6 => {
                    refused.max_retained_bytes = execution.retained_bytes - 1;
                    Resource::RetainedBytes
                },
                7 => {
                    refused.max_scratch_bytes = execution.peak_scratch_bytes - 1;
                    Resource::ScratchBytes
                },
                8 => {
                    refused.max_allocations = execution.allocations - 1;
                    Resource::Allocations
                },
                _ => unreachable!(),
            };
            let result = logical().execute(refused);
            assert!(
                matches!(&result, Err(Error::Limit { resource, .. }) if *resource == expected),
                "execution axis {axis} returned {result:?}"
            );
        }
    }
}
