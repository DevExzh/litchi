//! Checked aggregate accounting for one compact table-cell transaction.
//!
//! Publication is deliberately two phase.  A conservative directional cost
//! is checked and retained privately before the raw writer allocates its
//! output or reopens a candidate.  Only the infallible completion step moves
//! exact observed counters into [`Usage`].

use crate::table::cells::{Error, LimitKind, Path};

use std::mem::size_of;

use super::{
    super::Package,
    rewrite::{PublicationCost, PublicationReservation, ReopenCost},
};

const RETAINED_ELEMENTS_PER_UPDATE: u64 = 8;
const RETAINED_BYTES_PER_UPDATE: u64 = 192;
const SCRATCH_BYTES_PER_UPDATE: u64 = 128;
const SCRATCH_BYTES_PER_OUTPUT_BYTE: u64 = 8;
const ALLOCATIONS_PER_UPDATE: u64 = 4;
const WIRE_WORK_MULTIPLIER: u64 = 32;
const WIRE_FIELD_MULTIPLIER: u64 = 16;
const WIRE_WORK_PER_UPDATE: u64 = 64;
// Formula publication strictly visits the same bounded ArchiveInfo graph in
// resolver, list, metadata-plan, cache-plan, metadata-execute, and reopen
// phases. Keep those cumulative visits finite and independently governed.
const REFERENCE_MULTIPLIER: u64 = 64;
// Formula authoring runs strict list, dependency, evaluator, tile and reopen
// phases. Each cell may contribute both an AST node/edge visit and a physical
// cache transition in those independently reported phases.
const FORMULA_WORK_PER_UPDATE: u64 = 128;
const REOPEN_WORK_MULTIPLIER: u64 = 64;
// The allocation-free locality proof may compare the complete source and
// candidate artifacts through sixteen bounded ZIP/topology views. Both are
// individually bounded by `max_input_bytes`, so the transaction-wide ceiling
// must admit 16 * (source + candidate) = 32 work units per output byte.
const LOCALITY_WORK_PER_OUTPUT_BYTE: u64 = 32;

/// Finite independent ceilings for one changed transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct TransactionLimits {
    pub(super) max_updates: u64,
    pub(super) max_owned_value_bytes: u64,
    pub(super) max_retained_elements: u64,
    pub(super) max_retained_bytes: u64,
    pub(super) max_scratch_bytes: u64,
    pub(super) max_allocation_events: u64,
    pub(super) max_wire_bytes: u64,
    pub(super) max_wire_fields: u64,
    pub(super) max_wire_work: u64,
    pub(super) max_objects: u64,
    pub(super) max_references: u64,
    pub(super) max_formula_work: u64,
    pub(super) max_output_bytes: u64,
    pub(super) max_reopen_work: u64,
    pub(super) max_transaction_work: u64,
}

#[cfg(test)]
pub(super) mod testing {
    use std::cell::RefCell;

    use super::{TransactionLimits, Usage};

    #[derive(Default)]
    struct Observation {
        limits: Option<TransactionLimits>,
        last_usage: Option<Usage>,
    }

    std::thread_local! {
        static OBSERVATION: RefCell<Observation> = RefCell::new(Observation::default());
    }

    struct Reset(Option<TransactionLimits>);

    impl Drop for Reset {
        fn drop(&mut self) {
            OBSERVATION.with(|slot| slot.borrow_mut().limits = self.0.take());
        }
    }

    /// Run one rooted production transaction with an optional private limit
    /// override, returning the exact usage captured when its stack-local
    /// `TransactionBudget` is dropped.
    pub(crate) fn observe<T>(
        limits: Option<TransactionLimits>,
        operation: impl FnOnce() -> T,
    ) -> (T, Option<Usage>) {
        let previous = OBSERVATION.with(|slot| {
            let mut observation = slot.borrow_mut();
            observation.last_usage = None;
            core::mem::replace(&mut observation.limits, limits)
        });
        let reset = Reset(previous);
        let value = operation();
        let usage = OBSERVATION.with(|slot| slot.borrow_mut().last_usage.take());
        drop(reset);
        (value, usage)
    }

    pub(crate) fn package_limits(source: &crate::Package) -> TransactionLimits {
        TransactionLimits::from_package(source).expect("test package limits must be representable")
    }

    pub(super) fn limit_override() -> Option<TransactionLimits> {
        OBSERVATION.with(|slot| slot.borrow().limits)
    }

    pub(super) fn record(usage: Usage) {
        OBSERVATION.with(|slot| slot.borrow_mut().last_usage = Some(usage));
    }
}

impl TransactionLimits {
    /// Derive finite transaction ceilings from the package's selected ingress
    /// and semantic limits.  No derived ceiling saturates on overflow.
    pub(super) fn from_package(source: &Package) -> Result<Self, Error> {
        let archive = source.state.options.archive();
        let semantic = source.state.options.semantic();
        let updates = as_u64(semantic.max_materialized_cells(), LimitKind::Updates)?;
        let owned = as_u64(semantic.max_output_text_bytes(), LimitKind::OwnedValueBytes)?;
        let wire = as_u64(archive.max_iwa_stream_bytes(), LimitKind::WireWork)?;
        let output = archive.max_input_bytes();
        let objects = as_u64(semantic.max_objects(), LimitKind::Objects)?;
        let references = as_u64(semantic.max_references(), LimitKind::References)?;

        let max_retained_elements = checked_mul_limit(
            updates,
            RETAINED_ELEMENTS_PER_UPDATE,
            LimitKind::RetainedElements,
        )?;
        let max_retained_bytes = checked_add_limit(
            checked_mul_limit(updates, RETAINED_BYTES_PER_UPDATE, LimitKind::RetainedBytes)?,
            owned,
            LimitKind::RetainedBytes,
        )?;
        let max_scratch_bytes = checked_add_limit(
            checked_mul_limit(
                updates,
                SCRATCH_BYTES_PER_UPDATE,
                LimitKind::PeakScratchBytes,
            )?,
            checked_add_limit(
                wire,
                checked_mul_limit(
                    output,
                    SCRATCH_BYTES_PER_OUTPUT_BYTE,
                    LimitKind::PeakScratchBytes,
                )?,
                LimitKind::PeakScratchBytes,
            )?,
            LimitKind::PeakScratchBytes,
        )?;
        let max_allocation_events = checked_add_limit(
            checked_mul_limit(updates, ALLOCATIONS_PER_UPDATE, LimitKind::TransactionWork)?,
            objects,
            LimitKind::TransactionWork,
        )?;
        let max_wire_bytes = checked_mul_limit(wire, WIRE_WORK_MULTIPLIER, LimitKind::WireWork)?;
        let max_wire_fields =
            checked_mul_limit(wire, WIRE_FIELD_MULTIPLIER, LimitKind::WireFields)?;
        let max_wire_work = checked_add_limit(
            checked_mul_limit(wire, WIRE_WORK_MULTIPLIER, LimitKind::WireWork)?,
            checked_mul_limit(updates, WIRE_WORK_PER_UPDATE, LimitKind::WireWork)?,
            LimitKind::WireWork,
        )?;
        let max_references =
            checked_mul_limit(references, REFERENCE_MULTIPLIER, LimitKind::References)?;
        let max_formula_work = checked_add_limit(
            checked_mul_limit(updates, FORMULA_WORK_PER_UPDATE, LimitKind::FormulaWork)?,
            as_u64(semantic.max_formula_render_work(), LimitKind::FormulaWork)?,
            LimitKind::FormulaWork,
        )?;
        let max_reopen_work =
            checked_mul_limit(output, REOPEN_WORK_MULTIPLIER, LimitKind::ReopenWork)?;

        // Transaction work is an independent aggregate ceiling, not an
        // alias for any one leaf limit.  Its construction is checked at every
        // addition so a platform cannot silently acquire an unlimited budget.
        let max_transaction_work = [
            max_retained_elements,
            max_retained_bytes,
            max_scratch_bytes,
            max_allocation_events,
            max_wire_bytes,
            max_wire_fields,
            max_wire_work,
            objects,
            max_references,
            max_formula_work,
            output,
            max_reopen_work,
            checked_mul_limit(
                output,
                LOCALITY_WORK_PER_OUTPUT_BYTE,
                LimitKind::TransactionWork,
            )?,
        ]
        .into_iter()
        .try_fold(0u64, |total, value| {
            checked_add_limit(total, value, LimitKind::TransactionWork)
        })?;

        Ok(Self {
            max_updates: updates,
            max_owned_value_bytes: owned,
            max_retained_elements,
            max_retained_bytes,
            max_scratch_bytes,
            max_allocation_events,
            max_wire_bytes,
            max_wire_fields,
            max_wire_work,
            max_objects: objects,
            max_references,
            max_formula_work,
            max_output_bytes: output,
            max_reopen_work,
            max_transaction_work,
        })
    }
}

/// Exact observed counters for one transaction.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct Usage {
    pub(super) updates: u64,
    pub(super) input_value_bytes: u64,
    pub(super) retained_elements: u64,
    pub(super) retained_bytes: u64,
    pub(super) scratch_bytes: u64,
    pub(super) peak_scratch_bytes: u64,
    pub(super) allocation_events: u64,
    pub(super) wire_bytes: u64,
    pub(super) wire_fields: u64,
    pub(super) wire_work: u64,
    pub(super) objects: u64,
    pub(super) references: u64,
    pub(super) lookups: u64,
    pub(super) tile_reads: u64,
    pub(super) tile_writes: u64,
    pub(super) header_reads: u64,
    pub(super) header_writes: u64,
    pub(super) row_reads: u64,
    pub(super) row_writes: u64,
    pub(super) list_reads: u64,
    pub(super) list_writes: u64,
    pub(super) string_work: u64,
    pub(super) rich_text_work: u64,
    pub(super) formula_graph_builds: u64,
    pub(super) formula_nodes: u64,
    pub(super) formula_edges: u64,
    pub(super) range_candidates: u64,
    pub(super) cache_hosts: u64,
    pub(super) authored_formula_writes: u64,
    pub(super) formula_work: u64,
    pub(super) component_encodes: u64,
    pub(super) components_reassembled: u64,
    pub(super) reassembly_bytes: u64,
    pub(super) preview_bytes_deleted: u64,
    pub(super) output_artifact_allocations: u64,
    pub(super) output_bytes: u64,
    pub(super) candidate_reopens: u64,
    pub(super) reopen_references: u64,
    pub(super) reopen_work: u64,
    pub(super) locality_bytes: u64,
    pub(super) transaction_work: u64,
}

/// A set of counters reserved atomically.
pub(super) type UsageDelta = Usage;
/// Exact capacity still available before entering the next leaf.
pub(super) type Remaining = Usage;

/// Tightest tile-local work ceiling whose eventual transaction report is
/// guaranteed to fit `remaining`.
///
/// Tile wire work already charges every scanned slot and the output bytes.
/// The transaction adapter deliberately reports those two observations again,
/// plus written slots.  Therefore `transaction_work <= 2 * wire_work + writes`.
pub(super) fn tile_work_ceiling(
    remaining: Remaining,
    package_maximum: u64,
    maximum_writes: u64,
) -> Result<u64, Error> {
    let available = remaining
        .transaction_work
        .checked_sub(maximum_writes)
        .ok_or_else(|| {
            limit_error(
                LimitKind::TransactionWork,
                maximum_writes,
                remaining.transaction_work,
            )
        })?;
    Ok(package_maximum
        .min(remaining.wire_work)
        .min(remaining.peak_scratch_bytes)
        .min(available / 2))
}

/// Exact retained ownership for a `Vec<T>` moved into `Arc<Vec<T>>`.
///
/// The payload is charged by capacity rather than length because
/// `try_reserve_exact` may still return extra capacity. The Arc allocation
/// retains its `Vec` header plus the strong/weak control words; the Vec's
/// backing allocation is a second event when capacity is nonzero.
pub(super) fn arc_vec_retained_usage<T>(length: usize, capacity: usize) -> Result<Usage, Error> {
    if length > capacity {
        return Err(limit_error(
            LimitKind::RetainedElements,
            as_u64(length, LimitKind::RetainedElements)?,
            as_u64(capacity, LimitKind::RetainedElements)?,
        ));
    }
    let payload_bytes = capacity
        .checked_mul(size_of::<T>())
        .ok_or_else(|| limit_error(LimitKind::RetainedBytes, u64::MAX, u64::MAX - 1))?;
    let allocation_bytes = payload_bytes
        .checked_add(size_of::<Vec<T>>())
        .and_then(|bytes| bytes.checked_add(size_of::<usize>().checked_mul(2)?))
        .ok_or_else(|| limit_error(LimitKind::RetainedBytes, u64::MAX, u64::MAX - 1))?;
    Ok(Usage {
        retained_elements: as_u64(length, LimitKind::RetainedElements)?,
        retained_bytes: as_u64(allocation_bytes, LimitKind::RetainedBytes)?,
        allocation_events: 1 + u64::from(capacity != 0),
        transaction_work: as_u64(length, LimitKind::TransactionWork)?,
        ..Usage::default()
    })
}

#[derive(Debug, Clone, Copy)]
struct PendingPublication {
    directional: PublicationReservation,
    reserved: Usage,
}

#[derive(Debug, Clone, Copy)]
struct PendingAuthorization {
    baseline: Usage,
    reserved: Usage,
    available: Usage,
}

/// Aggregate budget and exact usage for one compact batch.
#[derive(Debug)]
pub(super) struct TransactionBudget {
    limits: TransactionLimits,
    usage: Usage,
    pending_authorization: Option<PendingAuthorization>,
    pending_publication: Option<PendingPublication>,
}

impl TransactionBudget {
    pub(super) fn new(source: &Package) -> Result<Self, Error> {
        #[cfg(test)]
        if let Some(limits) = testing::limit_override() {
            return Self::from_limits(limits);
        }
        Self::from_limits(TransactionLimits::from_package(source)?)
    }

    pub(super) fn from_limits(limits: TransactionLimits) -> Result<Self, Error> {
        Ok(Self {
            limits,
            usage: Usage::default(),
            pending_authorization: None,
            pending_publication: None,
        })
    }

    pub(super) const fn limits(&self) -> TransactionLimits {
        self.limits
    }

    #[cfg(test)]
    pub(super) fn required_limits_for(&self, envelope: Usage) -> Result<TransactionLimits, Error> {
        // This helper is an observation hook: it reports the exact ceiling
        // required by the current usage plus a prepared envelope.  Deriving
        // that ceiling must not itself be constrained by the deliberately
        // reduced test limit, otherwise a max-minus-one run cannot observe
        // the same requirement that it is meant to refuse.
        let unlimited = TransactionLimits {
            max_updates: u64::MAX,
            max_owned_value_bytes: u64::MAX,
            max_retained_elements: u64::MAX,
            max_retained_bytes: u64::MAX,
            max_scratch_bytes: u64::MAX,
            max_allocation_events: u64::MAX,
            max_wire_bytes: u64::MAX,
            max_wire_fields: u64::MAX,
            max_wire_work: u64::MAX,
            max_objects: u64::MAX,
            max_references: u64::MAX,
            max_formula_work: u64::MAX,
            max_output_bytes: u64::MAX,
            max_reopen_work: u64::MAX,
            max_transaction_work: u64::MAX,
        };
        let candidate = add_usage(self.usage, envelope, unlimited)?;
        let max_formula_work = [
            candidate.formula_graph_builds,
            candidate.formula_nodes,
            candidate.formula_edges,
            candidate.range_candidates,
            candidate.cache_hosts,
            candidate.authored_formula_writes,
            candidate.formula_work,
        ]
        .into_iter()
        .max()
        .unwrap_or(0);
        let max_transaction_work = [
            candidate.lookups,
            candidate.tile_reads,
            candidate.tile_writes,
            candidate.header_reads,
            candidate.header_writes,
            candidate.row_reads,
            candidate.row_writes,
            candidate.list_reads,
            candidate.list_writes,
            candidate.string_work,
            candidate.rich_text_work,
            candidate.components_reassembled,
            candidate.reassembly_bytes,
            candidate.preview_bytes_deleted,
            candidate.output_artifact_allocations,
            candidate.candidate_reopens,
            candidate.locality_bytes,
            candidate.transaction_work,
        ]
        .into_iter()
        .max()
        .unwrap_or(0);
        Ok(TransactionLimits {
            max_updates: candidate.updates,
            max_owned_value_bytes: candidate.input_value_bytes,
            max_retained_elements: candidate.retained_elements,
            max_retained_bytes: candidate.retained_bytes,
            max_scratch_bytes: candidate.peak_scratch_bytes.max(candidate.scratch_bytes),
            max_allocation_events: candidate.allocation_events,
            max_wire_bytes: candidate.wire_bytes,
            max_wire_fields: candidate.wire_fields,
            max_wire_work: candidate.wire_work,
            max_objects: candidate.objects.max(candidate.component_encodes),
            max_references: candidate.references.max(candidate.reopen_references),
            max_formula_work,
            max_output_bytes: candidate.output_bytes,
            max_reopen_work: candidate.reopen_work,
            max_transaction_work,
        })
    }

    /// Return exact per-counter capacity after observed and privately pending
    /// reservations.  A caller must snapshot this value, authorize an
    /// envelope no larger than it, and bind the same retained/scratch/wire/
    /// allocation/work ceilings into the leaf it is about to call.
    pub(super) fn remaining(&self) -> Result<Remaining, Error> {
        remaining_usage(self.effective_usage()?, self.limits)
    }

    /// Check every addition and every independent ceiling before assigning
    /// any counter.  A failed reservation leaves the complete usage unchanged.
    pub(super) fn reserve(&mut self, delta: UsageDelta) -> Result<(), Error> {
        let effective = self.effective_usage()?;
        let _authorized_candidate = add_usage(effective, delta, self.limits)?;
        let observed_candidate = add_usage(self.usage, delta, self.limits)?;
        self.usage = observed_candidate;
        Ok(())
    }

    /// Preauthorize one conservative leaf envelope without mutating observed
    /// usage.  The parent must call this before entering a resolver, tile,
    /// list, cache, rich-text, sparse, or grouped-writer callback.
    pub(super) fn authorize(&mut self, envelope: UsageDelta) -> Result<(), Error> {
        self.ensure_no_pending_authorization()?;
        let _candidate = add_usage(self.usage, envelope, self.limits)?;
        self.pending_authorization = Some(PendingAuthorization {
            baseline: self.usage,
            reserved: envelope,
            available: envelope,
        });
        Ok(())
    }

    /// Exercise one aggregate authorization against an independently
    /// selected exact ceiling without changing any earlier planning limits.
    /// This lets rooted tests prove max-minus-one preemption at the aggregate
    /// barrier itself instead of accidentally steering an earlier planner.
    #[cfg(test)]
    pub(super) fn authorize_under_limits(
        &mut self,
        envelope: UsageDelta,
        limits: TransactionLimits,
    ) -> Result<(), Error> {
        self.ensure_no_pending_authorization()?;
        let _candidate = add_usage(self.usage, envelope, limits)?;
        self.pending_authorization = Some(PendingAuthorization {
            baseline: self.usage,
            reserved: envelope,
            available: envelope,
        });
        Ok(())
    }

    /// Atomically merge exact leaf observations after a successful callback.
    /// This cannot limit-fail because `actual` must fit the admitted envelope.
    pub(super) fn record_authorized(&mut self, actual: Usage) -> Result<(), Error> {
        if self.pending_publication.is_some() {
            return Err(pending_error(self.limits));
        }
        let pending = self
            .pending_authorization
            .ok_or_else(|| pending_error(self.limits))?;
        validate_actual_le(actual, pending.available)?;
        let candidate = add_usage(self.usage, actual, self.limits)?;
        self.pending_authorization = None;
        self.usage = candidate;
        Ok(())
    }

    /// Cancel a leaf envelope after its callback returns an error.  Since the
    /// envelope was private, cancellation does not alter observed usage.
    pub(super) fn cancel_authorization(&mut self) {
        self.pending_authorization = None;
    }

    pub(super) const fn authorization_is_pending(&self) -> bool {
        self.pending_authorization.is_some()
    }

    /// Return the portion of the current leaf envelope not consumed by a
    /// nested publication reservation. A writer that can report component
    /// work separately must bind that work to this value before publication.
    pub(super) fn authorization_remaining(&self) -> Result<Remaining, Error> {
        self.pending_authorization
            .map(|pending| pending.available)
            .ok_or_else(|| pending_error(self.limits))
    }

    /// Admit the complete apply/reopen/locality envelope before entering the
    /// callback, then atomically replace that private envelope with the exact
    /// successful observation returned by the callback.
    ///
    /// `envelope` must cover Package reopen scratch and allocation events,
    /// reopened topology/reference/work counters, compact locality-plan
    /// staging, and the subsequent allocation-free locality proof. A refused
    /// envelope never invokes `operation`; a callback error cancels it without
    /// changing observed usage. The returned exact usage must include every
    /// retained element/byte and allocation created by Package reopen; the
    /// target payload itself is excluded when it is merely Arc-shared from the
    /// patch, but the new Package state and parsed projections are not.
    pub(super) fn with_apply_authorization<T>(
        &mut self,
        envelope: UsageDelta,
        operation: impl FnOnce() -> Result<(T, Usage), Error>,
    ) -> Result<T, Error> {
        self.authorize(envelope)?;
        let (value, actual) = match operation() {
            Ok(success) => success,
            Err(error) => {
                self.cancel_authorization();
                return Err(error);
            },
        };
        self.record_authorized(actual)?;
        Ok(value)
    }

    /// Atomically account for fallibly retained transaction storage.
    pub(super) fn reserve_retained(
        &mut self,
        elements: u64,
        bytes: u64,
        allocation_events: u64,
    ) -> Result<(), Error> {
        self.reserve(Usage {
            retained_elements: elements,
            retained_bytes: bytes,
            allocation_events,
            ..Usage::default()
        })
    }

    /// Release transaction-local retained plan storage after it has been
    /// consumed or abandoned. Allocation, work, and peak counters remain
    /// cumulative; only the live retained footprint is returned.
    pub(super) fn release_retained(&mut self, elements: u64, bytes: u64) -> Result<(), Error> {
        if self.pending_authorization.is_some() || self.pending_publication.is_some() {
            return Err(pending_error(self.limits));
        }
        let retained_elements = self
            .usage
            .retained_elements
            .checked_sub(elements)
            .ok_or_else(|| pending_error(self.limits))?;
        let retained_bytes = self
            .usage
            .retained_bytes
            .checked_sub(bytes)
            .ok_or_else(|| pending_error(self.limits))?;
        self.usage.retained_elements = retained_elements;
        self.usage.retained_bytes = retained_bytes;
        Ok(())
    }

    /// Enter a temporary allocation only after its peak and allocation event
    /// have both been admitted atomically.
    pub(super) fn reserve_scratch(
        &mut self,
        bytes: u64,
        allocation_events: u64,
    ) -> Result<(), Error> {
        self.reserve(Usage {
            scratch_bytes: bytes,
            allocation_events,
            ..Usage::default()
        })
    }

    /// Release temporary bytes.  Peak and allocation counters remain exact.
    pub(super) fn release_scratch(&mut self, bytes: u64) {
        self.usage.scratch_bytes = self
            .usage
            .scratch_bytes
            .checked_sub(bytes)
            .expect("scratch release must match an admitted reservation");
    }

    /// Run a fallible operation under a temporary scratch reservation.
    pub(super) fn with_scratch<T>(
        &mut self,
        bytes: u64,
        allocation_events: u64,
        operation: impl FnOnce(&mut Self) -> Result<T, Error>,
    ) -> Result<T, Error> {
        self.reserve_scratch(bytes, allocation_events)?;
        let result = operation(self);
        self.release_scratch(bytes);
        result
    }

    /// Reserve the complete conservative publication envelope without
    /// recording any output, reassembly, reopen, or locality event.
    pub(super) fn preauthorize_publication(
        &mut self,
        reservation: PublicationReservation,
    ) -> Result<(), Error> {
        if self.pending_publication.is_some() {
            return Err(pending_error(self.limits));
        }
        let reserved = publication_usage_reservation(reservation)?;
        // A grouped-writer leaf authorization may already be live.  The
        // publication reservation is nested inside it because the writer can
        // determine exact reassembly/reopen bounds only after component
        // preparation.  In that case it consumes a subset of the already
        // private writer envelope; adding it again would double-count the same
        // component/output capacity.  Without an outer writer envelope it is
        // checked as an independent reservation.
        if let Some(mut pending) = self.pending_authorization {
            pending.available = subtract_usage_envelope(pending.available, reserved)?;
            self.pending_authorization = Some(pending);
        } else {
            let _candidate = add_usage(self.usage, reserved, self.limits)?;
        }
        self.pending_publication = Some(PendingPublication {
            directional: reservation,
            reserved,
        });
        Ok(())
    }

    /// Record exact successful publication counters.  This is infallible: the
    /// writer may call it only after a successful preauthorization whose
    /// conservative envelope contains `actual`.
    pub(super) fn finish_publication(&mut self, actual: PublicationCost) {
        let pending = self
            .pending_publication
            .take()
            .expect("publication completion requires preauthorization");
        let exact = publication_usage_cost(actual)
            .expect("an admitted publication cost must remain representable");
        assert_publication_le(actual, pending.directional);
        assert_usage_le(exact, pending.reserved);
        // The outer grouped-writer report owns locality verification.  Once
        // the physical candidate has succeeded, return the unused portion of
        // the conservative publication envelope to that report.  In
        // particular, a publication with zero locality observations leaves
        // the outer authorisation able to settle the exact locality report
        // after its proof has completed.
        if let Some(mut authorization) = self.pending_authorization {
            authorization.available = subtract_usage_envelope(authorization.reserved, exact)
                .expect("exact publication must fit the outer authorization");
            self.pending_authorization = Some(authorization);
        }
        self.usage = add_usage(self.usage, exact, self.limits)
            .expect("exact publication must fit its admitted reservation");
    }

    /// Drop a conservative envelope when a post-precharge physical operation
    /// fails.  No observed publication event has occurred.
    pub(super) fn cancel_publication(&mut self) {
        if self.pending_publication.take().is_some() {
            if let Some(mut pending) = self.pending_authorization {
                pending.available = pending.reserved;
                self.pending_authorization = Some(pending);
            }
        }
    }

    pub(super) const fn publication_is_authorized(&self) -> bool {
        self.pending_publication.is_some()
    }

    fn effective_usage(&self) -> Result<Usage, Error> {
        match (self.pending_authorization, self.pending_publication) {
            // Publication is a checked subset of the grouped-writer envelope.
            (Some(leaf), _) => add_usage(leaf.baseline, leaf.reserved, self.limits),
            (None, Some(publication)) => add_usage(self.usage, publication.reserved, self.limits),
            (None, None) => Ok(self.usage),
        }
    }

    fn ensure_no_pending_authorization(&self) -> Result<(), Error> {
        if self.pending_authorization.is_none() && self.pending_publication.is_none() {
            return Ok(());
        }
        Err(pending_error(self.limits))
    }
}

#[cfg(test)]
impl Drop for TransactionBudget {
    fn drop(&mut self) {
        testing::record(self.usage);
    }
}

fn remaining_usage(used: Usage, limits: TransactionLimits) -> Result<Remaining, Error> {
    macro_rules! remaining {
        ($field:ident, $kind:expr, $maximum:expr) => {
            $maximum
                .checked_sub(used.$field)
                .ok_or_else(|| limit_error($kind, used.$field, $maximum))?
        };
    }
    macro_rules! activity {
        ($field:ident) => {
            remaining!(
                $field,
                LimitKind::TransactionWork,
                limits.max_transaction_work
            )
        };
    }
    Ok(Remaining {
        updates: remaining!(updates, LimitKind::Updates, limits.max_updates),
        input_value_bytes: remaining!(
            input_value_bytes,
            LimitKind::OwnedValueBytes,
            limits.max_owned_value_bytes
        ),
        retained_elements: remaining!(
            retained_elements,
            LimitKind::RetainedElements,
            limits.max_retained_elements
        ),
        retained_bytes: remaining!(
            retained_bytes,
            LimitKind::RetainedBytes,
            limits.max_retained_bytes
        ),
        scratch_bytes: remaining!(
            scratch_bytes,
            LimitKind::PeakScratchBytes,
            limits.max_scratch_bytes
        ),
        // Leaf peak reports are phase-local maxima. `add_usage` combines them
        // with the observed transaction peak using max, so the next leaf may
        // independently use the complete peak ceiling.
        peak_scratch_bytes: limits.max_scratch_bytes,
        allocation_events: remaining!(
            allocation_events,
            LimitKind::TransactionWork,
            limits.max_allocation_events
        ),
        wire_bytes: remaining!(wire_bytes, LimitKind::WireWork, limits.max_wire_bytes),
        wire_fields: remaining!(wire_fields, LimitKind::WireFields, limits.max_wire_fields),
        wire_work: remaining!(wire_work, LimitKind::WireWork, limits.max_wire_work),
        objects: remaining!(objects, LimitKind::Objects, limits.max_objects),
        references: remaining!(references, LimitKind::References, limits.max_references),
        lookups: activity!(lookups),
        tile_reads: activity!(tile_reads),
        tile_writes: activity!(tile_writes),
        header_reads: activity!(header_reads),
        header_writes: activity!(header_writes),
        row_reads: activity!(row_reads),
        row_writes: activity!(row_writes),
        list_reads: activity!(list_reads),
        list_writes: activity!(list_writes),
        string_work: activity!(string_work),
        rich_text_work: activity!(rich_text_work),
        formula_graph_builds: remaining!(
            formula_graph_builds,
            LimitKind::FormulaWork,
            limits.max_formula_work
        ),
        formula_nodes: remaining!(
            formula_nodes,
            LimitKind::FormulaWork,
            limits.max_formula_work
        ),
        formula_edges: remaining!(
            formula_edges,
            LimitKind::FormulaWork,
            limits.max_formula_work
        ),
        range_candidates: remaining!(
            range_candidates,
            LimitKind::FormulaWork,
            limits.max_formula_work
        ),
        cache_hosts: remaining!(cache_hosts, LimitKind::FormulaWork, limits.max_formula_work),
        authored_formula_writes: remaining!(
            authored_formula_writes,
            LimitKind::FormulaWork,
            limits.max_formula_work
        ),
        formula_work: remaining!(
            formula_work,
            LimitKind::FormulaWork,
            limits.max_formula_work
        ),
        component_encodes: remaining!(component_encodes, LimitKind::Objects, limits.max_objects),
        components_reassembled: activity!(components_reassembled),
        reassembly_bytes: activity!(reassembly_bytes),
        preview_bytes_deleted: activity!(preview_bytes_deleted),
        output_artifact_allocations: activity!(output_artifact_allocations),
        output_bytes: remaining!(
            output_bytes,
            LimitKind::OutputBytes,
            limits.max_output_bytes
        ),
        candidate_reopens: activity!(candidate_reopens),
        reopen_references: remaining!(
            reopen_references,
            LimitKind::References,
            limits.max_references
        ),
        reopen_work: remaining!(reopen_work, LimitKind::ReopenWork, limits.max_reopen_work),
        locality_bytes: activity!(locality_bytes),
        transaction_work: remaining!(
            transaction_work,
            LimitKind::TransactionWork,
            limits.max_transaction_work
        ),
    })
}

/// Partition a conservative nested reservation out of an already admitted
/// leaf envelope. Additive counters are removed from the capacity available
/// to the leaf's final report. Scratch peak is phase-local, so both phases may
/// independently use the admitted peak without adding their maxima together.
fn subtract_usage_envelope(envelope: Usage, nested: Usage) -> Result<Usage, Error> {
    validate_actual_le(nested, envelope)?;
    macro_rules! subtract {
        ($field:ident, $kind:expr) => {
            envelope
                .$field
                .checked_sub(nested.$field)
                .ok_or_else(|| limit_error($kind, nested.$field, envelope.$field))?
        };
    }
    macro_rules! activity {
        ($field:ident) => {
            subtract!($field, LimitKind::TransactionWork)
        };
    }
    Ok(Usage {
        updates: subtract!(updates, LimitKind::Updates),
        input_value_bytes: subtract!(input_value_bytes, LimitKind::OwnedValueBytes),
        retained_elements: subtract!(retained_elements, LimitKind::RetainedElements),
        retained_bytes: subtract!(retained_bytes, LimitKind::RetainedBytes),
        scratch_bytes: subtract!(scratch_bytes, LimitKind::PeakScratchBytes),
        peak_scratch_bytes: envelope.peak_scratch_bytes,
        allocation_events: activity!(allocation_events),
        wire_bytes: subtract!(wire_bytes, LimitKind::WireWork),
        wire_fields: subtract!(wire_fields, LimitKind::WireFields),
        wire_work: subtract!(wire_work, LimitKind::WireWork),
        objects: subtract!(objects, LimitKind::Objects),
        references: subtract!(references, LimitKind::References),
        lookups: activity!(lookups),
        tile_reads: activity!(tile_reads),
        tile_writes: activity!(tile_writes),
        header_reads: activity!(header_reads),
        header_writes: activity!(header_writes),
        row_reads: activity!(row_reads),
        row_writes: activity!(row_writes),
        list_reads: activity!(list_reads),
        list_writes: activity!(list_writes),
        string_work: activity!(string_work),
        rich_text_work: activity!(rich_text_work),
        formula_graph_builds: subtract!(formula_graph_builds, LimitKind::FormulaWork),
        formula_nodes: subtract!(formula_nodes, LimitKind::FormulaWork),
        formula_edges: subtract!(formula_edges, LimitKind::FormulaWork),
        range_candidates: subtract!(range_candidates, LimitKind::FormulaWork),
        cache_hosts: subtract!(cache_hosts, LimitKind::FormulaWork),
        authored_formula_writes: subtract!(authored_formula_writes, LimitKind::FormulaWork),
        formula_work: subtract!(formula_work, LimitKind::FormulaWork),
        component_encodes: subtract!(component_encodes, LimitKind::Objects),
        components_reassembled: activity!(components_reassembled),
        reassembly_bytes: activity!(reassembly_bytes),
        preview_bytes_deleted: activity!(preview_bytes_deleted),
        output_artifact_allocations: activity!(output_artifact_allocations),
        output_bytes: subtract!(output_bytes, LimitKind::OutputBytes),
        candidate_reopens: activity!(candidate_reopens),
        reopen_references: subtract!(reopen_references, LimitKind::References),
        reopen_work: subtract!(reopen_work, LimitKind::ReopenWork),
        locality_bytes: activity!(locality_bytes),
        transaction_work: activity!(transaction_work),
    })
}

pub(super) fn publication_usage_reservation(value: PublicationReservation) -> Result<Usage, Error> {
    publication_usage(
        value.components_reassembled,
        value.reassembly_bytes,
        value.preview_bytes_deleted,
        value.locality_bytes,
        value.locality_work,
        value.output_artifact_allocations,
        value.output_bytes,
        value.candidate_reopens,
        value.source_reopen,
        value.target_reopen,
        true,
    )
}

fn publication_usage_cost(value: PublicationCost) -> Result<Usage, Error> {
    publication_usage(
        value.components_reassembled,
        value.reassembly_bytes,
        value.preview_bytes_deleted,
        value.locality_bytes,
        value.locality_work,
        value.output_artifact_allocations,
        value.output_bytes,
        value.candidate_reopens,
        value.source_reopen,
        value.target_reopen,
        false,
    )
}

#[allow(clippy::too_many_arguments)]
fn publication_usage(
    components_reassembled: u64,
    reassembly_bytes: u64,
    preview_bytes_deleted: u64,
    locality_bytes: u64,
    locality_work: u64,
    output_allocations: u64,
    output_bytes: u64,
    candidate_reopens: u64,
    source_reopen: ReopenCost,
    target_reopen: ReopenCost,
    reserve_physical_resources: bool,
) -> Result<Usage, Error> {
    let reopen_work = checked_add_limit(
        source_reopen.work,
        target_reopen.work,
        LimitKind::ReopenWork,
    )?;
    let reopen_references = checked_add_limit(
        source_reopen.references,
        target_reopen.references,
        LimitKind::References,
    )?;
    // Component decode/serialization/compression belongs exclusively to the
    // outer component phase.  This nested reservation covers only ZIP
    // reassembly, the output owner, candidate reopen, and locality proof.
    let peak_scratch_bytes = if reserve_physical_resources {
        checked_mul_limit(
            reassembly_bytes.max(output_bytes),
            5,
            LimitKind::PeakScratchBytes,
        )?
    } else {
        0
    };
    let allocation_events = if reserve_physical_resources {
        [output_allocations, candidate_reopens, 4]
            .into_iter()
            .try_fold(0u64, |total, amount| {
                checked_add_limit(total, amount, LimitKind::TransactionWork)
            })?
    } else {
        0
    };
    let transaction_work = [
        components_reassembled,
        peak_scratch_bytes,
        allocation_events,
        reassembly_bytes,
        preview_bytes_deleted,
        // Locality has two independent observations: bytes compared are
        // reported in `Usage::locality_bytes`, while the work required to
        // prove those bytes unchanged is charged only to transaction work.
        locality_bytes,
        locality_work,
        output_allocations,
        output_bytes,
        candidate_reopens,
        reopen_work,
        reopen_references,
    ]
    .into_iter()
    .try_fold(0u64, |total, amount| {
        checked_add_limit(total, amount, LimitKind::TransactionWork)
    })?;
    Ok(Usage {
        peak_scratch_bytes,
        allocation_events,
        components_reassembled,
        reassembly_bytes,
        preview_bytes_deleted,
        output_artifact_allocations: output_allocations,
        output_bytes,
        candidate_reopens,
        references: reopen_references,
        reopen_references,
        reopen_work,
        locality_bytes,
        transaction_work,
        ..Usage::default()
    })
}

fn add_usage(base: Usage, delta: Usage, limits: TransactionLimits) -> Result<Usage, Error> {
    macro_rules! add {
        ($field:ident, $kind:expr, $maximum:expr) => {{
            let observed = base
                .$field
                .checked_add(delta.$field)
                .ok_or_else(|| limit_error($kind, u64::MAX, $maximum))?;
            if observed > $maximum {
                return Err(limit_error($kind, observed, $maximum));
            }
            observed
        }};
    }
    macro_rules! activity {
        ($field:ident) => {
            add!(
                $field,
                LimitKind::TransactionWork,
                limits.max_transaction_work
            )
        };
    }

    let scratch_bytes = add!(
        scratch_bytes,
        LimitKind::PeakScratchBytes,
        limits.max_scratch_bytes
    );
    let requested_peak = base
        .peak_scratch_bytes
        .max(delta.peak_scratch_bytes)
        .max(scratch_bytes);
    if requested_peak > limits.max_scratch_bytes {
        return Err(limit_error(
            LimitKind::PeakScratchBytes,
            requested_peak,
            limits.max_scratch_bytes,
        ));
    }

    Ok(Usage {
        updates: add!(updates, LimitKind::Updates, limits.max_updates),
        input_value_bytes: add!(
            input_value_bytes,
            LimitKind::OwnedValueBytes,
            limits.max_owned_value_bytes
        ),
        retained_elements: add!(
            retained_elements,
            LimitKind::RetainedElements,
            limits.max_retained_elements
        ),
        retained_bytes: add!(
            retained_bytes,
            LimitKind::RetainedBytes,
            limits.max_retained_bytes
        ),
        scratch_bytes,
        peak_scratch_bytes: requested_peak,
        allocation_events: add!(
            allocation_events,
            LimitKind::TransactionWork,
            limits.max_allocation_events
        ),
        wire_bytes: add!(wire_bytes, LimitKind::WireWork, limits.max_wire_bytes),
        wire_fields: add!(wire_fields, LimitKind::WireFields, limits.max_wire_fields),
        wire_work: add!(wire_work, LimitKind::WireWork, limits.max_wire_work),
        objects: add!(objects, LimitKind::Objects, limits.max_objects),
        references: add!(references, LimitKind::References, limits.max_references),
        lookups: activity!(lookups),
        tile_reads: activity!(tile_reads),
        tile_writes: activity!(tile_writes),
        header_reads: activity!(header_reads),
        header_writes: activity!(header_writes),
        row_reads: activity!(row_reads),
        row_writes: activity!(row_writes),
        list_reads: activity!(list_reads),
        list_writes: activity!(list_writes),
        string_work: activity!(string_work),
        rich_text_work: activity!(rich_text_work),
        formula_graph_builds: add!(
            formula_graph_builds,
            LimitKind::FormulaWork,
            limits.max_formula_work
        ),
        formula_nodes: add!(
            formula_nodes,
            LimitKind::FormulaWork,
            limits.max_formula_work
        ),
        formula_edges: add!(
            formula_edges,
            LimitKind::FormulaWork,
            limits.max_formula_work
        ),
        range_candidates: add!(
            range_candidates,
            LimitKind::FormulaWork,
            limits.max_formula_work
        ),
        cache_hosts: add!(cache_hosts, LimitKind::FormulaWork, limits.max_formula_work),
        authored_formula_writes: add!(
            authored_formula_writes,
            LimitKind::FormulaWork,
            limits.max_formula_work
        ),
        formula_work: add!(
            formula_work,
            LimitKind::FormulaWork,
            limits.max_formula_work
        ),
        component_encodes: add!(component_encodes, LimitKind::Objects, limits.max_objects),
        components_reassembled: activity!(components_reassembled),
        reassembly_bytes: activity!(reassembly_bytes),
        preview_bytes_deleted: activity!(preview_bytes_deleted),
        output_artifact_allocations: activity!(output_artifact_allocations),
        output_bytes: add!(
            output_bytes,
            LimitKind::OutputBytes,
            limits.max_output_bytes
        ),
        candidate_reopens: activity!(candidate_reopens),
        reopen_references: add!(
            reopen_references,
            LimitKind::References,
            limits.max_references
        ),
        reopen_work: add!(reopen_work, LimitKind::ReopenWork, limits.max_reopen_work),
        locality_bytes: activity!(locality_bytes),
        transaction_work: add!(
            transaction_work,
            LimitKind::TransactionWork,
            limits.max_transaction_work
        ),
    })
}

fn assert_usage_le(actual: Usage, reserved: Usage) {
    macro_rules! check {
        ($field:ident) => {
            assert!(
                actual.$field <= reserved.$field,
                "exact publication {} exceeds its preauthorization",
                stringify!($field)
            );
        };
    }
    check!(component_encodes);
    check!(components_reassembled);
    check!(reassembly_bytes);
    check!(preview_bytes_deleted);
    check!(output_artifact_allocations);
    check!(output_bytes);
    check!(candidate_reopens);
    check!(references);
    check!(reopen_references);
    check!(reopen_work);
    check!(locality_bytes);
    check!(transaction_work);
}

fn validate_actual_le(actual: Usage, reserved: Usage) -> Result<(), Error> {
    macro_rules! check {
        ($field:ident, $kind:expr) => {
            if actual.$field > reserved.$field {
                return Err(limit_error($kind, actual.$field, reserved.$field));
            }
        };
    }
    macro_rules! activity {
        ($field:ident) => {
            check!($field, LimitKind::TransactionWork)
        };
    }
    check!(updates, LimitKind::Updates);
    check!(input_value_bytes, LimitKind::OwnedValueBytes);
    check!(retained_elements, LimitKind::RetainedElements);
    check!(retained_bytes, LimitKind::RetainedBytes);
    check!(scratch_bytes, LimitKind::PeakScratchBytes);
    check!(peak_scratch_bytes, LimitKind::PeakScratchBytes);
    activity!(allocation_events);
    check!(wire_bytes, LimitKind::WireWork);
    check!(wire_fields, LimitKind::WireFields);
    check!(wire_work, LimitKind::WireWork);
    check!(objects, LimitKind::Objects);
    check!(references, LimitKind::References);
    activity!(lookups);
    activity!(tile_reads);
    activity!(tile_writes);
    activity!(header_reads);
    activity!(header_writes);
    activity!(row_reads);
    activity!(row_writes);
    activity!(list_reads);
    activity!(list_writes);
    activity!(string_work);
    activity!(rich_text_work);
    check!(formula_graph_builds, LimitKind::FormulaWork);
    check!(formula_nodes, LimitKind::FormulaWork);
    check!(formula_edges, LimitKind::FormulaWork);
    check!(range_candidates, LimitKind::FormulaWork);
    check!(cache_hosts, LimitKind::FormulaWork);
    check!(authored_formula_writes, LimitKind::FormulaWork);
    check!(formula_work, LimitKind::FormulaWork);
    check!(component_encodes, LimitKind::Objects);
    activity!(components_reassembled);
    activity!(reassembly_bytes);
    activity!(preview_bytes_deleted);
    activity!(output_artifact_allocations);
    check!(output_bytes, LimitKind::OutputBytes);
    activity!(candidate_reopens);
    check!(reopen_references, LimitKind::References);
    check!(reopen_work, LimitKind::ReopenWork);
    activity!(locality_bytes);
    activity!(transaction_work);
    Ok(())
}

fn assert_publication_le(actual: PublicationCost, reserved: PublicationReservation) {
    macro_rules! check {
        ($field:ident) => {
            assert!(
                actual.$field <= reserved.$field,
                "exact publication {} exceeds its directional preauthorization",
                stringify!($field)
            );
        };
    }
    check!(components_reassembled);
    check!(reassembly_bytes);
    check!(preview_bytes_deleted);
    check!(locality_bytes);
    check!(output_artifact_allocations);
    check!(output_bytes);
    check!(candidate_reopens);
    assert!(actual.source_reopen.work <= reserved.source_reopen.work);
    assert!(actual.source_reopen.references <= reserved.source_reopen.references);
    assert!(actual.target_reopen.work <= reserved.target_reopen.work);
    assert!(actual.target_reopen.references <= reserved.target_reopen.references);
}

fn as_u64(value: usize, kind: LimitKind) -> Result<u64, Error> {
    u64::try_from(value).map_err(|_error| limit_error(kind, u64::MAX, u64::MAX - 1))
}

fn checked_add_limit(left: u64, right: u64, kind: LimitKind) -> Result<u64, Error> {
    left.checked_add(right)
        .ok_or_else(|| limit_error(kind, u64::MAX, u64::MAX - 1))
}

fn checked_mul_limit(left: u64, right: u64, kind: LimitKind) -> Result<u64, Error> {
    left.checked_mul(right)
        .ok_or_else(|| limit_error(kind, u64::MAX, u64::MAX - 1))
}

const fn limit_error(kind: LimitKind, observed: u64, maximum: u64) -> Error {
    Error::LimitExceeded {
        kind,
        observed,
        maximum,
        path: Path::Package,
    }
}

fn pending_error(limits: TransactionLimits) -> Error {
    limit_error(
        LimitKind::TransactionWork,
        // All constructed limits are finite, checked values. Should a caller
        // manually provide the representational maximum, retain the typed
        // rejection without manufacturing a saturated observation.
        match limits.max_transaction_work.checked_add(1) {
            Some(observed) => observed,
            None => limits.max_transaction_work,
        },
        limits.max_transaction_work,
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    fn limits(max_transaction_work: u64) -> TransactionLimits {
        TransactionLimits {
            max_updates: 100_000,
            max_owned_value_bytes: 100_000,
            max_retained_elements: 1_000_000,
            max_retained_bytes: 10_000_000,
            max_scratch_bytes: 10_000_000,
            max_allocation_events: 1_000_000,
            max_wire_bytes: 10_000_000,
            max_wire_fields: 10_000_000,
            max_wire_work: 10_000_000,
            max_objects: 1_000_000,
            max_references: 10_000_000,
            max_formula_work: 10_000_000,
            max_output_bytes: 10_000_000,
            max_reopen_work: 10_000_000,
            max_transaction_work,
        }
    }

    fn publication(output: u64) -> PublicationReservation {
        PublicationReservation {
            components_reassembled: 1,
            reassembly_bytes: output,
            preview_bytes_deleted: 3,
            locality_bytes: output * 2,
            locality_work: output * 3,
            output_artifact_allocations: 2,
            output_bytes: output,
            candidate_reopens: 1,
            source_reopen: ReopenCost {
                work: 700,
                references: 11,
            },
            target_reopen: ReopenCost {
                work: 800,
                references: 13,
            },
        }
    }

    #[test]
    fn reservation_is_atomic_and_scratch_peak_is_exact() {
        let mut budget = TransactionBudget::from_limits(limits(1_000_000)).unwrap();
        budget.reserve_retained(4, 80, 1).unwrap();
        budget.reserve_scratch(64, 1).unwrap();
        budget.reserve_scratch(32, 1).unwrap();
        budget.release_scratch(32);
        budget.release_scratch(64);
        let before = budget.usage;
        assert_eq!(before.scratch_bytes, 0);
        assert_eq!(before.peak_scratch_bytes, 96);
        assert_eq!(before.allocation_events, 3);

        let delta = Usage {
            retained_elements: 1,
            retained_bytes: 20_000_000,
            ..Usage::default()
        };
        assert!(budget.reserve(delta).is_err());
        assert_eq!(budget.usage, before);
    }

    #[test]
    fn retained_plan_release_restores_capacity_but_not_allocations() {
        let mut budget = TransactionBudget::from_limits(limits(1_000_000)).unwrap();
        budget.reserve_retained(4, 80, 1).unwrap();
        let after_reserve = budget.remaining().unwrap();
        budget.release_retained(4, 80).unwrap();
        let after_release = budget.remaining().unwrap();
        assert_eq!(
            after_release.retained_elements,
            after_reserve.retained_elements + 4
        );
        assert_eq!(
            after_release.retained_bytes,
            after_reserve.retained_bytes + 80
        );
        assert_eq!(
            after_release.allocation_events,
            after_reserve.allocation_events
        );
        assert!(budget.release_retained(1, 0).is_err());

        budget.reserve_retained(2, 10, 1).unwrap();
        let before_failed_release = budget.usage;
        assert!(budget.release_retained(1, 11).is_err());
        assert_eq!(budget.usage, before_failed_release);
    }

    #[test]
    fn leaf_authorization_is_non_observing_and_records_only_actual_usage() {
        let mut budget = TransactionBudget::from_limits(limits(1_000_000)).unwrap();
        let envelope = Usage {
            retained_elements: 128,
            retained_bytes: 16_384,
            peak_scratch_bytes: 8_192,
            allocation_events: 32,
            transaction_work: 64_000,
            ..Usage::default()
        };
        budget.authorize(envelope).unwrap();
        assert_eq!(budget.usage, Usage::default());
        assert!(budget.authorization_is_pending());

        let actual = Usage {
            retained_elements: 64,
            retained_bytes: 8_192,
            peak_scratch_bytes: 4_096,
            allocation_events: 7,
            transaction_work: 31_000,
            ..Usage::default()
        };
        budget.record_authorized(actual).unwrap();
        assert_eq!(budget.usage, actual);
        assert!(!budget.authorization_is_pending());
    }

    #[test]
    fn remaining_is_exact_and_oversized_actual_returns_without_mutation() {
        let mut budget = TransactionBudget::from_limits(limits(1_000_000)).unwrap();
        budget
            .reserve(Usage {
                retained_elements: 7,
                wire_fields: 11,
                allocation_events: 3,
                peak_scratch_bytes: 512,
                transaction_work: 19,
                ..Usage::default()
            })
            .unwrap();
        let before = budget.usage;
        let remaining = budget.remaining().unwrap();
        assert_eq!(remaining.retained_elements, 1_000_000 - 7);
        assert_eq!(remaining.wire_fields, 10_000_000 - 11);
        assert_eq!(remaining.allocation_events, 1_000_000 - 3);
        assert_eq!(remaining.peak_scratch_bytes, 10_000_000);
        assert_eq!(remaining.transaction_work, 1_000_000 - 19);

        budget
            .authorize(Usage {
                retained_elements: 1,
                transaction_work: 1,
                ..Usage::default()
            })
            .unwrap();
        let error = budget
            .record_authorized(Usage {
                retained_elements: 2,
                transaction_work: 1,
                ..Usage::default()
            })
            .unwrap_err();
        assert!(matches!(
            error,
            Error::LimitExceeded {
                kind: LimitKind::RetainedElements,
                observed: 2,
                maximum: 1,
                ..
            }
        ));
        assert_eq!(budget.usage, before);
        assert!(budget.authorization_is_pending());
        budget.cancel_authorization();
    }

    #[test]
    fn max_minus_one_leaf_authorization_prevents_callback_entry() {
        let required = 587_222;
        let mut budget = TransactionBudget::from_limits(limits(required - 1)).unwrap();
        let envelope = Usage {
            transaction_work: required,
            ..Usage::default()
        };
        let mut callback_executed = false;
        let result = (|| -> Result<(), Error> {
            budget.authorize(envelope)?;
            callback_executed = true;
            budget.record_authorized(Usage::default())?;
            Ok(())
        })();
        assert!(matches!(
            result,
            Err(Error::LimitExceeded {
                kind: LimitKind::TransactionWork,
                observed,
                maximum,
                ..
            }) if observed == required && maximum == required - 1
        ));
        assert!(!callback_executed);
        assert_eq!(budget.usage, Usage::default());
        assert!(!budget.authorization_is_pending());
    }

    #[test]
    fn max_minus_one_apply_authorization_prevents_reopen_callback_entry() {
        let required = 587_222;
        let mut budget = TransactionBudget::from_limits(limits(required - 1)).unwrap();
        let envelope = Usage {
            peak_scratch_bytes: 32_768,
            allocation_events: 128,
            objects: 64,
            references: 32,
            candidate_reopens: 1,
            reopen_references: 32,
            reopen_work: 100_000,
            locality_bytes: 200_000,
            transaction_work: required,
            ..Usage::default()
        };
        let mut callback_executed = false;
        let result = budget.with_apply_authorization(envelope, || {
            callback_executed = true;
            Ok(((), Usage::default()))
        });
        assert!(matches!(
            result,
            Err(Error::LimitExceeded {
                kind: LimitKind::TransactionWork,
                observed,
                maximum,
                ..
            }) if observed == required && maximum == required - 1
        ));
        assert!(!callback_executed);
        assert_eq!(budget.usage, Usage::default());
        assert!(!budget.authorization_is_pending());
    }

    #[test]
    fn apply_authorization_settles_exact_reopen_and_locality_usage() {
        let mut budget = TransactionBudget::from_limits(limits(1_000_000)).unwrap();
        let envelope = Usage {
            peak_scratch_bytes: 65_536,
            allocation_events: 256,
            objects: 128,
            references: 64,
            candidate_reopens: 1,
            reopen_references: 64,
            reopen_work: 200_000,
            locality_bytes: 400_000,
            transaction_work: 900_000,
            ..Usage::default()
        };
        let actual = Usage {
            peak_scratch_bytes: 24_576,
            allocation_events: 19,
            objects: 37,
            references: 23,
            candidate_reopens: 1,
            reopen_references: 23,
            reopen_work: 90_000,
            locality_bytes: 140_000,
            transaction_work: 310_000,
            ..Usage::default()
        };
        let value = budget
            .with_apply_authorization(envelope, || Ok((41_u8, actual)))
            .unwrap();
        assert_eq!(value, 41);
        assert_eq!(budget.usage, actual);
        assert!(!budget.authorization_is_pending());
    }

    #[test]
    fn arc_vec_retention_counts_capacity_header_control_and_allocations() {
        let usage = arc_vec_retained_usage::<u64>(3, 4).unwrap();
        assert_eq!(usage.retained_elements, 3);
        assert_eq!(
            usage.retained_bytes,
            u64::try_from(4 * size_of::<u64>() + size_of::<Vec<u64>>() + 2 * size_of::<usize>())
                .unwrap()
        );
        assert_eq!(usage.allocation_events, 2);
        assert_eq!(usage.transaction_work, 3);

        let empty = arc_vec_retained_usage::<u64>(0, 0).unwrap();
        assert_eq!(
            empty.retained_bytes,
            u64::try_from(size_of::<Vec<u64>>() + 2 * size_of::<usize>()).unwrap()
        );
        assert_eq!(empty.allocation_events, 1);
        assert!(arc_vec_retained_usage::<u64>(2, 1).is_err());
    }

    #[test]
    fn compact_4096_to_8192_counters_are_linear_or_fixed() {
        fn measured(updates: u64) -> Usage {
            Usage {
                updates,
                retained_elements: updates * 4,
                retained_bytes: updates * 96,
                peak_scratch_bytes: updates * 64,
                allocation_events: updates + 8,
                tile_reads: updates,
                tile_writes: updates,
                row_reads: updates,
                row_writes: updates,
                string_work: updates,
                formula_graph_builds: 1,
                formula_nodes: updates,
                formula_edges: updates * 2,
                range_candidates: updates,
                cache_hosts: updates,
                authored_formula_writes: updates,
                formula_work: updates * 8,
                component_encodes: 1,
                components_reassembled: 1,
                output_artifact_allocations: 2,
                candidate_reopens: 1,
                ..Usage::default()
            }
        }
        let small = measured(4_096);
        let large = measured(8_192);
        for (left, right) in [
            (small.updates, large.updates),
            (small.retained_elements, large.retained_elements),
            (small.retained_bytes, large.retained_bytes),
            (small.peak_scratch_bytes, large.peak_scratch_bytes),
            (small.formula_nodes, large.formula_nodes),
            (small.formula_edges, large.formula_edges),
            (small.range_candidates, large.range_candidates),
            (small.cache_hosts, large.cache_hosts),
            (small.authored_formula_writes, large.authored_formula_writes),
            (small.formula_work, large.formula_work),
        ] {
            assert!(right * 10 <= left * 22);
        }
        assert_eq!(
            (small.formula_graph_builds, large.formula_graph_builds),
            (1, 1)
        );
        assert_eq!((small.component_encodes, large.component_encodes), (1, 1));
        assert_eq!(
            (small.components_reassembled, large.components_reassembled),
            (1, 1)
        );
        assert_eq!(
            (
                small.output_artifact_allocations,
                large.output_artifact_allocations
            ),
            (2, 2)
        );
        assert_eq!((small.candidate_reopens, large.candidate_reopens), (1, 1));
    }

    #[test]
    fn max_minus_one_refuses_before_publication_events() {
        let reservation = publication(8_192);
        let mapped = publication_usage_reservation(reservation).unwrap();
        assert_eq!(mapped.peak_scratch_bytes, 8_192 * 5);
        assert_eq!(mapped.allocation_events, 7);
        assert_eq!(mapped.component_encodes, 0);
        let required = mapped.transaction_work;
        let mut budget =
            TransactionBudget::from_limits(limits(required.checked_sub(1).unwrap())).unwrap();
        let error = budget.preauthorize_publication(reservation).unwrap_err();
        assert!(matches!(
            error,
            Error::LimitExceeded {
                kind: LimitKind::TransactionWork,
                observed,
                maximum,
                ..
            } if observed == required && maximum == required - 1
        ));
        let usage = budget.usage;
        assert_eq!(usage.component_encodes, 0);
        assert_eq!(usage.components_reassembled, 0);
        assert_eq!(usage.reassembly_bytes, 0);
        assert_eq!(usage.output_artifact_allocations, 0);
        assert_eq!(usage.output_bytes, 0);
        assert_eq!(usage.candidate_reopens, 0);
        assert_eq!(usage.locality_bytes, 0);
        assert!(!budget.publication_is_authorized());
    }

    #[test]
    fn publication_observations_are_exact_not_reserved_bounds() {
        let mut budget = TransactionBudget::from_limits(limits(1_000_000)).unwrap();
        budget.preauthorize_publication(publication(8_192)).unwrap();
        assert_eq!(budget.usage.output_artifact_allocations, 0);
        assert_eq!(budget.usage.candidate_reopens, 0);
        let actual = PublicationCost {
            components_reassembled: 1,
            reassembly_bytes: 4_096,
            preview_bytes_deleted: 3,
            locality_bytes: 8_192,
            locality_work: 12_288,
            output_artifact_allocations: 2,
            output_bytes: 4_096,
            candidate_reopens: 1,
            source_reopen: ReopenCost {
                work: 700,
                references: 11,
            },
            target_reopen: ReopenCost {
                work: 700,
                references: 12,
            },
        };
        budget.finish_publication(actual);
        let usage = budget.usage;
        assert_eq!(usage.component_encodes, 0);
        assert_eq!(usage.components_reassembled, 1);
        assert_eq!(usage.output_bytes, 4_096);
        assert_eq!(usage.reassembly_bytes, 4_096);
        assert_eq!(usage.reopen_work, 1_400);
        assert_eq!(usage.reopen_references, 23);
        assert_eq!(usage.output_artifact_allocations, 2);
        assert_eq!(usage.candidate_reopens, 1);
    }

    #[test]
    fn publication_can_nest_inside_a_grouped_writer_authorization() {
        let mut budget = TransactionBudget::from_limits(limits(1_000_000)).unwrap();
        let writer_envelope = budget.remaining().unwrap();
        budget.authorize(writer_envelope).unwrap();
        let publication_reservation = publication(8_192);
        let reserved = publication_usage_reservation(publication_reservation).unwrap();
        budget
            .preauthorize_publication(publication_reservation)
            .unwrap();
        assert_eq!(budget.usage, Usage::default());
        let component_envelope = budget.authorization_remaining().unwrap();
        assert_eq!(
            component_envelope.transaction_work,
            writer_envelope.transaction_work - reserved.transaction_work
        );
        assert_eq!(
            component_envelope.allocation_events,
            writer_envelope.allocation_events - reserved.allocation_events
        );
        assert_eq!(
            component_envelope.peak_scratch_bytes,
            writer_envelope.peak_scratch_bytes
        );
        assert_eq!(
            component_envelope.component_encodes,
            writer_envelope.component_encodes
        );

        let actual_publication = PublicationCost {
            components_reassembled: 1,
            reassembly_bytes: 4_096,
            preview_bytes_deleted: 3,
            locality_bytes: 8_192,
            locality_work: 12_288,
            output_artifact_allocations: 2,
            output_bytes: 4_096,
            candidate_reopens: 1,
            source_reopen: ReopenCost {
                work: 700,
                references: 11,
            },
            target_reopen: ReopenCost {
                work: 700,
                references: 12,
            },
        };
        budget.finish_publication(actual_publication);
        budget
            .record_authorized(Usage {
                component_encodes: 1,
                transaction_work: component_envelope.transaction_work,
                ..Usage::default()
            })
            .unwrap();
        assert_eq!(budget.usage.peak_scratch_bytes, 0);
        assert_eq!(budget.usage.allocation_events, 0);
        assert_eq!(budget.usage.component_encodes, 1);
        assert_eq!(budget.usage.components_reassembled, 1);
        assert!(budget.usage.transaction_work <= 1_000_000);
        assert!(!budget.authorization_is_pending());
        assert!(!budget.publication_is_authorized());
    }

    #[test]
    fn nested_publication_returns_locality_capacity_to_outer_proof() {
        let mut budget = TransactionBudget::from_limits(limits(1_000_000)).unwrap();
        let writer_envelope = budget.remaining().unwrap();
        budget.authorize(writer_envelope).unwrap();
        let reservation = publication(8_192);
        budget.preauthorize_publication(reservation).unwrap();

        // Publication deliberately records no locality: the exact locality
        // proof is produced only after the candidate has been compared by the
        // outer grouped writer.
        let publication = PublicationCost {
            components_reassembled: 1,
            reassembly_bytes: 4_096,
            preview_bytes_deleted: 3,
            locality_bytes: 0,
            locality_work: 0,
            output_artifact_allocations: 2,
            output_bytes: 4_096,
            candidate_reopens: 1,
            source_reopen: ReopenCost {
                work: 700,
                references: 11,
            },
            target_reopen: ReopenCost {
                work: 700,
                references: 12,
            },
        };
        let exact_publication = publication_usage_cost(publication).unwrap();
        budget.finish_publication(publication);

        let settlement = budget.authorization_remaining().unwrap();
        assert_eq!(settlement.locality_bytes, writer_envelope.locality_bytes);
        assert_eq!(
            settlement.transaction_work,
            writer_envelope.transaction_work - exact_publication.transaction_work
        );

        let locality_report = Usage {
            locality_bytes: reservation.locality_bytes,
            transaction_work: reservation.locality_work,
            ..Usage::default()
        };
        budget.record_authorized(locality_report).unwrap();
        let usage = budget.usage;
        assert_eq!(usage.locality_bytes, reservation.locality_bytes);
        assert_eq!(
            usage.transaction_work,
            exact_publication.transaction_work + reservation.locality_work
        );
        assert!(!budget.authorization_is_pending());
    }

    #[test]
    fn tile_work_ceiling_accounts_for_double_charged_wire_work_and_writes() {
        let remaining = Usage {
            wire_work: 1_000,
            peak_scratch_bytes: 900,
            transaction_work: 1_001,
            ..Usage::default()
        };
        assert_eq!(tile_work_ceiling(remaining, 2_000, 1).unwrap(), 500);

        let error = tile_work_ceiling(remaining, 2_000, 1_002).unwrap_err();
        assert!(matches!(
            error,
            Error::LimitExceeded {
                kind: LimitKind::TransactionWork,
                observed: 1_002,
                maximum: 1_001,
                ..
            }
        ));
    }
}
