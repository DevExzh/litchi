//! One operation-wide resource ledger for Keynote media lifecycle edits.
//!
//! The ledger is deliberately private to the package owner.  It is consumed
//! by graph selection, raw payload cloning, metadata transitions, archive
//! encoding, ZIP reassembly, and candidate readback; helpers must receive the
//! same mutable instance instead of manufacturing a fresh budget for a phase.

use super::{Package, SlideMediaLifecycleError, SlideMediaLifecycleLimitKind};

/// Finite resources charged by one duplicate/remove transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(in crate::package) struct LifecycleBudget {
    max_input: u64,
    max_output: u64,
    max_entries: usize,
    max_references: usize,
    max_media_bytes: u64,
    max_wire_fields: usize,
    max_wire_work: u64,
    max_nesting: u32,
    max_allocations: usize,
    input: u64,
    output: u64,
    entries: usize,
    references: usize,
    media_bytes: u64,
    wire_fields: usize,
    wire_work: u64,
    observed_nesting: u32,
    allocation_bytes: u64,
    allocations: usize,
}

impl LifecycleBudget {
    /// Build a ledger from the package's already checked physical/semantic
    /// profiles.  The maxima are finite even when an individual source is
    /// small; every later phase still consumes the same counters.
    pub(in crate::package) fn for_package(
        package: &Package,
    ) -> Result<Self, SlideMediaLifecycleError> {
        let physical = package.limits();
        let semantic = package.semantic_limits();
        let max_input = physical.max_input_bytes();
        let max_output = physical
            .max_total_bytes()
            .checked_mul(16)
            .ok_or(SlideMediaLifecycleError::InvalidSource)?;
        let max_entries = physical.max_entries();
        let max_references = semantic.max_references();
        let max_media_bytes = physical.max_total_bytes();
        let max_wire_fields = semantic
            .max_references()
            .checked_mul(8)
            .ok_or(SlideMediaLifecycleError::InvalidSource)?
            .max(1);
        let max_wire_work = u64::try_from(physical.max_iwa_stream_bytes())
            .map_err(|_| SlideMediaLifecycleError::InvalidSource)?
            .checked_mul(128)
            .ok_or(SlideMediaLifecycleError::InvalidSource)?
            .max(1);
        let max_nesting = u32::try_from(
            package
                .semantic_wire_limits()
                .map_err(|_| SlideMediaLifecycleError::InvalidSource)?
                .max_nesting(),
        )
        .map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
        let max_allocations = max_entries.saturating_mul(16).max(1);
        let mut budget = Self {
            max_input,
            max_output,
            max_entries,
            max_references,
            max_media_bytes,
            max_wire_fields,
            max_wire_work,
            max_nesting,
            max_allocations,
            input: 0,
            output: 0,
            entries: 0,
            references: 0,
            media_bytes: 0,
            wire_fields: 0,
            wire_work: 0,
            observed_nesting: 0,
            allocation_bytes: 0,
            allocations: 0,
        };
        budget.charge_input(package.source_bytes().len())?;
        Ok(budget)
    }

    pub(in crate::package) fn charge_input(
        &mut self,
        amount: usize,
    ) -> Result<(), SlideMediaLifecycleError> {
        let amount = as_u64(amount)?;
        let observed = checked_add(self.input, amount)?;
        if observed > self.max_input {
            return Err(limit(
                SlideMediaLifecycleLimitKind::InputBytes,
                observed,
                self.max_input,
            ));
        }
        self.input = observed;
        Ok(())
    }

    pub(in crate::package) fn charge_output(
        &mut self,
        amount: usize,
    ) -> Result<(), SlideMediaLifecycleError> {
        let amount = as_u64(amount)?;
        let observed = checked_add(self.output, amount)?;
        if observed > self.max_output {
            return Err(limit(
                SlideMediaLifecycleLimitKind::OutputBytes,
                observed,
                self.max_output,
            ));
        }
        self.output = observed;
        Ok(())
    }

    pub(in crate::package) fn charge_entries(
        &mut self,
        amount: usize,
    ) -> Result<(), SlideMediaLifecycleError> {
        let observed = self.entries.checked_add(amount).ok_or_else(|| {
            limit_usize(
                SlideMediaLifecycleLimitKind::Entries,
                usize::MAX,
                self.max_entries,
            )
        })?;
        if observed > self.max_entries {
            return Err(limit_usize(
                SlideMediaLifecycleLimitKind::Entries,
                observed,
                self.max_entries,
            ));
        }
        self.entries = observed;
        Ok(())
    }

    pub(in crate::package) fn charge_references(
        &mut self,
        amount: usize,
    ) -> Result<(), SlideMediaLifecycleError> {
        let observed = self.references.checked_add(amount).ok_or_else(|| {
            limit_usize(
                SlideMediaLifecycleLimitKind::References,
                usize::MAX,
                self.max_references,
            )
        })?;
        if observed > self.max_references {
            return Err(limit_usize(
                SlideMediaLifecycleLimitKind::References,
                observed,
                self.max_references,
            ));
        }
        self.references = observed;
        Ok(())
    }

    pub(in crate::package) fn charge_media_bytes(
        &mut self,
        amount: usize,
    ) -> Result<(), SlideMediaLifecycleError> {
        let amount = as_u64(amount)?;
        let observed = checked_add(self.media_bytes, amount)?;
        if observed > self.max_media_bytes {
            return Err(limit(
                SlideMediaLifecycleLimitKind::MediaBytes,
                observed,
                self.max_media_bytes,
            ));
        }
        self.media_bytes = observed;
        Ok(())
    }

    pub(in crate::package) fn charge_wire_fields(
        &mut self,
        amount: usize,
    ) -> Result<(), SlideMediaLifecycleError> {
        let observed = self.wire_fields.checked_add(amount).ok_or_else(|| {
            limit_usize(
                SlideMediaLifecycleLimitKind::WireFields,
                usize::MAX,
                self.max_wire_fields,
            )
        })?;
        if observed > self.max_wire_fields {
            return Err(limit_usize(
                SlideMediaLifecycleLimitKind::WireFields,
                observed,
                self.max_wire_fields,
            ));
        }
        self.wire_fields = observed;
        Ok(())
    }

    pub(in crate::package) fn charge_wire_work(
        &mut self,
        amount: usize,
    ) -> Result<(), SlideMediaLifecycleError> {
        let amount = as_u64(amount)?;
        let observed = checked_add(self.wire_work, amount)?;
        if observed > self.max_wire_work {
            return Err(limit(
                SlideMediaLifecycleLimitKind::WireWork,
                observed,
                self.max_wire_work,
            ));
        }
        self.wire_work = observed;
        Ok(())
    }

    /// Charge both allocation bytes and one allocation event.  Callers must
    /// invoke this before `try_reserve` or any other fallible allocation.
    pub(in crate::package) fn charge_allocations(
        &mut self,
        amount: usize,
    ) -> Result<(), SlideMediaLifecycleError> {
        self.charge_allocation_plan(amount, 1)
    }

    /// Atomically admit a bounded library call's bytes and allocation events.
    pub(in crate::package) fn charge_allocation_plan(
        &mut self,
        amount: usize,
        events: usize,
    ) -> Result<(), SlideMediaLifecycleError> {
        let bytes = as_u64(amount)?;
        let observed_bytes = checked_add(self.allocation_bytes, bytes)?;
        if observed_bytes > self.max_output {
            return Err(limit(
                SlideMediaLifecycleLimitKind::Allocations,
                observed_bytes,
                self.max_output,
            ));
        }
        let observed_events = self.allocations.checked_add(events).ok_or_else(|| {
            limit_usize(
                SlideMediaLifecycleLimitKind::Allocations,
                usize::MAX,
                self.max_allocations,
            )
        })?;
        if observed_events > self.max_allocations {
            return Err(limit_usize(
                SlideMediaLifecycleLimitKind::Allocations,
                observed_events,
                self.max_allocations,
            ));
        }
        self.allocation_bytes = observed_bytes;
        self.allocations = observed_events;
        Ok(())
    }

    pub(in crate::package) fn charge_nesting(
        &mut self,
        amount: usize,
    ) -> Result<(), SlideMediaLifecycleError> {
        let amount = u32::try_from(amount).map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
        if amount > self.max_nesting {
            return Err(limit(
                SlideMediaLifecycleLimitKind::WireNesting,
                u64::from(amount),
                u64::from(self.max_nesting),
            ));
        }
        self.observed_nesting = self.observed_nesting.max(amount);
        Ok(())
    }
}

fn as_u64(amount: usize) -> Result<u64, SlideMediaLifecycleError> {
    u64::try_from(amount).map_err(|_| SlideMediaLifecycleError::InvalidSource)
}

fn checked_add(left: u64, right: u64) -> Result<u64, SlideMediaLifecycleError> {
    left.checked_add(right)
        .ok_or(SlideMediaLifecycleError::InvalidSource)
}

fn limit(
    kind: SlideMediaLifecycleLimitKind,
    observed: u64,
    maximum: u64,
) -> SlideMediaLifecycleError {
    SlideMediaLifecycleError::LimitExceeded {
        kind,
        observed,
        maximum,
    }
}

fn limit_usize(
    kind: SlideMediaLifecycleLimitKind,
    observed: usize,
    maximum: usize,
) -> SlideMediaLifecycleError {
    let observed = match u64::try_from(observed) {
        Ok(value) => value,
        Err(_) => return SlideMediaLifecycleError::InvalidSource,
    };
    let maximum = match u64::try_from(maximum) {
        Ok(value) => value,
        Err(_) => return SlideMediaLifecycleError::InvalidSource,
    };
    limit(kind, observed, maximum)
}
