//! The finite work budget of the ported CalculationEngine phase.
//!
//! Every codec call already runs under its own limits; this budget bounds
//! the aggregate across the phase, derived from the package's size so a
//! hostile dependency graph cannot make a removal unbounded.

use litchi_iwa_common::WireLimits;

use super::{Fault, FaultLimit};

/// The aggregate codec and payload work of one phase.
#[derive(Debug)]
pub struct WireBudget {
    limits: WireLimits,
    work: usize,
    maximum_work: usize,
    fields: usize,
    references: usize,
}

fn exceeded(observed: usize, maximum: usize) -> Fault {
    Fault::Limit {
        kind: FaultLimit::PayloadBytes,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
    }
}

impl WireBudget {
    /// A budget proportional to `source_bytes` (decoded package size).
    pub fn new(source_bytes: usize) -> Self {
        let maximum_work = source_bytes
            .saturating_mul(64)
            .clamp(1 << 20, WireLimits::MAX_REWRITE_WORK.saturating_mul(16));
        Self {
            limits: WireLimits::default(),
            work: 0,
            maximum_work,
            fields: 0,
            references: 0,
        }
    }

    pub const fn wire_limits(&self) -> WireLimits {
        self.limits
    }

    pub fn remaining_wire_fields(&self) -> usize {
        WireLimits::MAX_FIELDS.saturating_sub(self.fields)
    }

    pub fn remaining_wire_work(&self) -> usize {
        self.maximum_work.saturating_sub(self.work)
    }

    pub fn remaining_payload_references(&self) -> usize {
        WireLimits::MAX_FIELDS.saturating_sub(self.references)
    }

    pub fn charge_payload_work(&mut self, amount: usize) -> Result<(), Fault> {
        self.work = self.work.saturating_add(amount);
        if self.work > self.maximum_work {
            return Err(exceeded(self.work, self.maximum_work));
        }
        Ok(())
    }

    pub fn charge_payload_bytes(&mut self, amount: usize) -> Result<(), Fault> {
        self.charge_payload_work(amount)
    }

    pub fn charge_output_bytes(&mut self, amount: usize) -> Result<(), Fault> {
        self.charge_payload_work(amount)
    }

    pub fn charge_sort_work(&mut self, items: usize) -> Result<(), Fault> {
        let steps = usize::BITS - items.leading_zeros();
        self.charge_payload_work(items.saturating_mul(steps as usize))
    }

    pub fn charge_codec_report(
        &mut self,
        fields: usize,
        work: usize,
        depth: u32,
        references: usize,
    ) -> Result<(), Fault> {
        let _ = depth;
        self.fields = self.fields.saturating_add(fields);
        self.references = self.references.saturating_add(references);
        if self.fields > WireLimits::MAX_FIELDS {
            return Err(exceeded(self.fields, WireLimits::MAX_FIELDS));
        }
        self.charge_payload_work(work)
    }
}
