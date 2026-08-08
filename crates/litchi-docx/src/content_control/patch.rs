//! Retry-safe exact-source and package patch gates.

use std::sync::{
    Arc,
    atomic::{AtomicU8, Ordering},
};

use crate::{Error, Result};

use super::snapshot::Source;
use super::{Limits, Snapshot};

const READY: u8 = 0;
const IN_FLIGHT: u8 = 1;
const APPLIED: u8 = 2;

/// Reversible, exact-source content-control XML replacement.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Source,
    after: Source,
    limits: Limits,
    gate: Gate,
}

impl Patch {
    pub(crate) fn new(before: Source, after: Source, limits: Limits) -> Self {
        Self {
            before,
            after,
            limits,
            gate: Gate::new(),
        }
    }

    /// Exact source required by the patch.
    #[must_use]
    pub fn before_bytes(&self) -> &[u8] {
        self.before.as_slice()
    }

    /// Fully reparsed candidate bytes.
    #[must_use]
    pub fn after_bytes(&self) -> &[u8] {
        self.after.as_slice()
    }

    /// Whether every source byte is retained.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.as_slice() == self.after.as_slice()
    }

    /// Whether this patch or one of its clones was successfully applied.
    #[must_use]
    pub fn is_applied(&self) -> bool {
        self.gate.is_applied()
    }

    /// Apply once to the exact detached snapshot. Stale failures remain retryable.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        let claim = self.gate.claim("content-control patch")?;
        if source.source() != self.before.as_slice() || source.limits() != &self.limits {
            return Err(Error::Invalid(
                "content-control patch source does not match its exact precondition".into(),
            ));
        }
        let candidate = Snapshot::from_source(self.after.clone(), self.limits.clone())?;
        claim.finalize();
        Ok(candidate)
    }

    /// Construct a fresh inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self::new(self.after.clone(), self.before.clone(), self.limits.clone())
    }
}

#[derive(Debug, Clone)]
pub(crate) struct Gate(Arc<AtomicU8>);

impl Gate {
    pub(crate) fn new() -> Self {
        Self(Arc::new(AtomicU8::new(READY)))
    }

    pub(crate) fn is_applied(&self) -> bool {
        self.0.load(Ordering::Acquire) == APPLIED
    }

    pub(crate) fn claim(&self, subject: &str) -> Result<Claim> {
        match self
            .0
            .compare_exchange(READY, IN_FLIGHT, Ordering::AcqRel, Ordering::Acquire)
        {
            Ok(_) => Ok(Claim {
                gate: self.clone(),
                finalized: false,
            }),
            Err(IN_FLIGHT) => Err(Error::Invalid(format!(
                "{subject} publication is already in flight"
            ))),
            Err(_) => Err(Error::Invalid(format!("{subject} was already applied"))),
        }
    }
}

pub(crate) struct Claim {
    gate: Gate,
    finalized: bool,
}

impl Claim {
    pub(crate) fn finalize(mut self) {
        self.gate.0.store(APPLIED, Ordering::Release);
        self.finalized = true;
    }
}

impl Drop for Claim {
    fn drop(&mut self) {
        if !self.finalized {
            let _ =
                self.gate
                    .0
                    .compare_exchange(IN_FLIGHT, READY, Ordering::AcqRel, Ordering::Acquire);
        }
    }
}
