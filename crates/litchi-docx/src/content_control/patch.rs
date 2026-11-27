//! Retry-safe exact-source and package patch gates.

use crate::publication_gate::{ClaimRefusal, PublicationClaim, PublicationGate};
use crate::{Error, Result};

use super::snapshot::Source;
use super::{Limits, Snapshot};

/// Reversible, exact-source content-control XML replacement.
#[derive(Debug, Clone)]
pub(crate) struct DraftPatch {
    before: Source,
    after: Source,
    limits: Limits,
    gate: Gate,
}

impl DraftPatch {
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
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
pub(crate) struct Gate(PublicationGate);

impl Gate {
    pub(crate) fn new() -> Self {
        Self(PublicationGate::new())
    }

    pub(crate) fn is_applied(&self) -> bool {
        self.0.is_applied()
    }

    pub(crate) fn claim(&self, subject: &str) -> Result<PublicationClaim> {
        self.0.claim().map_err(|refusal| match refusal {
            ClaimRefusal::InFlight => {
                Error::Invalid(format!("{subject} publication is already in flight"))
            },
            ClaimRefusal::Applied => Error::Invalid(format!("{subject} was already applied")),
        })
    }
}
