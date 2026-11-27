use super::{Change, PackageLimits, Snapshot};
use crate::package::{Error, Result};
use std::sync::Arc;

/// Deterministic fast precheck; exact bytes are always checked as well.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Revision(u64);

impl Revision {
    #[must_use]
    pub const fn value(self) -> u64 {
        self.0
    }
    pub(crate) fn from_bytes(bytes: &[u8]) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325u64;
        for byte in bytes {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Self(value)
    }
}

/// Reversible, exact-source whole-CFB patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Patch {
    pub(crate) base: Revision,
    pub(crate) target: Revision,
    pub(crate) before: Arc<[u8]>,
    pub(crate) after: Arc<[u8]>,
    pub(crate) changes: Vec<Change>,
    pub(crate) limits: PackageLimits,
}

#[cfg_attr(
    not(test),
    expect(
        dead_code,
        reason = "record-level accessors the unit tests exercise; the public family is the package-level Edit"
    )
)]
impl Patch {
    /// Apply once, or accept an exact already-applied target for retry safety.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn apply(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.revision() == self.target && current.bytes() == self.after.as_ref() {
            return Ok(current.clone());
        }
        if current.revision() != self.base || current.bytes() != self.before.as_ref() {
            return Err(Error::InvalidFormat(
                "cannot apply font patch to a different CFB source".into(),
            ));
        }
        Snapshot::from_arc(self.after.clone(), self.limits)
    }

    /// Revert to the base source, or accept the exact base for retry safety.
    ///
    /// # Errors
    ///
    /// Returns an error if `current` is neither the target nor the exact base
    /// source of this patch.
    pub(crate) fn undo(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.revision() == self.base && current.bytes() == self.before.as_ref() {
            return Ok(current.clone());
        }
        if current.revision() != self.target || current.bytes() != self.after.as_ref() {
            return Err(Error::InvalidFormat(
                "cannot undo font patch from a different CFB source".into(),
            ));
        }
        Snapshot::from_arc(self.before.clone(), self.limits)
    }
}
