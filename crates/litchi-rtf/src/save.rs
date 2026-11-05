//! Atomic publication policy for immutable RTF packages.

use litchi_core::Durability;
pub use litchi_core::atomic::{Error, Overwrite};

/// Explicit per-call publication policy. Defaults to full durability and replacement.
#[derive(Clone, Copy, Debug)]
pub struct Options {
    durability: Durability,
    overwrite: Overwrite,
}

impl Options {
    /// The ordinary atomic save policy.
    #[must_use]
    pub fn new() -> Self {
        Self {
            durability: Durability::Full,
            overwrite: Overwrite::Replace,
        }
    }

    /// Select filesystem synchronization for this call.
    #[must_use]
    pub const fn durability(mut self, value: Durability) -> Self {
        self.durability = value;
        self
    }

    /// Select whether an existing destination may be replaced.
    #[must_use]
    pub const fn overwrite(mut self, value: Overwrite) -> Self {
        self.overwrite = value;
        self
    }

    /// Selected synchronization policy.
    #[must_use]
    pub const fn durability_value(self) -> Durability {
        self.durability
    }

    /// Selected replacement policy.
    #[must_use]
    pub const fn overwrite_value(self) -> Overwrite {
        self.overwrite
    }
}

impl Default for Options {
    fn default() -> Self {
        Self::new()
    }
}
