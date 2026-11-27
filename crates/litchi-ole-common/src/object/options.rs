//! Ingress and edit policies for the immutable OLE object package.

use super::{Limits, Targets};
use litchi_cfb::{OleFileLimits, SectorLayoutPolicy};

/// Caller-selected policy for opening one object package.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpenOptions {
    targets: Targets,
    limits: Limits,
    cfb_limits: OleFileLimits,
}

impl Default for OpenOptions {
    fn default() -> Self {
        Self::new()
    }
}

impl OpenOptions {
    /// Starts from the finite default object and CFB policies.
    #[must_use]
    pub fn new() -> Self {
        Self {
            targets: Targets::default(),
            limits: Limits::default(),
            cfb_limits: OleFileLimits::default(),
        }
    }

    /// Select the host-owned storage targets.
    #[must_use]
    pub fn targets(mut self, targets: Targets) -> Self {
        self.targets = targets;
        self
    }

    /// Select the bounded object-capture policy.
    #[must_use]
    pub const fn limits(mut self, limits: Limits) -> Self {
        self.limits = limits;
        self
    }

    /// Select the bounded low-level CFB admission policy.
    #[must_use]
    pub const fn cfb_limits(mut self, limits: OleFileLimits) -> Self {
        self.cfb_limits = limits;
        self
    }

    /// Returns the selected targets.
    #[must_use]
    pub fn targets_value(&self) -> &Targets {
        &self.targets
    }

    /// Returns the selected object-capture policy.
    #[must_use]
    pub const fn limits_value(&self) -> Limits {
        self.limits
    }

    /// Returns the selected low-level CFB admission policy.
    #[must_use]
    pub const fn cfb_limits_value(&self) -> OleFileLimits {
        self.cfb_limits
    }
}

/// Caller-selected policy for one edit of an opened object package.
///
/// Unset numeric fields inherit the package's admitted policies. Calling a
/// setter marks that field as an explicit narrower request; the package still
/// intersects it with the retained admission guarantee, so an edit cannot
/// widen the package's limits.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EditOptions {
    limits: Limits,
    limits_overridden: bool,
    cfb_limits: OleFileLimits,
    cfb_limits_overridden: bool,
    sector_layout: SectorLayoutPolicy,
    sector_layout_overridden: bool,
}

impl Default for EditOptions {
    fn default() -> Self {
        Self::new()
    }
}

impl EditOptions {
    /// Starts with inherited package limits and the default sector layout.
    #[must_use]
    pub fn new() -> Self {
        Self {
            limits: Limits::default(),
            limits_overridden: false,
            cfb_limits: OleFileLimits::default(),
            cfb_limits_overridden: false,
            sector_layout: SectorLayoutPolicy::default(),
            sector_layout_overridden: false,
        }
    }

    /// Select the object-capture bounds for this edit.
    #[must_use]
    pub const fn limits(mut self, limits: Limits) -> Self {
        self.limits = limits;
        self.limits_overridden = true;
        self
    }

    /// Select the low-level CFB bounds for this edit.
    #[must_use]
    pub const fn cfb_limits(mut self, limits: OleFileLimits) -> Self {
        self.cfb_limits = limits;
        self.cfb_limits_overridden = true;
        self
    }

    /// Select the physical sector layout used by publication.
    #[must_use]
    pub const fn sector_layout(mut self, policy: SectorLayoutPolicy) -> Self {
        self.sector_layout = policy;
        self.sector_layout_overridden = true;
        self
    }

    /// Returns the requested object-capture bounds.
    #[must_use]
    pub const fn limits_value(&self) -> Limits {
        self.limits
    }

    /// Returns the requested low-level CFB bounds.
    #[must_use]
    pub const fn cfb_limits_value(&self) -> OleFileLimits {
        self.cfb_limits
    }

    /// Returns the requested physical sector layout.
    #[must_use]
    pub const fn sector_layout_value(&self) -> SectorLayoutPolicy {
        self.sector_layout
    }

    pub(crate) const fn has_limits_override(&self) -> bool {
        self.limits_overridden
    }

    pub(crate) const fn has_cfb_limits_override(&self) -> bool {
        self.cfb_limits_overridden
    }

    pub(crate) const fn has_sector_layout_override(&self) -> bool {
        self.sector_layout_overridden
    }
}
