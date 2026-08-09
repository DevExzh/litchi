//! Explicit protected-package edit policy.

/// Policy for a source package carrying digital signatures.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum SignaturePolicy {
    /// Preserve exact no-ops and refuse every changed publication.
    #[default]
    PreserveExactOnly,
}

/// Policy for a source package carrying encrypted entries.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum EncryptionPolicy {
    /// Preserve exact no-ops and refuse every changed publication.
    #[default]
    PreserveExactOnly,
}

/// Explicit policy applied before a changed ODB package is published.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct EditPolicy {
    signature: SignaturePolicy,
    encryption: EncryptionPolicy,
}

impl EditPolicy {
    /// Creates the fail-closed protected-package policy.
    #[must_use]
    pub const fn preserve_exact_only() -> Self {
        Self {
            signature: SignaturePolicy::PreserveExactOnly,
            encryption: EncryptionPolicy::PreserveExactOnly,
        }
    }

    /// Returns the configured signature policy.
    #[must_use]
    pub const fn signature(self) -> SignaturePolicy {
        self.signature
    }

    /// Returns the configured encryption policy.
    #[must_use]
    pub const fn encryption(self) -> EncryptionPolicy {
        self.encryption
    }
}

/// Inert protection inventory for an opened package.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct ProtectionStatus {
    signed: bool,
    encrypted: bool,
}

/// Explicit handling for modeled objects that depend on a removed owner.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum DependencyDisposition {
    /// Refuse while any modeled incoming or nested dependency would orphan.
    #[default]
    Refuse,
    /// Remove the selected owner and the modeled dependent key/index owners.
    Cascade,
}

impl ProtectionStatus {
    pub(crate) const fn new(signed: bool, encrypted: bool) -> Self {
        Self { signed, encrypted }
    }

    /// Whether a package signature member is present.
    #[must_use]
    pub const fn is_signed(self) -> bool {
        self.signed
    }

    /// Whether the ODF manifest declares encrypted entries.
    #[must_use]
    pub const fn is_encrypted(self) -> bool {
        self.encrypted
    }
}
