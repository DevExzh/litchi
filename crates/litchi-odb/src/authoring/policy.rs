//! Explicit protected-package edit policy.

/// Policy for a source package carrying digital signatures.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum SignaturePolicy {
    /// Preserve exact no-ops and refuse every changed publication.
    #[default]
    PreserveExactOnly,
    /// Publish changed XML while deliberately removing invalidated signature
    /// members through the shared ODF package writer.
    RemoveInvalidated,
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

    /// Selects the signature lifecycle for a changed publication.
    #[must_use]
    pub const fn with_signature(mut self, value: SignaturePolicy) -> Self {
        self.signature = value;
        self
    }

    /// Selects the encryption lifecycle for a changed publication.
    #[must_use]
    pub const fn with_encryption(mut self, value: EncryptionPolicy) -> Self {
        self.encryption = value;
        self
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

/// Stable root-level capability contract for protected ODB publication.
///
/// Re-signing and re-encryption are deliberately unsupported: callers can
/// inspect this value before beginning an edit instead of discovering an
/// implicit or producer-dependent downgrade during publication.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ProtectionCapabilities {
    signature: SignatureCapability,
    encryption: EncryptionCapability,
}

/// Supported signature lifecycle for an ODB package.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum SignatureCapability {
    /// Verify signature math or remove invalidated signatures, but never sign.
    VerificationAndInvalidationRemovalOnly,
}

/// Supported encryption lifecycle for an ODB package.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum EncryptionCapability {
    /// Preserve an encrypted package only when publication is byte-exact.
    ExactPreservationOnly,
}

impl SignatureCapability {
    /// Whether signature math can be verified without making a trust claim.
    #[must_use]
    pub const fn can_verify(self) -> bool {
        match self {
            Self::VerificationAndInvalidationRemovalOnly => true,
        }
    }

    /// Whether invalidated signature members can be explicitly removed.
    #[must_use]
    pub const fn can_remove_invalidated(self) -> bool {
        match self {
            Self::VerificationAndInvalidationRemovalOnly => true,
        }
    }

    /// Whether a new signature can be produced.
    #[must_use]
    pub const fn can_re_sign(self) -> bool {
        match self {
            Self::VerificationAndInvalidationRemovalOnly => false,
        }
    }
}

impl EncryptionCapability {
    /// Whether a changed package can be newly encrypted.
    #[must_use]
    pub const fn can_re_encrypt(self) -> bool {
        match self {
            Self::ExactPreservationOnly => false,
        }
    }
}

/// Protection inventory before and after one committed publication.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ProtectionTransition {
    before: ProtectionStatus,
    after: ProtectionStatus,
}

/// Explicit handling for modeled objects that depend on a removed owner.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum DependencyDisposition {
    /// Refuse removal or transfer while a modeled dependency would orphan.
    #[default]
    Refuse,
    /// Remove modeled dependents during deletion or recursively copy missing
    /// modeled dependencies during transfer.
    Cascade,
}

/// Explicit handling for potentially active content linked by a component.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum ActiveContentDisposition {
    /// Refuse a transfer when the linked subtree contains active declarations.
    #[default]
    Refuse,
    /// Copy the bytes inertly without interpreting or activating them.
    CopyInert,
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

impl ProtectionCapabilities {
    pub(crate) const fn odb() -> Self {
        Self {
            signature: SignatureCapability::VerificationAndInvalidationRemovalOnly,
            encryption: EncryptionCapability::ExactPreservationOnly,
        }
    }

    /// Returns the typed signature lifecycle.
    #[must_use]
    pub const fn signature(self) -> SignatureCapability {
        self.signature
    }

    /// Returns the typed encryption lifecycle.
    #[must_use]
    pub const fn encryption(self) -> EncryptionCapability {
        self.encryption
    }

    /// Whether document-signature math can be verified without a trust claim.
    #[must_use]
    pub const fn can_verify_signatures(self) -> bool {
        self.signature.can_verify()
    }

    /// Whether a changed publication may explicitly remove invalid signatures.
    #[must_use]
    pub const fn can_remove_invalidated_signatures(self) -> bool {
        self.signature.can_remove_invalidated()
    }

    /// Whether this crate can produce a new database package signature.
    #[must_use]
    pub const fn can_re_sign(self) -> bool {
        self.signature.can_re_sign()
    }

    /// Whether this crate can produce a newly encrypted database package.
    #[must_use]
    pub const fn can_re_encrypt(self) -> bool {
        self.encryption.can_re_encrypt()
    }
}

impl ProtectionTransition {
    pub(crate) const fn new(before: ProtectionStatus, after: ProtectionStatus) -> Self {
        Self { before, after }
    }

    /// Returns the source-package inventory.
    #[must_use]
    pub const fn before(self) -> ProtectionStatus {
        self.before
    }

    /// Returns the committed-package inventory.
    #[must_use]
    pub const fn after(self) -> ProtectionStatus {
        self.after
    }

    /// Whether a source signature was deliberately removed after mutation.
    #[must_use]
    pub const fn signature_was_removed(self) -> bool {
        self.before.signed && !self.after.signed
    }
}
