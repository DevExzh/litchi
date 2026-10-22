//! Configuration for XLSX package ingress.

use crate::leniency::Leniency;
use litchi_core::ExecutionContext;
use litchi_opc::{SourceCacheLimits, SourceLimits};

/// Policies for opening an XLSX [`crate::Package`].
///
/// Source limits bound the clear OPC archive; the leniency selects the
/// opt-in repairs described in [`crate::leniency`]. With the `encryption`
/// feature, encryption limits independently bound the outer encrypted
/// container and decryption work on the password-bearing ingress routes.
/// Passwords are supplied to those routes and are never retained here.
///
/// ```
/// use litchi_xlsx::{OpenOptions, Package, SourceLimits};
/// # fn example(bytes: &[u8]) -> litchi_xlsx::Result<Package> {
/// let limits = SourceLimits::builder().max_input_bytes(32 * 1024 * 1024)?.build()?;
/// Package::from_bytes_with_options(bytes, OpenOptions::new().source_limits(limits))
/// # }
/// ```
///
/// The source-backed owner ([`crate::SourceBackedWorkbook`]) also reads the
/// deferred-payload cache limits and the caller-owned execution context; an
/// owning [`crate::Package`] ignores both.
///
/// Defaults retain the standard finite source and encryption policies and
/// the strict leniency. Setters consume already constructed policy values
/// and perform no I/O.
#[derive(Debug, Clone, Default)]
pub struct OpenOptions {
    pub(crate) source_limits: SourceLimits,
    pub(crate) leniency: Leniency,
    pub(crate) semantic_limits: crate::SemanticLimits,
    #[cfg(feature = "encryption")]
    pub(crate) encryption_limits: crate::encryption::Limits,
    pub(crate) cache_limits: Option<SourceCacheLimits>,
    pub(crate) execution_context: Option<ExecutionContext>,
}

impl OpenOptions {
    /// The default ingress policies.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Select the validated physical limits for the clear OPC archive.
    #[must_use]
    pub const fn source_limits(mut self, limits: SourceLimits) -> Self {
        self.source_limits = limits;
        self
    }

    /// Select the opt-in repairs applied while the workbook is read.
    #[must_use]
    pub const fn leniency(mut self, leniency: Leniency) -> Self {
        self.leniency = leniency;
        self
    }

    /// Select the semantic budgets the package keeps for its bounded reads
    /// and edits ([`crate::semantic_limits`]).
    #[must_use]
    pub const fn semantic_limits(mut self, limits: crate::SemanticLimits) -> Self {
        self.semantic_limits = limits;
        self
    }

    /// The semantic budgets the package keeps.
    #[must_use]
    pub const fn semantic_limits_value(&self) -> &crate::SemanticLimits {
        &self.semantic_limits
    }

    /// The physical limits of the clear OPC archive.
    #[must_use]
    pub const fn source_limits_value(&self) -> SourceLimits {
        self.source_limits
    }

    /// The opt-in repairs applied while the workbook is read.
    #[must_use]
    pub const fn leniency_value(&self) -> Leniency {
        self.leniency
    }

    /// The outer encryption policy of password-bearing ingress, which an
    /// opened package keeps for its encrypted outputs.
    #[cfg(feature = "encryption")]
    #[must_use]
    pub const fn encryption_limits_value(&self) -> crate::encryption::Limits {
        self.encryption_limits
    }

    /// Select the outer encryption policy for password-bearing ingress.
    ///
    /// The encryption owner validates this policy during decryption. It does
    /// not replace the clear archive's source limits and has no effect when
    /// opening plaintext.
    #[cfg(feature = "encryption")]
    #[must_use]
    pub const fn encryption_limits(mut self, limits: crate::encryption::Limits) -> Self {
        self.encryption_limits = limits;
        self
    }

    /// Select the finite deferred-payload cache of a source-backed owner
    /// ([`crate::SourceBackedWorkbook`]); the default is the OPC default
    /// cache.
    #[must_use]
    pub const fn cache_limits(mut self, limits: SourceCacheLimits) -> Self {
        self.cache_limits = Some(limits);
        self
    }

    /// The deferred-payload cache limits of a source-backed owner, when
    /// selected.
    #[must_use]
    pub const fn cache_limits_value(&self) -> Option<SourceCacheLimits> {
        self.cache_limits
    }

    /// Select the caller-owned execution context of a source-backed owner
    /// ([`crate::SourceBackedWorkbook`]): cancellation is checked before
    /// mandatory reads and deferred materializations, and retained payloads
    /// are charged to its memory budget.
    #[must_use]
    pub fn execution_context(mut self, context: ExecutionContext) -> Self {
        self.execution_context = Some(context);
        self
    }

    /// The execution context of a source-backed owner, when selected.
    #[must_use]
    pub const fn execution_context_value(&self) -> Option<&ExecutionContext> {
        self.execution_context.as_ref()
    }
}
