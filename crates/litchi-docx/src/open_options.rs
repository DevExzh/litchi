//! Configuration for DOCX ingress.

use crate::SourceLimits;
use litchi_core::ExecutionContext;
use litchi_opc::{SourceCacheLimits, SourceReadPolicy};

/// Policies for owning and positional-source DOCX ingress.
///
/// Cache, read-ahead and execution settings apply to
/// [`crate::source_backed::Package::from_read_at_with_options`] and the path,
/// reader and byte-slice counterparts. Owning constructors reject these settings before input reads;
/// their payloads cannot yet retain the requested lazy resource policies.
///
/// Source limits bound the clear OPC archive. With the `encryption` feature,
/// encryption limits independently bound the outer encrypted container and
/// decryption work on password-bearing ingress routes. Passwords are supplied
/// to those routes, never retained in this configuration.
///
/// # Examples
///
/// ```
/// use litchi_docx::{OpenOptions, Package, SourceLimits};
/// # fn example(bytes: &[u8]) -> litchi_docx::Result<Package> {
/// let limits = SourceLimits::builder().max_input_bytes(32 * 1024 * 1024)?.build()?;
/// Package::from_bytes_with_options(bytes, OpenOptions::default().source_limits(limits))
/// # }
/// ```
///
/// Defaults retain the standard finite source and encryption policies.
/// Setters consume already constructed policy values and do not perform I/O.
#[derive(Debug, Clone, Default)]
pub struct OpenOptions {
    pub(crate) source_limits: SourceLimits,
    pub(crate) cache_limits: Option<SourceCacheLimits>,
    pub(crate) source_read_policy: Option<SourceReadPolicy>,
    pub(crate) execution_context: Option<ExecutionContext>,
    pub(crate) semantic_limits: crate::SemanticLimits,
    #[cfg(feature = "encryption")]
    pub(crate) encryption_limits: crate::encryption::Limits,
}

impl OpenOptions {
    /// Select the validated physical limits for the clear OPC archive.
    #[must_use]
    pub const fn source_limits(mut self, limits: SourceLimits) -> Self {
        self.source_limits = limits;
        self
    }

    /// Select a validated finite payload-cache policy for lazy ingress.
    #[must_use]
    pub const fn cache_limits(mut self, limits: SourceCacheLimits) -> Self {
        self.cache_limits = Some(limits);
        self
    }

    /// Select exact reads or validated bounded read-ahead for lazy ingress.
    /// Publication disables read-ahead and releases its retained reservation.
    #[must_use]
    pub const fn source_read_policy(mut self, policy: SourceReadPolicy) -> Self {
        self.source_read_policy = Some(policy);
        self
    }

    /// Use the caller's execution context for opening and deferred reads.
    /// Cloned options share its budget and cancellation token.
    #[must_use]
    pub fn execution_context(mut self, context: ExecutionContext) -> Self {
        self.execution_context = Some(context);
        self
    }

    /// Bound the opened package's semantic reads by `limits` (record 0355);
    /// the package keeps them for every bounded read and edit.
    #[must_use]
    pub fn semantic_limits(mut self, limits: crate::SemanticLimits) -> Self {
        self.semantic_limits = limits;
        self
    }

    /// The configured semantic read budgets.
    #[must_use]
    pub const fn semantic_limits_value(&self) -> &crate::SemanticLimits {
        &self.semantic_limits
    }

    pub(crate) fn check_owning(&self) -> crate::Result<()> {
        let policy = if self.execution_context.is_some() {
            Some(OpenPolicy::ExecutionContext)
        } else if self.cache_limits.is_some() {
            Some(OpenPolicy::CacheLimits)
        } else if self.source_read_policy.is_some() {
            Some(OpenPolicy::SourceReadPolicy)
        } else {
            None
        };
        if let Some(policy) = policy {
            return Err(crate::Error::UnsupportedOpenPolicy { policy });
        }
        Ok(())
    }

    /// Select the outer encryption policy for password-bearing ingress.
    ///
    /// The encryption owner validates this policy during decryption. This
    /// setting does not replace the clear archive's source limits and has no
    /// effect when opening plaintext or adopting an already-decrypted package.
    #[cfg(feature = "encryption")]
    #[must_use]
    pub const fn encryption_limits(mut self, limits: crate::encryption::Limits) -> Self {
        self.encryption_limits = limits;
        self
    }
}

/// A resource policy that requires positional-source DOCX ownership.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum OpenPolicy {
    /// Retention in the finite deferred payload cache.
    CacheLimits,
    /// Exact or bounded read-ahead source access.
    SourceReadPolicy,
    /// Hierarchical budgets and cooperative cancellation.
    ExecutionContext,
}
