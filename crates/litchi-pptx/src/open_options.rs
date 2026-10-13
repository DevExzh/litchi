//! Configuration for PPTX ingress.

use litchi_core::ExecutionContext;
use litchi_opc::{SourceCacheLimits, SourceLimits};

/// Policies for owning and positional-source PPTX ingress.
///
/// One envelope serves every configured open route:
///
/// - the owning [`crate::Package`] (`open_with_options`,
///   `from_reader_with_options`, `from_bytes_with_options` and, with the
///   `encryption` feature, the `*_with_password_and_options` routes);
/// - the positional-source owners
///   [`crate::SourceBackedPresentation::open_with_options`] and
///   [`crate::SourceBackedPresentationEditor::open_with_options`].
///
/// Source limits bound the clear OPC archive. Cache and execution settings
/// apply only to positional-source owners; owning constructors reject them
/// with [`crate::Error::UnsupportedOpenPolicy`] before any input is read,
/// because an owning package retains no deferred payload cache or execution
/// session to charge them to. With the `encryption` feature, encryption
/// limits independently bound the outer encrypted container and decryption
/// work on password-bearing routes. Passwords are arguments of those routes
/// and are never retained here.
///
/// # Examples
///
/// ```
/// use litchi_pptx::{OpenOptions, Package, SourceLimits};
/// # fn example(bytes: &[u8]) -> litchi_pptx::Result<Package> {
/// let limits = SourceLimits::builder().max_input_bytes(32 * 1024 * 1024)?.build()?;
/// Package::from_bytes_with_options(bytes, OpenOptions::default().source_limits(limits))
/// # }
/// ```
///
/// Defaults retain the standard finite source and encryption policies.
/// Setters consume already validated policy values and perform no I/O.
#[derive(Debug, Clone, Default)]
pub struct OpenOptions {
    pub(crate) source_limits: SourceLimits,
    pub(crate) cache_limits: Option<SourceCacheLimits>,
    pub(crate) execution_context: Option<ExecutionContext>,
    #[cfg(feature = "encryption")]
    pub(crate) encryption_limits: litchi_crypto::ooxml::Limits,
}

impl OpenOptions {
    /// Select the validated physical limits for the clear OPC archive.
    #[must_use]
    pub const fn source_limits(mut self, limits: SourceLimits) -> Self {
        self.source_limits = limits;
        self
    }

    /// Select a validated finite payload-cache policy for positional-source
    /// ingress.
    #[must_use]
    pub const fn cache_limits(mut self, limits: SourceCacheLimits) -> Self {
        self.cache_limits = Some(limits);
        self
    }

    /// Use the caller's execution context for positional-source opening and
    /// deferred reads. Cloned options share its budget and cancellation token.
    #[must_use]
    pub fn execution_context(mut self, context: ExecutionContext) -> Self {
        self.execution_context = Some(context);
        self
    }

    /// Select the outer encryption policy for password-bearing ingress.
    ///
    /// The encryption owner validates this policy during decryption. It does
    /// not replace the clear archive's source limits and has no effect on
    /// plaintext ingress.
    #[cfg(feature = "encryption")]
    #[must_use]
    pub const fn encryption_limits(mut self, limits: litchi_crypto::ooxml::Limits) -> Self {
        self.encryption_limits = limits;
        self
    }

    /// Refuse lazy-ingress policies on an owning route.
    pub(crate) fn check_owning(&self) -> crate::Result<()> {
        let policy = if self.execution_context.is_some() {
            Some(OpenPolicy::ExecutionContext)
        } else if self.cache_limits.is_some() {
            Some(OpenPolicy::CacheLimits)
        } else {
            None
        };
        match policy {
            Some(policy) => Err(crate::Error::UnsupportedOpenPolicy { policy }),
            None => Ok(()),
        }
    }
}

/// A resource policy that requires positional-source PPTX ownership.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum OpenPolicy {
    /// Retention in the finite deferred payload cache.
    CacheLimits,
    /// Hierarchical budgets and cooperative cancellation.
    ExecutionContext,
}

impl std::fmt::Display for OpenPolicy {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::CacheLimits => "payload cache limits",
            Self::ExecutionContext => "execution context",
        })
    }
}
