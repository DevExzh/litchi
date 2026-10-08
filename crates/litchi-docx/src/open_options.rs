//! Configuration for opening an owning DOCX package.

use crate::SourceLimits;

/// Policies applied while opening an owning [`crate::Package`].
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
#[derive(Debug, Clone, Copy, Default)]
pub struct OpenOptions {
    pub(crate) source_limits: SourceLimits,
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
