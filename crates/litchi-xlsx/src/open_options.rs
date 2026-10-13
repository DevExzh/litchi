//! Configuration for XLSX package ingress.

use crate::leniency::Leniency;
use litchi_opc::SourceLimits;

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
/// Defaults retain the standard finite source and encryption policies and
/// the strict leniency. Setters consume already constructed policy values
/// and perform no I/O.
#[derive(Debug, Clone, Copy, Default)]
pub struct OpenOptions {
    pub(crate) source_limits: SourceLimits,
    pub(crate) leniency: Leniency,
    #[cfg(feature = "encryption")]
    pub(crate) encryption_limits: crate::encryption::Limits,
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
}
