//! Shared provenance and downgrade policy for encrypted OOXML packages.
//!
//! This module records whether clear package content came from an encrypted
//! package and prevents silently writing it as an ordinary package. That is a
//! Litchi project policy, not a requirement imposed by `[MS-OFFCRYPTO]`.
//! Cryptographic transforms, passwords, OPC ownership, and serialization stay
//! with their respective layers.

pub use litchi_crypto::ooxml::Mode;
use thiserror::Error;

/// Encryption provenance retained alongside an opened OOXML document.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub struct PackageEncryption {
    mode: Option<Mode>,
}

impl PackageEncryption {
    /// Construct provenance for an ordinary, unencrypted input package.
    #[must_use]
    pub const fn plain() -> Self {
        Self { mode: None }
    }

    /// Construct provenance for an input package encrypted with `mode`.
    #[must_use]
    pub const fn encrypted(mode: Mode) -> Self {
        Self { mode: Some(mode) }
    }

    /// Return the retained input encryption mode, if any.
    #[must_use]
    pub const fn mode(self) -> Option<Mode> {
        self.mode
    }

    /// Enforce the project policy for writing an ordinary package.
    ///
    /// # Errors
    ///
    /// Returns [`PolicyError::OrdinaryOutputWouldRemoveEncryption`] when the
    /// opened input carried encryption provenance.
    pub const fn ordinary_output(self) -> Result<(), PolicyError> {
        match self.mode {
            Some(mode) => Err(PolicyError::OrdinaryOutputWouldRemoveEncryption { mode }),
            None => Ok(()),
        }
    }

    /// Return the mode required to retain encryption on output.
    ///
    /// # Errors
    ///
    /// Returns [`PolicyError::NoRetainedMode`] for provenance created from an
    /// ordinary input package.
    pub const fn require_retained_mode(self) -> Result<Mode, PolicyError> {
        match self.mode {
            Some(mode) => Ok(mode),
            None => Err(PolicyError::NoRetainedMode),
        }
    }

    /// Record the mode applied by a successful encryption operation.
    pub const fn mark_encrypted(&mut self, mode: Mode) {
        self.mode = Some(mode);
    }
}

/// Host-independent failures from package encryption policy checks.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
pub enum PolicyError {
    /// Ordinary output would discard retained encryption provenance.
    #[error("ordinary package output would remove retained {mode} encryption")]
    OrdinaryOutputWouldRemoveEncryption {
        /// Mode that must be retained instead.
        mode: Mode,
    },
    /// No source encryption mode is available to retain.
    #[error("no encryption mode is retained")]
    NoRetainedMode,
}

#[cfg(test)]
mod tests {
    use super::{Mode, PackageEncryption, PolicyError};

    #[test]
    fn plain_provenance_allows_ordinary_output() {
        let provenance = PackageEncryption::plain();

        assert_eq!(provenance.mode(), None);
        assert_eq!(provenance.ordinary_output(), Ok(()));
        assert_eq!(
            provenance.require_retained_mode(),
            Err(PolicyError::NoRetainedMode)
        );
        assert_eq!(PackageEncryption::default(), provenance);
    }

    #[test]
    fn encrypted_provenance_requires_its_retained_mode() {
        let provenance = PackageEncryption::encrypted(Mode::Agile);

        assert_eq!(provenance.mode(), Some(Mode::Agile));
        assert_eq!(provenance.require_retained_mode(), Ok(Mode::Agile));
        assert_eq!(
            provenance.ordinary_output(),
            Err(PolicyError::OrdinaryOutputWouldRemoveEncryption { mode: Mode::Agile })
        );
    }

    #[test]
    fn mark_encrypted_updates_plain_provenance() {
        let mut provenance = PackageEncryption::plain();

        provenance.mark_encrypted(Mode::Standard);

        assert_eq!(provenance.mode(), Some(Mode::Standard));
        assert_eq!(provenance.require_retained_mode(), Ok(Mode::Standard));
    }
}
