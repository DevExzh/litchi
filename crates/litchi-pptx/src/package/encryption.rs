//! Optional managed `[MS-OFFCRYPTO]` package encryption.

use std::fs::File;
use std::io::{Read, Write};
use std::path::Path;

use litchi_crypto::ooxml::{self, Mode};
use litchi_ooxml_common::package_encryption::PackageEncryption;
use litchi_opc::{PackageWriter, SourceLimits};

use super::Package;
use crate::{Error, Result};

impl Package {
    /// Return the mode retained from encrypted ingress.
    #[must_use]
    pub const fn encryption(&self) -> Option<Mode> {
        self.encryption.mode()
    }

    /// Open an ordinary or encrypted package using a password and safe limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn open_with_password<P: AsRef<Path>>(path: P, password: &str) -> Result<Self> {
        Self::open_with_password_and_options(path, password, crate::OpenOptions::default())
    }

    /// Open an ordinary or encrypted package with explicit ingress options.
    ///
    /// The options' encryption limits bound the outer container and
    /// decryption; their source limits bound the clear OPC archive.
    ///
    /// # Errors
    ///
    /// Returns [`Error::UnsupportedOpenPolicy`] for cache or execution
    /// policies, and an error if the input cannot be read or is malformed.
    pub fn open_with_password_and_options<P: AsRef<Path>>(
        path: P,
        password: &str,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        options.check_owning()?;
        let input = File::open(path)?;
        Self::from_reader_with_password_and_options(input, password, options)
    }

    /// Read an ordinary or encrypted package using a password and safe limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_reader_with_password<R: Read>(reader: R, password: &str) -> Result<Self> {
        Self::from_reader_with_password_and_options(reader, password, crate::OpenOptions::default())
    }

    /// Read an ordinary or encrypted package with explicit ingress options.
    ///
    /// # Errors
    ///
    /// Returns [`Error::UnsupportedOpenPolicy`] for cache or execution
    /// policies, and an error if the input cannot be read or is malformed.
    pub fn from_reader_with_password_and_options<R: Read>(
        reader: R,
        password: &str,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        options.check_owning()?;
        let opened = ooxml::load(reader, password, &options.encryption_limits)?;
        let mut package = Self::from_opened_encryption(opened, options.source_limits)?;
        package.encryption_limits = options.encryption_limits;
        Ok(package)
    }

    /// Open a borrowed ordinary or encrypted package buffer using a password.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_bytes_with_password(bytes: &[u8], password: &str) -> Result<Self> {
        Self::from_bytes_with_password_and_options(bytes, password, crate::OpenOptions::default())
    }

    /// Open a borrowed package buffer with explicit ingress options.
    ///
    /// # Errors
    ///
    /// Returns [`Error::UnsupportedOpenPolicy`] for cache or execution
    /// policies, and an error if the input cannot be read or is malformed.
    pub fn from_bytes_with_password_and_options(
        bytes: &[u8],
        password: &str,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        options.check_owning()?;
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(bytes.len())
            .map_err(|source| Error::Allocation {
                resource: "encrypted package input",
                source,
            })?;
        owned.extend_from_slice(bytes);
        let opened = ooxml::open(owned, password, &options.encryption_limits)?;
        let mut package = Self::from_opened_encryption(opened, options.source_limits)?;
        package.encryption_limits = options.encryption_limits;
        Ok(package)
    }

    /// Explicitly serialize a clear OPC package, permitting an encryption downgrade.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_plain_bytes(&self) -> Result<Vec<u8>> {
        Ok(PackageWriter::to_bytes(&self.opc)?)
    }

    /// Explicitly save a clear OPC package, permitting an encryption downgrade.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn save_plain<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        self.save_plain_with_durability(path, litchi_core::Durability::Full)
    }

    /// Explicitly save a clear OPC package at a caller-chosen
    /// [`Durability`](litchi_core::Durability), permitting an encryption
    /// downgrade; see [`Self::save_with_durability`].
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn save_plain_with_durability<P: AsRef<Path>>(
        &self,
        path: P,
        durability: litchi_core::Durability,
    ) -> Result<()> {
        PackageWriter::write_with_durability(path, &self.opc, durability)?;
        Ok(())
    }

    /// Serialize and encrypt the package in the selected mode, under the
    /// encryption limits the package was opened with.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_encrypted(&self, password: &str, mode: Mode) -> Result<Vec<u8>> {
        self.encrypt_inner(password, mode)
    }

    /// Serialize and encrypt using the retained input mode, under the
    /// encryption limits the package was opened with.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package without a retained
    /// mode, and an error if the operation fails.
    pub fn to_reencrypted(&self, password: &str) -> Result<Vec<u8>> {
        let mode = self.retained_mode("to_reencrypted")?;
        self.encrypt_inner(password, mode)
    }

    /// Write the package encrypted in its retained mode.
    ///
    /// The clear ZIP is staged under the retained encryption limits'
    /// `max_plaintext_bytes` before the native crypto writer runs. The
    /// writer receives ciphertext only after encryption succeeds. This does
    /// not change the package.
    ///
    /// # Errors
    /// Refuses a package without a retained mode. Crypto policy, allocation
    /// and sink failures retain their typed sources.
    pub fn write_reencrypted(&self, mut writer: impl Write, password: &str) -> Result<()> {
        let mode = self.retained_mode("write_reencrypted")?;
        let limits = &self.encryption_limits;
        let mut clear = ClearBuffer {
            bytes: zeroize::Zeroizing::new(Vec::new()),
            limit: limits.max_plaintext_bytes,
            failure: None,
        };
        let staged = PackageWriter::write_to(&mut clear, &self.opc);
        if let Some(error) = clear.failure.take() {
            return Err(error);
        }
        staged?;
        let encrypted = ooxml::encrypt(std::mem::take(&mut *clear.bytes), password, mode, limits)?;
        writer.write_all(&encrypted)?;
        writer.flush()?;
        Ok(())
    }

    /// Atomically save the package encrypted in the selected mode.
    ///
    /// The package is unchanged: its retained mode stays the one it was
    /// opened with.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn save_encrypted<P: AsRef<Path>>(
        &self,
        path: P,
        password: &str,
        mode: Mode,
    ) -> Result<()> {
        let bytes = self.encrypt_inner(password, mode)?;
        atomic_replace(path.as_ref(), &bytes)
    }

    /// Atomically save the package encrypted in its retained mode.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package without a retained
    /// mode, and an error if the output cannot be encoded or written.
    pub fn save_reencrypted<P: AsRef<Path>>(&self, path: P, password: &str) -> Result<()> {
        let bytes = self.to_reencrypted(password)?;
        atomic_replace(path.as_ref(), &bytes)
    }

    fn retained_mode(&self, operation: &'static str) -> Result<Mode> {
        self.encryption
            .require_retained_mode()
            .map_err(|source| Error::EncryptionPolicy { operation, source })
    }

    /// Open a package that was already decrypted, keeping its encryption
    /// mode.
    ///
    /// `opened` comes from [`litchi_crypto::ooxml::open`] or one of its limit
    /// and policy variants. The package keeps the source mode exactly as
    /// [`Self::from_reader_with_password_and_options`] does, so plain output
    /// stays refused until the caller chooses re-encryption or plaintext.
    /// This lets one decryption serve both a caller's own classification of
    /// the clear package and the open.
    ///
    /// # Errors
    ///
    /// Returns an error if the clear package violates the selected source
    /// limits or is not a valid PPTX package.
    pub fn from_decrypted(opened: ooxml::Opened, options: crate::OpenOptions) -> Result<Self> {
        options.check_owning()?;
        let mut package = Self::from_opened_encryption(opened, options.source_limits)?;
        package.encryption_limits = options.encryption_limits;
        Ok(package)
    }

    fn from_opened_encryption(opened: ooxml::Opened, limits: SourceLimits) -> Result<Self> {
        let mode = opened.mode();
        let mut package = Self::from_opc_package_with_provenance(
            litchi_opc::OpcPackage::from_vec(opened.into_bytes(), limits)?,
            true,
        )?;
        package.encryption = match mode {
            Some(mode) => PackageEncryption::encrypted(mode),
            None => PackageEncryption::plain(),
        };
        Ok(package)
    }

    fn encrypt_inner(&self, password: &str, mode: Mode) -> Result<Vec<u8>> {
        let clear = PackageWriter::to_bytes(&self.opc)?;
        Ok(ooxml::encrypt(
            clear,
            password,
            mode,
            &self.encryption_limits,
        )?)
    }
}

fn atomic_replace(path: &Path, bytes: &[u8]) -> Result<()> {
    litchi_opc::atomic::replace(path, |temporary| {
        temporary.write_all(bytes)?;
        Ok(())
    })?;
    Ok(())
}

/// Fallible, bounded clear staging; failed serialization erases staged bytes.
struct ClearBuffer {
    bytes: zeroize::Zeroizing<Vec<u8>>,
    limit: usize,
    failure: Option<Error>,
}
impl Write for ClearBuffer {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        let next = self.bytes.len().checked_add(bytes.len());
        if next.is_none_or(|next| next > self.limit) {
            self.failure = Some(Error::from(ooxml::Error::Limit {
                resource: "plaintext",
                actual: next.unwrap_or(usize::MAX) as u64,
                maximum: self.limit as u64,
            }));
            return Err(std::io::Error::other(
                "clear package exceeds encryption limit",
            ));
        }
        let next = next.unwrap_or(self.limit);
        if next > self.bytes.capacity() {
            let capacity = self
                .bytes
                .capacity()
                .saturating_mul(2)
                .max(next)
                .min(self.limit);
            let additional = capacity - self.bytes.len();
            if let Err(source) = self.bytes.try_reserve_exact(additional) {
                self.failure = Some(Error::Allocation {
                    resource: "clear encryption staging",
                    source,
                });
                return Err(std::io::Error::other("clear package allocation failed"));
            }
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }
    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}
