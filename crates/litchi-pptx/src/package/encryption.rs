//! Optional managed `[MS-OFFCRYPTO]` package encryption.

use std::fs::File;
use std::io::{Read, Write};
use std::path::Path;

use litchi_crypto::ooxml::{self, Limits, Mode};
use litchi_ooxml_common::package_encryption::PackageEncryption;
use litchi_opc::{PackageWriter, SourceLimits};

use super::Package;
use crate::{Error, Result};

impl Package {
    /// Return the mode retained from encrypted ingress or the latest
    /// successful encrypted file save.
    #[must_use]
    pub const fn encryption(&self) -> Option<Mode> {
        self.encryption.mode()
    }

    /// Alias for [`Self::encryption`]. Byte-only encryption and explicit plain
    /// output do not change this source-or-successful-save mode.
    #[must_use]
    pub const fn encryption_mode(&self) -> Option<Mode> {
        self.encryption()
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
        let opened = ooxml::load_with(reader, password, &options.encryption_limits)?;
        Self::from_opened_encryption(opened, options.source_limits)
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
        let opened = ooxml::open_with(owned, password, &options.encryption_limits)?;
        Self::from_opened_encryption(opened, options.source_limits)
    }

    /// Explicitly serialize a clear OPC package, permitting an encryption downgrade.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_plain_bytes(&mut self) -> Result<Vec<u8>> {
        self.flush_presentation()?;
        let bytes = PackageWriter::to_bytes(&self.opc)?;
        Ok(bytes)
    }

    /// Explicitly save a clear OPC package, permitting an encryption downgrade.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn save_plain<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
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
        &mut self,
        path: P,
        durability: litchi_core::Durability,
    ) -> Result<()> {
        self.flush_presentation()?;
        PackageWriter::write_with_durability(path, &self.opc, durability)?;
        Ok(())
    }

    /// Serialize and encrypt the package in the selected mode.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_encrypted(&mut self, password: &str, mode: Mode) -> Result<Vec<u8>> {
        self.to_encrypted_with_limits(password, mode, Limits::default())
    }

    /// Serialize and encrypt the package in the selected mode under explicit limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_encrypted_with_limits(
        &mut self,
        password: &str,
        mode: Mode,
        limits: Limits,
    ) -> Result<Vec<u8>> {
        self.encrypt_inner(password, mode, &limits)
    }

    /// Compatibility alias for [`Self::to_encrypted`].
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_encrypted_bytes(&mut self, password: &str, mode: Mode) -> Result<Vec<u8>> {
        self.to_encrypted(password, mode)
    }

    /// Serialize and encrypt using the retained input/output mode.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_reencrypted(&mut self, password: &str) -> Result<Vec<u8>> {
        self.to_reencrypted_with_limits(password, Limits::default())
    }

    /// Serialize and encrypt using the retained mode under explicit limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_reencrypted_with_limits(
        &mut self,
        password: &str,
        limits: Limits,
    ) -> Result<Vec<u8>> {
        let mode =
            self.encryption
                .require_retained_mode()
                .map_err(|source| Error::EncryptionPolicy {
                    operation: "to_reencrypted",
                    source,
                })?;
        let bytes = self.encrypt_inner(password, mode, &limits)?;
        Ok(bytes)
    }

    /// Write an immutable opened package encrypted in its retained mode.
    ///
    /// The clear ZIP is staged under `limits.max_plaintext_bytes` before the
    /// native crypto writer runs. The writer receives ciphertext only after
    /// encryption succeeds.
    /// This does not change encryption provenance or flush mutable authoring.
    ///
    /// # Errors
    /// Refuses pending authoring state, a stale graph or absent retained mode.
    /// Crypto policy, allocation and sink failures retain their typed sources.
    pub fn write_reencrypted(
        &self,
        mut writer: impl Write,
        password: &str,
        limits: Limits,
    ) -> Result<()> {
        self.ensure_graph_current("write_reencrypted")?;
        if self.mutable_pres.is_some() {
            return Err(Error::UnsafeEdit {
                operation: "write_reencrypted",
                reason: "immutable publication requires an already-opened package",
            });
        }
        let mode =
            self.encryption
                .require_retained_mode()
                .map_err(|source| Error::EncryptionPolicy {
                    operation: "write_reencrypted",
                    source,
                })?;
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
        let encrypted =
            ooxml::encrypt_with(std::mem::take(&mut *clear.bytes), password, mode, &limits)?;
        writer.write_all(&encrypted)?;
        writer.flush()?;
        Ok(())
    }

    /// Compatibility alias for [`Self::to_reencrypted`].
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_reencrypted_bytes(&mut self, password: &str) -> Result<Vec<u8>> {
        self.to_reencrypted(password)
    }

    /// Atomically save the package encrypted in the selected mode.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn save_encrypted<P: AsRef<Path>>(
        &mut self,
        path: P,
        password: &str,
        mode: Mode,
    ) -> Result<()> {
        self.save_encrypted_with_limits(path, password, mode, Limits::default())
    }

    /// Atomically save selected-mode encryption under explicit crypto limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn save_encrypted_with_limits<P: AsRef<Path>>(
        &mut self,
        path: P,
        password: &str,
        mode: Mode,
        limits: Limits,
    ) -> Result<()> {
        let bytes = self.encrypt_inner(password, mode, &limits)?;
        atomic_replace(path.as_ref(), &bytes)?;
        self.encryption.mark_encrypted(mode);
        Ok(())
    }

    /// Atomically save the package encrypted in its retained mode.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn save_reencrypted<P: AsRef<Path>>(&mut self, path: P, password: &str) -> Result<()> {
        self.save_reencrypted_with_limits(path, password, Limits::default())
    }

    /// Atomically save retained-mode encryption under explicit crypto limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn save_reencrypted_with_limits<P: AsRef<Path>>(
        &mut self,
        path: P,
        password: &str,
        limits: Limits,
    ) -> Result<()> {
        let mode =
            self.encryption
                .require_retained_mode()
                .map_err(|source| Error::EncryptionPolicy {
                    operation: "save_reencrypted",
                    source,
                })?;
        let bytes = self.encrypt_inner(password, mode, &limits)?;
        atomic_replace(path.as_ref(), &bytes)?;
        self.encryption.mark_encrypted(mode);
        Ok(())
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
    /// Returns an error if the clear package violates `limits` or is not a
    /// valid PPTX package.
    pub fn from_decrypted(opened: ooxml::Opened, limits: SourceLimits) -> Result<Self> {
        Self::from_opened_encryption(opened, limits)
    }

    fn from_opened_encryption(opened: ooxml::Opened, limits: SourceLimits) -> Result<Self> {
        let mode = opened.mode();
        let mut package = Self::from_opc_package_with_provenance(
            litchi_opc::OpcPackage::from_vec_with_limits(opened.into_bytes(), limits)?,
            true,
        )?;
        package.encryption = match mode {
            Some(mode) => PackageEncryption::encrypted(mode),
            None => PackageEncryption::plain(),
        };
        Ok(package)
    }

    fn encrypt_inner(&mut self, password: &str, mode: Mode, limits: &Limits) -> Result<Vec<u8>> {
        self.flush_presentation()?;
        let clear = PackageWriter::to_bytes(&self.opc)?;
        Ok(ooxml::encrypt_with(clear, password, mode, limits)?)
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
