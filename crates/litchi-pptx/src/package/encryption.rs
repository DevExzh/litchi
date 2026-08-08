//! Optional managed `[MS-OFFCRYPTO]` package encryption.

use std::fs::File;
use std::io::{Read, Write};
use std::path::Path;

use litchi_crypto::ooxml::{self, Limits, Mode};
use litchi_ooxml_common::package_encryption::PackageEncryption;
use litchi_opc::{PackageWriter, ReadLimits};

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
    pub fn open_with_password<P: AsRef<Path>>(path: P, password: &str) -> Result<Self> {
        Self::open_with_password_and_limits(
            path,
            password,
            Limits::default(),
            ReadLimits::default(),
        )
    }

    /// Open an ordinary or encrypted package with independent crypto and OPC limits.
    pub fn open_with_password_and_limits<P: AsRef<Path>>(
        path: P,
        password: &str,
        encryption_limits: Limits,
        opc_limits: ReadLimits,
    ) -> Result<Self> {
        let input = File::open(path)?;
        Self::from_reader_with_password_and_limits(input, password, encryption_limits, opc_limits)
    }

    /// Read an ordinary or encrypted package using a password and safe limits.
    pub fn from_reader_with_password<R: Read>(reader: R, password: &str) -> Result<Self> {
        Self::from_reader_with_password_and_limits(
            reader,
            password,
            Limits::default(),
            ReadLimits::default(),
        )
    }

    /// Read an ordinary or encrypted package with independent crypto and OPC limits.
    pub fn from_reader_with_password_and_limits<R: Read>(
        reader: R,
        password: &str,
        encryption_limits: Limits,
        opc_limits: ReadLimits,
    ) -> Result<Self> {
        let opened = ooxml::load_with(reader, password, &encryption_limits)?;
        Self::from_opened_encryption(opened, opc_limits)
    }

    /// Open an owned ordinary or encrypted package buffer using a password.
    pub fn from_vec_with_password(bytes: Vec<u8>, password: &str) -> Result<Self> {
        Self::from_vec_with_password_and_limits(
            bytes,
            password,
            Limits::default(),
            ReadLimits::default(),
        )
    }

    /// Open an owned package buffer with independent crypto and OPC limits.
    pub fn from_vec_with_password_and_limits(
        bytes: Vec<u8>,
        password: &str,
        encryption_limits: Limits,
        opc_limits: ReadLimits,
    ) -> Result<Self> {
        let opened = ooxml::open_with(bytes, password, &encryption_limits)?;
        Self::from_opened_encryption(opened, opc_limits)
    }

    /// Open a borrowed ordinary or encrypted package buffer using a password.
    pub fn from_bytes_with_password(bytes: &[u8], password: &str) -> Result<Self> {
        Self::from_bytes_with_password_and_limits(
            bytes,
            password,
            Limits::default(),
            ReadLimits::default(),
        )
    }

    /// Open a borrowed package buffer with independent crypto and OPC limits.
    pub fn from_bytes_with_password_and_limits(
        bytes: &[u8],
        password: &str,
        encryption_limits: Limits,
        opc_limits: ReadLimits,
    ) -> Result<Self> {
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(bytes.len())
            .map_err(|source| Error::Allocation {
                resource: "encrypted package input",
                source,
            })?;
        owned.extend_from_slice(bytes);
        Self::from_vec_with_password_and_limits(owned, password, encryption_limits, opc_limits)
    }

    /// Explicitly serialize a clear OPC package, permitting an encryption downgrade.
    pub fn to_plain_bytes(&mut self) -> Result<Vec<u8>> {
        self.flush_presentation()?;
        let bytes = PackageWriter::to_bytes(&self.opc)?;
        Ok(bytes)
    }

    /// Explicitly save a clear OPC package, permitting an encryption downgrade.
    pub fn save_plain<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        self.flush_presentation()?;
        PackageWriter::write(path, &self.opc)?;
        Ok(())
    }

    /// Serialize and encrypt the package in the selected mode.
    pub fn to_encrypted(&mut self, password: &str, mode: Mode) -> Result<Vec<u8>> {
        self.to_encrypted_with_limits(password, mode, Limits::default())
    }

    /// Serialize and encrypt the package in the selected mode under explicit limits.
    pub fn to_encrypted_with_limits(
        &mut self,
        password: &str,
        mode: Mode,
        limits: Limits,
    ) -> Result<Vec<u8>> {
        self.encrypt_inner(password, mode, &limits)
    }

    /// Compatibility alias for [`Self::to_encrypted`].
    pub fn to_encrypted_bytes(&mut self, password: &str, mode: Mode) -> Result<Vec<u8>> {
        self.to_encrypted(password, mode)
    }

    /// Serialize and encrypt using the retained input/output mode.
    pub fn to_reencrypted(&mut self, password: &str) -> Result<Vec<u8>> {
        self.to_reencrypted_with_limits(password, Limits::default())
    }

    /// Serialize and encrypt using the retained mode under explicit limits.
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

    /// Compatibility alias for [`Self::to_reencrypted`].
    pub fn to_reencrypted_bytes(&mut self, password: &str) -> Result<Vec<u8>> {
        self.to_reencrypted(password)
    }

    /// Atomically save the package encrypted in the selected mode.
    pub fn save_encrypted<P: AsRef<Path>>(
        &mut self,
        path: P,
        password: &str,
        mode: Mode,
    ) -> Result<()> {
        self.save_encrypted_with_limits(path, password, mode, Limits::default())
    }

    /// Atomically save selected-mode encryption under explicit crypto limits.
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
    pub fn save_reencrypted<P: AsRef<Path>>(&mut self, path: P, password: &str) -> Result<()> {
        self.save_reencrypted_with_limits(path, password, Limits::default())
    }

    /// Atomically save retained-mode encryption under explicit crypto limits.
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

    fn from_opened_encryption(opened: ooxml::Opened, limits: ReadLimits) -> Result<Self> {
        let mode = opened.mode();
        let mut package = Self::from_vec_with_limits(opened.into_bytes(), limits)?;
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
