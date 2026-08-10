//! Concise family entry points.

use litchi_core::{Error, Metadata, Result};
use std::path::Path;

pub use crate::authoring::Builder;

/// Immutable document snapshot.
#[derive(Clone)]
pub struct Database {
    pub(crate) package: crate::package::Snapshot,
}

impl Database {
    /// Opens a database package from a file path.
    ///
    /// # Errors
    ///
    /// Returns an error if the file cannot be read or is not a valid package.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        crate::package::Snapshot::open(path).map(|package| Self { package })
    }

    /// Opens a database package from in-memory bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the bytes are not a valid package.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        crate::package::Snapshot::from_bytes(bytes).map(|package| Self { package })
    }

    /// Opens a password-encrypted database package from a file path.
    ///
    /// The password is used only to decode package entries. No database
    /// credential, driver, connection, query, macro, form, or report executes.
    ///
    /// # Errors
    ///
    /// Returns an error if the file cannot be read, the password is incorrect,
    /// or the decrypted package is invalid.
    pub fn open_with_password(path: impl AsRef<Path>, password: impl Into<String>) -> Result<Self> {
        Self::from_bytes_with_password(std::fs::read(path)?, password)
    }

    /// Opens a password-encrypted database package from in-memory bytes.
    ///
    /// # Errors
    ///
    /// Returns an error for an incorrect password or invalid decrypted package.
    pub fn from_bytes_with_password(bytes: Vec<u8>, password: impl Into<String>) -> Result<Self> {
        crate::package::Snapshot::from_bytes_with_password(bytes, password)
            .map(|package| Self { package })
    }

    /// Returns the `content.xml` document.
    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.package.content_xml()
    }

    /// Returns the `styles.xml` document, if present.
    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.package.styles_xml()
    }

    /// Returns the document metadata, if present.
    #[must_use]
    pub fn metadata(&self) -> Option<&Metadata> {
        self.package.metadata()
    }

    /// Returns the raw package bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    /// Lists the file entries stored in the package.
    ///
    /// # Errors
    ///
    /// Returns an error if the package entries cannot be enumerated.
    pub fn files(&self) -> Result<Vec<String>> {
        self.package.files()
    }

    /// Reads a bounded, source-bound inert schema and query catalog.
    ///
    /// This never opens a connection, resolves a driver, or executes a query.
    /// Unknown source markup remains preserved in the immutable package bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the content part is structurally invalid or
    /// exceeds the default semantic discovery limits.
    pub fn catalog(&self) -> Result<crate::Catalog<'_>> {
        self.catalog_with(crate::Limits::default())
    }

    /// Reads an inert schema and query catalog with caller-selected finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the content part is structurally invalid or
    /// exceeds `limits`.
    pub fn catalog_with(&self, limits: crate::Limits) -> Result<crate::Catalog<'_>> {
        crate::Catalog::parse(self.content_xml(), limits)
    }

    /// Inventories opaque direct producer-extension subtrees without
    /// interpreting or activating them.
    ///
    /// # Errors
    ///
    /// Returns an error when the content XML is malformed or an extension
    /// range cannot be represented safely.
    pub fn producer_extensions(&self) -> Result<Vec<crate::ProducerExtension>> {
        crate::authoring::producer_extensions(self.content_xml())
    }

    /// Inventories macros, scripts, event listeners, form controls, actions,
    /// DDE declarations, and embedded objects without activating any member.
    ///
    /// # Errors
    ///
    /// Returns an error when a candidate XML member is malformed or the
    /// bounded scan limits are exceeded.
    pub fn active_content(&self) -> Result<crate::ActiveContentInventory> {
        self.package.active_content()
    }

    /// Inventories active-content dependencies inside one local form/report
    /// package subtree without following external IRIs or activating content.
    ///
    /// # Errors
    ///
    /// Returns an error when the component selector is absent or ambiguous,
    /// or when the bounded active-content inventory cannot be decoded.
    pub fn component_active_content(
        &self,
        kind: crate::ComponentKind,
        name: &str,
    ) -> Result<crate::ActiveContentInventory> {
        let catalog = self.catalog()?;
        let mut matches = catalog
            .components()
            .iter()
            .filter(|component| component.kind() == kind && component.name() == Some(name));
        let component = matches.next().ok_or_else(|| {
            Error::InvalidFormat("ODB active-content component does not exist".to_string())
        })?;
        if matches.next().is_some() {
            return Err(Error::InvalidFormat(
                "ODB active-content component selector is ambiguous".to_string(),
            ));
        }
        if component.link_kind() != crate::ComponentLinkKind::LocalPackage {
            return Ok(crate::ActiveContentInventory::new(Vec::new()));
        }
        let prefix = format!("{}/", component.href().unwrap_or("").trim_end_matches('/'));
        let entries = self
            .active_content()?
            .entries()
            .iter()
            .filter(|entry| entry.package_path().starts_with(&prefix))
            .cloned()
            .collect();
        Ok(crate::ActiveContentInventory::new(entries))
    }

    /// Inventories the exact bounded package closure of one linked component.
    ///
    /// Payload files/directories are distinguished from shared members reached
    /// through package-local XML `xlink:href` chains. External IRIs remain
    /// inert and produce an empty inventory. No component, database command,
    /// embedded database, macro, or connection is opened or executed.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent/ambiguous component, malformed or unsafe
    /// package link, missing dependency, encrypted unreadable member, or
    /// dependency budget violation.
    pub fn component_dependencies(
        &self,
        kind: crate::ComponentKind,
        name: &str,
    ) -> Result<crate::ComponentDependencyInventory> {
        let catalog = self.catalog()?;
        let mut matches = catalog
            .components()
            .iter()
            .filter(|component| component.kind() == kind && component.name() == Some(name));
        let component = matches.next().ok_or_else(|| {
            Error::InvalidFormat("ODB component dependency selector does not exist".to_string())
        })?;
        if matches.next().is_some() {
            return Err(Error::InvalidFormat(
                "ODB component dependency selector is ambiguous".to_string(),
            ));
        }
        if component.link_kind() != crate::ComponentLinkKind::LocalPackage {
            return Ok(crate::ComponentDependencyInventory::new(
                Vec::new(),
                0,
                crate::ComponentTransferSupport::Supported,
            ));
        }
        let prefix = format!("{}/", component.href().unwrap_or("").trim_end_matches('/'));
        let (files, directories) = self.package.component_payload(&prefix, &prefix)?;
        let active_content_count = crate::package::Snapshot::payload_active_content_count(&files)?;
        let transfer_support = crate::package::Snapshot::payload_transfer_support(&files);
        let capacity = files.len().checked_add(directories.len()).ok_or_else(|| {
            Error::InvalidFormat("ODB component dependency count overflow".to_string())
        })?;
        let mut entries = Vec::new();
        entries
            .try_reserve(capacity)
            .map_err(|source| Error::Allocation {
                resource: "ODB component dependency inventory",
                source,
            })?;
        entries.extend(files.into_iter().map(|file| {
            let kind = if file.path.starts_with(&prefix) {
                crate::ComponentDependencyKind::PayloadFile
            } else {
                crate::ComponentDependencyKind::LinkedFile
            };
            let byte_len = file.bytes.len();
            crate::ComponentDependency::new(kind, file.path, file.media_type, Some(byte_len))
        }));
        entries.extend(directories.into_iter().map(|(path, media_type)| {
            let kind = if path.starts_with(&prefix) {
                crate::ComponentDependencyKind::PayloadDirectory
            } else {
                crate::ComponentDependencyKind::LinkedDirectory
            };
            crate::ComponentDependency::new(kind, path, media_type, None)
        }));
        Ok(crate::ComponentDependencyInventory::new(
            entries,
            active_content_count,
            transfer_support,
        ))
    }

    /// Inventories signatures and manifest-declared encryption without
    /// validating, decrypting, or executing protected content.
    ///
    /// # Errors
    ///
    /// Returns an error if the package entries cannot be inspected.
    pub fn protection_status(&self) -> Result<crate::ProtectionStatus> {
        self.package.protection_status()
    }

    /// Returns the stable protected-publication capability contract.
    ///
    /// ODB signature verification and explicit invalid-signature removal are
    /// supported. Re-signing and re-encryption are explicitly unsupported;
    /// changed encrypted publication remains fail-closed.
    #[must_use]
    pub const fn protection_capabilities(&self) -> crate::ProtectionCapabilities {
        crate::ProtectionCapabilities::odb()
    }

    /// Returns explicit support for one protected-package operation.
    #[must_use]
    pub const fn protection_support(
        &self,
        operation: crate::ProtectionOperation,
    ) -> crate::ProtectionSupport {
        self.protection_capabilities().support(operation)
    }

    /// Returns the exact typed support/refusal decision for one protected
    /// package operation.
    #[must_use]
    pub const fn protection_decision(
        &self,
        operation: crate::ProtectionOperation,
    ) -> crate::ProtectionDecision {
        self.protection_capabilities().decision(operation)
    }

    /// Enforces the root protection boundary before a workflow gathers
    /// credentials or signing material.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Unsupported`] for re-signing and re-encryption. No
    /// package bytes are changed and no credential is requested.
    pub fn require_protection_operation(
        &self,
        operation: crate::ProtectionOperation,
    ) -> Result<()> {
        match self.protection_decision(operation) {
            crate::ProtectionDecision::Supported => Ok(()),
            crate::ProtectionDecision::Refused(reason) => Err(Error::Unsupported(format!(
                "ODB protected-package operation is refused: {reason:?}"
            ))),
        }
    }

    /// Reads inert document and macro signature metadata without verifying it
    /// or activating macro content.
    ///
    /// # Errors
    ///
    /// Returns an error if signature XML cannot be decoded safely.
    pub fn digital_signatures(&self) -> Result<crate::DigitalSignatures> {
        self.package.digital_signatures()
    }

    /// Verifies package/document signature math without making a certificate
    /// trust decision and without executing database or macro content.
    ///
    /// # Errors
    ///
    /// Returns an error if signature metadata or referenced bytes cannot be
    /// decoded or verified.
    pub fn verify_document_signatures(&self) -> Result<Vec<crate::SignatureVerification>> {
        self.package.verify_document_signatures()
    }

    /// Starts a source-bound unified database-front-end transaction.
    ///
    /// Connection targets, query commands, component links, and producer
    /// extensions remain inert data and are never opened or executed.
    #[must_use]
    pub fn edit(&self) -> crate::Edit<'_> {
        crate::Edit::new(self, crate::EditPolicy::default())
    }

    /// Starts a unified transaction with explicit protected-package policy.
    #[must_use]
    pub fn edit_with_policy(&self, policy: crate::EditPolicy) -> crate::Edit<'_> {
        crate::Edit::new(self, policy)
    }

    /// Consumes the snapshot and returns the raw package bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "test code panics on failure; unwrap keeps assertions concise"
)]
mod tests {
    use super::{Builder, Database};

    #[test]
    fn builder_opens_as_validated_snapshot() {
        let bytes = Builder::new().build().unwrap();
        let document = Database::from_bytes(bytes).unwrap();
        assert!(document.content_xml().contains("<office:database"));
        assert!(!document.as_bytes().is_empty());
        assert!(document.catalog().unwrap().tables().is_empty());
    }
}
