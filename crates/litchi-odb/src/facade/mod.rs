//! Concise family entry points.

use litchi_core::{Error as CoreError, Metadata, Result as CoreResult};
use std::{
    path::Path,
    sync::{Arc, OnceLock},
};

pub use crate::authoring::Builder;

type FamilyResult<T> = Result<T, litchi_odf_common::Error>;

/// The immutable, cheap-to-clone database package owner.
///
/// It owns every state verb (open, edit, apply, save); its read-only semantic
/// projection is [`Database`], obtained from [`Package::database`].
#[derive(Clone)]
pub struct Package {
    pub(crate) database: Database,
}

impl std::fmt::Debug for Package {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Package")
            .field("bytes", &self.as_bytes().len())
            .finish_non_exhaustive()
    }
}

impl Package {
    /// The retained physical package (metadata transactions, record 0288).
    pub(crate) fn owned_package(&self) -> &litchi_odf_common::OwnedPackage {
        self.database.package.owned_package()
    }

    /// Creates the minimal valid database package (an empty data source).
    ///
    /// Creation is deterministic: no clock or randomness is consulted.
    ///
    /// # Errors
    ///
    /// Returns an error if package writing fails.
    pub fn new() -> Result<Self, litchi_odf_common::Error> {
        Builder::new().build()
    }

    /// Opens a database package from a file path under the default limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the file cannot be read or is not a valid package.
    pub fn open(path: impl AsRef<Path>) -> Result<Self, litchi_odf_common::Error> {
        Self::open_with_options(path, litchi_odf_common::OpenOptions::new())
    }

    /// Opens a database package from a file path under caller-selected
    /// source and semantic limits.
    ///
    /// # Errors
    ///
    /// Returns [`litchi_odf_common::Error::Limit`] when a selected limit is
    /// exceeded, and another error for an unreadable or invalid package.
    pub fn open_with_options(
        path: impl AsRef<Path>,
        options: litchi_odf_common::OpenOptions,
    ) -> Result<Self, litchi_odf_common::Error> {
        crate::package::Snapshot::open_with_options(path, options).map(Self::from_snapshot)
    }

    /// Opens a database package from in-memory bytes (copied once).
    ///
    /// # Errors
    ///
    /// Returns an error if the bytes are not a valid package.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, litchi_odf_common::Error> {
        Self::from_bytes_with_options(bytes, litchi_odf_common::OpenOptions::new())
    }

    /// Opens a database package from in-memory bytes under caller-selected
    /// limits.
    ///
    /// # Errors
    ///
    /// As for [`Self::open_with_options`].
    pub fn from_bytes_with_options(
        bytes: &[u8],
        options: litchi_odf_common::OpenOptions,
    ) -> Result<Self, litchi_odf_common::Error> {
        crate::package::Snapshot::from_bytes_with_options(copy_input(bytes)?, options)
            .map(Self::from_snapshot)
    }

    /// Reads a database package from a stream under the default limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the stream cannot be read or is not a valid package.
    pub fn from_reader(reader: impl std::io::Read) -> Result<Self, litchi_odf_common::Error> {
        Self::from_reader_with_options(reader, litchi_odf_common::OpenOptions::new())
    }

    /// Reads a database package from a stream under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As for [`Self::open_with_options`].
    pub fn from_reader_with_options(
        reader: impl std::io::Read,
        options: litchi_odf_common::OpenOptions,
    ) -> Result<Self, litchi_odf_common::Error> {
        crate::package::Snapshot::from_reader_with_options(reader, options).map(Self::from_snapshot)
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
    pub fn open_with_password(
        path: impl AsRef<Path>,
        password: impl Into<String>,
    ) -> Result<Self, litchi_odf_common::Error> {
        Self::open_with_password_and_options(path, password, litchi_odf_common::OpenOptions::new())
    }

    /// Opens a password-encrypted database package from a file path under
    /// caller-selected limits.
    ///
    /// # Errors
    ///
    /// As for [`Self::open_with_password`] and [`Self::open_with_options`].
    pub fn open_with_password_and_options(
        path: impl AsRef<Path>,
        password: impl Into<String>,
        options: litchi_odf_common::OpenOptions,
    ) -> Result<Self, litchi_odf_common::Error> {
        crate::package::Snapshot::open_with_password_and_options(path, password, options)
            .map(Self::from_snapshot)
    }

    /// Opens a password-encrypted database package from in-memory bytes.
    ///
    /// # Errors
    ///
    /// Returns an error for an incorrect password or invalid decrypted package.
    pub fn from_bytes_with_password(
        bytes: &[u8],
        password: impl Into<String>,
    ) -> Result<Self, litchi_odf_common::Error> {
        Self::from_bytes_with_password_and_options(
            bytes,
            password,
            litchi_odf_common::OpenOptions::new(),
        )
    }

    /// Opens a password-encrypted database package from in-memory bytes
    /// under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As for [`Self::from_bytes_with_password`] and
    /// [`Self::open_with_options`].
    pub fn from_bytes_with_password_and_options(
        bytes: &[u8],
        password: impl Into<String>,
        options: litchi_odf_common::OpenOptions,
    ) -> Result<Self, litchi_odf_common::Error> {
        crate::package::Snapshot::from_bytes_with_password_and_options(
            copy_input(bytes)?,
            password,
            options,
        )
        .map(Self::from_snapshot)
    }

    /// Adopts the packaged ODF detection result without a second ZIP index.
    ///
    /// # Errors
    ///
    /// Returns an error when detection classified another format or the
    /// package is not a valid database.
    pub fn from_prepared_package(
        prepared: litchi_odf_common::PreparedPackage,
    ) -> Result<Self, litchi_odf_common::Error> {
        Self::from_prepared_package_with_options(prepared, litchi_odf_common::OpenOptions::new())
    }

    /// Adopt a prepared package under the semantic limits in `options`.
    /// The source limits were applied when the archive was prepared; its
    /// existing index and source allocation are reused.
    ///
    /// # Errors
    ///
    /// As [`Self::from_prepared_package`], plus a typed semantic limit failure.
    pub fn from_prepared_package_with_options(
        prepared: litchi_odf_common::PreparedPackage,
        options: litchi_odf_common::OpenOptions,
    ) -> Result<Self, litchi_odf_common::Error> {
        if prepared.format() != litchi_core::detection::FileFormat::Odb {
            return Err(CoreError::UnexpectedFormat {
                detected: prepared.format(),
            }
            .into());
        }
        crate::package::Snapshot::from_owned_package_with_options(prepared.into_package(), options)
            .map(Self::from_snapshot)
    }

    pub(crate) fn from_snapshot(package: crate::package::Snapshot) -> Self {
        Self {
            database: Database {
                package,
                settings: Arc::new(OnceLock::new()),
            },
        }
    }

    /// Reopens bytes this crate published (default limits).
    pub(crate) fn from_published(bytes: Vec<u8>) -> CoreResult<Self> {
        crate::package::Snapshot::from_bytes(bytes).map(Self::from_snapshot)
    }

    /// Returns the tolerated deviations recorded while opening the package.
    #[must_use]
    pub fn tolerance_report(&self) -> &litchi_odf_common::leniency::ToleranceReport {
        self.database.package.tolerance_report()
    }

    /// Borrows the read-only semantic projection; clone it for an owned,
    /// cheap handle.
    #[must_use]
    pub const fn database(&self) -> &Database {
        &self.database
    }

    /// Returns the exact package bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.database.package.as_bytes()
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

    /// Replays a committed patch onto this exact source package.
    ///
    /// # Errors
    ///
    /// Returns an error when this package is not the patch's source.
    pub fn apply(&self, patch: &crate::Patch) -> Result<crate::Commit, litchi_odf_common::Error> {
        if !patch.is_applicable_to(self) {
            return Err(litchi_odf_common::Error::patch_conflict(
                "odb.apply",
                "the package is not the patch source",
            ));
        }
        Ok(crate::Commit::replayed(patch.apply(self)?, patch.clone()))
    }

    /// Atomically saves the package (sibling temporary file, then rename,
    /// with file and directory synchronization).
    ///
    /// # Errors
    ///
    /// Returns an error when the staging file cannot be written or renamed;
    /// the destination is then unchanged.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<(), litchi_odf_common::Error> {
        self.save_with_durability(path, litchi_core::Durability::Full)
    }

    /// Atomically saves the package with an explicit durability level.
    ///
    /// # Errors
    ///
    /// As for [`Self::save`].
    pub fn save_with_durability(
        &self,
        path: impl AsRef<Path>,
        durability: litchi_core::Durability,
    ) -> Result<(), litchi_odf_common::Error> {
        litchi_odf_common::publish::save_bytes(path.as_ref(), self.as_bytes(), durability)
    }

    /// Returns a copy of the package bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the copy cannot be allocated.
    pub fn to_bytes(&self) -> Result<Vec<u8>, litchi_odf_common::Error> {
        Ok(copy_input(self.as_bytes())?)
    }

    /// Writes the package bytes to `sink`.
    ///
    /// # Errors
    ///
    /// Returns an error when the sink fails.
    pub fn write_to(&self, mut sink: impl std::io::Write) -> Result<(), litchi_odf_common::Error> {
        sink.write_all(self.as_bytes())
            .map_err(litchi_odf_common::Error::from)?;
        Ok(())
    }
}

fn copy_input(bytes: &[u8]) -> CoreResult<Vec<u8>> {
    let mut copy = Vec::new();
    copy.try_reserve_exact(bytes.len())
        .map_err(|source| CoreError::Allocation {
            resource: "ODB package bytes",
            source,
        })?;
    copy.extend_from_slice(bytes);
    Ok(copy)
}

/// The read-only semantic projection of a database [`Package`].
///
/// Cheap to clone; it carries no state verbs.
#[derive(Clone)]
pub struct Database {
    pub(crate) package: crate::package::Snapshot,
    pub(crate) settings: Arc<OnceLock<crate::DatabaseSettings>>,
}

impl std::fmt::Debug for Database {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Database").finish_non_exhaustive()
    }
}

impl Database {
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
    pub fn files(&self) -> FamilyResult<Vec<String>> {
        Ok(self.package.files()?)
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
    pub fn catalog(&self) -> FamilyResult<crate::Catalog<'_>> {
        self.catalog_with(crate::Limits::default())
    }

    /// Reads an inert schema and query catalog with caller-selected finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the content part is structurally invalid or
    /// exceeds the mandatory catalog limits. Optional settings limits are
    /// checked when [`crate::Catalog::settings`] is first requested.
    pub fn catalog_with(&self, limits: crate::Limits) -> FamilyResult<crate::Catalog<'_>> {
        let settings_cache = if limits == crate::Limits::default() {
            self.settings.clone()
        } else {
            Arc::new(OnceLock::new())
        };
        Ok(crate::Catalog::parse_with_settings(
            self.content_xml(),
            limits,
            settings_cache,
        )?)
    }

    /// Reads bounded inert login, driver, application, and table-filter
    /// settings from the data source.  The returned values are metadata only;
    /// no driver or credential is resolved. A successful projection is cached
    /// on this immutable snapshot and returned by reference on later reads;
    /// failed projections are not cached. Sources that omit the normative
    /// connection-data owner, its single connection target, or a required
    /// settings child remain readable as preservation-only snapshots, but do
    /// not produce a typed projection. Unknown producer markup is retained;
    /// it never substitutes for a normative `oneOrMore` child.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the optional settings XML is malformed or
    /// exceeds its finite limits.
    pub fn settings(&self) -> FamilyResult<&crate::DatabaseSettings> {
        if let Some(settings) = self.settings.get() {
            return Ok(settings);
        }
        let settings = self.catalog()?.settings()?.clone();
        let _ = self.settings.set(settings);
        self.settings.get().ok_or_else(|| {
            litchi_odf_common::Error::from(CoreError::InvalidFormat(
                "ODB settings projection was not materialized".into(),
            ))
        })
    }

    /// Inventories opaque direct producer-extension subtrees without
    /// interpreting or activating them.
    ///
    /// # Errors
    ///
    /// Returns an error when the content XML is malformed or an extension
    /// range cannot be represented safely.
    pub fn producer_extensions(&self) -> FamilyResult<Vec<crate::ProducerExtension>> {
        Ok(crate::authoring::producer_extensions(self.content_xml())?)
    }

    /// Inventories macros, scripts, event listeners, form controls, actions,
    /// DDE declarations, and embedded objects without activating any member.
    ///
    /// # Errors
    ///
    /// Returns an error when a candidate XML member is malformed or the
    /// bounded scan limits are exceeded.
    pub fn active_content(&self) -> FamilyResult<crate::ActiveContentInventory> {
        Ok(self.package.active_content()?)
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
    ) -> FamilyResult<crate::ActiveContentInventory> {
        let catalog = self.catalog()?;
        let mut matches = catalog
            .components()
            .iter()
            .filter(|component| component.kind() == kind && component.name() == Some(name));
        let component = matches.next().ok_or_else(|| {
            litchi_odf_common::Error::from(CoreError::InvalidFormat(
                "ODB active-content component does not exist".to_string(),
            ))
        })?;
        if matches.next().is_some() {
            return Err(litchi_odf_common::Error::from(CoreError::InvalidFormat(
                "ODB active-content component selector is ambiguous".to_string(),
            )));
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
    ) -> FamilyResult<crate::ComponentDependencyInventory> {
        let catalog = self.catalog()?;
        let mut matches = catalog
            .components()
            .iter()
            .filter(|component| component.kind() == kind && component.name() == Some(name));
        let component = matches.next().ok_or_else(|| {
            litchi_odf_common::Error::from(CoreError::InvalidFormat(
                "ODB component dependency selector does not exist".to_string(),
            ))
        })?;
        if matches.next().is_some() {
            return Err(litchi_odf_common::Error::from(CoreError::InvalidFormat(
                "ODB component dependency selector is ambiguous".to_string(),
            )));
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
            CoreError::InvalidFormat("ODB component dependency count overflow".to_string())
        })?;
        let mut entries = Vec::new();
        entries
            .try_reserve(capacity)
            .map_err(|source| CoreError::Allocation {
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
    pub fn protection_status(&self) -> FamilyResult<crate::ProtectionStatus> {
        Ok(self.package.protection_status()?)
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
    /// Returns [`litchi_odf_common::Error::Refused`] for re-signing and
    /// re-encryption. Package bytes and credentials remain untouched.
    pub fn require_protection_operation(
        &self,
        operation: crate::ProtectionOperation,
    ) -> FamilyResult<()> {
        match self.protection_decision(operation) {
            crate::ProtectionDecision::Supported => Ok(()),
            crate::ProtectionDecision::Refused(_reason) => Err(litchi_odf_common::Error::from(
                CoreError::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::Protected,
                    "odb.facade",
                    "ODB protected-package operation is refused",
                )),
            )),
        }
    }

    /// Reads inert document and macro signature metadata without verifying it
    /// or activating macro content.
    ///
    /// # Errors
    ///
    /// Returns an error if signature XML cannot be decoded safely.
    pub fn digital_signatures(&self) -> FamilyResult<crate::DigitalSignatures> {
        Ok(self.package.digital_signatures()?)
    }

    /// Verifies package/document signature math without making a certificate
    /// trust decision and without executing database or macro content.
    ///
    /// # Errors
    ///
    /// Returns an error if signature metadata or referenced bytes cannot be
    /// decoded or verified.
    pub fn verify_document_signatures(&self) -> FamilyResult<Vec<crate::SignatureVerification>> {
        Ok(self.package.verify_document_signatures()?)
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "test code panics on failure; unwrap keeps assertions concise"
)]
mod tests {
    use super::Builder;
    use crate::model::{parse_count, parse_test_lock, reset_parse_count};

    #[test]
    fn builder_opens_as_validated_snapshot() {
        let package = Builder::new().build().unwrap();
        let document = package.database();
        assert!(document.content_xml().contains("<office:database"));
        assert!(!document.as_bytes().is_empty());
        assert!(document.catalog().unwrap().tables().is_empty());
    }

    #[test]
    fn settings_reads_materialize_lazily_and_reuse_success() {
        let _lock = parse_test_lock().lock().unwrap();
        reset_parse_count();
        let package = Builder::new().build().unwrap();
        let document = package.database();
        assert_eq!(parse_count(), 0);
        let catalog = document.catalog().unwrap();
        assert!(catalog.tables().is_empty());
        assert_eq!(parse_count(), 0);

        let first = catalog.settings().unwrap();
        assert_eq!(parse_count(), 1);
        let second = document.settings().unwrap();
        assert!(std::ptr::eq(first, second));
        assert_eq!(parse_count(), 1);
    }

    #[test]
    fn custom_catalog_limits_do_not_reuse_default_settings_cache() {
        let _lock = parse_test_lock().lock().unwrap();
        reset_parse_count();
        let package = Builder::new()
                .content_xml(
                    r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:db="urn:oasis:names:tc:opendocument:xmlns:database:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:database><db:data-source><db:connection-data><db:connection-resource xlink:href="" xlink:type="simple"/><db:login db:user-name="alice"/></db:connection-data></db:data-source></office:database></office:body></office:document-content>"#,
                )
                .build()
                .unwrap();
        let document = package.database();
        document.settings().unwrap();
        assert_eq!(parse_count(), 1);

        let catalog = document
            .catalog_with(crate::Limits::default().with_max_attribute_bytes(0))
            .unwrap();
        assert!(catalog.settings().is_err());
        assert_eq!(parse_count(), 2);
        assert!(document.settings().is_ok());
        assert_eq!(parse_count(), 2);
    }
}
