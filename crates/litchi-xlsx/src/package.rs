//! Validated XLSX package ownership.
//!
//! [`Package`] is the physical-package boundary for the standalone XLSX
//! crate. It owns the OPC graph, delegates archive I/O to `litchi-opc`, and
//! validates the `SpreadsheetML` workbook graph before exposing a package
//! handle. Semantic reads and transactional edits live in [`crate::workbook`]
//! and [`crate::edit`].

/// Inert worksheet Printer Settings parts and references.
pub mod printer_settings;
/// Deterministic minimal package resources.
pub mod template;

use std::io::{Read, Write};
use std::path::Path;

use litchi_ooxml_common::custom::Host;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, ReadLimits, TargetMode};

use crate::Workbook;
use crate::custom::Props;
use crate::error::{Result, invalid};
use crate::writer;

#[cfg(feature = "encryption")]
use litchi_ooxml_common::package_encryption::PackageEncryption;

/// A validated physical XLSX package.
///
/// The ZIP/container implementation remains below `litchi-opc`; this type
/// exposes only XLSX package operations and does not leak archive-specific
/// readers, writers, or errors.
#[derive(Debug, Clone)]
pub struct Package(OpcPackage, #[cfg(feature = "encryption")] PackageEncryption);

impl Package {
    /// Create a deterministic minimal XLSX package with one visible worksheet.
    pub fn create() -> Result<Self> {
        Self::from_opc(build_minimal_package()?)
    }

    /// Open and validate an XLSX package from a filesystem path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_limits(path, ReadLimits::default())
    }

    /// Open and validate an XLSX package from a filesystem path with explicit
    /// OPC resource limits.
    pub fn open_with_limits(path: impl AsRef<Path>, limits: ReadLimits) -> Result<Self> {
        Self::from_opc(OpcPackage::open_with_limits(path, limits)?)
    }

    /// Open an ordinary or encrypted XLSX package with safe independent
    /// encryption and OPC limits.
    #[cfg(feature = "encryption")]
    pub fn open_with_password(path: impl AsRef<Path>, password: &str) -> Result<Self> {
        Self::open_with_password_and_limits(
            path,
            password,
            &crate::encryption::Limits::default(),
            ReadLimits::default(),
        )
    }

    /// Open an ordinary or encrypted XLSX package with independent outer
    /// encryption and inner OPC resource policies.
    #[cfg(feature = "encryption")]
    pub fn open_with_password_and_limits(
        path: impl AsRef<Path>,
        password: &str,
        encryption_limits: &crate::encryption::Limits,
        opc_limits: ReadLimits,
    ) -> Result<Self> {
        let file = std::fs::File::open(path).map_err(crate::encryption::Error::Io)?;
        let opened = crate::encryption::load_with(file, password, encryption_limits)?;
        Self::from_opened_with_limits(opened, opc_limits)
    }

    /// Read and validate an XLSX package from owned bytes.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, ReadLimits::default())
    }

    /// Read and validate an XLSX package from owned bytes with explicit OPC
    /// resource limits.
    pub fn from_bytes_with_limits(bytes: Vec<u8>, limits: ReadLimits) -> Result<Self> {
        Self::from_opc(OpcPackage::from_vec_with_limits(bytes, limits)?)
    }

    /// Move ordinary or encrypted bytes into the XLSX parser using safe
    /// independent resource policies.
    #[cfg(feature = "encryption")]
    pub fn from_bytes_with_password(bytes: Vec<u8>, password: &str) -> Result<Self> {
        Self::from_bytes_with_password_and_limits(
            bytes,
            password,
            &crate::encryption::Limits::default(),
            ReadLimits::default(),
        )
    }

    /// Move ordinary or encrypted bytes into the XLSX parser with independent
    /// outer encryption and inner OPC resource policies.
    #[cfg(feature = "encryption")]
    pub fn from_bytes_with_password_and_limits(
        bytes: Vec<u8>,
        password: &str,
        encryption_limits: &crate::encryption::Limits,
        opc_limits: ReadLimits,
    ) -> Result<Self> {
        let opened = crate::encryption::open_with(bytes, password, encryption_limits)?;
        Self::from_opened_with_limits(opened, opc_limits)
    }

    /// Read and validate an XLSX package from a borrowed byte slice.
    pub fn from_slice(bytes: &[u8]) -> Result<Self> {
        Self::from_slice_with_limits(bytes, ReadLimits::default())
    }

    /// Read and validate an XLSX package from a borrowed byte slice with
    /// explicit OPC resource limits.
    pub fn from_slice_with_limits(bytes: &[u8], limits: ReadLimits) -> Result<Self> {
        Self::from_opc(OpcPackage::from_bytes_with_limits(bytes, limits)?)
    }

    /// Read an ordinary or encrypted borrowed byte slice with safe independent
    /// resource policies.
    #[cfg(feature = "encryption")]
    pub fn from_slice_with_password(bytes: &[u8], password: &str) -> Result<Self> {
        Self::from_slice_with_password_and_limits(
            bytes,
            password,
            &crate::encryption::Limits::default(),
            ReadLimits::default(),
        )
    }

    /// Read an ordinary or encrypted borrowed byte slice with independent
    /// outer encryption and inner OPC resource policies.
    #[cfg(feature = "encryption")]
    pub fn from_slice_with_password_and_limits(
        bytes: &[u8],
        password: &str,
        encryption_limits: &crate::encryption::Limits,
        opc_limits: ReadLimits,
    ) -> Result<Self> {
        let opened = crate::encryption::load_with(bytes, password, encryption_limits)?;
        Self::from_opened_with_limits(opened, opc_limits)
    }

    /// Read and validate an XLSX package from a synchronous reader.
    pub fn from_reader(reader: impl Read) -> Result<Self> {
        Self::from_reader_with_limits(reader, ReadLimits::default())
    }

    /// Read and validate an XLSX package from a synchronous reader with
    /// explicit OPC resource limits.
    pub fn from_reader_with_limits(reader: impl Read, limits: ReadLimits) -> Result<Self> {
        Self::from_opc(OpcPackage::from_reader_with_limits(reader, limits)?)
    }

    /// Read an ordinary or encrypted package from a synchronous reader using
    /// safe independent resource policies.
    #[cfg(feature = "encryption")]
    pub fn from_reader_with_password(reader: impl Read, password: &str) -> Result<Self> {
        Self::from_reader_with_password_and_limits(
            reader,
            password,
            &crate::encryption::Limits::default(),
            ReadLimits::default(),
        )
    }

    /// Read an ordinary or encrypted package from a synchronous reader with
    /// independent outer encryption and inner OPC resource policies.
    #[cfg(feature = "encryption")]
    pub fn from_reader_with_password_and_limits(
        reader: impl Read,
        password: &str,
        encryption_limits: &crate::encryption::Limits,
        opc_limits: ReadLimits,
    ) -> Result<Self> {
        let opened = crate::encryption::load_with(reader, password, encryption_limits)?;
        Self::from_opened_with_limits(opened, opc_limits)
    }

    /// Validate and adopt an already parsed OPC package.
    ///
    /// This raw compatibility boundary deliberately classifies the adopted
    /// clear OPC graph as plaintext. Callers converting decrypted content to
    /// `OpcPackage` therefore explicitly declassify its encryption provenance.
    pub fn from_opc(package: OpcPackage) -> Result<Self> {
        // Run the complete workbook/relationship validation once at the
        // package boundary. The cloned graph shares immutable OPC payloads.
        Workbook::from_package(package.clone())?;
        Ok(Self(
            package,
            #[cfg(feature = "encryption")]
            PackageEncryption::plain(),
        ))
    }

    #[cfg(feature = "encryption")]
    fn from_opened_with_limits(
        opened: crate::encryption::Opened,
        opc_limits: ReadLimits,
    ) -> Result<Self> {
        let provenance = opened
            .mode()
            .map_or_else(PackageEncryption::plain, PackageEncryption::encrypted);
        let package = OpcPackage::from_vec_with_limits(opened.into_bytes(), opc_limits)?;
        Workbook::from_package_with_encryption(package.clone(), provenance)?;
        Ok(Self(package, provenance))
    }

    /// Materialize an immutable workbook snapshot from this package.
    pub fn workbook(&self) -> Result<Workbook> {
        #[cfg(feature = "encryption")]
        return Workbook::from_package_with_encryption(self.0.clone(), self.1);
        #[cfg(not(feature = "encryption"))]
        Workbook::from_package(self.0.clone())
    }

    /// Consume this package and materialize its immutable workbook snapshot.
    pub fn into_workbook(self) -> Result<Workbook> {
        #[cfg(feature = "encryption")]
        return Workbook::from_package_with_encryption(self.0, self.1);
        #[cfg(not(feature = "encryption"))]
        Workbook::from_package(self.0)
    }

    /// Consume this facade and expose its clear OPC graph explicitly.
    ///
    /// This plaintext-named compatibility boundary deliberately declassifies
    /// any retained encryption provenance. The returned `OpcPackage` carries
    /// no managed-package encryption policy.
    #[must_use]
    pub fn into_plain_opc(self) -> OpcPackage {
        self.0
    }

    /// Read the workbook's source-bound calculation metadata.
    pub fn calculation_metadata(&self) -> Result<crate::calculation_properties::Snapshot> {
        crate::calculation_properties::Snapshot::load(&self.0)
    }

    /// Read calculation metadata with a caller-supplied resource policy.
    pub fn calculation_metadata_with_limits(
        &self,
        limits: &crate::calculation_properties::Limits,
    ) -> Result<crate::calculation_properties::Snapshot> {
        crate::calculation_properties::Snapshot::load_with_limits(&self.0, limits)
    }

    /// Read the table-owned Survey catalog without changing the package graph.
    pub fn surveys(&self) -> Result<crate::survey::Snapshot> {
        crate::survey::Snapshot::load(&self.0)
    }

    /// Read the table-owned Survey catalog with explicit resource limits.
    pub fn surveys_with_limits(
        &self,
        limits: &crate::survey::Limits,
    ) -> Result<crate::survey::Snapshot> {
        crate::survey::Snapshot::load_with_limits(&self.0, limits)
    }

    /// Read the source-bound workbook Data Model graph, if present.
    pub fn data_model(&self) -> Result<crate::workbook::data_model::Snapshot> {
        crate::workbook::data_model::Snapshot::load(&self.0)
    }

    /// Start a source-bound Data Model transaction.
    pub fn edit_data_model(&mut self) -> Result<crate::workbook::data_model::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_data_model")?;
        crate::workbook::data_model::Transaction::new(&mut self.0)
    }

    /// Atomically create/import or replace a Data Model with outer validation.
    ///
    /// Creation/import validates the typed descriptor, XLDM storage profile,
    /// OPC ownership, and relationship closure while retaining the binary
    /// payload as opaque bytes. A time-grouping write additionally requires a
    /// complete neutral XLDM identity closure for its table, source column,
    /// and calculated columns. For an existing source-bound model,
    /// payload-only replacement with an unchanged typed descriptor is
    /// supported; other structural replacement remains refused until its
    /// complete inner proof exists.
    pub fn put_data_model(&mut self, model: crate::workbook::data_model::Model) -> Result<()> {
        self.ensure_mutation_allowed("put_data_model")?;
        let mut transaction = crate::workbook::data_model::Transaction::new(&mut self.0)?;
        transaction.set(model)?;
        transaction.commit().map(|_commit| ())
    }

    /// Atomically remove the workbook Data Model graph.
    pub fn remove_data_model(&mut self) -> Result<()> {
        self.ensure_mutation_allowed("remove_data_model")?;
        let mut transaction = crate::workbook::data_model::Transaction::new(&mut self.0)?;
        transaction.remove()?;
        transaction.commit().map(|_commit| ())
    }

    /// Apply an exact source-bound Data Model patch.
    pub fn apply_data_model_patch(
        &mut self,
        patch: &crate::workbook::data_model::Patch,
    ) -> Result<()> {
        self.ensure_mutation_allowed("apply_data_model_patch")?;
        patch.apply(&mut self.0)
    }

    /// Read the source-bound Custom Data and Custom Data Properties catalog.
    pub fn custom_data(&self) -> Result<crate::custom_data::Snapshot> {
        crate::custom_data::Snapshot::load(&self.0)
    }

    /// Read Custom Data with explicit payload and storage limits.
    pub fn custom_data_with_limits(
        &self,
        limits: &crate::custom_data::Limits,
    ) -> Result<crate::custom_data::Snapshot> {
        crate::custom_data::Snapshot::load_with_limits(&self.0, limits)
    }

    /// Alias emphasizing that the catalog includes Custom Data Properties.
    pub fn custom_data_properties(&self) -> Result<crate::custom_data::Snapshot> {
        self.custom_data()
    }

    /// Read Custom Data and Custom Data Properties with explicit limits.
    pub fn custom_data_properties_with_limits(
        &self,
        limits: &crate::custom_data::Limits,
    ) -> Result<crate::custom_data::Snapshot> {
        self.custom_data_with_limits(limits)
    }

    /// Start a source-bound Custom Data transaction.
    pub fn edit_custom_data(&mut self) -> Result<crate::custom_data::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_custom_data")?;
        crate::custom_data::Transaction::new(&mut self.0)
    }

    /// Start a Custom Data transaction with explicit limits.
    pub fn edit_custom_data_with_limits(
        &mut self,
        limits: &crate::custom_data::Limits,
    ) -> Result<crate::custom_data::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_custom_data_with_limits")?;
        crate::custom_data::Transaction::with_limits(&mut self.0, limits)
    }

    /// Atomically insert or replace one inert Custom Data storage.
    pub fn put_custom_data(&mut self, value: crate::custom_data::CustomData) -> Result<()> {
        self.ensure_mutation_allowed("put_custom_data")?;
        let mut transaction = crate::custom_data::Transaction::new(&mut self.0)?;
        transaction.upsert(value)?;
        transaction.commit().map(|_commit| ())
    }

    /// Atomically insert or replace one inert Custom Data storage with
    /// explicit resource limits.
    pub fn put_custom_data_with_limits(
        &mut self,
        value: crate::custom_data::CustomData,
        limits: &crate::custom_data::Limits,
    ) -> Result<()> {
        self.ensure_mutation_allowed("put_custom_data_with_limits")?;
        let mut transaction = crate::custom_data::Transaction::with_limits(&mut self.0, limits)?;
        transaction.upsert(value)?;
        transaction.commit().map(|_commit| ())
    }

    /// Atomically remove a Custom Data storage by UID.  Missing UIDs are an
    /// idempotent no-op. Referenced storages require an explicit disposition.
    pub fn remove_custom_data(&mut self, id: &str) -> Result<()> {
        self.remove_custom_data_with(id, crate::custom_data::RemovalDisposition::RejectReferenced)
    }

    /// Atomically remove one Custom Data storage with explicit limits.
    /// Referenced storages require an explicit disposition through
    /// [`Self::remove_custom_data_with_and_limits`].
    pub fn remove_custom_data_with_limits(
        &mut self,
        id: &str,
        limits: &crate::custom_data::Limits,
    ) -> Result<()> {
        self.remove_custom_data_with_and_limits(
            id,
            crate::custom_data::RemovalDisposition::RejectReferenced,
            limits,
        )
    }

    /// Rename a Custom Data storage and its recognized connection references.
    pub fn rename_custom_data(&mut self, id: &str, new_id: impl Into<String>) -> Result<bool> {
        self.ensure_mutation_allowed("rename_custom_data")?;
        let mut transaction = crate::custom_data::Transaction::new(&mut self.0)?;
        let index = transaction
            .entries()
            .iter()
            .position(|entry| entry.id() == id)
            .ok_or_else(|| crate::Error::Invalid("Custom Data storage UID is absent".into()))?;
        let changed = transaction.rename(index, new_id)?;
        transaction.commit()?;
        Ok(changed)
    }

    /// Rename a Custom Data storage and its recognized connection references
    /// with explicit resource limits.
    pub fn rename_custom_data_with_limits(
        &mut self,
        id: &str,
        new_id: impl Into<String>,
        limits: &crate::custom_data::Limits,
    ) -> Result<bool> {
        self.ensure_mutation_allowed("rename_custom_data_with_limits")?;
        let mut transaction = crate::custom_data::Transaction::with_limits(&mut self.0, limits)?;
        let index = transaction
            .entries()
            .iter()
            .position(|entry| entry.id() == id)
            .ok_or_else(|| crate::Error::Invalid("Custom Data storage UID is absent".into()))?;
        let changed = transaction.rename(index, new_id)?;
        transaction.commit()?;
        Ok(changed)
    }

    /// Remove a Custom Data storage with an explicit connection disposition.
    pub fn remove_custom_data_with(
        &mut self,
        id: &str,
        disposition: crate::custom_data::RemovalDisposition,
    ) -> Result<()> {
        self.ensure_mutation_allowed("remove_custom_data")?;
        let mut transaction = crate::custom_data::Transaction::new(&mut self.0)?;
        if let Some(index) = transaction
            .entries()
            .iter()
            .position(|entry| entry.id() == id)
        {
            transaction.remove_with(index, disposition)?;
        }
        transaction.commit().map(|_commit| ())
    }

    /// Remove a Custom Data storage with an explicit connection disposition
    /// and resource limits.
    pub fn remove_custom_data_with_and_limits(
        &mut self,
        id: &str,
        disposition: crate::custom_data::RemovalDisposition,
        limits: &crate::custom_data::Limits,
    ) -> Result<()> {
        self.ensure_mutation_allowed("remove_custom_data_with_and_limits")?;
        let mut transaction = crate::custom_data::Transaction::with_limits(&mut self.0, limits)?;
        if let Some(index) = transaction
            .entries()
            .iter()
            .position(|entry| entry.id() == id)
        {
            transaction.remove_with(index, disposition)?;
        }
        transaction.commit().map(|_commit| ())
    }

    /// Apply an exact source-bound Custom Data patch.
    pub fn apply_custom_data_patch(&mut self, patch: &crate::custom_data::Patch) -> Result<()> {
        self.ensure_mutation_allowed("apply_custom_data_patch")?;
        patch.apply(&mut self.0)
    }

    /// Start a source-bound Survey transaction.
    pub fn edit_surveys(&mut self) -> Result<crate::survey::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_surveys")?;
        crate::survey::Transaction::new(&mut self.0)
    }

    /// Start a Survey transaction with explicit resource limits.
    pub fn edit_surveys_with_limits(
        &mut self,
        limits: &crate::survey::Limits,
    ) -> Result<crate::survey::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_surveys_with_limits")?;
        crate::survey::Transaction::with_limits(&mut self.0, limits)
    }

    /// Apply an exact source-bound Survey patch.
    pub fn apply_surveys_patch(&mut self, patch: &crate::survey::Patch) -> Result<()> {
        self.ensure_mutation_allowed("apply_surveys_patch")?;
        patch.apply(&mut self.0)
    }

    /// Start a source-bound calculation-metadata transaction.
    pub fn edit_calculation_metadata(
        &mut self,
    ) -> Result<crate::calculation_properties::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_calculation_metadata")?;
        crate::calculation_properties::Transaction::new(&mut self.0)
    }

    /// Start a calculation-metadata transaction with explicit limits.
    pub fn edit_calculation_metadata_with_limits(
        &mut self,
        limits: &crate::calculation_properties::Limits,
    ) -> Result<crate::calculation_properties::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_calculation_metadata_with_limits")?;
        crate::calculation_properties::Transaction::with_limits(&mut self.0, limits)
    }

    /// Atomically apply an exact source-bound calculation-metadata patch.
    pub fn apply_calculation_metadata_patch(
        &mut self,
        patch: &crate::calculation_properties::Patch,
    ) -> Result<()> {
        self.ensure_mutation_allowed("apply_calculation_metadata_patch")?;
        patch.apply(&mut self.0)
    }

    /// Read one worksheet's source-bound row and column page breaks.
    ///
    /// # Errors
    ///
    /// Returns an error when the selector or package is invalid, the selected
    /// sheet is not a worksheet, or its page-break XML is invalid.
    pub fn page_breaks<'a>(
        &self,
        selector: impl Into<crate::Selector<'a>>,
    ) -> Result<crate::page_breaks::Snapshot> {
        crate::page_breaks::Snapshot::load(&self.0, selector)
    }

    /// Start a source-bound page-break transaction for one worksheet.
    ///
    /// # Errors
    ///
    /// Returns an error when mutation is prohibited or loading the selected
    /// worksheet fails.
    pub fn edit_page_breaks<'package, 'selector>(
        &'package mut self,
        selector: impl Into<crate::Selector<'selector>>,
    ) -> Result<crate::page_breaks::Transaction<'package>> {
        self.ensure_mutation_allowed("edit_page_breaks")?;
        crate::page_breaks::Transaction::new(&mut self.0, selector)
    }

    /// Atomically apply an exact source-bound worksheet page-break patch.
    ///
    /// # Errors
    ///
    /// Returns an error when mutation is prohibited, the patch source is
    /// stale, or publication/readback fails.
    pub fn apply_page_breaks_patch(&mut self, patch: &crate::page_breaks::Patch) -> Result<()> {
        self.ensure_mutation_allowed("apply_page_breaks_patch")?;
        patch.apply(&mut self.0)
    }

    /// Read the inert typed package-level custom document properties.
    ///
    /// An absent custom-properties relationship and part produce the shared
    /// empty [`Props`] value. The custom-properties package graph follows
    /// MS-OI29500 section 3.11.
    ///
    /// Values are metadata only: this method does not evaluate formulas or
    /// execute macros, controls, VBA, links, external code, or payloads.
    pub fn custom_props(&self) -> Result<Props> {
        Ok(Props::read_for(&self.0, Host::Excel)?)
    }

    /// Atomically publish inert typed package-level custom document properties.
    ///
    /// An empty collection removes the custom-properties graph. The change is
    /// clone-staged and only replaces this facade's OPC package after the
    /// shared OOXML owner has validated and serialized the complete graph.
    /// Changed metadata invalidates existing package signatures.
    pub fn put_custom_props(&mut self, props: Props) -> Result<()> {
        self.ensure_mutation_allowed("put_custom_props")?;
        self.stage(|package| {
            props.write_for(package, Host::Excel)?;
            Ok(())
        })
    }

    /// Atomically remove the package-level custom document properties.
    ///
    /// This idempotent operation rejects corrupt, external, or duplicate
    /// graphs instead of treating them as absent. The custom-properties graph
    /// is defined by MS-OI29500 section 3.11.
    pub fn remove_custom_props(&mut self) -> Result<()> {
        self.ensure_mutation_allowed("remove_custom_props")?;
        self.stage(|package| {
            Props::read_for(package, Host::Excel)?;
            Props::new().write_for(package, Host::Excel)?;
            Ok(())
        })
    }

    /// Serialize the package into owned bytes.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.ensure_ordinary_output("to_bytes")?;
        self.to_plain_bytes()
    }

    /// Explicitly serialize a plaintext OPC package, declassifying encrypted
    /// source provenance for this output only.
    pub fn to_plain_bytes(&self) -> Result<Vec<u8>> {
        writer::to_bytes(&self.0)
    }

    /// Stream the package into a sequential sink.
    pub fn write_to(&self, sink: impl Write) -> Result<()> {
        self.ensure_ordinary_output("write_to")?;
        self.write_plain_to(sink)
    }

    /// Explicitly stream a plaintext OPC package to a sequential sink.
    pub fn write_plain_to(&self, sink: impl Write) -> Result<()> {
        writer::write_to(&self.0, sink)
    }

    /// Encryption profile retained from source ingress or the latest
    /// successful encrypted save. In-memory byte generation is side-effect-free.
    #[cfg(feature = "encryption")]
    #[must_use]
    pub const fn encryption(&self) -> Option<crate::encryption::Mode> {
        self.1.mode()
    }

    /// Serialize and encrypt using an explicitly selected profile.
    #[cfg(feature = "encryption")]
    pub fn to_encrypted(&self, password: &str, mode: crate::encryption::Mode) -> Result<Vec<u8>> {
        self.to_encrypted_with_limits(password, mode, &crate::encryption::Limits::default())
    }

    /// Serialize and encrypt using an explicit encryption resource policy.
    #[cfg(feature = "encryption")]
    pub fn to_encrypted_with_limits(
        &self,
        password: &str,
        mode: crate::encryption::Mode,
        limits: &crate::encryption::Limits,
    ) -> Result<Vec<u8>> {
        self.encrypt_with_mode(password, mode, limits)
    }

    /// Serialize and encrypt using the source package's retained profile.
    #[cfg(feature = "encryption")]
    pub fn to_reencrypted(&self, password: &str) -> Result<Vec<u8>> {
        self.to_reencrypted_with_limits(password, &crate::encryption::Limits::default())
    }

    /// Re-encrypt using the retained profile and an explicit resource policy.
    #[cfg(feature = "encryption")]
    pub fn to_reencrypted_with_limits(
        &self,
        password: &str,
        limits: &crate::encryption::Limits,
    ) -> Result<Vec<u8>> {
        let mode = self.retained_mode("to_reencrypted")?;
        self.to_encrypted_with_limits(password, mode, limits)
    }

    /// Start a clone-staged transaction for worksheet slicers.
    ///
    /// Dropping the returned transaction rolls back; `commit` publishes the
    /// validated feature graph without rebuilding unrelated workbook parts.
    pub fn edit_slicers(&mut self) -> Result<crate::slicer::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_slicers")?;
        crate::slicer::Transaction::new(&mut self.0)
    }

    /// Start a clone-staged transaction for worksheet timelines.
    ///
    /// Timeline cache references are workbook-owned, so the workbook part
    /// identity is captured once at transaction start and is not exposed by
    /// the ordinary package facade.
    pub fn edit_timelines(&mut self) -> Result<crate::timeline::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_timelines")?;
        let workbook = self.0.main_document_part()?.partname().clone();
        crate::timeline::Transaction::new(&mut self.0, &workbook)
    }

    /// Load the inert Office Add-in task-pane graph, when present.
    pub fn task_panes(&self) -> Result<Option<litchi_ooxml_common::web::Panes>> {
        crate::task_panes::load(&self.0)
    }

    /// Start a clone-staged transaction for the package-level task-pane graph.
    ///
    /// Existing XML conformance is retained automatically. New graphs use
    /// Transitional `SpreadsheetML` relationships unless the explicit
    /// [`Self::edit_task_panes_with`] entry point is selected.
    pub fn edit_task_panes(&mut self) -> Result<crate::task_panes::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_task_panes")?;
        let conformance = crate::task_panes::existing_conformance(&self.0)?;
        crate::task_panes::Transaction::new(&mut self.0, conformance)
    }

    /// Start a task-pane transaction with an explicit XML relationship
    /// conformance. The staged graph is published only by
    /// [`crate::task_panes::Transaction::commit`].
    pub fn edit_task_panes_with(
        &mut self,
        conformance: litchi_ooxml_common::web::Conformance,
    ) -> Result<crate::task_panes::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_task_panes_with")?;
        crate::task_panes::Transaction::new(&mut self.0, conformance)
    }

    /// Read one worksheet's inert smart-tag annotations by semantic selector.
    pub fn smart_tags<'a>(
        &self,
        sheet: impl Into<crate::workbook::Selector<'a>>,
    ) -> Result<Option<crate::smart_tags::Collection>> {
        let workbook = self.workbook()?;
        let worksheet = workbook
            .sheet(sheet)?
            .ok_or_else(|| invalid("worksheet selector did not match a sheet"))?;
        worksheet.smart_tags()
    }

    /// Start an atomic smart-tag transaction for a semantic worksheet.
    pub fn edit_smart_tags<'a>(
        &mut self,
        sheet: impl Into<crate::workbook::Selector<'a>>,
    ) -> Result<crate::smart_tags::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_smart_tags")?;
        let worksheet = {
            let workbook = self.workbook()?;
            let worksheet = workbook
                .sheet(sheet)?
                .ok_or_else(|| invalid("worksheet selector did not match a sheet"))?;
            if worksheet.kind() != crate::workbook::WorksheetKind::Worksheet {
                return Err(crate::Error::NotWorksheet {
                    sheet: worksheet.name().to_owned(),
                });
            }
            worksheet.part_uri().clone()
        };
        crate::smart_tags::Transaction::new(&mut self.0, worksheet)
    }

    /// Atomically save the package to a filesystem path.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.save_with_durability(path, litchi_core::Durability::Full)
    }

    /// Atomically save the package to a filesystem path at a caller-chosen
    /// [`Durability`](litchi_core::Durability).
    ///
    /// [`Self::save`] is this method at `Durability::Full`. Every level
    /// publishes the same bytes through the same sibling temporary file and
    /// rename; a weaker level only skips the parent-directory sync
    /// (`FileOnly`) or both synchronizations (`NoSync`), with the crash
    /// guarantees `Durability` states. The level is never stored on the
    /// package.
    pub fn save_with_durability(
        &self,
        path: impl AsRef<Path>,
        durability: litchi_core::Durability,
    ) -> Result<()> {
        self.ensure_ordinary_output("save")?;
        self.save_plain_with_durability(path, durability)
    }

    /// Explicitly save a plaintext OPC package atomically.
    pub fn save_plain(&self, path: impl AsRef<Path>) -> Result<()> {
        self.save_plain_with_durability(path, litchi_core::Durability::Full)
    }

    /// Explicitly save a plaintext OPC package atomically at a caller-chosen
    /// [`Durability`](litchi_core::Durability); see
    /// [`Self::save_with_durability`].
    pub fn save_plain_with_durability(
        &self,
        path: impl AsRef<Path>,
        durability: litchi_core::Durability,
    ) -> Result<()> {
        writer::save(&self.0, path, durability)
    }

    /// Atomically save with an explicitly selected encryption profile.
    #[cfg(feature = "encryption")]
    pub fn save_encrypted(
        &mut self,
        path: impl AsRef<Path>,
        password: &str,
        mode: crate::encryption::Mode,
    ) -> Result<()> {
        self.save_encrypted_with_limits(path, password, mode, &crate::encryption::Limits::default())
    }

    /// Atomically save with explicit encryption profile and resource policy.
    #[cfg(feature = "encryption")]
    pub fn save_encrypted_with_limits(
        &mut self,
        path: impl AsRef<Path>,
        password: &str,
        mode: crate::encryption::Mode,
        limits: &crate::encryption::Limits,
    ) -> Result<()> {
        let output = self.encrypt_with_mode(password, mode, limits)?;
        writer::save_encrypted(&output, path)?;
        self.1.mark_encrypted(mode);
        Ok(())
    }

    /// Atomically save using the encrypted source's retained profile.
    #[cfg(feature = "encryption")]
    pub fn save_reencrypted(&mut self, path: impl AsRef<Path>, password: &str) -> Result<()> {
        self.save_reencrypted_with_limits(path, password, &crate::encryption::Limits::default())
    }

    /// Atomically save using the retained profile and explicit resource policy.
    #[cfg(feature = "encryption")]
    pub fn save_reencrypted_with_limits(
        &mut self,
        path: impl AsRef<Path>,
        password: &str,
        limits: &crate::encryption::Limits,
    ) -> Result<()> {
        let mode = self.retained_mode("save_reencrypted")?;
        self.save_encrypted_with_limits(path, password, mode, limits)
    }

    fn stage<T>(&mut self, mutation: impl FnOnce(&mut OpcPackage) -> Result<T>) -> Result<T> {
        let mut staged = self.0.clone();
        let value = mutation(&mut staged)?;
        self.0 = staged;
        Ok(value)
    }

    fn ensure_ordinary_output(&self, operation: &'static str) -> Result<()> {
        #[cfg(feature = "encryption")]
        self.1
            .ordinary_output()
            .map_err(|source| crate::Error::EncryptionPolicy { operation, source })?;
        #[cfg(not(feature = "encryption"))]
        let _ = operation;
        Ok(())
    }

    fn ensure_mutation_allowed(&self, operation: &'static str) -> Result<()> {
        self.ensure_ordinary_output(operation)
    }

    #[cfg(feature = "encryption")]
    fn retained_mode(&self, operation: &'static str) -> Result<crate::encryption::Mode> {
        self.1
            .require_retained_mode()
            .map_err(|source| crate::Error::EncryptionPolicy { operation, source })
    }

    #[cfg(feature = "encryption")]
    fn encrypt_with_mode(
        &self,
        password: &str,
        mode: crate::encryption::Mode,
        limits: &crate::encryption::Limits,
    ) -> Result<Vec<u8>> {
        let plaintext = self.to_plain_bytes()?;
        Ok(crate::encryption::encrypt_with(
            plaintext, password, mode, limits,
        )?)
    }
}

/// Plaintext-build compatibility conversion.
///
/// Encryption-enabled builds intentionally require
/// [`Package::into_plain_opc`] so provenance loss is explicit.
/// This compatibility shim is deprecated; new code should use the explicit
/// plaintext-named method in all feature configurations.
#[cfg(not(feature = "encryption"))]
impl From<Package> for OpcPackage {
    fn from(package: Package) -> Self {
        package.0
    }
}

pub(crate) fn build_minimal_package() -> Result<OpcPackage> {
    let mut package = OpcPackage::new();
    let workbook_uri = PackURI::new("/xl/workbook.xml").map_err(invalid)?;
    let worksheet_uri = PackURI::new("/xl/worksheets/sheet1.xml").map_err(invalid)?;
    let styles_uri = PackURI::new("/xl/styles.xml").map_err(invalid)?;

    let mut workbook = BlobPart::new(
        workbook_uri,
        ct::SML_SHEET_MAIN.to_string(),
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            r#"<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>"#,
            r#"</workbook>"#
        )
        .as_bytes()
        .to_vec(),
    );
    workbook.rels_mut().try_add_relationship(
        rt::WORKSHEET.to_owned(),
        "worksheets/sheet1.xml".to_owned(),
        "rId1".to_owned(),
        TargetMode::Internal,
    )?;
    workbook.rels_mut().try_add_relationship(
        rt::STYLES.to_owned(),
        "styles.xml".to_owned(),
        "rId2".to_owned(),
        TargetMode::Internal,
    )?;
    package.try_add_part(Box::new(workbook))?;
    package.try_add_part(Box::new(BlobPart::new(
        worksheet_uri,
        ct::SML_WORKSHEET.to_string(),
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
            r#"<dimension ref="A1"/><sheetData/></worksheet>"#
        )
        .as_bytes()
        .to_vec(),
    )))?;
    package.try_add_part(Box::new(BlobPart::new(
        styles_uri,
        ct::SML_STYLES.to_string(),
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
            r#"<fonts count="1"><font/></fonts>"#,
            r#"<fills count="2"><fill><patternFill patternType="none"/></fill>"#,
            r#"<fill><patternFill patternType="gray125"/></fill></fills>"#,
            r#"<borders count="1"><border/></borders>"#,
            r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
            r#"<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>"#,
            r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#,
            r#"</styleSheet>"#
        )
        .as_bytes()
        .to_vec(),
    )))?;
    package.rels_mut().try_add_relationship(
        rt::OFFICE_DOCUMENT.to_owned(),
        "xl/workbook.xml".to_owned(),
        "rId1".to_owned(),
        TargetMode::Internal,
    )?;
    Ok(package)
}

/// Inert analytical-model package resources.
///
/// The nested codecs live in the neutral [`litchi_xldm`] crate. The two
/// package-level entry points retain the historical XLSX error type so code
/// using `crate::package::xldm::inspect` continues to match [`crate::Error`].
pub mod xldm {
    pub use litchi_xldm::{
        BackupLog, Compression, FileEntry, FileGroup, FileGroupClass, FileKind, GeneratedNameKind,
        GeneratedPath, Header, LoggedFile, Offset, OlapProofLimits, PartitionMarker, Size, Storage,
        StorageProfile, WriteAccess, XLDM_PAGE_SIZE, XLDM_STREAM_SIGNATURE, Xldm140Closure,
        Xldm140ClosureMember, Xldm140ColumnBinding, Xldm140ColumnIdentity, Xldm140FileReplacement,
        Xldm140IdentityProjection, Xldm140InversePatch, Xldm140MemberSection, Xldm140Patch,
        Xldm140PatchBytes, Xldm140RelationshipIdentity, Xldm140TableIdentity,
        Xldm140TimeGroupingBinding, Xldm140TimeGroupingContentType, XmlEncoding, compression,
        crypt, generated, identity, metadata, native, olap, olapproof, project_xldm140_identity,
        project_xldm140_identity_with_closure, prove_xldm140_closure,
        validate_xldm140_identity_closure,
    };

    /// Classify a generated path while preserving the XLSX facade error type.
    pub fn classify_generated_path(path: &str) -> crate::error::Result<GeneratedPath> {
        litchi_xldm::classify_generated_path(path).map_err(super::map_xldm_error)
    }

    /// Inspect an XLDM stream while preserving the XLSX facade error type.
    pub fn inspect(bytes: &[u8]) -> crate::error::Result<Storage<'_>> {
        litchi_xldm::inspect(bytes).map_err(super::map_xldm_error)
    }

    /// Inspect a borrowed XLDM stream through the source-sharing API.
    pub fn inspect_shared(bytes: &[u8]) -> crate::error::Result<Storage<'_>> {
        litchi_xldm::inspect_shared(bytes).map_err(super::map_xldm_error)
    }

    /// Write an unchanged or explicitly edited XLDM storage snapshot.
    pub fn write(storage: &Storage<'_>) -> crate::error::Result<Vec<u8>> {
        litchi_xldm::write(storage).map_err(super::map_xldm_error)
    }
}

fn map_xldm_error(error: litchi_xldm::Error) -> crate::error::Error {
    match error {
        litchi_xldm::Error::Invalid(message) => crate::error::Error::Invalid(message),
        litchi_xldm::Error::Unsupported { feature } => crate::error::Error::Unsupported { feature },
        litchi_xldm::Error::Allocation { resource, source } => {
            crate::error::Error::Allocation { resource, source }
        },
        litchi_xldm::Error::Xml(message) => {
            crate::error::Error::Xml(litchi_ooxml_common::XmlError::Malformed(message))
        },
        other => crate::error::Error::Invalid(other.to_string()),
    }
}

#[cfg(test)]
pub(crate) mod xldm_test_support;
