//! The immutable XLSX package owner.
//!
//! [`Package`] is the ADR 0034 owner of every XLSX state verb: opening,
//! creation, the primary [`Package::edit`] transaction and its
//! [`Package::apply`] replay, publication and the security verbs. It owns the
//! validated OPC graph, delegates archive I/O to `litchi-opc`, and exposes the
//! read-only semantic projection through [`Package::workbook`]. A package is
//! a cheap-to-clone snapshot: committing an edit publishes a new package and
//! never changes the source.
//!
//! ```
//! use litchi_xlsx::Package;
//!
//! let source = Package::new()?;
//! let mut edit = source.edit()?;
//! edit.sheet(0_usize)?.expect("one worksheet").set("A1", 42)?;
//! let commit = edit.commit()?;
//! assert!(commit.diagnostics().changed());
//! assert!(!commit.patch().is_noop());
//! let next = commit.into_package();
//! let sheet = next.workbook().sheet(0_usize)?.expect("one worksheet");
//! assert!(sheet.cell("A1")?.stored().is_some());
//! // The source snapshot is unchanged.
//! let source_sheet = source.workbook().sheet(0_usize)?.expect("one worksheet");
//! assert!(source_sheet.cell("A1")?.stored().is_none());
//! # Ok::<(), litchi_xlsx::Error>(())
//! ```

/// Inert worksheet Printer Settings parts and references.
pub mod printer_settings;
mod target;
/// Deterministic minimal package resources.
pub mod template;

pub(crate) use target::Target;

use std::io::{Read, Write};
use std::path::Path;

use litchi_ooxml_common::custom::Host;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};

use crate::OpenOptions;
use crate::Workbook;
use crate::custom::Props;
use crate::error::{Result, invalid};
use crate::workbook::{Commit, Edit, Patch};
use crate::writer;

#[cfg(feature = "encryption")]
use litchi_ooxml_common::package_encryption::PackageEncryption;

/// The immutable, cheap-to-share XLSX package snapshot.
///
/// The package owns the validated OPC graph and its encryption provenance.
/// The ZIP/container implementation remains below `litchi-opc`; this type
/// does not leak archive-specific readers, writers, or errors. Semantic
/// reads go through the owned [`Workbook`] projection returned by
/// [`Self::workbook`].
#[derive(Debug, Clone)]
pub struct Package {
    workbook: Workbook,
}

impl Package {
    /// Create a deterministic minimal workbook package with one visible
    /// worksheet.
    pub fn new() -> Result<Self> {
        Self::from_opc(build_minimal_package()?)
    }

    /// Open and validate an XLSX package from a filesystem path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_options(path, OpenOptions::default())
    }

    /// Open and validate an XLSX package from a filesystem path with explicit
    /// ingress policies.
    pub fn open_with_options(path: impl AsRef<Path>, options: OpenOptions) -> Result<Self> {
        Self::from_validated(
            OpcPackage::open_with_limits(path, options.source_limits)?,
            options,
        )
    }

    /// Read and validate an XLSX package from bytes.
    ///
    /// The bytes are copied once into the package's owned source archive, so
    /// unchanged members keep exact-source publication.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_options(bytes, OpenOptions::default())
    }

    /// Read and validate an XLSX package from bytes with explicit ingress
    /// policies.
    pub fn from_bytes_with_options(bytes: &[u8], options: OpenOptions) -> Result<Self> {
        Self::from_validated(
            OpcPackage::from_vec_with_limits(owned_input(bytes, options)?, options.source_limits)?,
            options,
        )
    }

    /// Read and validate an XLSX package from a synchronous reader.
    pub fn from_reader(reader: impl Read) -> Result<Self> {
        Self::from_reader_with_options(reader, OpenOptions::default())
    }

    /// Read and validate an XLSX package from a synchronous reader with
    /// explicit ingress policies.
    pub fn from_reader_with_options(reader: impl Read, options: OpenOptions) -> Result<Self> {
        Self::from_validated(
            OpcPackage::from_reader_with_limits(reader, options.source_limits)?,
            options,
        )
    }

    /// Open an ordinary or encrypted XLSX package with the default ingress
    /// policies.
    #[cfg(feature = "encryption")]
    pub fn open_with_password(path: impl AsRef<Path>, password: &str) -> Result<Self> {
        Self::open_with_password_and_options(path, password, OpenOptions::default())
    }

    /// Open an ordinary or encrypted XLSX package with independent outer
    /// encryption and inner OPC policies.
    #[cfg(feature = "encryption")]
    pub fn open_with_password_and_options(
        path: impl AsRef<Path>,
        password: &str,
        options: OpenOptions,
    ) -> Result<Self> {
        let file = std::fs::File::open(path).map_err(crate::encryption::Error::Io)?;
        let opened = crate::encryption::load_with(file, password, &options.encryption_limits)?;
        Self::from_decrypted(opened, options)
    }

    /// Read ordinary or encrypted bytes with the default ingress policies.
    #[cfg(feature = "encryption")]
    pub fn from_bytes_with_password(bytes: &[u8], password: &str) -> Result<Self> {
        Self::from_bytes_with_password_and_options(bytes, password, OpenOptions::default())
    }

    /// Read ordinary or encrypted bytes with independent outer encryption and
    /// inner OPC policies.
    #[cfg(feature = "encryption")]
    pub fn from_bytes_with_password_and_options(
        bytes: &[u8],
        password: &str,
        options: OpenOptions,
    ) -> Result<Self> {
        let opened = crate::encryption::load_with(bytes, password, &options.encryption_limits)?;
        Self::from_decrypted(opened, options)
    }

    /// Read an ordinary or encrypted package from a synchronous reader with
    /// the default ingress policies.
    #[cfg(feature = "encryption")]
    pub fn from_reader_with_password(reader: impl Read, password: &str) -> Result<Self> {
        Self::from_reader_with_password_and_options(reader, password, OpenOptions::default())
    }

    /// Read an ordinary or encrypted package from a synchronous reader with
    /// independent outer encryption and inner OPC policies.
    #[cfg(feature = "encryption")]
    pub fn from_reader_with_password_and_options(
        reader: impl Read,
        password: &str,
        options: OpenOptions,
    ) -> Result<Self> {
        let opened = crate::encryption::load_with(reader, password, &options.encryption_limits)?;
        Self::from_decrypted(opened, options)
    }

    /// Open a package that was already decrypted, keeping its encryption
    /// provenance.
    ///
    /// `opened` comes from [`crate::encryption::open`] or one of its limit
    /// and policy variants. The package keeps the source profile exactly as
    /// [`Self::from_bytes_with_password_and_options`] does, so plain output
    /// stays refused until the caller chooses re-encryption or plaintext.
    /// This lets one decryption serve both a caller's own classification of
    /// the clear package and the open. The encryption limits of `options` are
    /// not consulted; decryption already happened.
    ///
    /// # Errors
    ///
    /// Returns an error if the clear package violates the source limits or
    /// is not a valid XLSX package.
    #[cfg(feature = "encryption")]
    pub fn from_decrypted(opened: crate::encryption::Opened, options: OpenOptions) -> Result<Self> {
        let provenance = opened
            .mode()
            .map_or_else(PackageEncryption::plain, PackageEncryption::encrypted);
        let package = OpcPackage::from_vec_with_limits(opened.into_bytes(), options.source_limits)?;
        Ok(Self {
            workbook: Workbook::from_package_with_encryption(
                package,
                provenance,
                options.leniency,
            )?,
        })
    }

    /// Validate and adopt an already parsed OPC package under the strict
    /// leniency.
    ///
    /// This raw compatibility boundary deliberately classifies the adopted
    /// clear OPC graph as plaintext. Callers converting decrypted content to
    /// `OpcPackage` therefore explicitly declassify its encryption provenance.
    pub fn from_opc(package: OpcPackage) -> Result<Self> {
        Ok(Self {
            workbook: Workbook::from_package(package)?,
        })
    }

    fn from_validated(package: OpcPackage, options: OpenOptions) -> Result<Self> {
        Ok(Self {
            workbook: Workbook::from_package_with_leniency(package, options.leniency)?,
        })
    }

    pub(crate) const fn from_workbook(workbook: Workbook) -> Self {
        Self { workbook }
    }

    /// The owned, read-only semantic projection of this snapshot.
    ///
    /// The projection shares this package's immutable state; producing one
    /// costs one reference-count increment.
    #[must_use]
    pub fn workbook(&self) -> Workbook {
        self.workbook.clone()
    }

    /// Borrow the read-only semantic projection of this snapshot.
    #[must_use]
    pub const fn workbook_ref(&self) -> &Workbook {
        &self.workbook
    }

    pub(crate) fn opc(&self) -> &OpcPackage {
        self.workbook.opc_package()
    }

    /// Whether the package carries OPC digital-signature parts or
    /// relationships. Nothing is verified.
    #[must_use]
    pub fn is_signed(&self) -> bool {
        self.opc().is_signed()
    }

    /// Remove every package signature explicitly (ADR 0006).
    ///
    /// Every signature relationship and signature part is removed, and the
    /// resulting unsigned graph is authorized for publication
    /// ([`OpcPackage::unsign`]). Edits never remove a signature themselves: a
    /// change to a signed package is refused with [`crate::Error::Signed`]
    /// until this is called, because the change invalidates every signature.
    /// An exact no-op keeps the signature byte for byte without it. A package
    /// that carries no signature, and no signature-policy requirement, is left
    /// untouched, exact source included. Clones of this package taken before
    /// the call keep their signatures.
    ///
    /// ```
    /// use litchi_xlsx::Package;
    ///
    /// let mut package = Package::new()?;
    /// package.unsign()?;
    /// assert!(!package.is_signed());
    /// # Ok::<(), litchi_xlsx::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package opened with a
    /// password, whose mutations are refused.
    pub fn unsign(&mut self) -> Result<&mut Self> {
        self.workbook = self.workbook.unsigned()?;
        Ok(self)
    }

    /// The opt-in repair policy this package was opened with.
    #[must_use]
    pub fn leniency(&self) -> crate::Leniency {
        self.workbook.leniency()
    }

    /// Compare two snapshots' native OPC content without serializing either
    /// package.
    ///
    /// The comparison walks the package relationship graph and compares every
    /// accepted part's content type, payload bytes, and relationships. It is
    /// therefore suitable for deciding whether a sequence of immutable edits
    /// returned to its original state (for example `A -> B -> A`) is a
    /// logical no-op. Derived workbook caches and parser tolerance reports are
    /// not package content and are intentionally excluded.
    ///
    /// # Errors
    ///
    /// Returns the native package error when a deferred payload cannot be
    /// decoded. Packages containing caller-defined OPC part implementations,
    /// or separately opened opaque non-part ZIP members, return an explicit
    /// unsupported error because their behavior or bytes cannot be proved
    /// through the typed OPC graph.
    pub fn same_content(&self, other: &Self) -> Result<bool> {
        self.workbook.same_content(&other.workbook)
    }

    /// Read the package's core properties, recording the repairs the reader
    /// applies in `report`.
    ///
    /// This reads the clear OPC graph's core-properties part only; it does
    /// not expose the graph or change any retained encryption policy.
    ///
    /// # Errors
    ///
    /// Returns an error if the core-properties part is present but cannot be
    /// read.
    pub fn core_properties_with_report(
        &self,
        report: &mut litchi_ooxml_common::leniency::ToleranceReport,
    ) -> Result<Option<litchi_core::Metadata>> {
        Ok(litchi_ooxml_common::properties::read_with_report(
            self.opc(),
            report,
        )?)
    }

    /// Consume this package and expose its clear OPC graph explicitly.
    ///
    /// This plaintext-named compatibility boundary deliberately declassifies
    /// any retained encryption provenance. The returned `OpcPackage` carries
    /// no managed-package encryption policy.
    #[must_use]
    pub fn into_plain_opc(self) -> OpcPackage {
        self.workbook.into_opc_package()
    }

    /// Start an isolated semantic transaction over this immutable snapshot.
    ///
    /// Dropping the returned [`Edit`] is the rollback; its
    /// [`Edit::commit`] publishes a new package and leaves this one intact.
    pub fn edit(&self) -> Result<Edit> {
        self.workbook.ensure_mutation_allowed("edit")?;
        Edit::new(self.workbook.clone())
    }

    /// Apply a reversible patch after checking every expected source part.
    pub fn apply(&self, patch: &Patch) -> Result<Commit> {
        self.workbook.ensure_mutation_allowed("apply")?;
        patch.apply_to(&self.workbook)
    }

    /// Begin an ordinary semantic edit handle for one workbook PivotCache.
    /// Select the cache field with [`crate::pivot::PivotCacheEdit::field`]
    /// before staging an existing-name replacement.
    #[must_use]
    pub fn edit_pivot_cache(
        &self,
        selector: impl Into<crate::pivot::PivotCacheSelector>,
    ) -> crate::pivot::PivotCacheEdit {
        crate::pivot::cached_unique_names::edit_workbook(&self.workbook, selector)
    }

    /// Start an ordinary semantic PivotTable server-format edit.
    pub fn edit_pivot_table<'a>(
        &self,
        selector: impl Into<crate::pivot::PivotTableSelector<'a>>,
    ) -> Result<crate::pivot::PivotTableServerFormatsEdit> {
        self.workbook.ensure_mutation_allowed("edit_pivot_table")?;
        crate::pivot::server_formats::edit_workbook(&self.workbook, selector)
    }

    /// Apply an exact ordinary server-format patch.
    pub fn apply_pivot_table(
        &self,
        patch: &crate::pivot::PivotTableServerFormatsPackagePatch,
    ) -> Result<crate::pivot::PivotTableServerFormatsPackageCommit> {
        self.workbook.ensure_mutation_allowed("apply_pivot_table")?;
        patch.apply(self)
    }

    /// Begin an ordinary semantic scalar edit of one `pivotTableData` owner.
    pub fn edit_pivot_table_data<'a>(
        &self,
        selector: impl Into<crate::pivot::PivotTableSelector<'a>>,
    ) -> Result<crate::pivot::PivotTableDataEdit> {
        self.workbook
            .ensure_mutation_allowed("edit_pivot_table_data")?;
        crate::pivot::server_formats::table_data::edit_workbook(&self.workbook, selector)
    }

    /// Start a `pivotTableData` edit with a caller-lowered semantic resource
    /// policy.
    pub fn edit_pivot_table_data_with_limits<'a>(
        &self,
        selector: impl Into<crate::pivot::PivotTableSelector<'a>>,
        limits: &crate::pivot::PivotTableDataLimits,
    ) -> Result<crate::pivot::PivotTableDataEdit> {
        self.workbook
            .ensure_mutation_allowed("edit_pivot_table_data_with_limits")?;
        crate::pivot::server_formats::table_data::edit_workbook_with_limits(
            &self.workbook,
            selector,
            limits,
        )
    }

    /// Apply an exact ordinary `pivotTableData` patch.
    pub fn apply_pivot_table_data(
        &self,
        patch: &crate::pivot::PivotTableDataPackagePatch,
    ) -> Result<crate::pivot::PivotTableDataPackageCommit> {
        self.workbook
            .ensure_mutation_allowed("apply_pivot_table_data")?;
        patch.apply(self)
    }

    /// Apply an exact ordinary cached-unique-names patch.
    pub fn apply_pivot_cache(
        &self,
        patch: &crate::pivot::cached_unique_names::PackagePatch,
    ) -> Result<crate::pivot::cached_unique_names::PackageCommit> {
        self.workbook.ensure_mutation_allowed("apply_pivot_cache")?;
        patch.apply(self)
    }

    /// Read the workbook's source-bound calculation metadata.
    pub fn calculation_metadata(&self) -> Result<crate::calculation_properties::Snapshot> {
        crate::calculation_properties::Snapshot::load(self.opc())
    }

    /// Read calculation metadata with a caller-supplied resource policy.
    pub fn calculation_metadata_with_limits(
        &self,
        limits: &crate::calculation_properties::Limits,
    ) -> Result<crate::calculation_properties::Snapshot> {
        crate::calculation_properties::Snapshot::load_with_limits(self.opc(), limits)
    }

    /// Read table-associated survey metadata.
    pub fn surveys(&self) -> Result<crate::survey::Snapshot> {
        crate::survey::Snapshot::load(self.opc())
    }

    /// Read surveys with a caller-supplied resource policy.
    pub fn surveys_with_limits(
        &self,
        limits: &crate::survey::Limits,
    ) -> Result<crate::survey::Snapshot> {
        crate::survey::Snapshot::load_with_limits(self.opc(), limits)
    }

    /// Read the source-bound workbook Data Model graph, if present.
    pub fn data_model(&self) -> Result<crate::workbook::data_model::Snapshot> {
        crate::workbook::data_model::Snapshot::load(self.opc())
    }

    /// Start a source-bound Data Model transaction.
    pub fn edit_data_model(&mut self) -> Result<crate::workbook::data_model::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_data_model")?;
        crate::workbook::data_model::Transaction::for_package(self)
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
        let mut transaction = crate::workbook::data_model::Transaction::for_package(self)?;
        transaction.set(model)?;
        transaction.commit().map(|_commit| ())
    }

    /// Atomically remove the workbook Data Model graph.
    pub fn remove_data_model(&mut self) -> Result<()> {
        self.ensure_mutation_allowed("remove_data_model")?;
        let mut transaction = crate::workbook::data_model::Transaction::for_package(self)?;
        transaction.remove()?;
        transaction.commit().map(|_commit| ())
    }

    /// Apply an exact source-bound Data Model patch.
    pub fn apply_data_model_patch(
        &mut self,
        patch: &crate::workbook::data_model::Patch,
    ) -> Result<()> {
        self.ensure_mutation_allowed("apply_data_model_patch")?;
        self.stage(|package| patch.apply(package))
    }

    /// Read the source-bound Custom Data and Custom Data Properties catalog.
    pub fn custom_data(&self) -> Result<crate::custom_data::Snapshot> {
        crate::custom_data::Snapshot::load(self.opc())
    }

    /// Read Custom Data with explicit payload and storage limits.
    pub fn custom_data_with_limits(
        &self,
        limits: &crate::custom_data::Limits,
    ) -> Result<crate::custom_data::Snapshot> {
        crate::custom_data::Snapshot::load_with_limits(self.opc(), limits)
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
        crate::custom_data::Transaction::for_package(self, &crate::custom_data::Limits::default())
    }

    /// Start a Custom Data transaction with explicit limits.
    pub fn edit_custom_data_with_limits(
        &mut self,
        limits: &crate::custom_data::Limits,
    ) -> Result<crate::custom_data::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_custom_data_with_limits")?;
        crate::custom_data::Transaction::for_package(self, limits)
    }

    /// Atomically insert or replace one inert Custom Data storage.
    pub fn put_custom_data(&mut self, value: crate::custom_data::CustomData) -> Result<()> {
        self.put_custom_data_with_limits(value, &crate::custom_data::Limits::default())
    }

    /// Atomically insert or replace one inert Custom Data storage with
    /// explicit resource limits.
    pub fn put_custom_data_with_limits(
        &mut self,
        value: crate::custom_data::CustomData,
        limits: &crate::custom_data::Limits,
    ) -> Result<()> {
        self.ensure_mutation_allowed("put_custom_data")?;
        let mut transaction = crate::custom_data::Transaction::for_package(self, limits)?;
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
        self.rename_custom_data_with_limits(id, new_id, &crate::custom_data::Limits::default())
    }

    /// Rename a Custom Data storage and its recognized connection references
    /// with explicit resource limits.
    pub fn rename_custom_data_with_limits(
        &mut self,
        id: &str,
        new_id: impl Into<String>,
        limits: &crate::custom_data::Limits,
    ) -> Result<bool> {
        self.ensure_mutation_allowed("rename_custom_data")?;
        let mut transaction = crate::custom_data::Transaction::for_package(self, limits)?;
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
        self.remove_custom_data_with_and_limits(
            id,
            disposition,
            &crate::custom_data::Limits::default(),
        )
    }

    /// Remove a Custom Data storage with an explicit connection disposition
    /// and resource limits.
    pub fn remove_custom_data_with_and_limits(
        &mut self,
        id: &str,
        disposition: crate::custom_data::RemovalDisposition,
        limits: &crate::custom_data::Limits,
    ) -> Result<()> {
        self.ensure_mutation_allowed("remove_custom_data")?;
        let mut transaction = crate::custom_data::Transaction::for_package(self, limits)?;
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
        self.stage(|package| patch.apply(package))
    }

    /// Start a source-bound Survey transaction.
    pub fn edit_surveys(&mut self) -> Result<crate::survey::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_surveys")?;
        crate::survey::Transaction::for_package(self, &crate::survey::Limits::default())
    }

    /// Start a Survey transaction with explicit resource limits.
    pub fn edit_surveys_with_limits(
        &mut self,
        limits: &crate::survey::Limits,
    ) -> Result<crate::survey::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_surveys_with_limits")?;
        crate::survey::Transaction::for_package(self, limits)
    }

    /// Apply an exact source-bound Survey patch.
    pub fn apply_surveys_patch(&mut self, patch: &crate::survey::Patch) -> Result<()> {
        self.ensure_mutation_allowed("apply_surveys_patch")?;
        self.stage(|package| patch.apply(package))
    }

    /// Start a source-bound calculation-metadata transaction.
    pub fn edit_calculation_metadata(
        &mut self,
    ) -> Result<crate::calculation_properties::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_calculation_metadata")?;
        crate::calculation_properties::Transaction::for_package(
            self,
            &crate::calculation_properties::Limits::default(),
        )
    }

    /// Start a calculation-metadata transaction with explicit limits.
    pub fn edit_calculation_metadata_with_limits(
        &mut self,
        limits: &crate::calculation_properties::Limits,
    ) -> Result<crate::calculation_properties::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_calculation_metadata_with_limits")?;
        crate::calculation_properties::Transaction::for_package(self, limits)
    }

    /// Atomically apply an exact source-bound calculation-metadata patch.
    pub fn apply_calculation_metadata_patch(
        &mut self,
        patch: &crate::calculation_properties::Patch,
    ) -> Result<()> {
        self.ensure_mutation_allowed("apply_calculation_metadata_patch")?;
        self.stage(|package| patch.apply(package))
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
        crate::page_breaks::Snapshot::load(self.opc(), selector)
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
        crate::page_breaks::Transaction::for_package(self, selector)
    }

    /// Atomically apply an exact source-bound worksheet page-break patch.
    ///
    /// # Errors
    ///
    /// Returns an error when mutation is prohibited, the patch source is
    /// stale, or publication/readback fails.
    pub fn apply_page_breaks_patch(&mut self, patch: &crate::page_breaks::Patch) -> Result<()> {
        self.ensure_mutation_allowed("apply_page_breaks_patch")?;
        self.stage(|package| patch.apply(package))
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
        Ok(Props::read_for(self.opc(), Host::Excel)?)
    }

    /// Atomically publish inert typed package-level custom document properties.
    ///
    /// An empty collection removes the custom-properties graph. The change is
    /// clone-staged and only replaces this package's state after the shared
    /// OOXML owner has validated and serialized the complete graph.
    /// Properties that serialize to the stored bytes publish nothing, so a
    /// signed package keeps its signature; changed metadata on a signed
    /// package is refused with [`crate::Error::Signed`] and leaves the package
    /// unchanged (ADR 0006): call [`Self::unsign`] first.
    pub fn put_custom_props(&mut self, props: Props) -> Result<()> {
        self.ensure_mutation_allowed("put_custom_props")?;
        self.stage(|package| {
            props
                .write_for(package, Host::Excel)
                .map_err(crate::signature_policy::from_shared)?;
            Ok(())
        })
    }

    /// Atomically remove the package-level custom document properties.
    ///
    /// This idempotent operation rejects corrupt, external, or duplicate
    /// graphs instead of treating them as absent. The custom-properties graph
    /// is defined by MS-OI29500 section 3.11. Removing an absent graph keeps a
    /// signed package's signature; removing a present one from a signed
    /// package is refused with [`crate::Error::Signed`] (ADR 0006): call
    /// [`Self::unsign`] first.
    pub fn remove_custom_props(&mut self) -> Result<()> {
        self.ensure_mutation_allowed("remove_custom_props")?;
        self.stage(|package| {
            Props::read_for(package, Host::Excel)?;
            Props::new()
                .write_for(package, Host::Excel)
                .map_err(crate::signature_policy::from_shared)?;
            Ok(())
        })
    }

    /// Start a clone-staged transaction for worksheet slicers.
    ///
    /// Dropping the returned transaction rolls back; `commit` publishes the
    /// validated feature graph without rebuilding unrelated workbook parts.
    /// On a signed package every changing operation is refused with
    /// [`crate::Error::Signed`]; call [`Self::unsign`] first to edit it.
    pub fn edit_slicers(&mut self) -> Result<crate::slicer::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_slicers")?;
        crate::slicer::Transaction::for_package(self)
    }

    /// Start a clone-staged transaction for worksheet timelines.
    ///
    /// Timeline cache references are workbook-owned, so the workbook part
    /// identity is captured once at transaction start and is not exposed by
    /// the ordinary package facade. On a signed package every changing
    /// operation is refused with [`crate::Error::Signed`]; call
    /// [`Self::unsign`] first to edit it.
    pub fn edit_timelines(&mut self) -> Result<crate::timeline::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_timelines")?;
        let workbook = self.opc().main_document_part()?.partname().clone();
        crate::timeline::Transaction::for_package(self, &workbook)
    }

    /// Load the inert Office Add-in task-pane graph, when present.
    pub fn task_panes(&self) -> Result<Option<litchi_ooxml_common::web::Panes>> {
        crate::task_panes::load(self.opc())
    }

    /// Start a clone-staged transaction for the package-level task-pane graph.
    ///
    /// Existing XML conformance is retained automatically. New graphs use
    /// Transitional `SpreadsheetML` relationships unless the explicit
    /// [`Self::edit_task_panes_with`] entry point is selected.
    pub fn edit_task_panes(&mut self) -> Result<crate::task_panes::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_task_panes")?;
        let conformance = crate::task_panes::existing_conformance(self.opc())?;
        crate::task_panes::Transaction::for_package(self, conformance)
    }

    /// Start a task-pane transaction with an explicit XML relationship
    /// conformance. The staged graph is published only by
    /// [`crate::task_panes::Transaction::commit`].
    pub fn edit_task_panes_with(
        &mut self,
        conformance: litchi_ooxml_common::web::Conformance,
    ) -> Result<crate::task_panes::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_task_panes_with")?;
        crate::task_panes::Transaction::for_package(self, conformance)
    }

    /// Read one worksheet's inert smart-tag annotations by semantic selector.
    pub fn smart_tags<'a>(
        &self,
        sheet: impl Into<crate::workbook::Selector<'a>>,
    ) -> Result<Option<crate::smart_tags::Collection>> {
        let worksheet = self
            .workbook
            .sheet(sheet)?
            .ok_or_else(|| invalid("worksheet selector did not match a sheet"))?;
        worksheet.smart_tags()
    }

    /// Start an atomic smart-tag transaction for a semantic worksheet.
    ///
    /// On a signed package every operation that changes the worksheet is
    /// refused with [`crate::Error::Signed`]; call [`Self::unsign`] first to
    /// edit it.
    pub fn edit_smart_tags<'a>(
        &mut self,
        sheet: impl Into<crate::workbook::Selector<'a>>,
    ) -> Result<crate::smart_tags::Transaction<'_>> {
        self.ensure_mutation_allowed("edit_smart_tags")?;
        let worksheet = {
            let worksheet = self
                .workbook
                .sheet(sheet)?
                .ok_or_else(|| invalid("worksheet selector did not match a sheet"))?;
            if worksheet.kind() != crate::workbook::WorksheetKind::Worksheet {
                return Err(crate::Error::NotWorksheet {
                    sheet: worksheet.name().to_owned(),
                });
            }
            worksheet.part_uri().clone()
        };
        crate::smart_tags::Transaction::new(Target::Package(self), worksheet)
    }

    /// Serialize the package into owned bytes.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.workbook.ensure_ordinary_output("to_bytes")?;
        self.to_plain_bytes()
    }

    /// Explicitly serialize a plaintext OPC package, declassifying encrypted
    /// source provenance for this output only.
    pub fn to_plain_bytes(&self) -> Result<Vec<u8>> {
        writer::to_bytes(self.opc())
    }

    /// Stream the package into a sequential sink without seeking.
    ///
    /// A sink failure can leave caller-owned output incomplete. Use
    /// [`Self::save`] for atomic filesystem replacement.
    pub fn write_to(&self, sink: impl Write) -> Result<()> {
        self.workbook.ensure_ordinary_output("write_to")?;
        self.write_plain_to(sink)
    }

    /// Explicitly stream a plaintext OPC package to a sequential sink.
    pub fn write_plain_to(&self, sink: impl Write) -> Result<()> {
        writer::write_to(self.opc(), sink)
    }

    /// Atomically save through a finalized sibling temporary artifact.
    ///
    /// Serialization, flushing, and file synchronization finish before the
    /// destination is replaced. Existing symbolic-link destinations are
    /// refused instead of being followed or silently replaced. This is
    /// [`Self::save_with_durability`] at `Durability::Full`.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.save_with_durability(path, litchi_core::Durability::Full)
    }

    /// Atomically save at a caller-chosen
    /// [`Durability`](litchi_core::Durability).
    ///
    /// Every level serializes and flushes the same bytes into the same
    /// sibling temporary artifact, refuses the same destinations and replaces
    /// the destination with one rename, so a failure before the rename leaves
    /// it untouched. A weaker level only skips the parent-directory sync
    /// (`FileOnly`) or both the file and directory syncs (`NoSync`), with the
    /// crash guarantees `Durability` states. The level applies to this call
    /// only; it is never stored on the package.
    pub fn save_with_durability(
        &self,
        path: impl AsRef<Path>,
        durability: litchi_core::Durability,
    ) -> Result<()> {
        self.workbook.ensure_ordinary_output("save")?;
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
        writer::save(self.opc(), path, durability)
    }

    /// Encryption profile retained from source ingress or the latest
    /// successful encrypted save. In-memory byte generation is side-effect-free.
    #[cfg(feature = "encryption")]
    #[must_use]
    pub const fn encryption(&self) -> Option<crate::encryption::Mode> {
        self.workbook.encryption_state().mode()
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
        let plaintext = self.to_plain_bytes()?;
        Ok(crate::encryption::encrypt_with(
            plaintext, password, mode, limits,
        )?)
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
        let mode = self.workbook.retained_mode("to_reencrypted")?;
        self.to_encrypted_with_limits(password, mode, limits)
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
        let output = self.to_encrypted_with_limits(password, mode, limits)?;
        writer::save_encrypted(&output, path)?;
        self.workbook.mark_encrypted(mode);
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
        let mode = self.workbook.retained_mode("save_reencrypted")?;
        self.save_encrypted_with_limits(path, password, mode, limits)
    }

    /// Replace this snapshot with a validated publication of `staged`.
    ///
    /// The candidate is validated as a complete workbook before this package
    /// changes; on failure the package is unchanged. Snapshots cloned from
    /// this package earlier are never affected.
    pub(crate) fn republish(&mut self, staged: OpcPackage) -> Result<()> {
        self.workbook = Workbook::from_package_with_styles(staged, Some(&self.workbook))?;
        Ok(())
    }

    fn stage<T>(&mut self, mutation: impl FnOnce(&mut OpcPackage) -> Result<T>) -> Result<T> {
        let mut staged = self.opc().clone();
        let value = mutation(&mut staged)?;
        self.republish(staged)?;
        Ok(value)
    }

    fn ensure_mutation_allowed(&self, operation: &'static str) -> Result<()> {
        self.workbook.ensure_mutation_allowed(operation)
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
        package.into_plain_opc()
    }
}

/// Unit tests read the projection through the owner directly; the public
/// API always goes through [`Package::workbook`].
#[cfg(test)]
impl std::ops::Deref for Package {
    type Target = Workbook;

    fn deref(&self) -> &Workbook {
        &self.workbook
    }
}

/// Copy borrowed input into the owned source archive after the input-size
/// admission check, so an oversized input is refused before it is copied.
fn owned_input(bytes: &[u8], options: OpenOptions) -> Result<Vec<u8>> {
    let observed = u64::try_from(bytes.len()).unwrap_or(u64::MAX);
    let maximum = options.source_limits.max_input_bytes();
    if observed > maximum {
        return Err(litchi_opc::OpcError::Limit {
            kind: litchi_opc::SourceLimitKind::InputBytes,
            observed,
            maximum,
        }
        .into());
    }
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(bytes.len())
        .map_err(|source| crate::error::allocation("package input bytes", source))?;
    owned.extend_from_slice(bytes);
    Ok(owned)
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
