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
/// Deterministic minimal package resources.
pub mod template;

use std::io::{Read, Write};
use std::path::Path;

use litchi_ooxml_common::custom::Host;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, SourceLimits, TargetMode};

use crate::OpenOptions;
use crate::Workbook;
use crate::custom::Props;
use crate::error::{Result, invalid};
use crate::workbook::{Commit, Edit, EditOptions, Patch};
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
        Self::from_validated(OpcPackage::open(path, options.source_limits)?, options)
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
            OpcPackage::from_vec(owned_input(bytes, &options)?, options.source_limits)?,
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
            OpcPackage::from_reader(reader, options.source_limits)?,
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
        let opened = crate::encryption::load(file, password, &options.encryption_limits)?;
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
        let opened = crate::encryption::load(bytes, password, &options.encryption_limits)?;
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
        let opened = crate::encryption::load(reader, password, &options.encryption_limits)?;
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
        let package = OpcPackage::from_vec(opened.into_bytes(), options.source_limits)?;
        Ok(Self {
            workbook: Workbook::from_package_with_encryption(package, provenance, options)?,
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
            workbook: Workbook::from_package_with_options(package, options)?,
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

    /// Conservative retained-data charge for one immutable package endpoint.
    /// This is used by format adapters when a private signed no-op probe is
    /// discarded after its native identity patch is rebuilt. It does not
    /// force lazy cell or feature projections. Decoded OPC payloads remain
    /// charged; derived cell and shared-string read caches are excluded.
    pub fn retained_bytes(&self) -> usize {
        let archive_hint = self.opc().exact_source_len().unwrap_or(0);
        crate::workbook::edit::workbook_retained_bytes(self.workbook_ref(), archive_hint)
    }

    /// Whether the package carries OPC digital-signature parts or
    /// relationships. Nothing is verified.
    #[must_use]
    pub fn is_signed(&self) -> bool {
        self.opc().is_signed()
    }

    /// A copy of this package with every package signature removed
    /// explicitly (ADR 0006).
    ///
    /// Every signature relationship and signature part is removed, and the
    /// resulting unsigned graph is authorized for publication
    /// ([`OpcPackage::unsign`]). Edits never remove a signature themselves: a
    /// change to a signed package is refused with [`crate::Error::Signed`]
    /// until this is called, because the change invalidates every signature.
    /// An exact no-op keeps the signature byte for byte without it. A package
    /// that carries no signature, and no signature-policy requirement, is
    /// returned unchanged, exact source included. This package keeps its
    /// signature.
    ///
    /// ```
    /// use litchi_xlsx::Package;
    ///
    /// let package = Package::new()?.unsign()?;
    /// assert!(!package.is_signed());
    /// # Ok::<(), litchi_xlsx::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package opened with a
    /// password, whose mutations are refused.
    pub fn unsign(&self) -> Result<Self> {
        Ok(Self {
            workbook: self.workbook.unsigned()?,
        })
    }

    /// The semantic budgets this package keeps for its bounded reads and the
    /// default limits of its record-family edits ([`crate::semantic_limits`]).
    #[must_use]
    pub fn semantic_limits(&self) -> &crate::SemanticLimits {
        self.workbook.semantic_limits()
    }

    /// The OPC source policy retained by this immutable package.
    #[must_use]
    pub fn source_limits(&self) -> SourceLimits {
        self.workbook.opc_package().read_limits()
    }

    /// This package under other semantic budgets; the package state is
    /// shared, not copied.
    #[must_use]
    pub fn with_semantic_limits(&self, limits: crate::SemanticLimits) -> Self {
        Self {
            workbook: self.workbook.with_semantic_limits(limits),
        }
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
        self.edit_with_options(EditOptions::default())
    }

    /// Start an isolated semantic transaction with an optional retained-byte
    /// admission budget. The budget is checked before a worksheet action can
    /// allocate a changed store or compaction buffer.
    pub fn edit_with_options(&self, options: EditOptions) -> Result<Edit> {
        self.workbook.ensure_mutation_allowed("edit")?;
        Edit::new_with_options(self.workbook.clone(), options)
    }

    /// Apply a reversible patch after checking every expected source part.
    pub fn apply(&self, patch: &Patch) -> Result<Commit> {
        self.workbook.ensure_mutation_allowed("apply")?;
        patch.apply_to(&self.workbook)
    }

    /// Start a cached unique-name edit of one field of one workbook
    /// PivotCache.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused and the resolution and validation errors of the cache field.
    pub fn edit_pivot_cache<'a>(
        &self,
        cache: impl Into<crate::pivot::PivotCacheSelector>,
        field: crate::pivot::PivotCacheFieldSelector<'a>,
    ) -> Result<crate::pivot::cached_unique_names::Edit> {
        self.workbook.ensure_mutation_allowed("edit_pivot_cache")?;
        crate::pivot::cached_unique_names::edit_workbook(&self.workbook, cache, field)
    }

    /// Apply an exact cached unique-name patch to a new package; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] when this package is not the
    /// patch's before state and [`crate::Error::Signed`] for a changed patch
    /// of a signed workbook.
    pub fn apply_pivot_cache(
        &self,
        patch: &crate::pivot::cached_unique_names::Patch,
    ) -> Result<crate::pivot::cached_unique_names::Commit> {
        self.workbook.ensure_mutation_allowed("apply_pivot_cache")?;
        patch.apply(self)
    }

    /// Start a server-format edit of one semantic PivotTable.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused and the resolution and validation errors of the PivotTable.
    pub fn edit_pivot_table<'a>(
        &self,
        selector: impl Into<crate::pivot::PivotTableSelector<'a>>,
    ) -> Result<crate::pivot::server_formats::Edit> {
        self.workbook.ensure_mutation_allowed("edit_pivot_table")?;
        crate::pivot::server_formats::edit_workbook(&self.workbook, selector)
    }

    /// Apply an exact server-format patch to a new package; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] when this package is not the
    /// patch's before state and [`crate::Error::Signed`] for a changed patch
    /// of a signed workbook.
    pub fn apply_pivot_table(
        &self,
        patch: &crate::pivot::server_formats::Patch,
    ) -> Result<crate::pivot::server_formats::Commit> {
        self.workbook.ensure_mutation_allowed("apply_pivot_table")?;
        patch.apply(self)
    }

    /// Start a scalar edit of one PivotTable's `pivotTableData` owner under
    /// the package's semantic limits
    /// ([`crate::SemanticLimits::pivot_table_data`]).
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused and the resolution, limit and validation errors of the owner.
    pub fn edit_pivot_table_data<'a>(
        &self,
        selector: impl Into<crate::pivot::PivotTableSelector<'a>>,
    ) -> Result<crate::pivot::table_data::Edit> {
        self.edit_pivot_table_data_with_options(
            selector,
            crate::pivot::table_data::EditOptions::default(),
        )
    }

    /// Start a `pivotTableData` edit with explicit options.
    ///
    /// # Errors
    ///
    /// See [`Self::edit_pivot_table_data`].
    pub fn edit_pivot_table_data_with_options<'a>(
        &self,
        selector: impl Into<crate::pivot::PivotTableSelector<'a>>,
        options: crate::pivot::table_data::EditOptions,
    ) -> Result<crate::pivot::table_data::Edit> {
        self.workbook
            .ensure_mutation_allowed("edit_pivot_table_data")?;
        let limits = options.limits_or(self.semantic_limits().pivot_table_data_value());
        crate::pivot::table_data::edit_workbook_with_limits(&self.workbook, selector, &limits)
    }

    /// Apply an exact `pivotTableData` patch to a new package; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] when this package is not the
    /// patch's before state and [`crate::Error::Signed`] for a changed patch
    /// of a signed workbook.
    pub fn apply_pivot_table_data(
        &self,
        patch: &crate::pivot::table_data::Patch,
    ) -> Result<crate::pivot::table_data::Commit> {
        self.workbook
            .ensure_mutation_allowed("apply_pivot_table_data")?;
        patch.apply(self)
    }

    /// Read the workbook's source-bound calculation metadata under the
    /// package's semantic limits
    /// ([`crate::SemanticLimits::calculation_properties`]).
    pub fn calculation_metadata(&self) -> Result<crate::calculation_properties::Snapshot> {
        self.workbook.calculation_metadata()
    }

    /// Read table-associated survey metadata under the package's semantic
    /// limits ([`crate::SemanticLimits::survey`]).
    pub fn surveys(&self) -> Result<crate::survey::Snapshot> {
        crate::survey::Snapshot::load_with_limits(
            self.opc(),
            &self.semantic_limits().survey_value(),
        )
    }

    /// Read the source-bound workbook Data Model graph, if present.
    pub fn data_model(&self) -> Result<crate::workbook::data_model::Snapshot> {
        crate::workbook::data_model::Snapshot::load(self.opc())
    }

    /// Read the source-bound Custom Data and Custom Data Properties catalog
    /// under the package's semantic limits
    /// ([`crate::SemanticLimits::custom_data`]).
    pub fn custom_data(&self) -> Result<crate::custom_data::Snapshot> {
        crate::custom_data::Snapshot::load_with_limits(
            self.opc(),
            &self.semantic_limits().custom_data_value(),
        )
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

    /// Read the inert typed package-level custom document properties.
    ///
    /// An absent custom-properties relationship and part produce the shared
    /// empty [`Props`] value. The custom-properties package graph follows
    /// MS-OI29500 section 3.11. Edit them through the metadata family
    /// ([`crate::metadata::Edit::set_custom_props`]); use
    /// [`crate::metadata::Edit::remove_custom_props`] for explicit graph
    /// removal.
    ///
    /// Values are metadata only: this method does not evaluate formulas or
    /// execute macros, controls, VBA, links, external code, or payloads.
    pub fn custom_props(&self) -> Result<Props> {
        Ok(Props::read_for(self.opc(), Host::Excel)?)
    }

    /// Load the inert Office Add-in task-pane graph, when present.
    pub fn task_panes(&self) -> Result<Option<litchi_ooxml_common::web::Panes>> {
        crate::task_panes::load(self.opc())
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
    pub fn encryption(&self) -> Option<crate::encryption::Mode> {
        self.workbook.encryption_state().mode()
    }

    /// Serialize and encrypt using an explicitly selected profile, under the
    /// encryption limits this package was opened with.
    #[cfg(feature = "encryption")]
    pub fn to_encrypted(&self, password: &str, mode: crate::encryption::Mode) -> Result<Vec<u8>> {
        let plaintext = self.to_plain_bytes()?;
        Ok(crate::encryption::encrypt(
            plaintext,
            password,
            mode,
            &self.workbook.encryption_limits(),
        )?)
    }

    /// Serialize and encrypt using the source package's retained profile.
    #[cfg(feature = "encryption")]
    pub fn to_reencrypted(&self, password: &str) -> Result<Vec<u8>> {
        let mode = self.workbook.retained_mode("to_reencrypted")?;
        self.to_encrypted(password, mode)
    }

    /// Atomically save with an explicitly selected encryption profile.
    ///
    /// The package is unchanged: its retained encryption profile stays the
    /// one it was opened with.
    #[cfg(feature = "encryption")]
    pub fn save_encrypted(
        &self,
        path: impl AsRef<Path>,
        password: &str,
        mode: crate::encryption::Mode,
    ) -> Result<()> {
        let output = self.to_encrypted(password, mode)?;
        writer::save_encrypted(&output, path)
    }

    /// Atomically save using the encrypted source's retained profile.
    #[cfg(feature = "encryption")]
    pub fn save_reencrypted(&self, path: impl AsRef<Path>, password: &str) -> Result<()> {
        let mode = self.workbook.retained_mode("save_reencrypted")?;
        self.save_encrypted(path, password, mode)
    }

    /// The part name of the worksheet `sheet` selects, for a staged edit of
    /// `operation`: a selector that matches no sheet is a `NotFound` refusal,
    /// and a chart or dialog sheet is [`crate::Error::NotWorksheet`].
    pub(crate) fn worksheet_part<'a>(
        &self,
        sheet: impl Into<crate::workbook::Selector<'a>>,
        operation: &'static str,
    ) -> Result<PackURI> {
        let worksheet = self.workbook.sheet(sheet)?.ok_or(crate::Error::refused(
            litchi_core::edit::RefusalKind::NotFound,
            operation,
            "the worksheet selector matches no sheet",
        ))?;
        if worksheet.kind() != crate::workbook::WorksheetKind::Worksheet {
            return Err(crate::Error::NotWorksheet {
                sheet: worksheet.name().to_owned(),
            });
        }
        Ok(worksheet.part_uri().clone())
    }

    /// A new package publishing `opc`, validated as a complete workbook,
    /// under this package's policies (leniency, limits and encryption
    /// provenance) and sharing its style lineage. This package is unchanged.
    pub(crate) fn derive(&self, opc: OpcPackage) -> Result<Self> {
        self.workbook
            .adopt_published_package(opc)
            .map(Self::from_workbook)
    }

    /// Run `mutation` on a copy of this package's OPC graph and publish the
    /// result as a new package; this package is unchanged.
    #[allow(dead_code, reason = "used by the record families")]
    pub(crate) fn staged<T>(
        &self,
        mutation: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<(Self, T)> {
        let mut staged = self.opc().clone();
        let value = mutation(&mut staged)?;
        Ok((self.derive(staged)?, value))
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
fn owned_input(bytes: &[u8], options: &OpenOptions) -> Result<Vec<u8>> {
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
/// The nested codecs live in the neutral [`litchi_xldm`] crate. The three
/// package-level entry points retain the historical XLSX error type so code
/// using `crate::package::xldm::inspect` continues to match [`crate::Error`].
pub mod xldm {
    pub use litchi_xldm::{
        BackupLog, Compression, FileEntry, FileGroup, FileGroupClass, FileKind, GeneratedNameKind,
        GeneratedPath, Header, Limits, LoggedFile, Offset, OlapProofLimits, PartitionMarker, Size,
        Storage, StorageProfile, WriteAccess, XLDM_PAGE_SIZE, XLDM_STREAM_SIGNATURE,
        Xldm140Closure, Xldm140ClosureMember, Xldm140ColumnBinding, Xldm140ColumnIdentity,
        Xldm140FileReplacement, Xldm140IdentityProjection, Xldm140InversePatch,
        Xldm140MemberSection, Xldm140Patch, Xldm140PatchBytes, Xldm140RelationshipIdentity,
        Xldm140TableIdentity, Xldm140TimeGroupingBinding, Xldm140TimeGroupingContentType,
        XmlEncoding, compression, crypt, generated, identity, metadata, native, olap, olapproof,
        project_xldm140_identity, project_xldm140_identity_with_closure, prove_xldm140_closure,
        validate_xldm140_identity_closure,
    };

    /// Classify a generated path while preserving the XLSX facade error type.
    pub fn classify_generated_path(path: &str) -> crate::error::Result<GeneratedPath> {
        litchi_xldm::classify_generated_path(path).map_err(super::map_xldm_error)
    }

    /// Inspect an XLDM stream while preserving the XLSX facade error type.
    pub fn inspect<'a>(bytes: &'a [u8], limits: &Limits) -> crate::error::Result<Storage<'a>> {
        litchi_xldm::inspect(bytes, limits).map_err(super::map_xldm_error)
    }

    /// Inspect a borrowed XLDM stream through the source-sharing API.
    pub fn inspect_shared<'a>(
        bytes: &'a [u8],
        limits: &Limits,
    ) -> crate::error::Result<Storage<'a>> {
        litchi_xldm::inspect_shared(bytes, limits).map_err(super::map_xldm_error)
    }

    /// Write an unchanged or explicitly edited XLDM storage snapshot.
    pub fn write(storage: &Storage<'_>, limits: &Limits) -> crate::error::Result<Vec<u8>> {
        litchi_xldm::write(storage, limits).map_err(super::map_xldm_error)
    }
}

fn map_xldm_error(error: litchi_xldm::Error) -> crate::error::Error {
    match error {
        litchi_xldm::Error::Invalid(message) => crate::error::Error::Invalid(message),
        litchi_xldm::Error::Limit {
            kind,
            observed,
            maximum,
        } => crate::error::Error::Limit {
            kind: map_xldm_limit_kind(kind),
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_xldm::Error::InvalidLimit {
            kind,
            value,
            maximum,
        } => crate::error::Error::InvalidLimit {
            kind: map_xldm_limit_kind(kind),
            value: value as u64,
            maximum: maximum as u64,
        },
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

pub(crate) fn map_xldm_limit_kind(kind: litchi_xldm::LimitKind) -> crate::error::LimitKind {
    use crate::workbook::data_model::LimitKind as DataModel;
    let kind = match kind {
        litchi_xldm::LimitKind::OutputBytes => DataModel::XldmOutputBytes,
        litchi_xldm::LimitKind::PathTextBytes => DataModel::XldmPathTextBytes,
        litchi_xldm::LimitKind::Items => DataModel::XldmProofItems,
        litchi_xldm::LimitKind::StringBytes => DataModel::XldmProofStringBytes,
        litchi_xldm::LimitKind::SourceBytes => DataModel::XldmProofSourceBytes,
        litchi_xldm::LimitKind::Work => DataModel::XldmProofWork,
        litchi_xldm::LimitKind::XldmSourceBytes => DataModel::XldmSourceBytes,
        litchi_xldm::LimitKind::XldmOutputBytes => DataModel::XldmOutputBytes,
        litchi_xldm::LimitKind::DirectoryBytes => DataModel::XldmDirectoryBytes,
        litchi_xldm::LimitKind::Files => DataModel::XldmFiles,
        litchi_xldm::LimitKind::Partitions => DataModel::XldmPartitions,
        litchi_xldm::LimitKind::PathBytes => DataModel::XldmPathBytes,
        litchi_xldm::LimitKind::XmlNodes => DataModel::XldmXmlNodes,
        litchi_xldm::LimitKind::XmlDepth => DataModel::XldmXmlDepth,
        litchi_xldm::LimitKind::XmlTextBytes => DataModel::XldmXmlTextBytes,
        litchi_xldm::LimitKind::MetadataBytes => DataModel::XldmMetadataBytes,
        litchi_xldm::LimitKind::MetadataNodes => DataModel::XldmMetadataNodes,
        litchi_xldm::LimitKind::MetadataDepth => DataModel::XldmMetadataDepth,
        litchi_xldm::LimitKind::MetadataTextBytes => DataModel::XldmMetadataTextBytes,
        litchi_xldm::LimitKind::MetadataFiles => DataModel::XldmMetadataFiles,
        _ => DataModel::PayloadBytes,
    };
    crate::error::LimitKind::DataModel(kind)
}

#[cfg(test)]
pub(crate) mod xldm_test_support;
