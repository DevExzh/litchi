//! Validated XLSB package ownership.
//!
//! This layer owns the OPC package boundary and the package-scoped BIFF12
//! readers and writers. The raw record kernel remains in [`crate::raw`],
//! while workbook and sheet handles live in [`crate::workbook`] and
//! [`crate::sheet`].

use std::io::{Read, Write};
use std::path::Path;
use std::sync::{Arc, OnceLock};

use litchi_opc::{OpcPackage, PackageWriter};

use crate::OpenOptions;
use crate::Workbook;
use crate::builder::{self, Builder};
use crate::external_link::ExternalLinkLimits;
use crate::leniency::{Leniency, ToleranceReport};
use crate::package::error::Result;
use crate::workbook::DrawingLoadPolicy;

#[path = "../host/cell.rs"]
pub(crate) mod cell;
#[path = "../host/cells_reader/mod.rs"]
#[allow(
    dead_code,
    unreachable_pub,
    reason = "internal host modules are shared by included package and writer implementations"
)]
pub(crate) mod cells_reader;
#[path = "../host/chart_resources.rs"]
pub(crate) mod chart_resources;
#[path = "../host/chartsheet/mod.rs"]
pub mod chartsheet;
#[path = "../host/connections/mod.rs"]
pub mod connections;
#[path = "../host/data_validation.rs"]
pub mod data_validation;
#[path = "../host/drawing.rs"]
pub mod drawing;
#[path = "../host/drawing_image.rs"]
pub mod drawing_image;
#[path = "../host/drawing_write.rs"]
pub(crate) mod drawing_write;
#[path = "../host/error.rs"]
pub mod error;
pub(crate) mod external_link;
#[path = "../host/formula/mod.rs"]
pub mod formula;
#[path = "../host/frt.rs"]
pub(crate) mod frt;
#[path = "../host/named_ranges.rs"]
pub mod named_ranges;
pub(crate) mod owner_transaction;
#[path = "../host/pivot/mod.rs"]
pub mod pivot;
#[path = "../host/pivot_tables.rs"]
#[allow(
    dead_code,
    unreachable_pub,
    reason = "internal host modules are shared by included package and writer implementations"
)]
pub(crate) mod pivot_tables;
#[path = "../host/records.rs"]
#[allow(
    dead_code,
    unreachable_pub,
    reason = "internal host modules are shared by included package and writer implementations"
)]
pub(crate) mod records;
#[path = "../host/scenarios/mod.rs"]
pub mod scenarios;
#[path = "../host/shared_strings.rs"]
pub mod shared_strings;
#[path = "../host/sheet_view.rs"]
pub mod sheet_view;
#[path = "../host/styles/mod.rs"]
pub mod styles;
#[path = "../host/styles_table.rs"]
pub mod styles_table;
#[path = "../host/table/mod.rs"]
pub mod table;
#[path = "../host/template.rs"]
pub(crate) mod template;
#[path = "../host/utils.rs"]
#[allow(
    dead_code,
    unreachable_pub,
    reason = "internal host modules are shared by included package and writer implementations"
)]
pub(crate) mod utils;
#[cfg(feature = "vba-inspection")]
#[path = "../host/vba_project.rs"]
pub mod vba_project;
#[path = "../host/walker.rs"]
pub(crate) mod walker;
#[path = "../host/web_extension_bindings.rs"]
pub mod web_extension_bindings;

pub use cell::{Cell, FormulaOpacityReason, FormulaResolutionStatus};
pub use error::{Error as PackageError, LimitKind, Result as PackageResult};
pub use shared_strings::{
    PhoneticAlignment, PhoneticRun, PhoneticString, PhoneticType, SharedString, SharedStringRun,
};

/// A validated physical XLSB package.
///
/// ZIP/archive mechanics are delegated to `litchi-opc`; this type exposes
/// only the XLSB package boundary and validates the workbook graph before a
/// package is handed to callers.
#[derive(Debug, Clone)]
pub struct Package(
    // Shared: a clone of the package (every edit starts from one) is a
    // reference-count increment, and a commit's patch shares both graphs.
    Arc<OpcPackage>,
    ExternalLinkLimits,
    ToleranceReport,
    DrawingLoadPolicy,
    Leniency,
    // The validated read-only workbook projection of exactly this OPC graph.
    // Constructors that already parsed it (every validating open and every
    // reparsing commit) fill it; otherwise `workbook()` parses it once on
    // first use. Clones share it.
    Projection,
    // The semantic budgets every family read and edit defaults to.
    crate::SemanticLimits,
    // The raw BIFF12 policy every detached codec reached by this owner uses.
    crate::raw::Limits,
);

/// The lazily filled, shared workbook projection of a [`Package`].
#[derive(Debug, Clone, Default)]
pub(crate) struct Projection(OnceLock<Arc<Workbook>>);

impl Projection {
    fn filled(workbook: Workbook) -> Self {
        Self(OnceLock::from(Arc::new(workbook)))
    }
}

impl Package {
    /// Start the primary detached workbook transaction.
    pub fn edit(&self) -> Result<crate::cell_values::structure::Edit> {
        self.edit_with_options(crate::cell_values::structure::EditOptions::default())
    }

    /// Start the primary transaction with explicit finite transfer and patch
    /// limits.
    pub fn edit_with_options(
        &self,
        options: crate::cell_values::structure::EditOptions,
    ) -> Result<crate::cell_values::structure::Edit> {
        crate::cell_values::structure::Edit::new(self, options.get_transfer_limits())
    }

    /// Apply an exact-source structural patch to a new package; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale patch source, a changed patch of a
    /// signed workbook, and the workbook validation errors of the result.
    pub fn apply(
        &self,
        patch: &crate::cell_values::structure::Patch,
    ) -> Result<crate::cell_values::structure::Commit> {
        Ok(crate::cell_values::structure::Commit::new(
            patch.apply_to_package(self)?,
            patch.clone(),
            litchi_core::edit::Diagnostics::new(!patch.is_noop()),
        ))
    }

    pub(crate) const fn drawing_load_policy(&self) -> DrawingLoadPolicy {
        self.3
    }

    pub(crate) fn project_candidate(&self, candidate: OpcPackage) -> Result<Workbook> {
        let mut workbook =
            Workbook::reparse_candidate_with_policy(candidate, self.1, self.3, self.4, self.7)?;
        workbook.set_semantic_limits(self.6);
        Ok(workbook)
    }

    /// The semantic budgets this package keeps; every family read uses them
    /// and every family edit defaults to them.
    #[must_use]
    pub const fn semantic_limits(&self) -> crate::SemanticLimits {
        self.6
    }

    /// The BIFF12 raw policy retained by this package owner.
    #[must_use]
    pub const fn raw_limits_value(&self) -> crate::raw::Limits {
        self.7
    }

    /// Return this package with other semantic budgets; the OPC graph and
    /// the workbook projection are shared.
    #[must_use]
    pub fn with_semantic_limits(&self, limits: crate::SemanticLimits) -> Self {
        let mut package = self.clone();
        package.6 = limits;
        if let Some(workbook) = package.5.0.get() {
            if workbook.semantic_limits() != limits {
                package.5 = Projection::default();
            }
        }
        package
    }

    /// Create a deterministic workbook containing one visible worksheet.
    pub fn new() -> Result<Self> {
        let mut builder = Builder::new();
        builder.add_worksheet(builder::Worksheet::new("Sheet1"));
        builder.build()
    }

    /// Open and validate an XLSB package from a filesystem path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_options(path, OpenOptions::default())
    }

    /// Open and validate an XLSB package from a filesystem path under
    /// explicit ingress policies.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::UnsupportedOpenPolicy`] for a cache or
    /// execution policy (positional-source only), the OPC limit and archive
    /// errors, and the workbook validation errors.
    pub fn open_with_options(path: impl AsRef<Path>, options: OpenOptions) -> Result<Self> {
        options.check_owning()?;
        let package = OpcPackage::open(path, options.source_limits)?;
        Self::from_opc_with_options(package, options)
    }

    /// Read and validate an XLSB package from a borrowed byte slice.
    ///
    /// The bytes are copied once into the package's owned source archive, so
    /// unchanged members (and a signed source's exact bytes) keep exact-source
    /// publication.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_options(bytes, OpenOptions::default())
    }

    /// Read and validate an XLSB package from a borrowed byte slice under
    /// explicit ingress policies.
    ///
    /// The input-size limit is checked before the bytes are copied.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::open_with_options`].
    pub fn from_bytes_with_options(bytes: &[u8], options: OpenOptions) -> Result<Self> {
        options.check_owning()?;
        let package = OpcPackage::from_vec(owned_input(bytes, &options)?, options.source_limits)?;
        Self::from_opc_with_options(package, options)
    }

    /// Read and validate an XLSB package from a synchronous reader.
    pub fn from_reader(reader: impl Read) -> Result<Self> {
        Self::from_reader_with_options(reader, OpenOptions::default())
    }

    /// Read and validate an XLSB package from a synchronous reader under
    /// explicit ingress policies.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::open_with_options`].
    pub fn from_reader_with_options(reader: impl Read, options: OpenOptions) -> Result<Self> {
        options.check_owning()?;
        let package = OpcPackage::from_reader(reader, options.source_limits)?;
        Self::from_opc_with_options(package, options)
    }

    /// Validate and adopt an already parsed OPC package.
    pub fn from_opc(package: OpcPackage) -> Result<Self> {
        Self::from_opc_with_options(package, OpenOptions::default())
    }

    /// Validate and adopt an already parsed OPC package under explicit
    /// semantic policies.
    ///
    /// The package was read under its own source limits, so only the
    /// external-link limits, leniency and drawing policy apply here.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::UnsupportedOpenPolicy`] for a cache or
    /// execution policy, and the workbook validation errors.
    pub fn from_opc_with_options(package: OpcPackage, options: OpenOptions) -> Result<Self> {
        options.check_owning()?;
        let drawing_load_policy = options.drawing_load_policy();
        let mut workbook = Workbook::from_opc_package_with_external_link_limits_and_drawing_policy(
            package.clone(),
            options.external_link_limits,
            drawing_load_policy,
            options.raw_limits,
        )?;
        workbook.set_projection_leniency(options.leniency);
        workbook.set_semantic_limits(options.semantic_limits);
        let tolerance = workbook.tolerance_report().clone();
        Ok(Self(
            Arc::new(package),
            options.external_link_limits,
            tolerance,
            drawing_load_policy,
            options.leniency,
            Projection::filled(workbook),
            options.semantic_limits,
            options.raw_limits,
        ))
    }

    pub(crate) fn from_workbook(workbook: &Workbook) -> Self {
        Self(
            Arc::new(workbook.package.clone()),
            workbook.external_link_limits(),
            workbook.tolerance_report().clone(),
            workbook.drawing_load_policy(),
            workbook.leniency(),
            Projection::default(),
            workbook.semantic_limits(),
            workbook.raw_limits_value(),
        )
    }

    /// Adopt a validated workbook without serializing or reparsing its graph;
    /// the workbook becomes the package's projection.
    pub(crate) fn from_validated_workbook(workbook: Workbook) -> Self {
        let limits = workbook.external_link_limits();
        let tolerance = workbook.tolerance_report().clone();
        let drawing_policy = workbook.drawing_load_policy();
        let leniency = workbook.leniency();
        let semantic_limits = workbook.semantic_limits();
        let raw_limits = workbook.raw_limits_value();
        let package = workbook.package.clone();
        Self(
            Arc::new(package),
            limits,
            tolerance,
            drawing_policy,
            leniency,
            Projection::filled(workbook),
            semantic_limits,
            raw_limits,
        )
    }

    /// Revalidate a candidate under the owner's retained projection policies,
    /// keeping the validated projection so [`Self::workbook`] does not parse
    /// the candidate again.
    pub(crate) fn reparse_candidate_retaining_projection(
        &self,
        candidate: impl Into<Arc<OpcPackage>>,
    ) -> Result<Self> {
        let candidate = candidate.into();
        let workbook = self.project_candidate((*candidate).clone())?;
        let tolerance = workbook.tolerance_report().clone();
        Ok(Self(
            candidate,
            self.1,
            tolerance,
            self.3,
            self.4,
            Projection::filled(workbook),
            self.6,
            self.7,
        ))
    }

    pub(crate) fn reparse_candidate(&self, candidate: OpcPackage) -> Result<Self> {
        self.reparse_candidate_retaining_projection(candidate)
    }

    /// Publish a candidate whose workbook projection the change cannot
    /// affect (a signature, ribbon, task-pane, VBA or theme part change)
    /// without validating it again. The projection is parsed on first use.
    pub(crate) fn with_unparsed(&self, candidate: impl Into<Arc<OpcPackage>>) -> Self {
        Self(
            candidate.into(),
            self.1,
            self.2.clone(),
            self.3,
            self.4,
            Projection::default(),
            self.6,
            self.7,
        )
    }

    /// The worksheet repair policy retained by this package.
    #[must_use]
    pub const fn leniency(&self) -> Leniency {
        self.4
    }

    /// Whether workbook views decode drawings rather than retaining them as opaque parts.
    #[must_use]
    pub fn drawings_parsed(&self) -> bool {
        self.3 == DrawingLoadPolicy::Eager
    }

    /// Every repair the validating open applied while admitting this package.
    ///
    /// This is the [`Workbook::tolerance_report`] of the workbook graph that
    /// was validated; see [`crate::leniency`] for the repair contract. A
    /// package adopted without validation through `From<OpcPackage>` reports
    /// clean because nothing was parsed.
    #[must_use]
    pub fn tolerance_report(&self) -> &ToleranceReport {
        &self.2
    }

    /// Borrow the validated OPC package for lower-level package services.
    pub fn opc_package(&self) -> &OpcPackage {
        &self.0
    }

    /// Return the external-link resource policy retained for workbook
    /// reconstructions from this package.
    #[must_use]
    pub const fn external_link_limits(&self) -> ExternalLinkLimits {
        self.1
    }

    /// The read-only workbook projection of this package.
    ///
    /// A validating open and commits whose dependency closure changed already
    /// parsed it, so this borrows that projection. A dependency-disjoint raw
    /// commit, or a package adopted without validation (`From<OpcPackage>`),
    /// parses it once on first use. Clones of a package share it.
    ///
    /// # Errors
    ///
    /// Returns the workbook validation errors of a package whose projection
    /// was not parsed yet.
    pub fn workbook(&self) -> Result<&Workbook> {
        if let Some(workbook) = self.5.0.get() {
            return Ok(workbook);
        }
        let workbook = self.materialize_workbook()?;
        let workbook = self.5.0.get_or_init(|| Arc::new(workbook));
        Ok(workbook)
    }

    #[cfg(test)]
    pub(crate) fn projection_cached(&self) -> bool {
        self.5.0.get().is_some()
    }

    /// The shared workbook projection, parsed on first use.
    pub(crate) fn projection(&self) -> Result<Arc<Workbook>> {
        self.workbook()?;
        self.5.0.get().cloned().ok_or(PackageError::InvalidFormat(
            "the XLSB workbook projection is missing",
        ))
    }

    /// Whether this package carries package signatures.
    #[must_use]
    pub fn is_signed(&self) -> bool {
        self.0.is_signed()
    }

    /// Verify package signatures with the safe strict policy.
    ///
    /// # Errors
    ///
    /// Returns the signature parsing and verification errors.
    #[cfg(feature = "sign")]
    pub fn signatures(&self) -> litchi_opc::sign::Result<Vec<litchi_opc::sign::Report>> {
        self.0.signatures()
    }

    /// Verify package signatures with an explicit trust-neutral policy.
    ///
    /// # Errors
    ///
    /// Returns the signature parsing and verification errors.
    #[cfg(feature = "sign")]
    pub fn signatures_with(
        &self,
        policy: &litchi_sign::Policy,
    ) -> litchi_opc::sign::Result<Vec<litchi_opc::sign::Report>> {
        self.0.signatures_with(policy)
    }

    /// Return a new package with one more signature, preserving every
    /// existing valid signature; this package is unchanged. The new
    /// signature part is reported by [`Self::signatures`].
    ///
    /// # Errors
    ///
    /// Returns the signing errors.
    #[cfg(feature = "sign")]
    pub fn sign(&self, signer: &litchi_sign::Signer) -> litchi_opc::sign::Result<Self> {
        let mut candidate = (*self.0).clone();
        candidate.sign(signer)?;
        Ok(self.with_unparsed(candidate))
    }

    /// [`Self::sign`] with explicit authoring resource bounds.
    ///
    /// # Errors
    ///
    /// Returns the signing errors.
    #[cfg(feature = "sign")]
    pub fn sign_with(
        &self,
        signer: &litchi_sign::Signer,
        limits: &litchi_sign::Limits,
    ) -> litchi_opc::sign::Result<Self> {
        let mut candidate = (*self.0).clone();
        candidate.sign_with(signer, limits)?;
        Ok(self.with_unparsed(candidate))
    }

    /// Return a new package whose signatures are atomically replaced by one
    /// signature; this package is unchanged.
    ///
    /// # Errors
    ///
    /// Returns the signing errors.
    #[cfg(feature = "sign")]
    pub fn resign(&self, signer: &litchi_sign::Signer) -> litchi_opc::sign::Result<Self> {
        let mut candidate = (*self.0).clone();
        candidate.resign(signer)?;
        Ok(self.with_unparsed(candidate))
    }

    /// [`Self::resign`] with explicit authoring resource bounds.
    ///
    /// # Errors
    ///
    /// Returns the signing errors.
    #[cfg(feature = "sign")]
    pub fn resign_with(
        &self,
        signer: &litchi_sign::Signer,
        limits: &litchi_sign::Limits,
    ) -> litchi_opc::sign::Result<Self> {
        let mut candidate = (*self.0).clone();
        candidate.resign_with(signer, limits)?;
        Ok(self.with_unparsed(candidate))
    }

    /// Return a new package without its package signatures; this package
    /// keeps them.
    ///
    /// This is the explicit authorization ADR 0006 requires: an edit never
    /// removes a signature itself, so a changed commit of a signed package
    /// is refused with [`PackageError::Signed`] until this is called.
    #[must_use]
    pub fn unsign(&self) -> Self {
        let mut candidate = (*self.0).clone();
        candidate.unsign();
        self.with_unparsed(candidate)
    }

    /// Parse a fresh, owned workbook from this package under its retained
    /// policies (for crate paths that author into a donor workbook).
    pub(crate) fn materialize_workbook(&self) -> Result<Workbook> {
        let mut workbook = Workbook::from_opc_package_with_external_link_limits_and_drawing_policy(
            (*self.0).clone(),
            self.1,
            self.3,
            self.7,
        )?;
        workbook.set_projection_leniency(self.4);
        workbook.set_semantic_limits(self.6);
        Ok(workbook)
    }

    /// Consume this package and return its underlying OPC graph.
    pub fn into_opc(self) -> OpcPackage {
        Arc::try_unwrap(self.0).unwrap_or_else(|shared| (*shared).clone())
    }

    /// The shared OPC graph, for patches that keep it.
    pub(crate) fn opc_shared(&self) -> &Arc<OpcPackage> {
        &self.0
    }

    /// Serialize the validated package into owned bytes.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        Ok(PackageWriter::to_bytes(&self.0)?)
    }

    /// Stream the validated package into a sequential sink.
    pub fn write_to(&self, writer: impl Write) -> Result<()> {
        Ok(PackageWriter::write_to(writer, &self.0)?)
    }

    /// Atomically save the validated package to a filesystem path.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.save_with_durability(path, litchi_core::Durability::Full)
    }

    /// Atomically save the validated package to a filesystem path at a
    /// caller-chosen [`Durability`](litchi_core::Durability).
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
        Ok(PackageWriter::write_with_durability(
            path, &self.0, durability,
        )?)
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
        .map_err(|source| PackageError::Allocation {
            resource: "XLSB package input bytes",
            source,
        })?;
    owned.extend_from_slice(bytes);
    Ok(owned)
}

impl From<Package> for OpcPackage {
    fn from(package: Package) -> Self {
        package.into_opc()
    }
}

impl From<OpcPackage> for Package {
    fn from(package: OpcPackage) -> Self {
        Self(
            Arc::new(package),
            ExternalLinkLimits::default(),
            ToleranceReport::default(),
            DrawingLoadPolicy::Eager,
            Leniency::Strict,
            Projection::default(),
            crate::SemanticLimits::new(),
            crate::raw::Limits::DEFAULT,
        )
    }
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    reason = "package ingress fixtures use panic-on-failure extraction after constructing exact finite limits"
)]
mod tests {
    use std::io::Cursor;

    use litchi_opc::SourceLimits;

    use super::*;

    #[test]
    fn workbook_adoption_retains_physical_and_external_link_limits() {
        let bytes = Package::new().expect("create").to_bytes().expect("bytes");
        let source_limits = SourceLimits::builder()
            .max_input_bytes(u64::try_from(bytes.len()).expect("length"))
            .expect("input limit")
            .build()
            .expect("source limits");
        let external_limits = ExternalLinkLimits::builder()
            .max_links(1)
            .build()
            .expect("external limits");
        let raw_limits = crate::raw::Limits::new(4096, 1024);
        let workbook = Workbook::new_with_options(
            Cursor::new(bytes),
            OpenOptions::new()
                .source_limits(source_limits)
                .external_link_limits(external_limits)
                .raw_limits(raw_limits),
        )
        .expect("bounded workbook");
        let package = workbook.into_package();
        assert_eq!(package.external_link_limits(), external_limits);
        assert_eq!(package.raw_limits_value(), raw_limits);
        assert_eq!(
            package.opc_package().source_read_limits(),
            Some(source_limits)
        );
        let reconstructed = package.materialize_workbook().expect("projection");
        assert_eq!(reconstructed.external_link_limits(), external_limits);
        assert_eq!(reconstructed.raw_limits_value(), raw_limits);
        assert_eq!(
            reconstructed.opc_package().source_read_limits(),
            Some(source_limits)
        );
    }

    #[test]
    fn workbook_adoption_retains_leniency_without_changing_default_ingress() {
        let bytes = Package::new().expect("create").to_bytes().expect("bytes");
        let workbook = Workbook::new_with_options(
            Cursor::new(bytes.clone()),
            OpenOptions::new().leniency(Leniency::TolerateCachedNumbers),
        )
        .expect("opt-in workbook");
        let tolerance = workbook.tolerance_report().clone();
        let package = workbook.into_package();
        assert_eq!(package.leniency(), Leniency::TolerateCachedNumbers);
        assert_eq!(package.tolerance_report(), &tolerance);
        assert_eq!(
            package.workbook().expect("projection").leniency(),
            Leniency::TolerateCachedNumbers
        );
        let published = package.unsign();
        assert_eq!(
            published
                .materialize_workbook()
                .expect("projection")
                .leniency(),
            Leniency::TolerateCachedNumbers
        );
        assert_eq!(
            Package::from_bytes(&bytes)
                .expect("default ingress")
                .leniency(),
            Leniency::Strict
        );
    }

    #[test]
    fn ingress_honors_exact_read_limits() {
        let bytes = Package::new()
            .expect("create package")
            .to_bytes()
            .expect("serialize package");
        let input_bytes = u64::try_from(bytes.len()).expect("input length fits u64");
        let exact = SourceLimits::builder()
            .max_input_bytes(input_bytes)
            .expect("exact input limit")
            .build()
            .expect("valid exact limit");
        let over = SourceLimits::builder()
            .max_input_bytes(input_bytes - 1)
            .expect("smaller input limit")
            .build()
            .expect("valid smaller limit");
        let file = tempfile::NamedTempFile::new().expect("temporary workbook path");
        std::fs::write(file.path(), &bytes).expect("write workbook");

        assert!(Package::open(file.path()).is_ok());
        assert!(Package::from_bytes(&bytes).is_ok());
        assert!(Package::from_bytes(&bytes).is_ok());
        assert!(Package::from_reader(Cursor::new(bytes.clone())).is_ok());

        assert!(
            Package::open_with_options(file.path(), OpenOptions::new().source_limits(exact))
                .is_ok()
        );
        assert!(
            Package::from_bytes_with_options(&bytes, OpenOptions::new().source_limits(exact))
                .is_ok()
        );
        assert!(
            Package::from_bytes_with_options(&bytes, OpenOptions::new().source_limits(exact))
                .is_ok()
        );
        assert!(
            Package::from_reader_with_options(
                Cursor::new(bytes.clone()),
                OpenOptions::new().source_limits(exact)
            )
            .is_ok()
        );

        assert!(
            Package::open_with_options(file.path(), OpenOptions::new().source_limits(over))
                .is_err()
        );
        assert!(
            Package::from_bytes_with_options(&bytes, OpenOptions::new().source_limits(over))
                .is_err()
        );
        assert!(
            Package::from_bytes_with_options(&bytes, OpenOptions::new().source_limits(over))
                .is_err()
        );
        assert!(
            Package::from_reader_with_options(
                Cursor::new(bytes),
                OpenOptions::new().source_limits(over)
            )
            .is_err()
        );
    }
}
