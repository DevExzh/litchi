//! Validated XLSB package ownership.
//!
//! This layer owns the OPC package boundary and the package-scoped BIFF12
//! readers and writers. The raw record kernel remains in [`crate::raw`],
//! while workbook and sheet handles live in [`crate::workbook`] and
//! [`crate::sheet`].

use std::io::{Read, Write};
use std::path::Path;
use std::sync::Arc;

use litchi_core::{ExecutionContext, sheet::WorkbookTrait};
use litchi_opc::{OpcPackage, PackageWriter};

use crate::OpenOptions;
use crate::Workbook;
use crate::external_link::ExternalLinkLimits;
use crate::leniency::{Leniency, ToleranceReport};
use crate::package::error::Result;
use crate::workbook::DrawingLoadPolicy;
use crate::writer::{Builder, MutableWorksheet};

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
pub use error::{Error as PackageError, Result as PackageResult};
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
    OpcPackage,
    ExternalLinkLimits,
    ToleranceReport,
    DrawingLoadPolicy,
    Leniency,
    // A validated projection is retained only by constructors that already
    // paid to build it.  The physical package remains the owner snapshot;
    // this cache lets an un-cloned consuming projection reuse that validation
    // without reparsing the same admitted graph.
    Option<Arc<Workbook>>,
);

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

    /// Apply an exact-source structural patch and return a new package.
    pub fn apply(&self, patch: &crate::cell_values::structure::Patch) -> Result<Self> {
        patch.apply_to_package(self)
    }

    pub(crate) const fn drawing_load_policy(&self) -> DrawingLoadPolicy {
        self.3
    }

    pub(crate) fn project_candidate(&self, candidate: OpcPackage) -> Result<Workbook> {
        Workbook::reparse_candidate_with_policy(candidate, self.1, self.3, self.4)
    }

    /// Create a deterministic workbook containing one visible worksheet.
    pub fn new() -> Result<Self> {
        let mut builder = Builder::new();
        builder.add_worksheet(MutableWorksheet::new("Sheet1"));
        builder.build()
    }

    /// Read workbook slicer caches as an immutable, source-bound snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the slicer owner or its dependency closure is invalid.
    pub fn slicer_caches(&self) -> Result<crate::slicer::Snapshot> {
        crate::slicer::transaction::read_caches(&self.0, &self.workbook_uri()?)
    }

    /// Read worksheet slicer views by checked zero-based worksheet position.
    ///
    /// `Ok(None)` means the worksheet selector did not resolve. An existing
    /// worksheet always returns a snapshot, including when it has no view part.
    ///
    /// # Errors
    ///
    /// Returns an error when workbook resolution or the slicer owner is invalid.
    pub fn slicer_views(&self, worksheet_index: usize) -> Result<Option<crate::slicer::Snapshot>> {
        let Some(worksheet_uri) = self.worksheet_uri(worksheet_index)? else {
            return Ok(None);
        };
        crate::slicer::transaction::read_views(&self.0, &worksheet_uri).map(Some)
    }

    /// Read worksheet slicer views by exact worksheet name.
    ///
    /// # Errors
    ///
    /// Returns an error when workbook resolution or the slicer owner is invalid.
    pub fn slicer_views_by_name(
        &self,
        worksheet_name: &str,
    ) -> Result<Option<crate::slicer::Snapshot>> {
        let Some(worksheet_uri) = self.worksheet_uri_by_name(worksheet_name)? else {
            return Ok(None);
        };
        crate::slicer::transaction::read_views(&self.0, &worksheet_uri).map(Some)
    }

    /// Apply a source-checked slicer patch and return a new package snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the patch is stale or its exact readback fails.
    pub fn apply_slicer_patch(&self, patch: &crate::slicer::Patch) -> Result<Self> {
        let mut candidate = self.0.clone();
        crate::slicer::transaction::apply(&mut candidate, patch)?;
        self.reparse_candidate(candidate)
    }

    /// Read workbook timeline caches as an immutable, source-bound snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the timeline owner or its dependency closure is invalid.
    pub fn timeline_caches(&self) -> Result<crate::timeline::Snapshot> {
        crate::timeline::transaction::read_caches(&self.0, &self.workbook_uri()?)
    }

    /// Read worksheet timeline views by checked zero-based worksheet position.
    ///
    /// `Ok(None)` means the worksheet selector did not resolve. An existing
    /// worksheet always returns a snapshot, including when it has no view part.
    ///
    /// # Errors
    ///
    /// Returns an error when workbook resolution or the timeline owner is invalid.
    pub fn timeline_views(
        &self,
        worksheet_index: usize,
    ) -> Result<Option<crate::timeline::Snapshot>> {
        let Some(worksheet_uri) = self.worksheet_uri(worksheet_index)? else {
            return Ok(None);
        };
        crate::timeline::transaction::read_views(&self.0, &worksheet_uri).map(Some)
    }

    /// Read worksheet timeline views by exact worksheet name.
    ///
    /// # Errors
    ///
    /// Returns an error when workbook resolution or the timeline owner is invalid.
    pub fn timeline_views_by_name(
        &self,
        worksheet_name: &str,
    ) -> Result<Option<crate::timeline::Snapshot>> {
        let Some(worksheet_uri) = self.worksheet_uri_by_name(worksheet_name)? else {
            return Ok(None);
        };
        crate::timeline::transaction::read_views(&self.0, &worksheet_uri).map(Some)
    }

    /// Apply a source-checked timeline patch and return a new package snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the patch is stale or its exact readback fails.
    pub fn apply_timeline_patch(&self, patch: &crate::timeline::Patch) -> Result<Self> {
        let mut candidate = self.0.clone();
        crate::timeline::transaction::apply(&mut candidate, patch)?;
        self.reparse_candidate(candidate)
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
        let package = OpcPackage::open_with_limits(path, options.source_limits)?;
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
        let package =
            OpcPackage::from_vec_with_limits(owned_input(bytes, &options)?, options.source_limits)?;
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
        let package = OpcPackage::from_reader_with_limits(reader, options.source_limits)?;
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
        let workbook = Workbook::from_opc_package_with_external_link_limits_and_drawing_policy(
            package.clone(),
            options.external_link_limits,
            drawing_load_policy,
        )?;
        let tolerance = workbook.into_tolerance_report();
        Ok(Self(
            package,
            options.external_link_limits,
            tolerance,
            drawing_load_policy,
            options.leniency,
            None,
        ))
    }

    /// Read the source-bound XLSB Custom Data catalog.
    pub fn custom_data(&self) -> Result<crate::custom_data::Snapshot> {
        crate::custom_data::Snapshot::load(&self.0)
    }

    /// Read Custom Data with an explicit finite policy.
    pub fn custom_data_with_limits(
        &self,
        limits: crate::custom_data::Limits,
    ) -> Result<crate::custom_data::Snapshot> {
        crate::custom_data::Snapshot::load_with_limits(&self.0, limits)
    }

    /// Read Custom Data with an explicit policy and owned execution context.
    pub fn custom_data_with_limits_and_context(
        &self,
        limits: crate::custom_data::Limits,
        context: ExecutionContext,
    ) -> Result<crate::custom_data::Snapshot> {
        crate::custom_data::Snapshot::load_with_limits_and_context(&self.0, limits, Some(context))
    }

    /// Start a detached source-bound Custom Data transaction.
    pub fn edit_custom_data(&self) -> Result<crate::custom_data::Transaction> {
        crate::custom_data::Transaction::from_package(self.0.clone())
    }

    /// Start a detached Custom Data transaction with an explicit finite policy.
    pub fn edit_custom_data_with_limits(
        &self,
        limits: crate::custom_data::Limits,
    ) -> Result<crate::custom_data::Transaction> {
        crate::custom_data::Transaction::from_package_with_limits(self.0.clone(), limits)
    }

    /// Start a detached Custom Data transaction with an owned execution context.
    pub fn edit_custom_data_with_limits_and_context(
        &self,
        limits: crate::custom_data::Limits,
        context: ExecutionContext,
    ) -> Result<crate::custom_data::Transaction> {
        crate::custom_data::Transaction::from_package_with_limits_and_context(
            self.0.clone(),
            limits,
            Some(context),
        )
    }

    /// Apply a source-checked Custom Data patch and return a new package.
    pub fn apply_custom_data_patch(&self, patch: &crate::custom_data::Patch) -> Result<Self> {
        let mut candidate = self.0.clone();
        patch.apply_to_opc(&mut candidate)?;
        self.reparse_candidate(candidate)
    }

    /// Apply a source-checked Custom Data commit and return a new package.
    pub fn apply_custom_data(&self, commit: &crate::custom_data::Commit) -> Result<Self> {
        self.apply_custom_data_patch(commit.patch())
    }

    pub(crate) fn from_workbook(workbook: &Workbook) -> Self {
        Self(
            workbook.package.clone(),
            workbook.external_link_limits(),
            workbook.tolerance_report().clone(),
            workbook.drawing_load_policy(),
            workbook.leniency(),
            None,
        )
    }

    /// Adopt a validated workbook without serializing or reparsing its graph.
    pub(crate) fn from_validated_workbook(workbook: Workbook) -> Self {
        let limits = workbook.external_link_limits();
        let tolerance = workbook.tolerance_report().clone();
        let drawing_policy = workbook.drawing_load_policy();
        let leniency = workbook.leniency();
        Self(
            workbook.package,
            limits,
            tolerance,
            drawing_policy,
            leniency,
            None,
        )
    }

    /// Adopt a physically admitted workbook while retaining its validated
    /// projection for one consuming `into_workbook` call.
    ///
    /// This narrow constructor is reserved for publication paths that have
    /// already parsed and validated the candidate. Ordinary workbook-to-
    /// package conversion intentionally discards the projection so a package
    /// does not retain an unused duplicate semantic graph.
    pub(crate) fn from_validated_workbook_with_projection(workbook: Workbook) -> Self {
        let limits = workbook.external_link_limits();
        let tolerance = workbook.tolerance_report().clone();
        let drawing_policy = workbook.drawing_load_policy();
        let leniency = workbook.leniency();
        let package = workbook.package.clone();
        Self(
            package,
            limits,
            tolerance,
            drawing_policy,
            leniency,
            Some(Arc::new(workbook)),
        )
    }

    /// Revalidate a candidate under the owner's retained projection policies.
    /// [`Self::reparse_candidate`] that keeps the validated projection, so a
    /// following [`Self::into_workbook`] does not parse the candidate again.
    pub(crate) fn reparse_candidate_retaining_projection(
        &self,
        candidate: OpcPackage,
    ) -> Result<Self> {
        let workbook =
            Workbook::reparse_candidate_with_policy(candidate.clone(), self.1, self.3, self.4)?;
        let tolerance = workbook.tolerance_report().clone();
        Ok(Self(
            candidate,
            self.1,
            tolerance,
            self.3,
            self.4,
            Some(Arc::new(workbook)),
        ))
    }

    /// [`Self::reparse_candidate_retaining_projection`] for a candidate
    /// derived from `workbook`'s graph, with the workbook's policies.
    pub(crate) fn reparse_workbook_candidate(
        workbook: &Workbook,
        candidate: OpcPackage,
    ) -> Result<Self> {
        let limits = workbook.external_link_limits();
        let drawing_policy = workbook.drawing_load_policy();
        let leniency = workbook.leniency();
        let reparsed = Workbook::reparse_candidate_with_policy(
            candidate.clone(),
            limits,
            drawing_policy,
            leniency,
        )?;
        let tolerance = reparsed.tolerance_report().clone();
        Ok(Self(
            candidate,
            limits,
            tolerance,
            drawing_policy,
            leniency,
            Some(Arc::new(reparsed)),
        ))
    }

    pub(crate) fn reparse_candidate(&self, candidate: OpcPackage) -> Result<Self> {
        let workbook =
            Workbook::reparse_candidate_with_policy(candidate.clone(), self.1, self.3, self.4)?;
        let tolerance = workbook.into_tolerance_report();
        Ok(Self(candidate, self.1, tolerance, self.3, self.4, None))
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

    /// Materialize a workbook handle from this package.
    pub fn workbook(&self) -> Result<Workbook> {
        let mut workbook = Workbook::from_opc_package_with_external_link_limits_and_drawing_policy(
            self.0.clone(),
            self.1,
            self.3,
        )?;
        workbook.set_projection_leniency(self.4);
        Ok(workbook)
    }

    /// Consume this package and materialize its workbook handle.
    ///
    /// A uniquely held projection retained by cell publication is reused.
    /// Packages without one, including shared projections, are parsed under
    /// the retained source and format policies.
    pub fn into_workbook(self) -> Result<Workbook> {
        let Self(package, external_link_limits, _tolerance, drawing_load_policy, leniency, cached) =
            self;
        if let Some(cached) = cached {
            if let Ok(workbook) = Arc::try_unwrap(cached) {
                return Ok(workbook);
            }
        }
        let mut workbook = Workbook::from_opc_package_with_external_link_limits_and_drawing_policy(
            package,
            external_link_limits,
            drawing_load_policy,
        )?;
        workbook.set_projection_leniency(leniency);
        Ok(workbook)
    }

    /// Consume this package and return its underlying OPC graph.
    pub fn into_opc(self) -> OpcPackage {
        self.0
    }

    /// Serialize the validated package into owned bytes.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        Ok(PackageWriter::to_bytes(&self.0)?)
    }

    /// Stream the validated package into a sequential sink.
    pub fn write_to(&self, writer: impl Write) -> Result<()> {
        Ok(PackageWriter::write_to_stream(writer, &self.0)?)
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

    fn worksheet_uri(&self, index: usize) -> Result<Option<litchi_opc::PackURI>> {
        let workbook = self.workbook()?;
        if index >= workbook.worksheet_names().len() {
            return Ok(None);
        }
        workbook.worksheet_uri(index).map(Some)
    }

    fn worksheet_uri_by_name(&self, name: &str) -> Result<Option<litchi_opc::PackURI>> {
        let workbook = self.workbook()?;
        let Some(index) = workbook
            .worksheet_names()
            .iter()
            .position(|worksheet| worksheet == name)
        else {
            return Ok(None);
        };
        workbook.worksheet_uri(index).map(Some)
    }

    fn workbook_uri(&self) -> Result<litchi_opc::PackURI> {
        Ok(self.0.main_document_part()?.partname().clone())
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
            package,
            ExternalLinkLimits::default(),
            ToleranceReport::default(),
            DrawingLoadPolicy::Eager,
            Leniency::Strict,
            None,
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
        let workbook = Workbook::new_with_options(
            Cursor::new(bytes),
            OpenOptions::new()
                .source_limits(source_limits)
                .external_link_limits(external_limits),
        )
        .expect("bounded workbook");
        let package = workbook.into_package();
        assert_eq!(package.external_link_limits(), external_limits);
        assert_eq!(
            package.opc_package().source_read_limits(),
            Some(source_limits)
        );
        let reconstructed = package.into_workbook().expect("projection");
        assert_eq!(reconstructed.external_link_limits(), external_limits);
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
        let published = package.without_signatures().expect("publication");
        assert_eq!(
            published.into_workbook().expect("projection").leniency(),
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
