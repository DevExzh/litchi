//! Concise user-facing ODS entry points.

mod cell_locator;
mod selective;
mod source;
mod source_edit;

use litchi_core::{
    Result, SequentialTextWriter, TextOutputError, TextOutputOptions, TextOutputReport,
};
use std::{
    io::Write,
    path::Path,
    sync::{
        Arc, OnceLock,
        atomic::{AtomicUsize, Ordering},
    },
};

pub use crate::authoring::{Builder, MutableSpreadsheet};
use crate::model::names::{Definition, Expression, Range, Scope};
pub use litchi_odf_common::rdf::{Graph, Object, Subject, Triple};
pub use selective::{SheetCatalogEntry, SourceBackedSpreadsheetCatalog, SourceReadMetrics};
pub use source::{ReadLimits, SourceBackedSpreadsheet};
pub use source_edit::{
    FormulaChange, HyperlinkChange, HyperlinkOperation, SourceCellCommit, SourceCellEdit,
    SourceCellPatch, SourceCellPublicationReport, SourceCellSnapshot,
};

/// Maximum number of positional cell selectors accepted by one lookup batch.
///
/// The bound keeps result-vector allocation and lookup work finite even when
/// selectors originate outside the parsed document.  A batch at the bound is
/// accepted; larger batches fail before any cell lookup or index construction.
pub const MAX_CELL_SELECTORS: usize = 4_096;

/// A reusable selector for one logical ODS cell.
///
/// The sheet name is borrowed so callers can build selector arrays without
/// copying names.  Rows and columns are zero-based logical coordinates; ODF
/// repeated rows and cells remain represented by their physical owners in the
/// returned [`crate::worksheet::CellView`].
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct CellSelector<'a> {
    sheet_name: &'a str,
    row: usize,
    column: usize,
}

impl<'a> CellSelector<'a> {
    /// Construct a selector for one zero-based logical cell.
    #[must_use]
    pub const fn new(sheet_name: &'a str, row: usize, column: usize) -> Self {
        Self {
            sheet_name,
            row,
            column,
        }
    }

    /// Return the exact worksheet name selected by this value.
    #[must_use]
    pub const fn sheet_name(self) -> &'a str {
        self.sheet_name
    }

    /// Return the zero-based logical row selected by this value.
    #[must_use]
    pub const fn row(self) -> usize {
        self.row
    }

    /// Return the zero-based logical column selected by this value.
    #[must_use]
    pub const fn column(self) -> usize {
        self.column
    }
}

impl<'a> From<(&'a str, usize, usize)> for CellSelector<'a> {
    fn from((sheet_name, row, column): (&'a str, usize, usize)) -> Self {
        Self::new(sheet_name, row, column)
    }
}

fn validate_cell_batch_len(length: usize) -> Result<()> {
    if length > MAX_CELL_SELECTORS {
        return Err(litchi_core::Error::InvalidFormat(format!(
            "ODS cell lookup batch exceeds the {MAX_CELL_SELECTORS} selector safety limit"
        )));
    }
    Ok(())
}

fn map_sheet_metadata_execution(error: litchi_core::ExecutionError) -> litchi_core::Error {
    match error {
        litchi_core::ExecutionError::ResourceLimit(value) => {
            litchi_core::Error::ResourceLimit(value)
        },
        litchi_core::ExecutionError::Cancelled => {
            litchi_core::Error::Unsupported("ODS metadata operation cancelled".to_string())
        },
        other => litchi_core::Error::Unsupported(format!(
            "ODS metadata execution policy rejected operation: {other}"
        )),
    }
}

/// Immutable ODS document facade.
pub struct Spreadsheet {
    package: Arc<crate::package::Package>,
    definitions: Vec<Definition>,
    sheets: Vec<crate::worksheet::Sheet>,
    metadata: crate::metadata::Snapshot,
    settings: Option<crate::settings::Settings>,
    cell_queries: AtomicUsize,
    cell_locator: OnceLock<Option<cell_locator::CellLocator>>,
}

impl Spreadsheet {
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let package = crate::package::Package::open(path)?;
        Self::from_package(package)
    }

    /// Open a password-encrypted ODS path and fully decode the public semantic owners.
    ///
    /// # Errors
    ///
    /// Returns an error for file I/O, an incorrect password, malformed encryption metadata,
    /// invalid XML, or typed owner readback failure.
    pub fn open_with_password(path: impl AsRef<Path>, password: impl Into<String>) -> Result<Self> {
        let package = crate::package::Package::open_with_password(path, password)?;
        Self::from_package(package)
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        // Let the prepared detector perform the local MIME probe exactly
        // once.  A matching ODS result transfers its indexed archive; a
        // different ODF family still follows the historical package owner so
        // its MIME error is reported at the same boundary; rejected probes
        // recover the original allocation for the ordinary parser.
        match litchi_odf_common::detect::prepared_or_original(bytes) {
            Ok(prepared) if prepared.format() == litchi_core::detection::FileFormat::Ods => {
                Self::from_prepared_package(prepared)
            },
            Ok(prepared) => {
                let package = crate::package::Package::from_owned_package(prepared.into_package())?;
                Self::from_package(package)
            },
            Err(bytes) => {
                let package = crate::package::Package::from_bytes(bytes)?;
                Self::from_package(package)
            },
        }
    }

    /// Adopt the indexed package retained by smart ODF detection.
    ///
    /// This transfers the detector-owned archive index into the immutable
    /// spreadsheet package without a second ZIP central-directory scan.
    pub fn from_prepared_package(
        prepared: litchi_odf_common::core::PreparedPackage,
    ) -> Result<Self> {
        let package = crate::package::Package::from_prepared_package(prepared)?;
        Self::from_package(package)
    }

    /// Alias for [`Self::from_prepared_package`].
    #[inline]
    pub fn from_prepared(prepared: litchi_odf_common::core::PreparedPackage) -> Result<Self> {
        Self::from_prepared_package(prepared)
    }

    /// Return the identity of the archive index retained by smart detection.
    #[doc(hidden)]
    #[must_use]
    pub fn prepared_index_identity(&self) -> usize {
        self.package.prepared_index_identity()
    }

    /// Open password-encrypted ODS bytes and fully decode the public semantic owners.
    ///
    /// # Errors
    ///
    /// Returns an error for an incorrect password, malformed encryption metadata, invalid XML,
    /// or typed owner readback failure.
    pub fn from_bytes_with_password(bytes: Vec<u8>, password: impl Into<String>) -> Result<Self> {
        let package = crate::package::Package::from_bytes_with_password(bytes, password)?;
        Self::from_package(package)
    }

    pub(crate) fn from_package(package: crate::package::Package) -> Result<Self> {
        Self::from_shared_package(Arc::new(package))
    }

    pub(crate) fn from_owned_package(
        package: litchi_odf_common::core::OwnedPackage,
    ) -> Result<Self> {
        Self::from_package(crate::package::Package::from_owned_package(package)?)
    }

    pub(crate) fn from_shared_package(package: Arc<crate::package::Package>) -> Result<Self> {
        // Keep ordinary and source-backed opening on the same namespace-aware
        // content reader.  Besides avoiding five independent tokenizations,
        // this admits semantically normalized namespace declarations (for
        // example an entity-escaped `xmlns:xml` URI) without changing the
        // source bytes retained by the package.
        let (settings, outputs) =
            crate::open_parse::OpenParse::run(package.content_xml())?.finish()?;
        let definitions = outputs.definitions;
        let mut sheets = outputs.sheets;
        // `OpenParse` owns the fused worksheet pass, while table title and
        // description are deliberately applied in the small metadata pass
        // shared with the source-backed facade.  Keep this post-pass here so
        // ordinary and source-backed opens expose the same Sheet metadata.
        crate::worksheet::codec::apply_table_metadata(package.content_xml(), &mut sheets)?;
        let metadata = package.metadata_snapshot()?;
        Ok(Self {
            package,
            definitions,
            sheets,
            metadata,
            settings,
            cell_queries: AtomicUsize::new(0),
            cell_locator: OnceLock::new(),
        })
    }

    /// Capture the exact package as the unified immutable transaction owner.
    ///
    /// # Errors
    ///
    /// Returns an error when package bounds or complete facade readback fail.
    pub fn document_snapshot(&self) -> Result<crate::document::Snapshot> {
        if self
            .package
            .package()
            .package()?
            .manifest()
            .has_encrypted_entries()
        {
            return crate::document::Snapshot::from_shared_package(
                Arc::clone(&self.package),
                crate::document::Limits::default(),
            );
        }
        let package = self.package.clone_without_password()?;
        crate::document::Snapshot::from_shared_package(
            Arc::new(package),
            crate::document::Limits::default(),
        )
    }

    /// Apply one durable exact-source unified package patch.
    ///
    /// # Errors
    ///
    /// Returns an error for stale lineage, security refusal, package bounds, or candidate
    /// readback failure. This facade changes only after the target fully reopens.
    pub fn apply_document_patch(&mut self, patch: &crate::document::Patch) -> Result<()> {
        let commit = patch.apply(&self.document_snapshot()?)?;
        if commit.changed() {
            *self = Self::from_bytes(commit.snapshot().as_bytes().to_vec())?;
        }
        Ok(())
    }

    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.package.content_xml()
    }

    /// Return worksheet names in document order.
    #[must_use]
    pub fn sheet_names(&self) -> Vec<String> {
        self.sheets.iter().map(|sheet| sheet.name.clone()).collect()
    }

    /// Return the number of worksheets.
    #[must_use]
    pub fn sheet_count(&self) -> usize {
        self.sheets.len()
    }

    /// Extract displayed worksheet text using tab-separated cells and
    /// newline-separated rows, preserving sheet order.
    ///
    /// # Errors
    /// Returns an error when the bounded projection cannot reserve its output.
    pub fn text(&self) -> Result<String> {
        source::project_text(&self.sheets)
    }

    /// Write displayed worksheet text to a caller-owned sequential sink.
    ///
    /// Each logical worksheet row is a paragraph-like semantic object. Cells
    /// within a row are separated by tabs, repeated cells and rows are
    /// expanded in document order, and an empty worksheet is represented by
    /// an empty object so the default output retains the existing `text()`
    /// newline projection. The shared output policy controls separators,
    /// empty-object handling, and caller-selected resource limits.
    ///
    /// The conversion does not construct the document-wide `String` returned
    /// by [`Self::text`]. The existing 64 MiB ODS projected-text safety ceiling
    /// is charged incrementally before each row object is emitted. Output may
    /// be partial when a document limit, output limit, or sink failure is
    /// observed. This method never flushes or rolls back the caller-owned sink
    /// and has no cancellation hook.
    pub fn write_text_to<W: Write + ?Sized>(
        &self,
        output: &mut W,
        options: TextOutputOptions<'_>,
    ) -> std::result::Result<TextOutputReport, TextOutputError<litchi_core::Error>> {
        let mut writer = SequentialTextWriter::new(output, options);
        source::write_text_to_sheets(&self.sheets, &mut writer, options)?;
        Ok(writer.finish())
    }

    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.package.styles_xml()
    }

    /// Capture the standalone table-template catalog from styles.xml.
    ///
    /// The snapshot retains the exact styles source for no-op, stale-source,
    /// and inverse transaction checks.
    pub fn table_templates(&self) -> Result<crate::styles::table_template::Snapshot> {
        crate::styles::table_template::Snapshot::from_source(self.package.styles_xml())
    }

    /// Apply an exact-source table-template patch and rehydrate the facade.
    pub fn apply_table_template_patch(
        &mut self,
        patch: &crate::styles::table_template::Patch,
    ) -> Result<()> {
        let snapshot = self.table_templates()?;
        let commit = patch.apply(&snapshot)?;
        if commit.changed() {
            let package = self
                .package
                .replace_styles_xml(commit.snapshot().source_xml())?;
            *self = Self::from_package(package)?;
        }
        Ok(())
    }

    /// Stage a source-checked table-template edit and publish it atomically.
    pub fn edit_table_templates<F>(&mut self, update: F) -> Result<()>
    where
        F: FnOnce(&mut crate::styles::table_template::Edit) -> Result<()>,
    {
        let snapshot = self.table_templates()?;
        let mut edit = snapshot.edit();
        update(&mut edit)?;
        let commit = edit.commit()?;
        if commit.changed() {
            let package = self
                .package
                .replace_styles_xml(commit.snapshot().source_xml())?;
            *self = Self::from_package(package)?;
        }
        Ok(())
    }

    /// Capture document, sheet, and automatic cell-protection metadata in a
    /// source-checked immutable snapshot.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn protection(&self) -> Result<crate::protection::Snapshot> {
        crate::protection::Snapshot::parse(self.package.content_xml(), self.package.styles_xml())
    }

    /// Apply an exact-source reversible protection patch and fully rehydrate
    /// this facade only after the candidate has passed its typed readback.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn apply_protection_patch(&mut self, patch: &crate::protection::Patch) -> Result<()> {
        let commit = patch.apply(&self.protection()?)?;
        if commit.changed() {
            let package = self.package.replace_content_xml(commit.content_xml())?;
            *self = Self::from_package(package)?;
        }
        Ok(())
    }

    /// Apply a failure-atomic protection edit and rebuild only `content.xml`.
    /// Password values remain inert verifiers; this method never authenticates
    /// or enforces a protection policy.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn update_protection<F>(&mut self, edit: F) -> Result<()>
    where
        F: FnOnce(&mut crate::protection::Transaction) -> Result<()>,
    {
        let snapshot = self.protection()?;
        let commit = crate::protection::update(&snapshot, edit)?;
        if !commit.changed() {
            return Ok(());
        }
        let package = self.package.replace_content_xml(commit.content_xml())?;
        *self = Self::from_package(package)?;
        Ok(())
    }

    /// Capture the source-checked cell-annotation owner for this spreadsheet.
    ///
    /// The owner retains the exact `content.xml` source and resolves cells by
    /// sheet name plus zero-based logical coordinates.  It is parsed on
    /// demand so an immutable spreadsheet does not retain a second XML copy.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn annotations(&self) -> Result<crate::annotations::Snapshot> {
        crate::annotations::Snapshot::parse(self.package.content_xml())
    }

    /// Capture the presence-aware, exact-source tracked-change owner.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn tracked_changes(&self) -> Result<crate::tracked_changes::Snapshot> {
        crate::tracked_changes::Snapshot::parse(self.package.content_xml())
    }

    /// Capture tracked changes under an explicit resource budget.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn tracked_changes_with(
        &self,
        limits: crate::tracked_changes::Limits,
    ) -> Result<crate::tracked_changes::Snapshot> {
        crate::tracked_changes::Snapshot::parse_with_limits(self.package.content_xml(), limits)
    }

    /// Inspect all DDE declarations and cached tables as inert, source-bound data.
    ///
    /// This method never starts a DDE conversation, refreshes a cache, opens a
    /// linked document, or performs ambient I/O.
    ///
    /// # Errors
    ///
    /// Returns an error when the content XML has invalid or over-budget DDE
    /// metadata.
    pub fn dde(&self) -> Result<crate::dde::Snapshot> {
        crate::dde::Snapshot::parse(self.package.content_xml()).map_err(Into::into)
    }

    /// Capture inert DDE declarations and cached tables under explicit limits.
    pub fn dde_with(
        &self,
        limits: crate::dde::Limits,
        context: &litchi_core::ExecutionContext,
    ) -> Result<crate::dde::Snapshot> {
        crate::dde::Snapshot::parse_with_context(self.package.content_xml(), limits, context)
            .map_err(Into::into)
    }

    /// Stage and publish a failure-atomic DDE metadata transaction.
    ///
    /// Sources remain inert; publication never refreshes cached data.
    pub fn edit_dde<F>(&mut self, update: F) -> Result<()>
    where
        F: FnOnce(&mut crate::dde::Edit) -> Result<()>,
    {
        self.edit_dde_with_context(
            crate::dde::Limits::default(),
            &crate::dde::default_context(),
            update,
        )
    }

    /// Edit DDE metadata with explicit parsing, staging, and readback budgets.
    ///
    /// Package replacement and facade rehydration use the package policy. The
    /// supplied context is checked after rehydration and before publication.
    pub fn edit_dde_with_context<F>(
        &mut self,
        limits: crate::dde::Limits,
        context: &litchi_core::ExecutionContext,
        update: F,
    ) -> Result<()>
    where
        F: FnOnce(&mut crate::dde::Edit) -> Result<()>,
    {
        let snapshot = self.dde_with(limits, context)?;
        let mut edit = snapshot.edit();
        update(&mut edit)?;
        let commit = edit.commit(context)?;
        if commit.changed() {
            self.ensure_dde_publication_allowed()?;
            let package = self
                .package
                .replace_content_xml(commit.snapshot().source_xml())?;
            let candidate = Self::from_package(package)?;
            context.check().map_err(map_sheet_metadata_execution)?;
            *self = candidate;
        }
        Ok(())
    }

    /// Apply an exact-source DDE patch and rehydrate the accepted package.
    pub fn apply_dde_patch(&mut self, patch: &crate::dde::Patch) -> Result<()> {
        let context = crate::dde::default_context();
        let snapshot = self.dde_with(crate::dde::Limits::default(), &context)?;
        let commit = patch.apply(&snapshot)?;
        if commit.changed() {
            self.ensure_dde_publication_allowed()?;
            let package = self
                .package
                .replace_content_xml(commit.snapshot().source_xml())?;
            let candidate = Self::from_package(package)?;
            context.check().map_err(map_sheet_metadata_execution)?;
            *self = candidate;
        }
        Ok(())
    }

    fn ensure_dde_publication_allowed(&self) -> Result<()> {
        if self
            .package
            .package()
            .files()?
            .into_iter()
            .any(|path| litchi_odf_common::core::package::is_signature_owner_path(&path))
        {
            return Err(litchi_core::Error::Unsupported(
                "signed-source refusal: changed ODS DDE metadata requires explicit unsign/resign policy"
                    .to_string(),
            ));
        }
        Ok(())
    }

    /// Inspect typed scenario declarations without applying their values.
    ///
    /// # Errors
    ///
    /// Returns an error when the content XML has invalid or over-budget
    /// scenario metadata.
    pub fn scenarios(&self) -> Result<crate::scenario::Snapshot> {
        crate::scenario::Snapshot::parse(self.package.content_xml()).map_err(|error| {
            litchi_core::Error::InvalidFormat(format!(
                "ODS scenario metadata inspection failed: {error}"
            ))
        })
    }

    /// Apply an exact-source, inert scenario metadata patch and rehydrate the
    /// package only after the candidate has passed typed readback.
    ///
    /// Scenario declarations are metadata only.  This method never applies a
    /// what-if scenario, evaluates a formula, or refreshes external data.
    pub fn apply_scenario_patch(&mut self, patch: &crate::scenario::Patch) -> Result<()> {
        let snapshot = self.scenarios()?;
        let commit = patch.apply(&snapshot)?;
        if commit.changed() {
            let package = self
                .package
                .replace_content_xml(commit.snapshot().source_xml())?;
            *self = Self::from_package(package)?;
        }
        Ok(())
    }

    /// Stage and publish one failure-atomic scenario metadata edit.
    ///
    /// The closure edits typed declarations selected by exact worksheet name
    /// or source order.  A failed closure, stale source, invalid XML, or typed
    /// readback leaves this spreadsheet unchanged.
    pub fn edit_scenarios<F>(&mut self, update: F) -> Result<()>
    where
        F: FnOnce(&mut crate::scenario::Edit) -> Result<()>,
    {
        let snapshot = self.scenarios()?;
        let mut edit = snapshot.edit();
        update(&mut edit)?;
        let commit = edit.commit()?;
        if commit.changed() {
            let package = self
                .package
                .replace_content_xml(commit.snapshot().source_xml())?;
            *self = Self::from_package(package)?;
        }
        Ok(())
    }

    /// Capture the source-backed consolidation, label-range, cell-source, and
    /// detective metadata catalog for this owned spreadsheet.
    pub fn sheet_metadata(&self) -> Result<crate::sheet_metadata::Snapshot> {
        crate::sheet_metadata::Snapshot::parse(self.package.content_xml())
    }

    /// Capture sheet metadata under explicit finite limits and execution context.
    pub fn sheet_metadata_with(
        &self,
        limits: crate::sheet_metadata::Limits,
        context: &litchi_core::ExecutionContext,
    ) -> Result<crate::sheet_metadata::Snapshot> {
        crate::sheet_metadata::Snapshot::parse_with_context(
            self.package.content_xml(),
            limits,
            context,
        )
    }

    /// Stage and publish one failure-atomic sheet metadata edit.
    pub fn edit_sheet_metadata<F>(&mut self, update: F) -> Result<()>
    where
        F: FnOnce(&mut crate::sheet_metadata::Edit) -> Result<()>,
    {
        self.edit_sheet_metadata_with_context(
            crate::sheet_metadata::Limits::default(),
            &crate::sheet_metadata::default_context(),
            update,
        )
    }

    /// Stage and publish sheet metadata under explicit limits and context.
    ///
    /// The supplied context governs metadata parsing, staging, candidate
    /// rendering, and target readback. Owned package replacement and facade
    /// rehydration use the package's own bounded policy; the context is
    /// checked again after rehydration and before this spreadsheet is replaced.
    pub fn edit_sheet_metadata_with_context<F>(
        &mut self,
        limits: crate::sheet_metadata::Limits,
        context: &litchi_core::ExecutionContext,
        update: F,
    ) -> Result<()>
    where
        F: FnOnce(&mut crate::sheet_metadata::Edit) -> Result<()>,
    {
        let snapshot = self.sheet_metadata_with(limits, context)?;
        let mut edit = snapshot.edit();
        update(&mut edit)?;
        let commit = edit.commit(context)?;
        if commit.changed() {
            self.ensure_sheet_metadata_publication_allowed()?;
            let package = self
                .package
                .replace_content_xml(commit.snapshot().source_xml())?;
            let candidate = Self::from_package(package)?;
            context.check().map_err(map_sheet_metadata_execution)?;
            *self = candidate;
        }
        Ok(())
    }

    /// Apply an exact source patch and rehydrate the owned spreadsheet only
    /// after the candidate passes complete metadata readback.
    pub fn apply_sheet_metadata_patch(
        &mut self,
        patch: &crate::sheet_metadata::Patch,
    ) -> Result<()> {
        let snapshot = self.sheet_metadata()?;
        let commit = patch.apply(&snapshot)?;
        if commit.changed() {
            self.ensure_sheet_metadata_publication_allowed()?;
            let package = self
                .package
                .replace_content_xml(commit.snapshot().source_xml())?;
            *self = Self::from_package(package)?;
        }
        Ok(())
    }

    fn ensure_sheet_metadata_publication_allowed(&self) -> Result<()> {
        let signed = self
            .package
            .package()
            .files()?
            .into_iter()
            .any(|path| litchi_odf_common::core::package::is_signature_owner_path(&path));
        if signed {
            return Err(litchi_core::Error::Unsupported(
                "signed-source refusal: changed ODS sheet metadata requires explicit unsign/resign policy"
                    .to_string(),
            ));
        }
        Ok(())
    }

    /// Stage, validate, rebuild, and fully rehydrate one inert tracked-change edit.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn update_tracked_changes<F>(&mut self, edit: F) -> Result<()>
    where
        F: FnOnce(&mut crate::tracked_changes::Transaction) -> Result<()>,
    {
        let snapshot = self.tracked_changes()?;
        let commit = crate::tracked_changes::update(&snapshot, edit)?;
        if !commit.changed() {
            return Ok(());
        }
        let package = self.package.replace_tracked_changes(&commit)?;
        let candidate = Self::from_package(package)?;
        *self = candidate;
        Ok(())
    }

    /// Apply an exact-source tracked-change patch and fully rehydrate the candidate.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn apply_tracked_changes_patch(
        &mut self,
        patch: &crate::tracked_changes::Patch,
    ) -> Result<()> {
        let snapshot = self.tracked_changes()?;
        let commit = patch.apply(&snapshot)?;
        if !commit.changed() {
            return Ok(());
        }
        let package = self.package.replace_tracked_changes(&commit)?;
        let candidate = Self::from_package(package)?;
        *self = candidate;
        Ok(())
    }

    /// Publish a validated annotation transaction without rebuilding an
    /// unchanged package.
    pub(crate) fn publish_annotations(&mut self, content_xml: &str) -> Result<()> {
        let package = crate::annotations::replace_content(&self.package, content_xml)?;
        *self = Self::from_package(package)?;
        Ok(())
    }

    /// Borrow the compact cross-format metadata projection.
    #[must_use]
    pub fn metadata(&self) -> &litchi_core::Metadata {
        self.metadata.value()
    }

    /// Borrow the complete typed ODF metadata model.
    #[must_use]
    pub fn odf_metadata(&self) -> &crate::metadata::Metadata {
        self.metadata.odf()
    }

    /// Borrow the retained metadata snapshot, including bounded source XML.
    #[must_use]
    pub fn metadata_snapshot(&self) -> &crate::metadata::Snapshot {
        &self.metadata
    }

    /// Borrow spreadsheet calculation settings, if the document declares them.
    #[must_use]
    pub fn settings(&self) -> Option<&crate::settings::Settings> {
        self.settings.as_ref()
    }

    /// Alias whose name makes the content-level ODF owner explicit.
    #[must_use]
    pub fn calculation_settings(&self) -> Option<&crate::settings::Settings> {
        self.settings()
    }

    /// Discover the typed `DataPilot` catalog owned by this spreadsheet.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn data_pilots(&self) -> Result<crate::data_pilot::Catalog<'_>> {
        crate::data_pilot::Catalog::load(&self.package)
    }

    /// Capture the `DataPilot` owner as an immutable, exact-package snapshot.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn data_pilot_snapshot(&self) -> Result<crate::data_pilot::Snapshot> {
        crate::data_pilot::Snapshot::from_bytes(self.package.package().as_bytes().to_vec())
    }

    /// Discover inert database-range declarations owned by this spreadsheet.
    ///
    /// Database sources, filters, sorting, and subtotals are metadata only;
    /// this method never opens a database, refreshes a range, or evaluates a
    /// filter or subtotal.
    ///
    /// # Errors
    /// Returns an error when the content owner is malformed or over budget.
    pub fn database_ranges(&self) -> Result<crate::database_range::Catalog<'_>> {
        crate::database_range::Catalog::load(&self.package)
    }

    /// Capture database-range metadata as an immutable, exact-package
    /// snapshot for explicit Snapshot → Edit → Commit → Patch workflows.
    ///
    /// # Errors
    /// Returns an error when the package or typed owner exceeds its finite
    /// admission limits.
    pub fn database_range_snapshot(&self) -> Result<crate::database_range::Snapshot> {
        crate::database_range::Snapshot::from_package(&self.package)
    }

    /// Alias with a plural noun for callers that use the owner name.
    pub fn database_ranges_snapshot(&self) -> Result<crate::database_range::Snapshot> {
        self.database_range_snapshot()
    }

    /// Apply an exact-source database-range patch and rehydrate the full
    /// spreadsheet only after typed readback succeeds.
    ///
    /// # Errors
    /// Returns an error for stale lineage, invalid metadata, or failed
    /// candidate readback.
    pub fn apply_database_range_patch(
        &mut self,
        patch: &crate::database_range::Patch,
    ) -> Result<()> {
        let commit = patch.apply(&self.database_range_snapshot()?)?;
        if commit.changed() {
            *self = Self::from_bytes(commit.snapshot().as_bytes().to_vec())?;
        }
        Ok(())
    }

    /// Clone-stage inert database-range CRUD and publish one package edit.
    ///
    /// Unknown markup inside the owned XML is retained by no-op transactions
    /// and causes a changed transaction to fail before package bytes are
    /// rebuilt.
    ///
    /// # Errors
    /// Returns an error when the closure, source checks, package rebuild, or
    /// typed readback fails.
    pub fn edit_database_ranges<F>(&mut self, edit: F) -> Result<()>
    where
        F: for<'source> FnOnce(&mut crate::database_range::Editor<'_, 'source>) -> Result<()>,
    {
        let commit = {
            let catalog = self.database_ranges()?;
            let mut transaction = catalog.transaction();
            edit(&mut transaction.editor())?;
            transaction.commit()?
        };
        if commit.changed() {
            *self = Self::from_bytes(commit.into_owned_bytes())?;
        }
        Ok(())
    }

    /// Return the typed worksheet graph in document order.
    #[must_use]
    pub fn sheets(&self) -> &[crate::worksheet::Sheet] {
        &self.sheets
    }

    /// Capture worksheets as an immutable, exact-package snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the retained package or worksheet graph is invalid.
    pub fn worksheet_snapshot(&self) -> Result<crate::worksheet::Snapshot> {
        let package = self.package.clone_without_password()?;
        crate::worksheet::Snapshot::from_shared_package(Arc::new(package))
    }

    /// Apply an exact-source reversible worksheet patch.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale patch or invalid candidate package.
    pub fn apply_worksheet_patch(&mut self, patch: &crate::worksheet::Patch) -> Result<()> {
        let commit = patch.apply(&self.worksheet_snapshot()?)?;
        if commit.changed() {
            *self = Self::from_bytes(commit.snapshot().as_bytes().to_vec())?;
        }
        Ok(())
    }

    /// Discover embedded charts in content-level drawing order.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn charts(&self) -> Result<crate::charts::Inventory<'_>> {
        self.charts_with(crate::charts::Limits::default())
    }

    /// Discover embedded charts with an explicit resource budget.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn charts_with(
        &self,
        limits: crate::charts::Limits,
    ) -> Result<crate::charts::Inventory<'_>> {
        crate::charts::inventory(&self.package, limits)
    }

    /// Capture embedded charts as an immutable, exact-package snapshot.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn chart_snapshot(&self) -> Result<crate::charts::Snapshot> {
        self.chart_snapshot_with(crate::charts::Limits::default())
    }

    /// Capture embedded charts with an explicit resource budget.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn chart_snapshot_with(
        &self,
        limits: crate::charts::Limits,
    ) -> Result<crate::charts::Snapshot> {
        crate::charts::Snapshot::from_bytes_with(self.package.package().as_bytes().to_vec(), limits)
    }

    /// Select one embedded chart by exact drawing name or checked position.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn chart<'a, S>(&self, selector: S) -> Result<Option<crate::charts::Chart>>
    where
        S: Into<crate::charts::Selector<'a>>,
    {
        self.charts()?.get(selector).map(|chart| chart.cloned())
    }

    /// Find a worksheet by its exact ODF name.
    #[must_use]
    pub fn sheet(&self, name: &str) -> Option<&crate::worksheet::Sheet> {
        self.sheets.iter().find(|sheet| sheet.name == name)
    }

    /// Look up a logical cell while retaining the distinction between a
    /// missing coordinate and a physical repeated cell run.
    #[must_use]
    pub fn cell(
        &self,
        sheet_name: &str,
        row: usize,
        column: usize,
    ) -> Option<crate::worksheet::CellView<'_>> {
        self.cell_unchecked(CellSelector::new(sheet_name, row, column))
    }

    fn cell_unchecked(&self, selector: CellSelector<'_>) -> Option<crate::worksheet::CellView<'_>> {
        let sheet_index = self
            .sheets
            .iter()
            .position(|sheet| sheet.name == selector.sheet_name)?;
        let direct = || self.sheets[sheet_index].cell_view(selector.row, selector.column);

        if let Some(locator) = self.cell_locator.get() {
            return Some(locator.as_ref().map_or_else(direct, |locator| {
                locator.cell_view(&self.sheets, sheet_index, selector.row, selector.column)
            }));
        }

        let previous = self
            .cell_queries
            .fetch_update(Ordering::Relaxed, Ordering::Relaxed, |count| {
                Some(count.saturating_add(1))
            })
            .unwrap_or(usize::MAX);
        if previous.saturating_add(1) >= cell_locator::BUILD_QUERY_THRESHOLD {
            let locator = self
                .cell_locator
                .get_or_init(|| cell_locator::CellLocator::try_build(&self.sheets));
            return Some(locator.as_ref().map_or_else(direct, |locator| {
                locator.cell_view(&self.sheets, sheet_index, selector.row, selector.column)
            }));
        }

        Some(direct())
    }

    /// Look up an ordered batch of logical cells with one bounded result
    /// allocation.
    ///
    /// A missing worksheet produces `None` for that selector, while an
    /// existing worksheet with no physical cell at the coordinate produces
    /// `Some(CellView::Missing)`, exactly matching [`Self::cell`].  Results
    /// retain selector order and duplicate selectors are allowed.
    ///
    /// # Errors
    ///
    /// Returns a typed allocation error when the result vector cannot be
    /// reserved, or an invalid-format error when the selector bound is
    /// exceeded.  The bound is checked before lookup or locator construction.
    pub fn cell_batch(
        &self,
        selectors: &[CellSelector<'_>],
    ) -> Result<Vec<Option<crate::worksheet::CellView<'_>>>> {
        validate_cell_batch_len(selectors.len())?;
        let mut values = Vec::new();
        values
            .try_reserve_exact(selectors.len())
            .map_err(|source| litchi_core::Error::Allocation {
                resource: "ODS cell lookup batch results",
                source,
            })?;
        for &selector in selectors {
            values.push(self.cell_unchecked(selector));
        }
        Ok(values)
    }

    /// Alias for [`Self::cell_batch`].
    pub fn cells(
        &self,
        selectors: &[CellSelector<'_>],
    ) -> Result<Vec<Option<crate::worksheet::CellView<'_>>>> {
        self.cell_batch(selectors)
    }

    /// Discover package, inline, missing, and inert linked images.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn images(&self) -> Result<Vec<crate::media::Image>> {
        let package = self.package.package().package()?;
        crate::media::scan_package(
            self.package.content_xml(),
            self.package.styles_xml(),
            &package,
        )
    }

    /// Inspect inert conditional-format, sparkline, hyperlink, and in-table
    /// drawing source metadata without evaluating or dereferencing it.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn source_features(&self) -> Result<crate::source_features::Snapshot> {
        crate::source_features::Snapshot::parse(self.package.content_xml())
    }

    /// Inspect the source-backed content-validation catalog and compact cell bindings.
    ///
    /// This is a read-only ownership inventory. It does not authorize mutation or
    /// publication of the catalog.
    ///
    /// # Errors
    /// Returns a typed error for malformed ownership, unsupported MCE selection,
    /// allocation failure, or a resource-limit excess.
    pub fn content_validations(
        &self,
    ) -> crate::content_validation::Result<crate::content_validation::Snapshot<'_>> {
        crate::content_validation::Snapshot::parse(self.package.content_xml())
    }

    /// Discover package, inline, missing, and inert linked embedded objects.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn embedded_objects(&self) -> Result<Vec<crate::embedded::Object>> {
        let package = self.package.package().package()?;
        crate::embedded::scan_package(
            self.package.content_xml(),
            self.package.styles_xml(),
            &package,
        )
    }

    /// Return bytes only for inline or verified package-contained images.
    /// Linked and missing images remain inert and are never fetched.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn image_bytes(&self, image: &crate::media::Image) -> Result<Option<Vec<u8>>> {
        match &image.source {
            crate::media::Source::Inline { bytes, .. } => Ok(Some(bytes.clone())),
            crate::media::Source::PackagePart { path, .. } => {
                self.package.package().get_file(path).map(Some)
            },
            crate::media::Source::MissingPackagePart { .. }
            | crate::media::Source::Linked { .. }
            | crate::media::Source::Missing
            | _ => Ok(None),
        }
    }

    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        match Arc::try_unwrap(self.package) {
            Ok(package) => package.into_bytes(),
            Err(package) => package.package().as_bytes().to_vec(),
        }
    }

    /// Return all global and sheet-local named definitions in document order.
    #[must_use]
    pub fn definitions(&self) -> &[Definition] {
        &self.definitions
    }

    /// Capture named definitions as an immutable, exact-package snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the retained package cannot be reparsed.
    pub fn definitions_snapshot(&self) -> Result<crate::definitions::Snapshot> {
        crate::definitions::Snapshot::from_bytes(self.package.package().as_bytes().to_vec())
    }

    /// Apply an exact-source reversible named-definition patch.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale patch or invalid candidate package. This facade changes only
    /// after the complete target has been reparsed.
    pub fn apply_definitions_patch(&mut self, patch: &crate::definitions::Patch) -> Result<()> {
        let commit = patch.apply(&self.definitions_snapshot()?)?;
        if commit.changed() {
            *self = Self::from_bytes(commit.snapshot().as_bytes().to_vec())?;
        }
        Ok(())
    }

    /// Return named ranges in their document order.
    pub fn ranges(&self) -> impl Iterator<Item = &Range> {
        self.definitions
            .iter()
            .filter_map(|definition| match definition {
                Definition::Range(range) => Some(range),
                Definition::Expression(_) => None,
            })
    }

    /// Return named expressions in their document order.
    pub fn expressions(&self) -> impl Iterator<Item = &Expression> {
        self.definitions
            .iter()
            .filter_map(|definition| match definition {
                Definition::Range(_) => None,
                Definition::Expression(expression) => Some(expression),
            })
    }

    /// Find a named range by its exact name and visibility scope.
    #[must_use]
    pub fn range(&self, name: &str, scope: &Scope) -> Option<&Range> {
        self.ranges()
            .find(|range| range.name == name && &range.scope == scope)
    }

    /// Find a named expression by its exact name and visibility scope.
    #[must_use]
    pub fn expression(&self, name: &str, scope: &Scope) -> Option<&Expression> {
        self.expressions()
            .find(|expression| expression.name == name && &expression.scope == scope)
    }

    /// Atomically append a validated named range.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn add_range(&mut self, range: Range) -> Result<()> {
        self.add_definition(range.into())
    }

    /// Atomically append a validated named expression.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn add_expression(&mut self, expression: Expression) -> Result<()> {
        self.add_definition(expression.into())
    }

    /// Atomically append a validated named definition while preserving catalog order.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn add_definition(&mut self, definition: Definition) -> Result<()> {
        let mut candidate = self.definitions.clone();
        candidate.push(definition);
        self.set_definitions(candidate)
    }

    /// Atomically replace the complete ordered named-definition catalog.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_definitions(&mut self, definitions: Vec<Definition>) -> Result<()> {
        let updated = crate::codec::names::replace(self.package.content_xml(), &definitions)?;
        let package = self.package.replace_content_xml(&updated)?;
        self.package = Arc::new(package);
        self.definitions = definitions;
        Ok(())
    }

    /// Publish a validated worksheet snapshot as one package transaction.
    pub(crate) fn publish_sheets(&mut self, sheets: Vec<crate::worksheet::Sheet>) -> Result<()> {
        crate::worksheet::package::validate_owned_link_delta(&self.sheets, &sheets)?;
        let package = self.package.replace_sheets(&sheets)?;
        self.package = Arc::new(package);
        self.sheets = sheets;
        Ok(())
    }

    pub(crate) fn publish_metadata(&mut self, metadata: litchi_core::Metadata) -> Result<()> {
        let package = self.package.metadata_snapshot()?;
        let mut transaction = package.transaction();
        transaction.replace(metadata)?;
        let commit = transaction.commit()?;
        if !commit.changed() {
            return Ok(());
        }
        let metadata_xml = commit.into_owned_xml().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "changed ODS metadata transaction produced no XML".to_string(),
            )
        })?;
        let package = self.package.replace_metadata_xml(Some(&metadata_xml))?;
        *self = Self::from_package(package)?;
        Ok(())
    }

    pub(crate) fn remove_metadata(&mut self) -> Result<()> {
        let snapshot = self.package.metadata_snapshot()?;
        let mut transaction = snapshot.transaction();
        transaction.remove();
        let commit = transaction.commit()?;
        if !commit.changed() {
            return Ok(());
        }
        let package = self.package.replace_metadata_xml(None)?;
        *self = Self::from_package(package)?;
        Ok(())
    }

    pub(crate) fn publish_settings(
        &mut self,
        settings: Option<crate::settings::Settings>,
    ) -> Result<()> {
        if self.settings == settings {
            return Ok(());
        }
        let package = self
            .package
            .replace_calculation_settings(settings.as_ref())?;
        *self = Self::from_package(package)?;
        Ok(())
    }

    /// Read all inert RDF metadata graphs in package order.
    ///
    /// # Errors
    ///
    /// Returns an error when the manifest or a declared graph is invalid.
    pub fn rdf_graphs(&self) -> Result<Vec<Graph>> {
        litchi_odf_common::rdf::graphs(self.package.package())
    }

    /// Capture RDF graphs as an immutable, exact-package snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the package, manifest, or a declared graph is invalid.
    pub fn rdf_snapshot(&self) -> Result<crate::metadata_graphs::Snapshot> {
        crate::metadata_graphs::Snapshot::from_bytes(self.package.package().as_bytes().to_vec())
    }

    /// Apply an exact-source reversible RDF graph patch.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale patch or invalid candidate package.
    pub fn apply_rdf_patch(&mut self, patch: &crate::metadata_graphs::Patch) -> Result<()> {
        let commit = patch.apply(&self.rdf_snapshot()?)?;
        self.publish_rdf_commit(commit)
    }

    /// Add a graph and atomically replace this snapshot with the rebuilt package.
    ///
    /// # Errors
    ///
    /// Returns an error when the path, triples, compact XML, or rebuilt package is invalid.
    pub fn add_rdf_graph(
        &mut self,
        preferred_path: Option<&str>,
        triples: &[Triple],
    ) -> Result<String> {
        let snapshot = self.rdf_snapshot()?;
        let mut edit = snapshot.edit();
        let path = edit.add_graph(preferred_path, triples)?;
        self.publish_rdf_commit(edit.commit())?;
        Ok(path)
    }

    /// Replace one complete RDF graph and atomically publish the result.
    ///
    /// # Errors
    ///
    /// Returns an error when the graph, triples, compact XML, or rebuilt package is invalid.
    pub fn replace_rdf_graph(&mut self, path: &str, triples: &[Triple]) -> Result<()> {
        let snapshot = self.rdf_snapshot()?;
        let mut edit = snapshot.edit();
        edit.replace_graph(path, triples)?;
        self.publish_rdf_commit(edit.commit())
    }

    /// Remove one RDF graph after validating that no remaining graph references it.
    ///
    /// # Errors
    ///
    /// Returns an error when the graph is missing, referenced, or package rebuilding fails.
    pub fn remove_rdf_graph(&mut self, path: &str) -> Result<()> {
        let snapshot = self.rdf_snapshot()?;
        let mut edit = snapshot.edit();
        edit.remove_graph(path)?;
        self.publish_rdf_commit(edit.commit())
    }

    /// Append one triple to an existing graph and return its committed index.
    ///
    /// # Errors
    ///
    /// Returns an error when the graph, triple, compact XML, or rebuilt package is invalid.
    pub fn add_rdf_triple(&mut self, path: &str, triple: &Triple) -> Result<usize> {
        let snapshot = self.rdf_snapshot()?;
        let mut edit = snapshot.edit();
        let position = edit.add_triple(path, triple)?;
        self.publish_rdf_commit(edit.commit())?;
        Ok(position.get())
    }

    /// Replace one triple while preserving its description subject.
    ///
    /// # Errors
    ///
    /// Returns an error when the graph, position, triple, or rebuilt package is invalid.
    pub fn replace_rdf_triple(&mut self, path: &str, index: usize, triple: &Triple) -> Result<()> {
        let snapshot = self.rdf_snapshot()?;
        let mut edit = snapshot.edit();
        edit.replace_triple(path, litchi_core::Position::new(index), triple)?;
        self.publish_rdf_commit(edit.commit())
    }

    /// Remove one triple from a graph.
    ///
    /// # Errors
    ///
    /// Returns an error when the graph, position, compact XML, or rebuilt package is invalid.
    pub fn remove_rdf_triple(&mut self, path: &str, index: usize) -> Result<()> {
        let snapshot = self.rdf_snapshot()?;
        let mut edit = snapshot.edit();
        edit.remove_triple(path, litchi_core::Position::new(index))?;
        self.publish_rdf_commit(edit.commit())
    }

    /// Move one triple within its RDF description.
    ///
    /// # Errors
    ///
    /// Returns an error when either position, the graph, or rebuilt package is invalid.
    pub fn move_rdf_triple(&mut self, path: &str, from: usize, to: usize) -> Result<()> {
        let snapshot = self.rdf_snapshot()?;
        let mut edit = snapshot.edit();
        edit.move_triple(
            path,
            litchi_core::Position::new(from),
            litchi_core::Position::new(to),
        )?;
        self.publish_rdf_commit(edit.commit())
    }

    fn publish_rdf_commit(&mut self, commit: crate::metadata_graphs::Commit) -> Result<()> {
        if commit.changed() {
            let snapshot = commit.into_snapshot();
            *self = Self::from_bytes(snapshot.as_bytes().to_vec())?;
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Arc;

    const ANNOTATED_CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:vendor="urn:example:vendor" office:version="1.3"><office:body><office:spreadsheet><vendor:keep/><table:table table:name="Data"><table:table-row><table:table-cell><office:annotation><text:p>existing</text:p></office:annotation></table:table-cell><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;

    #[test]
    fn builder_round_trips_through_facade() {
        let bytes = Builder::new()
            .build()
            .expect("test fixture or operation should succeed");
        let spreadsheet =
            Spreadsheet::from_bytes(bytes).expect("test fixture or operation should succeed");
        assert!(spreadsheet.content_xml().contains("office:spreadsheet"));
    }

    #[test]
    fn worksheet_snapshot_reuses_the_facade_package_index() {
        let bytes = Builder::new()
            .build()
            .expect("test fixture or operation should succeed");
        let spreadsheet =
            Spreadsheet::from_bytes(bytes).expect("test fixture or operation should succeed");
        let snapshot = spreadsheet
            .worksheet_snapshot()
            .expect("test fixture or operation should succeed");

        assert_eq!(
            snapshot.prepared_index_identity(),
            spreadsheet.prepared_index_identity()
        );
    }

    #[test]
    fn cell_locator_builds_at_the_threshold_and_preserves_snapshot_traits() {
        fn assert_send_sync<T: Send + Sync>() {}
        assert_send_sync::<Spreadsheet>();

        let bytes = Builder::new()
            .content_xml(ANNOTATED_CONTENT)
            .build()
            .expect("test fixture or operation should succeed");
        let spreadsheet =
            Spreadsheet::from_bytes(bytes).expect("test fixture or operation should succeed");

        for _ in 1..cell_locator::BUILD_QUERY_THRESHOLD {
            assert!(matches!(
                spreadsheet.cell("Data", 0, 0),
                Some(crate::worksheet::CellView::Stored(_))
            ));
        }
        assert!(spreadsheet.cell_locator.get().is_none());
        assert!(matches!(
            spreadsheet.cell("Data", 0, 0),
            Some(crate::worksheet::CellView::Stored(_))
        ));
        assert!(matches!(spreadsheet.cell_locator.get(), Some(Some(_))));
    }

    #[test]
    fn cell_batch_matches_scalar_order_missing_distinction_and_identity() {
        let bytes = Builder::new()
            .content_xml(ANNOTATED_CONTENT)
            .build()
            .expect("test fixture or operation should succeed");
        let spreadsheet =
            Spreadsheet::from_bytes(bytes).expect("test fixture or operation should succeed");
        let selectors = [
            CellSelector::new("Data", 0, 1),
            CellSelector::new("Data", 0, 0),
            CellSelector::new("Data", 0, 2),
            CellSelector::new("Missing", 0, 0),
            CellSelector::new("Data", 0, 0),
            CellSelector::new("Data", usize::MAX, usize::MAX),
        ];

        let batch = spreadsheet
            .cell_batch(&selectors)
            .expect("test fixture or operation should succeed");
        let scalar = selectors
            .iter()
            .map(|selector| {
                spreadsheet.cell(selector.sheet_name(), selector.row(), selector.column())
            })
            .collect::<Vec<_>>();
        assert_eq!(batch, scalar);
        assert!(matches!(
            batch[2],
            Some(crate::worksheet::CellView::Missing)
        ));
        assert_eq!(batch[3], None);

        for (selector, actual) in selectors.iter().zip(batch) {
            if let Some(crate::worksheet::CellView::Stored(actual)) = actual {
                let Some(crate::worksheet::CellView::Stored(expected)) =
                    spreadsheet.cell(selector.sheet_name(), selector.row(), selector.column())
                else {
                    panic!("scalar lookup lost a stored cell");
                };
                assert!(std::ptr::eq(actual, expected));
            }
        }
    }

    #[test]
    fn cell_batch_is_empty_and_bounded_before_lookup_work() {
        let bytes = Builder::new()
            .content_xml(ANNOTATED_CONTENT)
            .build()
            .expect("test fixture or operation should succeed");
        let spreadsheet =
            Spreadsheet::from_bytes(bytes).expect("test fixture or operation should succeed");
        assert!(
            spreadsheet
                .cell_batch(&[])
                .expect("empty batch should succeed")
                .is_empty()
        );

        let exact = (0..MAX_CELL_SELECTORS)
            .map(|index| match index % 4 {
                0 => CellSelector::new("Data", 0, 0),
                1 => CellSelector::new("Data", 0, 2),
                2 => CellSelector::new("Missing", 0, 0),
                _ => CellSelector::new("Data", usize::MAX, usize::MAX),
            })
            .collect::<Vec<_>>();
        let values = spreadsheet
            .cell_batch(&exact)
            .expect("exact selector bound should succeed");
        assert_eq!(values.len(), MAX_CELL_SELECTORS);
        for (index, value) in values.iter().enumerate() {
            match index % 4 {
                0 => assert!(matches!(
                    value,
                    Some(crate::worksheet::CellView::Stored(cell)) if cell.text == "existing"
                )),
                1 | 3 => assert!(matches!(value, Some(crate::worksheet::CellView::Missing))),
                _ => assert_eq!(*value, None),
            }
        }

        let bounded = Spreadsheet::from_bytes(
            Builder::new()
                .content_xml(ANNOTATED_CONTENT)
                .build()
                .expect("test fixture or operation should succeed"),
        )
        .expect("test fixture or operation should succeed");
        let selectors = vec![CellSelector::new("Data", 0, 0); MAX_CELL_SELECTORS + 1];
        let error = bounded.cell_batch(&selectors).expect_err("bound must fail");
        assert!(matches!(
            error,
            litchi_core::Error::InvalidFormat(message)
                if message.contains("selector safety limit")
        ));
        assert!(bounded.cell_locator.get().is_none());
        assert_eq!(bounded.cell_queries.load(Ordering::Relaxed), 0);
    }

    #[test]
    fn cell_batch_is_send_sync_and_concurrent_locator_build_is_shared() {
        fn assert_send_sync<T: Send + Sync>() {}
        assert_send_sync::<CellSelector<'static>>();

        let bytes = Builder::new()
            .content_xml(ANNOTATED_CONTENT)
            .build()
            .expect("test fixture or operation should succeed");
        let spreadsheet = Arc::new(
            Spreadsheet::from_bytes(bytes).expect("test fixture or operation should succeed"),
        );
        let expected = match spreadsheet.sheet("Data").and_then(|sheet| sheet.cell(0, 0)) {
            Some(cell) => std::ptr::from_ref(cell) as usize,
            None => panic!("fixture cell"),
        };
        let selector = CellSelector::new("Data", 0, 0);
        let threads = (0..8)
            .map(|_| {
                let spreadsheet = Arc::clone(&spreadsheet);
                std::thread::spawn(move || {
                    for _ in 0..cell_locator::BUILD_QUERY_THRESHOLD {
                        let result = spreadsheet.cell_batch(&[selector]).expect("batch lookup");
                        let Some(crate::worksheet::CellView::Stored(cell)) =
                            result.first().copied().flatten()
                        else {
                            panic!("fixture cell");
                        };
                        assert_eq!(std::ptr::from_ref(cell) as usize, expected);
                    }
                })
            })
            .collect::<Vec<_>>();
        for thread in threads {
            thread.join().expect("cell batch thread");
        }
        assert!(matches!(spreadsheet.cell_locator.get(), Some(Some(_))));
    }

    #[test]
    fn concurrent_first_cell_locator_build_is_shared_and_identical() {
        let bytes = Builder::new()
            .content_xml(ANNOTATED_CONTENT)
            .build()
            .expect("test fixture or operation should succeed");
        let spreadsheet = Arc::new(
            Spreadsheet::from_bytes(bytes).expect("test fixture or operation should succeed"),
        );
        let expected = std::ptr::from_ref(
            spreadsheet
                .sheet("Data")
                .and_then(|sheet| sheet.cell(0, 0))
                .expect("test fixture or operation should succeed"),
        ) as usize;

        let threads = (0..8)
            .map(|_| {
                let spreadsheet = Arc::clone(&spreadsheet);
                std::thread::spawn(move || {
                    for _ in 0..cell_locator::BUILD_QUERY_THRESHOLD {
                        let Some(crate::worksheet::CellView::Stored(cell)) =
                            spreadsheet.cell("Data", 0, 0)
                        else {
                            panic!("test fixture or operation should succeed");
                        };
                        assert_eq!(std::ptr::from_ref(cell) as usize, expected);
                    }
                })
            })
            .collect::<Vec<_>>();
        for thread in threads {
            thread
                .join()
                .expect("test fixture or operation should succeed");
        }
        assert!(matches!(spreadsheet.cell_locator.get(), Some(Some(_))));
    }

    #[test]
    fn facade_replacement_discards_built_cell_locator() {
        let bytes = Builder::new()
            .content_xml(ANNOTATED_CONTENT)
            .build()
            .expect("test fixture or operation should succeed");
        let mut spreadsheet =
            Spreadsheet::from_bytes(bytes).expect("test fixture or operation should succeed");
        for _ in 0..cell_locator::BUILD_QUERY_THRESHOLD {
            assert!(spreadsheet.cell("Data", 0, 0).is_some());
        }
        assert!(matches!(spreadsheet.cell_locator.get(), Some(Some(_))));

        let updated = ANNOTATED_CONTENT.replace("existing", "replacement");
        spreadsheet
            .publish_annotations(&updated)
            .expect("test fixture or operation should succeed");
        assert!(spreadsheet.cell_locator.get().is_none());
        assert_eq!(spreadsheet.cell_queries.load(Ordering::Relaxed), 0);
        assert_eq!(
            spreadsheet
                .annotations()
                .expect("test fixture or operation should succeed")
                .cell("Data", 0, 0)
                .expect("test fixture or operation should succeed")
                .expect("test fixture or operation should succeed")
                .annotation()
                .text(),
            "replacement"
        );
    }

    #[test]
    fn shared_resource_inventory_is_available_from_spreadsheet() {
        let bytes = Builder::new()
            .content_xml(
                r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.3"><office:body><office:spreadsheet><draw:frame draw:name="Photo"><draw:image><office:binary-data>AQID</office:binary-data></draw:image></draw:frame><draw:object xlink:href="https://example.invalid/object" xlink:type="simple"/></office:spreadsheet></office:body></office:document-content>"#,
            )
            .build()
            .expect("test fixture or operation should succeed");
        let spreadsheet =
            Spreadsheet::from_bytes(bytes).expect("test fixture or operation should succeed");

        let images = spreadsheet
            .images()
            .expect("test fixture or operation should succeed");
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].inline_bytes(), Some(&[1, 2, 3][..]));
        assert_eq!(
            spreadsheet
                .image_bytes(&images[0])
                .expect("test fixture or operation should succeed"),
            Some(vec![1, 2, 3])
        );

        let objects = spreadsheet
            .embedded_objects()
            .expect("test fixture or operation should succeed");
        assert_eq!(objects.len(), 1);
        assert!(matches!(
            objects[0].source,
            crate::embedded::Source::Linked { ref href }
                if href == "https://example.invalid/object"
        ));
    }

    #[test]
    fn spreadsheet_and_mutable_facades_expose_contextual_annotation_edits() {
        let bytes = Builder::new()
            .content_xml(ANNOTATED_CONTENT)
            .build()
            .expect("test fixture or operation should succeed");
        let spreadsheet = Spreadsheet::from_bytes(bytes.clone())
            .expect("test fixture or operation should succeed");
        let annotations = spreadsheet
            .annotations()
            .expect("test fixture or operation should succeed");
        assert_eq!(
            annotations
                .cell("Data", 0, 0)
                .expect("test fixture or operation should succeed")
                .expect("test fixture or operation should succeed")
                .annotation()
                .text(),
            "existing"
        );

        let mut mutable = MutableSpreadsheet::from_bytes(bytes.clone())
            .expect("test fixture or operation should succeed");
        mutable
            .edit_annotations(|transaction| {
                transaction.set("Data", 0, 1, crate::annotations::Annotation::new("added"))
            })
            .expect("test fixture or operation should succeed");
        let edited = Spreadsheet::from_bytes(mutable.to_bytes())
            .expect("test fixture or operation should succeed");
        assert_eq!(
            edited
                .annotations()
                .expect("test fixture or operation should succeed")
                .cell("Data", 0, 1)
                .expect("test fixture or operation should succeed")
                .expect("test fixture or operation should succeed")
                .annotation()
                .text(),
            "added"
        );
        assert!(edited.content_xml().contains("vendor:keep"));

        let mut no_op = MutableSpreadsheet::from_bytes(bytes.clone())
            .expect("test fixture or operation should succeed");
        no_op
            .edit_annotations(|_| Ok(()))
            .expect("test fixture or operation should succeed");
        assert_eq!(no_op.to_bytes(), bytes);
    }
}
