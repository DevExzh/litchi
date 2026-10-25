//! Concise user-facing ODS entry points.

mod cell_locator;
mod selective;
pub(crate) mod source;
mod source_edit;

use litchi_core::{
    Result as CoreResult, SequentialTextWriter, TextOutputError, TextOutputOptions,
    TextOutputReport,
};
use std::{
    io::Write,
    sync::{
        Arc, OnceLock,
        atomic::{AtomicUsize, Ordering},
    },
};

pub use crate::authoring::Builder;
use crate::model::names::{Definition, Expression, Range, Scope};
pub use litchi_odf_common::rdf::{Graph, Object, Subject, Triple};
pub use selective::{SheetCatalogEntry, SourceBackedSpreadsheetCatalog, SourceReadMetrics};
pub use source::SourceBackedSpreadsheet;
pub(crate) use source_edit::SourceCellEdit;
pub use source_edit::{FormulaChange, HyperlinkChange, HyperlinkOperation};

type FamilyResult<T> = Result<T, litchi_odf_common::Error>;

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

fn validate_cell_batch_len(length: usize) -> CoreResult<()> {
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
            litchi_core::Error::Other("ODS metadata operation cancelled".to_string())
        },
        _other => litchi_core::Error::Other(
            "ODS metadata execution policy rejected operation".to_string(),
        ),
    }
}

/// Immutable ODS document facade.
/// The read-only semantic projection of an `OpenDocument` spreadsheet.
///
/// A `Workbook` is obtained from [`crate::Package::workbook`]; it is owned,
/// immutable and cheap to clone (one reference-count increment). Every state
/// verb (open, edit, apply, save) lives on [`crate::Package`].
#[derive(Clone)]
pub struct Workbook {
    inner: Arc<State>,
}

struct State {
    package: Arc<crate::package::Archive>,
    xml_limits: litchi_odf_common::XmlStreamLimits,
    definitions: Vec<Definition>,
    sheets: Vec<crate::worksheet::Sheet>,
    metadata: crate::metadata::Snapshot,
    settings: Option<crate::settings::Settings>,
    cell_queries: AtomicUsize,
    cell_locator: OnceLock<Option<cell_locator::CellLocator>>,
    tolerance: litchi_odf_common::leniency::ToleranceReport,
}

impl std::fmt::Debug for Workbook {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Workbook")
            .field("sheets", &self.inner.sheets.len())
            .finish_non_exhaustive()
    }
}

impl Workbook {
    /// Reparse complete ODS bytes produced inside the crate.
    pub(crate) fn from_bytes(bytes: Vec<u8>) -> CoreResult<Self> {
        // Let the prepared detector perform the local MIME probe exactly
        // once.  A matching ODS result transfers its indexed archive; a
        // different ODF family still follows the historical package owner so
        // its MIME error is reported at the same boundary; rejected probes
        // recover the original allocation for the ordinary parser.
        match litchi_odf_common::detect::prepared_or_original(
            bytes,
            litchi_odf_common::SourceLimits::default(),
        ) {
            Ok(prepared) if prepared.format() == litchi_core::detection::FileFormat::Ods => {
                Self::from_package(crate::package::Archive::from_prepared_package(prepared)?)
            },
            Ok(prepared) => {
                let package = crate::package::Archive::from_owned_package(prepared.into_package())?;
                Self::from_package(package)
            },
            Err(bytes) => {
                let package = crate::package::Archive::from_bytes(bytes)?;
                Self::from_package(package)
            },
        }
    }

    /// Return the identity of the archive index retained by smart detection.
    #[doc(hidden)]
    #[must_use]
    pub fn prepared_index_identity(&self) -> usize {
        self.inner.package.prepared_index_identity()
    }

    pub(crate) fn from_package(package: crate::package::Archive) -> CoreResult<Self> {
        let xml_limits = package.xml_limits_value();
        Self::from_shared_package(Arc::new(package), xml_limits)
    }

    pub(crate) fn from_shared_package(
        package: Arc<crate::package::Archive>,
        xml_limits: litchi_odf_common::XmlStreamLimits,
    ) -> CoreResult<Self> {
        // Keep ordinary and source-backed opening on the same namespace-aware
        // content reader.  Besides avoiding five independent tokenizations,
        // this admits semantically normalized namespace declarations (for
        // example an entity-escaped `xmlns:xml` URI) without changing the
        // source bytes retained by the package.
        let (settings, outputs) =
            crate::open_parse::OpenParse::run(package.content_xml(), xml_limits)?.finish()?;
        let definitions = outputs.definitions;
        let mut sheets = outputs.sheets;
        let mut tolerance = package.tolerance_report().clone();
        tolerance.append(&outputs.tolerance);
        // `OpenParse` owns the fused worksheet pass, while table title and
        // description are deliberately applied in the small metadata pass
        // shared with the source-backed facade.  Keep this post-pass here so
        // ordinary and source-backed opens expose the same Sheet metadata.
        crate::worksheet::codec::apply_table_metadata(
            package.content_xml(),
            &mut sheets,
            xml_limits,
        )?;
        let metadata = package.metadata_snapshot()?;
        Ok(Self {
            inner: Arc::new(State {
                package,
                xml_limits,
                definitions,
                sheets,
                metadata,
                settings,
                cell_queries: AtomicUsize::new(0),
                cell_locator: OnceLock::new(),
                tolerance,
            }),
        })
    }

    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.inner.package.content_xml()
    }

    /// Everything a tolerant read repaired while this spreadsheet was opened.
    ///
    /// An empty report means the document conformed; see
    /// [`litchi_odf_common::leniency::ToleranceReport`].
    #[inline]
    #[must_use]
    pub fn tolerance_report(&self) -> &litchi_odf_common::leniency::ToleranceReport {
        &self.inner.tolerance
    }

    /// Return worksheet names in document order.
    #[must_use]
    pub fn sheet_names(&self) -> Vec<String> {
        self.inner
            .sheets
            .iter()
            .map(|sheet| sheet.name.clone())
            .collect()
    }

    /// Return the number of worksheets.
    #[must_use]
    pub fn sheet_count(&self) -> usize {
        self.inner.sheets.len()
    }

    /// Extract displayed worksheet text using tab-separated cells and
    /// newline-separated rows, preserving sheet order.
    ///
    /// # Errors
    /// Returns an error when the bounded projection cannot reserve its output.
    pub fn text(&self) -> FamilyResult<String> {
        Ok(source::project_text(&self.inner.sheets)?)
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
    ) -> Result<TextOutputReport, TextOutputError<litchi_core::Error>> {
        let mut writer = SequentialTextWriter::new(output, options);
        source::write_text_to_sheets(&self.inner.sheets, &mut writer, options)?;
        Ok(writer.finish())
    }

    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.inner.package.styles_xml()
    }

    /// Capture the standalone table-template catalog from styles.xml.
    ///
    /// The snapshot retains the exact styles source for no-op, stale-source,
    /// and inverse transaction checks.
    pub fn table_templates(&self) -> FamilyResult<crate::styles::table_template::Snapshot> {
        Ok(crate::styles::table_template::Snapshot::from_source(
            self.inner.package.styles_xml(),
        )?)
    }

    /// Stage a source-checked table-template edit and publish it atomically.
    pub(crate) fn edit_table_templates<F>(&mut self, update: F) -> CoreResult<()>
    where
        F: FnOnce(&mut crate::styles::table_template::Edit) -> CoreResult<()>,
    {
        let snapshot = self.table_templates()?;
        let mut edit = snapshot.edit();
        update(&mut edit)?;
        let commit = edit.commit()?;
        if commit.changed() {
            let package = self
                .inner
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
    pub fn protection(&self) -> FamilyResult<crate::protection::Snapshot> {
        Ok(crate::protection::Snapshot::parse(
            self.inner.package.content_xml(),
            self.inner.package.styles_xml(),
            &self.inner.xml_limits,
        )?)
    }

    /// Apply a failure-atomic protection edit and rebuild only `content.xml`.
    /// Password values remain inert verifiers; this method never authenticates
    /// or enforces a protection policy.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub(crate) fn update_protection<F>(&mut self, edit: F) -> CoreResult<()>
    where
        F: FnOnce(&mut crate::protection::Edit) -> CoreResult<()>,
    {
        let snapshot = self.protection()?;
        let commit = crate::protection::update(&snapshot, edit)?;
        if !commit.changed() {
            return Ok(());
        }
        let package = self
            .inner
            .package
            .replace_content_xml(commit.content_xml())?;
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
    pub fn annotations(&self) -> FamilyResult<crate::annotations::Snapshot> {
        Ok(crate::annotations::Snapshot::parse(
            self.inner.package.content_xml(),
            &self.inner.package.xml_limits_value(),
        )?)
    }

    /// Capture the presence-aware, exact-source tracked-change owner.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn tracked_changes(&self) -> FamilyResult<crate::tracked_changes::Snapshot> {
        Ok(crate::tracked_changes::Snapshot::parse(
            self.inner.package.content_xml(),
        )?)
    }

    /// Capture tracked changes under an explicit resource budget.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn tracked_changes_with(
        &self,
        limits: crate::tracked_changes::Limits,
    ) -> FamilyResult<crate::tracked_changes::Snapshot> {
        Ok(crate::tracked_changes::Snapshot::parse_with_limits(
            self.inner.package.content_xml(),
            limits,
        )?)
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
    pub fn dde(&self) -> FamilyResult<crate::dde::Snapshot> {
        Ok(
            crate::dde::Snapshot::parse(self.inner.package.content_xml())
                .map_err(litchi_core::Error::from)?,
        )
    }

    /// Capture inert DDE declarations and cached tables under explicit limits.
    pub fn dde_with(
        &self,
        limits: crate::dde::Limits,
        context: &litchi_core::ExecutionContext,
    ) -> FamilyResult<crate::dde::Snapshot> {
        Ok(crate::dde::Snapshot::parse_with_context(
            self.inner.package.content_xml(),
            limits,
            context,
        )
        .map_err(litchi_core::Error::from)?)
    }

    /// Stage and publish a failure-atomic DDE metadata transaction.
    ///
    /// Sources remain inert; publication never refreshes cached data.
    pub(crate) fn edit_dde<F>(&mut self, update: F) -> CoreResult<()>
    where
        F: FnOnce(&mut crate::dde::Edit) -> CoreResult<()>,
    {
        self.edit_dde_with_context(
            crate::dde::Limits::default(),
            &crate::dde::default_context(),
            update,
        )
    }

    /// Edit DDE metadata with explicit parsing, staging, and readback budgets.
    ///
    /// Archive replacement and facade rehydration use the package policy. The
    /// supplied context is checked after rehydration and before publication.
    pub(crate) fn edit_dde_with_context<F>(
        &mut self,
        limits: crate::dde::Limits,
        context: &litchi_core::ExecutionContext,
        update: F,
    ) -> CoreResult<()>
    where
        F: FnOnce(&mut crate::dde::Edit) -> CoreResult<()>,
    {
        let snapshot = self.dde_with(limits, context)?;
        let mut edit = snapshot.edit();
        update(&mut edit)?;
        let commit = edit.commit(context)?;
        if commit.changed() {
            self.ensure_dde_publication_allowed()?;
            let package = self
                .inner
                .package
                .replace_content_xml(commit.snapshot().source_xml())?;
            let candidate = Self::from_package(package)?;
            context.check().map_err(map_sheet_metadata_execution)?;
            *self = candidate;
        }
        Ok(())
    }

    fn ensure_dde_publication_allowed(&self) -> CoreResult<()> {
        if self
            .inner
            .package
            .package()
            .files()?
            .into_iter()
            .any(|path| litchi_odf_common::core::package::is_signature_owner_path(&path))
        {
            return Err(litchi_core::Error::Refused(
                litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::Protected,
                    "ods.facade",
                    "signed-source refusal: changed ODS DDE metadata requires explicit unsign/resign policy",
                ),
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
    pub fn scenarios(&self) -> FamilyResult<crate::scenario::Snapshot> {
        Ok(
            crate::scenario::Snapshot::parse(self.inner.package.content_xml()).map_err(
                |error| {
                    litchi_core::Error::InvalidFormat(format!(
                        "ODS scenario metadata inspection failed: {error}"
                    ))
                },
            )?,
        )
    }

    /// Stage and publish one failure-atomic scenario metadata edit.
    ///
    /// The closure edits typed declarations selected by exact worksheet name
    /// or source order.  A failed closure, stale source, invalid XML, or typed
    /// readback leaves this spreadsheet unchanged.
    pub(crate) fn edit_scenarios<F>(&mut self, update: F) -> CoreResult<()>
    where
        F: FnOnce(&mut crate::scenario::Edit) -> CoreResult<()>,
    {
        let snapshot = self.scenarios()?;
        let mut edit = snapshot.edit();
        update(&mut edit)?;
        let commit = edit.commit()?;
        if commit.changed() {
            let package = self
                .inner
                .package
                .replace_content_xml(commit.snapshot().source_xml())?;
            *self = Self::from_package(package)?;
        }
        Ok(())
    }

    /// Capture the source-backed consolidation, label-range, cell-source, and
    /// detective metadata catalog for this owned spreadsheet.
    pub fn sheet_metadata(&self) -> FamilyResult<crate::sheet_metadata::Snapshot> {
        Ok(crate::sheet_metadata::Snapshot::parse(
            self.inner.package.content_xml(),
        )?)
    }

    /// Capture sheet metadata under explicit finite limits and execution context.
    pub fn sheet_metadata_with(
        &self,
        limits: crate::sheet_metadata::Limits,
        context: &litchi_core::ExecutionContext,
    ) -> FamilyResult<crate::sheet_metadata::Snapshot> {
        Ok(crate::sheet_metadata::Snapshot::parse_with_context(
            self.inner.package.content_xml(),
            limits,
            context,
        )?)
    }

    /// Stage and publish one failure-atomic sheet metadata edit.
    pub(crate) fn edit_sheet_metadata<F>(&mut self, update: F) -> CoreResult<()>
    where
        F: FnOnce(&mut crate::sheet_metadata::Edit) -> CoreResult<()>,
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
    pub(crate) fn edit_sheet_metadata_with_context<F>(
        &mut self,
        limits: crate::sheet_metadata::Limits,
        context: &litchi_core::ExecutionContext,
        update: F,
    ) -> CoreResult<()>
    where
        F: FnOnce(&mut crate::sheet_metadata::Edit) -> CoreResult<()>,
    {
        let snapshot = self.sheet_metadata_with(limits, context)?;
        let mut edit = snapshot.edit();
        update(&mut edit)?;
        let commit = edit.commit(context)?;
        if commit.changed() {
            self.ensure_sheet_metadata_publication_allowed()?;
            let package = self
                .inner
                .package
                .replace_content_xml(commit.snapshot().source_xml())?;
            let candidate = Self::from_package(package)?;
            context.check().map_err(map_sheet_metadata_execution)?;
            *self = candidate;
        }
        Ok(())
    }

    fn ensure_sheet_metadata_publication_allowed(&self) -> CoreResult<()> {
        let signed = self
            .inner
            .package
            .package()
            .files()?
            .into_iter()
            .any(|path| litchi_odf_common::core::package::is_signature_owner_path(&path));
        if signed {
            return Err(litchi_core::Error::Refused(
                litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::Protected,
                    "ods.facade",
                    "signed-source refusal: changed ODS sheet metadata requires explicit unsign/resign policy",
                ),
            ));
        }
        Ok(())
    }

    /// Stage, validate, rebuild, and fully rehydrate one inert tracked-change edit.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub(crate) fn update_tracked_changes<F>(&mut self, edit: F) -> CoreResult<()>
    where
        F: FnOnce(&mut crate::tracked_changes::Edit) -> CoreResult<()>,
    {
        let snapshot = self.tracked_changes()?;
        let commit = crate::tracked_changes::update(&snapshot, edit)?;
        if !commit.changed() {
            return Ok(());
        }
        let package = self.inner.package.replace_tracked_changes(&commit)?;
        let candidate = Self::from_package(package)?;
        *self = candidate;
        Ok(())
    }

    /// Publish a validated annotation transaction without rebuilding an
    /// unchanged package.
    pub(crate) fn publish_annotations(&mut self, content_xml: &str) -> CoreResult<()> {
        let package = crate::annotations::replace_content(&self.inner.package, content_xml)?;
        *self = Self::from_package(package)?;
        Ok(())
    }

    /// Borrow the compact cross-format metadata projection.
    #[must_use]
    pub fn metadata(&self) -> &litchi_core::Metadata {
        self.inner.metadata.value()
    }

    /// Borrow the complete typed ODF metadata model.
    #[must_use]
    pub fn odf_metadata(&self) -> &crate::metadata::Metadata {
        self.inner.metadata.odf()
    }

    /// Borrow the retained metadata snapshot, including bounded source XML.
    #[must_use]
    pub fn metadata_snapshot(&self) -> &crate::metadata::Snapshot {
        &self.inner.metadata
    }

    /// Borrow spreadsheet calculation settings, if the document declares them.
    #[must_use]
    pub fn settings(&self) -> Option<&crate::settings::Settings> {
        self.inner.settings.as_ref()
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
    pub fn data_pilots(&self) -> FamilyResult<crate::data_pilot::Catalog<'_>> {
        Ok(crate::data_pilot::Catalog::load(&self.inner.package)?)
    }

    /// Capture the `DataPilot` owner as an immutable, exact-package snapshot.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn data_pilot_snapshot(&self) -> FamilyResult<crate::data_pilot::Snapshot> {
        Ok(crate::data_pilot::Snapshot::from_bytes_with_xml_limits(
            self.inner.package.package().as_bytes().to_vec(),
            self.inner.xml_limits,
        )?)
    }

    /// Discover inert database-range declarations owned by this spreadsheet.
    ///
    /// Database sources, filters, sorting, and subtotals are metadata only;
    /// this method never opens a database, refreshes a range, or evaluates a
    /// filter or subtotal.
    ///
    /// # Errors
    /// Returns an error when the content owner is malformed or over budget.
    pub fn database_ranges(&self) -> FamilyResult<crate::database_range::Catalog<'_>> {
        Ok(crate::database_range::Catalog::load(&self.inner.package)?)
    }

    /// Capture database-range metadata as an immutable, exact-package
    /// snapshot for explicit Snapshot → Edit → Commit → Patch workflows.
    ///
    /// # Errors
    /// Returns an error when the package or typed owner exceeds its finite
    /// admission limits.
    pub fn database_range_snapshot(&self) -> FamilyResult<crate::database_range::Snapshot> {
        Ok(crate::database_range::Snapshot::from_package(
            &self.inner.package,
        )?)
    }

    /// Alias with a plural noun for callers that use the owner name.
    pub fn database_ranges_snapshot(&self) -> FamilyResult<crate::database_range::Snapshot> {
        self.database_range_snapshot()
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
    pub(crate) fn edit_database_ranges<F>(&mut self, edit: F) -> CoreResult<()>
    where
        F: for<'source> FnOnce(&mut crate::database_range::Editor<'_, 'source>) -> CoreResult<()>,
    {
        let commit = {
            let catalog = self.database_ranges()?;
            let mut transaction = catalog.transaction();
            edit(&mut transaction.editor())?;
            transaction.commit()?
        };
        if commit.changed() {
            let package = crate::package::Archive::from_bytes_with_xml_limits(
                commit.into_owned_bytes(),
                self.inner.xml_limits,
            )?;
            *self = Self::from_package(package)?;
        }
        Ok(())
    }

    /// Return the typed worksheet graph in document order.
    #[must_use]
    pub fn sheets(&self) -> &[crate::worksheet::Sheet] {
        &self.inner.sheets
    }

    /// Capture worksheets as an immutable, exact-package snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the retained package or worksheet graph is invalid.
    pub fn worksheet_snapshot(&self) -> FamilyResult<crate::worksheet::Snapshot> {
        let package = self.inner.package.clone_without_password()?;
        Ok(
            crate::worksheet::Snapshot::from_shared_package_with_xml_limits(
                Arc::new(package),
                self.inner.xml_limits,
            )?,
        )
    }

    /// Discover embedded charts in content-level drawing order.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn charts(&self) -> FamilyResult<crate::charts::Inventory<'_>> {
        self.charts_with(crate::charts::Limits::default())
    }

    /// Discover embedded charts with an explicit resource budget.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn charts_with(
        &self,
        limits: crate::charts::Limits,
    ) -> FamilyResult<crate::charts::Inventory<'_>> {
        Ok(crate::charts::inventory(&self.inner.package, limits)?)
    }

    /// Capture embedded charts as an immutable, exact-package snapshot.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn chart_snapshot(&self) -> FamilyResult<crate::charts::Snapshot> {
        self.chart_snapshot_with(crate::charts::Limits::default())
    }

    /// Capture embedded charts with an explicit resource budget.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn chart_snapshot_with(
        &self,
        limits: crate::charts::Limits,
    ) -> FamilyResult<crate::charts::Snapshot> {
        Ok(crate::charts::Snapshot::from_bytes_with_xml_limits(
            self.inner.package.package().as_bytes().to_vec(),
            limits,
            self.inner.xml_limits,
        )?)
    }

    /// Select one embedded chart by exact drawing name or checked position.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn chart<'a, S>(&self, selector: S) -> FamilyResult<Option<crate::charts::Chart>>
    where
        S: Into<crate::charts::Selector<'a>>,
    {
        Ok(self.charts()?.get(selector).map(|chart| chart.cloned())?)
    }

    /// Find a worksheet by its exact ODF name.
    #[must_use]
    pub fn sheet(&self, name: &str) -> Option<&crate::worksheet::Sheet> {
        self.inner.sheets.iter().find(|sheet| sheet.name == name)
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
            .inner
            .sheets
            .iter()
            .position(|sheet| sheet.name == selector.sheet_name)?;
        let direct = || self.inner.sheets[sheet_index].cell_view(selector.row, selector.column);

        if let Some(locator) = self.inner.cell_locator.get() {
            return Some(locator.as_ref().map_or_else(direct, |locator| {
                locator.cell_view(
                    &self.inner.sheets,
                    sheet_index,
                    selector.row,
                    selector.column,
                )
            }));
        }

        let previous = self
            .inner
            .cell_queries
            .fetch_update(Ordering::Relaxed, Ordering::Relaxed, |count| {
                Some(count.saturating_add(1))
            })
            .unwrap_or(usize::MAX);
        if previous.saturating_add(1) >= cell_locator::BUILD_QUERY_THRESHOLD {
            let locator = self
                .inner
                .cell_locator
                .get_or_init(|| cell_locator::CellLocator::try_build(&self.inner.sheets));
            return Some(locator.as_ref().map_or_else(direct, |locator| {
                locator.cell_view(
                    &self.inner.sheets,
                    sheet_index,
                    selector.row,
                    selector.column,
                )
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
    ) -> FamilyResult<Vec<Option<crate::worksheet::CellView<'_>>>> {
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
    ) -> FamilyResult<Vec<Option<crate::worksheet::CellView<'_>>>> {
        self.cell_batch(selectors)
    }

    /// Discover package, inline, missing, and inert linked images.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn images(&self) -> FamilyResult<Vec<crate::media::Image>> {
        let package = self.inner.package.package().package()?;
        Ok(crate::media::scan_package(
            self.inner.package.content_xml(),
            self.inner.package.styles_xml(),
            &package,
        )?)
    }

    /// Inspect inert conditional-format, sparkline, hyperlink, and in-table
    /// drawing source metadata without evaluating or dereferencing it.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn source_features(&self) -> FamilyResult<crate::source_features::Snapshot> {
        crate::source_features::Snapshot::parse(self.inner.package.content_xml())
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
        crate::content_validation::Snapshot::parse(self.inner.package.content_xml())
    }

    /// Discover package, inline, missing, and inert linked embedded objects.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn embedded_objects(&self) -> FamilyResult<Vec<crate::embedded::Object>> {
        let package = self.inner.package.package().package()?;
        Ok(crate::embedded::scan_package(
            self.inner.package.content_xml(),
            self.inner.package.styles_xml(),
            &package,
        )?)
    }

    /// Return bytes only for inline or verified package-contained images.
    /// Linked and missing images remain inert and are never fetched.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn image_bytes(&self, image: &crate::media::Image) -> FamilyResult<Option<Vec<u8>>> {
        match &image.source {
            crate::media::Source::Inline { bytes, .. } => Ok(Some(bytes.clone())),
            crate::media::Source::PackagePart { path, .. } => {
                Ok(self.inner.package.package().get_file(path).map(Some)?)
            },
            crate::media::Source::MissingPackagePart { .. }
            | crate::media::Source::Linked { .. }
            | crate::media::Source::Missing
            | _ => Ok(None),
        }
    }

    #[must_use]
    pub(crate) fn into_bytes(self) -> Vec<u8> {
        self.inner.package.package().as_bytes().to_vec()
    }

    /// Return all global and sheet-local named definitions in document order.
    #[must_use]
    pub fn definitions(&self) -> &[Definition] {
        &self.inner.definitions
    }

    /// Capture named definitions as an immutable, exact-package snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the retained package cannot be reparsed.
    pub fn definitions_snapshot(&self) -> FamilyResult<crate::definitions::Snapshot> {
        Ok(crate::definitions::Snapshot::from_bytes_with_xml_limits(
            self.inner.package.package().as_bytes().to_vec(),
            self.inner.xml_limits,
        )?)
    }

    /// Return named ranges in their document order.
    pub fn ranges(&self) -> impl Iterator<Item = &Range> {
        self.inner
            .definitions
            .iter()
            .filter_map(|definition| match definition {
                Definition::Range(range) => Some(range),
                Definition::Expression(_) => None,
            })
    }

    /// Return named expressions in their document order.
    pub fn expressions(&self) -> impl Iterator<Item = &Expression> {
        self.inner
            .definitions
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
    pub(crate) fn add_range(&mut self, range: Range) -> CoreResult<()> {
        self.add_definition(range.into())
    }

    /// Atomically append a validated named expression.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub(crate) fn add_expression(&mut self, expression: Expression) -> CoreResult<()> {
        self.add_definition(expression.into())
    }

    /// Atomically append a validated named definition while preserving catalog order.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub(crate) fn add_definition(&mut self, definition: Definition) -> CoreResult<()> {
        let mut candidate = self.inner.definitions.clone();
        candidate.push(definition);
        self.set_definitions(candidate)
    }

    /// Atomically replace the complete ordered named-definition catalog.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub(crate) fn set_definitions(&mut self, definitions: Vec<Definition>) -> CoreResult<()> {
        let updated = crate::codec::names::replace(self.inner.package.content_xml(), &definitions)?;
        let package = self.inner.package.replace_content_xml(&updated)?;
        *self = Self::from_package(package)?;
        Ok(())
    }

    /// Publish a validated worksheet snapshot as one package transaction.
    pub(crate) fn publish_sheets(
        &mut self,
        sheets: Vec<crate::worksheet::Sheet>,
    ) -> CoreResult<()> {
        crate::worksheet::package::validate_owned_link_delta(&self.inner.sheets, &sheets)?;
        let package = self.inner.package.replace_sheets(&sheets)?;
        *self = Self::from_package(package)?;
        Ok(())
    }

    pub(crate) fn publish_metadata(&mut self, metadata: litchi_core::Metadata) -> CoreResult<()> {
        let package = self.inner.package.metadata_snapshot()?;
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
        let package = self
            .inner
            .package
            .replace_metadata_xml(Some(&metadata_xml))?;
        *self = Self::from_package(package)?;
        Ok(())
    }

    pub(crate) fn remove_metadata(&mut self) -> CoreResult<()> {
        let snapshot = self.inner.package.metadata_snapshot()?;
        let mut transaction = snapshot.transaction();
        transaction.remove();
        let commit = transaction.commit()?;
        if !commit.changed() {
            return Ok(());
        }
        let package = self.inner.package.replace_metadata_xml(None)?;
        *self = Self::from_package(package)?;
        Ok(())
    }

    pub(crate) fn publish_settings(
        &mut self,
        settings: Option<crate::settings::Settings>,
    ) -> CoreResult<()> {
        if self.inner.settings == settings {
            return Ok(());
        }
        let package = self
            .inner
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
    pub fn rdf_graphs(&self) -> FamilyResult<Vec<Graph>> {
        Ok(litchi_odf_common::rdf::graphs(
            self.inner.package.package(),
        )?)
    }

    /// Capture RDF graphs as an immutable, exact-package snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the package, manifest, or a declared graph is invalid.
    pub fn rdf_snapshot(&self) -> FamilyResult<crate::metadata_graphs::Snapshot> {
        Ok(
            crate::metadata_graphs::Snapshot::from_bytes_with_xml_limits(
                self.inner.package.package().as_bytes().to_vec(),
                self.inner.xml_limits,
            )?,
        )
    }

    /// Apply an exact-source reversible RDF graph patch.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale patch or invalid candidate package.
    pub(crate) fn apply_rdf_patch(
        &mut self,
        patch: &crate::metadata_graphs::Patch,
    ) -> CoreResult<()> {
        let commit = patch.apply(&self.rdf_snapshot()?)?;
        self.publish_rdf_commit(commit)
    }

    /// Add a graph and atomically replace this snapshot with the rebuilt package.
    ///
    /// # Errors
    ///
    /// Returns an error when the path, triples, compact XML, or rebuilt package is invalid.
    pub(crate) fn add_rdf_graph(
        &mut self,
        preferred_path: Option<&str>,
        triples: &[Triple],
    ) -> CoreResult<String> {
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
    pub(crate) fn replace_rdf_graph(&mut self, path: &str, triples: &[Triple]) -> CoreResult<()> {
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
    pub(crate) fn remove_rdf_graph(&mut self, path: &str) -> CoreResult<()> {
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
    pub(crate) fn add_rdf_triple(&mut self, path: &str, triple: &Triple) -> CoreResult<usize> {
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
    pub(crate) fn replace_rdf_triple(
        &mut self,
        path: &str,
        index: usize,
        triple: &Triple,
    ) -> CoreResult<()> {
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
    pub(crate) fn remove_rdf_triple(&mut self, path: &str, index: usize) -> CoreResult<()> {
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
    pub(crate) fn move_rdf_triple(&mut self, path: &str, from: usize, to: usize) -> CoreResult<()> {
        let snapshot = self.rdf_snapshot()?;
        let mut edit = snapshot.edit();
        edit.move_triple(
            path,
            litchi_core::Position::new(from),
            litchi_core::Position::new(to),
        )?;
        self.publish_rdf_commit(edit.commit())
    }

    fn publish_rdf_commit(&mut self, commit: crate::metadata_graphs::Commit) -> CoreResult<()> {
        if commit.changed() {
            let snapshot = commit.into_snapshot();
            let package = crate::package::Archive::from_bytes_with_xml_limits(
                snapshot.as_bytes().to_vec(),
                self.inner.xml_limits,
            )?;
            *self = Self::from_package(package)?;
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::builder::Workbook as Draft;
    use litchi_core::Error as CoreError;
    use litchi_odf_common::Error as PackageError;
    use std::path::Path;
    use std::sync::Arc;

    const ANNOTATED_CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:vendor="urn:example:vendor" office:version="1.3"><office:body><office:spreadsheet><vendor:keep/><table:table table:name="Data"><table:table-row><table:table-cell><office:annotation><text:p>existing</text:p></office:annotation></table:table-cell><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;

    #[test]
    fn builder_round_trips_through_facade() {
        let bytes = Builder::new()
            .build()
            .expect("test fixture or operation should succeed")
            .to_bytes()
            .expect("test fixture or operation should succeed");
        let spreadsheet =
            Workbook::from_bytes(bytes).expect("test fixture or operation should succeed");
        assert!(spreadsheet.content_xml().contains("office:spreadsheet"));
    }

    #[test]
    fn worksheet_snapshot_reuses_the_facade_package_index() {
        let bytes = Builder::new()
            .build()
            .expect("test fixture or operation should succeed")
            .to_bytes()
            .expect("test fixture or operation should succeed");
        let spreadsheet =
            Workbook::from_bytes(bytes).expect("test fixture or operation should succeed");
        let snapshot = spreadsheet
            .worksheet_snapshot()
            .expect("test fixture or operation should succeed");

        assert_eq!(
            snapshot.prepared_index_identity(),
            spreadsheet.prepared_index_identity()
        );
    }

    /// A conforming ODS opens through every facade with a present but empty
    /// tolerance report: nothing was tolerated, and nothing was recorded.
    #[test]
    fn conformant_fixture_opens_with_an_empty_tolerance_report() {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/odf/corpus/calc-two-sheets.ods");
        let bytes = std::fs::read(&path).expect("test fixture or operation should succeed");

        // Owned open and the prepared-detection handoff.
        let spreadsheet =
            Workbook::from_bytes(bytes.clone()).expect("test fixture or operation should succeed");
        assert!(spreadsheet.tolerance_report().is_clean());
        assert_eq!(spreadsheet.tolerance_report().total(), 0);
        let prepared =
            litchi_odf_common::detect::prepared(bytes, litchi_odf_common::SourceLimits::default())
                .expect("test fixture or operation should succeed");
        let spreadsheet = crate::Package::from_prepared_package(prepared)
            .and_then(|package| package.workbook())
            .expect("test fixture or operation should succeed");
        assert!(spreadsheet.tolerance_report().is_clean());

        // Path-based open and the positional source-backed facade.
        let spreadsheet = crate::Package::open(&path)
            .and_then(|package| package.workbook())
            .expect("test fixture or operation should succeed");
        assert!(spreadsheet.tolerance_report().is_clean());
        let source = Arc::new(litchi_core::OwnedSource::new(
            std::fs::read(&path).expect("test fixture or operation should succeed"),
        ));
        let source_backed = SourceBackedSpreadsheet::from_read_at(source)
            .expect("test fixture or operation should succeed");
        assert!(source_backed.tolerance_report().is_clean());
    }

    #[test]
    fn cell_locator_builds_at_the_threshold_and_preserves_snapshot_traits() {
        fn assert_send_sync<T: Send + Sync>() {}
        assert_send_sync::<Workbook>();

        let bytes = Builder::new()
            .content_xml(ANNOTATED_CONTENT)
            .build()
            .expect("test fixture or operation should succeed")
            .to_bytes()
            .expect("test fixture or operation should succeed");
        let spreadsheet =
            Workbook::from_bytes(bytes).expect("test fixture or operation should succeed");

        for _ in 1..cell_locator::BUILD_QUERY_THRESHOLD {
            assert!(matches!(
                spreadsheet.cell("Data", 0, 0),
                Some(crate::worksheet::CellView::Stored(_))
            ));
        }
        assert!(spreadsheet.inner.cell_locator.get().is_none());
        assert!(matches!(
            spreadsheet.cell("Data", 0, 0),
            Some(crate::worksheet::CellView::Stored(_))
        ));
        assert!(matches!(
            spreadsheet.inner.cell_locator.get(),
            Some(Some(_))
        ));
    }

    #[test]
    fn cell_batch_matches_scalar_order_missing_distinction_and_identity() {
        let bytes = Builder::new()
            .content_xml(ANNOTATED_CONTENT)
            .build()
            .expect("test fixture or operation should succeed")
            .to_bytes()
            .expect("test fixture or operation should succeed");
        let spreadsheet =
            Workbook::from_bytes(bytes).expect("test fixture or operation should succeed");
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
            .expect("test fixture or operation should succeed")
            .to_bytes()
            .expect("test fixture or operation should succeed");
        let spreadsheet =
            Workbook::from_bytes(bytes).expect("test fixture or operation should succeed");
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

        let bounded = Workbook::from_bytes(
            Builder::new()
                .content_xml(ANNOTATED_CONTENT)
                .build()
                .expect("test fixture or operation should succeed")
                .to_bytes()
                .expect("test fixture or operation should succeed"),
        )
        .expect("test fixture or operation should succeed");
        let selectors = vec![CellSelector::new("Data", 0, 0); MAX_CELL_SELECTORS + 1];
        let error = bounded.cell_batch(&selectors).expect_err("bound must fail");
        assert!(matches!(
            error,
            PackageError::Core(CoreError::InvalidFormat(message))
                if message.contains("selector safety limit")
        ));
        assert!(bounded.inner.cell_locator.get().is_none());
        assert_eq!(bounded.inner.cell_queries.load(Ordering::Relaxed), 0);
    }

    #[test]
    fn cell_batch_is_send_sync_and_concurrent_locator_build_is_shared() {
        fn assert_send_sync<T: Send + Sync>() {}
        assert_send_sync::<CellSelector<'static>>();

        let bytes = Builder::new()
            .content_xml(ANNOTATED_CONTENT)
            .build()
            .expect("test fixture or operation should succeed")
            .to_bytes()
            .expect("test fixture or operation should succeed");
        let spreadsheet = Arc::new(
            Workbook::from_bytes(bytes).expect("test fixture or operation should succeed"),
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
        assert!(matches!(
            spreadsheet.inner.cell_locator.get(),
            Some(Some(_))
        ));
    }

    #[test]
    fn concurrent_first_cell_locator_build_is_shared_and_identical() {
        let bytes = Builder::new()
            .content_xml(ANNOTATED_CONTENT)
            .build()
            .expect("test fixture or operation should succeed")
            .to_bytes()
            .expect("test fixture or operation should succeed");
        let spreadsheet = Arc::new(
            Workbook::from_bytes(bytes).expect("test fixture or operation should succeed"),
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
        assert!(matches!(
            spreadsheet.inner.cell_locator.get(),
            Some(Some(_))
        ));
    }

    #[test]
    fn facade_replacement_discards_built_cell_locator() {
        let bytes = Builder::new()
            .content_xml(ANNOTATED_CONTENT)
            .build()
            .expect("test fixture or operation should succeed")
            .to_bytes()
            .expect("test fixture or operation should succeed");
        let mut spreadsheet =
            Workbook::from_bytes(bytes).expect("test fixture or operation should succeed");
        for _ in 0..cell_locator::BUILD_QUERY_THRESHOLD {
            assert!(spreadsheet.cell("Data", 0, 0).is_some());
        }
        assert!(matches!(
            spreadsheet.inner.cell_locator.get(),
            Some(Some(_))
        ));

        let updated = ANNOTATED_CONTENT.replace("existing", "replacement");
        spreadsheet
            .publish_annotations(&updated)
            .expect("test fixture or operation should succeed");
        assert!(spreadsheet.inner.cell_locator.get().is_none());
        assert_eq!(spreadsheet.inner.cell_queries.load(Ordering::Relaxed), 0);
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
            .expect("test fixture or operation should succeed").to_bytes().expect("test fixture or operation should succeed");
        let spreadsheet =
            Workbook::from_bytes(bytes).expect("test fixture or operation should succeed");

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
            .expect("test fixture or operation should succeed")
            .to_bytes()
            .expect("test fixture or operation should succeed");
        let spreadsheet =
            Workbook::from_bytes(bytes.clone()).expect("test fixture or operation should succeed");
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

        let mut mutable = Workbook::from_bytes(bytes.clone())
            .map(Draft::from_workbook)
            .expect("test fixture or operation should succeed");
        mutable
            .edit_annotations(|transaction| {
                transaction.set("Data", 0, 1, crate::annotations::Annotation::new("added"))
            })
            .expect("test fixture or operation should succeed");
        let edited = Workbook::from_bytes(
            mutable
                .build()
                .expect("the draft builds")
                .as_bytes()
                .to_vec(),
        )
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

        let mut no_op = Workbook::from_bytes(bytes.clone())
            .map(Draft::from_workbook)
            .expect("test fixture or operation should succeed");
        no_op
            .edit_annotations(|_| Ok(()))
            .expect("test fixture or operation should succeed");
        assert_eq!(
            no_op.build().expect("the draft builds").as_bytes().to_vec(),
            bytes
        );
    }
}
