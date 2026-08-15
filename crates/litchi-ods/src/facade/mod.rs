//! Concise user-facing ODS entry points.

mod cell_locator;
mod source;

use litchi_core::Result;
use std::{
    path::Path,
    sync::{
        Arc, OnceLock,
        atomic::{AtomicUsize, Ordering},
    },
};

pub use crate::authoring::{Builder, MutableSpreadsheet};
use crate::model::names::{Definition, Expression, Range, Scope};
pub use litchi_odf_common::rdf::{Graph, Object, Subject, Triple};
pub use source::{ReadLimits, SourceBackedSpreadsheet};

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
        let package = crate::package::Package::from_bytes(bytes)?;
        Self::from_package(package)
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
        let definitions = package.definitions()?;
        let sheets = package.sheets()?;
        let metadata = package.metadata_snapshot()?;
        let settings = package.calculation_settings()?;
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

    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.package.styles_xml()
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
        crate::dde::Snapshot::parse(self.package.content_xml()).map_err(|error| {
            litchi_core::Error::InvalidFormat(format!(
                "ODS DDE metadata inspection failed: {error}"
            ))
        })
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
