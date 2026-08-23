//! Immutable XLSX reads backed by a caller-provided positional source.
//!
//! This facade intentionally does not adapt into [`super::Workbook`]: that
//! snapshot owns a mutable OPC graph, while this type must keep ordinary part
//! payloads deferred. It exposes only semantic catalog and worksheet reads.
//!
//! Managed [`ExecutionContext`] opens are supported for this read-only
//! facade. Focused source-backed editors retain managed
//! [`litchi_opc::PartData`] handles in their immutable snapshots and expose
//! their own managed constructors; owning-package patch application remains a
//! separate, explicitly fallible boundary when it would require detaching an
//! allocation from that reservation.

use std::collections::HashMap;
#[cfg(any(unix, windows))]
use std::path::Path;
use std::sync::{Arc, OnceLock};

#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::{
    ExecutionContext, ExecutionError, ReadAt, Selector as CoreSelector, SourceVersion,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    PackURI, PartData, PartView, ReadLimits, SourceBackedPackage, SourceCacheDiagnostics,
    SourceCacheLimits,
};
use litchi_sheet::{Area, At, Cell as Address, Rect};

use super::{DateSystem, Flavor, Selector, Visibility, WorksheetKind, codec};
use crate::cell::{Cell, Store, Text, View};
use crate::error::{Error, Result, allocation, invalid};
use crate::raw;

const CHARTSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
const STRICT_CHARTSHEET_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet";
const DIALOGSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
const STRICT_DIALOGSHEET_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/dialogsheet";
const MACROSHEET_REL: &str = "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet";
const INTL_MACROSHEET_REL: &str =
    "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet";
const CHARTSHEET_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";

fn check_execution(context: Option<&ExecutionContext>) -> Result<()> {
    let Some(context) = context else {
        return Ok(());
    };
    context.check().map_err(|error| {
        Error::Package(match error {
            ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
            error => litchi_opc::OpcError::Execution(error),
        })
    })
}

struct SourceSheetData {
    position: usize,
    name: String,
    name_key: Box<str>,
    kind: WorksheetKind,
    visibility: Visibility,
    part_uri: PackURI,
    cells: OnceLock<Store>,
}

struct SourceInner {
    package: SourceBackedPackage,
    execution: Option<ExecutionContext>,
    // Keep the mandatory workbook payload pinned for managed facades. Besides
    // avoiding a second root extraction, this makes an exact managed budget
    // boundary meaningful: the selected worksheet cannot evict the root while
    // callers still hold the workbook snapshot. Compatibility facades retain
    // the historical finite-cache behavior and do not pin this payload.
    _catalog_data: Option<PartData>,
    shared_strings_uri: Option<PackURI>,
    shared_strings: OnceLock<Box<[Text]>>,
    styles_uri: Option<PackURI>,
    styles: OnceLock<raw::styles::Catalog>,
    flavor: Flavor,
    date_system: DateSystem,
    active_sheet: Option<usize>,
    sheets: Box<[Arc<SourceSheetData>]>,
    /// Sheet positions sorted by their canonical, case-insensitive names.
    ///
    /// The mandatory workbook catalog rejects duplicate canonical names, so a
    /// sorted position index is equivalent to the previous linear search while
    /// avoiding a second allocation of every key string.
    sheet_name_order: Box<[usize]>,
}

/// Read-only XLSX catalog and worksheet access over a positional source.
///
/// Opening validates the OPC catalog, package relationships, workbook part,
/// and workbook-to-sheet graph. Worksheet, shared-string, and style payloads
/// are not materialized or parsed until a selected semantic read requires
/// them. ZIP central-directory discovery can still perform bounded physical
/// tail reads through the positional source; laziness here means deferred
/// logical part extraction and semantic parsing. The type has no edit or
/// output APIs.
#[derive(Clone)]
pub struct SourceBackedWorkbook {
    inner: Arc<SourceInner>,
}

/// A lifetime-free read-only worksheet handle from [`SourceBackedWorkbook`].
#[derive(Clone)]
pub struct SourceWorksheet {
    owner: Arc<SourceInner>,
    data: Arc<SourceSheetData>,
}

/// Owned semantic state at one coordinate in a source-backed worksheet.
///
/// This mirrors [`crate::cell::View`] without borrowing a lazily retained
/// worksheet store.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum SourceCellView {
    /// No cell record or covering merge exists at this coordinate.
    Missing,
    /// The coordinate is covered by a merge whose anchor is `range.start()`.
    Covered(Rect),
    /// One physical cell record is stored at this coordinate.
    Stored(Cell),
}

impl SourceCellView {
    /// Borrow the owned stored cell, when this coordinate owns one.
    #[must_use]
    pub const fn stored(&self) -> Option<&Cell> {
        match self {
            Self::Stored(cell) => Some(cell),
            Self::Missing | Self::Covered(_) => None,
        }
    }

    /// Covering merged range, if this is a non-anchor coordinate.
    #[must_use]
    pub const fn merge(&self) -> Option<Rect> {
        match self {
            Self::Covered(range) => Some(*range),
            Self::Missing | Self::Stored(_) => None,
        }
    }
}

/// One owning sparse cell returned by [`SourceWorksheet::cells`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourceCell {
    /// Checked worksheet coordinate of the stored cell.
    pub address: Address,
    /// Owned semantic cell state.
    pub cell: Cell,
}

impl SourceBackedWorkbook {
    /// Open an ordinary XLSX package from a regular filesystem path.
    ///
    /// The path is held through an open [`FileSource`]; the package is not
    /// slurped into memory and subsequent worksheet reads remain positional
    /// and lazy. Replacing the pathname after this call does not retarget the
    /// open source, while a detected metadata change returns a typed source
    /// version error on the next operation.
    #[cfg(any(unix, windows))]
    pub fn from_path(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_path_with_limits(path, ReadLimits::default())
    }

    /// Open a filesystem-backed XLSX package with explicit OPC limits.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits(path: impl AsRef<Path>, limits: ReadLimits) -> Result<Self> {
        Self::from_read_at_with_limits(file_source(path)?, limits)
    }

    /// Open a filesystem-backed XLSX package with an explicit deferred-Part
    /// cache policy.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_cache_limits(
        path: impl AsRef<Path>,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_read_at_with_cache_limits(file_source(path)?, cache_limits)
    }

    /// Open a filesystem-backed XLSX package with explicit read and cache
    /// policies.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits_and_cache_limits(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits(file_source(path)?, limits, cache_limits)
    }

    /// Open a filesystem-backed XLSX package with explicit read and execution
    /// policies.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_execution_context(file_source(path)?, limits, context)
    }

    /// Open a filesystem-backed XLSX package with explicit read and execution
    /// policies while retaining the default finite cache.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits_and_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_execution_context(file_source(path)?, limits, context)
    }

    /// Open a filesystem-backed XLSX package with explicit read, cache, and
    /// execution policies.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits_and_cache_limits_and_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits_and_execution_context(
            file_source(path)?,
            limits,
            cache_limits,
            context,
        )
    }

    /// Open a filesystem-backed XLSX package from a regular path.
    #[cfg(any(unix, windows))]
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_path(path)
    }

    /// Open a filesystem-backed XLSX package with explicit OPC limits.
    #[cfg(any(unix, windows))]
    pub fn open_with_limits(path: impl AsRef<Path>, limits: ReadLimits) -> Result<Self> {
        Self::from_path_with_limits(path, limits)
    }

    /// Open a filesystem-backed XLSX package with an explicit finite cache.
    #[cfg(any(unix, windows))]
    pub fn open_with_cache_limits(
        path: impl AsRef<Path>,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_path_with_cache_limits(path, cache_limits)
    }

    /// Open a filesystem-backed XLSX package with explicit read and cache
    /// policies.
    #[cfg(any(unix, windows))]
    pub fn open_with_limits_and_cache_limits(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_path_with_limits_and_cache_limits(path, limits, cache_limits)
    }

    /// Open a filesystem-backed XLSX package with explicit read and execution
    /// policies.
    #[cfg(any(unix, windows))]
    pub fn open_with_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_path_with_execution_context(path, limits, context)
    }

    /// Open a filesystem-backed XLSX package with explicit read and execution
    /// policies while retaining the default finite cache.
    #[cfg(any(unix, windows))]
    pub fn open_with_limits_and_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_path_with_limits_and_execution_context(path, limits, context)
    }

    /// Open a filesystem-backed XLSX package with explicit read, cache, and
    /// execution policies.
    #[cfg(any(unix, windows))]
    pub fn open_with_limits_and_cache_limits_and_execution_context(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_path_with_limits_and_cache_limits_and_execution_context(
            path,
            limits,
            cache_limits,
            context,
        )
    }

    /// Open an ordinary XLSX package from a caller-provided positional source.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default())
    }

    /// Open from a positional source with explicit OPC resource limits.
    pub fn from_read_at_with_limits(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_limits(
            source, limits,
        )?)
    }

    /// Open an XLSX source with an explicit finite deferred-Part cache policy.
    ///
    /// This compatibility constructor remains unmanaged: the cache is
    /// bounded by [`SourceCacheLimits`] but is not charged to a hierarchical
    /// execution budget. Use
    /// [`Self::from_read_at_with_limits_and_cache_limits_and_execution_context`]
    /// when the caller owns a managed budget.
    pub fn from_read_at_with_cache_limits(
        source: Arc<dyn ReadAt>,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_read_at_with_limits_and_cache_limits(source, ReadLimits::default(), cache_limits)
    }

    /// Open an XLSX source with explicit read and finite cache policies.
    pub fn from_read_at_with_limits_and_cache_limits(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_source_backed_package(
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits(
                source,
                limits,
                cache_limits,
            )?,
        )
    }

    /// Open an XLSX source with an explicit caller-owned execution context.
    ///
    /// The context checks cancellation before mandatory open reads and before
    /// every deferred semantic materialization. Retained and in-flight
    /// payloads are charged to its hierarchical memory budget. Managed
    /// [`litchi_opc::PartData`] handles never escape as unbudgeted `Arc`
    /// allocations through this read-only facade.
    pub fn from_read_at_with_execution_context(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source,
            limits,
            context.clone(),
        )?;
        Self::from_source_backed_package_with_execution_context(package, Some(context))
    }

    /// Open from a positional source with explicit read and execution
    /// policies while retaining the default finite cache.
    pub fn from_read_at_with_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_execution_context(source, limits, context)
    }

    /// Open an XLSX source with explicit read, cache, and execution policies.
    ///
    /// This is the fully explicit managed constructor. The read-only
    /// workbook/sheet facade retains only owned semantic values and bounded
    /// OPC cache handles; it does not detach managed payloads.
    pub fn from_read_at_with_limits_and_cache_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        let package =
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
                source,
                limits,
                cache_limits,
                context.clone(),
            )?;
        Self::from_source_backed_package_with_execution_context(package, Some(context))
    }

    /// Build the read-only XLSX facade from a validated deferred OPC package.
    pub fn from_source_backed_package(package: SourceBackedPackage) -> Result<Self> {
        Self::from_source_backed_package_with_execution_context(package, None)
    }

    fn from_source_backed_package_with_execution_context(
        package: SourceBackedPackage,
        execution: Option<ExecutionContext>,
    ) -> Result<Self> {
        check_execution(execution.as_ref())?;
        let retain_catalog = execution.is_some() || package.cache_diagnostics().budget_managed;
        let workbook = package.main_document_part()?;
        let flavor = codec::flavor(workbook.content_type()).ok_or_else(|| {
            invalid(format!(
                "main part '{}' has non-XLSX content type '{}'",
                workbook.partname(),
                workbook.content_type()
            ))
        })?;
        let catalog_bytes = workbook.data()?;
        let catalog = raw::parse_catalog(catalog_bytes.as_bytes())?;
        check_execution(execution.as_ref())?;
        let sheet_parts = validate_sheet_graph(&package, &workbook, &catalog.sheets)?;
        check_execution(execution.as_ref())?;
        let shared_strings_uri = validate_shared_strings(&package, &workbook)?;
        let styles_uri = validate_styles(&package, &workbook)?;
        check_execution(execution.as_ref())?;

        let active_sheet = (!catalog.sheets.is_empty()).then_some(catalog.active_sheet_index);
        let mut sheets = Vec::new();
        sheets
            .try_reserve_exact(catalog.sheets.len())
            .map_err(|source| allocation("source-backed workbook sheets", source))?;
        for (position, (sheet, part)) in catalog.sheets.into_iter().zip(sheet_parts).enumerate() {
            sheets.push(Arc::new(SourceSheetData {
                position,
                name_key: crate::sheet::key(&sheet.name),
                name: sheet.name,
                kind: part.kind,
                visibility: codec::visibility(sheet.visibility),
                part_uri: part.uri,
                cells: OnceLock::new(),
            }));
        }
        let sheets = sheets.into_boxed_slice();
        let mut sheet_name_order = (0..sheets.len()).collect::<Vec<_>>();
        sheet_name_order.sort_unstable_by(|left, right| {
            sheets[*left]
                .name_key
                .as_ref()
                .cmp(sheets[*right].name_key.as_ref())
                .then_with(|| left.cmp(right))
        });
        let sheet_name_order = sheet_name_order.into_boxed_slice();
        // The source can change after the mandatory root and relationship
        // reads, including while the semantic sheet metadata above is being
        // allocated. Do not publish a facade whose catalog was assembled from
        // a stale positional snapshot. This also covers an empty sheet list.
        package.source_version()?;
        check_execution(execution.as_ref())?;

        Ok(Self {
            inner: Arc::new(SourceInner {
                package,
                execution,
                _catalog_data: retain_catalog.then_some(catalog_bytes),
                shared_strings_uri,
                shared_strings: OnceLock::new(),
                styles_uri,
                styles: OnceLock::new(),
                flavor,
                date_system: if catalog.uses_1904_date_system {
                    DateSystem::Excel1904
                } else {
                    DateSystem::Excel1900
                },
                active_sheet,
                sheets,
                sheet_name_order,
            }),
        })
    }

    /// Workbook flavor derived from the mandatory workbook catalog.
    #[must_use]
    pub fn flavor(&self) -> Flavor {
        self.inner.flavor
    }

    /// Date serial system derived from the mandatory workbook catalog.
    #[must_use]
    pub fn date_system(&self) -> DateSystem {
        self.inner.date_system
    }

    /// Content-free deferred-Part cache diagnostics.
    #[must_use]
    pub fn cache_diagnostics(&self) -> SourceCacheDiagnostics {
        self.inner.package.cache_diagnostics()
    }

    /// Exact source identity and revision captured by this facade.
    pub fn source_version(&self) -> Result<SourceVersion> {
        self.inner.package.source_version().map_err(Into::into)
    }

    /// Number of logical workbook sheets.
    #[must_use]
    pub fn len(&self) -> usize {
        self.inner.sheets.len()
    }

    /// Whether the workbook catalog contains no sheets.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.inner.sheets.is_empty()
    }

    /// Iterate lightweight sheet handles without reading worksheet payloads.
    #[must_use]
    pub fn sheets(
        &self,
    ) -> impl ExactSizeIterator<Item = SourceWorksheet> + DoubleEndedIterator + '_ {
        self.inner
            .sheets
            .iter()
            .cloned()
            .map(|data| SourceWorksheet {
                owner: Arc::clone(&self.inner),
                data,
            })
    }

    /// Look up a sheet by developer-facing name or checked zero-based position.
    pub fn sheet<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Option<SourceWorksheet>> {
        self.inner.execution_check()?;
        let data = match selector.into() {
            CoreSelector::Position(position) => self.inner.sheets.get(position.get()).cloned(),
            CoreSelector::Name(name) => {
                let key = crate::sheet::key(&name);
                self.inner
                    .sheet_name_order
                    .binary_search_by(|&position| {
                        self.inner.sheets[position]
                            .name_key
                            .as_ref()
                            .cmp(key.as_ref())
                    })
                    .ok()
                    .and_then(|order| self.inner.sheets.get(self.inner.sheet_name_order[order]))
                    .cloned()
            },
            CoreSelector::Id(never) => match never {},
            _ => return Err(Error::UnsupportedSelector),
        };
        Ok(data.map(|data| SourceWorksheet {
            owner: Arc::clone(&self.inner),
            data,
        }))
    }

    /// Return the active sheet, if the workbook catalog contains one.
    #[must_use]
    pub fn active_sheet(&self) -> Option<SourceWorksheet> {
        let data = self
            .inner
            .active_sheet
            .and_then(|position| self.inner.sheets.get(position))
            .cloned()?;
        Some(SourceWorksheet {
            owner: Arc::clone(&self.inner),
            data,
        })
    }
}

#[cfg(any(unix, windows))]
fn file_source(path: impl AsRef<Path>) -> Result<Arc<dyn ReadAt>> {
    Ok(Arc::new(FileSource::open(path).map_err(|error| {
        Error::Package(litchi_opc::OpcError::from(error))
    })?))
}

impl SourceWorksheet {
    /// Developer-facing sheet name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.data.name
    }

    /// Checked zero-based workbook position.
    #[must_use]
    pub fn position(&self) -> usize {
        self.data.position
    }

    /// Semantic sheet kind resolved from its workbook relationship.
    #[must_use]
    pub fn kind(&self) -> WorksheetKind {
        self.data.kind
    }

    /// Retained visibility state.
    #[must_use]
    pub fn visibility(&self) -> &Visibility {
        &self.data.visibility
    }

    /// Whether this is the active sheet in its source-backed workbook catalog.
    #[must_use]
    pub fn is_active(&self) -> bool {
        self.owner.active_sheet == Some(self.data.position)
    }

    /// Look up an exact logical cell state, returning an owning semantic value.
    pub fn cell<'a>(&self, at: impl Into<At<'a>>) -> Result<SourceCellView> {
        let address = at.into().resolve()?;
        let value = match self.store()?.view(address) {
            View::Missing => SourceCellView::Missing,
            View::Covered(range) => SourceCellView::Covered(range),
            View::Stored(cell) => SourceCellView::Stored(cell.clone()),
        };
        // `Store::view` returns an owned clone, so the source may change
        // during that clone without being observed by the earlier store
        // check. Revalidate immediately before publishing the value.
        self.owner.package.source_version()?;
        self.owner.execution_check()?;
        Ok(value)
    }

    /// Read every stored cell selected by a checked range into owning values.
    ///
    /// The selected worksheet XML is loaded and parsed on first use. The
    /// returned cells are independent of the source and may outlive this
    /// worksheet handle.
    pub fn cells<'a>(&self, area: impl Into<Area<'a>>) -> Result<Vec<SourceCell>> {
        let range = area.into().resolve()?;
        let mut values = Vec::new();
        for (address, cell) in self.store()?.cells(range) {
            values
                .try_reserve(1)
                .map_err(|source| allocation("source-backed selected cells", source))?;
            values.push(SourceCell {
                address,
                cell: cell.clone(),
            });
        }
        // The collection owns cloned cells and may take arbitrarily long for
        // a sparse range. Final source/version and cancellation checks keep a
        // returned collection tied to the exact opened snapshot.
        self.owner.package.source_version()?;
        self.owner.execution_check()?;
        Ok(values)
    }

    /// Visit every stored cell selected by a checked range without cloning it.
    ///
    /// The callback receives the immutable semantic cell state owned by this
    /// worksheet's parsed source snapshot. A callback may copy a cell if it
    /// needs to retain it, but the ordinary full-scan path does not allocate a
    /// result vector or clone formulas, shared-string text, or unknown-cell
    /// diagnostics. Formula caches, shared strings, styles, and MCE-selected
    /// worksheet markup have already been validated by `Self::store`.
    ///
    /// Cancellation is checked between callbacks. The source version is
    /// checked after the complete visit, so a source mutation during the
    /// callback cannot publish a semantically stale scan as successful.
    pub fn visit_cells<'a, F>(&self, area: impl Into<Area<'a>>, mut visit: F) -> Result<usize>
    where
        F: FnMut(Address, &Cell) -> Result<()>,
    {
        let range = area.into().resolve()?;
        let store = self.store()?;
        let mut visited = 0usize;
        for (address, cell) in store.cells(range) {
            self.owner.execution_check()?;
            visit(address, cell)?;
            visited = visited
                .checked_add(1)
                .ok_or_else(|| invalid("source-backed cell visit count overflow"))?;
        }
        self.owner.package.source_version()?;
        self.owner.execution_check()?;
        Ok(visited)
    }

    /// Bounding rectangle of stored cell records.
    ///
    /// The worksheet payload is loaded and parsed on first use. `None` means
    /// that the selected worksheet has no explicit cell records.
    pub fn stored_extent(&self) -> Result<Option<Rect>> {
        let extent = self.store()?.extents().stored();
        self.owner.package.source_version()?;
        self.owner.execution_check()?;
        Ok(extent)
    }

    fn store(&self) -> Result<&Store> {
        self.owner.execution_check()?;
        if self.data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: self.data.name.clone(),
            });
        }
        // `SourceBackedPackage::part` checks the captured source version even
        // when this worksheet's parsed semantic store is already retained.
        // Never return cached semantic data after a caller-visible source
        // change has invalidated the snapshot.
        let part = self.owner.package.part(&self.data.part_uri)?;
        if let Some(store) = self.data.cells.get() {
            // A cached semantic Store must still honor the managed package's
            // cancellation and source-version checks. `PartView::data()` is
            // a cache hit when the bounded payload is retained and does not
            // detach its managed `PartData` reservation.
            let _payload = part.data()?;
            self.owner.execution_check()?;
            self.owner.package.source_version()?;
            return Ok(store);
        }

        let data = part.data()?;
        let parsed = raw::worksheet::parse(data.as_bytes(), || self.owner.shared_strings())?;
        self.owner.validate_styles(&parsed)?;
        self.owner.execution_check()?;
        self.owner.package.source_version()?;
        let _publish_result = self.data.cells.set(parsed);
        self.data.cells.get().ok_or_else(|| {
            invalid("source-backed worksheet cache initialization did not publish a value")
        })
    }
}

impl SourceInner {
    fn execution_check(&self) -> Result<()> {
        check_execution(self.execution.as_ref())
    }

    fn shared_strings(&self) -> Result<Option<&[Text]>> {
        self.execution_check()?;
        let Some(uri) = self.shared_strings_uri.as_ref() else {
            return Ok(None);
        };
        if let Some(strings) = self.shared_strings.get() {
            let _payload = self.package.part(uri)?.data()?;
            self.execution_check()?;
            self.package.source_version()?;
            return Ok(Some(strings));
        }

        let data = self.package.part(uri)?.data()?;
        let parsed = raw::strings::parse(data.as_bytes())?;
        self.execution_check()?;
        self.package.source_version()?;
        let _publish_result = self.shared_strings.set(parsed);
        self.shared_strings
            .get()
            .map(|strings| Some(strings.as_ref()))
            .ok_or_else(|| {
                invalid("source-backed shared-string cache initialization did not publish a value")
            })
    }

    fn style_count(&self) -> Result<u32> {
        self.execution_check()?;
        let Some(uri) = self.styles_uri.as_ref() else {
            return Ok(0);
        };
        if let Some(styles) = self.styles.get() {
            let _payload = self.package.part(uri)?.data()?;
            self.execution_check()?;
            self.package.source_version()?;
            return Ok(styles.len());
        }

        let data = self.package.part(uri)?.data()?;
        let parsed = raw::styles::parse(data.as_bytes())?;
        self.execution_check()?;
        self.package.source_version()?;
        let _publish_result = self.styles.set(parsed);
        self.styles
            .get()
            .map(raw::styles::Catalog::len)
            .ok_or_else(|| {
                invalid("source-backed style cache initialization did not publish a value")
            })
    }

    fn validate_styles(&self, store: &Store) -> Result<()> {
        if !store.entries().iter().any(|entry| entry.style.is_some())
            && !store
                .row_entries()
                .iter()
                .any(|entry| entry.properties.style.is_some())
            && !store
                .column_entries()
                .iter()
                .any(|entry| entry.properties.style.is_some())
        {
            return Ok(());
        }
        let len = self.style_count()?;
        if let Some(entry) = store
            .entries()
            .iter()
            .find(|entry| entry.style.is_some_and(|key| key >= len))
        {
            return Err(invalid(format!(
                "worksheet cell {} references shared style {}, but the workbook contains {len} cell formats",
                entry.address,
                entry.style.unwrap_or_default()
            )));
        }
        if let Some(entry) = store
            .row_entries()
            .iter()
            .find(|entry| entry.properties.style.is_some_and(|key| key >= len))
        {
            return Err(invalid(format!(
                "worksheet row {} references shared style {}, but the workbook contains {len} cell formats",
                entry.index,
                entry.properties.style.unwrap_or_default()
            )));
        }
        if let Some(entry) = store
            .column_entries()
            .iter()
            .find(|entry| entry.properties.style.is_some_and(|key| key >= len))
        {
            return Err(invalid(format!(
                "worksheet column {} references shared style {}, but the workbook contains {len} cell formats",
                entry.first,
                entry.properties.style.unwrap_or_default()
            )));
        }
        Ok(())
    }
}

pub(crate) struct SheetPart {
    pub(crate) kind: WorksheetKind,
    pub(crate) uri: PackURI,
}

pub(crate) fn validate_sheet_graph(
    package: &SourceBackedPackage,
    workbook: &PartView<'_>,
    sheets: &[raw::Sheet],
) -> Result<Vec<SheetPart>> {
    let mut parts = Vec::new();
    parts
        .try_reserve_exact(sheets.len())
        .map_err(|source| allocation("source-backed workbook sheet graph", source))?;
    let mut targets = HashMap::<PackURI, usize>::new();
    targets
        .try_reserve(sheets.len())
        .map_err(|source| allocation("source-backed worksheet target lookup", source))?;
    for (position, sheet) in sheets.iter().enumerate() {
        let relationship = workbook.rels().get(&sheet.relationship_id).ok_or_else(|| {
            invalid(format!(
                "sheet '{}' references missing relationship '{}'",
                sheet.name, sheet.relationship_id
            ))
        })?;
        if relationship.is_external() {
            return Err(invalid(format!(
                "sheet '{}' relationship cannot be external",
                sheet.name
            )));
        }
        let target = relationship.target_partname()?;
        let part = package.part(&target)?;
        let kind = match relationship.reltype() {
            rt::WORKSHEET | rt::STRICT_WORKSHEET => {
                require_content_type(sheet, part.content_type(), ct::SML_WORKSHEET)?;
                WorksheetKind::Worksheet
            },
            CHARTSHEET_REL | STRICT_CHARTSHEET_REL => {
                require_content_type(sheet, part.content_type(), CHARTSHEET_CONTENT_TYPE)?;
                WorksheetKind::Chart
            },
            DIALOGSHEET_REL | STRICT_DIALOGSHEET_REL => WorksheetKind::Dialog,
            MACROSHEET_REL | INTL_MACROSHEET_REL => WorksheetKind::Macro,
            _ => WorksheetKind::Unknown,
        };
        let uri = part.partname().clone();
        if let Some(previous) = targets.insert(uri.clone(), position) {
            return Err(invalid(format!(
                "sheet part '{uri}' is referenced by both '{}' and '{}'",
                sheets[previous].name, sheet.name
            )));
        }
        parts.push(SheetPart { kind, uri });
    }
    Ok(parts)
}

fn validate_shared_strings(
    package: &SourceBackedPackage,
    workbook: &PartView<'_>,
) -> Result<Option<PackURI>> {
    let mut found = None;
    for relationship in workbook.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::SHARED_STRINGS | rt::STRICT_SHARED_STRINGS
        )
    }) {
        if found.is_some() {
            return Err(invalid("workbook has multiple shared-string relationships"));
        }
        if relationship.is_external() {
            return Err(invalid("shared-string relationship cannot be external"));
        }
        let uri = relationship.target_partname()?;
        let part = package.part(&uri)?;
        if part.content_type() != ct::SML_SHARED_STRINGS {
            return Err(invalid(format!(
                "shared-string part has content type '{}', expected '{}'",
                part.content_type(),
                ct::SML_SHARED_STRINGS
            )));
        }
        found = Some(uri);
    }
    Ok(found)
}

fn validate_styles(
    package: &SourceBackedPackage,
    workbook: &PartView<'_>,
) -> Result<Option<PackURI>> {
    let mut found = None;
    for relationship in workbook
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), rt::STYLES | rt::STRICT_STYLES))
    {
        if found.is_some() {
            return Err(invalid("workbook has multiple styles relationships"));
        }
        if relationship.is_external() {
            return Err(invalid("styles relationship cannot be external"));
        }
        let uri = relationship.target_partname()?;
        let part = package.part(&uri)?;
        if part.content_type() != ct::SML_STYLES {
            return Err(invalid(format!(
                "styles part has content type '{}', expected '{}'",
                part.content_type(),
                ct::SML_STYLES
            )));
        }
        found = Some(uri);
    }
    Ok(found)
}

fn require_content_type(sheet: &raw::Sheet, actual: &str, expected: &str) -> Result<()> {
    if actual != expected {
        return Err(invalid(format!(
            "sheet '{}' has content type '{actual}', expected '{expected}'",
            sheet.name
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::io;
    use std::num::{NonZeroU64, NonZeroUsize};
    use std::sync::Arc;
    use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

    use litchi_core::{
        Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
        ReadAt, Resource, SourceVersion,
    };
    use litchi_opc::constants::content_type as ct;
    use litchi_opc::{OpcError, OpcPackage, PackURI, SourceCacheLimits};
    use soapberry_zip::office::StreamingArchiveWriter;

    use super::{ReadLimits, SourceBackedWorkbook, SourceCellView};
    use crate::{Cell, Error, Value};

    const FIRST_MARKER: &[u8] = b"source-backed-requested-first-sheet";
    const SECOND_MARKER: &[u8] = b"source-backed-unrequested-second-sheet";

    struct CountingSource {
        bytes: Vec<u8>,
        first_marker_offset: usize,
        marker_offset: usize,
        first_body_marker_reads: AtomicUsize,
        second_body_marker_reads: AtomicUsize,
        revision: AtomicU64,
    }

    impl CountingSource {
        fn new(bytes: Vec<u8>) -> Self {
            let first_marker_offset = bytes
                .windows(FIRST_MARKER.len())
                .position(|window| window == FIRST_MARKER)
                .expect("first worksheet marker is stored in archive");
            let marker_offset = bytes
                .windows(SECOND_MARKER.len())
                .position(|window| window == SECOND_MARKER)
                .expect("second worksheet marker is stored in archive");
            Self {
                bytes,
                first_marker_offset,
                marker_offset,
                first_body_marker_reads: AtomicUsize::new(0),
                second_body_marker_reads: AtomicUsize::new(0),
                revision: AtomicU64::new(0),
            }
        }

        fn changed(&self) {
            self.revision.fetch_add(1, Ordering::SeqCst);
        }
    }

    /// Source whose revision flips immediately before one selected version
    /// observation. The returned version is therefore already different on
    /// that observation, deterministically exercising a final-check race.
    struct VersionFlipSource {
        bytes: Vec<u8>,
        version_calls: AtomicUsize,
        revision: AtomicU64,
        flip_before_call: Option<usize>,
    }

    impl VersionFlipSource {
        fn new(bytes: Vec<u8>, flip_before_call: Option<usize>) -> Self {
            Self {
                bytes,
                version_calls: AtomicUsize::new(0),
                revision: AtomicU64::new(0),
                flip_before_call,
            }
        }
    }

    impl ReadAt for VersionFlipSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let offset = usize::try_from(offset)
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            let call = self.version_calls.fetch_add(1, Ordering::SeqCst) + 1;
            if self.flip_before_call == Some(call) {
                self.revision.store(1, Ordering::SeqCst);
            }
            Ok(SourceVersion::new(92, self.revision.load(Ordering::SeqCst)))
        }
    }

    impl ReadAt for CountingSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let offset = usize::try_from(offset)
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            let end = offset + count;
            if offset < self.first_marker_offset + FIRST_MARKER.len()
                && self.first_marker_offset < end
            {
                self.first_body_marker_reads.fetch_add(1, Ordering::SeqCst);
            }
            if offset < self.marker_offset + SECOND_MARKER.len() && self.marker_offset < end {
                self.second_body_marker_reads.fetch_add(1, Ordering::SeqCst);
            }
            output[..count].copy_from_slice(&self.bytes[offset..end]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(91, self.revision.load(Ordering::SeqCst)))
        }
    }

    fn source_backed_xlsx() -> Vec<u8> {
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                format!(
                    r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="{}"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="{}"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="{}"/></Types>"#,
                    ct::SML_SHEET_MAIN,
                    ct::SML_WORKSHEET,
                    ct::SML_WORKSHEET,
                )
                .as_bytes(),
            )
            .unwrap();
        writer
            .write_stored(
                "_rels/.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "xl/workbook.xml",
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="First" sheetId="1" r:id="rId1"/><sheet name="Second" sheetId="2" r:id="rId2"/></sheets></workbook>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "xl/_rels/workbook.xml.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "xl/worksheets/sheet1.xml",
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><!--source-backed-requested-first-sheet--><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>"#,
            )
            .unwrap();
        let padding = "x".repeat(128 * 1024);
        let second = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><!--{SECOND_MARKER_TEXT}{padding}--><sheetData><row r="1"><c r="A1"><v>9</v></c></row></sheetData></worksheet>"#,
            SECOND_MARKER_TEXT = std::str::from_utf8(SECOND_MARKER).unwrap(),
        );
        writer
            .write_stored("xl/worksheets/sheet2.xml", second.as_bytes())
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn empty_source_backed_xlsx() -> Vec<u8> {
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                format!(
                    r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="{}"/></Types>"#,
                    ct::SML_SHEET_MAIN,
                )
                .as_bytes(),
            )
            .unwrap();
        writer
            .write_stored(
                "_rels/.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "xl/workbook.xml",
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "xl/_rels/workbook.xml.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#,
            )
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn part_len(bytes: &[u8], member: &str) -> u64 {
        let package = OpcPackage::from_bytes(bytes).unwrap();
        package
            .get_part(&PackURI::new(member).unwrap())
            .unwrap()
            .blob()
            .len() as u64
    }

    fn managed_context(memory: u64) -> (Budget, CancellationSource, ExecutionContext) {
        let budget = Budget::root(
            "xlsx-source-facade-test",
            Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let (cancellation_source, cancellation) = CancellationSource::pair();
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(memory.max(1)).unwrap(),
            0,
        )
        .unwrap();
        let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
        (budget, cancellation_source, context)
    }

    #[test]
    fn deferred_catalog_and_first_sheet_do_not_reach_second_worksheet_body_marker() {
        let source = Arc::new(CountingSource::new(source_backed_xlsx()));
        let workbook = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();

        assert_eq!(workbook.len(), 2);
        assert_eq!(
            workbook
                .sheets()
                .map(|sheet| sheet.name().to_owned())
                .collect::<Vec<_>>(),
            ["First", "Second"]
        );
        // This marker is inside the second worksheet body. It demonstrates
        // that catalog listing and first-sheet semantic reads do not extract
        // or parse that body. It does not claim ZIP indexing made no physical
        // reads of every byte range belonging to the unrequested member.
        assert_eq!(source.second_body_marker_reads.load(Ordering::SeqCst), 0);

        let first = workbook.sheet("First").unwrap().unwrap();
        assert!(matches!(
            first.cell("A1").unwrap(),
            SourceCellView::Stored(Cell::Value(Value::Number(ref value))) if value.as_str() == "7"
        ));
        assert_eq!(source.second_body_marker_reads.load(Ordering::SeqCst), 0);

        let selected = first.cells("A1:B2").unwrap();
        assert_eq!(selected.len(), 1);
        assert_eq!(selected[0].address.to_string(), "A1");
        assert_eq!(source.second_body_marker_reads.load(Ordering::SeqCst), 0);

        let second = workbook.sheet("Second").unwrap().unwrap();
        assert!(matches!(
            second.cell("A1").unwrap(),
            SourceCellView::Stored(Cell::Value(Value::Number(ref value))) if value.as_str() == "9"
        ));
        assert!(source.second_body_marker_reads.load(Ordering::SeqCst) > 0);
    }

    #[test]
    fn source_changes_are_returned_as_typed_opc_errors() {
        let source = Arc::new(CountingSource::new(source_backed_xlsx()));
        let workbook = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
        let sheet = workbook.sheet("First").unwrap().unwrap();
        assert!(matches!(
            sheet.cell("A1").unwrap(),
            SourceCellView::Stored(Cell::Value(Value::Number(ref value))) if value.as_str() == "7"
        ));
        source.changed();

        assert!(matches!(
            sheet.cell("A1"),
            Err(Error::Package(OpcError::SourceChanged { .. }))
        ));
    }

    #[test]
    fn final_catalog_check_rejects_revision_flip_after_empty_metadata_build() {
        let bytes = empty_source_backed_xlsx();
        let baseline_source = Arc::new(VersionFlipSource::new(bytes.clone(), None));
        let baseline = SourceBackedWorkbook::from_read_at(baseline_source.clone()).unwrap();
        assert!(baseline.is_empty());
        let final_check_call = baseline_source.version_calls.load(Ordering::SeqCst);
        drop(baseline);

        let source = Arc::new(VersionFlipSource::new(bytes, Some(final_check_call)));
        assert!(matches!(
            SourceBackedWorkbook::from_read_at(source),
            Err(Error::Package(OpcError::SourceChanged { .. }))
        ));
    }

    #[test]
    fn final_cell_and_collection_checks_reject_revision_flips_after_owned_clones() {
        let bytes = source_backed_xlsx();
        let baseline_source = Arc::new(VersionFlipSource::new(bytes.clone(), None));
        let baseline = SourceBackedWorkbook::from_read_at(baseline_source.clone()).unwrap();
        let baseline_sheet = baseline.sheet("First").unwrap().unwrap();
        baseline_sheet.cell("A1").unwrap();
        let cell_check_call = baseline_source.version_calls.load(Ordering::SeqCst);
        drop(baseline_sheet);
        drop(baseline);

        let source = Arc::new(VersionFlipSource::new(bytes.clone(), Some(cell_check_call)));
        let workbook = SourceBackedWorkbook::from_read_at(source).unwrap();
        let sheet = workbook.sheet("First").unwrap().unwrap();
        assert!(matches!(
            sheet.cell("A1"),
            Err(Error::Package(OpcError::SourceChanged { .. }))
        ));
        drop(sheet);
        drop(workbook);

        let baseline_source = Arc::new(VersionFlipSource::new(bytes.clone(), None));
        let baseline = SourceBackedWorkbook::from_read_at(baseline_source.clone()).unwrap();
        let baseline_sheet = baseline.sheet("First").unwrap().unwrap();
        baseline_sheet.cells("A1:B2").unwrap();
        let cells_check_call = baseline_source.version_calls.load(Ordering::SeqCst);
        drop(baseline_sheet);
        drop(baseline);

        let source = Arc::new(VersionFlipSource::new(bytes, Some(cells_check_call)));
        let workbook = SourceBackedWorkbook::from_read_at(source).unwrap();
        let sheet = workbook.sheet("First").unwrap().unwrap();
        assert!(matches!(
            sheet.cells("A1:B2"),
            Err(Error::Package(OpcError::SourceChanged { .. }))
        ));
    }

    #[test]
    fn managed_facade_selectively_materializes_with_exact_budget_and_releases_on_drop() {
        let archive = source_backed_xlsx();
        let workbook_bytes = part_len(&archive, "/xl/workbook.xml");
        let first_sheet_bytes = part_len(&archive, "/xl/worksheets/sheet1.xml");
        let exact = workbook_bytes + first_sheet_bytes;
        let source = Arc::new(CountingSource::new(archive));
        let (budget, _cancellation_source, context) = managed_context(exact);
        let cache_limits = SourceCacheLimits::new(usize::try_from(exact).unwrap(), 4).unwrap();
        let workbook =
            SourceBackedWorkbook::from_read_at_with_limits_and_cache_limits_and_execution_context(
                source.clone(),
                ReadLimits::default(),
                cache_limits,
                context,
            )
            .unwrap();

        assert_eq!(workbook.source_version().unwrap().revision(), 0);
        assert_eq!(workbook.cache_diagnostics().successful_loads, 1);
        assert!(workbook.cache_diagnostics().budget_managed);
        assert_eq!(budget.used(Resource::Memory), workbook_bytes);
        assert_eq!(source.first_body_marker_reads.load(Ordering::SeqCst), 0);
        assert_eq!(source.second_body_marker_reads.load(Ordering::SeqCst), 0);

        let first = workbook.sheet("First").unwrap().unwrap();
        assert!(matches!(
            first.cell("A1").unwrap(),
            SourceCellView::Stored(Cell::Value(Value::Number(ref value))) if value.as_str() == "7"
        ));
        assert!(source.first_body_marker_reads.load(Ordering::SeqCst) > 0);
        assert_eq!(source.second_body_marker_reads.load(Ordering::SeqCst), 0);
        let diagnostics = workbook.cache_diagnostics();
        assert_eq!(diagnostics.successful_loads, 2);
        assert_eq!(diagnostics.retained_bytes as u64, exact);
        assert_eq!(diagnostics.budget_cache_reserved_bytes, exact);
        assert_eq!(budget.used(Resource::Memory), exact);

        drop(first);
        drop(workbook);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_facade_one_under_budget_rejects_before_selected_part_io() {
        let archive = source_backed_xlsx();
        let workbook_bytes = part_len(&archive, "/xl/workbook.xml");
        let first_sheet_bytes = part_len(&archive, "/xl/worksheets/sheet1.xml");
        let exact = workbook_bytes + first_sheet_bytes;
        let source = Arc::new(CountingSource::new(archive));
        let (budget, _cancellation_source, context) = managed_context(exact - 1);
        let cache_limits = SourceCacheLimits::new(usize::try_from(exact).unwrap(), 4).unwrap();
        let workbook =
            SourceBackedWorkbook::from_read_at_with_limits_and_cache_limits_and_execution_context(
                source.clone(),
                ReadLimits::default(),
                cache_limits,
                context,
            )
            .unwrap();

        assert_eq!(workbook.cache_diagnostics().successful_loads, 1);
        assert_eq!(budget.used(Resource::Memory), workbook_bytes);
        let first = workbook.sheet("First").unwrap().unwrap();
        let result = first.cell("A1");
        assert!(matches!(
            result,
            Err(Error::Package(OpcError::Execution(
                ExecutionError::ResourceLimit(_)
            )))
        ));
        assert_eq!(source.first_body_marker_reads.load(Ordering::SeqCst), 0);
        assert_eq!(source.second_body_marker_reads.load(Ordering::SeqCst), 0);
        assert_eq!(workbook.cache_diagnostics().successful_loads, 1);
        assert!(workbook.cache_diagnostics().budget_reservation_failures > 0);
        drop(first);
        drop(workbook);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn managed_facade_cancellation_is_typed_before_open_and_on_cached_reads() {
        let archive = source_backed_xlsx();
        let source = Arc::new(CountingSource::new(archive.clone()));
        let (budget, cancellation_source, context) = managed_context(u64::MAX);
        cancellation_source.cancel();
        assert!(matches!(
            SourceBackedWorkbook::from_read_at_with_execution_context(
                source.clone(),
                ReadLimits::default(),
                context,
            ),
            Err(Error::Package(OpcError::Cancelled))
        ));
        assert_eq!(source.first_body_marker_reads.load(Ordering::SeqCst), 0);
        assert_eq!(source.second_body_marker_reads.load(Ordering::SeqCst), 0);
        assert_eq!(budget.used(Resource::Memory), 0);

        let source = Arc::new(CountingSource::new(archive));
        let (budget, cancellation_source, context) = managed_context(u64::MAX);
        let workbook = SourceBackedWorkbook::from_read_at_with_execution_context(
            source.clone(),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let first = workbook.sheet("First").unwrap().unwrap();
        assert!(first.cell("A1").is_ok());
        let reads_before_cancel = source.first_body_marker_reads.load(Ordering::SeqCst);
        cancellation_source.cancel();
        assert!(matches!(
            first.cell("A1"),
            Err(Error::Package(OpcError::Cancelled))
        ));
        assert_eq!(
            source.first_body_marker_reads.load(Ordering::SeqCst),
            reads_before_cancel
        );
        drop(first);
        drop(workbook);
        assert_eq!(budget.used(Resource::Memory), 0);
    }
}
