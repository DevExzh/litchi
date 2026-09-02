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
use std::io::Read;
#[cfg(any(unix, windows))]
use std::path::Path;
use std::sync::{Arc, OnceLock};

#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::{
    ExecutionContext, ExecutionError, ReadAt, Selector as CoreSelector, SourceVersion,
};
use litchi_ooxml_common::mce::{Capabilities, Error as MceError, StreamError, StreamLimits};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    PackURI, PartData, PartView, ReadLimits, SourceBackedPackage, SourceCacheDiagnostics,
    SourceCacheLimits, VerifiedDecodedReaderError,
};
use litchi_sheet::{Area, At, Cell as Address, Rect};

use super::{DateSystem, Flavor, Selector, Visibility, WorksheetKind, codec};
use crate::cell::{Cell, Store, Text, Value, View};
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
const MCE_HARD_MAX_BYTES: u64 = 2 * 1024 * 1024 * 1024;

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

    /// Open an XLSX package from a sequential reader with the default
    /// bounded OPC policy.
    ///
    /// The reader is consumed into the source-backed owner's bounded byte
    /// storage. Workbook metadata is validated at open, while worksheet and
    /// other ordinary payloads remain deferred until a selected read asks for
    /// them. The resulting owned source has stable freshness identity for the
    /// lifetime of this immutable snapshot.
    pub fn from_reader<R: Read>(reader: R) -> Result<Self> {
        Self::from_reader_with_limits(reader, ReadLimits::default())
    }

    /// Open an XLSX package from a sequential reader with explicit OPC limits.
    ///
    /// Input ingestion and all later deferred payload reads use the same
    /// bounded source-backed package policy; unselected worksheets remain
    /// cold until explicitly selected.
    pub fn from_reader_with_limits<R: Read>(reader: R, limits: ReadLimits) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_reader_with_limits(
            reader, limits,
        )?)
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
        let mut sheet_name_order = Vec::new();
        sheet_name_order
            .try_reserve_exact(sheets.len())
            .map_err(|source| allocation("source-backed workbook sheet name order", source))?;
        for position in 0..sheets.len() {
            sheet_name_order.push(position);
        }
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
        let value = if self.data.cells.get().is_some() || self.data.kind != WorksheetKind::Worksheet
        {
            self.eager_cell(address)
        } else {
            self.stream_cell(address)
        };
        self.finish_result(value)
    }

    fn eager_cell(&self, address: Address) -> Result<SourceCellView> {
        Ok(match self.store()?.view(address) {
            View::Missing => SourceCellView::Missing,
            View::Covered(range) => SourceCellView::Covered(range),
            View::Stored(cell) => SourceCellView::Stored(cell.clone()),
        })
    }

    #[expect(
        clippy::result_large_err,
        reason = "The stream error intentionally retains typed primary plus raw/active callback diagnostics; boxing it would change the established API."
    )]
    fn stream_cell(&self, address: Address) -> Result<SourceCellView> {
        self.owner.execution_check()?;
        let outcome = {
            let part = self.owner.package.part(&self.data.part_uri)?;
            let declared = part.declared_uncompressed_size()?;
            match selected_stream_limits(declared) {
                Some(limits) => {
                    let capabilities = Capabilities::default();
                    Some(
                        part.with_verified_decoded_reader(|reader| {
                            raw::selected_worksheet::scan(reader, &capabilities, &limits, address)
                        })
                        .map_err(map_verified_reader_error)?,
                    )
                },
                None => None,
            }
        };
        let Some(outcome) = outcome else {
            return self.eager_cell(address);
        };
        match outcome {
            raw::selected_worksheet::ScanOutcome::Eligible(selected) => {
                let dependencies = selected.dependencies;
                let mut shared_text = None;
                let mut fallback = false;

                if let Some(max_index) = dependencies.max_shared_string_index {
                    match self.owner.stream_shared_strings_dependency(
                        max_index,
                        dependencies.target_shared_string_index,
                    )? {
                        SelectedDependency::Ready(text) => shared_text = text,
                        SelectedDependency::Fallback => fallback = true,
                    }
                }
                if !fallback && let Some(max_index) = dependencies.max_direct_style_index {
                    if matches!(
                        self.owner.stream_styles_dependency(max_index)?,
                        SelectedDependency::Fallback
                    ) {
                        fallback = true;
                    }
                }

                // Dependency readers must be fully released before a stale
                // source can trigger eager materialization. Keep this fence
                // in addition to the publication fence in `cell`.
                self.owner.package.source_version()?;
                self.owner.execution_check()?;
                if fallback {
                    return self.eager_cell(address);
                }

                match (selected.cell, dependencies.target_shared_string_index) {
                    (Some(cell), _) => Ok(SourceCellView::Stored(cell)),
                    (None, Some(_)) => match shared_text {
                        Some(text) => Ok(SourceCellView::Stored(Cell::Value(Value::Text(text)))),
                        None => self.eager_cell(address),
                    },
                    (None, None) => Ok(SourceCellView::Missing),
                }
            },
            raw::selected_worksheet::ScanOutcome::NotEligible(_) => {
                self.owner.package.source_version()?;
                self.owner.execution_check()?;
                self.eager_cell(address)
            },
        }
    }

    /// Read every stored cell selected by a checked range into owning values.
    ///
    /// A cold worksheet uses the bounded sparse selection route when its
    /// structure is eligible; unsupported worksheet semantics fall back to
    /// the materialized store. The returned cells are independent of the
    /// source and may outlive this worksheet handle.
    pub fn cells<'a>(&self, area: impl Into<Area<'a>>) -> Result<Vec<SourceCell>> {
        let range = area.into().resolve()?;
        let values =
            if self.data.cells.get().is_some() || self.data.kind != WorksheetKind::Worksheet {
                self.eager_cells(range)
            } else {
                self.stream_cells(range)
            };
        self.finish_result(values)
    }

    /// Visit every stored cell selected by a checked range.
    ///
    /// The range is first read into the same verified owning values returned by
    /// [`Self::cells`]. Callbacks then receive references into that local
    /// vector, so no callback runs while a source reader is active.
    ///
    /// Cancellation is checked before each callback. Final source and
    /// execution fences run even when a callback fails, so a source mutation
    /// or cancellation remains primary over that callback error.
    pub fn visit_cells<'a, F>(&self, area: impl Into<Area<'a>>, mut visit: F) -> Result<usize>
    where
        F: FnMut(Address, &Cell) -> Result<()>,
    {
        let values = match self.cells(area) {
            Ok(values) => values,
            Err(error) => return self.finish_result(Err(error)),
        };
        let mut visited = 0usize;
        let mut result = Ok(());
        for source_cell in &values {
            if let Err(error) = self.owner.execution_check() {
                result = Err(error);
                break;
            }
            if let Err(error) = visit(source_cell.address, &source_cell.cell) {
                result = Err(error);
                break;
            }
            match visited.checked_add(1) {
                Some(count) => visited = count,
                None => {
                    result = Err(invalid("source-backed cell visit count overflow"));
                    break;
                },
            }
        }
        self.finish_result(result.map(|()| visited))
    }

    fn eager_cells(&self, range: Rect) -> Result<Vec<SourceCell>> {
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
        Ok(values)
    }

    #[expect(
        clippy::result_large_err,
        reason = "The selected stream error intentionally retains typed primary plus raw/active callback diagnostics; boxing it would change the established API."
    )]
    fn stream_cells(&self, range: Rect) -> Result<Vec<SourceCell>> {
        self.owner.execution_check()?;
        let outcome = {
            let part = self.owner.package.part(&self.data.part_uri)?;
            let declared = part.declared_uncompressed_size()?;
            match selected_stream_limits(declared) {
                Some(limits) => {
                    let capabilities = Capabilities::default();
                    Some(
                        part.with_verified_decoded_reader(|reader| {
                            raw::selected_worksheet::scan_range(
                                reader,
                                &capabilities,
                                &limits,
                                range,
                            )
                        })
                        .map_err(map_verified_reader_error)?,
                    )
                },
                None => None,
            }
        };
        let Some(outcome) = outcome else {
            return self.eager_cells(range);
        };
        let selected = match outcome {
            raw::selected_worksheet::RangeScanOutcome::Eligible(selected) => selected,
            raw::selected_worksheet::RangeScanOutcome::NotEligible(_) => {
                // The verified worksheet reader must be gone before the
                // materialized fallback can inspect the source again.
                self.owner.package.source_version()?;
                self.owner.execution_check()?;
                return self.eager_cells(range);
            },
        };

        let dependencies = selected.dependencies;
        let mut requested = Vec::new();
        requested
            .try_reserve_exact(selected.cells.len())
            .map_err(|source| allocation("source-backed selected shared-string indexes", source))?;
        let mut fallback = false;
        for record in &selected.cells {
            let Some(index) = record.shared_string_index else {
                continue;
            };
            let Ok(index) = usize::try_from(index) else {
                fallback = true;
                break;
            };
            requested.push(index);
        }
        requested.sort_unstable();
        requested.dedup();

        let mut shared_text = None;
        if !fallback {
            if let Some(max_index) = dependencies.max_shared_string_index {
                match self
                    .owner
                    .stream_shared_strings_dependencies(max_index, &requested)?
                {
                    SelectedDependency::Ready(text) => shared_text = Some(text),
                    SelectedDependency::Fallback => fallback = true,
                }
            } else if !requested.is_empty() {
                fallback = true;
            }
        }
        if !fallback && let Some(max_index) = dependencies.max_direct_style_index {
            if matches!(
                self.owner.stream_styles_dependency(max_index)?,
                SelectedDependency::Fallback
            ) {
                fallback = true;
            }
        }

        if !fallback {
            for record in &selected.cells {
                let Some(index) = record.shared_string_index else {
                    continue;
                };
                let Ok(index) = usize::try_from(index) else {
                    fallback = true;
                    break;
                };
                let Some(text) = shared_text.as_ref() else {
                    fallback = true;
                    break;
                };
                if text
                    .binary_search_by_key(&index, |(candidate, _)| *candidate)
                    .is_err()
                {
                    fallback = true;
                    break;
                }
            }
        }

        // Dependency readers must be fully released before a stale source can
        // trigger eager materialization. Keep this fence in addition to the
        // publication fence in `cells`.
        self.owner.package.source_version()?;
        self.owner.execution_check()?;
        if fallback {
            return self.eager_cells(range);
        }

        let mut values = Vec::new();
        values
            .try_reserve_exact(selected.cells.len())
            .map_err(|source| allocation("source-backed selected cells", source))?;
        for record in selected.cells {
            let cell = match (record.cell, record.shared_string_index) {
                (Some(cell), None) => cell,
                (None, Some(index)) => {
                    let index = usize::try_from(index)
                        .map_err(|_error| invalid("shared-string index exceeds this platform"))?;
                    let text = shared_text
                        .as_ref()
                        .and_then(|text| {
                            text.binary_search_by_key(&index, |(candidate, _)| *candidate)
                                .ok()
                                .and_then(|position| text.get(position))
                        })
                        .map(|(_, text)| text.clone())
                        .ok_or_else(|| {
                            invalid("selected shared-string dependency was not retained")
                        })?;
                    Cell::Value(Value::Text(text))
                },
                (Some(_), Some(_)) => {
                    return Err(invalid(
                        "selected worksheet record has both a semantic cell and shared-string dependency",
                    ));
                },
                (None, None) => {
                    return Err(invalid(
                        "selected worksheet record has neither a semantic cell nor dependency",
                    ));
                },
            };
            values.push(SourceCell {
                address: record.address,
                cell,
            });
        }
        Ok(values)
    }

    fn finish_result<T>(&self, result: Result<T>) -> Result<T> {
        let source: Result<()> = self
            .owner
            .package
            .source_version()
            .map(|_| ())
            .map_err(Into::into);
        let execution = self.owner.execution_check();
        source.and(execution).and(result)
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
            // deliberately not requested here: the semantic Store is already
            // the validated worksheet snapshot, and its retained value must
            // remain usable even when the bounded payload cache evicted or
            // bypassed the original worksheet bytes.
            self.owner.package.check_execution()?;
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

fn selected_stream_limits(declared: u64) -> Option<StreamLimits> {
    if declared == 0 || declared > MCE_HARD_MAX_BYTES {
        return None;
    }
    let declared = usize::try_from(declared).ok()?;
    let mut limits = StreamLimits::default();
    limits.processing.max_input_bytes = declared;
    limits.processing.max_output_bytes = declared;
    Some(limits)
}

enum SelectedDependency<T> {
    Ready(T),
    Fallback,
}

fn map_verified_reader_error(
    error: VerifiedDecodedReaderError<StreamError<Error, Error>>,
) -> Error {
    match error {
        VerifiedDecodedReaderError::Opc { error, .. } => Error::Package(error),
        VerifiedDecodedReaderError::Callback(error) => map_selected_stream_error(error),
        _ => invalid("unknown verified OPC decoded reader error"),
    }
}

fn map_selected_stream_error(error: StreamError<Error, Error>) -> Error {
    match error {
        StreamError::Input {
            raw_error: Some(raw_error),
            ..
        }
        | StreamError::Mce {
            raw_error: Some(raw_error),
            ..
        }
        | StreamError::Callback {
            raw_error: Some(raw_error),
            ..
        } => raw_error,
        StreamError::Input { error, .. } => Error::Package(litchi_opc::OpcError::IoError(error)),
        StreamError::Mce { error, .. } => map_selected_mce_error(error),
        StreamError::Callback {
            raw_error: None,
            active_error: Some(active_error),
        } => active_error,
        StreamError::Callback {
            raw_error: None,
            active_error: None,
        } => invalid("MCE stream callback failure without an observer error"),
        _ => invalid("unknown MCE stream error"),
    }
}

fn map_selected_mce_error(error: MceError) -> Error {
    match error {
        MceError::Xml(message) => invalid(format!("invalid worksheet extension XML: {message}")),
        error => Error::MarkupCompatibility(error),
    }
}

impl SourceInner {
    fn execution_check(&self) -> Result<()> {
        // The package owns the execution context used by managed source
        // caches. Always check it first so compatibility facades built around
        // an already-managed package cannot bypass cancellation or budget
        // policy merely because their optional facade context is absent.
        self.package.check_execution()?;
        // Constructors that create both layers retain the same logical
        // context in each handle; checking the optional facade handle as well
        // also covers callers that supply a distinct policy at this boundary.
        check_execution(self.execution.as_ref())
    }

    fn stream_shared_strings_dependency(
        &self,
        max_index: u32,
        target_index: Option<u32>,
    ) -> Result<SelectedDependency<Option<Text>>> {
        self.execution_check()?;
        let target_index = match target_index {
            Some(index) => match usize::try_from(index) {
                Ok(index) => Some(index),
                Err(_) => return Ok(SelectedDependency::Fallback),
            },
            None => None,
        };
        let mut requested = Vec::new();
        if let Some(index) = target_index {
            requested
                .try_reserve_exact(1)
                .map_err(|source| allocation("source-backed shared-string indexes", source))?;
            requested.push(index);
        }
        match self.stream_shared_strings_dependencies(max_index, &requested)? {
            SelectedDependency::Ready(mut selected) => Ok(SelectedDependency::Ready(
                selected.pop().map(|(_, text)| text),
            )),
            SelectedDependency::Fallback => Ok(SelectedDependency::Fallback),
        }
    }

    #[expect(
        clippy::result_large_err,
        reason = "The dependency stream error intentionally retains typed primary plus raw/active callback diagnostics; boxing it would change the established API."
    )]
    fn stream_shared_strings_dependencies(
        &self,
        max_index: u32,
        requested: &[usize],
    ) -> Result<SelectedDependency<Vec<(usize, Text)>>> {
        self.execution_check()?;
        let Some(uri) = self.shared_strings_uri.as_ref() else {
            return Ok(SelectedDependency::Fallback);
        };
        let max_index = match usize::try_from(max_index) {
            Ok(index) => index,
            Err(_) => return Ok(SelectedDependency::Fallback),
        };
        let selected = {
            let part = self.package.part(uri)?;
            let declared = part.declared_uncompressed_size()?;
            let Some(limits) = selected_stream_limits(declared) else {
                return Ok(SelectedDependency::Fallback);
            };
            let capabilities = Capabilities::default();
            part.with_verified_decoded_reader(|reader| {
                raw::strings::stream_selected(reader, &capabilities, &limits, requested)
            })
            .map_err(map_verified_reader_error)?
        };

        if selected.unsupported_rich
            || max_index >= selected.count
            || selected.requested.len() != requested.len()
            || selected
                .requested
                .iter()
                .zip(requested)
                .any(|((index, _), requested)| *index != *requested)
        {
            return Ok(SelectedDependency::Fallback);
        }
        Ok(SelectedDependency::Ready(selected.requested))
    }

    #[expect(
        clippy::result_large_err,
        reason = "The dependency stream error intentionally retains typed primary plus raw/active callback diagnostics; boxing it would change the established API."
    )]
    fn stream_styles_dependency(&self, max_index: u32) -> Result<SelectedDependency<()>> {
        self.execution_check()?;
        let Some(uri) = self.styles_uri.as_ref() else {
            return Ok(SelectedDependency::Fallback);
        };
        let count = {
            let part = self.package.part(uri)?;
            let declared = part.declared_uncompressed_size()?;
            let Some(limits) = selected_stream_limits(declared) else {
                return Ok(SelectedDependency::Fallback);
            };
            let capabilities = Capabilities::default();
            part.with_verified_decoded_reader(|reader| {
                raw::styles::stream_count(reader, &capabilities, &limits)
            })
            .map_err(map_verified_reader_error)?
        };

        if max_index >= count {
            Ok(SelectedDependency::Fallback)
        } else {
            Ok(SelectedDependency::Ready(()))
        }
    }

    fn shared_strings(&self) -> Result<Option<&[Text]>> {
        self.execution_check()?;
        let Some(uri) = self.shared_strings_uri.as_ref() else {
            return Ok(None);
        };
        if let Some(strings) = self.shared_strings.get() {
            let _part = self.package.part(uri)?;
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
            let _part = self.package.part(uri)?;
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
    use litchi_opc::{OpcError, OpcPackage, PackURI, SourceBackedPackage, SourceCacheLimits};
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
    fn cached_worksheet_queries_do_not_reload_an_evicted_payload() {
        let source = Arc::new(CountingSource::new(source_backed_xlsx()));
        let cache_limits = SourceCacheLimits::new(256 * 1024, 1).unwrap();
        let workbook = SourceBackedWorkbook::from_read_at_with_limits_and_cache_limits(
            source.clone(),
            ReadLimits::default(),
            cache_limits,
        )
        .unwrap();

        let first = workbook.sheet("First").unwrap().unwrap();
        assert!(first.stored_extent().unwrap().is_some());
        assert!(matches!(
            first.cell("A1").unwrap(),
            SourceCellView::Stored(Cell::Value(Value::Number(ref value)))
                if value.as_str() == "7"
        ));
        let first_reads_after_preload = source.first_body_marker_reads.load(Ordering::SeqCst);
        assert!(first_reads_after_preload > 0);

        // Loading the second worksheet evicts the first worksheet's bounded
        // PartData payload, while the parsed semantic Store remains retained
        // by the worksheet handle.
        let second = workbook.sheet("Second").unwrap().unwrap();
        assert!(matches!(
            second.cell("A1").unwrap(),
            SourceCellView::Stored(Cell::Value(Value::Number(ref value)))
                if value.as_str() == "9"
        ));
        assert!(workbook.cache_diagnostics().evictions > 0);

        let first_reads_before_cached_query = source.first_body_marker_reads.load(Ordering::SeqCst);
        assert!(matches!(
            first.cell("A1").unwrap(),
            SourceCellView::Stored(Cell::Value(Value::Number(ref value)))
                if value.as_str() == "7"
        ));
        assert_eq!(
            source.first_body_marker_reads.load(Ordering::SeqCst),
            first_reads_before_cached_query,
            "a retained semantic worksheet must not reload evicted PartData"
        );
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
        assert!(first.stored_extent().unwrap().is_some());
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
        let result = first.stored_extent();
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

    #[test]
    fn managed_package_facade_checks_package_cancellation_on_cached_reads() {
        let archive = source_backed_xlsx();
        let source = Arc::new(CountingSource::new(archive));
        let (budget, cancellation_source, context) = managed_context(u64::MAX);
        let cache_limits = SourceCacheLimits::new(256 * 1024, 4).unwrap();
        let package =
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
                source.clone(),
                ReadLimits::default(),
                cache_limits,
                context,
            )
            .unwrap();
        let workbook = SourceBackedWorkbook::from_source_backed_package(package).unwrap();
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
            reads_before_cancel,
            "a cancelled cached read must not reload worksheet payload"
        );
        drop(first);
        drop(workbook);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    mod streaming_0363_tests {
        use std::io::Cursor;
        use std::sync::Arc;
        use std::sync::atomic::Ordering;

        use litchi_opc::OpcError;
        use litchi_opc::constants::content_type as ct;
        use soapberry_zip::office::StreamingArchiveWriter;

        use super::{
            CountingSource, ReadLimits, SourceBackedWorkbook, SourceCellView, managed_context,
            source_backed_xlsx,
        };
        use crate::{Cell, Error, Value};

        const SPREADSHEETML_NAMESPACE: &str =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        fn one_sheet_xlsx(worksheet: &str) -> Vec<u8> {
            let mut writer = StreamingArchiveWriter::new();
            writer
                .write_stored(
                    "[Content_Types].xml",
                    format!(
                        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="{}"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="{}"/></Types>"#,
                        ct::SML_SHEET_MAIN,
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
                    br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
                )
                .unwrap();
            writer
                .write_stored(
                    "xl/_rels/workbook.xml.rels",
                    br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
                )
                .unwrap();
            writer
                .write_stored("xl/worksheets/sheet1.xml", worksheet.as_bytes())
                .unwrap();
            writer.finish_to_bytes().unwrap()
        }

        #[test]
        fn eligible_scalar_cell_queries_0363_preserve_sparse_states_without_store() {
            let bytes = one_sheet_xlsx(&format!(
                r#"<worksheet xmlns="{SPREADSHEETML_NAMESPACE}"><sheetData><row r="1"><c r="A1"><v>7</v></c><c r="B1"/></row></sheetData></worksheet>"#,
            ));
            let workbook = SourceBackedWorkbook::from_reader(Cursor::new(bytes)).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let cold = workbook.cache_diagnostics();

            assert!(sheet.data.cells.get().is_none());
            assert!(matches!(
                sheet.cell("A1").unwrap(),
                SourceCellView::Stored(Cell::Value(Value::Number(ref value)))
                    if value.as_str() == "7"
            ));
            assert!(sheet.data.cells.get().is_none());
            assert!(matches!(sheet.cell("C1").unwrap(), SourceCellView::Missing));
            assert!(sheet.data.cells.get().is_none());
            assert!(matches!(
                sheet.cell("B1").unwrap(),
                SourceCellView::Stored(Cell::Empty)
            ));
            assert!(sheet.data.cells.get().is_none());

            let after = workbook.cache_diagnostics();
            assert_eq!(after.successful_loads, cold.successful_loads);
            assert_eq!(after.retained_bytes, cold.retained_bytes);
        }

        #[test]
        fn eligible_scalar_cell_queries_0363_rescan_without_cache() {
            let source = Arc::new(CountingSource::new(source_backed_xlsx()));
            let workbook = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
            let sheet = workbook.sheet("First").unwrap().unwrap();
            let cold = workbook.cache_diagnostics();

            assert!(sheet.data.cells.get().is_none());
            assert!(matches!(
                sheet.cell("A1").unwrap(),
                SourceCellView::Stored(Cell::Value(Value::Number(ref value)))
                    if value.as_str() == "7"
            ));
            let first_reads = source.first_body_marker_reads.load(Ordering::SeqCst);
            assert!(first_reads > 0);
            assert!(sheet.data.cells.get().is_none());
            assert_eq!(
                workbook.cache_diagnostics().successful_loads,
                cold.successful_loads
            );

            assert!(matches!(
                sheet.cell("A1").unwrap(),
                SourceCellView::Stored(Cell::Value(Value::Number(ref value)))
                    if value.as_str() == "7"
            ));
            assert!(source.first_body_marker_reads.load(Ordering::SeqCst) > first_reads);
            assert!(sheet.data.cells.get().is_none());
            assert_eq!(
                workbook.cache_diagnostics().successful_loads,
                cold.successful_loads
            );
        }

        #[test]
        fn merge_not_eligible_0363_falls_back_and_initializes_store() {
            let bytes = one_sheet_xlsx(&format!(
                r#"<worksheet xmlns="{SPREADSHEETML_NAMESPACE}"><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>"#,
            ));
            let workbook = SourceBackedWorkbook::from_reader(Cursor::new(bytes)).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let cold = workbook.cache_diagnostics();
            assert!(sheet.data.cells.get().is_none());

            assert!(matches!(
                sheet.cell("B1").unwrap(),
                SourceCellView::Covered(range) if range.a1() == "A1:B1"
            ));
            assert!(sheet.data.cells.get().is_some());
            assert!(matches!(
                sheet.cell("A1").unwrap(),
                SourceCellView::Stored(Cell::Value(Value::Number(ref value)))
                    if value.as_str() == "7"
            ));
            assert!(workbook.cache_diagnostics().successful_loads > cold.successful_loads);
        }

        #[test]
        fn malformed_tail_0363_does_not_publish_store() {
            let bytes = one_sheet_xlsx(&format!(
                r#"<worksheet xmlns="{SPREADSHEETML_NAMESPACE}"><sheetData><row r="1"><c r="A1"><v>7</v></c>"#,
            ));
            let workbook = SourceBackedWorkbook::from_reader(Cursor::new(bytes)).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let cold = workbook.cache_diagnostics();

            assert!(sheet.cell("A1").is_err());
            assert!(sheet.data.cells.get().is_none());
            assert_eq!(
                workbook.cache_diagnostics().successful_loads,
                cold.successful_loads
            );
        }

        #[test]
        fn crc_failure_0363_does_not_publish_store() {
            let mut bytes = source_backed_xlsx();
            let value = b"<v>7</v>";
            let offset = bytes
                .windows(value.len())
                .position(|window| window == value)
                .expect("first worksheet value is present");
            bytes[offset + 3] = b'8';
            let workbook = SourceBackedWorkbook::from_reader(Cursor::new(bytes)).unwrap();
            let sheet = workbook.sheet("First").unwrap().unwrap();
            let cold = workbook.cache_diagnostics();

            assert!(sheet.cell("A1").is_err());
            assert!(sheet.data.cells.get().is_none());
            assert_eq!(
                workbook.cache_diagnostics().successful_loads,
                cold.successful_loads
            );
        }

        #[test]
        fn source_change_0363_does_not_publish_store() {
            let source = Arc::new(CountingSource::new(source_backed_xlsx()));
            let workbook = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
            let sheet = workbook.sheet("First").unwrap().unwrap();
            let cold = workbook.cache_diagnostics();
            source.changed();

            assert!(matches!(
                sheet.cell("A1"),
                Err(Error::Package(OpcError::SourceChanged { .. }))
            ));
            assert!(sheet.data.cells.get().is_none());
            assert_eq!(
                workbook.cache_diagnostics().successful_loads,
                cold.successful_loads
            );
        }

        #[test]
        fn cancellation_0363_does_not_publish_store() {
            let source = Arc::new(CountingSource::new(source_backed_xlsx()));
            let (_budget, cancellation_source, context) = managed_context(u64::MAX);
            let workbook = SourceBackedWorkbook::from_read_at_with_execution_context(
                source,
                ReadLimits::default(),
                context,
            )
            .unwrap();
            let sheet = workbook.sheet("First").unwrap().unwrap();
            let cold = workbook.cache_diagnostics();
            cancellation_source.cancel();

            assert!(matches!(
                sheet.cell("A1"),
                Err(Error::Package(OpcError::Cancelled))
            ));
            assert!(sheet.data.cells.get().is_none());
            assert_eq!(
                workbook.cache_diagnostics().successful_loads,
                cold.successful_loads
            );
        }
    }

    mod streaming_0364_tests {
        use std::io::Cursor;
        use std::sync::Arc;
        use std::sync::atomic::Ordering;

        use litchi_opc::OpcError;
        use litchi_opc::SourceCacheLimits;
        use litchi_opc::constants::content_type as ct;
        use soapberry_zip::office::StreamingArchiveWriter;

        use super::{
            CountingSource, ReadLimits, SourceBackedWorkbook, SourceCellView, VersionFlipSource,
            managed_context,
        };
        use crate::{Cell, Error, Value};

        const SPREADSHEETML_NAMESPACE: &str =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const SHARED_STRINGS_RELATIONSHIP: &str =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
        const STYLES_RELATIONSHIP: &str =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

        pub(super) fn dependency_xlsx(
            worksheet: &str,
            shared_strings: Option<&str>,
            styles: Option<&str>,
        ) -> Vec<u8> {
            let shared_override = shared_strings.map_or_else(String::new, |_| {
                format!(
                    r#"<Override PartName="/xl/sharedStrings.xml" ContentType="{}"/>"#,
                    ct::SML_SHARED_STRINGS,
                )
            });
            let styles_override = styles.map_or_else(String::new, |_| {
                format!(
                    r#"<Override PartName="/xl/styles.xml" ContentType="{}"/>"#,
                    ct::SML_STYLES,
                )
            });
            let shared_relationship = shared_strings.map_or_else(String::new, |_| {
                format!(
                    r#"<Relationship Id="rId2" Type="{SHARED_STRINGS_RELATIONSHIP}" Target="sharedStrings.xml"/>"#,
                )
            });
            let styles_relationship = styles.map_or_else(String::new, |_| {
                format!(
                    r#"<Relationship Id="rId3" Type="{STYLES_RELATIONSHIP}" Target="styles.xml"/>"#,
                )
            });

            let mut writer = StreamingArchiveWriter::new();
            writer
                .write_stored(
                    "[Content_Types].xml",
                    format!(
                        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="{}"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="{}"/>{shared_override}{styles_override}</Types>"#,
                        ct::SML_SHEET_MAIN,
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
                    br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
                )
                .unwrap();
            writer
                .write_stored(
                    "xl/_rels/workbook.xml.rels",
                    format!(
                        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>{shared_relationship}{styles_relationship}</Relationships>"#,
                    )
                    .as_bytes(),
                )
                .unwrap();
            writer
                .write_stored("xl/worksheets/sheet1.xml", worksheet.as_bytes())
                .unwrap();
            if let Some(shared_strings) = shared_strings {
                writer
                    .write_stored("xl/sharedStrings.xml", shared_strings.as_bytes())
                    .unwrap();
            }
            if let Some(styles) = styles {
                writer
                    .write_stored("xl/styles.xml", styles.as_bytes())
                    .unwrap();
            }
            writer.finish_to_bytes().unwrap()
        }

        fn worksheet(cells: &str) -> String {
            let first = std::str::from_utf8(super::FIRST_MARKER).unwrap();
            let second = std::str::from_utf8(super::SECOND_MARKER).unwrap();
            format!(
                r#"<worksheet xmlns="{SPREADSHEETML_NAMESPACE}"><!--{first} {second}--><sheetData><row r="1">{cells}</row></sheetData></worksheet>"#,
            )
        }

        pub(super) fn plain_shared_strings(values: &[&str]) -> String {
            let items = values
                .iter()
                .map(|value| format!(r#"<si><t>{value}</t></si>"#))
                .collect::<String>();
            format!(
                r#"<sst xmlns="{SPREADSHEETML_NAMESPACE}" count="{count}" uniqueCount="{count}">{items}</sst>"#,
                count = values.len(),
            )
        }

        pub(super) fn styles(count: usize) -> String {
            let formats = "<xf/>".repeat(count);
            format!(
                r#"<styleSheet xmlns="{SPREADSHEETML_NAMESPACE}"><cellXfs count="{count}">{formats}</cellXfs></styleSheet>"#,
            )
        }

        fn rich_shared_strings() -> String {
            format!(
                r#"<sst xmlns="{SPREADSHEETML_NAMESPACE}" count="1" uniqueCount="1"><si><r><rPr/><t>rich </t></r><r><t>text</t></r></si></sst>"#,
            )
        }

        fn assert_text(view: SourceCellView, expected: &str) {
            assert!(matches!(
                view,
                SourceCellView::Stored(Cell::Value(Value::Text(ref value)))
                    if value.as_str() == expected
            ));
        }

        fn assert_error_without_store(bytes: Vec<u8>) {
            let workbook = SourceBackedWorkbook::from_reader(Cursor::new(bytes)).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            assert!(sheet.cell("A1").is_err());
            assert!(sheet.data.cells.get().is_none());
        }

        #[test]
        fn plain_shared_string_and_style_stream_0364_avoids_materialization() {
            let bytes = dependency_xlsx(
                &worksheet(r#"<c r="A1" t="s" s="0"><v>0</v></c>"#),
                Some(&plain_shared_strings(&["plain"])),
                Some(&styles(1)),
            );
            let source = Arc::new(CountingSource::new(bytes));
            let workbook = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let cold = workbook.cache_diagnostics();

            assert_text(sheet.cell("A1").unwrap(), "plain");
            assert!(sheet.data.cells.get().is_none());
            assert!(workbook.inner.shared_strings.get().is_none());
            assert!(workbook.inner.styles.get().is_none());
            let after = workbook.cache_diagnostics();
            assert_eq!(after.cold_loads, cold.cold_loads);
            assert_eq!(after.successful_loads, cold.successful_loads);
            assert!(source.first_body_marker_reads.load(Ordering::SeqCst) > 0);
        }

        #[test]
        fn unselected_dependency_maxima_0364_are_validated() {
            let valid = dependency_xlsx(
                &worksheet(r#"<c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c>"#),
                Some(&plain_shared_strings(&["selected", "unselected"])),
                None,
            );
            let workbook = SourceBackedWorkbook::from_reader(Cursor::new(valid)).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            assert_text(sheet.cell("A1").unwrap(), "selected");
            assert!(sheet.data.cells.get().is_none());

            let invalid = dependency_xlsx(
                &worksheet(r#"<c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c>"#),
                Some(&plain_shared_strings(&["selected"])),
                None,
            );
            let workbook = SourceBackedWorkbook::from_reader(Cursor::new(invalid)).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            assert!(sheet.cell("A1").is_err());
            assert!(sheet.data.cells.get().is_none());
        }

        #[test]
        fn requested_plain_and_rich_sst_0364_have_eager_parity() {
            let plain = dependency_xlsx(
                &worksheet(r#"<c r="A1" t="s"><v>0</v></c>"#),
                Some(&plain_shared_strings(&["rich text"])),
                None,
            );
            let workbook = SourceBackedWorkbook::from_reader(Cursor::new(plain)).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            assert_text(sheet.cell("A1").unwrap(), "rich text");
            assert!(sheet.data.cells.get().is_none());
            assert!(workbook.inner.shared_strings.get().is_none());

            let rich = dependency_xlsx(
                &worksheet(r#"<c r="A1" t="s"><v>0</v></c>"#),
                Some(&rich_shared_strings()),
                None,
            );
            let workbook = SourceBackedWorkbook::from_reader(Cursor::new(rich)).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            assert_text(sheet.cell("A1").unwrap(), "rich text");
            assert!(sheet.data.cells.get().is_some());
            assert!(workbook.inner.shared_strings.get().is_some());
        }

        #[test]
        fn invalid_and_missing_dependency_refs_0364_fall_back_without_store() {
            assert_error_without_store(dependency_xlsx(
                &worksheet(r#"<c r="A1" t="s"><v>0</v></c>"#),
                None,
                None,
            ));
            assert_error_without_store(dependency_xlsx(
                &worksheet(r#"<c r="A1" t="s"><v>1</v></c>"#),
                Some(&plain_shared_strings(&["only zero"])),
                None,
            ));
            assert_error_without_store(dependency_xlsx(
                &worksheet(r#"<c r="A1" s="0"><v>7</v></c>"#),
                None,
                None,
            ));
            assert_error_without_store(dependency_xlsx(
                &worksheet(r#"<c r="A1" s="1"><v>7</v></c>"#),
                None,
                Some(&styles(1)),
            ));
        }

        #[test]
        fn late_dependency_errors_0364_are_primary_without_publication() {
            let late_shared = format!(
                r#"<sst xmlns="{SPREADSHEETML_NAMESPACE}" count="2" uniqueCount="2"><si><t>ready</t></si><si><t>late"#,
            );
            let bytes = dependency_xlsx(
                &worksheet(r#"<c r="A1" t="s"><v>0</v></c>"#),
                Some(&late_shared),
                None,
            );
            let workbook = SourceBackedWorkbook::from_reader(Cursor::new(bytes)).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let cold = workbook.cache_diagnostics();
            assert!(sheet.cell("A1").is_err());
            assert!(sheet.data.cells.get().is_none());
            assert!(workbook.inner.shared_strings.get().is_none());
            assert_eq!(
                workbook.cache_diagnostics().successful_loads,
                cold.successful_loads
            );

            let late_styles = format!(
                r#"<styleSheet xmlns="{SPREADSHEETML_NAMESPACE}"><cellXfs count="1"><xf/></cellXfs><cellStyles>"#,
            );
            let bytes = dependency_xlsx(
                &worksheet(r#"<c r="A1" s="0"><v>7</v></c>"#),
                None,
                Some(&late_styles),
            );
            let workbook = SourceBackedWorkbook::from_reader(Cursor::new(bytes)).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let cold = workbook.cache_diagnostics();
            assert!(sheet.cell("A1").is_err());
            assert!(sheet.data.cells.get().is_none());
            assert!(workbook.inner.styles.get().is_none());
            assert_eq!(
                workbook.cache_diagnostics().successful_loads,
                cold.successful_loads
            );
        }

        #[test]
        fn dependency_crc_errors_0364_do_not_publish_semantics() {
            let mut shared = dependency_xlsx(
                &worksheet(r#"<c r="A1" t="s"><v>0</v></c>"#),
                Some(&plain_shared_strings(&["crc"])),
                None,
            );
            let shared_marker = b"<t>crc</t>";
            let shared_offset = shared
                .windows(shared_marker.len())
                .position(|window| window == shared_marker)
                .unwrap();
            shared[shared_offset + 3] = b"C"[0];
            let workbook = SourceBackedWorkbook::from_reader(Cursor::new(shared)).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let cold = workbook.cache_diagnostics();
            assert!(sheet.cell("A1").is_err());
            assert!(sheet.data.cells.get().is_none());
            assert!(workbook.inner.shared_strings.get().is_none());
            assert_eq!(
                workbook.cache_diagnostics().successful_loads,
                cold.successful_loads
            );

            let mut style = dependency_xlsx(
                &worksheet(r#"<c r="A1" s="0"><v>7</v></c>"#),
                None,
                Some(&styles(1).replace("<xf/>", r#"<xf numFmtId="0"/>"#)),
            );
            let style_marker = b"numFmtId=\"0\"";
            let style_offset = style
                .windows(style_marker.len())
                .position(|window| window == style_marker)
                .unwrap();
            style[style_offset + 10] = b"1"[0];
            let workbook = SourceBackedWorkbook::from_reader(Cursor::new(style)).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let cold = workbook.cache_diagnostics();
            assert!(sheet.cell("A1").is_err());
            assert!(sheet.data.cells.get().is_none());
            assert!(workbook.inner.styles.get().is_none());
            assert_eq!(
                workbook.cache_diagnostics().successful_loads,
                cold.successful_loads
            );
        }

        #[test]
        fn dependency_source_change_and_cancellation_0364_stay_primary() {
            let bytes = dependency_xlsx(
                &worksheet(r#"<c r="A1" t="s"><v>0</v></c>"#),
                Some(&plain_shared_strings(&["changed"])),
                Some(&styles(1)),
            );
            let source = Arc::new(CountingSource::new(bytes.clone()));
            let workbook = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let cold = workbook.cache_diagnostics();
            source.changed();
            assert!(matches!(
                sheet.cell("A1"),
                Err(Error::Package(OpcError::SourceChanged { .. }))
            ));
            assert!(sheet.data.cells.get().is_none());
            assert_eq!(
                workbook.cache_diagnostics().successful_loads,
                cold.successful_loads
            );

            let source = Arc::new(CountingSource::new(bytes));
            let (_budget, cancellation_source, context) = managed_context(u64::MAX);
            let workbook = SourceBackedWorkbook::from_read_at_with_execution_context(
                source,
                ReadLimits::default(),
                context,
            )
            .unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let cold = workbook.cache_diagnostics();
            cancellation_source.cancel();
            assert!(matches!(
                sheet.cell("A1"),
                Err(Error::Package(OpcError::Cancelled))
            ));
            assert!(sheet.data.cells.get().is_none());
            assert_eq!(
                workbook.cache_diagnostics().successful_loads,
                cold.successful_loads
            );
        }

        #[test]
        fn repeated_eligible_dependency_query_0364_rescans_source() {
            let bytes = dependency_xlsx(
                &worksheet(r#"<c r="A1" t="s"><v>0</v></c>"#),
                Some(&plain_shared_strings(&["repeat"])),
                Some(&styles(1)),
            );
            let source = Arc::new(CountingSource::new(bytes));
            let workbook = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let cold = workbook.cache_diagnostics();
            assert_text(sheet.cell("A1").unwrap(), "repeat");
            let first_reads = source.first_body_marker_reads.load(Ordering::SeqCst);
            assert!(first_reads > 0);
            assert_text(sheet.cell("A1").unwrap(), "repeat");
            assert!(source.first_body_marker_reads.load(Ordering::SeqCst) > first_reads);
            assert!(sheet.data.cells.get().is_none());
            assert!(workbook.inner.shared_strings.get().is_none());
            assert!(workbook.inner.styles.get().is_none());
            assert_eq!(
                workbook.cache_diagnostics().successful_loads,
                cold.successful_loads
            );
        }

        #[test]
        fn semantic_dependency_caches_0364_survive_payload_eviction() {
            let bytes = dependency_xlsx(
                &worksheet(r#"<c r="A1" t="s" s="0"><v>0</v></c>"#),
                Some(&rich_shared_strings()),
                Some(&styles(1)),
            );
            let source = Arc::new(CountingSource::new(bytes));
            let cache_limits = SourceCacheLimits::new(4096, 1).unwrap();
            let workbook =
                SourceBackedWorkbook::from_read_at_with_cache_limits(source, cache_limits).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            assert_text(sheet.cell("A1").unwrap(), "rich text");
            assert!(sheet.data.cells.get().is_some());
            assert!(workbook.inner.shared_strings.get().is_some());
            assert!(workbook.inner.styles.get().is_some());

            let shared_uri = workbook.inner.shared_strings_uri.as_ref().unwrap().clone();
            let styles_uri = workbook.inner.styles_uri.as_ref().unwrap().clone();
            {
                let _data = workbook
                    .inner
                    .package
                    .part(&shared_uri)
                    .unwrap()
                    .data()
                    .unwrap();
            }
            {
                let _data = workbook
                    .inner
                    .package
                    .part(&styles_uri)
                    .unwrap()
                    .data()
                    .unwrap();
            }
            let evicted = workbook.cache_diagnostics();
            assert!(evicted.evictions > 0);
            let cold_loads = evicted.cold_loads;
            assert_eq!(
                workbook.inner.shared_strings().unwrap().unwrap()[0].as_str(),
                "rich text"
            );
            assert_eq!(workbook.inner.style_count().unwrap(), 1);
            assert_text(sheet.cell("A1").unwrap(), "rich text");
            assert_eq!(workbook.cache_diagnostics().cold_loads, cold_loads);
        }

        #[test]
        fn final_fence_0364_outranks_parser_error() {
            let first = std::str::from_utf8(super::FIRST_MARKER).unwrap();
            let second = std::str::from_utf8(super::SECOND_MARKER).unwrap();
            let worksheet = format!(
                r#"<worksheet xmlns="{SPREADSHEETML_NAMESPACE}"><!--{first} {second}--><sheetData><row r="1"><c r="A1"><v>7</v></c>"#,
            );
            let bytes = dependency_xlsx(&worksheet, None, None);
            let baseline_source = Arc::new(VersionFlipSource::new(bytes.clone(), None));
            let baseline = SourceBackedWorkbook::from_read_at(baseline_source.clone()).unwrap();
            let baseline_sheet = baseline.sheet("Sheet1").unwrap().unwrap();
            assert!(baseline_sheet.cell("A1").is_err());
            let final_check_call = baseline_source.version_calls.load(Ordering::SeqCst);
            drop(baseline_sheet);
            drop(baseline);

            let source = Arc::new(VersionFlipSource::new(bytes, Some(final_check_call)));
            let workbook = SourceBackedWorkbook::from_read_at(source).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            assert!(matches!(
                sheet.cell("A1"),
                Err(Error::Package(OpcError::SourceChanged { .. }))
            ));
            assert!(sheet.data.cells.get().is_none());
        }
    }

    mod streaming_0365_tests {
        use std::sync::Arc;
        use std::sync::atomic::Ordering;

        use litchi_opc::{OpcError, SourceBackedPackage, SourceCacheLimits};

        use super::super::SourceCell;
        use super::streaming_0364_tests::{dependency_xlsx, plain_shared_strings, styles};
        use super::{
            CountingSource, ReadLimits, SourceBackedWorkbook, SourceCellView, managed_context,
        };
        use crate::{Cell, Error, Value};

        const SPREADSHEETML_NAMESPACE: &str =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const CALLBACK_ERROR: &str = "streaming_0365 callback: stop";

        fn worksheet_xml(sheet_data: &str) -> String {
            worksheet_xml_after(sheet_data, "")
        }

        fn worksheet_xml_after(sheet_data: &str, after_root: &str) -> String {
            let first = std::str::from_utf8(super::FIRST_MARKER).unwrap();
            let second = std::str::from_utf8(super::SECOND_MARKER).unwrap();
            format!(
                r#"<worksheet xmlns="{SPREADSHEETML_NAMESPACE}"><!--{first} {second}-->{sheet_data}</worksheet>{after_root}"#,
            )
        }

        fn worksheet_xml_open(sheet_data: &str) -> String {
            let first = std::str::from_utf8(super::FIRST_MARKER).unwrap();
            let second = std::str::from_utf8(super::SECOND_MARKER).unwrap();
            format!(
                r#"<worksheet xmlns="{SPREADSHEETML_NAMESPACE}"><!--{first} {second}-->{sheet_data}"#,
            )
        }

        fn source_workbook(
            worksheet: &str,
            shared_strings: Option<&str>,
            styles: Option<&str>,
        ) -> (Arc<CountingSource>, SourceBackedWorkbook) {
            let source = Arc::new(CountingSource::new(dependency_xlsx(
                worksheet,
                shared_strings,
                styles,
            )));
            let workbook = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
            (source, workbook)
        }

        fn stored_fixture() -> String {
            worksheet_xml(
                r#"<sheetData>
                    <row r="1"><c r="A1"><v>7</v></c><c r="C1"/></row>
                    <row r="2"><c r="B2"><v>8</v></c></row>
                    <row r="4"><c r="A4"><v>9</v></c></row>
                </sheetData>"#,
            )
        }

        fn addresses(cells: &[SourceCell]) -> Vec<String> {
            cells.iter().map(|cell| cell.address.a1()).collect()
        }

        fn assert_zero_callbacks(bytes: Vec<u8>) {
            let source = Arc::new(CountingSource::new(bytes));
            let workbook = SourceBackedWorkbook::from_read_at(source).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let before = workbook.cache_diagnostics();
            let mut callbacks = Vec::new();

            let result = sheet.visit_cells("A1:D4", |address, _cell| {
                callbacks.push(address.a1());
                Ok(())
            });

            assert!(result.is_err());
            assert!(callbacks.is_empty());
            assert!(sheet.data.cells.get().is_none());
            let after = workbook.cache_diagnostics();
            assert_eq!(after.cold_loads, before.cold_loads);
            assert_eq!(after.successful_loads, before.successful_loads);
            assert_eq!(after.retained_bytes, before.retained_bytes);
        }

        #[test]
        fn cold_cells_0365_are_sparse_row_major_and_area_bounded() {
            let (_source, workbook) = source_workbook(&stored_fixture(), None, None);
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();

            let all = sheet.cells("A1:D4").unwrap();
            assert_eq!(
                addresses(&all),
                vec![
                    "A1".to_owned(),
                    "C1".to_owned(),
                    "B2".to_owned(),
                    "A4".to_owned()
                ]
            );
            assert!(matches!(&all[1].cell, Cell::Empty));
            assert!(matches!(
                &all[0].cell,
                Cell::Value(Value::Number(value)) if value.as_str() == "7"
            ));
            assert!(sheet.data.cells.get().is_none());

            let subset = sheet.cells("B1:C2").unwrap();
            assert_eq!(addresses(&subset), vec!["C1".to_owned(), "B2".to_owned()]);
            assert!(matches!(&subset[0].cell, Cell::Empty));
            assert!(matches!(
                &subset[1].cell,
                Cell::Value(Value::Number(value)) if value.as_str() == "8"
            ));

            assert!(sheet.cells("Z10:AA11").unwrap().is_empty());
            assert!(sheet.data.cells.get().is_none());
        }

        #[test]
        fn cold_visit_cells_0365_matches_sparse_sequence_and_count() {
            let (_source, workbook) = source_workbook(&stored_fixture(), None, None);
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let expected = sheet.cells("A1:D4").unwrap();
            let mut visited = Vec::new();

            let count = sheet
                .visit_cells("A1:D4", |address, cell| {
                    visited.push((address, cell.clone()));
                    Ok(())
                })
                .unwrap();

            let expected = expected
                .into_iter()
                .map(|cell| (cell.address, cell.cell))
                .collect::<Vec<_>>();
            assert_eq!(count, expected.len());
            assert_eq!(visited, expected);
            assert!(sheet.data.cells.get().is_none());
        }

        #[test]
        fn multiple_dependencies_0365_stay_cold_without_store_or_part_cache() {
            let worksheet = worksheet_xml(
                r#"<sheetData>
                    <row r="1">
                        <c r="A1" t="s" s="0"><v>2</v></c>
                        <c r="C1" t="s" s="1"><v>0</v></c>
                        <c r="E1" s="1"><v>4</v></c>
                    </row>
                    <row r="2"><c r="B2" t="s" s="0"><v>1</v></c></row>
                </sheetData>"#,
            );
            let (_source, workbook) = source_workbook(
                &worksheet,
                Some(&plain_shared_strings(&["zero", "one", "two"])),
                Some(&styles(2)),
            );
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let before = workbook.cache_diagnostics();

            let cells = sheet.cells("A1:E2").unwrap();
            assert_eq!(
                addresses(&cells),
                vec![
                    "A1".to_owned(),
                    "C1".to_owned(),
                    "E1".to_owned(),
                    "B2".to_owned(),
                ]
            );
            assert!(matches!(
                &cells[0].cell,
                Cell::Value(Value::Text(value)) if value.as_str() == "two"
            ));
            assert!(matches!(
                &cells[1].cell,
                Cell::Value(Value::Text(value)) if value.as_str() == "zero"
            ));
            assert!(matches!(
                &cells[2].cell,
                Cell::Value(Value::Number(value)) if value.as_str() == "4"
            ));
            assert!(matches!(
                &cells[3].cell,
                Cell::Value(Value::Text(value)) if value.as_str() == "one"
            ));
            assert!(sheet.data.cells.get().is_none());
            assert!(workbook.inner.shared_strings.get().is_none());
            assert!(workbook.inner.styles.get().is_none());
            let after = workbook.cache_diagnostics();
            assert_eq!(after.cold_loads, before.cold_loads);
            assert_eq!(after.successful_loads, before.successful_loads);
            assert_eq!(after.retained_bytes, before.retained_bytes);
        }

        #[test]
        fn repeated_cells_0365_rescan_without_cache() {
            let (source, workbook) = source_workbook(&stored_fixture(), None, None);
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let before = workbook.cache_diagnostics();

            let first = sheet.cells("A1:D4").unwrap();
            let first_reads = source.first_body_marker_reads.load(Ordering::SeqCst);
            let second = sheet.cells("A1:D4").unwrap();
            let second_reads = source.first_body_marker_reads.load(Ordering::SeqCst);

            assert_eq!(first, second);
            assert!(first_reads > 0);
            assert!(second_reads > first_reads);
            assert!(sheet.data.cells.get().is_none());
            assert_eq!(
                workbook.cache_diagnostics().successful_loads,
                before.successful_loads
            );
        }

        #[test]
        fn warm_store_cells_0365_match_cold_selection() {
            let (_source, workbook) = source_workbook(&stored_fixture(), None, None);
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();

            let cold = sheet.cells("A1:D4").unwrap();
            assert!(sheet.data.cells.get().is_none());
            assert!(sheet.stored_extent().unwrap().is_some());
            assert!(sheet.data.cells.get().is_some());
            let warm = sheet.cells("A1:D4").unwrap();

            assert_eq!(warm, cold);
        }

        #[test]
        fn merge_and_shared_formula_0365_fall_back_to_store() {
            let merge = worksheet_xml(
                r#"<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>"#,
            );
            let (_source, workbook) = source_workbook(&merge, None, None);
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let cells = sheet.cells("A1:B1").unwrap();
            assert_eq!(addresses(&cells), vec!["A1".to_owned()]);
            assert!(sheet.data.cells.get().is_some());
            assert!(matches!(
                sheet.cell("B1").unwrap(),
                SourceCellView::Covered(range) if range.a1() == "A1:B1"
            ));

            let shared = worksheet_xml(
                r#"<sheetData>
                    <row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="7">B1+$C$1</f><v>1</v></c></row>
                    <row r="2"><c r="A2"><f t="shared" si="7"/><v>2</v></c></row>
                </sheetData>"#,
            );
            let (_source, workbook) = source_workbook(&shared, None, None);
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let cells = sheet.cells("A1:A2").unwrap();
            assert_eq!(addresses(&cells), vec!["A1".to_owned(), "A2".to_owned()]);
            assert!(matches!(&cells[0].cell, Cell::Formula(_)));
            assert!(matches!(&cells[1].cell, Cell::Formula(_)));
            assert!(sheet.data.cells.get().is_some());
        }

        #[test]
        fn malformed_tail_and_worksheet_0365_have_zero_callbacks() {
            let malformed_tail = dependency_xlsx(
                &worksheet_xml_after(
                    r#"<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>"#,
                    "<tail>",
                ),
                None,
                None,
            );
            assert_zero_callbacks(malformed_tail);

            let malformed_worksheet = dependency_xlsx(
                &worksheet_xml_open(r#"<sheetData><row r="1"><c r="A1"><v>7</v></c>"#),
                None,
                None,
            );
            assert_zero_callbacks(malformed_worksheet);
        }

        #[test]
        fn dependency_crc_0365_has_zero_callbacks_without_publication() {
            let mut bytes = dependency_xlsx(
                &worksheet_xml(
                    r#"<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>"#,
                ),
                Some(&plain_shared_strings(&["crc"])),
                None,
            );
            let marker = b"<t>crc</t>";
            let offset = bytes
                .windows(marker.len())
                .position(|window| window == marker)
                .unwrap();
            bytes[offset + 3] = b'C';

            let source = Arc::new(CountingSource::new(bytes));
            let workbook = SourceBackedWorkbook::from_read_at(source).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let before = workbook.cache_diagnostics();
            let mut callbacks = 0usize;
            let result = sheet.visit_cells("A1:B1", |_address, _cell| {
                callbacks += 1;
                Ok(())
            });

            assert!(result.is_err());
            assert_eq!(callbacks, 0);
            assert!(sheet.data.cells.get().is_none());
            assert!(workbook.inner.shared_strings.get().is_none());
            assert_eq!(
                workbook.cache_diagnostics().successful_loads,
                before.successful_loads
            );
        }

        #[test]
        fn source_change_and_cancellation_0365_have_zero_callbacks() {
            let (source, workbook) = source_workbook(&stored_fixture(), None, None);
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let before = workbook.cache_diagnostics();
            source.changed();
            let mut callbacks = 0usize;
            let error = sheet
                .visit_cells("A1:D4", |_address, _cell| {
                    callbacks += 1;
                    Ok(())
                })
                .unwrap_err();

            assert!(matches!(
                error,
                Error::Package(OpcError::SourceChanged { .. })
            ));
            assert_eq!(callbacks, 0);
            assert!(sheet.data.cells.get().is_none());
            assert_eq!(
                workbook.cache_diagnostics().successful_loads,
                before.successful_loads
            );

            let source = Arc::new(CountingSource::new(dependency_xlsx(
                &stored_fixture(),
                None,
                None,
            )));
            let (_budget, cancellation_source, context) = managed_context(u64::MAX);
            let workbook = SourceBackedWorkbook::from_read_at_with_execution_context(
                source,
                ReadLimits::default(),
                context,
            )
            .unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let before = workbook.cache_diagnostics();
            cancellation_source.cancel();
            let mut callbacks = 0usize;
            let error = sheet
                .visit_cells("A1:D4", |_address, _cell| {
                    callbacks += 1;
                    Ok(())
                })
                .unwrap_err();

            assert!(matches!(error, Error::Package(OpcError::Cancelled)));
            assert_eq!(callbacks, 0);
            assert!(sheet.data.cells.get().is_none());
            assert_eq!(
                workbook.cache_diagnostics().successful_loads,
                before.successful_loads
            );
        }

        #[test]
        fn callback_error_0365_preserves_exact_prefix() {
            let (_source, workbook) = source_workbook(&stored_fixture(), None, None);
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let mut callbacks = 0usize;
            let error = sheet
                .visit_cells("A1:D4", |_address, _cell| {
                    callbacks += 1;
                    Err(Error::Invalid(CALLBACK_ERROR.into()))
                })
                .unwrap_err();

            match error {
                Error::Invalid(message) => assert_eq!(message, CALLBACK_ERROR),
                other => panic!("callback error changed type: {other:?}"),
            }
            assert_eq!(callbacks, 1);
            assert!(sheet.data.cells.get().is_none());
        }

        #[test]
        fn callback_cancellation_0365_uses_package_owned_context() {
            let source = Arc::new(CountingSource::new(dependency_xlsx(
                &stored_fixture(),
                None,
                None,
            )));
            let (_budget, cancellation_source, context) = managed_context(u64::MAX);
            let package =
                SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
                    source,
                    ReadLimits::default(),
                    SourceCacheLimits::new(256 * 1024, 4).unwrap(),
                    context,
                )
                .unwrap();
            let workbook = SourceBackedWorkbook::from_source_backed_package(package).unwrap();
            assert!(workbook.inner.package.execution_context().is_some());
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let mut callbacks = Vec::new();

            let error = sheet
                .visit_cells("A1:D4", |address, _cell| {
                    callbacks.push(address.a1());
                    cancellation_source.cancel();
                    Ok(())
                })
                .unwrap_err();

            assert!(matches!(error, Error::Package(OpcError::Cancelled)));
            assert_eq!(callbacks, vec!["A1".to_owned()]);
            assert!(sheet.data.cells.get().is_none());
        }

        #[test]
        fn final_source_change_0365_outranks_callback_error() {
            let source = Arc::new(CountingSource::new(dependency_xlsx(
                &stored_fixture(),
                None,
                None,
            )));
            let (_budget, _cancellation_source, context) = managed_context(u64::MAX);
            let package = SourceBackedPackage::from_read_at_with_execution_context(
                source.clone(),
                ReadLimits::default(),
                context,
            )
            .unwrap();
            let workbook = SourceBackedWorkbook::from_source_backed_package(package).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let mut callbacks = 0usize;

            let error = sheet
                .visit_cells("A1:D4", |_address, _cell| {
                    callbacks += 1;
                    source.changed();
                    Err(Error::Invalid(CALLBACK_ERROR.into()))
                })
                .unwrap_err();

            assert!(matches!(
                error,
                Error::Package(OpcError::SourceChanged { .. })
            ));
            assert_eq!(callbacks, 1);
            assert!(sheet.data.cells.get().is_none());
        }
    }
}
