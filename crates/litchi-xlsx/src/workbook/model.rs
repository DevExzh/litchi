//! Semantic workbook snapshot and worksheet models.

use std::collections::VecDeque;
use std::collections::{HashMap, HashSet};
use std::convert::Infallible;
use std::mem::size_of;
use std::sync::{Arc, Mutex, OnceLock};

use super::worksheet;
use super::{codec, package};
use litchi_core::Selector as CoreSelector;
use litchi_opc::{OpcPackage, PackURI, Relationships};
use litchi_sheet::{Area, At, Cell as Address, ColumnAt, Rect, RowAt};
use once_cell::sync::OnceCell;

use crate::cell::{Extents, SharedStringLineage, Store, Text};
use crate::error::{Error, Result, invalid};
use crate::leniency::{Leniency, ToleranceReport};
use crate::raw;
use crate::style::StyleLineage;
use crate::style::cell_format::{Format, FormatCache, FormatKey, FormatTable};
use crate::{Cells, Column, Columns, LocalStyle, Row, Rows, Style, Styles};

#[cfg(feature = "encryption")]
use litchi_ooxml_common::package_encryption::PackageEncryption;

/// Semantic selector accepted by [`Workbook::sheet`].
///
/// Names and checked zero-based positions are the ordinary entry points. The
/// uninhabited identity variant reserves room for a future lineage-checked
/// durable selector without exposing native `SpreadsheetML` IDs.
pub type Selector<'a> = litchi_sheet::SheetSelector<'a, Infallible>;

/// Runtime workbook flavor derived from the main-part content type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Flavor {
    Workbook,
    Template,
    MacroWorkbook,
    MacroTemplate,
}

impl Flavor {
    /// Whether this flavor can contain a VBA project without promotion.
    #[must_use]
    pub const fn allows_macros(self) -> bool {
        matches!(self, Self::MacroWorkbook | Self::MacroTemplate)
    }

    /// Whether opening the file is intended to create a new workbook.
    #[must_use]
    pub const fn is_template(self) -> bool {
        matches!(self, Self::Template | Self::MacroTemplate)
    }
}

/// Workbook date serial system.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DateSystem {
    Excel1900,
    Excel1904,
}

/// Semantic sheet kind resolved from the workbook relationship.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum WorksheetKind {
    Worksheet,
    Chart,
    Dialog,
    Macro,
    Unknown,
}

/// Worksheet visibility retained without approximating producer extensions.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Visibility {
    Visible,
    Hidden,
    VeryHidden,
    Unknown(Box<str>),
}

impl Visibility {
    /// Whether Excel displays this sheet tab.
    #[must_use]
    pub const fn is_visible(&self) -> bool {
        matches!(self, Self::Visible)
    }

    /// Whether this tab is hidden by either recognized mechanism.
    #[must_use]
    pub const fn is_hidden(&self) -> bool {
        matches!(self, Self::Hidden | Self::VeryHidden)
    }

    /// Whether Excel omits this tab from its ordinary Unhide dialog.
    #[must_use]
    pub const fn is_very_hidden(&self) -> bool {
        matches!(self, Self::VeryHidden)
    }
}

#[derive(Debug)]
pub(super) struct SheetData {
    pub(super) position: usize,
    pub(super) name: String,
    pub(super) name_key: Box<str>,
    pub(super) kind: WorksheetKind,
    pub(super) visibility: Visibility,
    pub(super) part_uri: PackURI,
    pub(super) cells: OnceLock<Store>,
    pub(super) web_bindings: OnceCell<crate::web::Bindings>,
    /// Exact-source page-break projection identity and bounded cache handle.
    pub(super) page_breaks: Arc<PageBreakProjection>,
    #[allow(
        dead_code,
        reason = "the cached column index supports internal workbook edit planning"
    )]
    pub(super) native_id: u32,
    pub(super) relationship_id: String,
}

const PAGE_BREAK_CACHE_MAX_ENTRIES: usize = 64;
const PAGE_BREAK_CACHE_MAX_WEIGHT: usize = 2 * 1024 * 1024;

#[derive(Debug)]
struct PageBreakCacheEntry {
    source: Arc<Vec<u8>>,
    value: Arc<crate::page_breaks::PageBreaks>,
    weight: usize,
}

#[derive(Debug, Default)]
struct PageBreakCacheState {
    entries: VecDeque<PageBreakCacheEntry>,
    weight: usize,
    hits: usize,
    misses: usize,
}

/// Finite weighted cache for successful worksheet page-break projections.
///
/// The cache is shared by immutable descendants of one workbook lineage, but
/// never grows with the number of snapshots. Exact worksheet blob allocation
/// identity is the invalidation key: publication of changed bytes cannot hit
/// an earlier projection even when the part URI is unchanged.
#[derive(Debug, Default)]
pub(super) struct PageBreakProjectionCache {
    state: Mutex<PageBreakCacheState>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(
    dead_code,
    reason = "cache statistics are consumed by focused projection tests"
)]
pub(super) struct PageBreakCacheStats {
    pub(super) entries: usize,
    pub(super) weight: usize,
    pub(super) hits: usize,
    pub(super) misses: usize,
}

impl PageBreakProjectionCache {
    fn lock(&self) -> std::sync::MutexGuard<'_, PageBreakCacheState> {
        match self.state.lock() {
            Ok(state) => state,
            Err(poisoned) => poisoned.into_inner(),
        }
    }

    fn cached(&self, source: &Arc<Vec<u8>>) -> Option<Arc<crate::page_breaks::PageBreaks>> {
        let mut state = self.lock();
        let index = state
            .entries
            .iter()
            .position(|entry| Arc::ptr_eq(&entry.source, source))?;
        let entry = state.entries.remove(index)?;
        let value = Arc::clone(&entry.value);
        state.entries.push_back(entry);
        Some(value)
    }

    pub(super) fn get_or_try_init(
        &self,
        source: Arc<Vec<u8>>,
        parse: impl FnOnce(&[u8]) -> Result<crate::page_breaks::PageBreaks>,
    ) -> Result<Arc<crate::page_breaks::PageBreaks>> {
        if let Some(value) = self.cached(&source) {
            let mut state = self.lock();
            state.hits = state.hits.saturating_add(1);
            return Ok(value);
        }

        let value = Arc::new(parse(source.as_slice())?);
        let weight = page_break_projection_weight(source.capacity(), &value);
        {
            let mut state = self.lock();
            state.misses = state.misses.saturating_add(1);
            // Another reader may have parsed and published the same source
            // while this reader was outside the mutex. Prefer that canonical
            // value so concurrent misses cannot accumulate duplicate entries
            // for one exact worksheet blob.
            if let Some(index) = state
                .entries
                .iter()
                .position(|entry| Arc::ptr_eq(&entry.source, &source))
            {
                if let Some(entry) = state.entries.remove(index) {
                    let existing = Arc::clone(&entry.value);
                    state.entries.push_back(entry);
                    state.hits = state.hits.saturating_add(1);
                    return Ok(existing);
                }
            }
            if weight <= PAGE_BREAK_CACHE_MAX_WEIGHT {
                while state.entries.len() >= PAGE_BREAK_CACHE_MAX_ENTRIES
                    || state.weight.saturating_add(weight) > PAGE_BREAK_CACHE_MAX_WEIGHT
                {
                    let Some(evicted) = state.entries.pop_front() else {
                        break;
                    };
                    state.weight = state.weight.saturating_sub(evicted.weight);
                }
                state.weight = state.weight.saturating_add(weight);
                state.entries.push_back(PageBreakCacheEntry {
                    source,
                    value: Arc::clone(&value),
                    weight,
                });
            }
        }
        Ok(value)
    }

    #[allow(
        dead_code,
        reason = "cache statistics are consumed by focused projection tests"
    )]
    pub(super) fn stats(&self) -> PageBreakCacheStats {
        let state = self.lock();
        PageBreakCacheStats {
            entries: state.entries.len(),
            weight: state.weight,
            hits: state.hits,
            misses: state.misses,
        }
    }
}

fn page_break_projection_weight(
    source_capacity: usize,
    value: &crate::page_breaks::PageBreaks,
) -> usize {
    let breaks = value
        .horizontal()
        .into_iter()
        .chain(value.vertical())
        .map(crate::page_breaks::Collection::len)
        .sum::<usize>();
    let collections = value.horizontal().into_iter().count() + value.vertical().into_iter().count();
    source_capacity
        .saturating_add(size_of::<Vec<u8>>())
        .saturating_add(size_of::<crate::page_breaks::PageBreaks>())
        .saturating_add(collections.saturating_mul(size_of::<crate::page_breaks::Collection>()))
        .saturating_add(breaks.saturating_mul(size_of::<crate::page_breaks::Break>()))
        .saturating_add(size_of::<PageBreakCacheEntry>())
}

/// Source identity plus the bounded projection cache used by one worksheet
/// handle. This small handle is copied across unchanged snapshots; the
/// parsed value itself lives only in the finite weighted cache above.
#[derive(Debug)]
pub(super) struct PageBreakProjection {
    cache: Arc<PageBreakProjectionCache>,
    #[allow(
        dead_code,
        reason = "source identity is inspected by focused cache invalidation tests"
    )]
    source: Arc<Vec<u8>>,
}

impl PageBreakProjection {
    fn new(cache: Arc<PageBreakProjectionCache>, source: Arc<Vec<u8>>) -> Self {
        Self { cache, source }
    }

    #[allow(
        dead_code,
        reason = "cached values are inspected by focused projection tests"
    )]
    pub(super) fn get(&self) -> Option<Arc<crate::page_breaks::PageBreaks>> {
        let mut state = self.cache.lock();
        let index = state
            .entries
            .iter()
            .position(|entry| Arc::ptr_eq(&entry.source, &self.source))?;
        let entry = state.entries.remove(index)?;
        let value = Arc::clone(&entry.value);
        state.entries.push_back(entry);
        Some(value)
    }

    pub(super) fn get_or_try_init(
        &self,
        source: Arc<Vec<u8>>,
        parse: impl FnOnce(&[u8]) -> Result<crate::page_breaks::PageBreaks>,
    ) -> Result<Arc<crate::page_breaks::PageBreaks>> {
        self.cache.get_or_try_init(source, parse)
    }
}

/// A commit-validated worksheet store tied to the exact published part bytes.
pub(super) struct ValidatedWorksheetStore {
    pub(super) uri: PackURI,
    pub(super) content: Arc<Vec<u8>>,
    pub(super) store: Store,
}

struct SourcePageBreakProjection {
    blob: Arc<Vec<u8>>,
    projection: Arc<PageBreakProjection>,
}

/// Build the source cache index once per descendant snapshot. Keeping the
/// lookup keyed by part URI avoids scanning every source sheet for each
/// target sheet during publication.
fn page_break_caches(source: Option<&Workbook>) -> HashMap<PackURI, SourcePageBreakProjection> {
    let Some(source) = source else {
        return HashMap::new();
    };
    let mut caches = HashMap::with_capacity(source.inner.sheets.len());
    for sheet in source.inner.sheets.iter() {
        let Ok(part) = source.inner.package.get_part(&sheet.part_uri) else {
            continue;
        };
        caches.insert(
            sheet.part_uri.clone(),
            SourcePageBreakProjection {
                blob: part.blob_arc(),
                projection: Arc::clone(&sheet.page_breaks),
            },
        );
    }
    caches
}

/// Reuse a successful worksheet projection only when the target snapshot
/// retains the exact source payload allocation. A publication that replaces
/// the worksheet bytes receives a fresh cache, so no edit can observe a
/// projection from an earlier generation.
fn page_break_cache(
    source_caches: &HashMap<PackURI, SourcePageBreakProjection>,
    package: &OpcPackage,
    part_uri: &PackURI,
    cache: &Arc<PageBreakProjectionCache>,
) -> Result<Arc<PageBreakProjection>> {
    let target_part = package.get_part(part_uri)?;
    let target_blob = target_part.blob_arc();
    let Some(source) = source_caches.get(part_uri) else {
        return Ok(Arc::new(PageBreakProjection::new(
            Arc::clone(cache),
            target_blob,
        )));
    };
    if Arc::ptr_eq(&source.blob, &target_blob) {
        Ok(Arc::clone(&source.projection))
    } else {
        Ok(Arc::new(PageBreakProjection::new(
            Arc::clone(&source.projection.cache),
            target_blob,
        )))
    }
}

/// Compare the public semantic fields of one OPC relationship collection.
///
/// The collection's source-preservation handle is intentionally ignored: it
/// describes how the value was admitted, not the value itself.
fn same_relationships(left: &Relationships, right: &Relationships) -> bool {
    left.base_uri() == right.base_uri()
        && left.source_uri() == right.source_uri()
        && left.len() == right.len()
        && left.iter().all(|relationship| {
            right.get(relationship.r_id()).is_some_and(|candidate| {
                candidate.reltype() == relationship.reltype()
                    && candidate.target_ref() == relationship.target_ref()
                    && candidate.target_mode() == relationship.target_mode()
                    && candidate.base_uri() == relationship.base_uri()
                    && candidate.source_uri() == relationship.source_uri()
            })
        })
}

#[derive(Debug)]
pub(crate) struct Inner {
    pub(super) package: OpcPackage,
    #[allow(
        dead_code,
        reason = "the cached row index supports internal workbook edit planning"
    )]
    pub(super) workbook_uri: PackURI,
    pub(super) shared_strings_uri: Option<PackURI>,
    pub(super) shared_strings: OnceLock<Box<[Text]>>,
    /// Repairs made while the shared-string table was parsed; published
    /// before the table itself (see `crate::leniency`).
    pub(super) shared_strings_tolerance: OnceLock<ToleranceReport>,
    /// Ascending indexes of the shared-string items that carry rich-text
    /// runs, scanned on first use (record 0273).
    pub(super) rich_shared_strings: OnceLock<Box<[usize]>>,
    pub(crate) shared_string_lineage: Arc<SharedStringLineage>,
    pub(super) styles_uri: Option<PackURI>,
    pub(super) styles: OnceLock<raw::styles::Catalog>,
    pub(super) task_panes: OnceCell<Option<litchi_ooxml_common::web::Panes>>,
    pub(super) surveys: OnceCell<Box<[crate::survey::Part]>>,
    pub(crate) style_lineage: Arc<StyleLineage>,
    pub(super) flavor: Flavor,
    pub(super) date_system: DateSystem,
    pub(super) active_sheet: Option<usize>,
    pub(super) sheets: Box<[Arc<SheetData>]>,
    /// Sheet positions sorted by their canonical, case-insensitive names.
    ///
    /// `raw::parse_catalog` rejects duplicate canonical names before this
    /// snapshot is built, so the order is a total lookup index rather than a
    /// lossy collision map. Keeping positions (instead of cloning keys into
    /// another map) preserves the existing name-key allocations.
    pub(super) sheet_name_order: Box<[usize]>,
    pub(super) defined_names: Box<[raw::DefinedName]>,
    pub(super) pivot_caches: Box<[raw::PivotCache]>,
    pub(super) external_reference_ids: Box<[String]>,
    pub(super) page_break_cache: Arc<PageBreakProjectionCache>,
    /// Weighted clean-value cache of the parsed format table, shared by the
    /// lineage's descendant snapshots and keyed by the styles payload
    /// allocation.
    pub(super) format_cache: Arc<FormatCache>,
    /// Memo of the chart inventory of the worksheet drawing read last,
    /// shared by the lineage like `format_cache` and keyed by the drawing
    /// payload allocation (record 0286).
    pub(super) chart_inventories: Arc<crate::chart::edit::InventoryCache>,
    /// Repairs made while the workbook catalog was parsed at open time.
    pub(super) tolerance: ToleranceReport,
    /// The ingress policies this workbook retains after open (its leniency,
    /// semantic limits and encryption limits); snapshots derived from it by
    /// edits re-open under the same policies.
    pub(super) options: WorkbookOptions,
}

/// The policies an owned workbook snapshot actually needs after ingress.
///
/// Physical source limits live on the validated OPC package, while cache
/// limits and an execution context belong only to a source-backed owner.
/// Keeping those source-owner fields out of every eager snapshot preserves
/// the compact immutable workbook layout without dropping any policy used by
/// owned edits or encrypted output.
#[derive(Debug, Clone, Copy, Default)]
pub(crate) struct WorkbookOptions {
    pub(super) leniency: Leniency,
    pub(super) semantic_limits: crate::SemanticLimits,
    #[cfg(feature = "encryption")]
    pub(super) encryption_limits: crate::encryption::Limits,
}

impl WorkbookOptions {
    fn from_open_options(options: &crate::OpenOptions) -> Self {
        Self {
            leniency: options.leniency,
            semantic_limits: options.semantic_limits,
            #[cfg(feature = "encryption")]
            encryption_limits: options.encryption_limits,
        }
    }
}

/// The owned, read-only semantic projection of an XLSX [`crate::Package`].
///
/// A projection is cheap to clone and shares the package's immutable state;
/// [`crate::Package::workbook`] produces one. It exposes the workbook catalog,
/// sheet handles, cells, styles and the other semantic reads. Every state
/// verb (open, edit, apply, save) lives on the package.
#[derive(Debug, Clone)]
pub struct Workbook {
    pub(super) inner: Arc<Inner>,
    /// The policies of this snapshot that do not belong to its state;
    /// snapshots derived from it by edits share them.
    ///
    /// The ordinary strict package path keeps this absent and reads its
    /// semantic defaults from `Inner::options`.  An override (or encrypted
    /// provenance) installs the shared policy lazily; this keeps the common
    /// eager open path allocation-free for policy state.
    policy: Option<Arc<Policy>>,
}

/// The semantic budgets and encryption provenance of a workbook snapshot.
#[derive(Debug, Clone, Copy)]
pub(super) struct Policy {
    semantic_limits: crate::SemanticLimits,
    #[cfg(feature = "encryption")]
    encryption: PackageEncryption,
}

impl Workbook {
    /// Validate a clear OPC graph and build its strict projection. The
    /// public ingress is [`crate::Package`]; this raw boundary is internal.
    pub(crate) fn from_package(package: OpcPackage) -> Result<Self> {
        Self::from_package_with_styles(package, None)
    }

    /// [`Self::from_package`] under the ingress policies `options`, which
    /// the projection keeps.
    pub(crate) fn from_package_with_options(
        package: OpcPackage,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        Self::open_package(package, None, WorkbookOptions::from_open_options(&options))
    }

    /// The semantic budgets of this snapshot.
    pub(crate) fn semantic_limits(&self) -> &crate::SemanticLimits {
        self.policy
            .as_deref()
            .map_or(&self.inner.options.semantic_limits, |policy| {
                &policy.semantic_limits
            })
    }

    /// This snapshot under the semantic budgets `limits`; the state is
    /// shared.
    pub(crate) fn with_semantic_limits(&self, limits: crate::SemanticLimits) -> Self {
        let mut policy = self.policy.as_deref().copied().unwrap_or_else(|| Policy {
            semantic_limits: self.inner.options.semantic_limits,
            #[cfg(feature = "encryption")]
            encryption: PackageEncryption::plain(),
        });
        policy.semantic_limits = limits;
        Self {
            inner: Arc::clone(&self.inner),
            policy: Some(Arc::new(policy)),
        }
    }

    #[cfg(feature = "encryption")]
    pub(crate) fn encryption_limits(&self) -> crate::encryption::Limits {
        self.inner.options.encryption_limits
    }

    /// The lineage's memo of worksheet drawing chart inventories.
    pub(crate) fn chart_inventories(&self) -> &crate::chart::edit::InventoryCache {
        &self.inner.chart_inventories
    }

    /// The opt-in repair policy the owning package was opened with.
    #[must_use]
    pub fn leniency(&self) -> Leniency {
        self.inner.options.leniency
    }

    #[cfg(feature = "encryption")]
    pub(crate) fn from_package_with_encryption(
        package: OpcPackage,
        encryption: PackageEncryption,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        let mut workbook =
            Self::open_package(package, None, WorkbookOptions::from_open_options(&options))?;
        workbook.policy = Some(Arc::new(Policy {
            encryption,
            semantic_limits: workbook.semantic_limits().to_owned(),
        }));
        Ok(workbook)
    }

    /// Encryption provenance retained from the owning package.
    #[cfg(feature = "encryption")]
    pub(crate) fn encryption_state(&self) -> PackageEncryption {
        self.policy
            .as_deref()
            .map_or_else(PackageEncryption::plain, |policy| policy.encryption)
    }

    pub(crate) fn from_package_with_styles(
        package: OpcPackage,
        source: Option<&Workbook>,
    ) -> Result<Self> {
        let options = source.map_or_else(WorkbookOptions::default, |source| source.inner.options);
        Self::open_package(package, source, options)
    }

    fn open_package(
        package: OpcPackage,
        source: Option<&Workbook>,
        options: WorkbookOptions,
    ) -> Result<Self> {
        let leniency = options.leniency;
        // The packaging repairs come first (they were made before the
        // workbook part could be read), then the open-time catalog sites
        // (findings 4.7/4.8); see `crate::leniency`.
        let mut tolerance = ToleranceReport::default();
        tolerance.append_package(package.tolerance_report());
        let (workbook_uri, flavor, catalog, sheet_parts, shared_strings_uri, styles_uri) = {
            let workbook = package.main_document_part()?;
            let flavor = codec::flavor(workbook.content_type()).ok_or_else(|| {
                invalid(format!(
                    "main part '{}' has non-XLSX content type '{}'",
                    workbook.partname(),
                    workbook.content_type()
                ))
            })?;
            let catalog = raw::parse_catalog_repairing(
                workbook.blob(),
                &package.read_limits(),
                leniency,
                &mut tolerance,
            )?;
            let sheet_parts = package::validate_sheet_graph(&package, workbook, &catalog.sheets)?;
            let shared_strings_uri = package::validate_shared_strings(&package, workbook)?;
            let styles_uri = package::validate_styles(&package, workbook)?;
            (
                workbook.partname().clone(),
                flavor,
                catalog,
                sheet_parts,
                shared_strings_uri,
                styles_uri,
            )
        };

        let active_sheet = if catalog.sheets.is_empty() {
            None
        } else {
            Some(catalog.active_sheet_index)
        };
        let style_lineage = match source {
            Some(source) if package::same_style_table(source, &package, styles_uri.as_ref())? => {
                Arc::clone(&source.inner.style_lineage)
            },
            Some(_) | None => Arc::new(StyleLineage),
        };
        let shared_string_lineage = match source {
            Some(source)
                if package::same_shared_string_table(
                    source,
                    &package,
                    shared_strings_uri.as_ref(),
                )? =>
            {
                Arc::clone(&source.inner.shared_string_lineage)
            },
            Some(_) | None => Arc::new(SharedStringLineage),
        };
        let lineage_page_break_cache = source.map_or_else(
            || Arc::new(PageBreakProjectionCache::default()),
            |source| Arc::clone(&source.inner.page_break_cache),
        );
        let source_page_break_caches = page_break_caches(source);
        // A descendant reuses the lineage's cache; its allocation key keeps a
        // rewritten styles payload from reusing an earlier table.
        let format_cache = source.map_or_else(
            || Arc::new(FormatCache::new()),
            |source| Arc::clone(&source.inner.format_cache),
        );
        let chart_inventories = source.map_or_else(Arc::default, |source| {
            Arc::clone(&source.inner.chart_inventories)
        });
        let sheets = catalog
            .sheets
            .into_iter()
            .zip(sheet_parts)
            .enumerate()
            .map(|(position, (sheet, part))| {
                let name_key = crate::sheet::key(&sheet.name);
                let part_uri = part.uri;
                let page_breaks = page_break_cache(
                    &source_page_break_caches,
                    &package,
                    &part_uri,
                    &lineage_page_break_cache,
                )?;
                Ok(Arc::new(SheetData {
                    position,
                    name: sheet.name,
                    name_key,
                    kind: part.kind,
                    visibility: codec::visibility(sheet.visibility),
                    part_uri,
                    cells: OnceLock::new(),
                    web_bindings: OnceCell::new(),
                    page_breaks,
                    native_id: sheet.sheet_id,
                    relationship_id: sheet.relationship_id,
                }))
            })
            .collect::<Result<Vec<_>>>()?
            .into_boxed_slice();
        let mut sheet_name_order = (0..sheets.len()).collect::<Vec<_>>();
        sheet_name_order.sort_unstable_by(|left, right| {
            sheets[*left]
                .name_key
                .as_ref()
                .cmp(sheets[*right].name_key.as_ref())
                .then_with(|| left.cmp(right))
        });
        let sheet_name_order = sheet_name_order.into_boxed_slice();

        let policy = source
            .and_then(|source| source.policy.as_ref().map(Arc::clone))
            .or_else(|| {
                let semantic_limits = options.semantic_limits;
                (semantic_limits != crate::SemanticLimits::default()).then(|| {
                    Arc::new(Policy {
                        semantic_limits,
                        #[cfg(feature = "encryption")]
                        encryption: PackageEncryption::plain(),
                    })
                })
            });
        Ok(Self {
            inner: Arc::new(Inner {
                package,
                workbook_uri,
                shared_strings_uri,
                shared_strings: OnceLock::new(),
                shared_strings_tolerance: OnceLock::new(),
                rich_shared_strings: OnceLock::new(),
                shared_string_lineage,
                styles_uri,
                styles: OnceLock::new(),
                task_panes: OnceCell::new(),
                surveys: OnceCell::new(),
                style_lineage,
                flavor,
                date_system: if catalog.uses_1904_date_system {
                    DateSystem::Excel1904
                } else {
                    DateSystem::Excel1900
                },
                active_sheet,
                sheets,
                sheet_name_order,
                defined_names: catalog.defined_names.into_boxed_slice(),
                pivot_caches: catalog.pivot_caches.into_boxed_slice(),
                external_reference_ids: catalog.external_reference_ids.into_boxed_slice(),
                page_break_cache: lineage_page_break_cache,
                format_cache,
                chart_inventories,
                tolerance,
                options,
            }),
            policy,
        })
    }

    /// Publish commit-time worksheet validation into an otherwise-cold target
    /// snapshot only when every identity boundary remains exact.
    pub(super) fn adopt_validated_worksheet_stores(
        &self,
        source: &Self,
        validated: Vec<ValidatedWorksheetStore>,
    ) -> Result<()> {
        if !Arc::ptr_eq(&self.inner.style_lineage, &source.inner.style_lineage)
            || !Arc::ptr_eq(
                &self.inner.shared_string_lineage,
                &source.inner.shared_string_lineage,
            )
        {
            return Ok(());
        }

        for validated in validated {
            let source_sheet = source
                .inner
                .sheets
                .iter()
                .find(|sheet| sheet.part_uri == validated.uri);
            let target_sheet = self
                .inner
                .sheets
                .iter()
                .find(|sheet| sheet.part_uri == validated.uri);
            let (Some(source_sheet), Some(target_sheet)) = (source_sheet, target_sheet) else {
                continue;
            };
            if source_sheet.kind != WorksheetKind::Worksheet
                || target_sheet.kind != WorksheetKind::Worksheet
            {
                continue;
            }
            let published = self.inner.package.get_part(&validated.uri)?.blob_arc();
            if !Arc::ptr_eq(&published, &validated.content) {
                continue;
            }
            target_sheet
                .cells
                .set(validated.store)
                .map_err(|_| invalid("validated worksheet cache was already initialized"))?;
        }
        Ok(())
    }

    /// Every repair made while reading this workbook so far, in sheet order.
    ///
    /// The report merges the repairs made when the workbook was opened with
    /// those of every worksheet whose cells have been parsed, each stamped
    /// with its sheet position. Worksheets parse lazily, on first cell access;
    /// a worksheet not yet read contributes nothing until it is. An empty
    /// report means everything parsed so far conformed. See
    /// [`crate::leniency`] for the contract.
    #[must_use]
    pub fn tolerance_report(&self) -> ToleranceReport {
        let mut report = self.inner.tolerance.clone();
        if let Some(strings) = self.inner.shared_strings_tolerance.get() {
            report.append(strings);
        }
        for sheet in self.inner.sheets.iter() {
            if let Some(store) = sheet.cells.get() {
                report.append_sheet(store.tolerance_report(), sheet.position);
            }
        }
        report
    }

    /// Workbook flavor derived from package content, never its filename.
    #[must_use]
    pub fn flavor(&self) -> Flavor {
        self.inner.flavor
    }

    /// Date serial system used by the workbook.
    #[must_use]
    pub fn date_system(&self) -> DateSystem {
        self.inner.date_system
    }

    /// Number of logical workbook sheets, including chart and dialog sheets.
    #[must_use]
    pub fn len(&self) -> usize {
        self.inner.sheets.len()
    }

    /// Whether the workbook catalog contains no sheets.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.inner.sheets.is_empty()
    }

    /// Compare two snapshots' native OPC content without serializing either
    /// workbook.
    ///
    /// The comparison walks the package relationship graph and compares every
    /// accepted part's content type, payload bytes, and relationships.  It is
    /// therefore suitable for deciding whether a sequence of immutable edits
    /// returned to its original state (for example `A -> B -> A`) is a logical
    /// no-op.  Derived workbook caches and parser tolerance reports are not
    /// package content and are intentionally excluded.
    ///
    /// # Errors
    ///
    /// Returns the native package error when a deferred payload cannot be
    /// decoded. Packages containing caller-defined OPC part implementations,
    /// or separately opened opaque non-part ZIP members, return an explicit
    /// unsupported error because their behavior or bytes cannot be proved
    /// through the typed OPC graph.
    pub(crate) fn same_content(&self, other: &Self) -> Result<bool> {
        if Arc::ptr_eq(&self.inner, &other.inner) {
            return Ok(true);
        }

        // An equal authorized exact source digest is stronger than the
        // semantic walk below: it includes the complete archive, including
        // ZIP members which the OPC graph does not expose. A differing digest
        // is not enough to prove different logical content because ZIP
        // framing, member order, and archive comments may vary.
        if let (Some(left), Some(right)) = (
            self.inner.package.exact_source_sha256(),
            other.inner.package.exact_source_sha256(),
        ) {
            if left == right {
                return Ok(true);
            }
        }

        let left_package = &self.inner.package;
        let right_package = &other.inner.package;

        // A custom Part can carry behavior that is not represented by its
        // public byte/relationship view. Equality of that view would not be
        // a proof of equality of the package we publish.
        if !left_package.holds_only_built_in_parts() || !right_package.holds_only_built_in_parts() {
            return Err(Error::Unsupported {
                feature: "content comparison of custom OPC parts",
            });
        }

        // Non-part members are deliberately kept opaque by litchi-opc. A
        // difference in their names/reasons proves inequality. Descendants
        // of one immutable source archive retain a shared-owner proof for
        // their bytes; separately opened archives need an explicit refusal
        // because their non-part payloads are outside the typed graph.
        if left_package.non_part_members() != right_package.non_part_members() {
            return Ok(false);
        }
        if !left_package.non_part_members().is_empty()
            && !left_package.shares_source_archive(right_package)
        {
            return Err(Error::Unsupported {
                feature: "content comparison of opaque OPC members",
            });
        }

        if left_package.part_count() != right_package.part_count()
            || !same_relationships(left_package.rels(), right_package.rels())
        {
            return Ok(false);
        }

        for metadata in left_package.iter_parts() {
            let left = left_package.get_part(metadata.partname())?;
            let right = match right_package.get_part(metadata.partname()) {
                Ok(part) => part,
                Err(litchi_opc::OpcError::PartNotFound(_)) => return Ok(false),
                Err(error) => return Err(error.into()),
            };
            if right.partname() != metadata.partname()
                || right.content_type() != metadata.content_type()
                || right.blob() != left.blob()
                || !same_relationships(metadata.rels(), right.rels())
            {
                return Ok(false);
            }
        }

        // Content types and relationship members are package-level
        // publication manifests rather than ordinary OPC parts. Their
        // modeled fields were covered above; retained source bytes are the
        // remaining proof against producer extensions and skipped markup.
        left_package
            .same_retained_publication_manifests(right_package)
            .map_err(Into::into)
    }

    /// Iterate lightweight sheet handles in workbook order.
    #[must_use]
    pub fn sheets(&self) -> impl ExactSizeIterator<Item = Worksheet> + DoubleEndedIterator + '_ {
        self.inner.sheets.iter().cloned().map(|data| Worksheet {
            owner: Arc::clone(&self.inner),
            data,
            policy: self.policy.as_ref().map(Arc::clone),
        })
    }

    /// Look up a sheet by developer-facing name or checked zero-based position.
    pub fn sheet<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Option<Worksheet>> {
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
        Ok(data.map(|data| Worksheet {
            owner: Arc::clone(&self.inner),
            data,
            policy: self.policy.as_ref().map(Arc::clone),
        }))
    }

    /// Return the active sheet when the workbook contains sheets.
    #[must_use]
    pub fn active_sheet(&self) -> Option<Worksheet> {
        let data = self
            .inner
            .active_sheet
            .and_then(|position| self.inner.sheets.get(position))
            .cloned()?;
        Some(Worksheet {
            owner: Arc::clone(&self.inner),
            data,
            policy: self.policy.as_ref().map(Arc::clone),
        })
    }

    /// Low-level inert defined-name records retained by the catalog parser.
    #[must_use]
    pub fn defined_names(&self) -> &[raw::DefinedName] {
        &self.inner.defined_names
    }

    /// The worksheet table whose display name, or failing that whose name,
    /// equals `name` without regard to case, with the worksheet that owns it.
    ///
    /// Every worksheet is checked for ambiguous matches. A name match is
    /// used only when no worksheet has a display-name match. Worksheets are
    /// searched in workbook order and non-cell sheets are skipped.
    ///
    /// # Errors
    ///
    /// As for [`Worksheet::tables`] on each worksheet searched, or
    /// [`Error::Ambiguous`] when multiple names match.
    pub fn table(&self, name: &str) -> Result<Option<(Worksheet, worksheet::TablePart)>> {
        worksheet::find_workbook_table(
            self.sheets(),
            |sheet| {
                if sheet.kind() == WorksheetKind::Worksheet {
                    sheet.tables().map(Some)
                } else {
                    Ok(None)
                }
            },
            name,
        )
    }

    /// Low-level workbook pivot-cache references.
    #[must_use]
    pub fn pivot_caches(&self) -> &[raw::PivotCache] {
        &self.inner.pivot_caches
    }

    /// Return a semantic handle for one workbook PivotCache.  Cache and field
    /// resolution remains fallible at the typed read boundary; no OPC
    /// relationship or Part identity is accepted here.
    #[must_use]
    pub fn pivot_cache(
        &self,
        selector: impl Into<crate::pivot::PivotCacheSelector>,
    ) -> crate::pivot::PivotCacheHandle {
        crate::pivot::cached_unique_names::workbook_cache(self, selector)
    }

    /// Resolve the typed `pivotTableServerFormats` metadata for one semantic
    /// workbook PivotTable.  XML source ranges and package identities remain
    /// behind the low-level `pivot_table_server_formats_source` method.
    pub fn pivot_table_server_formats<'a>(
        &self,
        selector: impl Into<crate::pivot::PivotTableSelector<'a>>,
    ) -> Result<crate::pivot::PivotTableView> {
        crate::pivot::server_formats::view_workbook(self, selector)
    }

    /// Read the source-bound owner for low-level diagnostics and exact source
    /// patch construction.  Ordinary callers should use
    /// [`Self::pivot_table_server_formats`].
    pub fn pivot_table_server_formats_source<'a>(
        &self,
        selector: impl Into<crate::pivot::PivotTableSelector<'a>>,
    ) -> Result<crate::pivot::server_formats::Snapshot> {
        crate::pivot::server_formats::load(&self.inner.package, selector)
    }

    /// Resolve one semantic PivotTable through the ordinary workbook facade.
    ///
    /// The returned view exposes only typed server-format values.  Source XML,
    /// relationship IDs, and Part names remain behind the source-bound owner.
    pub fn pivot_table<'a>(
        &self,
        selector: impl Into<crate::pivot::PivotTableSelector<'a>>,
    ) -> Result<crate::pivot::PivotTableView> {
        crate::pivot::server_formats::view_workbook(self, selector)
    }

    /// Resolve the typed `pivotTableData` payload for one semantic
    /// non-worksheet PivotTable.  A missing C444 owner is reported as `None`;
    /// malformed or ambiguous owners remain fallible diagnostics.
    pub fn pivot_table_data<'a>(
        &self,
        selector: impl Into<crate::pivot::PivotTableSelector<'a>>,
    ) -> Result<Option<crate::pivot::PivotTableDataView>> {
        crate::pivot::table_data::view_workbook_with_limits(
            self,
            selector,
            &self.semantic_limits().pivot_table_data_value(),
        )
    }

    pub(crate) fn pivot_package(&self) -> &OpcPackage {
        &self.inner.package
    }

    /// The validated OPC graph behind this projection.
    pub(crate) fn opc_package(&self) -> &OpcPackage {
        &self.inner.package
    }

    /// The styles part, when the workbook has one.
    pub(crate) fn styles_part_uri(&self) -> Option<&PackURI> {
        self.inner.styles_uri.as_ref()
    }

    /// The OPC graph behind this projection, moved out when this is its only
    /// owner and cloned (sharing every part payload) otherwise.
    pub(crate) fn into_opc_package(self) -> OpcPackage {
        match Arc::try_unwrap(self.inner) {
            Ok(inner) => inner.package,
            Err(inner) => inner.package.clone(),
        }
    }

    pub(crate) fn adopt_published_package(&self, package: OpcPackage) -> Result<Self> {
        Self::from_package_with_styles(package, Some(self))
    }

    /// Inert external-workbook relationship IDs, for package diagnostics.
    #[must_use]
    pub fn external_reference_ids(&self) -> &[String] {
        &self.inner.external_reference_ids
    }

    /// Read the workbook's source-bound calculation metadata under the
    /// owning package's semantic limits
    /// ([`crate::SemanticLimits::calculation_properties`]).
    pub fn calculation_metadata(&self) -> Result<crate::calculation_properties::Snapshot> {
        crate::calculation_properties::Snapshot::load_with_limits(
            &self.inner.package,
            &self.semantic_limits().calculation_properties_value(),
        )
    }

    /// Persisted Office Add-in task panes, when this package contains them.
    ///
    /// The complete package graph is validated on first access. Later calls
    /// borrow the model retained by this immutable workbook snapshot.
    pub fn task_panes(&self) -> Result<Option<&litchi_ooxml_common::web::Panes>> {
        let panes = self.inner.task_panes.get_or_try_init(|| {
            litchi_ooxml_common::web::load(
                &self.inner.package,
                &litchi_ooxml_common::web::Limits::standard(),
            )
        })?;
        Ok(panes.as_ref())
    }

    /// Table-associated Excel survey metadata, when present.
    ///
    /// Surveys are parsed lazily as inert descriptive data. This API neither
    /// renders survey controls nor evaluates formulas, validates answers, or
    /// submits data to an external service.
    ///
    /// # Errors
    ///
    /// Returns an error when a survey part has an invalid table relationship,
    /// duplicate survey ID, or malformed Survey XML.
    pub fn surveys(&self) -> Result<&[crate::survey::Part]> {
        let values = self.inner.surveys.get_or_try_init(|| {
            crate::survey::load_with_limits(
                &self.inner.package,
                &self.semantic_limits().survey_value(),
            )
            .map(Vec::into_boxed_slice)
        })?;
        Ok(values)
    }

    /// Workbook-level protection metadata, when `workbookProtection` is
    /// present. The values are passive verifier and lock metadata; they do
    /// not enforce an editing policy.
    pub fn workbook_protection_metadata(
        &self,
    ) -> Result<Option<crate::workbook_metadata::protection::Metadata>> {
        let workbook = self.inner.package.get_part(&self.inner.workbook_uri)?;
        crate::workbook_metadata::protection::parse_workbook_protection(workbook.blob())
    }

    /// The workbook's people, the authors and mentions of threaded comments
    /// (record 0322); `None` without a persons part.
    ///
    /// # Errors
    ///
    /// Returns a typed parse error for a malformed persons part.
    pub fn people(&self) -> Result<Option<crate::threaded_comments::People>> {
        crate::threaded_comments::package::people(&self.inner.package)
    }

    /// Parse the styles part's differential formats, the formats conditional
    /// formatting rules reference by `dxfId`; empty without a styles part.
    ///
    /// # Errors
    ///
    /// Returns a typed parse error for malformed `dxfs` markup.
    pub fn differential_formats(&self) -> Result<Vec<crate::conditional_formatting::Differential>> {
        let Some(uri) = self.inner.styles_uri.as_ref() else {
            return Ok(Vec::new());
        };
        crate::conditional_formatting::parse_differential_formats(
            self.inner.package.get_part(uri)?.blob(),
        )
    }

    /// Shared immutable cell formats in this workbook snapshot.
    pub fn styles(&self) -> Result<Styles> {
        let len = self.inner.style_count()?;
        Ok(Styles::new(Arc::clone(&self.inner), len))
    }

    /// Release the parsed format table retained for [`Style::format`],
    /// unless an active [`crate::style::Format`] still uses it, and return
    /// whether a table was released.
    ///
    /// The table is a clean cached value: releasing it changes no result, and
    /// the next formatting read parses the styles part again. Snapshots
    /// derived from this workbook by edits share the same cache.
    pub fn release_cached_formats(&self) -> bool {
        self.inner.format_cache.trim_clean()
    }

    /// Remove every package signature from a copy of this snapshot; see
    /// [`crate::Package::unsign`].
    pub(crate) fn unsigned(&self) -> Result<Self> {
        self.ensure_mutation_allowed("unsign")?;
        let package = &self.inner.package;
        if !package.is_signed() && !package.requires_signature_edit_policy() {
            return Ok(self.clone());
        }
        let mut package = package.clone();
        package.unsign();
        Self::from_package_with_styles(package, Some(self))
    }

    pub(crate) fn ensure_ordinary_output(&self, operation: &'static str) -> Result<()> {
        #[cfg(feature = "encryption")]
        self.policy.as_deref().map_or_else(
            || Ok(()),
            |policy| {
                policy
                    .encryption
                    .ordinary_output()
                    .map_err(|source| Error::EncryptionPolicy { operation, source })
            },
        )?;
        #[cfg(not(feature = "encryption"))]
        let _ = operation;
        Ok(())
    }

    pub(crate) fn ensure_mutation_allowed(&self, operation: &'static str) -> Result<()> {
        self.ensure_ordinary_output(operation)
    }

    #[cfg(feature = "encryption")]
    pub(crate) fn retained_mode(&self, operation: &'static str) -> Result<crate::encryption::Mode> {
        self.policy
            .as_deref()
            .map_or_else(PackageEncryption::plain, |policy| policy.encryption)
            .require_retained_mode()
            .map_err(|source| Error::EncryptionPolicy { operation, source })
    }
}

/// Lightweight lifetime-free handle to one sheet in a workbook snapshot.
#[derive(Debug, Clone)]
pub struct Worksheet {
    pub(super) owner: Arc<Inner>,
    pub(super) data: Arc<SheetData>,
    /// The owning package's policy, for the form-control read budget
    /// ([`crate::SemanticLimits::form_controls`]); `None` for internal
    /// handles that never read form controls.
    pub(super) policy: Option<Arc<Policy>>,
}

impl Worksheet {
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

    /// Semantic sheet kind resolved from its relationship.
    #[must_use]
    pub fn kind(&self) -> WorksheetKind {
        self.data.kind
    }

    pub(crate) fn part_uri(&self) -> &PackURI {
        &self.data.part_uri
    }

    /// Retained visibility state.
    #[must_use]
    pub fn visibility(&self) -> &Visibility {
        &self.data.visibility
    }

    /// Whether this is the active sheet in its immutable workbook snapshot.
    #[must_use]
    pub fn is_active(&self) -> bool {
        self.owner.active_sheet == Some(self.data.position)
    }

    /// Whether two handles belong to the same immutable workbook snapshot.
    #[must_use]
    pub fn same_workbook(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.owner, &other.owner)
    }

    /// Parse this worksheet's typed inert hyperlinks.
    pub fn hyperlinks(&self) -> Result<Vec<crate::hyperlinks::Hyperlink>> {
        worksheet::hyperlinks(self)
    }

    /// Read the worksheet's admitted SpreadsheetML form controls.
    pub fn form_controls(&self) -> Result<crate::form_control::FormControlCollection> {
        if self.data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: self.data.name.clone(),
            });
        }
        crate::form_control::eager_form_controls_for_sheet_with_limits(
            &self.owner.package,
            &self.data.part_uri,
            self.policy
                .as_ref()
                .map_or_else(
                    || self.owner.options.semantic_limits.form_controls_value(),
                    |policy| policy.semantic_limits.form_controls_value(),
                )
                .unwrap_or_default(),
        )
        .map_err(crate::form_control::owner_to_xlsx)
    }

    /// Parse this worksheet's legacy notes; `None` without a comments part.
    ///
    /// # Errors
    ///
    /// [`Error::NotWorksheet`] for a chart or dialog sheet, and a typed error
    /// for a malformed comments part or relationship.
    pub fn comments(&self) -> Result<Option<crate::workbook::comments::Comments>> {
        if self.data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: self.data.name.clone(),
            });
        }
        Ok(crate::workbook::comments::load_from_worksheet(
            &self.owner.package,
            &self.data.part_uri,
        )?
        .map(|part| part.comments))
    }

    /// Read this worksheet's threaded comments (record 0322); `None`
    /// without a threaded-comments part. Only the worksheet's relationships
    /// and that part are read; authors resolve through
    /// [`crate::Workbook::people`].
    pub fn threaded_comments(&self) -> Result<Option<crate::threaded_comments::Comments>> {
        if self.data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: self.data.name.clone(),
            });
        }
        crate::threaded_comments::package::comments(&self.owner.package, &self.data.part_uri)
    }

    /// Resolve one form control by checked position or exact authored name.
    pub fn form_control<'a>(
        &self,
        selector: impl Into<crate::form_control::ControlSelector<'a>>,
    ) -> Result<Option<crate::form_control::FormControlView>> {
        Ok(self.form_controls()?.get(selector)?.cloned())
    }

    /// Borrow this worksheet's Office Add-in range bindings.
    ///
    /// Bindings are decoded and validated on first access, then retained by
    /// the immutable sheet snapshot. Non-worksheet handles return a typed
    /// error instead of attempting to interpret another sheet kind as XML.
    pub fn web_bindings(&self) -> Result<&crate::web::Bindings> {
        if self.data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: self.data.name.clone(),
            });
        }
        self.data.web_bindings.get_or_try_init(|| {
            let part = self.owner.package.get_part(&self.data.part_uri)?;
            raw::web::read(part.blob(), &self.owner.package.read_limits())
        })
    }

    /// Parse the worksheet's direct auto-filter declaration.
    pub fn auto_filter(&self) -> Result<Option<crate::auto_filter::Definition>> {
        worksheet::auto_filter(self)
    }

    /// Parse all conditional-formatting containers and associate their
    /// differential formats with the workbook style resource.
    pub fn conditional_formattings(
        &self,
    ) -> Result<Vec<crate::conditional_formatting::Formatting>> {
        worksheet::conditional_formattings(self)
    }

    /// Parse worksheet data-validation collections.
    pub fn data_validations(&self) -> Result<Vec<crate::data_validation::Collection>> {
        worksheet::data_validations(self)
    }

    /// Parse the worksheet's optional data-consolidation declaration.
    pub fn data_consolidation(
        &self,
    ) -> Result<Option<crate::data_consolidation::DataConsolidation>> {
        worksheet::data_consolidation(self)
    }

    /// Parse core worksheet header/footer settings.
    pub fn header_footer(&self) -> Result<Option<crate::header_footer::Settings>> {
        worksheet::header_footer(self)
    }

    /// Parse ignored-error declarations.
    pub fn ignored_errors(&self) -> Result<Option<crate::ignored_errors::IgnoredErrors>> {
        worksheet::ignored_errors(self)
    }

    /// Parse this worksheet's inert smart-tag annotations.
    ///
    /// The metadata is never resolved to an action provider or executed.
    pub fn smart_tags(&self) -> Result<Option<crate::smart_tags::Collection>> {
        worksheet::smart_tags(self)
    }

    /// Parse the optional named-sheet-view relationship.
    pub fn named_sheet_views(&self) -> Result<Option<crate::named_sheet_view::Views>> {
        worksheet::named_sheet_views(self)
    }

    /// Parse worksheet outline properties.
    pub fn outline_properties(
        &self,
    ) -> Result<Option<crate::outline_properties::OutlineProperties>> {
        worksheet::outline_properties(self)
    }

    /// Parse worksheet page margins.
    pub fn page_margins(&self) -> Result<Option<crate::page_margins::Margins>> {
        worksheet::page_margins(self)
    }

    /// Parse the direct horizontal and vertical worksheet page breaks.
    ///
    /// # Errors
    ///
    /// Returns an error when the worksheet XML or page-break values are
    /// malformed, unsafe, or out of bounds.
    pub fn page_breaks(&self) -> Result<crate::page_breaks::PageBreaks> {
        worksheet::page_breaks(self)
    }

    /// Parse worksheet page setup.
    pub fn page_setup(&self) -> Result<Option<crate::page_setup::Setup>> {
        worksheet::page_setup(self)
    }

    /// Parse worksheet phonetic properties.
    pub fn phonetic_properties(
        &self,
    ) -> Result<Option<crate::phonetic_properties::PhoneticProperties>> {
        worksheet::phonetic_properties(self)
    }

    /// Parse worksheet print options.
    pub fn print_options(&self) -> Result<Option<crate::print_options::PrintOptions>> {
        worksheet::print_options(self)
    }

    /// Parse worksheet what-if scenarios.
    pub fn scenarios(&self) -> Result<Option<crate::scenarios::Collection>> {
        worksheet::scenarios(self)
    }

    /// Parse worksheet-level calculation properties.
    pub fn calculation_properties(
        &self,
    ) -> Result<Option<crate::sheet_calculation_properties::Properties>> {
        worksheet::calculation_properties(self)
    }

    /// Parse the complete worksheet protection projection.
    pub fn protection(&self) -> Result<crate::sheet_protection::Metadata> {
        worksheet::protection(self)
    }

    /// Parse the worksheet's ordinary sheet-view collection.
    pub fn views(&self) -> Result<Option<crate::sheet_view::Collection>> {
        worksheet::views(self)
    }

    /// Parse every table relationship owned by this worksheet.
    pub fn tables(&self) -> Result<Vec<worksheet::TablePart>> {
        worksheet::tables(self)
    }

    /// The table whose display name, or failing that whose name, equals
    /// `name` without regard to case.
    ///
    /// Display names are unique in a workbook, and table names within a
    /// sheet, both compared without regard to case, so a well-formed
    /// package has at most one match.
    ///
    /// # Errors
    ///
    /// As for [`Self::tables`], or [`Error::Ambiguous`] when multiple names match.
    pub fn table(&self, name: &str) -> Result<Option<worksheet::TablePart>> {
        worksheet::find_table(self.tables()?, name)
    }

    /// Load and validate every query-table part owned by this worksheet.
    pub fn query_tables(&self) -> Result<Vec<crate::query_table::Part>> {
        worksheet::query_tables(self)
    }

    /// Load this worksheet's inert `ActiveX` graph.
    pub fn active_x(&self) -> Result<crate::active_x::ControlSet> {
        worksheet::active_x(self)
    }

    /// Load timeline views associated with this worksheet.
    pub fn timelines(&self) -> Result<Vec<crate::timeline::Part>> {
        worksheet::timelines(self)
    }

    /// Return array-formula anchors without materializing a dense worksheet.
    pub fn array_formulas(&self) -> Result<Vec<worksheet::ArrayFormula>> {
        worksheet::array_formulas(self)
    }

    /// Look up one exact logical cell state by A1 or checked coordinate.
    ///
    /// [`crate::cell::View::Missing`], [`crate::cell::View::Covered`], and a
    /// producer-stored [`crate::Cell`] are distinct without materializing follower
    /// cells for large merged ranges.
    pub fn cell<'a>(&self, at: impl Into<At<'a>>) -> Result<crate::cell::View<'_>> {
        let address = at.into().resolve()?;
        Ok(self.store()?.view(address))
    }

    /// Exact local style state for a stored cell.
    ///
    /// `None` means no cell record exists. [`LocalStyle::Default`] means the
    /// record exists without an explicit shared-style reference.
    pub fn local_style<'a>(&self, at: impl Into<At<'a>>) -> Result<Option<LocalStyle>> {
        let address = at.into().resolve()?;
        let Some(entry) = self.store()?.entry(address) else {
            return Ok(None);
        };
        entry.style.map_or(Ok(Some(LocalStyle::Default)), |key| {
            self.owner.require_style(key)?;
            Ok(Some(LocalStyle::Shared(Style::from_raw(
                Arc::clone(&self.owner),
                key,
            ))))
        })
    }

    /// Effective shared style for a stored cell.
    ///
    /// Cells without a local style resolve to the base shared format. If the
    /// workbook has no style part, an unstyled cell resolves to `None`.
    pub fn style<'a>(&self, at: impl Into<At<'a>>) -> Result<Option<Style>> {
        let address = at.into().resolve()?;
        let Some(entry) = self.store()?.entry(address) else {
            return Ok(None);
        };
        let key = entry.style.unwrap_or(0);
        if self.owner.style_count()? == 0 {
            return Ok(None);
        }
        self.owner.require_style(key)?;
        Ok(Some(Style::from_raw(Arc::clone(&self.owner), key)))
    }

    /// Whether the stored text at `at` carries rich-text runs: an inline
    /// string with `r` runs, or a shared string whose item has them.
    /// Writing a plain value to such a cell drops the runs' formatting, so
    /// a text search refuses to rewrite it (record 0273).
    ///
    /// # Errors
    ///
    /// The errors of reading the worksheet's cell store or the
    /// shared-string part.
    pub fn has_rich_text<'a>(&self, at: impl Into<At<'a>>) -> Result<bool> {
        let address = at.into().resolve()?;
        let Some(entry) = self.store()?.entry(address) else {
            return Ok(false);
        };
        if entry.inline_rich {
            return Ok(true);
        }
        let Some(index) = entry.shared_string else {
            return Ok(false);
        };
        Ok(self
            .owner
            .rich_shared_strings()?
            .binary_search(&index)
            .is_ok())
    }

    /// The rich-text runs of the shared string at `at`, each with its run
    /// properties as stored (record 0295). `None` for a cell that is not a
    /// shared string with runs, or whose item a run rewrite cannot keep
    /// (phonetic runs, nested run-property markup). An inline rich string
    /// is read from the worksheet part (one scan of it).
    ///
    /// # Errors
    ///
    /// The errors of reading the worksheet's cell store or the
    /// shared-string part.
    pub fn rich_text<'a>(&self, at: impl Into<At<'a>>) -> Result<Option<crate::RichText>> {
        let address = at.into().resolve()?;
        let Some(entry) = self.store()?.entry(address) else {
            return Ok(None);
        };
        if entry.inline_rich {
            let part = self.owner.package.get_part(&self.data.part_uri)?;
            return inline_rich(part.blob(), &address.a1());
        }
        let Some(index) = entry.shared_string else {
            return Ok(None);
        };
        if self
            .owner
            .rich_shared_strings()?
            .binary_search(&index)
            .is_err()
        {
            return Ok(None);
        }
        let Some(uri) = self.owner.shared_strings_uri.as_ref() else {
            return Ok(None);
        };
        rich_item(self.owner.package.get_part(uri)?.blob(), index)
    }

    /// Read rich text for a bounded set of cells with one requested payload
    /// pass per physical source part. Shared-string callers retain the
    /// existing first-use rich-index census before that payload pass.
    ///
    /// The scalar [`Self::rich_text`] accessor is intentionally exact for one
    /// address, but its inline and shared-string readers both start at the
    /// beginning of their XML member. Transfers ask for a range of cells, so
    /// restarting that scan for every address turns a rich range into a
    /// quadratic walk. This helper retains only requested run values: cells
    /// outside `addresses` are traversed by the XML reader but never parsed
    /// into a `RichText` value.
    pub(crate) fn rich_text_many(
        &self,
        addresses: &[Address],
    ) -> Result<HashMap<Address, Option<crate::RichText>>> {
        let mut values = HashMap::new();
        values
            .try_reserve(addresses.len())
            .map_err(|source| Error::Allocation {
                resource: "rich text values",
                source,
            })?;
        if addresses.is_empty() {
            return Ok(values);
        }

        let store = self.store()?;
        let mut inline = HashMap::new();
        inline
            .try_reserve(addresses.len())
            .map_err(|source| Error::Allocation {
                resource: "inline rich-text addresses",
                source,
            })?;
        let mut shared = HashSet::new();
        shared
            .try_reserve(addresses.len())
            .map_err(|source| Error::Allocation {
                resource: "rich shared-string indexes",
                source,
            })?;
        let mut shared_addresses = Vec::new();
        shared_addresses
            .try_reserve(addresses.len())
            .map_err(|source| Error::Allocation {
                resource: "rich shared-string addresses",
                source,
            })?;
        for &address in addresses {
            values.entry(address).or_insert(None);
            let Some(entry) = store.entry(address) else {
                continue;
            };
            if entry.inline_rich {
                // The scalar reader compares the raw `r` attribute with the
                // canonical address string. Keep that exact source policy in
                // the bulk reader instead of normalizing `$A$1` or lowercase
                // references through `Address::from_a1`.
                inline
                    .entry(address.a1().into_bytes().into_boxed_slice())
                    .or_insert(address);
            } else if let Some(index) = entry.shared_string {
                shared.insert(index);
                shared_addresses.push((address, index));
            }
        }

        if !inline.is_empty() {
            let part = self.owner.package.get_part(&self.data.part_uri)?;
            values.extend(inline_rich_many(part.blob(), &inline)?);
        }

        if !shared.is_empty() {
            // The cached rich-index census keeps ordinary shared-string
            // transfers on the existing no-part-read path. Only when a
            // requested index is actually rich do we make the one payload
            // pass that projects its runs.
            let rich_indexes = self.owner.rich_shared_strings()?;
            shared.retain(|index| rich_indexes.binary_search(index).is_ok());
            if shared.is_empty() {
                return Ok(values);
            }
            let Some(uri) = self.owner.shared_strings_uri.as_ref() else {
                return Ok(values);
            };
            let part = self.owner.package.get_part(uri)?;
            let items = rich_items_many(part.blob(), &shared)?;
            for (address, index) in shared_addresses {
                if let Some(value) = items.get(&index) {
                    values.insert(address, value.clone());
                }
            }
        }

        Ok(values)
    }

    /// Lazily traverse stored cells selected by A1 range, raw zero-based
    /// half-open bounds, or a reusable checked rectangle.
    pub fn cells<'a>(&self, area: impl Into<Area<'a>>) -> Result<Cells<'_>> {
        let range = area.into().resolve()?;
        Ok(self.store()?.cells(range))
    }

    /// The stored record at a position [`Cells::indexed`] reported for this
    /// worksheet, with its coordinate, in O(1); `None` past the last record.
    ///
    /// # Errors
    ///
    /// The errors of reading the worksheet's cell store.
    pub fn record(
        &self,
        index: crate::RecordIndex,
    ) -> Result<Option<(litchi_sheet::Cell, &crate::Cell)>> {
        Ok(self.store()?.record(index))
    }

    /// Lazily traverse validated, non-overlapping merged ranges.
    pub fn merges(&self) -> Result<crate::merge::Merges<'_>> {
        Ok(self.store()?.merges())
    }

    /// Borrow one checked logical row, including an implicit default row.
    pub fn row(&self, at: impl Into<RowAt>) -> Result<Row<'_>> {
        let index = at.into().resolve()?;
        Ok(self.store()?.row(index))
    }

    /// Lazily traverse only explicit worksheet row records.
    pub fn rows(&self) -> Result<Rows<'_>> {
        Ok(self.store()?.rows())
    }

    /// Exact shared-style state contributed by a stored row record.
    ///
    /// `None` means the logical row is implicit. [`LocalStyle::Default`]
    /// means an explicit record applies without a shared-style reference.
    pub fn row_style(&self, at: impl Into<RowAt>) -> Result<Option<LocalStyle>> {
        let index = at.into().resolve()?;
        let Some(entry) = self.store()?.row_entry(index) else {
            return Ok(None);
        };
        entry
            .properties
            .style
            .map_or(Ok(Some(LocalStyle::Default)), |key| {
                self.owner.require_style(key)?;
                Ok(Some(LocalStyle::Shared(Style::from_raw(
                    Arc::clone(&self.owner),
                    key,
                ))))
            })
    }

    /// Borrow one checked logical column, including an implicit default
    /// column. A1 labels such as `"B"` are the primary entry; raw inputs are
    /// zero-based and validated before lookup.
    pub fn column<'a>(&self, at: impl Into<ColumnAt<'a>>) -> Result<Column<'_>> {
        let index = at.into().resolve()?;
        Ok(self.store()?.column(index))
    }

    /// Lazily traverse logical columns covered by explicit property records.
    /// Overlapping producer records have already been resolved using Excel's
    /// last-record-wins semantics.
    pub fn columns(&self) -> Result<Columns<'_>> {
        Ok(self.store()?.columns())
    }

    /// Stored worksheet-grid defaults, if the producer supplied them.
    ///
    /// Absence is preserved rather than guessing a font-dependent row height
    /// or column width.
    pub fn defaults(&self) -> Result<Option<&crate::layout::Defaults>> {
        Ok(self.store()?.defaults())
    }

    /// Exact shared-style state contributed by a column-property record.
    ///
    /// `None` means the logical column is implicit. [`LocalStyle::Default`]
    /// means an explicit record applies without a shared-style reference.
    pub fn column_style<'a>(&self, at: impl Into<ColumnAt<'a>>) -> Result<Option<LocalStyle>> {
        let index = at.into().resolve()?;
        let Some(entry) = self.store()?.column_entry(index) else {
            return Ok(None);
        };
        entry
            .properties
            .style
            .map_or(Ok(Some(LocalStyle::Default)), |key| {
                self.owner.require_style(key)?;
                Ok(Some(LocalStyle::Shared(Style::from_raw(
                    Arc::clone(&self.owner),
                    key,
                ))))
            })
    }

    /// Distinct declared, stored, content, and directly styled cell bounds.
    pub fn extents(&self) -> Result<&Extents> {
        Ok(self.store()?.extents())
    }

    /// Bounding rectangle of stored cell records, distinct from declared,
    /// formatted, and content extents.
    pub fn stored_extent(&self) -> Result<Option<Rect>> {
        Ok(self.store()?.extents().stored())
    }

    /// The repairs made while parsing this worksheet's cells.
    ///
    /// This parses the worksheet if it has not been read yet. Locations are
    /// relative to this worksheet, so [`crate::ToleratedDefect::sheet`] is
    /// `None`. A chart, dialog, or other non-cell sheet reports clean.
    pub fn tolerance_report(&self) -> Result<&ToleranceReport> {
        if self.data.kind != WorksheetKind::Worksheet {
            return Ok(&crate::leniency::CLEAN);
        }
        Ok(self.store()?.tolerance_report())
    }

    pub(super) fn store(&self) -> Result<&Store> {
        if self.data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: self.data.name.clone(),
            });
        }
        if let Some(store) = self.data.cells.get() {
            return Ok(store);
        }

        let part = self.owner.package.get_part(&self.data.part_uri)?;
        let mut parsed = raw::worksheet::parse(part.blob(), || self.owner.shared_strings())?;
        self.owner.repair_styles(&mut parsed)?;
        let _publish_result = self.data.cells.set(parsed);
        self.data
            .cells
            .get()
            .ok_or_else(|| invalid("worksheet cache initialization did not publish a value"))
    }
}

impl Inner {
    pub(super) fn shared_strings(&self) -> Result<Option<&[Text]>> {
        let Some(uri) = self.shared_strings_uri.as_ref() else {
            return Ok(None);
        };
        if let Some(strings) = self.shared_strings.get() {
            return Ok(Some(strings));
        }

        let part = self.package.get_part(uri)?;
        let mut tolerance = ToleranceReport::default();
        let parsed = raw::strings::parse_repairing(part.blob(), &mut tolerance)?;
        // The report is published first, so a published table always has
        // its report; a racing parse of the same bytes reports the same.
        let _publish_result = self.shared_strings_tolerance.set(tolerance);
        let _publish_result = self.shared_strings.set(parsed);
        self.shared_strings
            .get()
            .map(|strings| Some(strings.as_ref()))
            .ok_or_else(|| invalid("shared-string cache initialization did not publish a value"))
    }

    /// The ascending indexes of the shared-string items with rich runs.
    fn rich_shared_strings(&self) -> Result<&[usize]> {
        if let Some(indexes) = self.rich_shared_strings.get() {
            return Ok(indexes);
        }
        let indexes = match self.shared_strings_uri.as_ref() {
            Some(uri) => rich_items(self.package.get_part(uri)?.blob())?,
            None => Box::default(),
        };
        let _publish_result = self.rich_shared_strings.set(indexes);
        self.rich_shared_strings
            .get()
            .map(AsRef::as_ref)
            .ok_or_else(|| {
                invalid("rich shared-string cache initialization did not publish a value")
            })
    }

    fn style_catalog(&self) -> Result<Option<&raw::styles::Catalog>> {
        let Some(uri) = self.styles_uri.as_ref() else {
            return Ok(None);
        };
        if let Some(styles) = self.styles.get() {
            return Ok(Some(styles));
        }

        let part = self.package.get_part(uri)?;
        let parsed = raw::styles::parse(part.blob())?;
        let _publish_result = self.styles.set(parsed);
        self.styles
            .get()
            .map(Some)
            .ok_or_else(|| invalid("style cache initialization did not publish a value"))
    }

    pub(crate) fn style_count(&self) -> Result<u32> {
        Ok(self.style_catalog()?.map_or(0, raw::styles::Catalog::len))
    }

    /// Readable formatting of shared cell format `key`, read through the
    /// lineage's format-table cache.
    pub(crate) fn format(&self, key: u32) -> Result<Format> {
        let uri = self
            .styles_uri
            .as_ref()
            .ok_or_else(|| invalid("workbook has no styles part"))?;
        let payload = self.package.get_part(uri)?.blob_arc();
        let limits = self.package.read_limits();
        let table = self
            .format_cache
            .get_or_build(FormatKey::allocation(&payload), || {
                FormatTable::parse_with_limits(&payload, limits)
            })?;
        Format::new(table, key)
    }

    fn require_style(&self, key: u32) -> Result<()> {
        let len = self.style_count()?;
        if key >= len {
            return Err(invalid(format!(
                "worksheet cell references shared style {key}, but the workbook contains {len} cell formats"
            )));
        }
        Ok(())
    }

    /// Repair stale style indexes in a freshly parsed worksheet store before
    /// it is published to the read cache (audit finding 4.2).
    ///
    /// Each repair is recorded in the store's own tolerance report. The
    /// commit-time verification paths keep using `validate_styles`, which
    /// still refuses a stale index: an edit must not make the repair durable.
    pub(super) fn repair_styles(&self, store: &mut Store) -> Result<()> {
        if !store.references_styles() {
            return Ok(());
        }
        store.repair_style_indexes(self.style_count()?)
    }

    pub(super) fn validate_styles(&self, store: &Store) -> Result<()> {
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

    pub(crate) fn style_fan_out(self: &Arc<Self>, key: u32) -> Result<usize> {
        self.require_style(key)?;
        let mut count = 0usize;
        for data in &self.sheets {
            if data.kind != WorksheetKind::Worksheet {
                continue;
            }
            let sheet = Worksheet {
                owner: Arc::clone(self),
                data: Arc::clone(data),
                policy: None,
            };
            count = count
                .checked_add(
                    sheet
                        .store()?
                        .entries()
                        .iter()
                        .filter(|entry| entry.style.unwrap_or(0) == key)
                        .count(),
                )
                .ok_or_else(|| invalid("shared style fan-out count overflowed usize"))?;
        }
        Ok(count)
    }
}

/// The ascending indexes of the `si` items of a shared-string part that
/// hold `r` runs.
fn rich_items(xml: &[u8]) -> Result<Box<[usize]>> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut depth = 0usize;
    let mut item = None::<usize>;
    let mut count = 0usize;
    let mut rich = Vec::new();
    loop {
        match reader
            .read_event()
            .map_err(|error| invalid(format!("malformed shared-string part: {error}")))?
        {
            quick_xml::events::Event::Start(element) => {
                depth += 1;
                let local = element.local_name();
                if depth == 2 && local.as_ref() == b"si" {
                    item = Some(count);
                    count += 1;
                } else if depth == 3
                    && local.as_ref() == b"r"
                    && let Some(index) = item
                    && rich.last() != Some(&index)
                {
                    rich.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "rich shared-string indexes",
                        source,
                    })?;
                    rich.push(index);
                }
            },
            quick_xml::events::Event::Empty(element) => {
                if depth == 1 && element.local_name().as_ref() == b"si" {
                    count += 1;
                }
            },
            quick_xml::events::Event::End(_) => {
                if depth == 2 {
                    item = None;
                }
                depth = depth.saturating_sub(1);
            },
            quick_xml::events::Event::Eof => break,
            _ => {},
        }
    }
    Ok(rich.into_boxed_slice())
}

/// The runs of shared-string item `wanted`, each with its `rPr` children
/// as `(local name, raw attributes)`; `None` when the item has no runs or
/// holds what a run rewrite would drop (phonetic data, nested `rPr`
/// markup, text outside the runs).
pub(crate) fn rich_item(xml: &[u8], wanted: usize) -> Result<Option<crate::RichText>> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut depth = 0usize;
    let mut count = 0usize;
    loop {
        match reader
            .read_event()
            .map_err(|error| invalid(format!("malformed shared-string part: {error}")))?
        {
            Event::Start(element) => {
                depth += 1;
                if depth == 2 && element.local_name().as_ref() == b"si" {
                    if count == wanted {
                        return read_runs(&mut reader, false).map(RunsRead::into_value);
                    }
                    count += 1;
                }
            },
            Event::Empty(element) => {
                if depth == 1 && element.local_name().as_ref() == b"si" {
                    if count == wanted {
                        return Ok(None);
                    }
                    count += 1;
                }
            },
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::Eof => return Ok(None),
            _ => {},
        }
    }
}

/// Read the requested shared-string items in one XML pass. The item indexes
/// are the workbook's stable shared-string identities; unrelated items are
/// skipped without constructing run values.
pub(crate) fn rich_items_many(
    xml: &[u8],
    wanted: &HashSet<usize>,
) -> Result<HashMap<usize, Option<crate::RichText>>> {
    let mut values = HashMap::new();
    values
        .try_reserve(wanted.len())
        .map_err(|source| Error::Allocation {
            resource: "rich shared-string values",
            source,
        })?;
    for &index in wanted {
        values.insert(index, None);
    }
    if wanted.is_empty() {
        return Ok(values);
    }

    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut depth = 0usize;
    let mut count = 0usize;
    loop {
        match reader
            .read_event()
            .map_err(|error| invalid(format!("malformed shared-string part: {error}")))?
        {
            Event::Start(element) => {
                depth += 1;
                if depth == 2 && element.local_name().as_ref() == b"si" {
                    let index = count;
                    count += 1;
                    if wanted.contains(&index) {
                        let value = match read_runs(&mut reader, true)? {
                            RunsRead::Complete(value) => value,
                            // Scalar lookup returns `None` as soon as it
                            // sees an unsupported run shape. If the remainder
                            // is malformed, stop this bulk projection rather
                            // than turning that established `None` into an
                            // error while chasing later addresses.
                            RunsRead::Aborted => return Ok(values),
                        };
                        values.insert(index, value);
                        // A complete read consumes the closing `</si>`.
                        depth = 1;
                    }
                }
            },
            Event::Empty(element) => {
                if depth == 1 && element.local_name().as_ref() == b"si" {
                    let index = count;
                    count += 1;
                    if wanted.contains(&index) {
                        values.insert(index, None);
                    }
                }
            },
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::Eof => return Ok(values),
            _ => {},
        }
    }
}

/// The runs of the inline string of cell `a1` in worksheet `xml` (record
/// 0295), as [`rich_item`] reads an item. One scan of the worksheet part.
pub(crate) fn inline_rich(xml: &[u8], a1: &str) -> Result<Option<crate::RichText>> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut in_cell = false;
    let mut depth = 0usize;
    let mut cell_depth = 0usize;
    loop {
        match reader
            .read_event()
            .map_err(|error| invalid(format!("malformed worksheet part: {error}")))?
        {
            Event::Start(element) => {
                depth += 1;
                let local = element.local_name();
                if local.as_ref() == b"c" {
                    in_cell = element
                        .try_get_attribute("r")
                        .ok()
                        .flatten()
                        .is_some_and(|attribute| attribute.value.as_ref() == a1.as_bytes());
                    cell_depth = depth;
                } else if in_cell && depth == cell_depth + 1 && local.as_ref() == b"is" {
                    return read_runs(&mut reader, false).map(RunsRead::into_value);
                }
            },
            Event::End(_) => {
                if in_cell && depth == cell_depth {
                    return Ok(None);
                }
                depth = depth.saturating_sub(1);
            },
            Event::Eof => return Ok(None),
            _ => {},
        }
    }
}

/// Read requested inline rich cells in one worksheet XML pass. The parser
/// keeps only the requested addresses and leaves unrelated inline strings as
/// raw events, preserving the scalar reader's exact run/failure behavior.
pub(crate) fn inline_rich_many(
    xml: &[u8],
    wanted: &HashMap<Box<[u8]>, Address>,
) -> Result<HashMap<Address, Option<crate::RichText>>> {
    let mut values = HashMap::new();
    values
        .try_reserve(wanted.len())
        .map_err(|source| Error::Allocation {
            resource: "inline rich-text values",
            source,
        })?;
    for &address in wanted.values() {
        values.insert(address, None);
    }
    if wanted.is_empty() {
        return Ok(values);
    }

    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut in_cell = None;
    let mut depth = 0usize;
    let mut cell_depth = 0usize;
    loop {
        match reader
            .read_event()
            .map_err(|error| invalid(format!("malformed worksheet part: {error}")))?
        {
            Event::Start(element) => {
                depth += 1;
                let local = element.local_name();
                if local.as_ref() == b"c" {
                    in_cell = element
                        .try_get_attribute("r")
                        .ok()
                        .flatten()
                        .and_then(|attribute| wanted.get(attribute.value.as_ref()).copied());
                    cell_depth = depth;
                } else if in_cell.is_some() && depth == cell_depth + 1 && local.as_ref() == b"is" {
                    let address = in_cell.expect("inline rich cell is present");
                    let value = match read_runs(&mut reader, true)? {
                        RunsRead::Complete(value) => value,
                        RunsRead::Aborted => return Ok(values),
                    };
                    values.insert(address, value);
                    // A complete read consumes the closing `</is>`.
                    depth = cell_depth;
                }
            },
            Event::End(_) => {
                if in_cell.is_some() && depth == cell_depth {
                    in_cell = None;
                }
                depth = depth.saturating_sub(1);
            },
            Event::Eof => return Ok(values),
            _ => {},
        }
    }
}

/// Read the runs of the string item whose start tag `reader` just read, up
/// to its end tag.
enum RunsRead {
    Complete(Option<crate::RichText>),
    Aborted,
}

impl RunsRead {
    fn into_value(self) -> Option<crate::RichText> {
        match self {
            Self::Complete(value) => value,
            Self::Aborted => None,
        }
    }
}

fn read_runs(reader: &mut quick_xml::Reader<&[u8]>, consume_invalid: bool) -> Result<RunsRead> {
    use quick_xml::events::Event;
    let attributes = |element: &quick_xml::events::BytesStart<'_>| -> Result<Box<str>> {
        let raw = std::str::from_utf8(element.attributes_raw())
            .map_err(|error| invalid(format!("string attributes are not UTF-8: {error}")))?;
        Ok(raw.trim().into())
    };
    let local = |element: &quick_xml::events::BytesStart<'_>| -> Result<Box<str>> {
        std::str::from_utf8(element.local_name().as_ref())
            .map(Into::into)
            .map_err(|error| invalid(format!("string element name is not UTF-8: {error}")))
    };
    let mut runs: Vec<crate::RichRun> = Vec::new();
    let mut properties: Option<Vec<(Box<str>, Box<str>)>> = None;
    let mut text = String::new();
    // Depth below the item and where the reader is: 0 the item, 1 a run,
    // 2 its `rPr`, 3 its `t`.
    let mut depth = 0usize;
    let mut place = 0u8;
    let mut valid = true;
    loop {
        let event = match reader.read_event() {
            Ok(event) => event,
            Err(_error) if consume_invalid && !valid => return Ok(RunsRead::Aborted),
            Err(error) => return Err(invalid(format!("malformed string item: {error}"))),
        };
        match event {
            Event::Start(element) => {
                depth += 1;
                if !valid {
                    if !consume_invalid {
                        return Ok(RunsRead::Complete(None));
                    }
                    continue;
                }
                match (depth, element.local_name().as_ref(), place) {
                    (1, b"r", 0) => {
                        place = 1;
                        properties = None;
                        text.clear();
                    },
                    (2, b"rPr", 1) => {
                        place = 2;
                        properties = Some(Vec::new());
                    },
                    (2, b"t", 1) => place = 3,
                    // A plain `t` beside runs, `rPh`, `phoneticPr` or
                    // children with content are not kept by a run rewrite.
                    _ if consume_invalid => valid = false,
                    _ => return Ok(RunsRead::Complete(None)),
                }
            },
            Event::Empty(element) => {
                if !valid {
                    if !consume_invalid {
                        return Ok(RunsRead::Complete(None));
                    }
                    continue;
                }
                match (depth, place) {
                    (2, 2) => {
                        if let Some(properties) = properties.as_mut() {
                            properties.push((local(&element)?, attributes(&element)?));
                        }
                    },
                    (1, 1) if element.local_name().as_ref() == b"rPr" => {
                        properties = Some(Vec::new());
                    },
                    (1, 1) if element.local_name().as_ref() == b"t" => {},
                    _ if consume_invalid => valid = false,
                    _ => return Ok(RunsRead::Complete(None)),
                }
            },
            Event::Text(value) if valid && place == 3 => text.push_str(
                &value
                    .decode()
                    .map_err(|error| invalid(format!("string text: {error}")))?,
            ),
            Event::CData(value) if valid && place == 3 => text.push_str(
                &value
                    .decode()
                    .map_err(|error| invalid(format!("string text: {error}")))?,
            ),
            Event::GeneralRef(reference) if valid && place == 3 => {
                text.push_str(&litchi_ooxml_common::xml::decode_xml_reference(&reference)?);
            },
            Event::End(_) => {
                if depth == 0 {
                    return Ok(RunsRead::Complete(
                        (valid && !runs.is_empty()).then(|| crate::RichText::from_runs(runs)),
                    ));
                }
                if !valid {
                    if !consume_invalid {
                        return Ok(RunsRead::Complete(None));
                    }
                    depth = depth.saturating_sub(1);
                    continue;
                }
                match (depth, place) {
                    (2, 2 | 3) => place = 1,
                    (1, 1) => {
                        let decoded = raw::strings::decode_spreadsheet_text(&text)?;
                        runs.push(crate::RichRun::with_properties(
                            properties.take().map(Vec::into_boxed_slice),
                            decoded,
                        ));
                        place = 0;
                    },
                    _ => {},
                }
                depth = depth.saturating_sub(1);
            },
            Event::Eof => {
                return Ok(if consume_invalid {
                    RunsRead::Aborted
                } else {
                    RunsRead::Complete(None)
                });
            },
            _ => {},
        }
    }
}

#[cfg(test)]
mod rich_item_tests {
    use std::collections::{HashMap, HashSet};

    use litchi_sheet::Cell as Address;

    #[test]
    fn bulk_rich_readers_keep_requested_runs_only() {
        let inline = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>outside</t></r></is></c><c r="B1" t="inlineStr"><is><r><rPr><i/></rPr><t>inside</t></r></is></c><c r="C1"><v>3</v></c></row></sheetData></worksheet>"#;
        let requested = HashMap::from([
            (
                b"B1".to_vec().into_boxed_slice(),
                Address::from_a1("B1").expect("B1"),
            ),
            (
                b"C1".to_vec().into_boxed_slice(),
                Address::from_a1("C1").expect("C1"),
            ),
        ]);
        let values = super::inline_rich_many(inline, &requested).expect("inline values");
        assert_eq!(
            values[&Address::from_a1("B1").expect("B1")]
                .as_ref()
                .expect("requested runs")
                .text(),
            "inside"
        );
        assert!(values[&Address::from_a1("C1").expect("C1")].is_none());
        assert!(!values.contains_key(&Address::from_a1("A1").expect("A1")));

        // Bulk lookup keeps the scalar reader's raw-address policy: a
        // lowercase source attribute is not normalized into the requested
        // canonical address.
        let noncanonical = br#"<worksheet><sheetData><row><c r="b1" t="inlineStr"><is><r><t>lowercase</t></r></is></c></row></sheetData></worksheet>"#;
        let requested = HashMap::from([(
            b"B1".to_vec().into_boxed_slice(),
            Address::from_a1("B1").expect("B1"),
        )]);
        let values = super::inline_rich_many(noncanonical, &requested).expect("raw addresses");
        assert!(values[&Address::from_a1("B1").expect("B1")].is_none());

        // An unsupported run shape still yields the scalar reader's `None`
        // without turning a malformed trailing tail into a new parse error.
        let malformed_tail =
            br#"<worksheet><sheetData><row><c r="A1" t="inlineStr"><is><r><rPh/></is></c><broken"#;
        let requested = HashMap::from([(
            b"A1".to_vec().into_boxed_slice(),
            Address::from_a1("A1").expect("A1"),
        )]);
        let values = super::inline_rich_many(malformed_tail, &requested)
            .expect("unsupported rich shape remains non-fatal");
        assert!(values[&Address::from_a1("A1").expect("A1")].is_none());

        let shared = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>outside</t></si><si><r><rPr><b/></rPr><t>inside</t></r></si><si><r><rPr><i/></rPr><t>also outside</t></r></si></sst>"#;
        let requested = HashSet::from([1usize]);
        let values = super::rich_items_many(shared, &requested).expect("shared values");
        assert_eq!(
            values[&1].as_ref().expect("requested runs").text(),
            "inside"
        );
        assert!(!values.contains_key(&2));
    }

    #[test]
    fn rich_items_lists_the_items_with_runs() {
        let xml = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4"><si><t>plain</t></si><si><r><rPr><b/></rPr><t>rich</t></r><r><t> text</t></r></si><si/><si><t>x</t><rPh sb="0" eb="1"><t>y</t></rPh></si><si><r><t>z</t></r></si></sst>"#;
        assert_eq!(&*super::rich_items(xml).unwrap(), &[1, 4]);
        assert!(super::rich_items(b"<sst").is_err());
    }

    #[test]
    fn rich_item_keeps_each_runs_properties() {
        let xml = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>plain</t></si><si><r><rPr><b/><sz val="11"/><color rgb="FFFF0000"/></rPr><t>rich &amp; </t></r><r><t xml:space="preserve">text_x0009_</t></r></si><si><t>x</t><rPh sb="0" eb="1"><t>y</t></rPh></si><si><r><t>a</t></r><rPh sb="0" eb="1"><t>y</t></rPh></si></sst>"#;
        assert_eq!(super::rich_item(xml, 0).unwrap(), None);
        let rich = super::rich_item(xml, 1).unwrap().unwrap();
        assert_eq!(rich.text(), "rich & text\t");
        assert_eq!(rich.runs().len(), 2);
        assert!(rich.runs()[0].is_formatted());
        assert!(!rich.runs()[1].is_formatted());
        let properties = rich.runs()[0].properties.as_deref().unwrap();
        assert_eq!(&*properties[1].0, "sz");
        assert_eq!(&*properties[1].1, r#"val="11""#);
        assert_eq!(super::rich_item(xml, 2).unwrap(), None);
        assert_eq!(super::rich_item(xml, 3).unwrap(), None);
        assert_eq!(super::rich_item(xml, 9).unwrap(), None);
    }
}

#[cfg(test)]
mod rich_text_write_tests {
    use crate::{Package, RichText};

    /// A workbook whose `A1` shows a formatted shared string.
    fn rich_workbook() -> crate::Result<Package> {
        let mut edit = Package::new()?.edit()?;
        edit.sheet("Sheet1")?
            .expect("worksheet")
            .set_rich_text("A1", RichText::plain("x"))?;
        edit.commit()
            .map(crate::workbook::edit::Commit::into_package)
    }

    #[test]
    fn a_rewritten_rich_string_keeps_each_runs_properties() -> crate::Result<()> {
        // A shared-string item with two runs, through the reader.
        let item = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><rPr><b/><sz val="11"/></rPr><t>Hello </t></r><r><rPr><i/></rPr><t>world</t></r></si></sst>"#;
        let mut rich = super::rich_item(item, 0)?.expect("rich item");
        rich.set_run_text(0, "Bye ")?;
        let mut edit = rich_workbook()?.edit()?;
        edit.sheet("Sheet1")?
            .expect("worksheet")
            .set_rich_text("B2", rich)?;
        let package = edit.commit()?.into_package();
        let workbook = package.workbook();
        let sheet = workbook.sheet("Sheet1")?.expect("worksheet");
        assert!(sheet.has_rich_text("B2")?);
        let read = sheet.rich_text("B2")?.expect("inline runs");
        assert_eq!(read.runs().len(), 2);
        assert_eq!(read.text(), "Bye world");
        assert!(read.runs()[1].is_formatted());
        let mut edit = package.edit()?;
        let sheet = edit.sheet("Sheet1")?.expect("worksheet");
        let crate::State::Cell {
            content: crate::Cell::Value(crate::Value::Text(text)),
            ..
        } = sheet.state("B2")?
        else {
            panic!("a text cell");
        };
        assert_eq!(text.as_str(), "Bye world");
        Ok(())
    }

    /// Typed run fonts (record 0338): reads of stored markup, changes that
    /// keep the properties they do not name, and the source-backed read.
    #[test]
    fn run_fonts_read_and_change_in_place() -> crate::Result<()> {
        use crate::style::{FontChange, FontColor, Script, Underline};
        use crate::{Rgb, RichRun};

        let item = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><rPr><b/><i val="0"/><sz val="11"/><color theme="1" tint="-0.25"/><rFont val="Calibri &amp; Co"/><family val="2"/><scheme val="minor"/></rPr><t>A</t></r><r><t>B</t></r></si></sst>"#;
        let mut rich = super::rich_item(item, 0)?.expect("rich item");
        let font = rich.runs()[0].font()?;
        assert!(font.bold && !font.italic);
        assert_eq!(font.size, Some(11.0));
        assert_eq!(font.name.as_deref(), Some("Calibri & Co"));
        assert_eq!(font.family, Some(2));
        let color = font.color.expect("theme color");
        assert_eq!(color.kind(), crate::style::FontColorKind::Theme(1));
        assert_eq!(color.tint().map(|tint| tint.get()), Some(-0.25));
        assert!(!rich.runs()[1].font()?.bold);

        let run = rich.run_mut(0).expect("first run");
        run.apply_font(
            &FontChange::new()
                .bold(None)
                .underline(Some(Underline::Double))
                .script(Some(Script::Superscript))
                .color(Some(FontColor::rgb(Rgb::new(0xC0, 0, 0))))
                .size(Some(14.5))?,
        )?;
        run.set_text("Z");
        let font = run.font()?;
        assert!(!font.bold);
        assert_eq!(font.underline, Some(Underline::Double));
        assert_eq!(font.script, Some(Script::Superscript));
        assert_eq!(font.size, Some(14.5));
        assert_eq!(font.family, Some(2), "an unnamed property is kept");
        assert_eq!(font.name.as_deref(), Some("Calibri & Co"));
        let mut plain = RichRun::new("C");
        plain.apply_font(&FontChange::new().italic(Some(true)))?;
        rich.insert_run(2, plain)?;
        assert!(rich.insert_run(9, RichRun::new("x")).is_err());
        assert_eq!(rich.remove_run(1)?.text(), "B");
        assert_eq!(rich.text(), "ZC");

        let mut edit = rich_workbook()?.edit()?;
        edit.sheet("Sheet1")?
            .expect("worksheet")
            .set_rich_text("C3", rich.clone())?;
        let bytes = edit.commit()?.into_package().to_bytes()?;
        let reopened = Package::from_bytes(&bytes)?;
        let read = reopened
            .workbook()
            .sheet("Sheet1")?
            .expect("worksheet")
            .rich_text("C3")?
            .expect("runs");
        assert_eq!(read, rich);
        // The source-backed worksheet reads the same runs.
        let source = crate::SourceBackedWorkbook::from_reader(std::io::Cursor::new(bytes))?;
        let sheet = source.sheet("Sheet1")?.expect("worksheet");
        assert_eq!(sheet.rich_text("C3")?, Some(rich));
        assert_eq!(
            sheet.rich_text("A1")?.map(|rich| rich.text()),
            Some("x".to_owned())
        );
        assert_eq!(sheet.rich_text("Z9")?, None);
        // Removing every property leaves an unformatted run.
        let mut run = RichRun::new("y");
        run.apply_font(&FontChange::new().bold(Some(true)))?;
        assert!(run.is_formatted());
        run.apply_font(&FontChange::new().bold(None))?;
        assert!(!run.is_formatted());
        Ok(())
    }
}
