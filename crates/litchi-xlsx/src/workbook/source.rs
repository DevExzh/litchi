//! Immutable XLSX reads backed by a caller-provided positional source.
//!
//! This facade intentionally does not adapt into [`super::Workbook`]: that
//! snapshot owns a mutable OPC graph, while this type must keep ordinary part
//! payloads deferred. It exposes only semantic catalog and worksheet reads.
//!
//! Parsed worksheet values (a sheet's grid and the XML outside its
//! `sheetData`) are built once per worksheet and kept in a bounded, weighted
//! clean-value cache; see [`SourceBackedWorkbook::semantic_cache_diagnostics`].
//!
//! Managed [`ExecutionContext`] opens are supported for this read-only
//! facade. Focused source-backed editors retain managed
//! [`litchi_opc::PartData`] handles in their immutable snapshots and expose
//! their own managed constructors; owning-package patch application remains a
//! separate, explicitly fallible boundary when it would require detaching an
//! allocation from that reservation.

use std::borrow::Cow;
use std::collections::HashMap;
use std::io::Read;
use std::mem::size_of;
use std::ops::ControlFlow;
#[cfg(any(unix, windows))]
use std::path::Path;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex, MutexGuard, OnceLock, PoisonError};

#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::{
    ExecutionContext, ExecutionError, ReadAt, Selector as CoreSelector, SourceVersion,
};
use litchi_ooxml_common::mce::{Capabilities, Error as MceError, StreamError, StreamLimits};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    PackURI, PartData, PartView, SourceBackedPackage, SourceCacheDiagnostics, SourceCacheLimits,
    SourceLimits, SourceReadPolicy, VerifiedDecodedReaderError,
};
use litchi_sheet::{Area, At, Cell as Address, Rect};

use super::{DateSystem, Flavor, Selector, Visibility, WorksheetKind, codec};
use crate::cell::{Cell, Store, Text, Value, View};
use crate::error::{Error, Result, allocation, invalid};
use crate::leniency::ToleranceReport;
use crate::raw;
use crate::raw::selected_worksheet::SelectedPayload;

mod styles;
mod tables;

mod blocks;
mod cursor;
pub(crate) mod deferred;
mod grid;
#[cfg(test)]
mod m1_tests;
mod values;

pub use cursor::CellCursor;
pub use grid::Grid;
pub use styles::Record;
pub use values::CacheDiagnostics;

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

pub(crate) struct SourceSheetData {
    position: usize,
    /// The sheet name, its canonical case-insensitive key and the workbook
    /// relationship ID, in one allocation; `name_len` and `name_key_len`
    /// split them.
    keys: Box<str>,
    pub(crate) kind: WorksheetKind,
    pub(crate) visibility: Visibility,
    pub(crate) part_uri: PackURI,
    /// The validated workbook catalog coordinates retained for the tab-state
    /// source closure. Keeping these values avoids reparsing the same
    /// workbook XML when an already admitted read-only owner starts an edit.
    pub(crate) sheet_id: u32,
    name_key_len: u16,
    /// Sheet names are at most 31 characters, so at most 124 UTF-8 bytes.
    name_len: u8,
    /// Repairs made by the first successful parse of this worksheet's grid.
    ///
    /// The parsed grid itself lives in the workbook's evictable semantic
    /// cache; its repair report is kept here so the workbook's report stays
    /// independent of what the cache currently retains. A report holds at
    /// most a bounded number of entries.
    tolerance: OnceLock<ToleranceReport>,
    /// Serializes the builds of this worksheet's parsed values, so readers
    /// racing for the same cold value share one build instead of parsing the
    /// part once each. Reads of other worksheets are never blocked.
    build: Mutex<()>,
}

impl SourceSheetData {
    /// Build the combined name, name-key and relationship-ID storage.
    fn keys(name: &str, relationship_id: &str) -> Result<(Box<str>, u8, u16)> {
        let bound = || Error::Invalid("worksheet name exceeds its storage bound".into());
        let name_len = u8::try_from(name.len()).map_err(|_overflow| bound())?;
        let mut keys = if name.is_ascii() {
            let mut keys = String::with_capacity(2 * name.len() + relationship_id.len());
            keys.push_str(name);
            keys.extend(name.chars().map(|character| character.to_ascii_lowercase()));
            keys
        } else {
            let key = crate::sheet::key(name);
            let mut keys = String::with_capacity(name.len() + key.len() + relationship_id.len());
            keys.push_str(name);
            keys.push_str(&key);
            keys
        };
        let name_key_len = u16::try_from(keys.len() - name.len()).map_err(|_overflow| bound())?;
        keys.push_str(relationship_id);
        Ok((keys.into_boxed_str(), name_len, name_key_len))
    }

    /// The sheet name.
    pub(crate) fn name(&self) -> &str {
        &self.keys[..usize::from(self.name_len)]
    }

    /// The canonical case-insensitive key of the sheet name.
    pub(crate) fn name_key(&self) -> &str {
        let start = usize::from(self.name_len);
        &self.keys[start..start + usize::from(self.name_key_len)]
    }

    /// The workbook relationship ID of this sheet.
    pub(crate) fn relationship_id(&self) -> &str {
        &self.keys[usize::from(self.name_len) + usize::from(self.name_key_len)..]
    }

    fn lock_build(&self) -> MutexGuard<'_, ()> {
        // The guarded value is empty, so a poisoned lock carries no state.
        self.build.lock().unwrap_or_else(PoisonError::into_inner)
    }
}

#[derive(Debug)]
pub(crate) struct SourceCatalogProof {
    pub(crate) defined_names: Arc<[raw::DefinedName]>,
    pub(crate) limits: Arc<SourceLimits>,
}

/// The source and the resolved state of a source-backed edit.
///
/// Source-backed edit entry points must not turn a cheap immutable owner into
/// a selected-part snapshot merely because a caller is about to stage an edit.
/// This context therefore retains only the owner until the first operation
/// that needs the selected source.  `OnceLock` also makes repeated accessors
/// share the one authenticated snapshot without retaining a decoded payload
/// in every discarded edit.
pub(crate) struct LazySource<S> {
    workbook: SourceBackedWorkbook,
    resolved: OnceLock<S>,
}

impl<S> LazySource<S> {
    pub(crate) fn new(workbook: SourceBackedWorkbook) -> Self {
        Self {
            workbook,
            resolved: OnceLock::new(),
        }
    }

    pub(crate) fn workbook(&self) -> &SourceBackedWorkbook {
        &self.workbook
    }

    pub(crate) fn resolve(
        &self,
        load: impl FnOnce(&SourceBackedPackage) -> Result<S>,
    ) -> Result<&S> {
        if let Some(resolved) = self.resolved.get() {
            return Ok(resolved);
        }
        let package = self.workbook.edit_package()?;
        let resolved = load(package)?;
        // A concurrently resolving clone may win the one-time publication.
        // Its snapshot is equivalent under the same immutable source policy;
        // discard this duplicate and borrow the winner without panicking.
        let _ = self.resolved.set(resolved);
        self.resolved
            .get()
            .ok_or_else(|| invalid("lazy source snapshot was not retained"))
    }
}

/// Make a selector owned so a deferred edit can outlive the caller's input.
///
/// XLSX currently has no stable-identity selector variant (`Id` is
/// `Infallible`), but matching it explicitly keeps this helper correct if the
/// alias gains one in a future API revision.
pub(crate) fn own_selector(selector: Selector<'_>) -> Result<Selector<'static>> {
    match selector {
        CoreSelector::Name(name) => Ok(CoreSelector::Name(Cow::Owned(name.into_owned()))),
        CoreSelector::Position(position) => Ok(CoreSelector::Position(position)),
        CoreSelector::Id(never) => match never {},
        _ => Err(Error::UnsupportedSelector),
    }
}

/// Caches used only by shared-string and conditional-format readers.
///
/// Keeping these less common caches separate means the ordinary cell/style
/// path does not retain their state in its first lazy allocation.
struct DependencyState {
    shared_strings: OnceLock<Box<[Text]>>,
    shared_strings_tolerance: OnceLock<ToleranceReport>,
    differential_format_count: OnceLock<usize>,
}

impl DependencyState {
    fn new() -> Self {
        Self {
            shared_strings: OnceLock::new(),
            shared_strings_tolerance: OnceLock::new(),
            differential_format_count: OnceLock::new(),
        }
    }
}

/// The source-backed policy that is meaningful after the OPC package has
/// already admitted its physical source limits.
///
/// `OpenOptions` also carries the 192-byte `SourceLimits` value. The package
/// owns that value, so retaining a complete `OpenOptions` in every immutable
/// source owner duplicates the physical ingress policy in each owner and in
/// every reopened history step. Keep the facade policies here and rebuild the
/// public options value only at a deferred materialization boundary. Package
/// cache and execution policy are retained separately from an optional facade
/// context: an external package remains authoritative for its own cache and
/// budget, while a facade context adds a second cancellation/budget fence.
#[derive(Clone)]
struct SourceOwnerOptions {
    leniency: crate::Leniency,
    /// Default semantic budgets need no per-owner storage. Customized
    /// budgets are shared only when a deferred owner actually needs them.
    semantic_limits: Option<Arc<crate::SemanticLimits>>,
    #[cfg(feature = "encryption")]
    encryption_limits: crate::encryption::Limits,
    cache_limits: Option<SourceCacheLimits>,
    /// The execution context that owns the OPC package cache and publication
    /// budget. This is the package's policy, not a facade override.
    package_execution_context: Option<Arc<ExecutionContext>>,
    /// An optional facade-level context supplied alongside an already-opened
    /// package. It remains an additional fence when the package has its own
    /// managed context.
    facade_execution_context: Option<Arc<ExecutionContext>>,
}

impl SourceOwnerOptions {
    fn from_open_options(options: &crate::OpenOptions) -> Self {
        let execution_context = options
            .execution_context
            .as_ref()
            .map(|context| Arc::new(context.clone()));
        Self {
            leniency: options.leniency,
            semantic_limits: (options.semantic_limits != crate::SemanticLimits::default())
                .then(|| Arc::new(options.semantic_limits)),
            #[cfg(feature = "encryption")]
            encryption_limits: options.encryption_limits,
            cache_limits: options.cache_limits,
            package_execution_context: execution_context.clone(),
            facade_execution_context: execution_context,
        }
    }

    fn from_source_backed_package(
        package: &SourceBackedPackage,
        options: &crate::OpenOptions,
    ) -> Self {
        let mut owner_options = Self::from_open_options(options);
        let cache_limits = package.cache_limits();
        owner_options.cache_limits =
            (cache_limits != SourceCacheLimits::default()).then_some(cache_limits);
        owner_options.package_execution_context = package.execution_context().map(Arc::new);
        owner_options
    }

    fn to_open_options(&self, source_limits: SourceLimits) -> crate::OpenOptions {
        // `package_execution_context` is supplied separately through
        // `shared_owner_options` when the reopened package is constructed.
        // The public options value carries only the facade context so it
        // cannot replace the package's own cache/budget policy.
        let mut options = crate::OpenOptions::new()
            .source_limits(source_limits)
            .leniency(self.leniency)
            .semantic_limits(*self.semantic_limits());
        #[cfg(feature = "encryption")]
        {
            options = options.encryption_limits(self.encryption_limits);
        }
        if let Some(cache_limits) = self.cache_limits {
            options = options.cache_limits(cache_limits);
        }
        if let Some(execution_context) = self.facade_execution_context.as_deref() {
            options = options.execution_context(execution_context.clone());
        }
        options
    }

    fn semantic_limits(&self) -> &crate::SemanticLimits {
        static DEFAULT: OnceLock<crate::SemanticLimits> = OnceLock::new();
        self.semantic_limits
            .as_deref()
            .unwrap_or_else(|| DEFAULT.get_or_init(crate::SemanticLimits::default))
    }

    fn is_default(&self) -> bool {
        self.leniency == crate::Leniency::default()
            && self.semantic_limits.is_none()
            && self.cache_limits.is_none()
            && self.package_execution_context.is_none()
            && self.facade_execution_context.is_none()
            && {
                #[cfg(feature = "encryption")]
                {
                    self.encryption_limits == crate::encryption::Limits::default()
                }
                #[cfg(not(feature = "encryption"))]
                {
                    true
                }
            }
    }
}

pub(crate) struct SourceInner {
    package: SourceBackedPackage,
    /// One exact replay endpoint shared by every patch that observes this
    /// opened source owner. Adjacent deferred patches otherwise retain the
    /// same endpoint state twice: the earlier patch's `after` and the next
    /// patch's `before` are the same pending owner.
    pub(crate) restore_endpoint: OnceLock<Arc<deferred::RestoreEndpoint>>,
    /// The immutable tab-state source closure captured for this owner. A
    /// source-backed edit does not mutate that closure, so repeated edits can
    /// share its parsed graph, relationships and source payload metadata.
    pub(crate) tab_state_snapshot: OnceLock<Arc<crate::tab_state::Snapshot>>,
    /// The first edit on an owner is commonly a one-shot admission probe. Do
    /// not retain its snapshot solely because the edit was dropped; arm the
    /// immutable cache after the first successful snapshot load so repeated
    /// edits still share the closure without adding one to every new owner.
    pub(crate) tab_state_snapshot_seen: AtomicBool,
    /// The source-backed tab graph is already admitted by the workbook open.
    /// Keep its exact validation lazy for the edit path, while sharing the
    /// resulting graph across snapshots once a commit needs it.
    pub(crate) tab_state_graph: OnceLock<Arc<crate::tab_state::GraphBinding>>,
    /// One shared package handle for exact replay endpoints. Keeping this
    /// lazy leaves ordinary read-only opens allocation-free while letting a
    /// retained patch refer to one package owner per source lineage instead
    /// of embedding a full `SourceBackedPackage` value in every endpoint.
    package_handle: OnceLock<Arc<SourceBackedPackage>>,
    /// Immutable merge-family source closures keyed by worksheet URI. Source
    /// backed edits repeatedly reopen one worksheet from the same immutable
    /// owner; sharing its workbook and relationship metadata keeps retained
    /// patch history compact without sharing rewritten worksheet bytes.
    merge_source_states:
        OnceLock<Box<Mutex<HashMap<PackURI, Arc<crate::merge::source::SourceStateShared>>>>>,
    execution: Option<Arc<ExecutionContext>>,
    // Keep the mandatory workbook payload pinned for managed facades. Besides
    // avoiding a second root extraction, this makes an exact managed budget
    // boundary meaningful: the selected worksheet cannot evict the root while
    // callers still hold the workbook snapshot. Compatibility facades retain
    // the historical finite-cache behavior and do not pin this payload.
    _catalog_data: Option<PartData>,
    shared_strings_uri: Option<PackURI>,
    styles_uri: Option<PackURI>,
    /// Style identity remains part of the owner lineage. The common worksheet
    /// and style caches stay inline so their first reads do not allocate a
    /// separate reader-state box; less common dependency caches remain lazy.
    style_lineage: Arc<crate::style::StyleLineage>,
    values: values::ValueCache,
    styles: OnceLock<raw::styles::Catalog>,
    formats: crate::style::cell_format::FormatCache,
    /// Shared-string and conditional-format caches are less common than the
    /// worksheet/style path, so keep their state in a second lazy allocation.
    dependencies: OnceLock<Box<DependencyState>>,
    flavor: Flavor,
    date_system: DateSystem,
    pub(crate) active_sheet: Option<usize>,
    pub(crate) sheets: Box<[Arc<SourceSheetData>]>,
    /// Sheet positions sorted by their canonical, case-insensitive names.
    ///
    /// The mandatory workbook catalog rejects duplicate canonical names, so a
    /// sorted position index is equivalent to the previous linear search while
    /// avoiding a second allocation of every key string.
    sheet_name_order: Box<[usize]>,
    /// Repairs made while the workbook catalog was parsed at open time.
    tolerance: ToleranceReport,
    /// Defined names parsed with the mandatory workbook catalog at open.
    ///
    /// An ordinary owner keeps the parsed catalogue as a boxed slice (no
    /// shared-ownership header on the cold open path). An owner reopened
    /// with a shared catalog proof stores `None` and borrows the proof's
    /// slice, so every proof-backed owner of a lineage reads one shared copy;
    /// [`Self::catalog_proof`] copies an ordinary owner's names at most once,
    /// when that owner first starts a deferred lineage.
    defined_names: Option<Box<[raw::DefinedName]>>,
    /// Lazily shared ingress policy for source-backed snapshots. Keeping this
    /// on the immutable lineage avoids copying all policy fields into every
    /// retained before/after snapshot while leaving ordinary opens lazy.
    source_limits: OnceLock<Arc<SourceLimits>>,
    /// The immutable catalog proof shared by tab-state snapshots in this
    /// source lineage. It groups the unchanged defined-name catalog and its
    /// ingress policy behind one retained handle.
    catalog_proof: OnceLock<Arc<SourceCatalogProof>>,
    /// The ingress policies this workbook was opened with; a committed
    /// workbook reopens its output under them.
    owner_options: Option<Arc<SourceOwnerOptions>>,
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
    /// The pending output of a committed workbook (record 0360): `inner` is
    /// then the base the commit was made on.
    pending: Option<Arc<deferred::Deferred>>,
}

/// A lifetime-free read-only worksheet handle from [`SourceBackedWorkbook`].
///
/// The first read that needs the worksheet's parsed grid builds it in one
/// pass over the part and offers it to the workbook's bounded semantic cache
/// (see [`SourceBackedWorkbook::semantic_cache_diagnostics`]). The handle then
/// pins that grid: its later reads, and those of its clones, are lookups in
/// the pinned grid, and the grid stays resident until the handle and its
/// clones are dropped, whatever the cache evicts meanwhile.
#[derive(Clone)]
pub struct SourceWorksheet {
    owner: Arc<SourceInner>,
    data: Arc<SourceSheetData>,
    grid: OnceLock<Grid>,
}

impl std::fmt::Debug for SourceBackedWorkbook {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("SourceBackedWorkbook")
            .field("sheets", &self.current().sheets.len())
            .field("committed", &self.pending.is_some())
            .finish_non_exhaustive()
    }
}

// Shared handles: the unified operations layer keeps these in `Arc`-backed,
// `Send + Sync` file handles (record 0003 risk R3).
const _: () = {
    const fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<SourceBackedWorkbook>();
    assert_send_sync::<SourceWorksheet>();
};

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
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_options(path, crate::OpenOptions::default())
    }

    /// Open a filesystem-backed XLSX package under explicit ingress
    /// policies: source limits, deferred-payload cache limits, execution
    /// context, leniency and semantic limits ([`crate::OpenOptions`]).
    #[cfg(any(unix, windows))]
    pub fn open_with_options(path: impl AsRef<Path>, options: crate::OpenOptions) -> Result<Self> {
        Self::from_read_at_with_options(file_source(path)?, options)
    }

    /// Open an ordinary XLSX package from a caller-provided positional source.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_options(source, crate::OpenOptions::default())
    }

    /// Open from a positional source under explicit ingress policies.
    ///
    /// With an execution context, the context checks cancellation before
    /// mandatory open reads and before every deferred semantic
    /// materialization, and retained and in-flight payloads are charged to
    /// its hierarchical memory budget; managed [`litchi_opc::PartData`]
    /// handles never escape as unbudgeted `Arc` allocations through this
    /// facade. Without one, the deferred-payload cache is bounded by the
    /// selected [`SourceCacheLimits`] (or the OPC default) but unmanaged.
    pub fn from_read_at_with_options(
        source: Arc<dyn ReadAt>,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        Self::from_read_at_with_options_and_shared_catalog(source, options, None, None)
    }

    fn from_read_at_with_options_and_shared_catalog(
        source: Arc<dyn ReadAt>,
        options: crate::OpenOptions,
        shared_catalog: Option<Arc<SourceCatalogProof>>,
        shared_owner_options: Option<Arc<SourceOwnerOptions>>,
    ) -> Result<Self> {
        let limits = options.source_limits;
        let cache_limits = match shared_owner_options.as_ref() {
            Some(owner_options) => owner_options.cache_limits.unwrap_or_default(),
            None => options.cache_limits.unwrap_or_default(),
        };
        let owner_options = shared_owner_options.or_else(|| {
            let owner_options = SourceOwnerOptions::from_open_options(&options);
            (!owner_options.is_default()).then(|| Arc::new(owner_options))
        });
        let execution = owner_options
            .as_ref()
            .and_then(|owner_options| owner_options.package_execution_context.as_deref().cloned());
        let package = match execution {
            Some(context) => SourceBackedPackage::from_read_at(
                source,
                limits,
                cache_limits,
                SourceReadPolicy::exact(),
                Some(context),
            )?,
            None => SourceBackedPackage::from_read_at(
                source,
                limits,
                cache_limits,
                SourceReadPolicy::exact(),
                None,
            )?,
        };
        Self::open_source_backed_package_with_shared_catalog(
            package,
            options,
            shared_catalog,
            owner_options,
        )
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
        Self::from_reader_with_options(reader, crate::OpenOptions::default())
    }

    /// Open an XLSX package from a sequential reader under explicit ingress
    /// policies. Input ingestion and later deferred payload reads use the
    /// same source limits.
    pub fn from_reader_with_options<R: Read>(
        reader: R,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        let owner_options = SourceOwnerOptions::from_open_options(&options);
        let owner_options = (!owner_options.is_default()).then(|| Arc::new(owner_options));
        let package = SourceBackedPackage::from_reader(
            reader,
            options.source_limits,
            options.cache_limits.unwrap_or_default(),
            options.execution_context.clone(),
        )?;
        Self::open_source_backed_package_with_shared_catalog(package, options, None, owner_options)
    }

    /// Build the read-only XLSX facade from a validated deferred OPC package.
    pub fn from_source_backed_package(package: SourceBackedPackage) -> Result<Self> {
        Self::from_source_backed_package_with_options(package, crate::OpenOptions::default())
    }

    /// [`Self::from_source_backed_package`] under explicit policies. The
    /// package keeps its source, cache limits and package execution context;
    /// `options` contributes the leniency, semantic limits and an optional
    /// facade execution context. When both contexts are present, both are
    /// checked and the package context continues to govern OPC cache and
    /// publication admission.
    pub fn from_source_backed_package_with_options(
        package: SourceBackedPackage,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        Self::open_source_backed_package(package, options)
    }

    fn open_source_backed_package(
        package: SourceBackedPackage,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        Self::open_source_backed_package_with_shared_catalog(package, options, None, None)
    }

    fn open_source_backed_package_with_shared_catalog(
        package: SourceBackedPackage,
        options: crate::OpenOptions,
        shared_catalog: Option<Arc<SourceCatalogProof>>,
        shared_owner_options: Option<Arc<SourceOwnerOptions>>,
    ) -> Result<Self> {
        let facade_execution = options.execution_context.clone();
        let leniency = options.leniency;
        package.check_execution()?;
        check_execution(facade_execution.as_ref())?;
        if let Some(shared) = shared_catalog.as_ref()
            && package.read_limits() != *shared.limits
        {
            return Err(invalid(
                "reopened workbook changed its source-backed policy",
            ));
        }
        let retain_catalog =
            facade_execution.is_some() || package.cache_diagnostics().budget_managed;
        let workbook = package.main_document_part()?;
        let flavor = codec::flavor(workbook.content_type()).ok_or_else(|| {
            invalid(format!(
                "main part '{}' has non-XLSX content type '{}'",
                workbook.partname(),
                workbook.content_type()
            ))
        })?;
        // The packaging repairs come first (they were made before the
        // workbook part could be read), then the open-time catalog sites
        // (findings 4.7/4.8); see `crate::leniency`.
        let mut tolerance = ToleranceReport::default();
        tolerance.append_package(package.tolerance_report());
        let catalog_bytes = workbook.data()?;
        let catalog = raw::parse_catalog_repairing(
            catalog_bytes.as_bytes(),
            &package.read_limits(),
            leniency,
            &mut tolerance,
        )?;
        package.check_execution()?;
        check_execution(facade_execution.as_ref())?;
        let sheet_parts = validate_sheet_graph(&package, &workbook, &catalog.sheets)?;
        package.check_execution()?;
        check_execution(facade_execution.as_ref())?;
        let shared_strings_uri = validate_shared_strings(&package, &workbook)?;
        let styles_uri = validate_styles(&package, &workbook)?;
        package.check_execution()?;
        check_execution(facade_execution.as_ref())?;

        let active_sheet = (!catalog.sheets.is_empty()).then_some(catalog.active_sheet_index);
        let defined_names = match shared_catalog.as_ref() {
            Some(shared) => {
                if catalog.defined_names.as_slice() != shared.defined_names.as_ref() {
                    return Err(invalid(
                        "reopened workbook changed its defined-name catalog",
                    ));
                }
                None
            },
            None => Some(catalog.defined_names.into_boxed_slice()),
        };
        let mut sheets = Vec::new();
        sheets
            .try_reserve_exact(catalog.sheets.len())
            .map_err(|source| allocation("source-backed workbook sheets", source))?;
        for (position, (sheet, part)) in catalog.sheets.into_iter().zip(sheet_parts).enumerate() {
            let (keys, name_len, name_key_len) =
                SourceSheetData::keys(&sheet.name, &sheet.relationship_id)?;
            sheets.push(Arc::new(SourceSheetData {
                position,
                keys,
                name_len,
                name_key_len,
                kind: part.kind,
                visibility: codec::visibility(sheet.visibility),
                part_uri: part.uri,
                sheet_id: sheet.sheet_id,
                tolerance: OnceLock::new(),
                build: Mutex::new(()),
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
                .name_key()
                .cmp(sheets[*right].name_key())
                .then_with(|| left.cmp(right))
        });
        let sheet_name_order = sheet_name_order.into_boxed_slice();
        // The source can change after the mandatory root and relationship
        // reads, including while the semantic sheet metadata above is being
        // allocated. Do not publish a facade whose catalog was assembled from
        // a stale positional snapshot. This also covers an empty sheet list.
        package.source_version()?;
        package.check_execution()?;
        check_execution(facade_execution.as_ref())?;

        let source_limits = OnceLock::new();
        let catalog_proof = OnceLock::new();
        if let Some(shared) = shared_catalog {
            let _ = source_limits.set(Arc::clone(&shared.limits));
            let _ = catalog_proof.set(shared);
        }
        let owner_options = shared_owner_options.or_else(|| {
            let owner_options = SourceOwnerOptions::from_source_backed_package(&package, &options);
            (!owner_options.is_default()).then(|| Arc::new(owner_options))
        });

        Ok(Self {
            inner: Arc::new(SourceInner {
                package,
                restore_endpoint: OnceLock::new(),
                tab_state_snapshot: OnceLock::new(),
                tab_state_snapshot_seen: AtomicBool::new(false),
                tab_state_graph: OnceLock::new(),
                package_handle: OnceLock::new(),
                merge_source_states: OnceLock::new(),
                execution: owner_options
                    .as_ref()
                    .and_then(|owner_options| owner_options.facade_execution_context.clone()),
                _catalog_data: retain_catalog.then_some(catalog_bytes),
                shared_strings_uri,
                styles_uri,
                style_lineage: Arc::new(crate::style::StyleLineage),
                values: values::ValueCache::new(),
                styles: OnceLock::new(),
                formats: crate::style::cell_format::FormatCache::new(),
                dependencies: OnceLock::new(),
                flavor,
                date_system: if catalog.uses_1904_date_system {
                    DateSystem::Excel1904
                } else {
                    DateSystem::Excel1900
                },
                active_sheet,
                sheets,
                sheet_name_order,
                tolerance,
                defined_names,
                source_limits,
                catalog_proof,
                owner_options,
            }),
            pending: None,
        })
    }

    /// Every repair made while reading this workbook so far, in sheet order.
    ///
    /// The report merges the repairs made when the workbook was opened with
    /// those of every worksheet whose cells have been materialized, each
    /// stamped with its sheet position. Worksheets parse lazily: a worksheet
    /// joins the report once a read has parsed its grid, and stays in it even
    /// after the semantic cache releases the grid. A read of a part too large
    /// to retain goes through the streaming scan, which never repairs, and
    /// falls back to the materializing parse, which records, when the
    /// worksheet needs a repair. An empty report means everything parsed so
    /// far conformed. See [`crate::leniency`] for the contract.
    #[must_use]
    pub fn tolerance_report(&self) -> ToleranceReport {
        let inner = self.current();
        let mut report = inner.tolerance.clone();
        if let Some(dependencies) = inner.dependencies.get()
            && let Some(strings) = dependencies.shared_strings_tolerance.get()
        {
            report.append(strings);
        }
        for sheet in inner.sheets.iter() {
            if let Some(tolerance) = sheet.tolerance.get() {
                report.append_sheet(tolerance, sheet.position);
            }
        }
        report
    }

    /// Low-level inert defined-name records retained from the workbook
    /// catalog parsed at open.
    ///
    /// These are the records [`crate::Workbook::defined_names`] returns for
    /// the same workbook part; reading them performs no I/O.
    #[must_use]
    pub fn defined_names(&self) -> &[raw::DefinedName] {
        match self.pending_catalog() {
            Some(catalog) => &catalog.defined_names,
            None => self.current().defined_names_slice(),
        }
    }

    pub(crate) fn shared_catalog_proof(&self) -> Arc<SourceCatalogProof> {
        self.current().catalog_proof()
    }

    pub(crate) fn tab_state_snapshot(&self) -> &OnceLock<Arc<crate::tab_state::Snapshot>> {
        &self.current().tab_state_snapshot
    }

    pub(crate) fn tab_state_source_owner(&self) -> Arc<SourceInner> {
        Arc::clone(self.current())
    }

    pub(crate) fn tab_state_snapshot_should_cache(&self) -> bool {
        self.current()
            .tab_state_snapshot_seen
            .swap(true, Ordering::AcqRel)
    }

    /// Content-free diagnostics of the workbook's semantic cache.
    ///
    /// The cache retains parsed worksheet grids and the projections of their
    /// blocks outside `sheetData` under a bounded logical weight, releasing
    /// the least-recently-used values no worksheet handle pins. It is
    /// separate from the deferred-part payload cache reported by
    /// [`Self::cache_diagnostics`].
    #[must_use]
    pub fn semantic_cache_diagnostics(&self) -> CacheDiagnostics {
        self.current().values.diagnostics()
    }

    /// Release every parsed worksheet value the semantic cache retains that
    /// no worksheet handle pins, returning the logical weight released.
    ///
    /// Reads are unaffected apart from cost: a released value is parsed
    /// again from the same source snapshot when a read next needs it.
    pub fn trim_semantic_cache(&self) -> usize {
        self.current().values.trim_clean()
    }

    /// Workbook flavor derived from the mandatory workbook catalog.
    #[must_use]
    pub fn flavor(&self) -> Flavor {
        self.current().flavor
    }

    /// Date serial system derived from the mandatory workbook catalog.
    #[must_use]
    pub fn date_system(&self) -> DateSystem {
        self.current().date_system
    }

    /// Content-free deferred-Part cache diagnostics.
    #[must_use]
    pub fn cache_diagnostics(&self) -> SourceCacheDiagnostics {
        self.current().package.cache_diagnostics()
    }

    /// Exact source identity and revision captured by this facade.
    pub fn source_version(&self) -> Result<SourceVersion> {
        self.current().package.source_version().map_err(Into::into)
    }

    /// Borrow the validated deferred OPC package this catalog reads from.
    ///
    /// This is the package indexed when the workbook was opened: no ZIP
    /// directory, content-type catalog, or relationship graph is rebuilt.
    /// Borrowed package services, such as
    /// [`SourceBackedPackage::to_opc_package`], read through the same
    /// positional source and observe the same source-version fence.
    ///
    /// A committed workbook (record 0360) has no indexed package of its own
    /// until a read writes its output into memory; this does so once.
    ///
    /// # Errors
    ///
    /// Returns the publication error of a committed workbook's output.
    pub fn source_backed_package(&self) -> Result<&SourceBackedPackage> {
        Ok(&self.view()?.inner.package)
    }

    /// Conservative retained-data census for promotion to the owning
    /// workbook used by the edit kernel.
    ///
    /// Promotion keeps the source archive and eagerly materializes the
    /// workbook catalog plus every worksheet part used by the page-break
    /// projection.  The census accounts for those declared payloads and the
    /// native endpoint/catalog owners without decoding any deferred part.
    /// Callers can therefore reject a native edit before promotion allocates;
    /// the owning package's exact retained census remains authoritative after
    /// promotion and commit.
    pub fn retained_bytes_for_edit(&self) -> Result<usize> {
        let inner = self.current();
        let package = &inner.package;
        let mut endpoint = usize::try_from(package.source_artifact().len()).unwrap_or(usize::MAX);
        endpoint = endpoint
            .saturating_add(size_of::<super::Workbook>())
            .saturating_add(size_of::<super::Inner>())
            .saturating_add(size_of::<litchi_opc::OpcPackage>());

        let main = package.main_document_part()?;
        let main_payload =
            usize::try_from(main.declared_uncompressed_size()?).unwrap_or(usize::MAX);
        endpoint = endpoint
            // The owning catalog parses the workbook XML into independent
            // names, relationship targets and sheet records in addition to
            // retaining its source payload.  Reserve one more payload-sized
            // envelope for those derived strings and records.
            .saturating_add(main_payload)
            .saturating_add(main_payload);
        endpoint = endpoint
            .saturating_add(main.partname().as_str().len())
            .saturating_add(main.content_type().len());

        // Workbook::from_package builds a page-break projection for every
        // catalog sheet, which materializes each corresponding part.  Count
        // only those payloads here; unrelated media and extension parts stay
        // deferred and are charged by their family editor if touched.
        for sheet in inner.sheets.iter() {
            let part = package.part(&sheet.part_uri)?;
            endpoint = endpoint.saturating_add(
                usize::try_from(part.declared_uncompressed_size()?).unwrap_or(usize::MAX),
            );
            endpoint = endpoint
                .saturating_add(size_of::<super::model::SheetData>())
                .saturating_add(sheet.name().len())
                .saturating_add(sheet.name_key().len())
                .saturating_add(sheet.part_uri.as_str().len())
                .saturating_add(sheet.relationship_id().len());
        }

        endpoint = endpoint
            .saturating_add(
                inner
                    .shared_strings_uri
                    .as_ref()
                    .map_or(0, |uri| uri.as_str().len()),
            )
            .saturating_add(
                inner
                    .styles_uri
                    .as_ref()
                    .map_or(0, |uri| uri.as_str().len()),
            )
            .saturating_add(
                inner
                    .defined_names
                    .as_ref()
                    .map_or(0, |names| names.len())
                    .saturating_mul(size_of::<raw::DefinedName>()),
            );

        // The native retained census includes package metadata and
        // relationship strings even when their payloads remain deferred.
        for part in package.iter_parts() {
            endpoint = endpoint
                .saturating_add(size_of::<litchi_opc::PartMetadata<'_>>())
                .saturating_add(part.partname().as_str().len())
                .saturating_add(part.content_type().len())
                .saturating_add(relationship_bytes(part.rels()));
        }
        endpoint = endpoint.saturating_add(relationship_bytes(package.rels()));
        Ok(endpoint)
    }

    /// Recover the validated deferred OPC package without re-indexing it.
    ///
    /// The returned package is the one this catalog was opened from, so
    /// [`Self::from_source_backed_package`] can adopt it again without a
    /// second container index. The lazily parsed
    /// catalog, shared strings and styles, and every parsed worksheet value
    /// the semantic cache retains, are released.
    ///
    /// # Errors
    ///
    /// Returns `self` unchanged while another clone of this workbook, or a
    /// [`SourceWorksheet`] handle selected from it, still shares the catalog.
    /// A committed workbook (record 0360), which has no indexed package of
    /// its own, is returned unchanged too.
    pub fn try_into_source_backed_package(self) -> std::result::Result<SourceBackedPackage, Self> {
        if self.pending.is_some() {
            return Err(self);
        }
        match Arc::try_unwrap(self.inner) {
            Ok(inner) => Ok(inner.package),
            Err(inner) => Err(Self {
                inner,
                pending: None,
            }),
        }
    }

    /// Number of logical workbook sheets.
    #[must_use]
    pub fn len(&self) -> usize {
        self.current().sheets.len()
    }

    /// Whether the workbook catalog contains no sheets.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.current().sheets.is_empty()
    }

    /// Iterate lightweight sheet handles without reading worksheet payloads.
    ///
    /// # Errors
    ///
    /// Returns the publication error of a committed workbook's output, which
    /// the first read writes into memory once (record 0360).
    pub fn sheets(
        &self,
    ) -> Result<impl ExactSizeIterator<Item = SourceWorksheet> + DoubleEndedIterator + '_> {
        let inner = &self.view()?.inner;
        Ok(inner
            .sheets
            .iter()
            .cloned()
            .map(|data| SourceWorksheet::new(Arc::clone(inner), data)))
    }

    /// The developer-facing sheet names in catalog order, without reading
    /// any worksheet payload.
    ///
    /// A worksheet-level commit leaves the catalog unchanged, so a committed
    /// workbook answers from its base catalog without materializing its
    /// output. Workbook-level families retain catalog-only state in their
    /// deferred projection; a read that needs the reopened owner materializes
    /// and validates that output before replacing the base projection.
    #[must_use]
    pub fn sheet_names(&self) -> impl ExactSizeIterator<Item = &str> + DoubleEndedIterator + '_ {
        self.current().sheets.iter().map(|sheet| sheet.name())
    }

    /// The catalog position of the active sheet, if the catalog contains one;
    /// see [`Self::sheet_names`] for committed workbooks.
    #[must_use]
    pub fn active_sheet_position(&self) -> Option<usize> {
        match self.pending_catalog() {
            Some(catalog) => catalog.active_sheet,
            None => self.current().active_sheet,
        }
    }

    fn pending_catalog(&self) -> Option<&deferred::Catalog> {
        self.pending
            .as_ref()
            .filter(|pending| pending.materialized().is_none())
            .and_then(|pending| pending.catalog())
    }

    /// Look up a sheet by developer-facing name or checked zero-based position.
    pub fn sheet<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Option<SourceWorksheet>> {
        self.view()?.sheet_in_view(selector)
    }

    fn sheet_in_view<'a>(
        &self,
        selector: impl Into<Selector<'a>>,
    ) -> Result<Option<SourceWorksheet>> {
        self.inner.execution_check()?;
        let data = match selector.into() {
            CoreSelector::Position(position) => self.inner.sheets.get(position.get()).cloned(),
            CoreSelector::Name(name) => {
                let key = crate::sheet::key(&name);
                self.inner
                    .sheet_name_order
                    .binary_search_by(|&position| {
                        self.inner.sheets[position].name_key().cmp(key.as_ref())
                    })
                    .ok()
                    .and_then(|order| self.inner.sheets.get(self.inner.sheet_name_order[order]))
                    .cloned()
            },
            CoreSelector::Id(never) => match never {},
            _ => return Err(Error::UnsupportedSelector),
        };
        Ok(data.map(|data| SourceWorksheet::new(Arc::clone(&self.inner), data)))
    }

    /// Return the active sheet, if the workbook catalog contains one.
    ///
    /// # Errors
    ///
    /// Returns the publication error of a committed workbook's output, which
    /// the first read writes into memory once (record 0360).
    pub fn active_sheet(&self) -> Result<Option<SourceWorksheet>> {
        let inner = &self.view()?.inner;
        let Some(data) = inner
            .active_sheet
            .and_then(|position| inner.sheets.get(position))
            .cloned()
        else {
            return Ok(None);
        };
        Ok(Some(SourceWorksheet::new(Arc::clone(inner), data)))
    }

    /// The catalog state infallible accessors read: the reopened output of a
    /// committed workbook once a read has materialized it, otherwise the
    /// opened (or base) catalog, which a worksheet-level commit leaves
    /// unchanged. Catalog-only workbook-level projections remain deferred;
    /// accessors that need their reopened owner go through [`Self::view`].
    fn current(&self) -> &Arc<SourceInner> {
        match &self.pending {
            Some(pending) => pending
                .materialized()
                .map_or(&self.inner, |view| &view.inner),
            None => &self.inner,
        }
    }
}

fn relationship_bytes(relationships: &litchi_opc::Relationships) -> usize {
    relationships.iter().fold(
        size_of::<litchi_opc::Relationships>(),
        |bytes, relationship| {
            bytes
                .saturating_add(size_of::<litchi_opc::Relationship>())
                .saturating_add(relationship.r_id().len())
                .saturating_add(relationship.reltype().len())
                .saturating_add(relationship.target_ref().len())
                .saturating_add(relationship.base_uri().len())
                .saturating_add(relationship.source_uri().map_or(0, str::len))
        },
    )
}

#[cfg(any(unix, windows))]
fn file_source(path: impl AsRef<Path>) -> Result<Arc<dyn ReadAt>> {
    Ok(Arc::new(FileSource::open(path).map_err(|error| {
        Error::Package(litchi_opc::OpcError::from(error))
    })?))
}

impl SourceWorksheet {
    const fn new(owner: Arc<SourceInner>, data: Arc<SourceSheetData>) -> Self {
        Self {
            owner,
            data,
            grid: OnceLock::new(),
        }
    }

    /// Developer-facing sheet name.
    #[must_use]
    pub fn name(&self) -> &str {
        self.data.name()
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

    /// The number of ordinary charts on this worksheet, in drawing order
    /// (see [`crate::chart::edit`]). Only the worksheet's relationships and
    /// its drawing part are read.
    pub fn chart_count(&self) -> Result<usize> {
        if self.data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: self.data.name().to_owned(),
            });
        }
        crate::chart::edit::source_chart_count(&self.owner.package, &self.data.part_uri)
    }

    /// The chart-space XML of the ordinary chart at `index` (drawing order)
    /// on this worksheet, or `None` past the last chart.
    pub fn chart_xml(&self, index: usize) -> Result<Option<Vec<u8>>> {
        if self.data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: self.data.name().to_owned(),
            });
        }
        crate::chart::edit::source_chart_xml(&self.owner.package, &self.data.part_uri, index)
    }

    /// The two-cell anchor of the chart at `index` (drawing order) on this
    /// worksheet; `None` past the last chart or for a one-cell or absolute
    /// anchor. Only the worksheet's relationships and its drawing are read.
    pub fn chart_anchor(&self, index: usize) -> Result<Option<crate::builder::chart::Anchor>> {
        if self.data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: self.data.name().to_owned(),
            });
        }
        crate::chart::edit::source_chart_anchor(&self.owner.package, &self.data.part_uri, index)
    }

    /// The top-level objects of this worksheet's drawing, in drawing order
    /// (see [`crate::drawing::edit`]). Only the worksheet's relationships, its
    /// drawing part and the image entries are read.
    pub fn drawing_elements(&self) -> Result<Vec<crate::drawing::edit::Element>> {
        if self.data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: self.data.name().to_owned(),
            });
        }
        crate::drawing::edit::source_elements(&self.owner.package, &self.data.part_uri)
    }

    /// Read the worksheet's admitted SpreadsheetML form controls.
    pub fn form_controls(&self) -> Result<crate::form_control::FormControlCollection> {
        if self.data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: self.data.name().to_owned(),
            });
        }
        let package = &self.owner.package;
        let limits = self
            .owner
            .owner_options_ref()
            .semantic_limits()
            .form_controls_value()
            .unwrap_or_else(|| {
                crate::form_control::OwnerLimits::from_read_limits(package.read_limits())
            });
        crate::form_control::source_form_controls_for_sheet_with_limits(
            package,
            &self.data.part_uri,
            limits,
        )
    }

    /// Resolve one form control by checked position or exact authored name.
    pub fn form_control<'a>(
        &self,
        selector: impl Into<crate::form_control::ControlSelector<'a>>,
    ) -> Result<Option<crate::form_control::FormControlView>> {
        Ok(self.form_controls()?.get(selector)?.cloned())
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

    /// The repairs made while parsing this worksheet's cells.
    ///
    /// This materializes the worksheet if it has not been yet. Locations are
    /// relative to this worksheet, so [`crate::ToleratedDefect::sheet`] is
    /// `None`. A chart, dialog, or other non-cell sheet reports clean.
    pub fn tolerance_report(&self) -> Result<&ToleranceReport> {
        if self.data.kind != WorksheetKind::Worksheet {
            return Ok(&crate::leniency::CLEAN);
        }
        let store = self.store();
        self.finish_result(store).map(Store::tolerance_report)
    }

    /// Look up an exact logical cell state, returning an owning semantic value.
    ///
    /// A retained parsed grid answers by lookup. Otherwise the bounded
    /// selected scan retains only requested records, falling back to a parsed
    /// grid when the worksheet needs the full parser. [`Self::grid`] explicitly
    /// materializes and pins a grid for repeated or structural reads.
    pub fn cell<'a>(&self, at: impl Into<At<'a>>) -> Result<SourceCellView> {
        let address = at.into().resolve()?;
        let value = match self.streams_cells() {
            Ok(true) => self.stream_cell(address),
            Ok(false) => self.eager_cell(address),
            Err(error) => Err(error),
        };
        self.finish_result(value)
    }

    /// The rich-text runs of the cell at `at`, as
    /// [`Worksheet::rich_text`](crate::Worksheet::rich_text) reads them
    /// (record 0338): an inline rich string, or a shared-string item with
    /// runs; `None` for any other cell. The worksheet's grid is parsed, and
    /// the worksheet or shared-string part is scanned once.
    ///
    /// # Errors
    ///
    /// The errors of reading the worksheet's cells or the part holding the
    /// runs, and a changed source.
    pub fn rich_text<'a>(&self, at: impl Into<At<'a>>) -> Result<Option<crate::RichText>> {
        let address = at.into().resolve()?;
        let result = (|| {
            let (inline, shared) = match self.store()?.entry(address) {
                Some(entry) => (entry.inline_rich, entry.shared_string),
                None => return Ok(None),
            };
            if inline {
                let data = self.owner.package.part(&self.data.part_uri)?.data()?;
                return crate::workbook::model::inline_rich(data.as_bytes(), &address.a1());
            }
            let (Some(index), Some(uri)) = (shared, self.owner.shared_strings_uri.as_ref()) else {
                return Ok(None);
            };
            let data = self.owner.package.part(uri)?.data()?;
            crate::workbook::model::rich_item(data.as_bytes(), index)
        })();
        self.finish_result(result)
    }

    /// Whether a cell read of this worksheet streams its part instead of
    /// reading the parsed grid.
    ///
    /// A resident grid (pinned by this handle or retained by the workbook)
    /// always answers. A cold eligible worksheet keeps the selected streaming
    /// route; materialization is explicit or a parser fallback.
    fn streams_cells(&self) -> Result<bool> {
        if self.data.kind != WorksheetKind::Worksheet || self.grid.get().is_some() {
            return Ok(false);
        }
        // A retained grid is pinned here, in the one lookup this read makes;
        // the store read that follows runs the resident fences. A miss is
        // counted by that store read, not by this probe.
        if let Some(store) = self.cached_grid(false) {
            let _publish_result = self.grid.set(Grid::new(store));
            return Ok(false);
        }
        self.owner.execution_check()?;
        self.owner.package.source_version()?;
        Ok(true)
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

                if let Some(range) = selected.covering_merge
                    && range.start() != address
                {
                    return Ok(SourceCellView::Covered(range));
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
    /// The range takes the route [`Self::cell`] describes: the worksheet's
    /// resident parsed grid, or the bounded sparse selection route when a
    /// cold worksheet's structure is eligible. The returned cells are independent of
    /// the source and may outlive this worksheet handle.
    pub fn cells<'a>(&self, area: impl Into<Area<'a>>) -> Result<Vec<SourceCell>> {
        let range = area.into().resolve()?;
        let values = self
            .select_cells(range)
            .and_then(|selection| match selection {
                Selection::Stored(store) => collect_stored_cells(store, range),
                Selection::Selected { cells, shared_text } => {
                    let mut values = Vec::new();
                    values
                        .try_reserve_exact(cells.len())
                        .map_err(|source| allocation("source-backed selected cells", source))?;
                    for record in cells {
                        values.push(resolve_selected_record(record, shared_text.as_deref())?);
                    }
                    Ok(values)
                },
            });
        self.finish_result(values)
    }

    /// Visit every stored cell selected by a checked range.
    ///
    /// The range takes the same route, the same fences and the same verified
    /// values as [`Self::cells`], but each selected cell is produced as it is
    /// visited instead of after a whole-range vector of owning cells has been
    /// built. Callbacks still receive references into local values, so no
    /// callback runs while a source reader is active: the parsed grid is
    /// complete, or the bounded selected scan has reached worksheet EOF,
    /// resolved its shared-string and style dependencies and validated every
    /// retained record, before the first callback, so every refusal
    /// [`Self::cells`] raises for this range is still raised before any cell
    /// is visited. [`Self::visit_cells_until`] is the variant whose callback
    /// can stop the walk without an error.
    ///
    /// Cancellation is checked before each callback. Final source and
    /// execution fences run even when a callback fails, so a source mutation
    /// or cancellation remains primary over that callback error.
    pub fn visit_cells<'a, F>(&self, area: impl Into<Area<'a>>, mut visit: F) -> Result<usize>
    where
        F: FnMut(Address, &Cell) -> Result<()>,
    {
        let flow = self.visit_cells_until(area, |address, cell| match visit(address, cell) {
            Ok(()) => ControlFlow::Continue(()),
            Err(error) => ControlFlow::Break(error),
        })?;
        match flow {
            ControlFlow::Continue(visited) => Ok(visited),
            ControlFlow::Break(error) => Err(error),
        }
    }

    /// Choose the route one checked range read takes and finish every step
    /// that reads the source.
    ///
    /// Nothing in the returned selection holds a source reader, so a caller
    /// may convert and publish its records one at a time. Every refusal that a
    /// whole-range read of this rectangle can raise has already been raised
    /// when this returns.
    fn select_cells(&self, range: Rect) -> Result<Selection<'_>> {
        if self.streams_cells()? {
            return self.stream_selection(range);
        }
        Ok(Selection::Stored(self.store()?))
    }

    #[expect(
        clippy::result_large_err,
        reason = "The selected stream error intentionally retains typed primary plus raw/active callback diagnostics; boxing it would change the established API."
    )]
    fn stream_selection(&self, range: Rect) -> Result<Selection<'_>> {
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
            return Ok(Selection::Stored(self.store()?));
        };
        let selected = match outcome {
            raw::selected_worksheet::RangeScanOutcome::Eligible(selected) => selected,
            raw::selected_worksheet::RangeScanOutcome::NotEligible(_) => {
                // The verified worksheet reader must be gone before the
                // materialized fallback can inspect the source again.
                self.owner.package.source_version()?;
                self.owner.execution_check()?;
                return Ok(Selection::Stored(self.store()?));
            },
        };

        let dependencies = selected.dependencies;
        // Only shared-string records need this staging vector. Reserving for
        // every selected physical cell made numeric/inline ranges allocate a
        // temporary index table whose length was always zero after the scan.
        let requested_count = selected
            .cells
            .iter()
            .filter(|record| matches!(&record.payload, SelectedPayload::SharedString(_)))
            .count();
        let mut requested = Vec::new();
        requested
            .try_reserve_exact(requested_count)
            .map_err(|source| allocation("source-backed selected shared-string indexes", source))?;
        let mut fallback = false;
        for record in &selected.cells {
            let SelectedPayload::SharedString(index) = &record.payload else {
                continue;
            };
            let Ok(index) = usize::try_from(*index) else {
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
                let SelectedPayload::SharedString(index) = &record.payload else {
                    continue;
                };
                let Ok(index) = usize::try_from(*index) else {
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
            return Ok(Selection::Stored(self.store()?));
        }
        Ok(Selection::Selected {
            cells: selected.cells,
            shared_text,
        })
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

    /// The parsed grid of this worksheet, pinned by this handle.
    ///
    /// A grid this handle already pins is returned after the source and
    /// execution fences. Otherwise the workbook's semantic cache is asked,
    /// and on a miss the part is parsed once, under this worksheet's build
    /// lock so racing readers share one build, and the result is offered to
    /// the cache. Only a complete, successfully parsed grid is published.
    fn store(&self) -> Result<&Store> {
        self.owner.execution_check()?;
        if self.data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: self.data.name().to_owned(),
            });
        }
        // `SourceBackedPackage::part` checks the captured source version even
        // when this worksheet's parsed semantic store is already resident.
        // Never return cached semantic data after a caller-visible source
        // change has invalidated the snapshot.
        let part = self.owner.package.part(&self.data.part_uri)?;
        if let Some(grid) = self.grid.get() {
            return self.resident_store(grid);
        }
        if let Some(store) = self.cached_grid(true) {
            return self.pin(store);
        }
        let _build = self.data.lock_build();
        if let Some(store) = self.cached_grid(false) {
            return self.pin(store);
        }

        let data = part.data()?;
        let mut parsed = raw::worksheet::parse(data.as_bytes(), || self.owner.shared_strings())?;
        drop(data);
        self.owner.repair_styles(&mut parsed)?;
        self.owner.execution_check()?;
        self.owner.package.source_version()?;
        let _publish_result = self.data.tolerance.set(parsed.tolerance_report().clone());
        let weight = parsed.weight_bytes();
        let store = match self.owner.values().publish(
            self.data.position,
            values::Cached::Grid(Arc::new(parsed)),
            weight,
        ) {
            values::Cached::Grid(store) => store,
            values::Cached::Blocks(_) | values::Cached::Prefix(_) => {
                return Err(invalid(
                    "source-backed worksheet grid cache returned another value kind",
                ));
            },
        };
        let _publish_result = self.grid.set(Grid::new(store));
        self.grid.get().map(Grid::store).ok_or_else(|| {
            invalid("source-backed worksheet cache initialization did not publish a value")
        })
    }

    /// Return a resident store after the fences every cached read runs.
    fn resident_store<'a>(&self, grid: &'a Grid) -> Result<&'a Store> {
        // A cached semantic Store must still honor the managed package's
        // cancellation and source-version checks. `PartView::data()` is
        // deliberately not requested here: the semantic Store is already
        // the validated worksheet snapshot, and its retained value must
        // remain usable even when the bounded payload cache evicted or
        // bypassed the original worksheet bytes.
        self.owner.package.check_execution()?;
        self.owner.execution_check()?;
        self.owner.package.source_version()?;
        Ok(grid.store())
    }

    /// Pin a store the workbook's cache retains and return it.
    fn pin(&self, store: Arc<Store>) -> Result<&Store> {
        // A concurrent reader of this same handle may have pinned the same
        // cached store first; either value is the one the cache holds.
        let _publish_result = self.grid.set(Grid::new(store));
        let grid = self
            .grid
            .get()
            .ok_or_else(|| invalid("source-backed worksheet grid was not pinned"))?;
        self.resident_store(grid)
    }

    fn cached_grid(&self, record_miss: bool) -> Option<Arc<Store>> {
        match self
            .owner
            .values()
            .get(self.data.position, values::Axis::Grid, record_miss)?
        {
            values::Cached::Grid(store) => Some(store),
            values::Cached::Blocks(_) | values::Cached::Prefix(_) => None,
        }
    }

    /// Whether this handle pins, or the workbook's cache retains, the parsed
    /// grid of this worksheet.
    #[cfg(test)]
    fn store_resident(&self) -> bool {
        self.grid.get().is_some()
            || self
                .owner
                .values()
                .holds(self.data.position, values::Axis::Grid)
    }
}

/// Route one checked source-backed range read takes once every source read
/// it needs is complete.
///
/// Neither variant holds a source reader, so a caller may publish its cells
/// one at a time without a callback observing the source mid-read.
enum Selection<'a> {
    /// The materialized worksheet store is authoritative for this read.
    Stored(&'a Store),
    /// Physical records retained by the bounded selected scan, in source
    /// order, with the shared-string payloads they depend on resolved and
    /// every record validated.
    Selected {
        cells: Vec<raw::selected_worksheet::SelectedRecord>,
        shared_text: Option<Vec<(usize, Text)>>,
    },
}

/// Copy every stored cell of a checked range out of a materialized store.
fn collect_stored_cells(store: &Store, range: Rect) -> Result<Vec<SourceCell>> {
    // Count the sparse records that will actually be copied before reserving.
    // The old one-slot growth loop doubled a 75,770-cell whole-sheet result to
    // 131,072 slots.  A count pass keeps narrow ranges sparse while making the
    // owning result's capacity exact; the store is immutable for this read, so
    // the second traversal cannot observe a different set of records.
    let count = if range == Rect::ALL {
        store.stored_cell_count()
    } else {
        store.cells(range).count()
    };
    let mut values = Vec::new();
    values
        .try_reserve_exact(count)
        .map_err(|source| allocation("source-backed selected cells", source))?;
    for (address, cell) in store.cells(range) {
        values.push(SourceCell {
            address,
            cell: cell.clone(),
        });
    }
    Ok(values)
}

/// Bind one retained selected record to its resolved shared-string payload.
///
/// `SelectedRecord::payload` makes the semantic/deferred choice an invariant,
/// so this conversion has no late both/neither refusal branch. Dependency
/// presence and source fences are still checked before the selection is
/// published by `stream_selection`.
///
/// Always inlined: the record and the produced cell are both larger than a
/// register pair, so an out-of-line call copies each of them through the stack
/// once per cell. Change 0642 measured the out-of-line form at 688,015 more
/// instructions over a 65,536-cell range — about 10.5 per cell — on the pinned
/// toolchain.
#[inline(always)]
fn resolve_selected_record(
    record: raw::selected_worksheet::SelectedRecord,
    shared_text: Option<&[(usize, Text)]>,
) -> Result<SourceCell> {
    let cell = match record.payload {
        SelectedPayload::Cell(cell) => cell,
        SelectedPayload::SharedString(index) => {
            let index = usize::try_from(index)
                .map_err(|_error| invalid("shared-string index exceeds this platform"))?;
            let text = shared_text
                .and_then(|text| {
                    text.binary_search_by_key(&index, |(candidate, _)| *candidate)
                        .ok()
                        .and_then(|position| text.get(position))
                })
                .map(|(_, text)| text.clone())
                .ok_or_else(|| invalid("selected shared-string dependency was not retained"))?;
            Cell::Value(Value::Text(text))
        },
    };
    Ok(SourceCell {
        address: record.address,
        cell,
    })
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
    fn default_owner_options() -> &'static Arc<SourceOwnerOptions> {
        static DEFAULT: OnceLock<Arc<SourceOwnerOptions>> = OnceLock::new();
        DEFAULT.get_or_init(|| {
            Arc::new(SourceOwnerOptions::from_open_options(
                &crate::OpenOptions::default(),
            ))
        })
    }

    fn owner_options_ref(&self) -> &SourceOwnerOptions {
        self.owner_options
            .as_deref()
            .unwrap_or_else(|| Self::default_owner_options().as_ref())
    }

    fn owner_options(&self) -> Arc<SourceOwnerOptions> {
        self.owner_options
            .clone()
            .unwrap_or_else(|| Arc::clone(Self::default_owner_options()))
    }

    fn defined_names_slice(&self) -> &[raw::DefinedName] {
        match &self.defined_names {
            Some(names) => names,
            // A proof-backed owner is constructed with its catalog proof set.
            None => self
                .catalog_proof
                .get()
                .expect("proof-backed source owner must retain its catalog proof")
                .defined_names
                .as_ref(),
        }
    }

    pub(crate) fn tab_state_graph(
        &self,
        owner: std::sync::Weak<SourceInner>,
    ) -> Arc<crate::tab_state::GraphBinding> {
        Arc::clone(self.tab_state_graph.get_or_init(|| {
            Arc::new(crate::tab_state::GraphBinding::source(
                owner,
                Arc::from(self.sheets.clone()),
            ))
        }))
    }

    pub(crate) fn tab_state_package(&self) -> &SourceBackedPackage {
        &self.package
    }

    pub(crate) fn tab_state_catalog_is_clean(&self) -> bool {
        self.tolerance.is_clean()
    }

    pub(crate) fn source_limits(&self) -> Arc<SourceLimits> {
        Arc::clone(
            self.source_limits
                .get_or_init(|| Arc::new(self.package.read_limits())),
        )
    }

    /// The catalog proof, when this owner already holds one (a proof-backed
    /// reopen, or an owner whose proof an earlier edit captured).
    pub(crate) fn catalog_proof_if_ready(&self) -> Option<Arc<SourceCatalogProof>> {
        self.catalog_proof.get().cloned()
    }

    /// Whether this owner's package was admitted under `limits`.
    pub(crate) fn admitted_under(&self, limits: &SourceLimits) -> bool {
        match self.source_limits.get() {
            Some(admitted) => **admitted == *limits,
            None => self.package.read_limits() == *limits,
        }
    }

    pub(crate) fn catalog_proof(&self) -> Arc<SourceCatalogProof> {
        Arc::clone(self.catalog_proof.get_or_init(|| {
            Arc::new(SourceCatalogProof {
                defined_names: Arc::from(self.defined_names_slice()),
                limits: self.source_limits(),
            })
        }))
    }

    fn values(&self) -> &values::ValueCache {
        &self.values
    }

    fn dependencies(&self) -> &DependencyState {
        self.dependencies
            .get_or_init(|| Box::new(DependencyState::new()))
            .as_ref()
    }

    pub(crate) fn package_handle(&self) -> Arc<SourceBackedPackage> {
        Arc::clone(
            self.package_handle
                .get_or_init(|| Arc::new(self.package.clone())),
        )
    }

    pub(crate) fn merge_source_state(
        &self,
        worksheet_uri: &PackURI,
    ) -> Option<Arc<crate::merge::source::SourceStateShared>> {
        self.merge_source_states.get().and_then(|states| {
            states
                .lock()
                .unwrap_or_else(PoisonError::into_inner)
                .get(worksheet_uri)
                .cloned()
        })
    }

    pub(crate) fn cache_merge_source_state(
        &self,
        worksheet_uri: PackURI,
        state: Arc<crate::merge::source::SourceStateShared>,
    ) -> Arc<crate::merge::source::SourceStateShared> {
        let mut states = self
            .merge_source_states
            .get_or_init(|| Box::new(Mutex::new(HashMap::new())))
            .lock()
            .unwrap_or_else(PoisonError::into_inner);
        Arc::clone(states.entry(worksheet_uri).or_insert(state))
    }

    fn execution_check(&self) -> Result<()> {
        // The package owns the execution context used by managed source
        // caches. Always check it first so compatibility facades built around
        // an already-managed package cannot bypass cancellation or budget
        // policy merely because their optional facade context is absent.
        self.package.check_execution()?;
        // Constructors that create both layers retain the same logical
        // context in each handle; checking the optional facade handle as well
        // also covers callers that supply a distinct policy at this boundary.
        check_execution(self.execution.as_deref())
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
            || selected.irregular
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
        let dependencies = self.dependencies();
        if let Some(strings) = dependencies.shared_strings.get() {
            let _part = self.package.part(uri)?;
            self.execution_check()?;
            self.package.source_version()?;
            return Ok(Some(strings));
        }

        let data = self.package.part(uri)?.data()?;
        let mut tolerance = ToleranceReport::default();
        let parsed = raw::strings::parse_repairing(data.as_bytes(), &mut tolerance)?;
        self.execution_check()?;
        self.package.source_version()?;
        // The report is published first, so a published table always has
        // its report; a racing parse of the same bytes reports the same.
        let _publish_result = dependencies.shared_strings_tolerance.set(tolerance);
        let _publish_result = dependencies.shared_strings.set(parsed);
        dependencies
            .shared_strings
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

    /// Repair stale style indexes in a freshly parsed worksheet store before
    /// it is published to the read cache (audit finding 4.2).
    ///
    /// Each repair is recorded in the store's own tolerance report.
    fn repair_styles(&self, store: &mut Store) -> Result<()> {
        if !store.references_styles() {
            return Ok(());
        }
        store.repair_style_indexes(self.style_count()?)
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
    use std::io::{self, Cursor, Read};
    use std::num::{NonZeroU64, NonZeroUsize};
    use std::sync::Arc;
    use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

    use litchi_core::{
        Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
        ReadAt, Resource, SourceVersion,
    };
    use litchi_opc::constants::content_type as ct;
    use litchi_opc::{
        OpcError, OpcPackage, PackURI, SourceBackedPackage, SourceCacheLimits, SourceLimitKind,
        SourceReadPolicy,
    };
    use soapberry_zip::office::StreamingArchiveWriter;

    use super::{SourceBackedWorkbook, SourceCellView};
    use crate::{Cell, Error, Value};
    use litchi_opc::SourceLimits;

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

    struct CountingReader {
        inner: Cursor<Vec<u8>>,
        reads: Arc<AtomicUsize>,
        bytes_read: Arc<AtomicUsize>,
    }

    impl CountingReader {
        fn new(bytes: Vec<u8>, reads: Arc<AtomicUsize>) -> Self {
            Self {
                inner: Cursor::new(bytes),
                reads,
                bytes_read: Arc::new(AtomicUsize::new(0)),
            }
        }

        fn with_byte_counter(
            bytes: Vec<u8>,
            reads: Arc<AtomicUsize>,
            bytes_read: Arc<AtomicUsize>,
        ) -> Self {
            Self {
                inner: Cursor::new(bytes),
                reads,
                bytes_read,
            }
        }
    }

    impl Read for CountingReader {
        fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
            self.reads.fetch_add(1, Ordering::SeqCst);
            let read = self.inner.read(output)?;
            self.bytes_read.fetch_add(read, Ordering::SeqCst);
            Ok(read)
        }
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
        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
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
        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
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
        let package = OpcPackage::from_bytes(bytes, SourceLimits::default()).unwrap();
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
                .unwrap()
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
        let workbook = SourceBackedWorkbook::from_read_at_with_options(
            source.clone(),
            crate::OpenOptions::new()
                .source_limits(SourceLimits::default())
                .cache_limits(cache_limits),
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
    fn sequential_reader_preserves_deferred_cache_options() {
        let cache_limits = SourceCacheLimits::new(256 * 1024, 1).unwrap();
        let workbook = SourceBackedWorkbook::from_reader_with_options(
            Cursor::new(source_backed_xlsx()),
            crate::OpenOptions::new().cache_limits(cache_limits),
        )
        .unwrap();

        let first = workbook.sheet("First").unwrap().unwrap();
        assert!(first.stored_extent().unwrap().is_some());
        assert!(first.cell("A1").is_ok());
        let second = workbook.sheet("Second").unwrap().unwrap();
        assert!(second.stored_extent().unwrap().is_some());
        assert!(second.cell("A1").is_ok());
        assert!(
            workbook.cache_diagnostics().evictions > 0,
            "the sequential-reader path must retain the caller's one-entry cache policy"
        );
    }

    #[test]
    fn sequential_reader_retains_cancellation_after_open() {
        let archive = source_backed_xlsx();
        let reads = Arc::new(AtomicUsize::new(0));
        let bytes_read = Arc::new(AtomicUsize::new(0));
        let reader = CountingReader::with_byte_counter(
            archive.clone(),
            Arc::clone(&reads),
            Arc::clone(&bytes_read),
        );
        let (budget, cancellation_source, context) = managed_context(u64::MAX);
        let workbook = SourceBackedWorkbook::from_reader_with_options(
            reader,
            crate::OpenOptions::new()
                .cache_limits(SourceCacheLimits::new(256 * 1024, 1).unwrap())
                .execution_context(context),
        )
        .unwrap();

        // Catalog access succeeds after the sequential reader has completed,
        // while the worksheet body remains deferred. Cancellation must still
        // be observed by that later read through the retained OPC context.
        assert_eq!(workbook.len(), 2);
        assert_eq!(bytes_read.load(Ordering::SeqCst), archive.len());
        let first = workbook.sheet("First").unwrap().unwrap();
        let reads_before_cancel = reads.load(Ordering::SeqCst);
        cancellation_source.cancel();
        assert!(matches!(
            first.stored_extent(),
            Err(Error::Package(OpcError::Cancelled))
        ));
        assert_eq!(reads.load(Ordering::SeqCst), reads_before_cancel);
        drop(first);
        drop(workbook);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn sequential_reader_enforces_exact_and_over_source_input_limits() {
        let archive = source_backed_xlsx();
        let exact = SourceLimits::builder()
            .max_input_bytes(archive.len() as u64)
            .unwrap()
            .build()
            .unwrap();
        let workbook = SourceBackedWorkbook::from_reader_with_options(
            Cursor::new(archive.clone()),
            crate::OpenOptions::new().source_limits(exact),
        )
        .unwrap();
        assert_eq!(workbook.len(), 2);

        let maximum = archive.len() - 1;
        let under = SourceLimits::builder()
            .max_input_bytes(maximum as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            SourceBackedWorkbook::from_reader_with_options(
                Cursor::new(archive),
                crate::OpenOptions::new().source_limits(under),
            ),
            Err(Error::Package(OpcError::Limit {
                kind: SourceLimitKind::InputBytes,
                observed,
                maximum: configured,
            })) if observed == maximum as u64 + 1 && configured == maximum as u64
        ));
    }

    #[test]
    fn sequential_reader_cache_options_survive_committed_owner_reopen() {
        let cache_limits = SourceCacheLimits::new(256 * 1024, 1).unwrap();
        let workbook = SourceBackedWorkbook::from_reader_with_options(
            Cursor::new(source_backed_xlsx()),
            crate::OpenOptions::new().cache_limits(cache_limits),
        )
        .unwrap();
        let mut edit = workbook.edit_tab_state().unwrap();
        assert!(edit.hide("First").unwrap());
        let commit = edit.commit().unwrap();

        let committed = commit.package();
        let first = committed.sheet("First").unwrap().unwrap();
        assert!(first.stored_extent().unwrap().is_some());
        let second = committed.sheet("Second").unwrap().unwrap();
        assert!(second.stored_extent().unwrap().is_some());
        assert!(
            committed.cache_diagnostics().evictions > 0,
            "a committed sequential owner must retain its custom one-entry cache policy"
        );
    }

    #[test]
    fn sequential_reader_checks_cancellation_before_ingestion() {
        let reads = Arc::new(AtomicUsize::new(0));
        let reader = CountingReader::new(source_backed_xlsx(), Arc::clone(&reads));
        let (budget, cancellation_source, context) = managed_context(u64::MAX);
        cancellation_source.cancel();

        assert!(matches!(
            SourceBackedWorkbook::from_reader_with_options(
                reader,
                crate::OpenOptions::new().execution_context(context),
            ),
            Err(Error::Package(OpcError::Cancelled))
        ));
        assert_eq!(reads.load(Ordering::SeqCst), 0);
        assert_eq!(budget.used(Resource::Memory), 0);
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
        let workbook = SourceBackedWorkbook::from_read_at_with_options(
            source.clone(),
            crate::OpenOptions::new()
                .source_limits(SourceLimits::default())
                .cache_limits(cache_limits)
                .execution_context(context),
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
        let workbook = SourceBackedWorkbook::from_read_at_with_options(
            source.clone(),
            crate::OpenOptions::new()
                .source_limits(SourceLimits::default())
                .cache_limits(cache_limits)
                .execution_context(context),
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
            SourceBackedWorkbook::from_read_at_with_options(
                source.clone(),
                crate::OpenOptions::new()
                    .source_limits(SourceLimits::default())
                    .execution_context(context)
            ),
            Err(Error::Package(OpcError::Cancelled))
        ));
        assert_eq!(source.first_body_marker_reads.load(Ordering::SeqCst), 0);
        assert_eq!(source.second_body_marker_reads.load(Ordering::SeqCst), 0);
        assert_eq!(budget.used(Resource::Memory), 0);

        let source = Arc::new(CountingSource::new(archive));
        let (budget, cancellation_source, context) = managed_context(u64::MAX);
        let workbook = SourceBackedWorkbook::from_read_at_with_options(
            source.clone(),
            crate::OpenOptions::new()
                .source_limits(SourceLimits::default())
                .execution_context(context),
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
        let package = SourceBackedPackage::from_read_at(
            source.clone(),
            SourceLimits::default(),
            cache_limits,
            SourceReadPolicy::exact(),
            Some(context),
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

    #[test]
    fn externally_opened_package_preserves_custom_cache_after_commit() {
        let source = Arc::new(CountingSource::new(source_backed_xlsx()));
        let cache_limits = SourceCacheLimits::new(256 * 1024, 1).unwrap();
        let package = SourceBackedPackage::from_read_at(
            source,
            SourceLimits::default(),
            cache_limits,
            SourceReadPolicy::exact(),
            None,
        )
        .unwrap();
        // A facade option cannot replace the cache policy already admitted by
        // an externally supplied package. Use the OPC default here so a
        // regression would visibly widen the one-entry cache.
        let workbook = SourceBackedWorkbook::from_source_backed_package_with_options(
            package,
            crate::OpenOptions::new().cache_limits(SourceCacheLimits::default()),
        )
        .unwrap();
        let mut edit = workbook.edit_tab_state().unwrap();
        assert!(edit.hide("First").unwrap());
        let commit = edit.commit().unwrap();
        let committed = commit.package();

        assert_eq!(
            committed.source_backed_package().unwrap().cache_limits(),
            cache_limits,
            "materialization must retain the externally admitted cache policy"
        );
        assert!(
            committed
                .sheet("First")
                .unwrap()
                .unwrap()
                .stored_extent()
                .unwrap()
                .is_some()
        );
        assert!(
            committed
                .sheet("Second")
                .unwrap()
                .unwrap()
                .stored_extent()
                .unwrap()
                .is_some()
        );
        assert!(
            committed.cache_diagnostics().evictions > 0,
            "a committed external package must retain its one-entry cache"
        );
    }

    #[test]
    fn externally_opened_managed_package_context_survives_commit() {
        let source = Arc::new(CountingSource::new(source_backed_xlsx()));
        let (budget, cancellation_source, context) = managed_context(u64::MAX);
        let package = SourceBackedPackage::from_read_at(
            source,
            SourceLimits::default(),
            SourceCacheLimits::new(256 * 1024, 4).unwrap(),
            SourceReadPolicy::exact(),
            Some(context),
        )
        .unwrap();
        let workbook = SourceBackedWorkbook::from_source_backed_package(package).unwrap();
        let mut edit = workbook.edit_tab_state().unwrap();
        assert!(edit.hide("First").unwrap());
        let commit = edit.commit().unwrap();
        let committed = commit.package();

        assert!(
            committed
                .source_backed_package()
                .unwrap()
                .execution_context()
                .is_some()
        );
        let first = committed.sheet("First").unwrap().unwrap();
        assert!(first.stored_extent().unwrap().is_some());
        cancellation_source.cancel();
        assert!(matches!(
            first.stored_extent(),
            Err(Error::Package(OpcError::Cancelled))
        ));

        drop(first);
        drop(workbook);
        drop(commit);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn external_package_and_facade_contexts_both_survive_commit() {
        let source = Arc::new(CountingSource::new(source_backed_xlsx()));
        let (package_budget, _package_cancellation, package_context) = managed_context(u64::MAX);
        let package = SourceBackedPackage::from_read_at(
            source,
            SourceLimits::default(),
            SourceCacheLimits::new(256 * 1024, 4).unwrap(),
            SourceReadPolicy::exact(),
            Some(package_context),
        )
        .unwrap();
        let (facade_budget, facade_cancellation, facade_context) = managed_context(u64::MAX);
        let workbook = SourceBackedWorkbook::from_source_backed_package_with_options(
            package,
            crate::OpenOptions::new().execution_context(facade_context),
        )
        .unwrap();
        let mut edit = workbook.edit_tab_state().unwrap();
        assert!(edit.hide("First").unwrap());
        let commit = edit.commit().unwrap();
        let committed = commit.package();
        let first = committed.sheet("First").unwrap().unwrap();
        assert!(first.stored_extent().unwrap().is_some());

        facade_cancellation.cancel();
        assert!(matches!(
            first.stored_extent(),
            Err(Error::Package(OpcError::Cancelled))
        ));
        drop(first);
        drop(workbook);
        drop(commit);
        assert_eq!(package_budget.used(Resource::Memory), 0);
        assert_eq!(facade_budget.used(Resource::Memory), 0);
    }

    #[test]
    fn cancelled_facade_fences_uncached_deferred_view_and_write() {
        let source = Arc::new(CountingSource::new(source_backed_xlsx()));
        let package = SourceBackedPackage::from_read_at(
            source,
            SourceLimits::default(),
            SourceCacheLimits::new(256 * 1024, 4).unwrap(),
            SourceReadPolicy::exact(),
            None,
        )
        .unwrap();
        let (budget, cancellation_source, context) = managed_context(u64::MAX);
        let workbook = SourceBackedWorkbook::from_source_backed_package_with_options(
            package,
            crate::OpenOptions::new().execution_context(context),
        )
        .unwrap();
        let mut edit = workbook.edit_tab_state().unwrap();
        assert!(edit.hide("First").unwrap());
        let commit = edit.commit().unwrap();
        let committed = commit.package();
        assert!(
            committed
                .pending
                .as_ref()
                .is_some_and(|pending| pending.materialized().is_none()),
            "the regression must exercise the uncached deferred publication path"
        );

        cancellation_source.cancel();
        assert!(matches!(
            committed.sheet("First"),
            Err(Error::Package(OpcError::Cancelled))
        ));
        assert!(
            committed
                .pending
                .as_ref()
                .is_some_and(|pending| pending.materialized().is_none()),
            "a cancelled view must not publish or cache an output"
        );

        let mut output = Vec::new();
        assert!(matches!(
            committed.write_to(&mut output),
            Err(Error::Package(OpcError::Cancelled))
        ));
        assert!(
            output.is_empty(),
            "cancelled publication must not touch the sink"
        );

        drop(workbook);
        drop(commit);
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
            CountingSource, SourceBackedWorkbook, SourceCellView, SourceLimits, managed_context,
            source_backed_xlsx,
        };
        use crate::{Cell, Error, Value};

        const SPREADSHEETML_NAMESPACE: &str =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        fn one_sheet_xlsx(worksheet: &str) -> Vec<u8> {
            let mut writer = StreamingArchiveWriter::new(
                soapberry_zip::office::StreamingArchiveLimits::default(),
            );
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

            assert!(!sheet.store_resident());
            assert!(matches!(
                sheet.cell("A1").unwrap(),
                SourceCellView::Stored(Cell::Value(Value::Number(ref value)))
                    if value.as_str() == "7"
            ));
            assert!(!sheet.store_resident());
            assert!(matches!(sheet.cell("C1").unwrap(), SourceCellView::Missing));
            assert!(!sheet.store_resident());
            assert!(matches!(
                sheet.cell("B1").unwrap(),
                SourceCellView::Stored(Cell::Empty)
            ));
            assert!(!sheet.store_resident());

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

            assert!(!sheet.store_resident());
            assert!(matches!(
                sheet.cell("A1").unwrap(),
                SourceCellView::Stored(Cell::Value(Value::Number(ref value)))
                    if value.as_str() == "7"
            ));
            let first_reads = source.first_body_marker_reads.load(Ordering::SeqCst);
            assert!(first_reads > 0);
            assert!(!sheet.store_resident());
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
            assert!(!sheet.store_resident());
            assert_eq!(
                workbook.cache_diagnostics().successful_loads,
                cold.successful_loads
            );
        }

        #[test]
        fn dimension_metadata_0400_keeps_scalar_and_range_queries_streaming() {
            let bytes = one_sheet_xlsx(&format!(
                r#"<worksheet xmlns="{SPREADSHEETML_NAMESPACE}"><dimension ref="$B$2:F9"/><sheetData><row r="2"><c r="B2"><v>-0.000</v></c><c r="F2"><v>6.02E+23</v></c></row></sheetData></worksheet>"#,
            ));
            let workbook = SourceBackedWorkbook::from_reader(Cursor::new(bytes)).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();

            assert!(!sheet.store_resident());
            assert!(matches!(
                sheet.cell("B2").unwrap(),
                SourceCellView::Stored(Cell::Value(Value::Number(ref value)))
                    if value.as_str() == "-0.000"
            ));
            assert!(!sheet.store_resident());

            let selected = sheet.cells("A1:F9").unwrap();
            assert_eq!(selected.len(), 2);
            assert_eq!(selected[0].address.a1(), "B2");
            assert_eq!(selected[1].address.a1(), "F2");
            assert!(!sheet.store_resident());

            let extent = sheet.stored_extent().unwrap().expect("declared extent");
            assert_eq!(extent.a1(), "B2:F2");
            assert!(sheet.store_resident());
        }

        #[test]
        fn unsupported_dimension_metadata_0400_preserves_eager_fallback() {
            let bytes = one_sheet_xlsx(&format!(
                r#"<worksheet xmlns="{SPREADSHEETML_NAMESPACE}"><dimension ref="A1" future="preserve"/><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>"#,
            ));
            let workbook = SourceBackedWorkbook::from_reader(Cursor::new(bytes)).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();

            assert!(matches!(
                sheet.cell("A1").unwrap(),
                SourceCellView::Stored(Cell::Value(Value::Number(ref value)))
                    if value.as_str() == "7"
            ));
            assert!(sheet.store_resident());
        }

        #[test]
        fn merge_not_eligible_0363_falls_back_and_initializes_store() {
            let bytes = one_sheet_xlsx(&format!(
                r#"<worksheet xmlns="{SPREADSHEETML_NAMESPACE}"><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1" futureAttr="preserve"/></mergeCells></worksheet>"#,
            ));
            let workbook = SourceBackedWorkbook::from_reader(Cursor::new(bytes)).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let cold = workbook.cache_diagnostics();
            assert!(!sheet.store_resident());

            assert!(matches!(
                sheet.cell("B1").unwrap(),
                SourceCellView::Covered(range) if range.a1() == "A1:B1"
            ));
            assert!(sheet.store_resident());
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
            assert!(!sheet.store_resident());
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
            assert!(!sheet.store_resident());
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
            assert!(!sheet.store_resident());
            assert_eq!(
                workbook.cache_diagnostics().successful_loads,
                cold.successful_loads
            );
        }

        #[test]
        fn cancellation_0363_does_not_publish_store() {
            let source = Arc::new(CountingSource::new(source_backed_xlsx()));
            let (_budget, cancellation_source, context) = managed_context(u64::MAX);
            let workbook = SourceBackedWorkbook::from_read_at_with_options(
                source,
                crate::OpenOptions::new()
                    .source_limits(SourceLimits::default())
                    .execution_context(context),
            )
            .unwrap();
            let sheet = workbook.sheet("First").unwrap().unwrap();
            let cold = workbook.cache_diagnostics();
            cancellation_source.cancel();

            assert!(matches!(
                sheet.cell("A1"),
                Err(Error::Package(OpcError::Cancelled))
            ));
            assert!(!sheet.store_resident());
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
            CountingSource, SourceBackedWorkbook, SourceCellView, SourceLimits, VersionFlipSource,
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

            let mut writer = StreamingArchiveWriter::new(
                soapberry_zip::office::StreamingArchiveLimits::default(),
            );
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

        fn styles_cached(workbook: &SourceBackedWorkbook) -> bool {
            workbook.inner.styles.get().is_some()
        }

        fn shared_strings_cached(workbook: &SourceBackedWorkbook) -> bool {
            workbook
                .inner
                .dependencies
                .get()
                .is_some_and(|dependencies| dependencies.shared_strings.get().is_some())
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
            assert!(!sheet.store_resident());
            assert!(
                workbook
                    .inner
                    .dependencies
                    .get()
                    .is_none_or(|dependencies| dependencies.shared_strings.get().is_none())
            );
            assert!(!styles_cached(&workbook));
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
            assert!(!sheet.store_resident());

            let invalid = dependency_xlsx(
                &worksheet(r#"<c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c>"#),
                Some(&plain_shared_strings(&["selected"])),
                None,
            );
            // An unselected dangling reference still forces the materializing
            // fallback, which repairs B1 (audit finding 4.4) and records it.
            let workbook = SourceBackedWorkbook::from_reader(Cursor::new(invalid)).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            assert_text(sheet.cell("A1").unwrap(), "selected");
            assert!(sheet.store_resident());
            assert_eq!(sheet.tolerance_report().unwrap().total(), 1);
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
            assert!(!sheet.store_resident());
            assert!(
                workbook
                    .inner
                    .dependencies
                    .get()
                    .is_none_or(|dependencies| dependencies.shared_strings.get().is_none())
            );

            let rich = dependency_xlsx(
                &worksheet(r#"<c r="A1" t="s"><v>0</v></c>"#),
                Some(&rich_shared_strings()),
                None,
            );
            let workbook = SourceBackedWorkbook::from_reader(Cursor::new(rich)).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            assert_text(sheet.cell("A1").unwrap(), "rich text");
            assert!(sheet.store_resident());
            assert!(shared_strings_cached(&workbook));
        }

        #[test]
        fn invalid_and_missing_dependency_refs_0364_fall_back_to_a_recorded_repair() {
            use crate::leniency::{CellDefect, Defect};
            // Audit findings 4.4 and 4.2: an irregular dependency makes the
            // selected scanner fall back to the materializing parse, which
            // repairs the cell and records the repair in the sheet report.
            for (bytes, defect, style) in [
                (
                    dependency_xlsx(&worksheet(r#"<c r="A1" t="s"><v>0</v></c>"#), None, None),
                    CellDefect::MissingSharedStringTable,
                    None,
                ),
                (
                    dependency_xlsx(
                        &worksheet(r#"<c r="A1" t="s"><v>1</v></c>"#),
                        Some(&plain_shared_strings(&["only zero"])),
                        None,
                    ),
                    CellDefect::SharedStringIndexOutOfRange,
                    None,
                ),
                (
                    dependency_xlsx(&worksheet(r#"<c r="A1" s="0"><v>7</v></c>"#), None, None),
                    CellDefect::StyleIndexOutOfRange,
                    None,
                ),
                (
                    dependency_xlsx(
                        &worksheet(r#"<c r="A1" s="1"><v>7</v></c>"#),
                        None,
                        Some(&styles(1)),
                    ),
                    CellDefect::StyleIndexOutOfRange,
                    Some(0),
                ),
            ] {
                let workbook = SourceBackedWorkbook::from_reader(Cursor::new(bytes)).unwrap();
                assert!(workbook.tolerance_report().is_clean());
                let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
                let view = sheet.cell("A1").expect("repaired cell");
                if matches!(defect, CellDefect::StyleIndexOutOfRange) {
                    assert!(matches!(
                        view,
                        SourceCellView::Stored(Cell::Value(Value::Number(_)))
                    ));
                } else {
                    assert_text(view, "");
                }
                let store = sheet.grid.get().expect("fallback pinned a store").store();
                assert_eq!(store.entries()[0].style, style);
                let report = sheet.tolerance_report().unwrap();
                assert_eq!(report.total(), 1);
                assert_eq!(report.defects()[0].defect, Defect::Cell(defect));
                let aggregate = workbook.tolerance_report();
                assert_eq!(aggregate.total(), 1);
                assert_eq!(aggregate.defects()[0].sheet, Some(0));
            }
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
            assert!(!sheet.store_resident());
            assert!(
                workbook
                    .inner
                    .dependencies
                    .get()
                    .is_none_or(|dependencies| dependencies.shared_strings.get().is_none())
            );
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
            assert!(!sheet.store_resident());
            assert!(!styles_cached(&workbook));
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
            assert!(!sheet.store_resident());
            assert!(
                workbook
                    .inner
                    .dependencies
                    .get()
                    .is_none_or(|dependencies| dependencies.shared_strings.get().is_none())
            );
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
            assert!(!sheet.store_resident());
            assert!(!styles_cached(&workbook));
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
            assert!(!sheet.store_resident());
            assert_eq!(
                workbook.cache_diagnostics().successful_loads,
                cold.successful_loads
            );

            let source = Arc::new(CountingSource::new(bytes));
            let (_budget, cancellation_source, context) = managed_context(u64::MAX);
            let workbook = SourceBackedWorkbook::from_read_at_with_options(
                source,
                crate::OpenOptions::new()
                    .source_limits(SourceLimits::default())
                    .execution_context(context),
            )
            .unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let cold = workbook.cache_diagnostics();
            cancellation_source.cancel();
            assert!(matches!(
                sheet.cell("A1"),
                Err(Error::Package(OpcError::Cancelled))
            ));
            assert!(!sheet.store_resident());
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
            assert!(!sheet.store_resident());
            assert!(
                workbook
                    .inner
                    .dependencies
                    .get()
                    .is_none_or(|dependencies| dependencies.shared_strings.get().is_none())
            );
            assert!(!styles_cached(&workbook));
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
            let workbook = SourceBackedWorkbook::from_read_at_with_options(
                source,
                crate::OpenOptions::new().cache_limits(cache_limits),
            )
            .unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            assert_text(sheet.cell("A1").unwrap(), "rich text");
            assert!(sheet.store_resident());
            assert!(shared_strings_cached(&workbook));
            assert!(styles_cached(&workbook));

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
            assert!(!sheet.store_resident());
        }
    }

    mod streaming_0365_tests {
        use std::sync::Arc;
        use std::sync::atomic::Ordering;

        use litchi_opc::{OpcError, SourceBackedPackage, SourceCacheLimits, SourceReadPolicy};

        use super::super::SourceCell;
        use super::streaming_0364_tests::{dependency_xlsx, plain_shared_strings, styles};
        use super::{
            CountingSource, SourceBackedWorkbook, SourceCellView, SourceLimits, managed_context,
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
            assert!(!sheet.store_resident());
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
            assert!(!sheet.store_resident());

            let subset = sheet.cells("B1:C2").unwrap();
            assert_eq!(addresses(&subset), vec!["C1".to_owned(), "B2".to_owned()]);
            assert!(matches!(&subset[0].cell, Cell::Empty));
            assert!(matches!(
                &subset[1].cell,
                Cell::Value(Value::Number(value)) if value.as_str() == "8"
            ));

            assert!(sheet.cells("Z10:AA11").unwrap().is_empty());
            assert!(!sheet.store_resident());
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
            assert!(!sheet.store_resident());
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
            assert!(!sheet.store_resident());
            assert!(
                workbook
                    .inner
                    .dependencies
                    .get()
                    .is_none_or(|dependencies| dependencies.shared_strings.get().is_none())
            );
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
            assert!(!sheet.store_resident());
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
            assert!(!sheet.store_resident());
            assert!(sheet.stored_extent().unwrap().is_some());
            assert!(sheet.store_resident());
            let warm = sheet.cells("A1:D4").unwrap();

            assert_eq!(warm, cold);
        }

        #[test]
        fn merge_and_shared_formula_0365_fall_back_to_store() {
            let merge = worksheet_xml(
                r#"<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1" futureAttr="preserve"/></mergeCells>"#,
            );
            let (_source, workbook) = source_workbook(&merge, None, None);
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let cells = sheet.cells("A1:B1").unwrap();
            assert_eq!(addresses(&cells), vec!["A1".to_owned()]);
            assert!(sheet.store_resident());
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
            assert!(sheet.store_resident());
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
            assert!(!sheet.store_resident());
            assert!(
                workbook
                    .inner
                    .dependencies
                    .get()
                    .is_none_or(|dependencies| dependencies.shared_strings.get().is_none())
            );
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
            assert!(!sheet.store_resident());
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
            let workbook = SourceBackedWorkbook::from_read_at_with_options(
                source,
                crate::OpenOptions::new()
                    .source_limits(SourceLimits::default())
                    .execution_context(context),
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
            assert!(!sheet.store_resident());
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
            assert!(!sheet.store_resident());
        }

        #[test]
        fn callback_cancellation_0365_uses_package_owned_context() {
            let source = Arc::new(CountingSource::new(dependency_xlsx(
                &stored_fixture(),
                None,
                None,
            )));
            let (_budget, cancellation_source, context) = managed_context(u64::MAX);
            let package = SourceBackedPackage::from_read_at(
                source,
                SourceLimits::default(),
                SourceCacheLimits::new(256 * 1024, 4).unwrap(),
                SourceReadPolicy::exact(),
                Some(context),
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
            assert!(!sheet.store_resident());
        }

        #[test]
        fn final_source_change_0365_outranks_callback_error() {
            let source = Arc::new(CountingSource::new(dependency_xlsx(
                &stored_fixture(),
                None,
                None,
            )));
            let (_budget, _cancellation_source, context) = managed_context(u64::MAX);
            let package = SourceBackedPackage::from_read_at(
                source.clone(),
                SourceLimits::default(),
                SourceCacheLimits::default(),
                SourceReadPolicy::exact(),
                Some(context),
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
            assert!(!sheet.store_resident());
        }
    }

    mod streaming_0366_tests {
        use std::sync::Arc;

        use litchi_opc::SourceCacheDiagnostics;

        use super::super::SourceCell;
        use super::streaming_0364_tests::dependency_xlsx;
        use super::{CountingSource, SourceBackedWorkbook, SourceCellView};
        use crate::{Cell, Value};

        const SPREADSHEETML_NAMESPACE: &str =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        fn worksheet_xml(sheet_data: &str) -> String {
            let first = std::str::from_utf8(super::FIRST_MARKER).unwrap();
            let second = std::str::from_utf8(super::SECOND_MARKER).unwrap();
            format!(
                r#"<worksheet xmlns="{SPREADSHEETML_NAMESPACE}"><!--{first} {second}-->{sheet_data}</worksheet>"#,
            )
        }

        fn source_workbook(worksheet: &str) -> (Arc<CountingSource>, SourceBackedWorkbook) {
            let source = Arc::new(CountingSource::new(dependency_xlsx(worksheet, None, None)));
            let workbook = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
            (source, workbook)
        }

        fn reference_fixture() -> String {
            worksheet_xml(
                r#"<sheetData>
                    <row r="1">
                        <c r="A1" t="inlineStr"><is><t>inline&amp;&lt;&gt;&apos;&quot;</t></is></c>
                        <c r="C1" t="str"><v>value&amp;&lt;&gt;&apos;&quot;</v></c>
                        <c r="E1"><f>A1&amp;C1&lt;&gt;D1</f><v>1</v></c>
                    </row>
                    <row r="3">
                        <c r="A3"><f>A1&#x2b;C1</f><v>3&#x2e;5</v></c>
                        <c r="D3"><v>42&#x2e;25</v></c>
                    </row>
                </sheetData>"#,
            )
        }

        fn addresses(cells: &[SourceCell]) -> Vec<String> {
            cells.iter().map(|cell| cell.address.a1()).collect()
        }

        fn assert_no_part_cache_admission(
            before: SourceCacheDiagnostics,
            after: SourceCacheDiagnostics,
        ) {
            assert_eq!(after.cold_loads, before.cold_loads);
            assert_eq!(after.successful_loads, before.successful_loads);
            assert_eq!(after.retained_entries, before.retained_entries);
            assert_eq!(after.retained_bytes, before.retained_bytes);
        }

        #[test]
        fn cold_cell_0366_decodes_references_with_exact_eager_parity() {
            let (_source, workbook) = source_workbook(&reference_fixture());
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let before = workbook.cache_diagnostics();
            let references = ["A1", "C1", "E1", "A3", "D3"];

            let cold = references
                .into_iter()
                .map(|reference| sheet.cell(reference).unwrap())
                .collect::<Vec<_>>();

            assert!(matches!(
                &cold[0],
                SourceCellView::Stored(Cell::Value(Value::Text(value)))
                    if value.as_str() == "inline&<>'\""
            ));
            assert!(matches!(
                &cold[1],
                SourceCellView::Stored(Cell::Value(Value::Text(value)))
                    if value.as_str() == "value&<>'\""
            ));
            assert!(matches!(
                &cold[2],
                SourceCellView::Stored(Cell::Formula(formula))
                    if formula.text() == "A1&C1<>D1"
            ));
            assert!(matches!(
                &cold[3],
                SourceCellView::Stored(Cell::Formula(formula))
                    if formula.text() == "A1+C1"
            ));
            assert!(matches!(
                &cold[4],
                SourceCellView::Stored(Cell::Value(Value::Number(value)))
                    if value.as_str() == "42.25"
            ));
            assert!(!sheet.store_resident());
            assert!(
                workbook
                    .inner
                    .dependencies
                    .get()
                    .is_none_or(|dependencies| dependencies.shared_strings.get().is_none())
            );
            assert!(workbook.inner.styles.get().is_none());
            assert_no_part_cache_admission(before, workbook.cache_diagnostics());

            sheet.stored_extent().unwrap();
            let eager = ["A1", "C1", "E1", "A3", "D3"]
                .into_iter()
                .map(|reference| sheet.cell(reference).unwrap())
                .collect::<Vec<_>>();
            assert_eq!(cold, eager);
        }

        #[test]
        fn cold_cells_0366_are_sparse_row_major_and_stay_cold() {
            let (_source, workbook) = source_workbook(&reference_fixture());
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let before = workbook.cache_diagnostics();

            let cold = sheet.cells("A1:E3").unwrap();
            assert_eq!(
                addresses(&cold),
                vec![
                    "A1".to_owned(),
                    "C1".to_owned(),
                    "E1".to_owned(),
                    "A3".to_owned(),
                    "D3".to_owned(),
                ]
            );
            assert!(!sheet.store_resident());
            assert!(
                workbook
                    .inner
                    .dependencies
                    .get()
                    .is_none_or(|dependencies| dependencies.shared_strings.get().is_none())
            );
            assert!(workbook.inner.styles.get().is_none());
            assert_no_part_cache_admission(before, workbook.cache_diagnostics());

            sheet.stored_extent().unwrap();
            let eager = sheet.cells("A1:E3").unwrap();
            assert_eq!(cold, eager);
        }

        #[test]
        fn cold_visit_cells_0366_stages_decoded_callbacks_without_admission() {
            let (_source, workbook) = source_workbook(&reference_fixture());
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let before = workbook.cache_diagnostics();
            let mut visited = Vec::new();

            let count = sheet
                .visit_cells("A1:E3", |address, cell| {
                    visited.push((address, cell.clone()));
                    Ok(())
                })
                .unwrap();

            assert_eq!(count, visited.len());
            assert_eq!(
                visited
                    .iter()
                    .map(|(address, _)| address.a1())
                    .collect::<Vec<_>>(),
                vec![
                    "A1".to_owned(),
                    "C1".to_owned(),
                    "E1".to_owned(),
                    "A3".to_owned(),
                    "D3".to_owned(),
                ]
            );
            assert!(!sheet.store_resident());
            assert!(
                workbook
                    .inner
                    .dependencies
                    .get()
                    .is_none_or(|dependencies| dependencies.shared_strings.get().is_none())
            );
            assert!(workbook.inner.styles.get().is_none());
            assert_no_part_cache_admission(before, workbook.cache_diagnostics());

            sheet.stored_extent().unwrap();
            let eager = sheet
                .cells("A1:E3")
                .unwrap()
                .into_iter()
                .map(|cell| (cell.address, cell.cell))
                .collect::<Vec<_>>();
            assert_eq!(visited, eager);
        }

        fn assert_eager_fallback_parity(reference: &str, literal: &str) {
            let (_source, fallback_workbook) = source_workbook(reference);
            let fallback_sheet = fallback_workbook.sheet("Sheet1").unwrap().unwrap();
            let before = fallback_workbook.cache_diagnostics();

            let (_source, eager_workbook) = source_workbook(literal);
            let eager_sheet = eager_workbook.sheet("Sheet1").unwrap().unwrap();
            eager_sheet.stored_extent().unwrap();
            let expected = eager_sheet.cell("A1").unwrap();

            let actual = fallback_sheet.cell("A1").unwrap();
            assert_eq!(actual, expected);
            assert!(fallback_sheet.store_resident());
            let after = fallback_workbook.cache_diagnostics();
            assert!(after.successful_loads > before.successful_loads);
            assert!(after.retained_entries > before.retained_entries);
        }

        #[test]
        fn valid_overlong_and_numeric_inline_refs_0366_require_eager_fallback() {
            assert_eager_fallback_parity(
                &worksheet_xml(
                    r#"<sheetData><row r="1"><c r="A1" t="str"><v>&#00000000065;</v></c></row></sheetData>"#,
                ),
                &worksheet_xml(
                    r#"<sheetData><row r="1"><c r="A1" t="str"><v>A</v></c></row></sheetData>"#,
                ),
            );
            assert_eager_fallback_parity(
                &worksheet_xml(
                    r#"<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>&#65;</t></is></c></row></sheetData>"#,
                ),
                &worksheet_xml(
                    r#"<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>A</t></is></c></row></sheetData>"#,
                ),
            );
        }

        fn assert_zero_callbacks_without_publication(bytes: Vec<u8>) {
            let source = Arc::new(CountingSource::new(bytes));
            let workbook = SourceBackedWorkbook::from_read_at(source).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let before = workbook.cache_diagnostics();
            let mut callbacks = 0usize;

            let result = sheet.visit_cells("A1:E3", |_address, _cell| {
                callbacks += 1;
                Ok(())
            });

            assert!(result.is_err());
            assert_eq!(callbacks, 0);
            assert!(!sheet.store_resident());
            assert!(
                workbook
                    .inner
                    .dependencies
                    .get()
                    .is_none_or(|dependencies| dependencies.shared_strings.get().is_none())
            );
            assert!(workbook.inner.styles.get().is_none());
            assert_no_part_cache_admission(before, workbook.cache_diagnostics());
        }

        #[test]
        fn malformed_reference_tail_and_crc_0366_have_no_callbacks_or_publication() {
            assert_zero_callbacks_without_publication(dependency_xlsx(
                &worksheet_xml(
                    r#"<sheetData><row r="1"><c r="A1" t="str"><v>&unknown;</v></c></row></sheetData>"#,
                ),
                None,
                None,
            ));

            let malformed_tail = format!(
                "{}<tail>",
                worksheet_xml(r#"<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>"#,)
            );
            assert_zero_callbacks_without_publication(dependency_xlsx(&malformed_tail, None, None));

            let mut crc = dependency_xlsx(
                &worksheet_xml(r#"<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>"#),
                None,
                None,
            );
            let marker = b"<v>7</v>";
            let offset = crc
                .windows(marker.len())
                .position(|window| window == marker)
                .unwrap();
            crc[offset + 3] = b'8';
            assert_zero_callbacks_without_publication(crc);
        }
    }

    mod streaming_0367_tests {
        use std::sync::Arc;

        use litchi_opc::{OpcError, SourceCacheDiagnostics};

        use super::super::SourceCell;
        use super::streaming_0364_tests::dependency_xlsx;
        use super::{
            CountingSource, SourceBackedWorkbook, SourceCellView, SourceLimits, managed_context,
        };
        use crate::{Cell, Error, Value};

        const SPREADSHEETML_NAMESPACE: &str =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        fn worksheet_xml(sheet_data: &str) -> String {
            let first = std::str::from_utf8(super::FIRST_MARKER).unwrap();
            let second = std::str::from_utf8(super::SECOND_MARKER).unwrap();
            format!(
                r#"<worksheet xmlns="{SPREADSHEETML_NAMESPACE}"><!--{first} {second}-->{sheet_data}</worksheet>"#,
            )
        }

        fn worksheet_xml_after(sheet_data: &str, after_root: &str) -> String {
            let first = std::str::from_utf8(super::FIRST_MARKER).unwrap();
            let second = std::str::from_utf8(super::SECOND_MARKER).unwrap();
            format!(
                r#"<worksheet xmlns="{SPREADSHEETML_NAMESPACE}"><!--{first} {second}-->{sheet_data}</worksheet>{after_root}"#,
            )
        }

        fn source_workbook(
            worksheet: &str,
            styles: Option<&str>,
        ) -> (Arc<CountingSource>, SourceBackedWorkbook) {
            let source = Arc::new(CountingSource::new(dependency_xlsx(
                worksheet, None, styles,
            )));
            let workbook = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
            (source, workbook)
        }

        fn addresses(cells: &[SourceCell]) -> Vec<String> {
            cells.iter().map(|cell| cell.address.a1()).collect()
        }

        fn assert_no_part_cache_admission(
            before: SourceCacheDiagnostics,
            after: SourceCacheDiagnostics,
        ) {
            assert_eq!(after.cold_loads, before.cold_loads);
            assert_eq!(after.successful_loads, before.successful_loads);
            assert_eq!(after.retained_entries, before.retained_entries);
            assert_eq!(after.retained_bytes, before.retained_bytes);
        }

        fn assert_zero_callbacks(bytes: Vec<u8>) {
            let source = Arc::new(CountingSource::new(bytes));
            let workbook = SourceBackedWorkbook::from_read_at(source).unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let before = workbook.cache_diagnostics();
            let mut callbacks = Vec::new();

            let result = sheet.visit_cells("A1:E3", |address, _cell| {
                callbacks.push(address.a1());
                Ok(())
            });

            assert!(result.is_err());
            assert!(callbacks.is_empty());
            assert!(!sheet.store_resident());
            assert!(
                workbook
                    .inner
                    .dependencies
                    .get()
                    .is_none_or(|dependencies| dependencies.shared_strings.get().is_none())
            );
            assert!(workbook.inner.styles.get().is_none());
            assert_no_part_cache_admission(before, workbook.cache_diagnostics());
        }

        #[test]
        fn cold_merge_cell_views_0367_match_forced_eager_without_admission() {
            let worksheet = worksheet_xml(
                r#"<sheetData><row r="1"><c r="A1"><v>7</v></c><c r="B1"><v>8</v></c><c r="E1"/></row></sheetData><mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells>"#,
            );
            let (_source, workbook) = source_workbook(&worksheet, None);
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let before = workbook.cache_diagnostics();
            let references = ["A1", "B1", "C1", "D1", "E1"];

            let cold = references
                .into_iter()
                .map(|reference| sheet.cell(reference).unwrap())
                .collect::<Vec<_>>();

            assert!(matches!(
                &cold[0],
                SourceCellView::Stored(Cell::Value(Value::Number(value)))
                    if value.as_str() == "7"
            ));
            assert!(matches!(
                &cold[1],
                SourceCellView::Covered(range) if range.a1() == "A1:C1"
            ));
            assert!(matches!(
                &cold[2],
                SourceCellView::Covered(range) if range.a1() == "A1:C1"
            ));
            assert!(matches!(&cold[3], SourceCellView::Missing));
            assert!(matches!(&cold[4], SourceCellView::Stored(Cell::Empty)));
            assert!(!sheet.store_resident());
            assert!(
                workbook
                    .inner
                    .dependencies
                    .get()
                    .is_none_or(|dependencies| dependencies.shared_strings.get().is_none())
            );
            assert!(workbook.inner.styles.get().is_none());
            assert_no_part_cache_admission(before, workbook.cache_diagnostics());

            sheet.stored_extent().unwrap();
            let eager = references
                .into_iter()
                .map(|reference| sheet.cell(reference).unwrap())
                .collect::<Vec<_>>();
            assert_eq!(cold, eager);
        }

        #[test]
        fn cold_cells_and_visit_0367_are_sparse_physical_and_merge_coverage_is_not_synthetic() {
            let worksheet = worksheet_xml(
                r#"<sheetData><row r="1"><c r="A1"><v>7</v></c><c r="B1"><v>8</v></c><c r="E1"/></row><row r="3"><c r="C3"><v>9</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells>"#,
            );
            let (_source, workbook) = source_workbook(&worksheet, None);
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let before = workbook.cache_diagnostics();

            let cold = sheet.cells("A1:E3").unwrap();
            assert_eq!(
                addresses(&cold),
                vec![
                    "A1".to_owned(),
                    "B1".to_owned(),
                    "E1".to_owned(),
                    "C3".to_owned(),
                ]
            );
            assert!(matches!(
                &cold[1].cell,
                Cell::Value(Value::Number(value)) if value.as_str() == "8"
            ));
            assert!(matches!(&cold[2].cell, Cell::Empty));
            assert!(!sheet.store_resident());

            let mut visited = Vec::new();
            let count = sheet
                .visit_cells("A1:E3", |address, cell| {
                    visited.push((address, cell.clone()));
                    Ok(())
                })
                .unwrap();
            let expected = cold
                .iter()
                .map(|cell| (cell.address, cell.cell.clone()))
                .collect::<Vec<_>>();
            assert_eq!(count, expected.len());
            assert_eq!(visited, expected);
            assert!(!sheet.store_resident());
            assert!(
                workbook
                    .inner
                    .dependencies
                    .get()
                    .is_none_or(|dependencies| dependencies.shared_strings.get().is_none())
            );
            assert!(workbook.inner.styles.get().is_none());
            assert_no_part_cache_admission(before, workbook.cache_diagnostics());

            sheet.stored_extent().unwrap();
            assert_eq!(cold, sheet.cells("A1:E3").unwrap());
        }

        #[test]
        fn malformed_merge_metadata_0367_has_zero_callbacks_and_publication() {
            let malformed = worksheet_xml(
                r#"<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1">"#,
            );
            let overlap = worksheet_xml(
                r#"<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData><mergeCells count="2"><mergeCell ref="A1:B2"/><mergeCell ref="B2:C3"/></mergeCells>"#,
            );
            // A merge count disagreeing with its records is repaired and
            // recorded (leniency finding 4.10), so it is not listed here.
            let out_of_grid = worksheet_xml(
                r#"<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="XFE1:XFF1"/></mergeCells>"#,
            );
            let tail = worksheet_xml_after(
                r#"<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>"#,
                "<tail>",
            );

            for worksheet in [malformed, overlap, out_of_grid, tail] {
                assert_zero_callbacks(dependency_xlsx(&worksheet, None, None));
            }
        }

        #[test]
        fn crc_source_change_and_cancellation_0367_have_zero_callbacks() {
            let valid = worksheet_xml(
                r#"<sheetData><row r="1"><c r="A1"><v>7</v></c><c r="B1"><v>8</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>"#,
            );
            let mut crc = dependency_xlsx(&valid, None, None);
            let marker = b"<v>7</v>";
            let offset = crc
                .windows(marker.len())
                .position(|window| window == marker)
                .unwrap();
            crc[offset + 3] = b'9';
            assert_zero_callbacks(crc);

            let (source, workbook) = source_workbook(&valid, None);
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let before = workbook.cache_diagnostics();
            source.changed();
            let mut callbacks = 0usize;
            let error = sheet
                .visit_cells("A1:E3", |_address, _cell| {
                    callbacks += 1;
                    Ok(())
                })
                .unwrap_err();
            assert!(matches!(
                error,
                Error::Package(OpcError::SourceChanged { .. })
            ));
            assert_eq!(callbacks, 0);
            assert!(!sheet.store_resident());
            assert_no_part_cache_admission(before, workbook.cache_diagnostics());

            let source = Arc::new(CountingSource::new(dependency_xlsx(&valid, None, None)));
            let (_budget, cancellation_source, context) = managed_context(u64::MAX);
            let workbook = SourceBackedWorkbook::from_read_at_with_options(
                source,
                crate::OpenOptions::new()
                    .source_limits(SourceLimits::default())
                    .execution_context(context),
            )
            .unwrap();
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let before = workbook.cache_diagnostics();
            cancellation_source.cancel();
            let mut callbacks = 0usize;
            let error = sheet
                .visit_cells("A1:E3", |_address, _cell| {
                    callbacks += 1;
                    Ok(())
                })
                .unwrap_err();
            assert!(matches!(error, Error::Package(OpcError::Cancelled)));
            assert_eq!(callbacks, 0);
            assert!(!sheet.store_resident());
            assert_no_part_cache_admission(before, workbook.cache_diagnostics());
        }

        fn assert_fallback_parity(
            fallback: &str,
            styles: Option<&str>,
            eager: &str,
            references: &[&str],
        ) {
            let (_source, fallback_workbook) = source_workbook(fallback, styles);
            let fallback_sheet = fallback_workbook.sheet("Sheet1").unwrap().unwrap();
            let before = fallback_workbook.cache_diagnostics();
            let actual = references
                .iter()
                .map(|reference| fallback_sheet.cell(*reference).unwrap())
                .collect::<Vec<_>>();
            assert!(fallback_sheet.store_resident());
            let after = fallback_workbook.cache_diagnostics();
            assert!(after.successful_loads > before.successful_loads);

            let (_source, eager_workbook) = source_workbook(eager, styles);
            let eager_sheet = eager_workbook.sheet("Sheet1").unwrap().unwrap();
            eager_sheet.stored_extent().unwrap();
            let expected = references
                .iter()
                .map(|reference| eager_sheet.cell(*reference).unwrap())
                .collect::<Vec<_>>();
            assert_eq!(actual, expected);
        }

        #[test]
        fn merge_and_unsupported_semantics_0367_match_eager_with_store_admission() {
            let shared_formula = worksheet_xml(
                r#"<sheetData><row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="7">B1</f><v>1</v></c><c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="7"/><v>3</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>"#,
            );
            let normal_formula = worksheet_xml(
                r#"<sheetData><row r="1"><c r="A1"><f>B1</f><v>1</v></c><c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><f>B2</f><v>3</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>"#,
            );
            assert_fallback_parity(&shared_formula, None, &normal_formula, &["A1", "B1"]);

            let row_style = worksheet_xml(
                r#"<sheetData><row r="1" s="0"><c r="A1"><v>7</v></c></row></sheetData>"#,
            );
            let unstyled =
                worksheet_xml(r#"<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>"#);
            let style_catalog = super::streaming_0364_tests::styles(1);
            assert_fallback_parity(&row_style, Some(&style_catalog), &unstyled, &["A1"]);

            let general_reference = worksheet_xml(
                r#"<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>&#x41;</t></is></c></row></sheetData>"#,
            );
            let decoded = worksheet_xml(
                r#"<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>A</t></is></c></row></sheetData>"#,
            );
            assert_fallback_parity(&general_reference, None, &decoded, &["A1"]);
        }

        #[test]
        fn numeric_value_scratch_0400_preserves_late_fallback_parity() {
            let numeric_with_late_date = worksheet_xml(
                r#"<sheetData><row r="1"><c r="A1"><v>-0.000</v></c><c r="B1" t="d"><v>2026-07-31T12:34:56.250-07:00</v></c></row></sheetData>"#,
            );
            let eager_numeric = worksheet_xml(
                r#"<sheetData><row r="1"><c r="A1"><v>-0.000</v></c></row></sheetData>"#,
            );
            assert_fallback_parity(&numeric_with_late_date, None, &eager_numeric, &["A1"]);
        }
    }

    /// `visit_cells` streams its cells instead of materializing the range.
    ///
    /// Change 0642 stopped building a whole-range `Vec<SourceCell>` before the
    /// first callback. These tests pin what that must not change: the visited
    /// sequence on both routes, and the rule that a refusal the whole-range
    /// read raises still arrives before any cell is visited, including a
    /// refusal that the worksheet raises after the requested rectangle.
    mod streaming_0642_tests {
        use std::sync::Arc;

        use super::streaming_0364_tests::{dependency_xlsx, plain_shared_strings};
        use super::{CountingSource, SourceBackedWorkbook};
        use crate::Cell;

        const SPREADSHEETML_NAMESPACE: &str =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        fn worksheet_xml(sheet_data: &str) -> String {
            let first = std::str::from_utf8(super::FIRST_MARKER).unwrap();
            let second = std::str::from_utf8(super::SECOND_MARKER).unwrap();
            format!(
                r#"<worksheet xmlns="{SPREADSHEETML_NAMESPACE}"><!--{first} {second}-->{sheet_data}</worksheet>"#,
            )
        }

        fn source_workbook(
            worksheet: &str,
            shared_strings: Option<&str>,
        ) -> (Arc<CountingSource>, SourceBackedWorkbook) {
            let source = Arc::new(CountingSource::new(dependency_xlsx(
                worksheet,
                shared_strings,
                None,
            )));
            let workbook = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
            (source, workbook)
        }

        fn stored_fixture() -> String {
            worksheet_xml(
                r#"<sheetData>
                    <row r="1"><c r="A1"><v>7</v></c><c r="C1"/></row>
                    <row r="2"><c r="B2" t="s"><v>0</v></c></row>
                    <row r="4"><c r="A4"><v>9</v></c></row>
                </sheetData>"#,
            )
        }

        fn visited(workbook: &SourceBackedWorkbook, warm: bool) -> (usize, Vec<(String, Cell)>) {
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            if warm {
                sheet.stored_extent().unwrap();
                assert!(sheet.store_resident());
            }
            let mut cells = Vec::new();
            let count = sheet
                .visit_cells("A1:D4", |address, cell| {
                    cells.push((address.a1(), cell.clone()));
                    Ok(())
                })
                .unwrap();
            (count, cells)
        }

        #[test]
        fn visit_cells_0642_matches_cells_on_both_routes() {
            let fixture = stored_fixture();
            let strings = plain_shared_strings(&["shared"]);

            let (_source, workbook) = source_workbook(&fixture, Some(&strings));
            let expected = workbook
                .sheet("Sheet1")
                .unwrap()
                .unwrap()
                .cells("A1:D4")
                .unwrap()
                .into_iter()
                .map(|cell| (cell.address.a1(), cell.cell))
                .collect::<Vec<_>>();
            assert_eq!(expected.len(), 4);

            let (_source, workbook) = source_workbook(&fixture, Some(&strings));
            let (cold_count, cold) = visited(&workbook, false);
            assert_eq!(cold_count, expected.len());
            assert_eq!(cold, expected);

            let (_source, workbook) = source_workbook(&fixture, Some(&strings));
            let (warm_count, warm) = visited(&workbook, true);
            assert_eq!(warm_count, expected.len());
            assert_eq!(warm, expected);
        }

        /// A worksheet that refuses after the requested rectangle must refuse
        /// before the first callback, exactly as the whole-range read does.
        #[test]
        fn visit_cells_0642_refuses_a_later_row_before_any_visit() {
            let fixture = worksheet_xml(
                r#"<sheetData>
                    <row r="1"><c r="A1"><v>7</v></c></row>
                    <row r="9"><c r="A9" t="inlineStr"><is><t>x</t></is><v>3</v></c></row>
                </sheetData>"#,
            );

            let (_source, workbook) = source_workbook(&fixture, None);
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let whole_range = sheet.cells("A1:A2").unwrap_err().to_string();
            assert!(
                whole_range.contains("both inline text and a value"),
                "unexpected refusal: {whole_range}"
            );

            let (_source, workbook) = source_workbook(&fixture, None);
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let mut callbacks = 0usize;
            let streamed = sheet
                .visit_cells("A1:A2", |_address, _cell| {
                    callbacks += 1;
                    Ok(())
                })
                .unwrap_err()
                .to_string();

            assert_eq!(callbacks, 0);
            assert_eq!(streamed, whole_range);
        }

        /// The stored route counts and stops exactly like the streamed one.
        #[test]
        fn visit_cells_0642_stops_on_the_stored_route_at_the_failing_cell() {
            let strings = plain_shared_strings(&["shared"]);
            let (_source, workbook) = source_workbook(&stored_fixture(), Some(&strings));
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            sheet.stored_extent().unwrap();
            assert!(sheet.store_resident());

            let mut callbacks = 0usize;
            let error = sheet
                .visit_cells("A1:D4", |_address, _cell| {
                    callbacks += 1;
                    if callbacks == 3 {
                        return Err(crate::error::invalid("streaming_0642 callback: stop"));
                    }
                    Ok(())
                })
                .unwrap_err()
                .to_string();

            assert_eq!(callbacks, 3);
            assert!(
                error.contains("streaming_0642 callback: stop"),
                "unexpected refusal: {error}"
            );
        }

        #[test]
        fn cells_0672_reserve_exactly_the_stored_selection() {
            let fixture = worksheet_xml(
                r#"<sheetData><row r="1">
                    <c r="A1"><v>1</v></c><c r="B1"><v>2</v></c>
                    <c r="C1"><v>3</v></c><c r="D1"><v>4</v></c>
                    <c r="E1"><v>5</v></c>
                </row></sheetData>"#,
            );
            let (_source, workbook) = source_workbook(&fixture, None);
            let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
            sheet.stored_extent().unwrap();

            let cells = sheet.cells("A1:E1").unwrap();
            assert_eq!(cells.len(), 5);
            assert_eq!(cells.capacity(), cells.len());
        }
    }
}

#[cfg(test)]
mod package_handoff_tests {
    use std::io::Cursor;

    use super::SourceBackedWorkbook;

    fn fixture() -> Vec<u8> {
        std::fs::read(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/ooxml/xlsx/DateFormatTests.xlsx"
        ))
        .expect("read XLSX fixture")
    }

    #[test]
    fn package_handoff_returns_the_package_indexed_at_open() {
        let workbook = SourceBackedWorkbook::from_reader(Cursor::new(fixture())).unwrap();
        let lineage = workbook.source_backed_package().unwrap().source_lineage();
        assert_eq!(
            workbook
                .source_backed_package()
                .unwrap()
                .source_version()
                .unwrap(),
            workbook.source_version().unwrap()
        );
        let selected = workbook.sheet(0_usize).unwrap().expect("first sheet");

        // A selected worksheet handle and a workbook clone both share the
        // catalog, so the handoff refuses and returns the workbook intact.
        let workbook = workbook
            .try_into_source_backed_package()
            .err()
            .expect("a selected worksheet shares the catalog");
        drop(selected);
        let clone = workbook.clone();
        let workbook = workbook
            .try_into_source_backed_package()
            .err()
            .expect("a workbook clone shares the catalog");
        assert_eq!(clone.len(), workbook.len());
        drop(clone);

        let package = workbook
            .try_into_source_backed_package()
            .expect("the sole owner recovers its package");
        // Same lineage token: the very package opened above, not a re-index.
        assert!(package.source_lineage() == lineage);
        let workbook = SourceBackedWorkbook::from_source_backed_package(package).unwrap();
        workbook
            .edit_tab_state()
            .expect("a source-backed family edit adopts the recovered package");
    }
}
