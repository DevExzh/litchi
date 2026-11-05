//! Handle for Excel OOXML (`.xlsx`) inputs.
//!
//! [`Xlsx`] is one `Arc` over its state: the two-state owner of record 0003
//! §2.3 (`Source` for the deferred, source-backed catalog; `Eager` for the
//! owning workbook snapshot), the core-properties metadata and its repairs
//! read at open (§1.6), the lazily built worksheet-name list, and the
//! retention policy from [`OpenOptions`]. Cloning bumps one reference count.
//!
//! # Routes
//!
//! The constructors take what detection already produced, so an open does
//! one detection pass and one container index:
//!
//! | Open route | Detection result | Constructor |
//! |---|---|---|
//! | `open` (path) on `unix`/`windows`, `open_bytes`, `open_prepared` | source-backed OPC package | `Xlsx::from_source_package` |
//! | [`SourceMode::Eager`], and paths elsewhere | eager `OpcPackage` | `Xlsx::from_opc_package` |
//! | `Locked::unlock` | decrypted package | `Xlsx::from_package` |
//!
//! The facade's filesystem route read the core properties while it built the
//! catalog; the unified route reads them from the same package in
//! `Xlsx::from_source_package`, so the work is the same.
//!
//! # Source freshness
//!
//! `XlsxOwner::ensure_current` is the facade adapter's
//! `ensure_source_current`. Catalog reads (names, count, metadata, tolerance
//! report) check the source revision first, as the facade did. Reads that
//! touch worksheet payloads ([`Xlsx::text`], [`Xlsx::sheet`]) are bracketed by
//! `XlsxOwner::finish`, which checks before and after, so a changed source
//! wins over what the read produced. Cell reads through a selected or generic
//! worksheet are fenced by the source-backed leaf itself.
//!
//! # Generic sheet traits
//!
//! [`Xlsx`] implements [`WorkbookTrait`] directly, so the workbook view that
//! S6 adds delegates to it by a `match`, with no boxed workbook. Its
//! worksheets and cells are the statically dispatched
//! [`crate::view::sheet`] types, and its errors are typed [`crate::Error`]
//! values.

use std::mem::size_of;
use std::path::Path;
use std::sync::{Arc, Mutex, OnceLock};

use litchi_core::sheet::WorkbookTrait;
use litchi_core::{FileFormat, Metadata};
use litchi_opc::{OpcPackage, SourceBackedPackage};
use litchi_xlsx as xlsx;

use super::cell_text;
use super::signature::{self, Opened, Signature};
use crate::detect::workbook_metadata::{MetadataRepairs, read_source_backed_metadata};
use crate::error::{Code, Error, Result};
use crate::options::{OpenOptions, SourceMode};

#[cfg(feature = "test-support")]
#[doc(hidden)]
#[path = "xlsx/support.rs"]
pub mod __unified;
#[cfg(feature = "eval")]
pub(crate) mod calc;
mod cell_add;
mod chart;
mod children;
pub(crate) mod create;
mod dimension_write;
mod feature_write;
mod features;
mod header;
mod merge_write;
mod metadata;
mod nav;
mod node;
mod publish;
mod runs;
mod selection;
pub(crate) mod sheet;
mod sheet_write;
mod spec;
mod structure;
mod views;
pub(crate) mod write;

pub(crate) use chart::{AXIS_CAPABILITY_PROPS, CHART_CAPABILITY_PROPS, SERIES_CAPABILITY_PROPS};
pub(crate) use children::Children;
pub(crate) use features::FEATURE_ELEMENTS;
pub(crate) use nav::Nav;
pub use node::Node;
pub(crate) use node::Props;
pub(crate) use runs::RUN_WRITE_TABLE;
pub(crate) use spec::{
    CELL_CAPABILITY_PROPS, COLUMN_CAPABILITY_PROPS, RANGE_CAPABILITY_PROPS, ROW_CAPABILITY_PROPS,
    SHEET_CAPABILITY_PROPS, SPEC, WORKBOOK_CAPABILITY_PROPS,
};
pub(crate) use spec::{CELL_WRITE_TABLE, COLUMN_WRITE_TABLE, ROW_WRITE_TABLE};
pub(crate) use views::{Json as ViewJson, view_json, view_to};

pub use selection::{SelectedCell, SelectedCellView, SelectedWorksheet, WorksheetSelector};
use sheet::Sheet;

/// An opened Excel OOXML workbook.
///
/// Cheap to clone and `Send + Sync`. Read methods take `&self`; on a
/// source-backed workbook each read checks that the file has not changed
/// since it was opened and reports [`Code::IoError`] when it has.
///
/// ```no_run
/// # fn demo(workbook: &litchi_ops::format::xlsx::Xlsx) -> litchi_ops::Result<()> {
/// println!("{:?}", workbook.worksheet_names()?);
/// let text = workbook.text()?;
/// // Sheet names match case-insensitively, as in Excel.
/// let data = workbook.sheet("dAtA")?;
/// # let _ = (text, data);
/// # Ok(())
/// # }
/// ```
#[derive(Clone)]
pub struct Xlsx(Arc<Opened<XlsxState>>);

/// Shared state behind an [`Xlsx`] handle.
struct XlsxState {
    owner: XlsxOwner,
    /// Worksheet names in catalog order, built on the first generic-trait
    /// request rather than at open.
    names: OnceLock<Box<[String]>>,
    /// Core properties, read at open as the facade does.
    metadata: Metadata,
    /// Repairs of that core-properties read, folded into
    /// [`Xlsx::tolerance_report`].
    metadata_repairs: MetadataRepairs,
    /// Whether write verbs may build an editor ([`OpenOptions::retain_source`]).
    /// Both owner states already hold everything an edit needs, so this
    /// releases nothing; it only refuses.
    retain_source: bool,
    /// The OPC read limits admitted with this handle. Raw facade operations
    /// re-index published bytes under this same policy.
    ooxml_limits: litchi_opc::SourceLimits,
    /// A successfully promoted source-backed workbook. The lock is initialized
    /// lazily so read-only opens do not pay for an edit cache, and failed
    /// promotions are never retained.
    promotion: OnceLock<Mutex<Option<xlsx::Package>>>,
    /// A successfully unsigned workbook for a removal-authorized signature
    /// disposition. It is separate from `promotion`: the original promoted
    /// snapshot remains reusable by ordinary handles and signed aliases still
    /// refuse before either cache is consulted.
    unsigned_promotion: OnceLock<Mutex<Option<xlsx::Package>>>,
}

/// The two owner states of an XLSX handle (record 0003 §2.3).
pub(crate) enum XlsxOwner {
    /// Deferred catalog over the retained positional source.
    Source(xlsx::SourceBackedWorkbook),
    /// The owning package snapshot.
    Eager(xlsx::Package),
}

const _: () = {
    const fn assert_handle<T: Send + Sync + Clone>() {}
    assert_handle::<Xlsx>();
    assert!(size_of::<Xlsx>() == size_of::<usize>());
};

impl XlsxOwner {
    /// The facade's `ensure_source_current`: refuse a source-backed owner
    /// whose file changed since it was opened.
    pub(crate) fn ensure_current(&self) -> Result<()> {
        match self {
            Self::Source(workbook) => workbook.source_version().map(drop).map_err(Error::from),
            Self::Eager(_) => Ok(()),
        }
    }

    /// This owner's key for [`crate::path::fence::Scope`].
    pub(crate) fn fence_key(&self) -> usize {
        std::ptr::from_ref(self).addr()
    }

    /// Bracket `read` with freshness checks; a changed source wins over the
    /// read's own result. Inside a resolver or query step, whose bracket
    /// already checks around the whole step, `read` runs alone.
    pub(crate) fn finish<T>(&self, read: impl FnOnce() -> Result<T>) -> Result<T> {
        if crate::path::fence::is_active(self.fence_key()) {
            return read();
        }
        self.ensure_current()?;
        let result = read();
        self.ensure_current()?;
        result
    }

    fn len(&self) -> usize {
        match self {
            Self::Source(workbook) => workbook.len(),
            Self::Eager(package) => package.workbook_ref().len(),
        }
    }

    fn names(&self) -> Box<[String]> {
        match self {
            Self::Source(workbook) => workbook.sheet_names().map(str::to_owned).collect(),
            Self::Eager(package) => package
                .workbook_ref()
                .sheets()
                .map(|sheet| sheet.name().to_owned())
                .collect(),
        }
    }

    fn select<'a>(&self, selector: impl Into<WorksheetSelector<'a>>) -> Result<Option<Sheet>> {
        Ok(match self {
            Self::Source(workbook) => workbook.sheet(selector)?.map(Sheet::source),
            Self::Eager(package) => package.workbook_ref().sheet(selector)?.map(Sheet::owned),
        })
    }

    fn active(&self) -> Result<Option<Sheet>> {
        Ok(match self {
            Self::Source(workbook) => workbook.active_sheet()?.map(Sheet::source),
            Self::Eager(package) => package.workbook_ref().active_sheet().map(Sheet::owned),
        })
    }

    fn active_position(&self) -> Option<usize> {
        match self {
            Self::Source(workbook) => workbook.active_sheet_position(),
            Self::Eager(package) => package
                .workbook_ref()
                .active_sheet()
                .map(|sheet| sheet.position()),
        }
    }
}

impl Xlsx {
    /// Check the retained source before an operation that will not otherwise
    /// touch the owner. Write transactions use this at their outer fence so a
    /// stale source wins over target/property admission and native errors.
    pub(crate) fn check_source(&self) -> Result<()> {
        self.0.owner.ensure_current()
    }

    /// Check only retained-source freshness for a setter that may establish
    /// an exact signed no-op. Signature and encryption policy are applied by
    /// the native editor and the final byte-identity fence.
    pub(crate) fn check_source_for_replay(&self) -> Result<()> {
        self.0.owner.ensure_current()
    }

    /// Admit an edit without publishing one. The transaction coordinator uses
    /// the same signature, retention and leaf encryption policy as the write
    /// adapter before it begins a multi-step edit.
    pub(crate) fn admit_edit(&self) -> Result<()> {
        self.check_source()?;
        let workbook = self.editor()?;
        workbook
            .edit()
            .map(|_| ())
            .map_err(|error| Error::from(error).with_format(FileFormat::Xlsx))?;
        self.check_source()
    }

    /// Whether two handles are the same immutable open snapshot. A cheap clone
    /// shares its `Opened` allocation; a signature disposition gets a distinct
    /// handle and is intentionally not interchangeable for replay.
    pub(crate) fn same_snapshot(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0, &other.0)
    }

    /// Retained encryption provenance, inspected without decoding source data.
    #[cfg(feature = "encryption")]
    pub(crate) fn requires_reencryption(&self) -> bool {
        matches!(&self.0.owner, XlsxOwner::Eager(workbook) if workbook.encryption().is_some())
    }

    /// Whether this handle retains an edit source, without reading any payload.
    pub(crate) fn retains_source(&self) -> bool {
        self.0.retain_source
    }

    /// Adopt the source-backed package detection validated: read its core
    /// properties, then build the deferred catalog over it. No worksheet
    /// payload is read.
    ///
    /// Under [`SourceMode::Eager`] the package is materialized into the
    /// owning snapshot instead; the openers reach that only from
    /// `open_prepared`, because their other eager routes call
    /// [`Self::from_opc_package`].
    pub(crate) fn from_source_package(
        package: SourceBackedPackage,
        options: &OpenOptions,
    ) -> Result<Self> {
        Self::from_source_package_with_origin(package, options, None)
    }

    /// Adopt a source-backed package and retain the path which supplied it.
    /// The origin is metadata only; package classification still comes from
    /// the already-indexed source package.
    pub(crate) fn from_source_package_with_origin(
        package: SourceBackedPackage,
        options: &OpenOptions,
        origin: Option<Arc<Path>>,
    ) -> Result<Self> {
        let ooxml_limits = package.read_limits();
        let (metadata, metadata_repairs) =
            read_source_backed_metadata(&package).map_err(metadata_error)?;
        let owner = if options.source_mode == SourceMode::Eager {
            XlsxOwner::Eager(materialize(&package)?)
        } else {
            XlsxOwner::Source(xlsx::SourceBackedWorkbook::from_source_backed_package(
                package,
            )?)
        };
        Ok(
            Self::with_owner(owner, metadata, metadata_repairs, options, ooxml_limits)
                .with_origin(origin),
        )
    }

    /// Build the owning snapshot from the OPC package the byte detector
    /// already parsed (`DetectedFormat::Xlsx`).
    pub(crate) fn from_opc_package(package: OpcPackage, options: &OpenOptions) -> Result<Self> {
        Self::from_opc_package_with_origin(package, options, None)
    }

    /// Adopt an eager OPC package and retain the path which supplied it.
    pub(crate) fn from_opc_package_with_origin(
        package: OpcPackage,
        options: &OpenOptions,
        origin: Option<Arc<Path>>,
    ) -> Result<Self> {
        let ooxml_limits = package.read_limits();
        let mut metadata_repairs = MetadataRepairs::default();
        let metadata =
            litchi_ooxml_common::properties::read_with_report(&package, &mut metadata_repairs)
                .map_err(metadata_error)?
                .unwrap_or_default();
        let metadata = super::metadata::with_app_props(metadata, &package);
        let workbook = xlsx::Package::from_opc(package)?;
        Ok(Self::with_owner(
            XlsxOwner::Eager(workbook),
            metadata,
            metadata_repairs,
            options,
            ooxml_limits,
        )
        .with_origin(origin))
    }

    /// Adopt a newly created owning workbook without serializing and
    /// reopening its package.
    pub(crate) fn from_workbook(workbook: xlsx::Package, options: &OpenOptions) -> Result<Self> {
        Self::from_workbook_with_origin(workbook, options, None)
    }

    /// Adopt a newly created workbook and retain an optional path origin.
    pub(crate) fn from_workbook_with_origin(
        workbook: xlsx::Package,
        options: &OpenOptions,
        origin: Option<Arc<Path>>,
    ) -> Result<Self> {
        let ooxml_limits = workbook.source_limits();
        Ok(Self::with_owner(
            XlsxOwner::Eager(workbook),
            Metadata::default(),
            MetadataRepairs::default(),
            options,
            ooxml_limits,
        )
        .with_origin(origin))
    }

    /// Adopt an owning package another route built, reading its core
    /// properties now. [`crate::Locked::unlock`] uses this for a package the
    /// leaf decrypted: the snapshot keeps its encryption provenance.
    #[cfg(feature = "encryption")]
    pub(crate) fn from_package(package: xlsx::Package, options: &OpenOptions) -> Result<Self> {
        Self::from_package_with_origin(package, options, None)
    }

    /// Adopt an owning package with its optional path origin.
    #[cfg(feature = "encryption")]
    pub(crate) fn from_package_with_origin(
        package: xlsx::Package,
        options: &OpenOptions,
        origin: Option<Arc<Path>>,
    ) -> Result<Self> {
        let ooxml_limits = package.source_limits();
        let mut metadata_repairs = MetadataRepairs::default();
        let metadata = package
            .core_properties_with_report(&mut metadata_repairs)?
            .unwrap_or_default();
        let metadata = match package.app_props() {
            Ok(Some(app)) => {
                let mut metadata = metadata;
                app.project_into(&mut metadata);
                metadata
            },
            _ => metadata,
        };
        Ok(Self::with_owner(
            XlsxOwner::Eager(package),
            metadata,
            metadata_repairs,
            options,
            ooxml_limits,
        )
        .with_origin(origin))
    }

    fn with_owner(
        owner: XlsxOwner,
        metadata: Metadata,
        metadata_repairs: MetadataRepairs,
        options: &OpenOptions,
        ooxml_limits: litchi_opc::SourceLimits,
    ) -> Self {
        Self(Opened::new(XlsxState {
            owner,
            names: OnceLock::new(),
            metadata,
            metadata_repairs,
            retain_source: options.retain_source,
            ooxml_limits,
            promotion: OnceLock::new(),
            unsigned_promotion: OnceLock::new(),
        }))
    }

    /// The same workbook under another signature disposition (record 0013).
    pub(crate) fn with_signature(&self, signature: Signature) -> Self {
        Self(self.0.with_signature(signature))
    }

    /// The stable path origin captured by a path open, if any.
    pub(crate) fn origin(&self) -> Option<Arc<Path>> {
        self.0.origin()
    }

    /// Attach path metadata to a handle after a path route has finished
    /// classification. Byte and create routes leave this unset.
    pub(crate) fn with_origin(&self, origin: Option<Arc<Path>>) -> Self {
        Self(self.0.with_origin(origin))
    }

    /// Adopt a committed native snapshot without serializing and reopening its
    /// OPC package. The new handle carries the source retention policy,
    /// metadata repair sidecar and path origin of the input. A removal-
    /// authorized handle becomes ordinary unsigned state after its native
    /// editor has already removed the signature graph.
    pub(crate) fn adopt_workbook(&self, workbook: xlsx::Package) -> Self {
        let options = OpenOptions::new().retain_source(self.0.retain_source);
        let signature = match self.signature() {
            Signature::RemovalAuthorized => Signature::Unsigned,
            signature => signature,
        };
        Self::with_owner(
            XlsxOwner::Eager(workbook),
            self.0.metadata.clone(),
            self.0.metadata_repairs.clone(),
            &options,
            self.0.ooxml_limits,
        )
        .with_origin(self.origin())
        .with_signature(signature)
    }

    /// This handle's signature disposition (record 0013).
    pub(crate) fn signature(&self) -> Signature {
        self.0.signature()
    }

    /// The format of this handle.
    #[must_use]
    pub const fn format(&self) -> FileFormat {
        FileFormat::Xlsx
    }

    /// The OPC ingress policy retained by this handle for raw re-admission.
    pub(crate) fn ooxml_read_limits(&self) -> litchi_opc::SourceLimits {
        self.0.ooxml_limits
    }

    /// Whether the owner defers worksheet payloads to the retained source,
    /// for the tests that pin each route's owner.
    #[cfg(test)]
    pub(crate) fn is_source_backed(&self) -> bool {
        matches!(self.0.owner, XlsxOwner::Source(_))
    }

    /// Worksheet names in workbook order, every sheet kind included.
    ///
    /// # Errors
    ///
    /// Returns [`Code::IoError`] when a source-backed workbook's file changed.
    pub fn worksheet_names(&self) -> Result<Vec<String>> {
        self.0.owner.ensure_current()?;
        Ok(self.names().to_vec())
    }

    /// The number of sheets in the workbook catalog.
    ///
    /// # Errors
    ///
    /// Returns [`Code::IoError`] when a source-backed workbook's file changed.
    pub fn worksheet_count(&self) -> Result<usize> {
        self.0.owner.ensure_current()?;
        Ok(self.0.owner.len())
    }

    /// All cell text: each row of each sheet on its own line, cells
    /// separated by tabs. Non-grid sheets contribute nothing.
    ///
    /// # Errors
    ///
    /// Returns the XLSX error of the first worksheet that fails to read, or
    /// [`Code::IoError`] when a source-backed workbook's file changed.
    pub fn text(&self) -> Result<String> {
        let owner = &self.0.owner;
        owner.finish(|| {
            let mut out = String::new();
            for index in 0..owner.len() {
                let sheet = owner
                    .select(index)
                    .and_then(|sheet| sheet.ok_or_else(|| sheet_index_out_of_bounds(index)))?;
                for row in 0..sheet.row_end()? {
                    cell_text::append_row(&mut out, &sheet.row_values(row as u32)?);
                }
            }
            Ok(out)
        })
    }

    /// The core document properties read at open. Always `Some` when it
    /// succeeds; a package without core properties reads as empty metadata,
    /// as the facade reports it.
    ///
    /// # Errors
    ///
    /// Returns [`Code::IoError`] when a source-backed workbook's file changed.
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        self.0.owner.ensure_current()?;
        Ok(Some(self.0.metadata.clone()))
    }

    /// Everything the reader repaired so far: the open (package, workbook
    /// part, styles, theme), the shared strings once read, every worksheet
    /// whose cells have been parsed, and the core-properties read at open.
    ///
    /// # Errors
    ///
    /// Returns [`Code::IoError`] when a source-backed workbook's file changed.
    pub fn tolerance_report(&self) -> Result<xlsx::ToleranceReport> {
        self.0.owner.ensure_current()?;
        let mut report = match &self.0.owner {
            XlsxOwner::Source(workbook) => workbook.tolerance_report(),
            XlsxOwner::Eager(package) => package.workbook_ref().tolerance_report(),
        };
        report.append_properties(&self.0.metadata_repairs);
        Ok(report)
    }

    /// Select one worksheet by name or zero-based position.
    ///
    /// Names match case-insensitively, as XLSX defines them; `"dAtA"` selects
    /// `Data`. Selecting is catalog-only on a source-backed workbook: cell and
    /// range reads are deferred until requested through the returned
    /// [`SelectedWorksheet`]. A missing name or position returns `Ok(None)`.
    ///
    /// # Errors
    ///
    /// Returns [`Code::IoError`] when a source-backed workbook's file changed,
    /// and the XLSX error for an unsupported selector form.
    pub fn sheet<'a>(
        &self,
        selector: impl Into<WorksheetSelector<'a>>,
    ) -> Result<Option<SelectedWorksheet>> {
        let selector = selector.into();
        self.0.owner.finish(|| {
            Ok(match &self.0.owner {
                XlsxOwner::Source(workbook) => workbook
                    .sheet(selector)?
                    .map(SelectedWorksheet::from_source),
                XlsxOwner::Eager(package) => package
                    .workbook_ref()
                    .sheet(selector)?
                    .map(SelectedWorksheet::from_owned),
            })
        })
    }

    /// Build an immutable native workbook snapshot for the write kernel (M2).
    ///
    /// Source-backed handles promote through their already-indexed package and
    /// memoize only a successful native snapshot in the shared open state.
    /// Cloned handles and selected worksheets therefore reuse the same native
    /// snapshot without consuming the facade handle or re-indexing the source.
    /// Eager handles clone their existing snapshot. A removal-authorized
    /// signature disposition always takes the native unsigned path so the
    /// signature graph is removed explicitly.
    /// The returned native snapshot is detached; a unified write must also
    /// fence its complete admission, edit and commit with `XlsxOwner::finish`.
    ///
    /// # Errors
    ///
    /// Returns [`Code::DocumentProtected`] while the workbook is still
    /// signed, with a suggestion naming `Signed::unsign`;
    /// [`Code::UnsupportedCapability`] when the handle was opened with
    /// [`OpenOptions::retain_source`] `false`; [`Code::IoError`] when the
    /// file changed; and the XLSX error of a failed promotion.
    pub(crate) fn editor(&self) -> Result<xlsx::Package> {
        if self.0.signature() == Signature::Signed {
            return Err(signature::refusal(FileFormat::Xlsx));
        }
        self.editor_after_source_check(false)
    }

    /// Build a native editor for a setter that may prove an exact signed
    /// no-op. Changed commits are rejected by the setter/commit fence; the
    /// signed package is kept intact while the native edit is staged.
    pub(crate) fn editor_for_write(&self) -> Result<xlsx::Package> {
        if self.0.signature() == Signature::Signed && !self.0.retain_source {
            return Err(source_released());
        }
        self.editor_after_source_check(self.0.signature() == Signature::Signed)
    }

    /// Admit the single source-backed endpoint needed by the metadata route
    /// before promotion. Metadata's common editor performs the exact
    /// source/target/patch admission after promotion; this guard only keeps a
    /// deferred source from being promoted when that one endpoint already
    /// exceeds the caller's remaining snapshot budget.
    pub(crate) fn editor_for_metadata_with_snapshot_limit(
        &self,
        maximum: usize,
    ) -> Result<xlsx::Package> {
        if !self.0.retain_source {
            return Err(source_released());
        }
        if let XlsxOwner::Source(workbook) = &self.0.owner {
            self.check_source_for_replay()?;
            let observed = workbook.retained_bytes_for_edit()?;
            if observed > maximum {
                return Err(crate::transaction::Error::Limit {
                    kind: crate::transaction::LimitKind::SnapshotBytes,
                    observed,
                    maximum,
                }
                .into());
            }
        }
        self.editor_for_write()
    }

    /// Admit the native promotion against the transaction's remaining
    /// snapshot budget before it can copy the source archive or build the
    /// owning workbook graph. Eager owners can use the native endpoint
    /// census directly; deferred owners use the source archive plus the
    /// fixed owning-package/workbook handoff that promotion must retain.
    pub(crate) fn editor_with_snapshot_limit(&self, maximum: usize) -> Result<xlsx::Package> {
        self.preflight_editor_snapshot(maximum)?;
        self.editor()
    }

    /// Signed/removal-authorized variant of [`Self::editor_with_snapshot_limit`].
    pub(crate) fn editor_for_write_with_snapshot_limit(
        &self,
        maximum: usize,
    ) -> Result<xlsx::Package> {
        self.preflight_editor_snapshot(maximum)?;
        self.editor_for_write()
    }

    fn preflight_editor_snapshot(&self, maximum: usize) -> Result<()> {
        let endpoint = match &self.0.owner {
            XlsxOwner::Eager(package) => package.retained_bytes(),
            XlsxOwner::Source(workbook) => workbook.retained_bytes_for_edit()?,
        };
        // This mirrors the native workbook editor's CaptureBudget::new:
        // promotion reserves the two immutable endpoint owners. Semantic
        // worksheet/feature images and patch payloads are admitted by the
        // family edit before they are materialized; transient editor and
        // container headers are not retained snapshot data.
        let observed = endpoint.saturating_mul(2);
        if observed > maximum {
            return Err(crate::transaction::Error::Limit {
                kind: crate::transaction::LimitKind::SnapshotBytes,
                observed,
                maximum,
            }
            .into());
        }
        Ok(())
    }

    /// Materialize the original native package for a canonical signed no-op
    /// patch. `editor_for_write` intentionally calls `unsigned_workbook` so a
    /// changed publication cannot preserve signature parts; the identity
    /// patch itself must still be anchored to the signed package bytes.
    pub(crate) fn original_package_for_noop(&self) -> Result<xlsx::Package> {
        if !self.0.retain_source {
            return Err(source_released());
        }
        self.check_source_for_replay()?;
        match &self.0.owner {
            XlsxOwner::Source(workbook) => self.promote_source(workbook),
            XlsxOwner::Eager(package) => Ok(package.clone()),
        }
    }

    fn editor_after_source_check(&self, allow_signed: bool) -> Result<xlsx::Package> {
        if !self.0.retain_source {
            return Err(source_released());
        }
        if allow_signed {
            self.check_source_for_replay()?;
        } else {
            self.check_source()?;
        }
        match self.0.signature() {
            Signature::RemovalAuthorized => self.unsigned_workbook(),
            Signature::Unsigned => match &self.0.owner {
                XlsxOwner::Source(workbook) => self.promote_source(workbook),
                XlsxOwner::Eager(workbook) => Ok(workbook.clone()),
            },
            Signature::Signed if allow_signed => self.unsigned_workbook(),
            Signature::Signed => unreachable!("signed handles return before source fencing"),
        }
    }

    /// Promote a source-backed owner once and return a native snapshot clone.
    /// The conversion and cache publication are source-fenced independently;
    /// the native OPC conversion also keeps its own checks for each stage.
    fn promote_source(&self, source: &xlsx::SourceBackedWorkbook) -> Result<xlsx::Package> {
        let cache = self.0.promotion.get_or_init(|| Mutex::new(None));
        let mut cached = cache
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        if let Some(workbook) = cached.as_ref() {
            return self.0.owner.finish(|| Ok(workbook.clone()));
        }
        let workbook = self
            .0
            .owner
            .finish(|| materialize(source.source_backed_package()?))?;
        *cached = Some(workbook.clone());
        if let Err(error) = self.0.owner.ensure_current() {
            *cached = None;
            return Err(error);
        }
        Ok(workbook)
    }

    /// An owning copy of the workbook whose signature the leaf removed
    /// explicitly ([`xlsx::Package::unsign`]): promoted through the
    /// source-preserving OPC conversion for a source-backed owner, copied from
    /// the eager snapshot otherwise. An unlocked encrypted snapshot refuses,
    /// as its edits do.
    fn unsigned_workbook(&self) -> Result<xlsx::Package> {
        let cache = self.0.unsigned_promotion.get_or_init(|| Mutex::new(None));
        let mut cached = cache
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        if let Some(workbook) = cached.as_ref() {
            return match &self.0.owner {
                XlsxOwner::Source(_) => self.0.owner.finish(|| Ok(workbook.clone())),
                XlsxOwner::Eager(_) => Ok(workbook.clone()),
            };
        }
        let workbook = match &self.0.owner {
            XlsxOwner::Source(workbook) => self.promote_source(workbook)?,
            XlsxOwner::Eager(workbook) => workbook.clone(),
        };
        let workbook = workbook.unsign()?;
        *cached = Some(workbook.clone());
        if let Err(error) = self.0.owner.ensure_current() {
            *cached = None;
            return Err(error);
        }
        Ok(workbook)
    }

    fn names(&self) -> &[String] {
        self.0.names.get_or_init(|| self.0.owner.names())
    }

    fn sheet_at(&self, index: usize) -> Result<Sheet> {
        self.0.owner.finish(|| {
            self.0
                .owner
                .select(index)?
                .ok_or_else(|| sheet_index_out_of_bounds(index))
        })
    }
}

/// Build the owning snapshot from a source-backed package through its
/// existing index.
fn materialize(package: &SourceBackedPackage) -> Result<xlsx::Package> {
    let package = package
        .to_opc_package_preserving_source()
        .map_err(|error| Error::from(error).with_format(FileFormat::Xlsx))?;
    Ok(xlsx::Package::from_opc(package)?)
}

/// Classify a core-properties read failure. OPC failures keep their own
/// class (a changed source is [`Code::IoError`]); other failures are XLSX
/// package errors.
fn metadata_error(error: litchi_ooxml_common::Error) -> Error {
    match error {
        litchi_ooxml_common::Error::Opc(error) => Error::from(error).with_format(FileFormat::Xlsx),
        error => Error::from(xlsx::Error::from(error)),
    }
}

fn sheet_index_out_of_bounds(index: usize) -> Error {
    Error::new(
        Code::OutOfRange,
        format!("XLSX worksheet index {index} is out of bounds"),
    )
    .with_format(FileFormat::Xlsx)
}

fn source_released() -> Error {
    Error::new(
        Code::UnsupportedCapability,
        "the workbook was opened without retaining its source for editing",
    )
    .with_format(FileFormat::Xlsx)
    .with_suggestion("reopen with `OpenOptions::retain_source(true)` to edit")
}

impl std::fmt::Debug for Xlsx {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Xlsx")
            .field("worksheet_count", &self.0.owner.len())
            .field(
                "source_backed",
                &matches!(self.0.owner, XlsxOwner::Source(_)),
            )
            .finish_non_exhaustive()
    }
}

impl WorkbookTrait for Xlsx {
    type Error = Error;
    type Worksheet<'a> = crate::view::sheet::Worksheet<'a>;
    type Worksheets<'a> = crate::view::sheet::Worksheets<'a>;

    fn active_worksheet(&self) -> Result<Self::Worksheet<'_>> {
        let sheet = self.0.owner.finish(|| {
            self.0.owner.active()?.ok_or_else(|| {
                Error::new(Code::NotFound, "XLSX workbook has no active worksheet")
                    .with_format(FileFormat::Xlsx)
            })
        })?;
        Ok(crate::view::sheet::Worksheet::xlsx(sheet))
    }

    fn worksheet_names(&self) -> &[String] {
        self.names()
    }

    fn worksheet_by_name(&self, name: &str) -> Result<Self::Worksheet<'_>> {
        let sheet = self.0.owner.finish(|| {
            self.0.owner.select(name)?.ok_or_else(|| {
                Error::new(Code::NotFound, "XLSX worksheet was not found")
                    .with_format(FileFormat::Xlsx)
            })
        })?;
        Ok(crate::view::sheet::Worksheet::xlsx(sheet))
    }

    fn worksheet_by_index(&self, index: usize) -> Result<Self::Worksheet<'_>> {
        Ok(crate::view::sheet::Worksheet::xlsx(self.sheet_at(index)?))
    }

    fn worksheets(&self) -> Self::Worksheets<'_> {
        crate::view::sheet::Worksheets::new(crate::view::sheet::WorksheetsInner::Xlsx(Worksheets {
            workbook: self,
            index: 0,
            done: false,
        }))
    }

    fn worksheet_count(&self) -> usize {
        self.0.owner.len()
    }

    fn active_sheet_index(&self) -> usize {
        self.0.owner.active_position().unwrap_or(0)
    }

    fn is_1904_date_system(&self) -> bool {
        let date_system = match &self.0.owner {
            XlsxOwner::Source(workbook) => workbook.date_system(),
            XlsxOwner::Eager(workbook) => workbook.workbook_ref().date_system(),
        };
        matches!(date_system, xlsx::DateSystem::Excel1904)
    }
}

/// The worksheets of an [`Xlsx`] handle, in catalog order.
pub(crate) struct Worksheets<'a> {
    workbook: &'a Xlsx,
    index: usize,
    done: bool,
}

impl Iterator for Worksheets<'_> {
    type Item = Result<Sheet>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        let result = self.workbook.0.owner.finish(|| {
            if self.index >= self.workbook.0.owner.len() {
                return Ok(None);
            }
            self.workbook.0.owner.select(self.index)
        });
        match result {
            Ok(Some(sheet)) => {
                self.index += 1;
                Some(Ok(sheet))
            },
            Ok(None) => {
                self.done = true;
                None
            },
            Err(error) => {
                self.done = true;
                Some(Err(error))
            },
        }
    }
}

#[cfg(test)]
mod tests;

#[cfg(test)]
mod unified_tests;
