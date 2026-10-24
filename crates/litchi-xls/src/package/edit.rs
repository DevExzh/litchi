//! The primary immutable XLS package transaction.
//!
//! `Package::edit` is the one package-owned write entry point for legacy XLS.
//! It delegates cell, resource, and worksheet-structure changes to the
//! checked [`crate::cell_values`] owner and adapts that owner to the canonical
//! `Edit`/`Commit`/`Patch` triad.  The package passed to an edit is never
//! mutated; dropping the edit is rollback.
//!
//! A small end-to-end example is enough to show the ownership boundary:
//!
//! ```
//! use litchi_xls::cell_values::{Reference, Value};
//! use litchi_xls::package::edit::Result;
//! use litchi_xls::Package;
//!
//! fn change_first_cell(bytes: &[u8]) -> Result<Vec<u8>> {
//!     let source = Package::from_bytes(bytes).map_err(Error::Package)?;
//!     let mut edit = source.edit()?;
//!     edit.set_value(0usize, Reference::new(0, 0).map_err(Error::CellValues)?,
//!         Value::Number(2.0))?;
//!     let commit = edit.commit()?;
//!     commit.into_package().to_bytes().map_err(Error::Package)
//! }
//! # use litchi_xls::package::edit::Error;
//! ```

use super::{Error as PackageError, Package, PackageState};
use crate::cell_values;
use crate::{OpenOptions, SourceBackedError};
use litchi_cfb::{OleFileLimits, SharedOleFile};
use litchi_core::SourceVersion;
use litchi_ole_common::object::Limits as ObjectLimits;
use std::fmt;
use std::sync::{Arc, OnceLock};

pub use crate::error::LimitKind;

/// Explicit bounds for one primary XLS edit.
///
/// The default profile is the leaf editor's checked default.  `Package`
/// intersects these bounds with the profile it retained at ingress, so an
/// edit cannot widen the source's admission policy by passing a larger value.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct EditOptions {
    object_limits: ObjectLimits,
    cfb_limits: OleFileLimits,
}

impl Default for EditOptions {
    fn default() -> Self {
        Self::new()
    }
}

impl EditOptions {
    /// Construct the checked default edit profile.
    #[must_use]
    pub fn new() -> Self {
        Self {
            object_limits: ObjectLimits::default(),
            cfb_limits: OleFileLimits::default(),
        }
    }

    /// Select the editable object-capture bounds.
    #[must_use]
    pub const fn object_limits(mut self, limits: ObjectLimits) -> Self {
        self.object_limits = limits;
        self
    }

    /// Select the low-level CFB admission bounds.
    #[must_use]
    pub const fn cfb_limits(mut self, limits: OleFileLimits) -> Self {
        self.cfb_limits = limits;
        self
    }

    /// Return the requested object-capture bounds.
    #[must_use]
    pub const fn object_limits_value(&self) -> ObjectLimits {
        self.object_limits
    }

    /// Return the requested CFB admission bounds.
    #[must_use]
    pub const fn cfb_limits_value(&self) -> OleFileLimits {
        self.cfb_limits
    }
}

/// Typed errors for the primary XLS package transaction.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// The immutable package owner rejected ingress, publication, or a source
    /// freshness check.
    Package(PackageError),
    /// The checked cell-values owner rejected a semantic operation.
    CellValues(crate::Error),
    /// A finite transaction or leaf admission bound was exceeded.
    Limit {
        /// Dimension whose bound was exceeded.
        kind: LimitKind,
        /// Observed amount.
        observed: usize,
        /// Configured maximum.
        maximum: usize,
    },
    /// A patch's exact source image does not match the package receiving it.
    PatchConflict,
}

impl From<PackageError> for Error {
    fn from(error: PackageError) -> Self {
        Self::Package(error)
    }
}

impl From<crate::Error> for Error {
    fn from(error: crate::Error) -> Self {
        // The leaf exposes this variant as a typed bound.  Keep the mapping
        // here so the package family has the canonical Limit shape while all
        // other leaf refusals remain available as typed sources.
        match error {
            crate::Error::Limit {
                kind: LimitKind::Operations,
                observed,
                maximum,
            } => Self::Limit {
                kind: LimitKind::Operations,
                observed: usize::try_from(observed).unwrap_or(usize::MAX),
                maximum: usize::try_from(maximum).unwrap_or(usize::MAX),
            },
            other => Self::CellValues(other),
        }
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package(error) => error.fmt(formatter),
            Self::CellValues(error) => error.fmt(formatter),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => write!(
                formatter,
                "XLS edit limit {kind:?}: {observed} exceeds {maximum}"
            ),
            Self::PatchConflict => formatter.write_str(
                "XLS package patch source does not match the exact package receiving it",
            ),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Package(error) => Some(error),
            Self::CellValues(error) => Some(error),
            Self::Limit { .. } | Self::PatchConflict => None,
        }
    }
}

impl Error {
    /// Return the exceeded limit dimension, if this is a limit failure.
    #[must_use]
    pub const fn limit_kind(&self) -> Option<LimitKind> {
        match self {
            Self::Limit { kind, .. } => Some(*kind),
            Self::Package(_) | Self::CellValues(_) | Self::PatchConflict => None,
        }
    }

    /// Return the observed and maximum values, if this is a limit failure.
    #[must_use]
    pub const fn limit_values(&self) -> Option<(usize, usize)> {
        match self {
            Self::Limit {
                observed, maximum, ..
            } => Some((*observed, *maximum)),
            Self::Package(_) | Self::CellValues(_) | Self::PatchConflict => None,
        }
    }
}

/// Result type for the canonical XLS primary transaction.
pub type Result<T> = std::result::Result<T, Error>;

/// Content-free facts about one primary XLS publication.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    operation_count: usize,
    changed_cells: usize,
    changed_number_fields: usize,
    touched_streams: usize,
}

impl Diagnostics {
    const fn applied(changed: bool, operation_count: usize) -> Self {
        Self {
            changed,
            operation_count,
            changed_cells: 0,
            changed_number_fields: 0,
            touched_streams: 0,
        }
    }

    fn from_leaf(value: cell_values::Diagnostics, operation_count: usize) -> Self {
        Self {
            changed: value.changed(),
            operation_count,
            changed_cells: value.changed_cells(),
            changed_number_fields: value.changed_number_fields(),
            touched_streams: value.touched_streams(),
        }
    }

    /// Whether exact package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of semantic operations retained by the patch.
    #[must_use]
    pub const fn operation_count(self) -> usize {
        self.operation_count
    }

    /// Number of semantically changed cells.
    #[must_use]
    pub const fn changed_cells(self) -> usize {
        self.changed_cells
    }

    /// Number of standalone `Number` fields changed.
    #[must_use]
    pub const fn changed_number_fields(self) -> usize {
        self.changed_number_fields
    }

    /// Number of CFB streams in the mutation closure.
    #[must_use]
    pub const fn touched_streams(self) -> usize {
        self.touched_streams
    }
}

/// Exact, reversible package patch produced by a primary edit.
#[derive(Clone)]
pub struct Patch {
    inner: cell_values::Patch,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("before_bytes", &self.before().len())
            .field("after_bytes", &self.after().len())
            .field("operation_count", &self.operation_count())
            .field("is_noop", &self.is_noop())
            .finish()
    }
}

impl Patch {
    fn from_leaf(inner: cell_values::Patch) -> Self {
        Self { inner }
    }

    /// Whether exact package bytes are identical.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.inner.is_noop()
    }

    /// Exact source image required before applying this patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        self.inner.before()
    }

    /// Exact target image produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        self.inner.after()
    }

    /// Number of semantic operations retained by this patch.
    #[must_use]
    pub fn operation_count(&self) -> usize {
        self.inner.semantic().len()
    }

    /// Returns the saturating byte charge for this exact patch.
    ///
    /// The charge includes both complete artifact endpoints and the forward
    /// and inverse semantic data retained by the native cell-values patch.
    /// It walks retained data directly without serializing or cloning the
    /// inverse patch.
    #[must_use]
    pub fn retained_bytes(&self) -> usize {
        self.inner.retained_bytes()
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self::from_leaf(self.inner.inverse())
    }

    /// The durable, reversible semantic patch: deterministic JSON
    /// ([`cell_values::SemanticPatch::to_deterministic_json`]), replayed with
    /// [`Package::apply_semantic`], merged with [`Package::plan_three_way`]
    /// and transferred with [`Package::plan_transfer`].
    #[must_use]
    pub fn semantic(&self) -> &cell_values::SemanticPatch {
        self.inner.semantic()
    }

    pub(crate) fn leaf(&self) -> &cell_values::Patch {
        &self.inner
    }
}

/// Successful immutable primary XLS publication.
#[derive(Clone)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl fmt::Debug for Commit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Commit")
            .field("package", &self.package)
            .field("patch", &self.patch)
            .field("diagnostics", &self.diagnostics)
            .finish()
    }
}

impl Commit {
    /// Published immutable package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Content-free publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Consume the commit into its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit into its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
    /// Record this commit's package in `history` as its next state, weighted
    /// by the patch's before and after artifact lengths.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the history's current package
    /// is not this commit's source, and a refusal when the history cannot
    /// hold it.
    pub fn record_in(self, history: &mut crate::package::History) -> Result<Vec<Package>> {
        if history
            .current()
            .to_bytes()
            .map_err(Error::Package)?
            .as_slice()
            != self.patch.before()
        {
            return Err(Error::PatchConflict);
        }
        let weight = u64::try_from(self.patch.before().len())
            .ok()
            .zip(u64::try_from(self.patch.after().len()).ok())
            .and_then(|(before, after)| before.checked_add(after))
            .ok_or(Error::CellValues(crate::Error::Internal(
                "XLS history weight overflow",
            )))?;
        history.record(self.package, weight).map_err(|_error| {
            Error::CellValues(crate::Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "edit cells",
                "the cell edit history cannot record this commit",
            ))
        })
    }
}

/// Detached primary XLS edit backed by the checked cell-values snapshot.
pub struct Edit {
    package: Package,
    inner: Option<cell_values::Edit>,
    lazy: LazySnapshot,
    options: EditOptions,
    source_guard: Option<SourceGuard>,
}

impl fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("package", &self.package)
            .field("options", &self.options)
            .field("materialized", &self.inner.is_some())
            .field("source_backed", &self.source_guard.is_some())
            .finish_non_exhaustive()
    }
}

#[derive(Clone)]
struct SourceGuard {
    source: Arc<SharedOleFile>,
    expected: SourceVersion,
    /// A deferred or handed-off source owner retains the original dynamic
    /// source separately from its composed CFB target.  Keep that owner in
    /// the primary package guard so a mutation is promoted to the typed
    /// `SourceChanged` result before the composed target can report I/O.
    original: Option<crate::SourceBackedWorkbook>,
}

/// Demand-driven semantic snapshot for the primary edit.
///
/// Package admission already retains the source owner and its policies, so
/// constructing an edit only clones those handles.  BIFF validation and the
/// cell indexes are built on the first source lookup, staged verb, or commit;
/// a successful snapshot is reused by all later operations.
struct LazySnapshot {
    package: Package,
    options: EditOptions,
    snapshot: OnceLock<cell_values::Snapshot>,
}

impl LazySnapshot {
    fn new(package: Package, options: EditOptions) -> Self {
        Self {
            package,
            options,
            snapshot: OnceLock::new(),
        }
    }

    fn get(&self) -> Result<&cell_values::Snapshot> {
        if let Some(snapshot) = self.snapshot.get() {
            return Ok(snapshot);
        }
        let snapshot = self
            .package
            .edit_snapshot_with_options(self.options)
            .map_err(Error::Package)?;
        let _ = self.snapshot.set(snapshot);
        self.snapshot
            .get()
            .ok_or(Error::Package(PackageError::InvalidData(
                "XLS edit snapshot was not retained",
            )))
    }

    fn into_edit(self) -> Result<cell_values::Edit> {
        if let Some(snapshot) = self.snapshot.into_inner() {
            return Ok(snapshot.edit());
        }
        self.package
            .edit_snapshot_with_options(self.options)
            .map_err(Error::Package)
            .map(|snapshot| snapshot.edit())
    }
}

impl Edit {
    pub(crate) fn from_package_with_options(
        package: Package,
        options: EditOptions,
    ) -> Result<Self> {
        super::ensure_publication_allowed(&package.inner).map_err(Error::Package)?;
        let source_guard = capture_source_guard(&package)?;
        check_source_guard(source_guard.as_ref())?;
        Ok(Self {
            lazy: LazySnapshot::new(package.clone(), options),
            package,
            inner: None,
            options,
            source_guard,
        })
    }

    /// A primary edit of `package` over a leaf edit already staged from its
    /// cell-value snapshot (a semantic replay or transfer).
    pub(crate) fn from_leaf(
        package: Package,
        inner: cell_values::Edit,
        options: EditOptions,
    ) -> Result<Self> {
        let source_guard = capture_source_guard(&package)?;
        Ok(Self {
            lazy: LazySnapshot::new(package.clone(), options),
            package,
            inner: Some(inner),
            options,
            source_guard,
        })
    }

    fn inner_mut(&mut self) -> Result<&mut cell_values::Edit> {
        if self.inner.is_none() {
            let inner = self.lazy.get()?.edit();
            self.inner = Some(inner);
        }
        self.inner
            .as_mut()
            .ok_or(Error::Package(PackageError::InvalidData(
                "XLS edit engine was not retained",
            )))
    }

    /// Join another edit of the same source package when the two are provably
    /// disjoint; the joined edit commits both atomically.
    ///
    /// # Errors
    ///
    /// Returns the structured join refusal with the rejected edit's
    /// conflicts.
    pub fn join(
        &mut self,
        incoming: Self,
    ) -> std::result::Result<&mut Self, cell_values::JoinError> {
        let incoming = match incoming.inner {
            Some(inner) => inner,
            None => incoming
                .lazy
                .into_edit()
                .map_err(|_| cell_values::JoinError::Allocation)?,
        };
        self.inner_mut()
            .map_err(|_| cell_values::JoinError::Allocation)?
            .join(incoming)?;
        Ok(self)
    }

    /// Immutable package that authorized this edit.
    #[must_use]
    pub const fn source(&self) -> &Package {
        &self.package
    }

    /// Requested edit profile. The package owner intersects it with the
    /// retained ingress profile before constructing the leaf snapshot.
    #[must_use]
    pub const fn options(&self) -> EditOptions {
        self.options
    }

    /// Look up a transaction-owned cell in the immutable source of this edit.
    ///
    /// This reuses the inventory admitted when the edit started. It reports
    /// the original cell, even after that coordinate has been staged for a
    /// change; commit before looking up a changed candidate. Source-backed
    /// edits check freshness before and after the lookup.
    ///
    /// ```
    /// use litchi_xls::{Package, cell_values::{Reference, Value}};
    /// let package = Package::new()?;
    /// let mut edit = package.edit()?;
    /// let cell = Reference::new(0, 0)?;
    /// assert!(edit.source_cell(0usize, cell)?.is_none());
    /// edit.insert_cell(0usize, cell, Value::Number(1.0))?;
    /// assert!(edit.source_cell(0usize, cell)?.is_none());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns a source-freshness error, a missing/ambiguous worksheet error,
    /// or an ambiguity error for duplicate physical cell records.
    pub fn source_cell<'a>(
        &self,
        selector: impl Into<cell_values::Selector<'a>>,
        reference: cell_values::Reference,
    ) -> Result<Option<&cell_values::Cell>> {
        check_source_guard(self.source_guard.as_ref())?;
        let result = self
            .lazy
            .get()?
            .source_cell(selector.into(), reference)
            .map_err(Error::from);
        check_source_guard(self.source_guard.as_ref())?;
        result
    }

    /// Set one existing numeric cell while preserving its storage family.
    pub fn set_numeric<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        reference: cell_values::Reference,
        value: f64,
    ) -> Result<()> {
        self.inner_mut()?
            .set_numeric(selector.into(), reference, value)
            .map_err(Error::from)
    }

    /// Set one existing standalone BIFF8 `Number` cell.
    pub fn set_number<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        reference: cell_values::Reference,
        value: f64,
    ) -> Result<()> {
        self.inner_mut()?
            .set_number(selector.into(), reference, value)
            .map_err(Error::from)
    }

    /// Copy a bounded scalar rectangle from another immutable XLS package.
    pub fn copy_scalar_cells_from<'source, 'target>(
        &mut self,
        donor: &Package,
        source_selector: impl Into<cell_values::Selector<'source>>,
        source_range: cell_values::CellRange,
        target_selector: impl Into<cell_values::Selector<'target>>,
        target_anchor: cell_values::Reference,
    ) -> Result<Option<&mut Self>> {
        let guard = capture_source_guard(donor)?;
        let options = donor.retained_edit_options()?;
        let donor = donor.edit_snapshot_with_options(options)?;
        let mut candidate = self.inner_mut()?.clone();
        let copied = candidate
            .copy_scalar_cells_from(
                &donor,
                source_selector,
                source_range,
                target_selector,
                target_anchor,
            )
            .map_err(Error::from)?
            .is_some();
        check_source_guard(guard.as_ref())?;
        if copied {
            self.inner = Some(candidate);
        }
        Ok(copied.then_some(self))
    }

    /// Copy one scalar cell from another immutable XLS package.
    pub fn copy_scalar_cell_from<'source, 'target>(
        &mut self,
        donor: &Package,
        source_selector: impl Into<cell_values::Selector<'source>>,
        source_reference: cell_values::Reference,
        target_selector: impl Into<cell_values::Selector<'target>>,
        target_reference: cell_values::Reference,
    ) -> Result<Option<&mut Self>> {
        let guard = capture_source_guard(donor)?;
        let options = donor.retained_edit_options()?;
        let donor = donor.edit_snapshot_with_options(options)?;
        let mut candidate = self.inner_mut()?.clone();
        let copied = candidate
            .copy_scalar_cell_from(
                &donor,
                source_selector,
                source_reference,
                target_selector,
                target_reference,
            )
            .map_err(Error::from)?
            .is_some();
        check_source_guard(guard.as_ref())?;
        if copied {
            self.inner = Some(candidate);
        }
        Ok(copied.then_some(self))
    }

    /// Set a typed value on an existing cell.
    pub fn set_value<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        reference: cell_values::Reference,
        value: cell_values::Value,
    ) -> Result<()> {
        self.inner_mut()?
            .set_value(selector.into(), reference, value)
            .map_err(Error::from)
    }

    /// Clear an existing formatting-only blank cell.
    pub fn clear_cell<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        reference: cell_values::Reference,
    ) -> Result<()> {
        self.inner_mut()?
            .clear_cell(selector.into(), reference)
            .map_err(Error::from)
    }

    /// Insert one scalar cell with the workbook's canonical default style.
    pub fn insert_cell<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        reference: cell_values::Reference,
        value: cell_values::Value,
    ) -> Result<()> {
        self.inner_mut()?
            .insert_cell(selector.into(), reference, value)
            .map_err(Error::from)
    }

    /// Insert one scalar cell with an existing workbook style resource.
    pub fn insert_cell_with_style<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        reference: cell_values::Reference,
        value: cell_values::Value,
        style: cell_values::StyleIndex,
    ) -> Result<()> {
        self.inner_mut()?
            .insert_cell_with_style(selector.into(), reference, value, style)
            .map_err(Error::from)
    }

    /// Insert one rich shared-string cell.
    pub fn insert_rich_text_cell<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        reference: cell_values::Reference,
        text: String,
        formatting_runs: Vec<crate::records::SharedStringFormatRun>,
    ) -> Result<()> {
        self.inner_mut()?
            .insert_rich_text_cell(selector.into(), reference, text, formatting_runs)
            .map_err(Error::from)
    }

    /// Insert one formula cell using the workbook's default style.
    pub fn insert_formula<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        reference: cell_values::Reference,
        expression: &str,
    ) -> Result<()> {
        self.inner_mut()?
            .insert_formula(selector.into(), reference, expression)
            .map_err(Error::from)
    }

    /// Insert one formula cell using an existing workbook style resource.
    pub fn insert_formula_with_style<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        reference: cell_values::Reference,
        expression: &str,
        style: cell_values::StyleIndex,
    ) -> Result<()> {
        self.inner_mut()?
            .insert_formula_with_style(selector.into(), reference, expression, style)
            .map_err(Error::from)
    }

    /// Write one formula cell, replacing an existing scalar cell in place.
    ///
    /// # Errors
    ///
    /// Returns a tokenizer, token-size, existing-formula, style, dependency
    /// or transaction-bound refusal without partially staging.
    pub fn set_formula<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        reference: cell_values::Reference,
        expression: &str,
    ) -> Result<()> {
        self.inner_mut()?
            .set_formula(selector.into(), reference, expression)
            .map_err(Error::from)
    }

    /// Set one row's height in twips and mark it custom.
    ///
    /// # Errors
    ///
    /// Returns a typed range, selector, axis-overlap or bound refusal.
    pub fn set_row_height<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        row: u16,
        twips: u16,
    ) -> Result<()> {
        self.inner_mut()?
            .set_row_height(selector.into(), row, twips)
            .map_err(Error::from)
    }

    /// Hide or show one row.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, axis-overlap or bound refusal.
    pub fn set_row_hidden<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        row: u16,
        hidden: bool,
    ) -> Result<()> {
        self.inner_mut()?
            .set_row_hidden(selector.into(), row, hidden)
            .map_err(Error::from)
    }

    /// Set one column's width in 1/256 character.
    ///
    /// # Errors
    ///
    /// Returns a typed range, selector, record-layout, axis-overlap or bound
    /// refusal.
    pub fn set_column_width<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        column: u8,
        width: u16,
    ) -> Result<()> {
        self.inner_mut()?
            .set_column_width(selector.into(), column, width)
            .map_err(Error::from)
    }

    /// Hide or show one column.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, record-layout, axis-overlap or bound
    /// refusal.
    pub fn set_column_hidden<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        column: u8,
        hidden: bool,
    ) -> Result<()> {
        self.inner_mut()?
            .set_column_hidden(selector.into(), column, hidden)
            .map_err(Error::from)
    }

    /// Change one cell's alignment, border styles or fill.
    ///
    /// # Errors
    ///
    /// Returns a typed indent, palette-color, selector, dependency,
    /// style-resource or bound refusal without staging.
    pub fn set_cell_format<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        reference: cell_values::Reference,
        change: &cell_values::FormatChange,
    ) -> Result<()> {
        self.inner_mut()?
            .set_cell_format(selector.into(), reference, change)
            .map_err(Error::from)
    }

    /// Remove one supported cell record.
    pub fn remove_cell<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        reference: cell_values::Reference,
    ) -> Result<()> {
        self.inner_mut()?
            .remove_cell(selector.into(), reference)
            .map_err(Error::from)
    }

    /// Set an existing cell's workbook XF style.
    pub fn set_style<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        reference: cell_values::Reference,
        style: cell_values::StyleIndex,
    ) -> Result<()> {
        self.inner_mut()?
            .set_style(selector.into(), reference, style)
            .map_err(Error::from)
    }

    /// Duplicate an existing XF resource for later style operations.
    pub fn duplicate_style(
        &mut self,
        source: cell_values::StyleIndex,
    ) -> Result<cell_values::StyleIndex> {
        self.inner_mut()?
            .duplicate_style(source)
            .map_err(Error::from)
    }

    /// Author one validated workbook XF resource for later style operations.
    pub fn author_style(&mut self, payload: &[u8]) -> Result<cell_values::StyleIndex> {
        self.inner_mut()?.author_style(payload).map_err(Error::from)
    }

    /// Insert rows and shift the checked worksheet closure.
    pub fn insert_rows<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        start: u16,
        count: u16,
    ) -> Result<()> {
        self.inner_mut()?
            .insert_rows(selector.into(), start, count)
            .map_err(Error::from)
    }

    /// Delete rows and shift the checked worksheet closure.
    pub fn delete_rows<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        start: u16,
        count: u16,
    ) -> Result<()> {
        self.inner_mut()?
            .delete_rows(selector.into(), start, count)
            .map_err(Error::from)
    }

    /// Insert columns and shift the checked worksheet closure.
    pub fn insert_columns<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        start: u8,
        count: u8,
    ) -> Result<()> {
        self.inner_mut()?
            .insert_columns(selector.into(), start, count)
            .map_err(Error::from)
    }

    /// Delete columns and shift the checked worksheet closure.
    pub fn delete_columns<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        start: u8,
        count: u8,
    ) -> Result<()> {
        self.inner_mut()?
            .delete_columns(selector.into(), start, count)
            .map_err(Error::from)
    }

    /// Rename one worksheet using the workbook's case-insensitive identity.
    pub fn rename_sheet<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
        name: &str,
    ) -> Result<()> {
        self.inner_mut()?
            .rename_sheet(selector.into(), name)
            .map_err(Error::from)
    }

    /// Reorder the complete tab catalog. The identity order is an exact
    /// no-op, and a tab-structure change must be the only operation of its
    /// edit.
    ///
    /// ```
    /// use litchi_xls::Package;
    /// use litchi_core::Position;
    /// use litchi_xls::cell_values::Selector;
    /// let mut builder = litchi_xls::Builder::new();
    /// builder.add_worksheet("A")?;
    /// builder.add_worksheet("B")?;
    /// let package = builder.build()?;
    /// let mut edit = package.edit()?;
    /// edit.reorder_sheets(&[
    ///     Selector::Position(Position::new(1)),
    ///     Selector::Position(Position::new(0)),
    /// ])?;
    /// let package = edit.commit()?.into_package();
    /// assert_eq!(package.workbook().worksheet_names()?, ["B", "A"]);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns a typed selector, protection, dependency-closure or
    /// exclusivity refusal without staging.
    pub fn reorder_sheets(&mut self, order: &[cell_values::Selector<'_>]) -> Result<()> {
        self.inner_mut()?.reorder_sheets(order).map_err(Error::from)
    }

    /// Insert one empty worksheet at a complete-catalog position.
    ///
    /// # Errors
    ///
    /// Returns a typed name, uniqueness, position, protection, macro-project
    /// or dependency-closure refusal without staging.
    pub fn insert_sheet(&mut self, position: litchi_core::Position, name: &str) -> Result<()> {
        self.inner_mut()?
            .insert_sheet(position, name)
            .map_err(Error::from)
    }

    /// Remove one ordinary worksheet tab.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, dependency, last-tab, last-visible-tab,
    /// protection or macro-project refusal without staging.
    pub fn remove_sheet<'a>(
        &mut self,
        selector: impl Into<cell_values::Selector<'a>>,
    ) -> Result<()> {
        self.inner_mut()?
            .remove_sheet(selector.into())
            .map_err(Error::from)
    }

    /// Publish the canonical immutable package commit.
    pub fn commit(self) -> Result<Commit> {
        check_source_guard(self.source_guard.as_ref())?;
        let source = self.package.clone();
        let inner = match self.inner {
            Some(inner) => inner,
            None => self.lazy.into_edit()?,
        };
        let leaf_commit = inner.commit().map_err(Error::from)?;
        check_source_guard(self.source_guard.as_ref())?;
        let leaf_diagnostics = leaf_commit.diagnostics();
        let (leaf_snapshot, leaf_patch) = leaf_commit.into_parts();
        let diagnostics = Diagnostics::from_leaf(leaf_diagnostics, leaf_patch.semantic().len());
        let package = if leaf_patch.is_noop() {
            let mut package = source;
            Arc::make_mut(&mut package.inner).edit_snapshot = Some(leaf_snapshot);
            package
        } else {
            package_from_snapshot(&source, leaf_snapshot)?
        };
        check_source_guard(self.source_guard.as_ref())?;
        Ok(Commit {
            package,
            patch: Patch::from_leaf(leaf_patch),
            diagnostics,
        })
    }
}

impl Package {
    /// The cell-value snapshot an edit of this package starts from, with
    /// the retained edit options.
    pub(crate) fn cell_source(&self) -> Result<(cell_values::Snapshot, EditOptions)> {
        super::ensure_publication_allowed(&self.inner).map_err(Error::Package)?;
        let options = self.retained_edit_options()?;
        let snapshot = self
            .edit_snapshot_with_options(options)
            .map_err(Error::Package)?;
        Ok((snapshot, options))
    }

    fn retained_edit_options(&self) -> Result<EditOptions> {
        if let Some(snapshot) = &self.inner.edit_snapshot {
            return Ok(EditOptions::new()
                .object_limits(snapshot.object_limits())
                .cfb_limits(snapshot.cfb_limits()));
        }
        if let Some(cfb) = self.source() {
            let selected = cfb.limits();
            let limits = OleFileLimits::new(selected.max_input_bytes())
                .and_then(|limits| limits.with_max_directory_bytes(selected.max_directory_bytes()))
                .and_then(|limits| {
                    limits.with_max_allocation_table_bytes(selected.max_allocation_table_bytes())
                })
                .map_err(PackageError::Cfb)?;
            return Ok(EditOptions::new().cfb_limits(limits));
        }
        Ok(EditOptions::default())
    }

    /// Start the primary XLS transaction under the retained default profile.
    pub fn edit(&self) -> Result<Edit> {
        Edit::from_package_with_options(self.clone(), self.retained_edit_options()?)
    }

    /// Start the primary XLS transaction under an explicitly requested profile.
    pub fn edit_with_options(&self, options: EditOptions) -> Result<Edit> {
        Edit::from_package_with_options(self.clone(), options)
    }

    /// Apply a primary exact-source patch without mutating this package.
    pub fn apply(&self, patch: &Patch) -> Result<Commit> {
        super::ensure_publication_allowed(&self.inner).map_err(Error::Package)?;
        let guard = capture_source_guard(self)?;
        let matches = if let Some(snapshot) = &self.inner.edit_snapshot {
            snapshot.bytes() == patch.before()
        } else {
            self.source_bytes().map_err(Error::Package)? == patch.before()
        };
        check_source_guard(guard.as_ref())?;
        if !matches {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            check_source_guard(guard.as_ref())?;
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics::applied(false, patch.operation_count()),
            });
        }
        if self.inner.signed {
            // Stale-base admission above remains first. Once the exact base
            // matches, a changed patch must be refused before the signed
            // probe or any target materialization can run.
            return Err(Error::CellValues(crate::Error::Signed));
        }
        let options = self.retained_edit_options()?;
        let snapshot = self
            .edit_snapshot_with_options(options)
            .map_err(Error::Package)?;
        let target = patch.leaf().apply(&snapshot).map_err(Error::from)?;
        check_source_guard(guard.as_ref())?;
        let package = package_from_snapshot(self, target)?;
        check_source_guard(guard.as_ref())?;
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: Diagnostics::applied(true, patch.operation_count()),
        })
    }
}

fn capture_source_guard(package: &Package) -> Result<Option<SourceGuard>> {
    match &package.inner.state {
        PackageState::Source { cfb, workbook, .. } => {
            let original = workbook.needs_source_fence().then(|| workbook.clone());
            check_original_source_guard(original.as_ref())?;
            let expected = match cfb.source_version() {
                Ok(expected) => expected,
                Err(error) => {
                    // The original check above and this target observation
                    // are separate windows.  A composed target can report
                    // an in-flight original mutation as CFB/I/O here, before
                    // a SourceGuard exists; promote the original error first.
                    check_original_source_guard(original.as_ref())?;
                    return Err(Error::Package(PackageError::Source(
                        SourceBackedError::from(error),
                    )));
                },
            };
            Ok(Some(SourceGuard {
                source: Arc::clone(cfb),
                expected,
                original,
            }))
        },
        PackageState::Eager { .. } => Ok(None),
    }
}

fn check_source_guard(guard: Option<&SourceGuard>) -> Result<()> {
    let Some(guard) = guard else {
        return Ok(());
    };
    check_original_source_guard(guard.original.as_ref())?;
    let observed = match guard.source.source_version() {
        Ok(observed) => observed,
        Err(error) => {
            // A composed target can report the original mutation as an
            // I/O/CFB error after the leading check. Re-fence the retained
            // original owner before exposing that target error, just as the
            // source-backed semantic owner does.
            check_original_source_guard(guard.original.as_ref())?;
            return Err(Error::Package(PackageError::Source(
                SourceBackedError::from(error),
            )));
        },
    };
    if observed != guard.expected {
        check_original_source_guard(guard.original.as_ref())?;
        return Err(Error::Package(PackageError::Source(
            SourceBackedError::SourceChanged {
                expected: guard.expected,
                observed,
            },
        )));
    }
    Ok(())
}

fn check_original_source_guard(workbook: Option<&crate::SourceBackedWorkbook>) -> Result<()> {
    if let Some(workbook) = workbook {
        workbook
            .ensure_current()
            .map_err(|error| Error::Package(PackageError::Source(error)))?;
    }
    Ok(())
}

pub(crate) fn package_from_replacement_bytes(source: &Package, bytes: Vec<u8>) -> Result<Package> {
    super::ensure_publication_allowed(&source.inner).map_err(Error::Package)?;
    let guard = capture_source_guard(source)?;
    let result = (|| {
        let snapshot = crate::raw::cell_values_snapshot_from_bytes(source, bytes)?;
        package_from_snapshot(source, snapshot)
    })();
    check_source_guard(guard.as_ref())?;
    result
}

fn package_from_snapshot(source: &Package, snapshot: cell_values::Snapshot) -> Result<Package> {
    if let super::RetainedPolicy::Source { limits, leniency } = source.inner.policy {
        let selected = snapshot.cfb_limits();
        let cfb_limits = litchi_cfb::SharedOleFileLimits::new(selected.max_input_bytes())
            .and_then(|limits| limits.with_max_directory_bytes(selected.max_directory_bytes()))
            .and_then(|limits| {
                limits.with_max_allocation_table_bytes(selected.max_allocation_table_bytes())
            })
            .map_err(PackageError::Cfb)?;
        let cfb = Arc::new(
            SharedOleFile::open_owned_vec(
                snapshot.bytes_shared(),
                snapshot.source_version(),
                cfb_limits,
            )
            .map_err(PackageError::Cfb)?,
        );
        let workbook = crate::SourceBackedWorkbook::from_shared_ole_file_with_options(
            cfb,
            crate::SourceBackedOptions::default()
                .limits(limits)
                .leniency(leniency),
        )
        .map_err(PackageError::Source)?;
        let mut package = Package::from_source(workbook)?;
        Arc::make_mut(&mut package.inner).edit_snapshot = Some(snapshot);
        return Ok(package);
    }
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(snapshot.bytes().len())
        .map_err(|_| Error::CellValues(crate::Error::Allocation("XLS primary read projection")))?;
    bytes.extend_from_slice(snapshot.bytes());
    let mut ole = litchi_cfb::OleFile::open(std::io::Cursor::new(bytes), snapshot.cfb_limits())
        .map_err(PackageError::Cfb)?;
    let metadata = super::eager_metadata(&mut ole);
    let options = OpenOptions::default()
        .leniency(source.leniency())
        .compatibility_profile(source.compatibility_profile().unwrap_or_default());
    let reader = crate::raw::Reader::from_admitted_ole_file(ole, options)?;
    let encryption = reader.encryption();
    Ok(Package {
        inner: Arc::new(super::PackageInner {
            state: PackageState::Eager {
                reader: Arc::new(reader),
                encryption,
            },
            policy: source.inner.policy,
            signed: source.inner.signed,
            metadata: OnceLock::from(metadata),
            catalog: OnceLock::new(),
            edit_snapshot: Some(snapshot),
            workbook_stream: OnceLock::new(),
        }),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_cfb::{SameLengthStreamSplice, SharedOleFile, StreamSpliceLimits};
    use litchi_core::{ReadAt, SourceVersion};
    use std::io;
    use std::sync::atomic::{AtomicU64, Ordering};

    struct AdmissionSource {
        bytes: Arc<Vec<u8>>,
        revision: AtomicU64,
        observations: AtomicU64,
        bump_after: AtomicU64,
    }

    impl AdmissionSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes: Arc::new(bytes),
                revision: AtomicU64::new(0),
                observations: AtomicU64::new(0),
                bump_after: AtomicU64::new(0),
            }
        }

        fn bump_after_observation(&self, ordinal: u64) {
            self.observations.store(0, Ordering::Relaxed);
            self.bump_after.store(ordinal, Ordering::Relaxed);
        }
    }

    impl ReadAt for AdmissionSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let start = usize::try_from(offset)
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset overflow"))?;
            if start >= self.bytes.len() || output.is_empty() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - start);
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            let ordinal = self.observations.fetch_add(1, Ordering::Relaxed) + 1;
            let observed =
                SourceVersion::new(0x584c_535f_4144_4d49, self.revision.load(Ordering::Relaxed));
            if self.bump_after.load(Ordering::Relaxed) == ordinal {
                self.revision.fetch_add(1, Ordering::Relaxed);
            }
            Ok(observed)
        }
    }

    #[test]
    fn capture_source_guard_promotes_target_error_to_original_source_changed() {
        let mut builder = crate::Builder::new();
        let sheet = builder.add_worksheet("Admission").unwrap();
        builder.write_number(sheet, 0, 0, 42.0).unwrap();
        let bytes = builder.build().unwrap().to_bytes().unwrap();
        let source = Arc::new(AdmissionSource::new(bytes));
        let source_cfb = Arc::new(
            SharedOleFile::open(source.clone(), litchi_cfb::SharedOleFileLimits::default())
                .unwrap(),
        );
        let original = crate::SourceBackedWorkbook::from_shared_ole_file_with_options(
            Arc::clone(&source_cfb),
            crate::SourceBackedOptions::default(),
        )
        .unwrap();

        let workbook_stream = source_cfb.open_stream(&["Workbook"]).unwrap();
        let expected = 42.0_f64.to_le_bytes();
        let offset = workbook_stream
            .windows(expected.len())
            .position(|window| window == expected)
            .expect("writer must retain the Number value in the Workbook stream");
        let replacement = 43.0_f64.to_le_bytes();
        let plan = source_cfb
            .plan_same_length_stream_splices(
                vec![SameLengthStreamSplice::new(
                    vec!["Workbook".to_string()],
                    offset as u64,
                    Arc::from(expected),
                    Arc::from(replacement),
                )],
                StreamSpliceLimits::default(),
            )
            .unwrap();
        let composed = Arc::new(plan.composed_source().unwrap());
        let target_cfb = Arc::new(
            SharedOleFile::open(composed, litchi_cfb::SharedOleFileLimits::default()).unwrap(),
        );
        let deferred = original.committed_deferred(plan).unwrap();
        let limits = deferred.retained_limits();
        let leniency = deferred.retained_leniency();
        let package = Package {
            inner: Arc::new(super::super::PackageInner {
                state: PackageState::Source {
                    cfb: target_cfb,
                    workbook: deferred,
                    encryption: None,
                    materialized: OnceLock::new(),
                },
                policy: super::super::RetainedPolicy::Source { limits, leniency },
                signed: false,
                metadata: OnceLock::new(),
                catalog: OnceLock::new(),
                edit_snapshot: None,
                workbook_stream: OnceLock::new(),
            }),
        };

        // The first observation is the retained original-owner fence. The
        // target CFB's composed-source version check then observes twice; the
        // mutation after its first observation must be reported as the typed
        // original SourceChanged error, before a SourceGuard exists.
        source.bump_after_observation(2);
        assert!(matches!(
            package.edit(),
            Err(Error::Package(PackageError::Source(
                SourceBackedError::SourceChanged { .. }
            )))
        ));
    }
}
