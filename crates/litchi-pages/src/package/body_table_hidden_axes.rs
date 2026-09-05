//! Exact-source hidden-row/hidden-column transactions for Pages tables.
//!
//! The public value is deliberately archive free.  This adapter resolves the
//! rooted table through `table_lock`, proves its hidden-state dependency closure, and
//! only then crosses into the strict hidden-state codec.  Payload surgery is
//! limited to the selected table-info/model fields; all unrelated wire bytes
//! and all unrelated package members remain source authoritative.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The focused transaction keeps its proof and rewrite helpers together."
)]

use std::fmt;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_iwa_archive::{SourceCatalog, package::EntryEdit};
use litchi_iwa_common::{decode_varint_from_bytes, varint::encoded_len};
use litchi_iwa_core::{Archive, ArchiveObject, MessageInfo, RawMessage};
use litchi_iwa_protos::numbers_table_physical_sort_codec as uid_codec;
use litchi_iwa_protos::pages_hidden_state_codec as codec;
use thiserror::Error;

use super::{Package, PackageError, page_layout, table_lock};
use crate::selector::BodyTableSelector;
use crate::table::hidden_axes::{AxisIndex, HiddenAxes};

const TABLE_INFO_MESSAGE_TYPES: &[u32] = &[6_000, 6_003];
const TABLE_MODEL_MESSAGE_TYPES: &[u32] = &[6_000, 6_001];
const TABLE_INFO_PIVOT_FIELD: u32 = 16;
const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
const HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE: u32 = 6_204;
const FILTER_SET_MESSAGE_TYPE: u32 = 6_220;
const UID_MAP_MESSAGE_TYPE: u32 = 6_267;
const LEGACY_UID_MAP_MESSAGE_TYPE: u32 = 6_200;
const CURRENT_MESSAGE_VERSIONS: &[u32] = &[1, 0, 5];
const NATIVE_TABLE_MODEL_MESSAGE_VERSIONS: &[u32] = &[3, 2, 10];
const NATIVE_FORMULA_OWNER_MESSAGE_VERSIONS: &[u32] = &[3, 2, 10];
const FORMULA_OWNER_REFERENCE_PATH: &[u32] = &[11];
const MODEL_PIVOT_OWNER_FIELD: u32 = 85;

/// The selected Pages producer profile determines which archive-header
/// version tuples and dependency metadata are authoritative.  The qualified
/// current profile remains the default; the native visible profile is
/// admitted only for the exact 6000/6001 role pair observed in the checked-in
/// Pages document.
#[derive(Clone, Copy, PartialEq, Eq)]
enum GraphProfile {
    Indexed,
    NativeVisible,
}

impl GraphProfile {
    const fn is_native(self) -> bool {
        matches!(self, Self::NativeVisible)
    }

    const fn info_versions(self) -> &'static [u32] {
        // Both admitted profiles use the current TableInfoArchive header.
        let _ = self;
        CURRENT_MESSAGE_VERSIONS
    }

    const fn model_versions(self) -> &'static [u32] {
        match self {
            Self::Indexed => CURRENT_MESSAGE_VERSIONS,
            Self::NativeVisible => NATIVE_TABLE_MODEL_MESSAGE_VERSIONS,
        }
    }

    const fn formula_owner_versions(self) -> &'static [u32] {
        match self {
            Self::Indexed => CURRENT_MESSAGE_VERSIONS,
            Self::NativeVisible => NATIVE_FORMULA_OWNER_MESSAGE_VERSIONS,
        }
    }
}

/// A content-free location associated with one hidden-axis operation.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyTableHiddenAxesPath {
    /// The complete Pages package.
    Package,
    /// One rooted table at a checked body position.
    Table { table: usize },
}

impl fmt::Debug for BodyTableHiddenAxesPath {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        // A table position is useful to the private adapter, but public
        // diagnostics are intentionally content-free.  In particular, do not
        // expose the selector, native proof, or any package member here.
        formatter.write_str(match self {
            Self::Package => "BodyTableHiddenAxesPath::Package",
            Self::Table { .. } => "BodyTableHiddenAxesPath::Table",
        })
    }
}

/// Finite resources governed by one hidden-axis transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyTableHiddenAxesLimitKind {
    InputBytes,
    OutputBytes,
    Entries,
    EntryBytes,
    TotalEntryBytes,
    PackageBytes,
    PayloadBytes,
    TotalPayloadBytes,
    PayloadObjects,
    PayloadMessages,
    PayloadItems,
    PayloadReferences,
    WireBytes,
    WireOutputBytes,
    WireFields,
    WireNesting,
    WireWork,
    TransactionWork,
}

impl fmt::Display for BodyTableHiddenAxesLimitKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalEntryBytes => "total entry bytes",
            Self::PackageBytes => "package bytes",
            Self::PayloadBytes => "payload bytes",
            Self::TotalPayloadBytes => "total payload bytes",
            Self::PayloadObjects => "payload objects",
            Self::PayloadMessages => "payload messages",
            Self::PayloadItems => "payload items",
            Self::PayloadReferences => "payload references",
            Self::WireBytes => "wire bytes",
            Self::WireOutputBytes => "wire output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::TransactionWork => "transaction work",
        })
    }
}

/// A redacted failure from a body-table hidden-axis read or transaction.
#[derive(Error, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum BodyTableHiddenAxesError {
    #[error("the Pages body has no table matching the requested selector")]
    TableNotFound,
    #[error("more than one Pages body table has the requested name")]
    AmbiguousTableName,
    #[error("the selected Pages body-table selector is ambiguous")]
    AmbiguousSelector,
    #[error("this Pages source does not support exact body-table hidden-axis editing")]
    UnsupportedSource,
    #[error("the selected Pages body-table hidden-axis source is invalid")]
    InvalidSource,
    #[error("the selected Pages body table is locked")]
    TableLocked,
    #[error("the selected Pages body-table hidden-axis dependency is unsupported")]
    UnsupportedDependency,
    #[error(
        "Pages body-table hidden axes {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        kind: BodyTableHiddenAxesLimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} units for Pages body-table hidden axes")]
    Allocation { amount: usize },
    #[error("the edited Pages body-table hidden axes failed semantic verification")]
    Verification,
    #[error("the Pages body-table hidden-axis patch does not match the exact source package")]
    PatchConflict,
}

impl fmt::Debug for BodyTableHiddenAxesError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::LimitExceeded {
                kind,
                observed,
                maximum,
            } => {
                // Resource categories and amounts are bounded diagnostics;
                // no source fingerprint, payload, UUID, or native identity is
                // ever included in this public representation.
                formatter
                    .debug_struct("BodyTableHiddenAxesError::LimitExceeded")
                    .field("kind", kind)
                    .field("observed", observed)
                    .field("maximum", maximum)
                    .finish_non_exhaustive()
            },
            Self::Allocation { amount } => formatter
                .debug_struct("BodyTableHiddenAxesError::Allocation")
                .field("allocation_units", amount)
                .finish_non_exhaustive(),
            Self::TableNotFound => formatter.write_str("BodyTableHiddenAxesError::TableNotFound"),
            Self::AmbiguousTableName => {
                formatter.write_str("BodyTableHiddenAxesError::AmbiguousTableName")
            },
            Self::AmbiguousSelector => {
                formatter.write_str("BodyTableHiddenAxesError::AmbiguousSelector")
            },
            Self::UnsupportedSource => {
                formatter.write_str("BodyTableHiddenAxesError::UnsupportedSource")
            },
            Self::InvalidSource => formatter.write_str("BodyTableHiddenAxesError::InvalidSource"),
            Self::TableLocked => formatter.write_str("BodyTableHiddenAxesError::TableLocked"),
            Self::UnsupportedDependency => {
                formatter.write_str("BodyTableHiddenAxesError::UnsupportedDependency")
            },
            Self::Verification => formatter.write_str("BodyTableHiddenAxesError::Verification"),
            Self::PatchConflict => formatter.write_str("BodyTableHiddenAxesError::PatchConflict"),
        }
    }
}

/// Selector-first immutable hidden-axis edit.
pub struct BodyTableHiddenAxesEdit<'a> {
    source: &'a Package,
    target: table_lock::BodyTableTarget,
    before: HiddenAxes,
    axes: HiddenAxes,
}

impl fmt::Debug for BodyTableHiddenAxesEdit<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("BodyTableHiddenAxesEdit")
            // Hidden row/column positions are semantic content.  Public
            // diagnostics may report only bounded counts and state changes;
            // the source, proof, and selected package member stay private.
            .field("before_count", &self.before.as_slice().len())
            .field("axes_count", &self.axes.as_slice().len())
            .field("changed", &(self.before != self.axes))
            .finish_non_exhaustive()
    }
}

impl BodyTableHiddenAxesEdit<'_> {
    #[must_use]
    pub const fn path(&self) -> BodyTableHiddenAxesPath {
        BodyTableHiddenAxesPath::Table {
            table: self.target.table_position,
        }
    }

    #[must_use]
    pub const fn before(&self) -> &HiddenAxes {
        &self.before
    }

    #[must_use]
    pub const fn axes(&self) -> &HiddenAxes {
        &self.axes
    }

    #[must_use]
    pub const fn hidden_axes(&self) -> &HiddenAxes {
        &self.axes
    }

    #[must_use]
    pub fn set(mut self, axes: HiddenAxes) -> Self {
        self.axes = axes;
        self
    }

    #[must_use]
    pub fn clear(self) -> Self {
        self.set(HiddenAxes::empty())
    }

    #[must_use]
    pub fn reset(self) -> Self {
        self.clear()
    }

    pub fn commit(self) -> Result<BodyTableHiddenAxesCommit, BodyTableHiddenAxesError> {
        commit_edit(self)
    }
}

/// Exact-source reversible hidden-axis patch.
#[derive(Clone, PartialEq, Eq)]
pub struct BodyTableHiddenAxesPatch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    proof: table_lock::BodyTableTarget,
    before: HiddenAxes,
    after: HiddenAxes,
    source_previews: usize,
    target_previews: usize,
    touched_components: usize,
}

impl fmt::Debug for BodyTableHiddenAxesPatch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("BodyTableHiddenAxesPatch")
            .field("before_count", &self.before.as_slice().len())
            .field("after_count", &self.after.as_slice().len())
            .field("changed", &(self.before != self.after))
            .field("noop", &self.is_noop())
            .field("touched_components", &self.touched_components)
            .finish_non_exhaustive()
    }
}

impl BodyTableHiddenAxesPatch {
    #[must_use]
    pub const fn path(&self) -> BodyTableHiddenAxesPath {
        BodyTableHiddenAxesPath::Table {
            table: self.proof.table_position,
        }
    }

    #[must_use]
    pub const fn before(&self) -> &HiddenAxes {
        &self.before
    }

    #[must_use]
    pub const fn after(&self) -> &HiddenAxes {
        &self.after
    }

    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
            && self.source_fingerprint == self.target_fingerprint
            && (Arc::ptr_eq(&self.source, &self.target) || self.source == self.target)
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            proof: self.proof.clone(),
            before: self.after.clone(),
            after: self.before.clone(),
            source_previews: self.target_previews,
            target_previews: self.source_previews,
            touched_components: self.touched_components,
        }
    }
}

/// Content-free publication diagnostics.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct BodyTableHiddenAxesDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl fmt::Debug for BodyTableHiddenAxesDiagnostics {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyTableHiddenAxesDiagnostics")
            .field("changed", &self.changed)
            .field("touched_components", &self.touched_components)
            .field("deleted_previews", &self.deleted_previews)
            .field("full_reparse_performed", &self.full_reparse_performed)
            .finish_non_exhaustive()
    }
}

impl BodyTableHiddenAxesDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(deleted_previews: usize, touched_components: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            deleted_previews,
            full_reparse_performed: true,
        }
    }

    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully reopened immutable result of one hidden-axis transaction.
#[must_use = "a body-table hidden-axis commit contains the validated package snapshot"]
pub struct BodyTableHiddenAxesCommit {
    package: Package,
    patch: BodyTableHiddenAxesPatch,
    diagnostics: BodyTableHiddenAxesDiagnostics,
}

impl fmt::Debug for BodyTableHiddenAxesCommit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        // Formatting `package` or `patch` recursively would make this type's
        // redaction depend on every nested implementation.  Keep the public
        // commit representation to the already-safe diagnostics summary.
        formatter
            .debug_struct("BodyTableHiddenAxesCommit")
            .field("changed", &self.diagnostics.changed)
            .field("touched_components", &self.diagnostics.touched_components)
            .field("deleted_previews", &self.diagnostics.deleted_previews)
            .field(
                "full_reparse_performed",
                &self.diagnostics.full_reparse_performed,
            )
            .finish_non_exhaustive()
    }
}

impl BodyTableHiddenAxesCommit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    #[must_use]
    pub const fn patch(&self) -> &BodyTableHiddenAxesPatch {
        &self.patch
    }

    #[must_use]
    pub const fn diagnostics(&self) -> &BodyTableHiddenAxesDiagnostics {
        &self.diagnostics
    }
}

#[derive(Clone)]
struct ModelValues {
    rows: u32,
    columns: u32,
    map: Option<NonZeroU64>,
    map_ref: Option<codec::ReferenceSnapshot>,
    hidden_rows: Option<u32>,
    hidden_columns: Option<u32>,
    filtered_rows: Option<u32>,
    user_rows: Option<u32>,
    user_columns: Option<u32>,
    formula_columns: Option<NonZeroU64>,
    formula_rows: Option<NonZeroU64>,
    formula_columns_ref: Option<codec::ReferenceSnapshot>,
    formula_rows_ref: Option<codec::ReferenceSnapshot>,
    owner: Option<codec::HiddenStatesOwnerSnapshot>,
    pivot: bool,
}

#[derive(Clone)]
struct InfoValues {
    model: NonZeroU64,
    model_ref: codec::ReferenceSnapshot,
    map: Option<NonZeroU64>,
    map_ref: Option<codec::ReferenceSnapshot>,
    hidden_uuid: Option<codec::UuidSnapshot>,
    pivot: bool,
}

#[derive(Clone)]
struct Graph {
    profile: GraphProfile,
    target: table_lock::BodyTableTarget,
    model: ModelValues,
    info: InfoValues,
    rows: Vec<codec::UuidSnapshot>,
    columns: Vec<codec::UuidSnapshot>,
    row_indices: UidIndex,
    column_indices: UidIndex,
    before: HiddenAxes,
}

/// A bounded binary-search index for the map's physical axis order.
///
/// The native UID map stores its stable UUIDs in sorted order and then a
/// permutation for physical order.  Semantic state records are physical, so
/// repeatedly scanning that permutation would turn a hostile sparse state
/// list into quadratic work.  Keep one sorted `(UUID, physical-index)` view
/// and use logarithmic lookups for every extent proof and rewrite.
#[derive(Clone)]
struct UidIndex {
    entries: Vec<(codec::UuidSnapshot, usize)>,
}

impl UidIndex {
    const fn empty() -> Self {
        Self {
            entries: Vec::new(),
        }
    }

    fn new(
        physical: &[codec::UuidSnapshot],
        budget: &mut table_lock::WireBudget,
    ) -> Result<Self, BodyTableHiddenAxesError> {
        let mut entries = Vec::new();
        entries.try_reserve_exact(physical.len()).map_err(|_| {
            BodyTableHiddenAxesError::Allocation {
                amount: physical.len(),
            }
        })?;
        for (index, uid) in physical.iter().copied().enumerate() {
            validate_uuid(uid)?;
            entries.push((uid, index));
        }
        entries.sort_unstable_by_key(|(uid, _)| (uid.lower(), uid.upper()));
        for pair in entries.windows(2) {
            if pair[0].0 == pair[1].0 {
                return Err(BodyTableHiddenAxesError::InvalidSource);
            }
        }
        let levels = if entries.len() <= 1 {
            0
        } else {
            (usize::BITS - (entries.len() - 1).leading_zeros()) as usize
        };
        budget
            .charge_payload_work(entries.len().checked_mul(levels).ok_or(
                BodyTableHiddenAxesError::LimitExceeded {
                    kind: BodyTableHiddenAxesLimitKind::TransactionWork,
                    observed: u64::MAX,
                    maximum:
                        u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
                },
            )?)
            .map_err(map_lock_error)?;
        Ok(Self { entries })
    }

    fn index_of(&self, uid: codec::UuidSnapshot) -> Option<usize> {
        self.entries
            .binary_search_by_key(&(uid.lower(), uid.upper()), |(candidate, _)| {
                (candidate.lower(), candidate.upper())
            })
            .ok()
            .map(|position| self.entries[position].1)
    }
}

impl Package {
    /// Read one rooted body-table's user-hidden rows and columns.
    pub fn body_table_hidden_axes<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<HiddenAxes, BodyTableHiddenAxesError> {
        let mut budget = transaction_budget(self)?;
        let graph = resolve_graph(self, selector.into(), &mut budget)?;
        Ok(graph.before)
    }

    /// Start a selector-first hidden-axis edit.
    pub fn edit_body_table_hidden_axes<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<BodyTableHiddenAxesEdit<'_>, BodyTableHiddenAxesError> {
        let mut budget = transaction_budget(self)?;
        let graph = resolve_graph(self, selector.into(), &mut budget)?;
        Ok(BodyTableHiddenAxesEdit {
            source: self,
            target: graph.target,
            before: graph.before.clone(),
            axes: graph.before,
        })
    }

    /// Apply a patch only when both its exact source and semantic proof match.
    pub fn apply_body_table_hidden_axes(
        &self,
        patch: &BodyTableHiddenAxesPatch,
    ) -> Result<BodyTableHiddenAxesCommit, BodyTableHiddenAxesError> {
        let mut budget = transaction_budget(self)?;
        let source = self.state.source.source_bytes();
        if page_layout::fingerprint(source) != patch.source_fingerprint
            || source != patch.source.as_ref()
        {
            return Err(BodyTableHiddenAxesError::PatchConflict);
        }
        let graph = resolve_graph(
            self,
            BodyTableSelector::index(patch.proof.table_position),
            &mut budget,
        )?;
        if graph.target.model_identifier != patch.proof.model_identifier
            || graph.before != patch.before
        {
            return Err(BodyTableHiddenAxesError::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(BodyTableHiddenAxesCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: BodyTableHiddenAxesDiagnostics::unchanged(),
            });
        }
        if graph.profile.is_native() {
            // The native visible profile is read/no-op only until a Pages
            // producer round-trip proves that its kind-1 dependency envelope
            // can be rewritten without changing unrelated native state.
            return Err(BodyTableHiddenAxesError::UnsupportedDependency);
        }
        if graph.target.explicit_locked == Some(true) {
            return Err(BodyTableHiddenAxesError::TableLocked);
        }
        if !self.state.source.source_is_exact()
            || page_layout::fingerprint(patch.target.as_ref()) != patch.target_fingerprint
        {
            return Err(BodyTableHiddenAxesError::PatchConflict);
        }
        let candidate = reopen(self, Arc::clone(&patch.target), &mut budget)?;
        let verified = resolve_graph(
            &candidate,
            BodyTableSelector::index(patch.proof.table_position),
            &mut budget,
        )?;
        if verified.target.model_identifier != patch.proof.model_identifier
            || verified.before != patch.after
        {
            return Err(BodyTableHiddenAxesError::Verification);
        }
        verify_locality(
            self,
            &candidate,
            &patch.proof,
            patch.touched_components,
            &mut budget,
        )?;
        Ok(BodyTableHiddenAxesCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: BodyTableHiddenAxesDiagnostics::published(
                patch.source_previews.saturating_sub(patch.target_previews),
                patch.touched_components,
            ),
        })
    }
}

fn commit_edit(
    edit: BodyTableHiddenAxesEdit<'_>,
) -> Result<BodyTableHiddenAxesCommit, BodyTableHiddenAxesError> {
    let source = edit.source;
    let mut budget = transaction_budget(source)?;
    let source_bytes = source.state.source.shared_source();
    let source_fingerprint = page_layout::fingerprint(source_bytes.as_ref());
    let source_previews = preview_count(source);
    if edit.before == edit.axes {
        // Re-read against the captured target so an edit cannot silently
        // publish after its source graph changed underneath it.
        let graph = resolve_graph(
            source,
            BodyTableSelector::index(edit.target.table_position),
            &mut budget,
        )?;
        if graph.target.model_identifier != edit.target.model_identifier
            || graph.before != edit.before
        {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        return Ok(BodyTableHiddenAxesCommit {
            package: source.snapshot(),
            patch: BodyTableHiddenAxesPatch {
                source: Arc::clone(&source_bytes),
                target: source_bytes,
                source_fingerprint,
                target_fingerprint: source_fingerprint,
                proof: edit.target,
                before: edit.before,
                after: edit.axes,
                source_previews,
                target_previews: source_previews,
                touched_components: 0,
            },
            diagnostics: BodyTableHiddenAxesDiagnostics::unchanged(),
        });
    }
    if !source.state.source.source_is_exact() {
        return Err(BodyTableHiddenAxesError::UnsupportedSource);
    }
    if edit.target.explicit_locked == Some(true) {
        return Err(BodyTableHiddenAxesError::TableLocked);
    }
    let graph = resolve_graph(
        source,
        BodyTableSelector::index(edit.target.table_position),
        &mut budget,
    )?;
    if graph.before != edit.before {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    if graph.profile.is_native() {
        // Native kind-1 formula ownership is qualified for reads and exact
        // no-ops only.  Refuse before any candidate allocation or publication.
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    if graph.model.pivot || graph.info.pivot {
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    // Creating this graph requires a producer-owned formula dependency and
    // package metadata creation records.  Until a native Pages producer path
    // can prove those records, fail closed before allocating candidate state.
    if graph.model.owner.is_none() && !edit.axes.is_empty() {
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    validate_axis_bounds(&graph, &edit.axes)?;
    let candidate = rewrite(source, &graph, &edit.axes, &mut budget)?;
    let verified = resolve_graph(
        &candidate,
        BodyTableSelector::index(edit.target.table_position),
        &mut budget,
    )?;
    if verified.target.model_identifier != edit.target.model_identifier
        || verified.before != edit.axes
    {
        return Err(BodyTableHiddenAxesError::Verification);
    }
    if preview_count(&candidate) != 0 {
        return Err(BodyTableHiddenAxesError::Verification);
    }
    verify_locality(source, &candidate, &edit.target, 1, &mut budget)?;
    let target = candidate.state.source.shared_source();
    let target_fingerprint = page_layout::fingerprint(target.as_ref());
    let target_previews = preview_count(&candidate);
    Ok(BodyTableHiddenAxesCommit {
        package: candidate,
        patch: BodyTableHiddenAxesPatch {
            source: source_bytes,
            target,
            source_fingerprint,
            target_fingerprint,
            proof: edit.target,
            before: edit.before,
            after: edit.axes,
            source_previews,
            target_previews,
            touched_components: 1,
        },
        diagnostics: BodyTableHiddenAxesDiagnostics::published(
            source_previews.saturating_sub(target_previews),
            1,
        ),
    })
}

fn transaction_budget(
    package: &Package,
) -> Result<table_lock::WireBudget, BodyTableHiddenAxesError> {
    let mut budget =
        table_lock::WireBudget::new(package.state.source.limits()).map_err(map_lock_error)?;
    budget
        .charge_source_catalog(&package.state.source)
        .map_err(map_lock_error)?;
    Ok(budget)
}

fn resolve_graph(
    package: &Package,
    selector: BodyTableSelector<'_>,
    budget: &mut table_lock::WireBudget,
) -> Result<Graph, BodyTableHiddenAxesError> {
    let target = package
        .resolve_body_table_with_budget(selector, budget)
        .map_err(map_lock_error)?;
    table_lock::validate_body_table_target(package, &target, budget).map_err(map_lock_error)?;
    if !TABLE_MODEL_MESSAGE_TYPES.contains(&target.model_message_type)
        || !TABLE_INFO_MESSAGE_TYPES.contains(&target.message_type)
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let profile = classify_graph_profile(package, &target)?;
    let model_raw = message_at(
        package,
        target.model_component_index,
        target.model_object_index,
        target.model_message_index,
        target.model_identifier,
        target.model_message_type,
        profile.model_versions(),
    )?;
    let info_raw = message_at(
        package,
        target.component_index,
        target.object_index,
        target.info_message_index,
        target.drawable_identifier,
        target.message_type,
        profile.info_versions(),
    )?;
    let model = decode_model(&model_raw.data, budget)?;
    let info = decode_info(&info_raw.data, budget)?;
    for reference in [
        Some(info.model_ref),
        info.map_ref,
        model.map_ref,
        model.formula_columns_ref,
        model.formula_rows_ref,
    ]
    .into_iter()
    .flatten()
    {
        validate_reference_shape(reference)?;
    }
    if let Some(map) = model.map {
        validate_reference_metadata_for_profile(
            package,
            target.model_component_index,
            target.model_object_index,
            target.model_message_index,
            map,
            &[46],
            profile.model_versions(),
            profile,
            budget,
        )?;
    }
    if let Some(map) = info.map {
        validate_reference_metadata_for_profile(
            package,
            target.component_index,
            target.object_index,
            target.info_message_index,
            map,
            &[6],
            profile.info_versions(),
            profile,
            budget,
        )?;
    }
    for (reference, path) in [(model.formula_columns, [34]), (model.formula_rows, [35])] {
        if let Some(reference) = reference {
            validate_reference_metadata_for_profile(
                package,
                target.model_component_index,
                target.model_object_index,
                target.model_message_index,
                reference,
                &path,
                profile.model_versions(),
                profile,
                budget,
            )?;
        } else if profile.is_native() {
            // The native visible profile carries both dependency roots.  An
            // absent root cannot be interpreted as an ownerless empty table.
            return Err(BodyTableHiddenAxesError::InvalidSource);
        } else {
            reject_reference_path_metadata(
                package,
                target.model_component_index,
                target.model_object_index,
                target.model_message_index,
                &path,
                budget,
            )?;
        }
    }
    let map = model.map.ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if info.map.is_some_and(|info_map| info_map != map) {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    if info.model != target.model_identifier {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    validate_reference_metadata_for_profile(
        package,
        target.component_index,
        target.object_index,
        target.info_message_index,
        info.model,
        &[2],
        profile.info_versions(),
        profile,
        budget,
    )?;
    if info.map.is_none() {
        reject_reference_identifier_metadata(
            package,
            target.component_index,
            target.object_index,
            target.info_message_index,
            map,
            budget,
        )?;
    }
    if let Some(hidden_uuid) = info.hidden_uuid {
        validate_uuid(hidden_uuid)?;
    }
    if profile.is_native() && model.owner.is_none() {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let (rows, columns, objects) = if model.owner.is_none() {
        // An ownerless table still proves its mandatory UID map, but does not
        // need a package-wide object index: no dependency edge can be edited
        // or resolved on this read/no-op path.
        let location = unique_object_location(package, map, budget)?;
        let (rows, columns) = decode_uid_map_at(package, location, &model, budget)?;
        (rows, columns, None)
    } else {
        let objects = global_objects(package, budget)?;
        let (rows, columns) = decode_uid_map(package, &objects, map, &model, budget)?;
        (rows, columns, Some(objects))
    };
    validate_cross_axis_uuid_identity(&rows, &columns, budget)?;
    if model.owner.is_none() && info.hidden_uuid.is_some() {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    if model.owner.is_none() && (model.formula_columns.is_some() || model.formula_rows.is_some()) {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    validate_model_counts(&model, info.hidden_uuid, budget)?;
    if model.owner.is_none() {
        // An ownerless table is a complete, valid empty read/no-op shape.  Do
        // not scan unrelated package-wide formula records or allocate axis
        // indexes that no changed operation can use.
        return Ok(Graph {
            profile,
            target,
            model,
            info,
            rows,
            columns,
            row_indices: UidIndex::empty(),
            column_indices: UidIndex::empty(),
            before: HiddenAxes::empty(),
        });
    }
    let objects = objects.ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let row_indices = UidIndex::new(&rows, budget)?;
    let column_indices = UidIndex::new(&columns, budget)?;
    let model_location = object_location(&objects, target.model_identifier)?;
    if model_location.component_index != target.model_component_index
        || model_location.object_index != target.model_object_index
        || model_location.identifier == target.drawable_identifier.get()
    {
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    let drawable_location = object_location(&objects, target.drawable_identifier)?;
    if drawable_location.component_index != target.component_index
        || drawable_location.object_index != target.object_index
    {
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    let formula_owner_messages = formula_owner_messages(package, &objects, budget)?;
    let formula_owner_uid = formula_owner_for(
        package,
        &formula_owner_messages,
        drawable_location,
        model_location,
        profile,
        budget,
    )?;
    if model.owner.is_some() && formula_owner_uid.is_none() {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    if model.owner.is_some() && (model.formula_columns.is_none() || model.formula_rows.is_none()) {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    if let Some(owner) = model.owner.as_ref() {
        let formula_uid = formula_owner_uid.ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        validate_uuid(formula_uid)?;
        let expected_owner_uid = codec::UuidSnapshot::new(
            formula_uid
                .lower()
                .checked_add(4)
                .ok_or(BodyTableHiddenAxesError::InvalidSource)?,
            formula_uid.upper(),
        );
        validate_uuid(expected_owner_uid)?;
        if owner.owner_uid() != expected_owner_uid || info.hidden_uuid != Some(owner.owner_uid()) {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
    }
    if profile.is_native() {
        validate_native_visible_owner(&model, &info)?;
    }
    validate_dependencies(
        package,
        &objects,
        &target,
        &model,
        info.hidden_uuid,
        &row_indices,
        &column_indices,
        profile,
        budget,
    )?;
    validate_model_counts(&model, info.hidden_uuid, budget)?;
    let before = hidden_axes(&model, &info, &row_indices, &column_indices, budget)?;
    Ok(Graph {
        profile,
        target,
        model,
        info,
        rows,
        columns,
        row_indices,
        column_indices,
        before,
    })
}

fn message_at(
    package: &Package,
    component_index: usize,
    object_index: usize,
    message_index: usize,
    identifier: NonZeroU64,
    type_: u32,
    versions: &[u32],
) -> Result<RawMessage, BodyTableHiddenAxesError> {
    let object = package
        .state
        .source
        .components()
        .get_index(component_index)
        .and_then(|c| c.archive().objects.get(object_index))
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if object.archive_info.identifier != Some(identifier.get()) {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let message = object
        .messages
        .get(message_index)
        .filter(|m| m.type_ == type_)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    validate_message_metadata_with_versions(object, message_index, type_, versions)?;
    Ok(message.clone())
}

fn classify_graph_profile(
    package: &Package,
    target: &table_lock::BodyTableTarget,
) -> Result<GraphProfile, BodyTableHiddenAxesError> {
    let info_object = package
        .state
        .source
        .components()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let info = info_object
        .archive_info
        .message_infos
        .get(target.info_message_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let model_object = package
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .and_then(|component| component.archive().objects.get(target.model_object_index))
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let model = model_object
        .archive_info
        .message_infos
        .get(target.model_message_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;

    if target.message_type == 6_000
        && target.model_message_type == 6_001
        && info.type_ == 6_000
        && info.versions.as_slice() == CURRENT_MESSAGE_VERSIONS
        && model.type_ == 6_001
        && model.versions.as_slice() == NATIVE_TABLE_MODEL_MESSAGE_VERSIONS
    {
        return Ok(GraphProfile::NativeVisible);
    }

    if info.versions.as_slice() == CURRENT_MESSAGE_VERSIONS
        && model.versions.as_slice() == CURRENT_MESSAGE_VERSIONS
    {
        return Ok(GraphProfile::Indexed);
    }

    Err(BodyTableHiddenAxesError::InvalidSource)
}

/// Validate the physical metadata paired with one selected raw payload.
///
/// `Archive::parse` already uses `MessageInfo.length` to frame each payload,
/// but the semantic owner deliberately repeats the identity/version/header
/// proof at every descendant edge.  This keeps a package assembled from
/// in-memory objects from being admitted merely because its decoded neutral
/// metadata happens to look plausible.
fn validate_message_metadata(
    object: &ArchiveObject,
    message_index: usize,
    expected_type: u32,
) -> Result<&MessageInfo, BodyTableHiddenAxesError> {
    validate_message_metadata_with_versions(
        object,
        message_index,
        expected_type,
        CURRENT_MESSAGE_VERSIONS,
    )
}

fn validate_message_metadata_with_versions<'object>(
    object: &'object ArchiveObject,
    message_index: usize,
    expected_type: u32,
    versions: &[u32],
) -> Result<&'object MessageInfo, BodyTableHiddenAxesError> {
    let message = object
        .messages
        .get(message_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if message.type_ != expected_type
        || info.type_ != message.type_
        || info.versions.as_slice() != versions
        || usize::try_from(info.length).ok() != Some(message.data.len())
        || object.header_length == 0
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    page_layout::validate_selected_metadata(object, message_index)
        .map_err(map_page_layout_error)?;
    Ok(info)
}

/// Prove that an object whose payload has no archive-object edges also has no
/// hidden MessageInfo/FieldInfo edge that could alias a selected dependency.
fn validate_no_reference_metadata(
    object: &ArchiveObject,
    message_index: usize,
    expected_type: u32,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    if object.messages.len() != 1
        || object.archive_info.message_infos.len() != 1
        || message_index != 0
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let info = validate_message_metadata(object, message_index, expected_type)?;
    charge_metadata_scan(info, 1, budget)?;
    if !info.object_references.is_empty()
        || !info.data_references.is_empty()
        || info
            .field_infos
            .iter()
            .any(|field| !field.object_references.is_empty() || !field.data_references.is_empty())
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    Ok(())
}

/// Type-4008 formula-owner records are the sole descendant with a native
/// outbound reference.  Current Pages fixtures carry one aggregate drawable
/// edge; producers that also emit FieldInfo must declare that edge exactly at
/// protobuf field 11.  A completely empty FieldInfo list is the established
/// native aggregate-only form and remains admissible.
fn validate_formula_owner_metadata(
    object: &ArchiveObject,
    message_index: usize,
    drawable_identifier: u64,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    if object.messages.len() != 1
        || object.archive_info.message_infos.len() != 1
        || message_index != 0
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let info = validate_message_metadata(object, message_index, FORMULA_OWNER_MESSAGE_TYPE)?;
    charge_metadata_scan(info, 2, budget)?;
    if info.object_references.as_slice() != [drawable_identifier]
        || !info.data_references.is_empty()
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let mut field_occurrences = 0usize;
    for field in &info.field_infos {
        if !field.data_references.is_empty() {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        if field.object_references.is_empty() {
            continue;
        }
        if field.object_references.as_slice() != [drawable_identifier]
            || field.path.as_slice() != FORMULA_OWNER_REFERENCE_PATH
        {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        field_occurrences = field_occurrences
            .checked_add(1)
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    }
    if !info.field_infos.is_empty() && field_occurrences != 1 {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    Ok(())
}

/// Native Pages' type-4008 owner keeps the selected drawable edge in the
/// payload while omitting both aggregate and FieldInfo declarations.  Accept
/// that omission only for the qualified native profile; a declaration that is
/// present still has to identify the payload-selected drawable at field 11.
fn validate_native_formula_owner_metadata(
    object: &ArchiveObject,
    message_index: usize,
    drawable_identifier: u64,
    versions: &[u32],
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    if object.messages.len() != 1
        || object.archive_info.message_infos.len() != 1
        || message_index != 0
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let info = validate_message_metadata_with_versions(
        object,
        message_index,
        FORMULA_OWNER_MESSAGE_TYPE,
        versions,
    )?;
    charge_metadata_scan(info, 2, budget)?;
    if !info.data_references.is_empty() {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    if info
        .object_references
        .iter()
        .any(|identifier| *identifier != drawable_identifier)
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let aggregate_occurrences = info
        .object_references
        .iter()
        .filter(|identifier| **identifier == drawable_identifier)
        .count();
    if aggregate_occurrences > 1 {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let mut field_occurrences = 0usize;
    for field in &info.field_infos {
        if !field.data_references.is_empty() {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        if field
            .object_references
            .iter()
            .any(|identifier| *identifier != drawable_identifier)
        {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        let occurrences = field
            .object_references
            .iter()
            .filter(|identifier| **identifier == drawable_identifier)
            .count();
        if occurrences != 0 {
            if occurrences != 1 || field.path.as_slice() != FORMULA_OWNER_REFERENCE_PATH {
                return Err(BodyTableHiddenAxesError::InvalidSource);
            }
            field_occurrences = field_occurrences
                .checked_add(1)
                .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        }
        if field.path.as_slice() == FORMULA_OWNER_REFERENCE_PATH
            && !field.object_references.is_empty()
            && (field.object_references.as_slice() != [drawable_identifier]
                || !field.data_references.is_empty())
        {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
    }
    if field_occurrences > 1 || field_occurrences > aggregate_occurrences {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    Ok(())
}

/// Model metadata carries filter-set ownership as aggregate-only edges.  No
/// current Pages field path is proven for these nested extent references, so
/// a FieldInfo declaration (or a data-reference alias) is rejected rather
/// than silently treated as equivalent.
fn validate_aggregate_only_reference_edge(
    inventory: &ReferenceInventory,
    identifier: u64,
) -> Result<(), BodyTableHiddenAxesError> {
    if ReferenceInventory::occurrences(&inventory.object_references, identifier) != 1
        || ReferenceInventory::occurrences(&inventory.data_references, identifier) != 0
        || ReferenceInventory::occurrences(&inventory.field_references, identifier) != 0
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    Ok(())
}

/// Validate the native model's filter-set declarations.  Pages 14.4 emits
/// each filter identifier once in the model aggregate and once in the
/// `hidden_states_owner` field (field 70), while the indexed profile uses
/// aggregate-only declarations.  The path and identifier set are both
/// checked so an unrelated FieldInfo entry cannot be used to bypass the
/// dependency proof. Other model references belong to unrelated table state
/// and remain opaque; they cannot substitute for a selected filter edge.
fn validate_native_filter_metadata(
    info: &MessageInfo,
    filter_identifiers: &[NonZeroU64],
) -> Result<(), BodyTableHiddenAxesError> {
    for identifier in filter_identifiers {
        let identifier = identifier.get();
        if info
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count()
            != 1
            || info.data_references.contains(&identifier)
        {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
    }
    let mut owner_field_count = 0usize;
    let mut owner_field_reference_count = 0usize;
    for field in &info.field_infos {
        let at_hidden_states_owner = field.path.as_slice() == [70];
        if at_hidden_states_owner && !field.data_references.is_empty() {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        if field
            .data_references
            .iter()
            .any(|reference| filter_identifiers.iter().any(|id| id.get() == *reference))
        {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        if at_hidden_states_owner {
            owner_field_count = owner_field_count
                .checked_add(1)
                .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
            owner_field_reference_count = owner_field_reference_count
                .checked_add(field.object_references.len())
                .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        }
        for reference in &field.object_references {
            let is_filter = filter_identifiers
                .iter()
                .any(|identifier| identifier.get() == *reference);
            if is_filter != at_hidden_states_owner {
                return Err(BodyTableHiddenAxesError::InvalidSource);
            }
        }
    }
    if owner_field_count > 1 {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    if owner_field_count == 1 {
        if owner_field_reference_count != filter_identifiers.len() {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        for identifier in filter_identifiers {
            let occurrences = info
                .field_infos
                .iter()
                .filter(|field| field.path.as_slice() == [70])
                .flat_map(|field| field.object_references.iter())
                .filter(|candidate| **candidate == identifier.get())
                .count();
            if occurrences != 1 {
                return Err(BodyTableHiddenAxesError::InvalidSource);
            }
        }
    }
    Ok(())
}

fn validate_reference_metadata(
    package: &Package,
    component_index: usize,
    object_index: usize,
    message_index: usize,
    identifier: NonZeroU64,
    path: &[u32],
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    let object = package
        .state
        .source
        .components()
        .get_index(component_index)
        .and_then(|component| component.archive().objects.get(object_index))
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    charge_metadata_scan(info, 2, budget)?;
    budget
        .charge_payload_work(path.len().saturating_add(1))
        .map_err(map_lock_error)?;
    page_layout::validate_reference_metadata(object, message_index, identifier.get(), path)
        .map_err(map_page_layout_error)?;
    let info =
        validate_message_metadata(object, message_index, object.messages[message_index].type_)?;
    if !info.field_infos.iter().any(|field| {
        field.path.as_slice() == path && field.object_references.as_slice() == [identifier.get()]
    }) || info.data_references.contains(&identifier.get())
        || info
            .field_infos
            .iter()
            .any(|field| field.data_references.contains(&identifier.get()))
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    Ok(())
}

/// Validate one selected payload edge under a producer profile whose
/// `FieldInfo` declarations are optional.  Native Pages 14.4 keeps the
/// aggregate object reference for the model's selected edges but omits the
/// corresponding field-local declarations.  The payload decoder remains the
/// authority for the selected identifier; metadata is accepted only when any
/// declaration that is present agrees with that identifier and path.
fn validate_reference_metadata_for_profile(
    package: &Package,
    component_index: usize,
    object_index: usize,
    message_index: usize,
    identifier: NonZeroU64,
    path: &[u32],
    versions: &[u32],
    profile: GraphProfile,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    if !profile.is_native() {
        return validate_reference_metadata(
            package,
            component_index,
            object_index,
            message_index,
            identifier,
            path,
            budget,
        );
    }
    let object = package
        .state
        .source
        .components()
        .get_index(component_index)
        .and_then(|component| component.archive().objects.get(object_index))
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    charge_metadata_scan(info, 2, budget)?;
    budget
        .charge_payload_work(path.len().saturating_add(1))
        .map_err(map_lock_error)?;
    validate_message_metadata_with_versions(
        object,
        message_index,
        object
            .messages
            .get(message_index)
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?
            .type_,
        versions,
    )?;

    let identifier = identifier.get();
    let aggregate_occurrences = info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count();
    if aggregate_occurrences != 1 || info.data_references.contains(&identifier) {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let mut field_occurrences = 0usize;
    for field in &info.field_infos {
        if field.data_references.contains(&identifier) {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        let occurrences = field
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if field.path.as_slice() == path && !field.data_references.is_empty() {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        if occurrences != 0 {
            if occurrences != 1 || field.path.as_slice() != path {
                return Err(BodyTableHiddenAxesError::InvalidSource);
            }
            field_occurrences = field_occurrences
                .checked_add(1)
                .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        }
        if field.path.as_slice() == path && !field.object_references.is_empty() {
            if field.object_references.as_slice() != [identifier]
                || !field.data_references.is_empty()
            {
                return Err(BodyTableHiddenAxesError::InvalidSource);
            }
        }
    }
    if field_occurrences > 1 || field_occurrences > aggregate_occurrences {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    Ok(())
}

fn reject_reference_identifier_metadata(
    package: &Package,
    component_index: usize,
    object_index: usize,
    message_index: usize,
    identifier: NonZeroU64,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    let object = package
        .state
        .source
        .components()
        .get_index(component_index)
        .and_then(|component| component.archive().objects.get(object_index))
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let info = validate_message_metadata(
        object,
        message_index,
        object
            .messages
            .get(message_index)
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?
            .type_,
    )?;
    charge_metadata_scan(info, 1, budget)?;
    if info.object_references.contains(&identifier.get())
        || info.data_references.contains(&identifier.get())
        || info.field_infos.iter().any(|field| {
            field.object_references.contains(&identifier.get())
                || field.data_references.contains(&identifier.get())
        })
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    Ok(())
}

fn reject_reference_path_metadata(
    package: &Package,
    component_index: usize,
    object_index: usize,
    message_index: usize,
    path: &[u32],
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    let object = package
        .state
        .source
        .components()
        .get_index(component_index)
        .and_then(|component| component.archive().objects.get(object_index))
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let info = validate_message_metadata(
        object,
        message_index,
        object
            .messages
            .get(message_index)
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?
            .type_,
    )?;
    charge_metadata_scan(info, 1, budget)?;
    if info.field_infos.iter().any(|field| {
        field.path.as_slice() == path
            && (!field.object_references.is_empty() || !field.data_references.is_empty())
    }) {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    Ok(())
}

fn validate_reference_shape(
    reference: codec::ReferenceSnapshot,
) -> Result<(), BodyTableHiddenAxesError> {
    if reference.deprecated_type().is_some() || reference.deprecated_is_external().is_some() {
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    Ok(())
}

fn strict_varint(
    source: &[u8],
    number: u32,
    budget: &mut table_lock::WireBudget,
) -> Result<Option<u64>, BodyTableHiddenAxesError> {
    let view = budget.parse(source, 1).map_err(map_lock_error)?;
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        field
            .validate_canonical_framing()
            .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?;
        if field.wire_type() != 0 || found.is_some() {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        let (value, length) = decode_varint_from_bytes(field.payload())
            .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?;
        if length != field.payload().len() || encoded_len(value) != length {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        found = Some(value);
    }
    Ok(found)
}

fn strict_bool(
    source: &[u8],
    number: u32,
    budget: &mut table_lock::WireBudget,
) -> Result<Option<bool>, BodyTableHiddenAxesError> {
    match strict_varint(source, number, budget)? {
        None => Ok(None),
        Some(0) => Ok(Some(false)),
        Some(1) => Ok(Some(true)),
        Some(_) => Err(BodyTableHiddenAxesError::InvalidSource),
    }
}

fn validate_uuid(uuid: codec::UuidSnapshot) -> Result<(), BodyTableHiddenAxesError> {
    if uuid.lower() == 0 && uuid.upper() == 0 {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    Ok(())
}

fn strict_has_field(
    source: &[u8],
    number: u32,
    wire: u8,
    budget: &mut table_lock::WireBudget,
) -> Result<bool, BodyTableHiddenAxesError> {
    let view = budget.parse(source, 1).map_err(map_lock_error)?;
    let mut found = false;
    for field in view.fields().filter(|field| field.number() == number) {
        field
            .validate_canonical_framing()
            .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?;
        if field.wire_type() != wire || found {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        found = true;
    }
    Ok(found)
}

/// Filter rules are opaque to this owner.  Their presence would make a
/// hidden-axis edit depend on unowned rule semantics, even when the selected
/// scalar projection happens to decode successfully.
fn reject_filter_rule_fields(
    source: &[u8],
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    let view = budget.parse(source, 1).map_err(map_lock_error)?;
    for field in view.fields() {
        if !matches!(field.number(), 3 | 6 | 7) {
            continue;
        }
        field
            .validate_canonical_framing()
            .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?;
        let expected_wire = if field.number() == 6 { 0 } else { 2 };
        if field.wire_type() != expected_wire {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    Ok(())
}

fn decode_model(
    source: &[u8],
    budget: &mut table_lock::WireBudget,
) -> Result<ModelValues, BodyTableHiddenAxesError> {
    let options = codec_options(budget, source.len(), source.len())?;
    let (snapshot, report) =
        codec::decode_table_model_with_report(source, options).map_err(map_codec_error)?;
    charge_codec(budget, report)?;
    // Pivot topology is intentionally read but is never changed by this
    // focused writer.  Field 47 is the unrelated merge-owner envelope;
    // `pivot_owner` is the native model field 85.
    let pivot = strict_has_field(source, MODEL_PIVOT_OWNER_FIELD, 2, budget)?;
    Ok(ModelValues {
        rows: snapshot.number_of_rows(),
        columns: snapshot.number_of_columns(),
        map: snapshot
            .base_column_row_uids()
            .map(|reference| reference.identifier()),
        map_ref: snapshot.base_column_row_uids(),
        hidden_rows: snapshot.number_of_hidden_rows(),
        hidden_columns: snapshot.number_of_hidden_columns(),
        filtered_rows: snapshot.number_of_filtered_rows(),
        user_rows: snapshot.number_of_user_hidden_rows(),
        user_columns: snapshot.number_of_user_hidden_columns(),
        formula_columns: snapshot
            .hidden_state_formula_owner_for_columns()
            .map(|reference| reference.identifier()),
        formula_rows: snapshot
            .hidden_state_formula_owner_for_rows()
            .map(|reference| reference.identifier()),
        formula_columns_ref: snapshot.hidden_state_formula_owner_for_columns(),
        formula_rows_ref: snapshot.hidden_state_formula_owner_for_rows(),
        owner: snapshot.hidden_states_owner().cloned(),
        pivot,
    })
}

fn decode_info(
    source: &[u8],
    budget: &mut table_lock::WireBudget,
) -> Result<InfoValues, BodyTableHiddenAxesError> {
    let options = codec_options(budget, source.len(), source.len())?;
    let (snapshot, report) =
        codec::decode_table_info_with_report(source, options).map_err(map_codec_error)?;
    charge_codec(budget, report)?;
    let pivot = strict_bool(source, TABLE_INFO_PIVOT_FIELD, budget)? == Some(true);
    Ok(InfoValues {
        model: snapshot.table_model().identifier(),
        model_ref: snapshot.table_model(),
        map: snapshot
            .view_column_row_uids()
            .map(|reference| reference.identifier()),
        map_ref: snapshot.view_column_row_uids(),
        hidden_uuid: snapshot.hidden_states_uuid(),
        pivot,
    })
}

fn codec_options(
    budget: &table_lock::WireBudget,
    source_len: usize,
    required_output: usize,
) -> Result<codec::DecodeOptions, BodyTableHiddenAxesError> {
    let limits = budget.wire_limits();
    let fields = budget.remaining_wire_fields();
    let work = budget.remaining_wire_work();
    let input = source_len.min(limits.max_input_bytes());
    let output = required_output.min(limits.max_output_bytes());
    // Decoded collections and rewrite bookkeeping can exceed their encoded
    // payload length. These are ceilings, not reservations: the strict codec
    // preflights actual memory before allocation and reports it for charging.
    let retained = limits.max_output_bytes();
    let scratch = limits.max_output_bytes();
    let allocations = fields;
    let states = source_len.min(16_384);
    let recursion = u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX);
    if source_len > limits.max_input_bytes() {
        return Err(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::WireBytes,
            observed: u64::try_from(source_len).unwrap_or(u64::MAX),
            maximum: u64::try_from(limits.max_input_bytes()).unwrap_or(u64::MAX),
        });
    }
    if required_output > limits.max_output_bytes() {
        return Err(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::WireOutputBytes,
            observed: u64::try_from(required_output).unwrap_or(u64::MAX),
            maximum: u64::try_from(limits.max_output_bytes()).unwrap_or(u64::MAX),
        });
    }
    if source_len > output || source_len > retained || source_len > scratch {
        return Err(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::WireOutputBytes,
            observed: u64::try_from(source_len).unwrap_or(u64::MAX),
            maximum: u64::try_from(output).unwrap_or(u64::MAX),
        });
    }
    if source_len != 0 && (fields == 0 || work == 0 || recursion == 0 || states == 0) {
        return Err(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::WireWork,
            observed: 1,
            maximum: u64::try_from(work).unwrap_or(u64::MAX),
        });
    }
    Ok(
        codec::DecodeOptions::new(input, output, fields, work, recursion, states)
            .with_max_allocations(allocations)
            .with_max_retained_bytes(retained)
            .with_max_scratch_bytes(scratch),
    )
}

fn charge_codec(
    budget: &mut table_lock::WireBudget,
    report: codec::DecodeReport,
) -> Result<(), BodyTableHiddenAxesError> {
    let retained_work = report
        .states()
        .checked_add(report.allocations())
        .and_then(|value| value.checked_add(report.retained_bytes()))
        .and_then(|value| value.checked_add(report.scratch_bytes()))
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::TransactionWork,
            observed: u64::MAX,
            maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
        })?;
    budget
        .charge_codec_report(report.fields(), report.work_bytes(), report.max_depth(), 0)
        .map_err(map_lock_error)?;
    budget
        .charge_payload_work(retained_work)
        .map_err(map_lock_error)
}

#[derive(Clone, Copy)]
struct ObjectLocation {
    identifier: u64,
    component_index: usize,
    object_index: usize,
}

#[derive(Clone, Copy)]
struct MessageLocation {
    object: ObjectLocation,
    message_index: usize,
}

/// The reference projections needed by aggregate-only dependency checks.
/// Each vector is sorted once, so checking many filter references remains
/// logarithmic in the declared metadata rather than rescanning every field
/// for every reference.
struct ReferenceInventory {
    object_references: Vec<u64>,
    data_references: Vec<u64>,
    field_references: Vec<u64>,
}

impl ReferenceInventory {
    fn new(
        info: &MessageInfo,
        budget: &mut table_lock::WireBudget,
    ) -> Result<Self, BodyTableHiddenAxesError> {
        let field_reference_count = info
            .field_infos
            .iter()
            .try_fold(0usize, |count, field| {
                count
                    .checked_add(field.object_references.len())
                    .and_then(|value| value.checked_add(field.data_references.len()))
            })
            .ok_or(BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::PayloadReferences,
                observed: u64::MAX,
                maximum: u64::try_from(budget.maximum_payload_references()).unwrap_or(u64::MAX),
            })?;
        let retained_references = info
            .object_references
            .len()
            .checked_add(info.data_references.len())
            .and_then(|count| count.checked_add(field_reference_count))
            .ok_or(BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::PayloadReferences,
                observed: u64::MAX,
                maximum: u64::try_from(budget.maximum_payload_references()).unwrap_or(u64::MAX),
            })?;
        // Source-catalog accounting covers the original metadata. These are
        // additional retained copies, so admit their storage and sorting work
        // before any allocation, copying, or sort takes place.
        budget
            .charge_payload_references(retained_references)
            .map_err(map_lock_error)?;
        let sorting_work = [
            info.object_references.len(),
            info.data_references.len(),
            field_reference_count,
        ]
        .into_iter()
        .try_fold(0usize, |total, count| {
            total
                .checked_add(reference_inventory_sort_work(count)?)
                .ok_or(BodyTableHiddenAxesError::LimitExceeded {
                    kind: BodyTableHiddenAxesLimitKind::TransactionWork,
                    observed: u64::MAX,
                    maximum: u64::try_from(budget.wire_limits().max_rewrite_work())
                        .unwrap_or(u64::MAX),
                })
        })?;
        let work = metadata_scan_work(info)?
            .checked_mul(2)
            .and_then(|scan| scan.checked_add(sorting_work))
            .ok_or(BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::TransactionWork,
                observed: u64::MAX,
                maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
            })?;
        budget.charge_payload_work(work).map_err(map_lock_error)?;
        let mut object_references = Vec::new();
        object_references
            .try_reserve_exact(info.object_references.len())
            .map_err(|_| BodyTableHiddenAxesError::Allocation {
                amount: info.object_references.len(),
            })?;
        object_references.extend_from_slice(&info.object_references);
        let mut data_references = Vec::new();
        data_references
            .try_reserve_exact(info.data_references.len())
            .map_err(|_| BodyTableHiddenAxesError::Allocation {
                amount: info.data_references.len(),
            })?;
        data_references.extend_from_slice(&info.data_references);
        let mut field_references = Vec::new();
        field_references
            .try_reserve_exact(field_reference_count)
            .map_err(|_| BodyTableHiddenAxesError::Allocation {
                amount: field_reference_count,
            })?;
        for field in &info.field_infos {
            field_references.extend_from_slice(&field.object_references);
            field_references.extend_from_slice(&field.data_references);
        }
        object_references.sort_unstable();
        data_references.sort_unstable();
        field_references.sort_unstable();
        Ok(Self {
            object_references,
            data_references,
            field_references,
        })
    }

    fn occurrences(values: &[u64], identifier: u64) -> usize {
        let start = values.partition_point(|value| *value < identifier);
        let end = values.partition_point(|value| *value <= identifier);
        end.saturating_sub(start)
    }
}

fn reference_inventory_sort_work(count: usize) -> Result<usize, BodyTableHiddenAxesError> {
    let levels = if count <= 1 {
        0
    } else {
        (usize::BITS - (count - 1).leading_zeros()) as usize
    };
    count
        .checked_mul(levels)
        .and_then(|value| value.checked_add(count))
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::TransactionWork,
            observed: u64::MAX,
            maximum: u64::MAX,
        })
}

fn metadata_scan_work(info: &MessageInfo) -> Result<usize, BodyTableHiddenAxesError> {
    let mut work = info
        .versions
        .len()
        .checked_add(info.field_infos.len())
        .and_then(|value| value.checked_add(info.object_references.len()))
        .and_then(|value| value.checked_add(info.data_references.len()))
        .and_then(|value| value.checked_add(info.diff_merge_version.len()))
        .and_then(|value| value.checked_add(info.diff_read_version.len()))
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if let Some(path) = info.diff_field_path.as_ref() {
        work = work
            .checked_add(path.as_slice().len())
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    }
    for path in &info.fields_to_remove {
        work = work
            .checked_add(path.as_slice().len())
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    }
    for field in &info.field_infos {
        work = work
            .checked_add(field.path.as_slice().len())
            .and_then(|value| value.checked_add(field.object_references.len()))
            .and_then(|value| value.checked_add(field.data_references.len()))
            .and_then(|value| value.checked_add(field.known_field_version.len()))
            .and_then(|value| {
                value.checked_add(
                    field
                        .known_field_feature_identifier
                        .as_ref()
                        .map_or(0, String::len),
                )
            })
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    }
    Ok(work)
}

fn charge_metadata_scan(
    info: &MessageInfo,
    passes: usize,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    let work = metadata_scan_work(info)?.checked_mul(passes).ok_or(
        BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::TransactionWork,
            observed: u64::MAX,
            maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
        },
    )?;
    budget.charge_payload_work(work).map_err(map_lock_error)
}

fn global_objects(
    package: &Package,
    budget: &mut table_lock::WireBudget,
) -> Result<Vec<ObjectLocation>, BodyTableHiddenAxesError> {
    let object_count = package
        .state
        .source
        .components()
        .iter()
        .try_fold(0usize, |count, component| {
            count.checked_add(component.archive().objects.len())
        })
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::PayloadObjects,
            observed: u64::MAX,
            maximum: u64::MAX,
        })?;
    budget
        .charge_payload_work(object_count)
        .map_err(map_lock_error)?;
    let mut objects = Vec::new();
    objects
        .try_reserve_exact(object_count)
        .map_err(|_| BodyTableHiddenAxesError::Allocation {
            amount: object_count,
        })?;
    for (component_index, component) in package.state.source.components().iter().enumerate() {
        for (object_index, object) in component.archive().objects.iter().enumerate() {
            let id = object
                .archive_info
                .identifier
                .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
            if id == 0 {
                return Err(BodyTableHiddenAxesError::InvalidSource);
            }
            objects.push(ObjectLocation {
                identifier: id,
                component_index,
                object_index,
            });
        }
    }
    objects.sort_unstable_by_key(|entry| entry.identifier);
    if objects
        .windows(2)
        .any(|pair| pair[0].identifier == pair[1].identifier)
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let levels = if objects.len() <= 1 {
        0
    } else {
        (usize::BITS - (objects.len() - 1).leading_zeros()) as usize
    };
    budget
        .charge_payload_work(objects.len().checked_mul(levels).ok_or(
            BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::TransactionWork,
                observed: u64::MAX,
                maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
            },
        )?)
        .map_err(map_lock_error)?;
    Ok(objects)
}

/// Index every type-4008 message once before resolving the selected drawable.
/// The index is deliberately location-only: payload bytes and neutral
/// metadata remain borrowed from the source archive, while the bounded list
/// removes the old package-wide nested rescan from every owner lookup.
fn formula_owner_messages(
    package: &Package,
    objects: &[ObjectLocation],
    budget: &mut table_lock::WireBudget,
) -> Result<Vec<MessageLocation>, BodyTableHiddenAxesError> {
    let mut formula_messages = 0usize;
    for location in objects {
        let object = package
            .state
            .source
            .components()
            .get_index(location.component_index)
            .and_then(|component| component.archive().objects.get(location.object_index))
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        for message in &object.messages {
            // The first pass counts and the second pass materializes the
            // location list. Charge both traversals before reserving the
            // list, so a hostile message inventory cannot be processed
            // without consuming the shared transaction budget.
            budget.charge_payload_work(2).map_err(map_lock_error)?;
            if message.type_ == FORMULA_OWNER_MESSAGE_TYPE {
                formula_messages = formula_messages.checked_add(1).ok_or(
                    BodyTableHiddenAxesError::LimitExceeded {
                        kind: BodyTableHiddenAxesLimitKind::PayloadMessages,
                        observed: u64::MAX,
                        maximum: u64::MAX,
                    },
                )?;
            }
        }
    }
    let mut locations = Vec::new();
    locations.try_reserve_exact(formula_messages).map_err(|_| {
        BodyTableHiddenAxesError::Allocation {
            amount: formula_messages,
        }
    })?;
    for location in objects {
        let object = package
            .state
            .source
            .components()
            .get_index(location.component_index)
            .and_then(|component| component.archive().objects.get(location.object_index))
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        for (message_index, message) in object.messages.iter().enumerate() {
            if message.type_ == FORMULA_OWNER_MESSAGE_TYPE {
                locations.push(MessageLocation {
                    object: *location,
                    message_index,
                });
            }
        }
    }
    Ok(locations)
}

fn object_location(
    objects: &[ObjectLocation],
    identifier: NonZeroU64,
) -> Result<ObjectLocation, BodyTableHiddenAxesError> {
    objects
        .binary_search_by_key(&identifier.get(), |entry| entry.identifier)
        .ok()
        .map(|index| objects[index])
        .ok_or(BodyTableHiddenAxesError::InvalidSource)
}

fn unique_object_location(
    package: &Package,
    identifier: NonZeroU64,
    budget: &mut table_lock::WireBudget,
) -> Result<ObjectLocation, BodyTableHiddenAxesError> {
    let mut result = None;
    for (component_index, component) in package.state.source.components().iter().enumerate() {
        for (object_index, object) in component.archive().objects.iter().enumerate() {
            budget.charge_payload_work(1).map_err(map_lock_error)?;
            if object.archive_info.identifier == Some(identifier.get())
                && result
                    .replace(ObjectLocation {
                        identifier: identifier.get(),
                        component_index,
                        object_index,
                    })
                    .is_some()
            {
                return Err(BodyTableHiddenAxesError::InvalidSource);
            }
        }
    }
    result.ok_or(BodyTableHiddenAxesError::InvalidSource)
}

fn decode_uid_map(
    package: &Package,
    objects: &[ObjectLocation],
    map_id: NonZeroU64,
    model: &ModelValues,
    budget: &mut table_lock::WireBudget,
) -> Result<(Vec<codec::UuidSnapshot>, Vec<codec::UuidSnapshot>), BodyTableHiddenAxesError> {
    let location = object_location(objects, map_id)?;
    decode_uid_map_at(package, location, model, budget)
}

fn decode_uid_map_at(
    package: &Package,
    location: ObjectLocation,
    model: &ModelValues,
    budget: &mut table_lock::WireBudget,
) -> Result<(Vec<codec::UuidSnapshot>, Vec<codec::UuidSnapshot>), BodyTableHiddenAxesError> {
    let object = &package
        .state
        .source
        .components()
        .get_index(location.component_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?
        .archive()
        .objects[location.object_index];
    budget
        .charge_payload_work(object.messages.len())
        .map_err(map_lock_error)?;
    let mut selected = None;
    for (message_index, message) in object.messages.iter().enumerate() {
        if message.type_ == UID_MAP_MESSAGE_TYPE || message.type_ == LEGACY_UID_MAP_MESSAGE_TYPE {
            if selected.replace((message_index, message)).is_some() {
                return Err(BodyTableHiddenAxesError::InvalidSource);
            }
        }
    }
    let (message_index, message) = selected.ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let message_info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    // 6267 is the native Pages route.  A legacy 6200 route is accepted only
    // when its archive metadata explicitly identifies the same qualified
    // message; synthetic message IDs and unqualified legacy payloads are not
    // a supported ingress shape.
    let allow_legacy = message.type_ == LEGACY_UID_MAP_MESSAGE_TYPE
        && message_info.type_ == LEGACY_UID_MAP_MESSAGE_TYPE
        && message_info.versions.as_slice() == CURRENT_MESSAGE_VERSIONS;
    page_layout::validate_selected_metadata(object, message_index)
        .map_err(map_page_layout_error)?;
    codec::validate_column_row_uid_map_message_type(message.type_, allow_legacy)
        .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?;
    validate_no_reference_metadata(object, message_index, message.type_, budget)?;
    budget
        .charge_payload_work(message.data.len())
        .map_err(map_lock_error)?;
    let limits = budget.wire_limits();
    let fields = budget.remaining_wire_fields();
    let work = budget.remaining_wire_work();
    let recursion = u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX);
    let element_count = usize::try_from(model.columns)
        .and_then(|columns| usize::try_from(model.rows).map(|rows| (columns, rows)))
        .ok()
        .and_then(|(columns, rows)| {
            columns
                .checked_mul(3)
                .and_then(|value| rows.checked_mul(3).and_then(|rows| value.checked_add(rows)))
        })
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::PayloadItems,
            observed: u64::MAX,
            maximum: u64::try_from(budget.maximum_payload_references()).unwrap_or(u64::MAX),
        })?;
    if message.data.len() > limits.max_input_bytes() {
        return Err(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::WireBytes,
            observed: u64::try_from(message.data.len()).unwrap_or(u64::MAX),
            maximum: u64::try_from(limits.max_input_bytes()).unwrap_or(u64::MAX),
        });
    }
    if element_count > fields || recursion == 0 {
        return Err(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::WireFields,
            observed: u64::try_from(element_count).unwrap_or(u64::MAX),
            maximum: u64::try_from(fields).unwrap_or(u64::MAX),
        });
    }
    let uid_options = uid_codec::DecodeOptions::new(
        message.data.len(),
        fields,
        work,
        recursion,
        0,
        element_count,
        limits.max_output_bytes(),
        limits.max_output_bytes(),
    );
    let (map, report) = uid_codec::decode_column_row_uid_map(
        message.data.as_slice(),
        usize::try_from(model.columns).map_err(|_| BodyTableHiddenAxesError::InvalidSource)?,
        usize::try_from(model.rows).map_err(|_| BodyTableHiddenAxesError::InvalidSource)?,
        uid_options,
    )
    .map_err(map_uid_codec_error)?;
    budget
        .charge_codec_report(report.fields(), report.work_bytes(), report.max_depth(), 0)
        .map_err(map_lock_error)?;
    budget
        .charge_payload_work(
            report
                .records()
                .checked_add(report.elements())
                .and_then(|value| value.checked_add(element_count))
                .ok_or(BodyTableHiddenAxesError::LimitExceeded {
                    kind: BodyTableHiddenAxesLimitKind::TransactionWork,
                    observed: u64::MAX,
                    maximum: u64::try_from(budget.wire_limits().max_rewrite_work())
                        .unwrap_or(u64::MAX),
                })?,
        )
        .map_err(map_lock_error)?;
    let rows = physical_uids(
        map.sorted_row_uids(),
        map.row_uid_for_index(),
        model.rows,
        budget,
    )?;
    let columns = physical_uids(
        map.sorted_column_uids(),
        map.column_uid_for_index(),
        model.columns,
        budget,
    )?;
    Ok((rows, columns))
}

fn physical_uids(
    sorted: &[uid_codec::Uuid],
    uid_for_index: &[u32],
    expected: u32,
    budget: &mut table_lock::WireBudget,
) -> Result<Vec<codec::UuidSnapshot>, BodyTableHiddenAxesError> {
    let expected =
        usize::try_from(expected).map_err(|_| BodyTableHiddenAxesError::InvalidSource)?;
    if sorted.len() != expected || uid_for_index.len() != expected {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let mut physical = Vec::new();
    physical
        .try_reserve_exact(expected)
        .map_err(|_| BodyTableHiddenAxesError::Allocation { amount: expected })?;
    for stable_index in uid_for_index.iter().copied() {
        let stable_index =
            usize::try_from(stable_index).map_err(|_| BodyTableHiddenAxesError::InvalidSource)?;
        let uid = sorted
            .get(stable_index)
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        let uid = codec::UuidSnapshot::new(uid.lower(), uid.upper());
        validate_uuid(uid)?;
        physical.push(uid);
    }
    budget
        .charge_payload_work(expected)
        .map_err(map_lock_error)?;
    Ok(physical)
}

fn validate_cross_axis_uuid_identity(
    rows: &[codec::UuidSnapshot],
    columns: &[codec::UuidSnapshot],
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    let mut sorted_rows = Vec::new();
    sorted_rows
        .try_reserve_exact(rows.len())
        .map_err(|_| BodyTableHiddenAxesError::Allocation { amount: rows.len() })?;
    sorted_rows.extend_from_slice(rows);
    sorted_rows.sort_unstable_by_key(|uid| (uid.lower(), uid.upper()));
    let row_levels = if sorted_rows.len() <= 1 {
        0
    } else {
        (usize::BITS - (sorted_rows.len() - 1).leading_zeros()) as usize
    };
    let column_levels = if rows.is_empty() || columns.len() <= 1 {
        0
    } else {
        (usize::BITS - (rows.len() - 1).leading_zeros()) as usize
    };
    let work = sorted_rows
        .len()
        .checked_mul(row_levels)
        .and_then(|value| value.checked_add(columns.len().checked_mul(column_levels)?))
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::TransactionWork,
            observed: u64::MAX,
            maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
        })?;
    budget.charge_payload_work(work).map_err(map_lock_error)?;
    if columns.iter().any(|column| {
        sorted_rows
            .binary_search_by_key(&(column.lower(), column.upper()), |row| {
                (row.lower(), row.upper())
            })
            .is_ok()
    }) {
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    Ok(())
}

fn map_uid_codec_error(error: uid_codec::DecodeError) -> BodyTableHiddenAxesError {
    match error.resource_limit() {
        Some(uid_codec::DecodeLimit::Bytes { observed, maximum }) => {
            BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::WireBytes,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        Some(uid_codec::DecodeLimit::Fields { observed, maximum }) => {
            BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::WireFields,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        Some(uid_codec::DecodeLimit::Work { observed, maximum }) => {
            BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::WireWork,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        Some(uid_codec::DecodeLimit::Nesting { observed, maximum }) => {
            BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            }
        },
        Some(uid_codec::DecodeLimit::Allocation { requested }) => {
            BodyTableHiddenAxesError::Allocation { amount: requested }
        },
        Some(uid_codec::DecodeLimit::OutputBytes { observed, maximum })
        | Some(uid_codec::DecodeLimit::ScratchBytes { observed, maximum }) => {
            BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::PayloadBytes,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        Some(uid_codec::DecodeLimit::Records { observed, maximum })
        | Some(uid_codec::DecodeLimit::Elements { observed, maximum }) => {
            BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::PayloadItems,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        None | Some(_) => BodyTableHiddenAxesError::InvalidSource,
    }
}

fn formula_owner_for(
    package: &Package,
    formula_messages: &[MessageLocation],
    drawable: ObjectLocation,
    model: ObjectLocation,
    profile: GraphProfile,
    budget: &mut table_lock::WireBudget,
) -> Result<Option<codec::UuidSnapshot>, BodyTableHiddenAxesError> {
    if profile.is_native() {
        return native_formula_owner_for(
            package,
            formula_messages,
            drawable,
            model,
            profile.formula_owner_versions(),
            budget,
        );
    }
    let mut result = None;
    for location in formula_messages {
        budget.charge_payload_work(1).map_err(map_lock_error)?;
        let object = package
            .state
            .source
            .components()
            .get_index(location.object.component_index)
            .and_then(|component| {
                component
                    .archive()
                    .objects
                    .get(location.object.object_index)
            })
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        if object.archive_info.identifier != Some(location.object.identifier) {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        let message = object
            .messages
            .get(location.message_index)
            .filter(|message| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        // The type-4008 payload is a dependency root, not an arbitrary
        // generated message.  Qualify its paired header before decoding so a
        // redirected/merged helper cannot be mistaken for the selected
        // drawable's owner.
        validate_message_metadata(object, location.message_index, FORMULA_OWNER_MESSAGE_TYPE)?;
        let (owner, report) = codec::decode_formula_owner_dependencies_with_report(
            message.data.as_slice(),
            codec_options(budget, message.data.len(), message.data.len())?,
        )
        .map_err(map_codec_error)?;
        charge_codec(budget, report)?;
        if owner.internal_formula_owner_id() == 0
            || owner.has_dependencies()
            || owner.owner_kind().is_some()
            || owner.base_owner_uid().is_some()
        {
            return Err(BodyTableHiddenAxesError::UnsupportedDependency);
        }
        if let Some(reference) = owner.formula_owner() {
            validate_reference_shape(reference)?;
        }
        if owner
            .formula_owner()
            .is_some_and(|reference| reference.identifier().get() == drawable.identifier)
        {
            // A forged aggregate edge is not sufficient provenance.  The
            // selected 4008 and drawable must be sibling objects in the same
            // component, and a helper cannot be embedded in the drawable
            // object itself.
            if location.object.component_index != drawable.component_index
                || location.object.identifier == drawable.identifier
                || location.object.identifier == model.identifier
            {
                return Err(BodyTableHiddenAxesError::UnsupportedDependency);
            }
            validate_formula_owner_metadata(
                object,
                location.message_index,
                drawable.identifier,
                budget,
            )?;
            if result.replace(owner.formula_owner_uid()).is_some() {
                return Err(BodyTableHiddenAxesError::InvalidSource);
            }
        }
    }
    let result = result.map(|uid| codec::UuidSnapshot::new(uid.lower(), uid.upper()));
    if let Some(uid) = result {
        validate_uuid(uid)?;
    }
    Ok(result)
}

fn native_formula_owner_for(
    package: &Package,
    formula_messages: &[MessageLocation],
    drawable: ObjectLocation,
    model: ObjectLocation,
    versions: &[u32],
    budget: &mut table_lock::WireBudget,
) -> Result<Option<codec::UuidSnapshot>, BodyTableHiddenAxesError> {
    let mut result = None;
    for location in formula_messages {
        budget.charge_payload_work(1).map_err(map_lock_error)?;
        let object = package
            .state
            .source
            .components()
            .get_index(location.object.component_index)
            .and_then(|component| {
                component
                    .archive()
                    .objects
                    .get(location.object.object_index)
            })
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        if object.archive_info.identifier != Some(location.object.identifier) {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        let message = object
            .messages
            .get(location.message_index)
            .filter(|message| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        validate_message_metadata_with_versions(
            object,
            location.message_index,
            FORMULA_OWNER_MESSAGE_TYPE,
            versions,
        )?;
        let (owner, report) = codec::decode_formula_owner_dependencies_with_report(
            message.data.as_slice(),
            codec_options(budget, message.data.len(), message.data.len())?,
        )
        .map_err(map_codec_error)?;
        charge_codec(budget, report)?;
        let Some(reference) = owner.formula_owner() else {
            continue;
        };
        validate_reference_shape(reference)?;
        if reference.identifier().get() != drawable.identifier {
            continue;
        }
        if owner.internal_formula_owner_id() == 0
            || owner.owner_kind() != Some(1)
            || !owner.has_dependencies()
            || owner.base_owner_uid().is_some()
        {
            return Err(BodyTableHiddenAxesError::UnsupportedDependency);
        }
        // A forged aggregate edge is not sufficient provenance.  The
        // selected 4008 and drawable must be sibling objects in the same
        // component, and a helper cannot be embedded in the model object.
        if location.object.component_index != drawable.component_index
            || location.object.identifier == drawable.identifier
            || location.object.identifier == model.identifier
        {
            return Err(BodyTableHiddenAxesError::UnsupportedDependency);
        }
        validate_native_formula_owner_metadata(
            object,
            location.message_index,
            drawable.identifier,
            versions,
            budget,
        )?;
        if result.replace(owner.formula_owner_uid()).is_some() {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
    }
    let result = result.map(|uid| codec::UuidSnapshot::new(uid.lower(), uid.upper()));
    if let Some(uid) = result {
        validate_uuid(uid)?;
    }
    Ok(result)
}

fn validate_dependencies(
    package: &Package,
    objects: &[ObjectLocation],
    target: &table_lock::BodyTableTarget,
    model: &ModelValues,
    active_uuid: Option<codec::UuidSnapshot>,
    row_indices: &UidIndex,
    column_indices: &UidIndex,
    profile: GraphProfile,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    let Some(owner) = model.owner.as_ref() else {
        return Ok(());
    };
    validate_uuid(owner.owner_uid())?;
    if row_indices.index_of(owner.owner_uid()).is_some()
        || column_indices.index_of(owner.owner_uid()).is_some()
    {
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    let active_uuid = active_uuid.ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    validate_uuid(active_uuid)?;
    let mut state_uids = Vec::new();
    state_uids
        .try_reserve_exact(owner.hidden_states().len())
        .map_err(|_| BodyTableHiddenAxesError::Allocation {
            amount: owner.hidden_states().len(),
        })?;
    for state in owner.hidden_states() {
        validate_uuid(state.hidden_states_uid())?;
        state_uids.push(state.hidden_states_uid());
    }
    state_uids.sort_unstable_by_key(|uid| (uid.lower(), uid.upper()));
    if state_uids.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let levels = if state_uids.len() <= 1 {
        0
    } else {
        (usize::BITS - (state_uids.len() - 1).leading_zeros()) as usize
    };
    budget
        .charge_payload_work(state_uids.len().checked_mul(levels).ok_or(
            BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::TransactionWork,
                observed: u64::MAX,
                maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
            },
        )?)
        .map_err(map_lock_error)?;
    let active = owner
        .hidden_states()
        .iter()
        .find(|state| state.hidden_states_uid() == active_uuid)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if owner
        .hidden_states()
        .iter()
        .filter(|state| state.hidden_states_uid() == active_uuid)
        .count()
        != 1
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    if active.column_hidden_state_extent().direction() != codec::AxisDirection::Column
        || active.row_hidden_state_extent().direction() != codec::AxisDirection::Row
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    if active
        .column_hidden_state_extent()
        .needs_to_update_filter_set_for_import()
        == Some(true)
        || active
            .row_hidden_state_extent()
            .needs_to_update_filter_set_for_import()
            == Some(true)
    {
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    for state in owner.hidden_states() {
        validate_uuid(state.hidden_states_uid())?;
        validate_extent_shape(state, row_indices, true, budget)?;
        validate_extent_shape(state, column_indices, false, budget)?;
        let expected_column_uid = codec::UuidSnapshot::new(
            state
                .hidden_states_uid()
                .lower()
                .checked_add(7)
                .ok_or(BodyTableHiddenAxesError::InvalidSource)?,
            state.hidden_states_uid().upper(),
        );
        validate_uuid(expected_column_uid)?;
        if state.row_hidden_state_extent().hidden_state_extent_uid() != state.hidden_states_uid()
            || state.column_hidden_state_extent().hidden_state_extent_uid() != expected_column_uid
        {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        if row_indices
            .index_of(state.row_hidden_state_extent().hidden_state_extent_uid())
            .is_some()
            || column_indices
                .index_of(state.row_hidden_state_extent().hidden_state_extent_uid())
                .is_some()
            || row_indices
                .index_of(state.column_hidden_state_extent().hidden_state_extent_uid())
                .is_some()
            || column_indices
                .index_of(state.column_hidden_state_extent().hidden_state_extent_uid())
                .is_some()
        {
            return Err(BodyTableHiddenAxesError::UnsupportedDependency);
        }
    }
    let expected_column_extent_uid = codec::UuidSnapshot::new(
        owner
            .owner_uid()
            .lower()
            .checked_add(7)
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?,
        owner.owner_uid().upper(),
    );
    if active.hidden_states_uid() != owner.owner_uid()
        || active.row_hidden_state_extent().hidden_state_extent_uid() != owner.owner_uid()
        || active
            .column_hidden_state_extent()
            .hidden_state_extent_uid()
            != expected_column_extent_uid
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    if active
        .column_hidden_state_extent()
        .filter_set()
        .zip(active.row_hidden_state_extent().filter_set())
        .is_some_and(|(column, row)| column.identifier() == row.identifier())
    {
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    validate_extent_filters(package, target, owner, objects, model, profile, budget)?;
    let col_ref = model
        .formula_columns
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let row_ref = model
        .formula_rows
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if col_ref == row_ref {
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    validate_formula_object(
        package,
        objects,
        col_ref,
        active
            .column_hidden_state_extent()
            .hidden_state_extent_uid(),
        budget,
    )?;
    validate_formula_object(
        package,
        objects,
        row_ref,
        active.row_hidden_state_extent().hidden_state_extent_uid(),
        budget,
    )?;
    Ok(())
}

fn validate_extent_shape(
    state: &codec::HiddenStatesSnapshot,
    physical: &UidIndex,
    row: bool,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    let extent = if row {
        state.row_hidden_state_extent()
    } else {
        state.column_hidden_state_extent()
    };
    let expected = if row {
        codec::AxisDirection::Row
    } else {
        codec::AxisDirection::Column
    };
    if extent.direction() != expected {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    validate_uuid(extent.hidden_state_extent_uid())?;
    let mut identities = Vec::new();
    identities
        .try_reserve_exact(extent.base_hidden_states().len())
        .map_err(|_| BodyTableHiddenAxesError::Allocation {
            amount: extent.base_hidden_states().len(),
        })?;
    for hidden_state in extent.base_hidden_states() {
        budget.charge_payload_work(1).map_err(map_lock_error)?;
        validate_uuid(hidden_state.row_or_column_uid())?;
        if physical
            .index_of(hidden_state.row_or_column_uid())
            .is_none()
        {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
        identities.push(hidden_state.row_or_column_uid());
    }
    identities.sort_unstable_by_key(|uid| (uid.lower(), uid.upper()));
    if identities.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let levels = if identities.len() <= 1 {
        0
    } else {
        (usize::BITS - (identities.len() - 1).leading_zeros()) as usize
    };
    budget
        .charge_payload_work(identities.len().checked_mul(levels).ok_or(
            BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::TransactionWork,
                observed: u64::MAX,
                maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
            },
        )?)
        .map_err(map_lock_error)?;
    Ok(())
}

fn validate_extent_filters(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    owner: &codec::HiddenStatesOwnerSnapshot,
    objects: &[ObjectLocation],
    model: &ModelValues,
    profile: GraphProfile,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    let model_location = object_location(objects, target.model_identifier)?;
    let model_object = package
        .state
        .source
        .components()
        .get_index(model_location.component_index)
        .and_then(|component| component.archive().objects.get(model_location.object_index))
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let model_info = validate_message_metadata_with_versions(
        model_object,
        target.model_message_index,
        target.model_message_type,
        profile.model_versions(),
    )?;
    let model_references = ReferenceInventory::new(model_info, budget)?;
    let mut filter_identifiers = Vec::new();
    let filter_capacity = owner.hidden_states().len().checked_mul(2).ok_or(
        BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::PayloadItems,
            observed: u64::MAX,
            maximum: u64::try_from(budget.maximum_payload_references()).unwrap_or(u64::MAX),
        },
    )?;
    filter_identifiers
        .try_reserve_exact(filter_capacity)
        .map_err(|_| BodyTableHiddenAxesError::Allocation {
            amount: filter_capacity,
        })?;
    for state in owner.hidden_states() {
        for extent in [
            state.column_hidden_state_extent(),
            state.row_hidden_state_extent(),
        ] {
            if extent.needs_to_update_filter_set_for_import() == Some(true) {
                return Err(BodyTableHiddenAxesError::UnsupportedDependency);
            }
            if let Some(reference) = extent.filter_set() {
                validate_reference_shape(reference)?;
                filter_identifiers.push(reference.identifier());
            }
        }
    }
    filter_identifiers.sort_unstable_by_key(|identifier| identifier.get());
    if filter_identifiers.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    let levels = if filter_identifiers.len() <= 1 {
        0
    } else {
        (usize::BITS - (filter_identifiers.len() - 1).leading_zeros()) as usize
    };
    budget
        .charge_payload_work(filter_identifiers.len().checked_mul(levels).ok_or(
            BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::TransactionWork,
                observed: u64::MAX,
                maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
            },
        )?)
        .map_err(map_lock_error)?;
    if profile.is_native() {
        // The visible profile has one state and at most two filter edges.
        // Admit its aggregate, field-path, and per-filter consistency scans
        // separately from constructing the sorted reference inventory.
        charge_metadata_scan(model_info, 8, budget)?;
        validate_native_filter_metadata(model_info, &filter_identifiers)?;
    }
    for identifier in filter_identifiers {
        if !profile.is_native() {
            validate_aggregate_only_reference_edge(&model_references, identifier.get())?;
        }
        let location = object_location(objects, identifier)?;
        let object = &package
            .state
            .source
            .components()
            .get_index(location.component_index)
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?
            .archive()
            .objects[location.object_index];
        let messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == FILTER_SET_MESSAGE_TYPE);
        let mut selected = None;
        for message in messages {
            if selected.replace(message).is_some() {
                return Err(BodyTableHiddenAxesError::InvalidSource);
            }
        }
        let message = selected.ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        let message_index = object
            .messages
            .iter()
            .position(|candidate| std::ptr::eq(candidate, message))
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        validate_no_reference_metadata(object, message_index, FILTER_SET_MESSAGE_TYPE, budget)?;
        reject_filter_rule_fields(message.data.as_slice(), budget)?;
        let (filter, report) = codec::decode_filter_set_with_report(
            message.data.as_slice(),
            codec_options(budget, message.data.len(), message.data.len())?,
        )
        .map_err(map_codec_error)?;
        charge_codec(budget, report)?;
        if filter.needs_formula_rewrite_for_import() == Some(true)
            || filter
                .filter_type()
                .is_some_and(|value| !matches!(value, 0 | 1))
        {
            return Err(BodyTableHiddenAxesError::UnsupportedDependency);
        }
        let mut offsets = Vec::new();
        offsets
            .try_reserve_exact(filter.filter_offsets().len())
            .map_err(|_| BodyTableHiddenAxesError::Allocation {
                amount: filter.filter_offsets().len(),
            })?;
        offsets.extend_from_slice(filter.filter_offsets());
        offsets.sort_unstable();
        if offsets.windows(2).any(|pair| pair[0] == pair[1])
            || offsets.iter().any(|offset| *offset >= model.rows)
        {
            return Err(BodyTableHiddenAxesError::UnsupportedDependency);
        }
        budget
            .charge_payload_work(offsets.len())
            .map_err(map_lock_error)?;
    }
    Ok(())
}

fn validate_formula_object(
    package: &Package,
    objects: &[ObjectLocation],
    identifier: NonZeroU64,
    expected_uid: codec::UuidSnapshot,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    let location = object_location(objects, identifier)?;
    let object = &package
        .state
        .source
        .components()
        .get_index(location.component_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?
        .archive()
        .objects[location.object_index];
    let mut selected = None;
    for message in &object.messages {
        if message.type_ == HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE {
            if selected.replace(message).is_some() {
                return Err(BodyTableHiddenAxesError::InvalidSource);
            }
        }
    }
    let message = selected.ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let message_index = object
        .messages
        .iter()
        .position(|candidate| std::ptr::eq(candidate, message))
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    validate_no_reference_metadata(
        object,
        message_index,
        HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE,
        budget,
    )?;
    let (decoded, report) = codec::decode_hidden_state_formula_owner_with_report(
        message.data.as_slice(),
        codec_options(budget, message.data.len(), message.data.len())?,
    )
    .map_err(map_codec_error)?;
    charge_codec(budget, report)?;
    if decoded.needs_to_update_filter_set_for_import() == Some(true)
        || decoded
            .owner_id()
            .is_none_or(|owner| owner.uuid_bytes().is_some())
    {
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    let owner = decoded
        .owner_id()
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let uuid = owner
        .words_uuid()
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if uuid != expected_uid {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    Ok(())
}

fn hidden_axes(
    model: &ModelValues,
    info: &InfoValues,
    row_indices: &UidIndex,
    column_indices: &UidIndex,
    budget: &mut table_lock::WireBudget,
) -> Result<HiddenAxes, BodyTableHiddenAxesError> {
    let Some(owner) = model.owner.as_ref() else {
        return Ok(HiddenAxes::empty());
    };
    let active_uuid = info
        .hidden_uuid
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let active = owner
        .hidden_states()
        .iter()
        .find(|state| state.hidden_states_uid() == active_uuid)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if owner
        .hidden_states()
        .iter()
        .filter(|state| state.hidden_states_uid() == active_uuid)
        .count()
        != 1
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let axis_capacity = active
        .row_hidden_state_extent()
        .base_hidden_states()
        .len()
        .saturating_add(
            active
                .column_hidden_state_extent()
                .base_hidden_states()
                .len(),
        );
    let mut axes = Vec::new();
    axes.try_reserve_exact(axis_capacity)
        .map_err(|_| BodyTableHiddenAxesError::Allocation {
            amount: axis_capacity,
        })?;
    read_extent(
        active.row_hidden_state_extent(),
        row_indices,
        true,
        &mut axes,
        budget,
    )?;
    read_extent(
        active.column_hidden_state_extent(),
        column_indices,
        false,
        &mut axes,
        budget,
    )?;
    budget
        .charge_payload_work(axes.as_slice().len())
        .map_err(map_lock_error)?;
    HiddenAxes::new(axes).map_err(|_| BodyTableHiddenAxesError::InvalidSource)
}

fn read_extent(
    extent: &codec::HiddenStateExtentSnapshot,
    physical: &UidIndex,
    row: bool,
    output: &mut Vec<AxisIndex>,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    let expected = if row {
        codec::AxisDirection::Row
    } else {
        codec::AxisDirection::Column
    };
    if extent.direction() != expected {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    for state in extent.base_hidden_states() {
        budget.charge_payload_work(1).map_err(map_lock_error)?;
        let index = physical
            .index_of(state.row_or_column_uid())
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        if state.user_hidden() == Some(true) {
            output.push(if row {
                AxisIndex::row(index)
            } else {
                AxisIndex::column(index)
            });
        }
    }
    Ok(())
}

fn validate_native_visible_owner(
    model: &ModelValues,
    info: &InfoValues,
) -> Result<(), BodyTableHiddenAxesError> {
    let owner = model
        .owner
        .as_ref()
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if owner.hidden_states().len() != 1
        || info.hidden_uuid != Some(owner.owner_uid())
        || model.formula_columns.is_none()
        || model.formula_rows.is_none()
        || model.pivot
        || info.pivot
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let state = owner
        .hidden_states()
        .first()
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if state
        .row_hidden_state_extent()
        .base_hidden_states()
        .is_empty()
        && state
            .column_hidden_state_extent()
            .base_hidden_states()
            .is_empty()
    {
        Ok(())
    } else {
        Err(BodyTableHiddenAxesError::InvalidSource)
    }
}

fn validate_axis_bounds(graph: &Graph, axes: &HiddenAxes) -> Result<(), BodyTableHiddenAxesError> {
    for axis in axes.iter() {
        let limit = match axis {
            AxisIndex::Row(_) => usize::try_from(graph.model.rows)
                .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?,
            AxisIndex::Column(_) => usize::try_from(graph.model.columns)
                .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?,
        };
        if axis.index() >= limit {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
    }
    Ok(())
}

fn validate_model_counts(
    model: &ModelValues,
    active_uuid: Option<codec::UuidSnapshot>,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    let Some(owner) = model.owner.as_ref() else {
        for value in [
            model.hidden_rows,
            model.hidden_columns,
            model.filtered_rows,
            model.user_rows,
            model.user_columns,
        ]
        .into_iter()
        .flatten()
        {
            if value != 0 {
                return Err(BodyTableHiddenAxesError::InvalidSource);
            }
        }
        return Ok(());
    };
    let active_uuid = active_uuid.ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let active = owner
        .hidden_states()
        .iter()
        .find(|state| state.hidden_states_uid() == active_uuid)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let count_work = active
        .row_hidden_state_extent()
        .base_hidden_states()
        .len()
        .checked_add(
            active
                .column_hidden_state_extent()
                .base_hidden_states()
                .len(),
        )
        .and_then(|count| count.checked_mul(5))
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::TransactionWork,
            observed: u64::MAX,
            maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
        })?;
    budget
        .charge_payload_work(count_work)
        .map_err(map_lock_error)?;
    let total_rows = active
        .row_hidden_state_extent()
        .base_hidden_states()
        .iter()
        .filter(|state| {
            state.user_hidden() == Some(true)
                || state.filtered() == Some(true)
                || state.pivot_hidden() == Some(true)
        })
        .count();
    let total_columns = active
        .column_hidden_state_extent()
        .base_hidden_states()
        .iter()
        .filter(|state| {
            state.user_hidden() == Some(true)
                || state.filtered() == Some(true)
                || state.pivot_hidden() == Some(true)
        })
        .count();
    let user_rows = active
        .row_hidden_state_extent()
        .base_hidden_states()
        .iter()
        .filter(|state| state.user_hidden() == Some(true))
        .count();
    let user_columns = active
        .column_hidden_state_extent()
        .base_hidden_states()
        .iter()
        .filter(|state| state.user_hidden() == Some(true))
        .count();
    let filtered_rows = active
        .row_hidden_state_extent()
        .base_hidden_states()
        .iter()
        .filter(|state| state.filtered() == Some(true))
        .count();
    for (stored, expected) in [
        (model.hidden_rows, total_rows),
        (model.hidden_columns, total_columns),
        (model.user_rows, user_rows),
        (model.user_columns, user_columns),
        (model.filtered_rows, filtered_rows),
    ] {
        if let Some(stored) = stored
            && usize::try_from(stored).ok() != Some(expected)
        {
            return Err(BodyTableHiddenAxesError::InvalidSource);
        }
    }
    Ok(())
}

fn desired_owner(
    graph: &Graph,
    axes: &HiddenAxes,
    budget: &mut table_lock::WireBudget,
) -> Result<(codec::HiddenStatesOwnerSnapshot, codec::UuidSnapshot), BodyTableHiddenAxesError> {
    let owner = graph
        .model
        .owner
        .as_ref()
        .ok_or(BodyTableHiddenAxesError::UnsupportedDependency)?;
    let active_uuid = graph
        .info
        .hidden_uuid
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if owner.hidden_states().len() != 1
        || owner
            .hidden_states()
            .first()
            .is_none_or(|state| state.hidden_states_uid() != active_uuid)
    {
        // The strict owner codec can preserve opaque fields, but it cannot
        // prove byte-exact preservation of an inactive view while changing a
        // sibling.  Reject that shape before allocating a candidate.
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    let state = owner
        .hidden_states()
        .first()
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let column = update_extent(
        state.column_hidden_state_extent(),
        &graph.column_indices,
        &graph.columns,
        axes,
        false,
        budget,
    )?;
    let row = update_extent(
        state.row_hidden_state_extent(),
        &graph.row_indices,
        &graph.rows,
        axes,
        true,
        budget,
    )?;
    let additional_states =
        axes.as_slice()
            .len()
            .checked_mul(2)
            .ok_or(BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::TransactionWork,
                observed: u64::MAX,
                maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
            })?;
    charge_owner_storage(owner, additional_states, budget)?;
    let updated_state = codec::HiddenStatesSnapshot::new(active_uuid, column, row);
    let rebuilt = codec::HiddenStatesOwnerSnapshot::new(owner.owner_uid(), [updated_state])
        .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?;
    Ok((rebuilt, active_uuid))
}

fn update_extent(
    extent: &codec::HiddenStateExtentSnapshot,
    physical: &UidIndex,
    physical_order: &[codec::UuidSnapshot],
    axes: &HiddenAxes,
    row: bool,
    budget: &mut table_lock::WireBudget,
) -> Result<codec::HiddenStateExtentSnapshot, BodyTableHiddenAxesError> {
    let loop_work = extent
        .base_hidden_states()
        .len()
        .checked_add(physical_order.len())
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::TransactionWork,
            observed: u64::MAX,
            maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
        })?;
    budget
        .charge_payload_work(loop_work)
        .map_err(map_lock_error)?;
    let mut existing = Vec::new();
    existing
        .try_reserve_exact(extent.base_hidden_states().len())
        .map_err(|_| BodyTableHiddenAxesError::Allocation {
            amount: extent.base_hidden_states().len(),
        })?;
    for state in extent.base_hidden_states() {
        validate_uuid(state.row_or_column_uid())?;
        existing.push(state.row_or_column_uid());
    }
    existing.sort_unstable_by_key(|uid| (uid.lower(), uid.upper()));
    if existing.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let levels = if existing.len() <= 1 {
        0
    } else {
        (usize::BITS - (existing.len() - 1).leading_zeros()) as usize
    };
    budget
        .charge_payload_work(existing.len().checked_mul(levels).ok_or(
            BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::TransactionWork,
                observed: u64::MAX,
                maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
            },
        )?)
        .map_err(map_lock_error)?;
    let mut states = Vec::new();
    states
        .try_reserve_exact(extent.base_hidden_states().len())
        .map_err(|_| BodyTableHiddenAxesError::Allocation {
            amount: extent.base_hidden_states().len(),
        })?;
    for state in extent.base_hidden_states() {
        let index = physical
            .index_of(state.row_or_column_uid())
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        let hidden = axes.contains(if row {
            AxisIndex::row(index)
        } else {
            AxisIndex::column(index)
        });
        states.push(
            codec::RowOrColumnStateSnapshot::new(state.row_or_column_uid())
                .with_user_hidden(if hidden {
                    Some(true)
                } else {
                    state.user_hidden().filter(|value| !*value)
                })
                .with_filtered(state.filtered())
                .with_pivot_hidden(state.pivot_hidden()),
        );
    }
    for (index, uid) in physical_order.iter().copied().enumerate() {
        if axes.contains(if row {
            AxisIndex::row(index)
        } else {
            AxisIndex::column(index)
        }) && existing
            .binary_search_by_key(&(uid.lower(), uid.upper()), |candidate| {
                (candidate.lower(), candidate.upper())
            })
            .is_err()
        {
            states
                .try_reserve(1)
                .map_err(|_| BodyTableHiddenAxesError::Allocation { amount: 1 })?;
            states.push(codec::RowOrColumnStateSnapshot::new(uid).with_user_hidden(Some(true)));
        }
    }
    codec::HiddenStateExtentSnapshot::new(
        extent.hidden_state_extent_uid(),
        extent.direction(),
        states,
    )
    .map(|value| {
        value
            .with_needs_to_update_filter_set_for_import(
                extent.needs_to_update_filter_set_for_import(),
            )
            .with_filter_set(extent.filter_set())
    })
    .map_err(|_| BodyTableHiddenAxesError::InvalidSource)
}

fn archive_source_length(archive: &Archive) -> Result<usize, BodyTableHiddenAxesError> {
    archive.objects.iter().try_fold(0usize, |end, object| {
        let object_end = usize::try_from(object.header_offset)
            .ok()
            .and_then(|offset| offset.checked_add(usize::try_from(object.header_length).ok()?))
            .and_then(|offset| offset.checked_add(usize::try_from(object.data_length).ok()?))
            .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
        Ok(end.max(object_end))
    })
}

fn charge_owner_storage(
    owner: &codec::HiddenStatesOwnerSnapshot,
    additional: usize,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    let mut count = owner.hidden_states().len();
    for state in owner.hidden_states() {
        count = count
            .checked_add(2)
            .and_then(|value| {
                value.checked_add(
                    state
                        .column_hidden_state_extent()
                        .base_hidden_states()
                        .len(),
                )
            })
            .and_then(|value| {
                value.checked_add(state.row_hidden_state_extent().base_hidden_states().len())
            })
            .ok_or(BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::TransactionWork,
                observed: u64::MAX,
                maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
            })?;
    }
    count = count
        .checked_add(additional)
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::TransactionWork,
            observed: u64::MAX,
            maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
        })?;
    budget.charge_payload_work(count).map_err(map_lock_error)
}

fn charge_rewrite_requirements(
    budget: &mut table_lock::WireBudget,
    requirements: codec::RewriteExecutionRequirements,
) -> Result<(), BodyTableHiddenAxesError> {
    let retained_work = requirements
        .states()
        .checked_add(requirements.allocations())
        .and_then(|value| value.checked_add(requirements.retained_bytes()))
        .and_then(|value| value.checked_add(requirements.scratch_bytes()))
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::TransactionWork,
            observed: u64::MAX,
            maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
        })?;
    budget
        .charge_output_bytes(requirements.output_bytes())
        .and_then(|_| {
            budget.charge_codec_report(
                requirements.fields(),
                requirements.work_bytes(),
                requirements.max_depth(),
                0,
            )
        })
        .and_then(|_| budget.charge_payload_work(retained_work))
        .map_err(map_lock_error)
}

fn desired_model_snapshot(
    graph: &Graph,
    owner: &codec::HiddenStatesOwnerSnapshot,
    active_uuid: codec::UuidSnapshot,
    budget: &mut table_lock::WireBudget,
) -> Result<codec::TableModelSnapshot, BodyTableHiddenAxesError> {
    let active = owner
        .hidden_states()
        .first()
        .filter(|state| state.hidden_states_uid() == active_uuid)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    charge_owner_storage(owner, 0, budget)?;
    let count_work = active
        .row_hidden_state_extent()
        .base_hidden_states()
        .len()
        .checked_add(
            active
                .column_hidden_state_extent()
                .base_hidden_states()
                .len(),
        )
        .and_then(|count| count.checked_mul(5))
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::TransactionWork,
            observed: u64::MAX,
            maximum: u64::try_from(budget.wire_limits().max_rewrite_work()).unwrap_or(u64::MAX),
        })?;
    budget
        .charge_payload_work(count_work)
        .map_err(map_lock_error)?;
    let total_rows = active
        .row_hidden_state_extent()
        .base_hidden_states()
        .iter()
        .filter(|state| {
            state.user_hidden() == Some(true)
                || state.filtered() == Some(true)
                || state.pivot_hidden() == Some(true)
        })
        .count();
    let total_columns = active
        .column_hidden_state_extent()
        .base_hidden_states()
        .iter()
        .filter(|state| {
            state.user_hidden() == Some(true)
                || state.filtered() == Some(true)
                || state.pivot_hidden() == Some(true)
        })
        .count();
    let user_rows = active
        .row_hidden_state_extent()
        .base_hidden_states()
        .iter()
        .filter(|state| state.user_hidden() == Some(true))
        .count();
    let user_columns = active
        .column_hidden_state_extent()
        .base_hidden_states()
        .iter()
        .filter(|state| state.user_hidden() == Some(true))
        .count();
    let filtered_rows = active
        .row_hidden_state_extent()
        .base_hidden_states()
        .iter()
        .filter(|state| state.filtered() == Some(true))
        .count();
    Ok(
        codec::TableModelSnapshot::new(graph.model.rows, graph.model.columns)
            .with_number_of_hidden_rows(
                graph
                    .model
                    .hidden_rows
                    .map(|_| u32::try_from(total_rows))
                    .transpose()
                    .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?,
            )
            .with_number_of_hidden_columns(
                graph
                    .model
                    .hidden_columns
                    .map(|_| u32::try_from(total_columns))
                    .transpose()
                    .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?,
            )
            .with_number_of_filtered_rows(
                graph
                    .model
                    .filtered_rows
                    .map(|_| u32::try_from(filtered_rows))
                    .transpose()
                    .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?,
            )
            .with_number_of_user_hidden_rows(
                graph
                    .model
                    .user_rows
                    .map(|_| u32::try_from(user_rows))
                    .transpose()
                    .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?,
            )
            .with_number_of_user_hidden_columns(
                graph
                    .model
                    .user_columns
                    .map(|_| u32::try_from(user_columns))
                    .transpose()
                    .map_err(|_| BodyTableHiddenAxesError::InvalidSource)?,
            )
            .with_hidden_state_formula_owner_for_columns(graph.model.formula_columns_ref)
            .with_hidden_state_formula_owner_for_rows(graph.model.formula_rows_ref)
            .with_base_column_row_uids(graph.model.map_ref)
            .with_hidden_states_owner(Some(owner.clone())),
    )
}

fn rewrite(
    source: &Package,
    graph: &Graph,
    axes: &HiddenAxes,
    budget: &mut table_lock::WireBudget,
) -> Result<Package, BodyTableHiddenAxesError> {
    let component = source
        .state
        .source
        .components()
        .get_index(graph.target.model_component_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if graph.target.component_index != graph.target.model_component_index {
        return Err(BodyTableHiddenAxesError::UnsupportedDependency);
    }
    let component_name = component.name();
    let entry = source
        .state
        .source
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(BodyTableHiddenAxesError::UnsupportedSource);
    }
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let (owner, active_uuid) = desired_owner(graph, axes, budget)?;
    let desired_model = desired_model_snapshot(graph, &owner, active_uuid, budget)?;
    let desired_info = codec::TableInfoSnapshot::new(graph.info.model_ref)
        .with_view_column_row_uids(graph.info.map_ref)
        .with_hidden_states_uuid(Some(active_uuid));
    let source_archive = component.archive();
    let model_message = source_archive
        .objects
        .get(graph.target.model_object_index)
        .and_then(|object| object.messages.get(graph.target.model_message_index))
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let info_message = source_archive
        .objects
        .get(graph.target.object_index)
        .and_then(|object| object.messages.get(graph.target.info_message_index))
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if source_archive
        .objects
        .get(graph.target.model_object_index)
        .and_then(|object| object.archive_info.identifier)
        != Some(graph.target.model_identifier.get())
        || source_archive
            .objects
            .get(graph.target.object_index)
            .and_then(|object| object.archive_info.identifier)
            != Some(graph.target.drawable_identifier.get())
    {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let model_prepared = codec::prepare_table_model_rewrite(
        model_message.data.as_slice(),
        &desired_model,
        codec_options(
            budget,
            model_message.data.len(),
            model_message.data.len().checked_add(512).ok_or(
                BodyTableHiddenAxesError::LimitExceeded {
                    kind: BodyTableHiddenAxesLimitKind::WireOutputBytes,
                    observed: u64::MAX,
                    maximum: source.state.source.limits().max_input_bytes(),
                },
            )?,
        )?,
    )
    .map_err(map_codec_error)?;
    let info_prepared = codec::prepare_table_info_rewrite(
        info_message.data.as_slice(),
        &desired_info,
        codec_options(
            budget,
            info_message.data.len(),
            info_message.data.len().checked_add(128).ok_or(
                BodyTableHiddenAxesError::LimitExceeded {
                    kind: BodyTableHiddenAxesLimitKind::WireOutputBytes,
                    observed: u64::MAX,
                    maximum: source.state.source.limits().max_input_bytes(),
                },
            )?,
        )?,
    )
    .map_err(map_codec_error)?;
    let model_requirements = model_prepared.execution_requirements();
    let info_requirements = info_prepared.execution_requirements();
    charge_rewrite_requirements(budget, model_requirements)?;
    charge_rewrite_requirements(budget, info_requirements)?;

    let stream_length = archive_source_length(source_archive)?;
    let old_payload_length = model_message
        .data
        .len()
        .checked_add(info_message.data.len())
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let new_payload_length = model_requirements
        .output_bytes()
        .checked_add(info_requirements.output_bytes())
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let rewritten_bound = stream_length
        .checked_sub(old_payload_length)
        .and_then(|value| value.checked_add(new_payload_length))
        .and_then(|value| value.checked_add(64))
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::OutputBytes,
            observed: u64::MAX,
            maximum: source.state.source.limits().max_input_bytes(),
        })?;
    let compressed_bound = table_lock::snappy_compressed_bound(rewritten_bound).ok_or(
        BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::OutputBytes,
            observed: u64::MAX,
            maximum: source.state.source.limits().max_input_bytes(),
        },
    )?;
    let replacement_compressed_bound = match entry.metadata().central().compression_method() {
        0 => compressed_bound,
        8 => table_lock::deflate_compressed_bound(compressed_bound).ok_or(
            BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::OutputBytes,
                observed: u64::MAX,
                maximum: source.state.source.limits().max_input_bytes(),
            },
        )?,
        _ => return Err(BodyTableHiddenAxesError::UnsupportedSource),
    };
    let old_compressed_size =
        usize::try_from(entry.metadata().compressed_size()).map_err(|_| {
            BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::EntryBytes,
                observed: u64::MAX,
                maximum: source.state.source.limits().max_entry_bytes(),
            }
        })?;
    let package_output_bound = source
        .state
        .source
        .source_bytes()
        .len()
        .checked_sub(old_compressed_size)
        .and_then(|value| value.checked_add(replacement_compressed_bound))
        .ok_or(BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::OutputBytes,
            observed: u64::MAX,
            maximum: source.state.source.limits().max_input_bytes(),
        })?;
    budget
        .charge_output_bytes(rewritten_bound)
        .and_then(|_| budget.charge_output_bytes(compressed_bound))
        .and_then(|_| budget.charge_output_bytes(replacement_compressed_bound))
        .and_then(|_| budget.charge_output_bytes(package_output_bound))
        .and_then(|_| budget.charge_payload_bytes(rewritten_bound))
        .and_then(|_| budget.charge_total_payload_bytes(rewritten_bound))
        .and_then(|_| budget.charge_payload_work(rewritten_bound))
        .and_then(|_| budget.charge_payload_work(compressed_bound))
        .and_then(|_| budget.charge_payload_work(replacement_compressed_bound))
        .and_then(|_| budget.charge_payload_work(package_output_bound))
        .and_then(|_| budget.charge_payload_work(info_requirements.output_bytes()))
        .map_err(map_lock_error)?;
    budget
        .precharge_candidate_reopen(
            &source.state.source,
            package_output_bound,
            graph.target.model_component_index,
            compressed_bound,
            rewritten_bound,
            graph.target.model_object_index,
            graph.target.model_message_index,
            model_requirements.output_bytes(),
        )
        .map_err(map_lock_error)?;

    let mut archive = page_layout::editable_archive(source, component_name)
        .map_err(map_page_layout_error)?
        .0;
    let model_object = archive
        .objects
        .get_mut(graph.target.model_object_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    if model_object.archive_info.identifier != Some(graph.target.model_identifier.get()) {
        return Err(BodyTableHiddenAxesError::InvalidSource);
    }
    let new_model = model_prepared
        .execute(codec::RewriteExecutionLimits::exact(model_requirements))
        .map_err(map_codec_error)?
        .into_bytes();
    model_object
        .replace_message_preserving_header_with_limits(
            graph.target.model_message_index,
            RawMessage {
                type_: graph.target.model_message_type,
                data: new_model,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let info_object = archive
        .objects
        .get_mut(graph.target.object_index)
        .ok_or(BodyTableHiddenAxesError::InvalidSource)?;
    let new_info = info_prepared
        .execute(codec::RewriteExecutionLimits::exact(info_requirements))
        .map_err(map_codec_error)?
        .into_bytes();
    info_object
        .replace_message_preserving_header_with_limits(
            graph.target.info_message_index,
            RawMessage {
                type_: graph.target.message_type,
                data: new_info,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let compressed =
        page_layout::compress_archive(archive, archive_limits).map_err(map_page_layout_error)?;
    if compressed.len() > compressed_bound {
        return Err(BodyTableHiddenAxesError::Verification);
    }
    let mut previews = Vec::new();
    previews
        .try_reserve_exact(page_layout::PREVIEW_ENTRY_NAMES.len())
        .map_err(|_| BodyTableHiddenAxesError::Allocation {
            amount: page_layout::PREVIEW_ENTRY_NAMES.len(),
        })?;
    for name in page_layout::PREVIEW_ENTRY_NAMES.iter().copied() {
        if source
            .state
            .source
            .package()
            .iter()
            .any(|entry| entry.name() == name)
        {
            previews.push(name);
        }
    }
    let edits = [EntryEdit::new(component_name, &compressed)];
    let prepared = source
        .state
        .source
        .package()
        .prepare_reassembly_with_deletions(&edits, &previews, source.state.source.limits())
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    budget
        .charge_output_bytes(prepared.output_bytes())
        .and_then(|_| budget.charge_payload_work(requirements.scratch_bytes()))
        .and_then(|_| budget.charge_payload_work(requirements.allocations()))
        .map_err(map_lock_error)?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    if output.len() > package_output_bound {
        return Err(BodyTableHiddenAxesError::Verification);
    }
    let candidate_source =
        SourceCatalog::from_shared_bytes_with_limits(output.into(), source.state.source.limits())
            .map_err(map_archive_error)?;
    Package::from_source_catalog(candidate_source).map_err(map_package_error)
}

fn preview_count(package: &Package) -> usize {
    page_layout::PREVIEW_ENTRY_NAMES
        .iter()
        .filter(|name| {
            package
                .state
                .source
                .package()
                .iter()
                .any(|entry| entry.name() == **name)
        })
        .count()
}

fn reopen(
    source: &Package,
    target: Arc<[u8]>,
    budget: &mut table_lock::WireBudget,
) -> Result<Package, BodyTableHiddenAxesError> {
    budget
        .charge_input_source(target.as_ref())
        .map_err(map_lock_error)?;
    let catalog =
        SourceCatalog::from_shared_bytes_with_limits(target, source.state.source.limits())
            .map_err(map_archive_error)?;
    budget
        .charge_source_catalog(&catalog)
        .map_err(map_lock_error)?;
    table_lock::charge_reopen_work(&catalog, budget).map_err(map_lock_error)?;
    Package::from_source_catalog(catalog).map_err(map_package_error)
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    target: &table_lock::BodyTableTarget,
    touched_components: usize,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableHiddenAxesError> {
    if touched_components != 1
        || source.state.source.components().len() != candidate.state.source.components().len()
        || target.model_identifier == target.drawable_identifier
    {
        return Err(BodyTableHiddenAxesError::Verification);
    }
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    // Exact inverse patches restore the source previews. Compare the same
    // non-preview member sequence in both directions; changed commits enforce
    // preview invalidation before reaching this symmetric locality check.
    let mut candidate_entries = candidate
        .state
        .source
        .package()
        .iter()
        .filter(|entry| !page_layout::PREVIEW_ENTRY_NAMES.contains(&entry.name()));
    for before_entry in source.state.source.package().iter() {
        budget
            .charge_payload_work(
                before_entry
                    .data()
                    .len()
                    .checked_add(before_entry.raw_name().len())
                    .and_then(|value| {
                        value.checked_add(before_entry.metadata().local().extra().len())
                    })
                    .and_then(|value| {
                        value.checked_add(before_entry.metadata().central().extra().len())
                    })
                    .ok_or(BodyTableHiddenAxesError::Verification)?,
            )
            .map_err(map_lock_error)?;
        if page_layout::PREVIEW_ENTRY_NAMES
            .iter()
            .any(|name| *name == before_entry.name())
        {
            continue;
        }
        let after_entry = candidate_entries
            .next()
            .ok_or(BodyTableHiddenAxesError::Verification)?;
        if before_entry.name() != after_entry.name()
            || before_entry.raw_name() != after_entry.raw_name()
        {
            return Err(BodyTableHiddenAxesError::Verification);
        }
        let selected_entry = before_entry.name()
            == source
                .state
                .source
                .components()
                .get_index(target.model_component_index)
                .ok_or(BodyTableHiddenAxesError::Verification)?
                .name();
        if selected_entry {
            if after_entry.is_opaque() || !same_entry_static_metadata(before_entry, after_entry) {
                return Err(BodyTableHiddenAxesError::Verification);
            }
        } else if before_entry.data() != after_entry.data()
            || before_entry.metadata() != after_entry.metadata()
        {
            return Err(BodyTableHiddenAxesError::Verification);
        }
    }
    if candidate_entries.next().is_some() {
        return Err(BodyTableHiddenAxesError::Verification);
    }
    let mut changed = 0usize;
    for (component_index, (before, after)) in source
        .state
        .source
        .components()
        .iter()
        .zip(candidate.state.source.components().iter())
        .enumerate()
    {
        if before.name() != after.name() {
            return Err(BodyTableHiddenAxesError::Verification);
        }
        if before.archive().objects.len() != after.archive().objects.len() {
            return Err(BodyTableHiddenAxesError::Verification);
        }
        for (object_index, before_object) in before.archive().objects.iter().enumerate() {
            let id = before_object
                .archive_info
                .identifier
                .ok_or(BodyTableHiddenAxesError::Verification)?;
            let after_object = after
                .archive()
                .objects
                .get(object_index)
                .ok_or(BodyTableHiddenAxesError::Verification)?;
            if after_object.archive_info.identifier != Some(id) {
                return Err(BodyTableHiddenAxesError::Verification);
            }
            let selected = component_index == target.model_component_index
                && (id == target.model_identifier.get() || id == target.drawable_identifier.get());
            if !selected {
                if !before_object.same_content_ignoring_offsets(after_object) {
                    return Err(BodyTableHiddenAxesError::Verification);
                }
            } else {
                changed = changed.saturating_add(1);
                let message_index = if id == target.model_identifier.get() {
                    target.model_message_index
                } else {
                    target.info_message_index
                };
                if !same_object_except_message(
                    before_object,
                    after_object,
                    message_index,
                    archive_limits,
                    budget,
                )? {
                    return Err(BodyTableHiddenAxesError::Verification);
                }
            }
            budget
                .charge_payload_work(object_index.saturating_add(1))
                .map_err(map_lock_error)?;
        }
    }
    if changed != 2 {
        return Err(BodyTableHiddenAxesError::Verification);
    }
    Ok(())
}

fn same_entry_static_metadata(
    before: &litchi_iwa_archive::package::Entry,
    after: &litchi_iwa_archive::package::Entry,
) -> bool {
    before.metadata().local() == after.metadata().local()
        && before.metadata().central() == after.metadata().central()
}

fn same_object_except_message(
    before: &ArchiveObject,
    after: &ArchiveObject,
    selected_message: usize,
    archive_limits: litchi_iwa_core::ArchiveLimits,
    budget: &mut table_lock::WireBudget,
) -> Result<bool, BodyTableHiddenAxesError> {
    if before.archive_info.identifier != after.archive_info.identifier
        || before.archive_info.should_merge != after.archive_info.should_merge
        || before.messages.len() != after.messages.len()
        || before.archive_info.message_infos.len() != after.archive_info.message_infos.len()
        || selected_message >= before.messages.len()
    {
        return Ok(false);
    }
    let before_payload_length = before.messages.iter().try_fold(0usize, |length, message| {
        length.checked_add(message.data.len())
    });
    let after_payload_length = after.messages.iter().try_fold(0usize, |length, message| {
        length.checked_add(message.data.len())
    });
    let expected_after_data_length = before.data_length.checked_sub(
        u64::try_from(before.messages[selected_message].data.len())
            .map_err(|_| BodyTableHiddenAxesError::Verification)?,
    );
    let expected_after_data_length = expected_after_data_length
        .and_then(|length| {
            length.checked_add(u64::try_from(after.messages[selected_message].data.len()).ok()?)
        })
        .ok_or(BodyTableHiddenAxesError::Verification)?;
    if before_payload_length.and_then(|length| u64::try_from(length).ok())
        != Some(before.data_length)
        || after_payload_length.and_then(|length| u64::try_from(length).ok())
            != Some(after.data_length)
        || after.data_length != expected_after_data_length
    {
        return Ok(false);
    }
    for (index, (before_message, after_message)) in before
        .messages
        .iter()
        .zip(after.messages.iter())
        .enumerate()
    {
        if index == selected_message {
            if before_message.type_ != after_message.type_ {
                return Ok(false);
            }
        } else if before_message != after_message {
            return Ok(false);
        }
    }
    for (index, (before_info, after_info)) in before
        .archive_info
        .message_infos
        .iter()
        .zip(after.archive_info.message_infos.iter())
        .enumerate()
    {
        if index == selected_message {
            if !same_message_info_except_length(before_info, after_info) {
                return Ok(false);
            }
        } else if before_info != after_info {
            return Ok(false);
        }
    }
    let after_payload_bytes = after
        .messages
        .iter()
        .try_fold(0usize, |length, message| {
            length.checked_add(message.data.len())
        })
        .ok_or(BodyTableHiddenAxesError::Verification)?;
    let before_selected_bytes = before.messages[selected_message].data.len();
    budget
        .charge_payload_work(
            after_payload_bytes
                .checked_add(before_selected_bytes)
                .ok_or(BodyTableHiddenAxesError::Verification)?,
        )
        .map_err(map_lock_error)?;
    let mut restored = after.clone();
    restored
        .replace_message_preserving_header_with_limits(
            selected_message,
            before.messages[selected_message].clone(),
            archive_limits,
        )
        .map_err(map_core_error)?;
    // `header_length` and `data_length` are source framing provenance.  The
    // reverse splice above proves the private raw/canonical header bytes;
    // restore the original framing values before the exact-content compare so
    // the only tolerated candidate differences are the selected length and
    // its consequent published framing.
    restored.header_length = before.header_length;
    restored.data_length = before.data_length;
    Ok(before.same_content_ignoring_offsets(&restored))
}

fn same_message_info_except_length(before: &MessageInfo, after: &MessageInfo) -> bool {
    before.type_ == after.type_
        && before.versions == after.versions
        && before.field_infos == after.field_infos
        && before.object_references == after.object_references
        && before.data_references == after.data_references
        && before.base_message_index == after.base_message_index
        && before.diff_merge_version == after.diff_merge_version
        && before.diff_field_path == after.diff_field_path
        && before.fields_to_remove == after.fields_to_remove
        && before.diff_read_version == after.diff_read_version
}

fn map_codec_error(error: codec::DecodeError) -> BodyTableHiddenAxesError {
    if let Some(amount) = error.allocation_amount() {
        return BodyTableHiddenAxesError::Allocation { amount };
    }
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            codec::DecodeLimit::InputBytes { observed, maximum } => {
                (BodyTableHiddenAxesLimitKind::WireBytes, observed, maximum)
            },
            codec::DecodeLimit::OutputBytes { observed, maximum } => (
                BodyTableHiddenAxesLimitKind::WireOutputBytes,
                observed,
                maximum,
            ),
            codec::DecodeLimit::Fields { observed, maximum } => {
                (BodyTableHiddenAxesLimitKind::WireFields, observed, maximum)
            },
            codec::DecodeLimit::WorkBytes { observed, maximum } => {
                (BodyTableHiddenAxesLimitKind::WireWork, observed, maximum)
            },
            codec::DecodeLimit::Nesting { observed, maximum } => (
                BodyTableHiddenAxesLimitKind::WireNesting,
                observed as usize,
                maximum as usize,
            ),
            codec::DecodeLimit::States { observed, maximum } => (
                BodyTableHiddenAxesLimitKind::PayloadItems,
                observed,
                maximum,
            ),
            codec::DecodeLimit::Allocations { observed, maximum } => (
                BodyTableHiddenAxesLimitKind::PayloadItems,
                observed,
                maximum,
            ),
            codec::DecodeLimit::RetainedBytes { observed, maximum } => (
                BodyTableHiddenAxesLimitKind::PayloadBytes,
                observed,
                maximum,
            ),
            codec::DecodeLimit::ScratchBytes { observed, maximum } => (
                BodyTableHiddenAxesLimitKind::PayloadBytes,
                observed,
                maximum,
            ),
            _ => return BodyTableHiddenAxesError::InvalidSource,
        };
        return BodyTableHiddenAxesError::LimitExceeded {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    BodyTableHiddenAxesError::InvalidSource
}

fn map_lock_error(error: table_lock::BodyTableLockError) -> BodyTableHiddenAxesError {
    match error {
        table_lock::BodyTableLockError::TableNotFound => BodyTableHiddenAxesError::TableNotFound,
        table_lock::BodyTableLockError::AmbiguousTableName => {
            BodyTableHiddenAxesError::AmbiguousTableName
        },
        table_lock::BodyTableLockError::AmbiguousSelector => {
            BodyTableHiddenAxesError::AmbiguousSelector
        },
        table_lock::BodyTableLockError::UnsupportedSource => {
            BodyTableHiddenAxesError::UnsupportedSource
        },
        table_lock::BodyTableLockError::InvalidSource => BodyTableHiddenAxesError::InvalidSource,
        table_lock::BodyTableLockError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => BodyTableHiddenAxesError::LimitExceeded {
            kind: map_lock_limit(kind),
            observed,
            maximum,
        },
        table_lock::BodyTableLockError::Allocation { amount } => {
            BodyTableHiddenAxesError::Allocation { amount }
        },
        table_lock::BodyTableLockError::Verification => BodyTableHiddenAxesError::Verification,
        table_lock::BodyTableLockError::PatchConflict => BodyTableHiddenAxesError::PatchConflict,
    }
}

fn map_lock_limit(kind: table_lock::BodyTableLockLimitKind) -> BodyTableHiddenAxesLimitKind {
    match kind {
        table_lock::BodyTableLockLimitKind::InputBytes => BodyTableHiddenAxesLimitKind::InputBytes,
        table_lock::BodyTableLockLimitKind::OutputBytes => {
            BodyTableHiddenAxesLimitKind::OutputBytes
        },
        table_lock::BodyTableLockLimitKind::Entries => BodyTableHiddenAxesLimitKind::Entries,
        table_lock::BodyTableLockLimitKind::EntryBytes => BodyTableHiddenAxesLimitKind::EntryBytes,
        table_lock::BodyTableLockLimitKind::TotalEntryBytes => {
            BodyTableHiddenAxesLimitKind::TotalEntryBytes
        },
        table_lock::BodyTableLockLimitKind::PackageBytes => {
            BodyTableHiddenAxesLimitKind::PackageBytes
        },
        table_lock::BodyTableLockLimitKind::PayloadBytes => {
            BodyTableHiddenAxesLimitKind::PayloadBytes
        },
        table_lock::BodyTableLockLimitKind::TotalPayloadBytes => {
            BodyTableHiddenAxesLimitKind::TotalPayloadBytes
        },
        table_lock::BodyTableLockLimitKind::PayloadObjects => {
            BodyTableHiddenAxesLimitKind::PayloadObjects
        },
        table_lock::BodyTableLockLimitKind::PayloadMessages => {
            BodyTableHiddenAxesLimitKind::PayloadMessages
        },
        table_lock::BodyTableLockLimitKind::PayloadItems => {
            BodyTableHiddenAxesLimitKind::PayloadItems
        },
        table_lock::BodyTableLockLimitKind::PayloadReferences => {
            BodyTableHiddenAxesLimitKind::PayloadReferences
        },
        table_lock::BodyTableLockLimitKind::WireBytes => BodyTableHiddenAxesLimitKind::WireBytes,
        table_lock::BodyTableLockLimitKind::WireFields => BodyTableHiddenAxesLimitKind::WireFields,
        table_lock::BodyTableLockLimitKind::WireNesting => {
            BodyTableHiddenAxesLimitKind::WireNesting
        },
        table_lock::BodyTableLockLimitKind::WireWork => BodyTableHiddenAxesLimitKind::WireWork,
    }
}

fn map_page_layout_error(error: page_layout::PageLayoutError) -> BodyTableHiddenAxesError {
    match error {
        page_layout::PageLayoutError::Allocation { amount } => {
            BodyTableHiddenAxesError::Allocation { amount }
        },
        page_layout::PageLayoutError::LimitExceeded {
            observed, maximum, ..
        } => BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::PayloadBytes,
            observed,
            maximum,
        },
        page_layout::PageLayoutError::UnsupportedSource => {
            BodyTableHiddenAxesError::UnsupportedSource
        },
        _ => BodyTableHiddenAxesError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> BodyTableHiddenAxesError {
    match error {
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            BodyTableHiddenAxesError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyTableHiddenAxesError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => {
                    BodyTableHiddenAxesLimitKind::InputBytes
                },
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    BodyTableHiddenAxesLimitKind::OutputBytes
                },
                litchi_iwa_archive::LimitKind::Entries => BodyTableHiddenAxesLimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => {
                    BodyTableHiddenAxesLimitKind::PackageBytes
                },
                litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes => {
                    BodyTableHiddenAxesLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes => {
                    BodyTableHiddenAxesLimitKind::TotalEntryBytes
                },
                litchi_iwa_archive::LimitKind::IwaStreamBytes => {
                    BodyTableHiddenAxesLimitKind::PayloadBytes
                },
                litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    BodyTableHiddenAxesLimitKind::TotalPayloadBytes
                },
            },
            observed,
            maximum,
        },
        _ => BodyTableHiddenAxesError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> BodyTableHiddenAxesError {
    match error {
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            BodyTableHiddenAxesError::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyTableHiddenAxesError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => BodyTableHiddenAxesLimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    BodyTableHiddenAxesLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderFields => {
                    BodyTableHiddenAxesLimitKind::WireFields
                },
                litchi_iwa_core::LimitKind::HeaderNesting => {
                    BodyTableHiddenAxesLimitKind::WireNesting
                },
                _ => BodyTableHiddenAxesLimitKind::PayloadBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        _ => BodyTableHiddenAxesError::InvalidSource,
    }
}

fn map_package_error(error: PackageError) -> BodyTableHiddenAxesError {
    match error {
        PackageError::Archive(error) => map_archive_error(error),
        PackageError::Allocation { amount } => BodyTableHiddenAxesError::Allocation { amount },
        PackageError::ObjectLimit { observed, limit } => BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::PayloadObjects,
            observed: observed as u64,
            maximum: limit as u64,
        },
        PackageError::PayloadLimit { observed, limit } => BodyTableHiddenAxesError::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::PayloadBytes,
            observed: observed as u64,
            maximum: limit as u64,
        },
        _ => BodyTableHiddenAxesError::InvalidSource,
    }
}

#[cfg(test)]
mod reference_inventory_tests {
    use super::*;
    use litchi_iwa_core::{FieldInfo, FieldPath};

    fn budget(references: usize) -> table_lock::WireBudget {
        let archive = litchi_iwa_core::Limits::default()
            .with_archive_bytes(4096)
            .expect("archive byte limit")
            .with_metadata_items(references)
            .expect("metadata item limit");
        let physical = litchi_iwa_archive::Limits::new(4096, 1, 4096, 4096, 4096)
            .expect("physical limits")
            .with_archive_limits(archive)
            .expect("archive limits");
        table_lock::WireBudget::new(physical).expect("wire budget")
    }

    fn metadata() -> MessageInfo {
        let mut info = MessageInfo::new(6001, 0);
        info.object_references = vec![9, 3];
        info.data_references = vec![5];
        let mut field = FieldInfo::new(FieldPath::new(vec![46]));
        field.object_references = vec![3];
        info.field_infos.push(field);
        info
    }

    #[test]
    fn retained_reference_copies_obey_inclusive_and_exceeded_limits() {
        let info = metadata();
        let mut exact = budget(4);
        let inventory = ReferenceInventory::new(&info, &mut exact).expect("inclusive limit");
        assert_eq!(inventory.object_references, [3, 9]);
        assert_eq!(inventory.data_references, [5]);
        assert_eq!(inventory.field_references, [3]);
        assert_eq!(exact.remaining_payload_references(), 0);

        assert!(matches!(
            ReferenceInventory::new(&info, &mut budget(3)),
            Err(BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::PayloadReferences,
                observed: 4,
                maximum: 3,
            })
        ));
        assert_eq!(info.object_references, [9, 3]);
    }

    #[test]
    fn sorting_work_is_admitted_before_building_reference_copies() {
        let info = metadata();
        // Nine metadata items scanned twice, plus sort/copy bounds of 4, 1,
        // and 1 for the three inventories. Retained references use the
        // separate total-work counter.
        let required_work = 9 * 2 + 4 + 1 + 1;
        let mut exact = budget(4);
        let available = 4096 * 16;
        exact
            .charge_payload_work(available - required_work)
            .expect("leave exact work allowance");
        ReferenceInventory::new(&info, &mut exact).expect("inclusive work limit");
        assert!(exact.charge_payload_work(1).is_err());

        let mut exhausted = budget(4);
        exhausted
            .charge_payload_work(available - required_work + 1)
            .expect("leave one fewer work unit");
        assert!(matches!(
            ReferenceInventory::new(&info, &mut exhausted),
            Err(BodyTableHiddenAxesError::LimitExceeded {
                kind: BodyTableHiddenAxesLimitKind::WireWork,
                ..
            })
        ));
    }
}
