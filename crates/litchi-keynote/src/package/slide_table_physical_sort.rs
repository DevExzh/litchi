//! Exact-source physical row sorting for one Keynote slide table.
//!
//! This owner is intentionally narrow.  It admits only the canonical BNC
//! table spine proved by [`super::slide_table_core`], moves complete raw row
//! envelopes, and keeps all package/archive/protobuf details below the
//! public semantic API.  The persisted sort order is read-only input; the
//! operation changes the physical row order and therefore invalidates the
//! root rendering previews.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::match_same_arms,
    clippy::needless_pass_by_value,
    clippy::too_many_arguments,
    clippy::type_complexity,
    reason = "The physical boundary keeps the complete admission and verification pipeline local."
)]

use std::cmp::Ordering as CompareOrdering;
use std::collections::{HashMap, HashSet};
use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{EntryEdit, ExactArtifacts};
use litchi_iwa_common::table::sort::planning::stable_body_row_permutation;
use litchi_iwa_common::{decode_varint_from_bytes, varint::encoded_len};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{
    numbers_table_cell_storage_codec as storage_codec,
    numbers_table_physical_sort_codec as physical_codec,
};
use litchi_numbers_wire::{BncCellView, CachedScalar, StoredValue};

use super::slide_table_core as core;
use super::slide_table_sort_order::SlideTableSortError;
use super::{Package, PayloadLimitKind, ReadError, SemanticLimitKind};
use crate::SlideSelector;
use crate::slide::table::TableSelector;
use crate::slide::table::physical_sort::{
    ColumnIndex, Diagnostics, Error, LimitKind, Order, Path, RowRange, Scope, UnsupportedFeature,
};

type Result<T> = std::result::Result<T, Error>;
type DecodeResult<T> = std::result::Result<T, storage_codec::DecodeError>;

/// Production boundary ratchet consumed by the crate-level migration checks.
///
/// Keep this token in the owner rather than in a test or compatibility bridge:
/// a build that accidentally drops this module must fail the boundary check.
pub const KEYNOTE_PHYSICAL_SORT_OWNER_ACTIVE: bool = true;

const TILE_MESSAGE_TYPE: u32 = 6_002;
const TABLE_DATA_LIST_MESSAGE_TYPE: u32 = 6_005;
const TABLE_DATA_LIST_NATIVE_MESSAGE_TYPE: u32 = 6_201;
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
const COLUMN_ROW_UID_MAP_LEGACY_MESSAGE_TYPE: u32 = 6_200;
const COLUMN_ROW_UID_MAP_MESSAGE_TYPE: u32 = 6_267;
const PHYSICAL_MUTABLE_MESSAGE_TYPES: [u32; 4] = [
    TILE_MESSAGE_TYPE,
    HEADER_BUCKET_MESSAGE_TYPE,
    COLUMN_ROW_UID_MAP_LEGACY_MESSAGE_TYPE,
    COLUMN_ROW_UID_MAP_MESSAGE_TYPE,
];
const HEADER_BUCKET_ROWS: u32 = 65_536;

/// Package-owned immutable edit staged against one exact selected table.
pub struct SlideTablePhysicalSortEdit<'a> {
    source: &'a Package,
    selection: PhysicalSelection,
    budget: core::Budget,
}

impl fmt::Debug for SlideTablePhysicalSortEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideTablePhysicalSortEdit")
            .field("path", &self.selection.path())
            .field("order", &self.selection.order)
            .field("range", &self.selection.range)
            .finish_non_exhaustive()
    }
}

impl SlideTablePhysicalSortEdit<'_> {
    /// Return the body-relative range, when this edit is row-scoped.
    #[must_use]
    pub const fn rows(&self) -> Option<RowRange> {
        self.selection.range
    }

    /// Return the persisted order used by this physical transaction.
    #[must_use]
    pub const fn order(&self) -> &Order {
        &self.selection.order
    }

    /// Return the selected semantic path.
    #[must_use]
    pub fn path(&self) -> Path {
        self.selection.path()
    }

    /// Start the immutable physical transaction.
    pub fn commit(self) -> Result<SlideTablePhysicalSortCommit> {
        let mut budget = self.budget;
        execute_transaction(self.source, &self.selection, &mut budget)
    }
}

/// Exact-source checked reversible physical row-sort patch.
#[derive(Clone, PartialEq, Eq)]
pub struct SlideTablePhysicalSortPatch {
    artifacts: ExactArtifacts,
    selection: PhysicalSelection,
    destination_by_source: Arc<[u32]>,
    moved_rows: usize,
}

impl fmt::Debug for SlideTablePhysicalSortPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideTablePhysicalSortPatch")
            .field("path", &self.selection.path())
            .field("moved_rows", &self.moved_rows)
            .field("source_fingerprint", &self.source_fingerprint())
            .field("target_fingerprint", &self.target_fingerprint())
            .finish_non_exhaustive()
    }
}

impl SlideTablePhysicalSortPatch {
    /// Return the compact exact-source diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the compact exact-target diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return the selected physical-sort path.
    #[must_use]
    pub fn path(&self) -> Path {
        self.selection.path()
    }

    /// Return the number of body rows whose destination changed.
    #[must_use]
    pub const fn moved_rows(&self) -> usize {
        self.moved_rows
    }

    /// Return whether the patch is an exact source byte no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.moved_rows == 0 && self.artifacts.is_byte_noop()
    }

    /// Return the exact inverse patch without copying package bytes.
    #[must_use]
    pub fn inverse(&self) -> Self {
        let inverse = invert_permutation(&self.destination_by_source)
            .expect("a published physical patch always contains a permutation");
        Self {
            artifacts: self.artifacts.inverse(),
            selection: self.selection.clone(),
            destination_by_source: Arc::from(
                inverse
                    .into_iter()
                    .map(|index| {
                        u32::try_from(index)
                            .expect("a published physical patch fits the checked row limit")
                    })
                    .collect::<Vec<_>>(),
            ),
            moved_rows: self.moved_rows,
        }
    }
}

/// Fully verified physical row-sort result.
#[must_use = "a Keynote physical-sort commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideTablePhysicalSortCommit {
    package: Package,
    patch: SlideTablePhysicalSortPatch,
    diagnostics: Diagnostics,
}

impl SlideTablePhysicalSortCommit {
    /// Borrow the reopened package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &SlideTablePhysicalSortPatch {
        &self.patch
    }

    /// Borrow publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

/// Canonical package-facing diagnostics alias.
pub type SlideTablePhysicalSortDiagnostics = Diagnostics;
/// Canonical package-facing error alias.
pub type SlideTablePhysicalSortError = Error;
/// Canonical package-facing finite limit alias.
pub type SlideTablePhysicalSortLimitKind = LimitKind;
/// Canonical package-facing semantic path alias.
pub type SlideTablePhysicalSortPath = Path;

#[derive(Clone, PartialEq, Eq)]
struct PhysicalSelection {
    target: core::Target,
    topology: core::PhysicalTableTopology,
    order: Order,
    scope: Scope,
    range: Option<RowRange>,
}

impl fmt::Debug for PhysicalSelection {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PhysicalSelection")
            .field("path", &self.path())
            .field("order", &self.order)
            .finish_non_exhaustive()
    }
}

impl PhysicalSelection {
    fn path(&self) -> Path {
        match self.range {
            Some(range) => Path::rows(
                self.target.slide_position,
                self.target.table_position,
                range,
            ),
            None => Path::table(self.target.slide_position, self.target.table_position),
        }
    }

    fn body_start(&self) -> Result<usize> {
        usize::try_from(self.topology.header_rows())
            .map_err(|_| Error::InvalidSource { path: self.path() })
    }

    fn body_rows(&self) -> Result<usize> {
        let total = usize::try_from(self.target.rows)
            .map_err(|_| Error::InvalidSource { path: self.path() })?;
        let head = usize::try_from(self.topology.header_rows())
            .map_err(|_| Error::InvalidSource { path: self.path() })?;
        let foot = usize::try_from(self.topology.footer_rows())
            .map_err(|_| Error::InvalidSource { path: self.path() })?;
        total
            .checked_sub(head)
            .and_then(|value| value.checked_sub(foot))
            .ok_or(Error::InvalidSource { path: self.path() })
    }

    fn range_bounds(&self) -> Result<(usize, usize)> {
        let body_rows = self.body_rows()?;
        match self.range {
            Some(range) => Ok((range.start(), range.end())),
            None => Ok((0, body_rows)),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
enum PhysicalScalar {
    Text(Arc<str>),
    Number(f64),
    Boolean(bool),
    Date(f64),
    Duration(f64),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PhysicalScalarKind {
    Text,
    Number,
    Boolean,
    Date,
    Duration,
}

impl PhysicalScalar {
    fn kind(&self) -> PhysicalScalarKind {
        match self {
            Self::Text(_) => PhysicalScalarKind::Text,
            Self::Number(_) => PhysicalScalarKind::Number,
            Self::Boolean(_) => PhysicalScalarKind::Boolean,
            Self::Date(_) => PhysicalScalarKind::Date,
            Self::Duration(_) => PhysicalScalarKind::Duration,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
struct PhysicalScan {
    selected_keys: Vec<Vec<PhysicalScalar>>,
    uid_map: physical_codec::ColumnRowUidMapSnapshot,
    row_header_indices: Vec<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct PhysicalPermutation {
    destination_by_source: Arc<[u32]>,
    moved_rows: usize,
}

impl PhysicalPermutation {
    fn is_identity(&self) -> bool {
        self.moved_rows == 0
    }
}

#[derive(Debug, Default)]
struct StringVisitor {
    entries: Vec<(u32, String)>,
    has_segments: bool,
    has_rich_text_entries: bool,
    has_comment_entries: bool,
    has_unsupported_sidecar: bool,
}

#[derive(Debug, Default)]
struct StringShapeVisitor {
    entry_count: usize,
    text_bytes: usize,
    has_segments: bool,
    has_rich_text_entries: bool,
    has_comment_entries: bool,
    has_unsupported_sidecar: bool,
}

impl StringVisitor {
    fn try_with_capacity(capacity: usize, path: Path) -> Result<Self> {
        let mut visitor = Self::default();
        visitor
            .entries
            .try_reserve_exact(capacity)
            .map_err(|_| Error::Allocation {
                amount: capacity,
                path,
            })?;
        Ok(visitor)
    }
}

impl storage_codec::StorageVisitor for StringShapeVisitor {
    fn visit_list_entry_record(
        &mut self,
        record: storage_codec::TableDataListEntryRecord<'_>,
    ) -> DecodeResult<()> {
        let snapshot = record.snapshot();
        self.entry_count = self
            .entry_count
            .checked_add(1)
            .ok_or_else(|| storage_codec::DecodeError::allocation(1))?;
        if let Some(value) = snapshot.string_value() {
            self.text_bytes = self
                .text_bytes
                .checked_add(value.len())
                .ok_or_else(|| storage_codec::DecodeError::allocation(value.len()))?;
        }
        if snapshot.rich_text_payload().is_some() {
            self.has_rich_text_entries = true;
        }
        if snapshot.comment_storage().is_some() {
            self.has_comment_entries = true;
        }
        self.has_unsupported_sidecar |= snapshot.reference().is_some()
            || snapshot.formula().is_some()
            || snapshot.format().is_some()
            || snapshot.custom_format().is_some()
            || snapshot.import_warning_set().is_some()
            || snapshot.cell_spec().is_some()
            || snapshot.key() == 0
            || snapshot.ref_count() == 0
            || snapshot.string_value().is_none();
        Ok(())
    }

    fn visit_list_segment(
        &mut self,
        _reference: storage_codec::ReferenceRecord<'_>,
    ) -> DecodeResult<()> {
        self.has_segments = true;
        Ok(())
    }
}

impl storage_codec::StorageVisitor for StringVisitor {
    fn visit_list_entry_record(
        &mut self,
        record: storage_codec::TableDataListEntryRecord<'_>,
    ) -> DecodeResult<()> {
        let snapshot = record.snapshot();
        if snapshot.rich_text_payload().is_some() {
            self.has_rich_text_entries = true;
        }
        if snapshot.comment_storage().is_some() {
            self.has_comment_entries = true;
        }
        self.has_unsupported_sidecar |= snapshot.reference().is_some()
            || snapshot.formula().is_some()
            || snapshot.format().is_some()
            || snapshot.custom_format().is_some()
            || snapshot.import_warning_set().is_some()
            || snapshot.cell_spec().is_some()
            || snapshot.key() == 0
            || snapshot.ref_count() == 0
            || snapshot.string_value().is_none();
        if let Some(value) = snapshot.string_value() {
            if self.entries.len() == self.entries.capacity() {
                // The shape pass pre-reserves every possible list entry.  A
                // mismatch here must fail closed rather than grow after the
                // owner's allocation ledger was admitted.
                return Err(storage_codec::DecodeError::allocation(1));
            }
            self.entries.push((snapshot.key(), value.to_owned()));
        }
        Ok(())
    }

    fn visit_list_segment(
        &mut self,
        _reference: storage_codec::ReferenceRecord<'_>,
    ) -> DecodeResult<()> {
        self.has_segments = true;
        Ok(())
    }
}

impl Package {
    /// Execute the persisted entire-table order as a physical row movement.
    ///
    /// Selection is resolved before physical admission, and the persisted
    /// order remains byte-for-byte untouched.  A changed transaction rewrites
    /// the canonical tile/header/UID storage and removes only root previews.
    pub fn execute_slide_table_sort_order<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<SlideTablePhysicalSortCommit> {
        let slide = slide.into();
        let table = table.into();
        let mut budget = core::Budget::new(self).map_err(|error| map_core(error, Path::Package))?;
        let selection = select_physical(self, slide, table, Scope::EntireTable, None, &mut budget)?;
        execute_transaction(self, &selection, &mut budget)
    }

    /// Execute the persisted selected-row order over a body-relative range.
    pub fn execute_slide_table_sort_order_to_rows<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
        rows: RowRange,
    ) -> Result<SlideTablePhysicalSortCommit> {
        let slide = slide.into();
        let table = table.into();
        let mut budget = core::Budget::new(self).map_err(|error| map_core(error, Path::Package))?;
        let selection = select_physical(
            self,
            slide,
            table,
            Scope::SelectedRows,
            Some(rows),
            &mut budget,
        )?;
        execute_transaction(self, &selection, &mut budget)
    }

    /// Start an exact immutable physical-sort edit.
    pub fn edit_slide_table_physical_sort<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<SlideTablePhysicalSortEdit<'_>> {
        let mut budget = core::Budget::new(self).map_err(|error| map_core(error, Path::Package))?;
        let selection = select_physical(
            self,
            slide.into(),
            table.into(),
            Scope::EntireTable,
            None,
            &mut budget,
        )?;
        Ok(SlideTablePhysicalSortEdit {
            source: self,
            selection,
            budget,
        })
    }

    /// Start an exact immutable physical-sort edit for a body-relative range.
    pub fn edit_slide_table_physical_sort_to_rows<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
        rows: RowRange,
    ) -> Result<SlideTablePhysicalSortEdit<'_>> {
        let mut budget = core::Budget::new(self).map_err(|error| map_core(error, Path::Package))?;
        let selection = select_physical(
            self,
            slide.into(),
            table.into(),
            Scope::SelectedRows,
            Some(rows),
            &mut budget,
        )?;
        Ok(SlideTablePhysicalSortEdit {
            source: self,
            selection,
            budget,
        })
    }

    /// Apply an exact-source physical-sort patch and reopen its candidate.
    pub fn apply_slide_table_physical_sort(
        &self,
        patch: &SlideTablePhysicalSortPatch,
    ) -> Result<SlideTablePhysicalSortCommit> {
        let path = patch.selection.path();
        let mut budget = core::Budget::new(self).map_err(|error| map_core(error, path))?;
        let catalog =
            core::physical_catalog(self).map_err(|error| map_core(error, Path::Package))?;
        if !patch.artifacts.authorizes_source(&catalog.shared_source()) {
            return Err(Error::PatchConflict);
        }
        let current = select_physical(
            self,
            SlideSelector::position(patch.selection.target.slide_position),
            TableSelector::position(patch.selection.target.table_position),
            patch.selection.scope,
            patch.selection.range,
            &mut budget,
        )?;
        if current != patch.selection {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read)?;
            return Ok(SlideTablePhysicalSortCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        let target = patch.artifacts.target();
        budget
            .preflight_candidate(self, target.len())
            .map_err(|error| map_core(error, path))?;
        let candidate =
            Package::from_source_with_options(target, self.state.options).map_err(map_read)?;
        candidate.validate().map_err(map_read)?;
        verify_candidate(
            self,
            &candidate,
            &patch.selection,
            &patch.destination_by_source,
            patch.moved_rows,
            &mut budget,
        )?;
        let deleted_previews = preview_deletion_count(self, &candidate);
        Ok(SlideTablePhysicalSortCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(
                patch.moved_rows,
                patch.selection.topology_component_count(),
                deleted_previews,
            ),
        })
    }
}

fn select_physical(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    requested_scope: Scope,
    range: Option<RowRange>,
    budget: &mut core::Budget,
) -> Result<PhysicalSelection> {
    let target = core::select_table(package, slide, table, budget)
        .map_err(|error| map_core(error, Path::Package))?;
    let path = Path::table(target.slide_position, target.table_position);
    let configured = package
        .slide_table_sort_order(slide, table)
        .map_err(|error| map_sort_error(error, path))?
        .ok_or(Error::SortOrderMissing { path })?;
    if configured.scope() != requested_scope {
        return Err(Error::ScopeMismatch {
            path: range.map_or(path, |value| {
                Path::rows(target.slide_position, target.table_position, value)
            }),
            configured: configured.scope(),
            requested: requested_scope,
        });
    }
    validate_order_columns(&configured, target.columns, path)?;
    if target.locked {
        return Err(Error::TableLocked { path });
    }
    let topology = core::admit_physical_table(package, &target, budget)
        .map_err(|error| map_core(error, path))?;
    if topology.has_formula_entries() {
        return Err(Error::UnsupportedFeature {
            path,
            feature: UnsupportedFeature::Formula,
        });
    }
    if topology.has_formula_error_entries() {
        return Err(Error::UnsupportedFeature {
            path,
            feature: UnsupportedFeature::FormulaError,
        });
    }
    if topology.has_conditional_style_entries() {
        return Err(Error::UnsupportedFeature {
            path,
            feature: UnsupportedFeature::ConditionalStyles,
        });
    }
    if topology.has_comment_entries() {
        return Err(Error::UnsupportedFeature {
            path,
            feature: UnsupportedFeature::CommentAnchors,
        });
    }
    if topology.has_rich_text_entries() {
        return Err(Error::UnsupportedFeature {
            path,
            feature: UnsupportedFeature::UnsupportedCell,
        });
    }
    let body_rows = body_row_count(&target, &topology, path)?;
    if let Some(selected) = range {
        if selected.end() > body_rows {
            return Err(Error::RowRangeOutOfBounds {
                path: Path::rows(target.slide_position, target.table_position, selected),
                range: selected,
                body_rows,
            });
        }
    }
    Ok(PhysicalSelection {
        target,
        topology,
        order: configured,
        scope: requested_scope,
        range,
    })
}

fn execute_transaction(
    source: &Package,
    selection: &PhysicalSelection,
    budget: &mut core::Budget,
) -> Result<SlideTablePhysicalSortCommit> {
    let path = selection.path();
    let scan = scan_table(source, selection, budget)?;
    let permutation = plan_permutation(selection, &scan, budget)?;
    if permutation.is_identity() {
        let bytes = core::physical_catalog(source)
            .map_err(|error| map_core(error, path))?
            .shared_source();
        return Ok(SlideTablePhysicalSortCommit {
            package: source.snapshot(),
            patch: SlideTablePhysicalSortPatch {
                artifacts: ExactArtifacts::new(Arc::clone(&bytes), bytes),
                selection: selection.clone(),
                destination_by_source: permutation.destination_by_source,
                moved_rows: 0,
            },
            diagnostics: Diagnostics::unchanged(),
        });
    }
    let catalog = core::physical_catalog(source).map_err(|error| map_core(error, path))?;
    if !catalog.source_is_exact() {
        return Err(Error::UnsupportedSource);
    }
    let previews = super::rendering_invalidation::root_preview_deletions(catalog.package())
        .map_err(|_| Error::Verification { path })?;
    let candidate = rewrite_physical(
        source,
        selection,
        &permutation.destination_by_source,
        previews.names(),
        budget,
    )?;
    verify_candidate(
        source,
        &candidate,
        selection,
        &permutation.destination_by_source,
        permutation.moved_rows,
        budget,
    )?;
    let source_artifact = core::physical_catalog(source)
        .map_err(|error| map_core(error, path))?
        .shared_source();
    let target_artifact = core::physical_catalog(&candidate)
        .map_err(|error| map_core(error, path))?
        .shared_source();
    Ok(SlideTablePhysicalSortCommit {
        package: candidate,
        patch: SlideTablePhysicalSortPatch {
            artifacts: ExactArtifacts::new(source_artifact, Arc::clone(&target_artifact)),
            selection: selection.clone(),
            destination_by_source: permutation.destination_by_source,
            moved_rows: permutation.moved_rows,
        },
        diagnostics: Diagnostics::published(
            permutation.moved_rows,
            selection.topology_component_count(),
            previews.len(),
        ),
    })
}

fn plan_permutation(
    selection: &PhysicalSelection,
    scan: &PhysicalScan,
    budget: &mut core::Budget,
) -> Result<PhysicalPermutation> {
    let path = selection.path();
    let (range_start, range_end) = selection.range_bounds()?;
    let selected_len = range_end
        .checked_sub(range_start)
        .ok_or(Error::InvalidPermutation { path })?;
    if scan.selected_keys.len() != selected_len {
        return Err(Error::InvalidPermutation { path });
    }
    // The common planner owns a source-offset vector and reserves it before
    // returning its requirements.  Charge that storage before entering the
    // planner so a ledger refusal cannot occur after the allocation.
    let planner_bytes = selected_len
        .checked_mul(size_of::<usize>())
        .ok_or(Error::InvalidSource { path })?;
    budget
        .allocations(selected_len)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(planner_bytes)
        .map_err(|error| map_core(error, path))?;
    let planner = stable_body_row_permutation(
        &scan.selected_keys,
        &selection.order,
        PhysicalScalar::kind,
        compare_scalars,
    )
    .map_err(|error| match error {
        litchi_iwa_common::table::sort::planning::Error::MixedScalarKinds { row, rule } => {
            Error::MixedSortKeyKinds {
                path,
                row: Position::new(range_start.saturating_add(row.get())),
                rule: Position::new(rule),
            }
        },
        litchi_iwa_common::table::sort::planning::Error::Allocation { rows } => {
            Error::Allocation { amount: rows, path }
        },
        _ => Error::InvalidPermutation { path },
    })?;
    let body_rows = selection.body_rows()?;
    let total_rows =
        usize::try_from(selection.target.rows).map_err(|_| Error::InvalidSource { path })?;
    let destination_bytes = total_rows
        .checked_mul(size_of::<u32>())
        .ok_or(Error::InvalidSource { path })?;
    budget
        .allocations(total_rows)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(destination_bytes)
        .map_err(|error| map_core(error, path))?;
    let mut destination_by_source = Vec::new();
    destination_by_source
        .try_reserve_exact(total_rows)
        .map_err(|_| Error::Allocation {
            amount: total_rows,
            path,
        })?;
    for index in 0..total_rows {
        destination_by_source
            .push(u32::try_from(index).map_err(|_| Error::InvalidSource { path })?);
    }
    for (destination_offset, source_offset) in planner.sources_by_destination().iter().enumerate() {
        let source_body = range_start
            .checked_add(source_offset.get())
            .ok_or(Error::InvalidPermutation { path })?;
        let destination_body = range_start
            .checked_add(destination_offset)
            .ok_or(Error::InvalidPermutation { path })?;
        if source_body >= body_rows || destination_body >= body_rows {
            return Err(Error::InvalidPermutation { path });
        }
        let source_global = usize::try_from(selection.topology.header_rows())
            .ok()
            .and_then(|head| head.checked_add(source_body))
            .ok_or(Error::InvalidPermutation { path })?;
        let destination_global = usize::try_from(selection.topology.header_rows())
            .ok()
            .and_then(|head| head.checked_add(destination_body))
            .ok_or(Error::InvalidPermutation { path })?;
        destination_by_source[source_global] =
            u32::try_from(destination_global).map_err(|_| Error::InvalidPermutation { path })?;
    }
    let moved_rows = destination_by_source
        .iter()
        .enumerate()
        .filter(|(source, destination)| usize::try_from(**destination).ok() != Some(*source))
        .count();
    if total_rows != 0 {
        // Converting a Vec into an Arc slice may allocate a new Arc control
        // block while the Vec is still live.  Reserve that second event and
        // peak retained storage before the conversion can take place.
        budget
            .allocations(1)
            .map_err(|error| map_core(error, path))?;
        budget
            .retained(destination_bytes)
            .map_err(|error| map_core(error, path))?;
    }
    Ok(PhysicalPermutation {
        destination_by_source: Arc::from(destination_by_source),
        moved_rows,
    })
}

fn scan_table(
    package: &Package,
    selection: &PhysicalSelection,
    budget: &mut core::Budget,
) -> Result<PhysicalScan> {
    let path = selection.path();
    let strings = read_string_table(package, selection, budget)?;
    let (range_start, range_end) = selection.range_bounds()?;
    let selected_len = range_end
        .checked_sub(range_start)
        .ok_or(Error::InvalidPermutation { path })?;
    let total_rows =
        usize::try_from(selection.target.rows).map_err(|_| Error::InvalidSource { path })?;
    let body_start = selection.body_start()?;
    let body_end = body_start
        .checked_add(selection.body_rows()?)
        .ok_or(Error::InvalidSource { path })?;
    let selected_global_start = body_start
        .checked_add(range_start)
        .ok_or(Error::InvalidSource { path })?;
    let selected_global_end = body_start
        .checked_add(range_end)
        .ok_or(Error::InvalidSource { path })?;
    if selected_global_end > body_end || selected_global_end > total_rows {
        return Err(Error::InvalidSource { path });
    }

    let seen_bytes = total_rows
        .checked_mul(size_of::<bool>())
        .ok_or(Error::InvalidSource { path })?;
    budget
        .allocations(total_rows)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(seen_bytes)
        .map_err(|error| map_core(error, path))?;
    let selected_key_slots_bytes = selected_len
        .checked_mul(size_of::<Option<Vec<PhysicalScalar>>>())
        .ok_or(Error::InvalidSource { path })?;
    budget
        .allocations(selected_len)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(selected_key_slots_bytes)
        .map_err(|error| map_core(error, path))?;
    let mut seen = Vec::new();
    seen.try_reserve_exact(total_rows)
        .map_err(|_| Error::Allocation {
            amount: total_rows,
            path,
        })?;
    seen.resize(total_rows, false);
    let mut selected_keys: Vec<Option<Vec<PhysicalScalar>>> = Vec::new();
    selected_keys
        .try_reserve_exact(selected_len)
        .map_err(|_| Error::Allocation {
            amount: selected_len,
            path,
        })?;
    selected_keys.resize_with(selected_len, || None);

    for tile_reference in selection.topology.tile_references() {
        let object = core::object_at(package, &tile_reference.location)
            .map_err(|error| map_core(error, path))?;
        let payload = message_payload(object, TILE_MESSAGE_TYPE, path)?;
        let options = physical_options(
            package,
            budget,
            payload.len(),
            selection.target.rows,
            selection.target.columns,
        )?;
        let prepared = physical_codec::plan_tile_rows_rewrite(
            payload,
            selection.topology.tile_size(),
            &[],
            options,
        )
        .map_err(|error| map_physical_codec(error, path))?;
        charge_rewrite_requirements(prepared.requirements(), budget, path)?;
        let tile = prepared.tile();
        if tile.num_rows() > selection.topology.tile_size()
            || tile.storage_version() != Some(5)
            || tile.last_saved_in_bnc() != Some(true)
        {
            return Err(Error::InvalidSource { path });
        }
        if tile.should_use_wide_rows() != selection.topology.wide_rows() {
            return Err(Error::UnsupportedTopology { path });
        }
        for record in prepared.row_records() {
            let local = record.snapshot().tile_row_index();
            let global = tile_reference
                .tile_id
                .checked_mul(selection.topology.tile_size())
                .and_then(|base| base.checked_add(local))
                .ok_or(Error::InvalidSource { path })?;
            let global_usize =
                usize::try_from(global).map_err(|_| Error::InvalidSource { path })?;
            if global_usize >= total_rows || std::mem::replace(&mut seen[global_usize], true) {
                return Err(Error::InvalidSource { path });
            }
            let values =
                decode_row_cells(record.snapshot(), selection, &strings, global_usize, budget)?;
            if (selected_global_start..selected_global_end).contains(&global_usize) {
                let offset = global_usize
                    .checked_sub(selected_global_start)
                    .ok_or(Error::InvalidSource { path })?;
                let rule_count = selection.order.rules().len();
                let key_bytes = rule_count
                    .checked_mul(size_of::<PhysicalScalar>())
                    .ok_or(Error::InvalidSource { path })?;
                budget
                    .allocations(rule_count)
                    .map_err(|error| map_core(error, path))?;
                budget
                    .retained(key_bytes)
                    .map_err(|error| map_core(error, path))?;
                let mut keys = Vec::new();
                keys.try_reserve_exact(selection.order.rules().len())
                    .map_err(|_| Error::Allocation {
                        amount: selection.order.rules().len(),
                        path,
                    })?;
                for (rule_index, rule) in selection.order.rules().iter().copied().enumerate() {
                    let column = rule.column();
                    let Some(value) = values.get(column.get()) else {
                        return Err(Error::InvalidSource { path });
                    };
                    let Some(value) = value.clone() else {
                        return Err(Error::MissingSortKey {
                            path,
                            row: Position::new(
                                global_usize
                                    .checked_sub(body_start)
                                    .ok_or(Error::InvalidSource { path })?,
                            ),
                            column,
                        });
                    };
                    let _ = rule_index;
                    keys.push(value);
                }
                selected_keys[offset] = Some(keys);
            }
        }
    }
    if seen.iter().any(|present| !present) || selected_keys.iter().any(Option::is_none) {
        return Err(Error::InvalidSource { path });
    }
    let keys_bytes = selected_len
        .checked_mul(size_of::<Vec<PhysicalScalar>>())
        .ok_or(Error::InvalidSource { path })?;
    budget
        .allocations(selected_len)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(keys_bytes)
        .map_err(|error| map_core(error, path))?;
    let mut keys = Vec::new();
    keys.try_reserve_exact(selected_len)
        .map_err(|_| Error::Allocation {
            amount: selected_len,
            path,
        })?;
    for value in selected_keys {
        keys.push(value.ok_or(Error::InvalidSource { path })?);
    }
    let row_header_indices = read_header_indices(package, selection, budget)?;
    let uid_map = read_uid_map(package, selection, budget)?;
    Ok(PhysicalScan {
        selected_keys: keys,
        uid_map,
        row_header_indices,
    })
}

fn read_string_table(
    package: &Package,
    selection: &PhysicalSelection,
    budget: &mut core::Budget,
) -> Result<HashMap<u32, Arc<str>>> {
    let path = selection.path();
    let object = core::object_at(package, selection.topology.string_table())
        .map_err(|error| map_core(error, path))?;
    let payload = message_payload_any(
        object,
        &[
            TABLE_DATA_LIST_MESSAGE_TYPE,
            TABLE_DATA_LIST_NATIVE_MESSAGE_TYPE,
        ],
        path,
    )?;
    let options = storage_options(package, budget, payload.len())?;
    let mut shape = StringShapeVisitor::default();
    let (snapshot, report) =
        storage_codec::decode_table_data_list_with_visitor(payload, options, &mut shape)
            .map_err(|error| map_storage_codec(error, path))?;
    charge_storage_report(report, budget, path)?;
    if snapshot.list_type() != 1 || shape.has_segments || shape.has_unsupported_sidecar {
        return Err(Error::UnsupportedTopology { path });
    }
    if shape.has_rich_text_entries {
        return Err(Error::UnsupportedFeature {
            path,
            feature: UnsupportedFeature::UnsupportedCell,
        });
    }
    if shape.has_comment_entries {
        return Err(Error::UnsupportedFeature {
            path,
            feature: UnsupportedFeature::CommentAnchors,
        });
    }

    // Decode the same source a second time only after its entry shape is
    // known.  This lets the owner precharge the exact string-vector and text
    // allocations before the visitor starts cloning values.
    let entry_bytes = shape
        .entry_count
        .checked_mul(size_of::<(u32, String)>())
        .ok_or(Error::InvalidSource { path })?;
    budget
        .allocations(1)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(entry_bytes)
        .map_err(|error| map_core(error, path))?;
    if shape.text_bytes != 0 {
        budget
            .allocations(shape.text_bytes)
            .map_err(|error| map_core(error, path))?;
        budget
            .retained(shape.text_bytes)
            .map_err(|error| map_core(error, path))?;
    }
    let options = storage_options(package, budget, payload.len())?;
    let mut visitor = StringVisitor::try_with_capacity(shape.entry_count, path)?;
    let (second_snapshot, second_report) =
        storage_codec::decode_table_data_list_with_visitor(payload, options, &mut visitor)
            .map_err(|error| map_storage_codec(error, path))?;
    charge_storage_report(second_report, budget, path)?;
    if second_snapshot != snapshot
        || visitor.has_segments != shape.has_segments
        || visitor.has_rich_text_entries != shape.has_rich_text_entries
        || visitor.has_comment_entries != shape.has_comment_entries
        || visitor.has_unsupported_sidecar != shape.has_unsupported_sidecar
    {
        return Err(Error::InvalidSource { path });
    }
    let map_bytes = visitor
        .entries
        .len()
        .checked_mul(
            size_of::<(u32, Arc<str>)>()
                .checked_add(size_of::<usize>())
                .ok_or(Error::InvalidSource { path })?,
        )
        .ok_or(Error::InvalidSource { path })?;
    budget
        .allocations(1)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(map_bytes)
        .map_err(|error| map_core(error, path))?;
    let mut strings = HashMap::new();
    strings
        .try_reserve(visitor.entries.len())
        .map_err(|_| Error::Allocation {
            amount: visitor.entries.len(),
            path,
        })?;
    for (key, value) in visitor.entries {
        if strings.contains_key(&key) {
            return Err(Error::InvalidSource { path });
        }
        if strings.len() == strings.capacity() {
            return Err(Error::Allocation { amount: 1, path });
        }
        budget
            .owned_value(value.len())
            .map_err(|error| map_core(error, path))?;
        strings.insert(key, Arc::<str>::from(value));
    }
    Ok(strings)
}

fn decode_row_cells(
    snapshot: physical_codec::TileRowInfoSnapshot<'_>,
    selection: &PhysicalSelection,
    strings: &HashMap<u32, Arc<str>>,
    global_row: usize,
    budget: &mut core::Budget,
) -> Result<Vec<Option<PhysicalScalar>>> {
    let path = selection.path();
    if snapshot.storage_version() != Some(5) {
        return Err(Error::UnsupportedTopology { path });
    }
    let storage = snapshot
        .cell_storage_buffer()
        .ok_or(Error::UnsupportedTopology { path })?;
    let offsets = snapshot
        .cell_offsets()
        .ok_or(Error::UnsupportedTopology { path })?;
    let columns =
        usize::try_from(selection.target.columns).map_err(|_| Error::InvalidSource { path })?;
    let expected_cells =
        usize::try_from(snapshot.cell_count()).map_err(|_| Error::InvalidSource { path })?;
    let pre_storage = snapshot.cell_storage_buffer_pre_bnc();
    let pre_offsets = snapshot.cell_offsets_pre_bnc();
    if pre_storage.is_empty() != pre_offsets.is_empty() {
        return Err(Error::InvalidSource { path });
    }
    if !pre_storage.is_empty() {
        validate_padded_offsets(
            pre_offsets,
            pre_storage.len(),
            columns,
            expected_cells,
            1,
            path,
        )?;
        validate_pre_bnc_sentinels(pre_storage, pre_offsets, columns, path)?;
        budget
            .work(pre_storage.len())
            .map_err(|error| map_core(error, path))?;
    }
    let expected_wide_rows = selection
        .topology
        .wide_rows()
        .ok_or(Error::UnsupportedTopology { path })?;
    let has_wide_offsets = snapshot
        .has_wide_offsets()
        .ok_or(Error::UnsupportedTopology { path })?;
    if has_wide_offsets != expected_wide_rows {
        return Err(Error::UnsupportedTopology { path });
    }
    let wide_multiplier = if has_wide_offsets { 4usize } else { 1usize };
    validate_padded_offsets(
        offsets,
        storage.len(),
        columns,
        expected_cells,
        wide_multiplier,
        path,
    )?;
    let starts_bytes = columns
        .checked_mul(size_of::<Option<usize>>())
        .ok_or(Error::InvalidSource { path })?;
    let values_bytes = columns
        .checked_mul(size_of::<Option<PhysicalScalar>>())
        .ok_or(Error::InvalidSource { path })?;
    budget
        .allocations(columns)
        .map_err(|error| map_core(error, path))?;
    budget
        .scratch(starts_bytes)
        .map_err(|error| map_core(error, path))?;
    budget
        .allocations(columns)
        .map_err(|error| map_core(error, path))?;
    budget
        .scratch(values_bytes)
        .map_err(|error| map_core(error, path))?;
    let mut starts: Vec<Option<usize>> = Vec::new();
    starts
        .try_reserve_exact(columns)
        .map_err(|_| Error::Allocation {
            amount: columns,
            path,
        })?;
    for bytes in offsets.chunks_exact(2).take(columns) {
        let raw = u16::from_le_bytes([bytes[0], bytes[1]]);
        let start = if raw == u16::MAX {
            None
        } else {
            Some(
                usize::from(raw)
                    .checked_mul(wide_multiplier)
                    .ok_or(Error::InvalidSource { path })?,
            )
        };
        starts.push(start);
    }
    let mut previous = 0usize;
    let mut present = 0usize;
    for start in starts.iter().flatten().copied() {
        if start >= storage.len()
            || (present == 0 && start != 0)
            || (present != 0 && start <= previous)
        {
            return Err(Error::InvalidSource { path });
        }
        previous = start;
        present = present
            .checked_add(1)
            .ok_or(Error::InvalidSource { path })?;
    }
    if present != expected_cells {
        return Err(Error::InvalidSource { path });
    }
    let mut values = Vec::new();
    values
        .try_reserve_exact(columns)
        .map_err(|_| Error::Allocation {
            amount: columns,
            path,
        })?;
    for (column, start) in starts.iter().copied().enumerate() {
        let Some(start) = start else {
            values.push(None);
            continue;
        };
        let next_column = column.checked_add(1).ok_or(Error::InvalidSource { path })?;
        let end = starts
            .iter()
            .skip(next_column)
            .flatten()
            .copied()
            .next()
            .unwrap_or(storage.len());
        if start >= end || end > storage.len() {
            return Err(Error::InvalidSource { path });
        }
        let view =
            BncCellView::parse(&storage[start..end]).map_err(|_| Error::InvalidSource { path })?;
        let value = decode_cell(view, selection, strings, global_row, column)?;
        budget
            .work(
                end.checked_sub(start)
                    .ok_or(Error::InvalidSource { path })?,
            )
            .map_err(|error| map_core(error, path))?;
        values.push(value);
    }
    Ok(values)
}

/// Validate the legacy mirror carried by native BNC-v5 rows.
///
/// Keynote 14.4 emits one canonical empty v4 sentinel for each materialized
/// modern cell.  Those sentinels contain no value, formula, comment, rich
/// text, format, control, or inline-border state.  The physical owner moves
/// the complete row envelope but deliberately refuses any richer legacy cell
/// until that second representation has its own typed dependency census.
fn validate_pre_bnc_sentinels(
    storage: &[u8],
    offsets: &[u8],
    columns: usize,
    path: Path,
) -> Result<()> {
    const EMPTY_V4_SENTINEL: [u8; 12] = [4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];

    let mut starts = offsets
        .chunks_exact(2)
        .take(columns)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .filter(|offset| *offset != u16::MAX)
        .map(usize::from);
    let mut start = starts.next().ok_or(Error::InvalidSource { path })?;
    for end in starts.chain(std::iter::once(storage.len())) {
        if start >= end
            || end > storage.len()
            || storage.get(start..end) != Some(EMPTY_V4_SENTINEL.as_slice())
        {
            return Err(Error::UnsupportedFeature {
                path,
                feature: UnsupportedFeature::RowAffineDependency,
            });
        }
        start = end;
    }
    Ok(())
}

fn validate_padded_offsets(
    offsets: &[u8],
    storage_len: usize,
    columns: usize,
    expected_cells: usize,
    unit: usize,
    path: Path,
) -> Result<()> {
    if storage_len == 0 || !offsets.len().is_multiple_of(2) || expected_cells > columns {
        return Err(Error::InvalidSource { path });
    }
    let slot_count = offsets.len() / 2;
    if slot_count < columns
        || offsets
            .chunks_exact(2)
            .skip(columns)
            .any(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]) != u16::MAX)
    {
        return Err(Error::InvalidSource { path });
    }
    let mut present = 0usize;
    let mut previous = 0usize;
    for bytes in offsets.chunks_exact(2).take(columns) {
        let raw = u16::from_le_bytes([bytes[0], bytes[1]]);
        if raw == u16::MAX {
            continue;
        }
        let offset = usize::from(raw)
            .checked_mul(unit)
            .ok_or(Error::InvalidSource { path })?;
        if offset >= storage_len
            || (present == 0 && offset != 0)
            || (present != 0 && offset <= previous)
        {
            return Err(Error::InvalidSource { path });
        }
        previous = offset;
        present = present
            .checked_add(1)
            .ok_or(Error::InvalidSource { path })?;
    }
    if present == expected_cells {
        Ok(())
    } else {
        Err(Error::InvalidSource { path })
    }
}

fn decode_cell(
    view: BncCellView<'_>,
    selection: &PhysicalSelection,
    strings: &HashMap<u32, Arc<str>>,
    global_row: usize,
    column: usize,
) -> Result<Option<PhysicalScalar>> {
    let path = selection.path();
    let body_start = selection.body_start()?;
    let identifiers = [
        view.style_identifier(),
        view.text_style_identifier(),
        view.cell_format_kind(),
        view.format_identifier(),
        view.secondary_format_identifier(),
    ];
    if identifiers
        .into_iter()
        .flatten()
        .any(|identifier| identifier == 0)
        || view.cell_format_kind().is_some() != view.format_identifier().is_some()
    {
        return Err(Error::InvalidSource { path });
    }
    if view.comment_identifier().is_some() {
        return Err(Error::UnsupportedFeature {
            path,
            feature: UnsupportedFeature::CommentAnchors,
        });
    }
    if let Some(identifier) = view.rich_text_identifier() {
        if identifier == 0 {
            return Err(Error::InvalidSource { path });
        }
        return Err(Error::UnsupportedFeature {
            path,
            feature: UnsupportedFeature::UnsupportedCell,
        });
    }
    if view.formula_error_identifier().is_some() {
        return Err(Error::UnsupportedFeature {
            path,
            feature: UnsupportedFeature::FormulaError,
        });
    }
    if view.conditional_style_identifier().is_some()
        || view.conditional_style_applied_rule().is_some()
    {
        return Err(Error::UnsupportedFeature {
            path,
            feature: UnsupportedFeature::ConditionalStyles,
        });
    }
    if view.control_cell_spec_identifier().is_some()
        || view.has_reserved_known_field()
        || !view.opaque_tail().is_empty()
    {
        return Err(Error::UnsupportedFeature {
            path,
            feature: UnsupportedFeature::RowAffineDependency,
        });
    }
    let unsupported = |feature| Error::UnsupportedFeature { path, feature };
    match view.stored_value() {
        StoredValue::Empty => Ok(None),
        StoredValue::Text(identifier) => {
            let value = strings
                .get(&identifier)
                .cloned()
                .ok_or(Error::InvalidSource { path })?;
            Ok(Some(PhysicalScalar::Text(value)))
        },
        StoredValue::Formula(_) => Err(unsupported(UnsupportedFeature::Formula)),
        StoredValue::RichText(_) => Err(unsupported(UnsupportedFeature::UnsupportedCell)),
        StoredValue::Error => Err(unsupported(UnsupportedFeature::FormulaError)),
        StoredValue::Unsupported(_) => Err(unsupported(UnsupportedFeature::UnsupportedCell)),
        StoredValue::Number => match view.cached_scalar() {
            Some(CachedScalar::Number(value)) => Ok(Some(PhysicalScalar::Number(value.get()))),
            _ => Err(Error::UnsupportedCell {
                path,
                row: Position::new(global_row.saturating_sub(body_start)),
                column: ColumnIndex::new(column).map_err(|_| Error::InvalidSource { path })?,
            }),
        },
        StoredValue::Boolean => match view.cached_scalar() {
            Some(CachedScalar::Boolean(value)) => Ok(Some(PhysicalScalar::Boolean(value))),
            _ => Err(Error::UnsupportedCell {
                path,
                row: Position::new(global_row.saturating_sub(body_start)),
                column: ColumnIndex::new(column).map_err(|_| Error::InvalidSource { path })?,
            }),
        },
        StoredValue::Date => match view.cached_scalar() {
            Some(CachedScalar::Date(value)) => Ok(Some(PhysicalScalar::Date(value.get()))),
            _ => Err(Error::UnsupportedCell {
                path,
                row: Position::new(global_row.saturating_sub(body_start)),
                column: ColumnIndex::new(column).map_err(|_| Error::InvalidSource { path })?,
            }),
        },
        StoredValue::Duration => match view.cached_scalar() {
            Some(CachedScalar::Duration(value)) => Ok(Some(PhysicalScalar::Duration(value.get()))),
            _ => Err(Error::UnsupportedCell {
                path,
                row: Position::new(global_row.saturating_sub(body_start)),
                column: ColumnIndex::new(column).map_err(|_| Error::InvalidSource { path })?,
            }),
        },
    }
}

fn read_header_indices(
    package: &Package,
    selection: &PhysicalSelection,
    budget: &mut core::Budget,
) -> Result<Vec<u32>> {
    let path = selection.path();
    let header_capacity =
        selection
            .topology
            .row_header_buckets()
            .iter()
            .try_fold(0usize, |total, bucket| {
                total
                    .checked_add(bucket.header_count)
                    .ok_or(Error::InvalidSource { path })
            })?;
    let header_bytes = header_capacity
        .checked_mul(
            size_of::<u32>()
                .checked_add(size_of::<usize>())
                .ok_or(Error::InvalidSource { path })?,
        )
        .ok_or(Error::InvalidSource { path })?;
    budget
        .allocations(2)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(
            header_bytes
                .checked_mul(2)
                .ok_or(Error::InvalidSource { path })?,
        )
        .map_err(|error| map_core(error, path))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(header_capacity)
        .map_err(|_| Error::Allocation {
            amount: header_capacity,
            path,
        })?;
    let mut seen = HashSet::new();
    seen.try_reserve(header_capacity)
        .map_err(|_| Error::Allocation {
            amount: header_capacity,
            path,
        })?;
    for bucket in selection.topology.row_header_buckets() {
        let object =
            core::object_at(package, &bucket.location).map_err(|error| map_core(error, path))?;
        let payload = message_payload(object, HEADER_BUCKET_MESSAGE_TYPE, path)?;
        let bucket_start = bucket
            .bucket_index
            .checked_mul(HEADER_BUCKET_ROWS)
            .ok_or(Error::InvalidSource { path })?;
        let row_limit = bucket_start
            .checked_add(HEADER_BUCKET_ROWS)
            .map_or(selection.target.rows, |end| end.min(selection.target.rows));
        let options = physical_options(
            package,
            budget,
            payload.len(),
            row_limit,
            selection.target.columns,
        )?;
        let prepared =
            physical_codec::plan_header_storage_bucket_rows(payload, row_limit, &[], options)
                .map_err(|error| map_physical_codec(error, path))?;
        charge_rewrite_requirements(prepared.requirements(), budget, path)?;
        if prepared.bucket_hash_function() == 0 {
            return Err(Error::InvalidSource { path });
        }
        for record in prepared.records() {
            let snapshot = record.snapshot();
            if snapshot.hiding_state() != 0 {
                return Err(Error::UnsupportedFeature {
                    path,
                    feature: UnsupportedFeature::NonPositionalHiddenState,
                });
            }
            if snapshot.number_of_cells() != selection.target.columns {
                return Err(Error::InvalidSource { path });
            }
            if !seen.contains(&snapshot.index()) && seen.len() == seen.capacity() {
                return Err(Error::Allocation { amount: 1, path });
            }
            if !seen.insert(snapshot.index()) {
                return Err(Error::InvalidSource { path });
            }
            if output.len() == output.capacity() {
                return Err(Error::Allocation { amount: 1, path });
            }
            output.push(snapshot.index());
        }
    }
    // Validate the column-axis records too.  They are not moved, but an
    // active hidden axis would make a row-only permutation semantically false.
    let object = core::object_at(package, selection.topology.column_headers())
        .map_err(|error| map_core(error, path))?;
    let payload = message_payload(object, HEADER_BUCKET_MESSAGE_TYPE, path)?;
    let options = physical_options(
        package,
        budget,
        payload.len(),
        selection.target.columns,
        selection.target.columns,
    )?;
    let prepared = physical_codec::plan_header_storage_bucket_rows(
        payload,
        selection.target.columns,
        &[],
        options,
    )
    .map_err(|error| map_physical_codec(error, path))?;
    charge_rewrite_requirements(prepared.requirements(), budget, path)?;
    for record in prepared.records() {
        if record.snapshot().hiding_state() != 0 {
            return Err(Error::UnsupportedFeature {
                path,
                feature: UnsupportedFeature::NonPositionalHiddenState,
            });
        }
    }
    output.sort_unstable();
    Ok(output)
}

fn read_uid_map(
    package: &Package,
    selection: &PhysicalSelection,
    budget: &mut core::Budget,
) -> Result<physical_codec::ColumnRowUidMapSnapshot> {
    let path = selection.path();
    let object = core::object_at(package, selection.topology.row_uid_map())
        .map_err(|error| map_core(error, path))?;
    let payload = message_payload_any(
        object,
        &[
            COLUMN_ROW_UID_MAP_MESSAGE_TYPE,
            COLUMN_ROW_UID_MAP_LEGACY_MESSAGE_TYPE,
        ],
        path,
    )?;
    let rows = usize::try_from(selection.target.rows).map_err(|_| Error::InvalidSource { path })?;
    let columns =
        usize::try_from(selection.target.columns).map_err(|_| Error::InvalidSource { path })?;
    let options = physical_options(
        package,
        budget,
        payload.len(),
        selection.target.rows,
        selection.target.columns,
    )?;
    let (snapshot, report) =
        physical_codec::decode_column_row_uid_map(payload, columns, rows, options)
            .map_err(|error| map_physical_codec(error, path))?;
    charge_decode_report(report, budget, path)?;
    Ok(snapshot)
}

struct StagedComponent {
    name: Arc<str>,
    archive: Archive,
}

struct EncodedComponent {
    name: Arc<str>,
    compressed: Vec<u8>,
}

fn staged_component_archive<'a>(
    source: &Package,
    staged: &'a mut Vec<StagedComponent>,
    name: &Arc<str>,
    budget: &mut core::Budget,
    path: Path,
) -> Result<&'a mut Archive> {
    if let Some(index) = staged
        .iter()
        .position(|component| component.name.as_ref() == name.as_ref())
    {
        return Ok(&mut staged[index].archive);
    }
    if staged.len() == staged.capacity() {
        return Err(Error::InvalidSource { path });
    }
    let archive = core::component_archive(source, name.as_ref(), budget)
        .map_err(|error| map_core(error, path))?;
    staged.push(StagedComponent {
        name: Arc::clone(name),
        archive,
    });
    let index = staged
        .len()
        .checked_sub(1)
        .ok_or(Error::InvalidSource { path })?;
    Ok(&mut staged[index].archive)
}

fn rewrite_physical(
    source: &Package,
    selection: &PhysicalSelection,
    destination_by_source: &[u32],
    deleted_previews: &[&str],
    budget: &mut core::Budget,
) -> Result<Package> {
    let path = selection.path();
    let catalog = core::physical_catalog(source).map_err(|error| map_core(error, path))?;
    let physical_limits = source.state.options.archive();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    let component_capacity = selection.topology_component_count();
    budget
        .allocations(component_capacity)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(
            component_capacity
                .checked_mul(size_of::<StagedComponent>())
                .ok_or(Error::InvalidSource { path })?,
        )
        .map_err(|error| map_core(error, path))?;
    let mut staged_components = Vec::new();
    staged_components
        .try_reserve_exact(component_capacity)
        .map_err(|_| Error::Allocation {
            amount: component_capacity,
            path,
        })?;
    let tile_size = selection.topology.tile_size();

    // TileStorage is coordinate-addressed.  The first owner refuses a
    // cross-tile move, because a whole-row rewrite must never synthesize a
    // new tile or silently alter tile affinity.
    for tile in selection.topology.tile_references() {
        let move_count = destination_by_source.iter().copied().enumerate().try_fold(
            0usize,
            |count, (source_global, destination)| {
                let source_global =
                    u32::try_from(source_global).map_err(|_| Error::InvalidSource { path })?;
                if source_global == destination {
                    return Ok(count);
                }
                let source_tile = source_global
                    .checked_div(tile_size)
                    .ok_or(Error::InvalidSource { path })?;
                let destination_tile = destination
                    .checked_div(tile_size)
                    .ok_or(Error::InvalidSource { path })?;
                if source_tile != tile.tile_id && destination_tile != tile.tile_id {
                    return Ok(count);
                }
                if source_tile != destination_tile {
                    return Err(Error::UnsupportedTopology { path });
                }
                if source_tile == tile.tile_id {
                    count.checked_add(1).ok_or(Error::InvalidSource { path })
                } else {
                    Ok(count)
                }
            },
        )?;
        if move_count == 0 {
            continue;
        }
        let move_bytes = move_count
            .checked_mul(size_of::<physical_codec::RowMove>())
            .ok_or(Error::InvalidSource { path })?;
        budget
            .allocations(1)
            .map_err(|error| map_core(error, path))?;
        budget
            .retained(move_bytes)
            .map_err(|error| map_core(error, path))?;
        let mut moves = Vec::new();
        moves
            .try_reserve_exact(move_count)
            .map_err(|_| Error::Allocation {
                amount: move_count,
                path,
            })?;
        for (source_global, destination) in destination_by_source.iter().copied().enumerate() {
            let source_global =
                u32::try_from(source_global).map_err(|_| Error::InvalidSource { path })?;
            if source_global == destination {
                continue;
            }
            let source_tile = source_global
                .checked_div(tile_size)
                .ok_or(Error::InvalidSource { path })?;
            let destination_tile = destination
                .checked_div(tile_size)
                .ok_or(Error::InvalidSource { path })?;
            if source_tile != tile.tile_id && destination_tile != tile.tile_id {
                continue;
            }
            if source_tile != destination_tile {
                return Err(Error::UnsupportedTopology { path });
            }
            if source_tile == tile.tile_id {
                let source_local = source_global % tile_size;
                let destination_local = destination % tile_size;
                if moves.len() == moves.capacity() {
                    return Err(Error::Allocation { amount: 1, path });
                }
                moves.push(physical_codec::RowMove::new(
                    source_local,
                    destination_local,
                ));
            }
        }
        debug_assert_eq!(moves.len(), move_count);
        let archive = staged_component_archive(
            source,
            &mut staged_components,
            &tile.location.component,
            budget,
            path,
        )?;
        let object = archive
            .object(tile.location.identifier)
            .ok_or(Error::InvalidSource { path })?;
        let payload = message_payload(object, TILE_MESSAGE_TYPE, path)?;
        let options = physical_options(
            source,
            budget,
            payload.len(),
            tile_size,
            selection.target.columns,
        )?;
        let prepared = physical_codec::plan_tile_rows_rewrite(payload, tile_size, &moves, options)
            .map_err(|error| map_physical_codec(error, path))?;
        charge_rewrite_requirements(prepared.requirements(), budget, path)?;
        let tile_snapshot = prepared.tile();
        if tile_snapshot.storage_version() != Some(5)
            || tile_snapshot.last_saved_in_bnc() != Some(true)
        {
            return Err(Error::InvalidSource { path });
        }
        if tile_snapshot.should_use_wide_rows() != selection.topology.wide_rows() {
            return Err(Error::UnsupportedTopology { path });
        }
        if prepared
            .row_records()
            .any(|record| record.snapshot().has_wide_offsets() != selection.topology.wide_rows())
        {
            return Err(Error::UnsupportedTopology { path });
        }
        let (rewritten, report) = physical_codec::execute_tile_rows_rewrite(prepared, options)
            .map_err(|error| map_physical_codec(error, path))?;
        if report.changed_records() != moves.len() {
            return Err(Error::Verification { path });
        }
        charge_decode_report(report.result(), budget, path)?;
        replace_message(
            archive,
            tile.location.identifier,
            TILE_MESSAGE_TYPE,
            rewritten,
            archive_limits,
            path,
        )?;
    }

    // Header buckets are sparse; only records that actually exist are moved.
    // Empty slots are never materialized.  Each record remains in its source
    // bucket, and the strict codec retains all unknown fields byte-for-byte.
    for bucket in selection.topology.row_header_buckets() {
        let bucket_start = bucket
            .bucket_index
            .checked_mul(HEADER_BUCKET_ROWS)
            .ok_or(Error::InvalidSource { path })?;
        let row_limit = bucket_start
            .checked_add(HEADER_BUCKET_ROWS)
            .map_or(selection.target.rows, |end| end.min(selection.target.rows));
        let archive = staged_component_archive(
            source,
            &mut staged_components,
            &bucket.location.component,
            budget,
            path,
        )?;
        let object = archive
            .object(bucket.location.identifier)
            .ok_or(Error::InvalidSource { path })?;
        let payload = message_payload(object, HEADER_BUCKET_MESSAGE_TYPE, path)?;
        let options = physical_options(
            source,
            budget,
            payload.len(),
            row_limit,
            selection.target.columns,
        )?;
        let identity =
            physical_codec::plan_header_storage_bucket_rows(payload, row_limit, &[], options)
                .map_err(|error| map_physical_codec(error, path))?;
        charge_rewrite_requirements(identity.requirements(), budget, path)?;
        let move_count = identity.records().try_fold(0usize, |count, record| {
            let source_index = record.snapshot().index();
            let source_usize =
                usize::try_from(source_index).map_err(|_| Error::InvalidSource { path })?;
            let destination = destination_by_source
                .get(source_usize)
                .copied()
                .ok_or(Error::InvalidPermutation { path })?;
            if source_index == destination {
                return Ok(count);
            }
            if source_index / HEADER_BUCKET_ROWS != destination / HEADER_BUCKET_ROWS {
                return Err(Error::UnsupportedTopology { path });
            }
            count.checked_add(1).ok_or(Error::InvalidSource { path })
        })?;
        if move_count == 0 {
            continue;
        }
        let move_bytes = move_count
            .checked_mul(size_of::<physical_codec::HeaderRowMove>())
            .ok_or(Error::InvalidSource { path })?;
        budget
            .allocations(1)
            .map_err(|error| map_core(error, path))?;
        budget
            .retained(move_bytes)
            .map_err(|error| map_core(error, path))?;
        let mut moves = Vec::new();
        moves
            .try_reserve_exact(move_count)
            .map_err(|_| Error::Allocation {
                amount: move_count,
                path,
            })?;
        for record in identity.records() {
            let source_index = record.snapshot().index();
            let source_usize =
                usize::try_from(source_index).map_err(|_| Error::InvalidSource { path })?;
            let Some(destination) = destination_by_source.get(source_usize).copied() else {
                return Err(Error::InvalidPermutation { path });
            };
            if source_index != destination {
                if source_index / HEADER_BUCKET_ROWS != destination / HEADER_BUCKET_ROWS {
                    return Err(Error::UnsupportedTopology { path });
                }
                if moves.len() == moves.capacity() {
                    return Err(Error::Allocation { amount: 1, path });
                }
                moves.push(physical_codec::HeaderRowMove::new(
                    source_index,
                    destination,
                ));
            }
        }
        debug_assert_eq!(moves.len(), move_count);
        let options = physical_options(
            source,
            budget,
            payload.len(),
            row_limit,
            selection.target.columns,
        )?;
        let prepared =
            physical_codec::plan_header_storage_bucket_rows(payload, row_limit, &moves, options)
                .map_err(|error| map_physical_codec(error, path))?;
        charge_rewrite_requirements(prepared.requirements(), budget, path)?;
        let (rewritten, report) =
            physical_codec::execute_header_storage_bucket_rows(prepared, options)
                .map_err(|error| map_physical_codec(error, path))?;
        if report.changed_records() != moves.len() {
            return Err(Error::Verification { path });
        }
        charge_decode_report(report.result(), budget, path)?;
        replace_message(
            archive,
            bucket.location.identifier,
            HEADER_BUCKET_MESSAGE_TYPE,
            rewritten,
            archive_limits,
            path,
        )?;
    }

    let row_count =
        usize::try_from(selection.target.rows).map_err(|_| Error::InvalidSource { path })?;
    let column_count =
        usize::try_from(selection.target.columns).map_err(|_| Error::InvalidSource { path })?;
    let uid_archive = staged_component_archive(
        source,
        &mut staged_components,
        &selection.topology.row_uid_map().component,
        budget,
        path,
    )?;
    let uid_object = uid_archive
        .object(selection.topology.row_uid_map().identifier)
        .ok_or(Error::InvalidSource { path })?;
    let uid_payload = message_payload_any(
        uid_object,
        &[
            COLUMN_ROW_UID_MAP_MESSAGE_TYPE,
            COLUMN_ROW_UID_MAP_LEGACY_MESSAGE_TYPE,
        ],
        path,
    )?;
    let permutation_bytes = destination_by_source
        .len()
        .checked_mul(size_of::<u32>())
        .ok_or(Error::InvalidSource { path })?;
    budget
        .allocations(1)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(permutation_bytes)
        .map_err(|error| map_core(error, path))?;
    let permutation = physical_codec::RowUidPermutation::new(destination_by_source)
        .map_err(|error| map_physical_codec(error, path))?;
    let options = physical_options(
        source,
        budget,
        uid_payload.len(),
        selection.target.rows,
        selection.target.columns,
    )?;
    let prepared = physical_codec::plan_column_row_uid_map_rewrite(
        uid_payload,
        column_count,
        row_count,
        &permutation,
        options,
    )
    .map_err(|error| map_physical_codec(error, path))?;
    charge_rewrite_requirements(prepared.requirements(), budget, path)?;
    let (rewritten_uid, report) = physical_codec::execute_column_row_uid_map_rewrite(
        prepared,
        column_count,
        row_count,
        options,
    )
    .map_err(|error| map_physical_codec(error, path))?;
    charge_decode_report(report.result(), budget, path)?;
    let uid_message_type = find_message_type(
        uid_archive
            .object(selection.topology.row_uid_map().identifier)
            .ok_or(Error::InvalidSource { path })?,
        &[
            COLUMN_ROW_UID_MAP_MESSAGE_TYPE,
            COLUMN_ROW_UID_MAP_LEGACY_MESSAGE_TYPE,
        ],
        path,
    )?;
    replace_message(
        uid_archive,
        selection.topology.row_uid_map().identifier,
        uid_message_type,
        rewritten_uid,
        archive_limits,
        path,
    )?;

    let encoded_capacity = staged_components.len();
    budget
        .allocations(
            encoded_capacity
                .checked_mul(2)
                .ok_or(Error::InvalidSource { path })?,
        )
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(
            encoded_capacity
                .checked_mul(size_of::<EncodedComponent>())
                .ok_or(Error::InvalidSource { path })?,
        )
        .map_err(|error| map_core(error, path))?;
    let mut encoded_components = Vec::new();
    encoded_components
        .try_reserve_exact(encoded_capacity)
        .map_err(|_| Error::Allocation {
            amount: encoded_capacity,
            path,
        })?;
    let mut compressed_total = 0usize;
    for component in staged_components {
        let encoded_len = component
            .archive
            .encoded_len_with_limits(archive_limits)
            .map_err(|_| Error::InvalidSource { path })?;
        let compressed_bound = SnappyStream::maximum_compressed_len(encoded_len)
            .map_err(|_| Error::InvalidSource { path })?;
        budget
            .output(
                encoded_len
                    .checked_add(compressed_bound)
                    .ok_or(Error::InvalidSource { path })?,
            )
            .map_err(|error| map_core(error, path))?;
        budget
            .allocations(1)
            .map_err(|error| map_core(error, path))?;
        budget
            .retained(encoded_len)
            .map_err(|error| map_core(error, path))?;
        budget
            .scratch(encoded_len)
            .map_err(|error| map_core(error, path))?;
        let decoded = component
            .archive
            .to_bytes_with_limits(archive_limits)
            .map_err(|_| Error::InvalidSource { path })?;
        if decoded.len() != encoded_len {
            return Err(Error::Verification { path });
        }
        budget
            .allocations(1)
            .map_err(|error| map_core(error, path))?;
        budget
            .retained(compressed_bound)
            .map_err(|error| map_core(error, path))?;
        budget
            .scratch(compressed_bound)
            .map_err(|error| map_core(error, path))?;
        let compressed =
            SnappyStream::compress(&decoded).map_err(|_| Error::InvalidSource { path })?;
        if compressed.len() > compressed_bound {
            return Err(Error::Verification { path });
        }
        compressed_total = compressed_total
            .checked_add(compressed.len())
            .ok_or(Error::InvalidSource { path })?;
        if encoded_components.len() == encoded_components.capacity() {
            return Err(Error::Allocation { amount: 1, path });
        }
        encoded_components.push(EncodedComponent {
            name: component.name,
            compressed,
        });
    }
    let edit_count = encoded_components.len();
    budget
        .allocations(1)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(
            edit_count
                .checked_mul(size_of::<EntryEdit<'_>>())
                .ok_or(Error::InvalidSource { path })?,
        )
        .map_err(|error| map_core(error, path))?;
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(encoded_components.len())
        .map_err(|_| Error::Allocation {
            amount: encoded_components.len(),
            path,
        })?;
    for component in &encoded_components {
        if edits.len() == edits.capacity() {
            return Err(Error::Allocation { amount: 1, path });
        }
        edits.push(EntryEdit::new(
            component.name.as_ref(),
            component.compressed.as_slice(),
        ));
    }
    budget
        .preflight_reassembly(catalog, compressed_total, deleted_previews.len())
        .map_err(|error| map_core(error, path))?;
    let prepared = catalog
        .prepare_reassembly_with_deletions(&edits, deleted_previews, physical_limits)
        .map_err(map_archive)?;
    let requirements = prepared.execution_requirements();
    budget
        .reassembly(requirements)
        .map_err(|error| map_core(error, path))?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive)?;
    budget
        .preflight_candidate(source, output.len())
        .map_err(|error| map_core(error, path))?;
    // `Arc<[u8]>` is the publication owner for the candidate.  Charge its
    // control/storage allocation while the reassembly `Vec` is still live;
    // the candidate parser's own inventory is charged by the preflight above.
    budget
        .allocations(1)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(output.len())
        .map_err(|error| map_core(error, path))?;
    Package::from_source_with_options(Arc::<[u8]>::from(output), source.state.options)
        .map_err(map_read)
}

fn verify_candidate(
    source: &Package,
    candidate: &Package,
    selection: &PhysicalSelection,
    destination_by_source: &[u32],
    moved_rows: usize,
    budget: &mut core::Budget,
) -> Result<()> {
    let path = selection.path();
    let target = core::select_table(
        candidate,
        SlideSelector::position(selection.target.slide_position),
        TableSelector::position(selection.target.table_position),
        budget,
    )
    .map_err(|error| map_core(error, path))?;
    if !core::same_target(&target, &selection.target) || target.locked {
        return Err(Error::Verification { path });
    }
    let topology = core::admit_physical_table(candidate, &target, budget)
        .map_err(|error| map_core(error, path))?;
    if topology != selection.topology {
        return Err(Error::Verification { path });
    }
    let persisted = candidate
        .slide_table_sort_order(
            SlideSelector::position(selection.target.slide_position),
            TableSelector::position(selection.target.table_position),
        )
        .map_err(|error| map_sort_error(error, path))?;
    if persisted.as_ref() != Some(&selection.order) {
        return Err(Error::Verification { path });
    }
    let source_scan = scan_table(source, selection, budget)?;
    let order_bytes = selection
        .order
        .rules()
        .len()
        .checked_mul(size_of::<litchi_iwa_common::table::sort::Rule>())
        .ok_or(Error::InvalidSource { path })?;
    budget
        .allocations(1)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(order_bytes)
        .map_err(|error| map_core(error, path))?;
    let candidate_selection = PhysicalSelection {
        target,
        topology,
        order: selection.order.clone(),
        scope: selection.scope,
        range: selection.range,
    };
    let candidate_scan = scan_table(candidate, &candidate_selection, budget)?;
    let inverse = invert_permutation_checked(destination_by_source, budget, path)?;
    let (range_start, range_end) = selection.range_bounds()?;
    let body_start = selection.body_start()?;
    let selected_len = range_end
        .checked_sub(range_start)
        .ok_or(Error::InvalidPermutation { path })?;
    if candidate_scan.selected_keys.len() != selected_len {
        return Err(Error::Verification { path });
    }
    for destination_offset in 0..selected_len {
        let destination_body = range_start
            .checked_add(destination_offset)
            .ok_or(Error::Verification { path })?;
        let destination_global = body_start
            .checked_add(destination_body)
            .ok_or(Error::Verification { path })?;
        let source_global = *inverse
            .get(destination_global)
            .ok_or(Error::Verification { path })?;
        let source_body = source_global
            .checked_sub(body_start)
            .ok_or(Error::Verification { path })?;
        let source_offset = source_body
            .checked_sub(range_start)
            .ok_or(Error::Verification { path })?;
        if candidate_scan.selected_keys[destination_offset]
            != source_scan.selected_keys[source_offset]
        {
            return Err(Error::Verification { path });
        }
    }
    let source_uid = source_scan.uid_map.row_uid_for_index();
    let candidate_uid = candidate_scan.uid_map.row_uid_for_index();
    if source_uid.len() != candidate_uid.len() || source_uid.len() != destination_by_source.len() {
        return Err(Error::Verification { path });
    }
    let uid_mapping_bytes = source_uid
        .len()
        .checked_mul(size_of::<u32>())
        .ok_or(Error::InvalidSource { path })?;
    budget
        .allocations(2)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(
            uid_mapping_bytes
                .checked_mul(2)
                .ok_or(Error::InvalidSource { path })?,
        )
        .map_err(|error| map_core(error, path))?;
    let mut expected_uid = vec![0u32; source_uid.len()];
    let mut expected_row_index_for_uid = vec![0u32; source_uid.len()];
    for (source_index, destination) in destination_by_source.iter().copied().enumerate() {
        let destination = usize::try_from(destination).map_err(|_| Error::Verification { path })?;
        if destination >= expected_uid.len() {
            return Err(Error::Verification { path });
        }
        let uid = *source_uid
            .get(source_index)
            .ok_or(Error::Verification { path })?;
        expected_uid[destination] = uid;
        let uid_slot = usize::try_from(uid).map_err(|_| Error::Verification { path })?;
        *expected_row_index_for_uid
            .get_mut(uid_slot)
            .ok_or(Error::Verification { path })? =
            u32::try_from(destination).map_err(|_| Error::Verification { path })?;
    }
    if expected_uid != candidate_uid
        || expected_row_index_for_uid != candidate_scan.uid_map.row_index_for_uid()
        || source_scan.uid_map.sorted_column_uids() != candidate_scan.uid_map.sorted_column_uids()
        || source_scan.uid_map.column_index_for_uid()
            != candidate_scan.uid_map.column_index_for_uid()
        || source_scan.uid_map.column_uid_for_index()
            != candidate_scan.uid_map.column_uid_for_index()
        || source_scan.uid_map.sorted_row_uids() != candidate_scan.uid_map.sorted_row_uids()
    {
        return Err(Error::Verification { path });
    }
    let header_mapping_bytes = source_scan
        .row_header_indices
        .len()
        .checked_mul(size_of::<u32>())
        .ok_or(Error::InvalidSource { path })?;
    budget
        .allocations(2)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(
            header_mapping_bytes
                .checked_mul(2)
                .ok_or(Error::InvalidSource { path })?,
        )
        .map_err(|error| map_core(error, path))?;
    let mut expected_headers = source_scan.row_header_indices.clone();
    for index in &mut expected_headers {
        let source_index = usize::try_from(*index).map_err(|_| Error::Verification { path })?;
        if let Some(destination) = destination_by_source.get(source_index) {
            *index = *destination;
        }
    }
    expected_headers.sort_unstable();
    let mut candidate_headers = candidate_scan.row_header_indices.clone();
    candidate_headers.sort_unstable();
    if expected_headers != candidate_headers {
        return Err(Error::Verification { path });
    }
    if moved_rows == 0 {
        return Err(Error::Verification { path });
    }
    let source_catalog = core::physical_catalog(source).map_err(|error| map_core(error, path))?;
    let candidate_catalog =
        core::physical_catalog(candidate).map_err(|error| map_core(error, path))?;
    let source_previews =
        super::rendering_invalidation::root_preview_deletions(source_catalog.package())
            .map_err(|_| Error::Verification { path })?;
    let candidate_previews =
        super::rendering_invalidation::root_preview_deletions(candidate_catalog.package())
            .map_err(|_| Error::Verification { path })?;
    let (locality_source, locality_candidate, preview_deletions, locality_permutation) =
        if candidate_previews.names().is_empty() {
            (
                source,
                candidate,
                source_previews.names(),
                copy_permutation_checked(destination_by_source, budget, path)?,
            )
        } else if source_previews.names().is_empty() {
            // An inverse patch restores the exact source previews. Verify the
            // same physical/object delta in the reverse direction so preview
            // additions cannot become a general package-member wildcard.
            (
                candidate,
                source,
                candidate_previews.names(),
                invert_permutation_u32_checked(destination_by_source, budget, path)?,
            )
        } else {
            return Err(Error::Verification { path });
        };
    let changed_objects = changed_object_ids(
        selection,
        &source_scan.row_header_indices,
        &locality_permutation,
        budget,
    )?;
    let allowlist =
        core::LocalityAllowlist::with_physical_topology(&selection.topology, preview_deletions)
            .and_then(|value| value.with_changed_object_ids(&changed_objects))
            .map_err(|error| map_core(error, path))?;
    core::verify_physical_locality(locality_source, locality_candidate, &allowlist, budget)
        .map_err(|error| map_core(error, path))?;
    verify_physical_payloads(
        locality_source,
        locality_candidate,
        selection,
        &locality_permutation,
        budget,
    )?;
    verify_changed_object_framing(
        locality_source,
        locality_candidate,
        &allowlist,
        budget,
        path,
    )?;
    Ok(())
}

/// Compare the complete physical storage envelopes after applying the row
/// permutation.  The generated codec snapshots prove scalar meaning, while
/// this second pass keeps every non-rewritten wire field source-authoritative.
/// In particular, an unknown field in an untouched cell/header/UID column is
/// never silently normalized by candidate verification.
fn verify_physical_payloads(
    source: &Package,
    candidate: &Package,
    selection: &PhysicalSelection,
    destination_by_source: &[u32],
    budget: &mut core::Budget,
) -> Result<()> {
    let path = selection.path();
    let tile_size = selection.topology.tile_size();
    for tile_reference in selection.topology.tile_references() {
        let source_object = core::object_at(source, &tile_reference.location)
            .map_err(|error| map_core(error, path))?;
        let candidate_object = core::object_at(candidate, &tile_reference.location)
            .map_err(|error| map_core(error, path))?;
        let source_payload = message_payload(source_object, TILE_MESSAGE_TYPE, path)?;
        let candidate_payload = message_payload(candidate_object, TILE_MESSAGE_TYPE, path)?;
        let source_options = physical_options(
            source,
            budget,
            source_payload.len(),
            tile_size,
            selection.target.columns,
        )?;
        let source_plan =
            physical_codec::plan_tile_rows_rewrite(source_payload, tile_size, &[], source_options)
                .map_err(|error| map_physical_codec(error, path))?;
        charge_rewrite_requirements(source_plan.requirements(), budget, path)?;
        let candidate_options = physical_options(
            candidate,
            budget,
            candidate_payload.len(),
            tile_size,
            selection.target.columns,
        )?;
        let candidate_plan = physical_codec::plan_tile_rows_rewrite(
            candidate_payload,
            tile_size,
            &[],
            candidate_options,
        )
        .map_err(|error| map_physical_codec(error, path))?;
        charge_rewrite_requirements(candidate_plan.requirements(), budget, path)?;
        if source_plan.tile() != candidate_plan.tile() {
            return Err(Error::Verification { path });
        }
        same_wire_fields_except(source_payload, candidate_payload, &[5], budget, path)?;

        let source_count = source_plan.row_records().count();
        let candidate_count = candidate_plan.row_records().count();
        if source_count != candidate_count {
            return Err(Error::Verification { path });
        }
        let candidate_row_limit = usize::try_from(candidate_plan.tile().num_rows())
            .map_err(|_| Error::Verification { path })?;
        let candidate_rows = index_records_by_position(
            candidate_plan
                .row_records()
                .map(|record| (record.snapshot().tile_row_index(), record.raw())),
            0,
            candidate_row_limit,
            candidate_count,
            budget,
            path,
        )?;
        for source_record in source_plan.row_records() {
            let source_local = source_record.snapshot().tile_row_index();
            let source_global = tile_reference
                .tile_id
                .checked_mul(tile_size)
                .and_then(|base| base.checked_add(source_local))
                .ok_or(Error::Verification { path })?;
            let destination = *destination_by_source
                .get(usize::try_from(source_global).map_err(|_| Error::Verification { path })?)
                .ok_or(Error::Verification { path })?;
            if destination / tile_size != tile_reference.tile_id {
                return Err(Error::Verification { path });
            }
            let destination_local = destination % tile_size;
            let destination_local =
                usize::try_from(destination_local).map_err(|_| Error::Verification { path })?;
            let candidate_record = candidate_rows
                .get(destination_local)
                .and_then(|record| *record)
                .ok_or(Error::Verification { path })?;
            same_wire_fields_except(source_record.raw(), candidate_record, &[1], budget, path)?;
        }
    }

    for bucket in selection.topology.row_header_buckets() {
        let source_object =
            core::object_at(source, &bucket.location).map_err(|error| map_core(error, path))?;
        let candidate_object =
            core::object_at(candidate, &bucket.location).map_err(|error| map_core(error, path))?;
        let source_payload = message_payload(source_object, HEADER_BUCKET_MESSAGE_TYPE, path)?;
        let candidate_payload =
            message_payload(candidate_object, HEADER_BUCKET_MESSAGE_TYPE, path)?;
        let bucket_start = bucket
            .bucket_index
            .checked_mul(HEADER_BUCKET_ROWS)
            .ok_or(Error::Verification { path })?;
        let row_limit = bucket_start
            .checked_add(HEADER_BUCKET_ROWS)
            .map_or(selection.target.rows, |end| end.min(selection.target.rows));
        let source_options = physical_options(
            source,
            budget,
            source_payload.len(),
            row_limit,
            selection.target.columns,
        )?;
        let source_plan = physical_codec::plan_header_storage_bucket_rows(
            source_payload,
            row_limit,
            &[],
            source_options,
        )
        .map_err(|error| map_physical_codec(error, path))?;
        charge_rewrite_requirements(source_plan.requirements(), budget, path)?;
        let candidate_options = physical_options(
            candidate,
            budget,
            candidate_payload.len(),
            row_limit,
            selection.target.columns,
        )?;
        let candidate_plan = physical_codec::plan_header_storage_bucket_rows(
            candidate_payload,
            row_limit,
            &[],
            candidate_options,
        )
        .map_err(|error| map_physical_codec(error, path))?;
        charge_rewrite_requirements(candidate_plan.requirements(), budget, path)?;
        if source_plan.bucket_hash_function() != candidate_plan.bucket_hash_function() {
            return Err(Error::Verification { path });
        }
        same_wire_fields_except(source_payload, candidate_payload, &[2], budget, path)?;
        let source_count = source_plan.records().count();
        let candidate_count = candidate_plan.records().count();
        if source_count != candidate_count {
            return Err(Error::Verification { path });
        }
        let bucket_span = usize::try_from(
            row_limit
                .checked_sub(bucket_start)
                .ok_or(Error::Verification { path })?,
        )
        .map_err(|_| Error::Verification { path })?;
        let candidate_headers = index_records_by_position(
            candidate_plan
                .records()
                .map(|record| (record.snapshot().index(), record.raw())),
            bucket_start,
            bucket_span,
            candidate_count,
            budget,
            path,
        )?;
        for source_record in source_plan.records() {
            let source_index = source_record.snapshot().index();
            let source_usize =
                usize::try_from(source_index).map_err(|_| Error::Verification { path })?;
            let destination = *destination_by_source
                .get(source_usize)
                .ok_or(Error::Verification { path })?;
            if destination / HEADER_BUCKET_ROWS != bucket.bucket_index {
                return Err(Error::Verification { path });
            }
            let destination_local = usize::try_from(
                destination
                    .checked_sub(bucket_start)
                    .ok_or(Error::Verification { path })?,
            )
            .map_err(|_| Error::Verification { path })?;
            let candidate_record = candidate_headers
                .get(destination_local)
                .and_then(|record| *record)
                .ok_or(Error::Verification { path })?;
            same_wire_fields_except(source_record.raw(), candidate_record, &[1], budget, path)?;
        }
    }

    // Column headers have no row-affine fields and must remain exactly as
    // authored, including unknown records and their raw field ordering.
    let source_column_object = core::object_at(source, selection.topology.column_headers())
        .map_err(|error| map_core(error, path))?;
    let candidate_column_object = core::object_at(candidate, selection.topology.column_headers())
        .map_err(|error| map_core(error, path))?;
    let source_column_payload =
        message_payload(source_column_object, HEADER_BUCKET_MESSAGE_TYPE, path)?;
    let candidate_column_payload =
        message_payload(candidate_column_object, HEADER_BUCKET_MESSAGE_TYPE, path)?;
    same_wire_fields_exact(
        source_column_payload,
        candidate_column_payload,
        budget,
        path,
    )?;

    // The UID map's row fields are the only rewritten fields.  The complete
    // column arrays and sorted row UID set are checked by `verify_candidate`;
    // this raw comparison also fences unknown/root fields around them.
    let source_uid_object = core::object_at(source, selection.topology.row_uid_map())
        .map_err(|error| map_core(error, path))?;
    let candidate_uid_object = core::object_at(candidate, selection.topology.row_uid_map())
        .map_err(|error| map_core(error, path))?;
    let source_uid_payload = message_payload_any(
        source_uid_object,
        &[
            COLUMN_ROW_UID_MAP_MESSAGE_TYPE,
            COLUMN_ROW_UID_MAP_LEGACY_MESSAGE_TYPE,
        ],
        path,
    )?;
    let candidate_uid_payload = message_payload_any(
        candidate_uid_object,
        &[
            COLUMN_ROW_UID_MAP_MESSAGE_TYPE,
            COLUMN_ROW_UID_MAP_LEGACY_MESSAGE_TYPE,
        ],
        path,
    )?;
    same_wire_fields_except(
        source_uid_payload,
        candidate_uid_payload,
        &[4, 5, 6],
        budget,
        path,
    )?;
    Ok(())
}

#[derive(Clone, Copy)]
struct WireFieldSpan {
    number: u32,
    start: usize,
    end: usize,
}

fn next_wire_field(source: &[u8], cursor: &mut usize, path: Path) -> Result<Option<WireFieldSpan>> {
    if *cursor == source.len() {
        return Ok(None);
    }
    let start = *cursor;
    let remaining = source.get(start..).ok_or(Error::Verification { path })?;
    let (key, key_bytes) =
        decode_varint_from_bytes(remaining).map_err(|_| Error::Verification { path })?;
    if key_bytes != encoded_len(key) {
        return Err(Error::Verification { path });
    }
    let number = key >> 3;
    if number == 0 || number > 0x1fff_ffff {
        return Err(Error::Verification { path });
    }
    let wire_type = key & 7;
    let value_start = start
        .checked_add(key_bytes)
        .ok_or(Error::Verification { path })?;
    let end = match wire_type {
        0 => {
            let (value, value_bytes) = decode_varint_from_bytes(
                source
                    .get(value_start..)
                    .ok_or(Error::Verification { path })?,
            )
            .map_err(|_| Error::Verification { path })?;
            if value_bytes != encoded_len(value) {
                return Err(Error::Verification { path });
            }
            value_start
                .checked_add(value_bytes)
                .ok_or(Error::Verification { path })?
        },
        1 => value_start
            .checked_add(8)
            .ok_or(Error::Verification { path })?,
        2 => {
            let (length, length_bytes) = decode_varint_from_bytes(
                source
                    .get(value_start..)
                    .ok_or(Error::Verification { path })?,
            )
            .map_err(|_| Error::Verification { path })?;
            if length_bytes != encoded_len(length) {
                return Err(Error::Verification { path });
            }
            let length = usize::try_from(length).map_err(|_| Error::Verification { path })?;
            value_start
                .checked_add(length_bytes)
                .and_then(|payload_start| payload_start.checked_add(length))
                .ok_or(Error::Verification { path })?
        },
        5 => value_start
            .checked_add(4)
            .ok_or(Error::Verification { path })?,
        _ => return Err(Error::Verification { path }),
    };
    if end > source.len() {
        return Err(Error::Verification { path });
    }
    *cursor = end;
    Ok(Some(WireFieldSpan {
        number: u32::try_from(number).map_err(|_| Error::Verification { path })?,
        start,
        end,
    }))
}

fn same_wire_fields_exact(
    source: &[u8],
    candidate: &[u8],
    budget: &mut core::Budget,
    path: Path,
) -> Result<()> {
    let bytes = source
        .len()
        .checked_add(candidate.len())
        .ok_or(Error::InvalidSource { path })?;
    budget.work(bytes).map_err(|error| map_core(error, path))?;
    if source == candidate {
        Ok(())
    } else {
        Err(Error::Verification { path })
    }
}

fn same_wire_fields_except(
    source: &[u8],
    candidate: &[u8],
    excluded: &[u32],
    budget: &mut core::Budget,
    path: Path,
) -> Result<()> {
    let bytes = source
        .len()
        .checked_add(candidate.len())
        .ok_or(Error::InvalidSource { path })?;
    budget.work(bytes).map_err(|error| map_core(error, path))?;
    let mut source_cursor = 0;
    let mut candidate_cursor = 0;
    loop {
        let source_field =
            next_nonexcluded_wire_field(source, &mut source_cursor, excluded, budget, path)?;
        let candidate_field =
            next_nonexcluded_wire_field(candidate, &mut candidate_cursor, excluded, budget, path)?;
        match (source_field, candidate_field) {
            (None, None) => return Ok(()),
            (Some(_), None) | (None, Some(_)) => return Err(Error::Verification { path }),
            (Some(source_field), Some(candidate_field)) => {
                if source_field.number != candidate_field.number
                    || source.get(source_field.start..source_field.end)
                        != candidate.get(candidate_field.start..candidate_field.end)
                {
                    return Err(Error::Verification { path });
                }
            },
        }
    }
}

fn next_nonexcluded_wire_field(
    source: &[u8],
    cursor: &mut usize,
    excluded: &[u32],
    budget: &mut core::Budget,
    path: Path,
) -> Result<Option<WireFieldSpan>> {
    loop {
        let field = next_wire_field(source, cursor, path)?;
        let Some(field) = field else {
            return Ok(None);
        };
        budget.fields(1).map_err(|error| map_core(error, path))?;
        if !excluded.contains(&field.number) {
            return Ok(Some(field));
        }
    }
}

/// Reconstruct the source-authoritative object and compare its framing after
/// applying only the changed physical payload messages.  Core locality keeps
/// raw ZIP records and unchanged objects fenced; this owner-level check closes
/// the remaining changed-object hole by retaining the exact ArchiveInfo raw
/// header and canonical framing length while candidate messages are replaced.
fn verify_changed_object_framing(
    source: &Package,
    candidate: &Package,
    allowlist: &core::LocalityAllowlist,
    budget: &mut core::Budget,
    path: Path,
) -> Result<()> {
    let limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    for component_name in allowlist.changed_components() {
        let source_component = core::component(source, component_name.as_ref())
            .map_err(|error| map_core(error, path))?;
        let candidate_component = core::component(candidate, component_name.as_ref())
            .map_err(|error| map_core(error, path))?;
        for source_object in &source_component.archive().objects {
            let identifier = source_object
                .archive_info
                .identifier
                .ok_or(Error::Verification { path })?;
            if allowlist
                .changed_object_ids()
                .binary_search(&identifier)
                .is_err()
            {
                continue;
            }
            let candidate_object = candidate_component
                .archive()
                .object(identifier)
                .ok_or(Error::Verification { path })?;
            if source_object.messages.len() != candidate_object.messages.len() {
                return Err(Error::Verification { path });
            }
            let changed = source_object
                .messages
                .iter()
                .zip(&candidate_object.messages)
                .any(|(source_message, candidate_message)| {
                    source_message.data != candidate_message.data
                });
            if !changed {
                // `verify_physical_locality` already requires exact content
                // for this case, including raw/canonical ArchiveInfo bytes.
                continue;
            }

            let (clone_allocations, clone_bytes, replacement_bytes) =
                archive_clone_reservation(source_object, candidate_object, path)?;
            budget
                .allocations(clone_allocations)
                .map_err(|error| map_core(error, path))?;
            budget
                .retained(clone_bytes)
                .map_err(|error| map_core(error, path))?;
            budget
                .scratch(replacement_bytes)
                .map_err(|error| map_core(error, path))?;
            let mut expected = source_object.clone();
            for (index, (source_message, candidate_message)) in source_object
                .messages
                .iter()
                .zip(&candidate_object.messages)
                .enumerate()
            {
                if source_message.data == candidate_message.data {
                    continue;
                }
                if source_message.type_ != candidate_message.type_
                    || !PHYSICAL_MUTABLE_MESSAGE_TYPES.contains(&source_message.type_)
                {
                    return Err(Error::Verification { path });
                }
                expected
                    .replace_message_preserving_header_with_limits(
                        index,
                        RawMessage {
                            type_: candidate_message.type_,
                            data: candidate_message.data.clone(),
                        },
                        limits,
                    )
                    .map_err(|_| Error::Verification { path })?;
            }
            // Relative offsets are intentionally not stable across package
            // reassembly.  Header/data lengths are stable framing facts and
            // make the canonical varint prefix width observable as well.
            expected.header_length = candidate_object.header_length;
            expected.data_length = candidate_object.data_length;
            if !expected.same_content_ignoring_offsets(candidate_object) {
                return Err(Error::Verification { path });
            }
        }
    }
    Ok(())
}

fn archive_clone_reservation(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    path: Path,
) -> Result<(usize, usize, usize)> {
    let mut allocations = 2usize; // object plus its message-info/message vectors
    let mut retained = size_of::<ArchiveObject>();
    let source_payload = source.messages.iter().try_fold(0usize, |sum, message| {
        sum.checked_add(message.data.len())
            .ok_or(Error::InvalidSource { path })
    })?;
    retained = retained
        .checked_add(source_payload)
        .and_then(|bytes| {
            bytes.checked_add(
                source
                    .archive_info
                    .message_infos
                    .len()
                    .checked_mul(size_of::<litchi_iwa_core::MessageInfo>())?,
            )
        })
        .and_then(|bytes| {
            bytes.checked_add(source.messages.len().checked_mul(size_of::<RawMessage>())?)
        })
        .and_then(|bytes| {
            bytes.checked_add(
                usize::try_from(source.header_length)
                    .map_err(|_| Error::InvalidSource { path })
                    .ok()?,
            )
        })
        .ok_or(Error::InvalidSource { path })?;
    allocations = allocations
        .checked_add(source.messages.len())
        .and_then(|value| value.checked_add(source.archive_info.message_infos.len()))
        .ok_or(Error::InvalidSource { path })?;
    let replacement_bytes = source
        .messages
        .iter()
        .zip(&candidate.messages)
        .filter(|(source_message, candidate_message)| source_message.data != candidate_message.data)
        .try_fold(0usize, |sum, (_, candidate_message)| {
            sum.checked_add(candidate_message.data.len())
                .ok_or(Error::InvalidSource { path })
        })?;
    allocations = allocations
        .checked_add(
            source
                .messages
                .iter()
                .zip(&candidate.messages)
                .filter(|(source_message, candidate_message)| {
                    source_message.data != candidate_message.data
                })
                .count(),
        )
        .ok_or(Error::InvalidSource { path })?;
    let source_header_bytes =
        usize::try_from(source.header_length).map_err(|_| Error::InvalidSource { path })?;
    let candidate_header_bytes =
        usize::try_from(candidate.header_length).map_err(|_| Error::InvalidSource { path })?;
    let header_scratch = source_header_bytes
        .checked_add(candidate_header_bytes)
        .and_then(|bytes| bytes.checked_mul(3))
        .ok_or(Error::InvalidSource { path })?;
    retained = retained
        .checked_add(replacement_bytes)
        .ok_or(Error::InvalidSource { path })?;
    Ok((
        allocations,
        retained,
        replacement_bytes
            .checked_add(header_scratch)
            .ok_or(Error::InvalidSource { path })?,
    ))
}

/// Build a bounded direct-position index for one already validated record
/// envelope.  The old verification path searched the candidate records from
/// the beginning for every source record, turning a sparse or adversarial
/// envelope into quadratic work.  A direct index touches each candidate once
/// and bounds retained memory by the envelope's declared row span.
fn index_records_by_position<'source, I>(
    records: I,
    index_base: u32,
    index_limit: usize,
    record_count: usize,
    budget: &mut core::Budget,
    path: Path,
) -> Result<Vec<Option<&'source [u8]>>>
where
    I: IntoIterator<Item = (u32, &'source [u8])>,
{
    let work = index_limit
        .checked_add(record_count)
        .ok_or(Error::InvalidSource { path })?;
    let retained = index_limit
        .checked_mul(size_of::<Option<&'source [u8]>>())
        .ok_or(Error::InvalidSource { path })?;
    budget.work(work).map_err(|error| map_core(error, path))?;
    if index_limit != 0 {
        budget
            .allocations(1)
            .map_err(|error| map_core(error, path))?;
        budget
            .retained(retained)
            .map_err(|error| map_core(error, path))?;
    }
    let mut index = Vec::new();
    index
        .try_reserve_exact(index_limit)
        .map_err(|_| Error::Allocation {
            amount: index_limit,
            path,
        })?;
    index.resize_with(index_limit, || None);

    populate_record_index(records, index_base, &mut index, record_count, path)?;
    Ok(index)
}

fn populate_record_index<'source, I>(
    records: I,
    index_base: u32,
    index: &mut [Option<&'source [u8]>],
    record_count: usize,
    path: Path,
) -> Result<()>
where
    I: IntoIterator<Item = (u32, &'source [u8])>,
{
    let base = usize::try_from(index_base).map_err(|_| Error::Verification { path })?;
    let mut observed = 0usize;
    for (position, raw) in records {
        observed = observed
            .checked_add(1)
            .ok_or(Error::InvalidSource { path })?;
        let local = usize::try_from(position)
            .ok()
            .and_then(|position| position.checked_sub(base))
            .ok_or(Error::Verification { path })?;
        let slot = index.get_mut(local).ok_or(Error::Verification { path })?;
        if slot.replace(raw).is_some() {
            return Err(Error::Verification { path });
        }
    }
    if observed != record_count {
        return Err(Error::Verification { path });
    }
    Ok(())
}

fn changed_object_ids(
    selection: &PhysicalSelection,
    row_header_indices: &[u32],
    destination_by_source: &[u32],
    budget: &mut core::Budget,
) -> Result<Vec<u64>> {
    let path = selection.path();
    let capacity = 1usize
        .checked_add(selection.topology.tile_references().len())
        .and_then(|value| value.checked_add(selection.topology.row_header_buckets().len()))
        .ok_or(Error::InvalidSource { path })?;
    budget
        .allocations(1)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(
            capacity
                .checked_mul(size_of::<u64>())
                .ok_or(Error::InvalidSource { path })?,
        )
        .map_err(|error| map_core(error, path))?;
    let mut ids = Vec::new();
    ids.try_reserve_exact(capacity)
        .map_err(|_| Error::Allocation {
            amount: capacity,
            path,
        })?;
    ids.push(selection.topology.row_uid_map().identifier);
    let tile_size = selection.topology.tile_size();
    for tile in selection.topology.tile_references() {
        let changed =
            destination_by_source
                .iter()
                .copied()
                .enumerate()
                .any(|(source, destination)| {
                    u32::try_from(source).is_ok_and(|source| {
                        source != destination && source / tile_size == tile.tile_id
                    })
                });
        if changed {
            ids.push(tile.location.identifier);
        }
    }
    for bucket in selection.topology.row_header_buckets() {
        let changed = row_header_indices.iter().copied().any(|source| {
            usize::try_from(source)
                .ok()
                .and_then(|index| destination_by_source.get(index).copied())
                .is_some_and(|destination| {
                    source != destination && source / HEADER_BUCKET_ROWS == bucket.bucket_index
                })
        });
        if changed {
            ids.push(bucket.location.identifier);
        }
    }
    ids.sort_unstable();
    ids.dedup();
    Ok(ids)
}

impl PhysicalSelection {
    fn topology_component_count(&self) -> usize {
        self.topology.mutable_component_count()
    }
}

fn compare_scalars(left: &PhysicalScalar, right: &PhysicalScalar) -> CompareOrdering {
    match (left, right) {
        (PhysicalScalar::Text(left), PhysicalScalar::Text(right)) => left.cmp(right),
        (PhysicalScalar::Number(left), PhysicalScalar::Number(right))
        | (PhysicalScalar::Date(left), PhysicalScalar::Date(right))
        | (PhysicalScalar::Duration(left), PhysicalScalar::Duration(right)) => {
            if left == right {
                CompareOrdering::Equal
            } else {
                left.total_cmp(right)
            }
        },
        (PhysicalScalar::Boolean(left), PhysicalScalar::Boolean(right)) => left.cmp(right),
        _ => CompareOrdering::Equal,
    }
}

fn invert_permutation(values: &[u32]) -> Option<Vec<usize>> {
    let mut inverse = vec![usize::MAX; values.len()];
    for (source, destination) in values.iter().copied().enumerate() {
        let destination = usize::try_from(destination).ok()?;
        if destination >= values.len() || inverse[destination] != usize::MAX {
            return None;
        }
        inverse[destination] = source;
    }
    inverse
        .iter()
        .all(|value| *value != usize::MAX)
        .then_some(inverse)
}

fn invert_permutation_checked(
    values: &[u32],
    budget: &mut core::Budget,
    path: Path,
) -> Result<Vec<usize>> {
    let bytes = values
        .len()
        .checked_mul(size_of::<usize>())
        .ok_or(Error::InvalidSource { path })?;
    budget
        .allocations(1)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(bytes)
        .map_err(|error| map_core(error, path))?;
    invert_permutation(values).ok_or(Error::InvalidPermutation { path })
}

fn copy_permutation_checked(
    values: &[u32],
    budget: &mut core::Budget,
    path: Path,
) -> Result<Vec<u32>> {
    let bytes = values
        .len()
        .checked_mul(size_of::<u32>())
        .ok_or(Error::InvalidSource { path })?;
    budget
        .allocations(1)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(bytes)
        .map_err(|error| map_core(error, path))?;
    let mut copied = Vec::new();
    copied
        .try_reserve_exact(values.len())
        .map_err(|_| Error::Allocation {
            amount: values.len(),
            path,
        })?;
    copied.extend_from_slice(values);
    Ok(copied)
}

fn invert_permutation_u32_checked(
    values: &[u32],
    budget: &mut core::Budget,
    path: Path,
) -> Result<Vec<u32>> {
    let bytes = values
        .len()
        .checked_mul(size_of::<u32>())
        .ok_or(Error::InvalidSource { path })?;
    budget
        .allocations(1)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(bytes)
        .map_err(|error| map_core(error, path))?;
    let mut inverse = Vec::new();
    inverse
        .try_reserve_exact(values.len())
        .map_err(|_| Error::Allocation {
            amount: values.len(),
            path,
        })?;
    inverse.resize(values.len(), u32::MAX);
    for (source, destination) in values.iter().copied().enumerate() {
        let destination =
            usize::try_from(destination).map_err(|_| Error::InvalidPermutation { path })?;
        let slot = inverse
            .get_mut(destination)
            .ok_or(Error::InvalidPermutation { path })?;
        if *slot != u32::MAX {
            return Err(Error::InvalidPermutation { path });
        }
        *slot = u32::try_from(source).map_err(|_| Error::InvalidPermutation { path })?;
    }
    if inverse.contains(&u32::MAX) {
        return Err(Error::InvalidPermutation { path });
    }
    Ok(inverse)
}

fn body_row_count(
    target: &core::Target,
    topology: &core::PhysicalTableTopology,
    path: Path,
) -> Result<usize> {
    usize::try_from(target.rows)
        .ok()
        .and_then(|rows| {
            usize::try_from(topology.header_rows())
                .ok()
                .and_then(|head| rows.checked_sub(head))
        })
        .and_then(|rows| {
            usize::try_from(topology.footer_rows())
                .ok()
                .and_then(|foot| rows.checked_sub(foot))
        })
        .ok_or(Error::InvalidSource { path })
}

fn validate_order_columns(order: &Order, columns: u32, path: Path) -> Result<()> {
    for rule in order.rules() {
        if rule.column().native_value() >= columns {
            return Err(Error::InvalidSource { path });
        }
    }
    Ok(())
}

fn message_payload(object: &ArchiveObject, message_type: u32, path: Path) -> Result<&[u8]> {
    let mut selected = None;
    for message in &object.messages {
        if message.type_ == message_type {
            if selected.is_some() {
                return Err(Error::InvalidSource { path });
            }
            selected = Some(message.data.as_slice());
        }
    }
    selected.ok_or(Error::InvalidSource { path })
}

fn message_payload_any<'a>(
    object: &'a ArchiveObject,
    message_types: &[u32],
    path: Path,
) -> Result<&'a [u8]> {
    let mut selected = None;
    for message in &object.messages {
        if message_types.contains(&message.type_) {
            if selected.is_some() {
                return Err(Error::InvalidSource { path });
            }
            selected = Some(message.data.as_slice());
        }
    }
    selected.ok_or(Error::InvalidSource { path })
}

fn find_message_type(object: &ArchiveObject, message_types: &[u32], path: Path) -> Result<u32> {
    let mut selected = None;
    for message in &object.messages {
        if message_types.contains(&message.type_) {
            if selected.is_some() {
                return Err(Error::InvalidSource { path });
            }
            selected = Some(message.type_);
        }
    }
    selected.ok_or(Error::InvalidSource { path })
}

fn replace_message(
    archive: &mut Archive,
    identifier: u64,
    message_type: u32,
    data: Vec<u8>,
    limits: litchi_iwa_core::Limits,
    path: Path,
) -> Result<()> {
    let object = archive
        .object_mut(identifier)
        .ok_or(Error::InvalidSource { path })?;
    let mut index = None;
    for (candidate, message) in object.messages.iter().enumerate() {
        if message.type_ == message_type && index.replace(candidate).is_some() {
            return Err(Error::InvalidSource { path });
        }
    }
    let index = index.ok_or(Error::InvalidSource { path })?;
    object
        .replace_message_preserving_header_with_limits(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
            limits,
        )
        .map_err(|_| Error::InvalidSource { path })?;
    Ok(())
}

fn physical_options(
    package: &Package,
    budget: &core::Budget,
    source_len: usize,
    records: u32,
    columns: u32,
) -> Result<physical_codec::DecodeOptions> {
    let limits = budget
        .residual(package)
        .map_err(|error| map_core(error, Path::Package))?;
    let row_count = usize::try_from(records).map_err(|_| Error::InvalidSource {
        path: Path::Package,
    })?;
    let column_count = usize::try_from(columns).map_err(|_| Error::InvalidSource {
        path: Path::Package,
    })?;
    let cell_elements = row_count
        .checked_mul(column_count)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let uid_records = row_count.checked_mul(2).ok_or(Error::InvalidSource {
        path: Path::Package,
    })?;
    let max_records = row_count.max(uid_records);
    let uid_elements = row_count
        .checked_add(column_count)
        .and_then(|value| value.checked_mul(6))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let max_elements = cell_elements.max(uid_elements);
    Ok(physical_codec::DecodeOptions::new(
        limits.max_input_bytes().min(source_len.max(1)),
        limits.max_fields(),
        budget
            .remaining_work()
            .map_err(|error| map_core(error, Path::Package))?
            .min(limits.max_rewrite_work()),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
        max_records,
        max_elements,
        budget
            .remaining_output()
            .map_err(|error| map_core(error, Path::Package))?
            .min(limits.max_output_bytes()),
        budget
            .remaining_scratch()
            .map_err(|error| map_core(error, Path::Package))?,
    ))
}

fn storage_options(
    package: &Package,
    budget: &core::Budget,
    source_len: usize,
) -> Result<storage_codec::DecodeOptions> {
    let limits = budget
        .residual(package)
        .map_err(|error| map_core(error, Path::Package))?;
    let references = budget
        .remaining_references()
        .map_err(|error| map_core(error, Path::Package))?;
    Ok(storage_codec::DecodeOptions::new(
        limits.max_input_bytes().min(source_len.max(1)),
        limits.max_fields(),
        limits.max_rewrite_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
        references,
        limits.max_input_bytes(),
    ))
}

fn charge_decode_report(
    report: physical_codec::DecodeReport,
    budget: &mut core::Budget,
    path: Path,
) -> Result<()> {
    budget
        .input(report.source_bytes())
        .map_err(|error| map_core(error, path))?;
    budget
        .fields(report.fields())
        .map_err(|error| map_core(error, path))?;
    budget
        .work(report.work_bytes())
        .map_err(|error| map_core(error, path))?;
    budget
        .nesting(report.max_depth() as usize)
        .map_err(|error| map_core(error, path))?;
    let allocations = report
        .records()
        .checked_add(report.elements())
        .ok_or(Error::InvalidSource { path })?;
    budget
        .allocations(allocations)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(report.source_bytes())
        .map_err(|error| map_core(error, path))?;
    Ok(())
}

fn charge_rewrite_requirements(
    requirements: physical_codec::RewriteRequirements,
    budget: &mut core::Budget,
    path: Path,
) -> Result<()> {
    charge_decode_report(requirements.source(), budget, path)?;
    budget
        .output(requirements.output_bytes())
        .map_err(|error| map_core(error, path))?;
    budget
        .work(requirements.work_bytes())
        .map_err(|error| map_core(error, path))?;
    budget
        .fields(requirements.fields())
        .map_err(|error| map_core(error, path))?;
    let allocations = requirements
        .records()
        .checked_add(requirements.elements())
        .ok_or(Error::InvalidSource { path })?;
    budget
        .allocations(allocations)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(requirements.output_bytes())
        .map_err(|error| map_core(error, path))?;
    Ok(())
}

fn charge_storage_report(
    report: storage_codec::DecodeReport,
    budget: &mut core::Budget,
    path: Path,
) -> Result<()> {
    budget
        .input(report.source_bytes())
        .map_err(|error| map_core(error, path))?;
    budget
        .fields(report.fields())
        .map_err(|error| map_core(error, path))?;
    budget
        .work(report.work_bytes())
        .map_err(|error| map_core(error, path))?;
    budget
        .references(report.references())
        .map_err(|error| map_core(error, path))?;
    budget
        .nesting(report.max_depth() as usize)
        .map_err(|error| map_core(error, path))?;
    budget
        .retained(report.text_bytes())
        .map_err(|error| map_core(error, path))?;
    Ok(())
}

fn map_physical_codec(error: physical_codec::DecodeError, path: Path) -> Error {
    let Some(limit) = error.resource_limit() else {
        return Error::InvalidSource { path };
    };
    match limit {
        physical_codec::DecodeLimit::Bytes { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireBytes,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        physical_codec::DecodeLimit::Fields { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        physical_codec::DecodeLimit::Work { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        physical_codec::DecodeLimit::Nesting { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
            path,
        },
        physical_codec::DecodeLimit::Records { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::PhysicalRows,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        physical_codec::DecodeLimit::Elements { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::PhysicalColumns,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        physical_codec::DecodeLimit::OutputBytes { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireOutputBytes,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        physical_codec::DecodeLimit::ScratchBytes { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireScratchBytes,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        physical_codec::DecodeLimit::Allocation { requested } => Error::Allocation {
            amount: requested,
            path,
        },
        _ => Error::InvalidSource { path },
    }
}

fn map_storage_codec(error: storage_codec::DecodeError, path: Path) -> Error {
    let Some(limit) = error.resource_limit() else {
        return Error::InvalidSource { path };
    };
    match limit {
        storage_codec::DecodeLimit::Bytes { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireBytes,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        storage_codec::DecodeLimit::Fields { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        storage_codec::DecodeLimit::Text { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireBytes,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        storage_codec::DecodeLimit::References { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::PayloadReferences,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        storage_codec::DecodeLimit::Work { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        storage_codec::DecodeLimit::Nesting { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
            path,
        },
        storage_codec::DecodeLimit::Allocation { requested } => Error::Allocation {
            amount: requested,
            path,
        },
        storage_codec::DecodeLimit::Retained { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireRetainedBytes,
            observed: observed as u64,
            maximum: maximum as u64,
            path,
        },
        _ => Error::InvalidSource { path },
    }
}

fn map_core(error: core::Error, path: Path) -> Error {
    match error {
        core::Error::UnsupportedSource => Error::UnsupportedSource,
        core::Error::UnsupportedDependency => Error::UnsupportedDependency {
            path,
            feature: UnsupportedFeature::RowAffineDependency,
        },
        core::Error::UnsupportedTopology => Error::UnsupportedTopology { path },
        core::Error::AmbiguousSelector => Error::AmbiguousSelector,
        core::Error::EmptySlideName => Error::EmptySlideName,
        core::Error::SlideNameNotFound => Error::SlideNameNotFound,
        core::Error::SlidePositionNotFound(position) => Error::SlidePositionNotFound { position },
        core::Error::TablePositionNotFound(position) => Error::TablePositionNotFound { position },
        core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: map_core_limit(kind),
            observed,
            maximum,
            path,
        },
        core::Error::Allocation(amount) => Error::Allocation { amount, path },
        core::Error::InvalidSource
        | core::Error::Read
        | core::Error::Wire
        | core::Error::Codec
        | core::Error::Archive
        | core::Error::Verification => Error::InvalidSource { path },
    }
}

fn map_core_limit(kind: core::LimitKind) -> LimitKind {
    match kind {
        core::LimitKind::InputBytes => LimitKind::InputBytes,
        core::LimitKind::OutputBytes => LimitKind::OutputBytes,
        core::LimitKind::Entries => LimitKind::Entries,
        core::LimitKind::EntryBytes => LimitKind::EntryBytes,
        core::LimitKind::TotalBytes => LimitKind::TotalEntryBytes,
        core::LimitKind::PayloadObjects => LimitKind::PayloadObjects,
        core::LimitKind::PayloadMessages => LimitKind::PayloadMessages,
        core::LimitKind::References => LimitKind::PayloadReferences,
        core::LimitKind::WireFields => LimitKind::WireFields,
        core::LimitKind::WireNesting => LimitKind::WireNesting,
        core::LimitKind::WireWork => LimitKind::WireWork,
        core::LimitKind::Allocations => LimitKind::WireAllocations,
        core::LimitKind::Retained => LimitKind::WireRetainedBytes,
        core::LimitKind::Scratch => LimitKind::WireScratchBytes,
        core::LimitKind::Components => LimitKind::Components,
    }
}

fn map_sort_error(error: SlideTableSortError, path: Path) -> Error {
    match error {
        SlideTableSortError::UnsupportedSource => Error::UnsupportedSource,
        SlideTableSortError::UnsupportedDependency => Error::UnsupportedDependency {
            path,
            feature: UnsupportedFeature::RowAffineDependency,
        },
        SlideTableSortError::UnsupportedTopology => Error::UnsupportedTopology { path },
        SlideTableSortError::AmbiguousSelector => Error::AmbiguousSelector,
        SlideTableSortError::EmptySlideName => Error::EmptySlideName,
        SlideTableSortError::SlideNameNotFound => Error::SlideNameNotFound,
        SlideTableSortError::SlidePositionNotFound { position } => {
            Error::SlidePositionNotFound { position }
        },
        SlideTableSortError::TablePositionNotFound { position } => {
            Error::TablePositionNotFound { position }
        },
        SlideTableSortError::Locked => Error::TableLocked { path },
        SlideTableSortError::InvalidSource => Error::InvalidSource { path },
        SlideTableSortError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: map_sort_limit(kind),
            observed,
            maximum,
            path,
        },
        SlideTableSortError::Allocation { amount } => Error::Allocation { amount, path },
        SlideTableSortError::Verification => Error::Verification { path },
        SlideTableSortError::PatchConflict => Error::PatchConflict,
    }
}

fn map_sort_limit(kind: super::slide_table_sort_order::SlideTableSortLimitKind) -> LimitKind {
    match kind {
        super::slide_table_sort_order::SlideTableSortLimitKind::InputBytes => LimitKind::InputBytes,
        super::slide_table_sort_order::SlideTableSortLimitKind::OutputBytes => {
            LimitKind::OutputBytes
        },
        super::slide_table_sort_order::SlideTableSortLimitKind::Entries => LimitKind::Entries,
        super::slide_table_sort_order::SlideTableSortLimitKind::EntryBytes => LimitKind::EntryBytes,
        super::slide_table_sort_order::SlideTableSortLimitKind::TotalBytes => {
            LimitKind::TotalEntryBytes
        },
        super::slide_table_sort_order::SlideTableSortLimitKind::PayloadObjects => {
            LimitKind::PayloadObjects
        },
        super::slide_table_sort_order::SlideTableSortLimitKind::PayloadMessages => {
            LimitKind::PayloadMessages
        },
        super::slide_table_sort_order::SlideTableSortLimitKind::References => {
            LimitKind::PayloadReferences
        },
        super::slide_table_sort_order::SlideTableSortLimitKind::WireFields => LimitKind::WireFields,
        super::slide_table_sort_order::SlideTableSortLimitKind::WireNesting => {
            LimitKind::WireNesting
        },
        super::slide_table_sort_order::SlideTableSortLimitKind::WireWork => LimitKind::WireWork,
        super::slide_table_sort_order::SlideTableSortLimitKind::Allocations => {
            LimitKind::WireAllocations
        },
        super::slide_table_sort_order::SlideTableSortLimitKind::Retained => {
            LimitKind::WireRetainedBytes
        },
        super::slide_table_sort_order::SlideTableSortLimitKind::Scratch => {
            LimitKind::WireScratchBytes
        },
        super::slide_table_sort_order::SlideTableSortLimitKind::Components => LimitKind::Components,
    }
}

fn map_read(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::References => LimitKind::PayloadReferences,
                SemanticLimitKind::Objects => LimitKind::PayloadObjects,
                SemanticLimitKind::Slides => LimitKind::PayloadItems,
                SemanticLimitKind::TextStorages
                | SemanticLimitKind::TextFragments
                | SemanticLimitKind::TextBytes => LimitKind::PayloadBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
            path: Path::Package,
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::LimitExceeded {
            kind: match kind {
                PayloadLimitKind::Bytes => LimitKind::InputBytes,
                PayloadLimitKind::Fields => LimitKind::WireFields,
                PayloadLimitKind::Nesting => LimitKind::WireNesting,
                PayloadLimitKind::Work => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
            path: Path::Package,
        },
        ReadError::Allocation { amount, .. } => Error::Allocation {
            amount,
            path: Path::Package,
        },
        ReadError::Archive(error) => map_archive(error),
        _ => Error::InvalidSource {
            path: Path::Package,
        },
    }
}

fn map_archive(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => LimitKind::TotalEntryBytes,
                _ => LimitKind::TransactionWork,
            },
            observed,
            maximum,
            path: Path::Package,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation {
            amount,
            path: Path::Package,
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_archive(error),
        _ => Error::InvalidSource {
            path: Path::Package,
        },
    }
}

fn map_core_archive(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => LimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => LimitKind::PayloadMessages,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                _ => LimitKind::TransactionWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
            path: Path::Package,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => Error::Allocation {
            amount: requested,
            path: Path::Package,
        },
        _ => Error::InvalidSource {
            path: Path::Package,
        },
    }
}

fn preview_deletion_count(source: &Package, candidate: &Package) -> usize {
    let source = core::physical_catalog(source).ok().and_then(|catalog| {
        super::rendering_invalidation::root_preview_deletions(catalog.package()).ok()
    });
    let candidate = core::physical_catalog(candidate).ok().and_then(|catalog| {
        super::rendering_invalidation::root_preview_deletions(catalog.package()).ok()
    });
    match (source, candidate) {
        (Some(source), Some(candidate)) => source.len().saturating_sub(candidate.len()),
        _ => 0,
    }
}

#[cfg(test)]
mod tests {
    use super::{Path, populate_record_index};

    #[test]
    fn direct_record_index_handles_reverse_adversarial_order_in_one_pass() {
        const RECORDS: usize = 16_384;
        let mut visits = 0usize;
        let mut index = vec![None; RECORDS];
        let result = populate_record_index(
            (0..u32::try_from(RECORDS).expect("test size fits u32"))
                .rev()
                .map(|position| {
                    visits = visits.saturating_add(1);
                    (position, b"record".as_slice())
                }),
            0,
            &mut index,
            RECORDS,
            Path::Package,
        );

        assert!(result.is_ok());
        assert_eq!(visits, RECORDS);
        assert_eq!(index.len(), RECORDS);
        for position in (0..RECORDS).rev() {
            assert!(index.get(position).is_some_and(|record| record.is_some()));
        }
    }
}
