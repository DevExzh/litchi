//! Strict generated-free Numbers formula dependency projections.
//!
//! Repeated dependency records are streamed through a caller visitor. Private
//! Buffa views provide borrowed parity only after the canonical handwritten
//! pass has validated every selected envelope with one aggregate budget.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Dependency wire helpers stay beside the snapshots they construct."
)]

use core::fmt;

use crate::buffa_numbers_table_cell_dependency_generated::LitchiIwaTableCellDependencyProjection as projection;
use crate::numbers_table_cell_storage_codec as wire;

pub use wire::{
    DecodeError, DecodeLimit, DecodeOptions, DecodeReport, ReferenceRecord, ReferenceSnapshot,
};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct UuidSnapshot {
    lower: u64,
    upper: u64,
}
impl UuidSnapshot {
    #[must_use]
    pub const fn lower(self) -> u64 {
        self.lower
    }
    #[must_use]
    pub const fn upper(self) -> u64 {
        self.upper
    }
}

/// Presence-preserving canonical `TSP.CFUUIDArchive` identity.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct CfuuidSnapshot<'source> {
    uuid_bytes: Option<&'source [u8]>,
    words: [Option<u32>; 4],
}
impl<'source> CfuuidSnapshot<'source> {
    #[must_use]
    pub const fn uuid_bytes(self) -> Option<&'source [u8]> {
        self.uuid_bytes
    }
    #[must_use]
    pub const fn words(self) -> [Option<u32>; 4] {
        self.words
    }
}

/// Exact optional forms of a native cell coordinate.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellCoordinateSnapshot {
    packed_data: Option<u32>,
    column: Option<u32>,
    row: Option<u32>,
}
impl CellCoordinateSnapshot {
    #[must_use]
    pub const fn packed_data(self) -> Option<u32> {
        self.packed_data
    }
    #[must_use]
    pub const fn column(self) -> Option<u32> {
        self.column
    }
    #[must_use]
    pub const fn row(self) -> Option<u32> {
        self.row
    }
}

/// Presence-preserving rectangle size; absent dimensions have native default 1.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ColumnRowSizeSnapshot {
    num_columns: Option<u32>,
    num_rows: Option<u32>,
}
impl ColumnRowSizeSnapshot {
    #[must_use]
    pub const fn num_columns(self) -> Option<u32> {
        self.num_columns
    }
    #[must_use]
    pub const fn num_rows(self) -> Option<u32> {
        self.num_rows
    }
}

/// Typed tiled-range rectangle.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellRectSnapshot {
    origin: CellCoordinateSnapshot,
    size: ColumnRowSizeSnapshot,
}
impl CellRectSnapshot {
    #[must_use]
    pub const fn origin(self) -> CellCoordinateSnapshot {
        self.origin
    }
    #[must_use]
    pub const fn size(self) -> ColumnRowSizeSnapshot {
        self.size
    }
}

/// Exact inclusive coordinate rectangle.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RangeCoordinateSnapshot {
    top_left_column: u32,
    top_left_row: u32,
    bottom_right_column: u32,
    bottom_right_row: u32,
}
impl RangeCoordinateSnapshot {
    #[must_use]
    pub const fn top_left_column(self) -> u32 {
        self.top_left_column
    }
    #[must_use]
    pub const fn top_left_row(self) -> u32 {
        self.top_left_row
    }
    #[must_use]
    pub const fn bottom_right_column(self) -> u32 {
        self.bottom_right_column
    }
    #[must_use]
    pub const fn bottom_right_row(self) -> u32 {
        self.bottom_right_row
    }
}

/// Table-identified range dependency.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct RangeReferenceSnapshot<'source> {
    table_id: CfuuidSnapshot<'source>,
    range: RangeCoordinateSnapshot,
}
impl<'source> RangeReferenceSnapshot<'source> {
    #[must_use]
    pub const fn table_id(self) -> CfuuidSnapshot<'source> {
        self.table_id
    }
    #[must_use]
    pub const fn range(self) -> RangeCoordinateSnapshot {
        self.range
    }
}

/// Internal-owner range dependency.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct InternalRangeReferenceSnapshot {
    owner_id: u32,
    range: RangeCoordinateSnapshot,
}
impl InternalRangeReferenceSnapshot {
    #[must_use]
    pub const fn owner_id(self) -> u32 {
        self.owner_id
    }
    #[must_use]
    pub const fn range(self) -> RangeCoordinateSnapshot {
        self.range
    }
}

/// One source-ordered scalar from the five parallel expanded-edge arrays.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ExpandedEdgeComponent {
    kind: ExpandedEdgeKind,
    index: usize,
    value: u32,
}
impl ExpandedEdgeComponent {
    #[must_use]
    pub const fn kind(self) -> ExpandedEdgeKind {
        self.kind
    }
    #[must_use]
    pub const fn index(self) -> usize {
        self.index
    }
    #[must_use]
    pub const fn value(self) -> u32 {
        self.value
    }
}

/// Canonical expanded-edge array identity.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExpandedEdgeKind {
    LocalRow,
    LocalColumn,
    ExternalRow,
    ExternalColumn,
    InternalOwner,
}

/// Validated cardinalities of the parallel expanded-edge arrays.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ExpandedEdgesSnapshot {
    local: usize,
    external: usize,
}
impl ExpandedEdgesSnapshot {
    #[must_use]
    pub const fn local(self) -> usize {
        self.local
    }
    #[must_use]
    pub const fn external(self) -> usize {
        self.external
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub struct CalculationEngineSnapshot<'source> {
    base_date_1904: Option<bool>,
    dependency_tracker: &'source [u8],
    named_reference_manager: Option<ReferenceSnapshot>,
    remote_data_store: Option<ReferenceSnapshot>,
    header_name_manager: Option<ReferenceSnapshot>,
    refs_to_dirty: Option<ReferenceSnapshot>,
}
impl<'source> CalculationEngineSnapshot<'source> {
    #[must_use]
    pub const fn base_date_1904(self) -> Option<bool> {
        self.base_date_1904
    }
    #[must_use]
    pub const fn dependency_tracker(self) -> &'source [u8] {
        self.dependency_tracker
    }
    #[must_use]
    pub const fn named_reference_manager(self) -> Option<ReferenceSnapshot> {
        self.named_reference_manager
    }
    #[must_use]
    pub const fn remote_data_store(self) -> Option<ReferenceSnapshot> {
        self.remote_data_store
    }
    #[must_use]
    pub const fn header_name_manager(self) -> Option<ReferenceSnapshot> {
        self.header_name_manager
    }
    #[must_use]
    pub const fn refs_to_dirty(self) -> Option<ReferenceSnapshot> {
        self.refs_to_dirty
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub struct DependencyTrackerSnapshot<'source> {
    owner_id_map: Option<&'source [u8]>,
    number_of_formulas: Option<u64>,
}
impl<'source> DependencyTrackerSnapshot<'source> {
    #[must_use]
    pub const fn owner_id_map(self) -> Option<&'source [u8]> {
        self.owner_id_map
    }
    #[must_use]
    pub const fn number_of_formulas(self) -> Option<u64> {
        self.number_of_formulas
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub struct FormulaOwnerDependenciesSnapshot<'source> {
    formula_owner_uid: UuidSnapshot,
    internal_formula_owner_id: u32,
    owner_kind: Option<u32>,
    cell_dependencies: Option<&'source [u8]>,
    range_dependencies: Option<&'source [u8]>,
    volatile_dependencies: Option<&'source [u8]>,
    spanning_column_dependencies: Option<&'source [u8]>,
    spanning_row_dependencies: Option<&'source [u8]>,
    whole_owner_dependencies: Option<&'source [u8]>,
    cell_errors: Option<&'source [u8]>,
    formula_owner: Option<ReferenceSnapshot>,
    base_owner_uid: Option<UuidSnapshot>,
    tiled_cell_dependencies: Option<&'source [u8]>,
    uuid_references: Option<&'source [u8]>,
    tiled_range_dependencies: Option<&'source [u8]>,
    spill_range_sizes: Option<&'source [u8]>,
}

macro_rules! owner_accessors {
    ($(($name:ident, $ty:ty)),+ $(,)?) => {$(
        #[must_use]
        pub const fn $name(self) -> $ty { self.$name }
    )+};
}
impl<'source> FormulaOwnerDependenciesSnapshot<'source> {
    owner_accessors!(
        (formula_owner_uid, UuidSnapshot),
        (internal_formula_owner_id, u32),
        (owner_kind, Option<u32>),
        (cell_dependencies, Option<&'source [u8]>),
        (range_dependencies, Option<&'source [u8]>),
        (volatile_dependencies, Option<&'source [u8]>),
        (spanning_column_dependencies, Option<&'source [u8]>),
        (spanning_row_dependencies, Option<&'source [u8]>),
        (whole_owner_dependencies, Option<&'source [u8]>),
        (cell_errors, Option<&'source [u8]>),
        (formula_owner, Option<ReferenceSnapshot>),
        (base_owner_uid, Option<UuidSnapshot>),
        (tiled_cell_dependencies, Option<&'source [u8]>),
        (uuid_references, Option<&'source [u8]>),
        (tiled_range_dependencies, Option<&'source [u8]>),
        (spill_range_sizes, Option<&'source [u8]>)
    );
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub struct CellRecordSnapshot<'source> {
    column: u32,
    row: u32,
    dirty_self_plus_precedents_count: Option<u64>,
    is_in_a_cycle: Option<bool>,
    has_calculated_precedents: Option<bool>,
    expanded_edges: Option<&'source [u8]>,
    expanded_edges_snapshot: Option<ExpandedEdgesSnapshot>,
}
impl<'source> CellRecordSnapshot<'source> {
    #[must_use]
    pub const fn column(self) -> u32 {
        self.column
    }
    #[must_use]
    pub const fn row(self) -> u32 {
        self.row
    }
    #[must_use]
    pub const fn dirty_self_plus_precedents_count(self) -> Option<u64> {
        self.dirty_self_plus_precedents_count
    }
    #[must_use]
    pub const fn is_in_a_cycle(self) -> Option<bool> {
        self.is_in_a_cycle
    }
    #[must_use]
    pub const fn has_calculated_precedents(self) -> Option<bool> {
        self.has_calculated_precedents
    }
    #[must_use]
    pub const fn expanded_edges(self) -> Option<&'source [u8]> {
        self.expanded_edges
    }
    #[must_use]
    pub const fn expanded_edges_snapshot(self) -> Option<ExpandedEdgesSnapshot> {
        self.expanded_edges_snapshot
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellRecordTileSnapshot {
    internal_owner_id: u32,
    tile_column_begin: u32,
    tile_row_begin: u32,
}
impl CellRecordTileSnapshot {
    #[must_use]
    pub const fn internal_owner_id(self) -> u32 {
        self.internal_owner_id
    }
    #[must_use]
    pub const fn tile_column_begin(self) -> u32 {
        self.tile_column_begin
    }
    #[must_use]
    pub const fn tile_row_begin(self) -> u32 {
        self.tile_row_begin
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub struct RangeBackDependencySnapshot<'source> {
    cell_coord_row: u32,
    cell_coord_column: u32,
    range_reference: Option<&'source [u8]>,
    internal_range_reference: Option<&'source [u8]>,
    decoded_range_reference: Option<RangeReferenceSnapshot<'source>>,
    decoded_internal_range_reference: Option<InternalRangeReferenceSnapshot>,
}
impl<'source> RangeBackDependencySnapshot<'source> {
    #[must_use]
    pub const fn cell_coord_row(self) -> u32 {
        self.cell_coord_row
    }
    #[must_use]
    pub const fn cell_coord_column(self) -> u32 {
        self.cell_coord_column
    }
    #[must_use]
    pub const fn range_reference(self) -> Option<&'source [u8]> {
        self.range_reference
    }
    #[must_use]
    pub const fn internal_range_reference(self) -> Option<&'source [u8]> {
        self.internal_range_reference
    }
    #[must_use]
    pub const fn decoded_range_reference(self) -> Option<RangeReferenceSnapshot<'source>> {
        self.decoded_range_reference
    }
    #[must_use]
    pub const fn decoded_internal_range_reference(self) -> Option<InternalRangeReferenceSnapshot> {
        self.decoded_internal_range_reference
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RangePrecedentsTileSnapshot {
    to_owner_id: u32,
}
impl RangePrecedentsTileSnapshot {
    #[must_use]
    pub const fn to_owner_id(self) -> u32 {
        self.to_owner_id
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub struct FromToRangeSnapshot<'source> {
    from_coord: &'source [u8],
    refers_to_rect: &'source [u8],
    decoded_from_coord: CellCoordinateSnapshot,
    decoded_refers_to_rect: CellRectSnapshot,
}
impl<'source> FromToRangeSnapshot<'source> {
    #[must_use]
    pub const fn from_coord(self) -> &'source [u8] {
        self.from_coord
    }
    #[must_use]
    pub const fn refers_to_rect(self) -> &'source [u8] {
        self.refers_to_rect
    }
    #[must_use]
    pub const fn decoded_from_coord(self) -> CellCoordinateSnapshot {
        self.decoded_from_coord
    }
    #[must_use]
    pub const fn decoded_refers_to_rect(self) -> CellRectSnapshot {
        self.decoded_refers_to_rect
    }
}

macro_rules! impl_redacted_debug {
    ($($snapshot:ident),+ $(,)?) => {$(
        impl fmt::Debug for $snapshot<'_> {
            fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
                formatter.write_str(concat!(stringify!($snapshot), " { payloads: <redacted> }"))
            }
        }
    )+};
}

impl_redacted_debug!(
    CfuuidSnapshot,
    RangeReferenceSnapshot,
    CalculationEngineSnapshot,
    DependencyTrackerSnapshot,
    FormulaOwnerDependenciesSnapshot,
    CellRecordSnapshot,
    RangeBackDependencySnapshot,
    FromToRangeSnapshot,
);

/// Streaming dependency hooks. Default methods retain no input-width state.
///
/// A callback receives a fully validated record, but enclosing validation and
/// Buffa parity may still fail afterwards. Successful callbacks are not rolled
/// back by the codec. Callers must stage mutations until the decode returns
/// `Ok`, or supply an explicit reversible rollback discipline.
pub trait DependencyVisitor {
    fn visit_formula_owner_dependency(
        &mut self,
        _reference: ReferenceRecord<'_>,
    ) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_tiled_cell_dependency(
        &mut self,
        _reference: ReferenceRecord<'_>,
    ) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_tiled_range_dependency(
        &mut self,
        _reference: ReferenceRecord<'_>,
    ) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_cell_record(&mut self, _record: CellRecordSnapshot<'_>) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_range_back_dependency(
        &mut self,
        _record: RangeBackDependencySnapshot<'_>,
    ) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_from_to_range(&mut self, _record: FromToRangeSnapshot<'_>) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_expanded_edge_component(
        &mut self,
        _component: ExpandedEdgeComponent,
    ) -> Result<(), DecodeError> {
        Ok(())
    }
}
impl DependencyVisitor for () {}

pub fn decode_calculation_engine(
    source: &[u8],
    options: DecodeOptions,
) -> Result<CalculationEngineSnapshot<'_>, DecodeError> {
    Ok(decode_calculation_engine_with_report(source, options)?.0)
}

pub fn decode_calculation_engine_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(CalculationEngineSnapshot<'_>, DecodeReport), DecodeError> {
    decode_calculation_engine_with_visitor(source, options, &mut ())
}

pub fn decode_calculation_engine_with_visitor<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    visitor: &mut dyn DependencyVisitor,
) -> Result<(CalculationEngineSnapshot<'source>, DecodeReport), DecodeError> {
    let mut budget = wire::Budget::new(source, options)?;
    let snapshot = decode_calculation_engine_in(source, &mut budget, 1, visitor)?;
    Ok((snapshot, budget.report()))
}

fn decode_calculation_engine_in<'source>(
    source: &'source [u8],
    budget: &mut wire::Budget,
    depth: u32,
    visitor: &mut dyn DependencyVisitor,
) -> Result<CalculationEngineSnapshot<'source>, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut base_date_1904 = None;
    let mut dependency_tracker = None;
    let mut named_reference_manager = None;
    let mut remote_data_store = None;
    let mut header_name_manager = None;
    let mut refs_to_dirty = None;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut base_date_1904, wire::canonical_bool(field.varint()?)?)?,
            2 => {
                let raw = field.bytes()?;
                if dependency_tracker.is_some() {
                    return Err(DecodeError::invalid());
                }
                let _ = decode_dependency_tracker_in(raw, budget, child_depth, visitor)?;
                dependency_tracker = Some(raw);
            },
            3 => {
                let raw = field.bytes()?;
                let reference = wire::decode_reference(raw, budget, child_depth)?;
                set_once(&mut named_reference_manager, (raw, reference))?;
            },
            12 => {
                let raw = field.bytes()?;
                let reference = wire::decode_reference(raw, budget, child_depth)?;
                set_once(&mut remote_data_store, (raw, reference))?;
            },
            14 => {
                let raw = field.bytes()?;
                let reference = wire::decode_reference(raw, budget, child_depth)?;
                set_once(&mut header_name_manager, (raw, reference))?;
            },
            15 => {
                let raw = field.bytes()?;
                let reference = wire::decode_reference(raw, budget, child_depth)?;
                set_once(&mut refs_to_dirty, (raw, reference))?;
            },
            _ => {},
        }
    }
    let snapshot = CalculationEngineSnapshot {
        base_date_1904,
        dependency_tracker: dependency_tracker.ok_or_else(DecodeError::invalid)?,
        named_reference_manager: named_reference_manager.map(|(_raw, reference)| reference),
        remote_data_store: remote_data_store.map(|(_raw, reference)| reference),
        header_name_manager: header_name_manager.map(|(_raw, reference)| reference),
        refs_to_dirty: refs_to_dirty.map(|(_raw, reference)| reference),
    };
    budget.message(source, depth)?;
    let view: projection::CalculationEngineArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.base_date_1904 != snapshot.base_date_1904
        || view.dependency_tracker != snapshot.dependency_tracker
        || view.named_reference_manager != named_reference_manager.map(|(raw, _reference)| raw)
        || view.remote_data_store != remote_data_store.map(|(raw, _reference)| raw)
        || view.header_name_manager != header_name_manager.map(|(raw, _reference)| raw)
        || view.refs_to_dirty != refs_to_dirty.map(|(raw, _reference)| raw)
    {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

pub fn decode_dependency_tracker(
    source: &[u8],
    options: DecodeOptions,
) -> Result<DependencyTrackerSnapshot<'_>, DecodeError> {
    Ok(decode_dependency_tracker_with_report(source, options)?.0)
}

pub fn decode_dependency_tracker_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(DependencyTrackerSnapshot<'_>, DecodeReport), DecodeError> {
    decode_dependency_tracker_with_visitor(source, options, &mut ())
}

pub fn decode_dependency_tracker_with_visitor<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    visitor: &mut dyn DependencyVisitor,
) -> Result<(DependencyTrackerSnapshot<'source>, DecodeReport), DecodeError> {
    let mut budget = wire::Budget::new(source, options)?;
    let snapshot = decode_dependency_tracker_in(source, &mut budget, 1, visitor)?;
    Ok((snapshot, budget.report()))
}

fn decode_dependency_tracker_in<'source>(
    source: &'source [u8],
    budget: &mut wire::Budget,
    depth: u32,
    visitor: &mut dyn DependencyVisitor,
) -> Result<DependencyTrackerSnapshot<'source>, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut owner_id_map = None;
    let mut number_of_formulas = None;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 | 2 | 4 => wire::scan_opaque_message(field.bytes()?, budget, child_depth)?,
            3 => {
                let raw = field.bytes()?;
                if owner_id_map.is_some() {
                    return Err(DecodeError::invalid());
                }
                wire::scan_opaque_message(raw, budget, child_depth)?;
                owner_id_map = Some(raw);
            },
            5 => set_once(&mut number_of_formulas, field.varint()?)?,
            6 => {
                let raw = field.bytes()?;
                let reference = wire::decode_reference(raw, budget, child_depth)?;
                visitor.visit_formula_owner_dependency(ReferenceRecord { raw, reference })?;
            },
            _ => {},
        }
    }
    let snapshot = DependencyTrackerSnapshot {
        owner_id_map,
        number_of_formulas,
    };
    budget.message(source, depth)?;
    let view: projection::DependencyTrackerArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.owner_id_map != snapshot.owner_id_map
        || view.number_of_formulas != snapshot.number_of_formulas
    {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

pub fn decode_formula_owner_dependencies(
    source: &[u8],
    options: DecodeOptions,
) -> Result<FormulaOwnerDependenciesSnapshot<'_>, DecodeError> {
    Ok(decode_formula_owner_dependencies_with_report(source, options)?.0)
}

pub fn decode_formula_owner_dependencies_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(FormulaOwnerDependenciesSnapshot<'_>, DecodeReport), DecodeError> {
    decode_formula_owner_dependencies_with_visitor(source, options, &mut ())
}

pub fn decode_formula_owner_dependencies_with_visitor<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    visitor: &mut dyn DependencyVisitor,
) -> Result<(FormulaOwnerDependenciesSnapshot<'source>, DecodeReport), DecodeError> {
    let mut budget = wire::Budget::new(source, options)?;
    let snapshot = decode_formula_owner_dependencies_in(source, &mut budget, 1, visitor)?;
    Ok((snapshot, budget.report()))
}

fn decode_formula_owner_dependencies_in<'source>(
    source: &'source [u8],
    budget: &mut wire::Budget,
    depth: u32,
    visitor: &mut dyn DependencyVisitor,
) -> Result<FormulaOwnerDependenciesSnapshot<'source>, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut raw: [Option<&'source [u8]>; 16] = [None; 16];
    let mut formula_owner_uid = None;
    let mut internal_formula_owner_id = None;
    let mut owner_kind = None;
    let mut formula_owner = None;
    let mut base_owner_uid = None;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        let index = usize::try_from(field.number).map_err(|_conversion| DecodeError::invalid())?;
        match field.number {
            1 => {
                let payload = field.bytes()?;
                if raw[0].is_some() {
                    return Err(DecodeError::invalid());
                }
                formula_owner_uid = Some(decode_uuid(payload, budget, child_depth)?);
                raw[0] = Some(payload);
            },
            2 => set_once(
                &mut internal_formula_owner_id,
                wire::canonical_u32(field.varint()?)?,
            )?,
            3 => set_once(&mut owner_kind, wire::canonical_u32(field.varint()?)?)?,
            4 => {
                let payload = unique_raw(&mut raw[index - 1], field.bytes()?)?;
                decode_cell_dependencies(payload, budget, child_depth, visitor)?;
            },
            5 => {
                let payload = unique_raw(&mut raw[index - 1], field.bytes()?)?;
                decode_range_dependencies(payload, budget, child_depth, visitor)?;
            },
            6..=10 | 14 | 16 => {
                let payload = unique_raw(&mut raw[index - 1], field.bytes()?)?;
                wire::scan_opaque_message(payload, budget, child_depth)?;
            },
            11 => {
                let payload = unique_raw(&mut raw[index - 1], field.bytes()?)?;
                formula_owner = Some(wire::decode_reference(payload, budget, child_depth)?);
            },
            12 => {
                let payload = unique_raw(&mut raw[index - 1], field.bytes()?)?;
                base_owner_uid = Some(decode_uuid(payload, budget, child_depth)?);
            },
            13 => {
                let payload = unique_raw(&mut raw[index - 1], field.bytes()?)?;
                decode_reference_container(payload, budget, child_depth, visitor, false)?;
            },
            15 => {
                let payload = unique_raw(&mut raw[index - 1], field.bytes()?)?;
                decode_reference_container(payload, budget, child_depth, visitor, true)?;
            },
            _ => {},
        }
    }
    let snapshot = FormulaOwnerDependenciesSnapshot {
        formula_owner_uid: formula_owner_uid.ok_or_else(DecodeError::invalid)?,
        internal_formula_owner_id: internal_formula_owner_id.ok_or_else(DecodeError::invalid)?,
        owner_kind,
        cell_dependencies: raw[3],
        range_dependencies: raw[4],
        volatile_dependencies: raw[5],
        spanning_column_dependencies: raw[6],
        spanning_row_dependencies: raw[7],
        whole_owner_dependencies: raw[8],
        cell_errors: raw[9],
        formula_owner,
        base_owner_uid,
        tiled_cell_dependencies: raw[12],
        uuid_references: raw[13],
        tiled_range_dependencies: raw[14],
        spill_range_sizes: raw[15],
    };
    budget.message(source, depth)?;
    let view: projection::FormulaOwnerDependenciesArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.formula_owner_uid != raw[0].ok_or_else(DecodeError::invalid)?
        || view.internal_formula_owner_id != snapshot.internal_formula_owner_id
        || view.owner_kind != snapshot.owner_kind
        || view.cell_dependencies != raw[3]
        || view.range_dependencies != raw[4]
        || view.volatile_dependencies != raw[5]
        || view.spanning_column_dependencies != raw[6]
        || view.spanning_row_dependencies != raw[7]
        || view.whole_owner_dependencies != raw[8]
        || view.cell_errors != raw[9]
        || view.formula_owner != raw[10]
        || view.base_owner_uid != raw[11]
        || view.tiled_cell_dependencies != raw[12]
        || view.uuid_references != raw[13]
        || view.tiled_range_dependencies != raw[14]
        || view.spill_range_sizes != raw[15]
    {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

pub fn decode_cell_record(
    source: &[u8],
    options: DecodeOptions,
) -> Result<CellRecordSnapshot<'_>, DecodeError> {
    Ok(decode_cell_record_with_report(source, options)?.0)
}

pub fn decode_cell_record_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(CellRecordSnapshot<'_>, DecodeReport), DecodeError> {
    decode_cell_record_with_visitor(source, options, &mut ())
}

pub fn decode_cell_record_with_visitor<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    visitor: &mut dyn DependencyVisitor,
) -> Result<(CellRecordSnapshot<'source>, DecodeReport), DecodeError> {
    let mut budget = wire::Budget::new(source, options)?;
    let snapshot = decode_cell_record_in(source, &mut budget, 1, visitor)?;
    Ok((snapshot, budget.report()))
}

fn decode_cell_record_in<'source>(
    source: &'source [u8],
    budget: &mut wire::Budget,
    depth: u32,
    visitor: &mut dyn DependencyVisitor,
) -> Result<CellRecordSnapshot<'source>, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut column = None;
    let mut row = None;
    let mut dirty = None;
    let mut cycle = None;
    let mut calculated = None;
    let mut expanded_edges = None;
    let mut expanded_edges_snapshot = None;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut column, wire::canonical_u32(field.varint()?)?)?,
            2 => set_once(&mut row, wire::canonical_u32(field.varint()?)?)?,
            3 => set_once(&mut dirty, field.varint()?)?,
            4 => set_once(&mut cycle, wire::canonical_bool(field.varint()?)?)?,
            5 => set_once(&mut calculated, wire::canonical_bool(field.varint()?)?)?,
            6 => {
                let raw = field.bytes()?;
                if expanded_edges.is_some() {
                    return Err(DecodeError::invalid());
                }
                expanded_edges_snapshot =
                    Some(decode_expanded_edges_in(raw, budget, child_depth, visitor)?);
                expanded_edges = Some(raw);
            },
            _ => {},
        }
    }
    let snapshot = CellRecordSnapshot {
        column: column.ok_or_else(DecodeError::invalid)?,
        row: row.ok_or_else(DecodeError::invalid)?,
        dirty_self_plus_precedents_count: dirty,
        is_in_a_cycle: cycle,
        has_calculated_precedents: calculated,
        expanded_edges,
        expanded_edges_snapshot,
    };
    budget.message(source, depth)?;
    let view: projection::CellRecordArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.column != snapshot.column
        || view.row != snapshot.row
        || view.dirty_self_plus_precedents_count != snapshot.dirty_self_plus_precedents_count
        || view.is_in_a_cycle != snapshot.is_in_a_cycle
        || view.has_calculated_precedents != snapshot.has_calculated_precedents
        || view.expanded_edges != snapshot.expanded_edges
    {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

pub fn decode_cell_record_tile(
    source: &[u8],
    options: DecodeOptions,
) -> Result<CellRecordTileSnapshot, DecodeError> {
    Ok(decode_cell_record_tile_with_report(source, options)?.0)
}

pub fn decode_cell_record_tile_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(CellRecordTileSnapshot, DecodeReport), DecodeError> {
    decode_cell_record_tile_with_visitor(source, options, &mut ())
}

pub fn decode_cell_record_tile_with_visitor(
    source: &[u8],
    options: DecodeOptions,
    visitor: &mut dyn DependencyVisitor,
) -> Result<(CellRecordTileSnapshot, DecodeReport), DecodeError> {
    let mut budget = wire::Budget::new(source, options)?;
    let snapshot = decode_cell_record_tile_in(source, &mut budget, 1, visitor)?;
    Ok((snapshot, budget.report()))
}

fn decode_cell_record_tile_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
    visitor: &mut dyn DependencyVisitor,
) -> Result<CellRecordTileSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut internal_owner_id = None;
    let mut tile_column_begin = None;
    let mut tile_row_begin = None;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(
                &mut internal_owner_id,
                wire::canonical_u32(field.varint()?)?,
            )?,
            2 => set_once(
                &mut tile_column_begin,
                wire::canonical_u32(field.varint()?)?,
            )?,
            3 => set_once(&mut tile_row_begin, wire::canonical_u32(field.varint()?)?)?,
            4 => {
                let record = decode_cell_record_in(field.bytes()?, budget, child_depth, visitor)?;
                visitor.visit_cell_record(record)?;
            },
            _ => {},
        }
    }
    let snapshot = CellRecordTileSnapshot {
        internal_owner_id: internal_owner_id.ok_or_else(DecodeError::invalid)?,
        tile_column_begin: tile_column_begin.ok_or_else(DecodeError::invalid)?,
        tile_row_begin: tile_row_begin.ok_or_else(DecodeError::invalid)?,
    };
    budget.message(source, depth)?;
    let view: projection::CellRecordTileArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.internal_owner_id != snapshot.internal_owner_id
        || view.tile_column_begin != snapshot.tile_column_begin
        || view.tile_row_begin != snapshot.tile_row_begin
    {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

pub fn decode_range_back_dependency(
    source: &[u8],
    options: DecodeOptions,
) -> Result<RangeBackDependencySnapshot<'_>, DecodeError> {
    Ok(decode_range_back_dependency_with_report(source, options)?.0)
}

pub fn decode_range_back_dependency_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(RangeBackDependencySnapshot<'_>, DecodeReport), DecodeError> {
    let mut budget = wire::Budget::new(source, options)?;
    let snapshot = decode_range_back_dependency_in(source, &mut budget, 1)?;
    Ok((snapshot, budget.report()))
}

fn decode_range_back_dependency_in<'source>(
    source: &'source [u8],
    budget: &mut wire::Budget,
    depth: u32,
) -> Result<RangeBackDependencySnapshot<'source>, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut row = None;
    let mut column = None;
    let mut range_reference = None;
    let mut internal_range_reference = None;
    let mut decoded_range_reference = None;
    let mut decoded_internal_range_reference = None;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut row, wire::canonical_u32(field.varint()?)?)?,
            2 => set_once(&mut column, wire::canonical_u32(field.varint()?)?)?,
            3 => {
                let raw = field.bytes()?;
                if range_reference.is_some() {
                    return Err(DecodeError::invalid());
                }
                decoded_range_reference =
                    Some(decode_range_reference_in(raw, budget, child_depth)?);
                range_reference = Some(raw);
            },
            4 => {
                let raw = field.bytes()?;
                if internal_range_reference.is_some() {
                    return Err(DecodeError::invalid());
                }
                decoded_internal_range_reference = Some(decode_internal_range_reference_in(
                    raw,
                    budget,
                    child_depth,
                )?);
                internal_range_reference = Some(raw);
            },
            _ => {},
        }
    }
    let snapshot = RangeBackDependencySnapshot {
        cell_coord_row: row.ok_or_else(DecodeError::invalid)?,
        cell_coord_column: column.ok_or_else(DecodeError::invalid)?,
        range_reference,
        internal_range_reference,
        decoded_range_reference,
        decoded_internal_range_reference,
    };
    budget.message(source, depth)?;
    let view: projection::RangeBackDependencyArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.cell_coord_row != snapshot.cell_coord_row
        || view.cell_coord_column != snapshot.cell_coord_column
        || view.range_reference != snapshot.range_reference
        || view.internal_range_reference != snapshot.internal_range_reference
    {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

pub fn decode_range_precedents_tile(
    source: &[u8],
    options: DecodeOptions,
) -> Result<RangePrecedentsTileSnapshot, DecodeError> {
    Ok(decode_range_precedents_tile_with_report(source, options)?.0)
}

pub fn decode_range_precedents_tile_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(RangePrecedentsTileSnapshot, DecodeReport), DecodeError> {
    decode_range_precedents_tile_with_visitor(source, options, &mut ())
}

pub fn decode_range_precedents_tile_with_visitor(
    source: &[u8],
    options: DecodeOptions,
    visitor: &mut dyn DependencyVisitor,
) -> Result<(RangePrecedentsTileSnapshot, DecodeReport), DecodeError> {
    let mut budget = wire::Budget::new(source, options)?;
    let snapshot = decode_range_precedents_tile_in(source, &mut budget, 1, visitor)?;
    Ok((snapshot, budget.report()))
}

fn decode_range_precedents_tile_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
    visitor: &mut dyn DependencyVisitor,
) -> Result<RangePrecedentsTileSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut to_owner_id = None;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut to_owner_id, wire::canonical_u32(field.varint()?)?)?,
            2 => visitor.visit_from_to_range(decode_from_to_range_in(
                field.bytes()?,
                budget,
                child_depth,
            )?)?,
            _ => {},
        }
    }
    let snapshot = RangePrecedentsTileSnapshot {
        to_owner_id: to_owner_id.ok_or_else(DecodeError::invalid)?,
    };
    budget.message(source, depth)?;
    let view: projection::RangePrecedentsTileArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.to_owner_id != snapshot.to_owner_id {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

pub fn decode_from_to_range(
    source: &[u8],
    options: DecodeOptions,
) -> Result<FromToRangeSnapshot<'_>, DecodeError> {
    Ok(decode_from_to_range_with_report(source, options)?.0)
}

pub fn decode_from_to_range_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(FromToRangeSnapshot<'_>, DecodeReport), DecodeError> {
    let mut budget = wire::Budget::new(source, options)?;
    let snapshot = decode_from_to_range_in(source, &mut budget, 1)?;
    Ok((snapshot, budget.report()))
}

fn decode_from_to_range_in<'source>(
    source: &'source [u8],
    budget: &mut wire::Budget,
    depth: u32,
) -> Result<FromToRangeSnapshot<'source>, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut from_coord = None;
    let mut refers_to_rect = None;
    let mut decoded_from_coord = None;
    let mut decoded_refers_to_rect = None;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => {
                let raw = field.bytes()?;
                if from_coord.is_some() {
                    return Err(DecodeError::invalid());
                }
                decoded_from_coord = Some(decode_cell_coordinate_in(raw, budget, child_depth)?);
                from_coord = Some(raw);
            },
            2 => {
                let raw = field.bytes()?;
                if refers_to_rect.is_some() {
                    return Err(DecodeError::invalid());
                }
                decoded_refers_to_rect = Some(decode_cell_rect_in(raw, budget, child_depth)?);
                refers_to_rect = Some(raw);
            },
            _ => {},
        }
    }
    let snapshot = FromToRangeSnapshot {
        from_coord: from_coord.ok_or_else(DecodeError::invalid)?,
        refers_to_rect: refers_to_rect.ok_or_else(DecodeError::invalid)?,
        decoded_from_coord: decoded_from_coord.ok_or_else(DecodeError::invalid)?,
        decoded_refers_to_rect: decoded_refers_to_rect.ok_or_else(DecodeError::invalid)?,
    };
    budget.message(source, depth)?;
    let view: projection::FromToRangeArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.from_coord != snapshot.from_coord || view.refers_to_rect != snapshot.refers_to_rect {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

pub fn decode_expanded_edges(
    source: &[u8],
    options: DecodeOptions,
) -> Result<ExpandedEdgesSnapshot, DecodeError> {
    Ok(decode_expanded_edges_with_report(source, options)?.0)
}

pub fn decode_expanded_edges_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(ExpandedEdgesSnapshot, DecodeReport), DecodeError> {
    decode_expanded_edges_with_visitor(source, options, &mut ())
}

pub fn decode_expanded_edges_with_visitor(
    source: &[u8],
    options: DecodeOptions,
    visitor: &mut dyn DependencyVisitor,
) -> Result<(ExpandedEdgesSnapshot, DecodeReport), DecodeError> {
    let mut budget = wire::Budget::new(source, options)?;
    let snapshot = decode_expanded_edges_in(source, &mut budget, 1, visitor)?;
    Ok((snapshot, budget.report()))
}

fn decode_expanded_edges_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
    visitor: &mut dyn DependencyVisitor,
) -> Result<ExpandedEdgesSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let mut counts = [0usize; 5];
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        if !(1..=5).contains(&field.number) {
            continue;
        }
        let index = usize::try_from(field.number - 1).map_err(|_error| DecodeError::invalid())?;
        let kind = match field.number {
            1 => ExpandedEdgeKind::LocalRow,
            2 => ExpandedEdgeKind::LocalColumn,
            3 => ExpandedEdgeKind::ExternalRow,
            4 => ExpandedEdgeKind::ExternalColumn,
            5 => ExpandedEdgeKind::InternalOwner,
            _ => return Err(DecodeError::invalid()),
        };
        if let Ok(value) = field.varint() {
            visit_edge_component(kind, value, &mut counts[index], visitor)?;
        } else {
            let mut packed = field.bytes()?;
            while !packed.is_empty() {
                let value = wire::take_canonical_varint(&mut packed)?;
                visit_edge_component(kind, value, &mut counts[index], visitor)?;
            }
        }
    }
    if counts[0] != counts[1] || counts[2] != counts[3] || counts[2] != counts[4] {
        return Err(DecodeError::invalid());
    }
    let snapshot = ExpandedEdgesSnapshot {
        local: counts[0],
        external: counts[2],
    };
    budget.message(source, depth)?;
    let _view: projection::ExpandedEdgesArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    Ok(snapshot)
}

fn visit_edge_component(
    kind: ExpandedEdgeKind,
    value: u64,
    count: &mut usize,
    visitor: &mut dyn DependencyVisitor,
) -> Result<(), DecodeError> {
    let index = *count;
    *count = count.checked_add(1).ok_or_else(DecodeError::invalid)?;
    visitor.visit_expanded_edge_component(ExpandedEdgeComponent {
        kind,
        index,
        value: wire::canonical_u32(value)?,
    })
}

pub fn decode_cell_coordinate(
    source: &[u8],
    options: DecodeOptions,
) -> Result<CellCoordinateSnapshot, DecodeError> {
    Ok(decode_cell_coordinate_with_report(source, options)?.0)
}

pub fn decode_cell_coordinate_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(CellCoordinateSnapshot, DecodeReport), DecodeError> {
    let mut budget = wire::Budget::new(source, options)?;
    let snapshot = decode_cell_coordinate_in(source, &mut budget, 1)?;
    Ok((snapshot, budget.report()))
}

fn decode_cell_coordinate_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
) -> Result<CellCoordinateSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let mut packed_data = None;
    let mut column = None;
    let mut row = None;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut packed_data, field.fixed32()?)?,
            2 => set_once(&mut column, wire::canonical_u32(field.varint()?)?)?,
            3 => set_once(&mut row, wire::canonical_u32(field.varint()?)?)?,
            _ => {},
        }
    }
    let snapshot = CellCoordinateSnapshot {
        packed_data,
        column,
        row,
    };
    budget.message(source, depth)?;
    let view: projection::CellCoordinateArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.packed_data != snapshot.packed_data
        || view.column != snapshot.column
        || view.row != snapshot.row
    {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

fn decode_column_row_size_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
) -> Result<ColumnRowSizeSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let mut num_columns = None;
    let mut num_rows = None;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut num_columns, wire::canonical_u32(field.varint()?)?)?,
            2 => set_once(&mut num_rows, wire::canonical_u32(field.varint()?)?)?,
            _ => {},
        }
    }
    let snapshot = ColumnRowSizeSnapshot {
        num_columns,
        num_rows,
    };
    budget.message(source, depth)?;
    let view: projection::ColumnRowSizeLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.num_columns != snapshot.num_columns || view.num_rows != snapshot.num_rows {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

fn decode_cell_rect_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
) -> Result<CellRectSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut origin_raw = None;
    let mut size_raw = None;
    let mut origin = None;
    let mut size = None;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => {
                let raw = field.bytes()?;
                set_once(
                    &mut origin,
                    decode_cell_coordinate_in(raw, budget, child_depth)?,
                )?;
                set_once(&mut origin_raw, raw)?;
            },
            2 => {
                let raw = field.bytes()?;
                set_once(
                    &mut size,
                    decode_column_row_size_in(raw, budget, child_depth)?,
                )?;
                set_once(&mut size_raw, raw)?;
            },
            _ => {},
        }
    }
    let snapshot = CellRectSnapshot {
        origin: origin.ok_or_else(DecodeError::invalid)?,
        size: size.ok_or_else(DecodeError::invalid)?,
    };
    budget.message(source, depth)?;
    let view: projection::CellRectArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.origin != origin_raw.ok_or_else(DecodeError::invalid)?
        || view.size != size_raw.ok_or_else(DecodeError::invalid)?
    {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

fn decode_cfuuid_in<'source>(
    source: &'source [u8],
    budget: &mut wire::Budget,
    depth: u32,
) -> Result<CfuuidSnapshot<'source>, DecodeError> {
    budget.message(source, depth)?;
    let mut uuid_bytes = None;
    let mut words = [None; 4];
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut uuid_bytes, field.bytes()?)?,
            2..=5 => {
                let index =
                    usize::try_from(field.number - 2).map_err(|_error| DecodeError::invalid())?;
                set_once(&mut words[index], wire::canonical_u32(field.varint()?)?)?;
            },
            _ => {},
        }
    }
    let snapshot = CfuuidSnapshot { uuid_bytes, words };
    budget.message(source, depth)?;
    let view: projection::CFUUIDArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.uuid_bytes != snapshot.uuid_bytes
        || view.uuid_w0 != snapshot.words[0]
        || view.uuid_w1 != snapshot.words[1]
        || view.uuid_w2 != snapshot.words[2]
        || view.uuid_w3 != snapshot.words[3]
    {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

pub fn decode_range_reference(
    source: &[u8],
    options: DecodeOptions,
) -> Result<RangeReferenceSnapshot<'_>, DecodeError> {
    Ok(decode_range_reference_with_report(source, options)?.0)
}

pub fn decode_range_reference_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(RangeReferenceSnapshot<'_>, DecodeReport), DecodeError> {
    let mut budget = wire::Budget::new(source, options)?;
    let snapshot = decode_range_reference_in(source, &mut budget, 1)?;
    Ok((snapshot, budget.report()))
}

fn decode_range_reference_in<'source>(
    source: &'source [u8],
    budget: &mut wire::Budget,
    depth: u32,
) -> Result<RangeReferenceSnapshot<'source>, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut table_id_raw = None;
    let mut table_id = None;
    let mut values = [None; 4];
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => {
                let raw = field.bytes()?;
                set_once(&mut table_id, decode_cfuuid_in(raw, budget, child_depth)?)?;
                set_once(&mut table_id_raw, raw)?;
            },
            2..=5 => {
                let index =
                    usize::try_from(field.number - 2).map_err(|_error| DecodeError::invalid())?;
                set_once(&mut values[index], wire::canonical_u32(field.varint()?)?)?;
            },
            _ => {},
        }
    }
    let range = RangeCoordinateSnapshot {
        top_left_column: values[0].ok_or_else(DecodeError::invalid)?,
        top_left_row: values[1].ok_or_else(DecodeError::invalid)?,
        bottom_right_column: values[2].ok_or_else(DecodeError::invalid)?,
        bottom_right_row: values[3].ok_or_else(DecodeError::invalid)?,
    };
    let snapshot = RangeReferenceSnapshot {
        table_id: table_id.ok_or_else(DecodeError::invalid)?,
        range,
    };
    budget.message(source, depth)?;
    let view: projection::RangeReferenceArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.table_id != table_id_raw.ok_or_else(DecodeError::invalid)?
        || view.top_left_column != range.top_left_column
        || view.top_left_row != range.top_left_row
        || view.bottom_right_column != range.bottom_right_column
        || view.bottom_right_row != range.bottom_right_row
    {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

fn decode_range_coordinate_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
) -> Result<RangeCoordinateSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let mut values = [None; 4];
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        if (1..=4).contains(&field.number) {
            let index =
                usize::try_from(field.number - 1).map_err(|_error| DecodeError::invalid())?;
            set_once(&mut values[index], wire::canonical_u32(field.varint()?)?)?;
        }
    }
    let snapshot = RangeCoordinateSnapshot {
        top_left_column: values[0].ok_or_else(DecodeError::invalid)?,
        top_left_row: values[1].ok_or_else(DecodeError::invalid)?,
        bottom_right_column: values[2].ok_or_else(DecodeError::invalid)?,
        bottom_right_row: values[3].ok_or_else(DecodeError::invalid)?,
    };
    budget.message(source, depth)?;
    let view: projection::RangeCoordinateArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.top_left_column != snapshot.top_left_column
        || view.top_left_row != snapshot.top_left_row
        || view.bottom_right_column != snapshot.bottom_right_column
        || view.bottom_right_row != snapshot.bottom_right_row
    {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

pub fn decode_internal_range_reference(
    source: &[u8],
    options: DecodeOptions,
) -> Result<InternalRangeReferenceSnapshot, DecodeError> {
    Ok(decode_internal_range_reference_with_report(source, options)?.0)
}

pub fn decode_internal_range_reference_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(InternalRangeReferenceSnapshot, DecodeReport), DecodeError> {
    let mut budget = wire::Budget::new(source, options)?;
    let snapshot = decode_internal_range_reference_in(source, &mut budget, 1)?;
    Ok((snapshot, budget.report()))
}

fn decode_internal_range_reference_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
) -> Result<InternalRangeReferenceSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut owner_id = None;
    let mut range_raw = None;
    let mut range = None;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut owner_id, wire::canonical_u32(field.varint()?)?)?,
            2 => {
                let raw = field.bytes()?;
                set_once(
                    &mut range,
                    decode_range_coordinate_in(raw, budget, child_depth)?,
                )?;
                set_once(&mut range_raw, raw)?;
            },
            _ => {},
        }
    }
    let snapshot = InternalRangeReferenceSnapshot {
        owner_id: owner_id.ok_or_else(DecodeError::invalid)?,
        range: range.ok_or_else(DecodeError::invalid)?,
    };
    budget.message(source, depth)?;
    let view: projection::InternalRangeReferenceArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if view.owner_id != snapshot.owner_id
        || view.range != range_raw.ok_or_else(DecodeError::invalid)?
    {
        return Err(DecodeError::invalid());
    }
    Ok(snapshot)
}

fn decode_uuid(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
) -> Result<UuidSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let mut lower = None;
    let mut upper = None;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut lower, field.varint()?)?,
            2 => set_once(&mut upper, field.varint()?)?,
            _ => {},
        }
    }
    Ok(UuidSnapshot {
        lower: lower.ok_or_else(DecodeError::invalid)?,
        upper: upper.ok_or_else(DecodeError::invalid)?,
    })
}

fn decode_cell_dependencies(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
    visitor: &mut dyn DependencyVisitor,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        if field.number == 1 {
            let record = decode_cell_record_in(field.bytes()?, budget, child_depth, visitor)?;
            visitor.visit_cell_record(record)?;
        }
    }
    Ok(())
}

fn decode_range_dependencies(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
    visitor: &mut dyn DependencyVisitor,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        if field.number == 2 {
            visitor.visit_range_back_dependency(decode_range_back_dependency_in(
                field.bytes()?,
                budget,
                child_depth,
            )?)?;
        }
    }
    Ok(())
}

fn decode_reference_container(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
    visitor: &mut dyn DependencyVisitor,
    range: bool,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        if field.number == 1 {
            let raw = field.bytes()?;
            let reference = wire::decode_reference(raw, budget, child_depth)?;
            let record = ReferenceRecord { raw, reference };
            if range {
                visitor.visit_tiled_range_dependency(record)?;
            } else {
                visitor.visit_tiled_cell_dependency(record)?;
            }
        }
    }
    Ok(())
}

fn unique_raw<'source>(
    slot: &mut Option<&'source [u8]>,
    value: &'source [u8],
) -> Result<&'source [u8], DecodeError> {
    if slot.is_some() {
        return Err(DecodeError::invalid());
    }
    *slot = Some(value);
    Ok(value)
}

fn set_once<T>(slot: &mut Option<T>, value: T) -> Result<(), DecodeError> {
    if slot.is_some() {
        return Err(DecodeError::invalid());
    }
    *slot = Some(value);
    Ok(())
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "Focused canonical dependency fixtures require exact construction and failures."
)]
mod tests {
    use super::*;

    fn varint(output: &mut Vec<u8>, mut value: u64) {
        loop {
            let mut byte = u8::try_from(value & 0x7f).unwrap();
            value >>= 7;
            if value != 0 {
                byte |= 0x80;
            }
            output.push(byte);
            if value == 0 {
                return;
            }
        }
    }
    fn key(output: &mut Vec<u8>, field: u32, wire: u8) {
        varint(output, (u64::from(field) << 3) | u64::from(wire));
    }
    fn v(output: &mut Vec<u8>, field: u32, value: u64) {
        key(output, field, 0);
        varint(output, value);
    }
    fn b(output: &mut Vec<u8>, field: u32, value: &[u8]) {
        key(output, field, 2);
        varint(output, u64::try_from(value.len()).unwrap());
        output.extend_from_slice(value);
    }
    fn fixed32(output: &mut Vec<u8>, field: u32, value: u32) {
        key(output, field, 5);
        output.extend_from_slice(&value.to_le_bytes());
    }
    fn packed(output: &mut Vec<u8>, field: u32, values: &[u32]) {
        let mut payload = Vec::new();
        for value in values {
            varint(&mut payload, u64::from(*value));
        }
        b(output, field, &payload);
    }
    fn reference(id: u64) -> Vec<u8> {
        let mut out = Vec::new();
        v(&mut out, 1, id);
        out
    }
    fn external_reference(id: u64) -> Vec<u8> {
        let mut out = reference(id);
        v(&mut out, 3, 1);
        out
    }
    fn uuid(seed: u64) -> Vec<u8> {
        let mut out = Vec::new();
        v(&mut out, 1, seed);
        v(&mut out, 2, seed + 1);
        out
    }
    fn cell_record(column: u32, row: u32) -> Vec<u8> {
        let mut out = Vec::new();
        v(&mut out, 1, u64::from(column));
        v(&mut out, 2, u64::from(row));
        out
    }
    fn cell_record_tile(records: usize) -> Vec<u8> {
        let mut out = Vec::new();
        v(&mut out, 1, 1);
        v(&mut out, 2, 0);
        v(&mut out, 3, 0);
        for index in 0..records {
            b(&mut out, 4, &cell_record(u32::try_from(index).unwrap(), 0));
        }
        out
    }
    fn tracker(references: usize) -> Vec<u8> {
        let mut out = Vec::new();
        v(&mut out, 5, u64::try_from(references).unwrap());
        for id in 1..=u64::try_from(references).unwrap() {
            b(&mut out, 6, &reference(id));
        }
        out
    }
    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(
            source.len().max(1),
            2_000_000,
            source.len().saturating_mul(24).max(1),
            64,
            20_000,
            0,
        )
    }

    #[derive(Default)]
    struct Counts {
        owners: usize,
        tiled_cells: usize,
        tiled_ranges: usize,
        cells: usize,
        backs: usize,
        ranges: usize,
        edge_components: usize,
    }
    impl DependencyVisitor for Counts {
        fn visit_formula_owner_dependency(
            &mut self,
            record: ReferenceRecord<'_>,
        ) -> Result<(), DecodeError> {
            self.owners += 1;
            assert_ne!(record.reference().identifier(), 0);
            Ok(())
        }
        fn visit_tiled_cell_dependency(
            &mut self,
            record: ReferenceRecord<'_>,
        ) -> Result<(), DecodeError> {
            self.tiled_cells += 1;
            assert_ne!(record.reference().identifier(), 0);
            Ok(())
        }
        fn visit_tiled_range_dependency(
            &mut self,
            record: ReferenceRecord<'_>,
        ) -> Result<(), DecodeError> {
            self.tiled_ranges += 1;
            assert_ne!(record.reference().identifier(), 0);
            Ok(())
        }
        fn visit_cell_record(&mut self, record: CellRecordSnapshot<'_>) -> Result<(), DecodeError> {
            self.cells += 1;
            assert_eq!(record.row(), 0);
            Ok(())
        }
        fn visit_range_back_dependency(
            &mut self,
            record: RangeBackDependencySnapshot<'_>,
        ) -> Result<(), DecodeError> {
            self.backs += 1;
            assert_eq!(record.cell_coord_row(), 1);
            Ok(())
        }
        fn visit_from_to_range(
            &mut self,
            record: FromToRangeSnapshot<'_>,
        ) -> Result<(), DecodeError> {
            self.ranges += 1;
            assert!(record.from_coord().is_empty());
            Ok(())
        }
        fn visit_expanded_edge_component(
            &mut self,
            _component: ExpandedEdgeComponent,
        ) -> Result<(), DecodeError> {
            self.edge_components += 1;
            Ok(())
        }
    }

    #[test]
    fn calculation_tracker_streams_references_and_crosschecks_buffa() {
        let tracker = tracker(3);
        let mut engine = Vec::new();
        v(&mut engine, 1, 1);
        b(&mut engine, 2, &tracker);
        b(&mut engine, 3, &reference(30));
        b(&mut engine, 12, &reference(120));
        b(&mut engine, 14, &reference(140));
        b(&mut engine, 15, &reference(150));
        let mut counts = Counts::default();
        let (snapshot, report) =
            decode_calculation_engine_with_visitor(&engine, options(&engine), &mut counts).unwrap();
        assert_eq!(snapshot.base_date_1904(), Some(true));
        assert_eq!(snapshot.named_reference_manager().unwrap().identifier(), 30);
        assert_eq!(snapshot.remote_data_store().unwrap().identifier(), 120);
        assert_eq!(snapshot.header_name_manager().unwrap().identifier(), 140);
        assert_eq!(snapshot.refs_to_dirty().unwrap().identifier(), 150);
        assert_eq!(counts.owners, 3);
        assert_eq!(report.references(), 7);
        assert_eq!(report.source_bytes(), engine.len());
        assert!(report.work_bytes() > engine.len() * 2);
    }

    #[test]
    fn formula_owner_streams_inline_and_tiled_dependency_envelopes() {
        let record = cell_record(2, 0);
        let mut cells = Vec::new();
        b(&mut cells, 1, &record);
        let mut back = Vec::new();
        v(&mut back, 1, 1);
        v(&mut back, 2, 2);
        let mut backs = Vec::new();
        b(&mut backs, 2, &back);
        let mut tiled_cells = Vec::new();
        b(&mut tiled_cells, 1, &reference(30));
        let mut tiled_ranges = Vec::new();
        b(&mut tiled_ranges, 1, &reference(31));
        let mut owner = Vec::new();
        b(&mut owner, 1, &uuid(4));
        v(&mut owner, 2, 9);
        v(&mut owner, 3, 2);
        b(&mut owner, 4, &cells);
        b(&mut owner, 5, &backs);
        b(&mut owner, 11, &reference(8));
        b(&mut owner, 12, &uuid(10));
        b(&mut owner, 13, &tiled_cells);
        b(&mut owner, 15, &tiled_ranges);
        let mut counts = Counts::default();
        let (snapshot, report) =
            decode_formula_owner_dependencies_with_visitor(&owner, options(&owner), &mut counts)
                .unwrap();
        assert_eq!(snapshot.internal_formula_owner_id(), 9);
        assert_eq!(snapshot.owner_kind(), Some(2));
        assert_eq!(
            (
                counts.cells,
                counts.backs,
                counts.tiled_cells,
                counts.tiled_ranges
            ),
            (1, 1, 1, 1)
        );
        assert_eq!(report.references(), 3);
        assert_eq!(report.max_depth(), 3);
    }

    #[test]
    fn standalone_cache_roots_preserve_presence_and_stream_records() {
        let record = cell_record(4, 5);
        assert_eq!(
            decode_cell_record(&record, options(&record))
                .unwrap()
                .column(),
            4
        );
        let tile = cell_record_tile(2);
        let mut counts = Counts::default();
        assert_eq!(
            decode_cell_record_tile_with_visitor(&tile, options(&tile), &mut counts)
                .unwrap()
                .0
                .internal_owner_id(),
            1
        );
        assert_eq!(counts.cells, 2);
        let mut back = Vec::new();
        v(&mut back, 1, 2);
        v(&mut back, 2, 3);
        let mut table_range = Vec::new();
        b(&mut table_range, 1, &[]);
        v(&mut table_range, 2, 1);
        v(&mut table_range, 3, 2);
        v(&mut table_range, 4, 3);
        v(&mut table_range, 5, 4);
        b(&mut back, 3, &table_range);
        assert_eq!(
            decode_range_back_dependency(&back, options(&back))
                .unwrap()
                .range_reference(),
            Some(table_range.as_slice())
        );
        let mut rect = Vec::new();
        b(&mut rect, 1, &[]);
        b(&mut rect, 2, &[]);
        let mut pair = Vec::new();
        b(&mut pair, 1, &[]);
        b(&mut pair, 2, &rect);
        let mut range_tile = Vec::new();
        v(&mut range_tile, 1, 7);
        b(&mut range_tile, 2, &pair);
        decode_range_precedents_tile_with_visitor(&range_tile, options(&range_tile), &mut counts)
            .unwrap();
        assert_eq!(counts.ranges, 1);
        assert!(decode_from_to_range(&pair, options(&pair)).is_ok());
    }

    #[test]
    fn expanded_edges_and_typed_ranges_close_formula_cache_routes() {
        let mut edges = Vec::new();
        packed(&mut edges, 1, &[2, 4]);
        v(&mut edges, 2, 3);
        v(&mut edges, 2, 5);
        packed(&mut edges, 3, &[6]);
        packed(&mut edges, 4, &[7]);
        packed(&mut edges, 5, &[8]);
        let mut counts = Counts::default();
        let (snapshot, report) =
            decode_expanded_edges_with_visitor(&edges, options(&edges), &mut counts).unwrap();
        assert_eq!((snapshot.local(), snapshot.external()), (2, 1));
        assert_eq!(counts.edge_components, 7);
        assert!(report.work_bytes() >= edges.len() * 2);

        let mut coordinate = Vec::new();
        fixed32(&mut coordinate, 1, 0x1234_5678);
        v(&mut coordinate, 2, 9);
        v(&mut coordinate, 3, 10);
        let decoded = decode_cell_coordinate(&coordinate, options(&coordinate)).unwrap();
        assert_eq!(decoded.packed_data(), Some(0x1234_5678));
        assert_eq!((decoded.column(), decoded.row()), (Some(9), Some(10)));

        let mut range = Vec::new();
        v(&mut range, 1, 1);
        v(&mut range, 2, 2);
        v(&mut range, 3, 3);
        v(&mut range, 4, 4);
        let mut internal = Vec::new();
        v(&mut internal, 1, 11);
        b(&mut internal, 2, &range);
        let decoded = decode_internal_range_reference(&internal, options(&internal)).unwrap();
        assert_eq!(decoded.owner_id(), 11);
        assert_eq!(decoded.range().bottom_right_row(), 4);

        let mut mismatched = edges.clone();
        v(&mut mismatched, 3, 99);
        assert!(decode_expanded_edges(&mismatched, options(&mismatched)).is_err());
    }

    #[test]
    fn canonical_required_duplicate_bool_and_external_failures_are_closed() {
        let missing = [0x08, 0x01];
        assert!(decode_cell_record(&missing, options(&missing)).is_err());
        let mut duplicate = cell_record(1, 2);
        v(&mut duplicate, 1, 3);
        assert!(decode_cell_record(&duplicate, options(&duplicate)).is_err());
        let mut bool_value = cell_record(1, 2);
        v(&mut bool_value, 4, 2);
        assert!(decode_cell_record(&bool_value, options(&bool_value)).is_err());
        let overlong = [0x88, 0x00, 0x01];
        assert!(decode_cell_record(&overlong, options(&overlong)).is_err());
        let mut bad_tracker = Vec::new();
        b(&mut bad_tracker, 6, &external_reference(2));
        assert!(decode_dependency_tracker(&bad_tracker, options(&bad_tracker)).is_err());

        let dependency_tracker = tracker(0);
        let mut external_header_manager = Vec::new();
        b(&mut external_header_manager, 2, &dependency_tracker);
        b(&mut external_header_manager, 14, &external_reference(14));
        assert!(
            decode_calculation_engine(&external_header_manager, options(&external_header_manager))
                .is_err()
        );
        let mut duplicate_header_manager = Vec::new();
        b(&mut duplicate_header_manager, 2, &dependency_tracker);
        b(&mut duplicate_header_manager, 14, &reference(14));
        b(&mut duplicate_header_manager, 14, &reference(15));
        assert!(
            decode_calculation_engine(
                &duplicate_header_manager,
                options(&duplicate_header_manager)
            )
            .is_err()
        );
        let mut wrong_wire_header_manager = Vec::new();
        b(&mut wrong_wire_header_manager, 2, &dependency_tracker);
        v(&mut wrong_wire_header_manager, 14, 14);
        assert!(
            decode_calculation_engine(
                &wrong_wire_header_manager,
                options(&wrong_wire_header_manager)
            )
            .is_err()
        );
    }

    #[test]
    fn exact_limits_are_inclusive_and_max_minus_one_is_typed() {
        let source = cell_record_tile(4);
        let (_, report) = decode_cell_record_tile_with_report(&source, options(&source)).unwrap();
        let exact = DecodeOptions::new(
            source.len(),
            report.fields(),
            report.work_bytes(),
            report.max_depth(),
            report.references(),
            0,
        );
        assert!(decode_cell_record_tile(&source, exact).is_ok());
        let fields = decode_cell_record_tile(
            &source,
            DecodeOptions::new(source.len(), report.fields() - 1, usize::MAX, 64, 0, 0),
        )
        .unwrap_err();
        assert!(matches!(
            fields.resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));
        let work = decode_cell_record_tile(
            &source,
            DecodeOptions::new(source.len(), usize::MAX, report.work_bytes() - 1, 64, 0, 0),
        )
        .unwrap_err();
        assert!(matches!(
            work.resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));
        let nesting = decode_cell_record_tile(
            &source,
            DecodeOptions::new(source.len(), usize::MAX, usize::MAX, 1, 0, 0),
        )
        .unwrap_err();
        assert!(matches!(
            nesting.resource_limit(),
            Some(DecodeLimit::Nesting { .. })
        ));
    }

    #[test]
    fn wide_4096_to_8192_dependency_routes_scale_linearly_and_preempt() {
        let small_source = cell_record_tile(4096);
        let large_source = cell_record_tile(8192);
        let (_, small) =
            decode_cell_record_tile_with_report(&small_source, options(&small_source)).unwrap();
        let (_, large) =
            decode_cell_record_tile_with_report(&large_source, options(&large_source)).unwrap();
        assert_eq!(large.fields() - 3, 2 * (small.fields() - 3));
        assert!(large.work_bytes() <= small.work_bytes() * 23 / 10 + 32);
        let field_error = decode_cell_record_tile(
            &large_source,
            DecodeOptions::new(large_source.len(), large.fields() - 1, usize::MAX, 64, 0, 0),
        )
        .unwrap_err();
        assert!(matches!(
            field_error.resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));
        let references = tracker(8192);
        let reference_error = decode_dependency_tracker(
            &references,
            DecodeOptions::new(references.len(), usize::MAX, usize::MAX, 64, 8191, 0),
        )
        .unwrap_err();
        assert!(matches!(
            reference_error.resource_limit(),
            Some(DecodeLimit::References {
                observed: 8192,
                maximum: 8191
            })
        ));
    }
}
