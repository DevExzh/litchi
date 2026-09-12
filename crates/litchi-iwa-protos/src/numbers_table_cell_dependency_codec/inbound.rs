//! Typed borrowed inbound facts for the dependency families that the compact
//! owner projection cannot safely leave opaque.
//!
//! The routines in this module deliberately decode only the native dependency
//! shapes needed by owner-removal proofs. They retain no generated values and
//! allocate no collection for repeated records: every validated owner ID or
//! UUID is sent to the caller's visitor while the shared strict `wire::Budget`
//! accounts for the complete nested traversal. The enclosing owner decoder
//! still runs its existing generated Buffa lazy projection for the supported
//! root and scalar fields; the dependency families here are represented as
//! opaque bytes in that projection, so this bounded strict pass supplies the
//! missing facts without materializing generated repeated arrays.

use super::{
    DecodeError, DependencyVisitor, UuidSnapshot, decode_cell_coordinate_in,
    decode_range_coordinate_in, decode_uuid, wire,
};

/// The native expanded dependency family that carries an internal owner ID.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FormulaOwnerInternalDependencyKind {
    /// `VolatileDependenciesExpandedArchive.volatile_geometry_cell_refs`.
    VolatileGeometry,
    /// `FormulaOwnerDependenciesArchive.spanning_column_dependencies`.
    SpanningColumn,
    /// `FormulaOwnerDependenciesArchive.spanning_row_dependencies`.
    SpanningRow,
    /// `WholeOwnerDependenciesExpandedArchive.dependent_cells`.
    WholeOwner,
}

impl FormulaOwnerInternalDependencyKind {
    /// Whether this fact came from a spanning dependency envelope.
    #[must_use]
    pub const fn is_spanning(self) -> bool {
        matches!(self, Self::SpanningColumn | Self::SpanningRow)
    }
}

/// One exact internal-owner inbound dependency from an expanded owner
/// envelope. The coordinate or extent that accompanies the ID is validated by
/// the decoder but intentionally not retained because owner-removal proofs
/// compare only the target internal owner identity.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct FormulaOwnerInternalDependencyFact {
    kind: FormulaOwnerInternalDependencyKind,
    owner_id: u32,
}

impl FormulaOwnerInternalDependencyFact {
    pub(super) const fn new(kind: FormulaOwnerInternalDependencyKind, owner_id: u32) -> Self {
        Self { kind, owner_id }
    }

    /// The dependency family that supplied this owner ID.
    #[must_use]
    pub const fn kind(self) -> FormulaOwnerInternalDependencyKind {
        self.kind
    }

    /// The exact internal formula-owner ID referenced by the dependency.
    #[must_use]
    pub const fn owner_id(self) -> u32 {
        self.owner_id
    }
}

/// The UUID-reference family that supplies an inbound owner UUID.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FormulaOwnerUuidDependencyKind {
    /// `UuidReferencesArchive.table_refs`.
    TableReference,
    /// `UuidReferencesArchive.table_uuid_refs`, including each nested
    /// `UuidRef` tuple.
    TableUuidReference,
}

/// One exact owner-UUID inbound dependency from `UuidReferencesArchive`.
///
/// `referenced_uuid` is populated for a nested `TableWithUuidRef.UuidRef`
/// tuple. A table-level record with no nested UUIDs still emits a fact with
/// `None`, because the native host treats the table owner entry itself as an
/// inbound dependency.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct FormulaOwnerUuidDependencyFact {
    kind: FormulaOwnerUuidDependencyKind,
    owner_uuid: UuidSnapshot,
    referenced_uuid: Option<UuidSnapshot>,
}

impl FormulaOwnerUuidDependencyFact {
    pub(super) const fn table_reference(owner_uuid: UuidSnapshot) -> Self {
        Self {
            kind: FormulaOwnerUuidDependencyKind::TableReference,
            owner_uuid,
            referenced_uuid: None,
        }
    }

    pub(super) const fn table_uuid_reference(
        owner_uuid: UuidSnapshot,
        referenced_uuid: Option<UuidSnapshot>,
    ) -> Self {
        Self {
            kind: FormulaOwnerUuidDependencyKind::TableUuidReference,
            owner_uuid,
            referenced_uuid,
        }
    }

    /// The UUID-reference family that supplied this fact.
    #[must_use]
    pub const fn kind(self) -> FormulaOwnerUuidDependencyKind {
        self.kind
    }

    /// The owner UUID targeted by the reference.
    #[must_use]
    pub const fn owner_uuid(self) -> UuidSnapshot {
        self.owner_uuid
    }

    /// The nested UUID referenced through a table UUID tuple, when present.
    #[must_use]
    pub const fn referenced_uuid(self) -> Option<UuidSnapshot> {
        self.referenced_uuid
    }
}

/// One dependency fact emitted by the formula-owner borrowed visitor.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FormulaDependencyFact {
    /// An expanded dependency points at an internal formula-owner ID.
    InternalOwner(FormulaOwnerInternalDependencyFact),
    /// A UUID-reference dependency points at an owner UUID.
    OwnerUuid(FormulaOwnerUuidDependencyFact),
}

/// Which spanning axis is being decoded. This is private to the owner field
/// router; callers receive the stable public kind on each emitted fact.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum SpanningDependencyAxis {
    Column,
    Row,
}

pub(super) fn decode_volatile_dependencies_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
    visitor: &mut dyn DependencyVisitor,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    // Fields 1..=5 and 7 are optional singular envelopes. Keep a tiny fixed
    // seen table so duplicate known fields fail without allocating.
    let mut seen = [false; 6];
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        let slot = match field.number {
            1..=5 => {
                usize::try_from(field.number - 1).map_err(|_conversion| DecodeError::invalid())?
            },
            7 => 5,
            _ => continue,
        };
        if seen[slot] {
            return Err(DecodeError::invalid());
        }
        seen[slot] = true;
        let payload = field.bytes()?;
        if field.number == 7 {
            decode_internal_cell_ref_set_in(
                payload,
                budget,
                child_depth,
                FormulaOwnerInternalDependencyKind::VolatileGeometry,
                visitor,
            )?;
        } else {
            decode_cell_coord_set_in(payload, budget, child_depth)?;
        }
    }
    Ok(())
}

pub(super) fn decode_spanning_dependencies_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
    axis: SpanningDependencyAxis,
    visitor: &mut dyn DependencyVisitor,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let kind = match axis {
        SpanningDependencyAxis::Column => FormulaOwnerInternalDependencyKind::SpanningColumn,
        SpanningDependencyAxis::Row => FormulaOwnerInternalDependencyKind::SpanningRow,
    };
    let mut total_range = false;
    let mut body_range = false;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => decode_cell_coord_refers_to_extents_in(
                field.bytes()?,
                budget,
                child_depth,
                kind,
                visitor,
            )?,
            2 => {
                if total_range {
                    return Err(DecodeError::invalid());
                }
                total_range = true;
                let _ = decode_range_coordinate_in(field.bytes()?, budget, child_depth)?;
            },
            3 => {
                if body_range {
                    return Err(DecodeError::invalid());
                }
                body_range = true;
                let _ = decode_range_coordinate_in(field.bytes()?, budget, child_depth)?;
            },
            _ => {},
        }
    }
    Ok(())
}

pub(super) fn decode_whole_owner_dependencies_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
    visitor: &mut dyn DependencyVisitor,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut dependent_cells = false;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        if field.number != 1 {
            continue;
        }
        if dependent_cells {
            return Err(DecodeError::invalid());
        }
        dependent_cells = true;
        decode_internal_cell_ref_set_in(
            field.bytes()?,
            budget,
            child_depth,
            FormulaOwnerInternalDependencyKind::WholeOwner,
            visitor,
        )?;
    }
    Ok(())
}

pub(super) fn decode_uuid_references_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
    visitor: &mut dyn DependencyVisitor,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => decode_table_reference_in(field.bytes()?, budget, child_depth, visitor)?,
            2 => decode_table_uuid_reference_in(field.bytes()?, budget, child_depth, visitor)?,
            _ => {},
        }
    }
    Ok(())
}

fn decode_cell_coord_set_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        if field.number == 1 {
            decode_cell_coord_column_entry_in(field.bytes()?, budget, child_depth)?;
        }
    }
    Ok(())
}

fn decode_cell_coord_column_entry_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut column = None;
    let mut row_set = false;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => super::set_once(&mut column, wire::canonical_u32(field.varint()?)?)?,
            2 => {
                if row_set {
                    return Err(DecodeError::invalid());
                }
                row_set = true;
                decode_index_set_in(field.bytes()?, budget, child_depth)?;
            },
            _ => {},
        }
    }
    if column.is_none() || !row_set {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn decode_index_set_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        if field.number == 1 {
            decode_index_set_entry_in(field.bytes()?, budget, child_depth)?;
        }
    }
    Ok(())
}

fn decode_index_set_entry_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let mut range_begin = None;
    let mut range_end = None;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => super::set_once(&mut range_begin, wire::canonical_int32(field.varint()?)?)?,
            2 => super::set_once(&mut range_end, wire::canonical_int32(field.varint()?)?)?,
            _ => {},
        }
    }
    if range_begin.is_none() {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn decode_internal_cell_ref_set_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
    kind: FormulaOwnerInternalDependencyKind,
    visitor: &mut dyn DependencyVisitor,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        if field.number == 1 {
            let owner_id =
                decode_internal_cell_ref_set_owner_entry_in(field.bytes()?, budget, child_depth)?;
            visitor.visit_formula_owner_internal_dependency(
                FormulaOwnerInternalDependencyFact::new(kind, owner_id),
            )?;
        }
    }
    Ok(())
}

fn decode_internal_cell_ref_set_owner_entry_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
) -> Result<u32, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut owner_id = None;
    let mut coord_set = false;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_owner_id(&mut owner_id, field.varint()?)?,
            2 => {
                if coord_set {
                    return Err(DecodeError::invalid());
                }
                coord_set = true;
                decode_cell_coord_set_in(field.bytes()?, budget, child_depth)?;
            },
            _ => {},
        }
    }
    if !coord_set {
        return Err(DecodeError::invalid());
    }
    owner_id.ok_or_else(DecodeError::invalid)
}

fn decode_cell_coord_refers_to_extents_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
    kind: FormulaOwnerInternalDependencyKind,
    visitor: &mut dyn DependencyVisitor,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut coordinate = None;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => super::set_once(
                &mut coordinate,
                decode_cell_coordinate_in(field.bytes()?, budget, child_depth)?,
            )?,
            2 => {
                let owner_id =
                    decode_extent_range_with_table_context_in(field.bytes()?, budget, child_depth)?;
                visitor.visit_formula_owner_internal_dependency(
                    FormulaOwnerInternalDependencyFact::new(kind, owner_id),
                )?;
            },
            _ => {},
        }
    }
    if coordinate.is_none() {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn decode_extent_range_with_table_context_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
) -> Result<u32, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut owner_id = None;
    let mut range_context = None;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_owner_id(&mut owner_id, field.varint()?)?,
            2 => super::set_once(&mut range_context, range_context_value(field.varint()?)?)?,
            3 => decode_extent_range_in(field.bytes()?, budget, child_depth)?,
            _ => {},
        }
    }
    if range_context.is_none() {
        return Err(DecodeError::invalid());
    }
    owner_id.ok_or_else(DecodeError::invalid)
}

fn decode_extent_range_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let mut extent_begin = None;
    let mut extent_end = None;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => super::set_once(&mut extent_begin, wire::canonical_u32(field.varint()?)?)?,
            2 => super::set_once(&mut extent_end, wire::canonical_u32(field.varint()?)?)?,
            _ => {},
        }
    }
    if extent_begin.is_none() {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn decode_table_reference_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
    visitor: &mut dyn DependencyVisitor,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut owner_uuid = None;
    let mut coord_set = false;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => super::set_once(
                &mut owner_uuid,
                decode_uuid(field.bytes()?, budget, child_depth)?,
            )?,
            2 => {
                if coord_set {
                    return Err(DecodeError::invalid());
                }
                coord_set = true;
                decode_cell_coord_set_in(field.bytes()?, budget, child_depth)?;
            },
            _ => {},
        }
    }
    let owner_uuid = owner_uuid.ok_or_else(DecodeError::invalid)?;
    visitor.visit_formula_owner_uuid_dependency(FormulaOwnerUuidDependencyFact::table_reference(
        owner_uuid,
    ))
}

fn decode_table_uuid_reference_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
    visitor: &mut dyn DependencyVisitor,
) -> Result<(), DecodeError> {
    // The owner UUID is required but protobuf does not require field order.
    // Validate the complete message once, then replay its nested UUID tuples
    // after the owner has been established, retaining a fixed-memory visitor
    // path for arbitrarily many repeated tuples.
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut owner_uuid = None;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => super::set_once(
                &mut owner_uuid,
                decode_uuid(field.bytes()?, budget, child_depth)?,
            )?,
            2 => {
                let _ = decode_uuid_reference_in(field.bytes()?, budget, child_depth, None, None)?;
            },
            _ => {},
        }
    }
    let owner_uuid = owner_uuid.ok_or_else(DecodeError::invalid)?;

    let mut nested_count = 0usize;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            // The first pass already validated the one required owner UUID;
            // replay only repeated nested tuples so their facts stay source
            // ordered without charging that scalar a third time.
            1 => {},
            2 => {
                decode_uuid_reference_in(
                    field.bytes()?,
                    budget,
                    child_depth,
                    Some(owner_uuid),
                    Some(visitor),
                )?;
                nested_count = nested_count
                    .checked_add(1)
                    .ok_or_else(DecodeError::invalid)?;
            },
            _ => {},
        }
    }
    if nested_count == 0 {
        visitor.visit_formula_owner_uuid_dependency(
            FormulaOwnerUuidDependencyFact::table_uuid_reference(owner_uuid, None),
        )?;
    }
    Ok(())
}

fn decode_uuid_reference_in(
    source: &[u8],
    budget: &mut wire::Budget,
    depth: u32,
    owner_uuid: Option<UuidSnapshot>,
    visitor: Option<&mut dyn DependencyVisitor>,
) -> Result<UuidSnapshot, DecodeError> {
    budget.message(source, depth)?;
    let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
    let mut uuid = None;
    let mut coord_set = false;
    let mut remaining = source;
    while let Some(field) = wire::next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => super::set_once(&mut uuid, decode_uuid(field.bytes()?, budget, child_depth)?)?,
            2 => {
                if coord_set {
                    return Err(DecodeError::invalid());
                }
                coord_set = true;
                decode_cell_coord_set_in(field.bytes()?, budget, child_depth)?;
            },
            _ => {},
        }
    }
    let uuid = uuid.ok_or_else(DecodeError::invalid)?;
    if let (Some(owner_uuid), Some(visitor)) = (owner_uuid, visitor) {
        visitor.visit_formula_owner_uuid_dependency(
            FormulaOwnerUuidDependencyFact::table_uuid_reference(owner_uuid, Some(uuid)),
        )?;
    }
    Ok(uuid)
}

fn set_owner_id(slot: &mut Option<u32>, value: u64) -> Result<(), DecodeError> {
    super::set_once(slot, wire::canonical_u32(value)?)
}

fn range_context_value(value: u64) -> Result<u32, DecodeError> {
    let value = wire::canonical_u32(value)?;
    match value {
        0 | 1 => Ok(value),
        _ => Err(DecodeError::invalid()),
    }
}
