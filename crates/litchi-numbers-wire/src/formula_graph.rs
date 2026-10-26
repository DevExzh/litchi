//! Bounded, source-preserving rewrites for the Numbers formula graph.
//!
//! The formula graph is stored in several native messages rather than in the
//! table model itself.  This module owns the small physical rewrite seam used
//! by table and sheet duplication: strict handwritten/Buffa projections prove
//! the source shape, then raw wire helpers change only the identity-bearing
//! fields.  Generated Prost values never cross this boundary.

use core::fmt;
use std::mem::size_of;

use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes,
    wire::{WireDescent, WireField, preflight_wire_tree},
};
use litchi_iwa_protos::numbers_table_cell_dependency_codec as dependency;

use crate::formula_owner::Uuid as FormulaOwnerUuid;

mod wire;

use wire::{
    Appender, append_length_delimited_field, append_varint_field, parse_fields_within,
    patch_length_delimited_field, patch_varint_field, repeated_varint_values,
    rewrite_repeated_varint_fields, transform_length_delimited_fields_at_path,
};

/// A resource governed by one formula-graph rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes in the source payload.
    InputBytes,
    /// Bytes in the rewritten payload.
    OutputBytes,
    /// Parsed wire fields, aggregated over every nested pass.
    Fields,
    /// Aggregate rewrite work.
    WorkBytes,
    /// Nested wire depth.
    Nesting,
    /// Native reference envelopes decoded by the strict dependency codec.
    References,
    /// Native text bytes decoded by the strict dependency codec.
    TextBytes,
    /// Temporary bytes live while a source-preserving rewrite is emitted.
    ScratchBytes,
    /// Fallible allocation events admitted by one rewrite.
    Allocations,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Fields => "fields",
            Self::WorkBytes => "rewrite work",
            Self::Nesting => "nesting depth",
            Self::References => "references",
            Self::TextBytes => "text bytes",
            Self::ScratchBytes => "scratch bytes",
            Self::Allocations => "allocations",
        })
    }
}

/// A typed formula-graph rewrite failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The source, or a nested envelope it selects, is not valid native wire.
    InvalidSource,
    /// A supplied identity map is unsorted, not unique, zero, or misses an
    /// identity the rewritten owner requires.
    InvalidRemap,
    /// The source holds a dependency family this rewrite cannot remap while
    /// keeping native semantics (volatile geometry references, spanning or
    /// whole-owner edges, cell errors, spill sizes, unknown owner-map fields).
    UnsupportedDependency,
    /// A caller-selected budget was exceeded.
    Limit {
        /// The exhausted resource.
        kind: LimitKind,
        /// The observed or requested amount.
        observed: usize,
        /// The caller-selected maximum.
        maximum: usize,
    },
    /// A caller asked for a budget of zero or above the hard ceiling.
    InvalidLimit {
        /// The rejected resource.
        kind: LimitKind,
        /// The requested value.
        value: usize,
        /// The hard ceiling.
        maximum: usize,
    },
    /// A fallible output or staging allocation was refused.
    Allocation {
        /// The bytes or items requested.
        requested: usize,
    },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidSource => formatter.write_str("invalid Numbers formula graph wire"),
            Self::InvalidRemap => formatter.write_str("invalid Numbers formula graph remap"),
            Self::UnsupportedDependency => {
                formatter.write_str("unsupported Numbers formula dependency")
            },
            Self::Limit {
                kind,
                observed,
                maximum,
            } => write!(
                formatter,
                "Numbers formula graph {kind} limit exceeded: observed {observed}, maximum {maximum}"
            ),
            Self::InvalidLimit {
                kind,
                value,
                maximum,
            } => write!(
                formatter,
                "invalid Numbers formula graph {kind} limit {value}, expected 1..={maximum}"
            ),
            Self::Allocation { requested } => {
                write!(
                    formatter,
                    "could not allocate {requested} formula graph bytes"
                )
            },
        }
    }
}

impl std::error::Error for Error {}

impl From<litchi_iwa_common::Error> for Error {
    fn from(error: litchi_iwa_common::Error) -> Self {
        match error {
            litchi_iwa_common::Error::LimitExceeded {
                kind,
                observed,
                limit,
            } => Self::Limit {
                kind: map_limit_kind(kind),
                observed,
                maximum: limit,
            },
            litchi_iwa_common::Error::Allocation { amount, .. } => {
                Self::Allocation { requested: amount }
            },
            _ => Self::InvalidSource,
        }
    }
}

fn map_limit_kind(kind: litchi_iwa_common::LimitKind) -> LimitKind {
    use litchi_iwa_common::LimitKind as Wire;

    match kind {
        Wire::InputBytes => LimitKind::InputBytes,
        Wire::Fields => LimitKind::Fields,
        Wire::OutputBytes => LimitKind::OutputBytes,
        Wire::Nesting => LimitKind::Nesting,
        _ => LimitKind::WorkBytes,
    }
}

/// Caller-selected budgets for one formula-graph rewrite.
///
/// One profile governs every nested validation, rewrite and final decode of
/// an operation: the ledger narrows the field and work budgets before each
/// pass, so no nested payload resets them. [`Limits::DEFAULT`] is the set of
/// hard ceilings; owners tighten it from their own retained-source budgets.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    wire: WireLimits,
    max_references: usize,
    max_scratch_bytes: usize,
    max_allocations: usize,
}

impl Limits {
    /// Hard native-reference ceiling.
    pub const MAX_REFERENCES: usize = 1_000_000;
    /// Hard temporary-byte ceiling.
    pub const MAX_SCRATCH_BYTES: usize = 1_024 * 1024 * 1024;
    /// Hard allocation-event ceiling.
    pub const MAX_ALLOCATIONS: usize = 1_000_000;
    /// The hard ceilings of every dimension.
    pub const DEFAULT: Self = Self {
        wire: WireLimits::DEFAULT,
        max_references: Self::MAX_REFERENCES,
        max_scratch_bytes: Self::MAX_SCRATCH_BYTES,
        max_allocations: Self::MAX_ALLOCATIONS,
    };

    /// Use `wire` for the parse, field, nesting, work and output budgets.
    #[must_use]
    pub const fn with_wire(mut self, wire: WireLimits) -> Self {
        self.wire = wire;
        self
    }

    /// Tighten the native-reference budget.
    ///
    /// # Errors
    ///
    /// [`Error::InvalidLimit`] for zero or a value above the hard ceiling.
    pub const fn with_references(mut self, value: usize) -> Result<Self, Error> {
        match checked(value, Self::MAX_REFERENCES, LimitKind::References) {
            Ok(value) => {
                self.max_references = value;
                Ok(self)
            },
            Err(error) => Err(error),
        }
    }

    /// Tighten the temporary-byte budget.
    ///
    /// # Errors
    ///
    /// [`Error::InvalidLimit`] for zero or a value above the hard ceiling.
    pub const fn with_scratch_bytes(mut self, value: usize) -> Result<Self, Error> {
        match checked(value, Self::MAX_SCRATCH_BYTES, LimitKind::ScratchBytes) {
            Ok(value) => {
                self.max_scratch_bytes = value;
                Ok(self)
            },
            Err(error) => Err(error),
        }
    }

    /// Tighten the allocation-event budget.
    ///
    /// # Errors
    ///
    /// [`Error::InvalidLimit`] for zero or a value above the hard ceiling.
    pub const fn with_allocations(mut self, value: usize) -> Result<Self, Error> {
        match checked(value, Self::MAX_ALLOCATIONS, LimitKind::Allocations) {
            Ok(value) => {
                self.max_allocations = value;
                Ok(self)
            },
            Err(error) => Err(error),
        }
    }

    /// The wire budgets.
    #[must_use]
    pub const fn wire(self) -> WireLimits {
        self.wire
    }

    /// The native-reference budget.
    #[must_use]
    pub const fn max_references(self) -> usize {
        self.max_references
    }

    /// The temporary-byte budget.
    #[must_use]
    pub const fn max_scratch_bytes(self) -> usize {
        self.max_scratch_bytes
    }

    /// The allocation-event budget.
    #[must_use]
    pub const fn max_allocations(self) -> usize {
        self.max_allocations
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

const fn checked(value: usize, maximum: usize, kind: LimitKind) -> Result<usize, Error> {
    if value == 0 || value > maximum {
        return Err(Error::InvalidLimit {
            kind,
            value,
            maximum,
        });
    }
    Ok(value)
}

/// Aggregate resource ledger for one formula-graph operation.
///
/// Native dependency snapshots report their exact field, work, and reference
/// totals.  This ledger carries those totals across every nested validation,
/// rewrite, and final decode so each pass cannot reset the caller's budget.
#[derive(Debug)]
struct RewriteBudget {
    limits: Limits,
    base_live: usize,
    retained_live: usize,
    fields: usize,
    work: usize,
    references: usize,
    allocations: usize,
}

impl RewriteBudget {
    const fn new(limits: Limits) -> Self {
        Self {
            limits,
            base_live: 0,
            retained_live: 0,
            fields: 0,
            work: 0,
            references: 0,
            allocations: 0,
        }
    }

    fn admit_source(&mut self, source: &[u8]) -> Result<(), Error> {
        if source.len() > self.limits.wire.max_input_bytes() {
            return Err(Error::Limit {
                kind: LimitKind::InputBytes,
                observed: source.len(),
                maximum: self.limits.wire.max_input_bytes(),
            });
        }
        if source.len() > self.limits.wire.max_output_bytes() {
            return Err(Error::Limit {
                kind: LimitKind::OutputBytes,
                observed: source.len(),
                maximum: self.limits.wire.max_output_bytes(),
            });
        }
        self.base_live = source.len();
        Ok(())
    }

    fn add(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: LimitKind,
    ) -> Result<(), Error> {
        let observed = current.checked_add(amount).ok_or(Error::Limit {
            kind,
            observed: usize::MAX,
            maximum,
        })?;
        if observed > maximum {
            return Err(Error::Limit {
                kind,
                observed,
                maximum,
            });
        }
        *current = observed;
        Ok(())
    }

    fn charge_fields(&mut self, amount: usize) -> Result<(), Error> {
        Self::add(
            &mut self.fields,
            amount,
            self.limits.wire.max_fields(),
            LimitKind::Fields,
        )
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), Error> {
        Self::add(
            &mut self.work,
            amount,
            self.limits.wire.max_rewrite_work(),
            LimitKind::WorkBytes,
        )
    }

    fn charge_references(&mut self, amount: usize) -> Result<(), Error> {
        Self::add(
            &mut self.references,
            amount,
            self.limits.max_references,
            LimitKind::References,
        )
    }

    fn charge_allocation(&mut self) -> Result<(), Error> {
        Self::add(
            &mut self.allocations,
            1,
            self.limits.max_allocations,
            LimitKind::Allocations,
        )
    }

    fn charge_allocations(&mut self, amount: usize) -> Result<(), Error> {
        Self::add(
            &mut self.allocations,
            amount,
            self.limits.max_allocations,
            LimitKind::Allocations,
        )
    }

    fn live_bytes(&self, additional: usize) -> usize {
        self.base_live.max(self.retained_live).max(additional)
    }

    fn admit_scratch(&mut self, live: usize, temporary: usize) -> Result<(), Error> {
        let observed = self
            .live_bytes(live)
            .checked_add(temporary)
            .ok_or(Error::Limit {
                kind: LimitKind::ScratchBytes,
                observed: usize::MAX,
                maximum: self.limits.max_scratch_bytes,
            })?;
        if observed > self.limits.max_scratch_bytes {
            return Err(Error::Limit {
                kind: LimitKind::ScratchBytes,
                observed,
                maximum: self.limits.max_scratch_bytes,
            });
        }
        self.retained_live = self.retained_live.max(observed);
        Ok(())
    }

    fn admit_output(
        &mut self,
        live: usize,
        temporary: usize,
        output: usize,
    ) -> Result<Vec<u8>, Error> {
        if output > self.limits.wire.max_output_bytes() {
            return Err(Error::Limit {
                kind: LimitKind::OutputBytes,
                observed: output,
                maximum: self.limits.wire.max_output_bytes(),
            });
        }
        let scratch = self
            .live_bytes(live)
            .checked_add(temporary)
            .and_then(|value| value.checked_add(output))
            .ok_or(Error::Limit {
                kind: LimitKind::ScratchBytes,
                observed: usize::MAX,
                maximum: self.limits.max_scratch_bytes,
            })?;
        if scratch > self.limits.max_scratch_bytes {
            return Err(Error::Limit {
                kind: LimitKind::ScratchBytes,
                observed: scratch,
                maximum: self.limits.max_scratch_bytes,
            });
        }
        self.retained_live = self.retained_live.max(scratch);
        self.charge_allocation()?;
        self.charge_work(output)?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(output)
            .map_err(|_| Error::Allocation { requested: output })?;
        Ok(bytes)
    }

    /// Parse one message's root fields under the residual field budget,
    /// then charge the retained field table to the ledger.
    ///
    /// The shared parser enforces the residual field ceiling while it grows
    /// the table, so the table can never exceed what the ledger still admits.
    fn parse_fields(&mut self, source: &[u8], live: usize) -> Result<Vec<WireField>, Error> {
        let limits = self.child_wire_limits()?;
        let fields = parse_fields_within(source, limits).map_err(Error::from)?;
        let field_bytes =
            fields
                .len()
                .checked_mul(size_of::<WireField>())
                .ok_or(Error::Allocation {
                    requested: usize::MAX,
                })?;
        self.admit_scratch(live, field_bytes)?;
        self.charge_allocation()?;
        self.charge_fields(fields.len())?;
        self.charge_work(source.len().checked_add(field_bytes).ok_or(Error::Limit {
            kind: LimitKind::WorkBytes,
            observed: usize::MAX,
            maximum: self.limits.wire.max_rewrite_work(),
        })?)?;
        Ok(fields)
    }

    fn decode_options(&self, source: &[u8]) -> Result<dependency::DecodeOptions, Error> {
        let fields = self.remaining(
            self.limits.wire.max_fields(),
            self.fields,
            LimitKind::Fields,
        )?;
        let work = self.remaining(
            self.limits.wire.max_rewrite_work(),
            self.work,
            LimitKind::WorkBytes,
        )?;
        let references = self.remaining(
            self.limits.max_references,
            self.references,
            LimitKind::References,
        )?;
        Ok(dependency::DecodeOptions::new(
            source.len(),
            fields,
            work,
            u32::try_from(self.limits.wire.max_nesting()).unwrap_or(u32::MAX),
            references,
            0,
        ))
    }

    fn remaining(&self, maximum: usize, used: usize, kind: LimitKind) -> Result<usize, Error> {
        let remaining = maximum.checked_sub(used).ok_or(Error::Limit {
            kind,
            observed: usize::MAX,
            maximum,
        })?;
        if remaining == 0 {
            return Err(Error::Limit {
                kind,
                observed: used.saturating_add(1),
                maximum,
            });
        }
        Ok(remaining)
    }

    fn charge_decode_report(&mut self, report: dependency::DecodeReport) -> Result<(), Error> {
        self.charge_fields(report.fields())?;
        self.charge_work(report.work_bytes())?;
        self.charge_references(report.references())
    }

    fn child_wire_limits(&self) -> Result<WireLimits, Error> {
        let fields = self.remaining(
            self.limits.wire.max_fields(),
            self.fields,
            LimitKind::Fields,
        )?;
        let work = self.remaining(
            self.limits.wire.max_rewrite_work(),
            self.work,
            LimitKind::WorkBytes,
        )?;
        self.limits
            .wire
            .with_fields(fields)
            .and_then(|limits| limits.with_rewrite_work(work))
            .map_err(Error::from)
    }

    /// Reserve the aggregate ledger before a shared wire helper is allowed to
    /// allocate its replacement vector and final output.  The helper receives
    /// the residual field/work profile, so a repeated path cannot reset those
    /// ceilings for every nested payload.
    fn admit_transform(
        &mut self,
        source: &[u8],
        live: usize,
        path: &[u32],
    ) -> Result<WireLimits, Error> {
        let source_temporary = source
            .len()
            .checked_mul(3)
            .map(|value| value.min(self.limits.wire.max_output_bytes()))
            .ok_or(Error::Limit {
                kind: LimitKind::ScratchBytes,
                observed: usize::MAX,
                maximum: self.limits.max_scratch_bytes,
            })?;
        let child_limits = self.child_wire_limits()?;
        let preflight = preflight_wire_tree(source, child_limits, |visit| {
            let depth = visit.path().len();
            if depth < path.len()
                && visit.field().wire_type() == 2
                && visit.field().number() == path[depth]
            {
                Ok(WireDescent::Descend)
            } else {
                Ok(WireDescent::Skip)
            }
        })
        .map_err(Error::from)?;
        let metadata_per_field = size_of::<WireField>()
            .checked_mul(3)
            .and_then(|value| value.checked_add(size_of::<&[u8]>()))
            .and_then(|value| value.checked_add(size_of::<Vec<u8>>()))
            .ok_or(Error::Limit {
                kind: LimitKind::ScratchBytes,
                observed: usize::MAX,
                maximum: self.limits.max_scratch_bytes,
            })?;
        let metadata = preflight
            .fields()
            .checked_mul(metadata_per_field)
            .ok_or(Error::Limit {
                kind: LimitKind::ScratchBytes,
                observed: usize::MAX,
                maximum: self.limits.max_scratch_bytes,
            })?;
        let temporary = source_temporary.checked_add(metadata).ok_or(Error::Limit {
            kind: LimitKind::ScratchBytes,
            observed: usize::MAX,
            maximum: self.limits.max_scratch_bytes,
        })?;
        self.admit_scratch(live, temporary)?;
        // Each common transform message retains parsed fields and payloads,
        // then builds replacements, matching fields, and output. Reserve the
        // complete helper allocation footprint before it gets control.
        let helper_allocations = preflight.messages().checked_mul(6).ok_or(Error::Limit {
            kind: LimitKind::Allocations,
            observed: usize::MAX,
            maximum: self.limits.max_allocations,
        })?;
        self.charge_allocations(helper_allocations)?;
        self.charge_fields(preflight.fields().checked_mul(3).ok_or(Error::Limit {
            kind: LimitKind::Fields,
            observed: usize::MAX,
            maximum: self.limits.wire.max_fields(),
        })?)?;
        self.charge_work(
            preflight
                .scanned_bytes()
                .checked_mul(3)
                .ok_or(Error::Limit {
                    kind: LimitKind::WorkBytes,
                    observed: usize::MAX,
                    maximum: self.limits.wire.max_rewrite_work(),
                })?,
        )?;
        Ok(child_limits)
    }
}

/// A source-to-target object, internal-owner, and UUID remap used by a graph
/// clone.  Each map must be sorted by its source key and contain no duplicate
/// source or target identity.  Unlisted identities remain source-authoritative
/// so cross-table dependencies stay attached to their original owner.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Remap<'a> {
    /// Archive object identifiers, including the source table-info object and
    /// cloned formula-owner/tile objects.
    object_ids: &'a [(u64, u64)],
    /// Native internal formula-owner identifiers.
    owner_ids: &'a [(u32, u32)],
    /// TSP.UUID and CFUUID owner identities.
    uuids: &'a [(FormulaOwnerUuid, FormulaOwnerUuid)],
}

impl<'a> Remap<'a> {
    /// A checked borrowed graph remap.
    ///
    /// Each map must be sorted by source, with no zero identity and no
    /// repeated source or target. The check runs once here, so every rewrite
    /// that takes the remap trusts it (record 0368a).
    ///
    /// # Errors
    ///
    /// [`Error::InvalidRemap`] for an unsorted, repeated or zero identity;
    /// [`Error::Allocation`] when the target check cannot reserve its copy.
    pub fn new(
        object_ids: &'a [(u64, u64)],
        owner_ids: &'a [(u32, u32)],
        uuids: &'a [(FormulaOwnerUuid, FormulaOwnerUuid)],
    ) -> Result<Self, Error> {
        let remap = Self {
            object_ids,
            owner_ids,
            uuids,
        };
        // The maps are caller-owned memory; their check is bounded by their
        // own length, not by a rewrite's budget.
        let mut budget = RewriteBudget::new(Limits::DEFAULT);
        validate_u64_map(remap.object_ids, &mut budget)?;
        validate_u32_map(remap.owner_ids, &mut budget)?;
        validate_uuid_map(remap.uuids, &mut budget)?;
        Ok(remap)
    }
}

/// Rewrite one `FormulaOwnerDependenciesArchive` while retaining unknown
/// fields, original field ordering, and untouched nested bytes.
pub fn remap_owner(source: &[u8], remap: Remap<'_>, limits: &Limits) -> Result<Vec<u8>, Error> {
    let mut budget = RewriteBudget::new(*limits);
    budget.admit_source(source)?;
    let owner = decode_owner(source, &mut budget)?;
    validate_owner_nested_fields(source, &mut budget)?;
    reject_unsupported_dependencies(&owner, &mut budget)?;
    if lookup_uuid(
        remap.uuids,
        FormulaOwnerUuid::new(
            owner.formula_owner_uid().lower(),
            owner.formula_owner_uid().upper(),
        ),
    )
    .is_none()
        || lookup_u32(remap.owner_ids, owner.internal_formula_owner_id()).is_none()
    {
        return Err(Error::InvalidRemap);
    }
    if let Some(formula_owner) = owner.formula_owner() {
        // A present native table-owner reference is part of the identity
        // family.  Never silently preserve a zero or unmapped reference while
        // remapping the owner that contains it.
        if formula_owner.identifier() == 0 {
            return Err(Error::InvalidSource);
        }
        if lookup_u64(remap.object_ids, formula_owner.identifier()).is_none() {
            return Err(Error::InvalidRemap);
        }
    }

    let mut output = rewrite_paths(source, OWNER_RULES, remap, &mut budget)?;
    output = remap_owner_id_field(
        &output,
        2,
        owner.internal_formula_owner_id(),
        remap.owner_ids,
        &mut budget,
    )?;
    decode_owner(&output, &mut budget)?;
    Ok(output)
}

/// The identity-bearing leaf a path rule rewrites.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Leaf {
    /// A `TSP.UUID` remapped through [`Remap::uuids`].
    Uuid,
    /// A `TSP.CFUUIDArchive` remapped through [`Remap::uuids`].
    Cfuuid,
    /// A `TSP.Reference` remapped through [`Remap::object_ids`].
    Reference,
    /// An internal range reference whose owner (field 1) is remapped through
    /// [`Remap::owner_ids`].
    InternalId,
    /// Expanded edges whose repeated owner identifiers (field 5) are remapped
    /// through [`Remap::owner_ids`].
    EdgeIds,
}

/// The identity-bearing paths of a `FormulaOwnerDependenciesArchive`
/// (record 0368: the same paths the host's sequential rewrite visits).
const OWNER_RULES: &[(&[u32], Leaf)] = &[
    (&[1], Leaf::Uuid),
    (&[11], Leaf::Reference),
    (&[12], Leaf::Uuid),
    (&[4, 1, 6], Leaf::EdgeIds),
    (&[5, 2, 4], Leaf::InternalId),
    (&[5, 2, 3, 1], Leaf::Cfuuid),
    (&[13, 1], Leaf::Reference),
    (&[14, 1, 1], Leaf::Uuid),
    (&[14, 2, 1], Leaf::Uuid),
    (&[14, 2, 2, 1], Leaf::Uuid),
    (&[15, 1], Leaf::Reference),
];

/// The expanded-edge paths of a `CellRecordTileArchive`.
const CELL_TILE_RULES: &[(&[u32], Leaf)] = &[(&[4, 6], Leaf::EdgeIds)];

/// What a rule set says about field `number` inside the message at `prefix`.
fn rule_at(rules: &[(&[u32], Leaf)], prefix: &[u32], number: u32) -> Option<Option<Leaf>> {
    let depth = prefix.len();
    let mut descend = false;
    for (path, leaf) in rules {
        if path.len() > depth && path[..depth] == *prefix && path[depth] == number {
            if path.len() == depth + 1 {
                return Some(Some(*leaf));
            }
            descend = true;
        }
    }
    descend.then_some(None)
}

/// Rewrite every rule path of `source` in one bounded walk.
///
/// Each visited message is parsed once; nothing is preflighted separately.
///
/// The output equals applying each path's repeated transform in turn: a
/// message whose rewritten fields are all byte-equal is copied verbatim, and
/// otherwise every occurrence of a changed field number keeps its key and
/// takes a canonical length prefix, exactly as the shared repeated rewrite
/// does. Untouched fields, unknown fields and field order are preserved.
fn rewrite_paths(
    source: &[u8],
    rules: &[(&[u32], Leaf)],
    remap: Remap<'_>,
    budget: &mut RewriteBudget,
) -> Result<Vec<u8>, Error> {
    // Every message the walk visits is admitted by `parse_fields` before its
    // field table is retained, every rewritten message by `admit_output`, and
    // the recursion is bounded by the longest rule path.
    let depth = rules.iter().map(|(path, _)| path.len()).max().unwrap_or(0);
    let mut prefix = Vec::new();
    budget.charge_allocation()?;
    prefix
        .try_reserve_exact(depth)
        .map_err(|_| Error::Allocation {
            requested: depth * size_of::<u32>(),
        })?;
    match rewrite_message(source, rules, &mut prefix, remap, budget)? {
        Some(output) => Ok(output),
        None => clone_output(source, budget, 0),
    }
}

/// Rewrite one message; `None` means nothing below it changed.
fn rewrite_message(
    source: &[u8],
    rules: &[(&[u32], Leaf)],
    prefix: &mut Vec<u32>,
    remap: Remap<'_>,
    budget: &mut RewriteBudget,
) -> Result<Option<Vec<u8>>, Error> {
    let fields = budget.parse_fields(source, source.len())?;
    let mut replacements: Vec<(usize, Vec<u8>)> = Vec::new();
    let mut changed_numbers: Vec<u32> = Vec::new();
    for (index, field) in fields.iter().enumerate() {
        let Some(rule) = rule_at(rules, prefix, field.number()) else {
            continue;
        };
        if field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        let payload = field.payload(source).map_err(Error::from)?;
        let replacement = match rule {
            Some(leaf) => rewrite_leaf(payload, leaf, remap, budget)?,
            None => {
                prefix.push(field.number());
                let nested = rewrite_message(payload, rules, prefix, remap, budget);
                prefix.pop();
                nested?
            },
        };
        if let Some(replacement) = replacement {
            if replacement.as_slice() != payload {
                if !changed_numbers.contains(&field.number()) {
                    budget.charge_allocation()?;
                    changed_numbers.push(field.number());
                }
                budget.charge_allocation()?;
                replacements.push((index, replacement));
            }
        }
    }
    if changed_numbers.is_empty() {
        return Ok(None);
    }
    let mut output_len = source.len();
    let mut next = 0usize;
    for (index, field) in fields.iter().enumerate() {
        if !changed_numbers.contains(&field.number()) {
            continue;
        }
        let payload_len = match replacements.get(next) {
            Some((replacement_index, replacement)) if *replacement_index == index => {
                next += 1;
                replacement.len()
            },
            _ => field.end() - field.payload_start(),
        };
        output_len = output_len
            .checked_sub(field.end() - field.key_end())
            .and_then(|value| value.checked_add(encoded_len_usize(payload_len)))
            .and_then(|value| value.checked_add(payload_len))
            .ok_or(Error::InvalidSource)?;
    }
    let mut output = budget.admit_output(
        source.len(),
        fields
            .len()
            .checked_mul(size_of::<WireField>())
            .ok_or(Error::Allocation {
                requested: usize::MAX,
            })?,
        output_len,
    )?;
    let mut copied = 0usize;
    let mut next = 0usize;
    for (index, field) in fields.iter().enumerate() {
        if !changed_numbers.contains(&field.number()) {
            continue;
        }
        output.extend_from_slice(&source[copied..field.key_end()]);
        let payload = match replacements.get(next) {
            Some((replacement_index, replacement)) if *replacement_index == index => {
                next += 1;
                replacement.as_slice()
            },
            _ => &source[field.payload_start()..field.end()],
        };
        litchi_iwa_common::varint::encode_varint_into(&mut output, payload.len() as u64);
        output.extend_from_slice(payload);
        copied = field.end();
    }
    output.extend_from_slice(&source[copied..]);
    debug_assert_eq!(output.len(), output_len);
    Ok(Some(output))
}

fn encoded_len_usize(value: usize) -> usize {
    litchi_iwa_common::varint::encoded_len(value as u64)
}

/// Rewrite one leaf payload; `None` means its identity is not remapped.
fn rewrite_leaf(
    payload: &[u8],
    leaf: Leaf,
    remap: Remap<'_>,
    budget: &mut RewriteBudget,
) -> Result<Option<Vec<u8>>, Error> {
    match leaf {
        Leaf::Uuid => {
            if remap.uuids.is_empty() {
                return Ok(None);
            }
            let value = parse_uuid_with_budget(payload, budget)?;
            let Some((_, replacement)) = lookup_uuid(remap.uuids, value) else {
                return Ok(None);
            };
            patch_uuid(payload, replacement, budget).map(Some)
        },
        Leaf::Cfuuid => {
            if remap.uuids.is_empty() {
                return Ok(None);
            }
            let value = parse_cfuuid_with_budget(payload, budget)?;
            let Some((_, replacement)) = lookup_uuid(remap.uuids, value) else {
                return Ok(None);
            };
            patch_cfuuid(payload, replacement, budget).map(Some)
        },
        Leaf::Reference => {
            if remap.object_ids.is_empty() {
                return Ok(None);
            }
            let identifier = required_varint_with_budget(payload, 1, budget)?;
            let Some(replacement) = lookup_u64(remap.object_ids, identifier) else {
                return Ok(None);
            };
            patch_required_varint(payload, 1, identifier, replacement, budget).map(Some)
        },
        Leaf::InternalId => {
            if remap.owner_ids.is_empty() {
                return Ok(None);
            }
            let identifier = required_varint_with_budget(payload, 1, budget)?;
            let identifier = u32::try_from(identifier).map_err(|_| Error::InvalidSource)?;
            let replacement = lookup_u32(remap.owner_ids, identifier).unwrap_or(identifier);
            patch_required_varint(
                payload,
                1,
                u64::from(identifier),
                u64::from(replacement),
                budget,
            )
            .map(Some)
        },
        Leaf::EdgeIds => {
            if remap.owner_ids.is_empty() {
                return Ok(None);
            }
            let limits = budget.child_wire_limits()?;
            budget.charge_allocations(5)?;
            let values = repeated_varint_values(payload, 5, limits).map_err(Error::from)?;
            let replacement_bytes =
                values
                    .len()
                    .checked_mul(size_of::<u64>())
                    .ok_or(Error::Allocation {
                        requested: usize::MAX,
                    })?;
            budget.admit_scratch(payload.len(), replacement_bytes)?;
            let mut replacements = Vec::new();
            replacements
                .try_reserve_exact(values.len())
                .map_err(|_| Error::Allocation {
                    requested: replacement_bytes,
                })?;
            for value in values {
                let value = u32::try_from(value).map_err(|_| Error::InvalidSource)?;
                replacements.push(u64::from(
                    lookup_u32(remap.owner_ids, value).unwrap_or(value),
                ));
            }
            rewrite_repeated_varint_fields(payload, 5, &replacements, limits)
                .map(Some)
                .map_err(Error::from)
        },
    }
}

/// Rewrite one `CellRecordTileArchive`, remapping its owner and all internal
/// owner IDs in expanded edge arrays.
pub fn remap_cell_record_tile(
    source: &[u8],
    remap: Remap<'_>,
    limits: &Limits,
) -> Result<Vec<u8>, Error> {
    let mut budget = RewriteBudget::new(*limits);
    budget.admit_source(source)?;
    let tile = decode_cell_tile(source, &mut budget)?;
    if lookup_u32(remap.owner_ids, tile.internal_owner_id()).is_none() {
        return Err(Error::InvalidRemap);
    }
    let mut output = patch_required_varint(
        source,
        1,
        u64::from(tile.internal_owner_id()),
        lookup_u32(remap.owner_ids, tile.internal_owner_id())
            .map(u64::from)
            .ok_or(Error::InvalidRemap)?,
        &mut budget,
    )?;
    output = rewrite_paths(&output, CELL_TILE_RULES, remap, &mut budget)?;
    // The rewrite replaces only owner-identifier varints that the strict
    // source decode admitted, with checked non-zero `u32` targets, so the
    // result's structure is the source's; re-decoding it in release builds
    // would only repeat the source decode (record 0368a).
    #[cfg(debug_assertions)]
    decode_cell_tile(&output, &mut budget)?;
    Ok(output)
}

/// Rewrite one `RangePrecedentsTileArchive`, remapping its destination owner
/// when that owner belongs to the cloned family.
pub fn remap_range_precedents_tile(
    source: &[u8],
    remap: Remap<'_>,
    limits: &Limits,
) -> Result<Vec<u8>, Error> {
    let mut budget = RewriteBudget::new(*limits);
    budget.admit_source(source)?;
    let tile = decode_range_tile(source, &mut budget)?;
    let Some(replacement) = lookup_u32(remap.owner_ids, tile.to_owner_id()) else {
        return clone_output(source, &mut budget, 0);
    };
    let output = patch_required_varint(
        source,
        1,
        u64::from(tile.to_owner_id()),
        u64::from(replacement),
        &mut budget,
    )?;
    decode_range_tile(&output, &mut budget)?;
    Ok(output)
}

/// The first internal formula-owner identifier free in a calculation engine.
///
/// The identifier is one above the largest identifier in use anywhere the
/// engine registers one: every owner payload's `internal_formula_owner_id`
/// (field 2), and every owner-map entry and unregistered identifier of the
/// engine's dependency tracker. Every larger identifier is free as well, so a
/// clone of `n` owners takes `next..next + n`.
///
/// The legacy host counted only owner payloads, so on engines whose owner
/// map holds identifiers without an owner object it allocated identifiers
/// already in the map (record 0368a); this allocator never does.
///
/// # Errors
///
/// [`Error::InvalidSource`] for malformed wire, a missing owner identifier or
/// an exhausted identifier space; [`Error::UnsupportedDependency`] for an
/// owner map with unknown fields; [`Error::Limit`] / [`Error::Allocation`].
pub fn next_owner_id(engine: &[u8], owners: &[&[u8]], limits: &Limits) -> Result<u32, Error> {
    let mut budget = RewriteBudget::new(*limits);
    budget.admit_source(engine)?;
    let decoded = decode_engine(engine, &mut budget)?;
    let tracker = decode_tracker(decoded.dependency_tracker(), &mut budget)?;
    let map = validate_owner_id_map(tracker.owner_id_map(), &mut budget)?;
    let mut maximum = map
        .identities
        .iter()
        .map(|(internal, _)| *internal)
        .chain(map.unregistered.iter().copied())
        .max()
        .unwrap_or(0);
    for owner in owners {
        let identifier = required_varint_with_budget(owner, 2, &mut budget)?;
        let identifier = u32::try_from(identifier).map_err(|_| Error::InvalidSource)?;
        maximum = maximum.max(identifier);
    }
    maximum.checked_add(1).ok_or(Error::InvalidSource)
}

/// Append cloned formula-owner references and owner-map entries to a native
/// `CalculationEngineArchive` payload, updating its formula count atomically.
/// Existing tracker/root unknown fields remain source-preserving.
pub fn append_owners_to_engine(
    source: &[u8],
    owner_object_ids: &[u64],
    owner_map_entries: &[(u32, FormulaOwnerUuid)],
    formula_count: u64,
    limits: &Limits,
) -> Result<Vec<u8>, Error> {
    let mut budget = RewriteBudget::new(*limits);
    budget.admit_source(source)?;
    validate_append_inputs(owner_object_ids, owner_map_entries, &mut budget)?;
    let engine = decode_engine(source, &mut budget)?;
    let tracker = decode_tracker(engine.dependency_tracker(), &mut budget)?;
    let existing_owner_map = validate_owner_id_map(tracker.owner_id_map(), &mut budget)?;
    for (internal, uuid) in owner_map_entries {
        if existing_owner_map
            .identities
            .binary_search_by_key(internal, |(known, _)| *known)
            .is_ok()
            || existing_owner_map
                .uuid_keys
                .binary_search(&uuid_key(*uuid))
                .is_ok()
            || existing_owner_map
                .unregistered
                .binary_search(internal)
                .is_ok()
        {
            return Err(Error::InvalidRemap);
        }
    }
    validate_tracker_owner_references(engine.dependency_tracker(), owner_object_ids, &mut budget)?;
    let previous_count = tracker.number_of_formulas().unwrap_or(0);
    let next_count = previous_count
        .checked_add(formula_count)
        .ok_or(Error::InvalidSource)?;
    if owner_object_ids.is_empty() && formula_count == 0 {
        return clone_output(source, &mut budget, 0);
    }

    let child_limits = budget.admit_transform(source, source.len(), &[2])?;
    let output = transform_length_delimited_fields_at_path(
        source,
        &[2],
        |tracker_data| {
            let mut candidate = clone_output(tracker_data, &mut budget, 0)?;
            let owner_reference_capacity =
                owner_object_ids
                    .iter()
                    .try_fold(0usize, |total, identifier| {
                        let reference_len = encoded_varint_field_size(1, *identifier)?;
                        total
                            .checked_add(encoded_length_delimited_field_size(6, reference_len)?)
                            .ok_or(Error::InvalidSource)
                    })?;
            if owner_reference_capacity != 0 {
                budget.admit_scratch(candidate.len(), owner_reference_capacity)?;
                budget.charge_allocation()?;
                candidate
                    .try_reserve_exact(owner_reference_capacity)
                    .map_err(|_| Error::Allocation {
                        requested: owner_reference_capacity,
                    })?;
            }
            let mut appender =
                Appender::of(&candidate, budget.child_wire_limits()?).map_err(Error::from)?;
            for identifier in owner_object_ids {
                let reference = encode_reference(*identifier, &mut budget, candidate.len())?;
                let additional = encoded_length_delimited_field_size(6, reference.len())?;
                budget.charge_work(
                    candidate
                        .len()
                        .checked_add(additional)
                        .ok_or(Error::InvalidSource)?,
                )?;
                budget.charge_allocations(2)?;
                let append_limits = budget.child_wire_limits()?;
                appender
                    .length_delimited(&mut candidate, 6, &reference, append_limits)
                    .map_err(Error::from)?;
            }
            if tracker.owner_id_map().is_some() {
                let map_limits = budget.admit_transform(&candidate, candidate.len(), &[3])?;
                candidate = transform_length_delimited_fields_at_path(
                    &candidate,
                    &[3],
                    |map_data| {
                        let mut map = clone_output(map_data, &mut budget, candidate.len())?;
                        let map_capacity = owner_map_entries.iter().try_fold(
                            0usize,
                            |total, (internal_owner_id, uuid)| {
                                total
                                    .checked_add(encoded_owner_map_field_size(
                                        *internal_owner_id,
                                        *uuid,
                                    )?)
                                    .ok_or(Error::InvalidSource)
                            },
                        )?;
                        if map_capacity != 0 {
                            budget.admit_scratch(candidate.len(), map_capacity)?;
                            budget.charge_allocation()?;
                            map.try_reserve_exact(map_capacity)
                                .map_err(|_| Error::Allocation {
                                    requested: map_capacity,
                                })?;
                        }
                        let mut appender =
                            Appender::of(&map, budget.child_wire_limits()?).map_err(Error::from)?;
                        for (internal_owner_id, uuid) in owner_map_entries {
                            let entry = encode_owner_map_entry(
                                *internal_owner_id,
                                *uuid,
                                &mut budget,
                                map.len(),
                            )?;
                            let additional = encoded_length_delimited_field_size(1, entry.len())?;
                            budget.charge_work(
                                map.len()
                                    .checked_add(additional)
                                    .ok_or(Error::InvalidSource)?,
                            )?;
                            budget.charge_allocations(2)?;
                            let append_limits = budget.child_wire_limits()?;
                            appender
                                .length_delimited(&mut map, 1, &entry, append_limits)
                                .map_err(Error::from)?;
                        }
                        Ok::<Vec<u8>, Error>(map)
                    },
                    map_limits,
                )?;
            } else {
                let mut map = Vec::new();
                budget.charge_allocation()?;
                let map_capacity = owner_map_entries.iter().try_fold(
                    0usize,
                    |total, (internal_owner_id, uuid)| {
                        total
                            .checked_add(encoded_owner_map_field_size(*internal_owner_id, *uuid)?)
                            .ok_or(Error::InvalidSource)
                    },
                )?;
                budget.admit_scratch(candidate.len(), map_capacity)?;
                map.try_reserve_exact(map_capacity)
                    .map_err(|_| Error::Allocation {
                        requested: map_capacity,
                    })?;
                let mut appender =
                    Appender::of(&map, budget.child_wire_limits()?).map_err(Error::from)?;
                for (internal_owner_id, uuid) in owner_map_entries {
                    let entry =
                        encode_owner_map_entry(*internal_owner_id, *uuid, &mut budget, map.len())?;
                    let additional = encoded_length_delimited_field_size(1, entry.len())?;
                    budget.charge_work(
                        map.len()
                            .checked_add(additional)
                            .ok_or(Error::InvalidSource)?,
                    )?;
                    budget.charge_allocations(2)?;
                    let append_limits = budget.child_wire_limits()?;
                    appender
                        .length_delimited(&mut map, 1, &entry, append_limits)
                        .map_err(Error::from)?;
                }
                let candidate_work = candidate
                    .len()
                    .checked_add(encoded_length_delimited_field_size(3, map.len())?)
                    .ok_or(Error::InvalidSource)?;
                budget.charge_work(candidate_work)?;
                budget.charge_allocations(2)?;
                let patch_limits = budget.child_wire_limits()?;
                candidate =
                    patch_length_delimited_field(&candidate, 3, false, Some(&map), patch_limits)
                        .map_err(Error::from)?;
            }
            budget.charge_work(candidate.len())?;
            budget.charge_allocations(2)?;
            let patch_limits = budget.child_wire_limits()?;
            patch_varint_field(
                &candidate,
                5,
                tracker.number_of_formulas().is_some(),
                Some(next_count),
                patch_limits,
            )
            .map_err(Error::from)
        },
        child_limits,
    )?;
    budget.charge_work(output.len())?;
    if output.len() > limits.wire.max_output_bytes() {
        return Err(Error::Limit {
            kind: LimitKind::OutputBytes,
            observed: output.len(),
            maximum: limits.wire.max_output_bytes(),
        });
    }
    decode_engine(&output, &mut budget)?;
    Ok(output)
}

fn clone_output(source: &[u8], budget: &mut RewriteBudget, live: usize) -> Result<Vec<u8>, Error> {
    let mut output = budget.admit_output(live, 0, source.len())?;
    output.extend_from_slice(source);
    Ok(output)
}

fn encoded_length_delimited_field_size(
    field_number: u32,
    payload_len: usize,
) -> Result<usize, Error> {
    let key = (u64::from(field_number) << 3) | 2;
    let payload_len = u64::try_from(payload_len).map_err(|_| Error::InvalidSource)?;
    litchi_iwa_common::varint::encoded_len(key)
        .checked_add(litchi_iwa_common::varint::encoded_len(payload_len))
        .and_then(|value| value.checked_add(payload_len as usize))
        .ok_or(Error::InvalidSource)
}

fn encoded_varint_field_size(field_number: u32, value: u64) -> Result<usize, Error> {
    let key = u64::from(field_number) << 3;
    litchi_iwa_common::varint::encoded_len(key)
        .checked_add(litchi_iwa_common::varint::encoded_len(value))
        .ok_or(Error::InvalidSource)
}

fn encoded_owner_map_field_size(
    internal_owner_id: u32,
    uuid: FormulaOwnerUuid,
) -> Result<usize, Error> {
    let cfuuid_len =
        uuid.cfuuid_words()
            .into_iter()
            .enumerate()
            .try_fold(0usize, |total, (field, word)| {
                total
                    .checked_add(encoded_varint_field_size(
                        u32::try_from(field + 2).unwrap_or(5),
                        u64::from(word),
                    )?)
                    .ok_or(Error::InvalidSource)
            })?;
    let entry_len = encoded_varint_field_size(1, u64::from(internal_owner_id))?
        .checked_add(encoded_length_delimited_field_size(2, cfuuid_len)?)
        .ok_or(Error::InvalidSource)?;
    encoded_length_delimited_field_size(1, entry_len)
}

fn decode_owner<'a>(
    source: &'a [u8],
    budget: &mut RewriteBudget,
) -> Result<dependency::FormulaOwnerDependenciesSnapshot<'a>, Error> {
    let (snapshot, report) = dependency::decode_formula_owner_dependencies_with_report(
        source,
        budget.decode_options(source)?,
    )
    .map_err(map_decode_error)?;
    budget.charge_decode_report(report)?;
    Ok(snapshot)
}

fn decode_cell_tile(
    source: &[u8],
    budget: &mut RewriteBudget,
) -> Result<dependency::CellRecordTileSnapshot, Error> {
    let (snapshot, report) =
        dependency::decode_cell_record_tile_with_report(source, budget.decode_options(source)?)
            .map_err(map_decode_error)?;
    budget.charge_decode_report(report)?;
    Ok(snapshot)
}

fn decode_range_tile(
    source: &[u8],
    budget: &mut RewriteBudget,
) -> Result<dependency::RangePrecedentsTileSnapshot, Error> {
    let (snapshot, report) = dependency::decode_range_precedents_tile_with_report(
        source,
        budget.decode_options(source)?,
    )
    .map_err(map_decode_error)?;
    budget.charge_decode_report(report)?;
    Ok(snapshot)
}

fn decode_engine<'a>(
    source: &'a [u8],
    budget: &mut RewriteBudget,
) -> Result<dependency::CalculationEngineSnapshot<'a>, Error> {
    let (snapshot, report) =
        dependency::decode_calculation_engine_with_report(source, budget.decode_options(source)?)
            .map_err(map_decode_error)?;
    budget.charge_decode_report(report)?;
    Ok(snapshot)
}

fn decode_tracker<'a>(
    source: &'a [u8],
    budget: &mut RewriteBudget,
) -> Result<dependency::DependencyTrackerSnapshot<'a>, Error> {
    let (snapshot, report) =
        dependency::decode_dependency_tracker_with_report(source, budget.decode_options(source)?)
            .map_err(map_decode_error)?;
    budget.charge_decode_report(report)?;
    Ok(snapshot)
}

fn encode_reference(
    identifier: u64,
    budget: &mut RewriteBudget,
    live: usize,
) -> Result<Vec<u8>, Error> {
    let mut output = Vec::new();
    let output_capacity = encoded_varint_field_size(1, identifier)?;
    budget.admit_scratch(live, output_capacity)?;
    budget.charge_allocation()?;
    output
        .try_reserve_exact(output_capacity)
        .map_err(|_| Error::Allocation {
            requested: output_capacity,
        })?;
    budget.charge_allocations(2)?;
    let append_limits = budget.child_wire_limits()?;
    append_varint_field(&mut output, 1, identifier, append_limits).map_err(Error::from)?;
    budget.charge_work(output.len())?;
    Ok(output)
}

fn encode_owner_map_entry(
    internal_owner_id: u32,
    uuid: FormulaOwnerUuid,
    budget: &mut RewriteBudget,
    live: usize,
) -> Result<Vec<u8>, Error> {
    let mut cfuuid = Vec::new();
    budget.admit_scratch(live, 64)?;
    budget.charge_allocation()?;
    for (field, word) in uuid.cfuuid_words().into_iter().enumerate() {
        budget.charge_allocations(2)?;
        let append_limits = budget.child_wire_limits()?;
        append_varint_field(
            &mut cfuuid,
            u32::try_from(field + 2).unwrap_or(5),
            u64::from(word),
            append_limits,
        )
        .map_err(Error::from)?;
    }
    let mut output = Vec::new();
    budget.charge_allocation()?;
    let output_capacity = cfuuid.len().checked_add(16).ok_or(Error::Allocation {
        requested: usize::MAX,
    })?;
    budget.admit_scratch(live, output_capacity)?;
    output
        .try_reserve_exact(output_capacity)
        .map_err(|_| Error::Allocation {
            requested: output_capacity,
        })?;
    budget.charge_allocations(2)?;
    let append_limits = budget.child_wire_limits()?;
    append_varint_field(&mut output, 1, u64::from(internal_owner_id), append_limits)
        .map_err(Error::from)?;
    budget.charge_allocations(2)?;
    let append_limits = budget.child_wire_limits()?;
    append_length_delimited_field(&mut output, 2, &cfuuid, append_limits).map_err(Error::from)?;
    budget.charge_work(cfuuid.len().checked_add(output.len()).ok_or(Error::Limit {
        kind: LimitKind::WorkBytes,
        observed: usize::MAX,
        maximum: budget.limits.wire.max_rewrite_work(),
    })?)?;
    Ok(output)
}

fn validate_append_inputs(
    object_ids: &[u64],
    owner_map_entries: &[(u32, FormulaOwnerUuid)],
    budget: &mut RewriteBudget,
) -> Result<(), Error> {
    if object_ids.len() != owner_map_entries.len()
        || object_ids.contains(&0)
        || owner_map_entries
            .iter()
            .any(|(identifier, uuid)| *identifier == 0 || (uuid.lower() == 0 && uuid.upper() == 0))
    {
        return Err(Error::InvalidRemap);
    }
    let mut objects = Vec::new();
    let object_bytes = object_ids
        .len()
        .checked_mul(size_of::<u64>())
        .ok_or(Error::Allocation {
            requested: usize::MAX,
        })?;
    let bytes = owner_map_entries
        .len()
        .checked_mul(size_of::<u32>() + size_of::<(u64, u64)>())
        .ok_or(Error::Allocation {
            requested: usize::MAX,
        })?;
    let total_bytes = object_bytes.checked_add(bytes).ok_or(Error::Allocation {
        requested: usize::MAX,
    })?;
    budget.admit_scratch(0, total_bytes)?;
    budget.charge_allocations(3)?;
    objects
        .try_reserve_exact(object_ids.len())
        .map_err(|_| Error::Allocation {
            requested: object_ids.len(),
        })?;
    objects.extend_from_slice(object_ids);
    objects.sort_unstable();
    if objects.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(Error::InvalidRemap);
    }

    let mut owners = Vec::new();
    let mut uuids = Vec::new();
    owners
        .try_reserve_exact(owner_map_entries.len())
        .map_err(|_| Error::Allocation {
            requested: owner_map_entries.len(),
        })?;
    uuids
        .try_reserve_exact(owner_map_entries.len())
        .map_err(|_| Error::Allocation {
            requested: owner_map_entries.len(),
        })?;
    for (internal, uuid) in owner_map_entries {
        owners.push(*internal);
        uuids.push(uuid_key(*uuid));
    }
    owners.sort_unstable();
    uuids.sort_unstable();
    if owners.windows(2).any(|pair| pair[0] == pair[1])
        || uuids.windows(2).any(|pair| pair[0] == pair[1])
    {
        return Err(Error::InvalidRemap);
    }
    budget.charge_work(total_bytes)
}

/// Validate the opaque `OwnerIDMapArchive` before appending to a tracker.
/// Unknown fields are rejected so an unparsed nested envelope cannot hide an
/// owner identity that would escape the append remap.
#[derive(Debug)]
struct ValidatedOwnerMap {
    identities: Vec<(u32, FormulaOwnerUuid)>,
    uuid_keys: Vec<(u64, u64)>,
    unregistered: Vec<u32>,
}

fn validate_owner_id_map(
    source: Option<&[u8]>,
    budget: &mut RewriteBudget,
) -> Result<ValidatedOwnerMap, Error> {
    let Some(source) = source else {
        return Ok(ValidatedOwnerMap {
            identities: Vec::new(),
            uuid_keys: Vec::new(),
            unregistered: Vec::new(),
        });
    };
    let fields = budget.parse_fields(source, 0)?;
    let mut identities = Vec::new();
    let mut uuid_keys = Vec::new();
    let mut unregistered = Vec::new();
    let field_capacity = fields.len();
    let fields_bytes =
        field_capacity
            .checked_mul(size_of::<WireField>())
            .ok_or(Error::Allocation {
                requested: usize::MAX,
            })?;
    let identity_bytes = field_capacity
        .checked_mul(
            size_of::<(u32, FormulaOwnerUuid)>()
                .checked_add(size_of::<(u64, u64)>())
                .and_then(|value| value.checked_add(size_of::<u32>()))
                .ok_or(Error::Allocation {
                    requested: usize::MAX,
                })?,
        )
        .ok_or(Error::Allocation {
            requested: usize::MAX,
        })?;
    let retained_bytes = fields_bytes
        .checked_add(identity_bytes)
        .ok_or(Error::Allocation {
            requested: usize::MAX,
        })?;
    budget.admit_scratch(source.len(), retained_bytes)?;
    budget.charge_allocations(3)?;
    identities
        .try_reserve_exact(field_capacity)
        .map_err(|_| Error::Allocation {
            requested: field_capacity,
        })?;
    uuid_keys
        .try_reserve_exact(field_capacity)
        .map_err(|_| Error::Allocation {
            requested: field_capacity,
        })?;
    unregistered
        .try_reserve_exact(field_capacity)
        .map_err(|_| Error::Allocation {
            requested: field_capacity,
        })?;
    for field in fields {
        match field.number() {
            1 => {
                if field.wire_type() != 2 {
                    return Err(Error::InvalidSource);
                }
                field
                    .validate_canonical_framing(source)
                    .map_err(Error::from)?;
                let entry_source = field.payload(source).map_err(Error::from)?;
                let entry_live = source
                    .len()
                    .checked_add(fields_bytes)
                    .and_then(|value| value.checked_add(identity_bytes))
                    .ok_or(Error::Allocation {
                        requested: usize::MAX,
                    })?;
                let entry_fields = budget.parse_fields(entry_source, entry_live)?;
                let mut internal = None;
                let mut uuid = None;
                for entry_field in entry_fields {
                    match entry_field.number() {
                        1 if entry_field.wire_type() == 0 && internal.is_none() => {
                            let value = canonical_varint_from_field(entry_source, entry_field)?;
                            internal = Some(
                                u32::try_from(value)
                                    .ok()
                                    .filter(|value| *value != 0)
                                    .ok_or(Error::InvalidSource)?,
                            );
                        },
                        2 if entry_field.wire_type() == 2 && uuid.is_none() => {
                            entry_field
                                .validate_canonical_framing(entry_source)
                                .map_err(Error::from)?;
                            uuid = Some(parse_canonical_cfuuid_with_budget(
                                entry_field.payload(entry_source).map_err(Error::from)?,
                                budget,
                            )?);
                        },
                        _ => return Err(Error::UnsupportedDependency),
                    }
                }
                let internal = internal.ok_or(Error::InvalidSource)?;
                let uuid = uuid.ok_or(Error::InvalidSource)?;
                identities.push((internal, uuid));
                uuid_keys.push(uuid_key(uuid));
            },
            2 if field.wire_type() == 0 => {
                let value = u32::try_from(canonical_varint_from_field(source, field)?)
                    .ok()
                    .filter(|value| *value != 0)
                    .ok_or(Error::InvalidSource)?;
                unregistered.push(value);
            },
            _ => return Err(Error::UnsupportedDependency),
        }
    }
    uuid_keys.sort_unstable();
    unregistered.sort_unstable();
    identities.sort_unstable_by_key(|(internal, _)| *internal);
    if identities.windows(2).any(|pair| pair[0].0 == pair[1].0)
        || uuid_keys.windows(2).any(|pair| pair[0] == pair[1])
        || unregistered.windows(2).any(|pair| pair[0] == pair[1])
        || identities
            .iter()
            .any(|(value, _)| unregistered.binary_search(value).is_ok())
    {
        return Err(Error::InvalidSource);
    }
    Ok(ValidatedOwnerMap {
        identities,
        uuid_keys,
        unregistered,
    })
}

fn validate_tracker_owner_references(
    source: &[u8],
    appended: &[u64],
    budget: &mut RewriteBudget,
) -> Result<(), Error> {
    if appended.is_empty() {
        return Ok(());
    }
    let fields = budget.parse_fields(source, 0)?;
    let mut existing = Vec::new();
    let reference_capacity = fields.iter().filter(|field| field.number() == 6).count();
    let fields_bytes =
        fields
            .len()
            .checked_mul(size_of::<WireField>())
            .ok_or(Error::Allocation {
                requested: usize::MAX,
            })?;
    let existing_bytes =
        reference_capacity
            .checked_mul(size_of::<u64>())
            .ok_or(Error::Allocation {
                requested: usize::MAX,
            })?;
    budget.admit_scratch(
        source.len(),
        fields_bytes
            .checked_add(existing_bytes)
            .ok_or(Error::Allocation {
                requested: usize::MAX,
            })?,
    )?;
    budget.charge_allocation()?;
    existing
        .try_reserve_exact(reference_capacity)
        .map_err(|_| Error::Allocation {
            requested: reference_capacity,
        })?;
    for field in fields {
        if field.number() != 6 {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        let reference = field.payload(source).map_err(Error::from)?;
        let reference_fields = budget.parse_fields(
            reference,
            source
                .len()
                .checked_add(fields_bytes)
                .and_then(|value| value.checked_add(existing_bytes))
                .ok_or(Error::Allocation {
                    requested: usize::MAX,
                })?,
        )?;
        validate_reference_fields(&reference_fields)?;
        let identifier = required_varint_from_fields(reference, &reference_fields, 1)?;
        if identifier == 0 {
            return Err(Error::InvalidSource);
        }
        existing.push(identifier);
    }
    existing.sort_unstable();
    if existing.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(Error::InvalidSource);
    }
    for identifier in appended {
        if existing.binary_search(identifier).is_ok() {
            return Err(Error::InvalidRemap);
        }
    }
    Ok(())
}

fn map_decode_error(error: dependency::DecodeError) -> Error {
    match error.resource_limit() {
        Some(dependency::DecodeLimit::Bytes { observed, maximum }) => Error::Limit {
            kind: LimitKind::InputBytes,
            observed,
            maximum,
        },
        Some(dependency::DecodeLimit::Fields { observed, maximum }) => Error::Limit {
            kind: LimitKind::Fields,
            observed,
            maximum,
        },
        Some(dependency::DecodeLimit::Work { observed, maximum }) => Error::Limit {
            kind: LimitKind::WorkBytes,
            observed,
            maximum,
        },
        Some(dependency::DecodeLimit::Nesting { observed, maximum }) => Error::Limit {
            kind: LimitKind::Nesting,
            observed: observed as usize,
            maximum: maximum as usize,
        },
        Some(dependency::DecodeLimit::References { observed, maximum }) => Error::Limit {
            kind: LimitKind::References,
            observed,
            maximum,
        },
        Some(dependency::DecodeLimit::Text { observed, maximum }) => Error::Limit {
            kind: LimitKind::TextBytes,
            observed,
            maximum,
        },
        Some(dependency::DecodeLimit::Retained { observed, maximum }) => Error::Limit {
            kind: LimitKind::OutputBytes,
            observed,
            maximum,
        },
        Some(dependency::DecodeLimit::Allocation { requested }) => Error::Allocation { requested },
        Some(_) => Error::InvalidSource,
        None => Error::InvalidSource,
    }
}

fn validate_uuid_message(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let _ = validate_typed_fields(source, &[(1, 0), (2, 0)], &[1, 2], budget)?;
    Ok(())
}

fn validate_reference_message(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let _ = validate_typed_fields(source, &[(1, 0), (2, 0), (3, 0)], &[1], budget)?;
    Ok(())
}

fn validate_reference_fields(fields: &[WireField]) -> Result<(), Error> {
    if fields
        .iter()
        .any(|field| !matches!(field.number(), 1..=3) || field.wire_type() != 0)
    {
        return Err(Error::UnsupportedDependency);
    }
    Ok(())
}

fn validate_typed_fields(
    source: &[u8],
    allowed: &[(u32, u8)],
    required: &[u32],
    budget: &mut RewriteBudget,
) -> Result<Vec<WireField>, Error> {
    let fields = budget.parse_fields(source, source.len())?;
    for field in &fields {
        let Some((_, wire_type)) = allowed.iter().find(|(number, _)| *number == field.number())
        else {
            return Err(Error::UnsupportedDependency);
        };
        if field.wire_type() != *wire_type {
            return Err(Error::InvalidSource);
        }
    }
    for number in required {
        if !fields.iter().any(|field| field.number() == *number) {
            return Err(Error::InvalidSource);
        }
    }
    Ok(fields)
}

fn validate_cfuuid_message(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let _ = parse_cfuuid_with_budget(source, budget)?;
    Ok(())
}

fn validate_cell_coordinate(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let _ = validate_typed_fields(source, &[(1, 5), (2, 0), (3, 0)], &[], budget)?;
    Ok(())
}

fn validate_range_coordinate(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let _ = validate_typed_fields(
        source,
        &[(1, 0), (2, 0), (3, 0), (4, 0)],
        &[1, 2, 3, 4],
        budget,
    )?;
    Ok(())
}

fn validate_column_row_size(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let _ = validate_typed_fields(source, &[(1, 0), (2, 0)], &[], budget)?;
    Ok(())
}

fn validate_index_set(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let fields = validate_typed_fields(source, &[(1, 2)], &[], budget)?;
    for field in fields {
        validate_index_set_entry(field.payload(source).map_err(Error::from)?, budget)?;
    }
    Ok(())
}

fn validate_index_set_entry(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let _ = validate_typed_fields(source, &[(1, 0), (2, 0)], &[1], budget)?;
    Ok(())
}

fn validate_cell_coord_set(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let fields = validate_typed_fields(source, &[(1, 2)], &[], budget)?;
    for field in fields {
        let column = field.payload(source).map_err(Error::from)?;
        let column_fields = validate_typed_fields(column, &[(1, 0), (2, 2)], &[1, 2], budget)?;
        let row_set = column_fields
            .iter()
            .find(|nested| nested.number() == 2)
            .ok_or(Error::InvalidSource)?
            .payload(column)
            .map_err(Error::from)?;
        validate_index_set(row_set, budget)?;
    }
    Ok(())
}

fn validate_internal_cell_ref_set(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let fields = validate_typed_fields(source, &[(1, 2)], &[], budget)?;
    for field in fields {
        let owner = field.payload(source).map_err(Error::from)?;
        let owner_fields = validate_typed_fields(owner, &[(1, 0), (2, 2)], &[1, 2], budget)?;
        let coord_set = owner_fields
            .iter()
            .find(|nested| nested.number() == 2)
            .ok_or(Error::InvalidSource)?
            .payload(owner)
            .map_err(Error::from)?;
        validate_cell_coord_set(coord_set, budget)?;
    }
    Ok(())
}

fn validate_expanded_edges(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let fields = validate_typed_fields(
        source,
        &[(1, 0), (2, 0), (3, 0), (4, 0), (5, 0)],
        &[],
        budget,
    )?;
    for field in fields {
        let _ = required_varint_from_field(source, field)?;
    }
    Ok(())
}

fn validate_cell_record_expanded(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let fields = validate_typed_fields(
        source,
        &[(1, 0), (2, 0), (3, 0), (4, 0), (5, 0), (6, 2)],
        &[1, 2],
        budget,
    )?;
    for field in fields {
        if field.number() == 6 {
            validate_expanded_edges(field.payload(source).map_err(Error::from)?, budget)?;
        }
    }
    Ok(())
}

fn validate_cell_dependencies(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let fields = validate_typed_fields(source, &[(1, 2)], &[], budget)?;
    for field in fields {
        validate_cell_record_expanded(field.payload(source).map_err(Error::from)?, budget)?;
    }
    Ok(())
}

fn validate_range_reference(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let fields = validate_typed_fields(
        source,
        &[(1, 2), (2, 0), (3, 0), (4, 0), (5, 0)],
        &[1, 2, 3, 4, 5],
        budget,
    )?;
    let table_id = fields
        .iter()
        .find(|field| field.number() == 1)
        .ok_or(Error::InvalidSource)?
        .payload(source)
        .map_err(Error::from)?;
    validate_cfuuid_message(table_id, budget)
}

fn validate_internal_range_reference(
    source: &[u8],
    budget: &mut RewriteBudget,
) -> Result<(), Error> {
    let fields = validate_typed_fields(source, &[(1, 0), (2, 2)], &[1, 2], budget)?;
    let range = fields
        .iter()
        .find(|field| field.number() == 2)
        .ok_or(Error::InvalidSource)?
        .payload(source)
        .map_err(Error::from)?;
    validate_range_coordinate(range, budget)
}

fn validate_range_dependencies(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let fields = validate_typed_fields(source, &[(2, 2)], &[], budget)?;
    for field in fields {
        let back = field.payload(source).map_err(Error::from)?;
        let back_fields =
            validate_typed_fields(back, &[(1, 0), (2, 0), (3, 2), (4, 2)], &[1, 2], budget)?;
        for back_field in back_fields {
            match back_field.number() {
                3 => validate_range_reference(
                    back_field.payload(back).map_err(Error::from)?,
                    budget,
                )?,
                4 => validate_internal_range_reference(
                    back_field.payload(back).map_err(Error::from)?,
                    budget,
                )?,
                _ => {},
            }
        }
    }
    Ok(())
}

fn validate_volatile_dependencies(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let fields = validate_typed_fields(
        source,
        &[(1, 2), (2, 2), (3, 2), (4, 2), (5, 2), (7, 2)],
        &[],
        budget,
    )?;
    for field in fields {
        match field.number() {
            1..=5 => validate_cell_coord_set(field.payload(source).map_err(Error::from)?, budget)?,
            7 => {
                validate_internal_cell_ref_set(field.payload(source).map_err(Error::from)?, budget)?
            },
            _ => {},
        }
    }
    Ok(())
}

fn validate_extent_range(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let _ = validate_typed_fields(source, &[(1, 0), (2, 0)], &[1], budget)?;
    Ok(())
}

fn validate_extent_range_with_context(
    source: &[u8],
    budget: &mut RewriteBudget,
) -> Result<(), Error> {
    let fields = validate_typed_fields(source, &[(1, 0), (2, 0), (3, 2)], &[1, 2], budget)?;
    let context = fields
        .iter()
        .find(|field| field.number() == 2)
        .ok_or(Error::InvalidSource)?;
    if !matches!(required_varint_from_field(source, *context)?, 0 | 1) {
        return Err(Error::InvalidSource);
    }
    for field in fields {
        if field.number() == 3 {
            validate_extent_range(field.payload(source).map_err(Error::from)?, budget)?;
        }
    }
    Ok(())
}

fn validate_cell_coord_refers_to_extents(
    source: &[u8],
    budget: &mut RewriteBudget,
) -> Result<(), Error> {
    let fields = validate_typed_fields(source, &[(1, 2), (2, 2)], &[1], budget)?;
    for field in fields {
        match field.number() {
            1 => validate_cell_coordinate(field.payload(source).map_err(Error::from)?, budget)?,
            2 => validate_extent_range_with_context(
                field.payload(source).map_err(Error::from)?,
                budget,
            )?,
            _ => {},
        }
    }
    Ok(())
}

fn validate_spanning_dependencies(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let fields = validate_typed_fields(source, &[(1, 2), (2, 2), (3, 2)], &[], budget)?;
    for field in fields {
        match field.number() {
            1 => validate_cell_coord_refers_to_extents(
                field.payload(source).map_err(Error::from)?,
                budget,
            )?,
            2 | 3 => {
                validate_range_coordinate(field.payload(source).map_err(Error::from)?, budget)?
            },
            _ => {},
        }
    }
    Ok(())
}

fn validate_whole_owner_dependencies(
    source: &[u8],
    budget: &mut RewriteBudget,
) -> Result<(), Error> {
    let fields = validate_typed_fields(source, &[(1, 2)], &[], budget)?;
    for field in fields {
        validate_internal_cell_ref_set(field.payload(source).map_err(Error::from)?, budget)?;
    }
    Ok(())
}

fn validate_cell_errors(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let fields = validate_typed_fields(source, &[(1, 2), (2, 2)], &[], budget)?;
    for field in fields {
        let error = field.payload(source).map_err(Error::from)?;
        let nested = if field.number() == 1 {
            validate_typed_fields(error, &[(1, 2), (2, 0), (3, 2)], &[1, 2], budget)?
        } else {
            validate_typed_fields(error, &[(1, 2), (2, 2), (3, 2), (4, 2)], &[1], budget)?
        };
        for nested_field in nested {
            match (field.number(), nested_field.number()) {
                (1, 1) | (2, 1) => validate_cell_coordinate(
                    nested_field.payload(error).map_err(Error::from)?,
                    budget,
                )?,
                (1, 3) | (2, 3) => validate_internal_cell_reference(
                    nested_field.payload(error).map_err(Error::from)?,
                    budget,
                )?,
                (2, 2) => validate_error_archive(
                    nested_field.payload(error).map_err(Error::from)?,
                    budget,
                )?,
                (2, 4) => validate_warning_archive(
                    nested_field.payload(error).map_err(Error::from)?,
                    budget,
                )?,
                _ => {},
            }
        }
    }
    Ok(())
}

fn validate_error_archive(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let fields = validate_typed_fields(source, &[(1, 0), (2, 2)], &[1], budget)?;
    for field in fields {
        if field.number() == 2 {
            let entry = field.payload(source).map_err(Error::from)?;
            let _ = validate_typed_fields(entry, &[(1, 2), (2, 2), (3, 1)], &[1], budget)?;
        }
    }
    Ok(())
}

fn validate_warning_archive(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let fields = validate_typed_fields(source, &[(1, 0), (2, 2), (3, 2)], &[1], budget)?;
    for field in fields {
        match field.number() {
            2 => {
                let entry = field.payload(source).map_err(Error::from)?;
                let _ = validate_typed_fields(entry, &[(1, 2), (2, 2), (3, 1)], &[1], budget)?;
            },
            3 => validate_range_reference(field.payload(source).map_err(Error::from)?, budget)?,
            _ => {},
        }
    }
    Ok(())
}

fn validate_internal_cell_reference(
    source: &[u8],
    budget: &mut RewriteBudget,
) -> Result<(), Error> {
    let fields = validate_typed_fields(source, &[(1, 2), (2, 0)], &[1, 2], budget)?;
    let coordinate = fields
        .iter()
        .find(|field| field.number() == 1)
        .ok_or(Error::InvalidSource)?
        .payload(source)
        .map_err(Error::from)?;
    validate_cell_coordinate(coordinate, budget)
}

fn validate_uuid_references(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let fields = validate_typed_fields(source, &[(1, 2), (2, 2)], &[], budget)?;
    for field in fields {
        let reference = field.payload(source).map_err(Error::from)?;
        match field.number() {
            1 => {
                let nested = validate_typed_fields(reference, &[(1, 2), (2, 2)], &[1], budget)?;
                for nested_field in nested {
                    match nested_field.number() {
                        1 => validate_uuid_message(
                            nested_field.payload(reference).map_err(Error::from)?,
                            budget,
                        )?,
                        2 => validate_cell_coord_set(
                            nested_field.payload(reference).map_err(Error::from)?,
                            budget,
                        )?,
                        _ => {},
                    }
                }
            },
            2 => {
                let nested = validate_typed_fields(reference, &[(1, 2), (2, 2)], &[1], budget)?;
                for nested_field in nested {
                    match nested_field.number() {
                        1 => validate_uuid_message(
                            nested_field.payload(reference).map_err(Error::from)?,
                            budget,
                        )?,
                        2 => {
                            let child = nested_field.payload(reference).map_err(Error::from)?;
                            let child_fields =
                                validate_typed_fields(child, &[(1, 2), (2, 2)], &[1], budget)?;
                            for child_field in child_fields {
                                match child_field.number() {
                                    1 => validate_uuid_message(
                                        child_field.payload(child).map_err(Error::from)?,
                                        budget,
                                    )?,
                                    2 => validate_cell_coord_set(
                                        child_field.payload(child).map_err(Error::from)?,
                                        budget,
                                    )?,
                                    _ => {},
                                }
                            }
                        },
                        _ => {},
                    }
                }
            },
            _ => {},
        }
    }
    Ok(())
}

fn validate_spill_sizes(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let fields = validate_typed_fields(source, &[(1, 2)], &[], budget)?;
    for field in fields {
        let spill = field.payload(source).map_err(Error::from)?;
        let nested = validate_typed_fields(spill, &[(1, 2), (2, 2)], &[1, 2], budget)?;
        for nested_field in nested {
            match nested_field.number() {
                1 => validate_cell_coordinate(
                    nested_field.payload(spill).map_err(Error::from)?,
                    budget,
                )?,
                2 => validate_column_row_size(
                    nested_field.payload(spill).map_err(Error::from)?,
                    budget,
                )?,
                _ => {},
            }
        }
    }
    Ok(())
}

fn validate_owner_nested_fields(source: &[u8], budget: &mut RewriteBudget) -> Result<(), Error> {
    let root = budget.parse_fields(source, 0)?;
    for field in root {
        if matches!(
            field.number(),
            1 | 4 | 5 | 6 | 7 | 8 | 9 | 10 | 11 | 12 | 13 | 14 | 15 | 16
        ) && field.wire_type() != 2
        {
            return Err(Error::InvalidSource);
        }
        if field.number() == 2 || field.number() == 3 {
            if field.wire_type() != 0 {
                return Err(Error::InvalidSource);
            }
        }
        let payload = match field.wire_type() {
            2 => field.payload(source).map_err(Error::from)?,
            _ => continue,
        };
        match field.number() {
            1 | 12 => validate_uuid_message(payload, budget)?,
            4 => validate_cell_dependencies(payload, budget)?,
            5 => validate_range_dependencies(payload, budget)?,
            6 => validate_volatile_dependencies(payload, budget)?,
            7 | 8 => validate_spanning_dependencies(payload, budget)?,
            9 => validate_whole_owner_dependencies(payload, budget)?,
            10 => validate_cell_errors(payload, budget)?,
            11 => validate_reference_message(payload, budget)?,
            13 | 15 => {
                let nested_fields = validate_typed_fields(payload, &[(1, 2)], &[], budget)?;
                for nested in nested_fields {
                    validate_reference_message(
                        nested.payload(payload).map_err(Error::from)?,
                        budget,
                    )?;
                }
            },
            14 => validate_uuid_references(payload, budget)?,
            16 => validate_spill_sizes(payload, budget)?,
            _ => {},
        }
    }
    Ok(())
}

fn reject_unsupported_dependencies(
    owner: &dependency::FormulaOwnerDependenciesSnapshot<'_>,
    budget: &mut RewriteBudget,
) -> Result<(), Error> {
    let unsupported = [
        (owner.volatile_dependencies(), &[7, 1][..]),
        (owner.spanning_column_dependencies(), &[1][..]),
        (owner.spanning_row_dependencies(), &[1][..]),
        (owner.whole_owner_dependencies(), &[1, 1][..]),
        (owner.cell_errors(), &[1][..]),
        (owner.cell_errors(), &[2][..]),
        (owner.spill_range_sizes(), &[1][..]),
    ];
    for (payload, path) in unsupported {
        if let Some(payload) = payload {
            if has_field_at_path(payload, path, budget)? {
                return Err(Error::UnsupportedDependency);
            }
        }
    }
    Ok(())
}

fn has_field_at_path(
    source: &[u8],
    path: &[u32],
    budget: &mut RewriteBudget,
) -> Result<bool, Error> {
    let Some((&field_number, rest)) = path.split_first() else {
        return Ok(true);
    };
    let fields = budget.parse_fields(source, 0)?;
    for field in fields {
        if field.number() != field_number {
            continue;
        }
        if rest.is_empty() {
            return Ok(true);
        }
        if field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        if has_field_at_path(field.payload(source).map_err(Error::from)?, rest, budget)? {
            return Ok(true);
        }
    }
    Ok(false)
}

fn remap_owner_id_field(
    source: &[u8],
    field_number: u32,
    current: u32,
    map: &[(u32, u32)],
    budget: &mut RewriteBudget,
) -> Result<Vec<u8>, Error> {
    let Some(replacement) = lookup_u32(map, current) else {
        return clone_output(source, budget, source.len());
    };
    patch_required_varint(
        source,
        field_number,
        u64::from(current),
        u64::from(replacement),
        budget,
    )
}

fn patch_required_varint(
    source: &[u8],
    field_number: u32,
    current: u64,
    replacement: u64,
    budget: &mut RewriteBudget,
) -> Result<Vec<u8>, Error> {
    let fields = budget.parse_fields(source, source.len())?;
    let mut selected = None;
    for field in fields
        .iter()
        .copied()
        .filter(|field| field.number() == field_number)
    {
        if selected.is_some() || field.wire_type() != 0 {
            return Err(Error::InvalidSource);
        }
        selected = Some(field);
    }
    let field = selected.ok_or(Error::InvalidSource)?;
    if required_varint_from_field(source, field)? != current {
        return Err(Error::InvalidSource);
    }
    let mut bytes = [0u8; litchi_iwa_common::varint::MAX_BYTES];
    let replacement_bytes =
        litchi_iwa_common::varint::encode_varint_to_buffer(replacement, &mut bytes);
    let output_len = source
        .len()
        .checked_sub(field.end() - field.start())
        .and_then(|value| value.checked_add(field.key_end() - field.start()))
        .and_then(|value| value.checked_add(replacement_bytes.len()))
        .ok_or(Error::InvalidSource)?;
    let mut output = budget.admit_output(
        source.len(),
        fields
            .len()
            .checked_mul(size_of::<WireField>())
            .ok_or(Error::Allocation {
                requested: usize::MAX,
            })?,
        output_len,
    )?;
    output.extend_from_slice(&source[..field.start()]);
    output.extend_from_slice(&source[field.start()..field.key_end()]);
    output.extend_from_slice(replacement_bytes);
    output.extend_from_slice(&source[field.end()..]);
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

fn patch_uuid(
    source: &[u8],
    replacement: FormulaOwnerUuid,
    budget: &mut RewriteBudget,
) -> Result<Vec<u8>, Error> {
    let current_lower = required_varint_with_budget(source, 1, budget)?;
    let output = patch_required_varint(source, 1, current_lower, replacement.lower(), budget)?;
    let current_upper = required_varint_with_budget(&output, 2, budget)?;
    patch_required_varint(&output, 2, current_upper, replacement.upper(), budget)
}

fn patch_cfuuid(
    source: &[u8],
    replacement: FormulaOwnerUuid,
    budget: &mut RewriteBudget,
) -> Result<Vec<u8>, Error> {
    let fields = budget.parse_fields(source, source.len())?;
    let words = replacement.cfuuid_words();
    let mut has_bytes = false;
    let mut has_words = false;
    for field in &fields {
        match field.number() {
            1 if field.wire_type() == 2 => {
                if has_bytes {
                    return Err(Error::InvalidSource);
                }
                has_bytes = true;
            },
            2..=5 if field.wire_type() == 0 => {
                has_words = true;
            },
            _ => return Err(Error::InvalidSource),
        }
    }
    if has_bytes && has_words {
        return Err(Error::InvalidSource);
    }
    let mut output_len = 0usize;
    for field in &fields {
        let field_len = match field.number() {
            1 => (field.key_end() - field.start())
                .checked_add(litchi_iwa_common::varint::encoded_len(16))
                .and_then(|value| value.checked_add(16))
                .ok_or(Error::InvalidSource)?,
            2..=5 => (field.key_end() - field.start())
                .checked_add(litchi_iwa_common::varint::encoded_len(u64::from(
                    words[usize::try_from(field.number() - 2).map_err(|_| Error::InvalidSource)?],
                )))
                .ok_or(Error::InvalidSource)?,
            _ => field.end() - field.start(),
        };
        output_len = output_len
            .checked_add(field_len)
            .ok_or(Error::InvalidSource)?;
    }
    let mut output = budget.admit_output(
        source.len(),
        fields
            .len()
            .checked_mul(size_of::<WireField>())
            .ok_or(Error::Allocation {
                requested: usize::MAX,
            })?,
        output_len,
    )?;
    for field in fields {
        match field.number() {
            1 => {
                output.extend_from_slice(&source[field.start()..field.key_end()]);
                litchi_iwa_common::varint::encode_varint_into(&mut output, 16);
                output.extend_from_slice(&uuid_bytes(replacement));
            },
            2..=5 => {
                let index =
                    usize::try_from(field.number() - 2).map_err(|_| Error::InvalidSource)?;
                output.extend_from_slice(&source[field.start()..field.key_end()]);
                litchi_iwa_common::varint::encode_varint_into(&mut output, u64::from(words[index]));
            },
            _ => output.extend_from_slice(field.raw(source).map_err(Error::from)?),
        }
    }
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

fn parse_uuid_from_fields(source: &[u8], fields: &[WireField]) -> Result<FormulaOwnerUuid, Error> {
    let lower = required_varint_from_fields(source, fields, 1)?;
    let upper = required_varint_from_fields(source, fields, 2)?;
    Ok(FormulaOwnerUuid::new(lower, upper))
}

fn parse_uuid_with_budget(
    source: &[u8],
    budget: &mut RewriteBudget,
) -> Result<FormulaOwnerUuid, Error> {
    let fields = budget.parse_fields(source, source.len())?;
    parse_uuid_from_fields(source, &fields)
}

fn parse_cfuuid_from_fields(
    source: &[u8],
    fields: &[WireField],
) -> Result<FormulaOwnerUuid, Error> {
    let mut bytes = None;
    let mut words = [None; 4];
    for field in fields.iter().copied() {
        match field.number() {
            1 if field.wire_type() == 2 => {
                if bytes.is_some() {
                    return Err(Error::InvalidSource);
                }
                let payload = field.payload(source).map_err(Error::from)?;
                let value: [u8; 16] = payload.try_into().map_err(|_| Error::InvalidSource)?;
                bytes = Some(value);
            },
            2..=5 if field.wire_type() == 0 => {
                let index =
                    usize::try_from(field.number() - 2).map_err(|_| Error::InvalidSource)?;
                if words[index].is_some() {
                    return Err(Error::InvalidSource);
                }
                words[index] = Some(
                    u32::try_from(required_varint_from_field(source, field)?)
                        .map_err(|_| Error::InvalidSource)?,
                );
            },
            _ => return Err(Error::InvalidSource),
        }
    }
    if bytes.is_some() && words.iter().any(Option::is_some) {
        return Err(Error::InvalidSource);
    }
    if words.iter().all(Option::is_some) {
        return Ok(FormulaOwnerUuid::new(
            u64::from(words[0].unwrap_or_default())
                | (u64::from(words[1].unwrap_or_default()) << 32),
            u64::from(words[2].unwrap_or_default())
                | (u64::from(words[3].unwrap_or_default()) << 32),
        ));
    }
    let value = bytes.ok_or(Error::InvalidSource)?;
    let value = u128::from_be_bytes(value);
    Ok(FormulaOwnerUuid::new(value as u64, (value >> 64) as u64))
}

#[cfg(test)]
fn parse_cfuuid(source: &[u8], limits: WireLimits) -> Result<FormulaOwnerUuid, Error> {
    let fields = litchi_iwa_common::wire::parse_wire_fields(source, limits).map_err(Error::from)?;
    parse_cfuuid_from_fields(source, &fields)
}

fn parse_cfuuid_with_budget(
    source: &[u8],
    budget: &mut RewriteBudget,
) -> Result<FormulaOwnerUuid, Error> {
    let fields = budget.parse_fields(source, source.len())?;
    parse_cfuuid_from_fields(source, &fields)
}

fn parse_canonical_cfuuid_with_budget(
    source: &[u8],
    budget: &mut RewriteBudget,
) -> Result<FormulaOwnerUuid, Error> {
    let fields = budget.parse_fields(source, source.len())?;
    for field in fields.iter().copied() {
        field
            .validate_canonical_framing(source)
            .map_err(Error::from)?;
        if field.wire_type() == 0 {
            let _ = canonical_varint_from_field(source, field)?;
        }
    }
    parse_cfuuid_from_fields(source, &fields)
}

fn uuid_bytes(value: FormulaOwnerUuid) -> [u8; 16] {
    (((value.upper() as u128) << 64) | value.lower() as u128).to_be_bytes()
}

fn required_varint_with_budget(
    source: &[u8],
    field_number: u32,
    budget: &mut RewriteBudget,
) -> Result<u64, Error> {
    let fields = budget.parse_fields(source, source.len())?;
    required_varint_from_fields(source, &fields, field_number)
}

fn required_varint_from_fields(
    source: &[u8],
    fields: &[WireField],
    field_number: u32,
) -> Result<u64, Error> {
    let mut value = None;
    for field in fields
        .iter()
        .copied()
        .filter(|field| field.number() == field_number)
    {
        if field.wire_type() != 0 || value.is_some() {
            return Err(Error::InvalidSource);
        }
        value = Some(required_varint_from_field(source, field)?);
    }
    value.ok_or(Error::InvalidSource)
}

fn required_varint_from_field(source: &[u8], field: WireField) -> Result<u64, Error> {
    let payload = field.payload(source).map_err(Error::from)?;
    let (value, consumed) = decode_varint_from_bytes(payload).map_err(|_| Error::InvalidSource)?;
    if consumed != payload.len() {
        return Err(Error::InvalidSource);
    }
    Ok(value)
}

fn canonical_varint_from_field(source: &[u8], field: WireField) -> Result<u64, Error> {
    if field.wire_type() != 0 {
        return Err(Error::InvalidSource);
    }
    field.validate_canonical_key(source).map_err(Error::from)?;
    let value = required_varint_from_field(source, field)?;
    let payload = field.payload(source).map_err(Error::from)?;
    if payload.len() != litchi_iwa_common::varint::encoded_len(value) {
        return Err(Error::InvalidSource);
    }
    Ok(value)
}

fn validate_u64_map(map: &[(u64, u64)], budget: &mut RewriteBudget) -> Result<(), Error> {
    if map
        .iter()
        .any(|(source, target)| *source == 0 || *target == 0)
    {
        return Err(Error::InvalidRemap);
    }
    if map.windows(2).any(|pair| pair[0].0 >= pair[1].0) {
        return Err(Error::InvalidRemap);
    }
    let mut targets = Vec::new();
    let bytes = map
        .len()
        .checked_mul(size_of::<u64>())
        .ok_or(Error::Allocation {
            requested: usize::MAX,
        })?;
    budget.admit_scratch(0, bytes)?;
    budget.charge_allocation()?;
    targets
        .try_reserve_exact(map.len())
        .map_err(|_| Error::Allocation {
            requested: map.len(),
        })?;
    targets.extend(map.iter().map(|(_, target)| *target));
    targets.sort_unstable();
    if targets.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(Error::InvalidRemap);
    }
    budget.charge_work(bytes)?;
    Ok(())
}

fn validate_u32_map(map: &[(u32, u32)], budget: &mut RewriteBudget) -> Result<(), Error> {
    if map
        .iter()
        .any(|(source, target)| *source == 0 || *target == 0)
    {
        return Err(Error::InvalidRemap);
    }
    if map.windows(2).any(|pair| pair[0].0 >= pair[1].0) {
        return Err(Error::InvalidRemap);
    }
    let mut targets = Vec::new();
    let bytes = map
        .len()
        .checked_mul(size_of::<u32>())
        .ok_or(Error::Allocation {
            requested: usize::MAX,
        })?;
    budget.admit_scratch(0, bytes)?;
    budget.charge_allocation()?;
    targets
        .try_reserve_exact(map.len())
        .map_err(|_| Error::Allocation {
            requested: map.len(),
        })?;
    targets.extend(map.iter().map(|(_, target)| *target));
    targets.sort_unstable();
    if targets.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(Error::InvalidRemap);
    }
    budget.charge_work(bytes)?;
    Ok(())
}

fn validate_uuid_map(
    map: &[(FormulaOwnerUuid, FormulaOwnerUuid)],
    budget: &mut RewriteBudget,
) -> Result<(), Error> {
    if map.iter().any(|(source, target)| {
        (source.lower() == 0 && source.upper() == 0) || (target.lower() == 0 && target.upper() == 0)
    }) {
        return Err(Error::InvalidRemap);
    }
    if map
        .windows(2)
        .any(|pair| uuid_key(pair[0].0) >= uuid_key(pair[1].0))
    {
        return Err(Error::InvalidRemap);
    }
    let mut targets = Vec::new();
    let bytes = map
        .len()
        .checked_mul(size_of::<(u64, u64)>())
        .ok_or(Error::Allocation {
            requested: usize::MAX,
        })?;
    budget.admit_scratch(0, bytes)?;
    budget.charge_allocation()?;
    targets
        .try_reserve_exact(map.len())
        .map_err(|_| Error::Allocation {
            requested: map.len(),
        })?;
    targets.extend(map.iter().map(|(_, target)| uuid_key(*target)));
    targets.sort_unstable();
    if targets.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(Error::InvalidRemap);
    }
    budget.charge_work(bytes)?;
    Ok(())
}

fn lookup_u64(map: &[(u64, u64)], source: u64) -> Option<u64> {
    map.binary_search_by_key(&source, |(source, _)| *source)
        .ok()
        .map(|index| map[index].1)
}

fn lookup_u32(map: &[(u32, u32)], source: u32) -> Option<u32> {
    map.binary_search_by_key(&source, |(source, _)| *source)
        .ok()
        .map(|index| map[index].1)
}

fn uuid_key(value: FormulaOwnerUuid) -> (u64, u64) {
    (value.lower(), value.upper())
}

fn lookup_uuid(
    map: &[(FormulaOwnerUuid, FormulaOwnerUuid)],
    source: FormulaOwnerUuid,
) -> Option<(FormulaOwnerUuid, FormulaOwnerUuid)> {
    map.binary_search_by_key(&uuid_key(source), |(source, _)| uuid_key(*source))
        .ok()
        .map(|index| map[index])
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_iwa_common::varint::encode_varint;

    fn decode_options(source: &[u8], limits: Limits) -> dependency::DecodeOptions {
        dependency::DecodeOptions::new(
            source.len(),
            limits.wire.max_fields(),
            limits.wire.max_rewrite_work(),
            u32::try_from(limits.wire.max_nesting()).unwrap_or(u32::MAX),
            limits.max_references,
            0,
        )
    }

    fn varint_field(output: &mut Vec<u8>, field: u32, value: u64) {
        output.extend(encode_varint(u64::from(field) << 3));
        output.extend(encode_varint(value));
    }

    fn bytes_field(output: &mut Vec<u8>, field: u32, value: &[u8]) {
        output.extend(encode_varint((u64::from(field) << 3) | 2));
        output.extend(encode_varint(u64::try_from(value.len()).unwrap_or(0)));
        output.extend_from_slice(value);
    }

    fn overlong_varint(value: u64) -> Vec<u8> {
        let mut output = encode_varint(value);
        let last = output.pop().unwrap_or(0);
        output.push(last | 0x80);
        output.push(0);
        output
    }

    fn uuid(value: FormulaOwnerUuid) -> Vec<u8> {
        let mut output = Vec::new();
        varint_field(&mut output, 1, value.lower());
        varint_field(&mut output, 2, value.upper());
        output
    }

    fn reference(identifier: u64) -> Vec<u8> {
        let mut output = Vec::new();
        varint_field(&mut output, 1, identifier);
        output
    }

    fn expanded_edges(owner: u32) -> Vec<u8> {
        let mut output = Vec::new();
        varint_field(&mut output, 1, 1);
        varint_field(&mut output, 2, 2);
        varint_field(&mut output, 3, 3);
        varint_field(&mut output, 4, 4);
        varint_field(&mut output, 5, u64::from(owner));
        output
    }

    fn owner_source() -> Vec<u8> {
        let table = FormulaOwnerUuid::new(0x101, 0x202);
        let mut owner = Vec::new();
        bytes_field(&mut owner, 1, &uuid(table));
        varint_field(&mut owner, 2, 7);

        let mut record = Vec::new();
        varint_field(&mut record, 1, 1);
        varint_field(&mut record, 2, 2);
        bytes_field(&mut record, 6, &expanded_edges(7));
        let mut cells = Vec::new();
        bytes_field(&mut cells, 1, &record);
        bytes_field(&mut owner, 4, &cells);

        let mut internal_range = Vec::new();
        varint_field(&mut internal_range, 1, 7);
        let mut range = Vec::new();
        varint_field(&mut range, 1, 0);
        varint_field(&mut range, 2, 0);
        varint_field(&mut range, 3, 1);
        varint_field(&mut range, 4, 1);
        bytes_field(&mut internal_range, 2, &range);
        let mut back = Vec::new();
        varint_field(&mut back, 1, 2);
        varint_field(&mut back, 2, 1);
        bytes_field(&mut back, 4, &internal_range);
        let mut ranges = Vec::new();
        bytes_field(&mut ranges, 2, &back);
        bytes_field(&mut owner, 5, &ranges);

        bytes_field(&mut owner, 11, &reference(100));
        bytes_field(&mut owner, 12, &uuid(table));
        let mut tiled = Vec::new();
        bytes_field(&mut tiled, 1, &reference(101));
        bytes_field(&mut owner, 13, &tiled);
        let mut uuid_refs = Vec::new();
        let mut table_ref = Vec::new();
        bytes_field(&mut table_ref, 1, &uuid(table));
        bytes_field(&mut uuid_refs, 1, &table_ref);
        bytes_field(&mut owner, 14, &uuid_refs);
        let mut range_tiles = Vec::new();
        bytes_field(&mut range_tiles, 1, &reference(102));
        bytes_field(&mut owner, 15, &range_tiles);
        varint_field(&mut owner, 97, 970);
        owner
    }

    fn cell_tile_source() -> Vec<u8> {
        let mut record = Vec::new();
        varint_field(&mut record, 1, 1);
        varint_field(&mut record, 2, 2);
        bytes_field(&mut record, 6, &expanded_edges(7));
        let mut tile = Vec::new();
        varint_field(&mut tile, 1, 7);
        varint_field(&mut tile, 2, 0);
        varint_field(&mut tile, 3, 0);
        bytes_field(&mut tile, 4, &record);
        tile
    }

    fn range_tile_source() -> Vec<u8> {
        let mut tile = Vec::new();
        varint_field(&mut tile, 1, 7);
        tile
    }

    fn engine_source() -> Vec<u8> {
        let old = FormulaOwnerUuid::new(0x101, 0x202);
        let mut budget = RewriteBudget::new(Limits::default());
        let entry = encode_owner_map_entry(7, old, &mut budget, 0).unwrap();
        let mut owner_map = Vec::new();
        bytes_field(&mut owner_map, 1, &entry);
        engine_with_owner_map(&owner_map)
    }

    fn engine_with_owner_map(owner_map: &[u8]) -> Vec<u8> {
        let mut tracker = Vec::new();
        bytes_field(&mut tracker, 3, owner_map);
        varint_field(&mut tracker, 5, 1);
        bytes_field(&mut tracker, 6, &reference(100));
        varint_field(&mut tracker, 97, 970);
        let mut engine = Vec::new();
        bytes_field(&mut engine, 2, &tracker);
        varint_field(&mut engine, 98, 980);
        engine
    }

    #[test]
    fn append_reservation_sizes_match_worst_case_wire() {
        let uuid = FormulaOwnerUuid::new(u64::from(u32::MAX), u64::from(u32::MAX));
        let mut budget = RewriteBudget::new(Limits::default());
        let entry = encode_owner_map_entry(u32::MAX, uuid, &mut budget, 0).unwrap();
        let mut encoded_map_field = Vec::new();
        bytes_field(&mut encoded_map_field, 1, &entry);
        assert_eq!(
            encoded_owner_map_field_size(u32::MAX, uuid).unwrap(),
            encoded_map_field.len(),
        );
        assert!(encoded_map_field.len() > 16);

        let reference = encode_reference(u64::MAX, &mut budget, 0).unwrap();
        assert_eq!(
            encoded_varint_field_size(1, u64::MAX).unwrap(),
            reference.len()
        );
    }

    #[test]
    fn remap_rejects_unsorted_target_duplicates() {
        let source = owner_source();
        let _ = source;
        assert_eq!(
            Remap::new(&[(1, 2), (2, 1), (3, 2)], &[], &[]),
            Err(Error::InvalidRemap)
        );
    }

    #[test]
    fn append_rejects_duplicate_input_identities() {
        let source = engine_source();
        let first = FormulaOwnerUuid::new(0x1001, 0x2002);
        let second = FormulaOwnerUuid::new(0x1002, 0x2003);
        let limits = Limits::default();
        assert_eq!(
            append_owners_to_engine(
                &source,
                &[101, 101],
                &[(17, first), (18, second)],
                2,
                &limits,
            ),
            Err(Error::InvalidRemap)
        );
        assert_eq!(
            append_owners_to_engine(
                &source,
                &[101, 102],
                &[(17, first), (17, second)],
                2,
                &limits,
            ),
            Err(Error::InvalidRemap)
        );
        assert_eq!(
            append_owners_to_engine(
                &source,
                &[101, 102],
                &[(17, first), (18, first)],
                2,
                &limits,
            ),
            Err(Error::InvalidRemap)
        );
    }

    #[test]
    fn append_rejects_identity_already_in_existing_owner_map() {
        let source = engine_source();
        let limits = Limits::default();
        assert_eq!(
            append_owners_to_engine(
                &source,
                &[101],
                &[(7, FormulaOwnerUuid::new(0x1001, 0x2002))],
                1,
                &limits,
            ),
            Err(Error::InvalidRemap)
        );
    }

    #[test]
    fn append_rejects_uuid_collision_when_owner_map_is_internal_sorted() {
        let high = FormulaOwnerUuid::new(0x9000, 0x1);
        let low = FormulaOwnerUuid::new(0x1000, 0x1);
        let mut budget = RewriteBudget::new(Limits::default());
        let first = encode_owner_map_entry(7, high, &mut budget, 0).unwrap();
        let second = encode_owner_map_entry(8, low, &mut budget, 0).unwrap();
        let mut owner_map = Vec::new();
        bytes_field(&mut owner_map, 1, &first);
        bytes_field(&mut owner_map, 1, &second);
        let mut tracker = Vec::new();
        bytes_field(&mut tracker, 3, &owner_map);
        varint_field(&mut tracker, 5, 2);
        let mut engine = Vec::new();
        bytes_field(&mut engine, 2, &tracker);
        assert_eq!(
            append_owners_to_engine(&engine, &[101], &[(17, low)], 1, &Limits::default(),),
            Err(Error::InvalidRemap)
        );
    }

    #[test]
    fn cfuuid_rejects_conflicting_bytes_and_words() {
        let value = FormulaOwnerUuid::new(0x101, 0x202);
        let mut source = Vec::new();
        bytes_field(&mut source, 1, &uuid_bytes(value));
        varint_field(&mut source, 2, u64::from(value.cfuuid_words()[0]));
        assert_eq!(
            parse_cfuuid(&source, WireLimits::default()),
            Err(Error::InvalidSource)
        );
    }

    #[test]
    fn append_rejects_unknown_owner_map_fields() {
        let old = FormulaOwnerUuid::new(0x101, 0x202);
        let mut budget = RewriteBudget::new(Limits::default());
        let entry = encode_owner_map_entry(7, old, &mut budget, 0).unwrap();
        let mut owner_map = Vec::new();
        bytes_field(&mut owner_map, 1, &entry);
        varint_field(&mut owner_map, 99, 1);
        let mut tracker = Vec::new();
        bytes_field(&mut tracker, 3, &owner_map);
        varint_field(&mut tracker, 5, 1);
        let mut engine = Vec::new();
        bytes_field(&mut engine, 2, &tracker);
        assert_eq!(
            append_owners_to_engine(
                &engine,
                &[101],
                &[(17, FormulaOwnerUuid::new(0x1001, 0x2002))],
                1,
                &Limits::default(),
            ),
            Err(Error::UnsupportedDependency)
        );
    }

    #[test]
    fn append_rejects_noncanonical_owner_map_framing_and_values() {
        let old = FormulaOwnerUuid::new(0x101, 0x202);
        let new = FormulaOwnerUuid::new(0x1001, 0x2002);
        let mut budget = RewriteBudget::new(Limits::default());
        let canonical_entry = encode_owner_map_entry(7, old, &mut budget, 0).unwrap();

        let mut root_key = Vec::new();
        root_key.extend([0x8a, 0x00]);
        root_key.extend(encode_varint(canonical_entry.len() as u64));
        root_key.extend_from_slice(&canonical_entry);

        let mut root_length = vec![0x0a];
        root_length.extend(overlong_varint(canonical_entry.len() as u64));
        root_length.extend_from_slice(&canonical_entry);

        let mut entry_key = vec![0x88, 0x00];
        entry_key.extend(encode_varint(7));
        let mut cfuuid = Vec::new();
        for (field, word) in old.cfuuid_words().into_iter().enumerate() {
            varint_field(
                &mut cfuuid,
                u32::try_from(field + 2).unwrap_or(5),
                u64::from(word),
            );
        }
        bytes_field(&mut entry_key, 2, &cfuuid);
        let mut nested_key = Vec::new();
        bytes_field(&mut nested_key, 1, &entry_key);

        let mut entry_length = Vec::new();
        varint_field(&mut entry_length, 1, 7);
        entry_length.extend([0x12]);
        entry_length.extend(overlong_varint(cfuuid.len() as u64));
        entry_length.extend_from_slice(&cfuuid);
        let mut nested_length = Vec::new();
        bytes_field(&mut nested_length, 1, &entry_length);

        let mut entry_value = Vec::new();
        entry_value.extend([0x08]);
        entry_value.extend(overlong_varint(7));
        bytes_field(&mut entry_value, 2, &cfuuid);
        let mut nested_value = Vec::new();
        bytes_field(&mut nested_value, 1, &entry_value);

        let mut cfuuid_value = Vec::new();
        for (field, word) in old.cfuuid_words().into_iter().enumerate() {
            let field = u32::try_from(field + 2).unwrap_or(5);
            if field == 2 {
                cfuuid_value.extend(encode_varint(u64::from(field) << 3));
                cfuuid_value.extend(overlong_varint(u64::from(word)));
            } else {
                varint_field(&mut cfuuid_value, field, u64::from(word));
            }
        }
        let mut entry_cfuuid_value = Vec::new();
        varint_field(&mut entry_cfuuid_value, 1, 7);
        bytes_field(&mut entry_cfuuid_value, 2, &cfuuid_value);
        let mut nested_cfuuid_value = Vec::new();
        bytes_field(&mut nested_cfuuid_value, 1, &entry_cfuuid_value);

        let cases = [
            ("root key", root_key),
            ("root length", root_length),
            ("entry key", nested_key),
            ("entry length", nested_length),
            ("entry value", nested_value),
            ("CFUUID value", nested_cfuuid_value),
        ];
        for (label, owner_map) in cases {
            let source = engine_with_owner_map(&owner_map);
            assert_eq!(
                append_owners_to_engine(&source, &[101], &[(17, new)], 1, &Limits::default(),),
                Err(Error::InvalidSource),
                "accepted noncanonical owner-map {label}",
            );
        }
    }

    #[test]
    fn unknown_nested_owner_fields_are_rejected() {
        let table = FormulaOwnerUuid::new(0x101, 0x202);
        let mut source = Vec::new();
        bytes_field(&mut source, 1, &uuid(table));
        varint_field(&mut source, 2, 7);
        let mut nested = Vec::new();
        varint_field(&mut nested, 99, 1);
        bytes_field(&mut source, 4, &nested);
        let owner_ids = [(7, 17)];
        assert_eq!(
            remap_owner(
                &source,
                Remap::new(&[], &owner_ids, &[]).unwrap(),
                &Limits::default(),
            ),
            Err(Error::UnsupportedDependency)
        );
    }

    #[test]
    fn owner_rewrite_requires_source_owned_identity_maps() {
        let source = owner_source();
        let old = FormulaOwnerUuid::new(0x101, 0x202);
        let root_uuid = [(old, old)];
        let root_owner = [(7, 7)];
        let object = [(100, 100)];
        assert_eq!(
            remap_owner(
                &source,
                Remap::new(&object, &root_owner, &[]).unwrap(),
                &Limits::default(),
            ),
            Err(Error::InvalidRemap)
        );
        assert_eq!(
            remap_owner(
                &source,
                Remap::new(&object, &[], &root_uuid).unwrap(),
                &Limits::default(),
            ),
            Err(Error::InvalidRemap)
        );
        assert_eq!(
            remap_owner(
                &source,
                Remap::new(&[], &root_owner, &root_uuid).unwrap(),
                &Limits::default(),
            ),
            Err(Error::InvalidRemap)
        );
    }

    #[test]
    fn owner_rewrite_rejects_family_specific_unknown_fields() {
        let old = FormulaOwnerUuid::new(0x101, 0x202);
        let new = FormulaOwnerUuid::new(0x1001, 0x2002);
        let owner_ids = [(7, 17)];
        let uuids = [(old, new)];
        let remap = Remap::new(&[], &owner_ids, &uuids).unwrap();
        for family in [6_u32, 7, 8] {
            let mut source = Vec::new();
            bytes_field(&mut source, 1, &uuid(old));
            varint_field(&mut source, 2, 7);
            let mut nested = Vec::new();
            varint_field(&mut nested, 99, 1);
            bytes_field(&mut source, family, &nested);
            assert_eq!(
                remap_owner(&source, remap, &Limits::default()),
                Err(Error::UnsupportedDependency),
                "family {family} accepted an unknown nested field",
            );
        }
    }

    #[test]
    fn owner_rewrite_remaps_family_and_keeps_unknown_tail() {
        let source = owner_source();
        let old = FormulaOwnerUuid::new(0x101, 0x202);
        let new = FormulaOwnerUuid::new(0x1001, 0x2002);
        let object_ids = [(100, 200), (101, 201), (102, 202)];
        let owner_ids = [(7, 17)];
        let uuids = [(old, new)];
        let rewritten = remap_owner(
            &source,
            Remap::new(&object_ids, &owner_ids, &uuids).unwrap(),
            &Limits::default(),
        )
        .unwrap();
        assert!(rewritten.ends_with(&[0x88, 0x06, 0xca, 0x07]));
        let owner = dependency::decode_formula_owner_dependencies(
            &rewritten,
            decode_options(&rewritten, Limits::default()),
        )
        .unwrap();
        assert_eq!(owner.formula_owner_uid().lower(), new.lower());
        assert_eq!(owner.formula_owner_uid().upper(), new.upper());
        assert_eq!(owner.internal_formula_owner_id(), 17);

        let reverse_object_ids = [(200, 100), (201, 101), (202, 102)];
        let reverse_owner_ids = [(17, 7)];
        let reverse_uuids = [(new, old)];
        assert_eq!(
            remap_owner(
                &rewritten,
                Remap::new(&reverse_object_ids, &reverse_owner_ids, &reverse_uuids,).unwrap(),
                &Limits::default(),
            )
            .unwrap(),
            source
        );
    }

    #[test]
    fn identity_map_is_required_to_be_source_sorted() {
        let object_ids = [(101, 201), (100, 200)];
        assert_eq!(Remap::new(&object_ids, &[], &[]), Err(Error::InvalidRemap));
    }

    #[test]
    fn unchanged_source_is_exact_when_no_maps_are_selected() {
        let source = owner_source();
        let owner_uuid = FormulaOwnerUuid::new(0x101, 0x202);
        let object_ids = [(100, 100), (101, 101), (102, 102)];
        let owner_ids = [(7, 7)];
        let uuids = [(owner_uuid, owner_uuid)];
        let rewritten = remap_owner(
            &source,
            Remap::new(&object_ids, &owner_ids, &uuids).unwrap(),
            &Limits::default(),
        )
        .unwrap();
        assert_eq!(rewritten, source);
    }

    #[test]
    fn output_limit_is_admitted_before_rewrite() {
        let source = owner_source();
        let limits = Limits {
            wire: WireLimits::default()
                .with_output_bytes(source.len() - 1)
                .unwrap(),
            ..Limits::default()
        };
        assert!(matches!(
            remap_owner(&source, Remap::new(&[], &[], &[]).unwrap(), &limits),
            Err(Error::Limit {
                kind: LimitKind::OutputBytes,
                ..
            })
        ));
    }

    #[test]
    fn scratch_and_allocation_limits_are_admitted_before_rewrite_buffers() {
        let source = owner_source();
        let scratch_limited = Limits::default().with_scratch_bytes(1).unwrap();
        assert!(matches!(
            remap_owner(
                &source,
                Remap::new(&[], &[], &[]).unwrap(),
                &scratch_limited,
            ),
            Err(Error::Limit {
                kind: LimitKind::ScratchBytes,
                ..
            })
        ));
        let allocation_limited = Limits::default().with_allocations(1).unwrap();
        assert!(matches!(
            remap_owner(
                &source,
                Remap::new(&[], &[], &[]).unwrap(),
                &allocation_limited,
            ),
            Err(Error::Limit {
                kind: LimitKind::Allocations,
                ..
            })
        ));
    }

    #[test]
    fn dependency_tiles_remap_root_and_expanded_owner_edges() {
        let owner_ids = [(7, 17)];
        let remap = Remap::new(&[], &owner_ids, &[]).unwrap();
        let limits = Limits::default();

        let cell_source = cell_tile_source();
        let cell_rewritten = remap_cell_record_tile(&cell_source, remap, &limits).unwrap();
        let cell = dependency::decode_cell_record_tile(
            &cell_rewritten,
            decode_options(&cell_rewritten, limits),
        )
        .unwrap();
        assert_eq!(cell.internal_owner_id(), 17);
        assert!(
            cell_rewritten
                .windows(2)
                .any(|window| window == [0x28, 0x11])
        );

        let range_source = range_tile_source();
        let range_rewritten = remap_range_precedents_tile(&range_source, remap, &limits).unwrap();
        let range = dependency::decode_range_precedents_tile(
            &range_rewritten,
            decode_options(&range_rewritten, limits),
        )
        .unwrap();
        assert_eq!(range.to_owner_id(), 17);

        let reverse = Remap::new(&[], &[(17, 7)], &[]).unwrap();
        assert_eq!(
            remap_cell_record_tile(&cell_rewritten, reverse, &limits).unwrap(),
            cell_source
        );
        assert_eq!(
            remap_range_precedents_tile(&range_rewritten, reverse, &limits).unwrap(),
            range_source
        );
    }

    #[test]
    fn repeated_owner_edges_admit_exact_replacement_scratch() {
        const RULES: &[(&[u32], Leaf)] = &[(&[6], Leaf::EdgeIds)];
        let mut nested = Vec::new();
        for _ in 0..32 {
            varint_field(&mut nested, 5, 7);
        }
        let mut source = Vec::new();
        bytes_field(&mut source, 6, &nested);
        let owner_ids = [(7, 17)];
        let remap = Remap::new(&[], &owner_ids, &[]).unwrap();
        let run = |scratch: usize| {
            let limits = Limits::default().with_scratch_bytes(scratch).unwrap();
            rewrite_paths(&source, RULES, remap, &mut RewriteBudget::new(limits))
        };
        let (mut lower, mut upper) = (1usize, Limits::MAX_SCRATCH_BYTES);
        while lower < upper {
            let middle = lower + (upper - lower) / 2;
            if run(middle).is_ok() {
                upper = middle;
            } else {
                lower = middle + 1;
            }
        }
        // The replacement vector of 32 owner identifiers is admitted.
        assert!(lower >= 32 * size_of::<u64>());
        let rewritten = run(lower).expect("exact replacement scratch should admit");
        let mut expected_nested = Vec::new();
        for _ in 0..32 {
            varint_field(&mut expected_nested, 5, 17);
        }
        let mut expected = Vec::new();
        bytes_field(&mut expected, 6, &expected_nested);
        assert_eq!(rewritten, expected);
        assert!(matches!(
            run(lower - 1),
            Err(Error::Limit {
                kind: LimitKind::ScratchBytes,
                ..
            })
        ));
    }

    #[test]
    fn engine_append_preserves_tracker_and_root_unknown_fields() {
        let source = engine_source();
        let limits = Limits::default();
        let new = FormulaOwnerUuid::new(0x1001, 0x2002);
        let rewritten = append_owners_to_engine(&source, &[101], &[(17, new)], 1, &limits).unwrap();
        assert!(rewritten.ends_with(&[0x90, 0x06, 0xd4, 0x07]));
        let mut budget = RewriteBudget::new(limits);
        let engine = decode_engine(&rewritten, &mut budget).unwrap();
        let tracker = decode_tracker(engine.dependency_tracker(), &mut budget).unwrap();
        assert_eq!(tracker.number_of_formulas(), Some(2));
        assert!(
            engine
                .dependency_tracker()
                .windows(4)
                .any(|window| window == [0x88, 0x06, 0xca, 0x07])
        );
        let owner_map = tracker.owner_id_map().unwrap();
        let new_entry = encode_owner_map_entry(17, new, &mut budget, 0).unwrap();
        assert!(
            owner_map
                .windows(new_entry.len())
                .any(|window| window == new_entry.as_slice())
        );

        assert_eq!(
            append_owners_to_engine(&source, &[], &[], 0, &limits).unwrap(),
            source
        );
    }

    #[test]
    fn sixteen_owner_append_has_exact_scratch_and_allocation_boundaries() {
        let source = engine_source();
        let object_ids: Vec<u64> = (0_u64..16).map(|offset| 101 + offset).collect();
        let owner_map_entries: Vec<(u32, FormulaOwnerUuid)> = (0_u32..16)
            .map(|offset| {
                (
                    17 + offset,
                    FormulaOwnerUuid::new(0x1001 + u64::from(offset), 0x2001 + u64::from(offset)),
                )
            })
            .collect();

        fn minimum_admitting_limit(
            maximum: usize,
            mut succeeds: impl FnMut(usize) -> bool,
        ) -> usize {
            assert!(succeeds(maximum));
            let mut lower = 0;
            let mut upper = maximum;
            while lower < upper {
                let middle = lower + (upper - lower) / 2;
                if succeeds(middle) {
                    upper = middle;
                } else {
                    lower = middle + 1;
                }
            }
            lower
        }

        let scratch_required = minimum_admitting_limit(Limits::MAX_SCRATCH_BYTES, |value| {
            let limits = Limits {
                max_scratch_bytes: value,
                ..Limits::default()
            };
            append_owners_to_engine(&source, &object_ids, &owner_map_entries, 16, &limits).is_ok()
        });
        assert!(scratch_required > 0);
        let exact_scratch = Limits {
            max_scratch_bytes: scratch_required,
            ..Limits::default()
        };
        assert!(
            append_owners_to_engine(&source, &object_ids, &owner_map_entries, 16, &exact_scratch,)
                .is_ok()
        );
        let mut below_scratch = exact_scratch;
        below_scratch.max_scratch_bytes = scratch_required - 1;
        assert!(matches!(
            append_owners_to_engine(&source, &object_ids, &owner_map_entries, 16, &below_scratch,),
            Err(Error::Limit {
                kind: LimitKind::ScratchBytes,
                ..
            })
        ));

        let allocations_required = minimum_admitting_limit(Limits::MAX_ALLOCATIONS, |value| {
            let limits = Limits {
                max_allocations: value,
                ..Limits::default()
            };
            append_owners_to_engine(&source, &object_ids, &owner_map_entries, 16, &limits).is_ok()
        });
        assert!(allocations_required > 0);
        let exact_allocations = Limits {
            max_allocations: allocations_required,
            ..Limits::default()
        };
        assert!(
            append_owners_to_engine(
                &source,
                &object_ids,
                &owner_map_entries,
                16,
                &exact_allocations,
            )
            .is_ok()
        );
        let mut below_allocations = exact_allocations;
        below_allocations.max_allocations = allocations_required - 1;
        assert!(matches!(
            append_owners_to_engine(
                &source,
                &object_ids,
                &owner_map_entries,
                16,
                &below_allocations,
            ),
            Err(Error::Limit {
                kind: LimitKind::Allocations,
                ..
            })
        ));
    }

    #[test]
    fn next_owner_id_clears_owner_map_entries_without_owner_objects() {
        let mut budget = RewriteBudget::new(Limits::default());
        let first =
            encode_owner_map_entry(7, FormulaOwnerUuid::new(0x101, 0x202), &mut budget, 0).unwrap();
        let orphan =
            encode_owner_map_entry(20, FormulaOwnerUuid::new(0x303, 0x404), &mut budget, 0)
                .unwrap();
        let mut owner_map = Vec::new();
        bytes_field(&mut owner_map, 1, &first);
        bytes_field(&mut owner_map, 1, &orphan);
        let engine = engine_with_owner_map(&owner_map);
        let owner = owner_source();
        // The owner objects alone would give 8; the orphan map entry wins.
        assert_eq!(
            next_owner_id(&engine, &[&owner], &Limits::default()),
            Ok(21)
        );
        assert_eq!(
            next_owner_id(&engine_source(), &[&owner], &Limits::default()),
            Ok(8)
        );
        assert_eq!(
            next_owner_id(&engine, &[&[0xff][..]], &Limits::default()),
            Err(Error::InvalidSource)
        );
    }
}
