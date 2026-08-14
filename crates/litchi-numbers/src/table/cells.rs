//! Presence-preserving semantic Numbers cell reads and exact-source batches.
//!
//! Start with [`crate::Package::table_cell`] for one checked coordinate, or
//! [`crate::Package::table_cells`] for a checked dense range. Both are
//! selector-first: select a sheet with [`crate::SheetSelector`], then a table
//! on that sheet with [`crate::TableSelector`]. Name selectors use exact
//! visible names and reject malformed duplicate matches instead of silently
//! choosing one. Index selectors are checked zero-based source positions.
//!
//! [`State`](crate::table::cells::State) preserves the requested coordinate
//! and its native materialization state.
//! [`Storage::Missing`](crate::table::cells::Storage::Missing) is distinct
//! from [`Storage::Stored`](crate::table::cells::Storage::Stored) containing
//! [`Value::Empty`](crate::cell::Value::Empty), so a caller never needs to
//! infer presence from an empty-looking value. Stored values may also be the
//! semantic formula or error values already represented by
//! [`Value`](crate::cell::Value). Changed batches accept finite scalar values,
//! bounded semantic formulas with optional typed caches, or explicit clears.
//! Error-value construction stays outside this boundary.
//!
//! A range is half-open and returned in row-major order. It is deliberately
//! dense: every addressed coordinate yields a
//! [`State`](crate::table::cells::State), including missing cells. Bounds are
//! checked against the selected table before allocation, and the requested
//! element count is bounded by the package's semantic materialized-cell
//! limit. An
//! over-limit, out-of-bounds, ambiguous, or malformed source fails as a typed
//! read error without returning a partial range.
//!
//! The reader is grounded in an Apple-authored Numbers workbook with
//! materialized text and number cells. Native identifiers, package members,
//! protobuf/wire records, and source bytes remain private to the package
//! reader. Mutation entry points retain exact source and target artifacts,
//! expose no native identifier, and refuse unsupported physical or formula
//! dependencies before publication.

use std::{fmt, sync::Arc};

use litchi_iwa_archive::package::{OwnedExactArtifacts, SharedBytes};

use crate::{
    Package,
    cell::{FiniteF64, FiniteF64Error, Type as ValueType, Value},
    formula::{CachedValue, Expression, Table},
};

use super::{CellPosition, CoordinateError, Dimensions};

pub use crate::package::table_cells::{DependencyKind, Error, LimitKind, Path};

/// One supported scalar or bounded semantic formula staged for a Numbers cell.
#[derive(Clone, PartialEq)]
#[non_exhaustive]
pub enum Input {
    /// User-entered text.
    Text(String),
    /// A finite numeric scalar.
    Number(FiniteF64),
    /// A Boolean scalar.
    Boolean(bool),
    /// A finite Apple-epoch date in seconds.
    Date(FiniteF64),
    /// A finite duration in seconds.
    Duration(FiniteF64),
    /// A bounded semantic formula and optional typed display cache.
    Formula {
        /// The opaque expression to compile at the package boundary.
        expression: Expression,
        /// An optional caller-supplied typed display cache.
        cached: Option<CachedValue>,
    },
}

impl Input {
    /// Fallibly copy user text into one exact-size buffer.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Allocation`] if the buffer cannot be reserved.
    pub fn text(value: impl AsRef<str>) -> Result<Self, Error> {
        let value = value.as_ref();
        let mut owned = String::new();
        owned
            .try_reserve_exact(value.len())
            .map_err(|_error| Error::Allocation {
                kind: LimitKind::OwnedValueBytes,
                amount: value.len(),
            })?;
        owned.push_str(value);
        Ok(Self::Text(owned))
    }

    /// Construct a finite numeric scalar.
    pub fn number(value: f64) -> Result<Self, FiniteF64Error> {
        FiniteF64::new(value).map(Self::Number)
    }

    /// Construct a Boolean scalar.
    #[must_use]
    pub const fn boolean(value: bool) -> Self {
        Self::Boolean(value)
    }

    /// Construct a finite Apple-epoch date in seconds.
    pub fn date(value: f64) -> Result<Self, FiniteF64Error> {
        FiniteF64::new(value).map(Self::Date)
    }

    /// Construct a finite duration in seconds.
    pub fn duration(value: f64) -> Result<Self, FiniteF64Error> {
        FiniteF64::new(value).map(Self::Duration)
    }

    /// Construct an authored formula without a supplied display cache.
    #[must_use]
    pub const fn formula(expression: Expression) -> Self {
        Self::Formula {
            expression,
            cached: None,
        }
    }

    /// Construct an authored formula with a typed display cache.
    #[must_use]
    pub const fn formula_cached(expression: Expression, cached: CachedValue) -> Self {
        Self::Formula {
            expression,
            cached: Some(cached),
        }
    }

    /// Return the scalar kind without exposing its content.
    #[must_use]
    pub const fn value_type(&self) -> ValueType {
        match self {
            Self::Text(_) => ValueType::Text,
            Self::Number(_) => ValueType::Number,
            Self::Boolean(_) => ValueType::Boolean,
            Self::Date(_) => ValueType::Date,
            Self::Duration(_) => ValueType::Duration,
            Self::Formula { .. } => ValueType::Formula,
        }
    }

    pub(crate) fn owned_bytes(&self) -> usize {
        match self {
            Self::Text(value) => value.len(),
            Self::Number(_) | Self::Boolean(_) | Self::Date(_) | Self::Duration(_) => 0,
            Self::Formula { expression, cached } => expression
                .owned_bytes()
                .saturating_add(cached.as_ref().map_or(0, CachedValue::owned_bytes)),
        }
    }

    pub(crate) const fn formula_parts(&self) -> Option<(&Expression, Option<&CachedValue>)> {
        match self {
            Self::Formula { expression, cached } => Some((expression, cached.as_ref())),
            Self::Text(_)
            | Self::Number(_)
            | Self::Boolean(_)
            | Self::Date(_)
            | Self::Duration(_) => None,
        }
    }

    pub(crate) fn matches_value(&self, value: &Value) -> bool {
        match (self, value) {
            (Self::Text(left), Value::Text(right)) => left == right,
            (Self::Number(left), Value::Number(right))
            | (Self::Date(left), Value::Date(right))
            | (Self::Duration(left), Value::Duration(right)) => left == right,
            (Self::Boolean(left), Value::Boolean(right)) => left == right,
            (Self::Formula { .. }, _) => false,
            _ => false,
        }
    }

    pub(crate) fn into_scalar_value(self) -> Option<Value> {
        match self {
            Self::Text(value) => Some(Value::Text(value)),
            Self::Number(value) => Some(Value::Number(value)),
            Self::Boolean(value) => Some(Value::Boolean(value)),
            Self::Date(value) => Some(Value::Date(value)),
            Self::Duration(value) => Some(Value::Duration(value)),
            Self::Formula { .. } => None,
        }
    }
}

impl fmt::Debug for Input {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Input")
            .field("kind", &self.value_type())
            .finish_non_exhaustive()
    }
}

/// One explicit cell replacement or clear.
#[derive(Clone, PartialEq)]
pub struct Change {
    position: CellPosition,
    input: Option<Input>,
}

impl Change {
    /// Stage a scalar replacement at a checked zero-based position.
    #[must_use]
    pub const fn set(position: CellPosition, input: Input) -> Self {
        Self {
            position,
            input: Some(input),
        }
    }

    /// Parse an A1 coordinate and stage a scalar replacement.
    pub fn set_a1(address: &str, input: Input) -> Result<Self, Error> {
        let position = CellPosition::from_a1(address).map_err(map_coordinate_error)?;
        Ok(Self::set(position, input))
    }

    /// Stage a bounded formula at a checked zero-based position.
    #[must_use]
    pub const fn set_formula(position: CellPosition, expression: Expression) -> Self {
        Self::set(position, Input::formula(expression))
    }

    /// Parse an A1 coordinate and stage a bounded formula.
    pub fn set_formula_a1(address: &str, expression: Expression) -> Result<Self, Error> {
        let position = CellPosition::from_a1(address).map_err(map_coordinate_error)?;
        Ok(Self::set_formula(position, expression))
    }

    /// Stage a bounded formula and typed display cache.
    #[must_use]
    pub const fn set_formula_cached(
        position: CellPosition,
        expression: Expression,
        cached: CachedValue,
    ) -> Self {
        Self::set(position, Input::formula_cached(expression, cached))
    }

    /// Parse an A1 coordinate and stage a bounded formula and display cache.
    pub fn set_formula_cached_a1(
        address: &str,
        expression: Expression,
        cached: CachedValue,
    ) -> Result<Self, Error> {
        let position = CellPosition::from_a1(address).map_err(map_coordinate_error)?;
        Ok(Self::set_formula_cached(position, expression, cached))
    }

    /// Stage an explicit clear at a checked zero-based position.
    #[must_use]
    pub const fn clear(position: CellPosition) -> Self {
        Self {
            position,
            input: None,
        }
    }

    /// Parse an A1 coordinate and stage an explicit clear.
    pub fn clear_a1(address: &str) -> Result<Self, Error> {
        let position = CellPosition::from_a1(address).map_err(map_coordinate_error)?;
        Ok(Self::clear(position))
    }

    /// Return the semantic target coordinate.
    #[must_use]
    pub const fn position(&self) -> CellPosition {
        self.position
    }

    /// Borrow the staged input, or return `None` for a clear.
    #[must_use]
    pub const fn input(&self) -> Option<&Input> {
        self.input.as_ref()
    }

    /// Return whether this change clears its coordinate.
    #[must_use]
    pub const fn is_clear(&self) -> bool {
        self.input.is_none()
    }

    pub(crate) const fn input_ref(&self) -> Option<&Input> {
        self.input.as_ref()
    }

    pub(crate) fn into_parts(self) -> (CellPosition, Option<Input>) {
        (self.position, self.input)
    }
}

impl fmt::Debug for Change {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Change")
            .field("position", &self.position)
            .field("operation", &if self.is_clear() { "clear" } else { "set" })
            .finish()
    }
}

/// Physical presence and semantic value at one requested coordinate.
#[derive(Clone, PartialEq)]
#[non_exhaustive]
pub enum Storage {
    /// No native cell is materialized at the coordinate.
    Missing,
    /// A native cell is materialized, including an explicitly stored
    /// [`Value::Empty`], formula, or error value.
    Stored(Value),
}

impl Storage {
    /// Return whether no native cell is materialized.
    #[must_use]
    pub const fn is_missing(&self) -> bool {
        matches!(self, Self::Missing)
    }

    /// Borrow the stored semantic value, if one is materialized.
    #[must_use]
    pub const fn value(&self) -> Option<&Value> {
        match self {
            Self::Missing => None,
            Self::Stored(value) => Some(value),
        }
    }
}

impl fmt::Debug for Storage {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Missing => formatter.write_str("Storage::Missing"),
            Self::Stored(value) => {
                write!(formatter, "Storage::Stored({})", value.cell_type().name())
            },
        }
    }
}

/// One presence-preserving semantic cell result.
#[derive(Clone, PartialEq)]
pub struct State {
    position: CellPosition,
    storage: Storage,
}

impl State {
    /// Return the requested semantic coordinate.
    #[must_use]
    pub const fn position(&self) -> CellPosition {
        self.position
    }

    /// Borrow the coordinate's presence-preserving storage state.
    #[must_use]
    pub const fn storage(&self) -> &Storage {
        &self.storage
    }

    pub(crate) const fn new(position: CellPosition, storage: Storage) -> Self {
        Self { position, storage }
    }
}

impl fmt::Debug for State {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("State")
            .field("position", &self.position)
            .field("storage", &self.storage)
            .finish()
    }
}

type CommitFn = for<'source> fn(Plan<'source>) -> Result<Commit, Error>;

/// One immutable selector-first cell batch under construction.
pub struct Edit<'source> {
    source: &'source Package,
    path: Path,
    dimensions: Dimensions,
    maximum_updates: usize,
    maximum_owned_bytes: usize,
    changes: Vec<Change>,
    owned_bytes: usize,
    formula_nodes: usize,
    change_allocation_events: usize,
    commit: CommitFn,
}

impl<'source> Edit<'source> {
    pub(crate) fn from_resolved(
        source: &'source Package,
        path: Path,
        dimensions: Dimensions,
        maximum_updates: usize,
        maximum_owned_bytes: usize,
        commit: CommitFn,
    ) -> Result<Self, Error> {
        Ok(Self {
            source,
            path,
            dimensions,
            maximum_updates,
            maximum_owned_bytes,
            changes: Vec::new(),
            owned_bytes: 0,
            formula_nodes: 0,
            change_allocation_events: 0,
            commit,
        })
    }

    /// Return the content-free semantic target.
    #[must_use]
    pub const fn path(&self) -> Path {
        self.path
    }

    /// Return the number of staged coordinates.
    #[must_use]
    pub const fn len(&self) -> usize {
        self.changes.len()
    }

    /// Return whether no coordinate has been staged.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.changes.is_empty()
    }

    /// Select an opaque cross-table formula handle from this exact source.
    ///
    /// The returned handle contains no native table identifier and is rejected
    /// if used with an edit from any independently opened or changed package.
    pub fn formula_table<'sheet, 'table>(
        &self,
        sheet: impl Into<crate::SheetSelector<'sheet>>,
        table: impl Into<crate::TableSelector<'table>>,
    ) -> Result<Table, Error> {
        crate::formula::resolve_table_handle(self.source, sheet, table)
    }

    /// Stage one scalar replacement.
    pub fn set(self, position: CellPosition, input: Input) -> Result<Self, Error> {
        self.change(Change::set(position, input))
    }

    /// Parse an A1 coordinate and stage one scalar replacement.
    pub fn set_a1(self, address: &str, input: Input) -> Result<Self, Error> {
        self.change(Change::set_a1(address, input)?)
    }

    /// Stage one bounded formula replacement.
    pub fn set_formula(
        self,
        position: CellPosition,
        expression: Expression,
    ) -> Result<Self, Error> {
        self.change(Change::set_formula(position, expression))
    }

    /// Parse an A1 coordinate and stage one bounded formula replacement.
    pub fn set_formula_a1(self, address: &str, expression: Expression) -> Result<Self, Error> {
        self.change(Change::set_formula_a1(address, expression)?)
    }

    /// Stage one bounded formula replacement with a typed display cache.
    pub fn set_formula_cached(
        self,
        position: CellPosition,
        expression: Expression,
        cached: CachedValue,
    ) -> Result<Self, Error> {
        self.change(Change::set_formula_cached(position, expression, cached))
    }

    /// Parse an A1 coordinate and stage one formula replacement with cache.
    pub fn set_formula_cached_a1(
        self,
        address: &str,
        expression: Expression,
        cached: CachedValue,
    ) -> Result<Self, Error> {
        self.change(Change::set_formula_cached_a1(address, expression, cached)?)
    }

    /// Stage one explicit clear.
    pub fn clear(self, position: CellPosition) -> Result<Self, Error> {
        self.change(Change::clear(position))
    }

    /// Parse an A1 coordinate and stage one explicit clear.
    pub fn clear_a1(self, address: &str) -> Result<Self, Error> {
        self.change(Change::clear_a1(address)?)
    }

    /// Stage one checked change.
    pub fn change(mut self, change: Change) -> Result<Self, Error> {
        let position = change.position();
        if position.row() >= self.dimensions.rows()
            || position.column() >= self.dimensions.columns()
        {
            return Err(Error::OutOfBounds {
                position,
                dimensions: self.dimensions,
            });
        }
        let observed = checked_staged_total(
            self.changes.len(),
            1,
            self.maximum_updates,
            LimitKind::Updates,
            self.path,
        )?;
        let formula_nodes = checked_staged_total(
            self.formula_nodes,
            change
                .input()
                .and_then(Input::formula_parts)
                .map_or(0, |(expression, _cached)| expression.node_count()),
            crate::formula::MAX_NODES,
            LimitKind::FormulaWork,
            self.path,
        )?;
        let owned_bytes = checked_staged_total(
            self.owned_bytes,
            change.input().map_or(0, Input::owned_bytes),
            self.maximum_owned_bytes,
            LimitKind::OwnedValueBytes,
            self.path,
        )?;
        if let Some((expression, _cached)) = change.input().and_then(Input::formula_parts) {
            expression.validate_for(self.source, self.dimensions)?;
        }
        let allocation_events =
            reserve_change_slot(&mut self.changes, self.change_allocation_events, self.path)?;
        self.changes.push(change);
        self.owned_bytes = owned_bytes;
        self.formula_nodes = formula_nodes;
        self.change_allocation_events = allocation_events;
        debug_assert_eq!(self.changes.len(), observed);
        Ok(self)
    }

    /// Stage a bounded iterator of changes without publishing partial work.
    pub fn extend(mut self, changes: impl IntoIterator<Item = Change>) -> Result<Self, Error> {
        for change in changes {
            self = self.change(change)?;
        }
        Ok(self)
    }

    /// Validate and atomically publish the final batch.
    pub fn commit(mut self) -> Result<Commit, Error> {
        self.changes.sort_unstable_by_key(Change::position);
        if let Some(position) = self.changes.windows(2).find_map(|pair| {
            (pair[0].position() == pair[1].position()).then_some(pair[0].position())
        }) {
            return Err(Error::DuplicatePosition { position });
        }
        (self.commit)(Plan {
            source: self.source,
            path: self.path,
            dimensions: self.dimensions,
            staging_usage: StagingUsage {
                change_capacity: self.changes.capacity(),
                allocation_events: self.change_allocation_events,
                formula_nodes: self.formula_nodes,
            },
            changes: self.changes,
            owned_bytes: self.owned_bytes,
        })
    }
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("path", &self.path)
            .field("changes", &self.changes.len())
            .finish_non_exhaustive()
    }
}

pub(crate) struct Plan<'source> {
    source: &'source Package,
    path: Path,
    dimensions: Dimensions,
    staging_usage: StagingUsage,
    changes: Vec<Change>,
    owned_bytes: usize,
}

impl<'source> Plan<'source> {
    pub(crate) fn into_parts(
        self,
    ) -> (
        &'source Package,
        Path,
        Dimensions,
        Vec<Change>,
        usize,
        StagingUsage,
    ) {
        (
            self.source,
            self.path,
            self.dimensions,
            self.changes,
            self.owned_bytes,
            self.staging_usage,
        )
    }
}

/// Actual private allocation retained by the public staging facade.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) struct StagingUsage {
    change_capacity: usize,
    allocation_events: usize,
    formula_nodes: usize,
}

impl StagingUsage {
    pub(crate) const fn change_capacity(self) -> usize {
        self.change_capacity
    }

    pub(crate) const fn allocation_events(self) -> usize {
        self.allocation_events
    }

    pub(crate) const fn formula_nodes(self) -> usize {
        self.formula_nodes
    }
}

/// One physical message coordinate retained by an exact directional patch.
///
/// This stays crate-private: component, object, and message indexes never
/// cross the semantic facade.
#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(crate) struct PhysicalLocation {
    pub(crate) component: usize,
    pub(crate) object: usize,
    pub(crate) message: usize,
}

/// The topology operation represented by one directional evidence item.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum EvidenceChangeKind {
    Replace,
    Append,
    Delete,
}

impl EvidenceChangeKind {
    const fn inverse(self) -> Self {
        match self {
            Self::Replace => Self::Replace,
            Self::Append => Self::Delete,
            Self::Delete => Self::Append,
        }
    }
}

/// One compact range into a patch-owned reference table.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) struct ReferenceSpan {
    start: usize,
    len: usize,
}

impl ReferenceSpan {
    pub(crate) const fn new(start: usize, len: usize) -> Self {
        Self { start, len }
    }

    fn end(self) -> Option<usize> {
        self.start.checked_add(self.len)
    }
}

/// Exact object-reference state for one `FieldInfo` ordinal.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) struct FieldReferenceRoute {
    field_index: usize,
    source: ReferenceSpan,
    target: ReferenceSpan,
}

impl FieldReferenceRoute {
    pub(crate) const fn new(
        field_index: usize,
        source: ReferenceSpan,
        target: ReferenceSpan,
    ) -> Self {
        Self {
            field_index,
            source,
            target,
        }
    }
}

/// Exact aggregate and complete field-reference state for one message.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) struct MessageReferenceRoute {
    source: ReferenceSpan,
    target: ReferenceSpan,
    fields: ReferenceSpan,
}

impl MessageReferenceRoute {
    pub(crate) const fn new(
        source: ReferenceSpan,
        target: ReferenceSpan,
        fields: ReferenceSpan,
    ) -> Self {
        Self {
            source,
            target,
            fields,
        }
    }
}

/// One flattened, patch-owned exact object-reference table.
///
/// Data references are intentionally absent: locality always requires exact
/// preservation of aggregate and every field-local data-reference list.
#[derive(Clone, PartialEq, Eq)]
pub(crate) struct ReferenceEvidence {
    routes: Arc<Vec<MessageReferenceRoute>>,
    fields: Arc<Vec<FieldReferenceRoute>>,
    object_references: Arc<Vec<u64>>,
}

impl ReferenceEvidence {
    /// Retain caller-built exact reference tables without allocating.
    pub(crate) fn new(
        routes: Arc<Vec<MessageReferenceRoute>>,
        fields: Arc<Vec<FieldReferenceRoute>>,
        object_references: Arc<Vec<u64>>,
        path: Path,
    ) -> Result<Self, Error> {
        let references_valid =
            |span: ReferenceSpan| span.end().is_some_and(|end| end <= object_references.len());
        let fields_valid = |span: ReferenceSpan| span.end().is_some_and(|end| end <= fields.len());
        let mut next_field = 0usize;
        for route in routes.iter() {
            if !references_valid(route.source)
                || !references_valid(route.target)
                || !fields_valid(route.fields)
                || route.fields.start != next_field
            {
                return Err(Error::InvalidSource { path });
            }
            let end = route.fields.end().ok_or(Error::InvalidSource { path })?;
            for (expected_index, field) in fields[route.fields.start..end].iter().enumerate() {
                if field.field_index != expected_index
                    || !references_valid(field.source)
                    || !references_valid(field.target)
                {
                    return Err(Error::InvalidSource { path });
                }
            }
            next_field = end;
        }
        if next_field != fields.len() || object_references.contains(&0) {
            return Err(Error::InvalidSource { path });
        }
        Ok(Self {
            routes,
            fields,
            object_references,
        })
    }

    pub(crate) fn allocation_shapes(&self) -> ((usize, usize), (usize, usize), (usize, usize)) {
        (
            (self.routes.len(), self.routes.capacity()),
            (self.fields.len(), self.fields.capacity()),
            (
                self.object_references.len(),
                self.object_references.capacity(),
            ),
        )
    }
}

/// Whether one changed message preserves or exactly transitions references.
#[derive(Clone, Copy, PartialEq, Eq)]
enum ReferenceTransition {
    Preserve,
    Exact { route: usize },
}

/// Exact source/target coordinates for one published native message.
///
/// `object_identifier` is the stable object identity shared by both
/// directions. The optional locations model replacement, append, and delete
/// without sentinels or a second allocation.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) struct DirectionalMessage {
    source: Option<PhysicalLocation>,
    target: Option<PhysicalLocation>,
    object_identifier: u64,
    expected_type: u32,
    kind: EvidenceChangeKind,
    references: ReferenceTransition,
}

impl DirectionalMessage {
    pub(crate) const fn new(
        source: Option<PhysicalLocation>,
        target: Option<PhysicalLocation>,
        object_identifier: u64,
        expected_type: u32,
        kind: EvidenceChangeKind,
    ) -> Self {
        Self {
            source,
            target,
            object_identifier,
            expected_type,
            kind,
            references: ReferenceTransition::Preserve,
        }
    }

    pub(crate) const fn with_reference_transition(mut self, route: usize) -> Self {
        self.references = ReferenceTransition::Exact { route };
        self
    }

    const fn inverse(self) -> Self {
        Self {
            source: self.target,
            target: self.source,
            object_identifier: self.object_identifier,
            expected_type: self.expected_type,
            kind: self.kind.inverse(),
            references: self.references,
        }
    }

    pub(crate) const fn source(self) -> Option<PhysicalLocation> {
        self.source
    }

    pub(crate) const fn target(self) -> Option<PhysicalLocation> {
        self.target
    }

    pub(crate) const fn object_identifier(self) -> u64 {
        self.object_identifier
    }

    pub(crate) const fn expected_type(self) -> u32 {
        self.expected_type
    }

    pub(crate) const fn kind(self) -> EvidenceChangeKind {
        self.kind
    }
}

/// One direction-normalized exact field-reference transition.
#[derive(Clone, Copy)]
pub(crate) struct DirectionalFieldReference<'evidence> {
    field_index: usize,
    source: &'evidence [u64],
    target: &'evidence [u64],
}

impl<'evidence> DirectionalFieldReference<'evidence> {
    pub(crate) const fn field_index(self) -> usize {
        self.field_index
    }

    pub(crate) const fn source(self) -> &'evidence [u64] {
        self.source
    }

    pub(crate) const fn target(self) -> &'evidence [u64] {
        self.target
    }
}

/// Borrowed exact object-reference transition in the patch's direction.
#[derive(Clone, Copy)]
pub(crate) struct DirectionalReferenceTransition<'evidence> {
    evidence: &'evidence ReferenceEvidence,
    route: MessageReferenceRoute,
    inverse: bool,
}

impl PartialEq for DirectionalReferenceTransition<'_> {
    fn eq(&self, other: &Self) -> bool {
        std::ptr::eq(self.evidence, other.evidence)
            && self.route == other.route
            && self.inverse == other.inverse
    }
}

impl Eq for DirectionalReferenceTransition<'_> {}

impl fmt::Debug for DirectionalReferenceTransition<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("DirectionalReferenceTransition")
            .field("fields", &self.route.fields.len)
            .field("inverse", &self.inverse)
            .finish_non_exhaustive()
    }
}

impl<'evidence> DirectionalReferenceTransition<'evidence> {
    pub(crate) fn source(&self) -> &'evidence [u64] {
        reference_slice(
            &self.evidence.object_references,
            if self.inverse {
                self.route.target
            } else {
                self.route.source
            },
        )
    }

    pub(crate) fn target(&self) -> &'evidence [u64] {
        reference_slice(
            &self.evidence.object_references,
            if self.inverse {
                self.route.source
            } else {
                self.route.target
            },
        )
    }

    pub(crate) fn fields(
        &self,
    ) -> impl ExactSizeIterator<Item = DirectionalFieldReference<'evidence>> + '_ {
        let inverse = self.inverse;
        let identifiers = self.evidence.object_references.as_slice();
        reference_fields(self.evidence, self.route)
            .iter()
            .map(move |field| DirectionalFieldReference {
                field_index: field.field_index,
                source: reference_slice(
                    identifiers,
                    if inverse { field.target } else { field.source },
                ),
                target: reference_slice(
                    identifiers,
                    if inverse { field.source } else { field.target },
                ),
            })
    }
}

/// Compact directional evidence retained for changed-patch apply.
#[derive(Clone, PartialEq, Eq)]
pub(crate) struct PatchEvidence {
    messages: Option<Arc<Vec<DirectionalMessage>>>,
    references: Option<ReferenceEvidence>,
    inverse: bool,
    preview_mask: u8,
    source_previews: usize,
    target_previews: usize,
}

impl PatchEvidence {
    /// Retain prebuilt directional evidence without allocating.
    ///
    /// The package owner constructs and charges the `Arc` before output
    /// publication. This constructor only moves that allocation into the
    /// exact patch.
    pub(crate) fn new(
        messages: Arc<Vec<DirectionalMessage>>,
        references: Option<ReferenceEvidence>,
        preview_mask: u8,
        source_previews: usize,
        target_previews: usize,
        path: Path,
    ) -> Result<Self, Error> {
        let valid_message = |message: &DirectionalMessage| {
            message.object_identifier() != 0
                && message.expected_type() != 0
                && matches!(
                    (message.kind(), message.source(), message.target()),
                    (EvidenceChangeKind::Replace, Some(_), Some(_))
                        | (EvidenceChangeKind::Append, None, Some(_))
                        | (EvidenceChangeKind::Delete, Some(_), None)
                )
        };
        if messages.iter().any(|message| !valid_message(message))
            || !locations_are_strict(messages.as_slice(), Direction::Source)
            || !locations_are_strict(messages.as_slice(), Direction::Target)
            || !reference_routes_are_exact(messages.as_slice(), references.as_ref())
        {
            return Err(Error::InvalidSource { path });
        }
        let preview_bits = usize::try_from(preview_mask.count_ones()).unwrap_or(usize::MAX);
        if preview_mask & !0b111 != 0
            || source_previews > preview_bits
            || target_previews > preview_bits
        {
            return Err(Error::InvalidSource { path });
        }
        Ok(Self {
            messages: Some(messages),
            references,
            inverse: false,
            preview_mask,
            source_previews,
            target_previews,
        })
    }

    fn empty() -> Self {
        Self {
            messages: None,
            references: None,
            inverse: false,
            preview_mask: 0,
            source_previews: 0,
            target_previews: 0,
        }
    }

    fn inverse(&self) -> Self {
        Self {
            messages: self.messages.as_ref().map(Arc::clone),
            references: self.references.clone(),
            inverse: !self.inverse,
            preview_mask: self.preview_mask,
            source_previews: self.target_previews,
            target_previews: self.source_previews,
        }
    }

    pub(crate) fn message_count(&self) -> usize {
        self.messages.as_deref().map_or(0, Vec::len)
    }

    pub(crate) const fn is_inverse(&self) -> bool {
        self.inverse
    }

    pub(crate) fn directional_message(&self, index: usize) -> Option<DirectionalMessage> {
        let message = *self.messages.as_deref()?.get(index)?;
        Some(if self.is_inverse() {
            message.inverse()
        } else {
            message
        })
    }

    pub(crate) fn directional_messages(
        &self,
    ) -> impl ExactSizeIterator<Item = DirectionalMessage> + '_ {
        let inverse = self.is_inverse();
        let messages: &[DirectionalMessage] = self.messages.as_deref().map_or(&[], Vec::as_slice);
        messages
            .iter()
            .copied()
            .map(move |message| if inverse { message.inverse() } else { message })
    }

    pub(crate) fn reference_transition(
        &self,
        message: DirectionalMessage,
    ) -> Option<DirectionalReferenceTransition<'_>> {
        let ReferenceTransition::Exact { route } = message.references else {
            return None;
        };
        let evidence = self.references.as_ref()?;
        Some(DirectionalReferenceTransition {
            evidence,
            route: *evidence.routes.get(route)?,
            inverse: self.inverse,
        })
    }

    pub(crate) const fn preview_mask(&self) -> u8 {
        self.preview_mask
    }

    pub(crate) const fn source_previews(&self) -> usize {
        self.source_previews
    }

    pub(crate) const fn target_previews(&self) -> usize {
        self.target_previews
    }
}

impl fmt::Debug for PatchEvidence {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PatchEvidence")
            .field("messages", &self.message_count())
            .field("inverse", &self.inverse)
            .field("source_previews", &self.source_previews)
            .field("target_previews", &self.target_previews)
            .finish_non_exhaustive()
    }
}

fn reference_routes_are_exact(
    messages: &[DirectionalMessage],
    evidence: Option<&ReferenceEvidence>,
) -> bool {
    let mut next_route = 0usize;
    for message in messages {
        match message.references {
            ReferenceTransition::Preserve => {},
            ReferenceTransition::Exact { route } if route == next_route => {
                let Some(next) = next_route.checked_add(1) else {
                    return false;
                };
                next_route = next;
            },
            ReferenceTransition::Exact { .. } => return false,
        }
    }
    evidence.map_or(next_route == 0, |table| next_route == table.routes.len())
}

fn reference_slice(values: &[u64], span: ReferenceSpan) -> &[u64] {
    &values[span.start..span.start + span.len]
}

fn reference_fields(
    evidence: &ReferenceEvidence,
    route: MessageReferenceRoute,
) -> &[FieldReferenceRoute] {
    &evidence.fields[route.fields.start..route.fields.start + route.fields.len]
}

#[derive(Clone, Copy)]
enum Direction {
    Source,
    Target,
}

fn locations_are_strict(messages: &[DirectionalMessage], direction: Direction) -> bool {
    let mut previous = None;
    for message in messages {
        let location = match direction {
            Direction::Source => message.source(),
            Direction::Target => message.target(),
        };
        let Some(location) = location else {
            continue;
        };
        if previous.is_some_and(|prior| prior >= location) {
            return false;
        }
        previous = Some(location);
    }
    true
}

/// Directional semantic snapshots retained beside exact package artifacts.
///
/// `Package` cloning is an `O(1)` clone of its private `Arc<State>`. Keeping
/// both snapshots lets apply and inverse reuse the already verified semantic
/// projection without reopening either artifact.
#[derive(Clone)]
struct PackagePair {
    source: Package,
    target: Package,
}

impl PackagePair {
    fn new(
        source: Package,
        target: Package,
        source_artifact: &SharedBytes,
        target_artifact: &SharedBytes,
        path: Path,
    ) -> Result<Self, Error> {
        if source.read_options() != target.read_options()
            || !std::ptr::eq(source.source_bytes(), source_artifact.as_ref())
            || !std::ptr::eq(target.source_bytes(), target_artifact.as_ref())
        {
            return Err(Error::Verification { path });
        }
        Ok(Self { source, target })
    }

    fn inverse(&self) -> Self {
        Self {
            source: self.target.snapshot(),
            target: self.source.snapshot(),
        }
    }
}

impl PartialEq for PackagePair {
    fn eq(&self, other: &Self) -> bool {
        // ExactArtifacts compares content separately. This private pair is a
        // process-local snapshot capability, so allocation identity plus the
        // complete read profile is the O(1) semantic-state identity.
        package_snapshot_eq(&self.source, &other.source)
            && package_snapshot_eq(&self.target, &other.target)
    }
}

impl Eq for PackagePair {}

impl fmt::Debug for PackagePair {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PackagePair")
            .finish_non_exhaustive()
    }
}

fn package_snapshot_eq(left: &Package, right: &Package) -> bool {
    left.read_options() == right.read_options()
        && std::ptr::eq(left.source_bytes(), right.source_bytes())
}

/// One process-local exact-source capability for a cell batch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    path: Path,
    requested: usize,
    changed: usize,
    artifacts: OwnedExactArtifacts,
    evidence: PatchEvidence,
    packages: Option<PackagePair>,
}

impl Patch {
    /// Return the content-free semantic target.
    #[must_use]
    pub const fn path(&self) -> Path {
        self.path
    }

    /// Return the number of requested coordinates.
    #[must_use]
    pub const fn len(&self) -> usize {
        self.requested
    }

    /// Return whether the transaction requested no coordinates.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.requested == 0
    }

    /// Return whether the exact target equals the exact source.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// Return a content-free diagnostic fingerprint of the source artifact.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return a content-free diagnostic fingerprint of the target artifact.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Construct the exact reverse capability in constant time.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            path: self.path,
            requested: self.requested,
            changed: self.changed,
            artifacts: self.artifacts.inverse(),
            evidence: self.evidence.inverse(),
            packages: self.packages.as_ref().map(PackagePair::inverse),
        }
    }

    pub(crate) fn authorizes_source(&self, source: &SharedBytes) -> bool {
        self.artifacts.authorizes_owner(source)
    }

    pub(crate) fn target_bytes(&self) -> SharedBytes {
        self.artifacts.target_owner()
    }

    pub(crate) fn changed_cells(&self) -> usize {
        self.changed
    }

    pub(crate) const fn evidence(&self) -> &PatchEvidence {
        &self.evidence
    }

    pub(crate) fn source_package(&self) -> Option<&Package> {
        self.packages.as_ref().map(|packages| &packages.source)
    }

    pub(crate) fn target_package(&self) -> Option<&Package> {
        self.packages.as_ref().map(|packages| &packages.target)
    }

    pub(crate) fn from_exact(
        path: Path,
        requested: usize,
        changed: usize,
        source: SharedBytes,
        target: SharedBytes,
    ) -> Self {
        Self {
            path,
            requested,
            changed,
            artifacts: OwnedExactArtifacts::new(source, target),
            evidence: PatchEvidence::empty(),
            packages: None,
        }
    }

    pub(crate) fn from_exact_with_evidence(
        path: Path,
        requested: usize,
        changed: usize,
        source: SharedBytes,
        target: SharedBytes,
        source_package: Package,
        target_package: Package,
        evidence: PatchEvidence,
    ) -> Result<Self, Error> {
        let packages = PackagePair::new(source_package, target_package, &source, &target, path)?;
        Ok(Self {
            path,
            requested,
            changed,
            artifacts: OwnedExactArtifacts::new(source, target),
            evidence,
            packages: Some(packages),
        })
    }
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("path", &self.path)
            .field("requested", &self.requested)
            .field("changed", &self.changed)
            .finish_non_exhaustive()
    }
}

/// Content-free observations from one committed batch.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    requested_cells: usize,
    changed_cells: usize,
    touched_components: usize,
    refreshed_formula_caches: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    /// Return whether the package artifact changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of requested coordinates.
    #[must_use]
    pub const fn requested_cells(self) -> usize {
        self.requested_cells
    }

    /// Return the number of semantically changed coordinates.
    #[must_use]
    pub const fn changed_cells(self) -> usize {
        self.changed_cells
    }

    /// Return the number of rewritten IWA components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return the number of formula caches refreshed from the final batch.
    #[must_use]
    pub const fn refreshed_formula_caches(self) -> usize {
        self.refreshed_formula_caches
    }

    /// Return the number of deleted root previews.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Return whether this result was constructed by fully reopening the
    /// changed candidate rather than reusing a retained verified snapshot.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }

    pub(crate) const fn unchanged(requested_cells: usize) -> Self {
        Self {
            requested_cells,
            ..Self::empty()
        }
    }

    pub(crate) const fn from_changed(
        requested_cells: usize,
        changed_cells: usize,
        touched_components: usize,
        refreshed_formula_caches: usize,
        deleted_previews: usize,
    ) -> Self {
        let applied = Self::applied(
            requested_cells,
            changed_cells,
            touched_components,
            refreshed_formula_caches,
            deleted_previews,
        );
        Self {
            full_reparse_performed: true,
            ..applied
        }
    }

    pub(crate) const fn applied(
        requested_cells: usize,
        changed_cells: usize,
        touched_components: usize,
        refreshed_formula_caches: usize,
        deleted_previews: usize,
    ) -> Self {
        Self {
            changed: true,
            requested_cells,
            changed_cells,
            touched_components,
            refreshed_formula_caches,
            deleted_previews,
            full_reparse_performed: false,
        }
    }

    const fn empty() -> Self {
        Self {
            changed: false,
            requested_cells: 0,
            changed_cells: 0,
            touched_components: 0,
            refreshed_formula_caches: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }
}

/// A verified target package, exact patch, and content-free diagnostics.
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the verified target package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return the verified target package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Return content-free commit diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    pub(crate) const fn new(package: Package, patch: Patch, diagnostics: Diagnostics) -> Self {
        Self {
            package,
            patch,
            diagnostics,
        }
    }
}

impl fmt::Debug for Commit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Commit")
            .field("patch", &self.patch)
            .field("diagnostics", &self.diagnostics)
            .finish_non_exhaustive()
    }
}

fn map_coordinate_error(_error: CoordinateError) -> Error {
    Error::InvalidAddress
}

fn checked_staged_total(
    current: usize,
    added: usize,
    maximum: usize,
    kind: LimitKind,
    path: Path,
) -> Result<usize, Error> {
    let observed = current.checked_add(added).ok_or(Error::LimitExceeded {
        kind,
        observed: u64::MAX,
        maximum: usize_to_u64(maximum),
        path,
    })?;
    if observed > maximum {
        return Err(Error::LimitExceeded {
            kind,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
            path,
        });
    }
    Ok(observed)
}

fn reserve_change_slot(
    changes: &mut Vec<Change>,
    allocation_events: usize,
    path: Path,
) -> Result<usize, Error> {
    let allocation_events = if changes.len() == changes.capacity() {
        checked_staged_total(
            allocation_events,
            1,
            usize::MAX,
            LimitKind::TransactionWork,
            path,
        )?
    } else {
        allocation_events
    };
    changes.try_reserve(1).map_err(|_error| Error::Allocation {
        kind: LimitKind::Updates,
        amount: 1,
    })?;
    Ok(allocation_events)
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn assert_send_sync<T: Send + Sync>() {}

    #[test]
    fn staged_limit_accepts_maximum_minus_one_then_reports_the_boundary() {
        let maximum = usize::MAX - 1;
        let accepted =
            checked_staged_total(maximum - 1, 1, maximum, LimitKind::Updates, Path::Package)
                .expect("maximum itself remains admissible");
        assert_eq!(accepted, maximum);
        assert!(matches!(
            checked_staged_total(
                accepted,
                1,
                maximum,
                LimitKind::Updates,
                Path::Package,
            ),
            Err(Error::LimitExceeded {
                kind: LimitKind::Updates,
                observed,
                maximum: reported_maximum,
                path: Path::Package,
            }) if observed == usize_to_u64(usize::MAX)
                && reported_maximum == usize_to_u64(maximum)
        ));
    }

    #[test]
    fn staged_usage_observes_the_four_to_eight_k_growth_boundary() {
        let mut changes = Vec::new();
        let mut allocation_events = 0;
        for row in 0..4_096 {
            allocation_events = reserve_change_slot(&mut changes, allocation_events, Path::Package)
                .expect("bounded staging allocation");
            changes.push(Change::clear(CellPosition::new(row, 0)));
        }
        assert_eq!(changes.capacity(), 4_096);
        let prior_events = allocation_events;

        allocation_events = reserve_change_slot(&mut changes, allocation_events, Path::Package)
            .expect("bounded geometric staging allocation");
        changes.push(Change::clear(CellPosition::new(4_096, 0)));

        let usage = StagingUsage {
            change_capacity: changes.capacity(),
            allocation_events,
            formula_nodes: 0,
        };
        assert_eq!(usage.change_capacity(), 8_192);
        assert_eq!(usage.allocation_events(), prior_events + 1);
    }

    #[test]
    fn state_preserves_presence_and_redacts_authored_content() {
        assert_send_sync::<Storage>();
        assert_send_sync::<State>();

        let position = CellPosition::new(2, 1);
        let missing = State::new(position, Storage::Missing);
        assert_eq!(missing.position(), position);
        assert!(missing.storage().is_missing());
        assert_eq!(missing.storage().value(), None);

        let stored = State::new(
            position,
            Storage::Stored(Value::Text("private cell text".to_owned())),
        );
        assert!(matches!(stored.storage(), Storage::Stored(Value::Text(_))));
        assert!(!format!("{stored:?}").contains("private cell text"));
        assert_eq!(
            stored.storage().value().map(Value::cell_type),
            Some(crate::cell::Type::Text)
        );
    }

    #[test]
    fn directional_evidence_inverts_by_sharing_one_prebuilt_slice() {
        const IDENTIFIER: u64 = 9_876_543_210;
        let forward = Arc::new(vec![
            DirectionalMessage::new(
                Some(PhysicalLocation {
                    component: 0,
                    object: 1,
                    message: 2,
                }),
                Some(PhysicalLocation {
                    component: 0,
                    object: 1,
                    message: 2,
                }),
                IDENTIFIER,
                6_001,
                EvidenceChangeKind::Replace,
            ),
            DirectionalMessage::new(
                None,
                Some(PhysicalLocation {
                    component: 1,
                    object: 4,
                    message: 0,
                }),
                IDENTIFIER + 1,
                6_178,
                EvidenceChangeKind::Append,
            ),
            DirectionalMessage::new(
                Some(PhysicalLocation {
                    component: 2,
                    object: 3,
                    message: 0,
                }),
                None,
                IDENTIFIER + 2,
                6_179,
                EvidenceChangeKind::Delete,
            ),
        ]);
        let evidence = PatchEvidence::new(Arc::clone(&forward), None, 0b11, 2, 0, Path::Package)
            .expect("valid directional evidence");
        let reverse = evidence.inverse();

        assert_eq!(evidence.message_count(), 3);
        assert!(!evidence.is_inverse());
        assert!(reverse.is_inverse());
        assert_eq!(reverse.source_previews(), 0);
        assert_eq!(reverse.target_previews(), 2);
        assert!(matches!(
            reverse.directional_message(1).map(DirectionalMessage::kind),
            Some(EvidenceChangeKind::Delete)
        ));
        assert!(
            reverse
                .directional_message(1)
                .and_then(DirectionalMessage::source)
                == forward[1].target()
        );
        assert_eq!(reverse.inverse(), evidence);
        assert_eq!(Arc::strong_count(&forward), 3);
        assert!(!format!("{evidence:?}").contains(&IDENTIFIER.to_string()));
    }

    #[test]
    fn exact_reference_evidence_is_complete_directional_and_redacted() {
        const PRIVATE_REFERENCE: u64 = 8_181_818_181;
        let routes = Arc::new(vec![MessageReferenceRoute::new(
            ReferenceSpan::new(0, 2),
            ReferenceSpan::new(2, 2),
            ReferenceSpan::new(0, 2),
        )]);
        let references = ReferenceEvidence::new(
            Arc::clone(&routes),
            Arc::new(vec![
                FieldReferenceRoute::new(0, ReferenceSpan::new(4, 1), ReferenceSpan::new(5, 1)),
                FieldReferenceRoute::new(1, ReferenceSpan::new(6, 0), ReferenceSpan::new(6, 0)),
            ]),
            Arc::new(vec![
                PRIVATE_REFERENCE,
                PRIVATE_REFERENCE + 1,
                PRIVATE_REFERENCE + 1,
                PRIVATE_REFERENCE + 2,
                PRIVATE_REFERENCE,
                PRIVATE_REFERENCE + 2,
            ]),
            Path::Package,
        )
        .expect("valid flattened reference evidence");
        let messages = Arc::new(vec![
            DirectionalMessage::new(
                Some(PhysicalLocation {
                    component: 0,
                    object: 0,
                    message: 0,
                }),
                Some(PhysicalLocation {
                    component: 0,
                    object: 0,
                    message: 0,
                }),
                1,
                6_001,
                EvidenceChangeKind::Replace,
            )
            .with_reference_transition(0),
        ]);
        let evidence = PatchEvidence::new(messages, Some(references), 0, 0, 0, Path::Package)
            .expect("valid exact reference transition");
        let message = evidence.directional_message(0).expect("message evidence");
        let transition = evidence
            .reference_transition(message)
            .expect("exact reference transition");
        assert_eq!(
            transition.source(),
            &[PRIVATE_REFERENCE, PRIVATE_REFERENCE + 1]
        );
        assert_eq!(
            transition.target(),
            &[PRIVATE_REFERENCE + 1, PRIVATE_REFERENCE + 2]
        );
        let fields: Vec<_> = transition.fields().collect();
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].field_index(), 0);
        assert_eq!(fields[0].source(), &[PRIVATE_REFERENCE]);
        assert_eq!(fields[0].target(), &[PRIVATE_REFERENCE + 2]);

        let inverse = evidence.inverse();
        let reverse = inverse
            .reference_transition(
                inverse
                    .directional_message(0)
                    .expect("inverse message evidence"),
            )
            .expect("inverse reference transition");
        assert_eq!(reverse.source(), transition.target());
        assert_eq!(reverse.target(), transition.source());
        assert!(inverse.inverse() == evidence);
        assert_eq!(Arc::strong_count(&routes), 3);
        assert!(!format!("{inverse:?}").contains(&PRIVATE_REFERENCE.to_string()));
    }

    #[test]
    fn exact_reference_evidence_rejects_missing_field_ordinals_and_routes() {
        let malformed = ReferenceEvidence::new(
            Arc::new(vec![MessageReferenceRoute::new(
                ReferenceSpan::new(0, 1),
                ReferenceSpan::new(1, 1),
                ReferenceSpan::new(0, 2),
            )]),
            Arc::new(vec![
                FieldReferenceRoute::new(0, ReferenceSpan::new(0, 1), ReferenceSpan::new(1, 1)),
                FieldReferenceRoute::new(2, ReferenceSpan::new(0, 1), ReferenceSpan::new(1, 1)),
            ]),
            Arc::new(vec![1, 2]),
            Path::Package,
        );
        assert!(matches!(
            malformed,
            Err(Error::InvalidSource {
                path: Path::Package
            })
        ));

        let messages = Arc::new(vec![
            DirectionalMessage::new(
                Some(PhysicalLocation {
                    component: 0,
                    object: 0,
                    message: 0,
                }),
                Some(PhysicalLocation {
                    component: 0,
                    object: 0,
                    message: 0,
                }),
                1,
                6_001,
                EvidenceChangeKind::Replace,
            )
            .with_reference_transition(0),
        ]);
        assert!(matches!(
            PatchEvidence::new(messages, None, 0, 0, 0, Path::Package,),
            Err(Error::InvalidSource {
                path: Path::Package
            })
        ));
    }

    #[test]
    fn patch_inverse_is_exact_shared_and_content_redacted() {
        let patch = Patch::from_exact(
            Path::Package,
            1,
            0,
            SharedBytes::from_shared_slice(Arc::from(&b"source"[..])),
            SharedBytes::from_shared_slice(Arc::from(&b"target"[..])),
        );
        let inverse = patch.inverse();

        assert_eq!(inverse.inverse(), patch);
        assert_eq!(inverse.source_fingerprint(), patch.target_fingerprint());
        assert_eq!(inverse.target_fingerprint(), patch.source_fingerprint());
        let debug = format!("{inverse:?}");
        assert!(!debug.contains("source"));
        assert!(!debug.contains("target"));
    }

    #[test]
    fn directional_evidence_rejects_inconsistent_topology_without_allocating() {
        let malformed = Arc::new(vec![DirectionalMessage::new(
            None,
            None,
            1,
            6_001,
            EvidenceChangeKind::Replace,
        )]);
        assert!(matches!(
            PatchEvidence::new(malformed, None, 0, 0, 0, Path::Package,),
            Err(Error::InvalidSource {
                path: Path::Package
            })
        ));
    }
}
