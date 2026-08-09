//! Canonical semantic values and validation for spreadsheet tracked changes.

use litchi_core::{Error, Result};
use std::{
    cmp::Ordering,
    collections::{HashMap, HashSet},
    fmt,
    num::{NonZeroU8, NonZeroU16, NonZeroU32, NonZeroU64, NonZeroU128, NonZeroUsize},
    str::FromStr,
};

use super::limits::Limits;

/// Canonical arbitrary-precision XML Schema `integer`.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct Integer(Box<str>);

impl Integer {
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn parse(value: &str) -> Result<Self> {
        Self::parse_with_limits(value, &Limits::default())
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn parse_with_limits(value: &str, limits: &Limits) -> Result<Self> {
        canonical_integer(value, false, limits.max_integer_digits()).map(Self)
    }

    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    #[must_use]
    pub fn digit_count(&self) -> usize {
        self.0.strip_prefix('-').unwrap_or(&self.0).len()
    }

    #[must_use]
    pub fn is_negative(&self) -> bool {
        self.0.starts_with('-')
    }

    #[must_use]
    pub fn is_zero(&self) -> bool {
        &*self.0 == "0"
    }

    #[must_use]
    pub fn is_positive(&self) -> bool {
        !self.is_negative() && !self.is_zero()
    }

    #[must_use]
    pub fn to_usize(&self) -> Option<usize> {
        (!self.is_negative()).then(|| self.0.parse().ok()).flatten()
    }
}

/// Canonical arbitrary-precision XML Schema `positiveInteger`.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct PositiveInteger(Box<str>);

impl PositiveInteger {
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn parse(value: &str) -> Result<Self> {
        Self::parse_with_limits(value, &Limits::default())
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn parse_with_limits(value: &str, limits: &Limits) -> Result<Self> {
        canonical_integer(value, true, limits.max_integer_digits()).map(Self)
    }

    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    #[must_use]
    pub fn digit_count(&self) -> usize {
        self.0.len()
    }

    #[must_use]
    pub fn to_usize(&self) -> Option<usize> {
        self.0.parse().ok()
    }
}

fn canonical_integer(value: &str, positive: bool, max_digits: usize) -> Result<Box<str>> {
    let value = value.trim_matches(is_xml_space);
    if value.is_empty() || value.chars().any(is_xml_space) {
        return tracked_invalid("integer has invalid XML Schema whitespace");
    }
    let (negative, digits) = match value.as_bytes().first() {
        Some(b'+') => (false, &value[1..]),
        Some(b'-') => (true, &value[1..]),
        _ => (false, value),
    };
    if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return tracked_invalid("integer has an invalid XML Schema lexical value");
    }
    let magnitude = digits.trim_start_matches('0');
    let magnitude = if magnitude.is_empty() { "0" } else { magnitude };
    if magnitude.len() > max_digits {
        return tracked_invalid("integer exceeds configured digit limit");
    }
    if positive && (negative || magnitude == "0") {
        return tracked_invalid("positiveInteger must be greater than zero");
    }
    let negative = negative && magnitude != "0";
    let mut canonical = String::new();
    canonical
        .try_reserve_exact(magnitude.len() + usize::from(negative))
        .map_err(|_error| allocation_invalid("canonical integer"))?;
    if negative {
        canonical.push('-');
    }
    canonical.push_str(magnitude);
    Ok(canonical.into_boxed_str())
}

fn is_xml_space(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}')
}

impl fmt::Display for Integer {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl fmt::Display for PositiveInteger {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl FromStr for Integer {
    type Err = Error;
    fn from_str(value: &str) -> Result<Self> {
        Self::parse(value)
    }
}

impl FromStr for PositiveInteger {
    type Err = Error;
    fn from_str(value: &str) -> Result<Self> {
        Self::parse(value)
    }
}

impl Ord for PositiveInteger {
    fn cmp(&self, other: &Self) -> Ordering {
        self.0
            .len()
            .cmp(&other.0.len())
            .then_with(|| self.0.cmp(&other.0))
    }
}

impl PartialOrd for PositiveInteger {
    fn partial_cmp(&self, other: &Self) -> Option<Ordering> {
        Some(self.cmp(other))
    }
}

impl Ord for Integer {
    fn cmp(&self, other: &Self) -> Ordering {
        match (self.is_negative(), other.is_negative()) {
            (true, false) => Ordering::Less,
            (false, true) => Ordering::Greater,
            (false, false) => magnitude_cmp(self.as_str(), other.as_str()),
            (true, true) => magnitude_cmp(&other.as_str()[1..], &self.as_str()[1..]),
        }
    }
}

impl PartialOrd for Integer {
    fn partial_cmp(&self, other: &Self) -> Option<Ordering> {
        Some(self.cmp(other))
    }
}

fn magnitude_cmp(left: &str, right: &str) -> Ordering {
    left.len().cmp(&right.len()).then_with(|| left.cmp(right))
}

macro_rules! integer_from {
    ($($type:ty),* $(,)?) => {$ (
        impl From<$type> for Integer {
            fn from(value: $type) -> Self {
                Self(value.to_string().into_boxed_str())
            }
        }
    )* };
}

integer_from!(
    i8, i16, i32, i64, i128, isize, u8, u16, u32, u64, u128, usize
);

macro_rules! positive_try_from {
    ($($type:ty),* $(,)?) => {$ (
        impl TryFrom<$type> for PositiveInteger {
            type Error = Error;
            fn try_from(value: $type) -> Result<Self> {
                Self::parse(&value.to_string())
            }
        }
    )* };
}

positive_try_from!(
    i8, i16, i32, i64, i128, isize, u8, u16, u32, u64, u128, usize
);

macro_rules! positive_from_nonzero {
    ($($type:ty),* $(,)?) => {$ (
        impl From<$type> for PositiveInteger {
            fn from(value: $type) -> Self {
                Self(value.get().to_string().into_boxed_str())
            }
        }
    )* };
}

positive_from_nonzero!(
    NonZeroU8,
    NonZeroU16,
    NonZeroU32,
    NonZeroU64,
    NonZeroU128,
    NonZeroUsize
);

impl From<PositiveInteger> for Integer {
    fn from(value: PositiveInteger) -> Self {
        Self(value.0)
    }
}

impl TryFrom<Integer> for PositiveInteger {
    type Error = Error;
    fn try_from(value: Integer) -> Result<Self> {
        if value.is_positive() {
            Ok(Self(value.0))
        } else {
            tracked_invalid("positiveInteger must be greater than zero")
        }
    }
}

/// Whether a recorded spreadsheet change is pending, accepted, or rejected.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum Acceptance {
    Accepted,
    Rejected,
    #[default]
    Pending,
}

/// The structural unit affected by a row, column, or table insertion/deletion.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Dimension {
    Row,
    Column,
    Table,
}

/// Author, date, and comments stored for one change.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Info {
    /// Required by ODF when the value is validated for authoring.
    pub creator: Option<String>,
    /// Required ODF `dateTime` when the value is validated for authoring.
    pub date: Option<String>,
    pub comments: Vec<String>,
}

/// Integer table/row/column coordinates used by the change-tracking vocabulary.
///
/// ODF deliberately uses signed `integer` values here so that an address may
/// identify a location outside the current valid spreadsheet range.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CellAddress {
    pub table: Integer,
    pub column: Integer,
    pub row: Integer,
}

/// A single cell or rectangular source/target range used by a movement.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum RangeAddress {
    Cell(CellAddress),
    Range {
        start: CellAddress,
        end: CellAddress,
    },
}

/// Typed scalar value preserved by `table:change-track-table-cell`.
#[derive(Clone, Debug)]
pub enum CellValue {
    Empty,
    Boolean(bool),
    Number(f64),
    Percentage(f64),
    Currency {
        value: f64,
        code: String,
    },
    Date(String),
    Time(String),
    Text(String),
    /// An ODF error value with an optional `office:string-value`.
    Error(Option<String>),
}

impl PartialEq for CellValue {
    fn eq(&self, other: &Self) -> bool {
        match (self, other) {
            (Self::Empty, Self::Empty) => true,
            (Self::Boolean(left), Self::Boolean(right)) => left == right,
            (Self::Number(left), Self::Number(right))
            | (Self::Percentage(left), Self::Percentage(right)) => equal_double(*left, *right),
            (
                Self::Currency {
                    value: left,
                    code: left_code,
                },
                Self::Currency {
                    value: right,
                    code: right_code,
                },
            ) => equal_double(*left, *right) && left_code == right_code,
            (Self::Date(left), Self::Date(right))
            | (Self::Time(left), Self::Time(right))
            | (Self::Text(left), Self::Text(right)) => left == right,
            (Self::Error(left), Self::Error(right)) => left == right,
            _ => false,
        }
    }
}

/// Former cell state embedded in a tracked change.
#[derive(Clone, Debug, PartialEq)]
pub struct Cell {
    pub address: Option<String>,
    /// Retained for source compatibility with older releases.
    ///
    /// ODF 1.4 does not permit `table:style-name` on
    /// `table:change-track-table-cell`; authoring validation rejects `Some`.
    pub style_name: Option<String>,
    pub matrix_covered: bool,
    pub formula: Option<String>,
    pub matrix_columns: Option<PositiveInteger>,
    pub matrix_rows: Option<PositiveInteger>,
    pub value: CellValue,
    pub display_text: String,
}

/// A deletion nested inside another tracked change.
#[derive(Clone, Debug, PartialEq)]
pub enum NestedDeletion {
    CellContent {
        change_id: Option<String>,
        address: Option<CellAddress>,
        cell: Option<Cell>,
    },
    Change {
        change_id: Option<String>,
    },
}

/// A location removed from a previously tracked insertion or movement.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum CutOff {
    Insertion {
        change_id: String,
        position: Integer,
    },
    MovementPoint {
        position: Integer,
    },
    MovementRange {
        start: Integer,
        end: Integer,
    },
}

/// Metadata common to every top-level spreadsheet change.
#[derive(Clone, Debug, PartialEq)]
pub struct Metadata {
    pub id: String,
    pub acceptance: Acceptance,
    pub rejecting_change_id: Option<String>,
    pub info: Info,
    pub dependencies: Vec<String>,
    pub deletions: Vec<NestedDeletion>,
}

/// A tracked row, column, or table insertion.
#[derive(Clone, Debug, PartialEq)]
pub struct Insertion {
    pub metadata: Metadata,
    pub dimension: Dimension,
    pub position: Integer,
    pub count: PositiveInteger,
    pub table: Option<Integer>,
}

/// A tracked row, column, or table deletion.
#[derive(Clone, Debug, PartialEq)]
pub struct Deletion {
    pub metadata: Metadata,
    pub dimension: Dimension,
    pub position: Integer,
    pub table: Option<Integer>,
    pub multi_deletion_spanned: Option<Integer>,
    pub cut_offs: Vec<CutOff>,
}

/// A tracked cell or range movement.
#[derive(Clone, Debug, PartialEq)]
pub struct Movement {
    pub metadata: Metadata,
    pub source: RangeAddress,
    pub target: RangeAddress,
}

/// A tracked replacement of one cell's content.
#[derive(Clone, Debug, PartialEq)]
pub struct ContentChange {
    pub metadata: Metadata,
    pub address: CellAddress,
    pub previous_change_id: Option<String>,
    pub previous: Cell,
}

/// One top-level spreadsheet change in document order.
#[derive(Clone, Debug, PartialEq)]
pub enum Change {
    Insertion(Insertion),
    Deletion(Deletion),
    Movement(Movement),
    CellContent(ContentChange),
}

impl Change {
    #[must_use]
    pub fn metadata(&self) -> &Metadata {
        match self {
            Self::Insertion(value) => &value.metadata,
            Self::Deletion(value) => &value.metadata,
            Self::Movement(value) => &value.metadata,
            Self::CellContent(value) => &value.metadata,
        }
    }

    /// Validate this record's local grammar and values without resolving IDs.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn validate_with_limits(&self, limits: &Limits) -> Result<()> {
        self.resources_with_limits(limits).map(|_| ())
    }

    /// Return the exact resource delta contributed by this record.
    ///
    /// `nodes` excludes the single `table:tracked-changes` owner node.
    pub(crate) fn resources_with_limits(&self, limits: &Limits) -> Result<Resources> {
        if limits.max_changes() == 0 {
            return tracked_invalid("spreadsheet tracked-change count exceeds resource limit");
        }
        let mut budget = Budget::new(limits);
        budget.add_nodes(2)?; // change and office:change-info
        budget.string(&self.metadata().id, "table:id", false)?;
        validate_change_body(self, &mut budget)?;
        Ok(Resources {
            changes: 1,
            nodes: budget.nodes,
            aggregate_bytes: budget.aggregate_bytes,
        })
    }

    /// Visit every outbound ID relation without allocating.
    pub(crate) fn for_each_relation(&self, mut visit: impl FnMut(&str, RelationKind)) {
        let metadata = self.metadata();
        if let Some(id) = metadata.rejecting_change_id.as_deref() {
            visit(id, RelationKind::Rejecting);
        }
        for id in &metadata.dependencies {
            visit(id, RelationKind::Dependency);
        }
        for deletion in &metadata.deletions {
            match deletion {
                NestedDeletion::CellContent {
                    change_id: Some(id),
                    ..
                } => visit(id, RelationKind::CellContentDeletion),
                NestedDeletion::Change {
                    change_id: Some(id),
                } => visit(id, RelationKind::ChangeDeletion),
                NestedDeletion::CellContent { .. } | NestedDeletion::Change { .. } => {},
            }
        }
        if let Change::Deletion(value) = self {
            for cut_off in &value.cut_offs {
                if let CutOff::Insertion { change_id, .. } = cut_off {
                    visit(change_id, RelationKind::InsertionCutOff);
                }
            }
        }
        if let Change::CellContent(value) = self
            && let Some(id) = value.previous_change_id.as_deref()
        {
            visit(id, RelationKind::Previous);
        }
    }
}

/// Kind of outbound tracked-change ID relation.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub(crate) enum RelationKind {
    Rejecting,
    Dependency,
    CellContentDeletion,
    ChangeDeletion,
    InsertionCutOff,
    Previous,
}

/// Exact resource contribution used by incremental tracked-change validation.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(crate) struct Resources {
    pub(crate) changes: usize,
    pub(crate) nodes: usize,
    pub(crate) aggregate_bytes: usize,
}

/// Cached result of one complete semantic validation pass.
#[derive(Clone, Debug)]
pub(crate) struct Validated {
    pub(crate) resources: Resources,
    pub(crate) record_resources: Vec<Resources>,
    pub(crate) ids: HashMap<String, usize>,
    pub(crate) limits: Limits,
}

/// Spreadsheet-wide tracked-change state and ordered change records.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Changes {
    pub enabled: bool,
    pub changes: Vec<Change>,
}

impl Changes {
    /// Validate with the default authoring resource limits.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn validate(&self) -> Result<()> {
        self.validate_with_limits(&Limits::default())
    }

    /// Validate ODF grammar, references, ordering, values and resource use.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn validate_with_limits(&self, limits: &Limits) -> Result<()> {
        self.validate_indexed_with_limits(limits).map(|_| ())
    }

    /// Validate once and retain exact resource totals and the owned ID index.
    pub(crate) fn validate_indexed_with_limits(&self, limits: &Limits) -> Result<Validated> {
        if self.changes.len() > limits.max_changes() {
            return tracked_invalid("spreadsheet tracked-change count exceeds resource limit");
        }

        let mut budget = Budget::new(limits);
        budget.add_nodes(1)?;

        let mut ids = HashMap::new();
        ids.try_reserve(self.changes.len())
            .map_err(|_error| allocation_invalid("tracked-change id index"))?;
        let mut record_resources = Vec::new();
        record_resources
            .try_reserve_exact(self.changes.len())
            .map_err(|_error| allocation_invalid("tracked-change resource index"))?;

        for (index, change) in self.changes.iter().enumerate() {
            let before_nodes = budget.nodes;
            let before_bytes = budget.aggregate_bytes;
            budget.add_nodes(2)?; // change element and office:change-info
            let metadata = change.metadata();
            budget.string(&metadata.id, "table:id", false)?;
            if ids
                .insert(try_clone_string(&metadata.id, "tracked-change id")?, index)
                .is_some()
            {
                return tracked_invalid(format!(
                    "duplicate spreadsheet tracked-change id '{}'",
                    metadata.id
                ));
            }
            validate_change_body(change, &mut budget)?;
            record_resources.push(Resources {
                changes: 1,
                nodes: budget.nodes - before_nodes,
                aggregate_bytes: budget.aggregate_bytes - before_bytes,
            });
        }

        validate_graph(&self.changes, &ids)?;
        Ok(Validated {
            resources: Resources {
                changes: self.changes.len(),
                nodes: budget.nodes,
                aggregate_bytes: budget.aggregate_bytes,
            },
            record_resources,
            ids,
            limits: *limits,
        })
    }

    /// Validate IDs, references, ordering and graph-wide record sets.
    ///
    /// Callers must locally validate changed records first. This split avoids
    /// re-walking large retained cell payloads during graph-only operations.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn validate_graph_with_limits(&self, limits: &Limits) -> Result<()> {
        if self.changes.len() > limits.max_changes() || self.changes.len() > limits.max_nodes() {
            return tracked_invalid("spreadsheet tracked-change graph exceeds resource limit");
        }
        let mut ids = HashMap::new();
        ids.try_reserve(self.changes.len())
            .map_err(|_error| allocation_invalid("tracked-change id index"))?;
        for (index, change) in self.changes.iter().enumerate() {
            let id = change.metadata().id.as_str();
            if id.is_empty() {
                return tracked_invalid("table:id must not be empty");
            }
            if ids
                .insert(try_clone_string(id, "tracked-change id")?, index)
                .is_some()
            {
                return tracked_invalid(format!("duplicate spreadsheet tracked-change id '{id}'"));
            }
        }
        validate_graph(&self.changes, &ids)
    }

    /// Validate outbound relations for one record against a retained ID index.
    pub(crate) fn validate_record_relations(
        &self,
        index: usize,
        ids: &HashMap<String, usize>,
    ) -> Result<()> {
        if index >= self.changes.len() {
            return tracked_invalid("tracked-change record index is out of bounds");
        }
        validate_record_relations_at(index, &self.changes, ids)
    }

    #[must_use]
    pub fn has_multi_deletion_groups(&self) -> bool {
        self.changes.iter().any(|change| {
            matches!(change, Change::Deletion(value) if value.multi_deletion_spanned.is_some())
        })
    }

    /// Validate a simultaneous-deletion group affected near one structural edit.
    ///
    /// # Errors
    ///
    /// Returns an error when a value violates the format or resource constraints.
    pub fn validate_multi_deletion_near(&self, index: usize) -> Result<()> {
        if self.changes.is_empty() {
            return Ok(());
        }
        let pivot = index.min(self.changes.len() - 1);
        if matches!(
            &self.changes[pivot],
            Change::Deletion(value) if value.multi_deletion_spanned.is_some()
        ) {
            validate_multi_deletion_group_at(&self.changes, pivot)?;
        }
        for start in (0..pivot).rev() {
            let Change::Deletion(value) = &self.changes[start] else {
                continue;
            };
            let Some(total) = value.multi_deletion_spanned.as_ref() else {
                continue;
            };
            let Some(total) = total.to_usize() else {
                return validate_multi_deletion_group_at(&self.changes, start);
            };
            if start.checked_add(total).is_none_or(|end| end > pivot) {
                validate_multi_deletion_group_at(&self.changes, start)?;
            }
            break;
        }
        Ok(())
    }

    /// Validate one known simultaneous-deletion group without searching for it.
    pub(crate) fn validate_multi_deletion_group_at(&self, index: usize) -> Result<()> {
        validate_multi_deletion_group_at(&self.changes, index)
    }

    /// Fallibly clone a validated semantic graph without attacker-sized
    /// infallible string or vector allocations.
    ///
    /// # Errors
    ///
    /// Returns an error when the clone would exceed the supplied limits.
    pub fn try_clone_with_limits(&self, limits: &Limits) -> Result<Self> {
        let validated = self.validate_indexed_with_limits(limits)?;
        self.try_clone_validated_with_limits(&validated, limits)
    }

    /// Fallibly clone using a cache produced by validation of this exact graph.
    pub(crate) fn try_clone_validated_with_limits(
        &self,
        validated: &Validated,
        limits: &Limits,
    ) -> Result<Self> {
        if !validated.limits.same_semantic_limits(limits)
            || validated.resources.changes != self.changes.len()
            || validated.record_resources.len() != self.changes.len()
            || validated.ids.len() != self.changes.len()
        {
            return tracked_invalid("tracked-change validation cache does not match clone input");
        }
        Ok(Self {
            enabled: self.enabled,
            changes: try_clone_vec(&self.changes, "tracked-change records", try_clone_change)?,
        })
    }
}

fn try_clone_vec<T, F>(values: &[T], label: &str, mut clone: F) -> Result<Vec<T>>
where
    F: FnMut(&T) -> Result<T>,
{
    let mut result = Vec::new();
    result
        .try_reserve_exact(values.len())
        .map_err(|_error| allocation_invalid(label))?;
    for value in values {
        result.push(clone(value)?);
    }
    Ok(result)
}

fn try_clone_optional_string(value: &Option<String>, label: &str) -> Result<Option<String>> {
    value
        .as_deref()
        .map(|value| try_clone_string(value, label))
        .transpose()
}

fn try_clone_integer(value: &Integer) -> Result<Integer> {
    Ok(Integer(
        try_clone_string(value.as_str(), "tracked integer")?.into_boxed_str(),
    ))
}

fn try_clone_positive_integer(value: &PositiveInteger) -> Result<PositiveInteger> {
    Ok(PositiveInteger(
        try_clone_string(value.as_str(), "tracked positive integer")?.into_boxed_str(),
    ))
}

fn try_clone_address(value: &CellAddress) -> Result<CellAddress> {
    Ok(CellAddress {
        table: try_clone_integer(&value.table)?,
        column: try_clone_integer(&value.column)?,
        row: try_clone_integer(&value.row)?,
    })
}

fn try_clone_range(value: &RangeAddress) -> Result<RangeAddress> {
    match value {
        RangeAddress::Cell(address) => Ok(RangeAddress::Cell(try_clone_address(address)?)),
        RangeAddress::Range { start, end } => Ok(RangeAddress::Range {
            start: try_clone_address(start)?,
            end: try_clone_address(end)?,
        }),
    }
}

fn try_clone_cell_value(value: &CellValue) -> Result<CellValue> {
    Ok(match value {
        CellValue::Empty => CellValue::Empty,
        CellValue::Boolean(value) => CellValue::Boolean(*value),
        CellValue::Number(value) => CellValue::Number(*value),
        CellValue::Percentage(value) => CellValue::Percentage(*value),
        CellValue::Currency { value, code } => CellValue::Currency {
            value: *value,
            code: try_clone_string(code, "tracked currency code")?,
        },
        CellValue::Date(value) => CellValue::Date(try_clone_string(value, "tracked date value")?),
        CellValue::Time(value) => CellValue::Time(try_clone_string(value, "tracked time value")?),
        CellValue::Text(value) => CellValue::Text(try_clone_string(value, "tracked text value")?),
        CellValue::Error(value) => {
            CellValue::Error(try_clone_optional_string(value, "tracked error value")?)
        },
    })
}

fn try_clone_cell(value: &Cell) -> Result<Cell> {
    Ok(Cell {
        address: try_clone_optional_string(&value.address, "tracked cell address")?,
        style_name: try_clone_optional_string(&value.style_name, "tracked cell style")?,
        matrix_covered: value.matrix_covered,
        formula: try_clone_optional_string(&value.formula, "tracked cell formula")?,
        matrix_columns: value
            .matrix_columns
            .as_ref()
            .map(try_clone_positive_integer)
            .transpose()?,
        matrix_rows: value
            .matrix_rows
            .as_ref()
            .map(try_clone_positive_integer)
            .transpose()?,
        value: try_clone_cell_value(&value.value)?,
        display_text: try_clone_string(&value.display_text, "tracked cell display text")?,
    })
}

fn try_clone_nested_deletion(value: &NestedDeletion) -> Result<NestedDeletion> {
    Ok(match value {
        NestedDeletion::CellContent {
            change_id,
            address,
            cell,
        } => NestedDeletion::CellContent {
            change_id: try_clone_optional_string(change_id, "nested deletion id")?,
            address: address.as_ref().map(try_clone_address).transpose()?,
            cell: cell.as_ref().map(try_clone_cell).transpose()?,
        },
        NestedDeletion::Change { change_id } => NestedDeletion::Change {
            change_id: try_clone_optional_string(change_id, "nested change deletion id")?,
        },
    })
}

fn try_clone_metadata(value: &Metadata) -> Result<Metadata> {
    Ok(Metadata {
        id: try_clone_string(&value.id, "tracked-change id")?,
        acceptance: value.acceptance,
        rejecting_change_id: try_clone_optional_string(
            &value.rejecting_change_id,
            "rejecting change id",
        )?,
        info: Info {
            creator: try_clone_optional_string(&value.info.creator, "change creator")?,
            date: try_clone_optional_string(&value.info.date, "change date")?,
            comments: try_clone_vec(&value.info.comments, "change comments", |value| {
                try_clone_string(value, "change comment")
            })?,
        },
        dependencies: try_clone_vec(&value.dependencies, "change dependencies", |value| {
            try_clone_string(value, "dependency id")
        })?,
        deletions: try_clone_vec(
            &value.deletions,
            "nested deletions",
            try_clone_nested_deletion,
        )?,
    })
}

fn try_clone_cutoff(value: &CutOff) -> Result<CutOff> {
    Ok(match value {
        CutOff::Insertion {
            change_id,
            position,
        } => CutOff::Insertion {
            change_id: try_clone_string(change_id, "insertion cut-off id")?,
            position: try_clone_integer(position)?,
        },
        CutOff::MovementPoint { position } => CutOff::MovementPoint {
            position: try_clone_integer(position)?,
        },
        CutOff::MovementRange { start, end } => CutOff::MovementRange {
            start: try_clone_integer(start)?,
            end: try_clone_integer(end)?,
        },
    })
}

fn try_clone_change(value: &Change) -> Result<Change> {
    Ok(match value {
        Change::Insertion(value) => Change::Insertion(Insertion {
            metadata: try_clone_metadata(&value.metadata)?,
            dimension: value.dimension,
            position: try_clone_integer(&value.position)?,
            count: try_clone_positive_integer(&value.count)?,
            table: value.table.as_ref().map(try_clone_integer).transpose()?,
        }),
        Change::Deletion(value) => Change::Deletion(Deletion {
            metadata: try_clone_metadata(&value.metadata)?,
            dimension: value.dimension,
            position: try_clone_integer(&value.position)?,
            table: value.table.as_ref().map(try_clone_integer).transpose()?,
            multi_deletion_spanned: value
                .multi_deletion_spanned
                .as_ref()
                .map(try_clone_integer)
                .transpose()?,
            cut_offs: try_clone_vec(&value.cut_offs, "deletion cut-offs", try_clone_cutoff)?,
        }),
        Change::Movement(value) => Change::Movement(Movement {
            metadata: try_clone_metadata(&value.metadata)?,
            source: try_clone_range(&value.source)?,
            target: try_clone_range(&value.target)?,
        }),
        Change::CellContent(value) => Change::CellContent(ContentChange {
            metadata: try_clone_metadata(&value.metadata)?,
            address: try_clone_address(&value.address)?,
            previous_change_id: try_clone_optional_string(
                &value.previous_change_id,
                "previous change id",
            )?,
            previous: try_clone_cell(&value.previous)?,
        }),
    })
}

fn validate_change_body(change: &Change, budget: &mut Budget<'_>) -> Result<()> {
    validate_metadata(change.metadata(), budget)?;
    match change {
        Change::Insertion(value) => {
            budget.integer(&value.position, "insertion position")?;
            budget.positive_integer(&value.count, "insertion count")?;
            if let Some(table) = &value.table {
                budget.integer(table, "insertion table")?;
            }
        },
        Change::Deletion(value) => {
            budget.integer(&value.position, "deletion position")?;
            if let Some(table) = &value.table {
                budget.integer(table, "deletion table")?;
            }
            if let Some(span) = &value.multi_deletion_spanned {
                budget.integer(span, "multi-deletion-spanned")?;
            }
            if !value.cut_offs.is_empty() {
                budget.add_nodes(1)?;
            }
            let mut insertion_seen = false;
            for (cut_off_index, cut_off) in value.cut_offs.iter().enumerate() {
                budget.add_nodes(1)?;
                match cut_off {
                    CutOff::Insertion {
                        change_id,
                        position,
                    } => {
                        budget.string(change_id, "insertion cut-off id", false)?;
                        budget.integer(position, "insertion cut-off position")?;
                        if insertion_seen || cut_off_index != 0 {
                            return tracked_invalid(
                                "insertion cut-off must occur at most once and first",
                            );
                        }
                        insertion_seen = true;
                    },
                    CutOff::MovementPoint { position } => {
                        budget.integer(position, "movement cut-off position")?;
                    },
                    CutOff::MovementRange { start, end } if start >= end => {
                        return tracked_invalid("movement cut-off start must precede end");
                    },
                    CutOff::MovementRange { start, end } => {
                        budget.integer(start, "movement cut-off start")?;
                        budget.integer(end, "movement cut-off end")?;
                    },
                }
            }
        },
        Change::Movement(value) => {
            budget.add_nodes(2)?;
            validate_range(&value.source, budget)?;
            validate_range(&value.target, budget)?;
        },
        Change::CellContent(value) => {
            budget.add_nodes(3)?;
            validate_address(&value.address, budget)?;
            validate_cell(&value.previous, budget)?;
            if let Some(id) = &value.previous_change_id {
                budget.string(id, "previous change id", false)?;
            }
        },
    }
    Ok(())
}

fn validate_graph(changes: &[Change], ids: &HashMap<String, usize>) -> Result<()> {
    validate_multi_deletion_sets(changes)?;
    validate_references(changes, ids)?;
    validate_dependency_cycles(changes, ids)
}

struct Budget<'a> {
    limits: &'a Limits,
    nodes: usize,
    aggregate_bytes: usize,
}

impl<'a> Budget<'a> {
    fn new(limits: &'a Limits) -> Self {
        Self {
            limits,
            nodes: 0,
            aggregate_bytes: 0,
        }
    }

    fn add_nodes(&mut self, amount: usize) -> Result<()> {
        self.nodes = self
            .nodes
            .checked_add(amount)
            .ok_or_else(|| Error::InvalidFormat("tracked-change node count overflow".into()))?;
        if self.nodes > self.limits.max_nodes() {
            return tracked_invalid("spreadsheet tracked-change node count exceeds resource limit");
        }
        Ok(())
    }

    fn string(&mut self, value: &str, label: &str, empty_allowed: bool) -> Result<()> {
        if !empty_allowed && value.is_empty() {
            return tracked_invalid(format!("{label} must not be empty"));
        }
        if value.len() > self.limits.max_value_bytes() {
            return tracked_invalid(format!("{label} exceeds per-value byte limit"));
        }
        if !value.chars().all(is_xml_10_scalar) {
            return tracked_invalid(format!("{label} contains a character forbidden by XML 1.0"));
        }
        self.add_bytes(value.len())
    }

    fn integer(&mut self, value: &Integer, label: &str) -> Result<()> {
        if value.digit_count() > self.limits.max_integer_digits() {
            return tracked_invalid(format!("{label} exceeds configured digit limit"));
        }
        self.add_bytes(value.as_str().len())
    }

    fn positive_integer(&mut self, value: &PositiveInteger, label: &str) -> Result<()> {
        if value.digit_count() > self.limits.max_integer_digits() {
            return tracked_invalid(format!("{label} exceeds configured digit limit"));
        }
        self.add_bytes(value.as_str().len())
    }

    fn add_bytes(&mut self, amount: usize) -> Result<()> {
        self.aggregate_bytes = self
            .aggregate_bytes
            .checked_add(amount)
            .ok_or_else(|| Error::InvalidFormat("tracked-change aggregate size overflow".into()))?;
        if self.aggregate_bytes > self.limits.max_aggregate_bytes() {
            return tracked_invalid("tracked-change values exceed aggregate byte limit");
        }
        Ok(())
    }
}

fn tracked_invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

fn allocation_invalid(label: &str) -> Error {
    Error::InvalidFormat(format!("unable to allocate bounded {label}"))
}

fn try_clone_string(value: &str, label: &str) -> Result<String> {
    let mut clone = String::new();
    clone
        .try_reserve_exact(value.len())
        .map_err(|_error| allocation_invalid(label))?;
    clone.push_str(value);
    Ok(clone)
}

fn is_xml_10_scalar(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(value as u32, 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

fn validate_address(address: &CellAddress, budget: &mut Budget<'_>) -> Result<()> {
    budget.integer(&address.table, "tracked cell table")?;
    budget.integer(&address.column, "tracked cell column")?;
    budget.integer(&address.row, "tracked cell row")
}

fn validate_range(range: &RangeAddress, budget: &mut Budget<'_>) -> Result<()> {
    match range {
        RangeAddress::Cell(address) => validate_address(address, budget)?,
        RangeAddress::Range { start, end } => {
            validate_address(start, budget)?;
            validate_address(end, budget)?;
        },
    }
    if let RangeAddress::Range { start, end } = range
        && (start.table > end.table || start.column > end.column || start.row > end.row)
    {
        return tracked_invalid("tracked range start must not follow its end");
    }
    Ok(())
}

fn validate_cell(cell: &Cell, budget: &mut Budget<'_>) -> Result<()> {
    if cell.style_name.is_some() {
        return tracked_invalid(
            "table:style-name is not permitted on an ODF 1.4 change-track-table-cell",
        );
    }
    if let Some(address) = cell.address.as_deref() {
        budget.string(address, "tracked cell address", false)?;
        if !is_odf_cell_address(address) {
            return tracked_invalid("tracked cell address is not an ODF cellAddress");
        }
    }
    if let Some(formula) = cell.formula.as_deref() {
        budget.string(formula, "tracked cell formula", false)?;
    }
    budget.string(&cell.display_text, "tracked cell display text", true)?;
    if !cell.display_text.is_empty() {
        budget.add_nodes(
            cell.display_text
                .bytes()
                .filter(|byte| *byte == b'\n')
                .count()
                + 1,
        )?;
    }
    if let Some(value) = &cell.matrix_columns {
        budget.positive_integer(value, "number-matrix-columns-spanned")?;
    }
    if let Some(value) = &cell.matrix_rows {
        budget.positive_integer(value, "number-matrix-rows-spanned")?;
    }
    match &cell.value {
        CellValue::Currency { value, code } => {
            let _ = value;
            budget.string(code, "tracked cell currency code", true)?;
        },
        CellValue::Date(value) => {
            budget.string(value, "tracked cell date value", false)?;
            if !is_xsd_date_or_datetime(value) {
                return tracked_invalid("tracked cell date has an invalid ODF date lexical value");
            }
        },
        CellValue::Time(value) => {
            budget.string(value, "tracked cell time value", false)?;
            if litchi_odf_common::datatype::Duration::decode_exact(value).is_err() {
                return tracked_invalid("tracked cell time has an invalid XML Schema duration");
            }
        },
        CellValue::Text(value) => {
            budget.string(value, "tracked cell value", true)?;
        },
        CellValue::Error(value) => {
            if let Some(value) = value {
                budget.string(value, "tracked cell error value", true)?;
            }
        },
        CellValue::Empty
        | CellValue::Boolean(_)
        | CellValue::Number(_)
        | CellValue::Percentage(_) => {},
    }
    Ok(())
}

fn equal_double(left: f64, right: f64) -> bool {
    (left.is_nan() && right.is_nan()) || left == right
}

fn is_odf_cell_address(value: &str) -> bool {
    let mut quoted = false;
    let mut separator = None;
    let bytes = value.as_bytes();
    let mut index = 0usize;
    while index < bytes.len() {
        match bytes[index] {
            b'\'' if quoted && bytes.get(index + 1) == Some(&b'\'') => index += 2,
            b'\'' => {
                quoted = !quoted;
                index += 1;
            },
            b'.' if !quoted => {
                if separator.replace(index).is_some() {
                    return false;
                }
                index += 1;
            },
            _ => index += 1,
        }
    }
    if quoted {
        return false;
    }
    let Some(separator) = separator else {
        return false;
    };
    validate_sheet_reference(&value[..separator])
        && validate_cell_reference(&value[separator + 1..])
}

fn validate_sheet_reference(value: &str) -> bool {
    let value = value.strip_prefix('$').unwrap_or(value);
    if value.is_empty() {
        return true;
    }
    if let Some(inner) = value
        .strip_prefix('\'')
        .and_then(|value| value.strip_suffix('\''))
    {
        if inner.is_empty() {
            return false;
        }
        let bytes = inner.as_bytes();
        let mut index = 0usize;
        while index < bytes.len() {
            if bytes[index] == b'\'' {
                if bytes.get(index + 1) != Some(&b'\'') {
                    return false;
                }
                index += 2;
            } else {
                index += 1;
            }
        }
        true
    } else {
        !value.chars().any(|ch| matches!(ch, '.' | ' ' | '\''))
    }
}

fn validate_cell_reference(value: &str) -> bool {
    let value = value.strip_prefix('$').unwrap_or(value);
    let column_end = value.bytes().take_while(u8::is_ascii_uppercase).count();
    if column_end == 0 {
        return false;
    }
    let row = value[column_end..]
        .strip_prefix('$')
        .unwrap_or(&value[column_end..]);
    !row.is_empty() && row.bytes().all(|byte| byte.is_ascii_digit())
}

fn validate_metadata(metadata: &Metadata, budget: &mut Budget<'_>) -> Result<()> {
    if let Some(id) = &metadata.rejecting_change_id {
        budget.string(id, "rejecting change id", false)?;
    }

    let creator = metadata.info.creator.as_deref().ok_or_else(|| {
        Error::InvalidFormat("office:change-info requires dc:creator".to_string())
    })?;
    budget.add_nodes(1)?;
    budget.string(creator, "dc:creator", false)?;

    let date =
        metadata.info.date.as_deref().ok_or_else(|| {
            Error::InvalidFormat("office:change-info requires dc:date".to_string())
        })?;
    budget.add_nodes(1)?;
    budget.string(date, "dc:date", false)?;
    if !is_xsd_datetime(date) {
        return tracked_invalid("dc:date must be an ODF dateTime lexical value");
    }

    budget.add_nodes(metadata.info.comments.len())?;
    for comment in &metadata.info.comments {
        budget.string(comment, "change comment", true)?;
    }

    if !metadata.dependencies.is_empty() {
        budget.add_nodes(1)?;
    }
    budget.add_nodes(metadata.dependencies.len())?;
    let mut seen = HashSet::new();
    seen.try_reserve(metadata.dependencies.len())
        .map_err(|_error| allocation_invalid("dependency set"))?;
    for dependency in &metadata.dependencies {
        budget.string(dependency, "dependency id", false)?;
        if dependency == &metadata.id || !seen.insert(dependency.as_str()) {
            return tracked_invalid(format!(
                "change '{}' has a self or duplicate dependency",
                metadata.id
            ));
        }
    }

    if !metadata.deletions.is_empty() {
        budget.add_nodes(1)?;
    }
    for deletion in &metadata.deletions {
        budget.add_nodes(1)?;
        match deletion {
            NestedDeletion::CellContent {
                change_id,
                address,
                cell,
            } => {
                if let Some(id) = change_id {
                    budget.string(id, "nested deletion id", false)?;
                }
                if let Some(address) = address {
                    budget.add_nodes(1)?;
                    validate_address(address, budget)?;
                }
                if let Some(cell) = cell {
                    budget.add_nodes(1)?;
                    validate_cell(cell, budget)?;
                }
            },
            NestedDeletion::Change { change_id } => {
                if let Some(id) = change_id {
                    budget.string(id, "nested change deletion id", false)?;
                }
            },
        }
    }
    Ok(())
}

fn validate_multi_deletion_sets(changes: &[Change]) -> Result<()> {
    for (index, change) in changes.iter().enumerate() {
        if matches!(change, Change::Deletion(value) if value.multi_deletion_spanned.is_some()) {
            validate_multi_deletion_group_at(changes, index)?;
        }
    }
    Ok(())
}

fn validate_multi_deletion_group_at(changes: &[Change], index: usize) -> Result<()> {
    let Some(Change::Deletion(first)) = changes.get(index) else {
        return Ok(());
    };
    let Some(total) = first.multi_deletion_spanned.as_ref() else {
        return Ok(());
    };
    if !total.is_positive() {
        return tracked_invalid("multi-deletion-spanned must be positive");
    }
    if first.dimension == Dimension::Table {
        return tracked_invalid("multi-deletion-spanned is only valid for rows or columns");
    }
    let Some(total) = total.to_usize() else {
        return tracked_invalid(
            "multi-deletion-spanned exceeds the following deletion record count",
        );
    };
    let end = index.checked_add(total).ok_or_else(|| {
        Error::InvalidFormat("multi-deletion-spanned record count overflow".into())
    })?;
    if end > changes.len() {
        return tracked_invalid(
            "multi-deletion-spanned exceeds the following deletion record count",
        );
    }
    for follower in &changes[index + 1..end] {
        let Change::Deletion(follower) = follower else {
            return tracked_invalid("multi-deletion-spanned records must be consecutive deletions");
        };
        if follower.dimension != first.dimension
            || follower.table != first.table
            || follower.position != first.position
        {
            return tracked_invalid(
                "multi-deletion-spanned records must have matching type, table and position",
            );
        }
        if follower.multi_deletion_spanned.is_some() {
            return tracked_invalid(
                "only the first record in a simultaneous deletion set may carry multi-deletion-spanned",
            );
        }
    }
    Ok(())
}

fn validate_references(changes: &[Change], ids: &HashMap<String, usize>) -> Result<()> {
    for index in 0..changes.len() {
        validate_record_relations_at(index, changes, ids)?;
    }
    Ok(())
}

fn validate_record_relations_at(
    index: usize,
    changes: &[Change],
    ids: &HashMap<String, usize>,
) -> Result<()> {
    let change = &changes[index];
    let metadata = change.metadata();
    if let Some(id) = metadata.rejecting_change_id.as_deref() {
        let target = earlier_target(id, index, changes, ids, "rejecting change")?;
        if target.metadata().acceptance != Acceptance::Rejected {
            return tracked_invalid(format!(
                "rejecting change id '{id}' must target a rejected change"
            ));
        }
    }

    for id in &metadata.dependencies {
        target(id, changes, ids, "dependency")?;
    }

    for deletion in &metadata.deletions {
        match deletion {
            NestedDeletion::CellContent {
                change_id: Some(id),
                address,
                ..
            } => {
                let target = earlier_target(id, index, changes, ids, "cell-content-deletion")?;
                let Change::CellContent(target) = target else {
                    return tracked_invalid(format!(
                        "cell-content-deletion id '{id}' must target a cell-content-change"
                    ));
                };
                if let Some(address) = address
                    && address != &target.address
                {
                    return tracked_invalid(format!(
                        "cell-content-deletion id '{id}' targets a different cell address"
                    ));
                }
            },
            NestedDeletion::Change {
                change_id: Some(id),
            } => {
                earlier_target(id, index, changes, ids, "change-deletion")?;
            },
            NestedDeletion::CellContent { .. } | NestedDeletion::Change { .. } => {},
        }
    }

    if let Change::Deletion(value) = change {
        for cut_off in &value.cut_offs {
            if let CutOff::Insertion { change_id, .. } = cut_off {
                let target = earlier_target(change_id, index, changes, ids, "insertion cut-off")?;
                if !matches!(target, Change::Insertion(_)) {
                    return tracked_invalid(format!(
                        "insertion cut-off id '{change_id}' must target an insertion"
                    ));
                }
            }
        }
    }

    if let Change::CellContent(value) = change
        && let Some(id) = value.previous_change_id.as_deref()
    {
        let target = earlier_target(id, index, changes, ids, "previous cell-content change")?;
        let Change::CellContent(target) = target else {
            return tracked_invalid(format!(
                "previous id '{id}' must target a cell-content-change"
            ));
        };
        if target.address != value.address {
            return tracked_invalid(format!(
                "previous id '{id}' must target the same cell address"
            ));
        }
    }
    Ok(())
}

fn target<'a>(
    id: &str,
    changes: &'a [Change],
    ids: &HashMap<String, usize>,
    label: &str,
) -> Result<&'a Change> {
    let Some(index) = ids.get(id).copied() else {
        return tracked_invalid(format!("{label} references unknown change id '{id}'"));
    };
    Ok(&changes[index])
}

fn earlier_target<'a>(
    id: &str,
    current: usize,
    changes: &'a [Change],
    ids: &HashMap<String, usize>,
    label: &str,
) -> Result<&'a Change> {
    let Some(index) = ids.get(id).copied() else {
        return tracked_invalid(format!("{label} references unknown change id '{id}'"));
    };
    if index >= current {
        return tracked_invalid(format!(
            "{label} id '{id}' must reference an earlier tracked change"
        ));
    }
    Ok(&changes[index])
}

fn validate_dependency_cycles(changes: &[Change], ids: &HashMap<String, usize>) -> Result<()> {
    let mut states = Vec::new();
    states
        .try_reserve_exact(changes.len())
        .map_err(|_error| allocation_invalid("dependency state table"))?;
    states.resize(changes.len(), 0u8);

    let mut stack = Vec::new();
    stack
        .try_reserve(changes.len())
        .map_err(|_error| allocation_invalid("dependency traversal stack"))?;

    for root in 0..changes.len() {
        if states[root] != 0 {
            continue;
        }
        states[root] = 1;
        stack.push((root, 0usize));
        while let Some((index, next_dependency)) = stack.last_mut() {
            let dependencies = &changes[*index].metadata().dependencies;
            if *next_dependency == dependencies.len() {
                states[*index] = 2;
                stack.pop();
                continue;
            }
            let dependency = ids[dependencies[*next_dependency].as_str()];
            *next_dependency += 1;
            match states[dependency] {
                1 => return tracked_invalid("spreadsheet tracked-change dependency cycle"),
                2 => {},
                _ => {
                    states[dependency] = 1;
                    stack.push((dependency, 0));
                },
            }
        }
    }
    Ok(())
}

fn is_xsd_date_or_datetime(value: &str) -> bool {
    if value.contains('T') {
        is_xsd_datetime(value)
    } else {
        split_timezone(value, false)
            .is_some_and(|(date, timezone)| validate_date(date) && validate_timezone(timezone))
    }
}

fn is_xsd_datetime(value: &str) -> bool {
    let Some((body, timezone)) = split_timezone(value, true) else {
        return false;
    };
    let Some((date, time)) = body.split_once('T') else {
        return false;
    };
    !time.contains('T') && validate_date(date) && validate_time(time) && validate_timezone(timezone)
}

fn split_timezone(value: &str, require_time: bool) -> Option<(&str, Option<&str>)> {
    if let Some(body) = value.strip_suffix('Z') {
        return Some((body, Some("Z")));
    }
    let time_start = if require_time {
        value.find('T')?.checked_add(1)?
    } else {
        1
    };
    let zone_start = value.char_indices().rev().find_map(|(index, ch)| {
        let suffix = &value[index..];
        (index >= time_start
            && matches!(ch, '+' | '-')
            && (require_time || (suffix.len() == 6 && suffix.as_bytes().get(3) == Some(&b':'))))
        .then_some(index)
    });
    match zone_start {
        Some(index) => Some((&value[..index], Some(&value[index..]))),
        None => Some((value, None)),
    }
}

fn validate_timezone(timezone: Option<&str>) -> bool {
    let Some(timezone) = timezone else {
        return true;
    };
    if timezone == "Z" {
        return true;
    }
    let bytes = timezone.as_bytes();
    if bytes.len() != 6
        || !matches!(bytes[0], b'+' | b'-')
        || bytes[3] != b':'
        || !bytes[1..3].iter().all(u8::is_ascii_digit)
        || !bytes[4..6].iter().all(u8::is_ascii_digit)
    {
        return false;
    }
    let hours = decimal2(&bytes[1..3]);
    let minutes = decimal2(&bytes[4..6]);
    minutes < 60 && (hours < 14 || (hours == 14 && minutes == 0))
}

fn validate_date(value: &str) -> bool {
    let value = value.strip_prefix('-').unwrap_or(value);
    let Some((year, rest)) = value.split_once('-') else {
        return false;
    };
    let Some((month, day)) = rest.split_once('-') else {
        return false;
    };
    if day.contains('-')
        || year.len() < 4
        || (year.len() > 4 && year.starts_with('0'))
        || !year.bytes().all(|byte| byte.is_ascii_digit())
        || year.bytes().all(|byte| byte == b'0')
        || month.len() != 2
        || day.len() != 2
        || !month.bytes().all(|byte| byte.is_ascii_digit())
        || !day.bytes().all(|byte| byte.is_ascii_digit())
    {
        return false;
    }
    let month = decimal2(month.as_bytes());
    let day = decimal2(day.as_bytes());
    let max_day = match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if year_mod(year.as_bytes(), 400) == 0
            || (year_mod(year.as_bytes(), 4) == 0 && year_mod(year.as_bytes(), 100) != 0) =>
        {
            29
        },
        2 => 28,
        _ => return false,
    };
    (1..=max_day).contains(&day)
}

fn validate_time(value: &str) -> bool {
    let mut parts = value.split(':');
    let (Some(hour), Some(minute), Some(second), None) =
        (parts.next(), parts.next(), parts.next(), parts.next())
    else {
        return false;
    };
    let (second, fraction) = second
        .split_once('.')
        .map_or((second, None), |(whole, fraction)| (whole, Some(fraction)));
    if hour.len() != 2
        || minute.len() != 2
        || second.len() != 2
        || !hour.bytes().all(|byte| byte.is_ascii_digit())
        || !minute.bytes().all(|byte| byte.is_ascii_digit())
        || !second.bytes().all(|byte| byte.is_ascii_digit())
        || fraction.is_some_and(|value| {
            value.is_empty() || !value.bytes().all(|byte| byte.is_ascii_digit())
        })
    {
        return false;
    }
    let hour = decimal2(hour.as_bytes());
    let minute = decimal2(minute.as_bytes());
    let second = decimal2(second.as_bytes());
    minute < 60
        && second < 60
        && (hour < 24
            || (hour == 24
                && minute == 0
                && second == 0
                && fraction.is_none_or(|value| value.bytes().all(|byte| byte == b'0'))))
}

fn decimal2(value: &[u8]) -> u8 {
    (value[0] - b'0') * 10 + (value[1] - b'0')
}

fn year_mod(value: &[u8], modulus: u16) -> u16 {
    value.iter().fold(0u16, |remainder, byte| {
        (remainder * 10 + u16::from(byte - b'0')) % modulus
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn metadata(id: &str) -> Metadata {
        Metadata {
            id: id.to_string(),
            acceptance: Acceptance::Pending,
            rejecting_change_id: None,
            info: Info {
                creator: Some("author".to_string()),
                date: Some("2026-08-08T12:00:00Z".to_string()),
                comments: Vec::new(),
            },
            dependencies: Vec::new(),
            deletions: Vec::new(),
        }
    }

    fn cell(value: CellValue) -> Cell {
        Cell {
            address: Some("'Sheet 1'.$A$1".to_string()),
            style_name: None,
            matrix_covered: false,
            formula: None,
            matrix_columns: None,
            matrix_rows: None,
            value,
            display_text: String::new(),
        }
    }

    fn deletion(id: &str, position: i64, span: Option<i64>) -> Change {
        Change::Deletion(Deletion {
            metadata: metadata(id),
            dimension: Dimension::Row,
            position: position.into(),
            table: Some((-1).into()),
            multi_deletion_spanned: span.map(Into::into),
            cut_offs: Vec::new(),
        })
    }

    #[test]
    fn validates_odf_cell_address_lexical_space() {
        for value in [".A1", "$Sheet.$A$01", "'Sheet 1'.$A$1", "'A.B'.Z9"] {
            assert!(is_odf_cell_address(value), "rejected {value}");
        }
        for value in ["A1", ".a1", "Sheet A.A1", "'Sheet'.A", "''.$A$1"] {
            assert!(!is_odf_cell_address(value), "accepted {value}");
        }
    }

    #[test]
    fn canonical_integers_support_values_beyond_machine_width() {
        let limits = Limits::default().with_max_integer_digits(512);
        let positive_digits = "9".repeat(300);
        let negative_lexical = format!(" \t-000{positive_digits}\r\n");
        let negative = Integer::parse_with_limits(&negative_lexical, &limits)
            .expect("test fixture or operation should succeed");
        let positive =
            PositiveInteger::parse_with_limits(&format!("\n+000{positive_digits}\t"), &limits)
                .expect("test fixture or operation should succeed");
        assert_eq!(negative.as_str(), format!("-{positive_digits}"));
        assert_eq!(positive.as_str(), positive_digits);
        assert!(negative < Integer::from(i128::MIN));
        assert!(
            positive
                > PositiveInteger::try_from(u128::MAX)
                    .expect("test fixture or operation should succeed")
        );

        let insertion = Change::Insertion(Insertion {
            metadata: metadata("huge"),
            dimension: Dimension::Column,
            position: negative,
            count: positive,
            table: Some(
                Integer::parse("-999999999999999999999999999999999999")
                    .expect("test fixture or operation should succeed"),
            ),
        });
        assert!(insertion.validate_with_limits(&limits).is_ok());
        assert!(
            PositiveInteger::parse_with_limits(
                &"1".repeat(513),
                &limits.with_max_integer_digits(512),
            )
            .is_err()
        );
    }

    #[test]
    fn validates_duration_date_error_and_double_families() {
        let limits = Limits::default();
        for value in [
            CellValue::Time("-P1Y2M3DT4H5M6.7S".to_string()),
            CellValue::Date("2024-02-29T24:00:00Z".to_string()),
            CellValue::Error(None),
            CellValue::Error(Some(String::new())),
            CellValue::Number(f64::INFINITY),
            CellValue::Percentage(f64::NEG_INFINITY),
            CellValue::Currency {
                value: f64::NAN,
                code: String::new(),
            },
        ] {
            let mut budget = Budget::new(&limits);
            assert!(validate_cell(&cell(value), &mut budget).is_ok());
        }
        let mut budget = Budget::new(&limits);
        assert!(validate_cell(&cell(CellValue::Time("PT".into())), &mut budget).is_err());
        assert_eq!(CellValue::Number(f64::NAN), CellValue::Number(f64::NAN));
    }

    #[test]
    fn validates_simultaneous_deletion_record_sets() {
        let valid = Changes {
            enabled: true,
            changes: vec![deletion("d1", -2, Some(2)), deletion("d2", -2, None)],
        };
        assert!(valid.validate().is_ok());
        assert!(valid.has_multi_deletion_groups());
        assert!(valid.validate_multi_deletion_near(1).is_ok());

        let invalid = Changes {
            enabled: true,
            changes: vec![deletion("d1", 4, Some(2)), deletion("d2", 5, None)],
        };
        assert!(invalid.validate().is_err());
        assert!(invalid.validate_multi_deletion_near(1).is_err());
    }

    #[test]
    fn fallible_clone_preserves_validated_semantics() {
        let changes = Changes {
            enabled: true,
            changes: vec![deletion("d1", -2, Some(2)), deletion("d2", -2, None)],
        };
        let limits = Limits::default();
        let validated = changes
            .validate_indexed_with_limits(&limits)
            .expect("test fixture or operation should succeed");
        let cloned = changes
            .try_clone_validated_with_limits(&validated, &limits)
            .expect("test fixture or operation should succeed");
        assert_eq!(cloned, changes);
        assert_eq!(
            cloned
                .validate_indexed_with_limits(&Limits::default())
                .expect("test fixture or operation should succeed")
                .resources,
            changes
                .validate_indexed_with_limits(&Limits::default())
                .expect("test fixture or operation should succeed")
                .resources
        );
    }
}
