//! Canonical semantic values and validation for spreadsheet tracked changes.

use litchi_core::{Error, Result};
use std::{
    collections::{HashMap, HashSet},
    num::NonZeroUsize,
};

use super::{MAX_NODES, MAX_VALUE_BYTES, append_size};

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
    pub creator: Option<String>,
    pub date: Option<String>,
    pub comments: Vec<String>,
}

/// Integer table/row/column coordinates used by the change-tracking vocabulary.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct CellAddress {
    pub table: i64,
    pub column: i64,
    pub row: i64,
}

/// A single cell or rectangular source/target range used by a movement.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum RangeAddress {
    Cell(CellAddress),
    Range {
        start: CellAddress,
        end: CellAddress,
    },
}

/// Typed scalar value preserved by `table:change-track-table-cell`.
#[derive(Clone, Debug, PartialEq)]
pub enum CellValue {
    Empty,
    Boolean(bool),
    Number(f64),
    Percentage(f64),
    Currency { value: f64, code: String },
    Date(String),
    Time(String),
    Text(String),
}

/// Former cell state embedded in a tracked change.
#[derive(Clone, Debug, PartialEq)]
pub struct Cell {
    pub address: Option<String>,
    /// Style associated with the historical cell value.
    pub style_name: Option<String>,
    pub matrix_covered: bool,
    pub formula: Option<String>,
    pub matrix_columns: Option<NonZeroUsize>,
    pub matrix_rows: Option<NonZeroUsize>,
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
    Insertion { change_id: String, position: i64 },
    MovementPoint { position: i64 },
    MovementRange { start: i64, end: i64 },
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
    pub position: i64,
    pub count: NonZeroUsize,
    pub table: Option<i64>,
}

/// A tracked row, column, or table deletion.
#[derive(Clone, Debug, PartialEq)]
pub struct Deletion {
    pub metadata: Metadata,
    pub dimension: Dimension,
    pub position: i64,
    pub table: Option<i64>,
    pub multi_deletion_spanned: Option<i64>,
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
    pub fn metadata(&self) -> &Metadata {
        match self {
            Self::Insertion(value) => &value.metadata,
            Self::Deletion(value) => &value.metadata,
            Self::Movement(value) => &value.metadata,
            Self::CellContent(value) => &value.metadata,
        }
    }
}

/// Spreadsheet-wide tracked-change state and ordered change records.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Changes {
    pub enabled: bool,
    pub changes: Vec<Change>,
}

impl Changes {
    /// Validate references, dependency ordering and values before authoring.
    pub fn validate(&self) -> Result<()> {
        if self.changes.len() > MAX_NODES {
            return tracked_invalid("spreadsheet tracked-change count exceeds resource limit");
        }
        let mut ids = HashMap::with_capacity(self.changes.len());
        let mut aggregate = 0usize;
        let mut nodes = 1usize;
        for (index, change) in self.changes.iter().enumerate() {
            let metadata = change.metadata();
            validate_tracked_string(&metadata.id, "table:id", false, &mut aggregate)?;
            if ids.insert(metadata.id.as_str(), index).is_some() {
                return tracked_invalid(format!(
                    "duplicate spreadsheet tracked-change id '{}'",
                    metadata.id
                ));
            }
            nodes = nodes
                .checked_add(2 + metadata.dependencies.len() + metadata.deletions.len())
                .ok_or_else(|| Error::InvalidFormat("tracked-change node count overflow".into()))?;
            validate_tracked_metadata(metadata, &mut aggregate)?;
            match change {
                Change::Insertion(value) => {
                    validate_tracked_position(value.position, "insertion position")?;
                    validate_optional_tracked_position(value.table, "insertion table")?;
                },
                Change::Deletion(value) => {
                    validate_tracked_position(value.position, "deletion position")?;
                    validate_optional_tracked_position(value.table, "deletion table")?;
                    validate_optional_tracked_position(
                        value.multi_deletion_spanned,
                        "multi-deletion-spanned",
                    )?;
                    nodes = nodes.saturating_add(value.cut_offs.len());
                    for cut_off in &value.cut_offs {
                        match cut_off {
                            CutOff::Insertion {
                                change_id,
                                position,
                            } => {
                                validate_tracked_string(
                                    change_id,
                                    "insertion cut-off id",
                                    false,
                                    &mut aggregate,
                                )?;
                                validate_tracked_position(*position, "insertion cut-off position")?;
                            },
                            CutOff::MovementPoint { position } => {
                                validate_tracked_position(*position, "movement cut-off position")?;
                            },
                            CutOff::MovementRange { start, end } => {
                                validate_tracked_position(*start, "movement cut-off start")?;
                                validate_tracked_position(*end, "movement cut-off end")?;
                                if start >= end {
                                    return tracked_invalid(
                                        "movement cut-off start must precede end",
                                    );
                                }
                            },
                        }
                    }
                },
                Change::Movement(value) => {
                    validate_tracked_range(&value.source)?;
                    validate_tracked_range(&value.target)?;
                    nodes = nodes.saturating_add(2);
                },
                Change::CellContent(value) => {
                    validate_tracked_address(&value.address)?;
                    validate_tracked_cell(&value.previous, &mut aggregate)?;
                    if let Some(id) = &value.previous_change_id {
                        validate_tracked_string(id, "previous change id", false, &mut aggregate)?;
                    }
                    nodes = nodes.saturating_add(3);
                },
            }
        }
        if nodes > MAX_NODES {
            return tracked_invalid("spreadsheet tracked-change node count exceeds resource limit");
        }

        for change in &self.changes {
            let metadata = change.metadata();
            validate_tracked_reference(metadata.rejecting_change_id.as_deref(), &ids)?;
            for id in &metadata.dependencies {
                validate_tracked_reference(Some(id), &ids)?;
            }
            for deletion in &metadata.deletions {
                let id = match deletion {
                    NestedDeletion::CellContent { change_id, .. }
                    | NestedDeletion::Change { change_id } => change_id.as_deref(),
                };
                validate_tracked_reference(id, &ids)?;
            }
            if let Change::Deletion(value) = change {
                for cut_off in &value.cut_offs {
                    if let CutOff::Insertion { change_id, .. } = cut_off {
                        validate_tracked_reference(Some(change_id), &ids)?;
                    }
                }
            }
            if let Change::CellContent(value) = change {
                validate_tracked_reference(value.previous_change_id.as_deref(), &ids)?;
            }
        }
        let mut states = vec![0u8; self.changes.len()];
        for index in 0..self.changes.len() {
            visit_tracked_dependencies(index, &self.changes, &ids, &mut states)?;
        }
        Ok(())
    }
}

fn tracked_invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

fn validate_tracked_string(
    value: &str,
    label: &str,
    empty_allowed: bool,
    aggregate: &mut usize,
) -> Result<()> {
    if !empty_allowed && value.is_empty() {
        return tracked_invalid(format!("{label} must not be empty"));
    }
    if value.len() > MAX_VALUE_BYTES {
        return tracked_invalid(format!("{label} exceeds 64 KiB"));
    }
    append_size(aggregate, value.len())
}

fn validate_tracked_position(value: i64, label: &str) -> Result<()> {
    if value < 0 {
        return tracked_invalid(format!("{label} must be nonnegative"));
    }
    Ok(())
}

fn validate_optional_tracked_position(value: Option<i64>, label: &str) -> Result<()> {
    value.map_or(Ok(()), |value| validate_tracked_position(value, label))
}

fn validate_tracked_address(address: &CellAddress) -> Result<()> {
    validate_tracked_position(address.table, "tracked cell table")?;
    validate_tracked_position(address.column, "tracked cell column")?;
    validate_tracked_position(address.row, "tracked cell row")
}

fn validate_tracked_range(range: &RangeAddress) -> Result<()> {
    match range {
        RangeAddress::Cell(address) => validate_tracked_address(address),
        RangeAddress::Range { start, end } => {
            validate_tracked_address(start)?;
            validate_tracked_address(end)?;
            if (start.table, start.row, start.column) > (end.table, end.row, end.column) {
                return tracked_invalid("tracked range start must not follow its end");
            }
            Ok(())
        },
    }
}

fn validate_tracked_cell(cell: &Cell, aggregate: &mut usize) -> Result<()> {
    for (value, label) in [
        (cell.address.as_deref(), "tracked cell address"),
        (cell.style_name.as_deref(), "tracked cell style name"),
        (cell.formula.as_deref(), "tracked cell formula"),
    ] {
        if let Some(value) = value {
            validate_tracked_string(value, label, false, aggregate)?;
        }
    }
    validate_tracked_string(
        &cell.display_text,
        "tracked cell display text",
        true,
        aggregate,
    )?;
    match &cell.value {
        CellValue::Number(value) | CellValue::Percentage(value) if !value.is_finite() => {
            return tracked_invalid("tracked cell numeric value must be finite");
        },
        CellValue::Currency { value, code } => {
            if !value.is_finite() {
                return tracked_invalid("tracked cell currency value must be finite");
            }
            validate_tracked_string(code, "tracked cell currency code", false, aggregate)?;
        },
        CellValue::Date(value) | CellValue::Time(value) | CellValue::Text(value) => {
            validate_tracked_string(value, "tracked cell value", true, aggregate)?;
        },
        _ => {},
    }
    Ok(())
}

fn validate_tracked_metadata(metadata: &Metadata, aggregate: &mut usize) -> Result<()> {
    if let Some(id) = &metadata.rejecting_change_id {
        validate_tracked_string(id, "rejecting change id", false, aggregate)?;
    }
    for value in [
        metadata.info.creator.as_deref(),
        metadata.info.date.as_deref(),
    ]
    .into_iter()
    .flatten()
    {
        validate_tracked_string(value, "change metadata", true, aggregate)?;
    }
    for comment in &metadata.info.comments {
        validate_tracked_string(comment, "change comment", true, aggregate)?;
    }
    let mut seen = HashSet::new();
    for dependency in &metadata.dependencies {
        validate_tracked_string(dependency, "dependency id", false, aggregate)?;
        if dependency == &metadata.id || !seen.insert(dependency) {
            return tracked_invalid(format!(
                "change '{}' has a self or duplicate dependency",
                metadata.id
            ));
        }
    }
    for deletion in &metadata.deletions {
        match deletion {
            NestedDeletion::CellContent {
                change_id,
                address,
                cell,
            } => {
                if let Some(id) = change_id {
                    validate_tracked_string(id, "nested deletion id", false, aggregate)?;
                }
                if let Some(address) = address {
                    validate_tracked_address(address)?;
                }
                if let Some(cell) = cell {
                    validate_tracked_cell(cell, aggregate)?;
                }
                if address.is_none() && cell.is_none() {
                    return tracked_invalid("cell-content-deletion requires an address or cell");
                }
            },
            NestedDeletion::Change { change_id } => {
                if let Some(id) = change_id {
                    validate_tracked_string(id, "nested change deletion id", false, aggregate)?;
                }
            },
        }
    }
    Ok(())
}

fn validate_tracked_reference(reference: Option<&str>, ids: &HashMap<&str, usize>) -> Result<()> {
    if let Some(reference) = reference
        && !ids.contains_key(reference)
    {
        return tracked_invalid(format!(
            "tracked change references unknown id '{reference}'"
        ));
    }
    Ok(())
}

fn visit_tracked_dependencies(
    index: usize,
    changes: &[Change],
    ids: &HashMap<&str, usize>,
    states: &mut [u8],
) -> Result<()> {
    match states[index] {
        1 => return tracked_invalid("spreadsheet tracked-change dependency cycle"),
        2 => return Ok(()),
        _ => {},
    }
    states[index] = 1;
    for dependency in &changes[index].metadata().dependencies {
        visit_tracked_dependencies(ids[dependency.as_str()], changes, ids, states)?;
    }
    states[index] = 2;
    Ok(())
}
