//! Implementation of the public [`crate::cell::conditional_highlight`]
//! transactions (record 0370d): the ordered conditional-highlight rules of
//! one cell.
//!
//! A cell's rules are a key in its table's conditional-style list whose
//! entry names a `TST.ConditionalStyleSetArchive`. The set holds each rule
//! twice, as a pre-pivot predicate with argument indexes and as a current
//! predicate with typed arguments, each with a canonical formula graph,
//! and names one paragraph style (bold, text colour) and one cell style
//! (fill) per rule. This is a port of the host editor's
//! `conditional_highlight` module onto the structural working set:
//!
//! - **set** on a cell without rules stores the list (with the table
//!   model's conditional-style owner) when the table has none, the style
//!   set and its styles, a list entry and the cell's key and applied rule;
//! - **set** on a cell whose entry it alone uses rewrites the graph in
//!   place, keeping the existing style identities, adding pairs for new
//!   rules and removing the surplus; a shared entry is released first;
//! - **clear** releases the entry and, when it went, removes the graph.
//!
//! Volatile date rules (today, yesterday, tomorrow, the next or last
//! period, an offset from today) need the calculation engine's volatile
//! owner family; they and tables that already carry it are refused.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the port keeps the host module's order"
)]

mod formula;
mod graph;
mod native;
mod proto;
mod read;
mod volatile;

use std::fmt;

use litchi_iwa_common::color::{RgbColorSpace, Rgba};
pub(in crate::package) use litchi_iwa_common::table::cell::conditional_highlight::{
    Condition, Date, DateRange, Rule,
};
use thiserror::Error as ThisError;

use self::native::{
    BINARY_FUNCTION_ARGUMENT_COUNT, BOOLEAN_VALUE_TYPE_CODE, BooleanPredicateKind,
    CELL_DATA_FORMAT_FUNCTION_INDEX, CHECKBOX_DATA_FORMAT_CODE,
    CONDITIONAL_FUNCTION_ARGUMENT_COUNT, CONDITIONAL_FUNCTION_INDEX, CellPredicateKind,
    CheckboxPredicateKind, DATE_DAY_FUNCTION_INDEX, DATE_MONTH_FUNCTION_INDEX,
    DATE_YEAR_FUNCTION_INDEX, FixedDatePredicateKind, IF_ERROR_FUNCTION_INDEX,
    IS_BLANK_FUNCTION_INDEX, IS_ERROR_FUNCTION_INDEX, IS_NUMBER_FUNCTION_INDEX,
    LOGICAL_AND_FUNCTION_INDEX, LOGICAL_NOT_FUNCTION_INDEX, LOGICAL_OR_FUNCTION_INDEX,
    NativePredicateKind, NumericPredicateKind, NumericSignPredicateKind, PREDICATE_ARGUMENT_DATE,
    PREDICATE_ARGUMENT_NONE, PREDICATE_ARGUMENT_NUMBER, PREDICATE_ARGUMENT_RELATIVE_CELL,
    PREDICATE_ARGUMENT_STRING, PREDICATE_CELL_ARGUMENT_INDEX, PREDICATE_DATE_ARGUMENT_INDEX,
    PREDICATE_DATE_EQUALITY_ARGUMENT_INDEX, PREDICATE_DATE_EQUALITY_CELL_ARGUMENT_INDEX,
    PREDICATE_NUMBER_ARGUMENT_INDEX, PREDICATE_QUALIFIER_NONE, PREDICATE_RANGE_CELL_ARGUMENT_INDEX,
    PREDICATE_RANGE_LOWER_ARGUMENT_INDEX, PREDICATE_RANGE_UPPER_ARGUMENT_INDEX,
    PREDICATE_TEXT_ARGUMENT_INDEX, PREDICATE_UNUSED_ARGUMENT_INDEX, RelativeDatePredicateKind,
    TERNARY_FUNCTION_ARGUMENT_COUNT, TEXT_LENGTH_FUNCTION_INDEX, TEXT_RIGHT_FUNCTION_INDEX,
    TEXT_SEARCH_FUNCTION_INDEX, TODAY_FUNCTION_INDEX, TextPredicateKind,
    UNARY_FUNCTION_ARGUMENT_COUNT, VALUE_TYPE_FUNCTION_INDEX, ZERO_FUNCTION_ARGUMENT_COUNT,
};
use self::proto::{Decode as _, Encode as _};
pub(in crate::package) use self::proto::{tsce, tsd, tsp, tst, tswp};
use super::Package;
use super::structure::{
    self, Fault, Published, Working,
    metadata::{self, Registry, Transition},
};
use super::table_cell_text_style as cells;
use super::table_topology::wire::{
    bytes_field, push_bytes_field, push_varint_field, rebuild, repeated_bytes, single_bytes,
    single_reference, single_varint, varint_field,
};
use super::table_topology::{
    self as topology, LIST_MESSAGE_TYPE, LIST_SEGMENT_MESSAGE_TYPE, TILE_MESSAGE_TYPE, storage,
};
use crate::selector::{SheetSelector, TableSelector};
use crate::table::At;

structure::structural_family!("conditional-highlight", Wire(litchi_iwa_common::LimitKind));

type RgbaColor = Rgba;
type Result<T, E = Error> = std::result::Result<T, E>;

/// The applied-rule value of a cell no rule matches.
const CONDITIONAL_STYLE_NO_APPLIED_RULE: u32 = 15;
const MAX_CONDITIONAL_HIGHLIGHT_RULES: usize = CONDITIONAL_STYLE_NO_APPLIED_RULE as usize;
const CONDITIONAL_LIST_MESSAGE_VERSION: &[u32] = &[3, 2, 10];
const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
/// `DataStore.stringTable`.
const STRING_LIST_FIELD: u32 = 4;
/// `DataStore.conditionalstyletable`.
const CONDITIONAL_LIST_FIELD: u32 = 18;
/// `TableModelArchive.conditional_style_formula_owner_id`.
const CONDITIONAL_OWNER_FIELD: u32 = 39;
/// `TableDataList.ListType.CONDITIONAL_STYLE`.
const CONDITIONAL_LIST_TYPE: u64 = 9;
const SECONDS_PER_DAY: f64 = 86_400.0;

/// A typed, content-redacted conditional-highlight failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed, or its rules use a representation
    /// this route does not read.
    #[error("rooted conditional-highlight source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("{kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} conditional-highlight items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("conditional-highlight candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("conditional-highlight patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// The table selector did not resolve one table of the sheet.
    #[error("table selector did not resolve")]
    TableNotFound,
    /// The cell is outside the table.
    #[error("cell outside the table")]
    CellNotFound,
    /// Commit was requested before staging a change.
    #[error("no conditional-highlight change is staged")]
    NoStagedOperation,
    /// A rule set must hold between one and fifteen rules.
    #[error("conditional highlighting takes one to fifteen rules")]
    InvalidRuleCount,
    /// An object outside the rule graph references the graph's styles.
    #[error("a surviving object depends on the conditional-highlight graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support conditional highlights")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::TableNotFound | Self::CellNotFound => RefusalKind::NotFound,
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
            Self::InvalidRuleCount => RefusalKind::InvalidValue,
            Self::UnsupportedDependency | Self::UnsupportedSource => {
                RefusalKind::UnsupportedContent
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<topology::Error> for Error {
    fn from(error: topology::Error) -> Self {
        from_cells(cells::from_topology(error))
    }
}

fn from_cells(error: cells::Error) -> Error {
    use cells::{Error as E, LimitKind as K, Refusal as R};
    match error {
        E::Refused(R::SheetNotFound) => Error::Refused(Refusal::SheetNotFound),
        E::Refused(R::TableNotFound) => Error::Refused(Refusal::TableNotFound),
        E::Refused(R::CellNotFound) => Error::Refused(Refusal::CellNotFound),
        E::Refused(R::UnsupportedDependency) => Error::Refused(Refusal::UnsupportedDependency),
        E::Refused(R::UnsupportedSource) => Error::Refused(Refusal::UnsupportedSource),
        E::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                K::PackageBytes => LimitKind::PackageBytes,
                K::PayloadBytes => LimitKind::PayloadBytes,
                K::Identifiers => LimitKind::Identifiers,
                K::Wire(kind) => LimitKind::Wire(kind),
            },
            observed,
            maximum,
        },
        E::Allocation { amount } => Error::Allocation { amount },
        E::Verification => Error::Verification,
        E::PatchConflict => Error::PatchConflict,
        _ => Error::InvalidSource,
    }
}

/// The host's message-carrying format errors, content-redacted.
fn invalid_format(_message: impl Into<String>) -> Error {
    Error::InvalidSource
}

fn color_to_native(color: Rgba) -> tsp::Color {
    tsp::Color {
        model: 1,
        r: Some(color.red()),
        g: Some(color.green()),
        b: Some(color.blue()),
        a: Some(color.alpha()),
        other_channels: false,
        rgbspace: Some(match color.color_space() {
            RgbColorSpace::Srgb => 1,
            RgbColorSpace::DisplayP3 => 2,
        }),
    }
}

fn color_from_native(color: &tsp::Color) -> Result<Rgba> {
    if color.model != 1 || color.other_channels {
        return Err(Error::InvalidSource);
    }
    let space = match color.rgbspace.ok_or(Error::InvalidSource)? {
        1 => RgbColorSpace::Srgb,
        2 => RgbColorSpace::DisplayP3,
        _ => return Err(Error::InvalidSource),
    };
    Rgba::new(
        color.r.ok_or(Error::InvalidSource)?,
        color.g.ok_or(Error::InvalidSource)?,
        color.b.ok_or(Error::InvalidSource)?,
        color.a.unwrap_or(1.0),
        space,
    )
    .map_err(|_error| Error::InvalidSource)
}

fn uuid_as_cfuuid(uuid: &tsp::Uuid) -> tsp::CfuuidArchive {
    tsp::CfuuidArchive {
        uuid_bytes: None,
        uuid_w0: Some(uuid.lower as u32),
        uuid_w1: Some((uuid.lower >> 32) as u32),
        uuid_w2: Some(uuid.upper as u32),
        uuid_w3: Some((uuid.upper >> 32) as u32),
    }
}

fn cfuuid_as_uuid(uuid: &tsp::CfuuidArchive) -> Option<tsp::Uuid> {
    let words = || {
        Some(tsp::Uuid {
            lower: u64::from(uuid.uuid_w0?) | (u64::from(uuid.uuid_w1?) << 32),
            upper: u64::from(uuid.uuid_w2?) | (u64::from(uuid.uuid_w3?) << 32),
        })
    };
    let bytes = || {
        let bytes: [u8; 16] = uuid.uuid_bytes.as_deref()?.try_into().ok()?;
        let value = u128::from_be_bytes(bytes);
        Some(tsp::Uuid {
            lower: value as u64,
            upper: (value >> 64) as u64,
        })
    };
    words().or_else(bytes)
}

/// The calculation-engine owner UUID of a table (the host's
/// `formula_owner_uuid_for_table`).
const fn formula_owner_uuid_for_table(table: &tsp::Uuid) -> tsp::Uuid {
    tsp::Uuid {
        lower: table.upper.swap_bytes(),
        upper: table.lower.swap_bytes(),
    }
}

fn parse_table_uuid(value: &str) -> Result<tsp::Uuid> {
    let compact = value.replace('-', "");
    if compact.len() != 32 || !compact.is_ascii() {
        return Err(Error::InvalidSource);
    }
    let mut bytes = [0_u8; 16];
    for (index, byte) in bytes.iter_mut().enumerate() {
        *byte = u8::from_str_radix(&compact[index * 2..index * 2 + 2], 16)
            .map_err(|_error| Error::InvalidSource)?;
    }
    let (upper, lower) = bytes.split_at(8);
    Ok(tsp::Uuid {
        upper: u64::from_be_bytes(upper.try_into().map_err(|_| Error::InvalidSource)?),
        lower: u64::from_be_bytes(lower.try_into().map_err(|_| Error::InvalidSource)?),
    })
}

const fn is_volatile_date_condition(condition: &Condition) -> bool {
    matches!(
        condition,
        Condition::DateIsToday
            | Condition::DateIsYesterday
            | Condition::DateIsTomorrow
            | Condition::DateIsInNext(_)
            | Condition::DateIsInLast(_)
            | Condition::DateIsOffsetFromToday(_)
    )
}

/// One staged change of one cell.
#[derive(Debug, Clone, PartialEq)]
enum Staged {
    Set(Vec<Rule>),
    Clear,
}

/// One immutable edit that changes the conditional highlighting of cells of
/// one table.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    table: usize,
    rows: u32,
    columns: u32,
    changes: Vec<(At, Staged)>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.sheet)
            .field("table", &self.table)
            .field("changes", &self.changes.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn stage(mut self, position: At, change: Staged) -> Result<Self> {
        if position.row() >= self.rows || position.column() >= self.columns {
            return Err(Error::Refused(Refusal::CellNotFound));
        }
        self.changes
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        self.changes.push((position, change));
        Ok(self)
    }

    /// Stage `rules`, in order, as the cell's conditional highlighting
    /// (replacing any it has).
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`], [`Refusal::InvalidRuleCount`] for no rule
    /// or more than fifteen, [`Error::Allocation`]. A volatile date rule
    /// registers the cell with the calculation engine's volatile owner
    /// family (record 0376a).
    pub fn set(self, position: At, rules: impl IntoIterator<Item = Rule>) -> Result<Self> {
        let mut staged = Vec::new();
        for rule in rules {
            staged
                .try_reserve(1)
                .map_err(|_error| Error::Allocation { amount: 1 })?;
            staged.push(rule);
        }
        if staged.is_empty() || staged.len() > MAX_CONDITIONAL_HIGHLIGHT_RULES {
            return Err(Error::Refused(Refusal::InvalidRuleCount));
        }
        self.stage(position, Staged::Set(staged))
    }

    /// Stage removing the cell's conditional highlighting, keeping its value
    /// and base style.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`], [`Error::Allocation`].
    pub fn clear(self, position: At) -> Result<Self> {
        self.stage(position, Staged::Clear)
    }

    /// Validate and atomically publish the staged changes, in order. A
    /// change that leaves the cell as it is is skipped.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], the refusals, source, resource and
    /// verification errors. Failure never publishes a partial package.
    pub fn commit(self) -> Result<Commit> {
        if self.changes.is_empty() {
            return Err(Error::Refused(Refusal::NoStagedOperation));
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        for (position, change) in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            if let Some(published) = apply_one(base, self.sheet, self.table, *position, change)? {
                let expected = match change {
                    Staged::Set(rules) => Some(rules.clone()),
                    Staged::Clear => None,
                };
                if rules_of(&published.package, self.sheet, self.table, *position)? != expected {
                    return Err(Error::Verification);
                }
                touched = touched.saturating_add(published.touched_components);
                current = Some(published.package);
            }
        }
        let Some(package) = current else {
            return Commit::unchanged(self.source);
        };
        Commit::new(
            self.source,
            Published {
                package,
                touched_components: touched,
                deleted_previews: 0,
            },
        )
    }
}

/// The cell's table: its model, data store and geometry.
struct Table {
    info: u64,
    model: u64,
    model_data: Vec<u8>,
    store: Vec<u8>,
    geometry: topology::Geometry,
}

fn table(working: &Working<'_>, sheet: usize, table: usize) -> Result<Table> {
    let target = topology::locate(working, sheet, table).map_err(Error::from)?;
    let (_, model) = topology::message(working, target.model, &[6_001]).map_err(Error::from)?;
    let geometry = topology::geometry(model).map_err(Error::from)?;
    let store = single_bytes(model, 4)?
        .ok_or(Error::InvalidSource)?
        .to_vec();
    Ok(Table {
        info: target.info,
        model: target.model,
        model_data: model.to_vec(),
        store,
        geometry,
    })
}

/// The entry of `key` in a table data list (root or segments) and the
/// object holding it.
fn list_entry(working: &Working<'_>, list: u64, key: u32) -> Result<(u64, Vec<u8>)> {
    let (_, data) = topology::message(working, list, &[LIST_MESSAGE_TYPE]).map_err(Error::from)?;
    let find = |data: &[u8]| -> Result<Option<Vec<u8>>> {
        for entry in repeated_bytes(data, 3)? {
            if single_varint(entry, 1)? == Some(u64::from(key)) {
                return Ok(Some(entry.to_vec()));
            }
        }
        Ok(None)
    };
    if let Some(found) = find(data)? {
        return Ok((list, found));
    }
    for reference in repeated_bytes(data, 4)? {
        let segment = single_varint(reference, 1)?.ok_or(Error::InvalidSource)?;
        let (_, segment_data) = topology::message(working, segment, &[LIST_SEGMENT_MESSAGE_TYPE])
            .map_err(Error::from)?;
        if let Some(found) = find(segment_data)? {
            return Ok((segment, found));
        }
    }
    Err(Error::InvalidSource)
}

/// The cell's BNC record, when it is stored.
fn stored_cell(
    working: &Working<'_>,
    table: &Table,
    position: At,
) -> Result<Option<litchi_numbers_wire::BncCell>> {
    let Some(located) = cells::locate_cell(working, &table.store, position).map_err(from_cells)?
    else {
        return Ok(None);
    };
    litchi_numbers_wire::BncCell::parse(
        &located.bytes,
        &crate::cell::wire::limits_for_source(&located.bytes),
    )
    .map(Some)
    .map_err(|_error| Error::InvalidSource)
}

/// The cell's conditional-style entry: its key, the list, refcount and
/// style set.
struct Entry {
    key: u32,
    list: u64,
    owner: u64,
    refcount: u64,
    style_set: u64,
}

fn entry(working: &Working<'_>, table: &Table, position: At) -> Result<Option<Entry>> {
    let Some(cell) = stored_cell(working, table, position)? else {
        return Ok(None);
    };
    let Some(key) = cell.conditional_style_identifier() else {
        return Ok(None);
    };
    if cell
        .conditional_style_applied_rule()
        .is_some_and(|rule| rule > CONDITIONAL_STYLE_NO_APPLIED_RULE)
    {
        return Err(Error::InvalidSource);
    }
    let list =
        single_reference(&table.store, CONDITIONAL_LIST_FIELD)?.ok_or(Error::InvalidSource)?;
    let (owner, data) = list_entry(working, list, key)?;
    Ok(Some(Entry {
        key,
        list,
        owner,
        refcount: single_varint(&data, 2)?.unwrap_or(0),
        style_set: single_reference(&data, 4)?.ok_or(Error::InvalidSource)?,
    }))
}

/// The cell's rules (`None` without conditional highlighting).
fn rules_of(
    package: &Package,
    sheet: usize,
    table_index: usize,
    position: At,
) -> Result<Option<Vec<Rule>>> {
    let working = Working::new(package)?;
    let table = table(&working, sheet, table_index)?;
    if position.row() >= table.geometry.rows || position.column() >= table.geometry.columns {
        return Err(Error::Refused(Refusal::CellNotFound));
    }
    let Some(entry) = entry(&working, &table, position)? else {
        return Ok(None);
    };
    let set = read::decode_style_set(&working, entry.style_set)?;
    read::decode_rules(&working, &set).map(Some)
}

/// The model's conditional-style owner UUID, when it has one.
fn conditional_owner(model: &[u8]) -> Result<Option<tsp::Uuid>> {
    let Some(owner) = single_bytes(model, CONDITIONAL_OWNER_FIELD)? else {
        return Ok(None);
    };
    Ok(cfuuid_as_uuid(&tsp::CfuuidArchive::decode(owner)?))
}

/// The style identifiers a style set names, pre-pivot rules first.
fn owned_style_ids(set: &tst::ConditionalStyleSetArchive) -> Vec<u64> {
    let mut identifiers = Vec::new();
    let pairs = set
        .rules_prepivot
        .iter()
        .map(|rule| (rule.cell_style.identifier, rule.text_style.identifier))
        .chain(set.rules.iter().flat_map(|rules| {
            rules
                .rule
                .iter()
                .map(|rule| (rule.cell_style.identifier, rule.text_style.identifier))
        }));
    for (cell, text) in pairs {
        for identifier in [cell, text] {
            if identifier != 0 && !identifiers.contains(&identifier) {
                identifiers.push(identifier);
            }
        }
    }
    identifiers
}

/// The rule style pairs of an existing set (the host's
/// `conditional_style_rule_identifiers`).
fn rule_identifiers(
    set: &tst::ConditionalStyleSetArchive,
) -> Result<Vec<graph::RuleStyleIdentifiers>> {
    let pairs: Vec<(u64, u64)> = if set.rules_prepivot.is_empty() {
        set.rules
            .iter()
            .flat_map(|rules| &rules.rule)
            .map(|rule| (rule.text_style.identifier, rule.cell_style.identifier))
            .collect()
    } else {
        set.rules_prepivot
            .iter()
            .map(|rule| (rule.text_style.identifier, rule.cell_style.identifier))
            .collect()
    };
    if pairs.len() != set.rule_count as usize {
        return Err(Error::InvalidSource);
    }
    pairs
        .into_iter()
        .map(|(text_style, cell_style)| {
            if text_style == 0 || cell_style == 0 {
                return Err(Error::InvalidSource);
            }
            Ok(graph::RuleStyleIdentifiers {
                text_style,
                cell_style,
            })
        })
        .collect()
}

/// No object but `owner` references any of `children` (the host's
/// `ensure_children_are_private`).
fn ensure_private(working: &Working<'_>, owner: u64, children: &[u64]) -> Result<()> {
    if children.is_empty() {
        return Ok(());
    }
    for name in working.component_names() {
        for object in &working.archive(&name)?.objects {
            if object.archive_info.identifier == Some(owner) {
                continue;
            }
            for info in &object.archive_info.message_infos {
                if info
                    .object_references
                    .iter()
                    .chain(
                        info.field_infos
                            .iter()
                            .flat_map(|field| &field.object_references),
                    )
                    .any(|identifier| children.contains(identifier))
                {
                    return Err(Error::Refused(Refusal::UnsupportedDependency));
                }
            }
        }
    }
    Ok(())
}

/// Write the cell's conditional-style key and applied rule (`None` removes
/// them), storing a cell when it is empty.
fn write_cell(
    working: &mut Working<'_>,
    table: &Table,
    position: At,
    key: Option<u32>,
    applied: Option<u32>,
) -> Result<()> {
    let Some(located) = cells::locate_cell(working, &table.store, position).map_err(from_cells)?
    else {
        if key.is_none() {
            return Ok(());
        }
        let mut cell = litchi_numbers_wire::BncCell::minimal();
        cell.set_conditional_style(key, applied);
        return cells::materialize_cell(
            working,
            &table.store,
            table.geometry.columns,
            position,
            &cell,
        )
        .map_err(from_cells);
    };
    let limits = crate::cell::wire::limits_for_source(&located.bytes);
    let mut cell = litchi_numbers_wire::BncCell::parse(&located.bytes, &limits)
        .map_err(|_error| Error::InvalidSource)?;
    cell.set_conditional_style(key, applied);
    let encoded = crate::cell::wire::encode(&cell, &limits).map_err(|_| Error::InvalidSource)?;
    let (_, tile_data) =
        topology::message(working, located.tile, &[TILE_MESSAGE_TYPE]).map_err(Error::from)?;
    let mut field_index = 0_usize;
    let rewritten = rebuild(tile_data, |field| {
        let this = field_index;
        field_index += 1;
        if this != located.row_field {
            return Ok(None);
        }
        let row = storage::replace_cell(field.payload, position.column(), &encoded)
            .map_err(Fault::from)?
            .ok_or(Fault::Invalid)?;
        bytes_field(5, &row).map(Some)
    })?;
    cells::replace_message(working, located.tile, located.tile_index, rewritten)?;
    Ok(())
}

/// The cell's value as the rule predicates see it.
enum ConditionalCellValue {
    Blank,
    Boolean(bool),
    Checkbox(bool),
    Date(f64),
    Number(f64),
    Other,
    Text(String),
}

fn cell_value(working: &Working<'_>, table: &Table, position: At) -> Result<ConditionalCellValue> {
    use litchi_numbers_wire::{CachedScalar, StoredValue};
    let Some(cell) = stored_cell(working, table, position)? else {
        return Ok(ConditionalCellValue::Blank);
    };
    Ok(match cell.stored_value() {
        StoredValue::Empty => ConditionalCellValue::Blank,
        StoredValue::Text(key) => {
            let list =
                single_reference(&table.store, STRING_LIST_FIELD)?.ok_or(Error::InvalidSource)?;
            let (_, entry) = list_entry(working, list, key)?;
            let text = single_bytes(&entry, 3)?.ok_or(Error::InvalidSource)?;
            ConditionalCellValue::Text(
                std::str::from_utf8(text)
                    .map_err(|_error| Error::InvalidSource)?
                    .to_owned(),
            )
        },
        _ => match cell
            .cached_scalar()
            .map_err(|_error| Error::InvalidSource)?
        {
            Some(CachedScalar::Number(value)) => ConditionalCellValue::Number(value.get()),
            Some(CachedScalar::Date(value)) => ConditionalCellValue::Date(value.get()),
            Some(CachedScalar::Boolean(value))
                if cell.cell_format_kind()
                    == Some(litchi_numbers_wire::CHECKBOX_CELL_FORMAT_KIND)
                    && cell.control_cell_spec_identifier().is_some() =>
            {
                ConditionalCellValue::Checkbox(value)
            },
            Some(CachedScalar::Boolean(value)) => ConditionalCellValue::Boolean(value),
            _ => ConditionalCellValue::Other,
        },
    })
}

/// The first rule the cell's value matches, or the no-rule value (the
/// host's `applied_rule_for_cell`; volatile rules are never written here).
fn applied_rule(value: &ConditionalCellValue, rules: &[Rule]) -> Result<u32> {
    rules
        .iter()
        .position(|rule| condition_matches(&rule.condition, value))
        .map(|index| u32::try_from(index).map_err(|_error| Error::InvalidSource))
        .transpose()
        .map(|matched| matched.unwrap_or(CONDITIONAL_STYLE_NO_APPLIED_RULE))
}

fn condition_matches(condition: &Condition, value: &ConditionalCellValue) -> bool {
    match (condition, value) {
        (Condition::CellIsBlank, ConditionalCellValue::Blank) => true,
        (
            Condition::CellIsNotBlank,
            ConditionalCellValue::Number(_)
            | ConditionalCellValue::Boolean(_)
            | ConditionalCellValue::Checkbox(_)
            | ConditionalCellValue::Date(_)
            | ConditionalCellValue::Other
            | ConditionalCellValue::Text(_),
        ) => true,
        (Condition::CheckboxIsChecked, ConditionalCellValue::Checkbox(value)) => *value,
        (Condition::CheckboxIsNotChecked, ConditionalCellValue::Checkbox(value)) => !*value,
        (
            Condition::BooleanIsTrue,
            ConditionalCellValue::Boolean(value) | ConditionalCellValue::Checkbox(value),
        ) => *value,
        (
            Condition::BooleanIsFalse,
            ConditionalCellValue::Boolean(value) | ConditionalCellValue::Checkbox(value),
        ) => !*value,
        (Condition::NumberIsPositive, ConditionalCellValue::Number(value)) => *value > 0.0,
        (Condition::NumberIsNegative, ConditionalCellValue::Number(value)) => *value < 0.0,
        (Condition::DateIs(date), ConditionalCellValue::Date(value)) => {
            let lower = date.apple_seconds();
            *value >= lower && *value < lower + SECONDS_PER_DAY
        },
        (Condition::DateIsBefore(date), ConditionalCellValue::Date(value)) => {
            *value < date.apple_seconds()
        },
        (Condition::DateIsAfter(date), ConditionalCellValue::Date(value)) => {
            *value >= date.apple_seconds() + SECONDS_PER_DAY
        },
        (Condition::DateIsBetween(range), ConditionalCellValue::Date(value)) => {
            *value >= range.lower().apple_seconds()
                && *value < range.upper().apple_seconds() + SECONDS_PER_DAY
        },
        (Condition::EqualTo(operand), ConditionalCellValue::Number(value)) => {
            *value == operand.get()
        },
        (Condition::NotEqualTo(operand), ConditionalCellValue::Number(value)) => {
            *value != operand.get()
        },
        (Condition::GreaterThan(operand), ConditionalCellValue::Number(value)) => {
            *value > operand.get()
        },
        (Condition::GreaterThanOrEqualTo(operand), ConditionalCellValue::Number(value)) => {
            *value >= operand.get()
        },
        (Condition::LessThan(operand), ConditionalCellValue::Number(value)) => {
            *value < operand.get()
        },
        (Condition::LessThanOrEqualTo(operand), ConditionalCellValue::Number(value)) => {
            *value <= operand.get()
        },
        (Condition::Between(range), ConditionalCellValue::Number(value)) => {
            *value >= range.lower().get() && *value <= range.upper().get()
        },
        (Condition::NotBetween(range), ConditionalCellValue::Number(value)) => {
            *value < range.lower().get() || *value > range.upper().get()
        },
        (Condition::TextEqualTo(needle), ConditionalCellValue::Text(value)) => {
            equals_case_insensitive(value, needle.as_str())
        },
        (Condition::TextNotEqualTo(needle), ConditionalCellValue::Text(value)) => {
            !equals_case_insensitive(value, needle.as_str())
        },
        (Condition::TextStartsWith(needle), ConditionalCellValue::Text(value)) => {
            starts_with_case_insensitive(value, needle.as_str())
        },
        (Condition::TextDoesNotStartWith(needle), ConditionalCellValue::Text(value)) => {
            !starts_with_case_insensitive(value, needle.as_str())
        },
        (Condition::TextEndsWith(needle), ConditionalCellValue::Text(value)) => {
            ends_with_case_insensitive(value, needle.as_str())
        },
        (Condition::TextDoesNotEndWith(needle), ConditionalCellValue::Text(value)) => {
            !ends_with_case_insensitive(value, needle.as_str())
        },
        (Condition::TextContains(needle), ConditionalCellValue::Text(value)) => {
            contains_case_insensitive(value, needle.as_str())
        },
        (Condition::TextDoesNotContain(needle), ConditionalCellValue::Text(value)) => {
            !contains_case_insensitive(value, needle.as_str())
        },
        _ => false,
    }
}

fn equals_case_insensitive(value: &str, expected: &str) -> bool {
    if value.is_ascii() && expected.is_ascii() {
        return value.eq_ignore_ascii_case(expected);
    }
    value.to_lowercase() == expected.to_lowercase()
}

fn starts_with_case_insensitive(value: &str, prefix: &str) -> bool {
    if value.is_ascii() && prefix.is_ascii() {
        return value
            .as_bytes()
            .get(..prefix.len())
            .is_some_and(|candidate| candidate.eq_ignore_ascii_case(prefix.as_bytes()));
    }
    value.to_lowercase().starts_with(&prefix.to_lowercase())
}

fn ends_with_case_insensitive(value: &str, suffix: &str) -> bool {
    if value.is_ascii() && suffix.is_ascii() {
        return value
            .as_bytes()
            .get(value.len().saturating_sub(suffix.len())..)
            .is_some_and(|candidate| {
                candidate.len() == suffix.len() && candidate.eq_ignore_ascii_case(suffix.as_bytes())
            });
    }
    value.to_lowercase().ends_with(&suffix.to_lowercase())
}

fn contains_case_insensitive(value: &str, needle: &str) -> bool {
    if value.is_ascii() && needle.is_ascii() {
        return value
            .as_bytes()
            .windows(needle.len())
            .any(|candidate| candidate.eq_ignore_ascii_case(needle.as_bytes()));
    }
    value.to_lowercase().contains(&needle.to_lowercase())
}

/// The publication state of one change.
struct Change<'a, 'p> {
    working: &'a mut Working<'p>,
    registry: &'a Registry,
    transition: &'a mut Transition,
}

impl Change<'_, '_> {
    fn component(&self, entry: &str) -> Result<u64> {
        Ok(self.registry.component_for_entry(entry)?.identifier)
    }

    /// Replace one message of `object`, appending `reference` to its
    /// references when given.
    fn replace(
        &mut self,
        object: u64,
        index: usize,
        data: Vec<u8>,
        reference: Option<u64>,
    ) -> Result<()> {
        cells::replace_message(self.working, object, index, data)?;
        if let Some(reference) = reference {
            let info = self
                .working
                .object_mut(object)?
                .archive_info
                .message_infos
                .get_mut(index)
                .ok_or(Error::InvalidSource)?;
            if !info.object_references.contains(&reference) {
                structure::reserve(&mut info.object_references, 1)?;
                info.object_references.push(reference);
            }
        }
        Ok(())
    }

    /// Release one reference of the cell's entry; returns whether it went
    /// (the host's `decrement_table_data_list_entry`).
    fn release_entry(&mut self, entry: &Entry) -> Result<bool> {
        let (index, data) = topology::message(
            self.working,
            entry.owner,
            &[LIST_MESSAGE_TYPE, LIST_SEGMENT_MESSAGE_TYPE],
        )
        .map_err(Error::from)?;
        let mut remaining = std::collections::BTreeMap::from([(entry.key, 1)]);
        let reclaim = entry.owner == entry.list;
        let (payload, pruned) =
            topology::lists::release_with(data, &mut remaining, reclaim).map_err(Fault::from)?;
        if !remaining.is_empty() {
            return Err(Error::InvalidSource);
        }
        if let Some(payload) = payload {
            cells::replace_with_references(
                self.working,
                entry.owner,
                index,
                payload,
                &pruned,
                &[],
            )?;
        }
        Ok(!pruned.is_empty())
    }

    /// Advance the save token of the component holding `object`.
    fn stamp_object(&mut self, object: u64) -> Result<()> {
        if let Some(name) = self.working.locate(object)? {
            let component = self.component(&name)?;
            self.transition.stamp(component)?;
        }
        Ok(())
    }

    /// Advance the save token of the component holding the tile of the
    /// cell at `position`.
    fn stamp_tile(&mut self, table: &Table, position: At) -> Result<()> {
        let (tiles, size) = topology::tiles(&table.store).map_err(Error::from)?;
        let key = u64::from(position.row()) / size;
        if let Some((_, tile)) = tiles
            .iter()
            .find(|(candidate, _)| u64::from(*candidate) == key)
        {
            self.stamp_object(*tile)?;
        }
        Ok(())
    }

    /// Remove `objects` with their registry records.
    fn remove_objects(&mut self, objects: &[u64]) -> Result<()> {
        self.transition.remove_objects(self.registry, objects)?;
        for identifier in objects {
            self.working.remove(*identifier)?;
        }
        Ok(())
    }
}

/// Clear the cell's highlighting (the host's `delete::clear_at_location`).
fn clear(change: &mut Change<'_, '_>, table: &Table, position: At) -> Result<bool> {
    let Some(entry) = entry(change.working, table, position)? else {
        return Ok(false);
    };
    let set = read::decode_style_set(change.working, entry.style_set)?;
    let mut owned = owned_style_ids(&set);
    if entry.refcount == 1 {
        ensure_private(change.working, entry.style_set, &owned)?;
    }
    owned.push(entry.style_set);
    // The host stamps the tile's and the list's components, and the
    // removed objects' components.
    change.stamp_tile(table, position)?;
    change.stamp_object(entry.list)?;
    let removed = change.release_entry(&entry)?;
    write_cell(change.working, table, position, None, None)?;
    if removed {
        for identifier in &owned {
            change.stamp_object(*identifier)?;
        }
        change.remove_objects(&owned)?;
        change.transition.release_suffix(&owned)?;
    }
    Ok(true)
}

/// Write `rules` on a cell without highlighting (the host's
/// `set_at_location`).
fn create(change: &mut Change<'_, '_>, table: &Table, position: At, rules: &[Rule]) -> Result<()> {
    let limits = change.working.archive_limits();
    let existing_list = single_reference(&table.store, CONDITIONAL_LIST_FIELD)?;
    let count = 1 + rules.len() * 2 + usize::from(existing_list.is_none());
    let identifiers = change.working.allocate(count)?;
    let (list, graph_ids) = match existing_list {
        Some(list) => (list, &identifiers[..]),
        None => (identifiers[0], &identifiers[1..]),
    };
    let style_set = graph_ids[0];
    let model_entry = change
        .working
        .locate(table.model)?
        .ok_or(Error::InvalidSource)?;
    let list_entry_name = if existing_list.is_some() {
        change.working.locate(list)?.ok_or(Error::InvalidSource)?
    } else {
        // The list beside the table model, with the model's conditional
        // owner (a fresh, deterministic UUID; the host draws a random one).
        let mut data = Vec::new();
        push_varint_field(&mut data, 1, CONDITIONAL_LIST_TYPE)?;
        push_varint_field(&mut data, 2, 1)?;
        push_varint_field(&mut data, 5, 1)?;
        let object =
            structure::new_object(list, None, LIST_MESSAGE_TYPE, data, Vec::new(), limits)?;
        change.working.insert(&model_entry, object)?;
        let seed = metadata::identity_seed(change.working, list)?;
        let owner = tsp::Uuid {
            lower: metadata::mix(seed).max(3),
            upper: metadata::mix(seed ^ 0x6c69_7463_6869_0370),
        };
        let index = structure::unique_message(change.working.object(table.model)?, &[6_001])?;
        let mut reference = Vec::new();
        push_varint_field(&mut reference, 1, list)?;
        let model = litchi_iwa_common::wire::patch_nested_length_delimited_field(
            &table.model_data,
            &[4, CONDITIONAL_LIST_FIELD],
            false,
            Some(&reference),
        )
        .map_err(|error| structure::map_wire(&error))?;
        let model = litchi_iwa_common::wire::patch_length_delimited_field(
            &model,
            CONDITIONAL_OWNER_FIELD,
            single_bytes(&table.model_data, CONDITIONAL_OWNER_FIELD)?.is_some(),
            Some(&uuid_as_cfuuid(&owner).encode_to_vec()),
        )
        .map_err(|error| structure::map_wire(&error))?;
        change.replace(table.model, index, model, Some(list))?;
        model_entry.clone()
    };
    let table_uuid = parse_table_uuid(
        std::str::from_utf8(single_bytes(&table.model_data, 1)?.ok_or(Error::InvalidSource)?)
            .map_err(|_error| Error::InvalidSource)?,
    )?;
    let formula_owner = formula_owner_uuid_for_table(&table_uuid);
    let pairs = graph::sequential_rule_style_identifiers(style_set, rules.len())?;
    for object in
        graph::conditional_style_graph_objects(style_set, rules, &formula_owner, &pairs, limits)?
    {
        change.working.insert(&list_entry_name, object)?;
    }
    // The entry naming the style set.
    let (index, data) =
        topology::message(change.working, list, &[LIST_MESSAGE_TYPE]).map_err(Error::from)?;
    let data = data.to_vec();
    let key = next_key(change.working, &data)?;
    let mut rebuilt = rebuild(&data, |field| {
        if field.number == 2 {
            return varint_field(2, u64::from(key) + 1).map(Some);
        }
        Ok(None)
    })?;
    let mut entry = Vec::new();
    push_varint_field(&mut entry, 1, u64::from(key))?;
    push_varint_field(&mut entry, 2, 1)?;
    let mut reference = Vec::new();
    push_varint_field(&mut reference, 1, style_set)?;
    push_bytes_field(&mut entry, 4, &reference)?;
    push_bytes_field(&mut rebuilt, 3, &entry)?;
    change.replace(list, index, rebuilt, Some(style_set))?;
    change
        .working
        .object_mut(list)?
        .archive_info
        .message_infos
        .get_mut(index)
        .ok_or(Error::InvalidSource)?
        .versions = CONDITIONAL_LIST_MESSAGE_VERSION.to_vec().into();
    // UUIDs for every new object, in the list's component.
    let component = change.component(&list_entry_name)?;
    let mut registered = Vec::new();
    if existing_list.is_none() {
        registered.push(list);
    }
    registered.extend_from_slice(graph_ids);
    for identifier in registered {
        let seed = metadata::identity_seed(change.working, identifier)?;
        change
            .transition
            .register_uuid(change.registry, component, identifier, seed)?;
    }
    // The host stamps the tile's and the list's components, and the model's
    // when it gave the table its list.
    change.stamp_tile(table, position)?;
    change.stamp_object(list)?;
    if existing_list.is_none() {
        change.stamp_object(table.model)?;
    }
    let value = cell_value(change.working, table, position)?;
    let applied = applied_rule(&value, rules)?;
    write_cell(change.working, table, position, Some(key), Some(applied))
}

/// The next key of a table data list (the host's
/// `next_table_data_list_key`).
fn next_key(working: &Working<'_>, list: &[u8]) -> Result<u32> {
    let mut next = single_varint(list, 2)?.unwrap_or(0).max(1);
    let mut visit = |data: &[u8]| -> Result<()> {
        for entry in repeated_bytes(data, 3)? {
            if let Some(key) = single_varint(entry, 1)? {
                next = next.max(key.checked_add(1).ok_or(Error::InvalidSource)?);
            }
        }
        Ok(())
    };
    visit(list)?;
    for reference in repeated_bytes(list, 4)? {
        let segment = single_varint(reference, 1)?.ok_or(Error::InvalidSource)?;
        let (_, data) = topology::message(working, segment, &[LIST_SEGMENT_MESSAGE_TYPE])
            .map_err(Error::from)?;
        visit(data)?;
    }
    u32::try_from(next).map_err(|_error| Error::Refused(Refusal::UnsupportedSource))
}

/// Rewrite the graph of an entry the cell alone uses (the host's
/// `replace::try_at_location`); `false` when its style pairs repeat.
fn replace(
    change: &mut Change<'_, '_>,
    table: &Table,
    position: At,
    entry: &Entry,
    rules: &[Rule],
) -> Result<bool> {
    let limits = change.working.archive_limits();
    let set = read::decode_style_set(change.working, entry.style_set)?;
    let existing = rule_identifiers(&set)?;
    let mut unique = Vec::new();
    for pair in &existing {
        for identifier in [pair.text_style, pair.cell_style] {
            if unique.contains(&identifier) {
                return Ok(false);
            }
            unique.push(identifier);
        }
    }
    let style_archive = change
        .working
        .locate(entry.style_set)?
        .ok_or(Error::InvalidSource)?;
    let retained = existing.len().min(rules.len());
    let mut pairs = existing[..retained].to_vec();
    let added_count = rules.len().saturating_sub(retained) * 2;
    let added = change.working.allocate(added_count)?;
    for chunk in added.chunks(2) {
        pairs.push(graph::RuleStyleIdentifiers {
            text_style: chunk[0],
            cell_style: chunk[1],
        });
    }
    let removed: Vec<u64> = existing
        .get(rules.len()..)
        .unwrap_or_default()
        .iter()
        .flat_map(|pair| [pair.text_style, pair.cell_style])
        .collect();
    ensure_private(change.working, entry.style_set, &removed)?;
    let table_uuid = parse_table_uuid(
        std::str::from_utf8(single_bytes(&table.model_data, 1)?.ok_or(Error::InvalidSource)?)
            .map_err(|_error| Error::InvalidSource)?,
    )?;
    let formula_owner = formula_owner_uuid_for_table(&table_uuid);
    let objects = graph::conditional_style_graph_objects(
        entry.style_set,
        rules,
        &formula_owner,
        &pairs,
        limits,
    )?;
    let text_archive = existing
        .first()
        .and_then(|pair| change.working.locate(pair.text_style).ok().flatten())
        .unwrap_or_else(|| style_archive.clone());
    let cell_archive = existing
        .first()
        .and_then(|pair| change.working.locate(pair.cell_style).ok().flatten())
        .unwrap_or_else(|| style_archive.clone());
    for (index, object) in objects.into_iter().enumerate() {
        let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
        let archive = match change.working.locate(identifier)? {
            Some(name) => name,
            None if index == 0 => style_archive.clone(),
            None if index % 2 == 1 => text_archive.clone(),
            None => cell_archive.clone(),
        };
        change.working.upsert(&archive, object)?;
    }
    // New styles: UUIDs in their component, and the style set's reference
    // across components.
    let owner = change.component(&style_archive)?;
    for pair in &pairs {
        for (identifier, archive) in [
            (pair.text_style, &text_archive),
            (pair.cell_style, &cell_archive),
        ] {
            if !added.contains(&identifier) {
                continue;
            }
            let target = change.component(archive)?;
            let seed = metadata::identity_seed(change.working, identifier)?;
            change
                .transition
                .register_uuid(change.registry, target, identifier, seed)?;
            change
                .transition
                .reference(change.registry, owner, target, identifier)?;
        }
    }
    change.stamp_tile(table, position)?;
    change.stamp_object(entry.list)?;
    change.stamp_object(entry.style_set)?;
    change.stamp_object(table.model)?;
    for pair in existing.iter().chain(&pairs) {
        change.stamp_object(pair.text_style)?;
        change.stamp_object(pair.cell_style)?;
    }
    change.remove_objects(&removed)?;
    if added.is_empty() {
        change.transition.release_suffix(&removed)?;
    }
    let value = cell_value(change.working, table, position)?;
    let applied = applied_rule(&value, rules)?;
    write_cell(
        change.working,
        table,
        position,
        Some(entry.key),
        Some(applied),
    )?;
    Ok(true)
}

fn apply_one(
    source: &Package,
    sheet: usize,
    table_index: usize,
    position: At,
    staged: &Staged,
) -> Result<Option<Published>> {
    let mut working = Working::new(source)?;
    working.keep_previews();
    let table = table(&working, sheet, table_index)?;
    if position.row() >= table.geometry.rows || position.column() >= table.geometry.columns {
        return Err(Error::Refused(Refusal::CellNotFound));
    }
    let registry =
        metadata::registry(&working)?.ok_or(Error::Refused(Refusal::UnsupportedSource))?;
    let mut transition = Transition::default();
    let mut change = Change {
        working: &mut working,
        registry: &registry,
        transition: &mut transition,
    };
    match staged {
        Staged::Clear => {
            // The host unregisters a volatile cell before it clears it.
            if entry(change.working, &table, position)?.is_some() {
                volatile::remove(&mut change, &table, position)?;
            }
            if !clear(&mut change, &table, position)? {
                return Ok(None);
            }
        },
        Staged::Set(rules) => {
            let current = entry(change.working, &table, position)?;
            // The host unregisters a volatile cell before it rewrites it,
            // and registers it again for volatile rules (record 0376a).
            if current.is_some() {
                volatile::remove(&mut change, &table, position)?;
            }
            match current {
                Some(entry) if entry.refcount > 1 => {
                    // A shared entry: release this cell's reference, then
                    // write a graph of its own.
                    change.release_entry(&entry)?;
                    create(&mut change, &table, position, rules)?;
                },
                Some(entry) => {
                    if !replace(&mut change, &table, position, &entry, rules)? {
                        clear(&mut change, &table, position)?;
                        create(&mut change, &table, position, rules)?;
                    }
                },
                None => create(&mut change, &table, position, rules)?,
            }
            if rules
                .iter()
                .any(|rule| is_volatile_date_condition(&rule.condition))
            {
                volatile::ensure(&mut change, &table, position)?;
            }
        },
    }
    Ok(Some(working.publish(&transition)?))
}

impl Package {
    /// The ordered conditional-highlight rules of one cell (`None` when it
    /// has none).
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::CellNotFound`], [`Error::InvalidSource`]
    /// for a rule representation this route does not read.
    pub fn table_cell_conditional_highlight<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
    ) -> Result<Option<Vec<Rule>>> {
        let (sheet, table) = self.resolve_conditional_target(sheet.into(), table.into())?;
        rules_of(self, sheet, table, position)
    }

    /// Start an immutable conditional-highlight edit of one table.
    ///
    /// # Errors
    ///
    /// Selector and source errors.
    pub fn edit_table_cell_conditional_highlights<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>> {
        let (sheet, table) = self.resolve_conditional_target(sheet.into(), table.into())?;
        let leaf = self.workbook_ref().sheets()[sheet]
            .tables()
            .nth(table)
            .ok_or(Error::Refused(Refusal::TableNotFound))?;
        Ok(Edit {
            source: self,
            sheet,
            table,
            rows: leaf.row_count(),
            columns: leaf.column_count(),
            changes: Vec::new(),
        })
    }

    /// Apply a conditional-highlight patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_table_cell_conditional_highlights(&self, patch: &Patch) -> Result<Commit> {
        apply_patch(self, patch)
    }

    fn resolve_conditional_target(
        &self,
        sheet: SheetSelector<'_>,
        table: TableSelector<'_>,
    ) -> Result<(usize, usize)> {
        let sheet = self
            .workbook_ref()
            .sheet(sheet)
            .map_err(|_error| Error::InvalidSource)?
            .ok_or(Error::Refused(Refusal::SheetNotFound))?
            .index();
        let table = topology::table_position(self, sheet, table).map_err(Error::from)?;
        Ok((sheet, table))
    }
}
