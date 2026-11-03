//! The shared table sort executor: a table's persisted sort order runs as a
//! physical row permutation, as Numbers' and Pages' **Sort Now** does
//! (record 0371, lifted from `litchi-pages` for Numbers and Keynote in
//! record 0421).
//!
//! The executor follows the migration host's scalar executor
//! (`numbers::editor::table_sort::apply`): every sort key must be a complete
//! column of one scalar kind (plain text, finite number, boolean, date or
//! duration); formula, error and unsupported body cells, filters, grouping,
//! pivots, spill state, conditional styles and a missing row UID map are
//! refused. The stable row permutation is the shared
//! `litchi_iwa_common::table::sort::planning` planner. [`execute`] rewrites,
//! in place on one structural working set over the exact source:
//!
//! - the tiles' row records (each record moves to its destination tile and
//!   local index; cells keep their bytes);
//! - the row header buckets;
//! - the row UID map;
//! - the stroke layers (borders stay with their cells);
//! - the user-hidden row states, which keep their positions.
//!
//! Header and footer rows stay put. The leaf locates the table
//! ([`Target`]), reads the persisted order, checks the scope and merges, and
//! supplies the user-hidden rows; [`Hooks`] carries the per-leaf policies.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::too_many_lines,
    reason = "the executor keeps its phases in plan order"
)]

use std::collections::{HashMap, HashSet};

use litchi_iwa_common::table::sort::planning::{RowPermutation, stable_body_row_permutation};
use litchi_iwa_common::table::sort::{Order, RowRange};
use litchi_iwa_common::wire::{
    patch_varint_field, rewrite_repeated_length_delimited_fields, rewrite_repeated_varint_fields,
};
use litchi_numbers_wire::{BncCellView, CachedScalar, StoredValue};

use crate::structure::wire::{
    repeated_bytes, repeated_varints, single_bytes, single_reference, single_varint,
};
use crate::structure::{Fault, Published, TableLeaf, Working, has_field, metadata::Transition};
use crate::topology::{message, replace_payload};

const TABLE_INFO_MESSAGE_TYPES: [u32; 2] = [6_000, 6_003];
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TILE_MESSAGE_TYPE: u32 = 6_002;
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
const UID_MAP_MESSAGE_TYPES: [u32; 2] = [6_267, 6_200];
const STROKE_SIDECAR_MESSAGE_TYPE: u32 = 6_305;
const STROKE_LAYER_MESSAGE_TYPE: u32 = 6_306;
const DATA_LIST_MESSAGE_TYPES: [u32; 2] = [6_005, 6_201];
const FORMULA_OWNER_DEPENDENCIES_MESSAGE_TYPE: u32 = 4_008;

const MODEL_DATA_STORE_FIELD: u32 = 4;
const MODEL_ROWS_FIELD: u32 = 6;
const MODEL_COLUMNS_FIELD: u32 = 7;
const MODEL_HEADER_ROWS_FIELD: u32 = 9;
const MODEL_FOOTER_ROWS_FIELD: u32 = 11;
const MODEL_HIDDEN_ROWS_FIELD: u32 = 14;
const MODEL_HIDDEN_COLUMNS_FIELD: u32 = 15;
const MODEL_HIDDEN_ROW_OWNER_FIELD: u32 = 35;
const MODEL_HIDDEN_COLUMN_OWNER_FIELD: u32 = 34;
const MODEL_ROW_FILTER_FIELD: u32 = 38;
const MODEL_FILTERED_ROWS_FIELD: u32 = 40;
const MODEL_USER_HIDDEN_ROWS_FIELD: u32 = 41;
const MODEL_USER_HIDDEN_COLUMNS_FIELD: u32 = 42;
const MODEL_UID_MAP_FIELD: u32 = 46;
const MODEL_STROKE_SIDECAR_FIELD: u32 = 49;
const MODEL_PIVOT_OWNER_FIELD: u32 = 85;
const MODEL_CATEGORY_OWNER_FIELD: u32 = 86;
const MODEL_HIDDEN_STATES_OWNER_FIELD: u32 = 70;
const MODEL_DEPRECATED_CATEGORY_OWNER_FIELD: u32 = 81;

const STORE_ROW_HEADERS_FIELD: u32 = 1;
const STORE_TILES_FIELD: u32 = 3;
const STORE_STRING_TABLE_FIELD: u32 = 4;
const STORE_CONDITIONAL_STYLE_FIELD: u32 = 18;
const HEADER_STORAGE_BUCKETS_FIELD: u32 = 2;
const TILE_STORAGE_TILES_FIELD: u32 = 1;
const TILE_STORAGE_SIZE_FIELD: u32 = 2;
const TILE_REFERENCE_ID_FIELD: u32 = 1;
const TILE_REFERENCE_TILE_FIELD: u32 = 2;
const TILE_NUMROWS_FIELD: u32 = 4;
const TILE_ROWS_FIELD: u32 = 5;
const ROW_INDEX_FIELD: u32 = 1;
const ROW_STORAGE_FIELD: u32 = 6;
const ROW_OFFSETS_FIELD: u32 = 7;
const ROW_WIDE_OFFSETS_FIELD: u32 = 8;
const BUCKET_HEADERS_FIELD: u32 = 2;
const HEADER_INDEX_FIELD: u32 = 1;
const UID_ROW_INDEX_FOR_UID_FIELD: u32 = 5;
const UID_ROW_UID_FOR_INDEX_FIELD: u32 = 6;
const UID_SORTED_ROWS_FIELD: u32 = 4;
const LIST_ENTRIES_FIELD: u32 = 3;
const LIST_ENTRY_KEY_FIELD: u32 = 1;
const LIST_ENTRY_STRING_FIELD: u32 = 3;
/// Rows per row-header bucket (the host's and Numbers' 65,536; the Pages
/// port used 256 and refused sorts that move a header past row 255).
const HEADER_BUCKET_ROWS: usize = 65_536;
const DEFAULT_TILE_SIZE: u64 = 256;

/// What a sort execution cannot move safely.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Unsupported {
    /// A body cell holds a formula.
    Formula,
    /// A body cell holds an error or a formula error.
    ErrorCell,
    /// A body cell has a cell type this executor does not move.
    CellType,
    /// A body row has no stored key cell, or a key cell has no scalar.
    MissingKey,
    /// A key column mixes scalar kinds.
    MixedKinds,
    /// Rows are filtered, grouped by category, pivoted or carry spill state.
    FilterOrGrouping,
    /// The table has merged cells.
    MergedCells,
    /// The table has conditional styles.
    ConditionalStyles,
    /// The table has no stable row UID map.
    NoRowUids,
    /// The table has hidden rows or columns.
    HiddenAxes,
    /// The table stores legacy (pre-BNC) rows or an unexpected layout.
    Layout,
}

/// A well-formed request the executor refuses.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// The row range is outside the table's body rows.
    RowsOutOfRange,
    /// The table holds content this executor does not move.
    Unsupported(Unsupported),
}

/// A typed sort-execution failure; a leaf maps it into its family's public
/// error.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request the executor refuses.
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    InvalidSource,
    /// A structural working-set fault (limits, allocation, verification).
    Fault(Fault),
    /// A bounded allocation failed.
    Allocation {
        /// Requested amount.
        amount: usize,
    },
}

impl From<Fault> for Error {
    fn from(fault: Fault) -> Self {
        match fault {
            Fault::Invalid => Self::InvalidSource,
            Fault::Allocation { amount } => Self::Allocation { amount },
            other => Self::Fault(other),
        }
    }
}

impl From<Unsupported> for Error {
    fn from(reason: Unsupported) -> Self {
        Self::Refused(Refusal::Unsupported(reason))
    }
}

fn wire(error: litchi_iwa_common::Error) -> Error {
    crate::structure::map_wire(&error).into()
}

/// The table drawable and model of one table.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Target {
    /// The table drawable (`TST.TableInfoArchive`).
    pub info: u64,
    /// The table model (`TST.TableModelArchive`).
    pub model: u64,
}

/// The per-leaf policies of one execution.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[non_exhaustive]
pub struct Hooks {
    /// Keep the root previews (Pages and Numbers, as the host does); a leaf
    /// that invalidates its rendering (Keynote) drops them.
    pub keep_previews: bool,
    /// Refuse a permutation that moves a row record into another storage
    /// tile (Keynote's physical owner admits only same-tile moves).
    pub same_tile_only: bool,
}

impl Hooks {
    /// The hooks of a leaf that keeps root previews and moves rows across
    /// tiles (Pages and Numbers).
    #[must_use]
    pub const fn keeping_previews() -> Self {
        Self {
            keep_previews: true,
            same_tile_only: false,
        }
    }
}

/// The result of one execution.
pub struct Executed<P> {
    /// The published candidate; `None` when the rows were already in order
    /// (or the range holds fewer than two rows).
    pub published: Option<Published<P>>,
    /// Body rows that changed position.
    pub moved: usize,
    /// The destination of every table row (`destination[row]`).
    pub destinations: Vec<usize>,
}

/// One stored row of a tile: its local index and raw record.
struct TileRow<'a> {
    local: u64,
    raw: &'a [u8],
}

/// The BNC cells of one row record, by column; `None` for a legacy
/// (pre-BNC) row.
///
/// # Errors
///
/// A malformed row record.
pub fn row_cells(row: &[u8]) -> Result<Option<Vec<Option<&[u8]>>>, Error> {
    let (Some(storage), Some(offsets)) = (
        single_bytes(row, ROW_STORAGE_FIELD)?,
        single_bytes(row, ROW_OFFSETS_FIELD)?,
    ) else {
        return Ok(None);
    };
    let unit = if single_varint(row, ROW_WIDE_OFFSETS_FIELD)?.unwrap_or(0) != 0 {
        4
    } else {
        1
    };
    if offsets.len() % 2 != 0 {
        return Err(Error::InvalidSource);
    }
    let starts = offsets
        .chunks_exact(2)
        .map(|pair| {
            let raw = u16::from_le_bytes([pair[0], pair[1]]);
            (raw != u16::MAX).then(|| usize::from(raw) * unit)
        })
        .collect::<Vec<_>>();
    let mut present = starts.iter().flatten().copied().collect::<Vec<_>>();
    if present.windows(2).any(|pair| pair[0] >= pair[1])
        || present.last().is_some_and(|last| *last >= storage.len())
    {
        return Err(Error::InvalidSource);
    }
    present.push(storage.len());
    let mut cells = Vec::with_capacity(starts.len());
    for start in starts {
        cells.push(start.map(|start| {
            let end = present
                .iter()
                .copied()
                .find(|candidate| *candidate > start)
                .unwrap_or(storage.len());
            &storage[start..end]
        }));
    }
    Ok(Some(cells))
}

#[derive(Clone, Copy, Debug, PartialEq)]
enum Scalar {
    Number(f64),
    Text(u32),
    Boolean(bool),
    Date(f64),
    Duration(f64),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
enum ScalarKind {
    Number,
    Text,
    Boolean,
    Date,
    Duration,
}

impl Scalar {
    const fn kind(self) -> ScalarKind {
        match self {
            Self::Number(_) => ScalarKind::Number,
            Self::Text(_) => ScalarKind::Text,
            Self::Boolean(_) => ScalarKind::Boolean,
            Self::Date(_) => ScalarKind::Date,
            Self::Duration(_) => ScalarKind::Duration,
        }
    }

    fn compare(self, other: Self, texts: &HashMap<u32, String>) -> std::cmp::Ordering {
        use std::cmp::Ordering;
        let kinds = self.kind().cmp(&other.kind());
        if kinds != Ordering::Equal {
            return kinds;
        }
        let numbers = |left: f64, right: f64| {
            if left < right {
                Ordering::Less
            } else if left > right {
                Ordering::Greater
            } else {
                Ordering::Equal
            }
        };
        match (self, other) {
            (Self::Number(left), Self::Number(right))
            | (Self::Date(left), Self::Date(right))
            | (Self::Duration(left), Self::Duration(right)) => numbers(left, right),
            (Self::Text(left), Self::Text(right)) => match (texts.get(&left), texts.get(&right)) {
                (Some(left), Some(right)) => left.cmp(right),
                _ => left.cmp(&right),
            },
            (Self::Boolean(left), Self::Boolean(right)) => left.cmp(&right),
            _ => Ordering::Equal,
        }
    }
}

fn movable_scalar(cell: &[u8], key: bool) -> Result<Option<Scalar>, Error> {
    let limits = litchi_numbers_wire::Limits::default();
    let cell = BncCellView::parse(cell, &limits).map_err(|_error| Error::InvalidSource)?;
    if cell.formula_error_identifier().is_some() {
        return Err(Unsupported::ErrorCell.into());
    }
    match cell.stored_value() {
        StoredValue::Formula(_) => return Err(Unsupported::Formula.into()),
        StoredValue::Error => return Err(Unsupported::ErrorCell.into()),
        StoredValue::Unsupported(_) => return Err(Unsupported::CellType.into()),
        _ => {},
    }
    if !key {
        return Ok(None);
    }
    if let StoredValue::Text(identifier) = cell.stored_value() {
        return Ok(Some(Scalar::Text(identifier)));
    }
    Ok(Some(match cell.cached_scalar() {
        Some(CachedScalar::Number(value)) => Scalar::Number(value.get()),
        Some(CachedScalar::Boolean(value)) => Scalar::Boolean(value),
        Some(CachedScalar::Date(value)) => Scalar::Date(value.get()),
        Some(CachedScalar::Duration(value)) => Scalar::Duration(value.get()),
        Some(CachedScalar::Unsupported(_)) => return Err(Unsupported::CellType.into()),
        None => return Err(Unsupported::MissingKey.into()),
    }))
}

/// The strings of a string table data list, by key.
fn strings(payload: &[u8]) -> Result<HashMap<u32, String>, Error> {
    let mut map = HashMap::new();
    for entry in repeated_bytes(payload, LIST_ENTRIES_FIELD)? {
        let key = single_varint(entry, LIST_ENTRY_KEY_FIELD)?.ok_or(Error::InvalidSource)?;
        let Some(text) = single_bytes(entry, LIST_ENTRY_STRING_FIELD)? else {
            continue;
        };
        let text = std::str::from_utf8(text).map_err(|_error| Error::InvalidSource)?;
        map.insert(
            u32::try_from(key).map_err(|_error| Error::InvalidSource)?,
            text.to_owned(),
        );
    }
    Ok(map)
}

/// The destination row of `row` under the plan (rows outside the sorted
/// range keep their index).
fn relocated(row: usize, start: usize, destinations: &[usize]) -> usize {
    row.checked_sub(start)
        .and_then(|offset| destinations.get(offset).copied())
        .unwrap_or(row)
}

/// The destination of every source row of a permutation.
fn destinations_by_source(permutation: &RowPermutation, start: usize) -> Result<Vec<usize>, Error> {
    let rows = permutation.len();
    let mut destinations = vec![usize::MAX; rows];
    for (destination, source) in permutation.sources_by_destination().iter().enumerate() {
        let source = source.get();
        if destinations.get(source) != Some(&usize::MAX) {
            return Err(Error::InvalidSource);
        }
        destinations[source] = start + destination;
    }
    if destinations.contains(&usize::MAX) {
        return Err(Error::InvalidSource);
    }
    Ok(destinations)
}

/// A tile payload with its rows relocated (the host's
/// `reorder_body_table_tile_rows`, one tile at a time): every row record
/// moves to the tile and local index of its destination.
fn tile_payload(
    original: &[u8],
    rows: &[(usize, Vec<u8>)],
    tile_size: usize,
) -> Result<Vec<u8>, Error> {
    let numrows = rows
        .iter()
        .map(|(row, _)| (row % tile_size) as u64 + 1)
        .max()
        .unwrap_or(0);
    let data =
        patch_varint_field(original, TILE_NUMROWS_FIELD, true, Some(numrows)).map_err(wire)?;
    let records = rows.iter().map(|(_, raw)| raw.clone()).collect::<Vec<_>>();
    rewrite_repeated_length_delimited_fields(&data, TILE_ROWS_FIELD, &records).map_err(wire)
}

fn declares_hidden(model: &[u8]) -> Result<bool, Error> {
    if single_bytes(model, MODEL_HIDDEN_STATES_OWNER_FIELD)?.is_some()
        || single_bytes(model, MODEL_HIDDEN_ROW_OWNER_FIELD)?.is_some()
        || single_bytes(model, MODEL_HIDDEN_COLUMN_OWNER_FIELD)?.is_some()
    {
        return Ok(true);
    }
    for field in [
        MODEL_HIDDEN_ROWS_FIELD,
        MODEL_HIDDEN_COLUMNS_FIELD,
        MODEL_USER_HIDDEN_ROWS_FIELD,
        MODEL_USER_HIDDEN_COLUMNS_FIELD,
    ] {
        if single_varint(model, field)?.unwrap_or(0) != 0 {
            return Ok(true);
        }
    }
    Ok(false)
}

/// One planned message replacement.
struct Rewrite {
    object: u64,
    index: usize,
    data: Vec<u8>,
}

/// Execute `order` over the body rows of the located table (`rows`: a
/// body-relative range, or every body row), moving the table's user-hidden
/// rows back to their positions. `hidden_rows` reads those positions; it is
/// called only when the table declares hidden axes and the rows move, after
/// every content refusal, so a leaf whose hidden-axes reader refuses a graph
/// reports the executor's refusal first.
///
/// The leaf has resolved the table, read `order` and checked its scope and
/// the table's merges.
///
/// # Errors
///
/// A typed [`Refusal`], a malformed source or a structural fault.
pub fn execute<P: TableLeaf>(
    source: &P,
    target: &Target,
    order: &Order,
    rows: Option<RowRange>,
    hidden_rows: impl FnOnce() -> Result<Option<HashSet<usize>>, Error>,
    hooks: Hooks,
) -> Result<Executed<P>, Error> {
    let mut working = Working::new(source)?;
    if hooks.keep_previews {
        working.keep_previews();
    }
    let unchanged = |moved| Executed {
        published: None,
        moved,
        destinations: Vec::new(),
    };

    // The model.
    let (_, model) = message(&working, target.model, &[TABLE_MODEL_MESSAGE_TYPE])?;
    let total_rows =
        usize::try_from(single_varint(model, MODEL_ROWS_FIELD)?.ok_or(Error::InvalidSource)?)
            .map_err(|_error| Error::InvalidSource)?;
    let columns =
        usize::try_from(single_varint(model, MODEL_COLUMNS_FIELD)?.ok_or(Error::InvalidSource)?)
            .map_err(|_error| Error::InvalidSource)?;
    let header_rows = single_varint(model, MODEL_HEADER_ROWS_FIELD)?.unwrap_or(0) as usize;
    let footer_rows = single_varint(model, MODEL_FOOTER_ROWS_FIELD)?.unwrap_or(0) as usize;
    let body_start = header_rows;
    let body_end = total_rows
        .checked_sub(footer_rows)
        .filter(|end| *end >= body_start)
        .ok_or(Error::InvalidSource)?;
    let (start, end) = match rows {
        None => (body_start, body_end),
        Some(range) => {
            if range.end() > body_end - body_start {
                return Err(Error::Refused(Refusal::RowsOutOfRange));
            }
            (body_start + range.start(), body_start + range.end())
        },
    };
    if order
        .rules()
        .iter()
        .any(|rule| rule.column().get() >= columns)
    {
        return Err(Error::InvalidSource);
    }
    if end.saturating_sub(start) < 2 {
        return Ok(unchanged(0));
    }

    // Feature refusals (the host's `validate_sort_features`).
    if single_varint(model, MODEL_FILTERED_ROWS_FIELD)?.unwrap_or(0) != 0
        || single_bytes(model, MODEL_PIVOT_OWNER_FIELD)?.is_some()
        || deprecated_grouping(model)?
        || filter_has_rows(&working, model)?
        || grouping_enabled(&working, model)?
        || spill_state(&working, target.info)?
    {
        return Err(Unsupported::FilterOrGrouping.into());
    }
    let declares_hidden = declares_hidden(model)?;
    let sidecar = single_reference(model, MODEL_STROKE_SIDECAR_FIELD)?;
    let uid_map = single_reference(model, MODEL_UID_MAP_FIELD)?.ok_or(Unsupported::NoRowUids)?;
    let store = single_bytes(model, MODEL_DATA_STORE_FIELD)?.ok_or(Error::InvalidSource)?;
    if let Some(conditional) = single_reference(store, STORE_CONDITIONAL_STYLE_FIELD)? {
        let (_, payload) = message(&working, conditional, &DATA_LIST_MESSAGE_TYPES)?;
        if !repeated_bytes(payload, LIST_ENTRIES_FIELD)?.is_empty() {
            return Err(Unsupported::ConditionalStyles.into());
        }
    }

    // Tiles.
    let tiles = single_bytes(store, STORE_TILES_FIELD)?.ok_or(Error::InvalidSource)?;
    let tile_size = single_varint(tiles, TILE_STORAGE_SIZE_FIELD)?.unwrap_or(DEFAULT_TILE_SIZE);
    let tile_size = usize::try_from(tile_size)
        .ok()
        .filter(|size| *size != 0)
        .ok_or(Error::InvalidSource)?;
    let mut tile_objects = Vec::new();
    for reference_payload in repeated_bytes(tiles, TILE_STORAGE_TILES_FIELD)? {
        let key = single_varint(reference_payload, TILE_REFERENCE_ID_FIELD)?
            .ok_or(Error::InvalidSource)?;
        let tile = single_reference(reference_payload, TILE_REFERENCE_TILE_FIELD)?
            .ok_or(Error::InvalidSource)?;
        tile_objects.push((
            usize::try_from(key).map_err(|_error| Error::InvalidSource)?,
            tile,
        ));
    }
    let tile_keys = tile_objects
        .iter()
        .map(|(key, _)| *key)
        .collect::<HashSet<_>>();
    if tile_keys.len() != tile_objects.len() {
        return Err(Error::InvalidSource);
    }

    // Keys of every body row in the range.
    let rule_columns = order
        .rules()
        .iter()
        .map(|rule| rule.column().get())
        .collect::<Vec<_>>();
    let mut keys: Vec<Option<Vec<Scalar>>> = vec![None; end - start];
    let mut tile_rows: Vec<(usize, u64, Vec<TileRow<'_>>)> = Vec::new();
    for (key, tile) in &tile_objects {
        let (_, payload) = message(&working, *tile, &[TILE_MESSAGE_TYPE])?;
        let mut rows_of_tile = Vec::new();
        let mut locals = HashSet::new();
        for raw in repeated_bytes(payload, TILE_ROWS_FIELD)? {
            let local = single_varint(raw, ROW_INDEX_FIELD)?.ok_or(Error::InvalidSource)?;
            if local >= tile_size as u64 || !locals.insert(local) {
                return Err(Error::InvalidSource);
            }
            let global = key * tile_size + local as usize;
            if global >= total_rows {
                return Err(Error::InvalidSource);
            }
            if (start..end).contains(&global) {
                let mut cells = row_cells(raw)?.ok_or(Unsupported::Layout)?;
                // Numbers keeps a fixed-width offset table: slots past the
                // last column must be empty.
                if cells.iter().skip(columns).any(Option::is_some) {
                    return Err(Error::InvalidSource);
                }
                cells.truncate(columns);
                let mut values = vec![None; rule_columns.len()];
                for (column, cell) in cells.into_iter().enumerate() {
                    let Some(cell) = cell else {
                        continue;
                    };
                    let position = rule_columns.iter().position(|rule| *rule == column);
                    let scalar = movable_scalar(cell, position.is_some())?;
                    if let (Some(position), Some(scalar)) = (position, scalar) {
                        values[position] = Some(scalar);
                    }
                }
                let values = values
                    .into_iter()
                    .collect::<Option<Vec<_>>>()
                    .ok_or(Unsupported::MissingKey)?;
                if keys[global - start].replace(values).is_some() {
                    return Err(Error::InvalidSource);
                }
            }
            rows_of_tile.push(TileRow { local, raw });
        }
        tile_rows.push((*key, *tile, rows_of_tile));
    }
    let keys = keys
        .into_iter()
        .collect::<Option<Vec<_>>>()
        .ok_or(Unsupported::MissingKey)?;
    let texts = if keys
        .iter()
        .flatten()
        .any(|key| matches!(key, Scalar::Text(_)))
    {
        let table =
            single_reference(store, STORE_STRING_TABLE_FIELD)?.ok_or(Error::InvalidSource)?;
        let (_, payload) = message(&working, table, &DATA_LIST_MESSAGE_TYPES)?;
        strings(payload)?
    } else {
        HashMap::new()
    };
    if keys.iter().flatten().any(|key| match key {
        Scalar::Text(identifier) => !texts.contains_key(identifier),
        _ => false,
    }) {
        return Err(Error::InvalidSource);
    }
    let permutation = stable_body_row_permutation(
        &keys,
        order,
        |scalar| scalar.kind(),
        |left, right| left.compare(*right, &texts),
    )
    .map_err(|error| match error {
        litchi_iwa_common::table::sort::planning::Error::MixedScalarKinds { .. } => {
            Error::from(Unsupported::MixedKinds)
        },
        litchi_iwa_common::table::sort::planning::Error::Allocation { rows } => {
            Error::Allocation { amount: rows }
        },
        _ => Error::InvalidSource,
    })?;
    if !permutation.changes_rows() {
        return Ok(unchanged(0));
    }
    let destinations = destinations_by_source(&permutation, start)?;
    let moved = destinations
        .iter()
        .enumerate()
        .filter(|(offset, destination)| start + offset != **destination)
        .count();
    if hooks.same_tile_only
        && destinations
            .iter()
            .enumerate()
            .any(|(offset, destination)| (start + offset) / tile_size != destination / tile_size)
    {
        return Err(Unsupported::Layout.into());
    }

    let mut rewrites = Vec::new();

    // Tile rows.
    let mut relocated_rows: Vec<(usize, Vec<u8>)> = Vec::new();
    for (key, _, rows_of_tile) in &tile_rows {
        for row in rows_of_tile {
            let source_row = key * tile_size + row.local as usize;
            let destination = relocated(source_row, start, &destinations);
            let raw = patch_varint_field(
                row.raw,
                ROW_INDEX_FIELD,
                true,
                Some((destination % tile_size) as u64),
            )
            .map_err(wire)?;
            relocated_rows.push((destination, raw));
        }
    }
    relocated_rows.sort_by_key(|(row, _)| *row);
    if relocated_rows.windows(2).any(|pair| pair[0].0 == pair[1].0)
        || relocated_rows
            .iter()
            .any(|(row, _)| !tile_keys.contains(&(row / tile_size)))
    {
        return Err(Unsupported::Layout.into());
    }
    for (key, tile, _) in &tile_rows {
        let (index, payload) = message(&working, *tile, &[TILE_MESSAGE_TYPE])?;
        let rows_here = relocated_rows
            .iter()
            .filter(|(row, _)| row / tile_size == *key)
            .cloned()
            .collect::<Vec<_>>();
        let data = tile_payload(payload, &rows_here, tile_size)?;
        if data != payload {
            rewrites.push(Rewrite {
                object: *tile,
                index,
                data,
            });
        }
    }

    // Stroke layers: explicit cell borders stay with their cells.
    if let Some(sidecar) = sidecar {
        rewrites.extend(stroke_rewrites(
            &working,
            sidecar,
            total_rows,
            start,
            &destinations,
        )?);
    }

    // Row header buckets.
    let row_headers = single_bytes(store, STORE_ROW_HEADERS_FIELD)?.ok_or(Error::InvalidSource)?;
    let mut buckets = Vec::new();
    for bucket in repeated_bytes(row_headers, HEADER_STORAGE_BUCKETS_FIELD)? {
        buckets.push(crate::structure::references::reference_identifier(bucket)?);
    }
    let mut headers: Vec<(usize, Vec<u8>)> = Vec::new();
    for bucket in &buckets {
        let (_, payload) = message(&working, *bucket, &[HEADER_BUCKET_MESSAGE_TYPE])?;
        for raw in repeated_bytes(payload, BUCKET_HEADERS_FIELD)? {
            let row = usize::try_from(
                single_varint(raw, HEADER_INDEX_FIELD)?.ok_or(Error::InvalidSource)?,
            )
            .map_err(|_error| Error::InvalidSource)?;
            if row >= total_rows {
                return Err(Error::InvalidSource);
            }
            let destination = relocated(row, start, &destinations);
            headers.push((
                destination,
                patch_varint_field(raw, HEADER_INDEX_FIELD, true, Some(destination as u64))
                    .map_err(wire)?,
            ));
        }
    }
    headers.sort_by_key(|(row, _)| *row);
    if headers.windows(2).any(|pair| pair[0].0 == pair[1].0)
        || headers
            .iter()
            .any(|(row, _)| row / HEADER_BUCKET_ROWS >= buckets.len())
    {
        return Err(Unsupported::Layout.into());
    }
    for (position, bucket) in buckets.iter().enumerate() {
        let (index, payload) = message(&working, *bucket, &[HEADER_BUCKET_MESSAGE_TYPE])?;
        let desired = headers
            .iter()
            .filter(|(row, _)| row / HEADER_BUCKET_ROWS == position)
            .map(|(_, raw)| raw.clone())
            .collect::<Vec<_>>();
        let data =
            rewrite_repeated_length_delimited_fields(payload, BUCKET_HEADERS_FIELD, &desired)
                .map_err(wire)?;
        if data != payload {
            rewrites.push(Rewrite {
                object: *bucket,
                index,
                data,
            });
        }
    }

    // The row UID map.
    let physical_row_uids: Vec<Vec<u8>>;
    {
        let (index, payload) = message(&working, uid_map, &UID_MAP_MESSAGE_TYPES)?;
        let uid_for_index = repeated_varints(payload, UID_ROW_UID_FOR_INDEX_FIELD)?;
        let index_for_uid = repeated_varints(payload, UID_ROW_INDEX_FOR_UID_FIELD)?;
        let sorted_uids = repeated_bytes(payload, UID_SORTED_ROWS_FIELD)?;
        if uid_for_index.len() != total_rows
            || index_for_uid.len() != total_rows
            || sorted_uids.len() != total_rows
        {
            return Err(Unsupported::NoRowUids.into());
        }
        let mut current = uid_for_index.clone();
        for (offset, destination) in destinations.iter().copied().enumerate() {
            current[destination] = uid_for_index[start + offset];
        }
        let mut inverse = vec![0_u64; total_rows];
        for (row, uid) in current.iter().copied().enumerate() {
            let uid = usize::try_from(uid)
                .ok()
                .filter(|uid| *uid < total_rows)
                .ok_or(Error::InvalidSource)?;
            inverse[uid] = row as u64;
        }
        physical_row_uids = current
            .iter()
            .map(|uid| {
                usize::try_from(*uid)
                    .ok()
                    .and_then(|uid| sorted_uids.get(uid))
                    .map(|raw| raw.to_vec())
                    .ok_or(Error::InvalidSource)
            })
            .collect::<Result<Vec<_>, _>>()?;
        let data = rewrite_repeated_varint_fields(payload, UID_ROW_INDEX_FOR_UID_FIELD, &inverse)
            .map_err(wire)?;
        let data = rewrite_repeated_varint_fields(&data, UID_ROW_UID_FOR_INDEX_FIELD, &current)
            .map_err(wire)?;
        rewrites.push(Rewrite {
            object: uid_map,
            index,
            data,
        });
    }
    let hidden = if declares_hidden {
        hidden_rows()?.filter(|rows| !rows.is_empty())
    } else {
        None
    };
    if let Some(rows) = &hidden {
        let (index, model) = message(&working, target.model, &[TABLE_MODEL_MESSAGE_TYPE])?;
        let (_, info) = message(&working, target.info, &TABLE_INFO_MESSAGE_TYPES)?;
        if let Some(data) = restored_hidden_rows(model, info, &physical_row_uids, rows)? {
            rewrites.push(Rewrite {
                object: target.model,
                index,
                data,
            });
        }
    }
    let mut full = Vec::with_capacity(total_rows);
    full.extend((0..total_rows).map(|row| relocated(row, start, &destinations)));
    for rewrite in rewrites {
        replace_payload(&mut working, rewrite.object, rewrite.index, rewrite.data).map_err(
            |error| match error {
                crate::topology::Error::Fault(fault) => Error::Fault(fault),
                crate::topology::Error::Allocation { amount } => Error::Allocation { amount },
                _ => Error::InvalidSource,
            },
        )?;
    }
    let published = working.publish(&Transition::default())?;
    Ok(Executed {
        published: Some(published),
        moved,
        destinations: full,
    })
}

const FILTER_ENABLED_FIELD: u32 = 2;
const FILTER_RULES_PREPIVOT_FIELD: u32 = 3;
const FILTER_ENABLED_FLAGS_FIELD: u32 = 6;
const FILTER_RULES_FIELD: u32 = 7;
const CATEGORY_GROUP_BY_FIELD: u32 = 1;
const DEPRECATED_CATEGORY_GROUP_BY_FIELD: u32 = 2;
const GROUP_BY_ENABLED_FIELD: u32 = 6;
const OWNER_FORMULA_OWNER_FIELD: u32 = 11;
const OWNER_SPILL_SIZES_FIELD: u32 = 16;
const SPILL_SIZES_SPILLS_FIELD: u32 = 1;

/// The first message payload of `identifier`.
fn first_payload<'w, P: TableLeaf>(
    working: &'w Working<'_, P>,
    identifier: u64,
) -> Result<&'w [u8], Error> {
    Ok(working
        .object(identifier)?
        .messages
        .first()
        .ok_or(Error::InvalidSource)?
        .data
        .as_slice())
}

/// Whether the model's pre-pivot row filter set has row state.
fn filter_has_rows<P: TableLeaf>(working: &Working<'_, P>, model: &[u8]) -> Result<bool, Error> {
    let Some(filter) = single_reference(model, MODEL_ROW_FILTER_FIELD)? else {
        return Ok(false);
    };
    let payload = first_payload(working, filter)?;
    let enabled = single_varint(payload, FILTER_ENABLED_FIELD)?.unwrap_or(1) != 0;
    let flags = repeated_varints(payload, FILTER_ENABLED_FLAGS_FIELD)?;
    Ok(enabled
        && (has_field(payload, FILTER_RULES_PREPIVOT_FIELD)?
            || has_field(payload, FILTER_RULES_FIELD)?
            || flags.iter().any(|flag| *flag != 0)))
}

/// Whether the category owner enables a group-by.
fn grouping_enabled<P: TableLeaf>(working: &Working<'_, P>, model: &[u8]) -> Result<bool, Error> {
    let Some(owner) = single_reference(model, MODEL_CATEGORY_OWNER_FIELD)? else {
        return Ok(false);
    };
    let payload = first_payload(working, owner)?;
    for group in repeated_bytes(payload, CATEGORY_GROUP_BY_FIELD)? {
        let group = crate::structure::references::reference_identifier(group)?;
        for message in &working.object(group)?.messages {
            if single_varint(&message.data, GROUP_BY_ENABLED_FIELD)?.unwrap_or(0) != 0 {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

/// Whether the deprecated inline category owner enables a group-by.
fn deprecated_grouping(model: &[u8]) -> Result<bool, Error> {
    let Some(owner) = single_bytes(model, MODEL_DEPRECATED_CATEGORY_OWNER_FIELD)? else {
        return Ok(false);
    };
    for group in repeated_bytes(owner, DEPRECATED_CATEGORY_GROUP_BY_FIELD)? {
        if single_varint(group, GROUP_BY_ENABLED_FIELD)?.unwrap_or(0) != 0 {
            return Ok(true);
        }
    }
    Ok(false)
}

/// Whether a formula owner of the table drawable holds spill sizes.
fn spill_state<P: TableLeaf>(working: &Working<'_, P>, drawable: u64) -> Result<bool, Error> {
    for object in working
        .source()
        .catalog()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
    {
        for message in &object.messages {
            if message.type_ != FORMULA_OWNER_DEPENDENCIES_MESSAGE_TYPE {
                continue;
            }
            if single_reference(&message.data, OWNER_FORMULA_OWNER_FIELD)? != Some(drawable) {
                continue;
            }
            if let Some(spills) = single_bytes(&message.data, OWNER_SPILL_SIZES_FIELD)? {
                if !repeated_bytes(spills, SPILL_SIZES_SPILLS_FIELD)?.is_empty() {
                    return Ok(true);
                }
            }
        }
    }
    Ok(false)
}

const INFO_HIDDEN_STATES_UUID_FIELD: u32 = 8;
const OWNER_HIDDEN_STATES_FIELD: u32 = 2;
const STATES_UID_FIELD: u32 = 1;
const STATES_ROW_EXTENT_FIELD: u32 = 3;
const EXTENT_BASE_STATES_FIELD: u32 = 2;
const STATE_UID_FIELD: u32 = 1;
const STATE_USER_HIDDEN_FIELD: u32 = 2;
const STATE_FILTERED_FIELD: u32 = 3;
const STATE_PIVOT_HIDDEN_FIELD: u32 = 4;

/// The model payload with user-hidden rows back at `rows` after a row
/// permutation (the host's `restore_positional_user_hidden_axes`, rows
/// only): the active row extent's states are re-keyed by the rows' new
/// physical UIDs. `None` when nothing changes.
fn restored_hidden_rows(
    model: &[u8],
    info: &[u8],
    physical_uids: &[Vec<u8>],
    rows: &HashSet<usize>,
) -> Result<Option<Vec<u8>>, Error> {
    let Some(owner) = single_bytes(model, MODEL_HIDDEN_STATES_OWNER_FIELD)? else {
        return Err(Unsupported::HiddenAxes.into());
    };
    let active =
        single_bytes(info, INFO_HIDDEN_STATES_UUID_FIELD)?.ok_or(Unsupported::HiddenAxes)?;
    let states = repeated_bytes(owner, OWNER_HIDDEN_STATES_FIELD)?;
    if states.len() != 1 {
        return Err(Unsupported::HiddenAxes.into());
    }
    if single_bytes(states[0], STATES_UID_FIELD)? != Some(active) {
        return Err(Error::InvalidSource);
    }
    let extent = single_bytes(states[0], STATES_ROW_EXTENT_FIELD)?.ok_or(Error::InvalidSource)?;
    let positions = physical_uids
        .iter()
        .enumerate()
        .map(|(index, uid)| (uid.as_slice(), index))
        .collect::<HashMap<_, _>>();
    let flag = |state: &[u8], number| -> Result<bool, Error> {
        Ok(single_varint(state, number)?.is_some_and(|value| value != 0))
    };
    let mut present = HashSet::new();
    let mut kept: Vec<(usize, Vec<u8>)> = Vec::new();
    for state in repeated_bytes(extent, EXTENT_BASE_STATES_FIELD)? {
        let uid = single_bytes(state, STATE_UID_FIELD)?.ok_or(Error::InvalidSource)?;
        if flag(state, STATE_FILTERED_FIELD)? || flag(state, STATE_PIVOT_HIDDEN_FIELD)? {
            return Err(Unsupported::FilterOrGrouping.into());
        }
        let Some(position) = positions.get(uid).copied() else {
            kept.push((usize::MAX, state.to_vec()));
            continue;
        };
        if rows.contains(&position) {
            present.insert(position);
            kept.push((
                position,
                patch_varint_field(
                    state,
                    STATE_USER_HIDDEN_FIELD,
                    single_varint(state, STATE_USER_HIDDEN_FIELD)?.is_some(),
                    Some(1),
                )
                .map_err(wire)?,
            ));
        }
    }
    let mut desired = rows.iter().copied().collect::<Vec<_>>();
    desired.sort_unstable();
    for position in desired {
        if present.contains(&position) {
            continue;
        }
        let uid = physical_uids.get(position).ok_or(Error::InvalidSource)?;
        let mut state = Vec::new();
        litchi_iwa_common::wire::append_length_delimited_field(
            &mut state,
            STATE_UID_FIELD,
            uid,
            litchi_iwa_common::WireLimits::default(),
        )
        .map_err(wire)?;
        litchi_iwa_common::wire::append_varint_field(&mut state, STATE_USER_HIDDEN_FIELD, 1)
            .map_err(wire)?;
        kept.push((position, state));
    }
    kept.sort_by_key(|(position, _)| *position);
    let records = kept.into_iter().map(|(_, raw)| raw).collect::<Vec<_>>();
    let new_extent =
        rewrite_repeated_length_delimited_fields(extent, EXTENT_BASE_STATES_FIELD, &records)
            .map_err(wire)?;
    if new_extent == extent {
        return Ok(None);
    }
    let new_state = litchi_iwa_common::wire::patch_length_delimited_field(
        states[0],
        STATES_ROW_EXTENT_FIELD,
        true,
        Some(&new_extent),
    )
    .map_err(wire)?;
    let new_owner =
        rewrite_repeated_length_delimited_fields(owner, OWNER_HIDDEN_STATES_FIELD, &[new_state])
            .map_err(wire)?;
    litchi_iwa_common::wire::patch_length_delimited_field(
        model,
        MODEL_HIDDEN_STATES_OWNER_FIELD,
        true,
        Some(&new_owner),
    )
    .map(Some)
    .map_err(wire)
}

const SIDECAR_LEFT_COLUMN_FIELD: u32 = 4;
const SIDECAR_TOP_ROW_FIELD: u32 = 6;
const SIDECAR_BOTTOM_ROW_FIELD: u32 = 7;
const LAYER_INDEX_FIELD: u32 = 1;
const LAYER_RUNS_FIELD: u32 = 2;
const RUN_ORIGIN_FIELD: u32 = 1;
const RUN_LENGTH_FIELD: u32 = 2;

/// One run of a column layer split into the destination ranges of its rows.
fn split_run(
    run: &[u8],
    rows: usize,
    start: usize,
    destinations: &[usize],
) -> Result<Vec<Vec<u8>>, Error> {
    let origin = single_varint(run, RUN_ORIGIN_FIELD)?.ok_or(Error::InvalidSource)?;
    let length = single_varint(run, RUN_LENGTH_FIELD)?.ok_or(Error::InvalidSource)?;
    let origin = usize::try_from(origin).map_err(|_error| Error::InvalidSource)?;
    let length = usize::try_from(length).map_err(|_error| Error::InvalidSource)?;
    if length == 0 || origin.checked_add(length).is_none_or(|end| end > rows) {
        return Err(Error::InvalidSource);
    }
    let mut moved = (origin..origin + length)
        .map(|row| relocated(row, start, destinations))
        .collect::<Vec<_>>();
    moved.sort_unstable();
    let mut ranges = Vec::new();
    let mut first = moved[0];
    let mut last = moved[0];
    for destination in moved.into_iter().skip(1) {
        if last + 1 == destination {
            last = destination;
            continue;
        }
        ranges.push((first, last));
        first = destination;
        last = destination;
    }
    ranges.push((first, last));
    ranges
        .into_iter()
        .map(|(first, last)| {
            let patched = patch_varint_field(run, RUN_ORIGIN_FIELD, true, Some(first as u64))
                .map_err(wire)?;
            patch_varint_field(
                &patched,
                RUN_LENGTH_FIELD,
                true,
                Some((last - first + 1) as u64),
            )
            .map_err(wire)
        })
        .collect()
}

/// The stroke-layer rewrites of a row permutation (the host's
/// `stroke_layers::reorder_rows`): a row layer moves with its row, a column
/// layer's runs split into the destination ranges of their rows.
fn stroke_rewrites<P: TableLeaf>(
    working: &Working<'_, P>,
    sidecar: u64,
    rows: usize,
    start: usize,
    destinations: &[usize],
) -> Result<Vec<Rewrite>, Error> {
    let (_, payload) = message(working, sidecar, &[STROKE_SIDECAR_MESSAGE_TYPE])?;
    let mut rewrites = Vec::new();
    let mut seen = HashSet::new();
    for field in SIDECAR_LEFT_COLUMN_FIELD..=SIDECAR_BOTTOM_ROW_FIELD {
        for layer in repeated_bytes(payload, field)? {
            let layer = crate::structure::references::reference_identifier(layer)?;
            if !seen.insert(layer) {
                return Err(Error::InvalidSource);
            }
            let (index, data) = message(working, layer, &[STROKE_LAYER_MESSAGE_TYPE])?;
            let rewritten = if matches!(field, SIDECAR_TOP_ROW_FIELD | SIDECAR_BOTTOM_ROW_FIELD) {
                let current =
                    single_varint(data, LAYER_INDEX_FIELD)?.ok_or(Error::InvalidSource)?;
                let current = usize::try_from(current)
                    .ok()
                    .filter(|current| *current <= rows)
                    .ok_or(Error::InvalidSource)?;
                let moved = relocated(current, start, destinations);
                patch_varint_field(data, LAYER_INDEX_FIELD, true, Some(moved as u64))
                    .map_err(wire)?
            } else {
                let mut runs = Vec::new();
                for run in repeated_bytes(data, LAYER_RUNS_FIELD)? {
                    runs.extend(split_run(run, rows, start, destinations)?);
                }
                rewrite_repeated_length_delimited_fields(data, LAYER_RUNS_FIELD, &runs)
                    .map_err(wire)?
            };
            if rewritten != data {
                rewrites.push(Rewrite {
                    object: layer,
                    index,
                    data: rewritten,
                });
            }
        }
    }
    Ok(rewrites)
}
