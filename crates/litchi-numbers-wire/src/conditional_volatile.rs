//! The calculation-engine owner family a volatile conditional-highlight rule
//! (today, yesterday, tomorrow, the next or last period, an offset from
//! today) needs, shared by the table owners (records 0376a, 0376b).
//!
//! The first volatile rule of a table registers twelve formula owners (the
//! conditional-style owner, its style-storage, global and global-name owners
//! and eight auxiliaries) and four cell-record tiles, in the host's object and
//! internal-identifier orders; every later volatile cell appends one record
//! to the conditional-style owner and its tile, and a removal takes the
//! record out again (the family stays, as the host keeps it). [`ensure`](crate::conditional_volatile::ensure) and
//! [`remove`](crate::conditional_volatile::remove) plan these payloads from the engine and the owner objects; the
//! owner allocates the identifiers it is handed, inserts the created objects
//! into the calculation engine's component, registers their UUIDs and
//! publishes. Unlike the host, whose global-name owner takes a random UUID,
//! the caller supplies a deterministic one.

use litchi_iwa_protos::numbers_table_cell_dependency_codec as dependency;

use crate::{
    Limits, formula_graph,
    formula_owner::Uuid,
    table_write::{fields, single_bytes, single_varint},
};

const REGISTERED_FORMULA_COUNT: u64 = 8;
const STYLE_STORAGE_OWNER_OFFSET: u64 = 35;
const CONDITIONAL_BASE_OWNER_OFFSET: u64 = 3;
const STYLE_STORAGE_ROW: u32 = 0;
const STYLE_STORAGE_COLUMN: u32 = 2;
const AUXILIARY_STYLE_ROW: u32 = 0;
const AUXILIARY_STYLE_EMPTY_END_COLUMN: u32 = 3;
const AUXILIARY_STYLE_LAST_COLUMN: u32 = 7;
const AUXILIARY_STYLE_SOURCE_COLUMN: u32 = 2;
const AUXILIARY_STYLE_STORAGE_COLUMN: u32 = 11;
const GLOBAL_OWNER_ROW: u32 = 1;
const GLOBAL_OWNER_COLUMN_COUNT: u32 = 2;
const GLOBAL_OWNER_UID: Uuid = Uuid::new(666, 0);
const GLOBAL_OWNER_BASE_UID: Uuid = Uuid::new(466, 0);
const WHOLE_ROW_COLUMN_SENTINEL: u64 = i16::MAX as u64;
const WHOLE_COLUMN_ROW_SENTINEL: u64 = i32::MAX as u64;
const TABLE_OWNER_KIND: u64 = 1;
/// `TSCE.CalculationEngineArchive.uuid_reference_map`.
const UUID_REFERENCE_MAP_FIELD: u32 = 7;

/// The host's `ArchiveInfo` message versions of the created objects.
pub const OWNER_MESSAGE_VERSION: [u32; 3] = [3, 2, 10];
/// The host's `ArchiveInfo` message version of a created cell-record tile.
pub const TILE_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];

/// A typed failure of the volatile owner-family planner.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The engine or an owner is malformed, or the family is incomplete.
    InvalidSource,
    /// The table has no formula owner.
    MissingTableOwner,
    /// The table's conditional-style owner is registered without the
    /// family's style-storage owner, volatile set or cell-record tile, as in
    /// files Numbers saved; the host refuses these too.
    IncompleteFamily,
    /// A shared-codec or engine-rewrite failure (limits included).
    Rewrite(formula_graph::Error),
    /// A fallible allocation failed.
    Allocation,
}

impl core::fmt::Display for Error {
    fn fmt(&self, formatter: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        match self {
            Self::InvalidSource => formatter.write_str("malformed volatile owner family"),
            Self::MissingTableOwner => formatter.write_str("the table has no formula owner"),
            Self::IncompleteFamily => {
                formatter.write_str("the conditional-style owner has no volatile owner family")
            },
            Self::Rewrite(error) => write!(formatter, "engine rewrite failed: {error}"),
            Self::Allocation => formatter.write_str("volatile owner family allocation failed"),
        }
    }
}

impl std::error::Error for Error {}

type Result<T> = core::result::Result<T, Error>;

/// One object's first message payload.
#[derive(Debug, Clone, Copy)]
pub struct Object<'a> {
    /// The object identifier.
    pub identifier: u64,
    /// The payload.
    pub payload: &'a [u8],
}

/// The engine, its formula owners, the table and the volatile cell.
#[derive(Debug, Clone, Copy)]
pub struct Source<'a> {
    /// The `TSCE.CalculationEngineArchive` object.
    pub engine: Object<'a>,
    /// Every `TSCE.FormulaOwnerDependenciesArchive` object of the engine.
    pub owners: &'a [Object<'a>],
    /// The table's drawable (`TST.TableInfoArchive`) identifier.
    pub table: u64,
    /// The table model's `conditional_style_formula_owner_id`.
    pub conditional_owner: Uuid,
    /// The cell row.
    pub row: u32,
    /// The cell column.
    pub column: u32,
    /// Whether the first volatile rule may give an existing family's
    /// conditional-style owner its set and tile (Numbers' own shape, record
    /// 0432). Only Numbers has native evidence.
    pub creates_first_tile: bool,
}

/// One object the family creates.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Created {
    /// The allocated identifier.
    pub identifier: u64,
    /// `TSCE.FormulaOwnerDependenciesArchive` (4008) or
    /// `TSCE.CellRecordTileArchive` (4009).
    pub message_type: u32,
    /// The payload.
    pub payload: Vec<u8>,
    /// The message's object references (an owner's tile).
    pub references: Vec<u64>,
}

/// The planned payloads.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Edit {
    /// Changed payloads of existing objects (engine, owner, tile).
    pub payloads: Vec<(u64, Vec<u8>)>,
    /// Objects to insert, in the host's insertion and UUID-registration
    /// order: each owner, then its tile.
    pub created: Vec<Created>,
    /// Owner objects the engine message now references (append to its
    /// `ArchiveInfo` references, with a `[2, 6]` object-reference field
    /// info).
    pub engine_references: Vec<u64>,
    /// `(owner, tile)`: a created tile an existing owner now references
    /// (append to the owner's `ArchiveInfo` references, with a `[13, 1]`
    /// object-reference field info).
    pub owner_references: Vec<(u64, u64)>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[repr(u32)]
enum Kind {
    ConditionalStyle = 3,
    Auxiliary4 = 4,
    Auxiliary5 = 5,
    Auxiliary6 = 6,
    GlobalName = 7,
    Auxiliary8 = 8,
    Auxiliary9 = 9,
    Auxiliary10 = 10,
    Auxiliary11 = 11,
    Auxiliary12 = 12,
    StyleStorage = 35,
    Global = 200,
}

impl Kind {
    const INTERNAL_ORDER: [Self; 12] = [
        Self::GlobalName,
        Self::Global,
        Self::Auxiliary4,
        Self::Auxiliary11,
        Self::ConditionalStyle,
        Self::Auxiliary10,
        Self::Auxiliary5,
        Self::Auxiliary6,
        Self::StyleStorage,
        Self::Auxiliary12,
        Self::Auxiliary8,
        Self::Auxiliary9,
    ];
    const OBJECT_ORDER: [Self; 12] = [
        Self::Auxiliary9,
        Self::Auxiliary8,
        Self::Auxiliary12,
        Self::StyleStorage,
        Self::Auxiliary6,
        Self::Auxiliary5,
        Self::Auxiliary10,
        Self::ConditionalStyle,
        Self::Auxiliary11,
        Self::Auxiliary4,
        Self::Global,
        Self::GlobalName,
    ];

    const fn has_tile(self) -> bool {
        matches!(
            self,
            Self::Auxiliary8 | Self::StyleStorage | Self::ConditionalStyle | Self::Global
        )
    }

    fn internal_offset(self) -> u32 {
        // Twelve kinds: the position always fits.
        Self::INTERNAL_ORDER
            .iter()
            .position(|kind| *kind == self)
            .map_or(0, |position| position as u32)
    }
}

/// The objects [`ensure`] creates for a table's first volatile rule.
pub const CREATED_OBJECTS: u64 = 16;

// ---------------------------------------------------------------------------
// Wire writing
// ---------------------------------------------------------------------------

fn put_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    put_varint(output, u64::from(number) << 3);
    put_varint(output, value);
}

fn message_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) {
    put_varint(output, (u64::from(number) << 3) | 2);
    put_varint(output, payload.len() as u64);
    output.extend_from_slice(payload);
}

fn uuid(value: Uuid) -> Vec<u8> {
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

/// A `TSCE.CellRecordExpandedArchive`.
struct Record {
    row: u32,
    column: u32,
    without_owner: Vec<(u32, u32)>,
    with_owner: Vec<(u32, u32, u32)>,
}

impl Record {
    fn encode(&self) -> Vec<u8> {
        let mut edges = Vec::new();
        for (row, _) in &self.without_owner {
            varint_field(&mut edges, 1, u64::from(*row));
        }
        for (_, column) in &self.without_owner {
            varint_field(&mut edges, 2, u64::from(*column));
        }
        for (row, _, _) in &self.with_owner {
            varint_field(&mut edges, 3, u64::from(*row));
        }
        for (_, column, _) in &self.with_owner {
            varint_field(&mut edges, 4, u64::from(*column));
        }
        for (_, _, owner) in &self.with_owner {
            varint_field(&mut edges, 5, u64::from(*owner));
        }
        let mut output = Vec::new();
        varint_field(&mut output, 1, u64::from(self.column));
        varint_field(&mut output, 2, u64::from(self.row));
        message_field(&mut output, 6, &edges);
        output
    }
}

/// A `TSCE.CellCoordSetArchive`: columns with row ranges.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
struct Coordinates(Vec<(u32, Vec<(i64, Option<i64>)>)>);

impl Coordinates {
    fn single(row: u32, column: u32) -> Self {
        Self(vec![(column, vec![(i64::from(row), None)])])
    }

    fn decode(payload: &[u8], limits: &Limits) -> Result<Self> {
        let mut columns = Vec::new();
        for entry in fields(payload, limits).map_err(|_| Error::InvalidSource)? {
            if entry.number != 1 {
                return Err(Error::InvalidSource);
            }
            let column = single_varint(entry.payload, 1, limits)
                .map_err(|_| Error::InvalidSource)?
                .and_then(|value| u32::try_from(value).ok())
                .ok_or(Error::InvalidSource)?;
            let rows = single_bytes(entry.payload, 2, limits)
                .map_err(|_| Error::InvalidSource)?
                .ok_or(Error::InvalidSource)?;
            let mut ranges = Vec::new();
            for range in fields(rows, limits).map_err(|_| Error::InvalidSource)? {
                if range.number != 1 {
                    return Err(Error::InvalidSource);
                }
                let begin = single_varint(range.payload, 1, limits)
                    .map_err(|_| Error::InvalidSource)?
                    .ok_or(Error::InvalidSource)?;
                let end =
                    single_varint(range.payload, 2, limits).map_err(|_| Error::InvalidSource)?;
                ranges.push((signed(begin), end.map(signed)));
            }
            columns.push((column, ranges));
        }
        Ok(Self(columns))
    }

    fn encode(&self) -> Vec<u8> {
        let mut output = Vec::new();
        for (column, ranges) in &self.0 {
            let mut rows = Vec::new();
            for (begin, end) in ranges {
                let mut range = Vec::new();
                varint_field(&mut range, 1, unsigned(*begin));
                if let Some(end) = end {
                    varint_field(&mut range, 2, unsigned(*end));
                }
                message_field(&mut rows, 1, &range);
            }
            let mut entry = Vec::new();
            varint_field(&mut entry, 1, u64::from(*column));
            message_field(&mut entry, 2, &rows);
            message_field(&mut output, 1, &entry);
        }
        output
    }

    fn contains(&self, row: i64, column: u32) -> bool {
        self.0.iter().any(|(candidate, ranges)| {
            *candidate == column
                && ranges
                    .iter()
                    .any(|(begin, end)| (*begin..=end.unwrap_or(*begin)).contains(&row))
        })
    }

    /// The host's `merge_coordinate_set` of one coordinate.
    fn add(&mut self, row: i64, column: u32) {
        let index = match self
            .0
            .iter()
            .position(|(candidate, _)| *candidate == column)
        {
            Some(index) => index,
            None => {
                self.0.push((column, Vec::new()));
                self.0.len() - 1
            },
        };
        let ranges = &mut self.0[index].1;
        if !ranges
            .iter()
            .any(|(begin, end)| (*begin..=end.unwrap_or(*begin)).contains(&row))
        {
            ranges.push((row, None));
        }
        ranges.sort_unstable_by_key(|(begin, _)| *begin);
        self.0.sort_unstable_by_key(|(column, _)| *column);
    }

    /// The host's `remove_coordinate`.
    fn remove(&mut self, row: i64, column: u32) -> bool {
        let Some(column_index) = self
            .0
            .iter()
            .position(|(candidate, _)| *candidate == column)
        else {
            return false;
        };
        let ranges = &mut self.0[column_index].1;
        let Some(index) = ranges
            .iter()
            .position(|(begin, end)| (*begin..=end.unwrap_or(*begin)).contains(&row))
        else {
            return false;
        };
        let (begin, end) = ranges[index];
        let end = end.unwrap_or(begin);
        let bounded = |begin: i64, end: i64| (begin, (begin != end).then_some(end));
        match (begin == row, end == row) {
            (true, true) => {
                ranges.remove(index);
            },
            (true, false) => ranges[index] = bounded(row + 1, end),
            (false, true) => ranges[index] = bounded(begin, row - 1),
            (false, false) => {
                ranges[index] = bounded(begin, row - 1);
                ranges.insert(index + 1, bounded(row + 1, end));
            },
        }
        if ranges.is_empty() {
            self.0.remove(column_index);
        }
        true
    }
}

const fn signed(value: u64) -> i64 {
    value as i32 as i64
}

const fn unsigned(value: i64) -> u64 {
    value as i32 as i64 as u64
}

fn empty_volatile() -> Vec<u8> {
    let mut output = Vec::new();
    for number in [1, 2, 3, 4, 5, 7] {
        message_field(&mut output, number, &[]);
    }
    output
}

fn empty_spanning() -> Vec<u8> {
    let mut sentinel = Vec::new();
    varint_field(&mut sentinel, 1, WHOLE_ROW_COLUMN_SENTINEL);
    varint_field(&mut sentinel, 2, WHOLE_COLUMN_ROW_SENTINEL);
    varint_field(&mut sentinel, 3, WHOLE_ROW_COLUMN_SENTINEL);
    varint_field(&mut sentinel, 4, WHOLE_COLUMN_ROW_SENTINEL);
    let mut output = Vec::new();
    message_field(&mut output, 2, &sentinel);
    message_field(&mut output, 3, &sentinel);
    output
}

/// One family owner, encoded in field order as the host's Prost encoder
/// writes it.
fn owner_payload(
    uid: Uuid,
    internal: u32,
    kind: Kind,
    base: Option<Uuid>,
    tile: Option<u64>,
    records: &[Vec<u8>],
    volatile_time: Option<&Coordinates>,
    table_reference: Option<(Uuid, &Coordinates)>,
) -> Vec<u8> {
    let mut output = Vec::new();
    message_field(&mut output, 1, &uuid(uid));
    varint_field(&mut output, 2, u64::from(internal));
    varint_field(&mut output, 3, kind as u64);
    let mut cells = Vec::new();
    for record in records {
        message_field(&mut cells, 1, record);
    }
    message_field(&mut output, 4, &cells);
    message_field(&mut output, 5, &[]);
    let volatile = match volatile_time {
        Some(coordinates) => {
            let mut volatile = Vec::new();
            message_field(&mut volatile, 1, &coordinates.encode());
            for number in [2, 3, 4, 5, 7] {
                message_field(&mut volatile, number, &[]);
            }
            volatile
        },
        None => empty_volatile(),
    };
    message_field(&mut output, 6, &volatile);
    message_field(&mut output, 7, &empty_spanning());
    message_field(&mut output, 8, &empty_spanning());
    let mut whole = Vec::new();
    message_field(&mut whole, 1, &[]);
    message_field(&mut output, 9, &whole);
    message_field(&mut output, 10, &[]);
    if let Some(base) = base {
        message_field(&mut output, 12, &uuid(base));
    }
    let mut tiled = Vec::new();
    if let Some(tile) = tile {
        message_field(&mut tiled, 1, &reference(tile));
    }
    message_field(&mut output, 13, &tiled);
    let mut references = Vec::new();
    if let Some((owner, coordinates)) = table_reference {
        let mut table_ref = Vec::new();
        message_field(&mut table_ref, 1, &uuid(owner));
        message_field(&mut table_ref, 2, &coordinates.encode());
        message_field(&mut references, 1, &table_ref);
    }
    message_field(&mut output, 14, &references);
    message_field(&mut output, 15, &[]);
    message_field(&mut output, 16, &[]);
    output
}

fn tile_payload(internal: u32, records: &[Vec<u8>]) -> Vec<u8> {
    let mut output = Vec::new();
    varint_field(&mut output, 1, u64::from(internal));
    varint_field(&mut output, 2, 0);
    varint_field(&mut output, 3, 0);
    for record in records {
        message_field(&mut output, 4, record);
    }
    output
}

// ---------------------------------------------------------------------------
// Reading
// ---------------------------------------------------------------------------

fn decode_uuid(payload: &[u8], limits: &Limits) -> Result<Uuid> {
    let lower = single_varint(payload, 1, limits).map_err(|_| Error::InvalidSource)?;
    let upper = single_varint(payload, 2, limits).map_err(|_| Error::InvalidSource)?;
    Ok(Uuid::new(lower.unwrap_or(0), upper.unwrap_or(0)))
}

struct OwnerFacts {
    uid: Uuid,
    internal: u32,
}

fn owner_facts(owner: Object<'_>, limits: &Limits) -> Result<OwnerFacts> {
    let uid = single_bytes(owner.payload, 1, limits)
        .map_err(|_| Error::InvalidSource)?
        .ok_or(Error::InvalidSource)?;
    let internal = single_varint(owner.payload, 2, limits)
        .map_err(|_| Error::InvalidSource)?
        .and_then(|value| u32::try_from(value).ok())
        .ok_or(Error::InvalidSource)?;
    Ok(OwnerFacts {
        uid: decode_uuid(uid, limits)?,
        internal,
    })
}

fn table_owner(source: Source<'_>, limits: &Limits) -> Result<OwnerFacts> {
    for owner in source.owners {
        let kind = single_varint(owner.payload, 3, limits).map_err(|_| Error::InvalidSource)?;
        let target = single_bytes(owner.payload, 11, limits)
            .map_err(|_| Error::InvalidSource)?
            .map(|reference| single_varint(reference, 1, limits))
            .transpose()
            .map_err(|_| Error::InvalidSource)?
            .flatten();
        if kind == Some(TABLE_OWNER_KIND) && target == Some(source.table) {
            return owner_facts(*owner, limits);
        }
    }
    Err(Error::MissingTableOwner)
}

fn conditional_owner<'a>(source: Source<'a>, limits: &Limits) -> Result<Option<Object<'a>>> {
    for owner in source.owners {
        if owner_facts(*owner, limits)?.uid == source.conditional_owner {
            return Ok(Some(*owner));
        }
    }
    Ok(None)
}

fn offset(base: Uuid, offset: u64) -> Result<Uuid> {
    Ok(Uuid::new(
        base.lower()
            .checked_add(offset)
            .ok_or(Error::InvalidSource)?,
        base.upper(),
    ))
}

fn host_record(source: Source<'_>, style_internal: u32, table_internal: u32) -> Vec<u8> {
    Record {
        row: source.row,
        column: source.column,
        without_owner: Vec::new(),
        with_owner: vec![
            (STYLE_STORAGE_ROW, STYLE_STORAGE_COLUMN, style_internal),
            (source.row, source.column, table_internal),
        ],
    }
    .encode()
}

// ---------------------------------------------------------------------------
// Planning
// ---------------------------------------------------------------------------

/// Plan registering the volatile cell `source.row`/`source.column`: the
/// family (twelve owners and four tiles from `next_identifier` on, with
/// `global_name` as the global-name owner's UUID) when the table has none,
/// otherwise one record on the conditional-style owner and its tile. A cell
/// already registered needs no edit.
///
/// # Errors
///
/// [`Error::MissingTableOwner`]; [`Error::IncompleteFamily`] for a
/// conditional-style owner without the family's style-storage owner, volatile
/// set or tile; [`Error::InvalidSource`] for a malformed engine or owner;
/// [`Error::Rewrite`] for an engine the shared rewrite refuses.
pub fn ensure<'a>(
    source: Source<'a>,
    next_identifier: u64,
    global_name: Uuid,
    object: &dyn Fn(u64) -> Option<&'a [u8]>,
    limits: &Limits,
) -> Result<Edit> {
    let table = table_owner(source, limits)?;
    let style_uid = offset(table.uid, STYLE_STORAGE_OWNER_OFFSET)?;
    let graph_limits = formula_graph::Limits::DEFAULT;
    if let Some(owner) = conditional_owner(source, limits)? {
        let style = source
            .owners
            .iter()
            .map(|owner| owner_facts(*owner, limits))
            .collect::<Result<Vec<_>>>()?
            .into_iter()
            .find(|facts| facts.uid == style_uid)
            .ok_or(Error::IncompleteFamily)?;
        let record = host_record(source, style.internal, table.internal);
        return append(
            source,
            owner,
            &record,
            table.uid,
            Some(next_identifier),
            object,
            limits,
        );
    }
    let payloads: Vec<&[u8]> = source.owners.iter().map(|owner| owner.payload).collect();
    let first_internal =
        formula_graph::next_owner_id(source.engine.payload, &payloads, &graph_limits)
            .map_err(Error::Rewrite)?;
    let internal = |kind: Kind| -> Result<u32> {
        first_internal
            .checked_add(kind.internal_offset())
            .ok_or(Error::InvalidSource)
    };
    let style_internal = internal(Kind::StyleStorage)?;
    let record = host_record(source, style_internal, table.internal);
    let coordinates = Coordinates::single(source.row, source.column);
    let first_tile = next_identifier
        .checked_add(Kind::OBJECT_ORDER.len() as u64)
        .ok_or(Error::InvalidSource)?;
    let mut next_tile = first_tile;
    let mut created = Vec::new();
    let mut owner_ids = Vec::new();
    let mut map = Vec::new();
    for (position, kind) in Kind::OBJECT_ORDER.iter().copied().enumerate() {
        let identifier = next_identifier
            .checked_add(position as u64)
            .ok_or(Error::InvalidSource)?;
        let tile = if kind.has_tile() {
            let tile = next_tile;
            next_tile = next_tile.checked_add(1).ok_or(Error::InvalidSource)?;
            Some(tile)
        } else {
            None
        };
        let uid = match kind {
            Kind::GlobalName => global_name,
            Kind::Global => GLOBAL_OWNER_UID,
            Kind::ConditionalStyle => source.conditional_owner,
            _ => offset(table.uid, kind as u64)?,
        };
        let base = match kind {
            Kind::GlobalName => None,
            Kind::Global => Some(GLOBAL_OWNER_BASE_UID),
            Kind::ConditionalStyle => Some(Uuid::new(
                source
                    .conditional_owner
                    .lower()
                    .checked_sub(CONDITIONAL_BASE_OWNER_OFFSET)
                    .ok_or(Error::InvalidSource)?,
                source.conditional_owner.upper(),
            )),
            _ => Some(table.uid),
        };
        let records = records(kind, style_internal, &record);
        let (volatile, table_reference) = if kind == Kind::ConditionalStyle {
            (Some(&coordinates), Some((table.uid, &coordinates)))
        } else {
            (None, None)
        };
        let owner_internal = internal(kind)?;
        owner_ids.push(identifier);
        created.push(Created {
            identifier,
            message_type: 4_008,
            payload: owner_payload(
                uid,
                owner_internal,
                kind,
                base,
                tile,
                &records,
                volatile,
                table_reference,
            ),
            references: tile.into_iter().collect(),
        });
        if let Some(tile) = tile {
            created.push(Created {
                identifier: tile,
                message_type: 4_009,
                payload: tile_payload(owner_internal, &records),
                references: Vec::new(),
            });
        }
        map.push((kind, owner_internal, uid));
    }
    let map_entries: Vec<(u32, Uuid)> = Kind::INTERNAL_ORDER
        .iter()
        .map(|kind| {
            map.iter()
                .find(|(candidate, _, _)| candidate == kind)
                .map(|(_, internal, uid)| (*internal, *uid))
                .ok_or(Error::InvalidSource)
        })
        .collect::<Result<_>>()?;
    let existing = engine_owner_references(source.engine.payload, limits)?;
    let mut engine = formula_graph::append_owners_to_engine(
        source.engine.payload,
        &owner_ids,
        &map_entries,
        REGISTERED_FORMULA_COUNT,
        &graph_limits,
    )
    .map_err(Error::Rewrite)?;
    if existing.len() == 1 {
        // The host's native order: the new owners but the last, the source
        // owner, then the last new owner.
        let mut order = owner_ids[..owner_ids.len() - 1].to_vec();
        order.extend(existing);
        order.push(owner_ids[owner_ids.len() - 1]);
        engine = reorder_engine_owners(&engine, &order, limits)?;
    }
    Ok(Edit {
        payloads: vec![(source.engine.identifier, engine)],
        created,
        engine_references: owner_ids,
        owner_references: Vec::new(),
    })
}

fn records(kind: Kind, style_internal: u32, host: &[u8]) -> Vec<Vec<u8>> {
    match kind {
        Kind::ConditionalStyle => vec![host.to_vec()],
        Kind::Global => (0..GLOBAL_OWNER_COLUMN_COUNT)
            .map(|column| {
                Record {
                    row: GLOBAL_OWNER_ROW,
                    column,
                    without_owner: vec![(0, 0)],
                    with_owner: Vec::new(),
                }
                .encode()
            })
            .collect(),
        Kind::Auxiliary8 => {
            let mut records: Vec<Vec<u8>> = (0..=AUXILIARY_STYLE_EMPTY_END_COLUMN)
                .map(|column| {
                    Record {
                        row: AUXILIARY_STYLE_ROW,
                        column,
                        without_owner: if column == 0 {
                            vec![(AUXILIARY_STYLE_ROW, AUXILIARY_STYLE_SOURCE_COLUMN)]
                        } else {
                            Vec::new()
                        },
                        with_owner: Vec::new(),
                    }
                    .encode()
                })
                .collect();
            records.push(
                Record {
                    row: AUXILIARY_STYLE_ROW,
                    column: AUXILIARY_STYLE_LAST_COLUMN,
                    without_owner: Vec::new(),
                    with_owner: vec![(
                        STYLE_STORAGE_ROW,
                        AUXILIARY_STYLE_STORAGE_COLUMN,
                        style_internal,
                    )],
                }
                .encode(),
            );
            records
        },
        _ => Vec::new(),
    }
}

/// Plan removing the volatile cell `source.row`/`source.column` from the
/// table's conditional-style owner and its tile; no edit when the cell is not
/// registered (the host's `remove_volatile_host`).
///
/// # Errors
///
/// [`Error::IncompleteFamily`] for a registered cell without the family's
/// volatile set or tile; [`Error::InvalidSource`] for a malformed owner,
/// tile or engine.
pub fn remove<'a>(
    source: Source<'a>,
    object: &dyn Fn(u64) -> Option<&'a [u8]>,
    limits: &Limits,
) -> Result<Edit> {
    let Some(owner) = conditional_owner(source, limits)? else {
        return Ok(Edit::default());
    };
    let table = table_owner(source, limits)?;
    append(source, owner, &[], table.uid, None, object, limits)
}

/// The shared append/remove of one host record on the conditional-style
/// owner, its volatile-time set, its UUID references and its tile, with the
/// engine's formula count. `add` holds the next free identifier when adding.
///
/// Every Numbers-saved table already registers the whole owner family; only
/// its conditional-style owner has no volatile-time set and no tile until the
/// first volatile rule (compare `basic.numbers` with
/// `volatile-highlight-native.numbers`, record 0432). The first add creates
/// them as Numbers does.
fn append<'a>(
    source: Source<'a>,
    owner: Object<'_>,
    record: &[u8],
    table_uid: Uuid,
    add: Option<u64>,
    object: &dyn Fn(u64) -> Option<&'a [u8]>,
    limits: &Limits,
) -> Result<Edit> {
    let next_identifier = add;
    let add = add.is_some();
    let row = i64::from(source.row);
    let column = source.column;
    let mut cells = Vec::new();
    let mut present = false;
    if let Some(dependencies) =
        single_bytes(owner.payload, 4, limits).map_err(|_| Error::InvalidSource)?
    {
        for entry in fields(dependencies, limits).map_err(|_| Error::InvalidSource)? {
            let same = single_varint(entry.payload, 1, limits).map_err(|_| Error::InvalidSource)?
                == Some(u64::from(column))
                && single_varint(entry.payload, 2, limits).map_err(|_| Error::InvalidSource)?
                    == Some(u64::from(source.row));
            present |= same;
            if !same || add {
                cells.push(entry.raw.to_vec());
            }
        }
    }
    if add == present {
        return Ok(Edit::default());
    }
    if add {
        let mut field = Vec::new();
        message_field(&mut field, 1, record);
        cells.push(field);
    }
    let owner_internal = single_varint(owner.payload, 2, limits)
        .map_err(|_| Error::InvalidSource)?
        .and_then(|internal| u32::try_from(internal).ok())
        .ok_or(Error::InvalidSource)?;
    // Numbers' own map; Keynote and Pages hold to the host's shape.
    let listed = |engine: Vec<u8>| {
        if source.creates_first_tile {
            with_uuid_reference(engine, table_uid, owner_internal, row, column, add, limits)
        } else {
            Ok(engine)
        }
    };
    let mut rebuilt = Vec::new();
    let mut tile = None;
    let mut created_tile = None;
    for field in fields(owner.payload, limits).map_err(|_| Error::InvalidSource)? {
        match field.number {
            4 => message_field(&mut rebuilt, 4, &cells.concat()),
            6 => {
                let mut volatile = Vec::new();
                let mut seen = false;
                for entry in fields(field.payload, limits).map_err(|_| Error::InvalidSource)? {
                    if entry.number == 1 {
                        seen = true;
                        let mut set = Coordinates::decode(entry.payload, limits)?;
                        if add {
                            set.add(row, column);
                        } else if !set.remove(row, column) {
                            return Err(Error::InvalidSource);
                        }
                        message_field(&mut volatile, 1, &set.encode());
                    } else {
                        volatile.extend_from_slice(entry.raw);
                    }
                }
                if !seen {
                    if !add || !source.creates_first_tile {
                        return Err(Error::IncompleteFamily);
                    }
                    // Field 1 first, as Numbers writes the message.
                    let mut first = Vec::new();
                    message_field(
                        &mut first,
                        1,
                        &Coordinates::single(source.row, column).encode(),
                    );
                    first.extend_from_slice(&volatile);
                    volatile = first;
                }
                message_field(&mut rebuilt, 6, &volatile);
            },
            14 => {
                let mut references = Vec::new();
                let mut matched = false;
                for entry in fields(field.payload, limits).map_err(|_| Error::InvalidSource)? {
                    if entry.number != 1 {
                        references.extend_from_slice(entry.raw);
                        continue;
                    }
                    let owner_uuid = single_bytes(entry.payload, 1, limits)
                        .map_err(|_| Error::InvalidSource)?
                        .ok_or(Error::InvalidSource)?;
                    let owner_uuid = decode_uuid(owner_uuid, limits)?;
                    let mut set = match single_bytes(entry.payload, 2, limits)
                        .map_err(|_| Error::InvalidSource)?
                    {
                        Some(set) => Coordinates::decode(set, limits)?,
                        None => Coordinates::default(),
                    };
                    let target = if add {
                        owner_uuid == table_uid
                    } else {
                        !matched && set.contains(row, column)
                    };
                    if !target {
                        references.extend_from_slice(entry.raw);
                        continue;
                    }
                    matched = true;
                    if add {
                        set.add(row, column);
                    } else {
                        set.remove(row, column);
                        if set.0.is_empty() {
                            continue;
                        }
                    }
                    let mut table_ref = Vec::new();
                    message_field(&mut table_ref, 1, &uuid(owner_uuid));
                    message_field(&mut table_ref, 2, &set.encode());
                    message_field(&mut references, 1, &table_ref);
                }
                if add && !matched {
                    let mut table_ref = Vec::new();
                    message_field(&mut table_ref, 1, &uuid(table_uid));
                    message_field(
                        &mut table_ref,
                        2,
                        &Coordinates::single(source.row, column).encode(),
                    );
                    message_field(&mut references, 1, &table_ref);
                } else if !add && !matched {
                    return Err(Error::InvalidSource);
                }
                message_field(&mut rebuilt, 14, &references);
            },
            13 => {
                tile = fields(field.payload, limits)
                    .map_err(|_| Error::InvalidSource)?
                    .first()
                    .map(|entry| single_varint(entry.payload, 1, limits))
                    .transpose()
                    .map_err(|_| Error::InvalidSource)?
                    .flatten();
                match (tile, next_identifier) {
                    (None, Some(identifier))
                        if field.payload.is_empty() && source.creates_first_tile =>
                    {
                        let mut tiled = Vec::new();
                        message_field(&mut tiled, 1, &reference(identifier));
                        message_field(&mut rebuilt, 13, &tiled);
                        created_tile = Some(identifier);
                    },
                    _ => rebuilt.extend_from_slice(field.raw),
                }
            },
            _ => rebuilt.extend_from_slice(field.raw),
        }
    }
    if let Some(identifier) = created_tile {
        let engine = formula_graph::append_owners_to_engine(
            source.engine.payload,
            &[],
            &[],
            1,
            &formula_graph::Limits::DEFAULT,
        )
        .map_err(Error::Rewrite)?;
        return Ok(Edit {
            payloads: vec![
                (source.engine.identifier, listed(engine)?),
                (owner.identifier, rebuilt),
            ],
            created: vec![Created {
                identifier,
                message_type: 4_009,
                payload: tile_payload(owner_internal, &[record.to_vec()]),
                references: Vec::new(),
            }],
            engine_references: Vec::new(),
            owner_references: vec![(owner.identifier, identifier)],
        });
    }
    let tile = tile.ok_or(Error::IncompleteFamily)?;
    let tile_payload = object(tile).ok_or(Error::IncompleteFamily)?;
    let mut tile_rebuilt = Vec::new();
    let mut tile_present = false;
    for field in fields(tile_payload, limits).map_err(|_| Error::InvalidSource)? {
        if field.number == 4 {
            let same = single_varint(field.payload, 1, limits).map_err(|_| Error::InvalidSource)?
                == Some(u64::from(column))
                && single_varint(field.payload, 2, limits).map_err(|_| Error::InvalidSource)?
                    == Some(u64::from(source.row));
            if same {
                tile_present = true;
                if !add {
                    continue;
                }
            }
        }
        tile_rebuilt.extend_from_slice(field.raw);
    }
    if add {
        message_field(&mut tile_rebuilt, 4, record);
    } else if !tile_present {
        return Err(Error::InvalidSource);
    }
    let engine = if add {
        formula_graph::append_owners_to_engine(
            source.engine.payload,
            &[],
            &[],
            1,
            &formula_graph::Limits::DEFAULT,
        )
        .map_err(Error::Rewrite)?
    } else {
        let options = dependency::DecodeOptions::new(
            limits.max_input_bytes(),
            limits.max_fields(),
            limits.max_input_bytes().saturating_mul(32),
            64,
            limits.max_items(),
            limits.max_output_bytes(),
        );
        dependency::rewrite_calculation_engine_owner_removal(
            source.engine.payload,
            &[],
            &[],
            1,
            options,
        )
        .map(|(data, _)| data)
        .map_err(|_| Error::InvalidSource)?
    };
    Ok(Edit {
        payloads: vec![
            (source.engine.identifier, listed(engine)?),
            (owner.identifier, rebuilt),
            (tile, tile_rebuilt),
        ],
        created: Vec::new(),
        engine_references: Vec::new(),
        owner_references: Vec::new(),
    })
}

/// The engine with one coordinate added to or taken out of
/// `uuid_reference_map` (field 7) under the table's UUID and the
/// conditional-style owner, as Numbers lists its rule cells
/// (`volatile-highlight-native.numbers`, record 0433). An engine without the
/// field (a host-built document) stays as it is, and so does a removal of a
/// coordinate the map does not list: leaf output of earlier records holds none.
fn with_uuid_reference(
    engine: Vec<u8>,
    table_uid: Uuid,
    owner_internal: u32,
    row: i64,
    column: u32,
    add: bool,
    limits: &Limits,
) -> Result<Vec<u8>> {
    let mut rebuilt = Vec::with_capacity(engine.len() + 32);
    for field in fields(&engine, limits).map_err(|_| Error::InvalidSource)? {
        if field.number == UUID_REFERENCE_MAP_FIELD {
            let map = uuid_reference_map(
                field.payload,
                table_uid,
                owner_internal,
                row,
                column,
                add,
                limits,
            )?;
            message_field(&mut rebuilt, UUID_REFERENCE_MAP_FIELD, &map);
        } else {
            rebuilt.extend_from_slice(field.raw);
        }
    }
    Ok(rebuilt)
}

/// `TSCE.UuidReferenceMapArchive` with the coordinate added or removed.
fn uuid_reference_map(
    map: &[u8],
    table_uid: Uuid,
    owner_internal: u32,
    row: i64,
    column: u32,
    add: bool,
    limits: &Limits,
) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(map.len() + 32);
    let mut matched = false;
    for entry in fields(map, limits).map_err(|_| Error::InvalidSource)? {
        let uid = single_bytes(entry.payload, 1, limits)
            .map_err(|_| Error::InvalidSource)?
            .ok_or(Error::InvalidSource)?;
        if entry.number != 1 || decode_uuid(uid, limits)? != table_uid {
            output.extend_from_slice(entry.raw);
            continue;
        }
        matched = true;
        let mut rewritten = Vec::with_capacity(entry.payload.len() + 16);
        let mut has_refs = false;
        for inner in fields(entry.payload, limits).map_err(|_| Error::InvalidSource)? {
            if inner.number == 3 {
                has_refs = true;
                let refs = owner_cell_refs(
                    inner.payload,
                    owner_internal,
                    row,
                    column,
                    add,
                    limits,
                )?;
                if !refs.is_empty() {
                    message_field(&mut rewritten, 3, &refs);
                }
            } else {
                rewritten.extend_from_slice(inner.raw);
            }
        }
        if add && !has_refs {
            let refs = owner_cell_refs(&[], owner_internal, row, column, add, limits)?;
            message_field(&mut rewritten, 3, &refs);
        }
        // The UUID alone says nothing: the entry goes with its last cell.
        if single_bytes(&rewritten, 3, limits)
            .map_err(|_| Error::InvalidSource)?
            .is_some()
            || single_bytes(&rewritten, 2, limits)
                .map_err(|_| Error::InvalidSource)?
                .is_some()
        {
            message_field(&mut output, 1, &rewritten);
        }
    }
    if add && !matched {
        let mut entry = Vec::new();
        message_field(&mut entry, 1, &uuid(table_uid));
        let refs = owner_cell_refs(&[], owner_internal, row, column, add, limits)?;
        message_field(&mut entry, 3, &refs);
        message_field(&mut output, 1, &entry);
    }
    Ok(output)
}

/// `TSCE.InternalCellRefSetArchive` with the coordinate added or removed.
fn owner_cell_refs(
    refs: &[u8],
    owner_internal: u32,
    row: i64,
    column: u32,
    add: bool,
    limits: &Limits,
) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(refs.len() + 16);
    let mut matched = false;
    let owner_entry = |set: &Coordinates| {
        let mut entry = Vec::new();
        varint_field(&mut entry, 1, u64::from(owner_internal));
        message_field(&mut entry, 2, &set.encode());
        entry
    };
    for entry in fields(refs, limits).map_err(|_| Error::InvalidSource)? {
        let owner = single_varint(entry.payload, 1, limits).map_err(|_| Error::InvalidSource)?;
        if entry.number != 1 || owner != Some(u64::from(owner_internal)) {
            output.extend_from_slice(entry.raw);
            continue;
        }
        matched = true;
        let mut set = match single_bytes(entry.payload, 2, limits)
            .map_err(|_| Error::InvalidSource)?
        {
            Some(set) => Coordinates::decode(set, limits)?,
            None => Coordinates::default(),
        };
        if add {
            set.add(row, column);
        } else {
            set.remove(row, column);
        }
        if !set.0.is_empty() {
            message_field(&mut output, 1, &owner_entry(&set));
        }
    }
    if add && !matched {
        let mut set = Coordinates::default();
        set.add(row, column);
        message_field(&mut output, 1, &owner_entry(&set));
    }
    Ok(output)
}

/// The engine's formula-owner references (tracker field 6), in order.
fn engine_owner_references(engine: &[u8], limits: &Limits) -> Result<Vec<u64>> {
    let tracker = single_bytes(engine, 2, limits)
        .map_err(|_| Error::InvalidSource)?
        .ok_or(Error::InvalidSource)?;
    let mut references = Vec::new();
    for field in fields(tracker, limits).map_err(|_| Error::InvalidSource)? {
        if field.number == 6 {
            references.push(
                single_varint(field.payload, 1, limits)
                    .map_err(|_| Error::InvalidSource)?
                    .ok_or(Error::InvalidSource)?,
            );
        }
    }
    Ok(references)
}

/// The engine with its tracker's formula-owner references in `order`; the
/// references keep their position among the tracker's fields.
fn reorder_engine_owners(engine: &[u8], order: &[u64], limits: &Limits) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    for field in fields(engine, limits).map_err(|_| Error::InvalidSource)? {
        if field.number != 2 {
            output.extend_from_slice(field.raw);
            continue;
        }
        let mut tracker = Vec::new();
        let mut written = false;
        for entry in fields(field.payload, limits).map_err(|_| Error::InvalidSource)? {
            if entry.number == 6 {
                if !written {
                    for identifier in order {
                        message_field(&mut tracker, 6, &reference(*identifier));
                    }
                    written = true;
                }
                continue;
            }
            tracker.extend_from_slice(entry.raw);
        }
        message_field(&mut output, 2, &tracker);
    }
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;

    const TABLE: u64 = 900;
    const TABLE_UID: Uuid = Uuid::new(0x1000, 0x2000);
    const CONDITIONAL_UID: Uuid = Uuid::new(0x5000, 0x6000);

    fn owner(uid: Uuid, internal: u64, kind: u64, table: Option<u64>) -> Vec<u8> {
        let mut output = Vec::new();
        message_field(&mut output, 1, &uuid(uid));
        varint_field(&mut output, 2, internal);
        varint_field(&mut output, 3, kind);
        if let Some(table) = table {
            let mut reference = Vec::new();
            varint_field(&mut reference, 1, table);
            message_field(&mut output, 11, &reference);
        }
        output
    }

    fn source<'a>(owners: &'a [Object<'a>]) -> Source<'a> {
        Source {
            engine: Object {
                identifier: 1,
                payload: &[],
            },
            owners,
            table: TABLE,
            conditional_owner: CONDITIONAL_UID,
            row: 1,
            column: 1,
            creates_first_tile: false,
        }
    }

    #[test]
    fn coordinates_follow_the_host_merge_and_removal() {
        let mut set = Coordinates::default();
        set.add(3, 2);
        set.add(1, 2);
        set.add(5, 0);
        set.add(3, 2);
        assert_eq!(
            set,
            Coordinates(vec![(0, vec![(5, None)]), (2, vec![(1, None), (3, None)])])
        );
        assert!(set.contains(3, 2));
        assert!(!set.contains(2, 2));

        let mut range = Coordinates(vec![(4, vec![(10, Some(14))])]);
        assert!(range.remove(12, 4));
        assert_eq!(
            range,
            Coordinates(vec![(4, vec![(10, Some(11)), (13, Some(14))])])
        );
        assert!(range.remove(10, 4));
        assert!(range.remove(14, 4));
        assert_eq!(range, Coordinates(vec![(4, vec![(11, None), (13, None)])]));
        assert!(!range.remove(12, 4));
        assert!(range.remove(11, 4));
        assert!(range.remove(13, 4));
        assert_eq!(range, Coordinates::default());

        let limits = Limits::DEFAULT;
        let set = Coordinates(vec![(2, vec![(1, None), (3, Some(7))])]);
        assert_eq!(Coordinates::decode(&set.encode(), &limits), Ok(set));
    }

    #[test]
    fn a_table_without_an_owner_is_refused() {
        let other = owner(TABLE_UID, 0, TABLE_OWNER_KIND, Some(TABLE + 1));
        let owners = [Object {
            identifier: 10,
            payload: &other,
        }];
        let limits = Limits::DEFAULT;
        assert_eq!(
            ensure(source(&owners), 100, Uuid::new(7, 8), &|_| None, &limits),
            Err(Error::MissingTableOwner)
        );
    }

    #[test]
    fn a_registered_conditional_owner_without_the_family_is_refused() {
        // Numbers-saved tables register the conditional-style owner alone.
        let table = owner(TABLE_UID, 0, TABLE_OWNER_KIND, Some(TABLE));
        let conditional = owner(CONDITIONAL_UID, 1, Kind::ConditionalStyle as u64, None);
        let owners = [
            Object {
                identifier: 10,
                payload: &table,
            },
            Object {
                identifier: 11,
                payload: &conditional,
            },
        ];
        let limits = Limits::DEFAULT;
        assert_eq!(
            ensure(source(&owners), 100, Uuid::new(7, 8), &|_| None, &limits),
            Err(Error::IncompleteFamily)
        );
        // A cell that is not registered needs no removal.
        assert_eq!(
            remove(source(&owners), &|_| None, &limits),
            Ok(Edit::default())
        );
    }

    #[test]
    fn removal_without_a_conditional_owner_is_a_no_op() {
        let table = owner(TABLE_UID, 0, TABLE_OWNER_KIND, Some(TABLE));
        let owners = [Object {
            identifier: 10,
            payload: &table,
        }];
        assert_eq!(
            remove(source(&owners), &|_| None, &Limits::DEFAULT),
            Ok(Edit::default())
        );
    }
}
