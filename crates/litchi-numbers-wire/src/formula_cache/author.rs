//! Formula authoring and removal in a table of the calculation graph, for the
//! table owners that stage payloads by object identifier (Keynote slide
//! tables, Pages body tables; record 0376a).
//!
//! [`encode`](crate::formula_cache::author::encode) lowers a
//! [`FormulaExpression`](litchi_iwa_common::formula::FormulaExpression) into the native formula archive and its dependency
//! facts. [`plan`](crate::formula_cache::author::plan) then plans the formula
//! side of setting or removing one cell's formula. That covers the
//! formula-owner dependency records, the cell-record tiles and the
//! calculation engine's formula count (through
//! [`formula_dependencies`](crate::formula_dependencies)), and the table's
//! formula list (through [`formula_list`](crate::formula_list)). The plan
//! returns the changed payloads and the new formula's list key; the owner
//! writes the cell (`table_write::{formula_free, with_formula}`) and
//! publishes. Nothing is allocated: a formula whose dependency record would
//! need a new cell-record tile object is refused, and so is a reference to
//! another table, which `FormulaExpression` names by native identifier.

use litchi_iwa_common::formula::{FormulaBinaryOperator, FormulaExpression};
use litchi_iwa_protos::{
    numbers_formula_codec::{
        self as codec, BinaryOperator as WireBinary, DecodeOptions, FormulaWriteAxisReference,
        FormulaWriteCellReference, FormulaWriteContext, FormulaWriteDependencyLimits,
        FormulaWriteDependencyVisitor, FormulaWriteNode, FormulaWritePrecedent, FormulaWriteRange,
    },
    numbers_table_cell_dependency_codec as dependency,
};

use super::refresh::{TableStorage, geometry, map_write, push, reference, tiled_references};
use super::{Error, LimitKind, TableIdentity, Unsupported, decode_options, map_decode_failure};
use crate::{
    BncCellView, Limits, StoredValue, formula_dependencies, formula_list, function_map, table_write,
};

const TABLE_OWNER_KIND: u32 = 1;
const DEPENDENCY_TILE_ROWS: u32 = 128;
const DEPENDENCY_TILE_COLUMNS: u32 = 32;
const FORMULA_NESTING: u32 = 32;

/// One object's first message payload and its `ArchiveInfo` object references.
#[derive(Debug, Clone, Copy)]
pub struct Object<'a> {
    /// The object identifier.
    pub identifier: u64,
    /// The payload.
    pub payload: &'a [u8],
    /// The message's object references, in source order.
    pub references: &'a [u64],
}

/// The calculation graph of the document and the edited table.
#[derive(Debug, Clone, Copy)]
pub struct Source<'a> {
    /// The `TSCE.CalculationEngineArchive` object.
    pub engine: Object<'a>,
    /// Every `TSCE.FormulaOwnerDependenciesArchive` object of the engine.
    pub owners: &'a [Object<'a>],
    /// The edited table's drawable (`TST.TableInfoArchive`) identifier.
    pub table: u64,
}

/// A lowered formula: the native archive and its dependency facts.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Formula {
    /// The encoded `TSCE.FormulaArchive`.
    pub bytes: Vec<u8>,
    precedents: Vec<formula_dependencies::Precedent>,
    ranges: Vec<formula_dependencies::Range>,
}

/// The formula change of one cell.
#[derive(Debug, Clone, Copy)]
pub enum Change<'f> {
    /// Remove the cell's formula (the cell's value is written by the owner).
    Remove,
    /// Set the cell's formula, replacing any formula it holds.
    Author(&'f Formula),
}

/// The formula side of one cell's change.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Edit {
    /// Changed payloads: engine, formula owners, cell-record tiles and the
    /// formula list, by object identifier.
    pub payloads: Vec<(u64, Vec<u8>)>,
    /// The authored formula's formula-list key.
    pub key: Option<u32>,
    /// The edited table's formula-owner object.
    pub owner: u64,
    /// A cell-record tile the authored record needs, created with the
    /// caller's `next_identifier`; its payload is in `payloads`, and the
    /// owner object gains a reference to it.
    pub created_tile: Option<u64>,
    /// A cell-record tile the removal left empty: the owner no longer
    /// references it (in `payloads`) and the caller deletes the object, as
    /// the host does.
    pub removed_tile: Option<u64>,
}

/// The internal formula-owner identifier and geometry of the table
/// `source.table`, or `None` when the table has no formula owner.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a malformed owner or table.
pub fn table_owner<'a>(
    source: Source<'a>,
    object: &dyn Fn(u64) -> Option<Object<'a>>,
    limits: &Limits,
) -> Result<Option<(u32, u32, u32)>, Error> {
    let cache = super::refresh::cache_limits(limits);
    let mut found = None;
    for owner in source.owners {
        let (snapshot, _) = dependency::decode_formula_owner_dependencies_with_report(
            owner.payload,
            decode_options(cache),
        )
        .map_err(map_decode_failure)?;
        if snapshot.owner_kind() != Some(TABLE_OWNER_KIND)
            || snapshot
                .formula_owner()
                .map(|reference| reference.identifier())
                != Some(source.table)
        {
            continue;
        }
        let info = object(source.table).ok_or(Error::InvalidSource)?;
        let model = reference(info.payload, 2, limits)?.ok_or(Error::InvalidSource)?;
        let model = object(model).ok_or(Error::InvalidSource)?;
        let identity = TableIdentity {
            owner: snapshot.internal_formula_owner_id(),
            uuid_lower: snapshot.formula_owner_uid().lower(),
            uuid_upper: snapshot.formula_owner_uid().upper(),
        };
        let geometry = geometry(identity, model.payload, limits)?;
        if found
            .replace((identity.owner, geometry.rows, geometry.columns))
            .is_some()
        {
            return Err(Error::InvalidSource);
        }
    }
    Ok(found)
}

/// Lower `expression` for the cell `row`/`column` of a `rows` × `columns`
/// table whose internal formula owner is `owner`.
///
/// # Errors
///
/// [`Error::UnsupportedDependency`] (`Unsupported::Formula`) for a reference
/// to another table, a pivot category, an unknown function or a value the
/// native formula cannot hold; [`Error::Limit`] when `limits` is exceeded.
pub fn encode(
    expression: &FormulaExpression,
    owner: u32,
    row: u32,
    column: u32,
    rows: u32,
    columns: u32,
    limits: &Limits,
) -> Result<Formula, Error> {
    let mut nodes = Vec::new();
    lower(expression, &mut nodes, 0)?;
    let nodes_count = nodes.len();
    let text_bytes = nodes
        .iter()
        .map(|node| match node {
            FormulaWriteNode::Text(text) => text.len(),
            _ => 0,
        })
        .sum::<usize>();
    let options = DecodeOptions::new(
        limits.max_input_bytes(),
        limits.max_fields(),
        limits.max_input_bytes().saturating_mul(32),
        FORMULA_NESTING,
        limits.max_items().min(nodes_count.max(1)),
        text_bytes.max(1),
    );
    let context = FormulaWriteContext::new(owner, row, column, rows, columns);
    let plan = codec::plan_resolved_formula_archive(
        &nodes,
        context,
        &[],
        FormulaWriteDependencyLimits::new(nodes_count, nodes_count),
        options,
    )
    .map_err(map_formula_codec)?;
    let mut collector = Collector::default();
    let (bytes, _report) =
        codec::execute_formula_archive_plan_with_visitor(plan, options, &mut collector)
            .map_err(map_formula_codec)?;
    let mut precedents = Vec::new();
    for precedent in collector.precedents {
        push(
            &mut precedents,
            formula_dependencies::Precedent {
                target_owner: precedent.internal_owner(),
                row: precedent.row(),
                column: precedent.column(),
            },
        )?;
    }
    precedents.sort_unstable();
    precedents.dedup();
    let mut ranges = Vec::new();
    for range in collector.ranges {
        push(
            &mut ranges,
            formula_dependencies::Range {
                target_owner: range.internal_owner(),
                top: range.top(),
                left: range.left(),
                bottom: range.bottom(),
                right: range.right(),
            },
        )?;
    }
    ranges.sort_unstable();
    ranges.dedup();
    Ok(Formula {
        bytes,
        precedents,
        ranges,
    })
}

#[derive(Default)]
struct Collector {
    precedents: Vec<FormulaWritePrecedent>,
    ranges: Vec<FormulaWriteRange>,
}

impl FormulaWriteDependencyVisitor for Collector {
    fn visit_precedent(
        &mut self,
        precedent: FormulaWritePrecedent,
    ) -> Result<(), codec::DecodeError> {
        self.precedents
            .try_reserve(1)
            .map_err(|_| codec::DecodeError::allocation(1))?;
        self.precedents.push(precedent);
        Ok(())
    }

    fn visit_range(&mut self, range: FormulaWriteRange) -> Result<(), codec::DecodeError> {
        self.ranges
            .try_reserve(1)
            .map_err(|_| codec::DecodeError::allocation(1))?;
        self.ranges.push(range);
        Ok(())
    }
}

fn unsupported<T>() -> Result<T, Error> {
    Err(Error::UnsupportedDependency(Unsupported::Formula))
}

fn index(value: usize) -> Result<u32, Error> {
    u32::try_from(value).map_err(|_| Error::UnsupportedDependency(Unsupported::Formula))
}

fn cell(
    reference: litchi_iwa_common::formula::FormulaCellReference,
) -> Result<FormulaWriteCellReference, Error> {
    Ok(FormulaWriteCellReference::new(
        index(reference.row)?,
        index(reference.column)?,
        reference.absolute_row,
        reference.absolute_column,
    ))
}

fn axis(
    reference: litchi_iwa_common::formula::FormulaAxisReference,
) -> Result<FormulaWriteAxisReference, Error> {
    Ok(FormulaWriteAxisReference::new(
        index(reference.index)?,
        reference.absolute,
    ))
}

fn lower<'e>(
    expression: &'e FormulaExpression,
    output: &mut Vec<FormulaWriteNode<'e>>,
    depth: u32,
) -> Result<(), Error> {
    if depth > FORMULA_NESTING {
        return Err(Error::Limit {
            kind: LimitKind::Nesting,
            observed: u64::from(depth),
            maximum: u64::from(FORMULA_NESTING),
        });
    }
    let depth = depth + 1;
    match expression {
        FormulaExpression::Number(value) => {
            if !value.is_finite() {
                return unsupported();
            }
            if *value != 0.0 && value.is_sign_negative() {
                push(output, FormulaWriteNode::Number(value.abs()))?;
                push(output, FormulaWriteNode::Negation)?;
            } else {
                push(output, FormulaWriteNode::Number(*value))?;
            }
        },
        FormulaExpression::Text(text) => push(output, FormulaWriteNode::Text(text))?,
        FormulaExpression::Boolean(value) => push(output, FormulaWriteNode::Boolean(*value))?,
        FormulaExpression::Cell(reference) => push(
            output,
            FormulaWriteNode::ResolvedCellReference {
                owner: None,
                reference: cell(*reference)?,
            },
        )?,
        FormulaExpression::Range { start, end } => push(
            output,
            FormulaWriteNode::ResolvedRange {
                owner: None,
                start: cell(*start)?,
                end: cell(*end)?,
            },
        )?,
        FormulaExpression::Rows { start, end } => push(
            output,
            FormulaWriteNode::WholeRows {
                owner: None,
                start: axis(*start)?,
                end: axis(*end)?,
            },
        )?,
        FormulaExpression::Columns { start, end } => push(
            output,
            FormulaWriteNode::WholeColumns {
                owner: None,
                start: axis(*start)?,
                end: axis(*end)?,
            },
        )?,
        FormulaExpression::Function { name, arguments } => {
            for argument in arguments {
                lower(argument, output, depth)?;
            }
            let identifier = function_map::function_identifier(name)
                .ok_or(Error::UnsupportedDependency(Unsupported::Formula))?;
            push(
                output,
                FormulaWriteNode::Function {
                    identifier,
                    argument_count: index(arguments.len())?,
                },
            )?;
        },
        FormulaExpression::Binary {
            operator,
            left,
            right,
        } => {
            lower(left, output, depth)?;
            lower(right, output, depth)?;
            push(output, FormulaWriteNode::Binary(binary(*operator)))?;
        },
        FormulaExpression::Negate(value) => {
            lower(value, output, depth)?;
            push(output, FormulaWriteNode::Negation)?;
        },
        FormulaExpression::Percent(value) => {
            lower(value, output, depth)?;
            push(output, FormulaWriteNode::Percent)?;
        },
        // Other tables are named by native identifier, and pivot
        // categories by native UUIDs: not admitted (record 0376a).
        FormulaExpression::PivotCategory(_)
        | FormulaExpression::TableCell { .. }
        | FormulaExpression::TableRange { .. }
        | FormulaExpression::TableRows { .. }
        | FormulaExpression::TableColumns { .. } => return unsupported(),
    }
    Ok(())
}

const fn binary(operator: FormulaBinaryOperator) -> WireBinary {
    match operator {
        FormulaBinaryOperator::Add => WireBinary::Add,
        FormulaBinaryOperator::Subtract => WireBinary::Subtract,
        FormulaBinaryOperator::Multiply => WireBinary::Multiply,
        FormulaBinaryOperator::Divide => WireBinary::Divide,
        FormulaBinaryOperator::Power => WireBinary::Power,
        FormulaBinaryOperator::Concatenate => WireBinary::Concatenate,
        FormulaBinaryOperator::GreaterThan => WireBinary::GreaterThan,
        FormulaBinaryOperator::GreaterThanOrEqual => WireBinary::GreaterThanOrEqual,
        FormulaBinaryOperator::LessThan => WireBinary::LessThan,
        FormulaBinaryOperator::LessThanOrEqual => WireBinary::LessThanOrEqual,
        FormulaBinaryOperator::Equal => WireBinary::Equal,
        FormulaBinaryOperator::NotEqual => WireBinary::NotEqual,
    }
}

fn map_formula_codec(error: codec::DecodeError) -> Error {
    match error.resource_limit() {
        Some(codec::DecodeLimit::Allocation { requested }) => {
            Error::Allocation { amount: requested }
        },
        Some(_) => Error::Limit {
            kind: LimitKind::FormulaWork,
            observed: u64::MAX,
            maximum: 0,
        },
        None => Error::UnsupportedDependency(Unsupported::Formula),
    }
}

/// Plan setting or removing the formula of the cell `row`/`column` of
/// `source.table`. `object` resolves an object of the (staged) source.
///
/// # Errors
///
/// [`Error::UnsupportedDependency`] when the table has no formula owner,
/// `Change::Remove` names a cell without a formula, or the record would need
/// a new cell-record tile; [`Error::InvalidSource`] for a malformed graph,
/// list or table; [`Error::Limit`] or [`Error::Allocation`] when `limits` is
/// exceeded.
pub fn plan<'a>(
    source: Source<'a>,
    row: u32,
    column: u32,
    change: Change<'_>,
    next_identifier: u64,
    object: &dyn Fn(u64) -> Option<Object<'a>>,
    limits: &Limits,
) -> Result<Edit, Error> {
    let cache = super::refresh::cache_limits(limits);
    let resolve = |identifier: u64| object(identifier).map(|object| object.payload);
    // The owners and the selected table.
    let mut selected = None;
    let mut owners = Vec::new();
    for owner in source.owners {
        let (snapshot, _) = dependency::decode_formula_owner_dependencies_with_report(
            owner.payload,
            decode_options(cache),
        )
        .map_err(map_decode_failure)?;
        let internal = snapshot.internal_formula_owner_id();
        let is_selected = snapshot.owner_kind() == Some(TABLE_OWNER_KIND)
            && snapshot
                .formula_owner()
                .map(|reference| reference.identifier())
                == Some(source.table);
        let mut cells = Vec::new();
        for tile in tiled_references(owner.payload, 13, limits)? {
            let tile = object(tile).ok_or(Error::InvalidSource)?;
            let begin = |number| -> Result<u32, Error> {
                let value = table_write::single_varint(tile.payload, number, limits)
                    .map_err(map_write)?
                    .unwrap_or(0);
                u32::try_from(value).map_err(|_| Error::InvalidSource)
            };
            push(
                &mut cells,
                formula_dependencies::CellTileSource {
                    message: message(tile),
                    source_present: true,
                    tile_column_begin: begin(2)?,
                    tile_row_begin: begin(3)?,
                },
            )?;
        }
        let mut ranges = Vec::new();
        for tile in tiled_references(owner.payload, 15, limits)? {
            let tile = object(tile).ok_or(Error::InvalidSource)?;
            let target = table_write::single_varint(tile.payload, 1, limits)
                .map_err(map_write)?
                .ok_or(Error::InvalidSource)?;
            push(
                &mut ranges,
                formula_dependencies::RangeTileSource {
                    message: message(tile),
                    target_owner: u32::try_from(target).map_err(|_| Error::InvalidSource)?,
                },
            )?;
        }
        if is_selected {
            let inline = table_write::single_bytes(owner.payload, 4, limits)
                .map_err(map_write)?
                .is_some();
            let tiled = table_write::single_bytes(owner.payload, 13, limits)
                .map_err(map_write)?
                .is_some();
            if selected
                .replace((internal, inline, tiled, owner.identifier))
                .is_some()
            {
                return Err(Error::InvalidSource);
            }
        }
        push(
            &mut owners,
            (
                *owner,
                internal,
                snapshot.formula_owner_uid().lower(),
                snapshot.formula_owner_uid().upper(),
                cells,
                ranges,
            ),
        )?;
    }
    let Some((table_owner, inline, tiled, owner_object)) = selected else {
        return Err(Error::UnsupportedDependency(Unsupported::MissingOwner));
    };

    // The table's formula cells, its formula list and the changed host.
    let info = resolve(source.table).ok_or(Error::InvalidSource)?;
    let model = reference(info, 2, limits)?.ok_or(Error::InvalidSource)?;
    let model = resolve(model).ok_or(Error::InvalidSource)?;
    let storage = TableStorage::read(model, &resolve, limits)?;
    let mut hosts = Vec::new();
    for (_, base, payload) in &storage.tiles {
        for (local, cell_column, bytes) in
            table_write::stored_cells(payload, limits).map_err(map_write)?
        {
            let view = BncCellView::parse(bytes, limits).map_err(|_| Error::InvalidSource)?;
            if let StoredValue::Formula(key) = view.stored_value() {
                push(
                    &mut hosts,
                    formula_list::SourceHost {
                        row: base.checked_add(local).ok_or(Error::InvalidSource)?,
                        column: cell_column,
                        key,
                    },
                )?;
            }
        }
    }
    hosts.sort_unstable();
    let old_key = hosts
        .iter()
        .find(|host| (host.row, host.column) == (row, column))
        .map(|host| host.key);
    if old_key.is_none() && matches!(change, Change::Remove) {
        return Err(Error::UnsupportedDependency(Unsupported::Formula));
    }
    let list_id = storage.formula_list.ok_or(Error::InvalidSource)?;
    let root = object(list_id).ok_or(Error::InvalidSource)?;
    let mut segments = Vec::new();
    let mut expected = 0usize;
    for field in table_write::fields(root.payload, limits).map_err(map_write)? {
        match field.number {
            3 => expected = expected.checked_add(1).ok_or(Error::InvalidSource)?,
            4 => {
                let identifier = table_write::single_varint(field.payload, 1, limits)
                    .map_err(map_write)?
                    .ok_or(Error::InvalidSource)?;
                let segment = object(identifier).ok_or(Error::InvalidSource)?;
                expected = expected
                    .checked_add(
                        table_write::fields(segment.payload, limits)
                            .map_err(map_write)?
                            .iter()
                            .filter(|field| field.number == 3)
                            .count(),
                    )
                    .ok_or(Error::InvalidSource)?;
                push(
                    &mut segments,
                    formula_list::SourceMessage {
                        object_id: segment.identifier,
                        payload: segment.payload,
                        object_references: segment.references,
                    },
                )?;
            },
            _ => {},
        }
    }

    // The formula list.
    let deltas = [formula_list::HostDelta {
        row,
        column,
        old_formula_key: old_key,
        new_formula: match change {
            Change::Remove => None,
            Change::Author(formula) => Some(&formula.bytes),
        },
    }];
    let list_source = formula_list::SourceList {
        root: formula_list::SourceMessage {
            object_id: root.identifier,
            payload: root.payload,
            object_references: root.references,
        },
        segments: &segments,
        expected_entries: expected,
        source_hosts: &hosts,
    };
    let list_limits = list_limits(limits, expected, hosts.len());
    let list = formula_list::prepare(list_source, &deltas, list_limits)
        .and_then(formula_list::PreparedPlan::logical)
        .and_then(|plan| plan.execute(list_limits))
        .map_err(map_list)?;
    let key = list
        .assignments
        .first()
        .and_then(|assignment| assignment.key);
    if matches!(change, Change::Author(_)) && key.is_none() {
        return Err(Error::InvalidSource);
    }

    // The host keeps a record inline when the owner has inline records and
    // in the cell-record tile of its origin when the owner has a tiled
    // envelope (or no inline records), creating that tile if needed.
    let mut created_tile = None;
    let cell_tile_object_id = if matches!(change, Change::Author(_)) && (tiled || !inline) {
        let origin = (
            row / DEPENDENCY_TILE_ROWS * DEPENDENCY_TILE_ROWS,
            column / DEPENDENCY_TILE_COLUMNS * DEPENDENCY_TILE_COLUMNS,
        );
        let selected = owners
            .iter_mut()
            .find(|owner| owner.1 == table_owner)
            .ok_or(Error::InvalidSource)?;
        match selected
            .4
            .iter()
            .find(|tile| (tile.tile_row_begin, tile.tile_column_begin) == origin)
        {
            Some(tile) => Some(tile.message.object_id),
            None => {
                push(
                    &mut selected.4,
                    formula_dependencies::CellTileSource {
                        message: formula_dependencies::SourceMessage {
                            object_id: next_identifier,
                            payload: &[],
                            object_references: &[],
                        },
                        source_present: false,
                        tile_column_begin: origin.1,
                        tile_row_begin: origin.0,
                    },
                )?;
                created_tile = Some(next_identifier);
                Some(next_identifier)
            },
        }
    } else {
        None
    };
    // The dependency graph.
    let source_owners: Vec<formula_dependencies::SourceOwner<'_, '_>> = owners
        .iter()
        .map(
            |(owner, internal, lower, upper, cells, ranges)| formula_dependencies::SourceOwner {
                message: message(*owner),
                internal_owner: *internal,
                uid_lower: *lower,
                uid_upper: *upper,
                cell_tiles: cells,
                range_tiles: ranges,
            },
        )
        .collect();
    let existing: Vec<formula_dependencies::HostKey> = hosts
        .iter()
        .map(|host| formula_dependencies::HostKey {
            owner: table_owner,
            row: host.row,
            column: host.column,
        })
        .collect();
    let new = match change {
        Change::Remove => None,
        Change::Author(formula) => Some(formula_dependencies::FormulaHost::authored(
            table_owner,
            row,
            column,
            &formula.precedents,
            &formula.ranges,
        )),
    };
    let changes = [formula_dependencies::HostChange {
        old: old_key.map(|_| formula_dependencies::HostKey {
            owner: table_owner,
            row,
            column,
        }),
        new,
        cell_tile_object_id,
    }];
    let graph = formula_dependencies::SourceGraph {
        engine: message(source.engine),
        owners: &source_owners,
        table_owner,
        existing_formula_hosts: &existing,
        changes: &changes,
    };
    // Every dependency record, precedent and range costs at least one source
    // byte: the source sizes bound the planner's host tables.
    let mut bytes = source.engine.payload.len();
    let mut messages = 1usize;
    for (owner, _, _, _, cells, ranges) in &owners {
        bytes = bytes.saturating_add(owner.payload.len());
        messages = messages.saturating_add(1 + cells.len() + ranges.len());
        for tile in cells {
            bytes = bytes.saturating_add(tile.message.payload.len());
        }
        for tile in ranges {
            bytes = bytes.saturating_add(tile.message.payload.len());
        }
    }
    let graph_limits = graph_limits(limits, bytes.saturating_add(hosts.len()) + 2, messages);
    let artifact = formula_dependencies::prepare_graph(graph, graph_limits)
        .and_then(|prepared| prepared.execute(graph_limits))
        .map_err(map_graph)?;

    // A removal that empties a cell-record tile: the host drops the tile
    // and the owner's reference to it.
    let mut removed_tile = None;
    if matches!(change, Change::Remove) {
        for edit in &artifact.cell_tiles {
            let Some(payload) = &edit.payload else {
                continue;
            };
            let records = table_write::fields(payload, limits)
                .map_err(map_write)?
                .iter()
                .filter(|field| field.number == 4)
                .count();
            if records == 0 && removed_tile.replace(edit.object_id).is_some() {
                return Err(Error::InvalidSource);
            }
        }
    }
    let mut payloads = Vec::new();
    for edit in core::iter::once(artifact.engine)
        .chain(artifact.owners)
        .chain(artifact.cell_tiles)
        .chain(artifact.range_tiles)
    {
        if Some(edit.object_id) == removed_tile {
            continue;
        }
        let source_owner = owners
            .iter()
            .find(|owner| owner.0.identifier == edit.object_id)
            .map(|owner| owner.0.payload);
        let payload = match (edit.payload, source_owner) {
            (Some(payload), Some(source)) => Some(keep_empty_envelopes(source, &payload, limits)?),
            (payload, _) => payload,
        };
        let payload = match (payload, removed_tile) {
            (Some(payload), Some(tile)) if edit.object_id == owner_object => {
                Some(without_tile_reference(&payload, tile, limits)?)
            },
            (None, Some(tile)) if edit.object_id == owner_object => {
                let source = owners
                    .iter()
                    .find(|owner| owner.0.identifier == owner_object)
                    .ok_or(Error::InvalidSource)?;
                Some(without_tile_reference(source.0.payload, tile, limits)?)
            },
            (payload, _) => payload,
        };
        if let Some(payload) = payload {
            push(&mut payloads, (edit.object_id, payload))?;
        }
    }
    for edit in core::iter::once(list.root).chain(list.segments) {
        if let Some(mut payload) = edit.payload {
            if edit.object_id == root.identifier
                && let Change::Remove = change
                && let Some(released) = old_key
            {
                payload = lower_next_key(&payload, released, limits)?;
            }
            push(&mut payloads, (edit.object_id, payload))?;
        }
    }
    Ok(Edit {
        payloads,
        key,
        owner: owner_object,
        created_tile,
        removed_tile,
    })
}

/// The rewritten owner `output` with every dependency envelope `source`
/// held kept, empty, in source field order: the shared rewrite drops an
/// envelope it empties, while the host and the native files keep it
/// (record 0376a). Fields are singular, so they are matched by number.
fn keep_empty_envelopes(source: &[u8], output: &[u8], limits: &Limits) -> Result<Vec<u8>, Error> {
    const ENVELOPES: [u32; 4] = [4, 5, 13, 15];
    let produced = table_write::fields(output, limits).map_err(map_write)?;
    let mut rebuilt = Vec::new();
    let mut seen = Vec::new();
    for field in table_write::fields(source, limits).map_err(map_write)? {
        if let Some(present) = produced
            .iter()
            .find(|candidate| candidate.number == field.number)
        {
            extend_bytes(&mut rebuilt, present.raw)?;
            push(&mut seen, field.number)?;
        } else if ENVELOPES.contains(&field.number) {
            length_delimited(&mut rebuilt, field.number, &[])?;
        }
    }
    for field in &produced {
        if !seen.contains(&field.number) {
            extend_bytes(&mut rebuilt, field.raw)?;
        }
    }
    Ok(rebuilt)
}

/// `owner` without the `TSP.Reference` to `tile` in its tiled cell
/// dependencies (field 13); every other byte stays.
fn without_tile_reference(owner: &[u8], tile: u64, limits: &Limits) -> Result<Vec<u8>, Error> {
    let mut output = Vec::new();
    for field in table_write::fields(owner, limits).map_err(map_write)? {
        if field.number != 13 {
            extend_bytes(&mut output, field.raw)?;
            continue;
        }
        let mut envelope = Vec::new();
        for entry in table_write::fields(field.payload, limits).map_err(map_write)? {
            let names_tile = entry.number == 1
                && table_write::single_varint(entry.payload, 1, limits).map_err(map_write)?
                    == Some(tile);
            if !names_tile {
                extend_bytes(&mut envelope, entry.raw)?;
            }
        }
        length_delimited(&mut output, 13, &envelope)?;
    }
    Ok(output)
}

/// The host's formula-list release: removing the entry just below
/// `nextListID` lowers `nextListID` to it.
fn lower_next_key(list: &[u8], released: u32, limits: &Limits) -> Result<Vec<u8>, Error> {
    let fields = table_write::fields(list, limits).map_err(map_write)?;
    let still_present = fields.iter().any(|field| {
        field.number == 3
            && table_write::single_varint(field.payload, 1, limits)
                .ok()
                .flatten()
                == Some(u64::from(released))
    });
    let next = table_write::single_varint(list, 2, limits).map_err(map_write)?;
    if still_present || next != Some(u64::from(released) + 1) {
        return super::refresh::copy(list);
    }
    let mut output = Vec::new();
    for field in fields {
        if field.number == 2 {
            varint_field(&mut output, 2, u64::from(released))?;
        } else {
            extend_bytes(&mut output, field.raw)?;
        }
    }
    Ok(output)
}

fn extend_bytes(output: &mut Vec<u8>, bytes: &[u8]) -> Result<(), Error> {
    output
        .try_reserve(bytes.len())
        .map_err(|_| Error::Allocation {
            amount: bytes.len(),
        })?;
    output.extend_from_slice(bytes);
    Ok(())
}

fn varint(output: &mut Vec<u8>, mut value: u64) -> Result<(), Error> {
    loop {
        let byte = (value & 0x7f) as u8;
        value >>= 7;
        if value == 0 {
            return extend_bytes(output, &[byte]);
        }
        extend_bytes(output, &[byte | 0x80])?;
    }
}

fn varint_field(output: &mut Vec<u8>, number: u32, value: u64) -> Result<(), Error> {
    varint(output, u64::from(number) << 3)?;
    varint(output, value)
}

fn length_delimited(output: &mut Vec<u8>, number: u32, payload: &[u8]) -> Result<(), Error> {
    varint(output, (u64::from(number) << 3) | 2)?;
    varint(output, payload.len() as u64)?;
    extend_bytes(output, payload)
}

fn message(object: Object<'_>) -> formula_dependencies::SourceMessage<'_> {
    formula_dependencies::SourceMessage {
        object_id: object.identifier,
        payload: object.payload,
        object_references: object.references,
    }
}

fn graph_limits(limits: &Limits, facts: usize, messages: usize) -> formula_dependencies::Limits {
    let items = limits.max_items().min(facts);
    formula_dependencies::Limits {
        max_source_bytes: limits.max_input_bytes(),
        max_output_bytes: limits.max_output_bytes(),
        max_fields: limits.max_fields(),
        max_work_bytes: limits.max_input_bytes().saturating_mul(4096),
        max_references: limits.max_items(),
        max_messages: limits.max_items().min(messages),
        max_hosts: items,
        max_precedents: items,
        max_ranges: items,
        max_retained_bytes: limits.max_output_bytes().saturating_mul(4),
        max_scratch_bytes: limits.max_output_bytes().saturating_mul(4),
        max_allocations: limits.max_items(),
        recursion_limit: 64,
    }
}

fn list_limits(limits: &Limits, entries: usize, hosts: usize) -> formula_list::Limits {
    let items = limits.max_items().min(entries.max(hosts).saturating_add(2));
    formula_list::Limits {
        max_input_bytes: limits.max_input_bytes(),
        max_output_bytes: limits.max_output_bytes(),
        max_fields: limits.max_fields(),
        max_work: limits.max_input_bytes().saturating_mul(4096),
        max_entries: items,
        max_hosts: items,
        max_references: limits.max_items(),
        max_retained_elements: limits.max_items(),
        max_retained_bytes: limits.max_output_bytes().saturating_mul(4),
        max_scratch_bytes: limits.max_output_bytes().saturating_mul(4),
        max_allocations: limits.max_items(),
    }
}

fn map_graph(error: formula_dependencies::Error) -> Error {
    match error {
        formula_dependencies::Error::Limit {
            observed, maximum, ..
        } => Error::Limit {
            kind: LimitKind::GraphWork,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        },
        formula_dependencies::Error::Allocation { requested } => {
            Error::Allocation { amount: requested }
        },
        formula_dependencies::Error::InvalidGraph => {
            Error::UnsupportedDependency(Unsupported::Graph)
        },
        _ => Error::InvalidSource,
    }
}

fn map_list(error: formula_list::Error) -> Error {
    match error {
        formula_list::Error::Limit {
            observed, maximum, ..
        } => Error::Limit {
            kind: LimitKind::GraphWork,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        },
        formula_list::Error::Allocation { requested } => Error::Allocation { amount: requested },
        _ => Error::InvalidSource,
    }
}
