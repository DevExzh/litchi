//! Dependency-driven refresh of cached native formula results after cell writes.

use std::collections::{BTreeSet, HashMap, HashSet, VecDeque};

use super::*;
use crate::numbers::bnc::CachedScalar;
use crate::numbers::function_map::function_name;

const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
const CELL_DEPENDENCY_TILE_MESSAGE_TYPE: u32 = 4_009;
const RANGE_DEPENDENCY_TILE_MESSAGE_TYPE: u32 = 4_010;
const DEFAULT_TILE_SIZE_ROWS: u32 = 256;
const AVERAGE_FUNCTION_ID: u32 = 15;
const COUNT_FUNCTION_ID: u32 = 30;
const MAX_FUNCTION_ID: u32 = 84;
const MIN_FUNCTION_ID: u32 = 88;
const SUM_FUNCTION_ID: u32 = 168;
const MAX_CACHE_AGGREGATE_CELLS: u64 = 1_100_000;
const PERCENT_SCALE: f64 = 100.0;
const WHOLE_ROW_COLUMN_SENTINEL: u32 = i16::MAX as u32;
const WHOLE_COLUMN_ROW_SENTINEL: u32 = i32::MAX as u32;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
struct CellKey {
    owner_id: u32,
    row: u32,
    column: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct CellRange {
    owner_id: u32,
    top: u32,
    left: u32,
    bottom: u32,
    right: u32,
}

impl CellRange {
    fn contains(self, cell: CellKey) -> bool {
        self.owner_id == cell.owner_id
            && (self.top..=self.bottom).contains(&cell.row)
            && (self.left..=self.right).contains(&cell.column)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum FormulaPrecedent {
    Cell(CellKey),
    Range(CellRange),
}

#[derive(Debug, Clone)]
struct RuntimeTable {
    descriptor: TableDescriptor,
    formulas: HashMap<u32, tsce::FormulaArchive>,
}

struct DependencyGraph {
    locations: HashMap<u64, String>,
    tables: HashMap<u64, RuntimeTable>,
    owner_to_table: HashMap<u32, u64>,
    table_to_owner: HashMap<u64, u32>,
    uuid_to_owner: HashMap<(u64, u64), u32>,
    direct_dependents: HashMap<CellKey, HashSet<CellKey>>,
    range_dependents: Vec<(CellRange, CellKey)>,
    precedents: HashMap<CellKey, HashSet<FormulaPrecedent>>,
}

impl DependencyGraph {
    fn from_package(package: &IWorkPackage) -> Result<Option<Self>> {
        let Some(component) = package.calculation_engine_entry_name()? else {
            return Ok(None);
        };
        let descriptors = attached_table_descriptors(package)?;
        let descriptor_by_info = descriptors
            .iter()
            .map(|descriptor| (descriptor.table_info_id, descriptor))
            .collect::<HashMap<_, _>>();
        let locations = object_locations(package)?;
        let archive = package.archive(component)?;
        let owners = archive
            .objects
            .iter()
            .flat_map(|object| &object.messages)
            .filter(|message| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
            .map(|message| {
                tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())
                    .map_err(Error::from)
            })
            .collect::<Result<Vec<_>>>()?;

        let mut owner_to_table = HashMap::new();
        let mut table_to_owner = HashMap::new();
        let mut uuid_to_owner = HashMap::new();
        for owner in &owners {
            let Some(table_info_id) = owner
                .formula_owner
                .as_ref()
                .map(|reference| reference.identifier)
            else {
                continue;
            };
            let Some(descriptor) = descriptor_by_info.get(&table_info_id) else {
                continue;
            };
            if owner_to_table
                .insert(owner.internal_formula_owner_id, descriptor.object_id)
                .is_some()
            {
                return Err(Error::InvalidFormat(format!(
                    "Numbers formula owner ID {} is duplicated",
                    owner.internal_formula_owner_id
                )));
            }
            if table_to_owner
                .insert(descriptor.object_id, owner.internal_formula_owner_id)
                .is_some()
            {
                return Err(Error::InvalidFormat(format!(
                    "Numbers table {} has more than one formula owner",
                    descriptor.object_id
                )));
            }
            let uuid = uuid_key(&owner.formula_owner_uid);
            if uuid_to_owner
                .insert(uuid, owner.internal_formula_owner_id)
                .is_some()
            {
                return Err(Error::InvalidFormat(format!(
                    "Numbers formula owner UUID {:016x}-{:016x} is duplicated",
                    uuid.0, uuid.1
                )));
            }
        }

        let attached_owner_tables = owner_to_table.values().copied().collect::<HashSet<_>>();
        let mut tables = HashMap::with_capacity(attached_owner_tables.len());
        for descriptor in descriptors {
            if !attached_owner_tables.contains(&descriptor.object_id) {
                continue;
            }
            let formula_table_id = descriptor.model.base_data_store.formula_table.identifier;
            let resolved = resolve_table_data_list(
                package,
                &locations,
                formula_table_id,
                tst::table_data_list::ListType::Formula,
            )?;
            let formulas = resolved
                .entries
                .into_iter()
                .filter_map(|located| {
                    located
                        .entry
                        .formula
                        .map(|formula| (located.entry.key, formula))
                })
                .collect();
            tables.insert(
                descriptor.object_id,
                RuntimeTable {
                    descriptor,
                    formulas,
                },
            );
        }

        let mut graph = Self {
            locations,
            tables,
            owner_to_table,
            table_to_owner,
            uuid_to_owner,
            direct_dependents: HashMap::new(),
            range_dependents: Vec::new(),
            precedents: HashMap::new(),
        };
        for owner in &owners {
            if let Some(dependencies) = &owner.cell_dependencies {
                for record in &dependencies.cell_record {
                    graph.add_cell_record(owner.internal_formula_owner_id, record)?;
                }
            }
            for reference in owner
                .tiled_cell_dependencies
                .as_ref()
                .into_iter()
                .flat_map(|dependencies| &dependencies.cell_record_tiles)
            {
                let object = archive.object(reference.identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers dependency tile {} is missing",
                        reference.identifier
                    ))
                })?;
                let message = object
                    .messages
                    .iter()
                    .find(|message| message.type_ == CELL_DEPENDENCY_TILE_MESSAGE_TYPE)
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers dependency tile {} has no payload",
                            reference.identifier
                        ))
                    })?;
                let tile = tsce::CellRecordTileArchive::decode(message.data.as_slice())?;
                if tile.internal_owner_id != owner.internal_formula_owner_id {
                    return Err(Error::InvalidFormat(format!(
                        "Numbers dependency tile {} belongs to owner {}, expected {}",
                        reference.identifier,
                        tile.internal_owner_id,
                        owner.internal_formula_owner_id
                    )));
                }
                for record in &tile.cell_records {
                    graph.add_cell_record(owner.internal_formula_owner_id, record)?;
                }
            }
            if let Some(dependencies) = &owner.range_dependencies {
                for dependency in &dependencies.back_dependency {
                    let host = CellKey {
                        owner_id: owner.internal_formula_owner_id,
                        row: dependency.cell_coord_row,
                        column: dependency.cell_coord_column,
                    };
                    if dependency.range_reference.is_some()
                        && dependency.internal_range_reference.is_some()
                    {
                        return Err(Error::InvalidFormat(
                            "Numbers range dependency has both external and internal targets"
                                .to_owned(),
                        ));
                    }
                    if let Some(reference) = &dependency.internal_range_reference {
                        graph.insert_range(
                            host,
                            CellRange {
                                owner_id: reference.owner_id,
                                top: reference.range.top_left_row,
                                left: reference.range.top_left_column,
                                bottom: reference.range.bottom_right_row,
                                right: reference.range.bottom_right_column,
                            },
                        )?;
                    } else if let Some(reference) = &dependency.range_reference
                        && let Some(owner_id) = graph.owner_for_cfuuid(&reference.table_id)
                    {
                        graph.insert_range(
                            host,
                            CellRange {
                                owner_id,
                                top: reference.top_left_row,
                                left: reference.top_left_column,
                                bottom: reference.bottom_right_row,
                                right: reference.bottom_right_column,
                            },
                        )?;
                    }
                }
            }
            for reference in owner
                .tiled_range_dependencies
                .as_ref()
                .into_iter()
                .flat_map(|dependencies| &dependencies.range_precedents_tile)
            {
                let object = archive.object(reference.identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers range dependency tile {} is missing",
                        reference.identifier
                    ))
                })?;
                let message = object
                    .messages
                    .iter()
                    .find(|message| message.type_ == RANGE_DEPENDENCY_TILE_MESSAGE_TYPE)
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers range dependency tile {} has no payload",
                            reference.identifier
                        ))
                    })?;
                let tile = tsce::RangePrecedentsTileArchive::decode(message.data.as_slice())?;
                for dependency in &tile.from_to_range {
                    let host = CellKey {
                        owner_id: owner.internal_formula_owner_id,
                        row: explicit_coordinate(dependency.from_coord.row, "range host row")?,
                        column: explicit_coordinate(
                            dependency.from_coord.column,
                            "range host column",
                        )?,
                    };
                    let top = explicit_coordinate(
                        dependency.refers_to_rect.origin.row,
                        "range origin row",
                    )?;
                    let left = explicit_coordinate(
                        dependency.refers_to_rect.origin.column,
                        "range origin column",
                    )?;
                    let rows = dependency.refers_to_rect.size.num_rows.unwrap_or(1);
                    let columns = dependency.refers_to_rect.size.num_columns.unwrap_or(1);
                    if rows == 0 || columns == 0 {
                        return Err(Error::InvalidFormat(
                            "Numbers range dependency tile declares an empty rectangle".to_owned(),
                        ));
                    }
                    let bottom = top.checked_add(rows.saturating_sub(1)).ok_or_else(|| {
                        Error::ParseError("Numbers range row overflow".to_owned())
                    })?;
                    let right = left.checked_add(columns.saturating_sub(1)).ok_or_else(|| {
                        Error::ParseError("Numbers range column overflow".to_owned())
                    })?;
                    graph.insert_range(
                        host,
                        CellRange {
                            owner_id: tile.to_owner_id,
                            top,
                            left,
                            bottom,
                            right,
                        },
                    )?;
                }
            }
        }
        Ok(Some(graph))
    }

    fn add_cell_record(
        &mut self,
        host_owner_id: u32,
        record: &tsce::CellRecordExpandedArchive,
    ) -> Result<()> {
        let host = CellKey {
            owner_id: host_owner_id,
            row: record.row,
            column: record.column,
        };
        let Some(edges) = &record.expanded_edges else {
            return Ok(());
        };
        if edges.edge_without_owner_rows.len() != edges.edge_without_owner_columns.len() {
            return Err(Error::InvalidFormat(
                "Numbers local dependency rows and columns have different lengths".to_owned(),
            ));
        }
        for (&row, &column) in edges
            .edge_without_owner_rows
            .iter()
            .zip(&edges.edge_without_owner_columns)
        {
            self.insert_cell(
                host,
                CellKey {
                    owner_id: host_owner_id,
                    row,
                    column,
                },
            );
        }
        if edges.edge_with_owner_rows.len() != edges.edge_with_owner_columns.len()
            || edges.edge_with_owner_rows.len() != edges.internal_owner_id_for_edge.len()
        {
            return Err(Error::InvalidFormat(
                "Numbers external dependency rows, columns, and owners have different lengths"
                    .to_owned(),
            ));
        }
        for ((&row, &column), &owner_id) in edges
            .edge_with_owner_rows
            .iter()
            .zip(&edges.edge_with_owner_columns)
            .zip(&edges.internal_owner_id_for_edge)
        {
            self.insert_cell(
                host,
                CellKey {
                    owner_id,
                    row,
                    column,
                },
            );
        }
        Ok(())
    }

    fn insert_cell(&mut self, host: CellKey, precedent: CellKey) {
        self.direct_dependents
            .entry(precedent)
            .or_default()
            .insert(host);
        self.precedents
            .entry(host)
            .or_default()
            .insert(FormulaPrecedent::Cell(precedent));
    }

    fn insert_range(&mut self, host: CellKey, range: CellRange) -> Result<()> {
        if range.top > range.bottom || range.left > range.right {
            return Err(Error::InvalidFormat(
                "Numbers dependency range is inverted".to_owned(),
            ));
        }
        if self
            .precedents
            .entry(host)
            .or_default()
            .insert(FormulaPrecedent::Range(range))
        {
            self.range_dependents.push((range, host));
        }
        Ok(())
    }

    fn dependents_of(&self, precedent: CellKey) -> HashSet<CellKey> {
        let mut result = self
            .direct_dependents
            .get(&precedent)
            .cloned()
            .unwrap_or_default();
        result.extend(
            self.range_dependents
                .iter()
                .filter_map(|(range, host)| range.contains(precedent).then_some(*host)),
        );
        result
    }

    fn owner_for_cfuuid(&self, uuid: &tsp::CfuuidArchive) -> Option<u32> {
        self.uuid_to_owner.get(&cfuuid_key(uuid)?).copied()
    }

    fn table_for_owner(&self, owner_id: u32) -> Result<&RuntimeTable> {
        let table_id = self.owner_to_table.get(&owner_id).ok_or_else(|| {
            Error::ParseError(format!(
                "Cannot refresh a formula referencing unsupported owner {owner_id}"
            ))
        })?;
        self.tables.get(table_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers formula owner {owner_id} maps to missing table {table_id}"
            ))
        })
    }

    fn dimensions(&self, owner_id: u32) -> Result<(u32, u32)> {
        let table = self.table_for_owner(owner_id)?;
        Ok((
            table.descriptor.model.number_of_rows,
            table.descriptor.model.number_of_columns,
        ))
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum CachedFormulaValue {
    Number(f64),
    Boolean(bool),
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum EvalScalar {
    Empty,
    Number(f64),
    Boolean(bool),
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum EvalValue {
    Scalar(EvalScalar),
    Reference(CellKey),
    Range(CellRange),
}

pub(super) fn refresh_formula_caches_after_cell_write(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<usize> {
    refresh_formula_caches_after_cell_writes(package, table_id, &[(row, column)])
}

pub(super) fn refresh_formula_caches_after_cell_writes(
    package: &mut IWorkPackage,
    table_id: u64,
    coordinates: &[(usize, usize)],
) -> Result<usize> {
    if coordinates.is_empty() {
        return Ok(0);
    }
    let Some(graph) = DependencyGraph::from_package(package)? else {
        return Ok(0);
    };
    let seeds = formula_hosts_for_coordinates(&graph, table_id, coordinates)?;
    let impacted = downstream_formula_hosts(&graph, &seeds, false);
    refresh_formula_hosts(package, &graph, impacted)
}

/// Recalculate specific formula hosts and every formula that depends on them.
///
/// Structural edits use this when a formula's AST is rewritten even though no
/// ordinary cell value was written. Starting with the hosts themselves keeps
/// their native cached values coherent before iWork opens the document.
pub(super) fn refresh_formula_caches_at_hosts(
    package: &mut IWorkPackage,
    table_id: u64,
    coordinates: &[(usize, usize)],
) -> Result<usize> {
    if coordinates.is_empty() {
        return Ok(0);
    }
    let Some(graph) = DependencyGraph::from_package(package)? else {
        return Ok(0);
    };
    let seeds = formula_hosts_for_coordinates(&graph, table_id, coordinates)?;
    let impacted = downstream_formula_hosts(&graph, &seeds, true);
    refresh_formula_hosts(package, &graph, impacted)
}

fn formula_hosts_for_coordinates(
    graph: &DependencyGraph,
    table_id: u64,
    coordinates: &[(usize, usize)],
) -> Result<Vec<CellKey>> {
    let Some(&owner_id) = graph.table_to_owner.get(&table_id) else {
        return Ok(Vec::new());
    };
    coordinates
        .iter()
        .map(|&(row, column)| {
            Ok(CellKey {
                owner_id,
                row: u32::try_from(row)
                    .map_err(|_| Error::ParseError("Numbers row exceeds u32".to_owned()))?,
                column: u32::try_from(column)
                    .map_err(|_| Error::ParseError("Numbers column exceeds u32".to_owned()))?,
            })
        })
        .collect()
}

fn downstream_formula_hosts(
    graph: &DependencyGraph,
    seeds: &[CellKey],
    include_seeds: bool,
) -> HashSet<CellKey> {
    let mut impacted: HashSet<CellKey> = if include_seeds {
        seeds.iter().copied().collect()
    } else {
        HashSet::new()
    };
    let mut queue = seeds.iter().copied().collect::<VecDeque<_>>();
    while let Some(precedent) = queue.pop_front() {
        for dependent in graph.dependents_of(precedent) {
            if graph.owner_to_table.contains_key(&dependent.owner_id) && impacted.insert(dependent)
            {
                queue.push_back(dependent);
            }
        }
    }
    impacted
}

fn refresh_formula_hosts(
    package: &mut IWorkPackage,
    graph: &DependencyGraph,
    impacted: HashSet<CellKey>,
) -> Result<usize> {
    if impacted.is_empty() {
        return Ok(0);
    }

    let mut indegree = impacted
        .iter()
        .copied()
        .map(|host| (host, 0usize))
        .collect::<HashMap<_, _>>();
    let mut downstream = HashMap::<CellKey, HashSet<CellKey>>::new();
    for &precedent in &impacted {
        for dependent in graph.dependents_of(precedent) {
            if impacted.contains(&dependent)
                && downstream.entry(precedent).or_default().insert(dependent)
            {
                *indegree.get_mut(&dependent).ok_or_else(|| {
                    Error::InvalidFormat("Numbers formula topology disappeared".to_owned())
                })? += 1;
            }
        }
    }
    let mut ready = indegree
        .iter()
        .filter_map(|(&host, &degree)| (degree == 0).then_some(host))
        .collect::<BTreeSet<_>>();
    let mut ordered = Vec::with_capacity(impacted.len());
    while let Some(host) = ready.pop_first() {
        ordered.push(host);
        for dependent in downstream.get(&host).into_iter().flatten() {
            let degree = indegree.get_mut(dependent).ok_or_else(|| {
                Error::InvalidFormat("Numbers formula topology disappeared".to_owned())
            })?;
            *degree = degree.checked_sub(1).ok_or_else(|| {
                Error::InvalidFormat("Numbers formula topology underflow".to_owned())
            })?;
            if *degree == 0 {
                ready.insert(*dependent);
            }
        }
    }
    if ordered.len() != impacted.len() {
        return Err(Error::ParseError(
            "Cannot refresh cached results for a cyclic Numbers formula dependency".to_owned(),
        ));
    }

    let mut refreshed = HashMap::with_capacity(ordered.len());
    for &host in &ordered {
        let value = evaluate_formula(package, graph, &refreshed, host)?;
        refreshed.insert(host, value);
    }
    for host in &ordered {
        let value = refreshed.get(host).copied().ok_or_else(|| {
            Error::InvalidFormat("Numbers formula refresh result disappeared".to_owned())
        })?;
        let table_id = graph.table_for_owner(host.owner_id)?.descriptor.object_id;
        let encoded = match value {
            CachedFormulaValue::Number(number) => EncodedValue::FormulaCachedNumber(number),
            CachedFormulaValue::Boolean(boolean) => EncodedValue::FormulaCachedBoolean(boolean),
        };
        set_encoded_cell_value(
            package,
            table_id,
            host.row as usize,
            host.column as usize,
            encoded,
        )?;
    }
    Ok(ordered.len())
}

fn evaluate_formula(
    package: &IWorkPackage,
    graph: &DependencyGraph,
    refreshed: &HashMap<CellKey, CachedFormulaValue>,
    host: CellKey,
) -> Result<CachedFormulaValue> {
    let table = graph.table_for_owner(host.owner_id)?;
    let cell = read_cell(
        package,
        &graph.locations,
        &table.descriptor,
        host.row,
        host.column,
    )?
    .ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers formula host ({}, {}) is missing",
            host.row, host.column
        ))
    })?;
    let formula_id = match cell.stored_value() {
        StoredValue::Formula(identifier) => identifier,
        _ => {
            return Err(Error::InvalidFormat(format!(
                "Numbers dependency host ({}, {}) does not contain a formula",
                host.row, host.column
            )));
        },
    };
    let formula = table.formulas.get(&formula_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers formula table has no entry {formula_id} for host ({}, {})",
            host.row, host.column
        ))
    })?;
    FormulaEvaluator {
        package,
        graph,
        refreshed,
        host,
    }
    .evaluate(formula)
}

struct FormulaEvaluator<'a> {
    package: &'a IWorkPackage,
    graph: &'a DependencyGraph,
    refreshed: &'a HashMap<CellKey, CachedFormulaValue>,
    host: CellKey,
}

impl FormulaEvaluator<'_> {
    fn evaluate(&self, formula: &tsce::FormulaArchive) -> Result<CachedFormulaValue> {
        use tsce::ast_node_array_archive::AstNodeType;

        let mut stack = Vec::<EvalValue>::new();
        for node in &formula.ast_node_array.ast_node {
            match node.ast_node_type() {
                AstNodeType::NumberNode => stack.push(EvalValue::Scalar(EvalScalar::Number(
                    node.ast_number_node_number.ok_or_else(|| {
                        Error::InvalidFormat("Numbers formula number has no value".to_owned())
                    })?,
                ))),
                AstNodeType::BooleanNode => stack.push(EvalValue::Scalar(EvalScalar::Boolean(
                    node.ast_boolean_node_boolean.ok_or_else(|| {
                        Error::InvalidFormat("Numbers formula boolean has no value".to_owned())
                    })?,
                ))),
                AstNodeType::TokenNode => stack.push(EvalValue::Scalar(EvalScalar::Boolean(
                    node.ast_token_node_boolean.ok_or_else(|| {
                        Error::InvalidFormat("Numbers formula token has no value".to_owned())
                    })?,
                ))),
                AstNodeType::DateNode | AstNodeType::DurationNode => {
                    return Err(Error::ParseError(
                        "Cannot refresh cached date or duration formulas yet".to_owned(),
                    ));
                },
                AstNodeType::EmptyArgumentNode => {
                    stack.push(EvalValue::Scalar(EvalScalar::Empty));
                },
                AstNodeType::CellReferenceNode => {
                    stack.push(EvalValue::Reference(self.cell_reference(node)?));
                },
                AstNodeType::LocalCellReferenceNode => {
                    let reference = node
                        .ast_local_cell_reference_node_reference
                        .as_ref()
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "Numbers local formula reference has no coordinate".to_owned(),
                            )
                        })?;
                    let cell = CellKey {
                        owner_id: self.host.owner_id,
                        row: reference.row_handle,
                        column: reference.column_handle,
                    };
                    self.validate_cell(cell)?;
                    stack.push(EvalValue::Reference(cell));
                },
                AstNodeType::CrossTableCellReferenceNode => {
                    let reference = node
                        .ast_cross_table_cell_reference_node_reference
                        .as_ref()
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "Numbers cross-table formula reference has no coordinate"
                                    .to_owned(),
                            )
                        })?;
                    let owner_id = self
                        .graph
                        .owner_for_cfuuid(&reference.table_id)
                        .ok_or_else(|| {
                            Error::ParseError(
                                "Cannot refresh a formula with an unresolved cross-table reference"
                                    .to_owned(),
                            )
                        })?;
                    let cell = CellKey {
                        owner_id,
                        row: reference.row_handle,
                        column: reference.column_handle,
                    };
                    self.validate_cell(cell)?;
                    stack.push(EvalValue::Reference(cell));
                },
                AstNodeType::ColonNode | AstNodeType::ColonNodeWithUids => {
                    let right = pop_value(&mut stack, "range end")?;
                    let left = pop_value(&mut stack, "range start")?;
                    let (EvalValue::Reference(start), EvalValue::Reference(end)) = (left, right)
                    else {
                        return Err(Error::ParseError(
                            "Cannot refresh a formula with non-cell range endpoints".to_owned(),
                        ));
                    };
                    stack.push(EvalValue::Range(self.range_between(start, end)?));
                },
                AstNodeType::ColonTractNode => {
                    stack.push(EvalValue::Range(self.colon_tract_range(node)?));
                },
                AstNodeType::AdditionNode
                | AstNodeType::SubtractionNode
                | AstNodeType::MultiplicationNode
                | AstNodeType::DivisionNode
                | AstNodeType::PowerNode
                | AstNodeType::GreaterThanNode
                | AstNodeType::GreaterThanOrEqualToNode
                | AstNodeType::LessThanNode
                | AstNodeType::LessThanOrEqualToNode
                | AstNodeType::EqualToNode
                | AstNodeType::NotEqualToNode => {
                    let right = self.scalar(pop_value(&mut stack, "binary right operand")?)?;
                    let left = self.scalar(pop_value(&mut stack, "binary left operand")?)?;
                    stack.push(EvalValue::Scalar(evaluate_binary(
                        node.ast_node_type(),
                        left,
                        right,
                    )?));
                },
                AstNodeType::NegationNode => {
                    let value = self.number(pop_value(&mut stack, "negation operand")?)?;
                    stack.push(EvalValue::Scalar(EvalScalar::Number(-value)));
                },
                AstNodeType::PercentNode => {
                    let value = self.number(pop_value(&mut stack, "percent operand")?)?;
                    stack.push(EvalValue::Scalar(EvalScalar::Number(value / PERCENT_SCALE)));
                },
                AstNodeType::FunctionNode => {
                    let identifier = node.ast_function_node_index.ok_or_else(|| {
                        Error::InvalidFormat("Numbers function node has no identifier".to_owned())
                    })?;
                    let count = usize::try_from(node.ast_function_node_num_args.unwrap_or(0))
                        .map_err(|_| {
                            Error::ParseError("Numbers function argument count overflow".to_owned())
                        })?;
                    if stack.len() < count {
                        return Err(Error::InvalidFormat(
                            "Numbers function argument stack underflow".to_owned(),
                        ));
                    }
                    let arguments = stack.split_off(stack.len() - count);
                    stack.push(EvalValue::Scalar(
                        self.evaluate_function(identifier, &arguments)?,
                    ));
                },
                AstNodeType::PlusSignNode
                | AstNodeType::AppendWhitespaceNode
                | AstNodeType::PrependWhitespaceNode => {},
                unsupported => {
                    return Err(Error::ParseError(format!(
                        "Cannot refresh cached results for formula node {unsupported:?}"
                    )));
                },
            }
        }
        if stack.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Numbers formula evaluation left {} stack values",
                stack.len()
            )));
        }
        let value = stack.pop().ok_or_else(|| {
            Error::InvalidFormat("Numbers formula evaluation produced no result".to_owned())
        })?;
        match self.scalar(value)? {
            EvalScalar::Number(number) if number.is_finite() => {
                Ok(CachedFormulaValue::Number(number))
            },
            EvalScalar::Boolean(boolean) => Ok(CachedFormulaValue::Boolean(boolean)),
            EvalScalar::Empty => Ok(CachedFormulaValue::Number(0.0)),
            EvalScalar::Number(_) => Err(Error::ParseError(
                "Numbers formula evaluation produced a non-finite number".to_owned(),
            )),
        }
    }

    fn cell_reference(
        &self,
        node: &tsce::ast_node_array_archive::AstNodeArchive,
    ) -> Result<CellKey> {
        let column = node.ast_column.as_ref().ok_or_else(|| {
            Error::InvalidFormat("Numbers formula cell reference has no column".to_owned())
        })?;
        let row = node.ast_row.as_ref().ok_or_else(|| {
            Error::InvalidFormat("Numbers formula cell reference has no row".to_owned())
        })?;
        let owner_id = node
            .ast_cross_table_reference_extra_info
            .as_ref()
            .map(|reference| {
                self.graph
                    .owner_for_cfuuid(&reference.table_id)
                    .ok_or_else(|| {
                        Error::ParseError(
                            "Cannot refresh a formula with an unresolved cross-table reference"
                                .to_owned(),
                        )
                    })
            })
            .transpose()?
            .unwrap_or(self.host.owner_id);
        let cell = CellKey {
            owner_id,
            row: resolve_coordinate(self.host.row, row.row, row.absolute.unwrap_or(false), "row")?,
            column: resolve_coordinate(
                self.host.column,
                column.column,
                column.absolute.unwrap_or(false),
                "column",
            )?,
        };
        self.validate_cell(cell)?;
        Ok(cell)
    }

    fn range_between(&self, start: CellKey, end: CellKey) -> Result<CellRange> {
        if start.owner_id != end.owner_id {
            return Err(Error::ParseError(
                "Cannot refresh a formula with cross-owner range endpoints".to_owned(),
            ));
        }
        self.validate_cell(start)?;
        self.validate_cell(end)?;
        Ok(CellRange {
            owner_id: start.owner_id,
            top: start.row.min(end.row),
            left: start.column.min(end.column),
            bottom: start.row.max(end.row),
            right: start.column.max(end.column),
        })
    }

    fn colon_tract_range(
        &self,
        node: &tsce::ast_node_array_archive::AstNodeArchive,
    ) -> Result<CellRange> {
        let tract = node.ast_colon_tract.as_ref().ok_or_else(|| {
            Error::InvalidFormat("Numbers colon-tract formula has no range".to_owned())
        })?;
        let sticky = node.ast_sticky_bits.as_ref().ok_or_else(|| {
            Error::InvalidFormat("Numbers colon-tract formula has no sticky bits".to_owned())
        })?;
        let owner_id = node
            .ast_cross_table_reference_extra_info
            .as_ref()
            .map(|reference| {
                self.graph
                    .owner_for_cfuuid(&reference.table_id)
                    .ok_or_else(|| {
                        Error::ParseError(
                            "Cannot refresh a formula with an unresolved range owner".to_owned(),
                        )
                    })
            })
            .transpose()?
            .unwrap_or(self.host.owner_id);
        let (rows, columns) = self.graph.dimensions(owner_id)?;
        let whole_rows = tract
            .absolute_column
            .first()
            .is_some_and(|range| range.range_begin == WHOLE_ROW_COLUMN_SENTINEL);
        let whole_columns = tract
            .absolute_row
            .first()
            .is_some_and(|range| range.range_begin == WHOLE_COLUMN_ROW_SENTINEL);
        let (top, bottom) = if whole_columns {
            (0, rows.saturating_sub(1))
        } else {
            tract_axis(
                &tract.relative_row,
                &tract.absolute_row,
                self.host.row,
                sticky.begin_row_is_absolute,
                sticky.end_row_is_absolute,
                "row",
            )?
        };
        let (left, right) = if whole_rows {
            (0, columns.saturating_sub(1))
        } else {
            tract_axis(
                &tract.relative_column,
                &tract.absolute_column,
                self.host.column,
                sticky.begin_column_is_absolute,
                sticky.end_column_is_absolute,
                "column",
            )?
        };
        let range = CellRange {
            owner_id,
            top: top.min(bottom),
            left: left.min(right),
            bottom: top.max(bottom),
            right: left.max(right),
        };
        self.validate_cell(CellKey {
            owner_id,
            row: range.top,
            column: range.left,
        })?;
        self.validate_cell(CellKey {
            owner_id,
            row: range.bottom,
            column: range.right,
        })?;
        Ok(range)
    }

    fn validate_cell(&self, cell: CellKey) -> Result<()> {
        let (rows, columns) = self.graph.dimensions(cell.owner_id)?;
        if cell.row >= rows || cell.column >= columns {
            return Err(Error::ParseError(format!(
                "Numbers formula reference ({}, {}) is outside {rows}x{columns} owner {}",
                cell.row, cell.column, cell.owner_id
            )));
        }
        Ok(())
    }

    fn scalar(&self, value: EvalValue) -> Result<EvalScalar> {
        match value {
            EvalValue::Scalar(value) => Ok(value),
            EvalValue::Reference(reference) => self.cell_scalar(reference),
            EvalValue::Range(_) => Err(Error::ParseError(
                "Cannot use a range as a scalar cached formula value".to_owned(),
            )),
        }
    }

    fn number(&self, value: EvalValue) -> Result<f64> {
        scalar_number(self.scalar(value)?)
    }

    fn cell_scalar(&self, reference: CellKey) -> Result<EvalScalar> {
        if let Some(value) = self.refreshed.get(&reference) {
            return Ok(match value {
                CachedFormulaValue::Number(number) => EvalScalar::Number(*number),
                CachedFormulaValue::Boolean(boolean) => EvalScalar::Boolean(*boolean),
            });
        }
        let table = self.graph.table_for_owner(reference.owner_id)?;
        let Some(cell) = read_cell(
            self.package,
            &self.graph.locations,
            &table.descriptor,
            reference.row,
            reference.column,
        )?
        else {
            return Ok(EvalScalar::Empty);
        };
        match cell.cached_scalar()? {
            Some(CachedScalar::Number(number)) => Ok(EvalScalar::Number(number.get())),
            Some(CachedScalar::Boolean(boolean)) => Ok(EvalScalar::Boolean(boolean)),
            Some(
                CachedScalar::Date(_) | CachedScalar::Duration(_) | CachedScalar::Unsupported(_),
            ) => Err(Error::ParseError(
                "Cannot refresh a formula that reads a non-scalar cached value".to_owned(),
            )),
            None => Ok(EvalScalar::Empty),
        }
    }

    fn evaluate_function(&self, identifier: u32, arguments: &[EvalValue]) -> Result<EvalScalar> {
        if !matches!(
            identifier,
            AVERAGE_FUNCTION_ID
                | COUNT_FUNCTION_ID
                | MAX_FUNCTION_ID
                | MIN_FUNCTION_ID
                | SUM_FUNCTION_ID
        ) {
            return Err(Error::ParseError(format!(
                "Cannot refresh cached results for Numbers function {}",
                function_name(identifier).unwrap_or("UNKNOWN")
            )));
        }
        let mut values = AggregateAccumulator::default();
        for &argument in arguments {
            self.collect_numbers(argument, &mut values)?;
        }
        let result = match identifier {
            SUM_FUNCTION_ID => values.sum,
            COUNT_FUNCTION_ID => values.count as f64,
            AVERAGE_FUNCTION_ID => {
                if values.count == 0 {
                    return Err(Error::ParseError(
                        "Cannot cache AVERAGE of no numeric values".to_owned(),
                    ));
                }
                values.sum / values.count as f64
            },
            MIN_FUNCTION_ID => values.minimum.unwrap_or(0.0),
            MAX_FUNCTION_ID => values.maximum.unwrap_or(0.0),
            unsupported => {
                return Err(Error::InvalidFormat(format!(
                    "Numbers aggregate identifier {unsupported} escaped validation"
                )));
            },
        };
        if !result.is_finite() {
            return Err(Error::ParseError(
                "Numbers aggregate produced a non-finite cached value".to_owned(),
            ));
        }
        Ok(EvalScalar::Number(result))
    }

    fn collect_numbers(&self, value: EvalValue, output: &mut AggregateAccumulator) -> Result<()> {
        match value {
            EvalValue::Scalar(EvalScalar::Number(number)) => output.push(number)?,
            EvalValue::Scalar(EvalScalar::Boolean(_)) => {
                return Err(Error::ParseError(
                    "Cannot refresh an aggregate with a direct Boolean argument".to_owned(),
                ));
            },
            EvalValue::Scalar(EvalScalar::Empty) => {},
            EvalValue::Reference(reference) => {
                if let EvalScalar::Number(number) = self.cell_scalar(reference)? {
                    output.push(number)?;
                }
            },
            EvalValue::Range(range) => {
                let rows = u64::from(range.bottom - range.top) + 1;
                let columns = u64::from(range.right - range.left) + 1;
                let cells = rows.checked_mul(columns).ok_or_else(|| {
                    Error::ParseError("Numbers aggregate range size overflow".to_owned())
                })?;
                if cells > MAX_CACHE_AGGREGATE_CELLS {
                    return Err(Error::ParseError(format!(
                        "Cannot refresh an aggregate over {cells} cells; the limit is {MAX_CACHE_AGGREGATE_CELLS}"
                    )));
                }
                for row in range.top..=range.bottom {
                    for column in range.left..=range.right {
                        if let EvalScalar::Number(number) = self.cell_scalar(CellKey {
                            owner_id: range.owner_id,
                            row,
                            column,
                        })? {
                            output.push(number)?;
                        }
                    }
                }
            },
        }
        Ok(())
    }
}

fn read_cell(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    descriptor: &TableDescriptor,
    row: u32,
    column: u32,
) -> Result<Option<BncCell>> {
    let tile_size = descriptor
        .model
        .base_data_store
        .tiles
        .tile_size
        .unwrap_or(DEFAULT_TILE_SIZE_ROWS);
    if tile_size == 0 {
        return Err(Error::InvalidFormat(
            "Numbers table declares a zero tile size".to_owned(),
        ));
    }
    let tile_key = row / tile_size;
    let tile_row = row % tile_size;
    let Some(tile_id) = descriptor
        .model
        .base_data_store
        .tiles
        .tiles
        .iter()
        .find(|tile| tile.tileid == tile_key)
        .map(|tile| tile.tile.identifier)
    else {
        return Ok(None);
    };
    let archive_name = locations
        .get(&tile_id)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers tile object {tile_id} is missing")))?;
    read_tile_cell(package, archive_name, tile_id, tile_row, column as usize)?
        .as_deref()
        .map(BncCell::parse)
        .transpose()
        .map_err(Into::into)
}

#[derive(Debug, Default)]
struct AggregateAccumulator {
    sum: f64,
    count: u64,
    minimum: Option<f64>,
    maximum: Option<f64>,
}

impl AggregateAccumulator {
    fn push(&mut self, value: f64) -> Result<()> {
        if !value.is_finite() {
            return Err(Error::ParseError(
                "Numbers aggregate input is not finite".to_owned(),
            ));
        }
        self.sum += value;
        if !self.sum.is_finite() {
            return Err(Error::ParseError(
                "Numbers aggregate sum is not finite".to_owned(),
            ));
        }
        self.count = self
            .count
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("Numbers aggregate count overflow".to_owned()))?;
        self.minimum = Some(self.minimum.map_or(value, |current| current.min(value)));
        self.maximum = Some(self.maximum.map_or(value, |current| current.max(value)));
        Ok(())
    }
}

fn evaluate_binary(
    operator: tsce::ast_node_array_archive::AstNodeType,
    left: EvalScalar,
    right: EvalScalar,
) -> Result<EvalScalar> {
    use tsce::ast_node_array_archive::AstNodeType;

    let result = match operator {
        AstNodeType::AdditionNode => {
            EvalScalar::Number(scalar_number(left)? + scalar_number(right)?)
        },
        AstNodeType::SubtractionNode => {
            EvalScalar::Number(scalar_number(left)? - scalar_number(right)?)
        },
        AstNodeType::MultiplicationNode => {
            EvalScalar::Number(scalar_number(left)? * scalar_number(right)?)
        },
        AstNodeType::DivisionNode => {
            EvalScalar::Number(scalar_number(left)? / scalar_number(right)?)
        },
        AstNodeType::PowerNode => {
            EvalScalar::Number(scalar_number(left)?.powf(scalar_number(right)?))
        },
        AstNodeType::GreaterThanNode => {
            EvalScalar::Boolean(scalar_number(left)? > scalar_number(right)?)
        },
        AstNodeType::GreaterThanOrEqualToNode => {
            EvalScalar::Boolean(scalar_number(left)? >= scalar_number(right)?)
        },
        AstNodeType::LessThanNode => {
            EvalScalar::Boolean(scalar_number(left)? < scalar_number(right)?)
        },
        AstNodeType::LessThanOrEqualToNode => {
            EvalScalar::Boolean(scalar_number(left)? <= scalar_number(right)?)
        },
        AstNodeType::EqualToNode => EvalScalar::Boolean(left == right),
        AstNodeType::NotEqualToNode => EvalScalar::Boolean(left != right),
        _ => {
            return Err(Error::InvalidFormat(format!(
                "Unsupported cached formula binary operator {operator:?}"
            )));
        },
    };
    if let EvalScalar::Number(number) = result
        && !number.is_finite()
    {
        return Err(Error::ParseError(
            "Numbers binary formula produced a non-finite cached value".to_owned(),
        ));
    }
    Ok(result)
}

fn scalar_number(value: EvalScalar) -> Result<f64> {
    match value {
        EvalScalar::Empty => Ok(0.0),
        EvalScalar::Number(number) => Ok(number),
        EvalScalar::Boolean(boolean) => Ok(if boolean { 1.0 } else { 0.0 }),
    }
}

fn pop_value(stack: &mut Vec<EvalValue>, context: &str) -> Result<EvalValue> {
    stack
        .pop()
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers formula {context} is missing")))
}

fn resolve_coordinate(host: u32, value: i32, absolute: bool, axis: &str) -> Result<u32> {
    let coordinate = if absolute {
        i64::from(value)
    } else {
        i64::from(host) + i64::from(value)
    };
    u32::try_from(coordinate).map_err(|_| {
        Error::ParseError(format!(
            "Numbers formula {axis} coordinate is negative or exceeds u32"
        ))
    })
}

fn tract_axis(
    relative: &[tsce::ast_node_array_archive::ast_colon_tract_archive::AstColonTractRelativeRangeArchive],
    absolute: &[tsce::ast_node_array_archive::ast_colon_tract_archive::AstColonTractAbsoluteRangeArchive],
    host: u32,
    begin_absolute: bool,
    end_absolute: bool,
    axis: &str,
) -> Result<(u32, u32)> {
    let relative_range = relative.first();
    let absolute_range = absolute.first();
    let relative_begin = relative_range
        .map(|range| resolve_coordinate(host, range.range_begin, false, axis))
        .transpose()?;
    let relative_end = relative_range
        .map(|range| {
            resolve_coordinate(
                host,
                range.range_end.unwrap_or(range.range_begin),
                false,
                axis,
            )
        })
        .transpose()?;
    let absolute_begin = absolute_range.map(|range| range.range_begin);
    let absolute_end = absolute_range.map(|range| range.range_end.unwrap_or(range.range_begin));
    let begin = if begin_absolute {
        absolute_begin
    } else {
        relative_begin
    }
    .ok_or_else(|| Error::InvalidFormat(format!("Numbers range has no {axis} start")))?;
    let end = if end_absolute {
        absolute_end
    } else {
        relative_end
    }
    .ok_or_else(|| Error::InvalidFormat(format!("Numbers range has no {axis} end")))?;
    Ok((begin, end))
}

fn explicit_coordinate(value: Option<u32>, context: &str) -> Result<u32> {
    value.ok_or_else(|| Error::InvalidFormat(format!("Numbers {context} is missing")))
}

fn uuid_key(uuid: &tsp::Uuid) -> (u64, u64) {
    (uuid.lower, uuid.upper)
}

fn cfuuid_key(uuid: &tsp::CfuuidArchive) -> Option<(u64, u64)> {
    let words = || {
        Some((
            u64::from(uuid.uuid_w0?) | (u64::from(uuid.uuid_w1?) << 32),
            u64::from(uuid.uuid_w2?) | (u64::from(uuid.uuid_w3?) << 32),
        ))
    };
    let bytes = || {
        let bytes: [u8; 16] = uuid.uuid_bytes.as_deref()?.try_into().ok()?;
        let value = u128::from_be_bytes(bytes);
        Some((value as u64, (value >> 64) as u64))
    };
    words().or_else(bytes)
}
