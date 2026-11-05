//! Read the pivot aggregate categories of a workbook (record 0421, the
//! host's `numbers::Draft::pivot_categories`).
//!
//! A pivot table's group-by (`TST.GroupByArchive`, in the calculation
//! engine component) groups rows into a tree of group nodes; each node
//! carries, per aggregated column, the cell coordinate of its aggregate
//! formula. A category is one (group node, aggregated column, summary
//! function) of an enabled group-by whose formula owner kind names exactly
//! one formula owner. Numbers summarizes one source column by several
//! functions with one aggregator and one `column_agg_type` per function
//! (record 0432), so a pair can hold several functions; the host read only
//! the first. Pairs that two group-bys describe differently are dropped as
//! ambiguous, as the host drops them. The walk is a bounded wire read that
//! skips messages it cannot decode, as the host's Prost decode skips them.

use std::collections::hash_map::Entry;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::parse_wire_view;

use super::Package;
use crate::table::pivot::{Aggregate, Category, CategoryKey, Error};

const FORMULA_OWNER_DEPENDENCIES_MESSAGE_TYPE: u32 = 4_008;
const GROUP_BY_MESSAGE_TYPE: u32 = 6_373;
const AGGREGATOR_MESSAGE_TYPE: u32 = 6_382;
const GROUP_NODE_MESSAGE_TYPE: u32 = 6_383;
/// The most group nodes one group-by walk visits (the host's table UID
/// ceiling).
const MAX_NODES: usize = 1_100_000;

/// The fields of one message as (number, wire type, payload); `None` for an
/// undecodable message.
fn fields(data: &[u8]) -> Option<Vec<(u32, u8, &[u8])>> {
    let view = parse_wire_view(data, WireLimits::DEFAULT).ok()?;
    let fields = view
        .fields()
        .map(|field| (field.number(), field.wire_type(), field.payload()))
        .collect();
    Some(fields)
}

fn varint(payload: &[u8]) -> Option<u64> {
    litchi_iwa_common::varint::decode_varint_from_bytes(payload)
        .ok()
        .filter(|(_, length)| *length == payload.len())
        .map(|(value, _)| value)
}

/// The last varint of field `number` (Prost keeps the last occurrence).
fn last_varint(fields: &[(u32, u8, &[u8])], number: u32) -> Option<u64> {
    fields
        .iter()
        .rev()
        .find(|(field, wire_type, _)| *field == number && *wire_type == 0)
        .and_then(|(_, _, payload)| varint(payload))
}

fn last_bytes<'a>(fields: &[(u32, u8, &'a [u8])], number: u32) -> Option<&'a [u8]> {
    fields
        .iter()
        .rev()
        .find(|(field, wire_type, _)| *field == number && *wire_type == 2)
        .map(|(_, _, payload)| *payload)
}

fn last_double(fields: &[(u32, u8, &[u8])], number: u32) -> Option<f64> {
    fields
        .iter()
        .rev()
        .find(|(field, wire_type, _)| *field == number && *wire_type == 1)
        .and_then(|(_, _, payload)| <[u8; 8]>::try_from(*payload).ok())
        .map(f64::from_le_bytes)
}

fn all_bytes<'a>(fields: &[(u32, u8, &'a [u8])], number: u32) -> Vec<&'a [u8]> {
    fields
        .iter()
        .filter(|(field, wire_type, _)| *field == number && *wire_type == 2)
        .map(|(_, _, payload)| *payload)
        .collect()
}

/// A `TSP.UUID` as `(lower, upper)`.
fn uuid(payload: &[u8]) -> Option<(u64, u64)> {
    let fields = fields(payload)?;
    Some((
        last_varint(&fields, 1).unwrap_or(0),
        last_varint(&fields, 2).unwrap_or(0),
    ))
}

/// A `TSP.Reference`'s identifier.
fn reference(payload: &[u8]) -> Option<u64> {
    last_varint(&fields(payload)?, 1)
}

/// A `TSCE.CellCoordinateArchive` with both its row and column.
fn coordinate(payload: &[u8]) -> Option<(u32, u32)> {
    let fields = fields(payload)?;
    Some((
        u32::try_from(last_varint(&fields, 3)?).ok()?,
        u32::try_from(last_varint(&fields, 2)?).ok()?,
    ))
}

/// A group node's label: its group cell value's string, number, Boolean or
/// date (the host's `pivot_group_label`).
fn label(node: &[(u32, u8, &[u8])]) -> Option<String> {
    let value = fields(last_bytes(node, 7)?)?;
    if let Some(string) = last_bytes(&value, 5) {
        let string = fields(string)?;
        return std::str::from_utf8(last_bytes(&string, 1).unwrap_or_default())
            .ok()
            .map(str::to_owned);
    }
    if let Some(number) = last_bytes(&value, 4) {
        if let Some(number) = last_double(&fields(number)?, 1) {
            return Some(number.to_string());
        }
    }
    if let Some(boolean) = last_bytes(&value, 2) {
        let value = last_varint(&fields(boolean)?, 1).unwrap_or(0) != 0;
        return Some(if value { "TRUE" } else { "FALSE" }.to_owned());
    }
    last_bytes(&value, 3)
        .and_then(fields)
        .map(|date| last_double(&date, 1).unwrap_or(0.0).to_string())
}

/// One (group node, column) pair's identity and the native values that must
/// agree when two group-bys describe it.
#[derive(Debug, Clone, PartialEq)]
pub(crate) struct Found {
    /// The group-by's internal formula owner.
    pub(crate) owner: u64,
    /// The grouping-columns cell `(row, column)` of that owner.
    pub(crate) grouping: (u32, u32),
    /// The category's aggregate cell `(row, column)` of that owner.
    pub(crate) aggregate: (u32, u32),
    /// The native summary-function codes, in the pivot's value order.
    pub(crate) aggregate_types: Rc<[u32]>,
    pub(crate) level: i32,
    label: Option<String>,
}

/// Every unambiguous (group node, column) pair of the package by its native
/// identity (the host's `formula_pivot_categories`).
pub(crate) fn categories(package: &Package) -> Result<HashMap<CategoryKey, Found>, Error> {
    let mut engines = package
        .state
        .components
        .catalog()
        .iter()
        .filter(|component| is_calculation_engine(component.name()));
    let Some(engine) = engines.next() else {
        return Ok(HashMap::new());
    };
    if engines.next().is_some() {
        return Err(Error::InvalidSource);
    }
    let mut owners = HashMap::<u64, Vec<u64>>::new();
    let mut groups = Vec::new();
    let mut aggregators = HashMap::new();
    let mut nodes = HashMap::new();
    for object in &engine.archive().objects {
        let identifier = object.archive_info.identifier;
        for message in &object.messages {
            match message.type_ {
                FORMULA_OWNER_DEPENDENCIES_MESSAGE_TYPE => {
                    if let Some(fields) = fields(&message.data)
                        && let Some(kind) = last_varint(&fields, 3)
                    {
                        owners
                            .entry(kind & u64::from(u32::MAX))
                            .or_default()
                            .push(last_varint(&fields, 2).unwrap_or(0) & u64::from(u32::MAX));
                    }
                },
                GROUP_BY_MESSAGE_TYPE => {
                    if let Some(fields) = fields(&message.data) {
                        groups.push(fields);
                    }
                },
                AGGREGATOR_MESSAGE_TYPE => {
                    if let Some(identifier) = identifier
                        && fields(&message.data).is_some()
                    {
                        aggregators.insert(identifier, message.data.as_slice());
                    }
                },
                GROUP_NODE_MESSAGE_TYPE => {
                    if let Some(identifier) = identifier
                        && fields(&message.data).is_some()
                    {
                        nodes.insert(identifier, message.data.as_slice());
                    }
                },
                _ => {},
            }
        }
    }

    let mut result = HashMap::new();
    let mut ambiguous = HashSet::new();
    for group in groups
        .iter()
        .filter(|group| last_varint(group, 6).unwrap_or(0) != 0)
    {
        // `owner_index` is an `int32`: a negative value names no owner kind.
        let Some(kind) = last_varint(group, 14)
            .and_then(|value| i32::try_from(value as i64).ok())
            .and_then(|value| u64::try_from(value).ok())
        else {
            continue;
        };
        let Some(owner) = owners
            .get(&kind)
            .filter(|owners| owners.len() == 1)
            .and_then(|owners| owners.first())
            .copied()
        else {
            continue;
        };
        let Some(grouping) = last_bytes(group, 8).and_then(coordinate) else {
            continue;
        };
        let inline = all_bytes(group, 4);
        let aggregate_columns = if inline.is_empty() {
            all_bytes(group, 17)
                .into_iter()
                .filter_map(reference)
                .filter_map(|identifier| aggregators.get(&identifier).copied())
                .collect::<Vec<_>>()
        } else {
            inline
        };
        if aggregate_columns.is_empty() {
            continue;
        }
        let aggregate_columns = aggregate_columns
            .into_iter()
            .map(|column| {
                fields(column)
                    .and_then(|fields| last_bytes(&fields, 1).map(uuid))
                    .flatten()
            })
            .collect::<Vec<_>>();
        let declared_types = all_bytes(group, 5)
            .into_iter()
            .filter_map(fields)
            .map(|fields| {
                (
                    last_bytes(&fields, 1).and_then(uuid),
                    u32::try_from(last_varint(&fields, 3).unwrap_or(0) & u64::from(u32::MAX))
                        .unwrap_or(u32::MAX),
                )
            })
            .collect::<Vec<_>>();
        // Each column's summary functions, in value order, once per code.
        let aggregate_types = aggregate_columns
            .iter()
            .map(|column| {
                let mut codes = Vec::new();
                for (uid, code) in &declared_types {
                    if uid == column && !codes.contains(code) {
                        codes.push(*code);
                    }
                }
                Rc::<[u32]>::from(codes)
            })
            .collect::<Vec<_>>();
        let root_reference = last_bytes(group, 18).and_then(reference);
        let Some(root) = last_bytes(group, 3)
            .or_else(|| root_reference.and_then(|identifier| nodes.get(&identifier).copied()))
            .and_then(fields)
        else {
            continue;
        };
        let group_by = last_bytes(group, 1).and_then(uuid).unwrap_or((0, 0));
        let mut stack = vec![(root, 0_u32)];
        let mut visited = HashSet::new();
        if let Some(identifier) = root_reference {
            visited.insert(identifier);
        }
        let mut visited_nodes = 0_usize;
        while let Some((node, depth)) = stack.pop() {
            visited_nodes += 1;
            if visited_nodes > MAX_NODES {
                break;
            }
            let Ok(level) = i32::try_from(depth) else {
                continue;
            };
            let node_group = last_bytes(&node, 1).and_then(uuid).unwrap_or((0, 0));
            let coordinates = all_bytes(&node, 5);
            for (index, column) in aggregate_columns.iter().enumerate() {
                let Some(aggregate) = coordinates.get(index).copied().and_then(coordinate) else {
                    continue;
                };
                let types = &aggregate_types[index];
                if types.is_empty() {
                    continue;
                }
                let key = (group_by, column.unwrap_or((0, 0)), node_group);
                if ambiguous.contains(&key) {
                    continue;
                }
                let value = Found {
                    owner,
                    grouping,
                    aggregate,
                    aggregate_types: Rc::clone(types),
                    level,
                    label: label(&node).or_else(|| {
                        (level == 0 && node_group == (1, 0)).then(|| "Grand Total".to_owned())
                    }),
                };
                match result.entry(key) {
                    Entry::Vacant(slot) => {
                        slot.insert(value);
                    },
                    Entry::Occupied(slot) => {
                        if *slot.get() != value {
                            slot.remove();
                            ambiguous.insert(key);
                        }
                    },
                }
            }
            let child_depth = depth.saturating_add(1);
            stack.extend(
                all_bytes(&node, 3)
                    .into_iter()
                    .rev()
                    .filter_map(fields)
                    .map(|child| (child, child_depth)),
            );
            for identifier in all_bytes(&node, 10).into_iter().rev().filter_map(reference) {
                if visited.insert(identifier)
                    && let Some(child) = nodes.get(&identifier).copied().and_then(fields)
                {
                    stack.push((child, child_depth));
                }
            }
        }
    }
    Ok(result)
}

/// The host's calculation-engine component names: `CalculationEngine.iwa`
/// or `CalculationEngine-<digits>[-<digits>…].iwa`.
fn is_calculation_engine(name: &str) -> bool {
    name.rsplit('/').next().is_some_and(|file| {
        file == "CalculationEngine.iwa"
            || file
                .strip_prefix("CalculationEngine-")
                .and_then(|suffix| suffix.strip_suffix(".iwa"))
                .is_some_and(|version| {
                    !version.is_empty()
                        && version.split('-').all(|part| {
                            !part.is_empty() && part.bytes().all(|byte| byte.is_ascii_digit())
                        })
                })
    })
}

impl Package {
    /// Read the workbook's pivot aggregate categories: one per group node,
    /// aggregated column and summary function of each enabled pivot
    /// group-by, ordered by group-by, column, level, label, group and the
    /// pivot's value order.
    ///
    /// # Errors
    ///
    /// [`Error::InvalidSource`] when the package holds several calculation
    /// engines.
    pub fn pivot_categories(&self) -> Result<Vec<Category>, Error> {
        let mut entries = categories(self)?.into_iter().collect::<Vec<_>>();
        entries.sort_by(|(left_key, left), (right_key, right)| {
            left_key
                .0
                .cmp(&right_key.0)
                .then_with(|| left_key.1.cmp(&right_key.1))
                .then_with(|| left.level.cmp(&right.level))
                .then_with(|| left.label.cmp(&right.label))
                .then_with(|| left_key.2.cmp(&right_key.2))
        });
        Ok(entries
            .into_iter()
            .flat_map(|(key, found)| {
                let Found {
                    aggregate_types,
                    label,
                    level,
                    ..
                } = found;
                (0..aggregate_types.len()).map(move |index| {
                    Category::new(
                        key,
                        label.clone(),
                        level,
                        Aggregate::from_code(aggregate_types[index]),
                    )
                })
            })
            .collect())
    }
}
