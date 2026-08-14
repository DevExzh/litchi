//! Generated-free semantic formula lowering and aggregate authoring preflight.

use core::mem::size_of;

use litchi_iwa_protos::numbers_formula_codec::{
    self as codec, BinaryOperator as WireBinary, DecodeLimit, DecodeOptions,
    FormulaWriteAxisReference, FormulaWriteCellReference, FormulaWriteContext,
    FormulaWriteDependencyLimits, FormulaWriteDependencyVisitor, FormulaWriteNode,
    FormulaWriteOwnerUid, FormulaWritePrecedent, FormulaWriteRange, ResolvedFormulaWriteOwner,
};

use crate::{
    Package,
    formula::{BinaryOperator, Node},
    package::function_map::function_identifier,
    table::{
        CellPosition,
        cells::{Change, Error, Input, LimitKind, Path},
    },
};

use super::{budget, cache_commit, lists, resolve, tile};

#[derive(Debug)]
pub(super) struct AuthoredFormula {
    pub(super) change_index: usize,
    pub(super) position: CellPosition,
    pub(super) bytes: Vec<u8>,
    pub(super) precedents: Vec<FormulaWritePrecedent>,
    pub(super) ranges: Vec<FormulaWriteRange>,
}

#[derive(Debug, Default)]
pub(super) struct PreparedFormulaBatch {
    pub(super) formulas: Vec<AuthoredFormula>,
}

impl PreparedFormulaBatch {
    pub(super) fn canonical_match_work(
        &self,
        existing: &cache_commit::ExistingFormulaIndex<'_>,
        path: Path,
    ) -> Result<usize, Error> {
        let per_formula = search_work(existing.payloads.len())
            .checked_add(search_work(existing.entries.len()))
            .and_then(|work| work.checked_add(search_work(existing.cells.len())))
            .and_then(|work| work.checked_add(3))
            .ok_or(Error::InvalidSource { path })?;
        self.formulas.iter().try_fold(0usize, |work, formula| {
            work.checked_add(per_formula)
                .and_then(|work| work.checked_add(formula.bytes.len()))
                .ok_or(Error::InvalidSource { path })
        })
    }

    pub(super) fn canonical_bytes_match_existing(
        &self,
        existing: &cache_commit::ExistingFormulaIndex<'_>,
        path: Path,
    ) -> Result<bool, Error> {
        for formula in &self.formulas {
            let row = formula.position.row();
            let column = formula.position.column();
            let payload_index = existing.payloads.binary_search_by_key(
                &(existing.identity.owner, row, column),
                |payload| {
                    (
                        payload.owner,
                        payload.coordinate.row,
                        payload.coordinate.column,
                    )
                },
            );
            let Ok(payload_index) = payload_index else {
                return Ok(false);
            };
            let payload = &existing.payloads[payload_index];
            let entry_exists = existing
                .entries
                .binary_search_by_key(&(payload.owner, payload.key), |entry| {
                    (entry.owner, entry.key)
                })
                .is_ok();
            if !entry_exists {
                return Err(Error::Verification { path });
            }
            // Full source coverage has already proved the entry/payload byte
            // join. Only the newly authored comparison belongs to this pass.
            if payload.bytes != formula.bytes {
                return Ok(false);
            }
            let cell = existing
                .cells
                .binary_search_by_key(
                    &(
                        payload.owner,
                        payload.coordinate.row,
                        payload.coordinate.column,
                    ),
                    |cell| (cell.owner, cell.coordinate.row, cell.coordinate.column),
                )
                .ok()
                .and_then(|index| existing.cells.get(index));
            if cell.is_none_or(|cell| cell.cache_object == 0) {
                return Err(Error::Verification { path });
            }
        }
        Ok(true)
    }

    pub(super) fn supplied_caches_match_existing(
        &self,
        changes: &[Change],
        existing: &cache_commit::ExistingFormulaIndex<'_>,
        source_package: &Package,
        string_list: &resolve::ListRoute,
        budget: &mut budget::TransactionBudget,
        path: Path,
    ) -> Result<bool, Error> {
        for formula in &self.formulas {
            let change = changes
                .get(formula.change_index)
                .ok_or(Error::Verification { path })?;
            let Some(Input::Formula { cached, .. }) = change.input_ref() else {
                return Ok(false);
            };
            let Some(supplied) = cached else {
                continue;
            };
            let cache_lookup_work = search_work(existing.caches.len())
                .checked_add(2)
                .ok_or(Error::InvalidSource { path })?;
            let cache_lookup_usage = budget::Usage {
                lookups: 1,
                transaction_work: as_u64(cache_lookup_work),
                ..budget::Usage::default()
            };
            budget.authorize(cache_lookup_usage)?;
            let coordinate = (formula.position.row(), formula.position.column());
            let source = existing
                .caches
                .binary_search_by_key(
                    &(existing.identity.owner, coordinate.0, coordinate.1),
                    |cache| (cache.owner, cache.row, cache.column),
                )
                .ok()
                .and_then(|index| existing.caches.get(index));
            budget.record_authorized(cache_lookup_usage)?;
            let Some(source) = source else {
                return Ok(false);
            };
            if source.formula_error.is_some() {
                return Ok(false);
            }
            match supplied.kind() {
                crate::formula::CachedKind::Text(expected) => {
                    let Some(tile::FormulaCacheValue::TextKey(key)) = source.value.as_ref() else {
                        return Ok(false);
                    };
                    if !string_cache_matches_existing(
                        source_package,
                        string_list,
                        *key,
                        expected,
                        budget,
                        path,
                    )? {
                        return Ok(false);
                    }
                },
                supplied => {
                    if !cache_values_match(supplied, source.value.as_ref()) {
                        return Ok(false);
                    }
                },
            }
        }
        Ok(true)
    }
}

fn string_cache_matches_existing(
    source_package: &Package,
    route: &resolve::ListRoute,
    key: u32,
    expected: &str,
    budget: &mut budget::TransactionBudget,
    path: Path,
) -> Result<bool, Error> {
    let mut entries_scanned = 0usize;
    let mut key_occurrences = 0usize;
    let mut matched = false;
    let routes = core::iter::once((false, route.message)).chain(
        route
            .segments
            .iter()
            .copied()
            .map(|message| (true, message)),
    );
    for (segment, message) in routes {
        let payload = super::message_payload(source_package, message, path)?;
        let remaining = budget.remaining()?;
        let limits = super::string_list_limits(source_package, 1, path, remaining)?;
        budget.authorize(remaining)?;
        let probe = if segment {
            lists::string_segment_key_matches(payload, key, expected, limits)
        } else {
            lists::string_key_matches(payload, key, expected, limits)
        };
        let report = match probe {
            Ok(report) => report,
            Err(error) => {
                budget.cancel_authorization();
                return Err(super::map_list_error(error, path));
            },
        };
        let usage = match string_key_match_usage(report, path) {
            Ok(usage) => usage,
            Err(error) => {
                budget.cancel_authorization();
                return Err(error);
            },
        };
        if let Err(error) = budget.record_authorized(usage) {
            budget.cancel_authorization();
            return Err(error);
        }
        entries_scanned = entries_scanned
            .checked_add(report.entries_scanned())
            .ok_or(Error::InvalidSource { path })?;
        key_occurrences = key_occurrences
            .checked_add(report.key_occurrences())
            .ok_or(Error::InvalidSource { path })?;
        if key_occurrences > 1 {
            return Err(Error::InvalidSource { path });
        }
        if report.matched() {
            matched = true;
        }
    }
    if entries_scanned != route.entries {
        return Err(Error::Verification { path });
    }
    Ok(matched)
}

fn string_key_match_usage(
    report: lists::StringKeyMatchReport,
    path: Path,
) -> Result<budget::Usage, Error> {
    let decode = report.decode();
    let transaction_work = decode
        .work_bytes()
        .checked_add(decode.fields())
        .and_then(|work| work.checked_add(decode.references()))
        .and_then(|work| work.checked_add(report.entries_scanned()))
        .and_then(|work| work.checked_add(1))
        .ok_or(Error::InvalidSource { path })?;
    let string_work = report
        .entries_scanned()
        .checked_add(1)
        .ok_or(Error::InvalidSource { path })?;
    Ok(budget::Usage {
        wire_bytes: as_u64(decode.source_bytes()),
        wire_fields: as_u64(decode.fields()),
        wire_work: as_u64(decode.work_bytes()),
        references: as_u64(decode.references()),
        list_reads: 1,
        string_work: as_u64(string_work),
        transaction_work: as_u64(transaction_work),
        ..budget::Usage::default()
    })
}

fn cache_values_match(
    supplied: &crate::formula::CachedKind,
    source: Option<&tile::FormulaCacheValue>,
) -> bool {
    match (supplied, source) {
        (
            crate::formula::CachedKind::Number(left),
            Some(tile::FormulaCacheValue::Number(right)),
        )
        | (crate::formula::CachedKind::Date(left), Some(tile::FormulaCacheValue::Date(right)))
        | (
            crate::formula::CachedKind::Duration(left),
            Some(tile::FormulaCacheValue::Duration(right)),
        ) => left.get().to_bits() == right.get().to_bits(),
        (
            crate::formula::CachedKind::Boolean(left),
            Some(tile::FormulaCacheValue::Boolean(right)),
        ) => left == right,
        // Native text caches carry a StringList key. The governed list join owns
        // resolution to content; a raw key is never accepted as public text authority.
        (crate::formula::CachedKind::Text(_), Some(tile::FormulaCacheValue::TextKey(_)))
        | (_, None)
        | (_, Some(_)) => false,
    }
}

fn search_work(length: usize) -> usize {
    if length < 2 {
        1
    } else {
        usize::try_from(usize::BITS - (length - 1).leading_zeros()).unwrap_or(usize::MAX)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct ReferencedTablePath {
    pub(super) sheet: u32,
    pub(super) table: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct ResolvedOwner {
    pub(super) path: ReferencedTablePath,
    pub(super) uid_lower: u64,
    pub(super) uid_upper: u64,
    pub(super) internal_owner: u32,
    pub(super) rows: u32,
    pub(super) columns: u32,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct LocalOwner<'owners> {
    pub(super) path: ReferencedTablePath,
    pub(super) internal_owner: u32,
    pub(super) owners: &'owners [ResolvedOwner],
}

pub(super) fn referenced_table_paths(
    changes: &[Change],
    staged_nodes: usize,
    budget: &mut budget::TransactionBudget,
    path: Path,
) -> Result<Vec<ReferencedTablePath>, Error> {
    let scan = budget::Usage {
        formula_work: as_u64(staged_nodes),
        transaction_work: as_u64(staged_nodes),
        ..budget::Usage::default()
    };
    budget.authorize(scan)?;
    let occurrences = changes.iter().try_fold(0usize, |total, change| {
        let Some((expression, _)) = change.input_ref().and_then(Input::formula_parts) else {
            return Ok(total);
        };
        count_table_paths(expression.root(), total, path)
    });
    let occurrences = match occurrences {
        Ok(value) => {
            budget.record_authorized(scan)?;
            value
        },
        Err(error) => {
            budget.cancel_authorization();
            return Err(error);
        },
    };
    let bytes = occurrences
        .checked_mul(size_of::<ReferencedTablePath>())
        .ok_or(Error::InvalidSource { path })?;
    let work = occurrences
        .checked_add(sort_work(occurrences, path)?)
        .and_then(|work| work.checked_add(staged_nodes))
        .ok_or(Error::InvalidSource { path })?;
    let envelope = budget::Usage {
        retained_elements: as_u64(occurrences),
        retained_bytes: as_u64(bytes),
        peak_scratch_bytes: as_u64(bytes),
        allocation_events: u64::from(occurrences != 0),
        formula_work: as_u64(staged_nodes),
        transaction_work: as_u64(work),
        ..budget::Usage::default()
    };
    budget.authorize(envelope)?;
    let result = (|| {
        let mut output = Vec::new();
        output
            .try_reserve_exact(occurrences)
            .map_err(|_| Error::Allocation {
                kind: LimitKind::RetainedElements,
                amount: occurrences,
            })?;
        for change in changes {
            if let Some((expression, _)) = change.input_ref().and_then(Input::formula_parts) {
                collect_table_paths(expression.root(), &mut output);
            }
        }
        exact_capacity(&output, occurrences, path)?;
        output.sort_unstable_by_key(|path| (path.sheet, path.table));
        output.dedup();
        Ok(output)
    })();
    match result {
        Ok(output) => {
            budget.record_authorized(envelope)?;
            Ok(output)
        },
        Err(error) => {
            budget.cancel_authorization();
            Err(error)
        },
    }
}

fn count_table_paths(node: &Node, mut total: usize, path: Path) -> Result<usize, Error> {
    match node {
        Node::TableCell(..)
        | Node::TableRange(..)
        | Node::TableRows(..)
        | Node::TableColumns(..) => total = checked_add(total, 1, path)?,
        Node::Function { arguments, .. } => {
            for argument in arguments {
                total = count_table_paths(argument.root(), total, path)?;
            }
        },
        Node::Binary { operands, .. } => {
            for operand in operands {
                total = count_table_paths(operand.root(), total, path)?;
            }
        },
        Node::Negate(operands) | Node::Percent(operands) => {
            total = count_table_paths(operands[0].root(), total, path)?;
        },
        Node::Number(_)
        | Node::Text(_)
        | Node::Boolean(_)
        | Node::Cell(_)
        | Node::Range(..)
        | Node::Rows(..)
        | Node::Columns(..) => {},
    }
    Ok(total)
}

fn collect_table_paths(node: &Node, output: &mut Vec<ReferencedTablePath>) {
    match node {
        Node::TableCell(table, _)
        | Node::TableRange(table, ..)
        | Node::TableRows(table, ..)
        | Node::TableColumns(table, ..) => output.push(table_path(table)),
        Node::Function { arguments, .. } => {
            for argument in arguments {
                collect_table_paths(argument.root(), output);
            }
        },
        Node::Binary { operands, .. } => {
            for operand in operands {
                collect_table_paths(operand.root(), output);
            }
        },
        Node::Negate(operands) | Node::Percent(operands) => {
            collect_table_paths(operands[0].root(), output);
        },
        Node::Number(_)
        | Node::Text(_)
        | Node::Boolean(_)
        | Node::Cell(_)
        | Node::Range(..)
        | Node::Rows(..)
        | Node::Columns(..) => {},
    }
}

#[derive(Debug, Clone, Copy, Default)]
struct Preflight {
    formulas: usize,
    lowered_nodes: usize,
    precedent_upper: usize,
    range_upper: usize,
    external_occurrences: usize,
    output_upper: usize,
    fields_upper: usize,
    work_upper: usize,
    formula_work: usize,
    retained_bytes: usize,
    peak_bytes: usize,
}

fn preflight(changes: &[Change], rows: u32, columns: u32, path: Path) -> Result<Preflight, Error> {
    let mut total = Preflight::default();
    for change in changes {
        let Some((expression, _cached)) = change.input_ref().and_then(Input::formula_parts) else {
            continue;
        };
        let stats = expression_preflight(
            expression.root(),
            expression.owned_bytes(),
            rows,
            columns,
            path,
        )?;
        total.formulas = checked_add(total.formulas, 1, path)?;
        total.lowered_nodes = checked_add(total.lowered_nodes, stats.lowered_nodes, path)?;
        total.precedent_upper = checked_add(total.precedent_upper, stats.precedent_upper, path)?;
        total.range_upper = checked_add(total.range_upper, stats.range_upper, path)?;
        total.external_occurrences =
            checked_add(total.external_occurrences, stats.external_occurrences, path)?;
        total.output_upper = checked_add(total.output_upper, stats.output_upper, path)?;
        total.fields_upper = checked_add(total.fields_upper, stats.fields_upper, path)?;
        total.work_upper = checked_add(total.work_upper, stats.work_upper, path)?;
        total.formula_work = checked_add(total.formula_work, stats.formula_work, path)?;
        total.retained_bytes = checked_add(total.retained_bytes, stats.retained_bytes, path)?;
        total.peak_bytes = total.peak_bytes.max(stats.peak_bytes);
    }
    total.retained_bytes = checked_add(
        total.retained_bytes,
        total
            .formulas
            .checked_mul(size_of::<AuthoredFormula>())
            .ok_or(Error::InvalidSource { path })?,
        path,
    )?;
    Ok(total)
}

fn expression_preflight(
    root: &Node,
    text_bytes: usize,
    rows: u32,
    columns: u32,
    path: Path,
) -> Result<Preflight, Error> {
    let mut stats = Preflight {
        formulas: 1,
        output_upper: 32,
        fields_upper: 1,
        work_upper: 96,
        ..Preflight::default()
    };
    preflight_node(root, rows, columns, &mut stats, path)?;
    let node_bytes = stats
        .lowered_nodes
        .checked_mul(size_of::<FormulaWriteNode<'_>>())
        .ok_or(Error::InvalidSource { path })?;
    let precedent_bytes = stats
        .precedent_upper
        .checked_mul(size_of::<FormulaWritePrecedent>())
        .ok_or(Error::InvalidSource { path })?;
    let range_bytes = stats
        .range_upper
        .checked_mul(size_of::<FormulaWriteRange>())
        .ok_or(Error::InvalidSource { path })?;
    // Every canonical node can fit within this conservative fixed header plus
    // the caller-owned literal bytes. The strict codec will settle exact use.
    stats.output_upper = checked_add(
        stats.output_upper,
        stats
            .lowered_nodes
            .checked_mul(256)
            .and_then(|bytes| bytes.checked_add(text_bytes))
            .ok_or(Error::InvalidSource { path })?,
        path,
    )?;
    stats.fields_upper = checked_add(
        stats.fields_upper,
        stats
            .lowered_nodes
            .checked_mul(32)
            .ok_or(Error::InvalidSource { path })?,
        path,
    )?;
    let precedent_sort_work = sort_work(stats.precedent_upper, path)?;
    let codec_work = stats
        .output_upper
        .checked_mul(4)
        .and_then(|work| work.checked_add(precedent_sort_work))
        .ok_or(Error::InvalidSource { path })?;
    stats.work_upper = checked_add(stats.work_upper, codec_work, path)?;
    stats.formula_work = checked_add(
        stats.lowered_nodes,
        checked_add(stats.precedent_upper, precedent_sort_work, path)?,
        path,
    )?;
    stats.retained_bytes = checked_add(
        checked_add(stats.output_upper, precedent_bytes, path)?,
        range_bytes,
        path,
    )?;
    stats.peak_bytes = checked_add(
        checked_add(
            checked_add(node_bytes, precedent_bytes, path)?,
            range_bytes,
            path,
        )?,
        stats.output_upper,
        path,
    )?;
    Ok(stats)
}

fn preflight_node(
    node: &Node,
    local_rows: u32,
    local_columns: u32,
    stats: &mut Preflight,
    path: Path,
) -> Result<(), Error> {
    match node {
        Node::Number(value) => {
            stats.lowered_nodes = checked_add(
                stats.lowered_nodes,
                if value.get() != 0.0 && value.get().is_sign_negative() {
                    2
                } else {
                    1
                },
                path,
            )?;
        },
        Node::Cell(_) | Node::TableCell(_, _) => {
            stats.lowered_nodes = checked_add(stats.lowered_nodes, 1, path)?;
            stats.precedent_upper = checked_add(stats.precedent_upper, 1, path)?;
            stats.external_occurrences = checked_add(
                stats.external_occurrences,
                usize::from(matches!(node, Node::TableCell(_, _))),
                path,
            )?;
        },
        Node::Range(start, end) | Node::TableRange(_, start, end) => {
            stats.lowered_nodes = checked_add(stats.lowered_nodes, 1, path)?;
            let (start_row, start_column) = start.coordinates();
            let (end_row, end_column) = end.coordinates();
            let rows = end_row
                .checked_sub(start_row)
                .and_then(|value| value.checked_add(1))
                .ok_or(Error::InvalidSource { path })?;
            let columns = end_column
                .checked_sub(start_column)
                .and_then(|value| value.checked_add(1))
                .ok_or(Error::InvalidSource { path })?;
            stats.precedent_upper = checked_add(
                stats.precedent_upper,
                rows.checked_mul(columns)
                    .ok_or(Error::InvalidSource { path })?,
                path,
            )?;
            stats.range_upper = checked_add(stats.range_upper, 1, path)?;
            stats.external_occurrences = checked_add(
                stats.external_occurrences,
                usize::from(matches!(node, Node::TableRange(_, _, _))),
                path,
            )?;
        },
        Node::Rows(start, end) | Node::TableRows(_, start, end) => {
            stats.lowered_nodes = checked_add(stats.lowered_nodes, 1, path)?;
            let (start, _) = start.parts();
            let (end, _) = end.parts();
            let selected = end
                .checked_sub(start)
                .and_then(|value| value.checked_add(1))
                .ok_or(Error::InvalidSource { path })?;
            let perpendicular = match node {
                Node::Rows(_, _) => local_columns,
                Node::TableRows(table, _, _) => table.dimensions().columns(),
                _ => unreachable!(),
            };
            stats.precedent_upper = checked_add(
                stats.precedent_upper,
                selected
                    .checked_mul(
                        usize::try_from(perpendicular)
                            .map_err(|_| Error::InvalidSource { path })?,
                    )
                    .ok_or(Error::InvalidSource { path })?,
                path,
            )?;
            stats.range_upper = checked_add(stats.range_upper, 1, path)?;
            stats.external_occurrences = checked_add(
                stats.external_occurrences,
                usize::from(matches!(node, Node::TableRows(_, _, _))),
                path,
            )?;
        },
        Node::Columns(start, end) | Node::TableColumns(_, start, end) => {
            stats.lowered_nodes = checked_add(stats.lowered_nodes, 1, path)?;
            let (start, _) = start.parts();
            let (end, _) = end.parts();
            let selected = end
                .checked_sub(start)
                .and_then(|value| value.checked_add(1))
                .ok_or(Error::InvalidSource { path })?;
            let perpendicular = match node {
                Node::Columns(_, _) => local_rows,
                Node::TableColumns(table, _, _) => table.dimensions().rows(),
                _ => unreachable!(),
            };
            stats.precedent_upper = checked_add(
                stats.precedent_upper,
                selected
                    .checked_mul(
                        usize::try_from(perpendicular)
                            .map_err(|_| Error::InvalidSource { path })?,
                    )
                    .ok_or(Error::InvalidSource { path })?,
                path,
            )?;
            stats.range_upper = checked_add(stats.range_upper, 1, path)?;
            stats.external_occurrences = checked_add(
                stats.external_occurrences,
                usize::from(matches!(node, Node::TableColumns(_, _, _))),
                path,
            )?;
        },
        Node::Text(_) | Node::Boolean(_) => {
            stats.lowered_nodes = checked_add(stats.lowered_nodes, 1, path)?;
        },
        Node::Function { arguments, .. } => {
            for argument in arguments {
                preflight_node(argument.root(), local_rows, local_columns, stats, path)?;
            }
            stats.lowered_nodes = checked_add(stats.lowered_nodes, 1, path)?;
        },
        Node::Binary { operands, .. } => {
            for operand in operands {
                preflight_node(operand.root(), local_rows, local_columns, stats, path)?;
            }
            stats.lowered_nodes = checked_add(stats.lowered_nodes, 1, path)?;
        },
        Node::Negate(operands) | Node::Percent(operands) => {
            preflight_node(operands[0].root(), local_rows, local_columns, stats, path)?;
            stats.lowered_nodes = checked_add(stats.lowered_nodes, 1, path)?;
        },
    }
    Ok(())
}

pub(super) fn prepare(
    source: &Package,
    changes: &[Change],
    staged_nodes: usize,
    owner: LocalOwner<'_>,
    rows: u32,
    columns: u32,
    budget: &mut budget::TransactionBudget,
    path: Path,
) -> Result<PreparedFormulaBatch, Error> {
    let recorded_nodes = changes.iter().try_fold(0usize, |total, change| {
        let Some((expression, _)) = change.input_ref().and_then(Input::formula_parts) else {
            return Ok(total);
        };
        checked_add(total, expression.node_count(), path)
    })?;
    if recorded_nodes != staged_nodes {
        return Err(Error::Verification { path });
    }
    let scan = budget::Usage {
        formula_work: as_u64(staged_nodes),
        transaction_work: as_u64(staged_nodes),
        ..budget::Usage::default()
    };
    budget.authorize(scan)?;
    let preflight = match preflight(changes, rows, columns, path) {
        Ok(preflight) => {
            budget.record_authorized(scan)?;
            preflight
        },
        Err(error) => {
            budget.cancel_authorization();
            return Err(error);
        },
    };
    if preflight.formulas == 0 {
        return Ok(PreparedFormulaBatch::default());
    }
    let envelope = budget::Usage {
        retained_elements: as_u64(
            preflight
                .formulas
                .checked_add(preflight.precedent_upper)
                .and_then(|count| count.checked_add(preflight.range_upper))
                .ok_or(Error::InvalidSource { path })?,
        ),
        retained_bytes: as_u64(preflight.retained_bytes),
        peak_scratch_bytes: as_u64(
            preflight
                .peak_bytes
                .checked_add(
                    owner
                        .owners
                        .len()
                        .checked_mul(
                            size_of::<ResolvedFormulaWriteOwner>()
                                .checked_add(size_of::<u32>())
                                .ok_or(Error::InvalidSource { path })?,
                        )
                        .ok_or(Error::InvalidSource { path })?,
                )
                .ok_or(Error::InvalidSource { path })?,
        ),
        allocation_events: as_u64(
            preflight
                .formulas
                .checked_mul(4)
                .and_then(|n| n.checked_add(1 + usize::from(!owner.owners.is_empty()) * 2))
                .ok_or(Error::InvalidSource { path })?,
        ),
        wire_bytes: as_u64(preflight.output_upper),
        wire_fields: as_u64(preflight.fields_upper),
        wire_work: as_u64(preflight.work_upper),
        formula_nodes: as_u64(preflight.lowered_nodes),
        formula_edges: as_u64(preflight.precedent_upper),
        authored_formula_writes: as_u64(preflight.formulas),
        formula_work: as_u64(
            preflight
                .formula_work
                .checked_add(staged_nodes)
                .ok_or(Error::InvalidSource { path })?,
        ),
        transaction_work: as_u64(
            preflight
                .work_upper
                .checked_add(preflight.formula_work)
                .and_then(|work| {
                    owner_registry_work(owner.owners.len())
                        .and_then(|owner_work| work.checked_add(owner_work))
                })
                .and_then(|work| {
                    preflight
                        .external_occurrences
                        .checked_mul(owner.owners.len())
                        .and_then(|lookup_work| work.checked_add(lookup_work))
                })
                .and_then(|work| work.checked_add(staged_nodes))
                .ok_or(Error::InvalidSource { path })?,
        ),
        ..budget::Usage::default()
    };
    budget.authorize(envelope)?;
    let result = prepare_authorized(
        source,
        changes,
        staged_nodes,
        owner,
        rows,
        columns,
        envelope,
        preflight,
        path,
    );
    match result {
        Ok((prepared, actual)) => {
            budget.record_authorized(actual)?;
            Ok(prepared)
        },
        Err(error) => {
            budget.cancel_authorization();
            Err(error)
        },
    }
}

fn prepare_authorized(
    _source: &Package,
    changes: &[Change],
    staged_nodes: usize,
    owner: LocalOwner<'_>,
    rows: u32,
    columns: u32,
    limits: budget::Usage,
    aggregate: Preflight,
    path: Path,
) -> Result<(PreparedFormulaBatch, budget::Usage), Error> {
    let wire_owners = wire_owners(owner, path)?;
    let mut formulas = Vec::new();
    formulas
        .try_reserve_exact(aggregate.formulas)
        .map_err(|_| Error::Allocation {
            kind: LimitKind::RetainedElements,
            amount: aggregate.formulas,
        })?;
    exact_capacity(&formulas, aggregate.formulas, path)?;
    let owner_scratch_bytes = wire_owners
        .capacity()
        .checked_mul(
            size_of::<ResolvedFormulaWriteOwner>()
                .checked_add(size_of::<u32>())
                .ok_or(Error::InvalidSource { path })?,
        )
        .ok_or(Error::InvalidSource { path })?;
    let mut actual = budget::Usage {
        peak_scratch_bytes: as_u64(owner_scratch_bytes),
        allocation_events: u64::from(aggregate.formulas != 0)
            .checked_add(u64::from(!wire_owners.is_empty()) * 2)
            .ok_or(Error::InvalidSource { path })?,
        formula_work: as_u64(staged_nodes),
        transaction_work: as_u64(
            owner_registry_work(owner.owners.len())
                .and_then(|work| {
                    aggregate
                        .external_occurrences
                        .checked_mul(owner.owners.len())
                        .and_then(|lookups| work.checked_add(lookups))
                })
                .and_then(|work| work.checked_add(staged_nodes))
                .ok_or(Error::InvalidSource { path })?,
        ),
        ..budget::Usage::default()
    };
    for (change_index, change) in changes.iter().enumerate() {
        let Some((expression, _cached)) = change.input_ref().and_then(Input::formula_parts) else {
            continue;
        };
        let stats = expression_preflight(
            expression.root(),
            expression.owned_bytes(),
            rows,
            columns,
            path,
        )?;
        let mut nodes = Vec::new();
        nodes
            .try_reserve_exact(stats.lowered_nodes)
            .map_err(|_| Error::Allocation {
                kind: LimitKind::RetainedElements,
                amount: stats.lowered_nodes,
            })?;
        exact_capacity(&nodes, stats.lowered_nodes, path)?;
        let mut precedents = Vec::new();
        precedents
            .try_reserve_exact(stats.precedent_upper)
            .map_err(|_| Error::Allocation {
                kind: LimitKind::RetainedElements,
                amount: stats.precedent_upper,
            })?;
        exact_capacity(&precedents, stats.precedent_upper, path)?;
        let mut ranges = Vec::new();
        ranges
            .try_reserve_exact(stats.range_upper)
            .map_err(|_| Error::Allocation {
                kind: LimitKind::RetainedElements,
                amount: stats.range_upper,
            })?;
        exact_capacity(&ranges, stats.range_upper, path)?;
        lower(expression.root(), owner, &mut nodes, path)?;
        if nodes.len() != stats.lowered_nodes {
            return Err(Error::Verification { path });
        }
        let options = DecodeOptions::new(
            as_usize(limits.wire_bytes, path)?,
            as_usize(limits.wire_fields, path)?,
            as_usize(limits.wire_work, path)?,
            32,
            as_usize(limits.formula_nodes, path)?,
            expression.owned_bytes().max(1),
        );
        let context = FormulaWriteContext::new(
            owner.internal_owner,
            change.position().row(),
            change.position().column(),
            rows,
            columns,
        );
        let plan = codec::plan_resolved_formula_archive(
            &nodes,
            context,
            &wire_owners,
            FormulaWriteDependencyLimits::new(stats.precedent_upper, stats.range_upper),
            options,
        )
        .map_err(|error| map_codec(error, path))?;
        let requirements = plan.requirements();
        let mut facts = DependencyCollector {
            precedents: &mut precedents,
            ranges: &mut ranges,
        };
        let (bytes, report) =
            codec::execute_formula_archive_plan_with_visitor(plan, options, &mut facts)
                .map_err(|error| map_codec(error, path))?;
        if report.requirements() != requirements {
            return Err(Error::Verification { path });
        }
        exact_capacity(&bytes, requirements.output_bytes(), path)?;
        if precedents.len() != requirements.precedent_count()
            || ranges.len() != requirements.range_count()
        {
            return Err(Error::Verification { path });
        }
        merge_requirement_usage(&mut actual, requirements, precedents.len(), path)?;
        let simultaneous_scratch = nodes
            .capacity()
            .checked_mul(size_of::<FormulaWriteNode<'_>>())
            .and_then(|bytes| {
                precedents
                    .capacity()
                    .checked_mul(size_of::<FormulaWritePrecedent>())
                    .and_then(|precedent_bytes| bytes.checked_add(precedent_bytes))
            })
            .and_then(|bytes| {
                ranges
                    .capacity()
                    .checked_mul(size_of::<FormulaWriteRange>())
                    .and_then(|range_bytes| bytes.checked_add(range_bytes))
            })
            .and_then(|bytes| bytes.checked_add(requirements.output_bytes()))
            .ok_or(Error::InvalidSource { path })?;
        actual.peak_scratch_bytes = actual.peak_scratch_bytes.max(as_u64(
            simultaneous_scratch
                .checked_add(owner_scratch_bytes)
                .ok_or(Error::InvalidSource { path })?,
        ));
        actual.formula_work = actual
            .formula_work
            .checked_add(as_u64(sort_work(stats.precedent_upper, path)?))
            .ok_or(Error::InvalidSource { path })?;
        actual.transaction_work = actual
            .transaction_work
            .checked_add(as_u64(sort_work(stats.precedent_upper, path)?))
            .ok_or(Error::InvalidSource { path })?;
        actual.allocation_events = actual
            .allocation_events
            .checked_add(4)
            .ok_or(Error::InvalidSource { path })?;
        formulas.push(AuthoredFormula {
            change_index,
            position: change.position(),
            bytes,
            precedents,
            ranges,
        });
    }
    actual.retained_elements = formulas
        .iter()
        .try_fold(formulas.len(), |total, formula| {
            total
                .checked_add(formula.precedents.capacity())
                .and_then(|n| n.checked_add(formula.ranges.capacity()))
                .ok_or(Error::InvalidSource { path })
        })
        .map(as_u64)?;
    actual.retained_bytes = formulas
        .iter()
        .try_fold(
            formulas
                .len()
                .checked_mul(size_of::<AuthoredFormula>())
                .ok_or(Error::InvalidSource { path })?,
            |total, formula| {
                total
                    .checked_add(formula.bytes.capacity())
                    .and_then(|n| {
                        formula
                            .precedents
                            .capacity()
                            .checked_mul(size_of::<FormulaWritePrecedent>())
                            .and_then(|p| n.checked_add(p))
                    })
                    .and_then(|n| {
                        formula
                            .ranges
                            .capacity()
                            .checked_mul(size_of::<FormulaWriteRange>())
                            .and_then(|r| n.checked_add(r))
                    })
                    .ok_or(Error::InvalidSource { path })
            },
        )
        .map(as_u64)?;
    Ok((PreparedFormulaBatch { formulas }, actual))
}

fn lower<'text>(
    node: &'text Node,
    owner: LocalOwner<'_>,
    output: &mut Vec<FormulaWriteNode<'text>>,
    path: Path,
) -> Result<(), Error> {
    match node {
        Node::Number(value) => {
            let value = value.get();
            if value != 0.0 && value.is_sign_negative() {
                output.push(FormulaWriteNode::Number(value.abs()));
                output.push(FormulaWriteNode::Negation);
            } else {
                output.push(FormulaWriteNode::Number(value));
            }
        },
        Node::Text(text) => output.push(FormulaWriteNode::Text(text)),
        Node::Boolean(value) => output.push(FormulaWriteNode::Boolean(*value)),
        Node::Cell(reference) => output.push(FormulaWriteNode::ResolvedCellReference {
            owner: None,
            reference: wire_cell(*reference, path)?,
        }),
        Node::TableCell(table, reference) => {
            output.push(FormulaWriteNode::ResolvedCellReference {
                owner: owner_uid(owner, table_path(table), path)?,
                reference: wire_cell(*reference, path)?,
            });
        },
        Node::Function { name, arguments } => {
            for argument in arguments {
                lower(argument.root(), owner, output, path)?;
            }
            let identifier = function_identifier(name).ok_or(Error::UnsupportedDependency {
                path,
                kind: crate::package::table_cells::DependencyKind::Formula,
            })?;
            output.push(FormulaWriteNode::Function {
                identifier,
                argument_count: u32::try_from(arguments.len())
                    .map_err(|_| Error::InvalidSource { path })?,
            });
        },
        Node::Binary { operator, operands } => {
            for operand in operands {
                lower(operand.root(), owner, output, path)?;
            }
            output.push(FormulaWriteNode::Binary(lower_binary(*operator)));
        },
        Node::Negate(operands) => {
            lower(operands[0].root(), owner, output, path)?;
            output.push(FormulaWriteNode::Negation);
        },
        Node::Percent(operands) => {
            lower(operands[0].root(), owner, output, path)?;
            output.push(FormulaWriteNode::Percent);
        },
        Node::Range(start, end) => output.push(FormulaWriteNode::ResolvedRange {
            owner: None,
            start: wire_cell(*start, path)?,
            end: wire_cell(*end, path)?,
        }),
        Node::TableRange(table, start, end) => {
            output.push(FormulaWriteNode::ResolvedRange {
                owner: owner_uid(owner, table_path(table), path)?,
                start: wire_cell(*start, path)?,
                end: wire_cell(*end, path)?,
            });
        },
        Node::Rows(start, end) => output.push(FormulaWriteNode::WholeRows {
            owner: None,
            start: wire_axis(*start, path)?,
            end: wire_axis(*end, path)?,
        }),
        Node::Columns(start, end) => output.push(FormulaWriteNode::WholeColumns {
            owner: None,
            start: wire_axis(*start, path)?,
            end: wire_axis(*end, path)?,
        }),
        Node::TableRows(table, start, end) => output.push(FormulaWriteNode::WholeRows {
            owner: owner_uid(owner, table_path(table), path)?,
            start: wire_axis(*start, path)?,
            end: wire_axis(*end, path)?,
        }),
        Node::TableColumns(table, start, end) => {
            output.push(FormulaWriteNode::WholeColumns {
                owner: owner_uid(owner, table_path(table), path)?,
                start: wire_axis(*start, path)?,
                end: wire_axis(*end, path)?,
            });
        },
    }
    Ok(())
}

fn wire_cell(
    reference: crate::formula::CellReference,
    path: Path,
) -> Result<FormulaWriteCellReference, Error> {
    let (row, column) = reference.coordinates();
    let (absolute_row, absolute_column) = reference.modes();
    Ok(FormulaWriteCellReference::new(
        u32::try_from(row).map_err(|_| Error::InvalidSource { path })?,
        u32::try_from(column).map_err(|_| Error::InvalidSource { path })?,
        absolute_row,
        absolute_column,
    ))
}

fn wire_axis(
    reference: crate::formula::AxisReference,
    path: Path,
) -> Result<FormulaWriteAxisReference, Error> {
    let (index, absolute) = reference.parts();
    Ok(FormulaWriteAxisReference::new(
        u32::try_from(index).map_err(|_| Error::InvalidSource { path })?,
        absolute,
    ))
}

fn table_path(table: &crate::formula::Table) -> ReferencedTablePath {
    let Path::Table { sheet, table } = table.path() else {
        unreachable!("formula table handles are always concrete")
    };
    ReferencedTablePath { sheet, table }
}

fn owner_uid(
    owner: LocalOwner<'_>,
    target: ReferencedTablePath,
    path: Path,
) -> Result<Option<FormulaWriteOwnerUid>, Error> {
    if target == owner.path {
        return Ok(None);
    }
    let resolved = owner
        .owners
        .iter()
        .find(|candidate| candidate.path == target)
        .ok_or(Error::InvalidSource { path })?;
    Ok(Some(FormulaWriteOwnerUid::from_halves(
        resolved.uid_lower,
        resolved.uid_upper,
    )))
}

fn wire_owners(owner: LocalOwner<'_>, path: Path) -> Result<Vec<ResolvedFormulaWriteOwner>, Error> {
    let mut output = Vec::new();
    let mut internal_owners = Vec::new();
    output
        .try_reserve_exact(owner.owners.len())
        .map_err(|_| Error::Allocation {
            kind: LimitKind::RetainedElements,
            amount: owner.owners.len(),
        })?;
    internal_owners
        .try_reserve_exact(owner.owners.len())
        .map_err(|_| Error::Allocation {
            kind: LimitKind::PeakScratchBytes,
            amount: owner
                .owners
                .len()
                .checked_mul(size_of::<u32>())
                .unwrap_or(usize::MAX),
        })?;
    exact_capacity(&output, owner.owners.len(), path)?;
    exact_capacity(&internal_owners, owner.owners.len(), path)?;
    let mut previous_uid = None;
    for resolved in owner.owners {
        let uid = (resolved.uid_lower, resolved.uid_upper);
        if resolved.path == owner.path
            || resolved.internal_owner == 0
            || resolved.uid_lower == 0 && resolved.uid_upper == 0
            || resolved.rows == 0
            || resolved.columns == 0
            || previous_uid.is_some_and(|previous| previous >= uid)
        {
            return Err(Error::InvalidSource { path });
        }
        previous_uid = Some(uid);
        internal_owners.push(resolved.internal_owner);
        output.push(ResolvedFormulaWriteOwner::new(
            FormulaWriteOwnerUid::from_halves(resolved.uid_lower, resolved.uid_upper),
            resolved.internal_owner,
            resolved.rows,
            resolved.columns,
        ));
    }
    internal_owners.sort_unstable();
    if internal_owners
        .windows(2)
        .any(|owners| owners[0] >= owners[1])
    {
        return Err(Error::InvalidSource { path });
    }
    Ok(output)
}

struct DependencyCollector<'facts> {
    precedents: &'facts mut Vec<FormulaWritePrecedent>,
    ranges: &'facts mut Vec<FormulaWriteRange>,
}

impl FormulaWriteDependencyVisitor for DependencyCollector<'_> {
    fn visit_precedent(
        &mut self,
        precedent: FormulaWritePrecedent,
    ) -> Result<(), codec::DecodeError> {
        self.precedents.push(precedent);
        Ok(())
    }

    fn visit_range(&mut self, range: FormulaWriteRange) -> Result<(), codec::DecodeError> {
        self.ranges.push(range);
        Ok(())
    }
}

const fn lower_binary(operator: BinaryOperator) -> WireBinary {
    match operator {
        BinaryOperator::Add => WireBinary::Add,
        BinaryOperator::Subtract => WireBinary::Subtract,
        BinaryOperator::Multiply => WireBinary::Multiply,
        BinaryOperator::Divide => WireBinary::Divide,
        BinaryOperator::Power => WireBinary::Power,
        BinaryOperator::Concatenate => WireBinary::Concatenate,
        BinaryOperator::GreaterThan => WireBinary::GreaterThan,
        BinaryOperator::GreaterThanOrEqual => WireBinary::GreaterThanOrEqual,
        BinaryOperator::LessThan => WireBinary::LessThan,
        BinaryOperator::LessThanOrEqual => WireBinary::LessThanOrEqual,
        BinaryOperator::Equal => WireBinary::Equal,
        BinaryOperator::NotEqual => WireBinary::NotEqual,
    }
}

fn merge_requirement_usage(
    usage: &mut budget::Usage,
    requirements: codec::FormulaWriteRequirements,
    precedents: usize,
    path: Path,
) -> Result<(), Error> {
    macro_rules! add {
        ($field:ident, $value:expr) => {
            usage.$field = usage
                .$field
                .checked_add(as_u64($value))
                .ok_or(Error::InvalidSource { path })?;
        };
    }
    add!(wire_bytes, requirements.output_bytes());
    add!(wire_fields, requirements.fields());
    add!(wire_work, requirements.work_bytes());
    add!(formula_nodes, requirements.nodes());
    add!(formula_edges, precedents);
    add!(authored_formula_writes, 1);
    add!(
        formula_work,
        requirements
            .nodes()
            .checked_add(precedents)
            .ok_or(Error::InvalidSource { path })?
    );
    add!(
        transaction_work,
        requirements
            .work_bytes()
            .checked_add(precedents)
            .ok_or(Error::InvalidSource { path })?
    );
    usage.peak_scratch_bytes = usage
        .peak_scratch_bytes
        .max(as_u64(requirements.output_bytes()));
    Ok(())
}

fn map_codec(error: codec::DecodeError, path: Path) -> Error {
    let (kind, observed, maximum) = match error.resource_limit() {
        Some(DecodeLimit::Bytes { observed, maximum }) => {
            (LimitKind::RetainedBytes, observed, maximum)
        },
        Some(DecodeLimit::Fields { observed, maximum }) => {
            (LimitKind::WireFields, observed, maximum)
        },
        Some(DecodeLimit::Work { observed, maximum }) => (LimitKind::WireWork, observed, maximum),
        Some(DecodeLimit::Nesting { observed, maximum }) => {
            (LimitKind::FormulaWork, observed as usize, maximum as usize)
        },
        Some(DecodeLimit::Nodes { observed, maximum }) => {
            (LimitKind::FormulaWork, observed, maximum)
        },
        Some(DecodeLimit::Text { observed, maximum }) => {
            (LimitKind::OwnedValueBytes, observed, maximum)
        },
        Some(DecodeLimit::Allocation { requested }) => {
            return Error::Allocation {
                kind: LimitKind::RetainedBytes,
                amount: requested,
            };
        },
        None => {
            return Error::UnsupportedDependency {
                path,
                kind: crate::package::table_cells::DependencyKind::Formula,
            };
        },
    };
    Error::LimitExceeded {
        kind,
        observed: as_u64(observed),
        maximum: as_u64(maximum),
        path,
    }
}

const fn as_u64(value: usize) -> u64 {
    value as u64
}

fn as_usize(value: u64, path: Path) -> Result<usize, Error> {
    usize::try_from(value).map_err(|_| Error::InvalidSource { path })
}

fn checked_add(left: usize, right: usize, path: Path) -> Result<usize, Error> {
    left.checked_add(right).ok_or(Error::InvalidSource { path })
}

fn sort_work(elements: usize, path: Path) -> Result<usize, Error> {
    if elements < 2 {
        return Ok(elements);
    }
    let log = usize::try_from(usize::BITS - (elements - 1).leading_zeros())
        .map_err(|_| Error::InvalidSource { path })?;
    elements
        .checked_mul(log)
        .and_then(|work| work.checked_add(elements))
        .ok_or(Error::InvalidSource { path })
}

fn owner_registry_work(owners: usize) -> Option<usize> {
    sort_work(owners, Path::Package).ok()?.checked_add(owners)
}

fn exact_capacity<T>(buffer: &Vec<T>, expected: usize, path: Path) -> Result<(), Error> {
    if buffer.capacity() != expected {
        return Err(Error::Allocation {
            kind: LimitKind::RetainedElements,
            amount: buffer.capacity().max(expected),
        });
    }
    let _ = path;
    Ok(())
}

#[cfg(test)]
mod tests {
    use crate::{cell::FiniteF64, formula::CachedValue};

    use super::{cache_values_match, tile};

    #[test]
    fn direct_formula_cache_comparison_is_typed_and_bit_exact() {
        let negative_zero = CachedValue::number(-0.0).expect("finite");
        let positive_zero = FiniteF64::new(0.0).expect("finite");
        assert!(!cache_values_match(
            negative_zero.kind(),
            Some(&tile::FormulaCacheValue::Number(positive_zero)),
        ));

        let date = CachedValue::date(42.0).expect("finite");
        let value = FiniteF64::new(42.0).expect("finite");
        assert!(cache_values_match(
            date.kind(),
            Some(&tile::FormulaCacheValue::Date(value)),
        ));
        assert!(!cache_values_match(
            date.kind(),
            Some(&tile::FormulaCacheValue::Duration(value)),
        ));
    }

    #[test]
    fn text_cache_never_matches_an_unresolved_native_key() {
        let text = CachedValue::text("exact").expect("bounded");
        assert!(!cache_values_match(
            text.kind(),
            Some(&tile::FormulaCacheValue::TextKey(7)),
        ));
    }
}
