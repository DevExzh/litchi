//! Coalesced Numbers cell writes for batches and closure edits (record
//! 0297).
//!
//! Consecutive direct `value`/`formula` assignments to distinct cells share
//! one leaf cell transaction per table: the run plans each assignment
//! against the snapshot it started from (cell writes change no path) and
//! stages the changes per table when it is flushed. A repeated coordinate,
//! a query, any other key or verb closes the run, so every operation still
//! observes the result of the ones before it.
//!
//! A run publishes what its assignments publish one at a time (record
//! 0314): a formula is published on its own (its new calculation objects
//! take identities derived from the package it is written to, and the
//! leaf's formula budget is checked against that package, so a change
//! written before it in the same transaction would change both), and a run
//! whose commit fails finds the assignment that fails first, so the error
//! names the item a one-at-a-time application stops at.

use std::collections::{BTreeMap, HashSet};
use std::hash::{BuildHasherDefault, Hasher};

use litchi_core::selector::{CellRef, Path, Step};

use litchi_numbers::table::cells::Change;

use super::super::write::{Chain, Commit, finish, node_at};
use super::super::{Iwork, Owner};
use super::{Address, Numbers, cell_address, cell_write, commit_cells, empty_cell, table_extent};
use crate::error::{Code, Error, Result};
use crate::path::IntoPath;
use crate::props::Props;
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};

/// One planned cell change.
#[derive(Debug, Clone)]
pub(crate) struct Plan {
    address: Address,
    change: Change,
    /// Whether the change can alter the cell (a clear of an empty cell
    /// cannot).
    changes: bool,
    /// Whether the change writes a formula.
    formula: bool,
    /// Whether the leaf stores the cell already.
    stored: bool,
}

/// A table a run's cell paths name: its leaf address and grid extent.
#[derive(Debug, Clone, Copy)]
pub(crate) struct TableInfo {
    sheet: usize,
    table: usize,
    rows: u32,
    columns: u32,
}

/// Table paths a run resolved already, by their steps (a run names few
/// tables, so a linear scan beats hashing the path).
pub(crate) type Tables = Vec<(Vec<Step<'static>>, Option<TableInfo>)>;

impl Iwork {
    /// Plan a direct single-cell `value` or `formula` assignment for a
    /// coalesced run; `None` when the call takes the ordinary path.
    ///
    /// A path whose last step is a bare A1 reference resolves its table
    /// once per run (`tables`) and reads whether the cell is stored from
    /// the leaf projection, instead of walking the tree for every cell.
    pub(crate) fn numbers_cell_plan(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        tables: Option<&mut Tables>,
    ) -> Result<Option<Plan>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        let mut entries = props.iter();
        let (Some((key, value)), None) = (entries.next(), entries.next()) else {
            return Ok(None);
        };
        let located = match tables {
            Some(tables) => self.cached_cell(path.path(), tables),
            None => None,
        };
        let (address, stored) = match located {
            Some(located) => located,
            None => match node_at::<Numbers>(self, path) {
                Ok(node) if node.item().type_name() == "cell" => (cell_address(&node), true),
                Ok(_) => return Ok(None),
                Err(error) if error.code() == Code::PathNotFound => {
                    match empty_cell(self, path.path()) {
                        Some(address) => (address, false),
                        None => return Ok(None),
                    }
                },
                // The ordinary path reports the same error.
                Err(_) => return Ok(None),
            },
        };
        let Some(write) = (match cell_write(key.as_str(), value) {
            Ok(write) => write,
            Err(_) => return Ok(None),
        }) else {
            return Ok(None);
        };
        let changes = stored || !matches!(write, super::CellWrite::Value(None));
        Ok(Some(Plan {
            address,
            change: write.change(address.position()),
            changes,
            formula: matches!(
                write,
                super::CellWrite::Formula(_) | super::CellWrite::FormulaCached { .. }
            ),
            stored,
        }))
    }

    /// The address of the bare A1 cell `path` names and whether the leaf
    /// stores it, through the run's table cache; `None` sends the caller to
    /// the tree walk.
    fn cached_cell(&self, path: &Path<'_>, tables: &mut Tables) -> Option<(Address, bool)> {
        let (last, parent) = path.steps().split_last()?;
        if !last.predicates().is_empty() || parent.is_empty() {
            return None;
        }
        let cell = CellRef::parse(last.name())?;
        let cached = tables
            .iter()
            .find(|(steps, _)| steps.as_slice() == parent)
            .map(|(_, info)| *info);
        let info = match cached {
            Some(info) => info,
            None => {
                let steps: Vec<Step<'static>> =
                    parent.iter().cloned().map(Step::into_owned).collect();
                let info = Path::from_steps(steps.iter().cloned())
                    .into_path()
                    .ok()
                    .and_then(|parent| {
                        let table = node_at::<Numbers>(self, &parent).ok()?;
                        if table.item().type_name() != "table" {
                            return None;
                        }
                        let [sheet, table_index, _, _] = table.item().leaf();
                        let (rows, columns) = table_extent(&table)?;
                        Some(TableInfo {
                            sheet: sheet as usize,
                            table: table_index as usize,
                            rows,
                            columns,
                        })
                    });
                tables.push((steps, info));
                info
            },
        }?;
        if cell.row() >= info.rows || cell.column() >= info.columns {
            return None;
        }
        let address = Address {
            sheet: info.sheet,
            table: info.table,
            row: cell.row(),
            column: cell.column(),
        };
        #[allow(
            irrefutable_let_patterns,
            reason = "a Numbers-only build has one owner variant"
        )]
        let Owner::Numbers(package) = self.owner() else {
            return None;
        };
        let stored = package
            .workbook_ref()
            .sheets()
            .get(info.sheet)?
            .tables()
            .nth(info.table)?
            .get(address.position())
            .is_some();
        Some((address, stored))
    }
}

/// A run of coalesced cell writes over one base snapshot.
pub(crate) struct Run {
    base: Iwork,
    /// The staged changes in staging order: table, change, item index.
    staged: Vec<((usize, usize), Change, usize)>,
    /// Whether the run holds a formula.
    formula: bool,
    /// Per table, the rows of the unstored cells the run writes, and the
    /// last of them: the leaf creates the row headers of new cells' rows in
    /// ascending order, one at a time in staging order.
    created: Vec<((usize, usize), u32, Vec<u32>)>,
    /// The cells planned so far, by [`cell_key`].
    planned: HashSet<(usize, usize, u64), BuildHasherDefault<CellHasher>>,
    tables: Tables,
    first: Option<usize>,
    operations: usize,
    limits: SemanticLimits,
}

impl Run {
    /// An empty run over `base`.
    pub(crate) fn new(base: &Iwork, limits: &SemanticLimits) -> Result<Self> {
        base.admit_edit()?;
        Ok(Self {
            base: base.clone(),
            staged: Vec::new(),
            formula: false,
            created: Vec::new(),
            planned: HashSet::default(),
            tables: Tables::new(),
            first: None,
            operations: 0,
            limits: *limits,
        })
    }

    /// Plan against the run's base snapshot.
    pub(crate) fn cell_plan(
        &mut self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
    ) -> Result<Option<Plan>> {
        self.base
            .numbers_cell_plan(target, props, Some(&mut self.tables))
    }

    /// Whether `plan` must start a new run: its cell was planned already,
    /// it or the run holds a formula, or it creates a cell in a row before
    /// the rows of the cells the run creates.
    pub(crate) fn requires_barrier(&self, plan: &Plan) -> bool {
        let table = (plan.address.sheet, plan.address.table);
        let row = plan.address.row;
        self.formula
            || (plan.formula && !self.staged.is_empty())
            || (!plan.stored
                && self.created.iter().any(|(held, last, rows)| {
                    *held == table && *last > row && rows.binary_search(&row).is_err()
                }))
            || self.planned.contains(&cell_key(plan.address))
    }

    /// The wire index of the run's first operation.
    pub(crate) const fn first_index(&self) -> Option<usize> {
        self.first
    }

    /// Stage `plan` as operation `index`; returns whether it can change a
    /// cell.
    pub(crate) fn stage(&mut self, plan: Plan, index: usize) -> Result<bool> {
        let operations = self.operations.saturating_add(1);
        if operations > self.limits.max_operations() {
            return Err(Error::from(TransactionError::Limit {
                kind: LimitKind::Operations,
                observed: operations,
                maximum: self.limits.max_operations(),
            })
            .with_format(litchi_core::FileFormat::Numbers));
        }
        self.operations = operations;
        self.first.get_or_insert(index);
        self.planned.insert(cell_key(plan.address));
        self.formula |= plan.formula;
        if !plan.stored {
            let table = (plan.address.sheet, plan.address.table);
            let row = plan.address.row;
            match self.created.iter_mut().find(|(held, _, _)| *held == table) {
                // Rows arrive ascending (an earlier new row is a barrier),
                // so the list stays sorted.
                Some((_, last, rows)) => {
                    if let Err(at) = rows.binary_search(&row) {
                        rows.insert(at, row);
                    }
                    *last = (*last).max(row);
                },
                None => self.created.push((table, row, vec![row])),
            }
        }
        self.staged
            .push(((plan.address.sheet, plan.address.table), plan.change, index));
        Ok(plan.changes)
    }

    /// Commit the run: one leaf cell transaction per table.
    ///
    /// # Errors
    ///
    /// The refusal of the first change that a one-at-a-time application
    /// refuses, with its item index (found by committing ever shorter
    /// prefixes of the run, on the error path only).
    pub(crate) fn commit(self) -> Result<(Iwork, Commit)> {
        let chain = match self.publish(self.staged.len()) {
            Ok(chain) => chain,
            Err(error) => return Err(self.first_failure(error)),
        };
        let commit = finish(
            &self.base,
            chain.changed,
            Owner::Numbers(chain.package),
            chain.patches,
            self.operations,
            &self.limits,
        )?;
        Ok((self.base, commit))
    }

    /// Commit the first `count` staged changes onto the base.
    fn publish(&self, count: usize) -> Result<Chain<litchi_numbers::Package>> {
        #[allow(
            irrefutable_let_patterns,
            reason = "a Numbers-only build has one owner variant"
        )]
        let Owner::Numbers(package) = self.base.owner() else {
            return Err(Error::new(
                Code::InternalError,
                "Numbers cell run lost its owner",
            ));
        };
        let mut cells: BTreeMap<(usize, usize), Vec<Change>> = BTreeMap::new();
        for (table, change, _) in self.staged.iter().take(count) {
            cells.entry(*table).or_default().push(change.clone());
        }
        let mut chain = Chain::new(package.clone());
        commit_cells(&mut chain, cells)?;
        Ok(chain)
    }

    /// The error of the shortest failing prefix of the run, with the item
    /// index of its last change; `error` (the whole run's) when no shorter
    /// prefix fails.
    fn first_failure(&self, error: Error) -> Error {
        let (mut low, mut high) = (1, self.staged.len());
        let mut found = None;
        while low < high {
            let middle = low + (high - low) / 2;
            match self.publish(middle) {
                Ok(_) => low = middle + 1,
                Err(error) => {
                    found = Some(error);
                    high = middle;
                },
            }
        }
        let error = found.unwrap_or(error);
        match self.staged.get(high.saturating_sub(1)) {
            Some((_, _, index)) if error.item().is_none() => error.with_item(*index),
            _ => error,
        }
    }
}

/// A planned cell's set key: sheet, table, and row and column in one word.
const fn cell_key(address: Address) -> (usize, usize, u64) {
    (
        address.sheet,
        address.table,
        ((address.row as u64) << 32) | address.column as u64,
    )
}

/// A multiply-rotate hasher for the integer keys of [`cell_key`] (record
/// 0302): the default SipHash cost about a third of a coalesced cell's
/// planning. The keys are coordinates inside the run's tables, so no
/// collision resistance against chosen keys is needed beyond what an
/// operation limit already bounds.
#[derive(Default)]
pub(crate) struct CellHasher(u64);

impl CellHasher {
    const SEED: u64 = 0x51_7c_c1_b7_27_22_0a_95;

    #[inline]
    fn add(&mut self, word: u64) {
        self.0 = (self.0.rotate_left(5) ^ word).wrapping_mul(Self::SEED);
    }
}

impl Hasher for CellHasher {
    fn write(&mut self, bytes: &[u8]) {
        for &byte in bytes {
            self.add(u64::from(byte));
        }
    }

    fn write_u64(&mut self, word: u64) {
        self.add(word);
    }

    fn write_usize(&mut self, word: usize) {
        self.add(word as u64);
    }

    fn finish(&self) -> u64 {
        // The product's high bits depend on every input bit; the table
        // indexes buckets by the low ones.
        self.0.rotate_left(26)
    }
}
