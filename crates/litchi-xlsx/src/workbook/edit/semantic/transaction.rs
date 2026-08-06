//! Public semantic transaction and worksheet-editing facade.

use std::collections::{BTreeMap, BTreeSet, HashMap, btree_map::Entry};
use std::sync::Arc;

use litchi_ooxml_common::web as common_web;
use litchi_opc::Part;
use litchi_sheet::{
    Area, At, Cell as Address, Column as ColumnIndex, ColumnAt, Row as RowIndex, RowAt,
};

use super::super::super::{Selector, Visibility, Workbook, Worksheet, WorksheetKind};
use crate::Style;
use crate::cell::{Cell, Content};
use crate::column::{OutlineAt, State as ColumnState, WidthAt};
use crate::error::{EditBlock, Error, RemoveBlock, Result, TabEditBlock, allocation, invalid};
use crate::layout;
use crate::raw;
use crate::raw::worksheet::edit::{
    Action, ColumnAction, DefaultsAction, DescentEffect, HeightEffect, MergePlan, OptionalEffect,
    Payload, Plan, RowAction, StyleEffect, WidthEffect,
};
use crate::row::{HeightAt, State as RowState};
use crate::sheet::Name;
use crate::style::StyleLineage;
use crate::web::{Binding as WebBinding, Bindings as WebBindings};

use super::super::model::{
    PartChange, StyleGuard, defaults_after, ensure_merge_area, merge_conflicts, project_merges,
};
use super::super::validation::{
    Added, FinalOrder, MergeIntent, MoveIntent, OrderPlan, PanesAction, Placement, SheetActions,
    TabAction, Target, pending_merge,
};
use super::super::{
    ActiveTab, Change, Commit, Conflict, ConflictSet, JoinError, JoinFailure, PackageChange, Patch,
    State,
};
use super::super::{codec, package};

use super::worksheet::{NewSheet, TabEdit, WorksheetEdit};

/// Isolated workbook transaction. Dropping it rolls back every pending change.
#[derive(Debug)]
pub struct Edit {
    pub(in crate::workbook::edit) base: Workbook,
    pub(in crate::workbook::edit) panes: Option<PanesAction>,
    pub(in crate::workbook::edit) active: Option<Target>,
    pub(in crate::workbook::edit) order: Option<OrderPlan>,
    pub(in crate::workbook::edit) sheets: BTreeMap<usize, SheetActions>,
    pub(in crate::workbook::edit) added: Vec<Added>,
    pub(in crate::workbook::edit) removed: BTreeSet<usize>,
}

impl Edit {
    pub(crate) fn new(base: Workbook) -> Result<Self> {
        codec::ensure_unsigned(&base)?;
        Ok(Self {
            base,
            panes: None,
            active: None,
            order: None,
            sheets: BTreeMap::new(),
            added: Vec::new(),
            removed: BTreeSet::new(),
        })
    }

    /// Create or replace the complete persisted Office Add-in task-pane graph.
    ///
    /// The owned model is moved into this isolated transaction. Its worksheet
    /// `appRef` dependencies are checked against the transaction's combined
    /// final state before any package part is changed.
    pub fn put_task_panes(
        &mut self,
        panes: common_web::Panes,
        conformance: common_web::Conformance,
    ) -> Result<&mut Self> {
        if !self.removed.is_empty() {
            return Err(self.remove_block(RemoveBlock::MixedEdit, "task panes"));
        }
        self.panes = Some(PanesAction::Put { panes, conformance });
        Ok(self)
    }

    /// Remove persisted task panes when no effective worksheet binding would
    /// dangle in the same transaction.
    pub fn remove_task_panes(&mut self) -> Result<&mut Self> {
        if !self.removed.is_empty() {
            return Err(self.remove_block(RemoveBlock::MixedEdit, "task panes"));
        }
        self.panes = Some(PanesAction::Remove);
        Ok(self)
    }

    /// Append a validated worksheet and borrow its transaction-local editor.
    ///
    /// The returned handle can populate cells and properties before the one
    /// atomic commit. Native sheet IDs, relationship IDs, and part names are
    /// allocated deterministically at commit and never enter the public API.
    pub fn add<T>(&mut self, name: T) -> Result<NewSheet<'_>>
    where
        T: TryInto<Name>,
        Error: From<T::Error>,
    {
        self.add_placed(name, Placement::Tail)
    }

    /// Insert a validated worksheet immediately before a semantic anchor.
    ///
    /// The anchor accepts the ordinary name or checked zero-based selector.
    /// `Ok(None)` means it did not resolve in the source snapshot.
    pub fn add_before<'e, 's, T>(
        &'e mut self,
        name: T,
        anchor: impl Into<Selector<'s>>,
    ) -> Result<Option<NewSheet<'e>>>
    where
        T: TryInto<Name>,
        Error: From<T::Error>,
    {
        if !self.removed.is_empty() {
            return Err(self.remove_block(RemoveBlock::MixedEdit, "transaction"));
        }
        let Some(anchor) = self.base.sheet(anchor)? else {
            return Ok(None);
        };
        self.add_placed(name, Placement::Before(anchor.position()))
            .map(Some)
    }

    /// Insert a validated worksheet immediately after a semantic anchor.
    ///
    /// Multiple additions at one anchor retain call order. `Ok(None)` means
    /// the name or checked zero-based anchor did not resolve.
    pub fn add_after<'e, 's, T>(
        &'e mut self,
        name: T,
        anchor: impl Into<Selector<'s>>,
    ) -> Result<Option<NewSheet<'e>>>
    where
        T: TryInto<Name>,
        Error: From<T::Error>,
    {
        if !self.removed.is_empty() {
            return Err(self.remove_block(RemoveBlock::MixedEdit, "transaction"));
        }
        let Some(anchor) = self.base.sheet(anchor)? else {
            return Ok(None);
        };
        self.add_placed(name, Placement::After(anchor.position()))
            .map(Some)
    }

    /// Select a worksheet for short transaction-scoped operations.
    pub fn sheet<'e, 's>(
        &'e mut self,
        selector: impl Into<Selector<'s>>,
    ) -> Result<Option<WorksheetEdit<'e>>> {
        if !self.removed.is_empty() {
            return Err(self.remove_block(RemoveBlock::MixedEdit, "transaction"));
        }
        let sheet = self.base.sheet(selector)?;
        let Some(sheet) = sheet else {
            return Ok(None);
        };
        if sheet.kind() != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: sheet.name().to_owned(),
            });
        }
        let position = sheet.position();
        Ok(Some(WorksheetEdit {
            edit: self,
            position,
        }))
    }

    /// Select any workbook sheet tab by its developer-facing name or checked
    /// zero-based position. Unlike [`Self::sheet`], this entry point also
    /// accepts chart, dialog, and macro sheets because visibility belongs to
    /// the workbook catalog rather than worksheet cell storage.
    pub fn tab<'e, 's>(
        &'e mut self,
        selector: impl Into<Selector<'s>>,
    ) -> Result<Option<TabEdit<'e>>> {
        if !self.removed.is_empty() {
            return Err(self.remove_block(RemoveBlock::MixedEdit, "transaction"));
        }
        let tab = self.base.sheet(selector)?;
        Ok(tab.map(|tab| TabEdit {
            edit: self,
            position: tab.position(),
        }))
    }

    /// Move one tab immediately before another by semantic selector.
    ///
    /// `Ok(None)` means either selector did not resolve in the source
    /// snapshot. Names are the ordinary entry point; checked numeric selectors
    /// remain available for import and positional workflows.
    pub fn move_before<'a, 'b>(
        &mut self,
        sheet: impl Into<Selector<'a>>,
        anchor: impl Into<Selector<'b>>,
    ) -> Result<Option<&mut Self>> {
        if !self.removed.is_empty() {
            return Err(self.remove_block(RemoveBlock::MixedEdit, "transaction"));
        }
        let Some(sheet) = self.base.sheet(sheet)? else {
            return Ok(None);
        };
        let Some(anchor) = self.base.sheet(anchor)? else {
            return Ok(None);
        };
        self.move_relative(sheet.position(), anchor.position(), false)?;
        Ok(Some(self))
    }

    /// Move one tab immediately after another by semantic selector.
    ///
    /// `Ok(None)` means either selector did not resolve in the source snapshot.
    pub fn move_after<'a, 'b>(
        &mut self,
        sheet: impl Into<Selector<'a>>,
        anchor: impl Into<Selector<'b>>,
    ) -> Result<Option<&mut Self>> {
        if !self.removed.is_empty() {
            return Err(self.remove_block(RemoveBlock::MixedEdit, "transaction"));
        }
        let Some(sheet) = self.base.sheet(sheet)? else {
            return Ok(None);
        };
        let Some(anchor) = self.base.sheet(anchor)? else {
            return Ok(None);
        };
        self.move_relative(sheet.position(), anchor.position(), true)?;
        Ok(Some(self))
    }

    /// Move a selected tab to a checked zero-based final position.
    ///
    /// `Ok(None)` means the source selector or destination position does not
    /// exist. Prefer [`Self::move_before`] and [`Self::move_after`] when a
    /// stable semantic anchor is available.
    pub fn move_to<'a>(
        &mut self,
        sheet: impl Into<Selector<'a>>,
        position: usize,
    ) -> Result<Option<&mut Self>> {
        if !self.removed.is_empty() {
            return Err(self.remove_block(RemoveBlock::MixedEdit, "transaction"));
        }
        let Some(sheet) = self.base.sheet(sheet)? else {
            return Ok(None);
        };
        if position >= self.base.len() {
            return Ok(None);
        }
        self.move_position(sheet.position(), position)?;
        Ok(Some(self))
    }

    /// Remove a worksheet selected by its developer-facing name or checked
    /// zero-based source position.
    ///
    /// `Ok(None)` means the selector did not resolve. The safe default refuses
    /// live formulas, unmodeled producer references, VBA projects, additional
    /// incoming relationships, and mixed mutation plans. Multiple independent
    /// worksheet removals may be collected in one atomic transaction.
    pub fn remove<'a>(&mut self, selector: impl Into<Selector<'a>>) -> Result<Option<&mut Self>> {
        let Some(sheet) = self.base.sheet(selector)? else {
            return Ok(None);
        };
        if sheet.kind() != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: sheet.name().to_owned(),
            });
        }
        let position = sheet.position();
        let name = sheet.name().to_owned();
        if self.has_non_removal() {
            return Err(Error::SheetRemoveBlocked {
                sheet: name,
                position,
                part: "transaction".to_owned(),
                reason: RemoveBlock::MixedEdit,
            });
        }
        self.removed.insert(position);
        Ok(Some(self))
    }

    pub fn len(&self) -> usize {
        let existing = self.sheets.values().fold(
            self.removed
                .len()
                .saturating_add(usize::from(self.panes.is_some()))
                .saturating_add(usize::from(self.active.is_some()))
                .saturating_add(
                    self.order
                        .as_ref()
                        .filter(|order| order.is_effective())
                        .map_or(0, |order| order.moves.len()),
                ),
            |len, actions| len.saturating_add(actions.len()),
        );
        self.added.iter().fold(existing, |len, added| {
            len.saturating_add(1)
                .saturating_add(usize::from(added.actions.defaults.is_some()))
                .saturating_add(usize::from(added.actions.web.is_some()))
                .saturating_add(added.actions.cells.len())
                .saturating_add(added.actions.rows.len())
                .saturating_add(added.actions.columns.len())
                .saturating_add(added.actions.merges.len())
        })
    }

    pub fn is_empty(&self) -> bool {
        self.active.is_none()
            && self.panes.is_none()
            && self
                .order
                .as_ref()
                .is_none_or(|order| !order.is_effective())
            && self.sheets.values().all(SheetActions::is_empty)
            && self.added.is_empty()
            && self.removed.is_empty()
    }

    /// Join an independently prepared edit when every effect is disjoint.
    ///
    /// Both edits must originate from the same immutable snapshot. On failure
    /// `self` is unchanged and [`JoinError`] returns ownership of `other`, so
    /// callers never lose prepared work while resolving a conflict.
    pub fn join(&mut self, other: Self) -> std::result::Result<&mut Self, JoinError> {
        if !Arc::ptr_eq(&self.base.inner, &other.base.inner) {
            return Err(JoinError {
                failure: JoinFailure::DifferentSnapshot,
                rejected: Box::new(other),
            });
        }
        if self.panes.is_some() && other.panes.is_some() {
            return Err(JoinError {
                failure: JoinFailure::TaskPanes,
                rejected: Box::new(other),
            });
        }
        let conflicts = self.conflicts_with(&other);
        if !conflicts.is_empty() {
            return Err(JoinError {
                failure: JoinFailure::Overlap(conflicts),
                rejected: Box::new(other),
            });
        }

        let added_offset = self.added.len();
        if self.panes.is_none() {
            self.panes = other.panes;
        }
        if self.active.is_none() {
            self.active = other.active.map(|target| match target {
                Target::Base(position) => Target::Base(position),
                Target::Added(index) => Target::Added(added_offset.saturating_add(index)),
            });
        }
        if self
            .order
            .as_ref()
            .is_none_or(|order| !order.is_effective())
        {
            self.order = other.order;
        }
        for (position, actions) in other.sheets {
            match self.sheets.entry(position) {
                Entry::Vacant(entry) => {
                    entry.insert(actions);
                },
                Entry::Occupied(entry) => {
                    let accepted = entry.into_mut();
                    if accepted.rename.is_none() {
                        accepted.rename = actions.rename;
                    }
                    if accepted.visibility.is_none() {
                        accepted.visibility = actions.visibility;
                    }
                    if accepted.web.is_none() {
                        accepted.web = actions.web;
                    }
                    match (accepted.defaults.as_mut(), actions.defaults) {
                        (None, defaults) => accepted.defaults = defaults,
                        (Some(accepted), Some(defaults)) => accepted.merge(defaults),
                        (Some(_), None) => {},
                    }
                    for (address, action) in actions.cells {
                        match accepted.cells.entry(address) {
                            Entry::Vacant(entry) => {
                                entry.insert(action);
                            },
                            Entry::Occupied(mut entry) => {
                                entry.get_mut().merge(action);
                            },
                        }
                    }
                    for (row, action) in actions.rows {
                        match accepted.rows.entry(row) {
                            Entry::Vacant(entry) => {
                                entry.insert(action);
                            },
                            Entry::Occupied(mut entry) => entry.get_mut().merge(action),
                        }
                    }
                    for (column, action) in actions.columns {
                        match accepted.columns.entry(column) {
                            Entry::Vacant(entry) => {
                                entry.insert(action);
                            },
                            Entry::Occupied(mut entry) => entry.get_mut().merge(action),
                        }
                    }
                    accepted.merges.extend(actions.merges);
                },
            }
        }
        self.added.extend(other.added);
        self.removed.extend(other.removed);
        Ok(self)
    }

    /// Validate and atomically publish a new immutable snapshot.
    pub fn commit(self) -> Result<Commit> {
        codec::ensure_unsigned(&self.base)?;
        if self.is_empty() {
            return Ok(Commit {
                workbook: self.base,
                patch: Patch::default(),
            });
        }
        if !self.removed.is_empty() {
            return package::commit_removals(self);
        }
        let Self {
            base,
            panes: requested_panes,
            active: requested_active,
            order: requested_order,
            mut sheets,
            added,
            removed: _,
        } = self;
        let validates_web = requested_panes.is_some()
            || sheets.values().any(|actions| actions.web.is_some())
            || added.iter().any(|sheet| sheet.actions.web.is_some());
        let web_refs = if validates_web {
            let final_panes = match &requested_panes {
                Some(PanesAction::Put { panes, .. }) => Some(panes),
                Some(PanesAction::Remove) => None,
                None => base.task_panes()?,
            };
            Some(match final_panes {
                Some(panes) => crate::web::Refs::from_panes(panes)?,
                None => crate::web::Refs::new(std::iter::empty::<&str>())?,
            })
        } else {
            None
        };
        if let Some(refs) = &web_refs {
            for (position, data) in base.inner.sheets.iter().enumerate() {
                if data.kind != WorksheetKind::Worksheet {
                    continue;
                }
                if let Some(bindings) = sheets
                    .get(&position)
                    .and_then(|actions| actions.web.as_ref())
                {
                    codec::check_web_bindings(refs, &data.name, bindings)?;
                } else {
                    let sheet = Worksheet {
                        owner: Arc::clone(&base.inner),
                        data: Arc::clone(data),
                    };
                    codec::check_web_bindings(refs, &data.name, sheet.web_bindings()?)?;
                }
            }
            for sheet in &added {
                if let Some(bindings) = &sheet.actions.web {
                    codec::check_web_bindings(refs, sheet.name.as_str(), bindings)?;
                }
            }
        }
        let mut changes = Vec::new();
        let mut package_changes = Vec::new();
        let mut parts = Vec::new();
        let mut needs_recalculation = false;

        let mut effective_renames = Vec::<(usize, Name)>::new();
        for (position, actions) in &mut sheets {
            let Some(name) = actions.rename.take() else {
                continue;
            };
            let data =
                base.inner.sheets.get(*position).ok_or_else(|| {
                    invalid(format!("renamed sheet position {position} disappeared"))
                })?;
            if name.as_str() != data.name {
                effective_renames.push((*position, name));
            }
        }
        if let Some((position, _)) = effective_renames.first() {
            let data = base
                .inner
                .sheets
                .get(*position)
                .ok_or_else(|| invalid("renamed tab disappeared during edit"))?;
            package::ensure_reorder_supported(&base, &data.name, *position)?;
        }
        let rename_by_position = effective_renames
            .iter()
            .map(|(position, name)| (*position, name))
            .collect::<HashMap<_, _>>();
        let final_len = base
            .inner
            .sheets
            .len()
            .checked_add(added.len())
            .ok_or_else(|| invalid("final worksheet count overflow"))?;
        if final_len > raw::catalog_edit::MAX_SHEETS {
            let first = added
                .first()
                .ok_or_else(|| invalid("worksheet limit exceeded without a creation"))?;
            return Err(Error::TabEditBlocked {
                sheet: first.name.as_str().to_owned(),
                position: base.inner.sheets.len(),
                reason: TabEditBlock::SheetLimit,
            });
        }
        let effective_order = requested_order.filter(OrderPlan::is_effective);
        if let Some(order) = &effective_order {
            package::validate_order_plan(order, base.inner.sheets.len())?;
        }
        let final_order =
            FinalOrder::plan(base.inner.sheets.len(), effective_order.as_ref(), &added)?;
        if let Some(first) = added.first() {
            let position = final_order
                .position(Target::Added(0))
                .ok_or_else(|| invalid("first created worksheet has no final position"))?;
            package::ensure_reorder_supported(&base, first.name.as_str(), position)?;
        }
        let mut final_names = HashMap::<&str, usize>::new();
        final_names
            .try_reserve(final_len)
            .map_err(|source| allocation("final sheet-name index", source))?;
        for (position, target) in final_order.targets.iter().copied().enumerate() {
            let (name, key) = match target {
                Target::Base(identity) => {
                    let data = base.inner.sheets.get(identity).ok_or_else(|| {
                        invalid("existing worksheet disappeared from final names")
                    })?;
                    rename_by_position
                        .get(&identity)
                        .map_or((data.name.as_str(), data.name_key.as_ref()), |name| {
                            (name.as_str(), name.identity_key())
                        })
                },
                Target::Added(index) => {
                    let created = added
                        .get(index)
                        .ok_or_else(|| invalid("created worksheet disappeared from final names"))?;
                    (created.name.as_str(), created.name.identity_key())
                },
            };
            if let Some(first) = final_names.insert(key, position) {
                return Err(Error::SheetNameConflict {
                    name: name.to_owned(),
                    first,
                    second: position,
                });
            }
        }
        for (position, after) in &effective_renames {
            let before = base
                .inner
                .sheets
                .get(*position)
                .ok_or_else(|| invalid("renamed tab disappeared during patch creation"))?;
            changes.push(Change::Rename {
                position: *position,
                before: before.name.clone().into_boxed_str(),
                after: after.as_str().into(),
            });
        }

        if let Some(order) = &effective_order {
            let first = order
                .moves
                .first()
                .ok_or_else(|| invalid("effective tab order has no semantic move"))?;
            let data = base
                .inner
                .sheets
                .get(first.sheet)
                .ok_or_else(|| invalid("moved tab disappeared during edit"))?;
            package::ensure_reorder_supported(&base, &data.name, first.from)?;
            for moved in &order.moves {
                let data = base
                    .inner
                    .sheets
                    .get(moved.sheet)
                    .ok_or_else(|| invalid("moved tab disappeared during patch creation"))?;
                changes.push(Change::Move {
                    sheet: data.name.clone().into_boxed_str(),
                    from: moved.from,
                    to: moved.to,
                });
            }
        }

        let mut effective_tabs = Vec::new();
        for (position, requested) in &sheets {
            let Some(action) = requested.visibility else {
                continue;
            };
            let data =
                base.inner.sheets.get(*position).ok_or_else(|| {
                    invalid(format!("edited sheet position {position} disappeared"))
                })?;
            let after = action.visibility();
            if data.visibility == after {
                continue;
            }
            effective_tabs.push((*position, action));
        }

        let final_base_is_visible = |position: usize| {
            sheets
                .get(&position)
                .and_then(|requested| requested.visibility)
                .map_or_else(
                    || {
                        base.inner
                            .sheets
                            .get(position)
                            .is_some_and(|sheet| sheet.visibility == Visibility::Visible)
                    },
                    |action| action == TabAction::Show,
                )
        };
        let added_is_visible = |index: usize| {
            added.get(index).is_some_and(|sheet| {
                sheet
                    .actions
                    .visibility
                    .is_none_or(|action| action == TabAction::Show)
            })
        };
        let any_visible = (0..base.inner.sheets.len()).any(&final_base_is_visible)
            || (0..added.len()).any(added_is_visible);
        if !effective_tabs.is_empty() && !any_visible {
            let (position, _) = effective_tabs
                .iter()
                .find(|(_, action)| *action != TabAction::Show)
                .copied()
                .ok_or_else(|| invalid("tab visibility invariant failed without a hide action"))?;
            let data = base
                .inner
                .sheets
                .get(position)
                .ok_or_else(|| invalid("last visible tab disappeared during edit"))?;
            return Err(Error::TabEditBlocked {
                sheet: data.name.clone(),
                position,
                reason: TabEditBlock::LastVisibleTab,
            });
        }

        let final_position = |target: Target| final_order.position(target);
        let final_is_visible = |target: Target| match target {
            Target::Base(position) => final_base_is_visible(position),
            Target::Added(index) => added_is_visible(index),
        };

        let current_active = base.inner.active_sheet;
        let current_target = current_active.map(Target::Base);
        let final_active = if let Some(target) = requested_active {
            let name = match target {
                Target::Base(identity) => {
                    let data =
                        base.inner.sheets.get(identity).ok_or_else(|| {
                            invalid("requested active tab disappeared during edit")
                        })?;
                    data.name.as_str()
                },
                Target::Added(index) => {
                    let data = added.get(index).ok_or_else(|| {
                        invalid("requested new active tab disappeared during edit")
                    })?;
                    data.name.as_str()
                },
            };
            let position = final_order
                .position(target)
                .ok_or_else(|| invalid("requested active tab has no final position"))?;
            if !final_is_visible(target) {
                return Err(Error::TabEditBlocked {
                    sheet: name.to_owned(),
                    position,
                    reason: TabEditBlock::NotVisible,
                });
            }
            Some(target)
        } else if effective_tabs.is_empty() || current_target.is_some_and(final_is_visible) {
            current_target
        } else {
            let len = final_order.len();
            if len == 0 {
                None
            } else {
                current_target
                    .and_then(final_position)
                    .and_then(|current_position| {
                        (1..=len)
                            .filter_map(|offset| {
                                current_position.checked_add(offset).map(|sum| sum % len)
                            })
                            .filter_map(|position| final_order.target(position))
                            .find(|target| final_is_visible(*target))
                    })
                    .or_else(|| {
                        final_order
                            .targets
                            .iter()
                            .copied()
                            .find(|target| final_is_visible(*target))
                    })
            }
        };
        let final_active_position = final_active.and_then(final_position);
        if let Some(position) = final_active_position
            && position > raw::catalog_edit::MAX_ACTIVE_TAB
        {
            let target =
                final_active.ok_or_else(|| invalid("active position has no sheet identity"))?;
            let name = match target {
                Target::Base(identity) => base
                    .inner
                    .sheets
                    .get(identity)
                    .map(|sheet| sheet.name.as_str()),
                Target::Added(index) => added.get(index).map(|sheet| sheet.name.as_str()),
            }
            .ok_or_else(|| invalid("replacement active tab disappeared during edit"))?;
            return Err(Error::TabEditBlocked {
                sheet: name.to_owned(),
                position,
                reason: TabEditBlock::ActiveTabLimit,
            });
        }

        let active_before = current_active
            .map(|identity| package::active_tab_at(&base, identity, identity, None))
            .transpose()?;
        let active_after = final_active
            .zip(final_active_position)
            .map(|(target, position)| match target {
                Target::Base(identity) => package::active_tab_at(
                    &base,
                    identity,
                    position,
                    rename_by_position.get(&identity).map(|name| name.as_str()),
                ),
                Target::Added(index) => added
                    .get(index)
                    .map(|sheet| ActiveTab {
                        name: sheet.name.as_str().into(),
                        position,
                    })
                    .ok_or_else(|| invalid("new active tab disappeared during patch creation")),
            })
            .transpose()?;
        let active_change = (current_target.zip(current_active)
            != final_active.zip(final_active_position))
        .then_some(final_active_position)
        .flatten();
        if active_change.is_some() {
            let before = active_before
                .ok_or_else(|| invalid("non-empty workbook has no source active tab"))?;
            let after = active_after
                .ok_or_else(|| invalid("non-empty workbook has no final active tab"))?;
            changes.push(Change::Active { before, after });
        }

        for (position, action) in &effective_tabs {
            let data = base
                .inner
                .sheets
                .get(*position)
                .ok_or_else(|| invalid("effective tab disappeared during edit"))?;
            changes.push(Change::Visibility {
                sheet: data.name.clone().into_boxed_str(),
                position: *position,
                before: data.visibility.clone(),
                after: action.visibility(),
            });
        }

        for (position, requested) in sheets {
            let data =
                base.inner.sheets.get(position).cloned().ok_or_else(|| {
                    invalid(format!("edited sheet position {position} disappeared"))
                })?;
            let SheetActions {
                rename: _,
                visibility: _,
                defaults,
                web,
                cells,
                rows,
                columns,
                merges,
            } = requested;
            if defaults.is_none()
                && web.is_none()
                && cells.is_empty()
                && rows.is_empty()
                && columns.is_empty()
                && merges.is_empty()
            {
                continue;
            }
            if data.kind != WorksheetKind::Worksheet {
                return Err(Error::NotWorksheet {
                    sheet: data.name.clone(),
                });
            }
            let sheet = Worksheet {
                owner: Arc::clone(&base.inner),
                data: Arc::clone(&data),
            };
            let store = sheet.store()?;
            let change_start = changes.len();
            let effective_web = match web {
                Some(after) => {
                    let before = sheet.web_bindings()?.clone();
                    if before == after {
                        None
                    } else {
                        changes.push(Change::Web {
                            sheet: data.name.clone().into_boxed_str(),
                            before: before.clone(),
                            after: after.clone(),
                        });
                        Some(after)
                    }
                },
                None => None,
            };
            let merge_projection = project_merges(&data.name, Some(store), merges, &cells)?;
            for (range, change) in &merge_projection.changes {
                changes.push(Change::Merge {
                    sheet: data.name.clone().into_boxed_str(),
                    range: *range,
                    change: *change,
                });
            }
            let mut effective_defaults = None;
            if let Some(action) = defaults {
                let before = store.defaults().cloned();
                let after = defaults_after(before.as_ref(), action).map_err(|reason| {
                    Error::DefaultsEditBlocked {
                        sheet: data.name.clone(),
                        reason,
                    }
                })?;
                if before != after {
                    effective_defaults = Some(action);
                    changes.push(Change::Defaults {
                        sheet: data.name.clone().into_boxed_str(),
                        before,
                        after,
                    });
                }
            }
            let mut effective_cells = BTreeMap::new();
            for (address, action) in cells {
                let before = store.entry(address);
                if before.is_some_and(|stored| matches!(stored.cell, Cell::Unknown(_)))
                    && action.payload().is_some()
                {
                    return Err(Error::EditBlocked {
                        sheet: data.name.clone(),
                        address,
                        reason: EditBlock::UnknownCell,
                    });
                }
                let before_state = State::read(before, &base);
                let after_state = State::after(before, &action, &base);
                if before_state == after_state {
                    continue;
                }
                needs_recalculation |= State::calculation_content(&before_state)
                    != State::calculation_content(&after_state);
                effective_cells.insert(address, action);
                changes.push(Change::Cell {
                    sheet: data.name.clone().into_boxed_str(),
                    address,
                    before: before_state,
                    after: after_state,
                });
            }
            let mut effective_rows = BTreeMap::new();
            for (row, action) in rows {
                let before = store.row_entry(row);
                let before_state = RowState::read(before, &base);
                let after_state = RowState::after(before, action, &base);
                if before_state == after_state {
                    continue;
                }
                effective_rows.insert(row, action);
                changes.push(Change::Row {
                    sheet: data.name.clone().into_boxed_str(),
                    row,
                    before: before_state,
                    after: after_state,
                });
            }
            let mut effective_columns = BTreeMap::new();
            for (column, action) in columns {
                let before = store.column_entry(column);
                let before_state = ColumnState::read(before, &base);
                let after_state = ColumnState::after(before, action, &base);
                if before_state == after_state {
                    continue;
                }
                effective_columns.insert(column, action);
                changes.push(Change::Column {
                    sheet: data.name.clone().into_boxed_str(),
                    column,
                    before: before_state,
                    after: after_state,
                });
            }
            if effective_defaults.is_none()
                && effective_web.is_none()
                && effective_cells.is_empty()
                && effective_rows.is_empty()
                && effective_columns.is_empty()
                && merge_projection.plan.is_empty()
            {
                continue;
            }

            let part = base.inner.package.get_part(&data.part_uri)?;
            let before = part.blob_arc();
            let MergePlan { add, remove } = merge_projection.plan;
            let mut after = if remove.is_empty() {
                None
            } else {
                Some(raw::worksheet::edit::rewrite_merges(
                    &before,
                    &data.name,
                    MergePlan {
                        add: Vec::new(),
                        remove,
                    },
                )?)
            };
            let ordinary = Plan {
                defaults: effective_defaults,
                cells: effective_cells,
                rows: effective_rows,
                columns: effective_columns,
            };
            if !ordinary.is_empty() {
                let input = after.as_deref().unwrap_or(&before);
                after = Some(raw::worksheet::edit::rewrite(input, &data.name, ordinary)?);
            }
            if !add.is_empty() {
                let input = after.as_deref().unwrap_or(&before);
                after = Some(raw::worksheet::edit::rewrite_merges(
                    input,
                    &data.name,
                    MergePlan {
                        add,
                        remove: Vec::new(),
                    },
                )?);
            }
            if let Some(bindings) = &effective_web {
                let input = after.as_deref().unwrap_or(&before);
                after = Some(raw::web::replace(input, bindings)?);
            }
            let after =
                after.ok_or_else(|| invalid("effective worksheet edit produced no bytes"))?;
            let parsed = raw::worksheet::parse(&after, || base.inner.shared_strings())?;
            let parsed_web = raw::web::read(&after)?;
            base.inner.validate_styles(&parsed)?;
            for change in &changes[change_start..] {
                match change {
                    Change::Create { .. }
                    | Change::Remove { .. }
                    | Change::Rename { .. }
                    | Change::Move { .. }
                    | Change::Active { .. }
                    | Change::Visibility { .. } => {},
                    Change::Web {
                        sheet,
                        after: expected,
                        ..
                    } => {
                        if &parsed_web != expected {
                            return Err(invalid(format!(
                                "worksheet web-binding verification failed at {sheet}"
                            )));
                        }
                    },
                    Change::Merge {
                        sheet,
                        range,
                        change,
                        ..
                    } => {
                        if parsed.merge_ranges().contains(range) != change.after() {
                            return Err(invalid(format!(
                                "worksheet merged-range verification failed at {sheet}!{range}"
                            )));
                        }
                    },
                    Change::Defaults { sheet, after, .. } => {
                        if parsed.defaults() != after.as_ref() {
                            return Err(invalid(format!(
                                "worksheet defaults edit verification failed at {sheet}"
                            )));
                        }
                    },
                    Change::Cell {
                        sheet,
                        address,
                        after,
                        ..
                    } => {
                        let actual = State::read(parsed.entry(*address), &base);
                        if actual != *after {
                            return Err(invalid(format!(
                                "worksheet edit verification failed at {sheet}!{address}"
                            )));
                        }
                    },
                    Change::Row {
                        sheet, row, after, ..
                    } => {
                        let actual = RowState::read(parsed.row_entry(*row), &base);
                        if actual != *after {
                            return Err(invalid(format!(
                                "worksheet row edit verification failed at {sheet}!row {}",
                                row.get()
                            )));
                        }
                    },
                    Change::Column {
                        sheet,
                        column,
                        after,
                        ..
                    } => {
                        let actual = ColumnState::read(parsed.column_entry(*column), &base);
                        if actual != *after {
                            return Err(invalid(format!(
                                "worksheet column edit verification failed at {sheet}!column {}",
                                column.get()
                            )));
                        }
                    },
                }
            }
            parts.push(PartChange {
                uri: data.part_uri.clone(),
                before,
                after: Arc::new(after),
            });
        }

        let active_added = match final_active {
            Some(Target::Added(index)) => Some(index),
            Some(Target::Base(_)) | None => None,
        };
        let mut created = codec::create_sheets(
            &base,
            added,
            &final_order.added_positions,
            active_added,
            &mut changes,
            &mut needs_recalculation,
        )?;

        if final_active != current_target {
            if let Some(old_active) = current_active {
                let data =
                    base.inner.sheets.get(old_active).ok_or_else(|| {
                        invalid("previous active sheet disappeared during tab edit")
                    })?;
                if data.kind == WorksheetKind::Unknown {
                    return Err(Error::TabEditBlocked {
                        sheet: data.name.clone(),
                        position: old_active,
                        reason: TabEditBlock::MarkupCompatibility,
                    });
                }
                codec::compose_part(&mut parts, &base, &data.part_uri, |content| {
                    raw::sheet_view_edit::rewrite(
                        content,
                        false,
                        raw::sheet_view_edit::Context {
                            sheet: &data.name,
                            position: old_active,
                        },
                    )
                })?;
            }
            let new_active = final_active
                .ok_or_else(|| invalid("active tab disappeared during selection rewrite"))?;
            if let Target::Base(new_active) = new_active {
                let data = base
                    .inner
                    .sheets
                    .get(new_active)
                    .ok_or_else(|| invalid("new active sheet disappeared during tab edit"))?;
                if data.kind == WorksheetKind::Unknown {
                    return Err(Error::TabEditBlocked {
                        sheet: data.name.clone(),
                        position: new_active,
                        reason: TabEditBlock::MarkupCompatibility,
                    });
                }
                codec::compose_part(&mut parts, &base, &data.part_uri, |content| {
                    raw::sheet_view_edit::rewrite(
                        content,
                        true,
                        raw::sheet_view_edit::Context {
                            sheet: &data.name,
                            position: final_active_position.unwrap_or(new_active),
                        },
                    )
                })?;
            }
        }

        let reference_renames = effective_renames
            .iter()
            .map(|(position, after)| {
                let before = base.inner.sheets.get(*position).ok_or_else(|| {
                    invalid("renamed sheet disappeared during reference planning")
                })?;
                Ok(raw::reference_edit::Rename {
                    before: &before.name,
                    after: after.as_str(),
                })
            })
            .collect::<Result<Vec<_>>>()?;
        let sheet_titles = if effective_renames.is_empty() {
            Vec::new()
        } else {
            base.inner
                .sheets
                .iter()
                .map(|sheet| sheet.name.as_str())
                .collect::<Vec<_>>()
        };
        if let Some((position, _)) = effective_renames.first() {
            let sheet = base
                .inner
                .sheets
                .get(*position)
                .ok_or_else(|| invalid("rename context sheet disappeared"))?;
            let mut reference_parts = base
                .inner
                .package
                .iter_parts()
                .filter(|part| {
                    part.partname() != &base.inner.workbook_uri && codec::reference_part(*part)
                })
                .map(|part| {
                    (
                        part.partname().clone(),
                        part.content_type()
                            == litchi_opc::constants::content_type::OFC_EXTENDED_PROPERTIES,
                    )
                })
                .collect::<Vec<_>>();
            reference_parts.sort_unstable_by(|left, right| left.0.as_str().cmp(right.0.as_str()));
            for (uri, extended_properties) in reference_parts {
                let part_name = uri.to_string();
                let part_titles = if extended_properties {
                    sheet_titles.as_slice()
                } else {
                    &[]
                };
                codec::compose_part_optional(&mut parts, &base, &uri, |content| {
                    raw::reference_edit::rewrite(
                        content,
                        &reference_renames,
                        raw::reference_edit::Context {
                            sheet: &sheet.name,
                            position: *position,
                            part: &part_name,
                            sheet_titles: part_titles,
                        },
                    )
                })?;
            }
            for created_sheet in &mut created {
                let part_name = created_sheet.graph.part.partname().to_string();
                if let Some(after) = raw::reference_edit::rewrite(
                    created_sheet.graph.part.blob(),
                    &reference_renames,
                    raw::reference_edit::Context {
                        sheet: &sheet.name,
                        position: *position,
                        part: &part_name,
                        sheet_titles: &[],
                    },
                )? {
                    created_sheet.graph.part.set_blob(after);
                }
                let parsed = raw::worksheet::parse(created_sheet.graph.part.blob(), || {
                    base.inner.shared_strings()
                })?;
                base.inner.validate_styles(&parsed)?;
                for change in &mut changes {
                    if let Change::Cell {
                        sheet,
                        address,
                        after,
                        ..
                    } = change
                        && sheet.as_ref() == created_sheet.name.as_str()
                    {
                        *after = State::read(parsed.entry(*address), &base);
                    }
                }
            }
            for change in &mut changes {
                let Change::Web { sheet, after, .. } = change else {
                    continue;
                };
                if let Some(data) = base
                    .inner
                    .sheets
                    .iter()
                    .find(|data| data.name.as_str() == sheet.as_ref())
                {
                    let content = parts
                        .iter()
                        .find(|part| part.uri == data.part_uri)
                        .map_or_else(
                            || base.inner.package.get_part(&data.part_uri).map(Part::blob),
                            |part| Ok(part.after.as_slice()),
                        )?;
                    *after = raw::web::read(content)?;
                } else if let Some(created) = created
                    .iter()
                    .find(|created| created.name.as_str() == sheet.as_ref())
                {
                    *after = raw::web::read(created.graph.part.blob())?;
                }
            }
        }
        if effective_order.is_some() || !created.is_empty() {
            let existing_titles = base
                .inner
                .sheets
                .iter()
                .enumerate()
                .map(|(position, sheet)| {
                    rename_by_position
                        .get(&position)
                        .map_or(sheet.name.as_str(), |name| name.as_str())
                })
                .collect::<Vec<_>>();
            let mut property_order = Vec::new();
            property_order
                .try_reserve_exact(final_order.len())
                .map_err(|source| allocation("extended-properties sheet order", source))?;
            for target in final_order.targets.iter().copied() {
                property_order.push(match target {
                    Target::Base(identity) => raw::properties_edit::Sheet::Existing(identity),
                    Target::Added(index) => {
                        let sheet = created.get(index).ok_or_else(|| {
                            invalid("created worksheet disappeared from property order")
                        })?;
                        raw::properties_edit::Sheet::New(sheet.name.as_str())
                    },
                });
            }
            let mut property_parts = base
                .inner
                .package
                .iter_parts()
                .filter(|part| {
                    part.content_type()
                        == litchi_opc::constants::content_type::OFC_EXTENDED_PROPERTIES
                })
                .map(|part| part.partname().clone())
                .collect::<Vec<_>>();
            property_parts.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
            for uri in property_parts {
                codec::compose_part_optional(&mut parts, &base, &uri, |content| {
                    raw::properties_edit::arrange_sheets(content, &existing_titles, &property_order)
                })?;
            }
        }

        let style_guard = changes
            .iter()
            .any(Change::uses_shared_style)
            .then(|| {
                let uri = base
                    .inner
                    .styles_uri
                    .as_ref()
                    .ok_or_else(|| invalid("shared style state has no styles part"))?
                    .clone();
                let content = base.inner.package.get_part(&uri)?.blob_arc();
                Ok::<_, Error>(StyleGuard { uri, content })
            })
            .transpose()?;
        let calculation_graph = if needs_recalculation {
            codec::calculation_chain_removal(&base)?
        } else {
            Vec::new()
        };

        if !effective_renames.is_empty()
            || !effective_tabs.is_empty()
            || active_change.is_some()
            || effective_order.is_some()
            || !created.is_empty()
            || needs_recalculation
        {
            let workbook_part = base.inner.package.get_part(&base.inner.workbook_uri)?;
            let before = workbook_part.blob_arc();
            let referenced_workbook = if let Some((position, _)) = effective_renames.first() {
                let sheet = base
                    .inner
                    .sheets
                    .get(*position)
                    .ok_or_else(|| invalid("rename workbook context disappeared"))?;
                raw::reference_edit::rewrite(
                    &before,
                    &reference_renames,
                    raw::reference_edit::Context {
                        sheet: &sheet.name,
                        position: *position,
                        part: base.inner.workbook_uri.as_str(),
                        sheet_titles: &[],
                    },
                )?
            } else {
                None
            };
            let workbook_input = referenced_workbook.as_deref().unwrap_or(&before);
            let scope_source = (effective_order.is_some() || !created.is_empty())
                .then(|| raw::parse_catalog(workbook_input))
                .transpose()?;
            let raw_order = effective_order
                .as_ref()
                .map(|order| {
                    let moved = order
                        .moves
                        .first()
                        .ok_or_else(|| invalid("effective tab order lost its move context"))?;
                    let data =
                        base.inner.sheets.get(moved.sheet).ok_or_else(|| {
                            invalid("tab reorder context disappeared during commit")
                        })?;
                    let relationship_ids = order
                        .positions
                        .iter()
                        .map(|identity| {
                            base.inner
                                .sheets
                                .get(*identity)
                                .map(|sheet| sheet.relationship_id.as_str())
                                .ok_or_else(|| {
                                    invalid("tab reorder target disappeared during commit")
                                })
                        })
                        .collect::<Result<Vec<_>>>()?;
                    Ok::<_, Error>(raw::catalog_edit::Order {
                        sheet: &data.name,
                        position: moved.from,
                        relationship_ids,
                        local_scopes: base
                            .inner
                            .defined_names
                            .iter()
                            .filter(|name| name.local_sheet_id.is_some())
                            .count(),
                    })
                })
                .transpose()?;
            let raw_renames = effective_renames
                .iter()
                .map(|(position, after)| {
                    let data =
                        base.inner.sheets.get(*position).ok_or_else(|| {
                            invalid("tab rename target disappeared during commit")
                        })?;
                    Ok(raw::catalog_edit::Rename {
                        sheet: &data.name,
                        position: *position,
                        relationship_id: &data.relationship_id,
                        name: after.as_str(),
                    })
                })
                .collect::<Result<Vec<_>>>()?;
            let has_existing_catalog_edit =
                !effective_renames.is_empty() || !effective_tabs.is_empty() || raw_order.is_some();
            let mut after = if has_existing_catalog_edit {
                let tabs = effective_tabs
                    .iter()
                    .map(|(position, action)| {
                        let data = base.inner.sheets.get(*position).ok_or_else(|| {
                            invalid("tab rewrite target disappeared during commit")
                        })?;
                        Ok(raw::catalog_edit::Tab {
                            sheet: &data.name,
                            position: *position,
                            relationship_id: &data.relationship_id,
                            state: action.raw(),
                        })
                    })
                    .collect::<Result<Vec<_>>>()?;
                raw::catalog_edit::rewrite(
                    workbook_input,
                    raw::catalog_edit::Plan {
                        tabs,
                        renames: raw_renames,
                        active: None,
                        order: raw_order,
                    },
                )?
            } else {
                workbook_input.to_vec()
            };
            for (index, sheet) in created.iter().enumerate() {
                let position = base
                    .inner
                    .sheets
                    .len()
                    .checked_add(index)
                    .ok_or_else(|| invalid("physical worksheet append position overflow"))?;
                after = raw::catalog_edit::append(
                    &after,
                    raw::catalog_edit::Create {
                        sheet: sheet.name.as_str(),
                        position,
                        sheet_id: sheet.sheet_id,
                        relationship_id: &sheet.relationship_id,
                        state: sheet.visibility.raw(),
                    },
                )?;
            }
            if !final_order.matches_appended(effective_order.as_ref()) {
                let mut relationship_ids = Vec::new();
                relationship_ids
                    .try_reserve_exact(final_order.len())
                    .map_err(|source| allocation("final catalog order", source))?;
                for target in final_order.targets.iter().copied() {
                    relationship_ids.push(match target {
                        Target::Base(identity) => base
                            .inner
                            .sheets
                            .get(identity)
                            .map(|sheet| sheet.relationship_id.as_str())
                            .ok_or_else(|| {
                                invalid("existing worksheet disappeared from final catalog order")
                            })?,
                        Target::Added(index) => created
                            .get(index)
                            .map(|sheet| sheet.relationship_id.as_str())
                            .ok_or_else(|| {
                                invalid("created worksheet disappeared from final catalog order")
                            })?,
                    });
                }
                let context = final_order
                    .targets
                    .first()
                    .copied()
                    .ok_or_else(|| invalid("final catalog order has no context sheet"))?;
                let context_name = match context {
                    Target::Base(identity) => base
                        .inner
                        .sheets
                        .get(identity)
                        .map(|sheet| sheet.name.as_str()),
                    Target::Added(index) => created.get(index).map(|sheet| sheet.name.as_str()),
                }
                .ok_or_else(|| invalid("final catalog context sheet disappeared"))?;
                after = raw::catalog_edit::rewrite(
                    &after,
                    raw::catalog_edit::Plan {
                        tabs: Vec::new(),
                        renames: Vec::new(),
                        active: None,
                        order: Some(raw::catalog_edit::Order {
                            sheet: context_name,
                            position: 0,
                            relationship_ids,
                            local_scopes: base
                                .inner
                                .defined_names
                                .iter()
                                .filter(|name| name.local_sheet_id.is_some())
                                .count(),
                        }),
                    },
                )?;
            }
            if let Some(position) = active_change {
                let target = final_active
                    .ok_or_else(|| invalid("active-tab rewrite has no final sheet identity"))?;
                let sheet = match target {
                    Target::Base(identity) => base
                        .inner
                        .sheets
                        .get(identity)
                        .map(|sheet| sheet.name.as_str()),
                    Target::Added(index) => created.get(index).map(|sheet| sheet.name.as_str()),
                }
                .ok_or_else(|| invalid("active-tab rewrite target disappeared"))?;
                after = raw::catalog_edit::rewrite(
                    &after,
                    raw::catalog_edit::Plan {
                        tabs: Vec::new(),
                        renames: Vec::new(),
                        active: Some(raw::catalog_edit::Active { sheet, position }),
                        order: None,
                    },
                )?;
            }
            if needs_recalculation {
                after = raw::recalc::invalidate(&after)?;
            }
            if has_existing_catalog_edit || !created.is_empty() || active_change.is_some() {
                let catalog = raw::parse_catalog(&after)?;
                if Some(catalog.active_sheet_index) != final_active_position {
                    return Err(invalid("workbook active-tab edit verification failed"));
                }
                if catalog.sheets.len() != final_order.len() {
                    return Err(invalid(
                        "workbook creation verification has the wrong sheet count",
                    ));
                }
                for (position, target) in final_order.targets.iter().copied().enumerate() {
                    let expected = match target {
                        Target::Base(identity) => base
                            .inner
                            .sheets
                            .get(identity)
                            .map(|sheet| sheet.relationship_id.as_str()),
                        Target::Added(index) => created
                            .get(index)
                            .map(|sheet| sheet.relationship_id.as_str()),
                    }
                    .ok_or_else(|| invalid("final worksheet order lost a sheet identity"))?;
                    let actual = catalog
                        .sheets
                        .get(position)
                        .ok_or_else(|| invalid("final worksheet order lost a catalog slot"))?;
                    if actual.relationship_id != expected {
                        return Err(invalid(format!(
                            "workbook order verification failed at position {position}"
                        )));
                    }
                }
                for (identity, action) in &effective_tabs {
                    let position = final_position(Target::Base(*identity)).ok_or_else(|| {
                        invalid("visible tab disappeared from the final workbook order")
                    })?;
                    let actual = catalog
                        .sheets
                        .get(position)
                        .ok_or_else(|| invalid("workbook tab edit verification lost a sheet"))?;
                    if !package::raw_visibility_matches(&actual.visibility, *action) {
                        let sheet = base
                            .inner
                            .sheets
                            .get(*identity)
                            .map_or("<missing sheet>", |sheet| sheet.name.as_str());
                        return Err(invalid(format!(
                            "workbook tab visibility verification failed at {sheet}"
                        )));
                    }
                }
                for (identity, expected_name) in &effective_renames {
                    let position = final_position(Target::Base(*identity)).ok_or_else(|| {
                        invalid("renamed tab disappeared from the final workbook order")
                    })?;
                    let actual = catalog
                        .sheets
                        .get(position)
                        .ok_or_else(|| invalid("workbook rename verification lost a sheet"))?;
                    let expected = base.inner.sheets.get(*identity).ok_or_else(|| {
                        invalid("renamed tab identity disappeared during verification")
                    })?;
                    if actual.relationship_id != expected.relationship_id
                        || actual.name != expected_name.as_str()
                    {
                        return Err(invalid(format!(
                            "workbook tab rename verification failed at position {position}"
                        )));
                    }
                }
                if effective_order.is_some() || !created.is_empty() {
                    let source = scope_source.as_ref().ok_or_else(|| {
                        invalid("defined-name scope source disappeared during verification")
                    })?;
                    package::verify_defined_name_scopes(
                        source,
                        &catalog,
                        base.inner.sheets.len(),
                        &final_order.targets,
                    )?;
                }
                for sheet in &created {
                    let actual = catalog.sheets.get(sheet.position).ok_or_else(|| {
                        invalid("created worksheet disappeared during catalog verification")
                    })?;
                    if actual.name != sheet.name.as_str()
                        || actual.relationship_id != sheet.relationship_id
                        || actual.sheet_id != sheet.sheet_id
                        || !package::raw_visibility_matches(&actual.visibility, sheet.visibility)
                    {
                        return Err(invalid(format!(
                            "workbook creation verification failed at position {}",
                            sheet.position
                        )));
                    }
                }
            }
            if after.as_slice() != before.as_slice() {
                parts.push(PartChange {
                    uri: base.inner.workbook_uri.clone(),
                    before,
                    after: Arc::new(after),
                });
            }
        }

        let mut graph = Vec::new();
        graph
            .try_reserve(created.len().saturating_add(calculation_graph.len()))
            .map_err(|source| allocation("package graph changes", source))?;
        graph.extend(created.into_iter().map(|sheet| sheet.graph));
        graph.extend(calculation_graph);

        let web_patch = match requested_panes {
            Some(PanesAction::Put { panes, conformance }) => {
                let after = panes.clone();
                let patch = common_web::plan_put(&base.inner.package, panes, conformance)?;
                if patch.is_empty() {
                    None
                } else {
                    package_changes.push(PackageChange::TaskPanes {
                        before: base.task_panes()?.cloned(),
                        after: Some(after),
                    });
                    Some(patch)
                }
            },
            Some(PanesAction::Remove) => {
                let patch = common_web::plan_remove(&base.inner.package)?;
                if patch.is_empty() {
                    None
                } else {
                    package_changes.push(PackageChange::TaskPanes {
                        before: base.task_panes()?.cloned(),
                        after: None,
                    });
                    Some(patch)
                }
            },
            None => None,
        };

        if changes.is_empty()
            && package_changes.is_empty()
            && parts.is_empty()
            && graph.is_empty()
            && web_patch.is_none()
        {
            return Ok(Commit {
                workbook: base,
                patch: Patch::default(),
            });
        }
        let mut package = base.inner.package.clone();
        for part in &parts {
            package
                .get_part_mut(&part.uri)?
                .set_blob_shared(Arc::clone(&part.after));
        }
        for change in &graph {
            change.validate(&package)?;
            change.apply(&mut package)?;
        }
        if let Some(web) = &web_patch {
            web.apply(&mut package)?;
        }
        let workbook = Workbook::from_package_with_styles(package, Some(&base))?;
        if web_patch.is_some()
            || changes
                .iter()
                .any(|change| matches!(change, Change::Web { .. }))
        {
            codec::validate_web_integrity(&workbook)?;
        }
        Ok(Commit {
            workbook,
            patch: Patch {
                changes: changes.into_boxed_slice(),
                package_changes: package_changes.into_boxed_slice(),
                parts: parts.into_boxed_slice(),
                graph: graph.into_boxed_slice(),
                web: web_patch,
                style_guard,
            },
        })
    }

    fn add_placed<T>(&mut self, name: T, placement: Placement) -> Result<NewSheet<'_>>
    where
        T: TryInto<Name>,
        Error: From<T::Error>,
    {
        if !self.removed.is_empty() {
            return Err(self.remove_block(RemoveBlock::MixedEdit, "transaction"));
        }
        let name = name.try_into().map_err(Error::from)?;
        let limit_position = self
            .base
            .len()
            .checked_add(self.added.len())
            .ok_or_else(|| invalid("worksheet position overflow"))?;
        if limit_position >= raw::catalog_edit::MAX_SHEETS {
            return Err(Error::TabEditBlocked {
                sheet: name.as_str().to_owned(),
                position: limit_position,
                reason: TabEditBlock::SheetLimit,
            });
        }
        self.added
            .try_reserve(1)
            .map_err(|source| allocation("worksheet creation", source))?;
        let index = self.added.len();
        self.added.push(Added {
            name,
            actions: SheetActions::default(),
            placement,
        });
        let effective_order = self.order.as_ref().filter(|order| order.is_effective());
        let position = match FinalOrder::plan(self.base.len(), effective_order, &self.added)
            .and_then(|order| {
                order
                    .position(Target::Added(index))
                    .ok_or_else(|| invalid("new worksheet has no projected position"))
            }) {
            Ok(position) => position,
            Err(error) => {
                if self.added.pop().is_none() {
                    return Err(invalid(
                        "new worksheet planning failed and its rollback target disappeared",
                    ));
                }
                return Err(error);
            },
        };
        if self.added.get(index).is_none() {
            let _ = self.added.pop();
            return Err(invalid("worksheet creation plan disappeared"));
        }
        let added = self
            .added
            .get_mut(index)
            .ok_or_else(|| invalid("worksheet creation plan disappeared after validation"))?;
        Ok(NewSheet {
            added,
            active: &mut self.active,
            style_lineage: &self.base.inner.style_lineage,
            index,
            position,
        })
    }

    pub(super) fn actions(&mut self, position: usize) -> &mut BTreeMap<Address, Action> {
        &mut self.sheets.entry(position).or_default().cells
    }

    pub(super) fn web_bindings(&mut self, position: usize) -> Result<&mut WebBindings> {
        let needs_source = self
            .sheets
            .get(&position)
            .is_none_or(|actions| actions.web.is_none());
        if needs_source {
            let data = self
                .base
                .inner
                .sheets
                .get(position)
                .cloned()
                .ok_or_else(|| invalid("web-binding target disappeared"))?;
            if data.kind != WorksheetKind::Worksheet {
                return Err(Error::NotWorksheet {
                    sheet: data.name.clone(),
                });
            }
            let sheet = Worksheet {
                owner: Arc::clone(&self.base.inner),
                data,
            };
            let source = sheet.web_bindings()?.clone();
            self.sheets.entry(position).or_default().web = Some(source);
        }
        self.sheets
            .get_mut(&position)
            .and_then(|actions| actions.web.as_mut())
            .ok_or_else(|| invalid("web-binding edit initialization failed"))
    }

    pub(super) fn set_visibility(&mut self, position: usize, action: TabAction) {
        self.sheets.entry(position).or_default().visibility = Some(action);
    }

    pub(super) fn set_name(&mut self, position: usize, name: Name) {
        self.sheets.entry(position).or_default().rename = Some(name);
    }

    pub(super) fn set_active(&mut self, position: usize) {
        self.active = Some(Target::Base(position));
    }

    fn order_plan(&mut self) -> Result<&mut OrderPlan> {
        if self.order.is_none() {
            let len = self.base.inner.sheets.len();
            let mut positions = Vec::new();
            positions
                .try_reserve_exact(len)
                .map_err(|source| allocation("tab-order plan", source))?;
            positions.extend(0..len);
            self.order = Some(OrderPlan {
                positions,
                moves: Vec::new(),
            });
        }
        self.order
            .as_mut()
            .ok_or_else(|| invalid("tab-order plan initialization failed"))
    }

    fn move_position(&mut self, identity: usize, to: usize) -> Result<()> {
        let order = self.order_plan()?;
        if to >= order.positions.len() {
            return Err(invalid("tab move destination exceeds the workbook order"));
        }
        let from = order
            .positions
            .iter()
            .position(|candidate| *candidate == identity)
            .ok_or_else(|| invalid("selected tab disappeared from the pending order"))?;
        if from == to {
            return Ok(());
        }
        order
            .moves
            .try_reserve(1)
            .map_err(|source| allocation("tab move", source))?;
        let identity = order.positions.remove(from);
        order.positions.insert(to, identity);
        order.moves.push(MoveIntent {
            sheet: identity,
            from,
            to,
        });
        Ok(())
    }

    fn move_relative(&mut self, identity: usize, anchor: usize, after: bool) -> Result<()> {
        if identity == anchor {
            return Ok(());
        }
        let order = self.order_plan()?;
        let from = order
            .positions
            .iter()
            .position(|candidate| *candidate == identity)
            .ok_or_else(|| invalid("selected tab disappeared from the pending order"))?;
        if !order.positions.contains(&anchor) {
            return Err(invalid("anchor tab disappeared from the pending order"));
        }
        order
            .moves
            .try_reserve(1)
            .map_err(|source| allocation("tab move", source))?;
        let identity = order.positions.remove(from);
        let anchor = order
            .positions
            .iter()
            .position(|candidate| *candidate == anchor)
            .ok_or_else(|| invalid("anchor tab disappeared during reorder"))?;
        let to = if after {
            anchor
                .checked_add(1)
                .ok_or_else(|| invalid("tab move position overflow"))?
        } else {
            anchor
        };
        order.positions.insert(to, identity);
        if from != to {
            order.moves.push(MoveIntent {
                sheet: identity,
                from,
                to,
            });
        }
        Ok(())
    }

    fn conflicts_with(&self, other: &Self) -> ConflictSet {
        let mut conflicts = Vec::new();
        let removal_conflict = self
            .removed
            .intersection(&other.removed)
            .next()
            .copied()
            .or_else(|| {
                (!self.removed.is_empty() && other.has_non_removal())
                    .then(|| self.removed.iter().next().copied())
                    .flatten()
            })
            .or_else(|| {
                (!other.removed.is_empty() && self.has_non_removal())
                    .then(|| other.removed.iter().next().copied())
                    .flatten()
            });
        if let Some(position) = removal_conflict {
            let sheet = self
                .base
                .inner
                .sheets
                .get(position)
                .map_or("<missing sheet>", |sheet| sheet.name.as_str());
            conflicts.push(Conflict::Remove {
                sheet: sheet.into(),
                position,
            });
        }
        if let (Some(left), Some(_)) = (
            self.order.as_ref().filter(|order| order.is_effective()),
            other.order.as_ref().filter(|order| order.is_effective()),
        ) {
            let moved = left.moves.first();
            let position = moved.map_or(0, |moved| moved.from);
            let sheet = moved
                .and_then(|moved| self.base.inner.sheets.get(moved.sheet))
                .map_or("<missing sheet>", |sheet| sheet.name.as_str());
            conflicts.push(Conflict::Order {
                sheet: sheet.into(),
                position,
            });
        }
        if let (Some(target), Some(_)) = (self.active, other.active) {
            let (sheet, position) = self.target_context(target);
            conflicts.push(Conflict::Active {
                sheet: sheet.into(),
                position,
            });
        }
        for (left_index, left) in self.added.iter().enumerate() {
            if other
                .added
                .iter()
                .any(|right| right.name.identity_key() == left.name.identity_key())
                || other.sheets.values().any(|actions| {
                    actions
                        .rename
                        .as_ref()
                        .is_some_and(|name| name.identity_key() == left.name.identity_key())
                })
            {
                conflicts.push(Conflict::Name {
                    sheet: left.name.as_str().into(),
                    position: self
                        .projected_position(Target::Added(left_index))
                        .unwrap_or_else(|| self.base.len().saturating_add(left_index)),
                });
            }
        }
        for (right_index, right) in other.added.iter().enumerate() {
            if self.sheets.values().any(|actions| {
                actions
                    .rename
                    .as_ref()
                    .is_some_and(|name| name.identity_key() == right.name.identity_key())
            }) {
                conflicts.push(Conflict::Name {
                    sheet: right.name.as_str().into(),
                    position: other
                        .projected_position(Target::Added(right_index))
                        .unwrap_or_else(|| other.base.len().saturating_add(right_index)),
                });
            }
        }
        for (position, left) in &self.sheets {
            let Some(right) = other.sheets.get(position) else {
                continue;
            };
            let sheet = self
                .base
                .inner
                .sheets
                .get(*position)
                .map_or("<missing sheet>", |sheet| sheet.name.as_str());
            if left.rename.is_some() && right.rename.is_some() {
                conflicts.push(Conflict::Name {
                    sheet: sheet.into(),
                    position: *position,
                });
            }
            if left.visibility.is_some() && right.visibility.is_some() {
                conflicts.push(Conflict::Tab {
                    sheet: sheet.into(),
                    position: *position,
                });
            }
            if left.web.is_some() && right.web.is_some() {
                conflicts.push(Conflict::Web {
                    sheet: sheet.into(),
                    position: *position,
                });
            }
            if let (Some(left), Some(right)) = (left.defaults, right.defaults) {
                let fields = left.fields() & right.fields();
                if left.overlaps(right) {
                    conflicts.push(Conflict::Defaults {
                        sheet: sheet.into(),
                        position: *position,
                        fields,
                    });
                }
            }
            let ranges = merge_conflicts(left, right);
            if !ranges.is_empty() {
                conflicts.push(Conflict::Merges {
                    sheet: sheet.into(),
                    position: *position,
                    ranges: ranges.into_boxed_slice(),
                });
            }
            let mut addresses = Vec::new();
            let mut left_cells = left.cells.iter().peekable();
            let mut right_cells = right.cells.iter().peekable();
            while let (Some((left_address, left_action)), Some((right_address, right_action))) =
                (left_cells.peek(), right_cells.peek())
            {
                match left_address.cmp(right_address) {
                    std::cmp::Ordering::Less => {
                        left_cells.next();
                    },
                    std::cmp::Ordering::Greater => {
                        right_cells.next();
                    },
                    std::cmp::Ordering::Equal => {
                        if left_action.overlaps(right_action) {
                            addresses.push(**left_address);
                        }
                        left_cells.next();
                        right_cells.next();
                    },
                }
            }
            if !addresses.is_empty() {
                conflicts.push(Conflict::Cells {
                    sheet: sheet.into(),
                    position: *position,
                    addresses: addresses.into_boxed_slice(),
                });
            }

            let rows = left
                .rows
                .iter()
                .filter_map(|(row, action)| {
                    right
                        .rows
                        .get(row)
                        .is_some_and(|other| action.overlaps(*other))
                        .then_some(*row)
                })
                .collect::<Vec<_>>();
            if !rows.is_empty() {
                conflicts.push(Conflict::Rows {
                    sheet: sheet.into(),
                    position: *position,
                    rows: rows.into_boxed_slice(),
                });
            }

            let columns = left
                .columns
                .iter()
                .filter_map(|(column, action)| {
                    right
                        .columns
                        .get(column)
                        .is_some_and(|other| action.overlaps(*other))
                        .then_some(*column)
                })
                .collect::<Vec<_>>();
            if !columns.is_empty() {
                conflicts.push(Conflict::Columns {
                    sheet: sheet.into(),
                    position: *position,
                    columns: columns.into_boxed_slice(),
                });
            }
        }
        ConflictSet {
            conflicts: conflicts.into_boxed_slice(),
        }
    }

    fn target_context(&self, target: Target) -> (&str, usize) {
        let projected = self.projected_position(target);
        match target {
            Target::Base(position) => (
                self.base
                    .inner
                    .sheets
                    .get(position)
                    .map_or("<missing sheet>", |sheet| sheet.name.as_str()),
                projected.unwrap_or(position),
            ),
            Target::Added(index) => self.added.get(index).map_or(
                ("<missing new sheet>", self.base.len().saturating_add(index)),
                |sheet| {
                    (
                        sheet.name.as_str(),
                        projected.unwrap_or_else(|| self.base.len().saturating_add(index)),
                    )
                },
            ),
        }
    }

    fn projected_position(&self, target: Target) -> Option<usize> {
        FinalOrder::plan(
            self.base.len(),
            self.order.as_ref().filter(|order| order.is_effective()),
            &self.added,
        )
        .ok()?
        .position(target)
    }

    pub(in crate::workbook::edit) fn has_non_removal(&self) -> bool {
        self.panes.is_some()
            || self.active.is_some()
            || self.order.as_ref().is_some_and(OrderPlan::is_effective)
            || self.sheets.values().any(|actions| !actions.is_empty())
            || !self.added.is_empty()
    }

    pub(in crate::workbook::edit) fn remove_block(&self, reason: RemoveBlock, part: &str) -> Error {
        let position = self.removed.iter().next().copied().unwrap_or(0);
        let sheet = self
            .base
            .inner
            .sheets
            .get(position)
            .map_or("<missing sheet>", |sheet| sheet.name.as_str());
        Error::SheetRemoveBlocked {
            sheet: sheet.to_owned(),
            position,
            part: part.to_owned(),
            reason,
        }
    }
}
