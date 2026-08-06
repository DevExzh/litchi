//! Transaction-scoped worksheet and workbook-tab editors.

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

use super::layout::{ColumnEdit, DefaultsEdit, RowEdit};
use super::transaction::Edit;

/// Transaction-scoped state editor for any workbook sheet tab.
#[derive(Debug)]
pub struct TabEdit<'a> {
    pub(super) edit: &'a mut Edit,
    pub(super) position: usize,
}

impl TabEdit<'_> {
    /// Rename this tab and every modeled local formula/reference dependency.
    ///
    /// Borrowed strings are validated and copied once. Passing a prevalidated
    /// [`crate::sheet::Name`] or owned `String` keeps the operation move-first.
    pub fn rename<T>(&mut self, name: T) -> Result<&mut Self>
    where
        T: TryInto<Name>,
        Error: From<T::Error>,
    {
        self.edit
            .set_name(self.position, name.try_into().map_err(Error::from)?);
        Ok(self)
    }

    /// Make this the active tab while preserving unrelated grouped selections.
    /// The tab must be visible in the transaction's final state.
    pub fn activate(&mut self) -> &mut Self {
        self.edit.set_active(self.position);
        self
    }

    /// Make this tab visible.
    pub fn show(&mut self) -> &mut Self {
        self.edit.set_visibility(self.position, TabAction::Show);
        self
    }

    /// Hide this tab while retaining it in Excel's ordinary Unhide dialog.
    pub fn hide(&mut self) -> &mut Self {
        self.edit.set_visibility(self.position, TabAction::Hide);
        self
    }

    /// Hide this tab from Excel's ordinary Unhide dialog (`veryHidden`).
    pub fn very_hide(&mut self) -> &mut Self {
        self.edit.set_visibility(self.position, TabAction::VeryHide);
        self
    }
}

/// Borrowed worksheet editor tied to one transaction.
#[derive(Debug)]
pub struct WorksheetEdit<'a> {
    pub(super) edit: &'a mut Edit,
    pub(super) position: usize,
}

impl WorksheetEdit<'_> {
    /// Replace all worksheet Office Add-in range bindings by moving one
    /// already-validated collection into the transaction.
    pub fn set_bindings(&mut self, bindings: WebBindings) -> &mut Self {
        self.edit.sheets.entry(self.position).or_default().web = Some(bindings);
        self
    }

    /// Insert or replace one range binding, keyed by its semantic `appRef`.
    pub fn bind(&mut self, binding: WebBinding) -> Result<&mut Self> {
        let _ = self.edit.web_bindings(self.position)?.put(binding)?;
        Ok(self)
    }

    /// Edit one range binding transactionally by semantic `appRef` or checked
    /// numeric position. A failed closure leaves the staged value unchanged.
    pub fn edit_binding<'key>(
        &mut self,
        selector: impl Into<crate::web::Selector<'key>>,
        edit: impl FnOnce(&mut WebBinding) -> Result<()>,
    ) -> Result<bool> {
        self.edit.web_bindings(self.position)?.edit(selector, edit)
    }

    /// Remove one range binding by semantic `appRef` or checked position.
    pub fn unbind<'key>(
        &mut self,
        selector: impl Into<crate::web::Selector<'key>>,
    ) -> Result<Option<WebBinding>> {
        Ok(self.edit.web_bindings(self.position)?.remove(selector))
    }

    /// Remove every worksheet range binding.
    pub fn clear_bindings(&mut self) -> Result<bool> {
        Ok(self.edit.web_bindings(self.position)?.clear())
    }

    /// Borrow the worksheet-wide grid-default editor.
    pub fn defaults(&mut self) -> DefaultsEdit<'_> {
        DefaultsEdit {
            action: &mut self.edit.sheets.entry(self.position).or_default().defaults,
        }
    }

    /// Select one checked row for short property-editing verbs.
    pub fn row(&mut self, at: impl Into<RowAt>) -> Result<RowEdit<'_>> {
        let row = at.into().resolve()?;
        let style_lineage = &self.edit.base.inner.style_lineage;
        let actions = &mut self.edit.sheets.entry(self.position).or_default().rows;
        Ok(RowEdit {
            actions,
            style_lineage,
            row,
        })
    }

    /// Select one column by its primary A1 label or a checked zero-based input.
    pub fn column<'a>(&mut self, at: impl Into<ColumnAt<'a>>) -> Result<ColumnEdit<'_>> {
        let column = at.into().resolve()?;
        let style_lineage = &self.edit.base.inner.style_lineage;
        let actions = &mut self.edit.sheets.entry(self.position).or_default().columns;
        Ok(ColumnEdit {
            actions,
            style_lineage,
            column,
        })
    }

    /// Merge a checked rectangular selection without discarding follower data.
    ///
    /// The top-left cell remains the visible anchor. Commit returns a typed
    /// refusal if another covered cell has content, a group formula intersects
    /// the range, or producer compatibility state owns the merge collection.
    pub fn merge<'a>(&mut self, area: impl Into<Area<'a>>) -> Result<&mut Self> {
        let range = area.into().resolve()?;
        let sheet = self
            .edit
            .base
            .inner
            .sheets
            .get(self.position)
            .ok_or_else(|| invalid("merged-range target disappeared"))?
            .name
            .as_str();
        ensure_merge_area(sheet, range)?;
        let intents = &mut self.edit.sheets.entry(self.position).or_default().merges;
        intents
            .try_reserve(1)
            .map_err(|source| allocation("merged-range edit plan", source))?;
        intents.push(MergeIntent::Add(range));
        Ok(self)
    }

    /// Unmerge the range containing one checked cell, if any.
    pub fn unmerge<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        let range = {
            let data = self
                .edit
                .base
                .inner
                .sheets
                .get(self.position)
                .cloned()
                .ok_or_else(|| invalid("unmerge target disappeared"))?;
            let sheet = Worksheet {
                owner: Arc::clone(&self.edit.base.inner),
                data,
            };
            let intents = self
                .edit
                .sheets
                .get(&self.position)
                .map_or(&[][..], |actions| actions.merges.as_slice());
            pending_merge(sheet.store()?.merge_ranges(), intents, address)
        };
        let Some(range) = range else {
            return Ok(self);
        };
        let intents = &mut self.edit.sheets.entry(self.position).or_default().merges;
        intents
            .try_reserve(1)
            .map_err(|source| allocation("merged-range edit plan", source))?;
        intents.push(MergeIntent::Remove(range));
        Ok(self)
    }

    /// Create or replace a cell's primary value or formula.
    pub fn set<'a>(
        &mut self,
        at: impl Into<At<'a>>,
        content: impl Into<Content>,
    ) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        let content = content.into();
        content.validate_for_write()?;
        match self.edit.actions(self.position).entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::set(content));
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_payload(Payload::Set(content)),
        }
        Ok(self)
    }

    /// Remove primary value/formula content while retaining the cell record,
    /// local style, metadata, comments, and unknown non-payload children.
    pub fn clear<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        let actions = self.edit.actions(self.position);
        let create = actions.get(&address).is_some_and(|action| {
            matches!(action.payload(), Some(Payload::Set(_) | Payload::Clear))
        });
        match actions.entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::clear(create));
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_payload(if create {
                Payload::Clear
            } else {
                Payload::ClearIfPresent
            }),
        }
        Ok(self)
    }

    /// Apply an existing shared style without copying its formatting payload.
    ///
    /// The handle must belong to this edit's shared-style table lineage.
    /// Styling a missing cell creates an explicit empty cell record.
    pub fn style<'a>(&mut self, at: impl Into<At<'a>>, style: &Style) -> Result<&mut Self> {
        if !Arc::ptr_eq(
            &self.edit.base.inner.style_lineage,
            &style.owner.style_lineage,
        ) {
            return Err(Error::ForeignStyle);
        }
        let address = at.into().resolve()?;
        let effect = StyleEffect::Set(style.raw());
        match self.edit.actions(self.position).entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::style(style.raw()));
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_style(effect),
        }
        Ok(self)
    }

    /// Remove an explicit local style reference, retaining the cell payload.
    ///
    /// A missing cell remains missing and resolves as a no-op at commit.
    pub fn reset_style<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        match self.edit.actions(self.position).entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::reset_style());
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_style(StyleEffect::Reset),
        }
        Ok(self)
    }

    /// Remove the cell record without shifting surrounding cells.
    pub fn remove<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        self.edit
            .actions(self.position)
            .insert(address, Action::Remove);
        Ok(self)
    }
}

/// Borrowed editor for one worksheet being created by this transaction.
///
/// The handle owns no native Office identity and cannot outlive its edit.
#[derive(Debug)]
pub struct NewSheet<'a> {
    pub(super) added: &'a mut Added,
    pub(super) active: &'a mut Option<Target>,
    pub(super) style_lineage: &'a Arc<StyleLineage>,
    pub(super) index: usize,
    pub(super) position: usize,
}

impl NewSheet<'_> {
    /// Current checked zero-based projected position.
    ///
    /// A later tab move or insertion in this transaction can shift it. The
    /// authoritative committed position is recorded by [`Change::Create`].
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Current validated developer-facing name in the pending transaction.
    pub fn name(&self) -> &str {
        self.added.name.as_str()
    }

    /// Replace the pending name without exposing a physical sheet identity.
    pub fn rename<T>(&mut self, name: T) -> Result<&mut Self>
    where
        T: TryInto<Name>,
        Error: From<T::Error>,
    {
        self.added.name = name.try_into().map_err(Error::from)?;
        Ok(self)
    }

    /// Make the new visible worksheet active at commit.
    pub fn activate(&mut self) -> &mut Self {
        *self.active = Some(Target::Added(self.index));
        self
    }

    /// Keep the new worksheet visible (the default).
    pub fn show(&mut self) -> &mut Self {
        self.added.actions.visibility = Some(TabAction::Show);
        self
    }

    /// Create the worksheet hidden from the ordinary tab strip.
    pub fn hide(&mut self) -> &mut Self {
        self.added.actions.visibility = Some(TabAction::Hide);
        self
    }

    /// Create the worksheet hidden from Excel's ordinary Unhide dialog.
    pub fn very_hide(&mut self) -> &mut Self {
        self.added.actions.visibility = Some(TabAction::VeryHide);
        self
    }

    /// Move a complete checked Office Add-in range-binding collection onto
    /// this pending worksheet.
    pub fn set_bindings(&mut self, bindings: WebBindings) -> &mut Self {
        self.added.actions.web = Some(bindings);
        self
    }

    /// Insert or replace one range binding by semantic `appRef`.
    pub fn bind(&mut self, binding: WebBinding) -> Result<&mut Self> {
        let bindings = self.added.actions.web.get_or_insert_with(WebBindings::new);
        let _ = bindings.put(binding)?;
        Ok(self)
    }

    /// Remove one pending binding by semantic `appRef` or checked position.
    pub fn unbind<'key>(
        &mut self,
        selector: impl Into<crate::web::Selector<'key>>,
    ) -> Option<WebBinding> {
        self.added.actions.web.as_mut()?.remove(selector)
    }

    /// Borrow the pending worksheet's grid-default editor.
    pub fn defaults(&mut self) -> DefaultsEdit<'_> {
        DefaultsEdit {
            action: &mut self.added.actions.defaults,
        }
    }

    /// Select one checked row for short property-editing verbs.
    pub fn row(&mut self, at: impl Into<RowAt>) -> Result<RowEdit<'_>> {
        let row = at.into().resolve()?;
        Ok(RowEdit {
            actions: &mut self.added.actions.rows,
            style_lineage: self.style_lineage,
            row,
        })
    }

    /// Select one column by its primary A1 label or a checked zero-based input.
    pub fn column<'b>(&mut self, at: impl Into<ColumnAt<'b>>) -> Result<ColumnEdit<'_>> {
        let column = at.into().resolve()?;
        Ok(ColumnEdit {
            actions: &mut self.added.actions.columns,
            style_lineage: self.style_lineage,
            column,
        })
    }

    /// Merge a checked range on this transaction-local worksheet.
    pub fn merge<'b>(&mut self, area: impl Into<Area<'b>>) -> Result<&mut Self> {
        let range = area.into().resolve()?;
        ensure_merge_area(self.added.name.as_str(), range)?;
        self.added
            .actions
            .merges
            .try_reserve(1)
            .map_err(|source| allocation("new-sheet merged-range plan", source))?;
        self.added.actions.merges.push(MergeIntent::Add(range));
        Ok(self)
    }

    /// Unmerge the range containing one cell in this pending worksheet.
    pub fn unmerge<'b>(&mut self, at: impl Into<At<'b>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        let Some(range) = pending_merge(&[], &self.added.actions.merges, address) else {
            return Ok(self);
        };
        self.added
            .actions
            .merges
            .try_reserve(1)
            .map_err(|source| allocation("new-sheet merged-range plan", source))?;
        self.added.actions.merges.push(MergeIntent::Remove(range));
        Ok(self)
    }

    /// Create or replace a cell's primary value or formula.
    pub fn set<'a>(
        &mut self,
        at: impl Into<At<'a>>,
        content: impl Into<Content>,
    ) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        let content = content.into();
        content.validate_for_write()?;
        match self.added.actions.cells.entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::set(content));
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_payload(Payload::Set(content)),
        }
        Ok(self)
    }

    /// Clear pending primary content while retaining an explicit cell when a
    /// previous operation in this transaction created it.
    pub fn clear<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        let actions = &mut self.added.actions.cells;
        let create = actions.get(&address).is_some_and(|action| {
            matches!(action.payload(), Some(Payload::Set(_) | Payload::Clear))
        });
        match actions.entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::clear(create));
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_payload(if create {
                Payload::Clear
            } else {
                Payload::ClearIfPresent
            }),
        }
        Ok(self)
    }

    /// Apply an existing shared style without copying its formatting payload.
    pub fn style<'a>(&mut self, at: impl Into<At<'a>>, style: &Style) -> Result<&mut Self> {
        if !Arc::ptr_eq(self.style_lineage, &style.owner.style_lineage) {
            return Err(Error::ForeignStyle);
        }
        let address = at.into().resolve()?;
        let effect = StyleEffect::Set(style.raw());
        match self.added.actions.cells.entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::style(style.raw()));
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_style(effect),
        }
        Ok(self)
    }

    /// Remove an explicit local style reference from the pending cell.
    pub fn reset_style<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        match self.added.actions.cells.entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::reset_style());
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_style(StyleEffect::Reset),
        }
        Ok(self)
    }

    /// Remove the pending cell record without shifting surrounding cells.
    pub fn remove<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        self.added.actions.cells.insert(address, Action::Remove);
        Ok(self)
    }
}
