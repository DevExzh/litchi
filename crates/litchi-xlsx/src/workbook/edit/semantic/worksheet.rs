//! Transaction-scoped worksheet and workbook-tab editors.

use std::collections::{BTreeMap, BTreeSet, HashMap, btree_map::Entry};
use std::mem::size_of;
use std::sync::Arc;

use litchi_ooxml_common::web as common_web;
use litchi_opc::Part;
use litchi_sheet::{
    Area, At, Cell as Address, Column as ColumnIndex, ColumnAt, Rect, Row as RowIndex, RowAt,
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
use crate::style::cell_format::{self, Format};
use crate::style::{
    AlignmentChange, BorderChange, BorderEdge, FillChange, FontChange, ProtectionChange,
};
use crate::web::{Binding as WebBinding, Bindings as WebBindings};

use super::super::model::{
    PartChange, StyleGuard, defaults_after, ensure_merge_area, merge_conflicts, project_merges,
};
use super::super::svg::{PictureSelector, SvgInput};
use super::super::validation::{
    Added, FinalOrder, MergeIntent, MoveIntent, OrderPlan, PanesAction, Placement, SheetActions,
    TabAction, Target, pending_merge,
};
use super::super::{
    ActiveTab, Change, Commit, Conflict, ConflictSet, JoinError, JoinFailure, PackageChange, Patch,
    State,
};
use super::super::{codec, package};

use super::layout::{Column, Defaults, Row};
use super::transaction::{AddedCaptureReservation, CaptureBudget, CaptureFacet, Edit};

/// Transaction-scoped state editor for any workbook sheet tab.
#[derive(Debug)]
pub struct Tab<'a> {
    pub(super) edit: &'a mut Edit,
    pub(super) position: usize,
}

impl Tab<'_> {
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
pub struct Sheet<'a> {
    pub(super) edit: &'a mut Edit,
    pub(super) position: usize,
}

/// Transaction-scoped scalar editor for one existing worksheet form control.
///
/// Staging validates the selected control and paired properties/VML edits.
/// The surrounding workbook transaction publishes them atomically at commit.
/// One control per worksheet batch is supported; independent form-control
/// branches on the same worksheet conflict when joined. In-memory patches
/// require the exact immutable source workbook; durable patches check the
/// complete serialized source.
#[derive(Debug)]
pub struct FormControl<'a> {
    pub(super) edit: &'a mut Edit,
    pub(super) position: usize,
    pub(super) selector: crate::form_control::OwnedSelector,
}

impl FormControl<'_> {
    /// Stage one typed scalar on the selected control.
    pub fn set_scalar(
        &mut self,
        field: crate::form_control::ScalarField,
        value: Option<crate::form_control::ScalarValue>,
    ) -> Result<&mut Self> {
        self.edit.stage_form_control_scalar(
            self.position,
            self.selector.selector(),
            field,
            value,
        )?;
        Ok(self)
    }
}

impl Sheet<'_> {
    fn source_storage(&self, address: Address) -> Result<Option<(Cell, bool, Option<usize>)>> {
        let sheet = self
            .edit
            .base
            .sheet(self.position)?
            .ok_or_else(|| invalid("worksheet edit target disappeared"))?;
        Ok(sheet.store()?.entry(address).map(|stored| {
            (
                stored.cell.clone(),
                stored.inline_rich,
                stored.shared_string,
            )
        }))
    }

    fn source_cell(&self, address: Address) -> Result<Option<Cell>> {
        Ok(self.source_storage(address)?.map(|(cell, _, _)| cell))
    }

    fn source_has_rich_text(&self, address: Address) -> Result<bool> {
        let sheet = self
            .edit
            .base
            .sheet(self.position)?
            .ok_or_else(|| invalid("worksheet edit target disappeared"))?;
        sheet.has_rich_text(address)
    }

    fn source_rich_text(&self, address: Address) -> Result<Option<crate::RichText>> {
        let sheet = self
            .edit
            .base
            .sheet(self.position)?
            .ok_or_else(|| invalid("worksheet edit target disappeared"))?;
        sheet.rich_text(address)
    }

    fn source_style(&self, address: Address) -> Result<Option<u32>> {
        let sheet = self
            .edit
            .base
            .sheet(self.position)?
            .ok_or_else(|| invalid("worksheet edit target disappeared"))?;
        Ok(sheet
            .store()?
            .entry(address)
            .and_then(|stored| stored.style))
    }

    fn source_format(&self, address: Address) -> Result<Option<Format>> {
        let sheet = self
            .edit
            .base
            .sheet(self.position)?
            .ok_or_else(|| invalid("worksheet edit target disappeared"))?;
        sheet
            .style(address)?
            .map(|style| style.format())
            .transpose()
    }

    /// Resolve the format that a typed request will actually see.  A pending
    /// complete style is the effective baseline for later typed facets; using
    /// the immutable source here would incorrectly retain a request that is
    /// already equal to that pending style.
    fn effective_typed_format(&self, address: Address) -> Result<Option<Format>> {
        let Some(action) = self
            .edit
            .sheets
            .get(&self.position)
            .and_then(|actions| actions.cells.get(&address))
        else {
            return self.source_format(address);
        };
        let Some(style) = (match action {
            Action::Update {
                style: Some(style), ..
            } => Some(style),
            _ => None,
        }) else {
            return self.source_format(address);
        };
        let raw = match style {
            StyleEffect::Set(raw) => *raw,
            StyleEffect::Reset => 0,
        };
        self.edit.base.inner.format(raw).map(Some)
    }

    fn has_pending_style(&self, address: Address) -> bool {
        self.edit.sheets.get(&self.position).is_some_and(|actions| {
            actions.fonts.contains_key(&address)
                || actions.alignments.contains_key(&address)
                || actions.fills.contains_key(&address)
                || actions.borders.contains_key(&address)
                || actions.protections.contains_key(&address)
                || actions.quote_prefixes.contains_key(&address)
                || actions.number_formats.contains_key(&address)
                || actions
                    .cells
                    .get(&address)
                    .is_some_and(|action| matches!(action, Action::Update { style: Some(_), .. }))
        })
    }

    fn has_pending_complete_style(&self, address: Address) -> bool {
        self.edit
            .sheets
            .get(&self.position)
            .and_then(|actions| actions.cells.get(&address))
            .is_some_and(|action| {
                matches!(
                    action,
                    Action::Update { style: Some(_), .. } | Action::Remove
                )
            })
    }

    /// A typed request after removal recreates the cell's style-bearing
    /// record. Keep the supported remove-then-style sequence instead of
    /// leaving an impossible `Remove` action beside deferred style maps.
    fn promote_remove_for_typed(&mut self, address: Address) {
        let removed = self
            .edit
            .sheets
            .get(&self.position)
            .and_then(|actions| actions.cells.get(&address))
            .is_some_and(|action| matches!(action, Action::Remove));
        if !removed {
            return;
        }
        self.edit
            .release_cell_capture_facet(self.position, address, CaptureFacet::Remove);
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .cells
            .insert(
                address,
                Action::Update {
                    payload: None,
                    style: None,
                },
            );
    }

    /// Discard typed style requests when a complete style operation resolves
    /// to the immutable source.  This path runs before capture admission, so
    /// it must release every facet explicitly rather than relying on a later
    /// complete-style replacement to overwrite the ledger.
    fn discard_typed_style_noop(&mut self, address: Address) -> Result<()> {
        let had_fill = self
            .edit
            .sheets
            .get(&self.position)
            .is_some_and(|actions| actions.fills.contains_key(&address));
        let had_border = self
            .edit
            .sheets
            .get(&self.position)
            .is_some_and(|actions| actions.borders.contains_key(&address));
        self.edit.discard_fill(self.position, address)?;
        self.edit.discard_border(self.position, address)?;
        let (font, alignment, protection, quote_prefix, number_format) = self
            .edit
            .sheets
            .get_mut(&self.position)
            .map_or((false, false, false, false, false), |actions| {
                (
                    actions.fonts.remove(&address).is_some(),
                    actions.alignments.remove(&address).is_some(),
                    actions.protections.remove(&address).is_some(),
                    actions.quote_prefixes.remove(&address).is_some(),
                    actions.number_formats.remove(&address).is_some(),
                )
            });
        if had_fill {
            self.edit
                .release_cell_capture_facet(self.position, address, CaptureFacet::Fill);
        }
        if had_border {
            self.edit
                .release_cell_capture_facet(self.position, address, CaptureFacet::Border);
        }
        if font {
            self.edit
                .release_cell_capture_facet(self.position, address, CaptureFacet::Font);
        }
        if alignment {
            self.edit
                .release_cell_capture_facet(self.position, address, CaptureFacet::Alignment);
        }
        if protection {
            self.edit
                .release_cell_capture_facet(self.position, address, CaptureFacet::Protection);
        }
        if quote_prefix {
            self.edit
                .release_cell_capture_facet(self.position, address, CaptureFacet::QuotePrefix);
        }
        if number_format {
            self.edit.release_cell_capture_facet(
                self.position,
                address,
                CaptureFacet::NumberFormat,
            );
        }
        Ok(())
    }

    /// Cancel a pending complete style when the later complete request is the
    /// immutable source style.  This is deliberately a bounded-edit boundary:
    /// ordinary edits keep the setter fast and defer the source comparison to
    /// the join/commit effect planner.  A bounded edit, however, must release
    /// the superseded complete-style capture before admitting the next
    /// operation, or a stale reservation can reject an otherwise fitting
    /// payload.
    fn discard_source_equivalent_complete_style(
        &mut self,
        address: Address,
        requested: Option<u32>,
    ) -> Result<bool> {
        if !self.edit.capture_budgeted()
            || !self
                .edit
                .sheets
                .get(&self.position)
                .and_then(|actions| actions.cells.get(&address))
                .is_some_and(|action| matches!(action, Action::Update { style: Some(_), .. }))
        {
            return Ok(false);
        }
        if self.source_style(address)? != requested {
            return Ok(false);
        }
        self.discard_typed_style_noop(address)?;
        let remove_action = self
            .edit
            .sheets
            .get_mut(&self.position)
            .and_then(|actions| actions.cells.get_mut(&address))
            .is_some_and(|action| match action {
                Action::Update { payload, style } => {
                    *style = None;
                    payload.is_none()
                },
                Action::Remove => false,
            });
        if remove_action {
            self.edit
                .sheets
                .get_mut(&self.position)
                .and_then(|actions| actions.cells.remove(&address));
        }
        self.edit
            .release_cell_capture_facet(self.position, address, CaptureFacet::Style);
        Ok(true)
    }

    fn has_pending_cell_effect(&self, address: Address) -> bool {
        self.edit.sheets.get(&self.position).is_some_and(|actions| {
            actions.cells.contains_key(&address)
                || actions.fonts.contains_key(&address)
                || actions.alignments.contains_key(&address)
                || actions.fills.contains_key(&address)
                || actions.borders.contains_key(&address)
                || actions.protections.contains_key(&address)
                || actions.quote_prefixes.contains_key(&address)
                || actions.number_formats.contains_key(&address)
        })
    }

    /// Stage one selector-first scalar on an existing form control.
    pub fn set_form_control_scalar<'a>(
        &mut self,
        selector: impl Into<crate::form_control::ControlSelector<'a>>,
        field: crate::form_control::ScalarField,
        value: Option<crate::form_control::ScalarValue>,
    ) -> Result<&mut Self> {
        self.edit
            .stage_form_control_scalar(self.position, selector, field, value)?;
        Ok(self)
    }

    /// Borrow a scalar editor pinned to one existing form-control selector.
    pub fn edit_form_control<'e, 's>(
        &'e mut self,
        selector: impl Into<crate::form_control::ControlSelector<'s>>,
    ) -> Result<FormControl<'e>> {
        let selector = selector.into();
        let target = crate::form_control::owned_selector(selector);
        {
            let data = self
                .edit
                .base
                .inner
                .sheets
                .get(self.position)
                .ok_or_else(|| invalid("form-control worksheet disappeared"))?;
            if data.kind != WorksheetKind::Worksheet {
                return Err(Error::NotWorksheet {
                    sheet: data.name.clone(),
                });
            }
            crate::form_control::validate_ordinary_form_control_selector(
                &self.edit.base.inner.package,
                &data.part_uri,
                target.selector(),
            )?;
        }
        Ok(FormControl {
            edit: self.edit,
            position: self.position,
            selector: target,
        })
    }

    /// Attach one borrowed SVG payload to an existing direct worksheet
    /// picture selected by source-order drawing and picture ordinals.
    ///
    /// The payload is validated against the transaction's finite input limit
    /// before a staged copy is allocated. Package IDs, relationship IDs, and
    /// media names remain transaction-owned.
    pub fn attach_svg<'svg>(
        &mut self,
        picture: PictureSelector,
        input: SvgInput<'svg>,
    ) -> Result<&mut Self> {
        self.edit.stage_svg_attach(self.position, picture, input)?;
        Ok(self)
    }

    /// Detach the admitted embedded SVG owner from one existing direct
    /// worksheet picture while retaining its raster fallback.
    pub fn detach_svg(&mut self, picture: PictureSelector) -> Result<&mut Self> {
        let _ = self.edit.stage_svg_detach(self.position, picture)?;
        Ok(self)
    }

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

    /// Insert or replace one typed worksheet hyperlink by its checked range.
    ///
    /// External targets are inert strings. The worksheet XML and its external
    /// relationship are staged together and published atomically at commit.
    pub fn put_hyperlink(&mut self, hyperlink: crate::hyperlinks::Hyperlink) -> Result<&mut Self> {
        self.edit.put_hyperlink_at(self.position, hyperlink)?;
        Ok(self)
    }

    /// Remove one hyperlink selected by its checked cell or range reference.
    pub fn remove_hyperlink(
        &mut self,
        reference: crate::hyperlinks::HyperlinkReference,
    ) -> Result<Option<crate::hyperlinks::Hyperlink>> {
        self.edit.remove_hyperlink_at(self.position, reference)
    }

    /// Borrow the worksheet-wide grid-default editor.
    pub fn defaults(&mut self) -> Defaults<'_> {
        Defaults {
            action: &mut self.edit.sheets.entry(self.position).or_default().defaults,
        }
    }

    /// Select one checked row for short property-editing verbs.
    pub fn row(&mut self, at: impl Into<RowAt>) -> Result<Row<'_>> {
        let row = at.into().resolve()?;
        let style_lineage = &self.edit.base.inner.style_lineage;
        let actions = &mut self.edit.sheets.entry(self.position).or_default().rows;
        Ok(Row {
            actions,
            style_lineage,
            row,
        })
    }

    /// Select one column by its primary A1 label or a checked zero-based input.
    pub fn column<'a>(&mut self, at: impl Into<ColumnAt<'a>>) -> Result<Column<'_>> {
        let column = at.into().resolve()?;
        let style_lineage = &self.edit.base.inner.style_lineage;
        let actions = &mut self.edit.sheets.entry(self.position).or_default().columns;
        Ok(Column {
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
        // Resolve the projected merge state before taking a finite snapshot
        // reservation. Re-adding an already-present range is a semantic
        // no-op, while an identical pending add is a duplicate request; both
        // must avoid charging a second structural capture.
        let data = self
            .edit
            .base
            .inner
            .sheets
            .get(self.position)
            .cloned()
            .ok_or_else(|| invalid("merged-range target disappeared"))?;
        let source_sheet = Worksheet {
            owner: Arc::clone(&self.edit.base.inner),
            data,
            policy: None,
        };
        let base_merges = source_sheet.store()?.merge_ranges().to_vec();
        let intents = self
            .edit
            .sheets
            .get(&self.position)
            .map_or(&[][..], |actions| actions.merges.as_slice());
        let projected = pending_merge(&base_merges, intents, range.start());
        if projected == Some(range) {
            if base_merges.contains(&range) {
                let keep_structural_capture =
                    if let Some(actions) = self.edit.sheets.get_mut(&self.position) {
                        actions.merges.retain(|intent| intent.range() != range);
                        actions
                            .merges
                            .iter()
                            .any(|intent| intent.range().start() == range.start())
                    } else {
                        false
                    };
                if !keep_structural_capture {
                    self.edit.release_cell_capture_facet(
                        self.position,
                        range.start(),
                        CaptureFacet::Structural,
                    );
                }
            }
            return Ok(self);
        }
        let capture = self.edit.preflight_cell_capture_with_facet(
            self.position,
            range.start(),
            CaptureFacet::Structural,
            0,
        )?;
        let reserve = self
            .edit
            .sheets
            .entry(self.position)
            .or_default()
            .merges
            .try_reserve(1);
        if let Err(source) = reserve {
            self.edit.rollback_cell_capture(capture);
            return Err(allocation("merged-range edit plan", source));
        }
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .merges
            .push(MergeIntent::Add(range));
        Ok(self)
    }

    /// Unmerge the range containing one checked cell, if any.
    pub fn unmerge<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        let range = (|| -> Result<Option<Rect>> {
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
                policy: None,
            };
            let intents = self
                .edit
                .sheets
                .get(&self.position)
                .map_or(&[][..], |actions| actions.merges.as_slice());
            Ok(pending_merge(
                sheet.store()?.merge_ranges(),
                intents,
                address,
            ))
        })()?;
        let Some(range) = range else {
            return Ok(self);
        };
        let capture = self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::Structural,
            0,
        )?;
        let cancel_add = self
            .edit
            .sheets
            .get(&self.position)
            .and_then(|actions| actions.merges.last())
            .is_some_and(|intent| *intent == MergeIntent::Add(range));
        if cancel_add {
            self.edit
                .sheets
                .entry(self.position)
                .or_default()
                .merges
                .pop();
            self.edit.rollback_cell_capture(capture);
            self.edit.release_cell_capture_facet(
                self.position,
                range.start(),
                CaptureFacet::Structural,
            );
            return Ok(self);
        }
        let reserve = self
            .edit
            .sheets
            .entry(self.position)
            .or_default()
            .merges
            .try_reserve(1);
        if let Err(source) = reserve {
            self.edit.rollback_cell_capture(capture);
            return Err(allocation("merged-range edit plan", source));
        }
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .merges
            .push(MergeIntent::Remove(range));
        Ok(self)
    }

    /// Create or replace a cell's primary value or formula.
    pub fn set<'a>(
        &mut self,
        at: impl Into<At<'a>>,
        content: impl Into<Content>,
    ) -> Result<&mut Self> {
        self.set_coalesced(at, content).map(|_| self)
    }

    /// Internal payload setter used by the unified XLSX coalescer.
    ///
    /// It returns `false` only when a bounded edit proves an exact
    /// source-equivalent no-op; otherwise it returns `true` after staging a
    /// payload. The hidden entry point keeps the same bounded capture,
    /// source-equivalent no-op, and action-map behavior as [`Self::set`].
    /// The unified coalescer calls it only for distinct, unsigned bounded
    /// targets. Unbounded or repeated callers must not interpret the result
    /// as a general before/after projection.
    #[doc(hidden)]
    pub fn set_coalesced<'a>(
        &mut self,
        at: impl Into<At<'a>>,
        content: impl Into<Content>,
    ) -> Result<bool> {
        let address = at.into().resolve()?;
        let content = content.into();
        content.validate_for_write()?;
        let source = if self.edit.capture_budgeted() {
            self.source_storage(address)?
        } else {
            None
        };
        let source_cell = source.as_ref().map(|(cell, _, _)| cell);
        let source_same = source_cell.is_some_and(|source| content.same_as_cell(source));
        let source_has_rich_text = if source_same {
            match source.as_ref() {
                Some((_, true, _)) => true,
                Some((_, false, Some(_))) => self.source_has_rich_text(address)?,
                _ => false,
            }
        } else {
            false
        };
        // Compare against the immutable source before reserving the authored
        // payload.  A bounded endpoint-only allowance must accept a genuine
        // no-op, while an unbounded edit keeps its historical fast path and
        // does not perform this source census.
        if self.edit.capture_budgeted() && source_same && !source_has_rich_text {
            let mut remove_action = false;
            let mut remove_all_capture = false;
            if let Some(action) = self
                .edit
                .sheets
                .get_mut(&self.position)
                .and_then(|actions| actions.cells.get_mut(&address))
            {
                match action {
                    Action::Update { payload, style } => {
                        *payload = None;
                        remove_action = style.is_none();
                    },
                    Action::Remove => {
                        remove_action = true;
                        remove_all_capture = true;
                    },
                }
            }
            if remove_action {
                self.edit
                    .sheets
                    .get_mut(&self.position)
                    .and_then(|actions| actions.cells.remove(&address));
            }
            if remove_all_capture {
                self.edit.release_cell_capture_all(self.position, address);
            } else {
                self.edit
                    .release_cell_capture_facet(self.position, address, CaptureFacet::Payload);
            }
            return Ok(false);
        }
        let _capture =
            self.edit
                .preflight_cell_capture(self.position, address, content.retained_bytes())?;
        match self.edit.actions(self.position).entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::set(content));
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_payload(Payload::Set(content)),
        }
        Ok(true)
    }

    /// Write `text` as the cell's own inline rich string: each run keeps
    /// its run properties (`rPr`) as read by
    /// [`Worksheet::rich_text`](crate::Worksheet::rich_text) (record 0295).
    /// The cell's style and other children are kept, as by [`Self::set`].
    ///
    /// # Errors
    ///
    /// Returns an invalid-input error for empty or over-long rich text.
    pub fn set_rich_text<'a>(
        &mut self,
        at: impl Into<At<'a>>,
        text: crate::RichText,
    ) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        text.validate_for_write()?;
        let source_same = if self.edit.capture_budgeted() {
            match self.source_storage(address)? {
                Some((_, true, _)) => match self.source_rich_text(address)? {
                    Some(source) => source == text,
                    None => false,
                },
                // Shared strings, including shared rich strings, must be
                // rewritten as an inline rich payload. Only an identical
                // inline-rich representation is a bounded setter no-op.
                Some((_, false, _)) | None => false,
            }
        } else {
            // The commit normalizer performs the exact representation-aware
            // comparison for unbounded edits. Keep the authoring fast path
            // free of source storage and rich-run scans.
            false
        };
        if source_same {
            let mut remove_action = false;
            let mut remove_all_capture = false;
            if let Some(action) = self
                .edit
                .sheets
                .get_mut(&self.position)
                .and_then(|actions| actions.cells.get_mut(&address))
            {
                match action {
                    Action::Update { payload, style } => {
                        *payload = None;
                        remove_action = style.is_none();
                    },
                    Action::Remove => {
                        remove_action = true;
                        remove_all_capture = true;
                    },
                }
            }
            if remove_action {
                self.edit
                    .sheets
                    .get_mut(&self.position)
                    .and_then(|actions| actions.cells.remove(&address));
            }
            if remove_all_capture {
                self.edit.release_cell_capture_all(self.position, address);
            } else {
                self.edit
                    .release_cell_capture_facet(self.position, address, CaptureFacet::Payload);
            }
            return Ok(self);
        }
        let _capture =
            self.edit
                .preflight_cell_capture(self.position, address, text.retained_bytes())?;
        match self.edit.actions(self.position).entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::rich(text));
            },
            Entry::Occupied(mut entry) => {
                entry.get_mut().set_payload(Payload::Rich(Box::new(text)))
            },
        }
        Ok(self)
    }

    /// Store `value` as the cached result of an existing scalar formula
    /// cell, keeping its formula markup (shared-formula members included)
    /// byte for byte; the cell type follows the value (record 0279).
    ///
    /// ```
    /// # fn main() -> litchi_xlsx::Result<()> {
    /// use litchi_xlsx::{Cell, Formula, Package, State, Value};
    ///
    /// let mut edit = Package::new()?.edit()?;
    /// edit.sheet("Sheet1")?.expect("worksheet").set("A1", Formula::new("1+1")?)?;
    /// let workbook = edit.commit()?.into_package();
    /// let mut edit = workbook.edit()?;
    /// let mut sheet = edit.sheet("Sheet1")?.expect("worksheet");
    /// sheet.set_cached("A1", 2)?;
    /// let State::Cell { content: Cell::Formula(formula), .. } = sheet.state("A1")? else {
    ///     unreachable!()
    /// };
    /// assert_eq!(formula.cached().map(|cache| cache.value()), Some(&Value::from(2)));
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an invalid-input error when the stored cell is not a scalar
    /// formula, the value cannot be written, or the same cell already has a
    /// payload change in this edit.
    pub fn set_cached<'a>(
        &mut self,
        at: impl Into<At<'a>>,
        value: impl Into<crate::Value>,
    ) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        let value = value.into();
        Content::Value(value.clone()).validate_for_write()?;
        let sheet = match self.edit.base.sheet(self.position) {
            Ok(Some(sheet)) => sheet,
            Ok(None) => return Err(invalid("worksheet edit target disappeared")),
            Err(error) => return Err(error),
        };
        let (scalar, source_same) = match sheet.store() {
            Ok(store) => match store.entry(address).map(|stored| &stored.cell) {
                Some(Cell::Formula(formula))
                    if matches!(formula.kind(), crate::formula::Kind::Scalar) =>
                {
                    (
                        true,
                        formula
                            .cached()
                            .is_some_and(|cache| cache.value() == &value),
                    )
                },
                _ => (false, false),
            },
            Err(error) => return Err(error),
        };
        if !scalar {
            return Err(invalid("a formula cache needs a stored scalar formula"));
        }
        let pending_payload = self
            .edit
            .sheets
            .get(&self.position)
            .and_then(|actions| actions.cells.get(&address))
            .and_then(Action::payload)
            .is_some();
        let pending_cache_same = self
            .edit
            .sheets
            .get(&self.position)
            .and_then(|actions| actions.cells.get(&address))
            .and_then(Action::payload)
            .is_some_and(|payload| matches!(payload, Payload::Cache(cached) if cached == &value));
        let pending_remove = self
            .edit
            .sheets
            .get(&self.position)
            .and_then(|actions| actions.cells.get(&address))
            .is_some_and(|action| matches!(action, Action::Remove));
        if source_same {
            if !pending_payload {
                if pending_remove {
                    self.edit
                        .sheets
                        .get_mut(&self.position)
                        .and_then(|actions| actions.cells.remove(&address));
                    self.edit.release_cell_capture_all(self.position, address);
                }
                return Ok(self);
            }
            if pending_cache_same {
                let mut remove_action = false;
                if let Some(action) = self
                    .edit
                    .sheets
                    .get_mut(&self.position)
                    .and_then(|actions| actions.cells.get_mut(&address))
                {
                    if let Action::Update { payload, style } = action {
                        *payload = None;
                        remove_action = style.is_none();
                    }
                }
                if remove_action {
                    self.edit
                        .sheets
                        .get_mut(&self.position)
                        .and_then(|actions| actions.cells.remove(&address));
                }
                self.edit
                    .release_cell_capture_facet(self.position, address, CaptureFacet::Payload);
                return Ok(self);
            }
        }
        if pending_payload || pending_remove {
            return Err(invalid(
                "a formula cache cannot join another payload change of the same cell",
            ));
        }
        let _capture =
            self.edit
                .preflight_cell_capture(self.position, address, value.retained_bytes())?;
        match self.edit.actions(self.position).entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::cache(value));
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_payload(Payload::Cache(value)),
        }
        Ok(self)
    }

    /// Remove primary value/formula content while retaining the cell record,
    /// local style, metadata, comments, and unknown non-payload children.
    pub fn clear<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        let source = if self.edit.capture_budgeted() {
            Some(self.source_cell(address)?)
        } else {
            None
        };
        let (source_exists, source_empty) = source.as_ref().map_or((true, false), |source| {
            (
                source.is_some(),
                source
                    .as_ref()
                    .is_some_and(|cell| matches!(cell, Cell::Empty)),
            )
        });
        if self.edit.capture_budgeted() {
            if !source_exists && !self.has_pending_cell_effect(address) {
                return Ok(self);
            }
            if source_empty {
                let (pending_clear, pending_payload, pending_style, pending_exists) = self
                    .edit
                    .sheets
                    .get(&self.position)
                    .and_then(|actions| actions.cells.get(&address))
                    .map_or((false, false, false, false), |action| {
                        (
                            matches!(
                                action,
                                Action::Update {
                                    payload: Some(Payload::Clear | Payload::ClearIfPresent),
                                    style: None,
                                }
                            ),
                            action.payload().is_some(),
                            matches!(action, Action::Update { style: Some(_), .. }),
                            true,
                        )
                    });
                if pending_clear {
                    self.edit
                        .sheets
                        .get_mut(&self.position)
                        .and_then(|actions| actions.cells.remove(&address));
                    self.edit.release_cell_capture_facet(
                        self.position,
                        address,
                        CaptureFacet::Payload,
                    );
                    return Ok(self);
                }
                if !pending_payload && !pending_style && !pending_exists {
                    return Ok(self);
                }
            }
        }
        let capture =
            self.edit
                .preflight_cell_capture(self.position, address, size_of::<Content>())?;
        if !source_exists && !self.has_pending_cell_effect(address) {
            self.edit.rollback_cell_capture(capture);
            return Ok(self);
        }
        let actions = self.edit.actions(self.position);
        let create = actions.get(&address).is_some_and(|action| {
            matches!(
                action.payload(),
                Some(
                    Payload::Set(_)
                        | Payload::SharedString { .. }
                        | Payload::Rich(_)
                        | Payload::Clear
                )
            )
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
        if source_empty {
            let mut remove = false;
            if let Some(action) = self
                .edit
                .sheets
                .get_mut(&self.position)
                .and_then(|actions| actions.cells.get_mut(&address))
            {
                match action {
                    Action::Update {
                        payload: Some(Payload::Clear | Payload::ClearIfPresent),
                        style,
                    } if style.is_none() => remove = true,
                    Action::Update {
                        payload: Some(Payload::Clear | Payload::ClearIfPresent),
                        ..
                    } => {},
                    _ => {},
                }
            }
            if remove {
                self.edit
                    .sheets
                    .get_mut(&self.position)
                    .and_then(|actions| actions.cells.remove(&address));
                self.edit
                    .release_cell_capture_facet(self.position, address, CaptureFacet::Payload);
            }
        }
        Ok(self)
    }

    /// Return the projected semantic cell state at one address.
    ///
    /// The projection combines the immutable source cell with the pending
    /// payload/style action for this worksheet. It is intentionally a
    /// cell-local staging view: commit-time validation, worksheet structure,
    /// and deferred style lowering are not projected here. A cell with any
    /// pending style request is therefore refused until the request has been
    /// lowered by commit planning. Shared-string keys are retained by the same
    /// state reader used by native patch generation.
    ///
    /// A worksheet scheduled for removal cannot produce a meaningful final
    /// cell state and is refused. Tab reordering does not change the base
    /// worksheet identity used by this handle, so it does not affect the
    /// returned cell state.
    ///
    /// ```
    /// # fn main() -> litchi_xlsx::Result<()> {
    /// use litchi_xlsx::{Cell, State, Value, Package};
    ///
    /// let workbook = Package::new()?;
    /// let mut edit = workbook.edit()?;
    /// let mut sheet = edit.sheet("Sheet1")?.expect("worksheet");
    /// sheet.set("A1", "draft")?;
    /// assert!(matches!(
    ///     sheet.state("A1")?,
    ///     State::Cell {
    ///         content: Cell::Value(Value::Text(value)),
    ///         ..
    ///     } if value.as_str() == "draft"
    /// ));
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// Formula/reference rewrites staged by other transaction operations and
    /// commit-time formula validation are not projected; callers must treat
    /// this as the cell-local payload/style view at the current staging point.
    ///
    /// # Errors
    ///
    /// Returns the coordinate error for an invalid address, the source
    /// worksheet/package error when the cell store cannot be read, or a typed
    /// unsupported error when the worksheet is scheduled for removal or the
    /// selected cell has deferred style lowering.
    pub fn state<'a>(&self, at: impl Into<At<'a>>) -> Result<State> {
        let address = at.into().resolve()?;
        if self.edit.removed.contains(&self.position) {
            return Err(Error::Unsupported {
                feature: "projecting cell state for a worksheet scheduled for removal",
            });
        }
        let Some(actions) = self.edit.sheets.get(&self.position) else {
            let sheet = self
                .edit
                .base
                .sheet(self.position)?
                .ok_or_else(|| invalid("worksheet edit target disappeared"))?;
            let stored = sheet.store()?.entry(address);
            return Ok(State::read(stored, &self.edit.base));
        };
        if actions.fonts.contains_key(&address)
            || actions.alignments.contains_key(&address)
            || actions.fills.contains_key(&address)
            || actions.borders.contains_key(&address)
            || actions.protections.contains_key(&address)
            || actions.quote_prefixes.contains_key(&address)
            || actions.number_formats.contains_key(&address)
        {
            return Err(Error::Unsupported {
                feature: "projecting cell state with pending style changes",
            });
        }
        let sheet = self
            .edit
            .base
            .sheet(self.position)?
            .ok_or_else(|| invalid("worksheet edit target disappeared"))?;
        let stored = sheet.store()?.entry(address);
        Ok(actions.cells.get(&address).map_or_else(
            || State::read(stored, &self.edit.base),
            |action| State::after(stored, action, &self.edit.base),
        ))
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
        if self.discard_source_equivalent_complete_style(address, Some(style.raw()))? {
            return Ok(self);
        }
        let had_pending_style = self.has_pending_style(address);
        let has_pending_complete_style = self.has_pending_complete_style(address);
        if had_pending_style
            && self
                .edit
                .sheets
                .get(&self.position)
                .and_then(|actions| actions.cells.get(&address))
                .is_some_and(|action| {
                    matches!(
                        action,
                        Action::Update {
                            payload: _,
                            style: Some(StyleEffect::Set(key)),
                        } if *key == style.raw()
                    )
                })
        {
            return Ok(self);
        }
        if self.edit.capture_budgeted() && !has_pending_complete_style {
            let source_style = self.source_style(address)?;
            if source_style == Some(style.raw()) {
                self.discard_typed_style_noop(address)?;
                return Ok(self);
            }
        }
        let capture = self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::Style,
            size_of::<StyleEffect>(),
        )?;
        // An explicit shared-style assignment is the later, complete style
        // operation. It supersedes any deferred property-level style request
        // staged for the same cell.
        if let Err(error) = self.edit.discard_fill(self.position, address) {
            self.edit.rollback_cell_capture(capture);
            return Err(error);
        }
        if let Err(error) = self.edit.discard_border(self.position, address) {
            self.edit.rollback_cell_capture(capture);
            return Err(error);
        }
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .fonts
            .remove(&address);
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .alignments
            .remove(&address);
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .protections
            .remove(&address);
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .quote_prefixes
            .remove(&address);
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .number_formats
            .remove(&address);
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
        if self.discard_source_equivalent_complete_style(address, None)? {
            return Ok(self);
        }
        let had_pending_style = self.has_pending_style(address);
        let has_pending_complete_style = self.has_pending_complete_style(address);
        if had_pending_style
            && self
                .edit
                .sheets
                .get(&self.position)
                .and_then(|actions| actions.cells.get(&address))
                .is_some_and(|action| {
                    matches!(
                        action,
                        Action::Update {
                            payload: _,
                            style: Some(StyleEffect::Reset),
                        }
                    )
                })
        {
            return Ok(self);
        }
        if self.edit.capture_budgeted() && !has_pending_complete_style {
            if self.source_style(address)?.is_none() {
                self.discard_typed_style_noop(address)?;
                return Ok(self);
            }
        }
        let source_missing = if self.edit.capture_budgeted() {
            Some(self.source_cell(address)?.is_none())
        } else {
            None
        };
        let capture = self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::Style,
            size_of::<StyleEffect>(),
        )?;
        if let Err(error) = self.edit.discard_fill(self.position, address) {
            self.edit.rollback_cell_capture(capture);
            return Err(error);
        }
        if let Err(error) = self.edit.discard_border(self.position, address) {
            self.edit.rollback_cell_capture(capture);
            return Err(error);
        }
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .fonts
            .remove(&address);
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .alignments
            .remove(&address);
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .protections
            .remove(&address);
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .quote_prefixes
            .remove(&address);
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .number_formats
            .remove(&address);
        match self.edit.actions(self.position).entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::reset_style());
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_style(StyleEffect::Reset),
        }
        let source_missing = source_missing.unwrap_or(false);
        let cancel_missing_style = source_missing
            && self
                .edit
                .sheets
                .get(&self.position)
                .and_then(|actions| actions.cells.get(&address))
                .is_some_and(|action| {
                    matches!(
                        action,
                        Action::Update {
                            payload: None,
                            style: Some(StyleEffect::Reset),
                        }
                    )
                });
        if cancel_missing_style {
            self.edit
                .sheets
                .get_mut(&self.position)
                .and_then(|actions| actions.cells.remove(&address));
            self.edit.release_cell_capture_all(self.position, address);
        }
        Ok(self)
    }

    /// Stage typed font properties for one cell.
    ///
    /// The request is resolved against the workbook's shared stylesheet at
    /// commit. No native font or cell-format identifier is exposed, and the
    /// source worksheet remains unchanged until the transaction publishes.
    pub fn font<'a>(&mut self, at: impl Into<At<'a>>, change: &FontChange) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        if change.is_empty() {
            return Ok(self);
        }
        let previous = self
            .edit
            .sheets
            .get(&self.position)
            .and_then(|actions| actions.fonts.get(&address))
            .cloned();
        let mut merged = previous.clone().unwrap_or_default();
        merged.merge(change);
        if self.edit.capture_budgeted()
            && font_change_matches_source(
                &merged,
                self.effective_typed_format(address)?
                    .as_ref()
                    .and_then(Format::font),
            )
        {
            self.edit
                .sheets
                .entry(self.position)
                .or_default()
                .fonts
                .remove(&address);
            self.edit
                .release_cell_capture_facet(self.position, address, CaptureFacet::Font);
            return Ok(self);
        }
        let capture = self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::Font,
            merged.retained_bytes(),
        )?;
        if previous.as_ref() == Some(&merged) {
            self.edit.rollback_cell_capture(capture);
            return Ok(self);
        }
        self.promote_remove_for_typed(address);
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .fonts
            .insert(address, merged);
        Ok(self)
    }

    /// Stage typed cell-alignment properties for one cell.
    ///
    /// The request is resolved against the workbook's shared stylesheet at
    /// commit. Font, alignment, and border requests for one address are
    /// composed into one native cell style, while unrelated properties remain
    /// independently joinable.
    pub fn alignment<'a>(
        &mut self,
        at: impl Into<At<'a>>,
        change: &AlignmentChange,
    ) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        if change.is_empty() {
            return Ok(self);
        }
        let previous = self
            .edit
            .sheets
            .get(&self.position)
            .and_then(|actions| actions.alignments.get(&address))
            .copied();
        let mut merged = previous.unwrap_or_default();
        merged.merge(change);
        if self.edit.capture_budgeted()
            && alignment_change_matches_source(
                &merged,
                self.effective_typed_format(address)?
                    .as_ref()
                    .and_then(Format::alignment),
            )
        {
            self.edit
                .sheets
                .entry(self.position)
                .or_default()
                .alignments
                .remove(&address);
            self.edit
                .release_cell_capture_facet(self.position, address, CaptureFacet::Alignment);
            return Ok(self);
        }
        let capture = self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::Alignment,
            merged.retained_bytes(),
        )?;
        if previous.as_ref() == Some(&merged) {
            self.edit.rollback_cell_capture(capture);
            return Ok(self);
        }
        self.promote_remove_for_typed(address);
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .alignments
            .insert(address, merged);
        Ok(self)
    }

    /// Stage typed cell-border properties for one cell.
    ///
    /// The request is resolved against the workbook's shared stylesheet at
    /// commit. Font, alignment, fill, border, protection, quote-prefix, and
    /// number-format requests for one address are composed into one native
    /// cell style, while unrelated properties remain independently joinable.
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_xlsx::{
    ///     style::{cell_format::{Line, Side}, BorderChange},
    ///     Package,
    /// };
    ///
    /// let workbook = Package::new()?;
    /// let mut edit = workbook.edit()?;
    /// let change = BorderChange::new().left(Some(Side::new(Line::Thin)));
    /// edit.sheet("Sheet1")?.expect("initial worksheet")
    ///     .border("A1", &change)?;
    /// # Ok::<(), litchi_xlsx::Error>(())
    /// ```
    pub fn border<'a>(
        &mut self,
        at: impl Into<At<'a>>,
        change: &BorderChange,
    ) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        if change.is_empty() {
            return Ok(self);
        }
        let previous = self
            .edit
            .sheets
            .get(&self.position)
            .and_then(|actions| actions.borders.get(&address))
            .cloned();
        let mut merged = previous.clone().unwrap_or_else(BorderChange::new);
        merged.merge(change);
        if self.edit.capture_budgeted()
            && border_change_matches_source(
                &merged,
                self.effective_typed_format(address)?
                    .as_ref()
                    .and_then(Format::border),
            )
        {
            self.edit.discard_border(self.position, address)?;
            self.edit
                .release_cell_capture_facet(self.position, address, CaptureFacet::Border);
            return Ok(self);
        }
        let capture = self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::Border,
            merged.retained_bytes(),
        )?;
        if previous.as_ref() == Some(&merged) {
            self.edit.rollback_cell_capture(capture);
            return Ok(self);
        }
        if let Err(error) = self.edit.stage_border(self.position, address, change) {
            self.edit.rollback_cell_capture(capture);
            return Err(error);
        }
        self.promote_remove_for_typed(address);
        Ok(self)
    }

    /// Stage typed cell-fill properties for one cell.
    ///
    /// The request is resolved against the workbook's shared stylesheet at
    /// commit. Font, alignment, fill, protection, quote-prefix, and
    /// number-format requests for one address are composed into one native
    /// cell style, while unrelated properties remain independently joinable.
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_xlsx::{style::{cell_format::Pattern, FillChange}, Package};
    ///
    /// let workbook = Package::new()?;
    /// let mut edit = workbook.edit()?;
    /// let change = FillChange::new().pattern(Some(Pattern::Solid));
    /// edit.sheet("Sheet1")?.expect("initial worksheet")
    ///     .fill("A1", &change)?;
    /// # Ok::<(), litchi_xlsx::Error>(())
    /// ```
    ///
    /// A complete gradient fill can be supplied as a validated replacement:
    ///
    /// ```
    /// use litchi_xlsx::{
    ///     style::{cell_format::{Color, Fill, GradientFill, GradientStop}, FillChange},
    ///     Package,
    /// };
    ///
    /// let workbook = Package::new()?;
    /// let mut edit = workbook.edit()?;
    /// let gradient = GradientFill::linear(
    ///     0.0,
    ///     vec![
    ///         GradientStop::new(0.0, Color::rgb(0xFF, 0x00, 0x00))?,
    ///         GradientStop::new(1.0, Color::rgb(0x00, 0x00, 0xFF))?,
    ///     ]
    ///     .into_boxed_slice(),
    /// )?;
    /// let change = FillChange::replace(Fill::Gradient(gradient));
    /// edit.sheet("Sheet1")?.expect("initial worksheet")
    ///     .fill("A1", &change)?;
    /// # Ok::<(), litchi_xlsx::Error>(())
    /// ```
    pub fn fill<'a>(&mut self, at: impl Into<At<'a>>, change: &FillChange) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        if let Some(replacement) = change.replacement() {
            replacement.validate()?;
        }
        if change.is_empty() {
            return Ok(self);
        }
        let previous = self
            .edit
            .sheets
            .get(&self.position)
            .and_then(|actions| actions.fills.get(&address))
            .cloned();
        let mut merged = previous.clone().unwrap_or_else(FillChange::new);
        merged.merge(change);
        if self.edit.capture_budgeted()
            && fill_change_matches_source(
                &merged,
                self.effective_typed_format(address)?
                    .as_ref()
                    .and_then(Format::fill),
            )
        {
            self.edit.discard_fill(self.position, address)?;
            self.edit
                .release_cell_capture_facet(self.position, address, CaptureFacet::Fill);
            return Ok(self);
        }
        let capture = self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::Fill,
            merged.retained_bytes(),
        )?;
        if previous.as_ref() == Some(&merged) {
            self.edit.rollback_cell_capture(capture);
            return Ok(self);
        }
        if let Err(error) = self.edit.stage_fill(self.position, address, change) {
            self.edit.rollback_cell_capture(capture);
            return Err(error);
        }
        self.promote_remove_for_typed(address);
        Ok(self)
    }

    /// Stage typed cell-protection properties for one cell.
    ///
    /// The request is resolved against the workbook's shared stylesheet at
    /// commit and composed with any pending font or alignment request for the
    /// same address.
    pub fn protection<'a>(
        &mut self,
        at: impl Into<At<'a>>,
        change: &ProtectionChange,
    ) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        if change.is_empty() {
            return Ok(self);
        }
        let previous = self
            .edit
            .sheets
            .get(&self.position)
            .and_then(|actions| actions.protections.get(&address))
            .copied();
        let mut merged = previous.unwrap_or_default();
        merged.merge(change);
        if self.edit.capture_budgeted()
            && protection_change_matches_source(
                &merged,
                self.effective_typed_format(address)?
                    .as_ref()
                    .and_then(Format::protection),
            )
        {
            self.edit
                .sheets
                .entry(self.position)
                .or_default()
                .protections
                .remove(&address);
            self.edit
                .release_cell_capture_facet(self.position, address, CaptureFacet::Protection);
            return Ok(self);
        }
        let capture = self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::Protection,
            merged.retained_bytes(),
        )?;
        if previous.as_ref() == Some(&merged) {
            self.edit.rollback_cell_capture(capture);
            return Ok(self);
        }
        self.promote_remove_for_typed(address);
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .protections
            .insert(address, merged);
        Ok(self)
    }

    /// Stage the authored `quotePrefix` value for one cell. `None` removes an
    /// authored attribute; `Some(false)` remains distinct from removal.
    pub fn quote_prefix<'a>(
        &mut self,
        at: impl Into<At<'a>>,
        value: Option<bool>,
    ) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        if self
            .edit
            .sheets
            .get(&self.position)
            .and_then(|actions| actions.quote_prefixes.get(&address))
            .is_some_and(|current| *current == value)
        {
            return Ok(self);
        }
        if self.edit.capture_budgeted()
            && quote_prefix_matches_source(
                value,
                self.effective_typed_format(address)?
                    .as_ref()
                    .map(Format::quote_prefix),
            )
        {
            self.edit
                .sheets
                .entry(self.position)
                .or_default()
                .quote_prefixes
                .remove(&address);
            self.edit
                .release_cell_capture_facet(self.position, address, CaptureFacet::QuotePrefix);
            return Ok(self);
        }
        let _capture = self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::QuotePrefix,
            size_of::<Option<bool>>(),
        )?;
        self.promote_remove_for_typed(address);
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .quote_prefixes
            .insert(address, value);
        Ok(self)
    }

    /// Stage an authored number-format code for one cell. `None` removes the
    /// local number-format attribute; a code is copied into transaction-owned
    /// storage after the native authoring validation succeeds.
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_xlsx::Package;
    ///
    /// let workbook = Package::new()?;
    /// let mut edit = workbook.edit()?;
    /// let mut sheet = edit.sheet("Sheet1")?.expect("the minimal workbook has Sheet1");
    /// sheet.number_format("A1", Some("0.00"))?;
    /// # Ok::<(), litchi_xlsx::Error>(())
    /// ```
    pub fn number_format<'a>(
        &mut self,
        at: impl Into<At<'a>>,
        code: Option<&str>,
    ) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        let code = code
            .map(crate::style::validate_number_format_code)
            .transpose()?;
        if self
            .edit
            .sheets
            .get(&self.position)
            .and_then(|actions| actions.number_formats.get(&address))
            .is_some_and(|current| *current == code)
        {
            return Ok(self);
        }
        if self.edit.capture_budgeted()
            && number_format_matches_source(
                code.as_deref(),
                self.effective_typed_format(address)?
                    .as_ref()
                    .and_then(|format| format.number_format().code()),
            )
        {
            self.edit
                .sheets
                .entry(self.position)
                .or_default()
                .number_formats
                .remove(&address);
            self.edit.release_cell_capture_facet(
                self.position,
                address,
                CaptureFacet::NumberFormat,
            );
            return Ok(self);
        }
        let _capture = self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::NumberFormat,
            size_of::<Option<Arc<str>>>()
                .saturating_add(code.as_ref().map_or(0, |value| value.len())),
        )?;
        self.promote_remove_for_typed(address);
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .number_formats
            .insert(address, code);
        Ok(self)
    }

    /// Remove the cell record without shifting surrounding cells.
    pub fn remove<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        let source_exists = if self.edit.capture_budgeted() {
            self.source_cell(address)?.is_some()
        } else {
            true
        };
        let pending_remove = self
            .edit
            .sheets
            .get(&self.position)
            .and_then(|actions| actions.cells.get(&address))
            .is_some_and(|action| matches!(action, Action::Remove));
        if pending_remove || (!source_exists && !self.has_pending_cell_effect(address)) {
            return Ok(self);
        }
        if !source_exists {
            (|| -> Result<()> {
                self.edit.discard_fill(self.position, address)?;
                self.edit.discard_border(self.position, address)?;
                Ok(())
            })()?;
            self.edit
                .sheets
                .entry(self.position)
                .or_default()
                .cells
                .remove(&address);
            self.edit
                .sheets
                .entry(self.position)
                .or_default()
                .fonts
                .remove(&address);
            self.edit
                .sheets
                .entry(self.position)
                .or_default()
                .alignments
                .remove(&address);
            self.edit
                .sheets
                .entry(self.position)
                .or_default()
                .protections
                .remove(&address);
            self.edit
                .sheets
                .entry(self.position)
                .or_default()
                .quote_prefixes
                .remove(&address);
            self.edit
                .sheets
                .entry(self.position)
                .or_default()
                .number_formats
                .remove(&address);
            self.edit.release_cell_capture_all(self.position, address);
            return Ok(self);
        }
        let capture = self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::Remove,
            size_of::<Action>(),
        )?;
        if let Err(error) = self.edit.discard_fill(self.position, address) {
            self.edit.rollback_cell_capture(capture);
            return Err(error);
        }
        if let Err(error) = self.edit.discard_border(self.position, address) {
            self.edit.rollback_cell_capture(capture);
            return Err(error);
        }
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .fonts
            .remove(&address);
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .alignments
            .remove(&address);
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .protections
            .remove(&address);
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .quote_prefixes
            .remove(&address);
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .number_formats
            .remove(&address);
        self.edit
            .actions(self.position)
            .insert(address, Action::Remove);
        Ok(self)
    }
}

fn requested_matches<T: PartialEq>(requested: Option<Option<T>>, current: Option<T>) -> bool {
    requested.is_none_or(|requested| requested == current)
}

fn font_color_matches(requested: &crate::style::FontColor, current: cell_format::Color) -> bool {
    let kind_matches = match (requested.kind(), current) {
        (crate::style::FontColorKind::Default, cell_format::Color::Default { .. }) => true,
        (crate::style::FontColorKind::Rgb(requested), cell_format::Color::Rgb { value, .. }) => {
            requested == value
        },
        (
            crate::style::FontColorKind::Theme(requested),
            cell_format::Color::Theme { index, .. },
        ) => requested == index,
        (
            crate::style::FontColorKind::Indexed(requested),
            cell_format::Color::Indexed { index, .. },
        ) => requested == index,
        (
            crate::style::FontColorKind::Auto(requested),
            cell_format::Color::Auto { enabled, .. },
        ) => requested == enabled,
        _ => false,
    };
    kind_matches && requested.tint() == current.tint()
}

pub(super) fn font_change_matches_source(
    change: &FontChange,
    source: Option<&cell_format::Font>,
) -> bool {
    let name = source.and_then(cell_format::Font::name);
    let color = source.and_then(cell_format::Font::color);
    let requested_color_matches = change
        .color_change()
        .is_none_or(|requested| match requested {
            None => color.is_none(),
            Some(requested) => color.is_some_and(|current| font_color_matches(requested, current)),
        });
    requested_matches(
        change.bold_change(),
        source.and_then(cell_format::Font::bold),
    ) && requested_matches(
        change.italic_change(),
        source.and_then(cell_format::Font::italic),
    ) && requested_matches(
        change.strike_change(),
        source.and_then(cell_format::Font::strike),
    ) && requested_matches(
        change.condense_change(),
        source.and_then(cell_format::Font::condense),
    ) && requested_matches(
        change.extend_change(),
        source.and_then(cell_format::Font::extend),
    ) && requested_matches(
        change.outline_change(),
        source.and_then(cell_format::Font::outline),
    ) && requested_matches(
        change.shadow_change(),
        source.and_then(cell_format::Font::shadow),
    ) && requested_matches(
        change.underline_change(),
        source.and_then(cell_format::Font::underline),
    ) && requested_color_matches
        && change
            .name_change()
            .is_none_or(|requested| requested == name)
        && requested_matches(
            change.size_change(),
            source.and_then(cell_format::Font::size),
        )
        && requested_matches(
            change.script_change(),
            source.and_then(cell_format::Font::script),
        )
        && requested_matches(
            change.family_change(),
            source.and_then(cell_format::Font::family),
        )
        && requested_matches(
            change.charset_change(),
            source.and_then(cell_format::Font::charset),
        )
        && requested_matches(
            change.scheme_change(),
            source.and_then(cell_format::Font::scheme),
        )
}

pub(super) fn alignment_change_matches_source(
    change: &AlignmentChange,
    source: Option<&cell_format::Alignment>,
) -> bool {
    requested_matches(
        change.horizontal_change(),
        source.and_then(|alignment| alignment.horizontal),
    ) && requested_matches(
        change.vertical_change(),
        source.and_then(|alignment| alignment.vertical),
    ) && requested_matches(
        change.text_rotation_change(),
        source.and_then(|alignment| alignment.text_rotation),
    ) && authored_alignment_bool_matches(change.wrap_text_change(), source, |alignment| {
        alignment.wrap_text
    }) && requested_matches(
        change.indent_change(),
        source.and_then(|alignment| alignment.indent),
    ) && requested_matches(
        change.relative_indent_change(),
        source.and_then(|alignment| alignment.relative_indent),
    ) && requested_matches(
        change.justify_last_line_change(),
        source.and_then(|alignment| alignment.justify_last_line),
    ) && authored_alignment_bool_matches(change.shrink_to_fit_change(), source, |alignment| {
        alignment.shrink_to_fit
    }) && requested_matches(
        change.reading_order_change(),
        source.and_then(|alignment| alignment.reading_order),
    )
}

fn authored_alignment_bool_matches(
    requested: Option<Option<bool>>,
    source: Option<&cell_format::Alignment>,
    value: impl Fn(&cell_format::Alignment) -> bool,
) -> bool {
    match requested {
        None => true,
        // The semantic format exposes boolean values but not whether a false
        // flag was authored.  A removal against an existing alignment record
        // must therefore remain an authored operation; otherwise an explicit
        // `wrapText="0"`/`shrinkToFit="0"` is silently retained.
        Some(None) => source.is_none(),
        Some(Some(true)) => source.is_some_and(value),
        // `Format` exposes the value of a boolean, not whether its XML
        // attribute was absent or explicitly false. Keep an authored false
        // request so source fidelity is never collapsed by ambiguity.
        Some(Some(false)) => false,
    }
}

fn source_border_side(
    source: Option<&cell_format::Border>,
    edge: BorderEdge,
) -> Option<&cell_format::Side> {
    let source = source?;
    match edge {
        BorderEdge::Start => source.start.as_ref(),
        BorderEdge::End => source.end.as_ref(),
        BorderEdge::Left => source.left.as_ref(),
        BorderEdge::Right => source.right.as_ref(),
        BorderEdge::Top => source.top.as_ref(),
        BorderEdge::Bottom => source.bottom.as_ref(),
        BorderEdge::Diagonal => source
            .diagonal
            .as_ref()
            .and_then(|diagonal| diagonal.side()),
        BorderEdge::Vertical => source.vertical.as_ref(),
        BorderEdge::Horizontal => source.horizontal.as_ref(),
    }
}

fn border_side_matches(
    change: &crate::style::BorderSideChange,
    source: Option<&cell_format::Side>,
) -> bool {
    let whole_matches = change
        .whole_change()
        .is_none_or(|requested| match requested {
            None => source.is_none(),
            Some(requested) => source == Some(requested),
        });
    let line_matches = requested_matches(change.line_change(), source.map(|side| side.line));
    let color_matches =
        requested_matches(change.color_change(), source.and_then(|side| side.color));
    whole_matches && line_matches && color_matches
}

pub(super) fn border_change_matches_source(
    change: &BorderChange,
    source: Option<&cell_format::Border>,
) -> bool {
    [
        BorderEdge::Start,
        BorderEdge::End,
        BorderEdge::Left,
        BorderEdge::Right,
        BorderEdge::Top,
        BorderEdge::Bottom,
        BorderEdge::Diagonal,
        BorderEdge::Vertical,
        BorderEdge::Horizontal,
    ]
    .into_iter()
    .all(|edge| {
        change
            .side_change(edge)
            .is_none_or(|side| border_side_matches(side, source_border_side(source, edge)))
    }) && requested_matches(
        change.diagonal_up_change(),
        source
            .and_then(|border| border.diagonal.as_ref().and_then(|diagonal| diagonal.dir()))
            .map(|direction| direction.is_up()),
    ) && requested_matches(
        change.diagonal_down_change(),
        source
            .and_then(|border| border.diagonal.as_ref().and_then(|diagonal| diagonal.dir()))
            .map(|direction| direction.is_down()),
    ) && requested_matches(
        change.outline_change(),
        source.and_then(|border| border.outline),
    )
}

pub(super) fn fill_change_matches_source(
    change: &FillChange,
    source: Option<&cell_format::Fill>,
) -> bool {
    if let Some(replacement) = change.replacement() {
        return match replacement {
            cell_format::Fill::Pattern(replacement) => {
                let Some(source) = source.and_then(cell_format::Fill::pattern) else {
                    return false;
                };
                let pattern = change.pattern_change().unwrap_or(replacement.pattern());
                let foreground = change
                    .foreground_change()
                    .unwrap_or(replacement.foreground());
                let background = change
                    .background_change()
                    .unwrap_or(replacement.background());
                source.pattern() == pattern
                    && source.foreground() == foreground
                    && source.background() == background
            },
            cell_format::Fill::Gradient(replacement) => {
                // Granular pattern fields do not apply to a gradient.  A
                // replacement followed by one of them must remain a real
                // request rather than being mistaken for the source fill.
                change.pattern_change().is_none()
                    && change.foreground_change().is_none()
                    && change.background_change().is_none()
                    && source
                        .and_then(cell_format::Fill::gradient)
                        .is_some_and(|source| source == replacement)
            },
        };
    }
    let pattern = source
        .and_then(cell_format::Fill::pattern)
        .and_then(|fill| fill.pattern());
    let foreground = source
        .and_then(cell_format::Fill::pattern)
        .and_then(|fill| fill.foreground());
    let background = source
        .and_then(cell_format::Fill::pattern)
        .and_then(|fill| fill.background());
    requested_matches(change.pattern_change(), pattern)
        && requested_matches(change.foreground_change(), foreground)
        && requested_matches(change.background_change(), background)
}

pub(super) fn protection_change_matches_source(
    change: &ProtectionChange,
    source: Option<cell_format::Protection>,
) -> bool {
    requested_matches(
        change.locked_change(),
        source.and_then(|protection| protection.locked()),
    ) && requested_matches(
        change.hidden_change(),
        source.and_then(|protection| protection.hidden()),
    )
}

pub(super) fn quote_prefix_matches_source(requested: Option<bool>, source: Option<bool>) -> bool {
    match requested {
        // `None` removes the authored attribute.  An explicit false remains
        // distinct from absence even though the public format projection
        // reports both as disabled.
        None => source.is_none(),
        Some(true) => source == Some(true),
        // `Format::quote_prefix` reports only an enabled flag, so an absent
        // and an explicit false attribute cannot be distinguished here.
        Some(false) => false,
    }
}

pub(super) fn number_format_matches_source(requested: Option<&str>, source: Option<&str>) -> bool {
    requested == source
}

/// Borrowed editor for one worksheet being created by this transaction.
///
/// The handle owns no native Office identity and cannot outlive its edit.
#[derive(Debug)]
pub struct NewSheet<'a> {
    pub(super) added: &'a mut Added,
    pub(super) active: &'a mut Option<Target>,
    pub(super) style_lineage: &'a Arc<StyleLineage>,
    pub(super) capture_budget: &'a mut Option<CaptureBudget>,
    pub(super) index: usize,
    pub(super) position: usize,
}

impl NewSheet<'_> {
    fn preflight_capture(
        &mut self,
        address: Address,
        facet: CaptureFacet,
        authored_bytes: usize,
    ) -> Result<AddedCaptureReservation> {
        if let Some(budget) = self.capture_budget.as_mut() {
            return budget.charge_new_sheet_cell(self.index, address, facet, authored_bytes);
        }
        Ok(AddedCaptureReservation::inactive())
    }

    fn rollback_capture(&mut self, reservation: AddedCaptureReservation) {
        if let Some(budget) = self.capture_budget.as_mut() {
            budget.rollback_added_cell(reservation);
        }
    }

    fn release_capture(&mut self, address: Address) {
        if let Some(budget) = self.capture_budget.as_mut() {
            budget.release_added_cell(self.index, address);
        }
    }

    /// Current checked zero-based projected position.
    ///
    /// A later tab move or insertion in this transaction can shift it. The
    /// authoritative committed position is recorded by [`Change::Create`].
    #[must_use]
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Current validated developer-facing name in the pending transaction.
    #[must_use]
    pub fn name(&self) -> &str {
        self.added.name.as_str()
    }

    /// Replace the pending name without exposing a physical sheet identity.
    pub fn rename<T>(&mut self, name: T) -> Result<&mut Self>
    where
        T: TryInto<Name>,
        Error: From<T::Error>,
    {
        let name = name.try_into().map_err(Error::from)?;
        if let Some(budget) = self.capture_budget.as_mut() {
            budget.charge_new_sheet(self.index, name.as_str().len())?;
        }
        self.added.name = name;
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
    pub fn defaults(&mut self) -> Defaults<'_> {
        Defaults {
            action: &mut self.added.actions.defaults,
        }
    }

    /// Select one checked row for short property-editing verbs.
    pub fn row(&mut self, at: impl Into<RowAt>) -> Result<Row<'_>> {
        let row = at.into().resolve()?;
        Ok(Row {
            actions: &mut self.added.actions.rows,
            style_lineage: self.style_lineage,
            row,
        })
    }

    /// Select one column by its primary A1 label or a checked zero-based input.
    pub fn column<'b>(&mut self, at: impl Into<ColumnAt<'b>>) -> Result<Column<'_>> {
        let column = at.into().resolve()?;
        Ok(Column {
            actions: &mut self.added.actions.columns,
            style_lineage: self.style_lineage,
            column,
        })
    }

    /// Merge a checked range on this transaction-local worksheet.
    pub fn merge<'b>(&mut self, area: impl Into<Area<'b>>) -> Result<&mut Self> {
        let range = area.into().resolve()?;
        ensure_merge_area(self.added.name.as_str(), range)?;
        let capture = self.preflight_capture(range.start(), CaptureFacet::Structural, 0)?;
        if self
            .added
            .actions
            .merges
            .last()
            .is_some_and(|intent| *intent == MergeIntent::Add(range))
        {
            self.rollback_capture(capture);
            return Ok(self);
        }
        if let Err(source) = self.added.actions.merges.try_reserve(1) {
            self.rollback_capture(capture);
            return Err(allocation("new-sheet merged-range plan", source));
        }
        self.added.actions.merges.push(MergeIntent::Add(range));
        Ok(self)
    }

    /// Unmerge the range containing one cell in this pending worksheet.
    pub fn unmerge<'b>(&mut self, at: impl Into<At<'b>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        let Some(range) = pending_merge(&[], &self.added.actions.merges, address) else {
            return Ok(self);
        };
        let capture = self.preflight_capture(range.start(), CaptureFacet::Structural, 0)?;
        if self
            .added
            .actions
            .merges
            .last()
            .is_some_and(|intent| *intent == MergeIntent::Add(range))
        {
            self.added.actions.merges.pop();
            self.rollback_capture(capture);
            return Ok(self);
        }
        if let Err(source) = self.added.actions.merges.try_reserve(1) {
            self.rollback_capture(capture);
            return Err(allocation("new-sheet merged-range plan", source));
        }
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
        self.preflight_capture(address, CaptureFacet::Payload, content.retained_bytes())?;
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
        self.preflight_capture(address, CaptureFacet::Payload, size_of::<Content>())?;
        let actions = &mut self.added.actions.cells;
        let create = actions.get(&address).is_some_and(|action| {
            matches!(
                action.payload(),
                Some(
                    Payload::Set(_)
                        | Payload::SharedString { .. }
                        | Payload::Rich(_)
                        | Payload::Clear
                )
            )
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
        self.preflight_capture(address, CaptureFacet::Style, size_of::<StyleEffect>())?;
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
        if !self.added.actions.cells.contains_key(&address) {
            return Ok(self);
        }
        let capture =
            self.preflight_capture(address, CaptureFacet::Style, size_of::<StyleEffect>())?;
        match self.added.actions.cells.entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::reset_style());
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_style(StyleEffect::Reset),
        }
        if matches!(
            self.added.actions.cells.get(&address),
            Some(Action::Update {
                payload: None,
                style: Some(StyleEffect::Reset),
            })
        ) {
            self.added.actions.cells.remove(&address);
            self.rollback_capture(capture);
            self.release_capture(address);
        }
        Ok(self)
    }

    /// Remove the pending cell record without shifting surrounding cells.
    pub fn remove<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        if !self.added.actions.cells.contains_key(&address) {
            return Ok(self);
        }
        let capture = self.preflight_capture(address, CaptureFacet::Remove, size_of::<Action>())?;
        if matches!(self.added.actions.cells.get(&address), Some(Action::Remove)) {
            self.rollback_capture(capture);
            return Ok(self);
        }
        self.added.actions.cells.remove(&address);
        self.rollback_capture(capture);
        self.release_capture(address);
        Ok(self)
    }
}
