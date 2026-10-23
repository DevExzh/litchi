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
use crate::style::{AlignmentChange, BorderChange, FillChange, FontChange, ProtectionChange};
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
use super::transaction::{CaptureBudget, CaptureFacet, Edit};

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
        self.edit.preflight_cell_capture_with_facet(
            self.position,
            range.start(),
            CaptureFacet::Structural,
            0,
        )?;
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
        // Admit the worksheet-before image before asking the source store for
        // its merge index. A deferred worksheet store can allocate its full
        // XML payload during that lookup.
        let capture = self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::Structural,
            0,
        )?;
        let range = match (|| -> Result<Option<Rect>> {
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
        })() {
            Ok(range) => range,
            Err(error) => {
                self.edit.rollback_cell_capture(capture);
                return Err(error);
            },
        };
        let Some(range) = range else {
            self.edit.rollback_cell_capture(capture);
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
        self.edit
            .preflight_cell_capture(self.position, address, content.retained_bytes())?;
        match self.edit.actions(self.position).entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::set(content));
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_payload(Payload::Set(content)),
        }
        Ok(self)
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
        if self
            .edit
            .sheets
            .get(&self.position)
            .and_then(|actions| actions.cells.get(&address))
            .is_some_and(|action| action.payload().is_some() || matches!(action, Action::Remove))
        {
            return Err(invalid(
                "a formula cache cannot join another payload change of the same cell",
            ));
        }
        // The capture admission deliberately precedes `store()`: reading the
        // source formula store may decode a deferred worksheet payload.
        let capture =
            self.edit
                .preflight_cell_capture(self.position, address, value.retained_bytes())?;
        let sheet = match self.edit.base.sheet(self.position) {
            Ok(Some(sheet)) => sheet,
            Ok(None) => {
                self.edit.rollback_cell_capture(capture);
                return Err(invalid("worksheet edit target disappeared"));
            },
            Err(error) => {
                self.edit.rollback_cell_capture(capture);
                return Err(error);
            },
        };
        let scalar = match sheet.store() {
            Ok(store) => store.entry(address).is_some_and(|stored| {
                matches!(&stored.cell, Cell::Formula(formula)
                    if matches!(formula.kind(), crate::formula::Kind::Scalar))
            }),
            Err(error) => {
                self.edit.rollback_cell_capture(capture);
                return Err(error);
            },
        };
        if !scalar {
            self.edit.rollback_cell_capture(capture);
            return Err(invalid("a formula cache needs a stored scalar formula"));
        }
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
        self.edit
            .preflight_cell_capture(self.position, address, size_of::<Content>())?;
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
        self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::Style,
            size_of::<StyleEffect>(),
        )?;
        // An explicit shared-style assignment is the later, complete style
        // operation. It supersedes any deferred property-level style request
        // staged for the same cell.
        self.edit.discard_fill(self.position, address)?;
        self.edit.discard_border(self.position, address)?;
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
        self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::Style,
            size_of::<StyleEffect>(),
        )?;
        self.edit.discard_fill(self.position, address)?;
        self.edit.discard_border(self.position, address)?;
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
        self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::Font,
            change.retained_bytes(),
        )?;
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .fonts
            .entry(address)
            .and_modify(|current| current.merge(change))
            .or_insert_with(|| change.clone());
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
        self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::Alignment,
            change.retained_bytes(),
        )?;
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .alignments
            .entry(address)
            .and_modify(|current| current.merge(change))
            .or_insert(*change);
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
        self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::Border,
            change.retained_bytes(),
        )?;
        self.edit.stage_border(self.position, address, change)?;
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
        self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::Fill,
            change.retained_bytes(),
        )?;
        self.edit.stage_fill(self.position, address, change)?;
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
        self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::Protection,
            change.retained_bytes(),
        )?;
        self.edit
            .sheets
            .entry(self.position)
            .or_default()
            .protections
            .entry(address)
            .and_modify(|current| current.merge(change))
            .or_insert(*change);
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
        self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::QuotePrefix,
            size_of::<Option<bool>>(),
        )?;
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
        self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::NumberFormat,
            size_of::<Option<Arc<str>>>()
                .saturating_add(code.as_ref().map_or(0, |value| value.len())),
        )?;
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
        self.edit.preflight_cell_capture_with_facet(
            self.position,
            address,
            CaptureFacet::Remove,
            size_of::<Action>(),
        )?;
        self.edit.discard_fill(self.position, address)?;
        self.edit.discard_border(self.position, address)?;
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
    ) -> Result<()> {
        if let Some(budget) = self.capture_budget.as_mut() {
            budget.charge_new_sheet_cell(self.index, address, facet, authored_bytes)?;
        }
        Ok(())
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
        self.preflight_capture(range.start(), CaptureFacet::Structural, 0)?;
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
        self.preflight_capture(range.start(), CaptureFacet::Structural, 0)?;
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
        self.preflight_capture(address, CaptureFacet::Style, size_of::<StyleEffect>())?;
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
        self.preflight_capture(address, CaptureFacet::Remove, size_of::<Action>())?;
        self.added.actions.cells.insert(address, Action::Remove);
        Ok(self)
    }
}
