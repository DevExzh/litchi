//! Transaction-scoped row, column, and worksheet-default editors.

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

/// Transaction-scoped worksheet-grid-default editor.
#[derive(Debug)]
pub struct DefaultsEdit<'a> {
    pub(super) action: &'a mut Option<DefaultsAction>,
}

impl DefaultsEdit<'_> {
    fn action(&mut self) -> &mut DefaultsAction {
        self.action.get_or_insert_default()
    }

    /// Set the base width in whole Normal-style characters (`0..=255`).
    pub fn base_width(&mut self, width: u8) -> &mut Self {
        self.action().update().base_width = Some(OptionalEffect::Set(width));
        self
    }

    /// Restore the implicit base width of eight characters.
    pub fn reset_base_width(&mut self) -> &mut Self {
        self.action().update().base_width = Some(OptionalEffect::Reset);
        self
    }

    /// Set a checked default column width in character units.
    pub fn width(&mut self, width: impl Into<layout::WidthAt>) -> Result<&mut Self> {
        let width = width.into().resolve()?;
        self.action().update().width = Some(OptionalEffect::Set(width));
        Ok(self)
    }

    /// Restore font-derived default column-width calculation.
    pub fn reset_width(&mut self) -> &mut Self {
        self.action().update().width = Some(OptionalEffect::Reset);
        self
    }

    /// Set a checked default row height and derive its custom-height marker.
    pub fn height(&mut self, height: impl Into<layout::HeightAt>) -> Result<&mut Self> {
        let height = height.into().resolve()?;
        self.action().update().height = Some(height);
        Ok(self)
    }

    /// Hide rows without an explicit row record by default.
    pub fn hide(&mut self) -> &mut Self {
        self.action().update().hidden = Some(true);
        self
    }

    /// Show implicit rows by removing the hidden-default marker.
    pub fn show(&mut self) -> &mut Self {
        self.action().update().hidden = Some(false);
        self
    }

    /// Request a thick top edge on default rows.
    pub fn thick_top(&mut self) -> &mut Self {
        self.action().update().thick_top = Some(true);
        self
    }

    /// Restore the normal top edge on default rows.
    pub fn normal_top(&mut self) -> &mut Self {
        self.action().update().thick_top = Some(false);
        self
    }

    /// Request a thick bottom edge on default rows.
    pub fn thick_bottom(&mut self) -> &mut Self {
        self.action().update().thick_bottom = Some(true);
        self
    }

    /// Restore the normal bottom edge on default rows.
    pub fn normal_bottom(&mut self) -> &mut Self {
        self.action().update().thick_bottom = Some(false);
        self
    }

    /// Set typographic descent in pixels at 100% worksheet zoom.
    pub fn descent(&mut self, value: impl Into<layout::DescentAt>) -> Result<&mut Self> {
        let value = value.into().resolve()?;
        self.action().update().descent = Some(DescentEffect::Set(value));
        Ok(self)
    }

    /// Remove the explicit typographic descent.
    pub fn reset_descent(&mut self) -> &mut Self {
        self.action().update().descent = Some(DescentEffect::Reset);
        self
    }

    /// Remove the complete stored defaults record.
    ///
    /// A later facet setter on this handle starts a fresh update instead.
    pub fn remove(&mut self) -> &mut Self {
        *self.action = Some(DefaultsAction::remove());
        self
    }
}

/// Transaction-scoped editor for one checked worksheet row.
#[derive(Debug)]
pub struct RowEdit<'a> {
    pub(super) actions: &'a mut BTreeMap<RowIndex, RowAction>,
    pub(super) style_lineage: &'a Arc<StyleLineage>,
    pub(super) row: RowIndex,
}

impl RowEdit<'_> {
    fn action(&mut self) -> &mut RowAction {
        self.actions.entry(self.row).or_default()
    }

    /// Hide this row while preserving all row and cell content.
    pub fn hide(&mut self) -> &mut Self {
        self.action().hidden = Some(true);
        self
    }

    /// Show this row while preserving all row and cell content.
    pub fn show(&mut self) -> &mut Self {
        self.action().hidden = Some(false);
        self
    }

    /// Set a checked row height in points and mark it as explicitly customized.
    ///
    /// Raw finite values in `0..=409` and reusable [`crate::row::Height`]
    /// values are both accepted.
    pub fn height(&mut self, height: impl Into<HeightAt>) -> Result<&mut Self> {
        let height = height.into().resolve()?;
        self.action().height = Some(HeightEffect::Set(height));
        Ok(self)
    }

    /// Remove the explicit height and its derived custom-height marker.
    pub fn reset_height(&mut self) -> &mut Self {
        self.action().height = Some(HeightEffect::Reset);
        self
    }

    /// Set typographic descent in pixels at 100% worksheet zoom.
    pub fn descent(&mut self, value: impl Into<layout::DescentAt>) -> Result<&mut Self> {
        let value = value.into().resolve()?;
        self.action().descent = Some(DescentEffect::Set(value));
        Ok(self)
    }

    /// Remove the row's explicit typographic descent.
    pub fn reset_descent(&mut self) -> &mut Self {
        self.action().descent = Some(DescentEffect::Reset);
        self
    }

    /// Apply an existing shared style as this row's default formatting.
    ///
    /// The handle must belong to the transaction's shared-style lineage.
    /// Cells with an explicit local style continue to take precedence.
    pub fn style(&mut self, style: &Style) -> Result<&mut Self> {
        if !Arc::ptr_eq(self.style_lineage, &style.owner.style_lineage) {
            return Err(Error::ForeignStyle);
        }
        self.action().style = Some(StyleEffect::Set(style.raw()));
        Ok(self)
    }

    /// Remove the row's explicit default style and custom-format marker.
    pub fn reset_style(&mut self) -> &mut Self {
        self.action().style = Some(StyleEffect::Reset);
        self
    }

    /// Set a checked outline level in Office's `0..=7` domain.
    pub fn outline(&mut self, level: impl Into<OutlineAt>) -> Result<&mut Self> {
        let level = level.into().resolve()?;
        self.action().outline = Some(level);
        Ok(self)
    }

    /// Store the affected outline in its collapsed state.
    pub fn collapse(&mut self) -> &mut Self {
        self.action().collapsed = Some(true);
        self
    }

    /// Store the affected outline in its expanded state.
    pub fn expand(&mut self) -> &mut Self {
        self.action().collapsed = Some(false);
        self
    }

    /// Request a thick top edge for this row.
    pub fn thick_top(&mut self) -> &mut Self {
        self.action().thick_top = Some(true);
        self
    }

    /// Restore the normal top edge for this row.
    pub fn normal_top(&mut self) -> &mut Self {
        self.action().thick_top = Some(false);
        self
    }

    /// Request a thick bottom edge for this row.
    pub fn thick_bottom(&mut self) -> &mut Self {
        self.action().thick_bottom = Some(true);
        self
    }

    /// Restore the normal bottom edge for this row.
    pub fn normal_bottom(&mut self) -> &mut Self {
        self.action().thick_bottom = Some(false);
        self
    }

    /// Show phonetic information by default for this row.
    pub fn show_phonetic(&mut self) -> &mut Self {
        self.action().phonetic = Some(true);
        self
    }

    /// Hide phonetic information by default for this row.
    pub fn hide_phonetic(&mut self) -> &mut Self {
        self.action().phonetic = Some(false);
        self
    }
}

/// Transaction-scoped editor for one checked worksheet column.
#[derive(Debug)]
pub struct ColumnEdit<'a> {
    pub(super) actions: &'a mut BTreeMap<ColumnIndex, ColumnAction>,
    pub(super) style_lineage: &'a Arc<StyleLineage>,
    pub(super) column: ColumnIndex,
}

impl ColumnEdit<'_> {
    fn action(&mut self) -> &mut ColumnAction {
        self.actions.entry(self.column).or_default()
    }

    /// Hide this column while preserving its other effective properties and
    /// every cell record.
    pub fn hide(&mut self) -> &mut Self {
        self.action().hidden = Some(true);
        self
    }

    /// Show this column while preserving its other effective properties and
    /// every cell record.
    pub fn show(&mut self) -> &mut Self {
        self.action().hidden = Some(false);
        self
    }

    /// Set a checked `SpreadsheetML` width and mark it as explicitly customized.
    ///
    /// Raw finite values in `0..=255` and reusable [`crate::column::Width`]
    /// values are both accepted.
    pub fn width(&mut self, width: impl Into<WidthAt>) -> Result<&mut Self> {
        let width = width.into().resolve()?;
        self.action().width = Some(WidthEffect::Set(width));
        Ok(self)
    }

    /// Remove the explicit width and its derived custom-width marker.
    pub fn reset_width(&mut self) -> &mut Self {
        self.action().width = Some(WidthEffect::Reset);
        self
    }

    /// Apply an existing shared style as this column's default formatting.
    ///
    /// The handle must belong to the transaction's shared-style lineage.
    /// Cells with an explicit local style continue to take precedence.
    /// When the column has no stored width, stage [`Self::width`] in the same
    /// transaction; commit otherwise rejects the style instead of producing
    /// a zero-width column in Excel.
    pub fn style(&mut self, style: &Style) -> Result<&mut Self> {
        if !Arc::ptr_eq(self.style_lineage, &style.owner.style_lineage) {
            return Err(Error::ForeignStyle);
        }
        self.action().style = Some(StyleEffect::Set(style.raw()));
        Ok(self)
    }

    /// Remove the column's explicit default style without changing its width.
    pub fn reset_style(&mut self) -> &mut Self {
        self.action().style = Some(StyleEffect::Reset);
        self
    }

    /// Mark this column for producer best-fit behavior without measuring text.
    pub fn best_fit(&mut self) -> &mut Self {
        self.action().best_fit = Some(true);
        self
    }

    /// Clear the producer best-fit marker while preserving the stored width.
    pub fn fixed(&mut self) -> &mut Self {
        self.action().best_fit = Some(false);
        self
    }

    /// Set a checked outline level in Office's `0..=7` domain.
    pub fn outline(&mut self, level: impl Into<OutlineAt>) -> Result<&mut Self> {
        let level = level.into().resolve()?;
        self.action().outline = Some(level);
        Ok(self)
    }

    /// Store the affected outline in its collapsed state.
    pub fn collapse(&mut self) -> &mut Self {
        self.action().collapsed = Some(true);
        self
    }

    /// Store the affected outline in its expanded state.
    pub fn expand(&mut self) -> &mut Self {
        self.action().collapsed = Some(false);
        self
    }

    /// Show phonetic information by default for this column.
    pub fn show_phonetic(&mut self) -> &mut Self {
        self.action().phonetic = Some(true);
        self
    }

    /// Hide phonetic information by default for this column.
    pub fn hide_phonetic(&mut self) -> &mut Self {
        self.action().phonetic = Some(false);
        self
    }
}
