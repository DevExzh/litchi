//! Semantic edit snapshots, changes, conflicts, and reversible patches.

use super::codec::{ensure_unsigned, same_part, same_relationship, validate_web_integrity};
use super::{Edit, MergeIntent, SheetActions};

use std::collections::BTreeMap;
use std::fmt;
use std::sync::Arc;

use litchi_ooxml_common::web as common_web;
use litchi_opc::{OpcPackage, PackURI, Part, Relationship};
use litchi_sheet::{Cell as Address, Column as ColumnIndex, Rect, Row as RowIndex};

use super::{Visibility, Workbook};
use crate::cell::{Cell, SharedStringKey, Stored};
use crate::column::{Flags as ColumnFlags, Outline, Props as ColumnProps, State as ColumnState};
use crate::error::{Error, MergeEditBlock, Result, allocation, invalid};
use crate::layout::{self, Defaults};
use crate::raw::worksheet::edit::{
    Action, ColumnAction, DefaultsAction, DescentEffect, HeightEffect, MergePlan, OptionalEffect,
    Payload, RowAction, StyleEffect, WidthEffect,
};
use crate::row::{Flags as RowFlags, Props as RowProps, State as RowState};
use crate::web::Bindings as WebBindings;
use crate::{StyleKey, StyleState};

/// Cell state recorded before or after one semantic change.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum State {
    Missing,
    Cell {
        content: Cell,
        style: StyleState,
        /// Exact workbook shared-string identity, including any rich runs.
        shared_string: Option<SharedStringKey>,
    },
}

impl State {
    pub(super) fn read(value: Option<&Stored>, workbook: &Workbook) -> Self {
        value.map_or(Self::Missing, |stored| Self::Cell {
            content: stored.cell.clone(),
            style: stored.style.map_or(StyleState::Default, |key| {
                StyleState::Shared(StyleKey::new(
                    key,
                    Arc::clone(&workbook.inner.style_lineage),
                ))
            }),
            shared_string: stored.shared_string.map(|index| {
                SharedStringKey::new(index, Arc::clone(&workbook.inner.shared_string_lineage))
            }),
        })
    }

    pub(super) fn after(before: Option<&Stored>, action: &Action, workbook: &Workbook) -> Self {
        let Action::Update { payload, style } = action else {
            return Self::Missing;
        };
        let exists = before.is_some() || action.creates_missing();
        if !exists {
            return Self::Missing;
        }
        let content = match payload {
            Some(Payload::Set(content)) => content.as_cell(),
            Some(Payload::SharedString { text, .. }) => {
                Cell::Value(crate::Value::Text(text.clone()))
            },
            Some(Payload::Clear | Payload::ClearIfPresent) => Cell::Empty,
            None => before.map_or(Cell::Empty, |stored| stored.cell.clone()),
        };
        let style = match style {
            Some(StyleEffect::Set(key)) => StyleState::Shared(StyleKey::new(
                *key,
                Arc::clone(&workbook.inner.style_lineage),
            )),
            Some(StyleEffect::Reset) => StyleState::Default,
            None => before
                .and_then(|stored| stored.style)
                .map_or(StyleState::Default, |key| {
                    StyleState::Shared(StyleKey::new(
                        key,
                        Arc::clone(&workbook.inner.style_lineage),
                    ))
                }),
        };
        let shared_string = match payload {
            Some(Payload::SharedString { index, .. }) => Some(SharedStringKey::new(
                *index,
                Arc::clone(&workbook.inner.shared_string_lineage),
            )),
            Some(Payload::Set(_) | Payload::Clear | Payload::ClearIfPresent) => None,
            None => before.and_then(|stored| stored.shared_string).map(|index| {
                SharedStringKey::new(index, Arc::clone(&workbook.inner.shared_string_lineage))
            }),
        };
        Self::Cell {
            content,
            style,
            shared_string,
        }
    }

    pub(super) fn rebind_style(&mut self, workbook: &Workbook) {
        if let Self::Cell {
            style,
            shared_string,
            ..
        } = self
        {
            style.rebind(&workbook.inner.style_lineage);
            if let Some(shared_string) = shared_string {
                shared_string.rebind(&workbook.inner.shared_string_lineage);
            }
        }
    }

    pub(super) const fn uses_shared_style(&self) -> bool {
        matches!(
            self,
            Self::Cell {
                style: StyleState::Shared(_),
                ..
            }
        )
    }

    pub(super) fn calculation_content(&self) -> Option<&Cell> {
        match self {
            Self::Cell { content, .. } if !matches!(content, Cell::Empty) => Some(content),
            Self::Missing | Self::Cell { .. } => None,
        }
    }
}

/// Semantic active-tab identity recorded in a patch without native Office IDs.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ActiveTab {
    pub(super) name: Box<str>,
    pub(super) position: usize,
}

impl ActiveTab {
    /// Developer-facing sheet name at the source snapshot.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Checked zero-based workbook position in the corresponding patch state.
    #[must_use]
    pub const fn position(&self) -> usize {
        self.position
    }
}

impl ColumnState {
    pub(super) fn read(value: Option<&crate::column::Stored>, workbook: &Workbook) -> Self {
        value.map_or(Self::Missing, |column| {
            Self::Stored(ColumnProps {
                width: column.properties.width,
                style: column.properties.style.map_or(StyleState::Default, |key| {
                    StyleState::Shared(StyleKey::new(
                        key,
                        Arc::clone(&workbook.inner.style_lineage),
                    ))
                }),
                outline: column.properties.outline,
                flags: column.properties.flags,
            })
        })
    }

    pub(super) fn after(
        before: Option<&crate::column::Stored>,
        action: ColumnAction,
        workbook: &Workbook,
    ) -> Self {
        if before.is_none() && !action.materializes() {
            return Self::Missing;
        }
        let mut properties = match Self::read(before, workbook) {
            Self::Missing => ColumnProps {
                width: None,
                style: StyleState::Default,
                outline: Outline::NONE,
                flags: ColumnFlags::empty(),
            },
            Self::Stored(properties) => properties,
        };
        if let Some(hidden) = action.hidden {
            properties.flags.set(ColumnFlags::HIDDEN, hidden);
        }
        if let Some(width) = action.width {
            match width {
                WidthEffect::Set(width) => {
                    properties.width = Some(width);
                    properties.flags.insert(ColumnFlags::CUSTOM_WIDTH);
                },
                WidthEffect::Reset => {
                    properties.width = None;
                    properties.flags.remove(ColumnFlags::CUSTOM_WIDTH);
                },
            }
        }
        if let Some(style) = action.style {
            properties.style = match style {
                StyleEffect::Set(key) => StyleState::Shared(StyleKey::new(
                    key,
                    Arc::clone(&workbook.inner.style_lineage),
                )),
                StyleEffect::Reset => StyleState::Default,
            };
        }
        if let Some(best_fit) = action.best_fit {
            properties.flags.set(ColumnFlags::BEST_FIT, best_fit);
        }
        if let Some(outline) = action.outline {
            properties.outline = outline;
        }
        if let Some(collapsed) = action.collapsed {
            properties.flags.set(ColumnFlags::COLLAPSED, collapsed);
        }
        if let Some(phonetic) = action.phonetic {
            properties.flags.set(ColumnFlags::PHONETIC, phonetic);
        }
        Self::Stored(properties)
    }
}

impl RowState {
    pub(super) fn read(value: Option<&crate::row::Stored>, workbook: &Workbook) -> Self {
        value.map_or(Self::Missing, |row| {
            Self::Stored(RowProps {
                height: row.properties.height,
                descent: row.properties.descent,
                style: row.properties.style.map_or(StyleState::Default, |key| {
                    StyleState::Shared(StyleKey::new(
                        key,
                        Arc::clone(&workbook.inner.style_lineage),
                    ))
                }),
                outline: row.properties.outline,
                flags: row.properties.flags,
            })
        })
    }

    pub(super) fn after(
        before: Option<&crate::row::Stored>,
        action: RowAction,
        workbook: &Workbook,
    ) -> Self {
        if before.is_none() && !action.materializes() {
            return Self::Missing;
        }
        let mut properties = match Self::read(before, workbook) {
            Self::Missing => RowProps {
                height: None,
                descent: None,
                style: StyleState::Default,
                outline: Outline::NONE,
                flags: RowFlags::empty(),
            },
            Self::Stored(properties) => properties,
        };
        if let Some(hidden) = action.hidden {
            properties.flags.set(RowFlags::HIDDEN, hidden);
        }
        if let Some(height) = action.height {
            match height {
                HeightEffect::Set(height) => {
                    properties.height = Some(height);
                    properties.flags.insert(RowFlags::CUSTOM_HEIGHT);
                },
                HeightEffect::Reset => {
                    properties.height = None;
                    properties.flags.remove(RowFlags::CUSTOM_HEIGHT);
                },
            }
        }
        if let Some(descent) = action.descent {
            properties.descent = match descent {
                DescentEffect::Set(value) => Some(value),
                DescentEffect::Reset => None,
            };
        }
        if let Some(style) = action.style {
            match style {
                StyleEffect::Set(key) => {
                    properties.style = StyleState::Shared(StyleKey::new(
                        key,
                        Arc::clone(&workbook.inner.style_lineage),
                    ));
                    properties.flags.insert(RowFlags::CUSTOM_FORMAT);
                },
                StyleEffect::Reset => {
                    properties.style = StyleState::Default;
                    properties.flags.remove(RowFlags::CUSTOM_FORMAT);
                },
            }
        }
        if let Some(outline) = action.outline {
            properties.outline = outline;
        }
        if let Some(collapsed) = action.collapsed {
            properties.flags.set(RowFlags::COLLAPSED, collapsed);
        }
        if let Some(thick_top) = action.thick_top {
            properties.flags.set(RowFlags::THICK_TOP, thick_top);
        }
        if let Some(thick_bottom) = action.thick_bottom {
            properties.flags.set(RowFlags::THICK_BOTTOM, thick_bottom);
        }
        if let Some(phonetic) = action.phonetic {
            properties.flags.set(RowFlags::PHONETIC, phonetic);
        }
        Self::Stored(properties)
    }
}

pub(super) fn defaults_after(
    before: Option<&Defaults>,
    action: DefaultsAction,
) -> std::result::Result<Option<Defaults>, crate::DefaultsEditBlock> {
    if action.is_remove() {
        return Ok(None);
    }
    if before.is_none() && !action.materializes() {
        return Ok(None);
    }
    let effects = action.effects();
    let mut defaults = match before {
        Some(value) => value.clone(),
        None => Defaults {
            base_width: None,
            width: None,
            height: effects
                .height
                .ok_or(crate::DefaultsEditBlock::NeedsHeight)?,
            descent: None,
            row_outline: None,
            column_outline: None,
            flags: layout::Flags::empty(),
            present: layout::Flags::empty(),
        },
    };
    if let Some(effect) = effects.base_width {
        defaults.base_width = match effect {
            OptionalEffect::Set(value) => Some(value),
            OptionalEffect::Reset => None,
        };
    }
    if let Some(effect) = effects.width {
        defaults.width = match effect {
            OptionalEffect::Set(value) => Some(value),
            OptionalEffect::Reset => None,
        };
    }
    if let Some(height) = effects.height {
        defaults.height = height;
        defaults.flags.insert(layout::Flags::CUSTOM_HEIGHT);
        defaults.present.insert(layout::Flags::CUSTOM_HEIGHT);
    }
    for (value, flag) in [
        (effects.hidden, layout::Flags::HIDDEN),
        (effects.thick_top, layout::Flags::THICK_TOP),
        (effects.thick_bottom, layout::Flags::THICK_BOTTOM),
    ] {
        if let Some(value) = value {
            defaults.flags.set(flag, value);
            defaults.present.set(flag, value);
        }
    }
    if let Some(effect) = effects.descent {
        defaults.descent = match effect {
            DescentEffect::Set(value) => Some(value),
            DescentEffect::Reset => None,
        };
    }
    Ok(Some(defaults))
}

pub(super) fn ensure_merge_area(sheet: &str, range: Rect) -> Result<()> {
    if range.rows() == 1 && range.columns() == 1 {
        return Err(Error::MergeEditBlocked {
            sheet: sheet.to_owned(),
            range,
            reason: MergeEditBlock::SingleCell,
        });
    }
    Ok(())
}

#[derive(Debug)]
pub(super) struct MergeProjection {
    pub(super) plan: MergePlan,
    pub(super) changes: Vec<(Rect, crate::merge::Change)>,
}

fn rect_key(range: Rect) -> (u32, u32, u32, u32) {
    let start = range.start();
    let (end_row, end_column) = range.end();
    (start.row().get(), start.column().get(), end_row, end_column)
}

fn has_content_after(before: Option<&Cell>, action: Option<&Action>) -> bool {
    match action {
        Some(Action::Remove) => false,
        Some(Action::Update {
            payload: Some(Payload::Set(_) | Payload::SharedString { .. }),
            ..
        }) => true,
        Some(Action::Update {
            payload: Some(Payload::Clear | Payload::ClearIfPresent),
            ..
        }) => false,
        Some(Action::Update { payload: None, .. }) | None => {
            before.is_some_and(|cell| !matches!(cell, Cell::Empty))
        },
    }
}

fn follower_content(
    store: Option<&crate::cell::Store>,
    cells: &BTreeMap<Address, Action>,
    range: Rect,
) -> Option<Address> {
    let anchor = range.start();
    let mut blocked = None;
    if let Some(store) = store {
        for (address, cell) in store.cells(range) {
            if address != anchor && has_content_after(Some(cell), cells.get(&address)) {
                blocked = Some(blocked.map_or(address, |current: Address| current.min(address)));
            }
        }
    }
    for (address, action) in cells.range(anchor..) {
        if address.row().get() >= range.end().0 {
            break;
        }
        if *address == anchor || !range.contains(*address) {
            continue;
        }
        if store.is_some_and(|store| store.entry(*address).is_some()) {
            continue;
        }
        if has_content_after(None, Some(action)) {
            blocked = Some(blocked.map_or(*address, |current| current.min(*address)));
        }
    }
    blocked
}

pub(super) fn project_merges(
    sheet: &str,
    store: Option<&crate::cell::Store>,
    intents: Vec<MergeIntent>,
    cells: &BTreeMap<Address, Action>,
) -> Result<MergeProjection> {
    let base = store.map_or(&[][..], crate::cell::Store::merge_ranges);
    let mut projected = Vec::new();
    projected
        .try_reserve_exact(base.len())
        .map_err(|source| allocation("merged-range projection", source))?;
    projected.extend_from_slice(base);
    for intent in intents {
        match intent {
            MergeIntent::Add(range) => {
                ensure_merge_area(sheet, range)?;
                if projected.contains(&range) {
                    continue;
                }
                if let Some(existing) = projected
                    .iter()
                    .copied()
                    .find(|existing| crate::merge::overlaps(*existing, range))
                {
                    return Err(Error::MergeEditBlocked {
                        sheet: sheet.to_owned(),
                        range,
                        reason: MergeEditBlock::Overlap { existing },
                    });
                }
                if let Some(address) = follower_content(store, cells, range) {
                    return Err(Error::MergeEditBlocked {
                        sheet: sheet.to_owned(),
                        range,
                        reason: MergeEditBlock::FollowerContent { address },
                    });
                }
                projected
                    .try_reserve(1)
                    .map_err(|source| allocation("merged-range projection", source))?;
                projected.push(range);
            },
            MergeIntent::Remove(range) => projected.retain(|candidate| *candidate != range),
        }
    }
    let projected = crate::merge::Index::new(projected)?;
    let projected = projected.as_slice();

    let mut remove = Vec::new();
    remove
        .try_reserve_exact(base.len())
        .map_err(|source| allocation("removed merged ranges", source))?;
    remove.extend(
        base.iter()
            .copied()
            .filter(|range| !projected.contains(range)),
    );
    let mut add = Vec::new();
    add.try_reserve_exact(projected.len())
        .map_err(|source| allocation("added merged ranges", source))?;
    add.extend(
        projected
            .iter()
            .copied()
            .filter(|range| !base.contains(range)),
    );
    remove.sort_unstable_by_key(|range| rect_key(*range));
    add.sort_unstable_by_key(|range| rect_key(*range));

    let mut changes = Vec::new();
    changes
        .try_reserve_exact(remove.len().saturating_add(add.len()))
        .map_err(|source| allocation("merged-range changes", source))?;
    changes.extend(
        remove
            .iter()
            .copied()
            .map(|range| (range, crate::merge::Change::Remove)),
    );
    changes.extend(
        add.iter()
            .copied()
            .map(|range| (range, crate::merge::Change::Add)),
    );
    Ok(MergeProjection {
        plan: MergePlan { add, remove },
        changes,
    })
}

fn intersection(left: Rect, right: Rect) -> Option<Rect> {
    if !crate::merge::overlaps(left, right) {
        return None;
    }
    let start_row = left.start().row().get().max(right.start().row().get());
    let start_column = left
        .start()
        .column()
        .get()
        .max(right.start().column().get());
    let (left_end_row, left_end_column) = left.end();
    let (right_end_row, right_end_column) = right.end();
    Rect::at(
        start_row,
        start_column,
        left_end_row.min(right_end_row),
        left_end_column.min(right_end_column),
    )
    .ok()
}

pub(super) fn merge_conflicts(left: &SheetActions, right: &SheetActions) -> Vec<Rect> {
    let mut ranges = Vec::new();
    for left_intent in &left.merges {
        for right_intent in &right.merges {
            if let Some(overlap) = intersection(left_intent.range(), right_intent.range()) {
                ranges.push(overlap);
            }
        }
    }
    for intent in &left.merges {
        if let MergeIntent::Add(range) = intent {
            ranges.extend(
                right
                    .cells
                    .iter()
                    .filter(|(address, action)| {
                        **address != range.start()
                            && range.contains(**address)
                            && has_content_after(None, Some(action))
                    })
                    .map(|(address, _)| Rect::single(*address)),
            );
        }
    }
    for intent in &right.merges {
        if let MergeIntent::Add(range) = intent {
            ranges.extend(
                left.cells
                    .iter()
                    .filter(|(address, action)| {
                        **address != range.start()
                            && range.contains(**address)
                            && has_content_after(None, Some(action))
                    })
                    .map(|(address, _)| Rect::single(*address)),
            );
        }
    }
    ranges.sort_unstable_by_key(|range| rect_key(*range));
    ranges.dedup();
    ranges
}

/// One deterministic semantic change in a reversible patch.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Change {
    /// A worksheet was added at a checked logical position.
    Create {
        sheet: Box<str>,
        position: usize,
        visibility: Visibility,
    },
    /// A worksheet was removed at a checked logical position.
    Remove {
        sheet: Box<str>,
        position: usize,
        visibility: Visibility,
    },
    Rename {
        position: usize,
        before: Box<str>,
        after: Box<str>,
    },
    Move {
        sheet: Box<str>,
        from: usize,
        to: usize,
    },
    Active {
        before: ActiveTab,
        after: ActiveTab,
    },
    Visibility {
        sheet: Box<str>,
        position: usize,
        before: Visibility,
        after: Visibility,
    },
    Defaults {
        sheet: Box<str>,
        before: Option<Defaults>,
        after: Option<Defaults>,
    },
    /// Worksheet Office Add-in range bindings changed as one validated set.
    Web {
        sheet: Box<str>,
        before: WebBindings,
        after: WebBindings,
    },
    Merge {
        sheet: Box<str>,
        range: Rect,
        change: crate::merge::Change,
    },
    Cell {
        sheet: Box<str>,
        address: Address,
        before: State,
        after: State,
    },
    Row {
        sheet: Box<str>,
        row: RowIndex,
        before: RowState,
        after: RowState,
    },
    Column {
        sheet: Box<str>,
        column: ColumnIndex,
        before: ColumnState,
        after: ColumnState,
    },
    /// Direct worksheet row and column page-break state changed.
    PageBreaks {
        sheet: Box<str>,
        before: crate::page_breaks::PageBreaks,
        after: crate::page_breaks::PageBreaks,
    },
    /// Direct worksheet page-margin state changed.
    PageMargins {
        sheet: Box<str>,
        before: Option<crate::page_margins::Margins>,
        after: Option<crate::page_margins::Margins>,
    },
    /// Direct worksheet page-setup state changed.
    PageSetup {
        sheet: Box<str>,
        before: Option<crate::page_setup::Setup>,
        after: Option<crate::page_setup::Setup>,
    },
    /// Direct worksheet print-option state changed.
    PrintOptions {
        sheet: Box<str>,
        before: Option<crate::print_options::PrintOptions>,
        after: Option<crate::print_options::PrintOptions>,
    },
}

impl Change {
    #[must_use]
    pub fn sheet(&self) -> &str {
        match self {
            Self::Create { sheet, .. } | Self::Remove { sheet, .. } => sheet,
            Self::Rename { after, .. } => after,
            Self::Active { after, .. } => after.name(),
            Self::Move { sheet, .. }
            | Self::Visibility { sheet, .. }
            | Self::Defaults { sheet, .. }
            | Self::Web { sheet, .. }
            | Self::Merge { sheet, .. }
            | Self::Cell { sheet, .. }
            | Self::Row { sheet, .. }
            | Self::Column { sheet, .. }
            | Self::PageBreaks { sheet, .. }
            | Self::PageMargins { sheet, .. }
            | Self::PageSetup { sheet, .. }
            | Self::PrintOptions { sheet, .. } => sheet,
        }
    }

    /// Ordered source/destination positions when this is a tab move.
    #[must_use]
    pub fn moved(&self) -> Option<(usize, usize)> {
        match self {
            Self::Move { from, to, .. } => Some((*from, *to)),
            Self::Create { .. }
            | Self::Remove { .. }
            | Self::Rename { .. }
            | Self::Active { .. }
            | Self::Visibility { .. }
            | Self::Defaults { .. }
            | Self::Web { .. }
            | Self::Merge { .. }
            | Self::Cell { .. }
            | Self::Row { .. }
            | Self::Column { .. }
            | Self::PageBreaks { .. }
            | Self::PageMargins { .. }
            | Self::PageSetup { .. }
            | Self::PrintOptions { .. } => None,
        }
    }

    /// Name transition when this is a dependency-aware sheet rename.
    #[must_use]
    pub fn renamed(&self) -> Option<(usize, &str, &str)> {
        match self {
            Self::Rename {
                position,
                before,
                after,
            } => Some((*position, before, after)),
            Self::Create { .. }
            | Self::Remove { .. }
            | Self::Move { .. }
            | Self::Active { .. }
            | Self::Visibility { .. }
            | Self::Defaults { .. }
            | Self::Web { .. }
            | Self::Merge { .. }
            | Self::Cell { .. }
            | Self::Row { .. }
            | Self::Column { .. }
            | Self::PageBreaks { .. }
            | Self::PageMargins { .. }
            | Self::PageSetup { .. }
            | Self::PrintOptions { .. } => None,
        }
    }

    /// Active-tab transition when this is a workbook-view change.
    #[must_use]
    pub fn active(&self) -> Option<(&ActiveTab, &ActiveTab)> {
        match self {
            Self::Active { before, after } => Some((before, after)),
            Self::Create { .. }
            | Self::Remove { .. }
            | Self::Rename { .. }
            | Self::Move { .. }
            | Self::Visibility { .. }
            | Self::Defaults { .. }
            | Self::Web { .. }
            | Self::Merge { .. }
            | Self::Cell { .. }
            | Self::Row { .. }
            | Self::Column { .. }
            | Self::PageBreaks { .. }
            | Self::PageMargins { .. }
            | Self::PageSetup { .. }
            | Self::PrintOptions { .. } => None,
        }
    }

    /// Visibility state tuple when this is a workbook tab change.
    #[must_use]
    pub fn visibility(&self) -> Option<(usize, &Visibility, &Visibility)> {
        match self {
            Self::Visibility {
                position,
                before,
                after,
                ..
            } => Some((*position, before, after)),
            Self::Create { .. }
            | Self::Remove { .. }
            | Self::Rename { .. }
            | Self::Move { .. }
            | Self::Active { .. }
            | Self::Defaults { .. }
            | Self::Web { .. }
            | Self::Merge { .. }
            | Self::Cell { .. }
            | Self::Row { .. }
            | Self::Column { .. }
            | Self::PageBreaks { .. }
            | Self::PageMargins { .. }
            | Self::PageSetup { .. }
            | Self::PrintOptions { .. } => None,
        }
    }

    /// Worksheet-default transition when this changes `sheetFormatPr`.
    #[must_use]
    pub fn defaults(&self) -> Option<(Option<&Defaults>, Option<&Defaults>)> {
        match self {
            Self::Defaults { before, after, .. } => Some((before.as_ref(), after.as_ref())),
            _ => None,
        }
    }

    /// Office Add-in range-binding transition, when applicable.
    #[must_use]
    pub fn web(&self) -> Option<(&WebBindings, &WebBindings)> {
        match self {
            Self::Web { before, after, .. } => Some((before, after)),
            _ => None,
        }
    }

    /// Merged-range membership transition, when applicable.
    #[must_use]
    pub const fn merged(&self) -> Option<(Rect, crate::merge::Change)> {
        match self {
            Self::Merge { range, change, .. } => Some((*range, *change)),
            _ => None,
        }
    }

    /// Cell state tuple when this is an ordinary cell change.
    #[must_use]
    pub fn cell(&self) -> Option<(Address, &State, &State)> {
        match self {
            Self::Cell {
                address,
                before,
                after,
                ..
            } => Some((*address, before, after)),
            Self::Create { .. }
            | Self::Remove { .. }
            | Self::Rename { .. }
            | Self::Move { .. }
            | Self::Active { .. }
            | Self::Visibility { .. }
            | Self::Defaults { .. }
            | Self::Web { .. }
            | Self::Merge { .. }
            | Self::Row { .. }
            | Self::Column { .. }
            | Self::PageBreaks { .. }
            | Self::PageMargins { .. }
            | Self::PageSetup { .. }
            | Self::PrintOptions { .. } => None,
        }
    }

    /// Row state tuple when this is a row-property change.
    #[must_use]
    pub fn row(&self) -> Option<(RowIndex, &RowState, &RowState)> {
        match self {
            Self::Row {
                row, before, after, ..
            } => Some((*row, before, after)),
            Self::Create { .. }
            | Self::Remove { .. }
            | Self::Rename { .. }
            | Self::Move { .. }
            | Self::Active { .. }
            | Self::Visibility { .. }
            | Self::Defaults { .. }
            | Self::Web { .. }
            | Self::Merge { .. }
            | Self::Cell { .. }
            | Self::Column { .. }
            | Self::PageBreaks { .. }
            | Self::PageMargins { .. }
            | Self::PageSetup { .. }
            | Self::PrintOptions { .. } => None,
        }
    }

    /// Column state tuple when this is a column-property change.
    #[must_use]
    pub fn column(&self) -> Option<(ColumnIndex, &ColumnState, &ColumnState)> {
        match self {
            Self::Column {
                column,
                before,
                after,
                ..
            } => Some((*column, before, after)),
            Self::Create { .. }
            | Self::Remove { .. }
            | Self::Rename { .. }
            | Self::Move { .. }
            | Self::Active { .. }
            | Self::Visibility { .. }
            | Self::Defaults { .. }
            | Self::Web { .. }
            | Self::Merge { .. }
            | Self::Cell { .. }
            | Self::Row { .. }
            | Self::PageBreaks { .. }
            | Self::PageMargins { .. }
            | Self::PageSetup { .. }
            | Self::PrintOptions { .. } => None,
        }
    }

    /// Added worksheet identity when this is a structural create.
    #[must_use]
    pub fn created(&self) -> Option<(usize, &str)> {
        match self {
            Self::Create {
                sheet, position, ..
            } => Some((*position, sheet)),
            _ => None,
        }
    }

    /// Removed worksheet identity when this is a structural delete or inverse.
    #[must_use]
    pub fn removed(&self) -> Option<(usize, &str)> {
        match self {
            Self::Remove {
                sheet, position, ..
            } => Some((*position, sheet)),
            _ => None,
        }
    }

    /// Direct page-break transition, when applicable.
    #[must_use]
    pub fn page_breaks(
        &self,
    ) -> Option<(
        &crate::page_breaks::PageBreaks,
        &crate::page_breaks::PageBreaks,
    )> {
        match self {
            Self::PageBreaks { before, after, .. } => Some((before, after)),
            _ => None,
        }
    }

    /// Direct page-margin transition, when applicable.
    #[must_use]
    pub fn page_margins(
        &self,
    ) -> Option<(
        Option<&crate::page_margins::Margins>,
        Option<&crate::page_margins::Margins>,
    )> {
        match self {
            Self::PageMargins { before, after, .. } => Some((before.as_ref(), after.as_ref())),
            _ => None,
        }
    }

    /// Direct page-setup transition, when applicable.
    #[must_use]
    pub fn page_setup(
        &self,
    ) -> Option<(
        Option<&crate::page_setup::Setup>,
        Option<&crate::page_setup::Setup>,
    )> {
        match self {
            Self::PageSetup { before, after, .. } => Some((before.as_ref(), after.as_ref())),
            _ => None,
        }
    }

    /// Direct print-option transition, when applicable.
    #[must_use]
    pub fn print_options(
        &self,
    ) -> Option<(
        Option<&crate::print_options::PrintOptions>,
        Option<&crate::print_options::PrintOptions>,
    )> {
        match self {
            Self::PrintOptions { before, after, .. } => Some((before.as_ref(), after.as_ref())),
            _ => None,
        }
    }

    pub(super) fn inverse(&self) -> Self {
        match self {
            Self::Create {
                sheet,
                position,
                visibility,
            } => Self::Remove {
                sheet: sheet.clone(),
                position: *position,
                visibility: visibility.clone(),
            },
            Self::Remove {
                sheet,
                position,
                visibility,
            } => Self::Create {
                sheet: sheet.clone(),
                position: *position,
                visibility: visibility.clone(),
            },
            Self::Rename {
                position,
                before,
                after,
            } => Self::Rename {
                position: *position,
                before: after.clone(),
                after: before.clone(),
            },
            Self::Move { sheet, from, to } => Self::Move {
                sheet: sheet.clone(),
                from: *to,
                to: *from,
            },
            Self::Active { before, after } => Self::Active {
                before: after.clone(),
                after: before.clone(),
            },
            Self::Visibility {
                sheet,
                position,
                before,
                after,
            } => Self::Visibility {
                sheet: sheet.clone(),
                position: *position,
                before: after.clone(),
                after: before.clone(),
            },
            Self::Defaults {
                sheet,
                before,
                after,
            } => Self::Defaults {
                sheet: sheet.clone(),
                before: after.clone(),
                after: before.clone(),
            },
            Self::Web {
                sheet,
                before,
                after,
            } => Self::Web {
                sheet: sheet.clone(),
                before: after.clone(),
                after: before.clone(),
            },
            Self::Merge {
                sheet,
                range,
                change,
            } => Self::Merge {
                sheet: sheet.clone(),
                range: *range,
                change: change.inverse(),
            },
            Self::Cell {
                sheet,
                address,
                before,
                after,
            } => Self::Cell {
                sheet: sheet.clone(),
                address: *address,
                before: after.clone(),
                after: before.clone(),
            },
            Self::Row {
                sheet,
                row,
                before,
                after,
            } => Self::Row {
                sheet: sheet.clone(),
                row: *row,
                before: after.clone(),
                after: before.clone(),
            },
            Self::Column {
                sheet,
                column,
                before,
                after,
            } => Self::Column {
                sheet: sheet.clone(),
                column: *column,
                before: after.clone(),
                after: before.clone(),
            },
            Self::PageBreaks {
                sheet,
                before,
                after,
            } => Self::PageBreaks {
                sheet: sheet.clone(),
                before: after.clone(),
                after: before.clone(),
            },
            Self::PageMargins {
                sheet,
                before,
                after,
            } => Self::PageMargins {
                sheet: sheet.clone(),
                before: *after,
                after: *before,
            },
            Self::PageSetup {
                sheet,
                before,
                after,
            } => Self::PageSetup {
                sheet: sheet.clone(),
                before: after.clone(),
                after: before.clone(),
            },
            Self::PrintOptions {
                sheet,
                before,
                after,
            } => Self::PrintOptions {
                sheet: sheet.clone(),
                before: *after,
                after: *before,
            },
        }
    }

    pub(super) fn rebind_style(&mut self, workbook: &Workbook) {
        match self {
            Self::Cell { before, after, .. } => {
                before.rebind_style(workbook);
                after.rebind_style(workbook);
            },
            Self::Column { before, after, .. } => {
                before.rebind_style(&workbook.inner.style_lineage);
                after.rebind_style(&workbook.inner.style_lineage);
            },
            Self::Row { before, after, .. } => {
                before.rebind_style(&workbook.inner.style_lineage);
                after.rebind_style(&workbook.inner.style_lineage);
            },
            _ => {},
        }
    }

    pub(super) fn uses_shared_style(&self) -> bool {
        match self {
            Self::Cell { before, after, .. } => {
                before.uses_shared_style() || after.uses_shared_style()
            },
            Self::Column { before, after, .. } => {
                before.uses_shared_style() || after.uses_shared_style()
            },
            Self::Row { before, after, .. } => {
                before.uses_shared_style() || after.uses_shared_style()
            },
            _ => false,
        }
    }
}

/// Overlapping effects on one logical workbook sheet.
#[derive(Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Conflict {
    Remove {
        sheet: Box<str>,
        position: usize,
    },
    Name {
        sheet: Box<str>,
        position: usize,
    },
    Order {
        sheet: Box<str>,
        position: usize,
    },
    Active {
        sheet: Box<str>,
        position: usize,
    },
    Tab {
        sheet: Box<str>,
        position: usize,
    },
    Defaults {
        sheet: Box<str>,
        position: usize,
        fields: layout::Fields,
    },
    Web {
        sheet: Box<str>,
        position: usize,
    },
    Merges {
        sheet: Box<str>,
        position: usize,
        ranges: Box<[Rect]>,
    },
    Cells {
        sheet: Box<str>,
        position: usize,
        addresses: Box<[Address]>,
    },
    Rows {
        sheet: Box<str>,
        position: usize,
        rows: Box<[RowIndex]>,
    },
    Columns {
        sheet: Box<str>,
        position: usize,
        columns: Box<[ColumnIndex]>,
    },
    /// Both branches replace page breaks on the same worksheet.
    PageBreaks {
        sheet: Box<str>,
        position: usize,
    },
    /// Both branches replace page margins on the same worksheet.
    PageMargins {
        sheet: Box<str>,
        position: usize,
    },
    /// Both branches replace page setup on the same worksheet.
    PageSetup {
        sheet: Box<str>,
        position: usize,
    },
    /// Both branches replace print options on the same worksheet.
    PrintOptions {
        sheet: Box<str>,
        position: usize,
    },
}

impl Conflict {
    /// Developer-facing sheet name.
    #[must_use]
    pub fn sheet(&self) -> &str {
        match self {
            Self::Remove { sheet, .. }
            | Self::Name { sheet, .. }
            | Self::Order { sheet, .. }
            | Self::Active { sheet, .. }
            | Self::Tab { sheet, .. }
            | Self::Defaults { sheet, .. }
            | Self::Web { sheet, .. }
            | Self::Merges { sheet, .. }
            | Self::Cells { sheet, .. }
            | Self::Rows { sheet, .. }
            | Self::Columns { sheet, .. }
            | Self::PageBreaks { sheet, .. }
            | Self::PageMargins { sheet, .. }
            | Self::PageSetup { sheet, .. }
            | Self::PrintOptions { sheet, .. } => sheet,
        }
    }

    /// Checked zero-based sheet position in the shared base snapshot.
    #[must_use]
    pub fn position(&self) -> usize {
        match self {
            Self::Remove { position, .. }
            | Self::Name { position, .. }
            | Self::Order { position, .. }
            | Self::Active { position, .. }
            | Self::Tab { position, .. }
            | Self::Defaults { position, .. }
            | Self::Web { position, .. }
            | Self::Merges { position, .. }
            | Self::Cells { position, .. }
            | Self::Rows { position, .. }
            | Self::Columns { position, .. }
            | Self::PageBreaks { position, .. }
            | Self::PageMargins { position, .. }
            | Self::PageSetup { position, .. }
            | Self::PrintOptions { position, .. } => *position,
        }
    }

    /// Whether both edits target this sheet's catalog name.
    #[must_use]
    pub const fn is_name(&self) -> bool {
        matches!(self, Self::Name { .. })
    }

    /// Whether both edits remove or otherwise overlap one removed sheet.
    #[must_use]
    pub const fn is_remove(&self) -> bool {
        matches!(self, Self::Remove { .. })
    }

    /// Whether both edits target this sheet tab's visibility.
    #[must_use]
    pub const fn is_tab(&self) -> bool {
        matches!(self, Self::Tab { .. })
    }

    /// Whether both edits target the workbook's one active-tab facet.
    #[must_use]
    pub const fn is_active(&self) -> bool {
        matches!(self, Self::Active { .. })
    }

    /// Whether both edits target the workbook's one tab-order facet.
    #[must_use]
    pub const fn is_order(&self) -> bool {
        matches!(self, Self::Order { .. })
    }

    /// Overlapping worksheet-default facets, when applicable.
    #[must_use]
    pub const fn defaults(&self) -> Option<layout::Fields> {
        match self {
            Self::Defaults { fields, .. } => Some(*fields),
            _ => None,
        }
    }

    /// Whether both edits replace bindings on the same worksheet.
    #[must_use]
    pub const fn is_web(&self) -> bool {
        matches!(self, Self::Web { .. })
    }

    /// Whether both edits replace page breaks on the same worksheet.
    #[must_use]
    pub const fn is_page_breaks(&self) -> bool {
        matches!(self, Self::PageBreaks { .. })
    }

    /// Whether both edits replace page margins on the same worksheet.
    #[must_use]
    pub const fn is_page_margins(&self) -> bool {
        matches!(self, Self::PageMargins { .. })
    }

    /// Whether both edits replace page setup on the same worksheet.
    #[must_use]
    pub const fn is_page_setup(&self) -> bool {
        matches!(self, Self::PageSetup { .. })
    }

    /// Whether both edits replace print options on the same worksheet.
    #[must_use]
    pub const fn is_print_options(&self) -> bool {
        matches!(self, Self::PrintOptions { .. })
    }

    /// Structurally overlapping merged ranges, when applicable.
    #[must_use]
    pub fn merges(&self) -> Option<&[Rect]> {
        match self {
            Self::Merges { ranges, .. } => Some(ranges),
            _ => None,
        }
    }

    /// Deterministically ordered cells written by both edits, when applicable.
    #[must_use]
    pub fn cells(&self) -> Option<&[Address]> {
        match self {
            Self::Cells { addresses, .. } => Some(addresses),
            Self::Remove { .. }
            | Self::Name { .. }
            | Self::Order { .. }
            | Self::Active { .. }
            | Self::Tab { .. }
            | Self::Defaults { .. }
            | Self::Web { .. }
            | Self::Merges { .. }
            | Self::Rows { .. }
            | Self::Columns { .. }
            | Self::PageBreaks { .. }
            | Self::PageMargins { .. }
            | Self::PageSetup { .. }
            | Self::PrintOptions { .. } => None,
        }
    }

    /// Deterministically ordered rows written by both edits, when applicable.
    #[must_use]
    pub fn rows(&self) -> Option<&[RowIndex]> {
        match self {
            Self::Rows { rows, .. } => Some(rows),
            Self::Remove { .. }
            | Self::Name { .. }
            | Self::Order { .. }
            | Self::Active { .. }
            | Self::Tab { .. }
            | Self::Defaults { .. }
            | Self::Web { .. }
            | Self::Merges { .. }
            | Self::Cells { .. }
            | Self::Columns { .. }
            | Self::PageBreaks { .. }
            | Self::PageMargins { .. }
            | Self::PageSetup { .. }
            | Self::PrintOptions { .. } => None,
        }
    }

    /// Deterministically ordered columns written by both edits, when
    /// applicable.
    #[must_use]
    pub fn columns(&self) -> Option<&[ColumnIndex]> {
        match self {
            Self::Columns { columns, .. } => Some(columns),
            Self::Remove { .. }
            | Self::Name { .. }
            | Self::Order { .. }
            | Self::Active { .. }
            | Self::Tab { .. }
            | Self::Defaults { .. }
            | Self::Web { .. }
            | Self::Merges { .. }
            | Self::Cells { .. }
            | Self::Rows { .. }
            | Self::PageBreaks { .. }
            | Self::PageMargins { .. }
            | Self::PageSetup { .. }
            | Self::PrintOptions { .. } => None,
        }
    }

    fn len(&self) -> usize {
        match self {
            Self::Remove { .. }
            | Self::Name { .. }
            | Self::Order { .. }
            | Self::Active { .. }
            | Self::Tab { .. }
            | Self::Web { .. }
            | Self::PageBreaks { .. }
            | Self::PageMargins { .. }
            | Self::PageSetup { .. }
            | Self::PrintOptions { .. } => 1,
            Self::Defaults { fields, .. } => fields.bits().count_ones() as usize,
            Self::Merges { ranges, .. } => ranges.len(),
            Self::Cells { addresses, .. } => addresses.len(),
            Self::Rows { rows, .. } => rows.len(),
            Self::Columns { columns, .. } => columns.len(),
        }
    }
}

/// Structured overlap report returned by [`Edit::join`].
#[derive(Debug, PartialEq, Eq)]
pub struct ConflictSet {
    pub(super) conflicts: Box<[Conflict]>,
}

impl ConflictSet {
    /// Conflicts in deterministic workbook-effect and sheet order.
    #[must_use]
    pub fn conflicts(&self) -> &[Conflict] {
        &self.conflicts
    }

    /// Number of overlapping effects across the workbook.
    pub fn len(&self) -> usize {
        self.conflicts.iter().map(Conflict::len).sum()
    }

    /// Whether no overlapping effects were found.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.conflicts.is_empty()
    }
}

impl fmt::Display for ConflictSet {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "{} overlapping effect(s) across {} conflict group(s)",
            self.len(),
            self.conflicts.len()
        )
    }
}

/// Why two independently prepared edits could not be joined.
#[derive(Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum JoinFailure {
    /// The edits were not created from the same immutable snapshot lineage.
    DifferentSnapshot,
    /// Both edits write at least one of the same effect facets.
    Overlap(ConflictSet),
    /// Both edits replace the workbook's one persisted task-pane graph.
    TaskPanes,
    /// Both edits replace the complete workbook defined-name catalog.
    DefinedNames,
    /// Both edits stage worksheet drawing graph transfers that cannot be joined.
    DrawingTransfer,
}

/// Recoverable join failure that returns ownership of the rejected edit.
pub struct JoinError {
    pub(super) failure: JoinFailure,
    pub(super) rejected: Box<Edit>,
}

impl fmt::Debug for JoinError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("JoinError")
            .field("failure", &self.failure)
            .field("rejected_effects", &self.rejected.len())
            .finish()
    }
}

impl JoinError {
    /// Structured reason the join was refused.
    #[must_use]
    pub fn failure(&self) -> &JoinFailure {
        &self.failure
    }

    /// Overlapping effects, or `None` for a lineage mismatch.
    #[must_use]
    pub fn conflicts(&self) -> Option<&ConflictSet> {
        match &self.failure {
            JoinFailure::Overlap(conflicts) => Some(conflicts),
            JoinFailure::DifferentSnapshot
            | JoinFailure::TaskPanes
            | JoinFailure::DefinedNames
            | JoinFailure::DrawingTransfer => None,
        }
    }

    /// Borrow the edit that was not merged.
    #[must_use]
    pub fn rejected(&self) -> &Edit {
        &self.rejected
    }

    /// Recover the edit that was not merged.
    #[must_use]
    pub fn into_rejected(self) -> Edit {
        *self.rejected
    }

    /// Recover both the structured reason and rejected edit.
    #[must_use]
    pub fn into_parts(self) -> (JoinFailure, Edit) {
        (self.failure, *self.rejected)
    }
}

impl fmt::Display for JoinError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.failure {
            JoinFailure::DifferentSnapshot => {
                formatter.write_str("edits belong to different workbook snapshots")
            },
            JoinFailure::Overlap(conflicts) => {
                write!(formatter, "edit effects overlap: {conflicts}")
            },
            JoinFailure::TaskPanes => {
                formatter.write_str("both edits replace the persisted task-pane graph")
            },
            JoinFailure::DefinedNames => {
                formatter.write_str("both edits replace the workbook defined-name catalog")
            },
            JoinFailure::DrawingTransfer => {
                formatter.write_str("worksheet drawing graph transfers overlap")
            },
        }
    }
}

impl std::error::Error for JoinError {}

/// One workbook-scoped semantic change that has no worksheet identity.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum PackageChange {
    /// Persisted Office Add-in task panes and their inert package graph changed.
    TaskPanes {
        before: Option<common_web::Panes>,
        after: Option<common_web::Panes>,
    },
    /// Complete inert workbook defined-name catalog changed.
    DefinedNames {
        before: Box<[crate::raw::DefinedName]>,
        after: Box<[crate::raw::DefinedName]>,
    },
    /// Selected worksheet drawing anchors and their image graph were cloned.
    DrawingTransfer {
        source: Box<str>,
        target: Box<str>,
        anchors: usize,
        added: bool,
    },
}

impl PackageChange {
    /// Borrow the task-pane transition represented by this change.
    #[must_use]
    pub fn task_panes(&self) -> (Option<&common_web::Panes>, Option<&common_web::Panes>) {
        match self {
            Self::TaskPanes { before, after } => (before.as_ref(), after.as_ref()),
            Self::DefinedNames { .. } | Self::DrawingTransfer { .. } => (None, None),
        }
    }

    /// Borrow the complete defined-name transition, when applicable.
    #[must_use]
    pub fn defined_names(
        &self,
    ) -> Option<(&[crate::raw::DefinedName], &[crate::raw::DefinedName])> {
        match self {
            Self::DefinedNames { before, after } => Some((before, after)),
            Self::TaskPanes { .. } | Self::DrawingTransfer { .. } => None,
        }
    }

    pub(super) fn inverse(&self) -> Self {
        match self {
            Self::TaskPanes { before, after } => Self::TaskPanes {
                before: after.clone(),
                after: before.clone(),
            },
            Self::DefinedNames { before, after } => Self::DefinedNames {
                before: after.clone(),
                after: before.clone(),
            },
            Self::DrawingTransfer {
                source,
                target,
                anchors,
                added,
            } => Self::DrawingTransfer {
                source: source.clone(),
                target: target.clone(),
                anchors: *anchors,
                added: !added,
            },
        }
    }
}

#[derive(Debug, Clone)]
pub(super) struct PartChange {
    pub(super) uri: PackURI,
    pub(super) before: Arc<Vec<u8>>,
    pub(super) after: Arc<Vec<u8>>,
}

#[derive(Debug, Clone)]
pub(super) struct StyleGuard {
    pub(super) uri: PackURI,
    pub(super) content: Arc<Vec<u8>>,
}

impl StyleGuard {
    fn validate(&self, workbook: &Workbook) -> Result<()> {
        let matches = workbook.inner.styles_uri.as_ref() == Some(&self.uri)
            && workbook
                .inner
                .package
                .get_part(&self.uri)
                .is_ok_and(|part| part.blob() == self.content.as_slice());
        if matches {
            Ok(())
        } else {
            Err(Error::PatchConflict {
                part: self.uri.to_string(),
            })
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum GraphAction {
    Add,
    Remove,
}

#[derive(Clone)]
pub(super) struct GraphChange {
    pub(super) action: GraphAction,
    pub(super) source: PackURI,
    pub(super) relationship: Relationship,
    pub(super) part: Box<dyn Part + Send + Sync>,
}

impl fmt::Debug for GraphChange {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("GraphChange")
            .field("action", &self.action)
            .field("source", &self.source)
            .field("relationship", &self.relationship.r_id())
            .field("part", self.part.partname())
            .finish()
    }
}

/// Versioned reversible patch produced by a successful transaction.
///
/// The in-memory representation keeps semantic public changes free of native
/// Office IDs while private deltas retain exact changed-part bytes and
/// relationship fields. [`Patch::durable`] projects the committed source and
/// target snapshots onto the common bounded deterministic wire.
#[derive(Debug, Clone, Default)]
pub struct Patch {
    pub(super) changes: Box<[Change]>,
    pub(super) package_changes: Box<[PackageChange]>,
    pub(super) parts: Box<[PartChange]>,
    pub(super) graph: Box<[GraphChange]>,
    pub(super) web: Option<common_web::Patch>,
    pub(super) style_guard: Option<StyleGuard>,
    pub(super) source: Option<Workbook>,
    pub(super) target: Option<Workbook>,
}

impl Patch {
    pub const VERSION: u16 = 1;

    #[must_use]
    pub const fn version(&self) -> u16 {
        Self::VERSION
    }

    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Workbook-scoped semantic changes, kept separate from sheet changes.
    #[must_use]
    pub fn package_changes(&self) -> &[PackageChange] {
        &self.package_changes
    }

    #[must_use]
    pub fn len(&self) -> usize {
        self.changes
            .len()
            .saturating_add(self.package_changes.len())
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.changes.is_empty() && self.package_changes.is_empty()
    }

    /// Build the inverse without copying part payloads.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            changes: self
                .changes
                .iter()
                .rev()
                .map(Change::inverse)
                .collect::<Vec<_>>()
                .into_boxed_slice(),
            package_changes: self
                .package_changes
                .iter()
                .rev()
                .map(PackageChange::inverse)
                .collect::<Vec<_>>()
                .into_boxed_slice(),
            parts: self
                .parts
                .iter()
                .rev()
                .map(|part| PartChange {
                    uri: part.uri.clone(),
                    before: Arc::clone(&part.after),
                    after: Arc::clone(&part.before),
                })
                .collect::<Vec<_>>()
                .into_boxed_slice(),
            graph: self
                .graph
                .iter()
                .rev()
                .map(|change| GraphChange {
                    action: match change.action {
                        GraphAction::Add => GraphAction::Remove,
                        GraphAction::Remove => GraphAction::Add,
                    },
                    source: change.source.clone(),
                    relationship: change.relationship.clone(),
                    part: change.part.clone(),
                })
                .collect::<Vec<_>>()
                .into_boxed_slice(),
            web: self.web.as_ref().map(common_web::Patch::inverse),
            style_guard: self.style_guard.clone(),
            source: self.target.clone(),
            target: self.source.clone(),
        }
    }

    /// Convert this exact in-memory patch to the bounded durable wire form.
    ///
    /// # Errors
    ///
    /// Returns an error when this patch was not produced by a transaction or
    /// application commit, a package exceeds finite bounds, or the core wire
    /// envelope cannot be constructed.
    pub fn durable(&self) -> Result<super::DurablePatch> {
        super::DurablePatch::from_patch(self)
    }

    pub(in crate::workbook) fn apply_to(&self, workbook: &Workbook) -> Result<Commit> {
        ensure_unsigned(workbook)?;
        let source = workbook.clone();
        if let Some(guard) = &self.style_guard {
            guard.validate(workbook)?;
        }
        if self.parts.is_empty() && self.graph.is_empty() && self.web.is_none() {
            let mut patch = self.clone();
            patch.source = Some(workbook.clone());
            patch.target = Some(workbook.clone());
            return Ok(Commit {
                workbook: workbook.clone(),
                patch,
            });
        }
        let mut package = workbook.inner.package.clone();
        for change in &self.parts {
            let part = package.get_part(&change.uri)?;
            if part.blob() != change.before.as_slice() {
                return Err(Error::PatchConflict {
                    part: change.uri.to_string(),
                });
            }
        }
        for change in &self.parts {
            package
                .get_part_mut(&change.uri)?
                .set_blob_shared(Arc::clone(&change.after));
        }
        for change in &self.graph {
            change.validate(&package)?;
            change.apply(&mut package)?;
        }
        if let Some(web) = &self.web {
            let _ = web.apply(&mut package)?;
        }
        let workbook = Workbook::from_package_with_styles(package, Some(workbook))?;
        if self.web.is_some()
            || self
                .changes
                .iter()
                .any(|change| matches!(change, Change::Web { .. }))
        {
            validate_web_integrity(&workbook)?;
        }
        let mut patch = self.clone();
        for change in &mut patch.changes {
            change.rebind_style(&workbook);
        }
        patch.source = Some(source);
        patch.target = Some(workbook.clone());
        Ok(Commit { workbook, patch })
    }
}

/// Successful atomic transaction result.
#[derive(Debug)]
pub struct Commit {
    pub(super) workbook: Workbook,
    pub(super) patch: Patch,
}

impl Commit {
    #[must_use]
    pub fn workbook(&self) -> &Workbook {
        &self.workbook
    }

    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    pub fn into_workbook(self) -> Workbook {
        self.workbook
    }

    #[must_use]
    pub fn into_parts(self) -> (Workbook, Patch) {
        (self.workbook, self.patch)
    }

    /// Record this committed snapshot in bounded undo/redo history.
    ///
    /// The retained transition weight is the complete canonical durable JSON
    /// length. The history's current snapshot must satisfy the patch's exact
    /// source precondition.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale history head, durable serialization
    /// failure, or a transition heavier than the history budget.
    pub fn record(&self, history: &mut crate::workbook::History) -> Result<Vec<Workbook>> {
        let durable = self.patch.durable()?;
        let next = durable.apply(history.current())?;
        let weight = u64::try_from(durable.to_deterministic_json()?.len())
            .map_err(|error| invalid(format!("durable patch weight exceeds u64: {error}")))?;
        Ok(history.record(next, weight)?)
    }
}

impl GraphChange {
    pub(super) fn validate(&self, package: &OpcPackage) -> Result<()> {
        let source = package.get_part(&self.source)?;
        match self.action {
            GraphAction::Remove => {
                let relationship =
                    source.rels().get(self.relationship.r_id()).ok_or_else(|| {
                        Error::PatchConflict {
                            part: self.source.to_string(),
                        }
                    })?;
                if !same_relationship(relationship, &self.relationship) {
                    return Err(Error::PatchConflict {
                        part: self.source.to_string(),
                    });
                }
                let part = package.get_part(self.part.partname())?;
                if !same_part(part, &*self.part) {
                    return Err(Error::PatchConflict {
                        part: self.part.partname().to_string(),
                    });
                }
            },
            GraphAction::Add => {
                if source.rels().get(self.relationship.r_id()).is_some()
                    || package
                        .validate_new_part_name(self.part.partname())
                        .is_err()
                {
                    return Err(Error::PatchConflict {
                        part: self.part.partname().to_string(),
                    });
                }
            },
        }
        Ok(())
    }

    pub(super) fn apply(&self, package: &mut OpcPackage) -> Result<()> {
        match self.action {
            GraphAction::Remove => {
                package
                    .get_part_mut(&self.source)?
                    .rels_mut()
                    .remove(self.relationship.r_id())
                    .ok_or_else(|| Error::PatchConflict {
                        part: self.source.to_string(),
                    })?;
                if !package.remove_part(self.part.partname()) {
                    return Err(Error::PatchConflict {
                        part: self.part.partname().to_string(),
                    });
                }
            },
            GraphAction::Add => {
                package.try_add_part(self.part.clone())?;
                package
                    .get_part_mut(&self.source)?
                    .rels_mut()
                    .try_add_relationship(
                        self.relationship.reltype().to_owned(),
                        self.relationship.target_ref().to_owned(),
                        self.relationship.r_id().to_owned(),
                        self.relationship.target_mode(),
                    )?;
            },
        }
        Ok(())
    }
}
