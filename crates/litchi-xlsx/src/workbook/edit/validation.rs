//! Transaction state and invariant planning for workbook edits.
//!
//! This layer owns the compact pending-state representation and the checked
//! tab-order projection used by the public semantic facade.

use std::collections::BTreeMap;

use litchi_ooxml_common::web as common_web;
use litchi_sheet::{Cell as Address, Column as ColumnIndex, Rect, Row as RowIndex};

use super::super::Visibility;
use super::model::GraphChange;
use crate::error::{Result, allocation, invalid};
use crate::raw;
use crate::raw::worksheet::edit::{Action, ColumnAction, DefaultsAction, RowAction};
use crate::sheet::Name;
use crate::web::Bindings as WebBindings;

#[derive(Debug, Default)]
pub(super) struct SheetActions {
    pub(super) rename: Option<Name>,
    pub(super) visibility: Option<TabAction>,
    pub(super) defaults: Option<DefaultsAction>,
    pub(super) web: Option<WebBindings>,
    pub(super) cells: BTreeMap<Address, Action>,
    pub(super) rows: BTreeMap<RowIndex, RowAction>,
    pub(super) columns: BTreeMap<ColumnIndex, ColumnAction>,
    pub(super) merges: Vec<MergeIntent>,
    pub(super) page_breaks: Option<crate::page_breaks::PageBreaks>,
    pub(super) page_margins: Option<OptionalAction<crate::page_margins::Margins>>,
    pub(super) page_setup: Option<OptionalAction<crate::page_setup::Setup>>,
    pub(super) print_options: Option<OptionalAction<crate::print_options::PrintOptions>>,
    pub(super) hyperlinks: BTreeMap<(u32, u32, u32, u32), HyperlinkAction>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum HyperlinkAction {
    Put(crate::hyperlinks::Hyperlink),
    Remove(crate::hyperlinks::HyperlinkReference),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum OptionalAction<T> {
    Put(T),
    Remove,
}

impl<T> OptionalAction<T> {
    pub(super) fn from_option(value: Option<T>) -> Self {
        value.map_or(Self::Remove, Self::Put)
    }

    pub(super) const fn as_option(&self) -> Option<&T> {
        match self {
            Self::Put(value) => Some(value),
            Self::Remove => None,
        }
    }
}

impl SheetActions {
    pub(super) fn len(&self) -> usize {
        usize::from(self.rename.is_some())
            .saturating_add(usize::from(self.visibility.is_some()))
            .saturating_add(usize::from(self.defaults.is_some()))
            .saturating_add(usize::from(self.web.is_some()))
            .saturating_add(self.cells.len())
            .saturating_add(self.rows.len())
            .saturating_add(self.columns.len())
            .saturating_add(self.merges.len())
            .saturating_add(usize::from(self.page_breaks.is_some()))
            .saturating_add(usize::from(self.page_margins.is_some()))
            .saturating_add(usize::from(self.page_setup.is_some()))
            .saturating_add(usize::from(self.print_options.is_some()))
            .saturating_add(self.hyperlinks.len())
    }

    pub(super) fn is_empty(&self) -> bool {
        self.rename.is_none()
            && self.visibility.is_none()
            && self.defaults.is_none()
            && self.web.is_none()
            && self.cells.is_empty()
            && self.rows.is_empty()
            && self.columns.is_empty()
            && self.merges.is_empty()
            && self.page_breaks.is_none()
            && self.page_margins.is_none()
            && self.page_setup.is_none()
            && self.print_options.is_none()
            && self.hyperlinks.is_empty()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum MergeIntent {
    Add(Rect),
    Remove(Rect),
}

impl MergeIntent {
    pub(super) const fn range(self) -> Rect {
        match self {
            Self::Add(range) | Self::Remove(range) => range,
        }
    }
}

pub(super) fn pending_merge(
    base: &[Rect],
    intents: &[MergeIntent],
    address: Address,
) -> Option<Rect> {
    let mut current = base.iter().copied().find(|range| range.contains(address));
    for intent in intents {
        match *intent {
            MergeIntent::Add(range) if range.contains(address) => current = Some(range),
            MergeIntent::Remove(range) if current == Some(range) => current = None,
            MergeIntent::Add(_) | MergeIntent::Remove(_) => {},
        }
    }
    current
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum TabAction {
    Show,
    Hide,
    VeryHide,
}

impl TabAction {
    pub(super) const fn visibility(self) -> Visibility {
        match self {
            Self::Show => Visibility::Visible,
            Self::Hide => Visibility::Hidden,
            Self::VeryHide => Visibility::VeryHidden,
        }
    }

    pub(super) const fn raw(self) -> raw::catalog_edit::State {
        match self {
            Self::Show => raw::catalog_edit::State::Visible,
            Self::Hide => raw::catalog_edit::State::Hidden,
            Self::VeryHide => raw::catalog_edit::State::VeryHidden,
        }
    }
}

#[derive(Debug)]
pub(super) struct CreatedSheet {
    pub(super) name: Name,
    pub(super) position: usize,
    pub(super) sheet_id: u32,
    pub(super) relationship_id: String,
    pub(super) visibility: TabAction,
    pub(super) graph: GraphChange,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct MoveIntent {
    pub(super) sheet: usize,
    pub(super) from: usize,
    pub(super) to: usize,
}

#[derive(Debug)]
pub(super) struct OrderPlan {
    pub(super) positions: Vec<usize>,
    pub(super) moves: Vec<MoveIntent>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Target {
    Base(usize),
    Added(usize),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Placement {
    Tail,
    Before(usize),
    After(usize),
}

#[derive(Debug)]
pub(super) struct Added {
    pub(super) name: Name,
    pub(super) actions: SheetActions,
    pub(super) placement: Placement,
}

#[derive(Debug)]
pub(super) enum PanesAction {
    Put {
        panes: common_web::Panes,
        conformance: common_web::Conformance,
    },
    Remove,
}

#[derive(Debug, Default)]
struct Around {
    before: Vec<usize>,
    after: Vec<usize>,
}

#[derive(Debug)]
pub(super) struct FinalOrder {
    pub(super) targets: Vec<Target>,
    pub(super) base_positions: Vec<usize>,
    pub(super) added_positions: Vec<usize>,
}

impl OrderPlan {
    pub(super) fn is_effective(&self) -> bool {
        self.positions
            .iter()
            .copied()
            .enumerate()
            .any(|(position, identity)| position != identity)
    }
}

impl FinalOrder {
    pub(super) fn plan(
        base_len: usize,
        order: Option<&OrderPlan>,
        added: &[Added],
    ) -> Result<Self> {
        let final_len = base_len
            .checked_add(added.len())
            .ok_or_else(|| invalid("final worksheet order length overflow"))?;
        let mut around = BTreeMap::<usize, Around>::new();
        let mut tail = Vec::new();
        for (index, sheet) in added.iter().enumerate() {
            match sheet.placement {
                Placement::Tail => {
                    tail.try_reserve(1)
                        .map_err(|source| allocation("appended worksheet order", source))?;
                    tail.push(index);
                },
                Placement::Before(anchor) | Placement::After(anchor) => {
                    if anchor >= base_len {
                        return Err(invalid("new worksheet anchor is outside the base catalog"));
                    }
                    let slots = around.entry(anchor).or_default();
                    let targets = if matches!(sheet.placement, Placement::Before(_)) {
                        &mut slots.before
                    } else {
                        &mut slots.after
                    };
                    targets
                        .try_reserve(1)
                        .map_err(|source| allocation("anchored worksheet order", source))?;
                    targets.push(index);
                },
            }
        }

        let mut targets = Vec::new();
        targets
            .try_reserve_exact(final_len)
            .map_err(|source| allocation("final worksheet order", source))?;
        let mut push_base = |identity: usize| -> Result<()> {
            if identity >= base_len {
                return Err(invalid("base worksheet order contains an unknown identity"));
            }
            let slots = around.remove(&identity).unwrap_or_default();
            targets.extend(slots.before.into_iter().map(Target::Added));
            targets.push(Target::Base(identity));
            targets.extend(slots.after.into_iter().map(Target::Added));
            Ok(())
        };
        if let Some(order) = order {
            for identity in order.positions.iter().copied() {
                push_base(identity)?;
            }
        } else {
            for identity in 0..base_len {
                push_base(identity)?;
            }
        }
        if !around.is_empty() {
            return Err(invalid(
                "new worksheet anchor disappeared from the final base order",
            ));
        }
        targets.extend(tail.into_iter().map(Target::Added));
        if targets.len() != final_len {
            return Err(invalid("final worksheet order has the wrong length"));
        }

        let mut base_positions = Vec::new();
        base_positions
            .try_reserve_exact(base_len)
            .map_err(|source| allocation("base tab positions", source))?;
        base_positions.resize(base_len, usize::MAX);
        let mut added_positions = Vec::new();
        added_positions
            .try_reserve_exact(added.len())
            .map_err(|source| allocation("new tab positions", source))?;
        added_positions.resize(added.len(), usize::MAX);
        for (position, target) in targets.iter().copied().enumerate() {
            let slot = match target {
                Target::Base(identity) => base_positions.get_mut(identity),
                Target::Added(index) => added_positions.get_mut(index),
            }
            .ok_or_else(|| invalid("final worksheet order contains an unknown target"))?;
            if *slot != usize::MAX {
                return Err(invalid("final worksheet order repeats a target"));
            }
            *slot = position;
        }
        if base_positions.contains(&usize::MAX) || added_positions.contains(&usize::MAX) {
            return Err(invalid("final worksheet order omits a target"));
        }
        Ok(Self {
            targets,
            base_positions,
            added_positions,
        })
    }

    pub(super) fn position(&self, target: Target) -> Option<usize> {
        match target {
            Target::Base(identity) => self.base_positions.get(identity),
            Target::Added(index) => self.added_positions.get(index),
        }
        .copied()
    }

    pub(super) fn target(&self, position: usize) -> Option<Target> {
        self.targets.get(position).copied()
    }

    pub(super) fn len(&self) -> usize {
        self.targets.len()
    }

    pub(super) fn matches_appended(&self, base_order: Option<&OrderPlan>) -> bool {
        let base_len = self.base_positions.len();
        self.targets
            .iter()
            .copied()
            .enumerate()
            .all(|(position, actual)| {
                let expected = if position < base_len {
                    let Some(identity) = base_order.map_or(Some(position), |order| {
                        order.positions.get(position).copied()
                    }) else {
                        return false;
                    };
                    Target::Base(identity)
                } else {
                    Target::Added(position - base_len)
                };
                actual == expected
            })
    }
}
