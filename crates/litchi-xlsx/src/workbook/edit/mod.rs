//! Isolated worksheet transactions, disjoint joins, and source-checked patches.
//!
//! The parent keeps a private context namespace for the existing codec and
//! package siblings, whose `super::*` imports are intentionally unchanged.
#![allow(
    unused_imports,
    reason = "the edit facade deliberately retains its complete crate-visible vocabulary"
)]

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet, btree_map::Entry};
use std::sync::Arc;

use litchi_ooxml_common::web as common_web;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, Relationship, TargetMode};
use litchi_sheet::{
    Area, At, Cell as Address, Column as ColumnIndex, ColumnAt, Rect, Row as RowIndex, RowAt,
};

use super::{Selector, Visibility, Workbook, Worksheet, WorksheetKind};
use crate::Style;
use crate::cell::{Cell, Content};
use crate::chain;
use crate::column::{OutlineAt, State as ColumnState, WidthAt};
/// The error of the primary transaction family.
///
/// The family currently reports the crate error; its limit failures use
/// [`Error::Limit`] with a [`LimitKind`]. The per-module split of the crate
/// error is staged with later ADR 0034 waves.
pub use crate::error::Error;
use crate::error::{EditBlock, RemoveBlock, Result, TabEditBlock, allocation, invalid};
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

mod codec;
mod drawing_transfer;
mod family;
mod model;
mod package;
mod planning;
mod semantic;
mod svg;
mod svg_lifecycle;
mod transfer;
mod validation;
mod wire;

use validation::{
    Added, CreatedSheet, FinalOrder, MergeIntent, MoveIntent, OptionalAction, OrderPlan,
    PanesAction, Placement, SheetActions, TabAction, Target,
};

#[cfg(test)]
mod tests;

pub use family::{Diagnostics, LimitKind};

/// Bounds for one primary workbook edit.
///
/// The optional retained-byte ceiling is an admission budget for source
/// promotion and staged worksheet captures. It is checked before a native
/// worksheet action allocates its changed store or compaction buffer; the
/// commit still performs exact patch accounting before publication.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct EditOptions {
    max_retained_bytes: Option<usize>,
}

impl EditOptions {
    /// Set the maximum retained bytes admitted by this edit.
    #[must_use]
    pub const fn with_max_retained_bytes(mut self, value: usize) -> Self {
        self.max_retained_bytes = Some(value);
        self
    }

    /// Return the retained-byte ceiling, if one was supplied.
    #[must_use]
    pub const fn max_retained_bytes_value(self) -> Option<usize> {
        self.max_retained_bytes
    }
}

/// Insert a `drawing` element naming `relationship_id` into worksheet XML
/// that has none, at its schema position.
pub(crate) fn attach_worksheet_drawing(xml: &[u8], relationship_id: &str) -> Result<Vec<u8>> {
    drawing_transfer::attach_worksheet(xml, relationship_id)
}
pub(crate) use model::workbook_retained_bytes;
pub use model::{
    ActiveTab, Change, Commit, Conflict, ConflictSet, JoinError, JoinFailure, JoinNormalization,
    PackageChange, Patch, State,
};
pub use planning::{MergeChoice, MergeLimits, ThreeWayPlan};
pub use semantic::{Column, Defaults, Edit, FormControl, NewSheet, Row, Sheet, Tab};
pub use svg::{PictureSelector, SvgInput};
pub use wire::{DurablePatch, SealedPatch};

use model::{
    GraphAction, GraphChange, PartChange, StyleGuard, defaults_after, ensure_merge_area,
    merge_conflicts, project_merges,
};
