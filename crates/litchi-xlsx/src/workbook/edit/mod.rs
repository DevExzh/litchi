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

mod codec;
mod model;
mod package;
mod semantic;
mod validation;

use validation::{
    Added, CreatedSheet, FinalOrder, MergeIntent, MoveIntent, OrderPlan, PanesAction, Placement,
    SheetActions, TabAction, Target,
};

#[cfg(test)]
mod tests;

pub use model::{
    ActiveTab, Change, Commit, Conflict, ConflictSet, JoinError, JoinFailure, PackageChange, Patch,
    State,
};
pub use semantic::{ColumnEdit, DefaultsEdit, Edit, NewSheet, RowEdit, TabEdit, WorksheetEdit};

use model::{
    GraphAction, GraphChange, PartChange, StyleGuard, defaults_after, ensure_merge_area,
    merge_conflicts, project_merges,
};
