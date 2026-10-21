//! A typed custom-function edit carried by a durable Web Extensions patch
//! ([`super::WebIntent::CustomFunctions`]).
//!
//! [`Change`] is the bounded semantic request a planner binds to its ordinary
//! source-checked [`super::Patch`]; it is a value, not a staging handle.

use super::{
    BackgroundAppData, ContainsCustomFunctions, CustomFunctionList, CustomFunctions, GraphMode,
    OwnerSelector, Pane,
};

/// One typed custom-function operation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Operation {
    Set(Option<CustomFunctions>),
    SetContains(Option<ContainsCustomFunctions>),
    SetBackground(Option<BackgroundAppData>),
    SetList(Option<CustomFunctionList>),
    InsertId { index: usize, id: String },
    ReplaceId { index: usize, id: String },
    RemoveId { index: usize },
}

/// A bounded semantic custom-function edit. `additions` is only admitted for
/// [`GraphMode::EnsureGraph`]. Complete graph replacement is represented by a
/// [`super::WebIntent::Put`] so no panes or resources can be silently omitted.
#[derive(Debug, Clone, PartialEq)]
pub struct Change {
    pub mode: GraphMode,
    pub operations: Vec<(OwnerSelector, Operation)>,
    pub additions: Vec<Pane>,
}

impl Change {
    #[must_use]
    pub fn existing(operations: Vec<(OwnerSelector, Operation)>) -> Self {
        Self {
            mode: GraphMode::ExistingOnly,
            operations,
            additions: Vec::new(),
        }
    }

    #[must_use]
    pub fn ensure(additions: Vec<Pane>, operations: Vec<(OwnerSelector, Operation)>) -> Self {
        Self {
            mode: GraphMode::EnsureGraph,
            operations,
            additions,
        }
    }
}
