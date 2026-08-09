//! Contextual ODS DataPilot ownership.
//!
//! The lower [`crate::model::data_pilot`] module owns the ODF vocabulary and
//! its standalone XML grammar.  This module binds that vocabulary to the
//! direct `office:spreadsheet` owner and supplies an immutable catalog plus
//! clone-staged package transactions.

mod codec;
mod model;
mod package;
mod snapshot;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{Catalog, Selector};
pub use snapshot::{Commit, Edit, OwnedEditor, Patch, Snapshot};
pub use transaction::{Commit as CatalogCommit, Editor, Transaction};

pub use crate::model::data_pilot::{
    DisplayInfo, DisplayMemberMode, Field, FieldReference, GrandTotal, GrandTotalElement,
    GrandTotalOrientation, Group, GroupBoundary, GroupBy, Groups, LayoutInfo, LayoutMode, Level,
    Member, Orientation, ReferenceMemberType, ReferenceType, SortInfo, SortMode, SortOrder, Source,
    Table,
};
