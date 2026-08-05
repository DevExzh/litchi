//! Layered, lossless worksheet XML edits.
//!
//! The facade keeps the workbook snapshot editor's crate-visible surface
//! stable while the implementation is divided by semantic ownership:
//! selector/edit models, XML codecs, validation, and package orchestration.
mod codec;
mod model;
mod package;
mod validation;

#[cfg(test)]
mod tests;

#[allow(
    unused_imports,
    reason = "the raw edit facade retains the complete crate-visible effect vocabulary"
)]
pub(crate) use model::{
    Action, ColumnAction, DefaultsAction, DefaultsEffects, DescentEffect, HeightEffect, MergePlan,
    OptionalEffect, Payload, Plan, RowAction, StyleEffect, WidthEffect,
};
pub(crate) use package::{rewrite, rewrite_merges};
