//! Semantic slicer values.
//!
//! These names are intentionally contextual. Package relationship identity is
//! represented only by [`Part`] and [`Cache`] at the package boundary; the
//! XML grammar remains below the semantic facade.

pub use crate::slicer_cache::views::{Slicer, SlicerExtensionList, Slicers};
pub use crate::slicer_cache::{Cache, Data, DataKind, Definition, ExtensionList, PivotTable};

/// One worksheet-owned slicers part and its relationship identity.
pub type Part = crate::slicer_cache::views::SlicerPart;
