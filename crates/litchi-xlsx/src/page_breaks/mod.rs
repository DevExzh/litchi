//! Typed, inert worksheet page-break ownership.
//!
//! The owner validates the core `rowBreaks` and `colBreaks` vocabulary,
//! applies Office's collection limits, and supports byte-preserving worksheet
//! rewrites. It does not paginate, print, or infer automatic breaks.

mod codec;
pub mod edit;
mod model;
mod patch;
mod snapshot;
pub mod source;

pub use codec::{parse, replace, write};
pub use model::{
    Axis, Break, Collection, LimitKind, MAX_HORIZONTAL_BREAKS, MAX_VERTICAL_BREAKS, PageBreaks,
};
pub(crate) use patch::Patch;
pub use snapshot::Snapshot;

/// Load one source-bound worksheet page-break snapshot.
///
/// # Errors
///
/// Returns an error when the package or selector is invalid, the selected
/// sheet is not a worksheet, or its page-break XML is invalid.
pub fn load<'a>(
    package: &litchi_opc::OpcPackage,
    selector: impl Into<crate::Selector<'a>>,
) -> crate::Result<Snapshot> {
    Snapshot::load(package, selector)
}

#[cfg(test)]
mod tests;
