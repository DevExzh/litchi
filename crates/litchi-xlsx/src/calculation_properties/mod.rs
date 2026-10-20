//! Immutable workbook calculation-properties read model.
//!
//! The facade exposes the typed calculation policy while keeping its
//! SpreadsheetML/MCE representation and regression coverage in dedicated
//! layers.

mod codec;
pub mod edit;
mod features;
mod limits;
mod model;
mod package;
mod patch;
#[allow(
    dead_code,
    reason = "the calculation-properties parser remains an internal rewriter seam"
)]
mod rewriter;
mod snapshot;
mod source;

#[cfg(test)]
mod tests;

pub use codec::{parse, parse_features, parse_features_with_limits, parse_with_limits};
pub use features::{Feature, Features};
pub use limits::{LimitKind, Limits};
pub use model::{Builder, Mode, Properties, ReferenceMode, Specified};
pub use package::load;
pub use patch::{Commit, Patch};
pub use snapshot::Snapshot;
pub use source::{SourceBackedEditor, SourceEdit};

#[allow(
    unused_imports,
    reason = "the facade intentionally retains all calculation-properties codec entry points"
)]
pub(crate) use codec::{Inspection, inspect};
#[allow(
    unused_imports,
    reason = "the facade intentionally retains the calculation-properties rewrite entry point"
)]
pub(crate) use rewriter::{invalidate_formulas, rewrite};
