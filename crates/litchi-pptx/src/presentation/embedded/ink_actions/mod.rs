//! Generic PresentationML owner for the PowerPoint 2014 InkAction MCE
//! extension.
//!
//! The owner resolves only existing slide relationships.  It reuses the
//! source-backed shared DrawingML action profile after validating the exact
//! MCE, OPC relationship, internal target, `text/xml` declaration, and root
//! closure.  It does not infer an InkML MIME/path, execute actions, or create
//! fresh producer-specific parts.

mod codec;
mod edit;
mod graph;
mod model;
mod package;
mod transaction;

pub use edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};
pub use litchi_opc::TargetMode;
pub use model::{
    ActionAnchor, AnchorFingerprint, AnchorSelector, Branch, InboundReference, RawAnchor,
    SlideSelector, Snapshot,
};
#[cfg(test)]
pub(crate) use package::apply_patch;
pub use package::{Limits, default_limits, generic_content_type, load_slide, load_snapshot};
pub(crate) use package::{apply_commit, load_snapshots};
#[cfg(test)]
pub(crate) use transaction::Prepared;
pub use transaction::{Revision, default_profile_limits};

#[cfg(test)]
mod package_tests;
#[cfg(test)]
mod tests;
