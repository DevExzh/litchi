//! Generic PresentationML owner for the PowerPoint 2014 InkAction MCE
//! extension.
//!
//! The owner resolves only existing slide relationships.  It reuses the
//! source-backed shared DrawingML action profile after validating the exact
//! MCE, OPC relationship, internal target, `text/xml` declaration, and root
//! closure.  It does not infer an InkML MIME/path, execute actions, or create
//! fresh producer-specific parts.

mod codec;
mod graph;
mod model;
mod package;
mod transaction;

pub use litchi_opc::TargetMode;
pub use model::{
    ActionAnchor, AnchorFingerprint, AnchorSelector, Branch, InboundReference, RawAnchor,
    SlideSelector, Snapshot,
};
pub(crate) use package::load_snapshots;
pub use package::{
    Limits, apply_commit, apply_patch, default_limits, generic_content_type, load_slide,
    load_snapshot,
};
pub use transaction::{Commit, Edit, Patch, Revision, default_profile_limits};

#[cfg(test)]
mod tests;
