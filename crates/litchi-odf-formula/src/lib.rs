//! Standalone `OpenDocument` Formula (`.odf` and `.otf`) support.
//!
//! [`Package`] is the immutable owner of one Formula artifact and owns every
//! state verb (ingress, [`Edit`], [`Package::apply`] and output);
//! [`Formula`] is its read-only `MathML` projection. The crate keeps `MathML`
//! tree ownership, XML validation, package handling, authoring, and the
//! public facade in separate layers. Formula markup is inert data: this crate
//! never evaluates it.
//!
//! The primary transaction family — [`Edit`], [`Patch`], [`Commit`],
//! [`Diagnostics`], [`Error`] and [`LimitKind`] — is exported from the crate
//! root. The `MathML` parser's own finite limits live in [`codec`]
//! ([`codec::Limits`], [`codec::LimitError`], [`codec::LimitKind`]).

#![forbid(unsafe_code)]

pub mod authoring;
pub mod codec;
pub mod facade;
pub mod metadata;
pub mod model;
mod package;

pub use authoring::Builder;
pub use facade::{
    ChangeKind, Commit, CommitRecord, DependencyConflict, DependencyConflictKind,
    DependencyTransfer, Diagnostics, Edit, Formula, History, MAX_COMMIT_HISTORY,
    MAX_COMMIT_HISTORY_BYTES, MAX_SEMANTIC_OPERATIONS, MAX_STARMATH_SOURCE_BYTES, NodePath,
    OpaqueStarMath, Package, Patch, Revision, RootChange, SemanticChange, StarMathAnnotation,
    StarMathVersion, ThreeWayPlan,
};
/// The family error of package-level verbs, [`Edit::commit`] and
/// [`Package::apply`]: typed limits plus the common-core payload.
pub use litchi_odf_common::Error;
/// The dimension named by [`Error::Limit`] and [`Error::InvalidLimit`].
pub use litchi_odf_common::LimitKind;
pub use litchi_odf_common::{
    OpenOptions, PreparedPackage, SemanticLimitKind, SemanticLimits, SourceLimitKind, SourceLimits,
};
pub use model::{Attribute, Content, ContentKind, ContentSymbol, Element, Kind};
