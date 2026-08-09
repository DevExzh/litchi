//! Standalone `OpenDocument` Formula (`.odf` and `.otf`) support.
//!
//! The crate keeps `MathML` tree ownership, XML validation, package handling,
//! authoring, and the public facade in separate layers. Formula markup is
//! inert data: this crate never evaluates it.

#![forbid(unsafe_code)]

pub mod authoring;
pub mod codec;
pub mod facade;
pub mod model;
pub mod package;

pub use authoring::Builder;
pub use codec::{LimitError, LimitKind, Limits};
pub use facade::{
    ChangeKind, Commit, CommitRecord, DependencyConflict, DependencyConflictKind,
    DependencyTransfer, Diagnostics, Edit, Formula, History, MAX_COMMIT_HISTORY,
    MAX_COMMIT_HISTORY_BYTES, MAX_SEMANTIC_OPERATIONS, MAX_STARMATH_SOURCE_BYTES, NodePath,
    OpaqueStarMath, Patch, Revision, RootChange, SemanticChange, StarMathAnnotation,
    StarMathVersion, ThreeWayPlan,
};
pub use model::{Attribute, Content, Element, Kind};
