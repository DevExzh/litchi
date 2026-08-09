//! Standalone `OpenDocument` Formula (`.odf` and `.otf`) support.
//!
//! The crate keeps `MathML` tree ownership, XML validation, package handling,
//! authoring, and the public facade in separate layers. Formula markup is
//! inert data: this crate never evaluates it.

#![forbid(unsafe_code)]
#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "formula transaction artifacts stay adjacent to their encode/decode phases"
)]

pub mod authoring;
pub mod codec;
pub mod facade;
pub mod model;
pub mod package;

pub use authoring::Builder;
pub use codec::{LimitError, LimitKind, Limits};
pub use facade::{
    ChangeKind, Commit, Diagnostics, Edit, Formula, History, NodePath, Patch, Revision, RootChange,
    SemanticChange, StarMathAnnotation, StarMathVersion,
};
pub use model::{Attribute, Content, Element, Kind};
