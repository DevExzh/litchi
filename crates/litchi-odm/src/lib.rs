//! `OpenDocument` Master Document support with semantic responsibility layers.
#![forbid(unsafe_code)]

mod authoring;
mod codec;
mod edit_ops;
mod facade;
pub mod link;
pub mod metadata;
mod model;
mod package;
pub mod title;
pub mod transaction;

pub use facade::{Builder, Document, Package};
pub use litchi_odf_common::{Error, LimitKind, OpenOptions};
pub use model::{resource, section, security, structure, style, subdocument};
