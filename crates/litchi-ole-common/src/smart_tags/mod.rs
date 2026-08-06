//! Shared [MS-OSHARED] smart-tag property-bag structures.
//!
//! Word and PowerPoint use the same `PropertyBagStore`, but wrap the property
//! bags differently. This module deliberately performs no recognition or
//! schema download; it only decodes inert metadata.

mod codec;
mod model;
mod patch;
mod snapshot;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use self::model::{
    Error, Limits, Property, PropertyBag, PropertyBagStore, PropertyBagString,
    PropertyBagStringEncoding, Type,
};
pub use self::patch::{Change, Patch};
pub use self::snapshot::Snapshot;
pub use self::transaction::{Commit, Revision, Transaction, update};
