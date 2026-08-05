//! Shared [MS-OSHARED] smart-tag property-bag structures.
//!
//! Word and PowerPoint use the same `PropertyBagStore`, but wrap the property
//! bags differently. This module deliberately performs no recognition or
//! schema download; it only decodes inert metadata.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use self::model::{
    Error, Limits, Property, PropertyBag, PropertyBagStore, PropertyBagString,
    PropertyBagStringEncoding, Type,
};
