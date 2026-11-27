//! Typed semantic models for `OfficeArt` properties.
//!
//! The facade keeps the concise `prop::{Id, Prop, Value, ...}` API while the
//! implementation is divided by ownership: identifier vocabulary, borrowed
//! value views, property-table records, and shape geometry are independent
//! semantic groups.

mod geometry;
mod identifiers;
mod property;
mod values;

pub(super) use super::{IS_BLIP, IS_COMPLEX, PROPERTY_ID_MASK};

pub use geometry::Anchor;
pub use identifiers::{Id, UnknownId};
pub use property::Prop;
pub use values::{Array, ColorRef, Value};
