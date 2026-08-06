//! Typed `[MS-OLEPS]` 2.23 property-set stream and storage bindings.
//!
//! A [`Binding`] is the format-neutral association between a property-set
//! FMTID and its standard CFB name.  The codec handles both the six named
//! bindings defined by the specification and the 26-character GUID form used
//! for every other FMTID.  Names are fixed-size, ASCII-compatible values, so
//! callers can pass them to CFB APIs without allocating a temporary `String`.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{Binding, BindingName, GLOBAL_INFO_FMTID, IMAGE_CONTENTS_FMTID, IMAGE_INFO_FMTID};
