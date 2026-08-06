//! Layered binary SPRM codecs for TAP properties.
//!
//! The facade keeps the historical `TapParser` associated methods intact while
//! the implementation is organized by the [MS-DOC] operand families.

mod borders;
mod cells;
mod definitions;
mod prelude;
mod primitives;
mod shading;

pub(in crate::parts::tap_parser) use primitives::binary_to_doc_result;
