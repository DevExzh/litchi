//! OfficeArt shape-property parsing (`Opt` records).
//!
//! Properties control shape appearance: position, size, colors, rotation, etc.
//! Based on MS-ODRAW specification section 2.3.
//!
//! # Complex Properties
//!
//! Properties can be simple (4-byte value) or complex (variable-length data).
//! Complex properties use a two-pass parsing approach:
//! 1. First pass: Parse all 6-byte property headers
//! 2. Second pass: Read complex data that follows the headers
//!
//! # Performance
//!
//! - Two-pass parsing minimizes data copying
//! - Zero-copy for complex data (borrows from source)
//! - Wire-order retention plus indexed property lookup
//! - Pre-allocated capacity based on the checked property count
//!
//! The public property types remain available directly under `crate::prop`.

const IS_BLIP: u16 = 0x4000;
const IS_COMPLEX: u16 = 0x8000;
const PROPERTY_ID_MASK: u16 = 0x3FFF;

mod codec;
pub mod geometry;
pub mod gradient;
mod model;
mod package;
pub mod picture;
#[cfg(test)]
mod tests;

pub use model::{Anchor, Array, ColorRef, Id, Prop, UnknownId, Value};
pub use package::Props;
