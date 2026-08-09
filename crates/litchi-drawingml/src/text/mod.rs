//! Checked, format-neutral `DrawingML` text primitives.
//!
//! The semantic model is shared by every `DrawingML` host. Numeric scalars,
//! lexical codecs, and the neutral text-body owner stay layered beneath this
//! facade so host crates consume one compact, prefix-free vocabulary.

mod codec;
mod model;
mod scalars;

pub mod body;

#[cfg(test)]
mod tests;

pub use codec::{ParseError, parse_bool, parse_on_off};
pub use model::{Anchor, Autofit, Direction, Underline, Wrap};
pub use scalars::{Columns, Coordinate32, TextSize};
