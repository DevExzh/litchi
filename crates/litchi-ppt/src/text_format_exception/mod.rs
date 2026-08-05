//! Layered legacy PowerPoint text-format exception ownership.
//!
//! The facade exposes semantic text defaults, while the codec owns the
//! bounded MS-PPT record representation and the tests exercise both layers.

pub(super) mod codec;
mod model;
#[cfg(test)]
mod tests;

pub use model::{BulletFlags, CFStyle, TextCFException, TextPFException, WrapFlags};
