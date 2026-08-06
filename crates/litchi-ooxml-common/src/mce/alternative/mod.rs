//! Typed, lossless ownership of one MCE `AlternateContent` element.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use codec::read;
pub use model::{Alternatives, Branch, Choice, Fallback, Limits};
