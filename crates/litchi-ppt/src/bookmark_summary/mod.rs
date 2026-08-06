//! Strict, inert PowerPoint document bookmark-summary metadata.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{Bookmark, Summary};
