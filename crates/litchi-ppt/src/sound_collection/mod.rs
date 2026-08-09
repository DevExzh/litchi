//! Strict, inert `PowerPoint` embedded sound-collection support.
//!
//! The semantic model is kept separate from the MS-PPT record codec. Sound
//! bytes are borrowed from the presentation and are never decoded, played,
//! fetched, or otherwise activated by this owner.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use model::{BuiltinId, Collection, Sound};
