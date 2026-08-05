//! DrawingML chart XML writer facade.
//!
//! The public writer path remains [`crate::chart::writer`]. Writer-specific
//! series capability state lives in `model`, while XML emission lives in
//! `codec`.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use codec::{write, write_with_rels};
