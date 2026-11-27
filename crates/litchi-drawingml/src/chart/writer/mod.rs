//! `DrawingML` chart XML writer facade.
//!
//! The public writer path remains [`crate::chart::writer`]. Writer-specific
//! series capability state lives in `model`, while XML emission lives in
//! `codec`.

mod codec;
mod model;
mod semantic;
mod validation;
mod xml;

#[cfg(test)]
mod tests;

pub use codec::{write, write_with_rels};
