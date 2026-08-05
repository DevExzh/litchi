//! Presentation and slide protection values with bounded XML handling.
//!
//! Password verifier publication is intentionally separated from the semantic
//! model. Parsing and inert re-publication are dependency-free; generating a
//! new Office password verifier is reported as an explicit integration error
//! until the crate-level crypto dependency set is wired.

mod codec;
mod model;

pub use model::{Algorithm, Settings, Slide, Type, Verifier};
