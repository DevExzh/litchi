//! Bounded, inert `PowerPoint` diagram-build metadata.
//!
//! This owner covers the fixed-shape records from MS-PPT §§2.8.13–2.8.14
//! and the `DiagramBuildEnum` vocabulary from §2.13.7.  It only retains
//! metadata; it never renders a diagram or schedules, plays, or executes an
//! animation.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse_bytes, parse_record};
pub use model::{Atom, Build, BuildType, Container, Kind, Limits};
