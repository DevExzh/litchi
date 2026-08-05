//! Column layout for one Word section.
//!
//! The semantic model lives in [`model`].  Binary `SEPX` encoding remains in
//! the section writer, so callers can work with checked column values without
//! depending on SPRM opcodes or byte layout.

mod model;

pub(crate) use model::WireView;
pub use model::{Column, Error, Layout};
