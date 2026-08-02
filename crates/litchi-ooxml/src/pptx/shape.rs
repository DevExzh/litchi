//! Borrowed semantic PowerPoint shapes.
//!
//! The canonical scene, selectors, and data-bearing [`Shape`] enum live in
//! `litchi-pptx`. Table content uses the separate [`crate::pptx::table`]
//! module, avoiding a collision between a table-shaped view and its grid.

pub use litchi_pptx::shape::*;
