//! Typed terminal structure of a legacy PowerPoint document container.

mod codec;
mod model;

pub use model::{CustomTableStylesPlacement, DocumentStructure};

#[cfg(test)]
mod tests;
