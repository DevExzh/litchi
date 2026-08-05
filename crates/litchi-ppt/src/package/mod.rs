//! Package implementation for legacy PowerPoint presentations (`.ppt`).

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use model::{EncryptionKind, Error, OpenOptions, Package, Result};
