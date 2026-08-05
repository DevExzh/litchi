//! Package implementation for legacy Word documents (.doc).

mod codec;
mod model;
#[cfg(test)]
mod tests;

pub use model::{EncryptionKind, Error, OpenOptions, Package, Result};
