//! Package implementation for legacy Word documents (.doc).

mod codec;
mod model;
#[cfg(test)]
mod tests;

pub use model::{DocEncryptionKind, DocError, DocOpenOptions, Package, Result};
