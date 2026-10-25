//! Lossless-bounded `PresentationML` package facade.
//!
//! The public package owner is split into a typed semantic facade, OPC graph
//! lifecycle/serialization, and focused regression tests. Consumers continue
//! to use `crate::package` and the re-exported [`Package`] type.

mod codec;
#[cfg(feature = "encryption")]
mod encryption;
mod model;
pub(crate) mod replay_index;
pub(crate) mod signature;

#[cfg(test)]
mod test_seams;
#[cfg(test)]
mod tests;

pub use model::Package;
pub(crate) use model::{
    OpenedOwnerState, resolve_slide, resolve_slide_identity, resolve_slide_position,
};
