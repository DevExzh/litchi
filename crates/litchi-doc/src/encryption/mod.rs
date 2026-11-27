//! Semantic facade for legacy Word password-to-open encryption.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub(crate) use codec::{
    decrypt_document_streams, encrypt_document_streams_for_write, validate_writer_password,
};
pub use model::EncryptionProfile;
