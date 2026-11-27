//! Low-level legacy PowerPoint decoding and encoding.
//!
//! The canonical [`crate::Package`] owner keeps source identity and
//! publication policy separate from this generic `Read + Seek` projection.
//! Callers that need the native mutable reader should use this explicit
//! namespace.

/// The legacy generic PowerPoint reader.
pub use crate::package::model::Reader;

/// Encode a mutable native [`crate::Builder`] into a seekable destination.
///
/// This is the deliberately low-level escape hatch for codec and harness
/// integrations. Normal callers should use [`crate::Builder::build`] and then
/// publish the returned immutable [`crate::Package`]. Encrypted output still
/// requires an explicit one-use `EncryptionEntropy` token on the
/// builder; this function never obtains ambient randomness.
pub fn encode<W: std::io::Write + std::io::Seek>(
    builder: &mut crate::Builder,
    writer: &mut W,
) -> Result<(), crate::WriteError> {
    crate::writer::core_encode(builder, writer)
}
